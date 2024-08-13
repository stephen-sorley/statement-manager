const baseUrl = 'https://api-m.paypal.com';
//const baseUrl = 'https://api-m.sandbox.paypal.com';

const oneHour_ms = 1 * 60*60*1000;

let defaultHeaders_; // set by _paypalRefreshToken()
let accessToken_; // set by _paypalRefreshToken()
let accssTokenExpires_; //set by _paypalRefreshToken()
let secret_; // set by _paypalRefreshToken()


function paypal_getTransactions(startDate="2024-07-30T00:00:00-0400", endDate=Date.now()) {
  paypal_refreshToken_();

  let start = new Date(startDate);
  let end = new Date(endDate);

  console.log(`Start: ${start.toISOString()}, End: ${end.toISOString()}`);

  let page = 1;
  let totalPages = 1;
  let requests = [];
  let results = [];
  do {
    const options = {
      url: main_buildUrl(baseUrl + '/v1/reporting/transactions', {
        // Query parameters:
        'start_date': start.toISOString(),
        'end_date': end.toISOString(),
        'page': page,
        'page_size': 1, //DEBUG_161
      }),
      contentType: 'application/json',
      headers: defaultHeaders_,
    };

    if (page === 1) {
      const resp = paypal_urlFetch(options.url, options);
      const data = JSON.parse(resp.getContentText());
      totalPages = data.total_pages;
      results.push(data.transaction_details);
    } else {
      requests.push(options);
    }
    page++;
  } while(page <= totalPages);

  if (requests.length > 0) {
    // If we need to do a bunch of requests, do them in parallel for speed.
    const resps = paypal_urlFetchAll(requests);
    for (const resp of resps) {
      data = JSON.parse(resp.getContentText());
      results.push(data.transaction_details);
    }
  }

  // DEBUG_161 BEGIN
  for (result of results) {
    console.log(JSON.stringify(result,null,2));
  }
  // DEBUG_161 END

  return results;
}


function paypal_refreshToken_(force = false) {
  if (!force && accessToken_) {
    console.log('using access token in memory');
    return;
  }
  
  defaultHeaders_ = {
    'Accept': 'application/json',
    'Accept-Language': 'en_US',
  }

  const ps = PropertiesService.getScriptProperties();

  if (!secret_) {
    // Get client ID, secret, and any stored access token from property storage.
    // (this particular credential is set up to only allow read access to PayPal)
    const props = ps.getProperties();
    secret_ = props.paypal_client_id + ':' + props.paypal_client_secret;
    secret_ = Utilities.base64Encode(secret_);

    // If there's a stored access token and we've still got at least an hour
    // before it's supposed to expire, use it.
    if (!force && props.paypal_access_token && props.paypal_access_token_expires) {
      if (props.paypal_access_token_expires > (Date.now() - oneHour_ms)) {
        accessToken_ = props.paypal_access_token;
        accessTokenExpires_ = props.paypal_access_token_expires;
        defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;
        console.log(`using access token from properties: ${accessToken_}`); //DEBUG_161
        return;
      }
    }
  }

  defaultHeaders_['Authorization'] = 'Basic ' + secret_;

  const requestTime = Date.now();
  const resp = UrlFetchApp.fetch(baseUrl + '/v1/oauth2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: defaultHeaders_,
    payload: {
     grant_type: 'client_credentials',
    },
  });

  data = JSON.parse(resp.getContentText());

  accessToken_ = data.access_token;
  accessTokenExpires_ = requestTime + data.expires_in * 1000;
  ps.setProperties({
    paypal_access_token: accessToken_,
    paypal_access_token_expires: accessTokenExpires_,
  })

  defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;

  console.log(`Got new access token, expires ${new Date(accessTokenExpires_).toISOString()}`);
}


function paypal_urlFetch(url, options={}) {
  // Silence HTTP exceptions so we can handle bad access tokens, but
  // record whether or not the caller wanted them.
  const muteRequested = options.muteHttpExceptions ?? false;
  options.muteHttpExceptions = true;

  let resp = UrlFetchApp.fetch(url, options);
  
  // If our access token has expired, refresh token and try once more.
  if (resp.getResponseCode() == 401) {
    paypal_refreshToken_(true);
    resp = UrlFetchApp.fetch(url, options);
  }

  // If caller didn't want to mute HTTP exceptions and there is one, throw an error.
  const code = resp.getResponseCode();
  if (!muteRequested && code >= 400) {
    console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
    throw new Error(`HTTP request to ${url} `
      + `failed with status code ${code}.`);
  }

  return resp;
}


function paypal_urlFetchAll(requests) {
  if (!requests || requests.length == 0) {
    return;
  }
  // Silence HTTP exceptions so we can handle bad access tokens, but
  // record whether or not the caller wanted them.
  for (const request of requests) {
    request.muteRequested = request.muteHttpExceptions ?? false;
    request.muteHttpExceptions = true;
  }

  let resps = UrlFetchApp.fetchAll(requests);

  let needsRefresh = false;
  let startIndex = requests.length;
  for (const [i, resp] of resps.entries()) {
    if (resp.getResponseCode() == 401) {
      needsRefresh = true;
      startIndex = i;
      break;
    }
  }
  
  // If our access token has expired, refresh token and try all the
  // requests from the failure onward over again.
  if (needsRefresh) {
    paypal_refreshToken_(true);
    let new_resps = UrlFetchApp.fetchAll(requests.slice(i,requests.length));
    // Replace the old responses that we redid with new ones.
    resps.splice(i, new_resps.length, ...new_resps);
  }

  // If caller didn't want to mute HTTP exceptions and there is one, throw an error.
  for (const [i, resp] of resps.entries()) {
    const muteRequested = requests[i].muteRequested;
    const code = resp.getResponseCode();
    if (!muteRequested && code >= 400) {
      console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
      throw new Error(`HTTP request to ${requests[i].url} `
        + `failed with status code ${code}.`);
    }
  }

  return resps;
}
