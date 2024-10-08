/* paypal_http.js
 *
 * Maintains an authenticated HTTP connection with PayPal, and allows you to
 * make HTTP requests across it.
 * 
 * Requires the following script properties to be set manually in your
 * Google script settings:
 *   paypal_client_id
 *   paypal_client_secret
 * 
 * Main functions
 * --------------
 *   
 * * paypal_http_fetch:
 * 
 *   A wrapper around UrlFetchApp.fetch(). Maintains the access token
 *   internally, will reacquire a new token and reissue the failed
 *   HTTP requests automatically.
 * 
 * * paypal_http_fetchAll:
 * 
 *   A wrapper around UrlFetchApp.fetchAll(). Supports the same token
 *   handling as paypal_http_fetch().
 */


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Constants and internal state.
 */
const PAYPAL_ONEHOUR_ms = 1 * 60*60*1000;

const paypal_defaultContentType_ = 'application/json';

const paypal_defaultHeaders_ = {
  'Accept': 'application/json',
  'Accept-Language': 'en_US',
  //Authorization header added by paypal_http_guaranteeToken_().
}

// These are all set or modified by paypal_http_guaranteeToken_().
let paypal_accessToken_;
let paypal_accessTokenExpires_;
let paypal_secret_;



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Public functions.
 */

function paypal_http_fetch(url, options={}) {
  paypal_http_guaranteeToken_();

  // Use internal defaults for contentType and headers, unless specifically
  // passed in by user.
  options.contentType = options.contentType ?? paypal_defaultContentType_;
  options.headers = options.headers ?? {};
  for (const [header, defaultValue] of Object.entries(paypal_defaultHeaders_)) {
    options.headers[header] = options.headers[header] ?? defaultValue;
  }

  // Silence HTTP exceptions so we can handle bad access tokens, but
  // record whether or not the caller wanted them.
  const muteRequested = options.muteHttpExceptions ?? false;
  options.muteHttpExceptions = true;

  let resp = UrlFetchApp.fetch(url, options);
  
  // If our access token has expired, refresh token and try once more.
  if (resp.getResponseCode() == 401) {
    paypal_http_guaranteeToken_(true);
    resp = UrlFetchApp.fetch(url, options);
  }

  // If caller didn't want to mute HTTP exceptions and one's there, throw error.
  const code = resp.getResponseCode();
  if (!muteRequested && code >= 400) {
    console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
    throw new Error(`HTTP request to ${url} `
      + `failed with status code ${code}.`);
  }

  // Try parsing to JSON.
  try {
    resp.json = JSON.parse(resp.getContentText());
  } catch(e) {
    resp.json = null;
  }

  return resp;
}
  
  
function paypal_http_fetchAll(requests) {
  if (!requests || requests.length == 0) {
    return [];
  }

  paypal_http_guaranteeToken_();

  for (const request of requests) {
    // Silence HTTP exceptions so we can handle bad access tokens, but
    // record whether or not the caller wanted them.
    request.muteRequested = request.muteHttpExceptions ?? false;
    request.muteHttpExceptions = true;

    // Use internal defaults for contentType and headers, unless specifically
    // passed in by user.
    request.options.contentType =
      request.options.contentType ?? paypal_defaultContentType_;
    request.options.headers = request.options.headers ?? {};
    for (const [header, defaultValue] of Object.entries(paypal_defaultHeaders_)) {
      request.options.headers[header] =
        request.options.headers[header] ?? defaultValue;
    }
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
    paypal_http_guaranteeToken_(true);
    let new_resps = UrlFetchApp.fetchAll(requests.slice(i,requests.length));
    // Replace the old responses that we redid with new ones.
    resps.splice(i, new_resps.length, ...new_resps);
  }

  
  for (const [i, resp] of resps.entries()) {
    // Try parsing to JSON.
    try {
      resp.json = JSON.parse(resp.getContentText());
    } catch(e) {
      resp.json = null;
    }

    // If the caller left HTTP exceptions enabled and we saw a bad response
    // code, throw an error.
    const muteRequested = requests[i].muteRequested;
    const code = resp.getResponseCode();
    if (!muteRequested && code >= 400) {
      console.error(JSON.stringify(resp.json, null, 2));
      throw new Error(`HTTP request to ${requests[i].url} `
        + `failed with status code ${code}.`);
    }
  }

  return resps;
}



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Internal helper functions.
 */

/* paypal_http_guaranteeToken_
 *
 * PayPal uses the "client_credentials" flow from OAuth 2.0.
 * 
 * 1) Stop if token in memory and not being forced.
 * 2) Load client secret and any stored previous token from properties.
 * 3) Stop if valid token loaded from properties and not being forced.
 * 4) Request new token from PayPal using client secret.
 * 5) Store new token and expiration date in properties.
 * 
 * Parameters:
 *   force: if true, ignore stored tokens and get a fresh one. Default: false
 */
function paypal_http_guaranteeToken_(force = false) {
  if (!force && paypal_accessToken_) {
    return;
  }

  const ps = PropertiesService.getScriptProperties();

  if (!paypal_secret_) {
    // Get client ID, secret, and any stored access token from property storage.
    const props = ps.getProperties();
    paypal_secret_ = props.paypal_client_id + ':' + props.paypal_client_secret;
    paypal_secret_ = Utilities.base64Encode(paypal_secret_);

    // If there's a stored access token and we've still got at least an hour
    // before it's supposed to expire, use it.
    paypal_accessToken_ = props.paypal_access_token;
    paypal_accessTokenExpires_ =
      new Date(props.paypal_access_token_expires ?? 0).getTime();
    if (!force && paypal_accessToken_ && paypal_accessTokenExpires_>(Date.now()-PAYPAL_ONEHOUR_ms)) {
      paypal_defaultHeaders_['Authorization'] = 'Bearer ' + paypal_accessToken_;
      console.log('loaded access token from properties, expires '
        + new Date(paypal_accessTokenExpires_));
      return;
    }
  }

  paypal_defaultHeaders_['Authorization'] = 'Basic ' + paypal_secret_;

  const requestTime = Date.now();
  const resp = UrlFetchApp.fetch(PAYPAL_BASEURL + '/v1/oauth2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: paypal_defaultHeaders_,
    payload: {
      grant_type: 'client_credentials',
    },
  });

  data = JSON.parse(resp.getContentText());

  paypal_accessToken_ = data.access_token;
  paypal_accessTokenExpires_ = requestTime + data.expires_in * 1000;
  ps.setProperties({
    paypal_access_token: paypal_accessToken_,
    paypal_access_token_expires: new Date(paypal_accessTokenExpires_).toISOString(),
  })

  paypal_defaultHeaders_['Authorization'] = 'Bearer ' + paypal_accessToken_;

  console.log('received new access token from PayPal, expires '
    + new Date(paypal_accessTokenExpires_)
  );
}


// Clear internal connection state. Useful only for debugging.
function paypal_http_reset() {
  paypal_secret_ = undefined;
  paypal_accessToken_ = undefined;
  paypal_accessTokenExpires_ = undefined;

  const ps = PropertiesService.getScriptProperties();
  ps.deleteProperty('paypal_access_token');
  ps.deleteProperty('paypal_access_token_expires');
}