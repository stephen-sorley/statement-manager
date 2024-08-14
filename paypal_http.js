/* paypal_http.js
 *
 * Maintains an authenticated HTTP connection with PayPal, and allows you to
 * make HTTP requests across it.
 * 
 * Requires the following two script properties to be set manually in your
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
const ONEHOUR_ms = 1 * 60*60*1000;

const defaultContentType_ = 'application/json';

const defaultHeaders_ = {
  'Accept': 'application/json',
  'Accept-Language': 'en_US',
  //Authorization header added by paypal_http_guaranteeToken_().
}

// These are all set or modified by paypal_http_guaranteeToken_().
let accessToken_;
let accessTokenExpires_;
let secret_;



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Public functions.
 */

function paypal_http_fetch(url, options={}) {
  paypal_http_guaranteeToken_();

  // Use internal defaults for contentType and headers, unless specifically
  // passed in by user.
  options.contentType = options.contentType ?? defaultContentType_;
  options.headers = options.headers ?? {};
  for (const [header, defaultValue] of Object.entries(defaultHeaders_)) {
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
      request.options.contentType ?? defaultContentType_;
    request.options.headers = request.options.headers ?? {};
    for (const [header, defaultValue] of Object.entries(defaultHeaders_)) {
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
  if (!force && accessToken_) {
    return;
  }

  const ps = PropertiesService.getScriptProperties();

  if (!secret_) {
    // Get client ID, secret, and any stored access token from property storage.
    const props = ps.getProperties();
    secret_ = props.paypal_client_id + ':' + props.paypal_client_secret;
    secret_ = Utilities.base64Encode(secret_);

    // If there's a stored access token and we've still got at least an hour
    // before it's supposed to expire, use it.
    accessToken_ = props.paypal_access_token;
    accessTokenExpires_ =
      new Date(props.paypal_access_token_expires ?? 0).getTime();
    if (!force && accessToken_ && accessTokenExpires_>(Date.now()-ONEHOUR_ms)) {
      defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;
      console.log('loaded access token from properties, expires '
        + new Date(accessTokenExpires_));
      return;
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
    paypal_access_token_expires: new Date(accessTokenExpires_).toISOString(),
  })

  defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;

  console.log('received new access token from PayPal, expires '
    + new Date(accessTokenExpires_)
  );
}


// Clear internal connection state. Useful only for debugging.
function paypal_http_reset() {
  secret_ = undefined;
  accessToken_ = undefined;
  accessTokenExpires_ = undefined;

  const ps = PropertiesService.getScriptProperties();
  ps.deleteProperty('paypal_access_token');
  ps.deleteProperty('paypal_access_token_expires');
}