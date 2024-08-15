/* stripe_http.js
 *
 * Maintains an authenticated HTTP connection with Stripe, and allows you to
 * make HTTP requests across it. Handles timeouts due to lock conflicts or
 * rate limiting by reissuing requests along an exponential backoff schedule.
 * 
 * Requires the following script properties to be set manually in your
 * Google script settings:
 *   stripe_client_secret
 */

/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Constants and internal state.
 */
const stripe_defaultContentType_ = 'application/x-www-form-urlencoded';

const stripe_defaultHeaders_ = {
  'Stripe-Version': '2024-06-20', // update this periodically, but make sure you test it.
  'Accept': 'application/json',
  //Authorization header added by stripe_http_getSecret_().
};

const STRIPE_SECRET_key = 'stripe_client_secret';

const STRIPE_RETRY_SCHED_ms = [ // Exponential backoff with factor of 4.
      15 * 1000, // 15 seconds delay
      60 * 1000, // 1 minute delay
  4 * 60 * 1000, // 4 minutes delay
];



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Public functions.
 */

function stripe_http_fetch(url, options={}) {
  stripe_http_getSecret_();

  // Use internal defaults for contentType and headers, unless specifically
  // passed in by user.
  options.contentType = options.contentType ?? stripe_defaultContentType_;
  options.headers = options.headers ?? {};
  for (const [header, defaultValue] of Object.entries(stripe_defaultHeaders_)) {
    options.headers[header] = options.headers[header] ?? defaultValue;
  }

  // Silence HTTP exceptions so we can handle lock timeouts, but
  // record whether or not the caller wanted them.
  const muteRequested = options.muteHttpExceptions ?? false;
  options.muteHttpExceptions = true;

  let resp = UrlFetchApp.fetch(url, options);
  
  // If hit a rate limit or lock timeout, try a couple more times, according to
  // the schedule defined in DELAY_SCHED_ms.
  let attempt = 0;
  while (resp.getResponseCode() == 429) {
    // If we're still getting the error after completing all retries, throw err.
    if (attempt >= STRIPE_RETRY_SCHED_ms.length) {
      console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
      throw new Error('exhausted all retry attempts, giving up.');
    }
    // Sleep.
    const delay_ms = STRIPE_RETRY_SCHED_ms[attempt];
    console.log(`timed out waiting for lock - sleeping ${delay_ms/1000} (s)`
      + ` before next attempt (retry attempt #${attempt + 1}).`);
    Utilities.sleep(delay_ms);
    // Try again.
    resp = UrlFetchApp.fetch(url, options);

    attempt++;
  }

  // If caller didn't want to mute HTTP exceptions and one's there, throw error.
  const code = resp.getResponseCode();
  if (!muteRequested && code >= 400) {
    if (resp.getContentText()) {
      console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
    }
    throw new Error(`HTTP request to ${url} `
      + `failed with status code ${code}.`);
  }

  return resp;
}
  


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Internal helper functions.
 */

/* stripe_http_getSecret_
 *
 * Stripe uses the Basic Auth mechanism for all requests.
 * 
 */
function stripe_http_getSecret_() {
  if (stripe_defaultHeaders_.Authorization) {
    return;
  }

  // Get secret from property storage.
  const ps = PropertiesService.getScriptProperties();
  const secret = ps.getProperty(STRIPE_SECRET_key);

  // Encode it and add as an authentication header.
  stripe_defaultHeaders_.Authorization
    = 'Basic ' + Utilities.base64Encode(secret);

  console.log(`retrieved secret from properties: '${STRIPE_SECRET_key}'`);
}


// Clear internal connection state. Useful only for debugging.
function stripe_http_reset() {
  _defaultHeaders.Authorization = undefined;
}