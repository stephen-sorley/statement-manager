// Main entry point when viewed as a web page.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


function test() {
  const startDate = "2024-06-01T00:00:00-0400";
  const endDate = "2024-08-01T00:00:00-0400";

  const paypal = paypal_makeReportOfx(startDate, endDate);
  const stripe = stripe_makeReportOfx(startDate, endDate);

  const startDatePretty = new Date(startDate).toLocaleString('en-US','America/New York');
  const endDatePretty = new Date(endDate).toLocaleString('en-US','America/New York');
  MailApp.sendEmail({
    to: 'treasurer@cccgainesville.org',
    subject: '[TEST] monthly statement',
    name: 'Google Bot',
    noReply: true,
    body: `Monthly statements covering dates:
${startDatePretty} (Eastern Time)
${endDatePretty} (Eastern Time)`,
    attachments: [
      ofx_makeBlob(paypal, "paypal.ofx"),
      ofx_makeBlob(stripe, "stripe.ofx"),
    ],
  });
}


// Global helper functions.

/* 
 * Converts the given 'params' object to a query string, then adds it to
 * the given url.
 * 
 * If the given URL already has a query string, this function will append
 * the new parameters to it.
 * 
 * This function properly encodes nested objects in the params list. For
 * example:
 *   params = {"created": {"gt":0,"lte":2}}
 * 
 *   becomes
 * 
 *   ?created[gt]=0&created[lte]=2
 */
function main_buildUrl(url, params) {
  const makeQueryString = (obj, parentKey = null) => {
    let params = [];
    
    for (const [key, value] of Object.entries(obj)) {
      // Construct the full key name
      const fullKey = parentKey ? `${parentKey}[${key}]` : key;
  
      if (typeof value === 'object' && value !== null) {
        // Recursively process nested objects
        queryString += makeQueryString(value, fullKey);
      } else {
        // Add key-value pair to our list of fields.
        params.push(
          `${encodeURIComponent(fullKey)}=${encodeURIComponent(value)}`
        );
      }
    }
    return params.join('&');
  }

  return url + (url.indexOf('?') >= 0 ? '&' : '?') + makeQueryString(params);
}
