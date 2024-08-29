
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Constants and internal state.
 */

// comma-separated list of email addresses.
// (REQUIRED)
const RECIPIENTS_KEY = 'email_recipients_list';

// Three letter ISO currency code that you wish to produce reports in.
// ex: "USD", "CNY", etc.
// (OPTIONAL - defaults to USD)
const CURRENCY_KEY = 'currency';

// Mode to produce reports in. 'gross' or 'net'
// (OPTIONAL - defaults to 'gross')
const MODE_KEY = 'mode';

// stores the start date of the next SincePrevious report for a given target.
// key is prefixed by target name - paypal, stripe
// (AUTO-GENERATED)
const SINCE_PREV_START_KEY = '_since_previous_startdate';


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Entry points.
 */

/* doPaypalSinceLast()
 * 
 * Create and send a PayPal report covering the time since the last report,
 * up to now.
 * 
 * If there's no record of a previous report, uses the beginning of the
 * current year as the start date.
 */
function doPaypalSinceLast() {
  main_doSinceLast_("PayPal");
}


/* doStripeSinceLast()
 * 
 * Create and send a Stripe report covering the time since the last report,
 * up to now.
 * 
 * If there's no record of a previous report, uses the beginning of the
 * current year as the start date.
 */
function doStripeSinceLast() {
  main_doSinceLast_("Stripe");
}



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Helper functions.
 */

function main_doSinceLast_(targetPretty) {
  const target = targetPretty.toLowerCase();

  const tz = Session.getScriptTimeZone();

  const ps = PropertiesService.getScriptProperties();

  const now = new Date(Date.now());

  let startDate = ps.getProperty(target + SINCE_PREV_START_KEY);
  if (!startDate) {
    // If there is no recorded previous run of this report, start at the
    // beginning of the current year, in the script's timezone.

    // Figure out how many months we need to go back to get to January in
    // the script's timezone.
    const monthDelta = parseInt(Utilities.formatDate(now, tz, "M")) - 1;

    // Make a Date() object representing some time in January in the
    // script's timezone.
    const jan = new Date(now);
    jan.setMonth(jan.getMonth() - monthDelta);

    // Create a timezone string for Jan 1 in the appropriate timezone BACK THEN.
    // Note that because of daylight savings time, the timezone in January might
    // be different than the one in other parts of the year.
    startDate = Utilities.formatDate(
      jan,
      tz,
      "yyyy-MM'-01T00:00:00'XXX"
    );
    console.log('No previous run detected, using start date: ' + startDate);
  }

  res = main_doReport_(targetPretty, startDate, now);

  // Store the end date of the returned report, so we know where to start the
  // next one. Note that end dates are EXCLUSIVE, so there's no chance of
  // duplicates here.
  if (res) {
    ps.setProperty(target + SINCE_PREV_START_KEY, new Date(res.endDate));
  }
}


// Calls one of the modules to make a report covering the given start and
// end dates, then sends an email to the recipients specified in script
// properties.
function main_doReport_(targetPretty, startDate, endDate=Date.now()) {
  const target = targetPretty.toLowerCase();

  const tz = Session.getScriptTimeZone();

  const ps = PropertiesService.getScriptProperties();

  const currency = ps.getProperty(CURRENCY_KEY) || 'USD';

  const mode = ps.getProperty(MODE_KEY) || 'gross';

  const recipients = ps.getProperty(RECIPIENTS_KEY);
  if (!recipients) {
    throw new Error(
      'Email recipients missing. Please set the script property '
      + RECIPIENTS_KEY + ' to a comma-separated list of the email addresses '
      + 'that you want to send this report to.'
    );
  }

  // Construct a report from previous end date (if any), up to as close to
  // the current time as we have data for.
  let res;
  let targetUrl;
  switch(target) {
    case 'paypal':
      res = paypal_makeReportOfx(startDate, endDate, currency, mode);
      targetUrl = 'https://paypal.com/mep/dashboard';
      targetColor = '#003087';
      break;
    case 'stripe':
      res = stripe_makeReportOfx(startDate, endDate, currency, mode);
      targetUrl = 'https://dashboard.stripe.com';
      targetColor = '#635BFF';
      break;
  }

  // If the start date was so new that there's no data available, pass the
  // null back to this function's caller as well. Don't send any emails or
  // throw any errors - just wait for the next trigger.
  if (!res) {
    console.log(`Skipping ... startDate ${main_prettyDate_(startDate)} `
      + 'was too new, no new data has been published yet.');
    return res;
  }


  const fileDate = Utilities.formatDate(
    new Date(res.reportDate), tz, 'yyyyMMdd_HHmmss'
  );
  const fileName = `${targetPretty}_${fileDate}.ofx`;

  const startDatePretty = main_prettyDate_(res.startDate, tz);
  const endDatePretty = main_prettyDate_(res.endDate, tz);
  const reportDatePretty = main_prettyDate_(res.reportDate, tz);

  const money = new Intl.NumberFormat(Session.getActiveUserLocale(), {
    style: 'currency',
    currency: currency
  });


  // Send email to specified recipients.
  MailApp.sendEmail({
    to: recipients,
    subject: `[Google Bot] ${targetPretty} statement since ${startDatePretty}`,
    name: 'Statement Manager',
    noReply: true,
    attachments: [
      ofx_makeBlob(res.ofx, fileName)
    ],
    htmlBody: `
<html>
<head>
<style>
  table {
    border: solid 12px ${targetColor};
    border-collapse: collapse;
    margin-top: 25px;
    margin-bottom: 25px;
  }
  tr {
    border-bottom: 1px solid ${targetColor};
  }
  th {
    text-align: right;
    padding-left: 10px;
    padding-right: 20px;
  }
  th, td {
    padding-top: 5px;
    padding-bottom: 5px;
  }
  td {
    padding-right: 10px;
  }
  body {
    font-size: 120%;
  }
</style>
</head>
<body>
<h2 style="color:${targetColor}">New Report from Statement Manager for ${targetPretty}</h2>

<table>
  <tr><th>Target</th><td><a href="${targetUrl}">${targetPretty}</a></td></tr>
  <tr><th>Start</th><td>${startDatePretty}</td></tr>
  <tr><th>End</th><td>${endDatePretty}</td></tr>
  <tr><th>Balance</th><td>${money.format(res.balance)}</td></tr>
  <tr><th>Transactions</th><td>${res.numTxns}</td></tr>
</table>

<p>The data used to produce this report was current as of ${reportDatePretty}.

<p><a href="https://qbo.intuit.com/app/newfileupload">
Upload to QuickBooks Online Here</a>

<br>
<hr/>
<p style="font-size: 100%">This email was sent automatically by Google, from the
<a href="https://script.google.com/home/projects/${ScriptApp.getScriptId()}">
Statement Manager</a> Apps Script.
</body>
</html>
`
  });

  return res;
}


function main_prettyDate_(date, tz=Session.getScriptTimeZone()) {
  return Utilities.formatDate(
    new Date(date), tz, 'yyyy-MM-dd HH:mm:ss z'
  );
}


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
