// Main entry point when viewed as a web page.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


function test() {
  const startDate = "2024-06-01T00:00:00-0400";
  const endDate = "2024-08-01T00:00:00-0400";

  const str = paypal_makeReportOfx(startDate, endDate);

  const startDatePretty = new Date(startDate).toLocaleString('en-US','America/New York');
  const endDatePretty = new Date(endDate).toLocaleString('en-US','America/New York');
  MailApp.sendEmail({
    to: 'treasurer@cccgainesville.org',
    subject: '[TEST] monthly statement',
    name: 'Google Bot',
    noReply: true,
    body: `Monthly statement covering dates:
${startDatePretty}
${endDatePretty}`,
    attachments: ofx_makeBlob(str, "test.ofx")
  });
}


// Global helper functions.
function main_buildUrl(url, params) {
  var paramString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + paramString;
}
