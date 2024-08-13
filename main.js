// Main entry point when viewed as a web page.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}








function test() {
  const startDate = "2024-01-01T00:00:00-0400";
  const endDate = "2024-02-01T00:00:00-0400";
  let str = ofx_makeHeader(Date.now(), startDate, endDate);

  str += ofx_makeTxn('CREDIT','2024-01-03T8:56:32-0400', 3.56,
    'e8764269-940b-4f88-acbe-3f0aa6f9cd40', 'a name!', 'a memo, too!');

  str += ofx_makeFooter(103.56, '2024-01-31T23:59:59-0400');

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
    attachments: ofxBlob(str, "test.ofx")
  });
}


// Global helper functions.
function main_buildUrl(url, params) {
  var paramString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + paramString;
}
