/* ofx.gs
 *
 * Contains helper functions for writing out OFX files.
 */


/* Transaction types allowed in ofxTxn().
 * 
 * It's not the full list from the spec, just the ones
 * I consider useful.
 */
const TXN_TYPES = new Set([
  'CREDIT', // Generic credit
  'DEBIT',  // Generic debit
  'XFER',   // Transfer
  'FEE',    // FI fee
  'INT',    // Interest earned or paid
]);

/* Returns the beginning of an OFX file as a string.
 *
 * Parameters:
 *   fileDate: datetime that the data in this file was retrieved
 *   startDate: start of datetime range that this report covers
 *   endDate: end of datetime range that this report covers
 *   bankId: max 9 alphanumeric characters
 *   acctId: max 22 alphanumeric characters
 * 
 * For example, if I create the report at 8am EDT on 08/24/2024, and I asked
 * for transactions from 01/01/2023 through 01/31/2023, you'd pass in dates
 * like this:
 *    fileDate = "2024-08-24T08:00:00-0400"
 *    startDate = "2024-01-01T00:00:00-0400"
 *    endDate = "2024-02-01T00:00:00-0400"
 */
function ofx_makeHeader(fileDate, startDate, endDate, bankId='00', acctId='00') {
    ret =
`OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SONRS>
<STATUS><CODE>0<SEVERITY>INFO</STATUS>
<DTSERVER>${ofx_date_(fileDate)}
<LANGUAGE>ENG
</SONRS>
<BANKMSGSRSV1><STMTTRNRS><TRNUID>0<STATUS><CODE>0<SEVERITY>INFO</STATUS>
<STMTRS>
  <CURDEF>USD
  <BANKACCTFROM>
    <BANKID>${ofx_escape_(bankId.substring(0,9))}
    <ACCTID>${ofx_escape_(acctId.substring(0,22))}
    <ACCTTYPE>CHECKING
  </BANKACCTFROM>
  <BANKTRANLIST>
    <DTSTART>${ofx_date_(startDate)}
    <DTEND>${ofx_date_(endDate)}`;
    return ret;
}

/* Encodes a transaction into OFX, returns as string.
 *
 * Call this multiple times, and append the results after
 * a single call to ofxHeader().
 * 
 * Parameters:
 *   type: OFX bank transaction type (see TXN_TYPES at top for list)
 *   date: date on which the transaction was posted to the account
 *   amount: amount of transaction
 *   id: unique transaction ID issued by financial institution (max 255 chars)
 *   name: name of payee or txn description (max 32 chars)
 *   memo: additional info not in name (max 255 chars)
 */
function ofx_makeTxn(type, date, amount, id, name, memo) {
  type = type.toUpperCase();
  if (!TXN_TYPES.has(type)) {
    throw new Error('Given transaction type ${type} is not valid.');
  }
  ret = `
    <STMTTRN>
      <TRNTYPE>${type}
      <DTPOSTED>${ofx_date_(date)}
      <TRNAMT>${amount.toFixed(2)}
      <FITID>${ofx_escape_(id.substring(0,255))}`;
  if (name) {
    ret += `
      <NAME>${ofx_escape_(name.substring(0,32))}`;
  }
  if (memo) {
    ret += `
      <MEMO>${ofx_escape_(memo.substring(0,255))}`;
  }
  ret += `
    </STMTTRN>`;
  return ret;
}


/* Returns the end of an OFX file as a string.
 *
 * Append this after the list of OFX transactions.
 * 
 * Parameters:
 *   balanceAmount: balance in account after all the transactions.
 *   asOfDate: datetime at which the account balance was the above amount.
 */
function ofx_makeFooter(balanceAmount, asOfDate) {
  return `
  </BANKTRANLIST>
  <LEDGERBAL>
      <BALAMT>${balanceAmount.toFixed(2)}
      <DTASOF>${ofx_date_(asOfDate)}
  </LEDGERBAL>
</STMTRS>
</STMTTRNRS></BANKMSGSRSV1></OFX>`;
}


/* Returns the given string as a binary blob with the correct MIME
 * type for this OFX file. This blob can then be attached to an email,
 * saved to drive, whatever.
 * 
 * Parameters: {
 *  str: string containing OFX file
 *  name: name to give the blob for when it's saved as a file
 * }
 */
function ofx_makeBlob(str, name) {
  if (!name.endsWith('.ofx')) {
    name += '.ofx';
  }
  return Utilities.newBlob(str, 'application/x-ofx;version="1.02"', name);
}



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Helper functions. */

/* Escape the given string according to OFX's requirements for text data
 * in element payloads.
 */
function ofx_escape_(str) {
  return str
        .replaceAll('&','&amp;')
        .replaceAll('<','&lt;')
        .replaceAll('>','&rt;');
}

/* Convert the given date string or unix timestamp to OFX's datetime
 * format. The returned string is always converted to UTC first.
 */
function ofx_date_(date) {
  return Utilities.formatDate(new Date(date), 'GMT', 'yyyyMMddHHmmss.SSS')
    + '[-0:GMT]'; // Consumers are supposed to default the timezone to GMT, but I'm untrusting.
}
