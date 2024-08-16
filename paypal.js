/* paypal.js
 *
 * Get transactions and balances over PayPal's REST API, then translate them
 * into OFX.
 * 
 * Paypal's transaction history only updates every 3 hours, so a transaction
 * may take up to that long to show up in the history. Therefore you should
 * wait to request a report until at least three hours after that report's
 * end date.
 * 
 * Requires the following script properties to be set manually in your
 * Google script settings:
 *   paypal_client_id
 *   paypal_client_secret
 * 
 * You can generate these by making a new app with your PayPal account here:
 * https://developer.paypal.com/dashboard/applications/live
 * 
 * This code only requires the "Transaction search" feature. I recommend
 * creating your app credential for this script with all other features
 * disabled, to enhance security.
 */


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Constants and internal state.
 */

const PAYPAL_BASEURL = 'https://api-m.paypal.com';
//const PAYPAL_BASEURL = 'https://api-m.sandbox.paypal.com'; //for debugging only

const PAYPAL_MAX_INTERVAL_ms = 31 * 24 * 60 * 60 * 1000; // 31 days


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Public Functions.
 */

/* paypal_makeReportOfx
 *
 * Produce a PayPal bank statement containing all balance-affecting transacions
 * over the given time interval.
 * 
 * Currently, this function only supports producing reports for one currency at
 * a time.
 * 
 * Note that PayPal only keeps transaction data for 3 years, requests for data
 * earlier than this may fail.
 * 
 * Parameters:
 *   startDate: datetime where the report begins.
 *   endDate: datetime where the report ends (exclusive). Default: current time
 *   currency: only report txns & balances done in this currency. Default: USD
 * 
 * Returns: {
 *   reportDate: data current as of this date
 *   startDate: report start date (adjusted for data availability)
 *   endDate: report end date (adjusted for data availability)
 *   balance: balance as of the report end date
 *   ofx: a string representing the report, formatted as OFX data.
 * }
 */
function paypal_makeReportOfx(startDate, endDate=Date.now(), currency='USD') {
  /*
    OFX reports: time interval DOES NOT include endDate.
    PayPal: time interval DOES include endDate, with 1 second resolution.

    Therefore we need to subtract 1 second from the OFX endDate before asking
    PayPal to generate a report, so that it uses the correct time range.
  */
  res = paypal_getTransactions_(
    startDate,
    new Date(endDate).getTime() - 1*1000,
    currency
  );

  // Convert returned time interval back from a PayPal interval definition
  // (inclusive end) to an OFX definition (exclusive end).
  res.endDate = new Date(res.endDate).getTime() + 1*1000;
  
  ofx = ofx_makeHeader(
    res.reportDate,
    res.startDate,
    res.endDate,
    "PayPal",
    res.accountId,
    currency
  );

  for (const txn of res.txns) {
    const ti = txn.transaction_info;
    const pi = txn.payer_info;

    const date = ti.transaction_updated_date ?? ti.transaction_initiation_date;
    const code = ti.transaction_event_code;
    const amountGross = Number(ti.transaction_amount.value);
    const amountFee = ti.fee_amount ? Number(ti.fee_amount.value) : 0;
    const txnTypeName = paypal_ofxTxnTypeName_(code, amountGross);

    let name;
    let memo = [];
    if (pi && pi.payer_name) {
      if (pi.email_address) {
        memo.push(pi.email_address);
      }
      memo.push(ti.transaction_id);

      // Prefer org name.
      name = (pi.payer_name.alternate_full_name ?? "").trim();
      // If not provided, use individual name.
      if (!name) {
        name = ((pi.payer_name.given_name ?? "") + ' '
          + (pi.payer_name.surname ?? "")).trim();
      }
      // If that wasn't present either, use the name of the transaction type.
      if (!name) {
        name = txnTypeName;
      } else {
        memo.push(txnTypeName);
      }
    } else {
      name = txnTypeName;
      memo.push(ti.transaction_id);
    }
    
    if (ti.transaction_subject){
      memo.push(ti.transaction_subject);
    }
    if (ti.paypal_account_id) {
      memo.push('PAYER:' + ti.paypal_account_id);
    }
    if (ti.paypal_reference_id) {
      memo.push(ti.paypal_reference_id_type + ':' + ti.paypal_reference_id);
    }
    if (ti.bank_reference_id) {
      memo.push('BANK:' + ti.bank_reference_id);
    }

    ofx += ofx_makeTxn(
      paypal_ofxTxnCode_(code, amountGross),
      date,
      amountGross,
      ti.transaction_id + '-' + code,
      name,
      memo.join(' // ')
    );

    if (amountFee != 0) {
      ofx += ofx_makeTxn(
        "FEE",
        date,
        amountFee,
        ti.transaction_id + '-1',
        'PayPal',
        'processing fee for:' + ti.transaction_id
      );
    }
  }

  ofx += ofx_makeFooter(res.balance, res.endDate);

  return {
    reportDate: res.reportDate,
    startDate: res.startDate,
    endDate: res.endDate,
    balance: res.balance,
    ofx: ofx,
  };
}



/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Internal helper functions.
 */

// Paypal report end date is INCLUSIVE with 1 SECOND precision.
// This function expects the report interval to be specified in PayPal format.
function paypal_getTransactions_(startDate, endDate, currency='USD') {
  startDate = new Date(startDate).getTime();
  endDate = new Date(endDate).getTime();

  if (startDate > endDate) {
    throw new Error('paypal_getTransactions_: invalid dates, startDate is later'
      + ' than endDate.');
  }

  const out = {
    reportDate: 0,
    startDate: Number.MAX_SAFE_INTEGER,
    endDate: 0,
    balance: undefined,
  };
  let results = [];

  // PayPal only lets us pull transaction data in 31 day chunks.
  // Divide full report interval into chunks, make http requests for each.
  let chunkDate = startDate;
  do {
    // Use as big a chunk as we can (max interval, or end of report, whichever's
    // smaller).
    chunkDate = Math.min(endDate, chunkDate + PAYPAL_MAX_INTERVAL_ms);

    console.log(`Chunk: ${new Date(startDate).toISOString()} to `
      + `${new Date(chunkDate).toISOString()}`); //DEBUG_161

    // Request transactions for this chunk. Response may require multiple pages.
    let page = 1;
    let totalPages = 1;
    let requests = [];
    do {
      const request = {
        url: main_buildUrl(PAYPAL_BASEURL + '/v1/reporting/transactions', {
          // Query parameters:
          'start_date': new Date(startDate).toISOString(),
          'end_date': new Date(chunkDate).toISOString(),
          'fields': 'transaction_info,payer_info',
          'transaction_currency': currency,
          'page': page
        })
      };
  
      if (page === 1) {
        const resp = paypal_http_fetch(request.url, request);
        const data = JSON.parse(resp.getContentText());
        console.log(JSON.stringify(data, null, 2));
        
        totalPages = data.total_pages;
        
        out.accountId = data.account_number;
        out.reportDate = Math.max(
          out.reportDate, Date.parse(data.last_refreshed_datetime));
        out.startDate = Math.min(
          out.startDate, Date.parse(data.start_date));
        out.endDate = Math.max(
          out.endDate, Date.parse(data.end_date));
        
        if (data.transaction_details.length > 0) {
          results.push(data.transaction_details);
        }
      } else {
        requests.push(request);
      }
      page++;
    } while(page <= totalPages);
  
    if (requests.length > 0) {
      // There are additional pages of results after the first, do them in
      // parallel for improved speed.
      const resps = paypal_http_fetchAll(requests);
      for (const resp of resps) {
        data = JSON.parse(resp.getContentText());
        results.push(data.transaction_details);
      }
    }

    // For PayPal, end datetime of a reporting interval is inclusive with 1
    // second resolution. So need to add a second for the next interval's start.
    startDate = chunkDate + 1000;
  } while (startDate < endDate); // Exit loop if past the end of the report interval.

  // Get balance as of the end date, save to output.
  const url = main_buildUrl(PAYPAL_BASEURL + '/v1/reporting/balances', {
    // Query parameters:
    as_of_time: new Date(out.endDate).toISOString(),
    currency_code: currency
  });
  const resp = paypal_http_fetch(url);
  const data = JSON.parse(resp.getContentText());
  console.log(JSON.stringify(data, null, 2));

  out.balance = Number(data.balances[0].total_balance.value);

  // If there were no transactions in the reporting period, return early.
  if (results.length == 0) {
    console.log(JSON.stringify(out, null, 2)); //DEBUG_161
    return out;
  }

  // Collapse all transactions into a flat array, instead of an array of arrays.
  out.txns = results.flat();

  // Sort txns in-place in ascending order, by updated date. Use initiation date
  // if updated date is not present.
  out.txns.sort((a,b) => {
    const ia = a.transaction_info;
    const ib = b.transaction_info;
    const ta = 
      Date.parse(ia.transaction_updated_date ?? ia.transaction_initiation_date);
    const tb = 
      Date.parse(ib.transaction_updated_date ?? ib.transaction_initiation_date);
    return ta - tb;
  });

  console.log(JSON.stringify(out, null, 2)); //DEBUG_161
  return out;
}


function paypal_ofxTxnCode_(code, amount) {
  // 'T0400' -> group is '04'
  const group = code.substring(1,3);
  switch(group) {
    case '00':
      return 'PAYMENT';
    case '01': // non-payment-related fee
      return 'FEE';
    case '03':
    case '04':
    case '17':
    case '20':
    case '22':
      return 'XFER';
    case '08':
    case '14':
      return 'INT';
  }
  return (amount < 0.0)? 'DEBIT' : 'CREDIT';
}


function paypal_ofxTxnTypeName_(code, amount) {
  switch(code) {
    case 'T0002': return 'recurring payment';
    case 'T0013': return 'donation payment';
  }

  const group = code.substring(1,3);
  switch(group) {
    case '00': return 'payment';
    case '01': return 'fee';
    case '03': return 'deposit from bank';
    case '04': return 'withdrawal to bank';
  }

  return (amount < 0.0)? 'account debit' : 'account credit';
}