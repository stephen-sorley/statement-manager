/* stripe.js
 *
 * Get transactions and balances over Stripe's REST API, then translate them
 * into OFX.
 * 
 * Requires the following script properties to be set manually in your
 * Google script settings:
 *   stripe_client_secret
 * 
 * You can make secret keys for your Stripe account here:
 * https://dashboard.stripe.com/apikeys
 * 
 * For enhanced security, I recommend making a new restricted key and only
 * enabling the following read permissions:
 *   - Balance
 *   - Balance transaction sources
 *   - Files
 *   - All Reporting resources (this one's near the bottom)
 */


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Constants and internal state.
 */

const STRIPE_RESPONSE_SIZE = 25;

const STRIPE_BASEURL = 'https://api.stripe.com';

const STRIPE_REPORT = 'ending_balance_reconciliation.summary.1';


function stripe_test() {
  const startDate = "2024-06-01T00:00:00-0400";
  const endDate = "2024-08-01T00:00:00-0400";
  stripe_makeReportOfx(startDate, endDate);
}


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Public Functions.
 */

/* stripe_makeReportOfx
 *
 * Produce a Stripe bank statement containing all balance-affecting transacions
 * over the given time interval.
 * 
 * Currently, this function only supports producing reports for one currency at
 * a time.
 * 
 * Parameters:
 *   startDate: datetime where the report begins (inclusive).
 *   endDate: datetime where the report ends (exclusive). Default: current time
 *   currency: only report txns & balances done in this currency. Default: USD
 * 
 * Returns: {
 *   reportDate: data current as of this date
 *   startDate: report start date (adjusted for data availability)
 *   endDate: report end date (adjusted for data availability)
 *   balance: balance as of the report end date
 *   numTxns: number of transactions that occurred in the report interval
 *   ofx: a string representing the full report, formatted as OFX data.
 * }
 * 
 * Returns 'null' if the start date was so new that Stripe doesn't have data
 * available yet. Stripe's publishing interval may be up to 24 hours.
 */
function stripe_makeReportOfx(startDate, endDate=Date.now(), currency='USD') {
  
  res = stripe_getTransactions_(
    startDate,
    endDate,
    currency
  );

  // If the start date was so new that there's no data available, pass the
  // null back to this function's caller as well.
  if (!res) {
    return res;
  }

  ofx = ofx_makeHeader(
    res.reportDate,
    res.startDate,
    res.endDate,
    "Stripe",
    "dashboard.stripe.com",
    currency
  );

  /*
    Format conversions for Stripe transaction objects:
  
    Datetime: uses Unix timestamps in seconds, must convert to milliseconds
              to be parseable by Javascript.
  
    Amount: uses number of cents, must divide by 100 to get number of dollars.
  */
   
  for (const txn of res.txns) {
    const date = txn.created * 1000;
    const amountGross = txn.amount / 100;
    const amountFee = txn.fee / 100;

    let memo = [];

    const src = txn.source ?? {};
    const billing = src.billing_details ?? {};

    if (billing.name) {
      memo.push(billing.name);
    }

    if (billing.email) {
      memo.push(billing.email);
    }

    memo.push(txn.reporting_category);

    memo.push(src.id ?? txn.id);

    if (src.customer) {
      memo.push('PAYER:' + src.customer);
    }

    if (src.destination) {
      memo.push('BANK:' + src.destination);
    }

    ofx += ofx_makeTxn(
      stripe_ofxTxnCode_(txn.reporting_category, amountGross),
      date,
      amountGross,
      txn.id,
      txn.description ?? txn.object,
      memo.join(' // ')
    );

    if (amountFee != 0) {
      ofx += ofx_makeTxn(
        "FEE",
        date,
        amountFee,
        txn.id + '-1',
        'Stripe processing fees',
        'for:' + src.id ?? ti.transaction_id
      );
    }
  }

  ofx += ofx_makeFooter(res.balance, res.endDate);

  return {
    reportDate: res.reportDate,
    startDate: res.startDate,
    endDate: res.endDate,
    balance: res.balance,
    numTxns: res.txns.length,
    ofx: ofx,
  };
}
  
  
  
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 * Internal helper functions.
 */

function stripe_getTransactions_(startDate, endDate, currency='USD') {
  // Stripe accepts all datetimes as Unix timestamps (seconds since 1970).
  // Javascript timestamps are in milliseconds, so we need to convert.
  startDate = stripe_timestamp(startDate);
  endDate = stripe_timestamp(endDate);

  if (startDate > endDate) {
    throw new Error('invalid dates, startDate is later than endDate.');
  }

  // Check to see what date range data is available for.
  const avail = stripe_getAvailableDates_();

  // If start date is so new that Stripe doesn't have any data available yet.
  if (startDate >= avail.endDate) {
    return null; // indicates to caller that no data is available yet
  }

  // Clamp report date range to available interval.
  endDate = Math.min(endDate, avail.endDate);

  const out = {
    startDate: startDate * 1000,
    endDate: endDate * 1000,
    reportDate: avail.endDate * 1000,
    balance: null,
    txns: null,
  };

  console.log(
  `  start: ${new Date(out.startDate)}
  end: ${new Date(out.endDate)},
  report: ${new Date(out.reportDate)}`); //DEBUG_161

  // Start generating an ending balance report at our new end date.
  // Wait to pull the report till after we pull transaction history,
  // to give it a little time to generate.
  const reportId = stripe_requestBalanceReport_(endDate, currency);

  // Request transactions. Response may require multiple pages.
  let results = [];
  let startAfter;
  do {
    const params = {
      limit: STRIPE_RESPONSE_SIZE,
      currency: 'USD',
      'expand[]': 'data.source', // include details looked up from original txn
      'created[gte]': startDate, //inclusive (>=)
      'created[lt]': endDate, //exclusive (<)
    };
    if (startAfter) {
      params.starting_after = startAfter;
    }
    const url = main_buildUrl(STRIPE_BASEURL + '/v1/balance_transactions', params);

    const resp = stripe_http_fetch(url);
    const data = resp.json.data;

    if (data && data.length > 0) {
      results.push(data);
    }

    startAfter = (resp.json.has_more) ? data[data.length - 1].id : null;
  } while(startAfter);

  // Collapse all transactions into a flat array, instead of an array of arrays.
  out.txns = results.flat();

  // Sort in-place in ascending order, by creation date.
  out.txns.sort((a,b) => a.created - b.created);

  for (const txn of out.txns) {
    console.log(JSON.stringify(txn, null, 2));
  }

  // Get balance as of the end date from the report we asked to have made
  // earlier. Save it to output.
  out.balance = stripe_getBalanceFromReport_(reportId);

  return out;
}


function stripe_getAvailableDates_() {
  const resp = stripe_http_fetch(STRIPE_BASEURL + '/v1/reporting/report_types/'
    + STRIPE_REPORT
  );

  const out = {
    startDate: resp.json.data_available_start,
    endDate: resp.json.data_available_end,
  }

  console.log('Available Dates\n' + JSON.stringify(out, null, 2));

  return out;
}


function stripe_requestBalanceReport_(date, currency='USD') {
  const payload = {
    report_type: STRIPE_REPORT,
    'parameters[currency]': currency,
    'parameters[interval_end]': ''+date,
    'parameters[columns[0]]': 'reporting_category',
    'parameters[columns[1]]': 'net'
  };

  const resp = stripe_http_fetch(STRIPE_BASEURL + '/v1/reporting/report_runs', {
    payload: payload
  });

  if (resp.json.status === 'failed') {
    console.err(JSON.stringify(resp.json, null, 2));
    throw new Error('balance report failed: ' + resp.json.error);
  }

  return resp.json.id;
}


// This returns the balance as the number of dollars (x.xx) instead of Stripe's
// normal amount representation (number of cents, xxx).
function stripe_getBalanceFromReport_(id) {
  // Poll while report is pending.
  let resp;
  let delay_ms = 0;
  do {
    if (delay_ms > 0) {
      Utilities.sleep(delay_ms);
    }
    // Increase delay by 15 seconds after each try until we're waiting a full
    // minute between attempts.
    delay_ms = Math.min(delay_ms + 15 * 1000, 60 * 1000);

    resp = stripe_http_fetch(STRIPE_BASEURL + '/v1/reporting/report_runs/' + id);

  } while(resp.json.status === 'pending');

  // Report failed.
  if (resp.json.status === 'failed') {
    console.err(JSON.stringify(resp.json, null, 2));
    throw new Error('balance report failed: ' + resp.json.error);
  }

  // Report succeeded - download contents.
  resp = stripe_http_fetch(resp.json.result.url);
  const csv = resp.getContentText();
  console.log('Balance report:\n' + csv);

  /* Parse the contents. It should look something like this:

     "reporting_category","net"
     "charge","100.00"
     "fee","-0.15"
     "total","-99.85"

     We want the "net" value for the "total" category ("-99.85").
   */
  const table = Utilities.parseCsv(csv, ',');
  let colCategory = 0;              // default to first column
  let colNet = table[0].length - 1; // default to last column
  for (const [i, heading] of table[0].entries()) {
    switch(heading.toLowerCase()) {
      case 'reporting_category': colCategory = i; break;
      case 'net': colNet = i; break;
    }
  }

  let balanceStr = table[table.length - 1][colNet]; // default to last row
  for (const [i, row] of table.entries()) {
    if (row[colCategory].trim().toLowerCase() === 'total') {
      balanceStr = row[colNet];
      console.log(`Using value from row #${i+1}, col #${colNet+1}`);
      break;
    }
  }
  
  return parseFloat(balanceStr);
}


function stripe_timestamp(date) {
  return Math.floor(new Date(date).getTime() / 1000);
}

function stripe_ofxTxnCode_(reportingCategory, amount) {
  const cat = reportingCategory.toLowerCase();
  
  switch(cat) {
    case 'charge':
      return 'PAYMENT';
    case 'fee': // non-payment-related fee
    case 'tax':
      return 'FEE';
    case 'payout':
    case 'payout_reversal':
    case 'topup':
    case 'topup_reversal':
    case 'transfer':
    case 'transfer_reversal':
      return 'XFER';
  }
  return (amount < 0.0)? 'DEBIT' : 'CREDIT';
}
