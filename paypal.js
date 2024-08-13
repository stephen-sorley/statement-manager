/* paypal.js
 *
 * Functions that connect to PayPal's REST API, pull transaction data from it,
 * and convert those transactions into an OFX file.
 * 
 */

const baseUrl = 'https://api-m.paypal.com';
//const baseUrl = 'https://api-m.sandbox.paypal.com';

const oneHour_ms = 1 * 60*60*1000;

let defaultHeaders_; // set by _paypalRefreshToken()
let accessToken_; // set by _paypalRefreshToken()
let accssTokenExpires_; //set by _paypalRefreshToken()
let secret_; // set by _paypalRefreshToken()


// NOTE: paypal report end date is INCLUSIVE at 1 SECOND precision.
function paypal_getTransactions_(startDate, endDate) {
  paypal_refreshToken_();

  let start = new Date(startDate);
  let end = new Date(endDate);

  console.log(`Start: ${start.toISOString()}, End: ${end.toISOString()}`);

  const out = {};

  let page = 1;
  let totalPages = 1;
  let requests = [];
  let results = [];
  do {
    const options = {
      url: main_buildUrl(baseUrl + '/v1/reporting/transactions', {
        // Query parameters:
        'start_date': start.toISOString(),
        'end_date': end.toISOString(),
        'transaction_currency': "USD",
        'page': page
      }),
      contentType: 'application/json',
      headers: defaultHeaders_,
    };

    if (page === 1) {
      const resp = paypal_urlFetch(options.url, options);
      const data = JSON.parse(resp.getContentText());
      console.log(JSON.stringify(data, null, 2));
      
      totalPages = data.total_pages;
      
      out.accountId = data.account_number;
      out.reportDate = data.last_refreshed_datetime;
      out.startDate = data.start_date;
      out.endDate = data.end_date;
      
      if (data.transaction_details.length > 0) {
        results.push(data.transaction_details);
      }
    } else {
      requests.push(options);
    }
    page++;
  } while(page <= totalPages);

  if (requests.length > 0) {
    // There are additional pages of results after the first, do them in
    // parallel for improved speed.
    const resps = paypal_urlFetchAll(requests);
    for (const resp of resps) {
      data = JSON.parse(resp.getContentText());
      results.push(data.transaction_details);
    }
  }

  // Get balance as of the end date.
  const url = main_buildUrl(baseUrl + '/v1/reporting/balances', {
    // Query parameters:
    as_of_time: out.endDate,
    currency_code: "USD"
  });
  const options = {
    contentType: 'application/json',
    headers: defaultHeaders_,
  };
  const resp = paypal_urlFetch(url, options);
  const data = JSON.parse(resp.getContentText());
  console.log(JSON.stringify(data, null, 2));

  out.balance = Number(data.balances[0].total_balance.value);

  // If there were no transactions in the reporting period, return early.
  if (results.length == 0) {
    console.log(out); //DEBUG_161
    return out;
  }

  // Collapse all transactions into a flat array, instead of an array of arrays.
  // Move one level deeper in the JSON hierarchy to the transaction_info obj.
  out.txns = results.flat().map((res) => res.transaction_info);

  // Sort txns in-place in ascending order, by update date.
  out.txns.sort((a,b) => {
    const ta = 
      Date.parse(a.transaction_updated_date ?? a.transaction_initiation_date);
    const tb = 
      Date.parse(b.transaction_updated_date ?? b.transaction_initiation_date);
    return ta - tb;
  });

  console.log(out); //DEBUG_161
  return out;
}


function paypal_makeReportOfx(startDate='2024-07-30T10:24:10-0000', endDate=Date.now()) {
  res = paypal_getTransactions_(startDate, endDate);
  
  ofx = ofx_makeHeader(
    res.reportDate,
    res.startDate,
    // add 1 second to end date, because OFX end dates are exclusive, but
    // paypal's are inclusive to 1-second precision.
    new Date(Date.parse(res.endDate) + 1*1000),
    "PayPal",
    res.accountId
  );

  for (const txn of res.txns) {
    const date = txn.transaction_updated_date ?? txn.transaction_initiation_date;
    const code = txn.transaction_event_code;
    const amountGross = Number(txn.transaction_amount.value);
    const amountFee = txn.fee_amount ? Number(txn.fee_amount.value) : 0;

    let memo = 'initiated ' + txn.transaction_initiation_date + '   ';
    if (txn.paypal_account_id) {
      memo += 'initiated by account ' + txn.paypal_account_id + '   ';
    }
    if (txn.paypal_reference_id) {
      memo += txn.paypal_reference_id_type + ' ref: ';
      memo += txn.paypal_reference_id + '   ';
    }
    if (txn.bank_reference_id) {
      memo += 'bank id: ' + txn.bank_reference_id + '   ';
    }

    ofx += ofx_makeTxn(
      paypal_ofxTxnCode(code, amountGross),
      date,
      amountGross,
      txn.transaction_id,
      code + ': ' + paypal_ofxTxnName(code, amountGross),
      memo
    );

    if (amountFee != 0) {
      ofx += ofx_makeTxn(
        "FEE",
        txn.transaction_updated_date,
        amountFee,
        txn.transaction_id + '-1',
        'payment processing fee',
        'fee for transaction ' + txn.transaction_id
      );
    }
  }

  ofx += ofx_makeFooter(res.balance, res.endDate);

  console.log(ofx);
  return ofx;
}


function paypal_ofxTxnCode(code, amount) {
  // 'T0400' -> group is '04'
  const group = code.substring(1,3);
  switch(group) {
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

function paypal_ofxTxnName(code, amount) {
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


function paypal_refreshToken_(force = false) {
  if (!force && accessToken_) {
    console.log('using access token in memory');
    return;
  }
  
  defaultHeaders_ = {
    'Accept': 'application/json',
    'Accept-Language': 'en_US',
  }

  const ps = PropertiesService.getScriptProperties();

  if (!secret_) {
    // Get client ID, secret, and any stored access token from property storage.
    // (this particular credential is set up to only allow read access to PayPal)
    const props = ps.getProperties();
    secret_ = props.paypal_client_id + ':' + props.paypal_client_secret;
    secret_ = Utilities.base64Encode(secret_);

    // If there's a stored access token and we've still got at least an hour
    // before it's supposed to expire, use it.
    if (!force && props.paypal_access_token && props.paypal_access_token_expires) {
      if (props.paypal_access_token_expires > (Date.now() - oneHour_ms)) {
        accessToken_ = props.paypal_access_token;
        accessTokenExpires_ = props.paypal_access_token_expires;
        defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;
        console.log(`using access token from properties: ${accessToken_}`); //DEBUG_161
        return;
      }
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
    paypal_access_token_expires: accessTokenExpires_,
  })

  defaultHeaders_['Authorization'] = 'Bearer ' + accessToken_;

  console.log(`Got new access token, expires ${new Date(accessTokenExpires_).toISOString()}`);
}


function paypal_urlFetch(url, options={}) {
  // Silence HTTP exceptions so we can handle bad access tokens, but
  // record whether or not the caller wanted them.
  const muteRequested = options.muteHttpExceptions ?? false;
  options.muteHttpExceptions = true;

  let resp = UrlFetchApp.fetch(url, options);
  
  // If our access token has expired, refresh token and try once more.
  if (resp.getResponseCode() == 401) {
    paypal_refreshToken_(true);
    resp = UrlFetchApp.fetch(url, options);
  }

  // If caller didn't want to mute HTTP exceptions and there is one, throw an error.
  const code = resp.getResponseCode();
  if (!muteRequested && code >= 400) {
    console.error(JSON.stringify(JSON.parse(resp.getContentText()), null, 2));
    throw new Error(`HTTP request to ${url} `
      + `failed with status code ${code}.`);
  }

  return resp;
}


function paypal_urlFetchAll(requests) {
  if (!requests || requests.length == 0) {
    return;
  }
  // Silence HTTP exceptions so we can handle bad access tokens, but
  // record whether or not the caller wanted them.
  for (const request of requests) {
    request.muteRequested = request.muteHttpExceptions ?? false;
    request.muteHttpExceptions = true;
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
    paypal_refreshToken_(true);
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
