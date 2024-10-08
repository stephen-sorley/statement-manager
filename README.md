[![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg)](https://github.com/google/clasp)

# Statement Manager

Provides a schedulable service that pulls transaction data from [PayPal](https://developer.paypal.com/api/rest/)
and [Stripe](https://docs.stripe.com/api), converts it to
[OFX v1.0.2](https://www.financialdataexchange.org/common/Uploaded%20files/OFX%20files/ofx1.0.2spec.zip)
format, and emails an OFX file to a list of recipients. These OFX files can then be easily imported into
[QuickBooks Online (QBO)](https://qbo.intuit.com/app/newfileupload).

This code runs on Google's scripting platform ([Google Apps Script](https://script.google.com)).

### Rationale

This is helpful because QuickBooks does not provide an automated way to import bank transactions from
Stripe and PayPal, and Stripe and PayPal do not provide transaction data downloads in OFX format (only
CSV).

Here are the reasons why this solution is better than importing CSV files into QuickBooks:

  1. QuickBooks cannot import bank balance data from CSV files, so it will not update the "Bank Balance"
     field for the account, as displayed [here](https://qbo.intuit.com/app/banking). This removes a lot of
     the usefulness of having bank transactions in QBO. Manually imported OFX files **do** update the
     "Bank Balance" field.

  2. QuickBooks requires that the user pick a date format and verify column selection every time a CSV
     is imported. This provides opportunities for the user to mess things up, and must be repeated on
     every import. OFX files import immediately without asking for any user input.

  3. The OFX files produced by this tool include unique id's for each transaction as provided by PayPal
     and Stripe. QuickBooks will automatically detect and exclude duplicate transactions using these.
     CSV imports do not support automatic duplicate detection or unique ID's, so the user is required
     to manually verify that the imported transactions are not duplicates, leading to errors.

  4. Both PayPal and Stripe put the gross payment amount and the fee amount on the same line of the CSV
     file. If you need to keep track of fees as expenditures, you can't just import the net amount - you
     need to have the fee payment listed on a separate line of the CSV file so that it's imported as a
     separate transaction. This requires manually editing the CSV files or writing your own tool to do so.
     Statement Manager can optionally be set to produce separate transactions for you (use 'gross' mode).

  6. Stripe does not offer any way to download all balance-affecting transactions in a single CSV file.
     You have to download payments in one file, and payouts (transfers to your linked bank account) in a
     separate file. This leads to either doing two separate imports into QuickBooks, or writing your own
     tool to combine the two files. Statement Manager reports all the required information in a single file.

### Setup Instructions

*UNDER CONSTRUCTION*
