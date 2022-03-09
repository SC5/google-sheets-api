# google-sheets-api

[![Build Status](https://travis-ci.org/SC5/google-sheets-api.svg?branch=master)](https://travis-ci.org/SC5/google-sheets-api)

An unofficial client for *reading* data from Google Sheets, since [googleapis does not come with one](https://github.com/google/google-api-nodejs-client/tree/master/apis).

**Table of contents**

<!-- MarkdownTOC depth=3 autolink=true bracket=round -->

- [google-sheets-api](#google-sheets-api)
  - [Usage](#usage)
  - [API](#api)
    - [Sheets(options)](#sheetsoptions)
    - [sheets.getSheets(id, sheetId)](#sheetsgetsheetsid-sheetid)
    - [sheets.getSheet(id, sheetId)](#sheetsgetsheetid-sheetid)
    - [sheets.getRange(id, sheetId, rangeInfo)](#sheetsgetrangeid-sheetid-rangeinfo)
    - [sheets.getCells(id, sheetId)](#sheetsgetcellsid-sheetid)
  - [Changelog](#changelog)
  - [License](#license)
  - [Credit](#credit)

<!-- /MarkdownTOC -->

## Usage

1.  Install module

    ```shell
    npm install google-sheets-api
    ```

2.  Create a project in [Google Developer Console](https://console.developers.google.com/project), for example: "Sheets App"
3.  Enable Drive API for project under *APIs & auth* > *APIs*
4.  Create service auth credentials for project under *APIs & auth* > *Credentials* > *Create new Client ID*: *Service account*
5.  Collect the listed service email address
6.  Regenerate and download the P12 key
7.  Convert the .p12 file into .pem format:

    ```shell
    openssl pkcs12 -in *.p12 -nodes -nocerts > sheets.pem
    ```

    when prompted for password, it's `notasecret`

8.  Share the Sheets document to *service email address* using the *Share* button
9.  Pick up the Sheets document id from URL or Share dialog. Example:

    ```shell
    # Sheets document browser URL
    https://docs.google.com/a/sc5.io/spreadsheets/d/1FHa0vyPxXj3BtqigQ3LcwPoa7ldlRtUDx6fFV6CqkNE/edit#gid=0
    # Sheets document id
    1FHa0vyPxXj3BtqigQ3LcwPoa7ldlRtUDx6fFV6CqkNE
    ```

9.  Put it all together:

    ```javascript
    var fs = require('fs');
    var Promise = require('polyfill-promise');
    var Sheets = require('google-sheets-api').Sheets;

    // TODO: Replace these values with yours
    var documentId = 'generated-by-sheets';
    var serviceEmail = 'generated-by-dev-console@developer.gserviceaccount.com';
    var serviceKey = fs.readFileSync('path/to/your/sheets.pem').toString();

    var sheets = new Sheets({ email: serviceEmail, key: serviceKey });

    sheets.getSheets(documentId)
    .then(function(sheetsInfo) {
      // NOTE: Using first sheet in this example
      var sheetInfo = sheetsInfo[0];
      return Promise.all([
        sheets.getSheet(documentId, sheetInfo.id),
        sheets.getRange(documentId, sheetInfo.id, 'A1:C3')
      ]);
    })
    .then(function(sheets) {
      console.log('Sheets metadata:', sheets[0]);
      console.log('Sheets contents:', sheets[1]);
    })
    .catch(function(err){
      console.error(err, 'Failed to read Sheets document');
    });
    ```

10. Success!


## API

Relevant API methods, see code for details and internal ones.

**NOTE:** All the methods returns a native (polyfilled when needed) Promise.

### Sheets(options)

Initialize Sheets client with provided options

* @param {Object} options        All the options
* @param {String} options.email  Service email address
* @param {String} options.key    Service .PEM key contents

### sheets.getSheets(id, sheetId)

Fetch info from one sheet

* @param  {String} id      Sheets document id
* @param  {String} sheetId Worksheet id (use getSheets to fetch them)
* @return {Promise}        A promise that resolves to a list of worksheet info

### sheets.getSheet(id, sheetId)

Fetch info from one sheet

* @param  {String} id      Sheets document id
* @param  {String} sheetId Worksheet id (use getSheets to fetch them)
* @return {Promise}        A promise that resolves to a worksheet info containing id, title, rowCount, colCount


### sheets.getRange(id, sheetId, rangeInfo)

Retrieve cells data based on given range

**NOTE:** 
*  * All below ranges are v4 compatible but full matrix are [SUPPORTED] only for few of them:
*
* -  [SUPPORTED] "Sheet1!A1:B2" refers to the first two cells in the top two rows of Sheet1.
* -  [SUPPORTED] "A3:" refres to all cells starts from 'A' column and 3rd row.
* - "Sheet1!A:A" refers to all the cells in the first column of Sheet1.
* - "Sheet1!1:2" refers to all the cells in the first two rows of Sheet1.
* - "Sheet1!A5:A" refers to all the cells of the first column of Sheet 1, from row 5 onward.
* - [SUPPORTED]"Sheet1" refers to all the cells in Sheet1.
* - "'My Custom Sheet'!A:A" refers to all the cells in a sheet named "My Custom Sheet."
*   Single quotes are required for sheet names with spaces, special characters, or an alphanumeric combination.
*
*
* @param  {String} id        Sheet document id
* @param  {String} sheetId   Sheet id
* @param  {Mixed} rangeInfo  Range info
* @return {Array}            Rows containing cells


### sheets.getCells(id, sheetId)

Fetch cell contents from one worksheet

* @param  {String} id      Sheets document id
* @param  {String} sheetId Worksheet id (use getSheets to fetch them)
* @return {Promise}        A promise that resolves to a list of rows

### sheets.getRowAndColCount(data)

Get total count of rows and columns in a data array

* @param {Array} data Title, name of the sheet
* @return {Array<number>} Total count of rows and columns

### sheets.paddedEmptyMatrix(data)

Generate a empty matrix from startRow,startCol

* @param {number} totalRow no of total rows in raw v4 data
* @param {number} totalCol no of total cols in raw v4 data
* @param {number} startRow start row number
* @param {number} startCol start col number
* @return {Array<[{row: number, column: string, content: string}]>} A full empty matrix

### sheets.paddedDataMatrix(data, rangePattern)

Creates a full, padded data matrix

* if the range is like 'A3:' or 'B1:C2' and there are missing
* cells (no content) this function adds them there (unlike other functions),
* thus you'll always have full matrix like B1:C2 -->
* [
    [
      { row: 1, column: "B", content: "B1" },
      { row: 1, column: "C", content: "C1" },
    ],
    [
      { row: 2, column: "B", content: "" },
      { row: 2, column: "C", content: "C2" },
    ],
  ]
* or in either case it will return raw v4 response like, A:B --> [[A1, B1], ['', B2]]
* 
* @param  {Array}   data     Sheet document id
* @param  {String}  rangePattern Range info
* @return {Array<[]>}  A full data matrix

## Changelog

- 1.0.0: Google sheet API migration from v3 to v4
- 0.4.3: Fixed JWT auth issue with recent Google API
- 0.4.2: Updated dependencies / fixed vulnerabilities
- 0.4.1: Fixed the double letter range issue, like: `A1:AA5`
- 0.4.0: Added support for setting auth scope (makes module usable with other Google APIs as well)
- 0.3.0: Using native promises if available, added `rowCount` and `colCount` to `getSheet()` response
- 0.2.3: Improved documentation
- 0.2.2: Fixed the issue the range with double digits, like `A1:C10`
- 0.2.1: Fixed the documentation
- 0.2.0: Added support for getRange()
- 0.1.0: Initial release

## License

Module is MIT -licensed

## Credit

Module is backed by

<a href="http://sc5.io">
  <img src="http://logo.sc5.io/78x33.png" style="padding: 4px 0;">
</a>
