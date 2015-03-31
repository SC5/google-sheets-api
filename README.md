# google-sheets-api

[![Build Status](https://travis-ci.org/SC5/google-sheets-api.svg?branch=master)](https://travis-ci.org/SC5/google-sheets-api)

An unofficial client for *reading* data from Google Sheets, since [googleapis does not come with one](https://github.com/google/google-api-nodejs-client/tree/master/apis).

## Usage

1.  Install module

    ```shell
    npm install google-sheets-api
    ```

2.  Create a project in https://console.developers.google.com/project - for example: "Sheets App"
3.  Enable Drive API for project under *APIs & auth* > *APIs*
4.  Create service auth credentials for project under
    *APIs & auth* > *Credentials* > *Create new Client ID*: *Service account*


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
    var Promise = require('bluebird');
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
    .spread(function(sheet, rows) {
      // TODO: Do something with cell within range
      console.log('Sheets contents:', sheet, rows);
    })
    .catch(function(err){
      console.error(err, 'Failed to read Sheets document');
    });
    ```

10. Success!


## API

Relevant API methods, see code for details and internal ones.

**NOTE:** All the methods returns a Promise.

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
* @return {Promise}        A promise that resolves to sheet info containing id, title and latest update info


### sheets.getRange(id, sheetId, rangeInfo)

Retrieve cells based on given range

**NOTE:** If there are missing cells (no content) this function adds them there (unlike other functions), thus you'll always have full matrix

* @param  {String} id        Sheet document id
* @param  {String} sheetId   Sheet id
* @param  {Mixed} rangeInfo  Range info as object or string like `A2:D5` or `A2:`
* @return {Array}            Rows containing cells, like `[[{A1}, {B1}], [{A2}, {B2}]]`


### sheets.getCells(id, sheetId)

Fetch cell contents from one worksheet

* @param  {String} id      Sheets document id
* @param  {String} sheetId Worksheet id (use getSheets to fetch them)
* @return {Promise}        A promise that resolves to a list of rows

## License

Module is MIT -licensed
