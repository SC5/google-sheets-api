"use strict";
var _ = require('lodash');
var google = require('googleapis');
var Promise = require('bluebird');
var request = require('request');


function authorize(options) {
  var authClient = new google.auth.JWT(
    options.email,
    null,
    options.key,
    ['https://spreadsheets.google.com/feeds/']
  );

  return new Promise(function(resolve, reject) {
    authClient.authorize(function(err, tokens) {
      if (err) {
        reject(err);
      }
      resolve(authClient);
    });
  });
}

/**
 * Initialize the Sheets API client
 * @param {Object} options        All the options
 * @param {String} options.email  Service email address
 * @param {String} options.key    Service .PEM key contents
 */
function Sheets(options) {
  this.baseUrl = "https://spreadsheets.google.com/feeds/";
  this.email = options.email;
  this.key = options.key;
  this.authorization_expiry = new Date();
  this.columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split('');
}

/**
 * Authorize the API requests
 * @return {Object} Authorization object
 */
Sheets.prototype.authorize = function() {
  var self = this;
  if (!this.authorization || (this.authorization_expiry < new Date())) {
    this.authorization = authorize({
      email: this.email,
      key: this.key
    });

    this.authorization.then(function(authClient) {
      self.authorization_expiry = new Date(authClient.credentials.expiry_date);
    });

  }
  return this.authorization;
};

/**
 * Internal method to communicate with Google Sheets API
 * @param  {String} path URL path
 * @return {Object}      Response
 */
Sheets.prototype.get = function get(path) {
  var self = this;
  return new Promise(function(resolve, reject) {
    self.authorize()
    .then(function(authClient) {
      request({
        url: self.baseUrl + path,
        auth: {
          bearer: authClient.credentials.access_token
        },
        json: true
      }, function(error, response, body) {
        if (error) {
          reject(new Error(error));
        }
        if (response.statusCode !== 200) {
          reject(new Error(response.body));
        }
        resolve(body);
      });
    })
    .catch(function(err) {
      reject(err);
    });
  });
};

/**
 * Fetch info from all the worksheets in document
 * @param  {String} id Sheets document id (see browser URL)
 * @return {Promise}   A promise that resolves to a list of worksheet info
 */
Sheets.prototype.getSheets = function getSheets(id) {
  var path = "worksheets/" + id + "/private/full?alt=json";
  return this.get(path).then(function(data) {
    return data.feed.entry.map(function(sheet) {
      var idArray = sheet.id["$t"].split("/");
      return {
        id: idArray[idArray.length - 1],
        updated: sheet.updated["$t"],
        title: sheet.title["$t"],
        colCount: parseInt(sheet['gs$colCount']["$t"]),
        rowCount: parseInt(sheet['gs$rowCount']["$t"])
      };
    });
  });
};

/**
 * Fetch info from one sheet
 * @param  {String} id      Sheets document id
 * @param  {String} sheetId Worksheet id (use getSheets to fetch them)
 * @return {Promise}        A promise that resolves to sheet info containing id, title and latest update info
 */
Sheets.prototype.getSheet = function getSheet(id, sheetId) {
  var path = "worksheets/" + id + "/private/full/" + sheetId + "?alt=json";
  return this.get(path).then(function(data) {
    var sheet = data.entry;
    var idArray = sheet.id["$t"].split("/");
    return {
      id: idArray[idArray.length - 1],
      updated: sheet.updated["$t"],
      title: sheet.title["$t"],
      colCount: parseInt(sheet['gs$colCount']["$t"]),
      rowCount: parseInt(sheet['gs$rowCount']["$t"])
    };
  });
};

/**
 * Return low level info about the worksheet
 * @param  {String} id      Sheets document id
 * @param  {String} sheetId Worksheet id (use getSheets to fetch them)
 * @return {Promise}        A promise that resolves to list of data
 */
Sheets.prototype.getList = function getList(id, sheetId) {
  // TODO: parse rows to sane data, now returns raw feed data
  var path = "list/" + id + "/" + sheetId + "/private/full?alt=json";
  return this.get(path).then(function(data) {
    return data.feed.entry;
  });
};

/**
 * Fetch cell contents from one worksheet
 * @param  {String} id      Sheets document id
 * @param  {String} sheetId Worksheet id (use getSheets to fetch them)
 * @return {Promise}        A promise that resolves to a list of rows
 */
Sheets.prototype.getCells = function getCells(id, sheetId) {
  var path = "cells/" + id + "/" + sheetId + "/private/full?alt=json";
  return this.get(path).then(function(data) {
    return data.feed.entry.map(function(cell) {
      var title = cell.title["$t"];
      return {
        row: title.substring(1),
        column: title[0],
        content: cell.content["$t"]
      };
    });
  });
};

/**
 * Retrieve cells based on given range
 * NOTE: If there are missing cells (no content) this function
 * adds them there (unlike other functions), thus you'll always
 * have full matrix
 *
 * @param  {String} id        Sheet document id
 * @param  {String} sheetId   Sheet id
 * @param  {Mixed} rangeInfo  Range info as object or string like 'A2:D5' or 'A2:'
 * @return {Array}            Rows containing cells, like [[{A1}, {B1}], [{A2}, {B2}]]
 */
Sheets.prototype.getRange = function(id, sheetId, rangeInfo) {
  var self = this;
  var rows = [];

  // Get empty default using parser function
  rangeInfo = rangeInfo || self.parseRangeInfo('');

  // Info range is given in string format, parse it
  if (_.isString(rangeInfo)) {
    rangeInfo = self.parseRangeInfo(rangeInfo);
  }

  // Retrieve cells (returns in one array)
  return self.getCells(id, sheetId)
    .then(function(cells) {
      var currentRow = parseInt(_.first(cells).row, 10);
      var currentColumn = _.first(cells).column;
      var currentColIndex = self.columnLetters.indexOf(currentColumn);
      var row = [];

      // Convert rows string presentation into integer
      _.each(cells, function(cell) {
        cell.row = parseInt(cell.row, 10);
      });

      // Get used columns from cells to pickup start and stop
      // NOTE: .sort() does not work with numbers
      var cellColumnLetters  = _.pluck(cells, 'column').sort();
      var cellRowNumbers  = _.chain(cells)
        .pluck('row')
        .sortBy()
        .value();

      rangeInfo.from.col = rangeInfo.from.col || _.first(cellColumnLetters);
      rangeInfo.from.row = rangeInfo.from.row || _.first(cellRowNumbers);
      rangeInfo.to.col = rangeInfo.to.col || _.last(cellColumnLetters);
      rangeInfo.to.row = rangeInfo.to.row || _.last(cellRowNumbers);

      // Pad missing cells: Go throw columns and rows and add missing
      var currentPadColumn;
      var cell;

      // Iterate through rows
      _.each(_.range(rangeInfo.from.row, rangeInfo.to.row + 1), function(currentPadRow) {
        row = [];

        // Iterate through columns
        _.each(_.range(self.columnLetters.indexOf(rangeInfo.from.col), self.columnLetters.indexOf(rangeInfo.to.col) + 1), function(columnIndex) {
          currentPadColumn = self.columnLetters[columnIndex];
          // Try to find the cell based on current row and column
          // if not found, place empty content there instead
          cell = _.findWhere(cells, { row: currentPadRow, column: currentPadColumn });
          if (cell) {
            row.push(cell);
          } else {
            row.push({ row: currentPadRow, column: currentPadColumn,  content: '' });
          }
        });

        rows.push(row);
      });

      return Promise.resolve(rows);
    });
};

/**
 * Parse given range from given string, like 'A2:C8' or 'A2:'
 * @param  {String} str Range info
 * @return {Object}     Range info { from: { col: 'A', row: 2 }, to: ...}
 */
Sheets.prototype.parseRangeInfo = function(str) {
  var rangeInfo = {
    from: {
      col: null, row: null
    },
    to: {
      col: null, row: null
    }
  };

  // Match full range: A1:B20
  var match = str.match(/(\w)(\d+):(\w)(\d+)/);

  if (match && match.length > 1) {
    rangeInfo.from.row = parseInt(match[2], 10);
    rangeInfo.from.col = match[1];
    rangeInfo.to.row = parseInt(match[4], 10);
    rangeInfo.to.col = match[3];

    return rangeInfo;
  }

  // Match full range: A1:
  var match = str.match(/(\w)(\d+):/);
  if (match && match.length > 1) {
    rangeInfo.from.row = parseInt(match[2], 10);
    rangeInfo.from.col = match[1];

    return rangeInfo;
  }

  return rangeInfo;
};

module.exports = Sheets;