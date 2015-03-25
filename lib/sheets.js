"use strict";
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
        title: sheet.title["$t"]
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
      title: sheet.title["$t"]
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

module.exports = Sheets;