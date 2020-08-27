"use strict";
var _ = require('lodash');
var google = require('googleapis');
var Promise = require('polyfill-promise');
var request = require('request');


function authorize(options) {
  var authClient = new google.Auth.JWT(
    options.email,
    null,
    options.key,
    [options.baseUrl]
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
  this.columnLetters = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ,CK,CL,CM,CN,CO,CP,CQ,CR,CS,CT,CU,CV,CW,CX,CY,CZ,DA,DB,DC,DD,DE,DF,DG,DH,DI,DJ,DK,DL,DM,DN,DO,DP,DQ,DR,DS,DT,DU,DV,DW,DX,DY,DZ,EA,EB,EC,ED,EE,EF,EG,EH,EI,EJ,EK,EL,EM,EN,EO,EP,EQ,ER,ES,ET,EU,EV,EW,EX,EY,EZ,FA,FB,FC,FD,FE,FF,FG,FH,FI,FJ,FK,FL,FM,FN,FO,FP,FQ,FR,FS,FT,FU,FV,FW,FX,FY,FZ,GA,GB,GC,GD,GE,GF,GG,GH,GI,GJ,GK,GL,GM,GN,GO,GP,GQ,GR,GS,GT,GU,GV,GW,GX,GY,GZ,HA,HB,HC,HD,HE,HF,HG,HH,HI,HJ,HK,HL,HM,HN,HO,HP,HQ,HR,HS,HT,HU,HV,HW,HX,HY,HZ,IA,IB,IC,ID,IE,IF,IG,IH,II,IJ,IK,IL,IM,IN,IO,IP,IQ,IR,IS,IT,IU,IV,IW,IX,IY,IZ,JA,JB,JC,JD,JE,JF,JG,JH,JI,JJ,JK,JL,JM,JN,JO,JP,JQ,JR,JS,JT,JU,JV,JW,JX,JY,JZ,KA,KB,KC,KD,KE,KF,KG,KH,KI,KJ,KK,KL,KM,KN,KO,KP,KQ,KR,KS,KT,KU,KV,KW,KX,KY,KZ,LA,LB,LC,LD,LE,LF,LG,LH,LI,LJ,LK,LL,LM,LN,LO,LP,LQ,LR,LS,LT,LU,LV,LW,LX,LY,LZ,MA,MB,MC,MD,ME,MF,MG,MH,MI,MJ,MK,ML,MM,MN,MO,MP,MQ,MR,MS,MT,MU,MV,MW,MX,MY,MZ,NA,NB,NC,ND,NE,NF,NG,NH,NI,NJ,NK,NL,NM,NN,NO,NP,NQ,NR,NS,NT,NU,NV,NW,NX,NY,NZ,OA,OB,OC,OD,OE,OF,OG,OH,OI,OJ,OK,OL,OM,ON,OO,OP,OQ,OR,OS,OT,OU,OV,OW,OX,OY,OZ,PA,PB,PC,PD,PE,PF,PG,PH,PI,PJ,PK,PL,PM,PN,PO,PP,PQ,PR,PS,PT,PU,PV,PW,PX,PY,PZ,QA,QB,QC,QD,QE,QF,QG,QH,QI,QJ,QK,QL,QM,QN,QO,QP,QQ,QR,QS,QT,QU,QV,QW,QX,QY,QZ,RA,RB,RC,RD,RE,RF,RG,RH,RI,RJ,RK,RL,RM,RN,RO,RP,RQ,RR,RS,RT,RU,RV,RW,RX,RY,RZ,SA,SB,SC,SD,SE,SF,SG,SH,SI,SJ,SK,SL,SM,SN,SO,SP,SQ,SR,SS,ST,SU,SV,SW,SX,SY,SZ,TA,TB,TC,TD,TE,TF,TG,TH,TI,TJ,TK,TL,TM,TN,TO,TP,TQ,TR,TS,TT,TU,TV,TW,TX,TY,TZ,UA,UB,UC,UD,UE,UF,UG,UH,UI,UJ,UK,UL,UM,UN,UO,UP,UQ,UR,US,UT,UU,UV,UW,UX,UY,UZ,VA,VB,VC,VD,VE,VF,VG,VH,VI,VJ,VK,VL,VM,VN,VO,VP,VQ,VR,VS,VT,VU,VV,VW,VX,VY,VZ,WA,WB,WC,WD,WE,WF,WG,WH,WI,WJ,WK,WL,WM,WN,WO,WP,WQ,WR,WS,WT,WU,WV,WW,WX,WY,WZ,XA,XB,XC,XD,XE,XF,XG,XH,XI,XJ,XK,XL,XM,XN,XO,XP,XQ,XR,XS,XT,XU,XV,XW,XX,XY,XZ,YA,YB,YC,YD,YE,YF,YG,YH,YI,YJ,YK,YL,YM,YN,YO,YP,YQ,YR,YS,YT,YU,YV,YW,YX,YY,YZ,ZA,ZB,ZC,ZD,ZE,ZF,ZG,ZH,ZI,ZJ,ZK,ZL,ZM,ZN,ZO,ZP,ZQ,ZR,ZS,ZT,ZU,ZV,ZW,ZX,ZY,ZZ".split(',');
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
      key: this.key,
      baseUrl: this.baseUrl
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
        colCount: parseInt(sheet['gs$colCount']["$t"], 10),
        rowCount: parseInt(sheet['gs$rowCount']["$t"], 10)
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
      colCount: parseInt(sheet['gs$colCount']["$t"], 10),
      rowCount: parseInt(sheet['gs$rowCount']["$t"], 10)
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
        row: parseInt(title.match(/\d+$/)[0], 10),
        column: title.replace(/\d+$/, ''),
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
      var currentColumn = _.first(cells).column;
      var row = [];

      // Convert rows string presentation into integer
      _.each(cells, function(cell) {
        cell.row = parseInt(cell.row, 10);
      });

      // Get used columns from cells to pickup start and stop
      // NOTE: .sort() does not work with numbers
      var cellColumnLetters  = _.map(cells, 'column').sort();
      var cellRowNumbers  = _.chain(cells)
        .map('row')
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
          cell = _.find(cells, { row: currentPadRow, column: currentPadColumn });
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
  var match = str.match(/([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d+)/);

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
