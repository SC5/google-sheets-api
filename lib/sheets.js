"use strict";
const { google } = require('googleapis');
var Promise = require('polyfill-promise');
var excelColumnName = require('excel-column-name');

/**
 * Authorizationn via GoogleApis
 * @param {Object} options        All the options
 * @param {String} options.email  Service email address
 * @param {String} options.key    Service .PEM key contents
 * @param {String} options.baseUrl Scope of the Service
 * @return {Promise}   A promise that resolves to a JWT object
 */
async function authorize(options) {
  const authClient = new google.auth.JWT(
    options.email,
    null,
    options.key,
    [options.baseUrl]
  );

  try {
    await authClient.authorize();
    return authClient;
  } catch (err) {
    throw new Error(err);
  }
}

/**
 * Initialize the Sheets API client
 * @param {Object} options        All the options
 * @param {String} options.email  Service email address
 * @param {String} options.key    Service .PEM key contents
 */
function Sheets(options) {
  this.baseUrl = "https://www.googleapis.com/auth/spreadsheets.readonly";
  this.email = options.email;
  this.key = options.key;
  this.authorization_expiry = new Date();
  this.columnLetters = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM,BN,BO,BP,BQ,BR,BS,BT,BU,BV,BW,BX,BY,BZ,CA,CB,CC,CD,CE,CF,CG,CH,CI,CJ,CK,CL,CM,CN,CO,CP,CQ,CR,CS,CT,CU,CV,CW,CX,CY,CZ,DA,DB,DC,DD,DE,DF,DG,DH,DI,DJ,DK,DL,DM,DN,DO,DP,DQ,DR,DS,DT,DU,DV,DW,DX,DY,DZ,EA,EB,EC,ED,EE,EF,EG,EH,EI,EJ,EK,EL,EM,EN,EO,EP,EQ,ER,ES,ET,EU,EV,EW,EX,EY,EZ,FA,FB,FC,FD,FE,FF,FG,FH,FI,FJ,FK,FL,FM,FN,FO,FP,FQ,FR,FS,FT,FU,FV,FW,FX,FY,FZ,GA,GB,GC,GD,GE,GF,GG,GH,GI,GJ,GK,GL,GM,GN,GO,GP,GQ,GR,GS,GT,GU,GV,GW,GX,GY,GZ,HA,HB,HC,HD,HE,HF,HG,HH,HI,HJ,HK,HL,HM,HN,HO,HP,HQ,HR,HS,HT,HU,HV,HW,HX,HY,HZ,IA,IB,IC,ID,IE,IF,IG,IH,II,IJ,IK,IL,IM,IN,IO,IP,IQ,IR,IS,IT,IU,IV,IW,IX,IY,IZ,JA,JB,JC,JD,JE,JF,JG,JH,JI,JJ,JK,JL,JM,JN,JO,JP,JQ,JR,JS,JT,JU,JV,JW,JX,JY,JZ,KA,KB,KC,KD,KE,KF,KG,KH,KI,KJ,KK,KL,KM,KN,KO,KP,KQ,KR,KS,KT,KU,KV,KW,KX,KY,KZ,LA,LB,LC,LD,LE,LF,LG,LH,LI,LJ,LK,LL,LM,LN,LO,LP,LQ,LR,LS,LT,LU,LV,LW,LX,LY,LZ,MA,MB,MC,MD,ME,MF,MG,MH,MI,MJ,MK,ML,MM,MN,MO,MP,MQ,MR,MS,MT,MU,MV,MW,MX,MY,MZ,NA,NB,NC,ND,NE,NF,NG,NH,NI,NJ,NK,NL,NM,NN,NO,NP,NQ,NR,NS,NT,NU,NV,NW,NX,NY,NZ,OA,OB,OC,OD,OE,OF,OG,OH,OI,OJ,OK,OL,OM,ON,OO,OP,OQ,OR,OS,OT,OU,OV,OW,OX,OY,OZ,PA,PB,PC,PD,PE,PF,PG,PH,PI,PJ,PK,PL,PM,PN,PO,PP,PQ,PR,PS,PT,PU,PV,PW,PX,PY,PZ,QA,QB,QC,QD,QE,QF,QG,QH,QI,QJ,QK,QL,QM,QN,QO,QP,QQ,QR,QS,QT,QU,QV,QW,QX,QY,QZ,RA,RB,RC,RD,RE,RF,RG,RH,RI,RJ,RK,RL,RM,RN,RO,RP,RQ,RR,RS,RT,RU,RV,RW,RX,RY,RZ,SA,SB,SC,SD,SE,SF,SG,SH,SI,SJ,SK,SL,SM,SN,SO,SP,SQ,SR,SS,ST,SU,SV,SW,SX,SY,SZ,TA,TB,TC,TD,TE,TF,TG,TH,TI,TJ,TK,TL,TM,TN,TO,TP,TQ,TR,TS,TT,TU,TV,TW,TX,TY,TZ,UA,UB,UC,UD,UE,UF,UG,UH,UI,UJ,UK,UL,UM,UN,UO,UP,UQ,UR,US,UT,UU,UV,UW,UX,UY,UZ,VA,VB,VC,VD,VE,VF,VG,VH,VI,VJ,VK,VL,VM,VN,VO,VP,VQ,VR,VS,VT,VU,VV,VW,VX,VY,VZ,WA,WB,WC,WD,WE,WF,WG,WH,WI,WJ,WK,WL,WM,WN,WO,WP,WQ,WR,WS,WT,WU,WV,WW,WX,WY,WZ,XA,XB,XC,XD,XE,XF,XG,XH,XI,XJ,XK,XL,XM,XN,XO,XP,XQ,XR,XS,XT,XU,XV,XW,XX,XY,XZ,YA,YB,YC,YD,YE,YF,YG,YH,YI,YJ,YK,YL,YM,YN,YO,YP,YQ,YR,YS,YT,YU,YV,YW,YX,YY,YZ,ZA,ZB,ZC,ZD,ZE,ZF,ZG,ZH,ZI,ZJ,ZK,ZL,ZM,ZN,ZO,ZP,ZQ,ZR,ZS,ZT,ZU,ZV,ZW,ZX,ZY,ZZ";
  this.sheets = null;
}

/**
 * Authorize the API requests
 * @return {Object} Authorization object
 */
Sheets.prototype.authorize = async function() {
  var self = this;
  try {
    if (!this.authorization || (this.authorization_expiry < new Date())) {
      this.authorization = await authorize({
        email: this.email,
        key: this.key,
        baseUrl: this.baseUrl
      });
  
      self.authorization_expiry = new Date(this.authorization.credentials.expiry_date);
    }
    return this.authorization;
  } catch (e) {
    throw new Error(e);
  }

};

/**
 * Internal method to communicate with Google Sheets API
 * @param  {String} path URL path
 * @return {Object}      Response
 */
Sheets.prototype.get = async function get() {
  var self = this;
  try{
    const authClient = await self.authorize();
    const sheets = google.sheets({
      version: 'v4',
      auth: authClient
    });
    return sheets;
    // this.sheets = sheets;
  } catch(e){
    throw new Error(e);
  }
};

/**
 * Fetch info from all the sheets in a spreadsheet
 * @param  {String} docId   spreadsheet document id (see browser URL)
 * @return {Promise}        A promise that resolves to a list of sheets info
 */
Sheets.prototype.getSheets = async function getSheets(docId) {
  try {
    const sheets = await this.get();
    const request = {
      spreadsheetId: docId
    }
    const response = await sheets.spreadsheets.get(request);
    
    console.log('v4 response-->', response.data.sheets )
    return(response.data.sheets.map((sheet) => {
      return {
        id: sheet.properties.sheetId,
        title: sheet.properties.title,
        colCount: sheet.properties.gridProperties.columnCount,
        rowCount: sheet.properties.gridProperties.rowCount
      };
    }));
  } catch(error) {
    throw new Error(error);
  }
};

/**
 * Fetch info from one sheet
 * @param  {String} docId     spreadsheet document id
 * @param  {String} sheetId   sheet id (use getSheets to fetch them)
 * @return {Promise}          A promise that resolves to sheet info containing id, title and latest update info
 */
Sheets.prototype.getSheet = async function getSheet(docId, sheetId) {
  try {
    const sheetsInfo = await this.getSheets(docId);
    console.log('-----sheetsInfo-----', sheetsInfo)
    const { id, title, rowCount, colCount } = sheetsInfo
    .find((sheet) => sheet.id === +sheetId);

    return {
      id,
      title,
      colCount,
      rowCount
    };
  } catch(error) {
    throw new Error(error);
  }
};

/**
 * Return low level info about a sheet
 * @param  {String} id      spreadsheet document id
 * @param  {String} sheetId sheet id (use getSheets to fetch them)
 * @return {Promise}        A promise that resolves to list of data from a sheet
 */
Sheets.prototype.getList = async function getList(docId, sheetId) {
  // TODO: parse rows to sane data, now returns raw feed data
  try {
    const { title } = await this.getSheet(docId, sheetId);
    const sheets = await this.get();

    const request = {
      spreadsheetId: docId,
      range: title
    }
    const sheet = await sheets.spreadsheets.values.get(request);
    // console.log(JSON.stringify(sheet.data.values, null, 2));
    return sheet.data.values;
  } catch(error) {
    throw new Error(error);
  }
};

/**
 * Fetch cell contents from one sheet
 * @param  {String} docId       spreadsheet document id
 * @param  {String} sheetId     sheet id (use getSheets to fetch them)
 * @return {Promise}            A promise that resolves to a list of rows
 */
Sheets.prototype.getCells = async function getCells(docId, sheetId) {
  try {
    const cells = await this.getRange(docId, sheetId);
    return cells.flat().filter(cell => cell.content != '');
  } catch (e) {
    throw new Error(e);
  }
};


/**
 * Convert rangeInfo to valid Google Sheets API v4 range
 *
 * Examples:
 *
 * - "Sheet1!A1:B2" refers to the first two cells in the top two rows of Sheet1.
 * - "Sheet1!A:A" refers to all the cells in the first column of Sheet1.
 * - "Sheet1!1:2" refers to all the cells in the first two rows of Sheet1.
 * - "Sheet1!A5:A" refers to all the cells of the first column of Sheet 1, from row 5 onward.
 * - "Sheet1" refers to all the cells in Sheet1.
 * - "'My Custom Sheet'!A:A" refers to all the cells in a sheet named "My Custom Sheet."
 *   Single quotes are required for sheet names with spaces, special characters, or an alphanumeric combination.
 *
 * @param {String} t Title, name of the sheet
 * @param {String} ri Range info, e.g "A1:B2", where "A" and "B" are columns and "1" and "2" are rows
 */
Sheets.prototype.toV4Range = (t, ri) => {
  // Titles with spaces need to be wrapped into single quotes
  let title = t.includes(" ") ? `'${t}'` : t;

  // Return just the title "Sheet1" if there is no range info
  if ([undefined, null].includes(ri)) {
    return title;
  }

  // "Sheet1!A3:" -> "Sheet1" -> start from A3 and go column per column on the response returning all cells that are on row 3 or above
  if (ri[ri.length - 1] === ":") {
    return title;
  }
  // "Sheet1!1:2" -> all the cells in the first two rows of Sheet1
  // const patternMatch = ri.match(/([a-zA-Z]*)(\d*):([a-zA-Z]*)(\d*)/);
  // if (patternMatch && patternMatch[1] === '' && patternMatch[3] === '') {
  //   return title;
  // }

  return `${title}!${ri}`;
  // return title;
};


 Sheets.prototype.fullDataMatrix = function (data, range) {
  const parts = range//'D1:D5'//'C5:A1'//'A1:B2'//'A1:C5'//'B2:C2';
  const partsMatch = parts.match(/([a-zA-Z]*)(\d*):([a-zA-Z]*)(\d*)/);
  let noOfColumns = 0;
  let startCol = 0;
  let noOfRows = 0;
  let startRow = 0;

  // No column specified --> '1:2'
  if(partsMatch[1] === '' && partsMatch[3] === ''){
    let col = 0;
    
    for (let r = 0; r < data.length; r += 1) {
      for (let c = 0; c < data[r].length; c += 1) {
        if (c >= col) {
          col = col + c;
        }
      }
    }
    noOfRows = Math.abs(+partsMatch[4] - +partsMatch[2]) + 1;
    startRow = Math.min(+partsMatch[2], +partsMatch[4]);
    noOfColumns = col;
    startCol = 1;
  } 
  // No row specified --> 'A:B' OR 'A:A' OR 'B:B'
  else if (partsMatch[2] === '' && partsMatch[4] === '') {
    noOfRows = data.length;
    startRow = 1;
    noOfColumns = Math.abs(
      excelColumnName.excelColToInt(partsMatch[3]) - excelColumnName.excelColToInt(partsMatch[1])
    ) + 1;
    startCol = Math.min(excelColumnName.excelColToInt(partsMatch[1]), excelColumnName.excelColToInt(partsMatch[3]))
  }
  // Only 1st part of a range specified --> 'A3:'
  else if (partsMatch[3] === '' && partsMatch[4] === '') {
    let col = 0;
    
    for (let r = 0; r < data.length; r += 1) {
      for (let c = 0; c < data[r].length; c += 1) {
        if (c >= col) {
          col = col + c;
        }
      }
    }
    startRow = +partsMatch[2];
    noOfRows = Math.abs(data.length - startRow) + 1;
    startCol = excelColumnName.excelColToInt(partsMatch[1]);
    noOfColumns = Math.abs(col - startCol) + 1;

    console.log('results-->', startRow, startCol, noOfRows, noOfColumns, partsMatch[1], excelColumnName.excelColToInt(partsMatch[1]))
  } 
  // full range specified --> 'A1:C5'
  else {
    noOfRows = Math.abs(+partsMatch[4] - +partsMatch[2]) + 1;
    startRow = Math.min(+partsMatch[2], +partsMatch[4]);
    noOfColumns = Math.abs(
      excelColumnName.excelColToInt(partsMatch[3]) - excelColumnName.excelColToInt(partsMatch[1])
    ) + 1;
    startCol = Math.min(excelColumnName.excelColToInt(partsMatch[1]), excelColumnName.excelColToInt(partsMatch[3]))
  }

  const nMatrix = [...Array(noOfRows)]
    .map((_ ,r) =>
      [...Array(noOfColumns)]
        .map((__, c) => ({ row: r+startRow, col: c+startCol, content: '' }))
    )
  console.log('nMatrix', nMatrix)
  for (let r = 0; r < nMatrix.length; r += 1) {
    for (let c = 0; c < nMatrix[r].length; c += 1) {
      // console.log('data--->', r, c, data[r+startRow-1][c+startCol-1], nMatrix[r][c])
      // const startRowIndex = startRow - 1;
      // const startColIndex = startCol - 1;
      // console.log('results-->', startRow, startCol, data[r+startRowIndex][c+startColIndex])
      nMatrix[r][c].content = (data  === undefined || data[r][c] === undefined) ? '' : data[r][c];
    }
  }
  return nMatrix;
}

/**
 * Retrieve cells based on given range
 * NOTE: If there are missing cells (no content) this function
 * adds them there (unlike other functions), thus you'll always
 * have full matrix
 *
 * @param  {String} docId     Sheet document id
 * @param  {String} sheetId   Sheet id
 * @param  {Mixed}  rangeInfo Range info. use sheets A1 notation for this https://developers.google.com/sheets/api/guides/concepts
 * @return {Promise<{ row: Number, column: String, content: String }[]>}          Rows containing cells, like [[{A1}, {B1}], [{A2}, {B2}]]
 */
Sheets.prototype.getRange = async function(docId, sheetId, rangeInfo) {
  try {
    const [{ title }, sheets] = await Promise.all([
      this.getSheet(docId, sheetId),
      this.get()
    ]);
    const range = this.toV4Range(title, rangeInfo);
    const request = {
      spreadsheetId: docId,
      range
    }
    const sheet = await sheets.spreadsheets.values.get(request);
    console.log('Actual data-->', sheet.data.values)
    return this.fullDataMatrix(sheet.data.values, rangeInfo);
    // const patternMatch = ![undefined, null].includes(rangeInfo) 
    //             ? rangeInfo.match(/([a-zA-Z]*)(\d*):([a-zA-Z]*)(\d*)/)
    //             : null
    // const sheet = await sheets.spreadsheets.values.get(request);
    // console.log('Actual data-->', sheet.data.values)
    // const fullMatrix = this.fullDataMatrix(sheet.data.values);
    
    // if(patternMatch && patternMatch[1] === '' && patternMatch[3] === ''){ // range 'Sheet1!1:2'
    //   return fullMatrix.slice(patternMatch[2] - 1, patternMatch[4]);
    // } else if (patternMatch && patternMatch[3] === '' && patternMatch[4] === '') { // range 'Sheet1!A3:'
    //   return fullMatrix.slice(patternMatch[2] - 1, fullMatrix.length);
    // } 
    // else { // range "Sheet1"
    //   return fullMatrix;
    // }
    
  } catch(error) {
    throw new Error(error);
  }
};

/**
 * Parse given range from given string, like 'A2:C8' or 'A2:'
 * @param  {String} str Range info
 * @return {Object}     Range info { from: { col: 'A', row: 2 }, to: ...}
 */
 Sheets.prototype.parseRangeInfo = function (str) {
  var rangeInfo = {
    from: {
      col: null,
      row: null,
    },
    to: {
      col: null,
      row: null,
    },
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