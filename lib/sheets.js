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
 * @return {Promise}          A promise that resolves to sheet info containing id, title, colCount and rowCount
 */
Sheets.prototype.getSheet = async function getSheet(docId, sheetId) {
  try {
    const sheetsInfo = await this.getSheets(docId);
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
  try {
    const { title } = await this.getSheet(docId, sheetId);
    const sheets = await this.get();

    const request = {
      spreadsheetId: docId,
      range: title
    }
    const sheet = await sheets.spreadsheets.values.get(request);
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
 * Sanitize range info before passing it to v4 request
 * @param {String} t Title, name of the sheet
 * @param {String} ri Range info, e.g "A1:B2", where "A" and "B" are columns and "1" and "2" are rows
*/
Sheets.prototype.v4RangeSanitizer = (t, ri) => {
  // Titles with spaces need to be wrapped into single quotes
  let title = t.includes(" ") ? `'${t}'` : t;
  return `${title}!${ri}`;
};

/**
 * Get total count of rows and columns in a data array
 * @param {Array} data Title, name of the sheet
 * @return {Array<number>} Total count of rows and columns
*/
Sheets.prototype.getRowAndColCount = (data) => {
  const dataArray = data || [];
  let col = 0;
  
  for (let r = 0; r < dataArray.length; r += 1) {
    for (let c = 0; c < dataArray[r].length; c += 1) {
      if (c >= col) {
        col = col + c;
      }
    }
  }
  return [dataArray.length, col];
}

/**
 * Generate a empty matrix from startRow,startCol
 * @param {number} totalRow no of total rows in raw v4 data
 * @param {number} totalCol no of total cols in raw v4 data
 * @param {number} startRow start row number
 * @param {number} startCol start col number
 * @return {Array<[{row: number, column: string, content: string}]>} A full empty matrix
*/
Sheets.prototype.paddedEmptyMatrix = (totalRow = 0, totalCol = 0, startRow = 1, startCol = 1) => {
  return [...Array(totalRow)]
  .map((_ ,r) =>
    [...Array(totalCol)]
       .map((__, c) => ({ row: r+startRow, column: excelColumnName.intToExcelCol(c+startCol), content: '' }))
  )
}

/**
 * Creates a full, padded data matrix
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
 */
Sheets.prototype.paddedDataMatrix = function (data, rangePattern) {
  let noOfCols = 0;
  let startCol = 0;
  let noOfRows = 0;
  let startRow = 0;
  let nMatrix = [];
  const [row, col] = this.getRowAndColCount(data);
  const [_, pattern1, pattern2, pattern3, pattern4] = rangePattern || [];
  const specialPatternType1 = [pattern3, pattern4].every(value => value === ''); // 'A3:'

  if(!rangePattern){
    nMatrix = this.paddedEmptyMatrix(row, col);
  } else if (specialPatternType1) {
    startRow = +pattern2;
    noOfRows = Math.abs(data.length - startRow) + 1;
    startCol = excelColumnName.excelColToInt(pattern1);
    noOfCols = Math.abs(col - startCol) + 1;
    nMatrix = this.paddedEmptyMatrix(noOfRows, noOfCols, startRow, startCol);
  } else {
    noOfRows = Math.abs(+pattern4 - +pattern2) + 1;
    startRow = Math.min(+pattern2, +pattern4);
    noOfCols = Math.abs(
      excelColumnName.excelColToInt(pattern3) - excelColumnName.excelColToInt(pattern1)
    ) + 1;
    startCol = Math.min(excelColumnName.excelColToInt(pattern1), excelColumnName.excelColToInt(pattern3));
    nMatrix = this.paddedEmptyMatrix(noOfRows, noOfCols, startRow, startCol);
  }

  for (let r = 0; r < nMatrix.length; r += 1) {
    for (let c = 0; c < nMatrix[r].length; c += 1) {
      if (specialPatternType1) {
        const startRowIndex = startRow - 1;
        const startColIndex = startCol - 1;
        nMatrix[r][c].content = (data  === undefined || data[r+startRowIndex][c+startColIndex] === undefined) 
        ? '' 
        : data[r+startRowIndex][c+startColIndex];
      } else {
        nMatrix[r][c].content = (data  === undefined || data[r][c] === undefined) 
        ? '' 
        : data[r][c];
      }

    }
  }
  return nMatrix;
}

/**
 * Retrieve cells data based on given range
 * 
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
 * @param  {String} docId     Sheet document id
 * @param  {String} sheetId   Sheet id
 * @param  {Mixed}  rangeInfo Range info.
 * @return {Promise}          Rows containing cells
 */
Sheets.prototype.getRange = async function(docId, sheetId, rangeInfo) {
  try {
    const [{ title }, sheets] = await Promise.all([
      this.getSheet(docId, sheetId),
      this.get()
    ]);

    const request = {
      spreadsheetId: docId,
      range: this.v4RangeSanitizer(title, rangeInfo)
    }

    if (!rangeInfo) {
      const req = {
        ...request,
        range: title
      }
      const sheet = await sheets.spreadsheets.values.get(req);
      return this.paddedDataMatrix(sheet.data.values);
    } else {
      const rangePattern = rangeInfo.match(/([a-zA-Z]*)(\d*):([a-zA-Z]*)(\d*)/);
      const [_, pattern1, pattern2, pattern3, pattern4] = rangePattern;
      const patternType1 = [pattern3, pattern4].every(value => value === ''); // 'A3:'
      const patternType2 = [pattern1, pattern2, pattern3, pattern4].every(value => value !== ''); // 'A1:C10'

      if (patternType1) {
        const req = {
          ...request,
          range: title
        }
        const sheet = await sheets.spreadsheets.values.get(req);
        return this.paddedDataMatrix(sheet.data.values, rangePattern);
      } else if(patternType2) {
        const sheet = await sheets.spreadsheets.values.get(request);
        return this.paddedDataMatrix(sheet.data.values, rangePattern);
      } else { // 'A:A' or '1:2' or 'A3:A' or 'Sheet1'
        const sheet = await sheets.spreadsheets.values.get(request);
        console.log('data->', sheet.data.values, request)
        return sheet.data.values || [];
      }
    }

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