"use strict";
const { google } = require('googleapis');
var Promise = require('polyfill-promise');

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
    return cells.flat();
  } catch (e) {
    throw new Error(e);
  }
};

/**
 * Retrieve cells based on given range
 * NOTE: If there are missing cells (no content) this function
 * adds them there (unlike other functions), thus you'll always
 * have full matrix
 *
 * @param  {String} docId     Sheet document id
 * @param  {String} sheetId   Sheet id
 * @param  {Mixed}  rangeInfo Range info. use sheets A1 notation for this https://developers.google.com/sheets/api/guides/concepts
 * @return {Promise}          Rows containing cells, like [[{A1}, {B1}], [{A2}, {B2}]]
 */
Sheets.prototype.getRange = async function(docId, sheetId, rangeInfo) {
  try {
    const [{ title }, sheets] = await Promise.all([
      this.getSheet(docId, sheetId),
      this.get()
    ]);

    const request = {
      spreadsheetId: docId,
      range: rangeInfo === undefined ? title : `${title}!${rangeInfo}`,
    }
    const sheet = await sheets.spreadsheets.values.get(request);
    return sheet.data.values.map((row, index) => {
      const rowIndex = index + 1;
      return row.map((cell, cellIndex) => {
        return {
          row: rowIndex,
          column: this.columnLetters.split(',')[cellIndex],
          content: cell
        };
      });
    });
  } catch(error) {
    throw new Error(error);
  }
};

module.exports = Sheets;