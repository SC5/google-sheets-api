var fs = require("fs");
var path = require("path");
var Promise = require("polyfill-promise");
var proxyquire = require("proxyquire");

describe("Sheets", function () {
  // Load test JSON
  var rawSpreadsheet = fs
    .readFileSync(path.join(__dirname, "spreadsheet.json"))
    .toString();
  var rawValues = fs
    .readFileSync(path.join(__dirname, "values.json"))
    .toString();
  var rawRangeOfValues1 = fs
    .readFileSync(path.join(__dirname, "range-of-values-1.json"))
    .toString();
  var rawRangeOfValues2 = fs
    .readFileSync(path.join(__dirname, "range-of-values-2.json"))
    .toString();
  var spreadsheet = Promise.resolve(JSON.parse(rawSpreadsheet));
  var values = Promise.resolve(JSON.parse(rawValues));
  var rangeOfValues1 = Promise.resolve(JSON.parse(rawRangeOfValues1));
  var rangeOfValues2 = Promise.resolve(JSON.parse(rawRangeOfValues2));

  // Mock googleapis
  var Sheets = proxyquire("../lib/sheets", {
    googleapis: {
      google: {
        auth: {
          JWT: function () {
            return {
              authorize: function () {
                return Promise.resolve({});
              },
              credentials: {
                expiry_date: 0,
              },
            };
          },
        },
        sheets: function () {
          return {
            spreadsheets: {
              get: function () {
                return Promise.resolve(spreadsheet);
              },
              values: {
                get: function ({range}) {
                  console.log('<range>', range)
                  // Partial range --> 'Sheet1!'
                  if(range.includes('!')) {
                    // 'Sheet1!A3:A'
                    const match = range.match(/([a-zA-Z]+)(\d+):([a-zA-Z]+)(\d*)/);
                    console.log('<match>', match)
                    if (match[1] === match[3]) {
                      return Promise.resolve(rangeOfValues2)
                    }
                    return Promise.resolve(rangeOfValues1);
                  }
                  // Full range --> 'Sheet1'
                  return Promise.resolve(values);
                },
              },
            },
          };
        },
      },
    },
  });

  var sheets = new Sheets({
    email: "test@company.com",
    key: "testkey",
  });

  it("gets cells", function (done) {
    sheets.getCells(null, 807593019).then(function (cells) {
      expect(cells.length).toBe(13);
      done();
    });
  });

  it("parses range", function () {
    // info = {
    //   from: { col: "A", row: 1 },
    //   to: { col: "B", row: 2 },
    // };
    var info = sheets.parseRangeInfo("A1:B2");

    expect(info.from.col).toBe("A");
    expect(info.from.row).toBe(1);
    expect(info.to.col).toBe("B");
    expect(info.to.row).toBe(2);

    // Test with double digits, too
    // info = {
    //   from: { col: "A", row: 1 },
    //   to: { col: "B", row: 10 },
    // };
    var info = sheets.parseRangeInfo("A1:B10");
    expect(info.from.row).toBe(1);
    expect(info.to.row).toBe(10);
  });

  it("parses partial range", function () {
    // info = {
    //   from: { col: "A", row: 1 },
    //   to: { col: null, row: null },
    // };
    var info = sheets.parseRangeInfo("A1:");
    expect(info.from.col).toBe("A");
    expect(info.from.row).toBe(1);
    expect(info.to.col).toBe(null);
    expect(info.to.row).toBe(null);
  });

  it("parses invalid range", function () {
    // info = {
    //   from: { col: null, row: null },
    //   to: { col: null, row: null },
    // };
    var info = sheets.parseRangeInfo("asdf");
    expect(info.from.col).toBe(null);
    expect(info.from.row).toBe(null);
    expect(info.to.col).toBe(null);
    expect(info.to.row).toBe(null);
  });

  it("returns rows with partial range", function (done) {
    sheets.getRange(null, 807593019, "A3:").then(function (rows) {
      // rows = [
      //   [ { row: 3, column: 'A', content: 'A3' }, { row: 3, column: 'B', content: '' }, { row: 3, column: 'C', content: '' } ],
      //   [ { row: 4, column: 'A', content: 'A4' }, { row: 4, column: 'B', content: '' }, { row: 4, column: 'C', content: '' } ],
      //   [ { row: 5, column: 'A', content: 'A5' }, { row: 5, column: 'B', content: '' }, { row: 5, column: 'C', content: '' } ],
      //   [ { row: 6, column: 'A', content: 'A5' }, { row: 6, column: 'B', content: '' }, { row: 6, column: 'C', content: '' } ],
      //   [ { row: 7, column: 'A', content: 'A7' }, { row: 7, column: 'B', content: '' }, { row: 7, column: 'C', content: '' } ],
      //   [ { row: 8, column: 'A', content: 'A8' }, { row: 8, column: 'B', content: '' }, { row: 8, column: 'C', content: '' } ],
      //   [ { row: 9, column: 'A', content: 'A9' }, { row: 9, column: 'B', content: '' }, { row: 9, column: 'C', content: '' } ],
      //   [ { row: 10, column: 'A', content: 'A10' }, { row: 10, column: 'B', content: '' }, { row: 10, column: 'C', content: '' } ],
      // ]

      // console.log('<ROWS>', rows)
      expect(rows.length).toBe(8);
      expect(rows[0].length).toBe(3);
      done();
    });
  });

  // DONE
  it("returns rows with full range", function (done) {
    sheets.getRange(null, 807593019, "B1:C2").then(function (rows) {
      // rows = [
      //   [
      //     { row: 1, column: "B", content: "B1" },
      //     { row: 1, column: "C", content: "C1" },
      //   ],
      //   [
      //     { row: 2, column: "B", content: "" },
      //     { row: 2, column: "C", content: "C2" },
      //   ],
      // ];
      expect(rows.length).toBe(2);
      expect(rows[0][0].content).toBe("B1");
      done();
    });
  });

  // DONE
  it("returns rows with no range", function (done) {
    sheets.getRange(null, 807593019, null).then(function (rows) {
      console.log('<Rows>', rows)
      // console.log('<<ROWS>>', rows)
      // rows = [
      //   [
      //     { row: 1, column: "A", content: "A1" },
      //     { row: 1, column: "B", content: "B1" },
      //     { row: 1, column: "C", content: "C1" },
      //   ],
      //   [
      //     { row: 2, column: "A", content: "A2" },
      //     { row: 2, column: "B", content: "" },
      //     { row: 2, column: "C", content: "C2" },
      //   ],
      //   [
      //     { row: 3, column: "A", content: "A3" },
      //     { row: 3, column: "B", content: "" },
      //     { row: 3, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 4, column: "A", content: "A4" },
      //     { row: 4, column: "B", content: "" },
      //     { row: 4, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 5, column: "A", content: "A5" },
      //     { row: 5, column: "B", content: "" },
      //     { row: 5, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 6, column: "A", content: "A5" },
      //     { row: 6, column: "B", content: "" },
      //     { row: 6, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 7, column: "A", content: "A7" },
      //     { row: 7, column: "B", content: "" },
      //     { row: 7, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 8, column: "A", content: "A8" },
      //     { row: 8, column: "B", content: "" },
      //     { row: 8, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 9, column: "A", content: "A9" },
      //     { row: 9, column: "B", content: "" },
      //     { row: 9, column: "C", content: "" },
      //   ],
      //   [
      //     { row: 10, column: "A", content: "A10" },
      //     { row: 10, column: "B", content: "" },
      //     { row: 10, column: "C", content: "" },
      //   ],
      // ];
      expect(rows.length).toBe(10);
      done();
    })
  });
});
