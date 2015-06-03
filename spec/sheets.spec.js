var fs = require('fs');
var path = require('path');
var Promise = require('polyfill-promise');
var proxyquire = require('proxyquire');


describe('Sheets', function() {

  // Load test JSON
  var raw = fs.readFileSync(path.join(__dirname, 'cells.json')).toString();
  var cells = Promise.resolve(JSON.parse(raw));

  // Mock googleapis
  var Sheets = proxyquire('../lib/sheets', {
    googleapis: {
      auth: {
        JWT: function() {
          return {
            authorize: function(cb) {
              return cb();
            },
            credentials: {
              expiry_date: 0
            }
          }
        }
      }
    },
    request: function(params, cb) {
      // Replace with actual respose
      cb(null, { statusCode: 200 }, {});
    }
  });

  Sheets.prototype.getCells = function() {
    return Promise.resolve(cells);
  };

  var sheets = new Sheets({
    email: 'test@company.com',
    key: 'testkey'
  });


  it('gets cells', function(done) {
    sheets.getCells()
    .then(function(cells){
      expect(cells.length).toBe(13);
      done();
    });
  });

  it('parses range', function() {
    var info = sheets.parseRangeInfo('A1:B2');
    expect(info.from.col).toBe('A');
    expect(info.from.row).toBe(1);
    expect(info.to.col).toBe('B');
    expect(info.to.row).toBe(2);

    // Test with double digits, too
    var info = sheets.parseRangeInfo('A1:B10');
    expect(info.from.row).toBe(1);
    expect(info.to.row).toBe(10);
  });

  it('parses partial range', function() {
    var info = sheets.parseRangeInfo('A1:');
    expect(info.from.col).toBe('A');
    expect(info.from.row).toBe(1);
    expect(info.to.col).toBe(null);
    expect(info.to.row).toBe(null);
  });

  it('parses invalid range', function() {
    var info = sheets.parseRangeInfo('asdf');
    expect(info.from.col).toBe(null);
    expect(info.from.row).toBe(null);
    expect(info.to.col).toBe(null);
    expect(info.to.row).toBe(null);
  });

  it('returns rows with partial range', function(done) {
    sheets.getRange(null, null, 'A3:')
    .then(function(rows){
      expect(rows.length).toBe(8);
      expect(rows[0].length).toBe(3);
      done();
    });
  });

  it('returns rows with full range', function(done) {
    sheets.getRange(null, null, 'B1:C2')
    .then(function(rows){
      expect(rows.length).toBe(2);
      expect(rows[0][0].content).toBe('B1');
      done();
    });
  });

  it('returns rows with no range', function(done) {
    sheets.getRange(null, null, null)
    .then(function(rows){
      expect(rows.length).toBe(10);
      done();
    });
  });

});