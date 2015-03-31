var gulp = require('gulp');
var plug = require('gulp-load-plugins')();
var _ = require('lodash');

var isProduction = function(file) {
  return process.env.NODE_ENV === 'production';
};
var isDebug = function(file) {
  return isProduction(file);
};

gulp.task('test', function () {
  return gulp.src('spec/*.js')
  .pipe(plug.jasmine({
    verbose: true,
    includeStackTrace: true
  }));
});

gulp.task('watch', ['test'], function() {
  gulp.watch(['lib/**/*.*'], ['test']);
});

gulp.task('default', ['test']);