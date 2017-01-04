'use script';

var gulp = require('gulp');
var webserver = require('gulp-webserver');
var $ = require('gulp-load-plugins')({ lazy: true });

var config = {
  release: './dist',
  localUrl:'https://localhost:8443/'
};

gulp.task('help', $.taskListing.withFilters(function (task) {
  var mainTasks = ['default', 'help', 'serve-static'];
  var isSubTask = mainTasks.indexOf(task) < 0;
  return isSubTask;
}));
gulp.task('default', ['help']);

/** +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ **/

/**
 * Startup static webserver.
 */
gulp.task('serve-static', function () {
  gulp.src('.')
    .pipe(webserver({
      https: true,
      port: '8443',
      host: 'localhost',
      directoryListing: true,
      fallback: 'index.html'
    }));
});