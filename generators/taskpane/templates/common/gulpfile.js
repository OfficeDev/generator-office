'use script';

var gulp = require('gulp');
var webserver = require('gulp-webserver');

gulp.task('serve-static', function(){
  gulp.src('.')
    .pipe(webserver({
      https: true,
      port: '8443',
      host: 'localhost',
      directoryListing: true,
      fallback: 'index.html'
    }));
});
