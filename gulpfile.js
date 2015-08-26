'use strict';

var gulp = require('gulp');
var spawn = require('child_process').spawn;
var inspector = require('gulp-node-inspector');
var chalk = require('chalk');
var which = require('which');
var path = require('path');
var istanbul = require('gulp-istanbul');
var mocha = require('gulp-mocha');

/**
 * test
 */
gulp.task('test', function (done) {
  gulp.src(['generators/**/.js'])
    .pipe(istanbul()) // covering files
    .pipe(istanbul.hookRequire()) // force 'require' to return coverd files
    .on('finish', function () {
      gulp.src(['test/*.js'])
        .pipe(mocha()) // run tests
        .pipe(istanbul.writeReports()) // write coverage reports
        .on('end', done);
    });
});

/**
 * Setup node inspector to debug app.
 */
gulp.task('node-inspector', function () {
  // start node inspector
  return gulp.src([])
    .pipe(inspector({
      debugPort: 5858,
      webHost: '127.0.0.1',
      webPort: 8080
    }));
});

/**
 * Run the Yeoman generator in debug mode.
 */
gulp.task('run-yo', function () {
  console.log(chalk.yellow('BE AWARE!!! - Running this with default options will scaffold the project in the generator\'s source folder.'));

  spawn('node',
    [
      '--debug',
      path.join(which.sync('yo'), '../../', 'lib/node_modules/yo/lib/cli.js'),
      'office',
      ' --skip-install'
    ],
    { stdio: 'inherit' });
});

gulp.task('debug-yo', ['run-yo', 'node-inspector']);

/* default gulp task */
gulp.task('default', ['node-inspector']);