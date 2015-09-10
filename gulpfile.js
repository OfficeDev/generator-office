'use strict';

var args = require('yargs').argv;
var spawn = require('child_process').spawn;
var chalk = require('chalk');
var which = require('which');
var path = require('path');

var config = require('./gulp.config.js');
var gulp = require('gulp');
var $ = require('gulp-load-plugins')({lazy: true});

var cwd = process.cwd();

/**
 * yargs variables can be passed in to alter the behavior, when present.
 * Example: gulp vet
 *      or: gulp vet --verbose
 *
 * --verbose  : Various tasks will produce more output to the console.
 */

/**
 * List the available gulp tasks
 */
gulp.task('help', $.taskListing);
gulp.task('default', ['help']);

/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

/**
 * Vets all JS for code style.
 */
gulp.task('vet', function(){
  log('Analyzing source with JSHint and JSCS');

  return gulp
    .src(config.allJs, {base: './'})
    .pipe($.if(args.verbose, $.print()))
    .pipe($.jshint())
    .pipe($.jshint.reporter('jshint-stylish', {verbose: true}))
    .pipe($.jscs());
});

/**
 * Run all tests with code coverage.
 */
gulp.task('test', function(done){
  gulp.src([path.join(cwd, 'generators/**/.js')])
    .pipe($.if(args.verbose, $.print()))
    .pipe($.istanbul()) // covering files
    .pipe($.istanbul.hookRequire()) // force 'require' to return coverd files
    .on('finish', function(){
      gulp.src([path.join(cwd, 'test/**/*.js')])
        .pipe($.mocha()) // run tests
        .on('error', handleError)
        .pipe($.istanbul.writeReports()) // write coverage reports
        .on('end', done);
    });
});

/**
 * Run the Yeoman generator in debug mode.
 */
gulp.task('run-yo', function(){
  log(chalk.yellow('BE AWARE!!! - Running this with default options will scaffold the project ' +
    'in the generator\'s source folder.'));

  spawn('node',
    [
      '--debug',
      path.join(which.sync('yo'), '../../', 'lib/node_modules/yo/lib/cli.js'),
      'office',
      ' --skip-install'
    ], {
      stdio: 'inherit'
    });
});

/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */
/*                         auto / watch tasks                                */
/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

/**
 * Watch for changes in any JS files to changes, then vet's & tests them.
 */
gulp.task('autotest', function(done){
  gulp.watch(['generators/**', 'test/**'], ['vet', 'test']);
});

/**
 * Watches for changes in scripts and auto-vets the JS with JSHint & JSCS.
 */
gulp.task('autovet', function(done){
  gulp.watch(['./.jscsrc', './.jshintrc', '**/*.js'], ['vet']);
});


/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */
/*                            utility methods                                */
/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

/**
 * Log a message or series of messages using chalk's blue color.
 * Can pass in a string, object or array.
 */
function log(msg){
  if (typeof (msg) === 'object') {
    for (var item in msg) {
      if (msg.hasOwnProperty(item)) {
        $.util.log($.util.colors.blue(msg[item]));
      }
    }
  } else {
    $.util.log($.util.colors.blue(msg));
  }
}

/**
 * Handle errors by writing to the log, then emit an end event.
 */
function handleError(err){
  $.log(err.toString());
  this.emit('end'); // jshint ignore:line
}

module.exports = gulp;
