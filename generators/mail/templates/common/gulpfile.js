'use script';

var gulp = require('gulp');
var webserver = require('gulp-webserver');
var fs = require('fs');
var minimist = require('minimist');
var xmllint = require('xmllint');
var chalk = require('chalk');
var $ = require('gulp-load-plugins')({ lazy: true });
var del = require('del');
var runSequence = require('run-sequence');

var config = {
  release: './dist'
};

gulp.task('help', $.taskListing.withFilters(function (task) {
  var mainTasks = ['default', 'help', 'serve-static', 'validate-xml', 'dist'];
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

/**
 * Validates the Office add-in manifest for submission to the store.
 */
gulp.task('validate-xml', function () {
  var options = minimist(process.argv.slice(2));
  var xsd = fs.readFileSync('./manifest.xsd');
  var xmlFilePath = options.xmlfile;
  var resultsAsJson = options.json || false;
  var xml = fs.readFileSync(xmlFilePath);

  if (!resultsAsJson) {
    console.log('\nValidating ' + chalk.blue(xmlFilePath.substring(xmlFilePath.lastIndexOf('/') + 1)) + ':');
  }
  
  // verify valid XML against the XSD schema
  var result = xmllint.validateXML({
    xml: xml,
    schema: xsd
  });

  // check the <HighResolutionIconUrl> property
  _validateHighResolutionIconUrl(xml, result);

  if (resultsAsJson) {
    console.log(JSON.stringify(result));
  }
  else {
    if (result.errors === null) {
      console.log(chalk.green('Valid'));
    }
    else {
      console.log(chalk.red('Invalid'));
      result.errors.forEach(function (e) {
        console.log(chalk.red(e));
      });
    }
  }
});

/**
 * Removes existing dist folder
 */
gulp.task('dist-remove', function () {
  return del(config.release);
});

/**
 * Copies files to the dist folder
 */
gulp.task('dist-copy-files', function() {
  return gulp.src([
    './app*/**/*',
    './bower_components/**/*',
    './content/**/*',
    './images/**/*',
    './scripts/**/*',
    './manifest-*.xml',
    './index.html',
    './package.json'
  ], { base: './' }).pipe(gulp.dest(config.release));
});

/**
 * Optimizes JavaScript and CSS files
 */
gulp.task('dist-minify', ['dist-minify-js', 'dist-minify-css'], function() {
});

/**
 * Minifies and uglifies JavaScript files
 */
gulp.task('dist-minify-js', function() {
  gulp.src([
    './app*/**/*.js',
    './scripts/**/*', '!./scripts/MicrosoftAjax.js'
  ], { base: './' })
    .pipe($.uglify())
    .pipe(gulp.dest(config.release));
});

/**
 * Minifies and uglifies CSS files
 */
gulp.task('dist-minify-css', function() {
  gulp.src([
    './app*/**/*.css',
    './content/**/*.css'
  ], { base: './' })
    .pipe($.minifyCss())
    .pipe(gulp.dest(config.release));
});

/**
 * Creates a release version of the project
 */
gulp.task('dist', function () {
  runSequence(
    ['dist-remove'],
    ['dist-copy-files'],
    ['dist-minify']
    );
});

/** +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ **/

/**
 * Ensures the <HighResolutionIconUrl> element is present and valid.
 * 
 * @param {object}  xml     - XML document to process.
 * @param {object}  result  - Result object from validating the XML.
 */
function _validateHighResolutionIconUrl(xml, result) {
  if (xml && result) {
    var xmlString = xml.toString();

    if (xmlString.indexOf('<HighResolutionIconUrl ') > -1 &&
      xmlString.indexOf('<HighResolutionIconUrl DefaultValue="https://') < 0) {
      if (result.errors === null) {
        result.errors = [];
      }

      result.errors.push('The value of the HighResolutionIconUrl attribute contains an unsupported URL.'
                       + ' You can only use https:// URLs.');
    }
  }
}
