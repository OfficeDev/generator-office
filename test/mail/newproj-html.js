
'use strict';

var fs = require('fs');
var path = require('path');
var mockery = require('mockery');
var assert = require('yeoman-generator').assert;
var helpers = require('yeoman-generator').test;

var Xml2Js = require('xml2js');
var validator = require('validator');
var chai = require('chai'),
  expect = chai.expect;

var util = require('./../_testUtils');


// sub:generator options 
var options = {};

/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

describe('office:mail', function () {

  before(function (done) {
    options = {
      name: 'My Office Add-in'
    };
    done();
  });
  
  /**
   * Test addin when running on empty folder.
   */
  describe('run on new project (empty folder)', function () {

    before(function (done) {
      // set to current folder
      options.rootPath = '';
      done();
    });

    /**
     * Test addin when technology = html
     */
    describe('addin technology:html', function () {

      before(function (done) {
        //set language to html
        options.tech = 'html';

        // run the generator
        helpers.run(path.join(__dirname, '../../generators/mail'))
          .withOptions(options)
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        var expected = [
          '.bowerrc',
          'bower.json',
          'package.json',
          'gulpfile.js',
          'manifest.xml',
          'tsd.json',
          'jsconfig.json',
          'appcompose/app.js',
          'appcompose/app.css',
          'appcompose/home/home.js',
          'appcompose/home/home.html',
          'appcompose/home/home.css',
          'appread/app.js',
          'appread/app.css',
          'appread/home/home.js',
          'appread/home/home.html',
          'appread/home/home.css',
          'content/Office.css',
          'images/close.png',
          'scripts/MicrosoftAjax.js'
        ];
        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1'
          }
        };

        assert.file('bower.json');
        util.assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          devDependencies: {
            gulp: '^3.9.0',
            'gulp-webserver': '^0.9.1'
          }
        };

        assert.file('package.json');
        util.assertJSONFileContains('package.json', expected);
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.FormSettings[0].Form[0].DesktopSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/appcompose/home/home.html');
          done();
        });

      }); //describe('manifest.xml contents')

      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function () {
      
        it('contains task \'serve-static\'', function (done) {
          
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });
      
      }); // describe('gulpfile.js contents')

    }); // describe('technology:html')
      
  }); // describe('run on new project (empty folder)')

});