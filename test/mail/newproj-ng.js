
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

  beforeEach(function (done) {
    options = {
      name: 'My Office Add-in'
    };
    done();
  });
   
  /**
   * Test scrubbing of name with illegal characters
   */
  it('project name is alphanumeric only', function (done) {
    options = {
      name: 'Some\'s bad * character$ ~!@#$%^&*()',
      rootPath: '',
      tech: 'ng',
      startPage: 'https://localhost:8443/manifest-only/index.html'
    };
    
    // run generator
    helpers.run(path.join(__dirname, '../../generators/mail'))
      .withOptions(options)
      .on('end', function () {
        var expected = {
          name: 'somes-bad-character',
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
  });
  
  /**
   * Test addin when running on empty folder.
   */
  describe('run on new project (empty folder)', function () {

    beforeEach(function (done) {
      // set to current folder
      options.rootPath = '';
      done();
    });

    /**
     * Test addin when technology = angular
     */
    describe('addin technology:ng', function () {

      beforeEach(function (done) {
        //set language to html
        options.tech = 'ng';

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
          'appcompose/index.html',
          'appcompose/app.module.js',
          'appcompose/app.routes.js',
          'appcompose/home/home.controller.js',
          'appcompose/home/home.html',
          'appcompose/services/data.service.js',
          'appread/index.html',
          'appread/app.module.js',
          'appread/app.routes.js',
          'appread/home/home.controller.js',
          'appread/home/home.html',
          'appread/services/data.service.js',
          'content/Office.css',
          'content/fabric.components.css',
          'content/fabric.components.min.css',
          'content/fabric.components.rtl.css',
          'content/fabric.components.rtl.min.css',
          'content/fabric.css',
          'content/fabric.min.css',
          'content/fabric.rtl.css',
          'content/fabric.rtl.min.css',
          'images/close.png',
          'scripts/MicrosoftAjax.js',
          'scripts/jquery.fabric.js',
          'scripts/jquery.fabric.min.js'
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
            jquery: '~1.9.1',
            angular: '~1.4.4',
            'angular-route': '~1.4.4',
            'angular-sanitize': '~1.4.4'
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

        beforeEach(function (done) {
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
          expect(manifest.OfficeApp.FormSettings[0].Form[0].DesktopSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/appcompose/index.html');
          done();
        });

      }); // describe('manifest.xml contents')
      
      /**
       * tsd.json is good
       */
      describe('tsd.json contents', function () {
        var tsd = {};

        beforeEach(function (done) {
          fs.readFile('tsd.json', 'utf8', function (err, tsdJson) {
            tsd = JSON.parse(tsdJson);

            done();
          });
        });

        it ('has correct *.d.ts references', function (done) {
          expect(tsd.installed).to.exist;
          expect(tsd.installed["jquery/jquery.d.ts"]).to.exist;
          expect(tsd.installed["angularjs/angular.d.ts"]).to.exist;
          expect(tsd.installed["angularjs/angular-route.d.ts"]).to.exist;
          expect(tsd.installed["angularjs/angular-sanitize.d.ts"]).to.exist;
          done();
        });

      }); // describe('tsd.json contents')

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

    }); // describe('technology:ng')
      
  }); // describe('run on new project (empty folder)')

});