
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

describe('office:content', function () {

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
      tech: 'html',
      startPage: 'https://localhost:8443/manifest-only/index.html'
    };
    
    // run generator
    helpers.run(path.join(__dirname, '../../generators/content'))
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
   * Test addin when running on an exsting folder.
   */
  describe('run on existing project (non-empty folder)', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      done();
    });


    /**
     * Test addin when technology = html
     */
    describe('addin technology:html', function () {

      beforeEach(function (done) {
        //set language to html
        options.tech = 'html';

        // run the generator
        helpers.run(path.join(__dirname, '../../generators/content'))
          .withOptions(options)
          .on('ready', function (gen) {
            util.setupExistingProject(gen);
          }.bind(this))
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
          'gulpfile.js',
          'package.json',
          'manifest.xml',
          'tsd.json',
          'jsconfig.json',
          addinRootPath + '/app/app.js',
          addinRootPath + '/app/app.css',
          addinRootPath + '/app/home/home.js',
          addinRootPath + '/app/home/home.html',
          addinRootPath + '/app/home/home.css',
          addinRootPath + '/content/Office.css',
          addinRootPath + '/images/close.png',
          addinRootPath + '/scripts/MicrosoftAjax.js'
        ];
        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'ProjectName',
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
          name: 'ProjectName',
          description: 'HTTPS site using Express and Node.js',
          version: '0.1.0',
          main: 'src/server/server.js',
          dependencies: {
            express: '^4.12.2'
          },
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
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/app/home/home.html');
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
    
  }); // describe('run on existing project (non-empty folder)')
  
});