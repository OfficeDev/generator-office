/* jshint expr:true */
'use strict';

var fs = require('fs');
var path = require('path');
var _ = require('lodash');
var mockery = require('mockery');
var assert = require('yeoman-assert');
var helpers = require('yeoman-test');

var Xml2Js = require('xml2js');
var validator = require('validator');
var chai = require('chai'),
  expect = chai.expect;

var util = require('./../_testUtils');


// sub:generator options
var options = {};


/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

describe('office:taskpane', function(){

  var projectDisplayName = 'My Office Add-in';
  var projectEscapedName = 'my-office-add-in';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  beforeEach(function(done){
    options = {
      name: projectDisplayName
    };
    done();
  });

  describe('run on existing project (non-empty folder)', function(){
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function(done){
      // set to current folder
      options.rootPath = addinRootPath;
      done();
    });

    /**
     * Test addin when technology = ng
     */
    describe('technology:ng-adal', function(){

      beforeEach(function(done){
        // set language to html
        options.tech = 'ng-adal';

        // set products
        options.clients = ['Document', 'Workbook', 'Presentation', 'Project'];
        
        options.appId = '03ad2348-c459-4573-8f7d-0ca44d822e7c';

        helpers.run(path.join(__dirname, '../../generators/taskpane'))
          .withOptions(options)
          .on('ready', function(gen){
            util.setupExistingProject(gen);
          }.bind(this))
          .on('end', done);
      });

      afterEach(function(){
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function(done){
        var expected = [
          '.bowerrc',
          'bower.json',
          'gulpfile.js',
          'package.json',
          manifestFileName,
          'manifest.xsd',
          'tsd.json',
          'jsconfig.json',
          'tsconfig.json',
          addinRootPath + '/index.html',
          addinRootPath + '/app/app.module.js',
          addinRootPath + '/app/app.adalconfig.js',
          addinRootPath + '/app/app.config.js',
          addinRootPath + '/app/app.routes.js',
          addinRootPath + '/app/home/home.controller.js',
          addinRootPath + '/app/home/home.html',
          addinRootPath + '/app/services/data.service.js',
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
      it('bower.json contains correct values', function(done){
        var expected = {
          name: 'ProjectName',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1',
            angular: '~1.4.4',
            'angular-route': '~1.4.4',
            'angular-sanitize': '~1.4.4',
            'adal-angular': '~1.0.5',
            'office-ui-fabric': '*'
          }
        };

        assert.file('bower.json');
        util.assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function(done){
        var expected = {
          name: 'ProjectName',
          description: 'HTTPS site using Express and Node.js',
          version: '0.1.0',
          main: 'src/server/server.js',
          dependencies: {
            express: '^4.12.2'
          },
          devDependencies: {
            chalk: '^1.1.1',
            del: '^2.1.0',
            gulp: '^3.9.0',
            'gulp-load-plugins': '^1.0.0',
            'gulp-minify-css': '^1.2.2',
            'gulp-task-listing': '^1.0.1',
            'gulp-uglify': '^1.5.1',
            "browser-sync": "^2.11.0",
            minimist: '^1.2.0',
            'run-sequence': '^1.1.5',
            xmllint: 'git+https://github.com/kripken/xml.js.git'
          }
        };

        assert.file('package.json');
        util.assertJSONFileContains('package.json', expected);
        done();
      });

      /**
       * manfiest-*.xml is good
       */
      describe('manifest-*.xml contents', function(){
        var manifest = {};

        beforeEach(function(done){
          var parser = new Xml2Js.Parser();
          fs.readFile(manifestFileName, 'utf8', function(err, manifestContent){
            parser.parseString(manifestContent, function(err, manifestJson){
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function(done){
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;
          done();
        });

        it('has correct display name', function(done){
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal(projectDisplayName);
          done();
        });

        it('has correct start page', function(done){
          var subject = manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue;
          expect(subject).to.equal('https://localhost:8443/index.html');
          done();
        });
        
        it('includes AAD App Domains', function(done){
          var loginWindowsNetFound = false;
          var loginMicrosoftonlineNetFound = false;
          var loginMicrosoftonlineComFound = false;
          
          _.forEach(manifest.OfficeApp.AppDomains[0].AppDomain, function(a){
            if (a === 'https://login.windows.net') {
              loginWindowsNetFound = true;
            }
            else if (a === 'https://login.microsoftonline.net') {
              loginMicrosoftonlineNetFound = true;
            }
            else if (a === 'https://login.microsoftonline.com') {
              loginMicrosoftonlineComFound = true;
            }
          });
          expect(loginWindowsNetFound, 'App Domain https://login.windows.net exist').to.be.true;
          expect(loginMicrosoftonlineNetFound, 'App Domain https://login.microsoftonline.net exist').to.be.true;
          expect(loginMicrosoftonlineComFound, 'App Domain https://login.microsoftonline.com exist').to.be.true;

          done();
        });

        /**
         * Word present in host entry.
         */
        it('includes Word in Hosts', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Hosts[0].Host, function(h){
            if (h.$.Name === 'Document') {
              found = true;
            }
          });
          expect(found, '<Host Name="Document"/> exist').to.be.true;

          done();
        });

        /**
         * Excel present in host entry.
         */
        it('includes Excel in Hosts', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Hosts[0].Host, function(h){
            if (h.$.Name === 'Workbook') {
              found = true;
            }
          });
          expect(found, '<Host Name="Workbook"/> exist').to.be.true;

          done();
        });

        /**
         * PowerPoint present in host entry.
         */
        it('includes PowerPoint in Hosts', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Hosts[0].Host, function(h){
            if (h.$.Name === 'Presentation') {
              found = true;
            }
          });
          expect(found, '<Host Name="Presentation"/> exist').to.be.true;

          done();
        });

        /**
         * Project present in host entry.
         */
        it('includes Project in Hosts', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Hosts[0].Host, function(h){
            if (h.$.Name === 'Project') {
              found = true;
            }
          });
          expect(found, '<Host Name="Project"/> exist').to.be.true;

          done();
        });

      }); // describe('manifest-*.xml contents')
      
      /**
       * app.config.js is good
       */
      describe('app.config.js contents', function(){
        it('contains correct appId', function(done){
          assert.file(addinRootPath + '/app/app.config.js');
          assert.fileContent(addinRootPath + '/app/app.config.js', '03ad2348-c459-4573-8f7d-0ca44d822e7c');
          done();
        });

      }); // describe('app.config.js contents')

      /**
       * tsd.json is good
       */
      describe('tsd.json contents', function(){
        var tsd = {};

        beforeEach(function(done){
          fs.readFile('tsd.json', 'utf8', function(err, tsdJson){
            tsd = JSON.parse(tsdJson);

            done();
          });
        });

        it('has correct *.d.ts references', function(done){
          expect(tsd.installed).to.exist;
          // make sure the existing ones are present (to verify we didn't overwrite, but rather update)
          expect(tsd.installed['lodash/lodash.d.ts']).to.exist;
          // make sure the new ones are present
          expect(tsd.installed['angularjs/angular.d.ts']).to.exist;
          expect(tsd.installed['angularjs/angular-route.d.ts']).to.exist;
          expect(tsd.installed['angularjs/angular-sanitize.d.ts']).to.exist;
          expect(tsd.installed['office-js/office-js.d.ts']).to.exist;
          done();
        });

      }); // describe('tsd.json contents')

      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function(){
        
        it('contains task \'help\'', function(done){
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'help\',');
          done();
        });
        
        it('contains task \'default\'', function(done){
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'default\',');
          done();
        });

        it('contains task \'serve-static\'', function(done){

          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });
        
        it('contains task \'validate\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'validate\',');
          done();
        });
        
        it('contains task \'validate-forcatalog\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'validate-forcatalog\',');
          done();
        });
        
        it('contains task \'validate-forstore\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'validate-forstore\',');
          done();
        });
        
        it('contains task \'validate-highResolutionIconUrl\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'validate-highResolutionIconUrl\',');
          done();
        });
        
        it('contains task \'validate-xml\'', function(done){
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'validate-xml\',');
          done();
        });
        
        it('contains task \'dist-remove\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist-remove\',');
          done();
        });
        
        it('contains task \'dist-copy-files\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist-copy-files\',');
          done();
        });
        
        it('contains task \'dist-minify\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist-minify\',');
          done();
        });
        
        it('contains task \'dist-minify-js\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist-minify-js\',');
          done();
        });
        
        it('contains task \'dist-minify-css\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist-minify-css\',');
          done();
        });
        
        it('contains task \'dist\'', function (done) {
          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'dist\',');
          done();
        });

      }); // describe('gulpfile.js contents')

    }); // describe('technology:ng')

  }); // describe('run on existing project (non-empty folder)')

});
