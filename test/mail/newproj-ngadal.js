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
var prompts = {};

/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

describe('office:mail', function () {

  var projectDisplayName = 'My Office Add-in';
  var projectEscapedName = 'my-office-add-in';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  beforeEach(function (done) {
    options = {
      name: projectDisplayName
    };
    
    // Since mail invokes commands, we
    // need to mock responding to the prompts for
    // info
    prompts = {
      buttonTypes: ['uiless'],
      functionFileUrl: 'https://localhost:8443/manifest-only/functions.html',
      iconUrl: 'https://localhost:8443/manifest-only/icon.png'
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
      tech: 'ng-adal',
      extensionPoint: [
        'MessageReadCommandSurface', 
        'MessageComposeCommandSurface', 
        'AppointmentAttendeeCommandSurface', 
        'AppointmentOrganizerCommandSurface'
      ],
      startPage: 'https://localhost:8443/manifest-only/index.html'
    };

    // run generator
    helpers.run(path.join(__dirname, '../../generators/mail'))
      .withOptions(options)
      .withPrompts(prompts)
      .on('end', function () {
        var expected = {
          name: 'somes-bad-character',
          version: '0.1.0',
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
            'xml2js': '^0.4.15',
            xmllint: 'git+https://github.com/kripken/xml.js.git'
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
    describe('addin technology:ng-adal', function () {

      describe('Outlook extension points:MessageReadCommandSurface, MessageComposeCommandSurface, '
             + 'AppointmentAttendeeCommandSurface, AppointmentOrganizerCommandSurface', 
      function () {

        beforeEach(function (done) {
          // set language to html
          options.tech = 'ng-adal';

          options.appId = '03ad2348-c459-4573-8f7d-0ca44d822e7c';
  
          // set outlook form type
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'MessageComposeCommandSurface', 
            'AppointmentAttendeeCommandSurface', 
            'AppointmentOrganizerCommandSurface'
          ];
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .withPrompts(prompts)
            .on('end', done);
        });

        afterEach(function () {
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
            manifestFileName,
            'manifest.xsd',
            'tsd.json',
            'jsconfig.json',
            'tsconfig.json',
            'appcompose/index.html',
            'appcompose/app.module.js',
            'appcompose/app.adalconfig.js',
            'appcompose/app.config.js',
            'appcompose/app.routes.js',
            'appcompose/home/home.controller.js',
            'appcompose/home/home.html',
            'appcompose/services/data.service.js',
            'appread/index.html',
            'appread/app.module.js',
            'appread/app.adalconfig.js',
            'appread/app.config.js',
            'appread/app.routes.js',
            'appread/home/home.controller.js',
            'appread/home/home.html',
            'appread/services/data.service.js',
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
            name: projectEscapedName,
            version: '0.1.0',
            dependencies: {
              'microsoft.office.js': '*',
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
        it('package.json contains correct values', function (done) {
          var expected = {
            name: projectEscapedName,
            version: '0.1.0',
            scripts: {
              postinstall: 'bower install'
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
              'xml2js': '^0.4.15',
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
        describe('manifest-*.xml contents', function () {
          var manifest = {};

          beforeEach(function (done) {
            var parser = new Xml2Js.Parser();
            fs.readFile(manifestFileName, 'utf8', function (err, manifestContent) {
              parser.parseString(manifestContent, function (err, manifestJson) {
                manifest = manifestJson;

                done();
              });
            });
          });

          it('has valid ID', function (done) {
            expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;
            done();
          });

          it('has correct display name', function (done) {
            expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal(projectDisplayName);
            done();
          });

          it('has correct start page', function (done) {
            var valid = false;
            var subject = manifest.OfficeApp.FormSettings[0].Form[0]
                                  .DesktopSettings[0].SourceLocation[0].$.DefaultValue;

            if (subject === 'https://localhost:8443/appcompose/index.html' ||
              subject === 'https://localhost:8443/appread/index.html') {
              valid = true;
            }

            expect(valid, 'start page is not valid compose or edit form').to.be.true;
            done();
          });

          it('includes AAD App Domains', function (done) {
            var loginWindowsNetFound = false;
            var loginMicrosoftonlineNetFound = false;
            var loginMicrosoftonlineComFound = false;

            _.forEach(manifest.OfficeApp.AppDomains[0].AppDomain, function (a) {
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
          * Form for ItemRead present
          */
          it('includes form for ItemRead', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemRead') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemRead"> exist').to.be.true;
            done();
          });
  
          /**
          * Form for ItemEdit present
          */
          it('includes form for ItemEdit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemEdit') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemEdit"> exist').to.be.true;
            done();
          });
  
          /**
          * Rule for Mail Read present
          */
          it('includes rule for mail read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Read') {
                found = true;
              }

            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Mail Edit present
          */
          it('includes rule for mail edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Appointment Read present
          */
          it('includes rule for appointment read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Read') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Appointment Edit present
          */
          it('includes rule for appointment edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />').to.be.true;
            done();
          });

        }); // describe('manifest-*.xml contents')
        
        /**
        * app.config.js is good
        */
        describe('app.config.js contents', function () {
          it('appcompose/app.config.js contains correct appId', function (done) {
            assert.file('appcompose/app.config.js');
            assert.fileContent('appcompose/app.config.js', '03ad2348-c459-4573-8f7d-0ca44d822e7c');
            done();
          });

          it('appread/app.config.js contains correct appId', function (done) {
            assert.file('appread/app.config.js');
            assert.fileContent('appread/app.config.js', '03ad2348-c459-4573-8f7d-0ca44d822e7c');
            done();
          });
        }); // describe('app.config.js contents')
  
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

          it('has correct *.d.ts references', function (done) {
            expect(tsd.installed).to.exist;
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
        describe('gulpfule.js contents', function () {

          it('contains task \'help\'', function (done) {
            assert.file('gulpfile.js');
            assert.fileContent('gulpfile.js', 'gulp.task(\'help\',');
            done();
          });

          it('contains task \'default\'', function (done) {
            assert.file('gulpfile.js');
            assert.fileContent('gulpfile.js', 'gulp.task(\'default\',');
            done();
          });

          it('contains task \'serve-static\'', function (done) {

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

          it('contains task \'validate-xml\'', function (done) {
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
        
      }); // describe('Outlook extension points:MessageReadCommandSurface, 
          // MessageComposeCommandSurface, AppointmentAttendeeCommandSurface, 
          // AppointmentOrganizerCommandSurface')
      
      describe('Outlook extension points:MessageReadCommandSurface, '
             + 'AppointmentAttendeeCommandSurface', 
      function () {

        beforeEach(function (done) {
          // set language to html
          options.tech = 'ng-adal';

          options.appId = '03ad2348-c459-4573-8f7d-0ca44d822e7c';
  
          // set outlook form type
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'AppointmentAttendeeCommandSurface'
          ];
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .withPrompts(prompts)
            .on('end', done);
        });

        afterEach(function () {
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
            manifestFileName,
            'manifest.xsd',
            'tsd.json',
            'jsconfig.json',
            'tsconfig.json',
            'appread/index.html',
            'appread/app.module.js',
            'appread/app.adalconfig.js',
            'appread/app.config.js',
            'appread/app.routes.js',
            'appread/home/home.controller.js',
            'appread/home/home.html',
            'appread/services/data.service.js',
            'content/Office.css',
            'images/close.png',
            'scripts/MicrosoftAjax.js'
          ];
          assert.file(expected);
          done();
        });
  
        /**
        * manfiest-*.xml is good
        */
        describe('manifest-*.xml contents', function () {
          var manifest = {};

          beforeEach(function (done) {
            var parser = new Xml2Js.Parser();
            fs.readFile(manifestFileName, 'utf8', function (err, manifestContent) {
              parser.parseString(manifestContent, function (err, manifestJson) {
                manifest = manifestJson;

                done();
              });
            });
          });
  
          /**
          * Form for ItemRead present
          */
          it('includes form for ItemRead', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemRead') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemRead"> exist').to.be.true;
            done();
          });
  
          /**
          * Form for ItemEdit not present
          */
          it('doesn\'t include form for ItemEdit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemEdit') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemEdit"> exist').to.be.false;
            done();
          });
  
          /**
          * Rule for Mail Read present
          */
          it('includes rule for mail read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Read') {
                found = true;
              }

            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Mail Edit not present
          */
          it('doesn\'t include rule for mail edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />').to.be.false;
            done();
          });
  
          /**
          * Rule for Appointment Read present
          */
          it('includes rule for appointment read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Read') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Appointment Edit not present
          */
          it('doesn\'t include rule for appointment edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />').to.be.false;
            done();
          });

        }); // describe('manifest-*.xml contents')
          
      }); // describe('Outlook extension points:MessageReadCommandSurface, 
          // AppointmentAttendeeCommandSurface')
      
      describe('Outlook extension points:MessageComposeCommandSurface, '
             + 'AppointmentOrganizerCommandSurface', 
      function () {

        beforeEach(function (done) {
          // set language to html
          options.tech = 'ng-adal';

          options.appId = '03ad2348-c459-4573-8f7d-0ca44d822e7c';
  
          // set outlook form type
          options.extensionPoint = [
            'MessageComposeCommandSurface', 
            'AppointmentOrganizerCommandSurface'
          ];
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .withPrompts(prompts)
            .on('end', done);
        });

        afterEach(function () {
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
            manifestFileName,
            'manifest.xsd',
            'tsd.json',
            'jsconfig.json',
            'tsconfig.json',
            'appcompose/index.html',
            'appcompose/app.module.js',
            'appcompose/app.adalconfig.js',
            'appcompose/app.config.js',
            'appcompose/app.routes.js',
            'appcompose/home/home.controller.js',
            'appcompose/home/home.html',
            'appcompose/services/data.service.js',
            'content/Office.css',
            'images/close.png',
            'scripts/MicrosoftAjax.js'
          ];
          assert.file(expected);
          done();
        });
  
        /**
        * manfiest-*.xml is good
        */
        describe('manifest-*.xml contents', function () {
          var manifest = {};

          beforeEach(function (done) {
            var parser = new Xml2Js.Parser();
            fs.readFile(manifestFileName, 'utf8', function (err, manifestContent) {
              parser.parseString(manifestContent, function (err, manifestJson) {
                manifest = manifestJson;

                done();
              });
            });
          });
  
          /**
          * Form for ItemRead not present
          */
          it('doesn\'t include form for ItemRead', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemRead') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemRead"> exist').to.be.false;
            done();
          });
  
          /**
          * Form for ItemEdit present
          */
          it('includes form for ItemEdit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.FormSettings[0].Form, function (formSetting) {
              if (formSetting.$['xsi:type'] === 'ItemEdit') {
                found = true;
              }
            });

            expect(found, '<Form xsi:type="ItemEdit"> exist').to.be.true;
            done();
          });
  
          /**
          * Rule for Mail Read not present
          */
          it('doesn\'t include rule for mail read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Read') {
                found = true;
              }

            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />').to.be.false;
            done();
          });
  
          /**
          * Rule for Mail Edit present
          */
          it('includes rule for mail edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Message' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />').to.be.true;
            done();
          });
  
          /**
          * Rule for Appointment Read not present
          */
          it('doesn\'t includes rule for appointment read', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Read') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />').to.be.false;
            done();
          });
  
          /**
          * Rule for Appointment Edit present
          */
          it('includes rule for appointment edit', function (done) {
            var found = false;
            _.forEach(manifest.OfficeApp.Rule[0].Rule, function (rule) {
              if (rule.$['xsi:type'] === 'ItemIs' &&
                rule.$.ItemType === 'Appointment' &&
                rule.$.FormType === 'Edit') {
                found = true;
              }
            });

            expect(found, '<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />').to.be.true;
            done();
          });

        }); // describe('manifest-*.xml contents')
          
      }); // describe('Outlook extension points:MessageComposeCommandSurface, 
          // AppointmentOrganizerCommandSurface')

    }); // describe('technology:ng')

  }); // describe('run on new project (empty folder)')

});
