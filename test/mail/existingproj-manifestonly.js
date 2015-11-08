/* jshint expr:true */
'use strict';

var fs = require('fs');
var path = require('path');
var _ = require('lodash');
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

  var projectDisplayName = 'My Office Add-in';
  var projectEscapedName = 'my-office-add-in';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  beforeEach(function (done) {
    options = {
      name: projectDisplayName
    };
    done();
  });

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
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function () {
      describe('Outlook form:mail-read, mail-compose, appointment-read, appointment-compose', function () {
        beforeEach(function (done) {
          // set language to html
          options.tech = 'manifest-only';
  
          // set outlook form type
          options.outlookForm = ['mail-read', 'mail-compose', 'appointment-read', 'appointment-compose'];

          options.startPage = 'https://localhost:8443/manifest-only/index.html';
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .on('end', done);
        });

        afterEach(function () {
          mockery.disable();
        });
  
        /**
        * All expected files are created.
        */
        it('creates expected files', function (done) {
          assert.file(manifestFileName);
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
            var subject = manifest.OfficeApp.FormSettings[0].Form[0]
                                  .DesktopSettings[0].SourceLocation[0].$.DefaultValue;
            expect(subject).to.equal('https://localhost:8443/manifest-only/index.html');
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

        }); // describe('manifest.xml contents')

      }); // describe('Outlook form:mail-read, mail-compose, appointment-read, appointment-compose')
      
      describe('Outlook form:mail-read, appointment-read', function () {
        beforeEach(function (done) {
          // set language to html
          options.tech = 'manifest-only';
  
          // set outlook form type
          options.outlookForm = ['mail-read', 'appointment-read'];

          options.startPage = 'https://localhost:8443/manifest-only/index.html';
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .on('end', done);
        });

        afterEach(function () {
          mockery.disable();
        });
  
        /**
        * All expected files are created.
        */
        it('creates expected files', function (done) {
          assert.file(manifestFileName);
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

        }); // describe('manifest.xml contents')

      }); // describe('Outlook form:mail-read, appointment-read')
      
      describe('Outlook form:mail-compose, appointment-compose', function () {
        beforeEach(function (done) {
          // set language to html
          options.tech = 'manifest-only';
  
          // set outlook form type
          options.outlookForm = ['mail-compose', 'appointment-compose'];

          options.startPage = 'https://localhost:8443/manifest-only/index.html';
  
          // run the generator
          helpers.run(path.join(__dirname, '../../generators/mail'))
            .withOptions(options)
            .on('end', done);
        });

        afterEach(function () {
          mockery.disable();
        });
  
        /**
        * All expected files are created.
        */
        it('creates expected files', function (done) {
          assert.file(manifestFileName);
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
          it('doesn\'t include rule for appointment read', function (done) {
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

      }); // describe('Outlook form:mail-compose, appointment-compose')

    }); // describe('technology:manifest-only')

  }); // describe('run on existing project (non-empty folder)')

});
