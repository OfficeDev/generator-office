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

describe('office:mail', function(){

  beforeEach(function(done){
    options = {
      name: 'My Office Add-in'
    };
    done();
  });

  /**
   * Test addin when running on empty folder.
   */
  describe('run on new project (empty folder)', function(){

    beforeEach(function(done){
      // set to current folder
      options.rootPath = '';
      done();
    });

    /**
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function(){
      beforeEach(function(done){
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

      afterEach(function(){
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function(done){
        assert.file('manifest.xml');
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function(){
        var manifest = {};

        beforeEach(function(done){
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function(err, manifestContent){
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
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function(done){
          var subject = manifest.OfficeApp.FormSettings[0].Form[0].DesktopSettings[0].SourceLocation[0].$.DefaultValue;
          expect(subject).to.equal('https://localhost:8443/manifest-only/index.html');
          done();
        });

        /**
        * Form for ItemRead present
        */
        it('includes form for ItemRead', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.FormSettings[0].Form, function(formSetting){
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
        it('includes form for ItemEdit', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.FormSettings[0].Form, function(formSetting){
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
        it('includes rule for mail read', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Rule[0].Rule, function(rule){
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
        it('includes rule for mail edit', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Rule[0].Rule, function(rule){
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
        it('includes rule for appointment read', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Rule[0].Rule, function(rule){
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
        it('includes rule for appointment edit', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Rule[0].Rule, function(rule){
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

    }); // describe('technology:manifest-only')

  }); // describe('run on new project (empty folder)')

});
