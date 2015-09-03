
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
   * Test addin when running on empty folder.
   */
  describe('run on new project (empty folder)', function () {

    beforeEach(function (done) {
      // set to current folder
      options.rootPath = '';
      done();
    });

    /**
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function () {
      beforeEach(function (done) {
        //set language to html
        options.tech = 'manifest-only';
        options.startPage = 'https://localhost:8443/manifest-only/index.html';

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
        assert.file('manifest.xml');
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
          expect(manifest.OfficeApp.FormSettings[0].Form[0].DesktopSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/manifest-only/index.html');
          done();
        });

      }); //describe('manifest.xml contents')

    }); // describe('technology:manifest-only')

  }); // describe('run on new project (empty folder)')

});