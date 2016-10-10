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
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function(){
      beforeEach(function(done){
        // set language to html
        options.tech = 'manifest-only';

        // set products
        options.clients = ['Document', 'Workbook', 'Presentation', 'Project','Notebook'];

        options.startPage = 'https://localhost:8443/manifest-only/index.html';

        // run the generator
        helpers.run(path.join(__dirname, '../../generators/taskpane'))
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
        assert.file(manifestFileName);
        done();
      });

      /**
       * manifest-*.xml is good
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
          expect(validator.isUUID(manifest.OfficeApp.Id + '')).to.be.true;
          done();
        });

        it('has correct display name', function(done){
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal(projectDisplayName);
          done();
        });

        it('has correct start page', function(done){
          var subject = manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue;
          expect(subject).to.equal('https://localhost:8443/manifest-only/index.html');
          done();
        });

        it('has valid icon URL', function (done) {
          expect(manifest.OfficeApp.IconUrl[0].$.DefaultValue)
            .to.match(/^https:\/\/.+\.(png|jpe?g|gif|bmp)$/i);
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
         * OneNote present in host entry.
         */
        it('includes OneNote in Hosts', function(done){
          var found = false;
          _.forEach(manifest.OfficeApp.Hosts[0].Host, function(h){
            if (h.$.Name === 'Notebook') {
              found = true;
            }
          });
          expect(found, '<Host Name="Notebook"/> exist').to.be.true;

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

    }); // describe('technology:manifest-only')

  }); // describe('run on existing project (non-empty folder)')

});
