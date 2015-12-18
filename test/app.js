'use strict';

var fs = require('fs');
var path = require('path');
var assert = require('yeoman-assert');
var helpers = require('yeoman-test');

var Xml2Js = require('xml2js');
var validator = require('validator');
var chai = require('chai'),
  expect = chai.expect;

// generator prompt responses
var promptResponses = {};



/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

describe('office:app', function(){

  // set timeouts to 5s
  this.timeout(10000);

  var projectEscapedName = 'my-first-addin';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  describe('runs mail subgenerator', function(){

    beforeEach(function(done){
      // setup generator prompts
      var options = {
        name: 'My First Addin',
        rootPath: '',
        type: 'mail',
        tech: 'html',
        extensionPoint: [
          'MessageReadCommandSurface', 
          'MessageComposeCommandSurface', 
          'AppointmentAttendeeCommandSurface', 
          'AppointmentOrganizerCommandSurface'
        ],
        'skip-install': true
      };

      // run the generator
      helpers.run(path.join(__dirname, '../generators/app'))
        .withPrompts(options)
        .on('end', done);
    });

    it('manifest-*.xml is for mail addin', function(done){

      // verify manifest-*.xml exists
      assert.file(manifestFileName);

      // load manifest-*.xml as JSON
      var parser = new Xml2Js.Parser();
      fs.readFile(manifestFileName, 'utf8', function(err, manifestContent){
        parser.parseString(manifestContent, function(err, manifestJson){

          // check addin type
          expect(manifestJson.OfficeApp.$['xsi:type']).to.equal('MailApp');

          done();
        });
      });

    });

  }); // describe('runs mail subgenerator')

  describe('runs taskpane subgenerator', function(){

    beforeEach(function(done){
      // setup generator prompts
      var options = {
        name: 'My First Addin',
        rootPath: '',
        type: 'taskpane',
        tech: 'ng-adal',
        clients: ['Document', 'Workbook'],
        'skip-install': true
      };

      // run the generator
      helpers.run(path.join(__dirname, '../generators/app'))
        .withPrompts(options)
        .on('end', done);
    });

    it('manifest-*.xml is for taskpane addin', function(done){

      // verify manifest-*.xml exists
      assert.file(manifestFileName);

      // load manifest-*.xml as JSON
      var parser = new Xml2Js.Parser();
      fs.readFile(manifestFileName, 'utf8', function(err, manifestContent){
        parser.parseString(manifestContent, function(err, manifestJson){

          // check addin type
          expect(manifestJson.OfficeApp.$['xsi:type']).to.equal('TaskPaneApp');

          done();
        });
      });

    });

  }); // describe('runs content subgenerator')

  describe('runs content subgenerator', function(){

    beforeEach(function(done){
      // setup generator prompts
      var options = {
        name: 'My First Addin',
        rootPath: '',
        type: 'content',
        tech: 'html',
        clients: ['Document', 'Workbook'],
        'skip-install': true
      };

      // run the generator
      helpers.run(path.join(__dirname, '../generators/app'))
        .withPrompts(options)
        .on('end', done);
    });

    it('manifest-*.xml is for content addin', function(done){

      // verify manifest-*.xml exists
      assert.file(manifestFileName);

      // load manifest-*.xml as JSON
      var parser = new Xml2Js.Parser();
      fs.readFile(manifestFileName, 'utf8', function(err, manifestContent){
        parser.parseString(manifestContent, function(err, manifestJson){

          // check addin type
          expect(manifestJson.OfficeApp.$['xsi:type']).to.equal('ContentApp');

          done();
        });
      });

    });

  }); // describe('runs content subgenerator')

}); // describe('office:app')
