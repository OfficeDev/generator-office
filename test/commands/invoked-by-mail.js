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

describe('office:mail -> office:commands', function () {

  var projectDisplayName = 'My Office Commands';
  var projectEscapedName = 'my-office-commands';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  beforeEach(function (done) {
    options = {
      name: projectDisplayName
    };
    done();
  });

  describe('called with custom command data', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      done();
    });

    /**
     * Set technology = html, the mail html template passes
     * custom command data
     */
    describe('addin technology:html', function () {
      
      describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface,' 
             + 'AppointmentAttendeeCommandSurface, AppointmentOrganizerCommandSurface, CustomPane', 
      function () {
      
        beforeEach(function (done) {
          // set language to html
          options.tech = 'html';
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'MessageComposeCommandSurface', 
            'AppointmentAttendeeCommandSurface', 
            'AppointmentOrganizerCommandSurface', 
            'CustomPane'
          ];
  
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
          var expected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png'
          ];
          assert.file(expected);
          done();
        });
        
        /**
          * manifest-*.xml is good
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
          * VersionOverrides is present and uses
          * correct xmlns for mail
          */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
          * Hosts is present and has only one
          * Host element with xsi:type=MailHost
          */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have.property('Host')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have.property('$')
              .with.property('xsi:type').equal('MailHost'); 
            done();
          });
          
          /**
          * ExtensionPoint for MessageReadCommandSurface is present 
          */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists').to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for MessageComposeCommandSurface is present 
          */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists').to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentAttendeeCommandSurface is present 
          */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists').to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentOrganizerCommandSurface is present 
          */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists').to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for CustomPane is present 
          */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
          * Resources node is present with correct
          * child nodes
          */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have.property
              ('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have.property
              ('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have.property
              ('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have.property
              ('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
      
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // MessageComposeCommandSurface, AppointmentAttendeeCommandSurface, 
          // AppointmentOrganizerCommandSurface, CustomPane')
      
      describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface', 
      function () {
      
        beforeEach(function (done) {
          // set language to html
          options.tech = 'html';
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'MessageComposeCommandSurface'
          ];
  
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
          var expected = [
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png'
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js'
          ];
          
          assert.noFile(unexpected);
          done();
        });
        
        /**
          * manifest-*.xml is good
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
          * VersionOverrides is present and uses
          * correct xmlns for mail
          */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
          * Hosts is present and has only one
          * Host element with xsi:type=MailHost
          */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
          * ExtensionPoint for MessageReadCommandSurface is present 
          */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for MessageComposeCommandSurface is present 
          */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
          */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
          */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for CustomPane is not present 
          */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.false;
            done();
          });
          
          /**
          * Resources node is present with correct
          * child nodes
          */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
      
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // MessageComposeCommandSurface')
          
      describe('Outlook extension points: CustomPane', 
      function () {
      
        beforeEach(function (done) {
          // set language to html
          options.tech = 'html';
  
          // set extension points
          options.extensionPoint = [
            'CustomPane'
          ];
  
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
          var expected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png'
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          
          assert.noFile(unexpected);
          done();
        });
        
        /**
          * manifest-*.xml is good
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
          * VersionOverrides is present and uses
          * correct xmlns for mail
          */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
          * Hosts is present and has only one
          * Host element with xsi:type=MailHost
          */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
          * ExtensionPoint for MessageReadCommandSurface is not present 
          */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for MessageComposeCommandSurface is not present 
          */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
          */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
          */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
          * ExtensionPoint for CustomPane is present 
          */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
          * Resources node is present with correct
          * child nodes
          */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
      
      }); // describe('Outlook extension points: CustomPane')
      
    }); // describe('technology:html')

  }); // describe('called with custom command data')
  
  describe('called without custom command data', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      done();
    });

    /**
     * Set technology = ng, the mail ng template doesn't
     * pass custom command data
     */
    describe('addin technology:ng', function () {
      
      describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface,' 
            +  'AppointmentAttendeeCommandSurface, AppointmentOrganizerCommandSurface, CustomPane', 
      function () {
        
        beforeEach(function (done) {
          // set language to ng
          options.tech = 'ng';
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'MessageComposeCommandSurface', 
            'AppointmentAttendeeCommandSurface', 
            'AppointmentOrganizerCommandSurface', 
            'CustomPane'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
  
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
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.file(expected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // MessageComposeCommandSurface, AppointmentAttendeeCommandSurface,
          // AppointmentOrganizerCommandSurface, CustomPane')
      
      describe('Outlook extension points: MessageReadCommandSurface, '
             + 'AppointmentAttendeeCommandSurface', 
      function () {
        
        beforeEach(function (done) {
          // set language to ng
          options.tech = 'ng';
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'AppointmentAttendeeCommandSurface'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['uiless'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png'
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is not present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is not present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.false;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // AppointmentAttendeeCommandSurface')
      
      describe('Outlook extension points: MessageComposeCommandSurface, '
             + 'AppointmentOrganizerCommandSurface', 
      function () {
        
        beforeEach(function (done) {
          // set language to ng
          options.tech = 'ng';
  
          // set extension points
          options.extensionPoint = [
            'MessageComposeCommandSurface', 
            'AppointmentOrganizerCommandSurface'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js'
          ];
          
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is not present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is not present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.false;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageComposeCommandSurface, 
          // AppointmentOrganizerCommandSurface')
      
      describe('Outlook extension points: CustomPane', function () {
        
        beforeEach(function (done) {
          // set language to ng
          options.tech = 'ng';
  
          // set extension points
          options.extensionPoint = ['CustomPane'];
          
          // set tabs
          //prompts.commandContainers = ['TabDefault'];
          //prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png'
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is not present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is not present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: CustomPane')
      
    }); // describe('technology:ng')

  }); // describe('called without custom command data')
  
  describe('called with manifest-only', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      options.tech = 'manifest-only';
      done();
    });
    
    describe('addin technology:manifest-only', function () {
      
      describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface,' 
            +  'AppointmentAttendeeCommandSurface, AppointmentOrganizerCommandSurface, CustomPane', 
      function () {
        
        beforeEach(function (done) {
          // set language to ng
          
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'MessageComposeCommandSurface', 
            'AppointmentAttendeeCommandSurface', 
            'AppointmentOrganizerCommandSurface', 
            'CustomPane'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
          prompts.startPage = 'https://localhost:8443/manifest-only/index.html';
          prompts.functionFileUrl = 'https://localhost:8443/manifest-only/function.html';
          prompts.taskPaneUrl = 'https://localhost:8443/manifest-only/taskpane.html';
          prompts.iconUrl = 'https://localhost:8443/manifest-only/icon.png';
  
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
            manifestFileName
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // MessageComposeCommandSurface, AppointmentAttendeeCommandSurface,
          // AppointmentOrganizerCommandSurface, CustomPane')
      
      describe('Outlook extension points: MessageReadCommandSurface, '
             + 'AppointmentAttendeeCommandSurface', 
      function () {
        
        beforeEach(function (done) {
  
          // set extension points
          options.extensionPoint = [
            'MessageReadCommandSurface', 
            'AppointmentAttendeeCommandSurface'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['uiless'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            manifestFileName
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is not present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is not present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.false;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageReadCommandSurface, 
          // AppointmentAttendeeCommandSurface')
      
      describe('Outlook extension points: MessageComposeCommandSurface, '
             + 'AppointmentOrganizerCommandSurface', 
      function () {
        
        beforeEach(function (done) {
  
          // set extension points
          options.extensionPoint = [
            'MessageComposeCommandSurface', 
            'AppointmentOrganizerCommandSurface'
          ];
          
          // set tabs
          prompts.commandContainers = ['TabDefault'];
          prompts.buttonTypes = ['menu', 'taskpane'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            manifestFileName
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is not present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.true;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is not present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.false;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: MessageComposeCommandSurface, 
          // AppointmentOrganizerCommandSurface')
      
      describe('Outlook extension points: CustomPane', function () {
        
        beforeEach(function (done) {
  
          // set extension points
          options.extensionPoint = ['CustomPane'];
          
          // set tabs
          //prompts.commandContainers = ['TabDefault'];
          //prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
  
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
         * All expected files are created,
         * unexpected are not created
         */
        it('creates expected files', function (done) {
          var expected = [
            manifestFileName
          ];
          assert.file(expected);
          
          var unexpected = [
            addinRootPath + '/CustomPane/CustomPane.html',
            addinRootPath + '/CustomPane/CustomPane.js',
            addinRootPath + '/FunctionFile/Functions.html',
            addinRootPath + '/FunctionFile/Functions.js',
            addinRootPath + '/images/icon-16.png',
            addinRootPath + '/images/icon-32.png',
            addinRootPath + '/images/icon-80.png',
            addinRootPath + '/TaskPane/TaskPane.html',
            addinRootPath + '/TaskPane/TaskPane.js'
          ];
          assert.noFile(unexpected);
          done();
        });
        
        /**
         * manifest-*.xml is good
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
           * VersionOverrides is present and uses
           * correct xmlns for mail
           */
          it('has valid VersionOverrides', function(done) {
            expect(manifest.OfficeApp).to.have.property('VersionOverrides').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('$').with
              .property('xmlns').equal(
                'http://schemas.microsoft.com/office/mailappversionoverrides');
            done();
          });
          
          /**
           * Hosts is present and has only one
           * Host element with xsi:type=MailHost
           */
          it('has MailHost Host entry', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have
              .property('Host').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0]).to.have
              .property('$').with.property('xsi:type').equal('MailHost');
            done();
          });
          
          /**
           * ExtensionPoint for MessageReadCommandSurface is not present 
           */
          it('has MessageReadCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageReadCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageReadCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for MessageComposeCommandSurface is not present 
           */
          it('has MessageComposeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'MessageComposeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentAttendeeCommandSurface is not present 
           */
          it('has AppointmentAttendeeCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentAttendeeCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for AppointmentOrganizerCommandSurface is not present 
           */
          it('has AppointmentOrganizerCommandSurface', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'AppointmentOrganizerCommandSurface') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists')
              .to.be.false;
            done();
          });
          
          /**
           * ExtensionPoint for CustomPane is present 
           */
          it('has CustomPane', function(done) {
            var found = false;
            
            _.forEach(manifest.OfficeApp.VersionOverrides[0].Hosts[0].Host[0].DesktopFormFactor[0]
              .ExtensionPoint, function(extPoint) {
              if(extPoint.$['xsi:type'] === 'CustomPane') {
                found = true;
              }
            });
            expect(found, '<ExtensionPoint xsi:type="CustomPane"> exists').to.be.true;
            done();
          });
          
          /**
           * Resources node is present with correct
           * child nodes
           */
          it('has valid Resources', function(done) {
            expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources')
              .with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Images').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:Urls').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:ShortStrings').with.length(1);
            expect(manifest.OfficeApp.VersionOverrides[0].Resources[0]).to.have
              .property('bt:LongStrings').with.length(1);
            done();
          });
          
        }); // describe('manifest-*.xml contents')
        
      }); // describe('Outlook extension points: CustomPane')
      
    }); // describe('technology:manifest-only')

  }); // describe('called with manifest-only')

}); // describe('office:mail -> office:commands')
