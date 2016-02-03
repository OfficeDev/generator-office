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

describe('office:commands', function () {
  var projectDisplayName = 'My Office Commands';
  var projectEscapedName = 'my-office-commands';
  var manifestFileName = 'manifest-' + projectEscapedName + '.xml';

  beforeEach(function (done) {
    options = {
      name: projectDisplayName
    };
    done();
  });

  describe('Outlook add-in', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      options.type = 'mail';
      options['manifest-file'] = manifestFileName;
      done();
    });
    
    describe('Called with non-existent manifest', function () {
      beforeEach(function (done) {

        // set extension points
        options.extensionPoint = ['CustomPane'];
        // run the generator
        helpers.run(path.join(__dirname, '../../generators/commands'))
          .withOptions(options)
          .on('end', function() {
            done();
          });
      });
      
      afterEach(function () {
        mockery.disable();
      });
      
      it('creates no files', function (done) {
        var unexpected = [
          manifestFileName,
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
    }); // describe('Called with non-existent manifest')
        
    describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface, '
           + 'AppointmentAttendeeCommandSurface, AppointmentOrganizerCommandSurface, CustomPane', 
    function () {
    
      beforeEach(function (done) {

        // set extension points
        options.extensionPoint = [
          'MessageReadCommandSurface', 
          'MessageComposeCommandSurface', 
          'AppointmentAttendeeCommandSurface', 
          'AppointmentOrganizerCommandSurface', 
          'CustomPane'
        ];
        
        prompts.buttonTypes = ['uiless', 'menu', 'taskpane'];
        prompts.commandContainers = ['TabDefault', 'TabCustom'];

        // run the generator
        helpers.run(path.join(__dirname, '../../generators/commands'))
          .withOptions(options)
          .withPrompts(prompts)
          .on('ready', function(gen) {
            util.setupExistingManifest(gen, manifestFileName);
          }.bind(this))
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
            .property('xmlns').equal('http://schemas.microsoft.com/office/mailappversionoverrides');
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
              // Validate tabs are present
              expect(extPoint, 'MessageReadCommandSurface has default tab').to.have.deep
                .property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'MessageReadCommandSurface has custom tab').to.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.true;
              
              hasUiless = false;
              hasMenu = false;
              hasTaskpane = false;
              
              _.forEach(extPoint.CustomTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Custom tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Custom tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Custom tab has taskpane button').to.be.true;
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
              
              // Validate tabs are present
              expect(extPoint, 'MessageComposeCommandSurface has default tab').to.have.deep.
                property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'MessageComposeCommandSurface has custom tab').to.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.true;
              
              hasUiless = false;
              hasMenu = false;
              hasTaskpane = false;
              
              _.forEach(extPoint.CustomTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Custom tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Custom tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Custom tab has taskpane button').to.be.true;
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
              
              // Validate tabs are present
              expect(extPoint, 'AppointmentAttendeeCommandSurface has default tab')
                .to.have.deep.property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'AppointmentAttendeeCommandSurface has custom tab')
                .to.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.true;
              
              hasUiless = false;
              hasMenu = false;
              hasTaskpane = false;
              
              _.forEach(extPoint.CustomTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Custom tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Custom tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Custom tab has taskpane button').to.be.true;
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
              
              // Validate tabs are present
              expect(extPoint, 'AppointmentOrganizerCommandSurface has default tab')
                .to.have.deep.property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'AppointmentOrganizerCommandSurface has custom tab')
                .to.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.true;
              
              hasUiless = false;
              hasMenu = false;
              hasTaskpane = false;
              
              _.forEach(extPoint.CustomTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Custom tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Custom tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Custom tab has taskpane button').to.be.true;
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
          expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources').with.length(1);
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
    
    describe('Outlook extension points: MessageReadCommandSurface, MessageComposeCommandSurface', 
    function () {
    
      beforeEach(function (done) {

        // set extension points
        options.extensionPoint = [
          'MessageReadCommandSurface', 
          'MessageComposeCommandSurface'
        ];
        
        prompts.buttonTypes = ['uiless', 'menu'];
        prompts.commandContainers = ['TabDefault'];

        // run the generator
        helpers.run(path.join(__dirname, '../../generators/commands'))
          .withOptions(options)
          .withPrompts(prompts)
          .on('ready', function(gen) {
            util.setupExistingManifest(gen, manifestFileName);
          }.bind(this))
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
          addinRootPath + '/TaskPane/TaskPane.html',
          addinRootPath + '/TaskPane/TaskPane.js',
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
            .property('xmlns').equal('http://schemas.microsoft.com/office/mailappversionoverrides');
          done();
        });
        
        /**
        * Hosts is present and has only one
        * Host element with xsi:type=MailHost
        */
        it('has MailHost Host entry', function(done) {
          expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Hosts').with.length(1);
          expect(manifest.OfficeApp.VersionOverrides[0].Hosts[0]).to.have.property('Host').with.length(1);
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
              
              // Validate tabs are present
              expect(extPoint, 'MessageReadCommandSurface has default tab').to.have.deep
                .property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'MessageReadCommandSurface has custom tab').to.not.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.false;
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
              
              // Validate tabs are present
              expect(extPoint, 'MessageComposeCommandSurface has default tab').to.have.deep
                .property('OfficeTab[0]').with.property('$').with.property('id').equal('TabDefault');
              expect(extPoint, 'MessageComposeCommandSurface has custom tab').to.not.have.deep.property('CustomTab[0]');
              
              var hasUiless = false;
              var hasMenu = false;
              var hasTaskpane = false;
              
              _.forEach(extPoint.OfficeTab[0].Group[0].Control, function(control) {
                if (control.$['xsi:type'] === 'Menu') {
                  hasMenu = true;
                }
                else {
                  if (control.Action[0].$['xsi:type'] === 'ExecuteFunction') {
                    hasUiless = true;
                  }
                  else if (control.Action[0].$['xsi:type'] === 'ShowTaskpane') {
                    hasTaskpane = true;
                  }
                }
              });
              
              // Validate tabs have a uiless button
              expect(hasUiless, 'Default tab has uiless button').to.be.true;
              
              // Validate tabs have a menu button
              expect(hasMenu, 'Default tab has menu button').to.be.true;
              
              // Validate tabs have a taskpane button
              expect(hasTaskpane, 'Default tab has taskpane button').to.be.false;
            }
          });
          expect(found, '<ExtensionPoint xsi:type="MessageComposeCommandSurface"> exists').to.be.true;
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
          expect(found, '<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface"> exists').to.be.false;
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
          expect(found, '<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface"> exists').to.be.false;
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
          expect(manifest.OfficeApp.VersionOverrides[0]).to.have.property('Resources').with.length(1);
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
      
  }); // describe('Outlook add-in')
  
  describe('Taskpane add-in', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      options.type = 'taskpane';
      options['manifest-file'] = manifestFileName;
      done();
    });
    
    describe('Called with non-existent manifest', function () {
      beforeEach(function (done) {

        // set extension points
        // run the generator
        helpers.run(path.join(__dirname, '../../generators/commands'))
          .withOptions(options)
          .on('end', function() {
            done();
          });
      });
      
      afterEach(function () {
        mockery.disable();
      });
      
      it('creates no files', function (done) {
        var unexpected = [
          manifestFileName,
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
    }); // describe('Called with non-existent manifest')
  }); // describe('Taskpane add-in')
  
  describe('Content add-in', function () {
    var addinRootPath = 'src/public';

    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    beforeEach(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      options.type = 'content';
      options['manifest-file'] = manifestFileName;
      done();
    });
    
    describe('Called with non-existent manifest', function () {
      beforeEach(function (done) {

        // set extension points
        // run the generator
        helpers.run(path.join(__dirname, '../../generators/commands'))
          .withOptions(options)
          .on('end', function() {
            done();
          });
      });
      
      afterEach(function () {
        mockery.disable();
      });
      
      it('creates no files', function (done) {
        var unexpected = [
          manifestFileName,
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
    }); // describe('Called with non-existent manifest')
  }); // describe('Content add-in')
  
}); // describe('office:commands')
