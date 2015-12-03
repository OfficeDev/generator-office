'use strict';

var generators = require('yeoman-generator');
var chalk = require('chalk');
var path = require('path');
var _ = require('lodash');
var extend = require('deep-extend');
var guid = require('uuid');
var Xml2Js = require('xml2js');

module.exports = generators.Base.extend({
  /**
   * Setup the generator
   */
  constructor: function(){
    generators.Base.apply(this, arguments);
    
    // Shared options
    this.option('type', {
      type: String,
      required: false,
      desc: 'Add-in type (mail, taskpane, content)'
    });
    
    this.option('tech', {
      type: String,
      desc: 'Technology to use for the Add-in (html = HTML; ng = Angular)',
      required: false
    });
    
    this.option('root-path', {
      type: String,
      desc: 'Relative path where the project should be created (blank = current directory)',
      required: false
    });
    
    this.option('manifest-file', {
      type: String,
      desc: 'Relative path to manifest file',
      required: false
    });
    
    this.option('extensionPoints', {
      type: String,
      desc: 'Supported extension points',
      required: false
    });
    
    // If commands are passed as an option, it will be a
    // JSON object with all the details we need
    // extensionPoints and clients will be ignored
    this.option('commands', {
      type: String,
      desc: 'A JSON-formatted string defining the commands to add',
      required: false
    });
    
    // Task-pane/Content pane-specific options
    this.option('clients', {
      type: String,
      desc: 'Office client product that can host the add-in',
      required: false
    });
    
    // create global config object on this generator
    this.genConfig = {};
  }, // constructor()
  
  /**
   * Prompt users for options
   */
  prompting: {
    
    askFor: function(){
      var done = this.async();
      
      var prompts = [
        {
          name: 'type',
          message: 'Office project type:',
          type: 'list',
          choices: [
            {
              name: 'Outlook Add-in',
              value: 'mail'
            },
            {
              name: 'Task Pane Add-in',
              value: 'taskpane'
            },

            {
              name: 'Content Add-in',
              value: 'content'
            }
          ],
          when: this.options.type === undefined
        },
        {
          name: 'root-path',
          message: 'Root folder of project?'
          + ' Default to current directory\n'
          + ' (' + this.destinationRoot() + '),'
          + ' or specify relative path\n'
          + ' from current (src / public): ',
          default: 'current folder',
          when: this.options['root-path'] === undefined,
          filter: /* istanbul ignore next */ function(response){
            if (response === 'current folder') {
              return '';
            } else {
              return response;
            }
          }
        },
        // technology used to create the addin (html / angular / etc)
        {
          name: 'tech',
          message: 'Technology to use:',
          type: 'list',
          when: this.options.tech === undefined,
          choices: [
            {
              name: 'HTML, CSS & JavaScript',
              value: 'html'
            }, {
              name: 'Angular',
              value: 'ng'
            }, {
              name: 'Manifest.xml only (no application source files)',
              value: 'manifest-only'
            }]
        },
        {
          name: 'manifest-file',
          message: 'Relative path to manifest file:',
          default: 'manifest.xml',
          when: this.options['manifest-file'] === undefined
        }
      ];
      
      // trigger prompts
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, this.options);
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));
    }, // askFor()
    
    /**
     * ask for hosts
     */
    askForHosts: function() {
      // If commands were passed as an option, then don't prompt
      if (this.genConfig.commands !== undefined) {
        return;
      }
      
      switch(this.genConfig.type){
        case 'mail':
          // Only one host for mail, so no need to prompt
          this.genConfig = extend(this.genConfig, { hosts: [ 'MailHost' ] });
          return;
        case 'taskpane':
          // TODO: Setup prompt for available hosts
          break;
        case 'content':
          // TODO: Setup prompt for available hosts
          break;
      }
    }, // askForHosts()
    
    /**
     * ask for form factors
     */
    askForFormFactors: function() {
      // Currently only Desktop is supported, so no need to prompt.
      // Just set form factors to only desktop and move on.
      
      // When support for other form factors is added will need to make this
      // into a prompt.
      
      // If commands were passed as an option, then don't prompt
      if (this.genConfig.commands !== undefined) {
        return;
      }
      
      this.genConfig = extend(this.genConfig, {formFactors: [ 'DesktopFormFactor' ]});
    }, // askForFormFactors()
    
    /**
     * ask for extension points based on add-in type
     */
    askForExtensionPoints: function() {
      // If commands were passed as an option, then don't prompt
      if (this.genConfig.commands !== undefined) {
        return;
      }
      
      var availableExtensionPoints = undefined;
      
      switch(this.genConfig.type) {
        case 'mail':
          availableExtensionPoints = [
            {
              name: 'Message read commands',
              value: 'MessageReadCommandSurface',
              checked: true
            },
            {
              name: 'Message compose commands',
              value: 'MessageComposeCommandSurface',
              checked: true
            },
            {
              name: 'Appointment organizer commands',
              value: 'AppointmentOrganizerCommandSurface',
              checked: true
            },
            {
              name: 'Appointment attendee commands',
              value: 'AppointmentAttendeeCommandSurface',
              checked: true
            },{
              name: 'Custom pane (for message read and appointment attendee)',
              value: 'CustomPane',
              checked: false
            }
          ];
          break;
        case 'taskpane':
          // TODO: set available extension points
          break;
        case 'content':
          // TODO: set available extension points
          break;
      }
      
      if (availableExtensionPoints !== undefined) {
        var prompts = [
          {
            name: 'extensionPoints',
            message: 'Supported extension points:',
            type: 'checkbox',
            when: this.genConfig.extensionPoints === undefined,
            choices: availableExtensionPoints,
            validate: function(answers) {
              if (answers.length < 1) {
                return 'Must select at least one extension point';
              }
              return true;
            }
          }
        ];
        
        var done = this.async();
        this.prompt(prompts, function(responses){
          this.genConfig = extend(this.genConfig, responses);
          done();
        }.bind(this));
      }
    }, // askForExtensionPoints()
    
    /**
     * ask for CustomPane details
     */
    askForCustomPane: function() {
      if (this.genConfig.commands !== undefined ||
          this.genConfig.type !== 'mail' || 
          this.genConfig.extensionPoints.indexOf('CustomPane') < 0) {
        return;
      }
      
      var prompts = [
        {
          name: 'requestedHeight',
          message: 'Requested height in pixels for custom pane',
          default: 200,
          filter: function(input) {
            if (typeof input === 'number') {
              return input;
            }
            return parseInt(input, 10);
          },
          validate: function(response) {
            var numVal = response;
            if (typeof response !== 'number')
            {
              numVal = parseInt(response, 10);
            }
            
            if (isNaN(numVal) || numVal < 32 || numVal > 450) {
              return 'Please enter a valid integer between 32 and 450'      
            }
            return true;
          } 
        },
        {
          name: 'sourceLocation',
          message: 'Relative path to source page for custom pane',
          default: '/CustomPane/CustomPane.html'
        }
        // Possible enhancements:
        // - Add rich rule building prompts here
        // - Ask for disable entity highlighting
      ];
      
      var done = this.async();
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, { customPaneOptions: responses });
        done();
      }.bind(this));
    },  // askForCustomPane()
    
    /**
     * ask for *CommandSurface details
     */
    askForCommandSurface: function() {
      if (this.genConfig.commands !== undefined ||
          !commandSurfaceIncluded(this.genConfig.extensionPoints)) {
        return;
      }
      
      var availableContainers = [];
      
      for (var i = 0; i < this.genConfig.hosts.length; i++) {
        switch (this.genConfig.hosts[i]) {
          case 'MailHost':
            availableContainers.push({ name: 'Default tab', value: 'TabDefault', checked: true });
            availableContainers.push({ name: 'Custom tab', value: 'TabCustom', checked: true });
            break;
          // TODO: Add other hosts here
        }
      }
      
      if (availableContainers.length > 0) {
        var prompts = [
          {
            name: 'commandContainers',
            message: 'Add buttons to:',
            type: 'checkbox',
            choices: availableContainers,
            validate: function(answers) {
              if (answers.length < 1) {
                return 'Must select at least one container to add buttons to';
              }
              return true;
            }
          },
          {
            name: 'buttonTypes',
            message: 'Supported button types:',
            type: 'checkbox',
            choices: [
              {
                name: 'UI-less button',
                value: 'uiless',
                checked: true
              },
              {
                name: 'Drop-down menu button',
                value: 'menu',
                checked: true
              },
              {
                name: 'Task-pane launcher button',
                value: 'taskpane',
                checked: true
              }
            ]
          }
        ]
        
        var done = this.async();
        this.prompt(prompts, function(response) {
          this.genConfig = extend(this.genConfig, response);
          done();
        }.bind(this));
      }
    }
  }, // prompting()
  
  /**
   * save configurations & config project
   */
  configuring: function(){
    
    // helper function to build path to the file off root path
      this._parseTargetPath = function(file){
        return path.join(this.genConfig['root-path'], file);
      };
    
    // Build up a JSON-representation of the VersionOverrides
    // element here.
    
    // If the caller passed this in, we can build directly from it
    if (this.genConfig.commands === undefined)
    {
      // Initialize resource arrays
      this.genConfig.resources = {
        urls: [],
        images: [],
        shortStrings: [],
        longStrings: []
      };
      
      // Determine if a function file is needed
      var needFuncFile = this.genConfig.buttonTypes.indexOf('uiless') >= 0;
      if (needFuncFile) {
        // Use the default function file
        this.genConfig.functionFileId = createUrlResource('funcFile', '',
          'https://localhost:8443/FunctionFile/Functions.html', this.genConfig);
      }
      
      // Set up control counters
      this.genConfig.customContainerCount = 0;
      this.genConfig.groupCount = 0;
      this.genConfig.uilessCount = 0;
      this.genConfig.menuCount = 0;
      this.genConfig.taskPaneCount = 0;
      
      _.forEach(this.genConfig.hosts, function(hostType){
        // foreach formfactor
          // foreach applicable extensionpoint
            // if custompane
              // do custompane resources
            // else
              // foreach command tab
                // foreach button type
                  // add resources
      });
    }
  }, // configuring()
  
  /**
   * write generator specific files
   */
  writing: {
    /**
     * Update the manifest
     */
    
    updateManifest: function() {
      var done = this.async();
      
      var manifestFile = this.genConfig['manifest-file'];
      var modifiedManifestFile = manifestFile.replace('.xml', '-cmd.xml');
      
      // workaround to 'this' context issue
      var yoGenerator = this;
      
      // make sure manifest exists
      if (!yoGenerator.fs.exists(manifestFile)){
        this.log('Specified manifest "', manifestFile, '" not found. Exiting...');
        return;
      }
      
      // load manifest XML
      var manifestXml = yoGenerator.fs.read(yoGenerator.destinationPath(manifestFile));
      
      // convert to JSON
      var parser = new Xml2Js.Parser();
      
      var test = '<?xml version="1.0" encoding="UTF-8"?>';
      test += '<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->';
      test += '<OfficeApp/>';
      
      parser.parseString(manifestXml, function(err, manifestJson) {
        
        // Add namespaces to the OfficeApp element
        var newNS = {
          '$': {
            'xmlns:bt': 'http://schemas.microsoft.com/office/officeappbasictypes/1.0',
            'xmlns:mailappor': 'http://schemas.microsoft.com/office/mailappversionoverrides/1.0'
          }
        };
        manifestJson.OfficeApp = extend(manifestJson.OfficeApp, newNS);
        
        // Create VersionOverrides
        manifestJson.OfficeApp.VersionOverrides = {
          '$': {
            xmlns: 'http://schemas.microsoft.com/office/mailappversionoverrides',
            'xsi:type': 'VersionOverridesV1_0'
          },
        };
        
        // Build array of Host elements
        var hosts = [];
        _.forEach(yoGenerator.genConfig.hosts, function(hostType){
          // Initialize host
          var host = {
            '$': {
              'xsi:type': hostType
            }
          };
          
          // Add form factors
          _.forEach(yoGenerator.genConfig.formFactors, function(factorType){
            host[factorType] = buildFormFactor(yoGenerator.genConfig);
          });
          
          hosts.push(host);
        });
        
        manifestJson.OfficeApp.VersionOverrides.Hosts = { Host: hosts };
        
        // Sort resources by id to make it easier
        // to find specific resources in the manifest
        yoGenerator.genConfig.resources.images.sort(function(a,b){
          return (a['$'].id.localeCompare(b['$'].id));
        });
        yoGenerator.genConfig.resources.urls.sort(function(a,b){
          return (a['$'].id.localeCompare(b['$'].id));
        });
        yoGenerator.genConfig.resources.shortStrings.sort(function(a,b){
          return (a['$'].id.localeCompare(b['$'].id));
        });
        yoGenerator.genConfig.resources.longStrings.sort(function(a,b){
          return (a['$'].id.localeCompare(b['$'].id));
        });
        
        manifestJson.OfficeApp.VersionOverrides.Resources = {
          'bt:Images': { 'bt:Image': yoGenerator.genConfig.resources.images },
          'bt:Urls': { 'bt:Url': yoGenerator.genConfig.resources.urls },
          'bt:ShortStrings': { 'bt:String': yoGenerator.genConfig.resources.shortStrings },
          'bt:LongStrings': { 'bt:String': yoGenerator.genConfig.resources.longStrings }
        };
        
        // convert JSON => XML
        var xmlBuilder = new Xml2Js.Builder();
        var updatedManifestXml = xmlBuilder.buildObject(manifestJson);
        
        // write updated manifest
        yoGenerator.fs.write(yoGenerator.destinationPath(modifiedManifestFile), updatedManifestXml);
        
        done();
      });
    }, // updateManifest();
    
    /**
     * Add supporting files
     */
    addFiles: function() {
      this.fs.copyTpl(this.templatePath('common/FunctionFile/Functions.js'),
                      this.destinationPath('FunctionFile/Functions.js'),
                      this.genConfig);
      this.fs.copy(this.templatePath('common/FunctionFile/Functions.html'),
                   this.destinationPath(this._parseTargetPath('FunctionFile/Functions.html')));
      this.fs.copy(this.templatePath('common/TaskPane/TaskPane.html'),
                   this.destinationPath(this._parseTargetPath('TaskPane/TaskPane.html')));
      this.fs.copy(this.templatePath('common/TaskPane/TaskPane.js'),
                   this.destinationPath(this._parseTargetPath('TaskPane/TaskPane.js')));  
      this.fs.copy(this.templatePath('common/Images/icon-16.png'),
                   this.destinationPath(this._parseTargetPath('Images/icon-16.png')));
      this.fs.copy(this.templatePath('common/Images/icon-32.png'),
                   this.destinationPath(this._parseTargetPath('Images/icon-32.png')));
      this.fs.copy(this.templatePath('common/Images/icon-80.png'),
                   this.destinationPath(this._parseTargetPath('Images/icon-80.png')));
    }
  } // writing()
});

/**
 * Returns true if any of the known command surface
 * extension points are in the array
 */
function commandSurfaceIncluded(extensionPoints) {
  // Be sure to add applicable command surfaces here
  return (extensionPoints.indexOf('MessageReadCommandSurface') >= 0 ||
          extensionPoints.indexOf('MessageComposeCommandSurface') >= 0 ||
          extensionPoints.indexOf('AppointmentAttendeeCommandSurface') >= 0 ||
          extensionPoints.indexOf('AppointmentOrganizerCommandSurface') >= 0);
}

/**
 * Builds out a host element as a JSON object
 */

function buildFormFactor(config) {
  var factor = {};
  
  if (config.functionFileId !== undefined) {
    factor.FunctionFile = { 
      '$': { resid: config.functionFileId }
    };
  }
  
  var extensionPoints = [];
  _.forEach(config.extensionPoints, function(extensionType){
    extensionPoints.push(buildExtensionPoint(extensionType, config));
  });
  
  factor.ExtensionPoint = extensionPoints;
  
  return factor;
}

/**
 * Builds out an extension point
 */
function buildExtensionPoint(type, config) {
  var extPoint = {
    '$': { 'xsi:type': type }
  };
  
  if (type === 'CustomPane') {
    // Build custom pane
  }
  else {
    // Build a command surface
    _.forEach(config.commandContainers, function(containerId){
      var container = buildControlContainer(containerId, config);
      extPoint[container.nodeName] = container.node;
    });
  }
  
  return extPoint;
}

/**
 * Builds out a control container
 */
function buildControlContainer(type, config) {
  
  var container = {};
  switch(type){
    case 'TabDefault': // Default tab (used by Outlook)
      container.nodeName = 'OfficeTab';
      container.node = {
        '$': { id: type }
      };
      break;
    case 'TabCustom': // Custom tab 
      config.customContainerCount++;
      container.nodeName = 'CustomTab'
      container.node = {
        '$': { id: type + config.customContainerCount },
        Label: { 
          '$': { 
            resid: createShortStringResource('customTabLabel', config.customContainerCount,
              'Custom Tab ' + config.customContainerCount, config)
          }
        }
      };

      break;
  }
  
  container.node.Group = buildGroup(config);
  return container;
}

/**
 * Builds out a group
 */
function buildGroup(config) {
  
  config.groupCount++;
  
  var group = {
    '$': { id: 'group' + config.groupCount },
    Label: { 
      '$': { 
        resid: createShortStringResource('groupLabel', config.groupCount, 
          'Group ' + config.groupCount, config) 
      } 
    }
  };
  
  var buttons = [];
  _.forEach(config.buttonTypes, function(buttonType){
    switch(buttonType) {
      case 'uiless':
        buttons.push(buildUiLessButton(config));
        break;
      case 'menu':
        buttons.push(buildMenu(config));
        break;
      case 'taskpane':
        buttons.push(buildTaskPaneButton(config));
        break;
    }
  });
  
  group.Control = buttons;
  
  return group;
}

/**
 * Builds out a uiless button
 */
function buildUiLessButton(config) {
  config.uilessCount++;
  
  var button = {
    '$': {
      'xsi:type': 'Button',
      id: 'uilessButton' + config.uilessCount
    },
    Label: { 
      '$': { 
        resid: createShortStringResource('uilessButtonLabel', config.uilessCount,
          'UI-less Button ' + config.uilessCount, config)
      } 
    },
    Tooltip: { 
      '$': { 
        resid: createLongStringResource('uilessButtonToolTip', config.uilessCount,
          'This is the tooltip for UI-less Button ' + config.uilessCount, config) 
      } 
    },
    Supertip: {
      Title: { 
        '$': { 
          resid: createShortStringResource('uilessButtonSuperTipTitle', config.uilessCount,
            'UI-less Button ' + config.uilessCount, config) 
        } 
      },
      Description: { 
        '$': { 
          resid: createLongStringResource('uilessButtonSuperTipDesc', config.uilessCount,
            'This is the description for UI-less Button ' + config.uilessCount, config)
        } 
      }
    },
    Icon: {
      'bt:Image': [
        {
          '$': {
            size: 16,
            resid: createImageResource('uilessButtonIcon', config.uilessCount + '-16',
              'https://localhost:8443/images/icon-16.png', config)
          }
        },
        {
          '$': {
            size: 32,
            resid: createImageResource('uilessButtonIcon', config.uilessCount + '-32',
              'https://localhost:8443/images/icon-32.png', config)
          }
        },
        {
          '$': {
            size: 80,
            resid: createImageResource('uilessButtonIcon', config.uilessCount + '-80',
              'https://localhost:8443/images/icon-80.png', config)
          }
        }
      ] 
    },
    Action: {
      '$': { 'xsi:type': 'ExecuteFunction' },
      FunctionName: 'buttonFunction' + config.uilessCount
    }
  };
  
  return button;
}

/**
 * Build out a menu button
 */
function buildMenu(config) {
  config.menuCount++;
  
  // Create a UI-less button to put inside the menu
  var uilessButton = buildUiLessButton(config);
  // Remove the 'xsi:type' attribute from the button, it isn't
  // used in menu items.
  delete uilessButton.$['xsi:type'];
  
  var menu = {
    '$': {
      'xsi:type': 'Menu',
      id: 'menu' + config.menuCount
    },
    Label: {
      '$': {
        resid: createShortStringResource('menuLabel', config.menuCount,
        'Menu ' + config.menuCount, config)
      }
    },
    Tooltip: {
      '$': {
        resid: createLongStringResource('menuToolTip', config.menuCount,
        'This is the tooltip for Menu ' + config.menuCount, config)
      }
    },
    Supertip: {
      Title: { 
        '$': { 
          resid: createShortStringResource('menuSuperTipTitle', config.menuCount,
            'Menu ' + config.menuCount, config) 
        } 
      },
      Description: { 
        '$': { 
          resid: createLongStringResource('menuSuperTipDesc', config.menuCount,
            'This is the description for Menu ' + config.menuCount, config)
        } 
      }
    },
    Icon: {
      'bt:Image': [
        {
          '$': {
            size: 16,
            resid: createImageResource('menuIcon', config.menuCount + '-16',
              'https://localhost:8443/images/icon-16.png', config)
          }
        },
        {
          '$': {
            size: 32,
            resid: createImageResource('menuIcon', config.menuCount + '-32',
              'https://localhost:8443/images/icon-32.png', config)
          }
        },
        {
          '$': {
            size: 80,
            resid: createImageResource('menuIcon', config.menuCount + '-80',
              'https://localhost:8443/images/icon-80.png', config)
          }
        }
      ] 
    },
    Items: {
      Item: uilessButton
    }
  };
  
  return menu;
}

/**
 * Build out a taskpane button
 */
function buildTaskPaneButton(config) {
  config.taskPaneCount++;
  
  var button = {
    '$': {
      'xsi:type': 'Button',
      id: 'taskpaneButton' + config.taskPaneCount
    },
    Label: { 
      '$': { 
        resid: createShortStringResource('taskpaneButtonLabel', config.taskPaneCount,
          'Taskpane Button ' + config.taskPaneCount, config)
      } 
    },
    Tooltip: { 
      '$': { 
        resid: createLongStringResource('taskpaneButtonToolTip', config.taskPaneCount,
          'This is the tooltip for Taskpane Button ' + config.taskPaneCount, config) 
      } 
    },
    Supertip: {
      Title: { 
        '$': { 
          resid: createShortStringResource('taskpaneButtonSuperTipTitle', config.taskPaneCount,
            'Taskpane Button ' + config.taskPaneCount, config) 
        } 
      },
      Description: { 
        '$': { 
          resid: createLongStringResource('taskpaneButtonSuperTipDesc', config.taskPaneCount,
            'This is the description for Taskpane Button ' + config.taskPaneCount, config)
        } 
      }
    },
    Icon: {
      'bt:Image': [
        {
          '$': {
            size: 16,
            resid: createImageResource('taskpaneButtonIcon', config.taskPaneCount + '-16',
              'https://localhost:8443/images/icon-16.png', config)
          }
        },
        {
          '$': {
            size: 32,
            resid: createImageResource('taskpaneButtonIcon', config.taskPaneCount + '-32',
              'https://localhost:8443/images/icon-32.png', config)
          }
        },
        {
          '$': {
            size: 80,
            resid: createImageResource('taskpaneButtonIcon', config.taskPaneCount + '-80',
              'https://localhost:8443/images/icon-80.png', config)
          }
        }
      ] 
    },
    Action: {
      '$': { 'xsi:type': 'ShowTaskpane' },
      SourceLocation: {
        '$': {
          resid: createUrlResource('taskPaneUrl', config.taskPaneCount,
            'https://localhost:8443/TaskPane/TaskPane.html', config)
        }
      }
    }
  };
  
  return button;
}

/**
 * Create URL resource
 */
function createUrlResource(prefix, suffix, value, config) {
  var resid = prefix + suffix;
  if (resid.length > 32) {
    throw "Invalid resource ID: must be 32 characters or less";
  }
  
  config.resources.urls.push({
    '$': {
      id: resid,
      DefaultValue: value
    }
  });
  
  return resid;
}

/**
 * Create image resource
 */
function createImageResource(prefix, suffix, value, config) {
  var resid = prefix + suffix;
  if (resid.length > 32) {
    throw "Invalid resource ID: must be 32 characters or less";
  }
  
  config.resources.images.push({
    '$': {
      id: resid,
      DefaultValue: value
    }
  });
  
  return resid;
}

/**
 * Create short string resource
 */
function createShortStringResource(prefix, suffix, value, config) {
  var resid = prefix + suffix;
  if (resid.length > 32) {
    throw "Invalid resource ID: must be 32 characters or less";
  }
  
  config.resources.shortStrings.push({
    '$': {
      id: resid,
      DefaultValue: value
    }
  });
  
  return resid;
}

/**
 * Create long string resource
 */
function createLongStringResource(prefix, suffix, value, config) {
  var resid = prefix + suffix;
  if (resid.length > 32) {
    throw "Invalid resource ID: must be 32 characters or less";
  }
  
  config.resources.longStrings.push({
    '$': {
      id: resid,
      DefaultValue: value
    }
  });
  
  return resid;
}