'use strict';

var generators = require('yeoman-generator');
var chalk = require('chalk');
var yosay = require('yosay');
var extend = require('deep-extend');

module.exports = generators.Base.extend({
  constructor: function(){

    generators.Base.apply(this, arguments);

    this.option('skip-install', {
      type: Boolean,
      desc: 'Skip running package managers (NPM, bower, etc) post scaffolding',
      required: false,
      defaults: false
    });

    this.option('name', {
      type: String,
      desc: 'Title of the Office Project',
      required: false
    });

    this.option('root-path', {
      type: String,
      desc: 'Relative path where the project should be created (blank = current directory)',
      required: false
    });

    this.option('tech', {
      type: String,
      desc: 'Technology to use for the project (html = HTML; ng = Angular)',
      required: false
    });

    this.option('clients', {
      type: String,
      desc: 'Office client product that can host the add-in',
      required: false
    });
    
    this.option('extensionPoint', {
      type: String,
      desc: 'Supported extension points',
      required: false
    });
    
    this.option('appId', {
      type: String,
      desc: 'Application ID as registered in Azure AD',
      required: false
    });

  }, // constructor()

  /**
   * Generator initalization
   */
  initializing: function(){
    this.log(yosay('Welcome to the ' +
      chalk.red('Office Project') +
      ' generator, by ' +
      chalk.red('@OfficeDev') +
      '! Let\'s create a project together!'));

    // generator configuration
    this.genConfig = {};
  }, // initializing()

  /**
   * Prompt users for options
   */
  prompting: {

    askFor: function(){
      var done = this.async();

      var prompts = [
        // friendly name of the generator
        {
          name: 'name',
          message: 'Project name (display name):',
          default: 'My Office Project',
          when: this.options.name === undefined
        },
        // root path where the addin should be created; should go in current folder where
        //  generator is being executed, or within a subfolder?
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
        // type of project - this will dictate which subgenerator to call
        {
          name: 'type',
          message: 'Office project type:',
          type: 'list',
          choices: [
            {
              name: 'Mail Add-in (read & compose forms)',
              value: 'mail'
            },
            {
              name: 'Task Pane Add-in',
              value: 'taskpane'
            },

            {
              name: 'Content Add-in',
              value: 'content'
            }]
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
              name: 'Angular ADAL',
              value: 'ng-adal'
            }, {
              name: 'Manifest.xml only (no application source files)',
              value: 'manifest-only'
            }]
        }];

      // trigger prompts
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, this.options);
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));

    }, // askFor()
    
    askForAdalConfig: function(){
      // if it's not an ADAL app, don't ask the questions
      if (this.genConfig.tech !== 'ng-adal') {
        return;
      }

      var done = this.async();

      // office client application that can host the addin
      var prompts = [{
        name: 'appId',
        message: 'Application ID as registered in Azure AD:',
        default: '00000000-0000-0000-0000-000000000000',
        when: this.options.appId === undefined
      }];

      // trigger prompts
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));

    }, // askForAdalConfig()

    askForOfficeClients: function(){
      // if it's a mail addin, don't ask for Office client
      if (this.genConfig.type === 'mail') {
        return;
      }

      var done = this.async();

      // office client application that can host the addin
      var prompts = [{
        name: 'clients',
        message: 'Supported Office applications:',
        type: 'checkbox',
        choices: [
          {
            name: 'Word',
            value: 'Document',
            checked: true
          },
          {
            name: 'Excel',
            value: 'Workbook',
            checked: true
          },
          {
            name: 'PowerPoint',
            value: 'Presentation',
            checked: true
          },
          {
            name: 'Project',
            value: 'Project',
            checked: true
          }
        ],
        when: this.options.clients === undefined,
        validate: /* istanbul ignore next */ function(clientsAnswer){
          if (clientsAnswer.length < 1) {
            return 'Must select at least one Office application';
          }
          return true;
        }
      }];

      // trigger prompts
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));

    } // askForOfficeClients()

  }, // prompting()

  default: function(){

    // determine which subgenerator to call
    switch (this.genConfig.type) {
      // Mail Office Add-in
      case 'mail':
        // execute subgenerator
        this.composeWith('office:mail', {
          options: {
            name: this.genConfig.name,
            'root-path': this.genConfig['root-path'],
            tech: this.genConfig.tech,
            outlookForm: this.genConfig.outlookForm,
            extensionPoint: this.genConfig.extensionPoint,
            appId: this.genConfig.appId,
            'skip-install': this.options['skip-install']
          }
        }, {
            local: require.resolve('../mail')
          });
        break;

      // Taskpane Office Add-in
      case 'taskpane':
        // execute subgenerator
        this.composeWith('office:taskpane', {
          options: {
            name: this.genConfig.name,
            'root-path': this.genConfig['root-path'],
            tech: this.genConfig.tech,
            appId: this.genConfig.appId,            
            clients: this.genConfig.clients,
            'skip-install': this.options['skip-install']
          }
        }, {
            local: require.resolve('../taskpane')
          });
        break;
      // Content Office Add-in
      case 'content':
        // execute subgenerator
        this.composeWith('office:content', {
          options: {
            name: this.genConfig.name,
            'root-path': this.genConfig['root-path'],
            tech: this.genConfig.tech,
            appId: this.genConfig.appId,
            clients: this.genConfig.clients,
            'skip-install': this.options['skip-install']
          }
        }, {
            local: require.resolve('../content')
          });
        break;
    }
  }, // default()

  /**
   * write generator specific files
   */
  // writing: { },

  /**
   * conflict resolution
   */
  // conflicts: { },

  /**
   * run installations (bower, npm, tsd, etc)
   */
  // install: { },

  /**
   * last cleanup, goodbye, etc
   */
  // end: { }
});
