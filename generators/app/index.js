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
      default: '.',
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
        }];

      // trigger prompts
      this.prompt(prompts, function(responses){
        this.genConfig = extend(this.genConfig, this.options);
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));
    }

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
