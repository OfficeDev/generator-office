'use strict'

const guid = require('uuid');
const yo = require('yeoman-generator');

import appInsight = require('applicationinsights');
import chalk = require('chalk');
import yosay = require('yosay');
import ncp = require('ncp');
import * as path from 'path';

var insight = appInsight.getClient();

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.option('skip-install', {
      type: Boolean,
      required: false,
      defaults: false,
      desc: 'Skip running package managers (NPM, bower, etc) post scaffolding'
    });

    this.option('name', {
      type: String,
      desc: 'Title of the Office Add-in',
      required: false
    });

    this.option('root-path', {
      type: String,
      desc: 'Relative path where the Add-in should be created (blank = current directory)',
      required: false
    });

    this.option('tech', {
      type: String,
      desc: 'Technology to use for the Add-in (html = HTML; ng = Angular)',
      required: false
    });

    this.option('is-project-new', {
      type: String,
      desc: 'To create a new project or update exisiting project',
      required: false
    });

    this.option('host', {
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
  initializing: function () {
    this.log(yosay('Welcome to the ' +
      chalk.red('Office Project') +
      ' generator, by ' +
      chalk.red('@OfficeDev') +
      '! Let\'s create a project together!'));

    // create global config object on this generator
    this.genConfig = {};
  }, // initializing()

  /**
   * Prompt users for options
   */
  prompting: async function () {
    let prompts = [

      // allow customer to create new project or update existing project
      {
        name: 'is-project-new',
        message: 'Create new project or update existing project:',
        type: 'list',
        default: 'New project',
        choices: [
          {
            name: 'Create new project',
            value: 'new'
          },
          {
            name: 'Update existing project',
            value: 'existing'
          }],
        when: this.options['is-project-new'] === undefined
      },

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
        filter: function (response) {
          if (response === 'current folder') {
            return '.';
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
            name: 'Angular ADAL',
            value: 'ng-adal'
          }, {
            name: 'Manifest.xml only (no application source files)',
            value: 'manifest-only'
          }]
      },

      // office client application that can host the addin
      {
        name: 'host',
        message: 'Supported Office application:',
        type: 'list',
        choices: [
          {
            name: 'Mail',
            value: 'mail'
          },
          {
            name: 'Word',
            value: 'document'
          },
          {
            name: 'Excel',
            value: 'workbook'
          },
          {
            name: 'PowerPoint',
            value: 'presentation'
          },
          {
            name: 'OneNote',
            value: 'notebook'
          },
          {
            name: 'Project',
            value: 'project'
          }
        ],
        when: this.options.host === undefined
      }];
    
    insight.trackTrace('User begins to choose options');
    var start = (new Date()).getTime();

    // trigger prompts and store user input
    await this.prompt(prompts).then(function (responses) {
      var end = (new Date()).getTime();
      var duration = (end - start)/1000;
      insight.trackEvent('WHYME', { Project_Type: this.genConfig.type }, { duration });

      this.genConfig = {
        name: responses.name,
        tech: responses.tech,
        'is-project-new': responses['is-project-new'],
        'root-path': responses['root-path'],
        host: responses.host
      };
    }.bind(this));
  },

  /**
   * save configurations & config project
   */
  configuring: function () {
    // take name submitted and strip everything out non-alphanumeric or space
    var projectName = this.genConfig.name;
    projectName = projectName.replace(/[^\w\s\-]/g, '');
    projectName = projectName.replace(/\s{2,}/g, ' ');
    projectName = projectName.trim();

    // add the result of the question to the generator configuration object
    this.genConfig.projectInternalName = projectName.toLowerCase().replace(/ /g, '-');
    this.genConfig.projectDisplayName = projectName;
    this.genConfig.rootPath = this.genConfig['root-path'];
    this.genConfig.isProjectNew = this.genConfig['is-project-new'];

    this.genConfig.projectId = guid.v4();
  }, // configuring()

  writing: {
    copyFiles: function () {
      /**
       * Output files
       */
      var manifestFilename = 'manifest-' + this.genConfig.host + '.xml';

      if (this.genConfig.isProjectNew === 'new')
      {
        ncp.ncp(this.templatePath('common-static'), this.destinationPath(), err => console.log(err));
        this.fs.copyTpl(this.templatePath('common-dynamic/package.json'), 
              this.destinationPath('package.json'),
              this.genConfig);
        this.fs.copyTpl(this.templatePath('manifest/' + manifestFilename), 
          this.destinationPath(manifestFilename),
          this.genConfig);

        switch (this.genConfig.tech) {
          case 'html':
            ncp.ncp(this.templatePath('tech/html'), this.destinationPath(), err => console.log(err));
            break;
          case 'ng':
            ncp.ncp(this.templatePath('tech/ng'), this.destinationPath(), err => console.log(err));
            break;
        };
      }
    }
  },

  install: function () {
    if (!this.options['skip-install'] && this.genConfig.tech !== 'manifest-only') {
      this.npmInstall();
    }
  }
} as any);
