let cpx = require('cpx');
import * as fs from 'fs';
import * as path from 'path';
let uuid = require('uuid');
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';
let yosay = require('yosay');
let yo = require('yeoman-generator');

let insight = appInsights.getClient('1fd62c46-f0ef-4cfb-9560-448c857ab690');

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
  },

  /**
   * Generator initalization
   */
  initializing: function () {
    let message = `Welcome to the ${chalk.red('Office Project')} generator, by ${chalk.red('@OfficeDev')}! Let\'s create a project together!`;
    this.log(yosay(message));
    this.genConfig = {};
  },

  /**
   * Prompt users for options
   */
  prompting: async function () {
    let prompts = [
      /** allow user to create new project or update existing project */
      {
        name: 'is-project-new',
        message: 'Create new Add-in or update existing Add-in:',
        type: 'list',
        default: 'new',
        choices: [
          {
            name: 'Create new Add-in',
            value: 'new'
          },
          {
            name: 'Update existing Add-in',
            value: 'existing'
          }
        ]
      },

      /** name for the project */
      {
        name: 'name',
        type: 'input',
        message: 'Name of the Add-in',
        default: 'My Office Add-in'
      },

      /**
       * root path where the addin should be created.
       * should go in current folder where generator is being executed,
       * or within a subfolder?
       */
      {
        name: 'root-path',
        message: `Root folder of project? Default to current directory\n (${this.destinationRoot()}), or specify relative path\n from current (src / public):`,
        default: 'current folder',
        filter: response => {
          if (response === 'current folder') {
            return '.';
          }
          else {
            return response;
          }
        }
      },

      /** use TypeScript for the project */
      {
        name: 'ts',
        type: 'confirm',
        message: 'Would you like to use TypeScript',
        default: false
      },

      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Framework to use',
        type: 'list',
        default: 'jquery',
        choices: [
          {
            name: 'jQuery',
            value: 'jquery'
          },
          {
            name: 'Angular',
            value: 'angular'
          },
          {
            name: 'Angular + ADAL',
            value: 'angular-adal'
          },
          {
            name: 'Manifest.xml only (no application source files)',
            value: 'manifest-only'
          }
        ]
      },

      /** office client application that can host the addin */
      {
        name: 'host',
        message: 'Create the add-in for',
        type: 'list',
        default: 'excel',
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
        ]
      }
    ];

    insight.trackTrace('User begins to choose options');
    let start = (new Date()).getTime();

    // trigger prompts and store user input
    let answers = await this.prompt(prompts);

    let end = (new Date()).getTime();
    let duration = (end - start) / 1000;
    insight.trackEvent('WHYME', { Project_Type: this.genConfig.type }, { duration });

    this.genConfig = {
      name: answers.name,
      framework: answers.framework,
      ts: answers.ts,
      'is-project-new': answers['is-project-new'],
      'root-path': answers['root-path'],
      host: answers.host
    };
  },

  /**
   * save configurations & config project
   */
  configuring: function () {
    // take name submitted and strip everything out non-alphanumeric or space
    let projectName = _.kebabCase(this.genConfig.name);

    // add the result of the question to the generator configuration object
    this.genConfig.projectInternalName = projectName;
    this.genConfig.projectDisplayName = this.genConfig.name;
    this.genConfig.rootPath = this.genConfig['root-path'];
    this.genConfig.isProjectNew = this.genConfig['is-project-new'];
    this.genConfig.projectId = uuid.v4();
  },

  writing: {
    copyFiles: function () {
      let manifestFilename = 'manifest-' + this.genConfig.host + '.xml';
      let folder = this.genConfig.ts ? 'ts' : 'js';

      if (this.genConfig.isProjectNew === 'new') {
        cpx.copy(this.templatePath(`${folder}/base/**`), this.destinationPath());
        this.fs.copyTpl(this.templatePath('manifest/' + manifestFilename), this.destinationPath(manifestFilename), this.genConfig);

        switch (this.genConfig.framework) {
          case 'jquery':
            this.fs.copyTpl(this.templatePath(`${folder}/jquery/**/*`), this.destinationPath(), this.genConfig);
            break;

          case 'angular':
            this._recurrsiveCopy(this.templatePath(`${folder}/angular/**/*`), this.destinationPath(), this.genConfig);
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
