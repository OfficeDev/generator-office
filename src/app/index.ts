import * as fs from 'fs';
import * as path from 'path';
let uuid = require('uuid/v4');
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';
let yosay = require('yosay');
let yo = require('yeoman-generator');
let opn = require('opn');
let insight = appInsights.getClient('1fd62c46-f0ef-4cfb-9560-448c857ab690');

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.argument('host', { type: String, required: false });
    this.argument('name', { type: String, required: false });

    this.option('skip-install', {
      type: Boolean,
      required: false,
      defaults: false,
      desc: 'Skip running `npm install` post scaffolding.'
    });

    this.option('js', {
      type: Boolean,
      required: false,
      desc: 'Use JavaScript templates instead of TypeScript.'
    });
  },

  /**
   * Generator initalization
   */
  initializing: function () {
    let message = `Welcome to the ${chalk.bold.green('Office Add-in')} generator, by ${chalk.bold.green('@OfficeDev')}! Let\'s create a project together!`;
    this.log(yosay(message));
    this.project = {};
  },

  /**
   * Prompt users for options
   */
  prompting: async function () {
    let prompts = [
      /** allow user to create new project or update existing project */
      {
        name: 'new',
        message: 'Would you like to create a new add-in?',
        type: 'confirm',
        default: 'true',
        when: (this.options.name == null)
      },

      /** name for the project */
      {
        name: 'name',
        type: 'input',
        message: 'Name of your add-in:',
        default: 'My Office Add-in',
        when: (this.options.name == null)
      },

      /**
       * root path where the addin should be created.
       * should go in current folder where generator is being executed,
       * or within a subfolder?
       */
      {
        name: 'folder',
        message: `Create a new folder?`,
        type: 'confirm',
        default: 'true',
        when: (this.options.name == null)
      },

      /** use TypeScript for the project */
      {
        name: 'ts',
        type: 'confirm',
        message: 'Would you like to use TypeScript?',
        default: true,
        when: (this.options.name == null)
      },

      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Choose a framework:',
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
            name: 'Manifest only (no application source files)',
            value: 'manifest-only'
          }
        ],
        when: (this.options.name == null)
      },

      /** office client application that can host the addin */
      {
        name: 'host',
        message: 'Create the add-in for:',
        type: 'list',
        default: 'workbook',
        choices: [
          {
            name: 'Excel',
            value: 'workbook'
          },
          {
            name: 'Word',
            value: 'document'
          },
          {
            name: 'PowerPoint',
            value: 'presentation'
          },
          {
            name: 'Mail',
            value: 'mail'
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
        when: (this.options.host == null)
      }
    ];

    insight.trackTrace('User begins to choose options');
    let start = (new Date()).getTime();

    // trigger prompts and store user input
    let answers = await this.prompt(prompts);

    let end = (new Date()).getTime();
    let duration = (end - start) / 1000;
    insight.trackEvent('WHYME', { Project_Type: this.project.type }, { duration });

    this.project = {
      name: answers.name || this.options.name,
      framework: answers.framework || 'jquery',
      ts: answers.ts || this.options.js || true,
      new: answers.new || true,
      folder: answers.folder || true,
      host: answers.host || this.options.host
    };
  },

  /**
   * save configs & config project
   */
  configuring: function () {
    this.project.projectInternalName = _.kebabCase(this.project.name);
    this.project.projectDisplayName = this.project.name;
    this.project.isNew = this.project.new;
    this.project.projectId = uuid();
    if (this.project.folder) {
      this.destinationRoot(this.project.projectInternalName);
    }
  },

  writing: {
    copyFiles: function () {
      let manifestFilename = 'manifest-' + this.project.host + '.xml';
      let language = this.project.ts ? 'ts' : 'js';

      if (this.project.isNew === true) {
        /** Copy the base template */
        this.fs.copy(this.templatePath(`${language}/base/**`), this.destinationPath());

        /** Copy the framework specific overrides */
        this.fs.copyTpl(this.templatePath(`${language}/${this.project.framework}/**`), this.destinationPath(), this.project);

        /** Copy the manifest */
        this.fs.copyTpl(this.templatePath('manifest/' + manifestFilename), this.destinationPath(manifestFilename), this.project);
      }
    }
  },

  install: function () {
    this.spawnCommand('project_readme.html');
    // opn(this.destinationPath(`${this.project.path}project_readme.html`));
    if (!this.options['skip-install'] && this.project.framework !== 'manifest-only') {
      this.npmInstall();
    }
  }
} as any);
