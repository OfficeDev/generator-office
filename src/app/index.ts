/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';

let opn = require('opn');
let uuid = require('uuid/v4');
let yosay = require('yosay');
let yo = require('yeoman-generator');
let insight = appInsights.getClient('68a8ef35-112c-4d33-a118-3c346947f2fe');

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.argument('name', { type: String, required: false });
    this.argument('host', { type: String, required: false });
    this.argument('framework', { type: String, required: false });

    this.option('skip-install', {
      type: Boolean,
      required: false,
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
    let jsTemplates = getDirectories(this.templatePath('js'));
    let tsTemplates = getDirectories(this.templatePath('ts'));
    let manifests = getFiles(this.templatePath('manifest')).map(manifest => manifest.replace('.xml', ''));

    let askForBasic = [
      /** whether to create a new folder for the project */
      {
        name: 'folder',
        message: 'Would you like to create a new subfolder for your project?',
        type: 'confirm',
        default: true
      },

      /** name for the project */
      {
        name: 'name',
        type: 'input',
        message: 'What do you want to name your add-in?',
        default: 'My Office Add-in',
        when: this.options.name == null
      },

      /** office client application that can host the addin */
      {
        name: 'host',
        message: 'Which Office client application would you like to support?',
        type: 'list',
        default: 'excel',
        choices: manifests.map(manifest => ({ name: _.capitalize(manifest), value: manifest })),
        when: this.options.host == null
      },

      /** set flag for manifest-only to prompt accordingly later */
      {
        name: 'isManifestOnly',
        message: 'Would you like to create a new project?',
        type: 'list',
        default: false,
        choices: [
          {
            name: 'Yes, I want a new project.',
            value: false
          },
          {
            name: 'No, I only need the manifest file.',
            value: true
          }
        ],
        when: this.options.framework == null
      }
    ];

    /**
     * Configure user input to have correct values
     */
    let answers = await this.prompt(askForBasic); // trigger prompts and store user input
    this.project = {
      folder: answers.folder,
      name: this.options.name || answers.name,
      host: this.options.host || answers.host,
      framework: this.options.framework || null,
      isManifestOnly: answers.isManifestOnly || null
    };
    if (answers.isManifestOnly) {
      this.project.framework = 'manifest-only';
    }
    if (this.options.framework === 'manifest-only') {
      this.project.isManifestOnly = true;
    }

    /** askForTs and askForFramework will only be triggered if it's not a manifest-only project */
    let askForTs = [
      /** use TypeScript for the project */
      {
        name: 'ts',
        type: 'confirm',
        message: 'Would you like to use TypeScript?',
        default: true,
        when: (this.options.js == null) && (!this.project.isManifestOnly)
      }
    ];
    let tsAnswers = await this.prompt(askForTs); // trigger prompts and store user input
    this.project.ts = tsAnswers.ts;
    if (!(this.options.js == null)) {
      this.project.ts = !this.options.js;
    }
    else {
      this.project.ts = tsAnswers.ts;
    }

    let askForFramework = [
      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: tsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
        when: (this.project.framework == null) && tsAnswers.ts && !answers.isManifestOnly
      },

      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: jsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
        when: (this.project.framework == null) && !tsAnswers.ts && !answers.isManifestOnly
      }
    ];
    let frameworkAnswers = await this.prompt(askForFramework); // trigger prompts and store user input
    if (!(this.options.framework == null)) {
      this.project.framework = this.options.framework;
    }
    else if (this.project.isManifestOnly === true) {
      this.project.framework = 'manifest-only';
    }
    else {
      this.project.framework = frameworkAnswers.framework;
    }
  },

  /**
   * save configs & config project
   */
  configuring: function () {
    this.project.projectInternalName = _.kebabCase(this.project.name);
    this.project.projectDisplayName = _.capitalize(this.project.name);
    this.project.hostDisplayName = _.capitalize(this.project.host);
    this.project.projectId = uuid();
    if (this.project.folder) {
      this.destinationRoot(this.project.projectInternalName);
    }

    insight.trackEvent('App_Data', { AppID: this.project.projectId, Host: this.project.host, Framework: this.project.framework, isTypeScript: this.project.ts });
  },

  writing: {
    copyFiles: function () {
      let language = this.project.ts ? 'ts' : 'js';

      /** Show type of project creating in progress */
      if (this.project.framework !== 'manifest-only') {
        this.log('----------------------------------------------------------------------------------\n');
        this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in using ${chalk.bold.magenta(language)} and ${chalk.bold.cyan(this.project.framework)}\n`);
        this.log('----------------------------------------------------------------------------------\n\n');
      }
      else {
        this.log('----------------------------------------------------------------------------------\n');
        this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} add-in`);
        this.log('----------------------------------------------------------------------------------\n\n');
      }

      /** Copy the manifest */
      this.fs.copyTpl(this.templatePath(`manifest/${this.project.host}.xml`), this.destinationPath(`${this.project.projectInternalName}-manifest.xml`), this.project);

      if (this.project.framework === 'manifest-only') {
        this.fs.copyTpl(this.templatePath(`manifest-only/**`), this.destinationPath(), this.project);
      }
      else {
        /** Copy the base template */
        this.fs.copy(this.templatePath(`${language}/base/**`), this.destinationPath());

        /** Copy the framework specific overrides */
        this.fs.copyTpl(this.templatePath(`${language}/${this.project.framework}/**`), this.destinationPath(), this.project);
      }
    }
  },

  install: function () {
    if (!this.options['skip-install'] && this.project.framework !== 'manifest-only') {
      this.installDependencies({
        npm: true,
        bower: false,
        callback: this._postInstallHints.bind(this)
      });
    }
    else {
      if (this.project.framework !== 'manifest-only') {
        this._postInstallHints();
      }
    }
  },

  _postInstallHints: async function () {
    /** Next steps and npm commands */
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    this.log(`      1. Launch your local web server via ${chalk.inverse(' npm start ')} (you may also need to trust`);
    this.log(`         the Self-Signed Certificate for the site, if you haven't done that before)`);
    this.log(`      2. Sideload the add-in into your Office application.\n`);
    this.log(`      We have created a resource.html file in your project for more info & resources on these topics.\n`);
    this.log('----------------------------------------------------------------------------------------------------------\n');
    let askForOpenResourcePage = [
      /** ask to open resource page */
      {
        name: 'open',
        type: 'confirm',
        message: 'Would you like to open resource.html file now?',
        default: true
      }
    ];
    let openResourcePageAnswers = await this.prompt(askForOpenResourcePage); // trigger prompts and store user input
    if (openResourcePageAnswers.open === true) {
      opn('resource.html');
    }
  }
} as any);

function getDirectories(root) {
  return fs.readdirSync(root).filter(file => {
    if (file === 'base') {
      return false;
    }
    return fs.statSync(path.join(root, file)).isDirectory();
  });
}

function getFiles(root) {
  return fs.readdirSync(root).filter(file => {
    return !(fs.statSync(path.join(root, file)).isDirectory());
  });
}
