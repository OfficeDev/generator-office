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
    let manifests = getFiles(this.templatePath('manifest')).map(manifest => _.capitalize(manifest.replace('.xml', '')));
    updateHostNames(manifests, 'Onenote', 'OneNote');
    updateHostNames(manifests, 'Powerpoint', 'PowerPoint');

    /** begin prompting */
    /** whether to create a new folder for the project */
    let startForFolder = (new Date()).getTime();
    let askForFolder = [{
      name: 'folder',
      message: 'Would you like to create a new subfolder for your project?',
      type: 'confirm',
      default: false
    }];
    let answerForFolder = await this.prompt(askForFolder);
    let endForFolder = (new Date()).getTime();
    let durationForFolder = (endForFolder - startForFolder) / 1000;

    /** name for the project */
    let startForName = (new Date()).getTime();
    let askForName = [{
      name: 'name',
      type: 'input',
      message: 'What do you want to name your add-in?',
      default: 'My Office Add-in',
      when: this.options.name == null
    }];
    let answerForName = await this.prompt(askForName);
    let endForName = (new Date()).getTime();
    let durationForName = (endForName - startForName) / 1000;

    /** office client application that can host the addin */
    let startForHost =  (new Date()).getTime();
    let askForHost = [{
      name: 'host',
      message: 'Which Office client application would you like to support?',
      type: 'list',
      default: 'Excel',
      choices: manifests.map(manifest => ({ name: manifest, value: manifest })),
      when: this.options.host == null
    }];
    let answerForHost = await this.prompt(askForHost);
    let endForHost = (new Date()).getTime();
    let durationForHost = (endForHost - startForHost) / 1000;

    /** set flag for manifest-only to prompt accordingly later */
    let startForManifestOnly = (new Date()).getTime();
    let askForManifestOnly = [{
      name: 'isManifestOnly',
      message: 'Would you like to create a new add-in?',
      type: 'list',
      default: false,
      choices: [
        {
          name: 'Yes, I want a new add-in.',
          value: false
        },
        {
          name: 'No, I only need the manifest file.',
          value: true
        }
      ],
      when: this.options.framework == null
    }];
    let answerForManifestOnly = await this.prompt(askForManifestOnly); // trigger prompts and store user input
    let endForManifestOnly = (new Date()).getTime();
    let durationForManifestOnly = (endForManifestOnly - startForManifestOnly) / 1000;

    /**
     * Configure user input to have correct values
     */
    this.project = {
      folder: answerForFolder.folder,
      name: this.options.name || answerForName.name,
      host: this.options.host || answerForHost.host,
      framework: this.options.framework || null,
      isManifestOnly: answerForManifestOnly.isManifestOnly
    };
    if (answerForManifestOnly.isManifestOnly) {
      this.project.framework = 'manifest-only';
    }
    if (this.options.framework != null) {
      if (this.options.framework === 'manifest-only') {
        this.project.isManifestOnly = true;
      } else {
        this.project.isManifestOnly = false;
      }
    }

    /** askForTs and askForFramework will only be triggered if it's not a manifest-only project */
    /** use TypeScript for the project */
    let startForTs = (new Date()).getTime();
    let askForTs = [
      {
        name: 'ts',
        type: 'confirm',
        message: 'Would you like to use TypeScript?',
        default: true,
        when: (this.options.js == null) && (!this.project.isManifestOnly)
      }
    ];
    let answerForTs = await this.prompt(askForTs);
    let endForTs = (new Date()).getTime();
    let durationForTs = (endForTs - startForTs) / 1000;

    this.project.ts = answerForTs.ts;
    if (!(this.options.js == null)) {
      this.project.ts = !this.options.js;
    }
    else {
      this.project.ts = answerForTs.ts;
    }

    /** technology used to create the addin (html / angular / etc) */
    let startForFramework = (new Date()).getTime();
    let askForFramework = [
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: tsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
        when: (this.project.framework == null) && answerForTs.ts && !answerForManifestOnly.isManifestOnly
      },
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: jsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
        when: (this.project.framework == null) && !answerForTs.ts && !answerForManifestOnly.isManifestOnly
      }
    ];
    let answerForFramework = await this.prompt(askForFramework);
    let endForFramework = (new Date()).getTime();
    let durationForFramework = (endForFramework - startForFramework) / 1000;

    if (!(this.options.framework == null)) {
      this.project.framework = this.options.framework;
    }
    else if (this.project.isManifestOnly === true) {
      this.project.framework = 'manifest-only';
    }
    else {
      this.project.framework = answerForFramework.framework;
    }

    let startForResourcePage = (new Date()).getTime();
    this.log('\nFor more information and resources on your next steps, we have created a resource.html file in your project.');
    let askForOpenResourcePage = [
      /** ask to open resource page */
      {
        name: 'open',
        type: 'confirm',
        message: 'Would you like to open it now while we finish creating your project?',
        default: true
      }
    ];
    let answerForOpenResourcePage = await this.prompt(askForOpenResourcePage);
    let endForResourcePage = (new Date()).getTime();
    let durationForResourcePage = (endForResourcePage - startForResourcePage) / 1000;
    this.project.isResourcePageOpened = answerForOpenResourcePage.open;

    /** appInsights logging */
    insight.trackEvent('Folder', { CreatedSubFolder: this.project.folder.toString() }, { durationForFolder });
    insight.trackEvent('Name', { Name: this.project.name }, { durationForName });
    insight.trackEvent('Host', { Host: this.project.host }, { durationForHost });
    insight.trackEvent('IsManifestOnly', { IsManifestOnly: this.project.isManifestOnly.toString() }, { durationForManifestOnly });
    insight.trackEvent('IsResourcePageOpened', { IsResourcePageOpened: this.project.isResourcePageOpened.toString() }, { durationForResourcePage });

    if (this.project.isManifestOnly === false) {
      insight.trackEvent('IsTs', { IsTs: this.project.ts.toString() }, { durationForTs });
      insight.trackEvent('Framework', { Framework: this.project.framework }, { durationForFramework });
    }
},

  /**
   * save configs & config project
   */
  configuring: function () {
    this.project.projectInternalName = _.kebabCase(this.project.name);
    this.project.projectDisplayName = _.capitalize(this.project.name);
    this.project.projectId = uuid();
    this.project.hostInternalName = _.toLower(this.project.host).replace(' ', '');

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
        this.log('\n----------------------------------------------------------------------------------\n');
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
    if (this.project.isResourcePageOpened) {
      opn(`resource.html`);
    }
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

  _postInstallHints: function () {
    /** Next steps and npm commands */
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    this.log(`      1. Launch your local web server via ${chalk.inverse(' npm start ')} (you may also need to trust`);
    this.log(`         the Self-Signed Certificate for the site, if you haven't done that before)`);
    this.log(`      2. Sideload the add-in into your Office application.\n`);
    this.log('----------------------------------------------------------------------------------------------------------\n');
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

function updateHostNames (arr, key, newval) {
  let match = _.some(arr, _.method('match', key));
  if (match) {
    let index = _.indexOf(arr, key);
    arr.splice(index, 1, newval);
  }
}
