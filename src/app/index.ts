/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';
import * as opn from 'opn';
import * as uuid from 'uuid/v4';
import * as yosay from 'yosay';
import * as yo from 'yeoman-generator';

import generateStarterCode from './config/starterCode';

let insight = appInsights.getClient('1ced6a2f-b3b2-4da5-a1b8-746512fbc840');

// Remove unwanted tags
delete insight.context.tags['ai.cloud.roleInstance'];
delete insight.context.tags['ai.device.osVersion'];
delete insight.context.tags['ai.device.osArchitecture'];
delete insight.context.tags['ai.device.osPlatform'];

const manifestOnly = 'manifest-only';

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.argument('name', { type: String, required: false });
    this.argument('host', { type: String, required: false });
    this.argument('projectType', { type: String, required: false });

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
  
    this.option('output', {
      type: String,
      required: false,
      desc: 'Project folder name if different from project name'
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
    try {
      let jsTemplates = getDirectories(this.templatePath('js'));
      jsTemplates.push(manifestOnly);
      let tsTemplates = getDirectories(this.templatePath('ts'));
      tsTemplates.push(manifestOnly);
      let manifests = getFiles(this.templatePath('manifest')).map(manifest => _.capitalize(manifest.replace('.xml', '')));
      updateHostNames(manifests, 'Onenote', 'OneNote');
      updateHostNames(manifests, 'Powerpoint', 'PowerPoint');

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
      let startForHost = (new Date()).getTime();
      let askForHost = [{
        name: 'host',
        message: 'Which Office client application would you like to support?',
        type: 'list',
        default: 'Excel',
        choices: manifests.map(manifest => ({ name: manifest, value: manifest })),
        when: this.options.host == null || (this.options.host != null && !this._isValidInput(this.options.host, manifests, true /* isHostParam */))
      }];
      let answerForHost = await this.prompt(askForHost);
      let endForHost = (new Date()).getTime();
      let durationForHost = (endForHost - startForHost) / 1000;

      /**
       * Configure user input to have correct values
       */
      this.project = {
        folder: null,
        name: this.options.name || answerForName.name,
        host: answerForHost.host || this.options.host,
        projectType: this.options.projectType || null,
        isManifestOnly: false
      };

      // Set folder to specified name if 'output' option is passed as an argument
      if (this.options.output != null) {
        this.project.folder = this.options.output;
      }
      else{
        this.project.folder = this.project.name;
      }
  
      // Set isManifestOnly flag to true if project-type argument is 'manifest-only'
      if (this.options.projectType === 'manifest-only') {
          this.project.isManifestOnly = true;
      }

      // Set js flag to true if 'js' option is passed as an argument
      if (this.options.js != null) {
        this.project.ts = !this.options.js;
      }
      else {
        this.project.ts = true;
      }
      if (this.options.projectType === 'react') {
        this.project.ts = true;
      }

      /** technology used to create the addin (jquery / angular / etc) */
      let startForProjectType = (new Date()).getTime();
      let askForProjectType = [
        {
          name: 'projectType',
          message: 'Choose a project-type:',
          type: 'list',
          default: 'react',
          choices: tsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
          when: (this.project.projectType == null || !this._isValidInput(this.options.projectType, tsTemplates, false /* isHostParam */)) && this.project.ts && !this.options.js
                && !this.project.isManifestOnly
        },
        {
          name: 'projectType',
          message: 'Choose a project-type:',
          type: 'list',
          default: 'jquery',
          choices: jsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
          when: (this.project.projectType == null || !this._isValidInput(this.options.projectType, jsTemplates, false /* isHostParam */)) && !this.project.ts && this.options.js
                && !this.project.isManifestOnly
        }
      ];
      let answerForProjectType = await this.prompt(askForProjectType);
      let endForProjectType = (new Date()).getTime();
      let durationForProjectType = (endForProjectType - startForProjectType) / 1000;
      this.project.projectType = answerForProjectType.projectType || this.options.projectType;

      /** appInsights logging */
      const noElapsedTime = 0;
      insight.trackEvent('Name', { Name: this.project.name }, { durationForName });
      insight.trackEvent('Folder', { CreatedSubFolder: this.project.folder.toString() }, { noElapsedTime }); 
      insight.trackEvent('Host', { Host: this.project.host }, { durationForHost });    
      insight.trackEvent('IsTs', { IsTs: this.project.ts.toString() }, { noElapsedTime });      
      insight.trackEvent('IsManifestOnly', { IsManifestOnly: this.project.isManifestOnly.toString() }, { noElapsedTime });
      insight.trackEvent('ProjectType', { ProjectType: this.project.projectType }, { durationForProjectType }); 
    } catch (err) {
      insight.trackException(new Error('Prompting Error: ' + err));
    }
  },

  /**
   * save configs & config project
   */
  configuring: function () {
    try {
      this.project.projectInternalName = _.kebabCase(this.project.name);
      this.project.projectDisplayName = this.project.name;
      this.project.projectId = uuid();
      this.project.hostInternalName = _.toLower(this.project.host);
      this.destinationRoot(this.project.folder);

      let duration = this.project.duration;
      insight.trackEvent('App_Data', { AppID: this.project.projectId, Host: this.project.host, ProjectType: this.project.projectType, isTypeScript: this.project.ts.toString() }, { duration });
    } catch (err) {
      insight.trackException(new Error('Configuration Error: ' + err));
    }
  },

  writing: {
    copyFiles: function () {
      try {
        let language = this.project.ts ? 'ts' : 'js';

        /** Show type of project creating in progress */
        if (this.project.projectType !== 'manifest-only') {
          this.log('\n----------------------------------------------------------------------------------\n');
          this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in for ${chalk.bold.yellow(this.project.host)} using ${chalk.bold.magenta(language)} and ${chalk.bold.cyan(this.project.projectType)}\n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }
        else {
          this.log('----------------------------------------------------------------------------------\n');
          this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} add-in\n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }

        const starterCode = generateStarterCode(this.project.host);
        const templateFills = Object.assign({}, this.project, starterCode);

        /** Copy the manifest */
        this.fs.copyTpl(this.templatePath(`manifest/${this.project.hostInternalName}.xml`), this.destinationPath(`${this.project.projectInternalName}-manifest.xml`), templateFills);

        if (this.project.projectType === 'manifest-only') {
          this.fs.copyTpl(this.templatePath(`manifest-only/**`), this.destinationPath(), templateFills);
        }
        else {
          /** Copy the base template */
          this.fs.copy(this.templatePath(`${language}/base/**`), this.destinationPath(), { globOptions: { ignore: `**/*.placeholder` }});

          /** Copy the project-type specific overrides */
          this.fs.copyTpl(this.templatePath(`${language}/${this.project.projectType}/**`), this.destinationPath(), templateFills, null, { globOptions: { ignore: `**/*.placeholder` }});
          
          /** Manually copy any dot files as yoeman can't handle them */

          /** .babelrc */
          const babelrcPath = this.templatePath(`${language}/${this.project.projectType}/babelrc.placeholder`);
          if (this.fs.exists(babelrcPath)) {
              this.fs.copy(babelrcPath, this.destinationPath('.babelrc'));
          }

          /** .gitignore */
          const gitignorePath = this.templatePath(`${language}/base/gitignore.placeholder`);
          if (this.fs.exists(gitignorePath)) {
              this.fs.copy(gitignorePath, this.destinationPath('.gitignore'));
          }
        }
      } catch (err) {
        insight.trackException(new Error('File Copy Error: ' + err));
      }
    }
  },

  install: function () {
    try {      
      if (this.options['skip-install']) {
        this.installDependencies({
          npm: false,
          bower: false,
          callback: this._postInstallHints.bind(this)
        });
      }
      else {
        this.installDependencies({
          npm: true,
          bower: false,
          callback: this._postInstallHints.bind(this)
        });    
      }
    } catch (err) {
      insight.trackException(new Error('Installation Error: ' + err));
      process.exitCode = 1;
    }
  },

  _postInstallHints: function () {
    /** Next steps and npm commands */
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    this.log(`      1. Launch your local web server via ${chalk.inverse(' npm start ')} (you may also need to`);
    this.log(`         trust the Self-Signed Certificate for the site if you haven't done that)`);
    this.log(`      2. Sideload the add-in into your Office application.\n`);
    this.log(`      Please refer to resource.html in your project for more information.`);
    this.log(`      Or visit our repo at: https://github.com/officeDev/generator-office \n`);
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this._exitProcess();
  },

  _isValidInput: function(input, inputArray, isHostParam) 
  {
    // validate host and project-type inputs
    for (var i = 0; i < inputArray.length; i++)
    {
      var element = inputArray[i];
      if (input.toLowerCase() == element.toLowerCase())
      {
        if (isHostParam)
        {
          this.options.host = element;
        }
        else
        {
          this.options['project-type'] = element;
        }
        return true;
      }
    }
    return false;
  },

  _exitProcess: function () {
    process.exit();
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

function updateHostNames(arr, key, newval) {
  let match = _.some(arr, _.method('match', key));
  if (match) {
    let index = _.indexOf(arr, key);
    arr.splice(index, 1, newval);
  }
}