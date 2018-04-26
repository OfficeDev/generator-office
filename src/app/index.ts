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
import { log } from 'util';

let insight = appInsights.getClient('1ced6a2f-b3b2-4da5-a1b8-746512fbc840');

// Remove unwanted tags
delete insight.context.tags['ai.cloud.roleInstance'];
delete insight.context.tags['ai.device.osVersion'];
delete insight.context.tags['ai.device.osArchitecture'];
delete insight.context.tags['ai.device.osPlatform'];

const manifest = 'manifest';

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.argument('projectType', { type: String, required: false }); 
    this.argument('name', { type: String, required: false }); 
    this.argument('host', { type: String, required: false });

    this.option('skip-install', {
      type: Boolean,
      required: false,
      desc: 'Skip running `npm install` post scaffolding.'
    });

    this.option('js', {
      type: Boolean,
      required: false,
      desc: 'Project uses JavaScript instead of TypeScript.'
    });

    this.option('ts', {
      type: Boolean,
      required: false,
      desc: 'Project uses TypeScript instead of JavaScript.'
    });
  
    this.option('output', {
      alias: 'o',
      type: String,
      required: false,
      desc: 'Project folder name if different from project name.'
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
      jsTemplates.push(`Manifest`);
      let tsTemplates = getDirectories(this.templatePath('ts'));
      tsTemplates.push(`Manifest`);
      let manifests = getFiles(this.templatePath('manifest')).map(manifest => (manifest.replace('.xml', '')));
      let isManifestProject = false;
      updateHostNames(manifests, 'Onenote', 'OneNote');
      updateHostNames(manifests, 'Powerpoint', 'PowerPoint');

      // Set isManifestProject to true if manifest project type passed as argument
      if (this.options.project != null && this.options.project.toLowerCase() == manifest){
        isManifestProject = true;
      }

      /** askForTs and askForProjectType will only be triggered if it's not a manifest-only project */
      let startForScriptType = (new Date()).getTime();
      let askForScriptType = [
        {
          name: 'scriptType',
          type: 'list',
          message: 'Choose a script type',
          choices: ['Typescript', 'Javascript'],
          default: 'Typescript',
          when: this.options.js == null && this.options.ts == null && this.options.projectType != manifest && !isManifestProject
        }
      ];
      let answerForScriptType = await this.prompt(askForScriptType);
      let endForScriptType = (new Date()).getTime();
      let durationForScriptType = (endForScriptType - startForScriptType) / 1000;

      /** Project type for the add-in (jquery / angular / react / manifest) */
      let startForProjectType = (new Date()).getTime();
      let askForProjectType = [
        {
          name: 'projectType',
          message: 'Choose a project type:',
          type: 'list',
          default: 'React',
          choices: tsTemplates.map(template => ({ name: template, value: template })),
          when: (this.options.projectType == null || !this._isValidInput(this.options.projectType, tsTemplates, false /* isHostParam */))
          && (this.options.ts != null || answerForScriptType.scriptType == 'Typescript') && !isManifestProject
        },
        {
          name: 'projectType',
          message: 'Choose a project type:',
          type: 'list',
          default: 'Jquery',
          choices: jsTemplates.map(template => ({ name: template, value: template })),
          when: (this.options.projectType == null || !this._isValidInput(this.options.projectType, jsTemplates, false /* isHostParam */))
          && (this.options.js != null || answerForScriptType.scriptType == 'Javascript') && !isManifestProject
        }
      ];
      let answerForProjectType = await this.prompt(askForProjectType);
      let endForProjectType = (new Date()).getTime();
      let durationForProjectType = (endForProjectType - startForProjectType) / 1000;
      
      if ((this.options.projectType != null && this.options.projectType.toLowerCase() == manifest) || (answerForProjectType.projectType != null
        && answerForProjectType.projectType.toLowerCase() == manifest)){ 
          isManifestProject = true; }

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
        when: this.options.host == null || this.options.host != null && !this._isValidInput(this.options.host, manifests, true /* isHostParam */)
      }];
      let answerForHost = await this.prompt(askForHost);
      let endForHost = (new Date()).getTime();
      let durationForHost = (endForHost - startForHost) / 1000;

       /**
       * Configure project properties based on user input or answers to prompts
       */
      this.project = {
        folder: this.options.name || answerForName.name || this.options.output,
        name: this.options.name || answerForName.name,
        host: this.options.host || answerForHost.host,
        projectType: this.options.projectType || answerForProjectType.projectType,
        isManifestOnly: isManifestProject,
        scriptType: answerForScriptType.scriptType
      };

      // Configure project properties based on any user options specified
      if (this.options.ts){
        this.project.scriptType = 'Typescript'; 
      }

      if (this.options.js){
        this.project.scriptType = 'Javascript';
      }

      // Ensure script type is set to Typescript if the project type is react
      if (this.project.projectType.toLowerCase() === 'react') {
        this.project.scriptType = 'Typescript';
      }

      if (this.options.output != null){
        this.project.folder = this.options.output;
      }
  
      /** appInsights logging */
      const noElapsedTime = 0;
      insight.trackEvent('Name', { Name: this.project.name }, { durationForName });
      insight.trackEvent('Folder', { CreatedSubFolder: this.project.folder.toString() }, { noElapsedTime }); 
      insight.trackEvent('Host', { Host: this.project.host }, { durationForHost });    
      insight.trackEvent('ScriptType', { ScriptType: this.project.scriptType }, { noElapsedTime });      
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

      // Check to to see if destination folder already exists. If so, we will exit and prompt the user to provide
      // a different project name or output folder
      // this._projectFolderExists();

      let duration = this.project.duration;
      insight.trackEvent('App_Data', { AppID: this.project.projectId, Host: this.project.host, ProjectType: this.project.projectType/* , isTypeScript: this.project.scriptType = 'Typescript' */ }, { duration });
    } catch (err) {
      insight.trackException(new Error('Configuration Error: ' + err));
    }
  },

  writing: {
    copyFiles: function () {
      try {
        let language = this.project.scriptType === 'Typescript' || this.options.ts ? 'ts' : 'js';

        /** Show type of project creating in progress */
        if (this.project.projectType.toLowerCase() !== manifest) {
          this.log('\n----------------------------------------------------------------------------------\n');
          this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in for ${chalk.bold.yellow(this.project.host)} using ${chalk.bold.magenta(language)} and ${chalk.bold.cyan(this.project.projectType)} in folder:${chalk.bold.green(this.project.folder)}\n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }
        else {
          this.log('----------------------------------------------------------------------------------\n');
          this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} add-in in folder: ${chalk.bold.magenta(this.project.folder)} \n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }

        const starterCode = generateStarterCode(this.project.host);
        const templateFills = Object.assign({}, this.project, starterCode);

        /** Copy the manifest */
        this.fs.copyTpl(this.templatePath(`manifest/${this.project.hostInternalName}.xml`), this.destinationPath(`${this.project.projectInternalName}-manifest.xml`), templateFills);

        if (this.project.projectType.toLowerCase() === manifest) {
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

  _projectFolderExists()
  {
    try
    {
      if (fs.existsSync(this._destinationRoot))
      {
        throw new Error('Folder already exists');
      }
    }
    catch(err)
    {
      this.log(`${chalk.bold.red(`\nFolder already exists at `) + this._destinationRoot + `. Please start over and choose a different project name or destination folder`}\n`); 
      insight.trackException(new Error('Installation Error: ' + err));
      this._exitProcess();
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