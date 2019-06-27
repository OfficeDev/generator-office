/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import * as _ from 'lodash';
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as childProcess from "child_process";
import * as uuid from 'uuid/v4';
import { promisify } from "util";
import * as yosay from 'yosay';
import * as yo from 'yeoman-generator';
import projectsJsonData from './config/projectsJsonData';
import { helperMethods } from './helpers/helperMethods';
import { modifyManifestFile } from 'office-addin-manifest';

let insight = appInsights.getClient('1ced6a2f-b3b2-4da5-a1b8-746512fbc840');
const childProcessExec = promisify(childProcess.exec);
const excelCustomFunctions = `excel-functions`;
const manifest = 'manifest';
const typescript = `TypeScript`;
const javascript = `JavaScript`;
let language;

/* Remove unwanted tags */
delete insight.context.tags['ai.cloud.roleInstance'];
delete insight.context.tags['ai.device.osVersion'];
delete insight.context.tags['ai.device.osArchitecture'];
delete insight.context.tags['ai.device.osPlatform'];

module.exports = yo.extend({
 /*  Setup the generator */
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

    this.option('prerelease', {
      type: String,
      required: false,
      desc: 'Use the prerelease version of the project template.'
    });

    this.option('details', {
      alias: 'd',
      type: Boolean,
      required: false,
      desc: 'Get more details on Yo Office arguments.'
    });
  },

  /* Generator initalization */
  initializing: function () {
    if (this.options.details){
     this._detailedHelp();
    }
    let message = `Welcome to the ${chalk.bold.green('Office Add-in')} generator, by ${chalk.bold.green('@OfficeDev')}! Let\'s create a project together!`;
    this.log(yosay(message));
    this.project = {};
  },

  /* Prompt user for project options */
  prompting: async function () {
    try {
      let jsonData = new projectsJsonData(this.templatePath());
      let isManifestProject = false;
      let isExcelFunctionsProject = false;

      // Normalize host name if passed as a command line argument
      if (this.options.host != null) {
        this.options.host = jsonData.getHostDisplayName(this.options.host);
      }

      /* askForProjectType will only be triggered if no project type was specified via command line projectType argument,
       * and the projectType argument input was indeed valid */
      let startForProjectType = (new Date()).getTime();
      let askForProjectType = [
        {
          name: 'projectType',
          message: 'Choose a project type:',
          type: 'list',
          default: 'React',
          choices: jsonData.getProjectTemplateNames().map(template => ({ name: jsonData.getProjectDisplayName(template), value: template })),
          when: this.options.projectType == null || !jsonData.isValidInput(this.options.projectType, false /* isHostParam */)
        }
      ];
      let answerForProjectType = await this.prompt(askForProjectType);
      let endForProjectType = (new Date()).getTime();
      let durationForProjectType = (endForProjectType - startForProjectType) / 1000;

      /* Set isManifestProject to true if Manifest project type selected from prompt or Manifest was specified via the command prompt */
      if ((answerForProjectType.projectType != null && _.toLower(answerForProjectType.projectType) == manifest)
      || (this.options.projectType != null && _.toLower(this.options.projectType)) == manifest) {
          isManifestProject = true; }

      /* Set isExcelFunctionsProject to true if ExcelexcelFunctions project type selected from prompt or Excel Functions was specified via the command prompt */
      if ((answerForProjectType.projectType != null  && answerForProjectType.projectType) == excelCustomFunctions
      || (this.options.projectType != null && _.toLower(this.options.projectType) == excelCustomFunctions)) {
        isExcelFunctionsProject = true; }
        
      let askForScriptType = [
        {
          name: 'scriptType',
          type: 'list',
          message: 'Choose a script type:',
          choices: [typescript, javascript],
          default: typescript,
          when: !this.options.js && !this.options.ts && !isManifestProject
        }
      ];
      let answerForScriptType = await this.prompt(askForScriptType);

      /* askforName will be triggered if no project name was specified via command line Name argument */
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

      /* askForHost will be triggered if no project name was specified via the command line Host argument, and the Host argument
       * input was in fact valid, and the project type is not Excel-Functions */
      let startForHost = (new Date()).getTime();
      let askForHost = [{
        name: 'host',
        message: 'Which Office client application would you like to support?',
        type: 'list',
        default: 'Excel',
        choices: jsonData.getHostTemplateNames().map(host => ({ name: host, value: host })),
        when: (this.options.host == null || this.options.host != null && !jsonData.isValidInput(this.options.host, true /* isHostParam */))
        && !isExcelFunctionsProject
      }];
      let answerForHost = await this.prompt(askForHost);
      let endForHost = (new Date()).getTime();
      let durationForHost = (endForHost - startForHost) / 1000;

      /* Configure project properties based on user input or answers to prompts */
      this._configureProject(answerForProjectType, answerForScriptType, answerForHost, answerForName, isManifestProject, isExcelFunctionsProject);

      /* Gnerate Insights logging */
      const noElapsedTime = 0;
      insight.trackEvent('Name', { Name: this.project.name }, { durationForName });
      insight.trackEvent('Host', { Host: this.project.host }, { durationForHost });
      insight.trackEvent('ScriptType', { ScriptType: this.project.scriptType }, { noElapsedTime });
      insight.trackEvent('IsManifestOnly', { IsManifestOnly: this.project.isManifestOnly.toString() }, { noElapsedTime });
      insight.trackEvent('ProjectType', { ProjectType: this.project.projectType }, { durationForProjectType });
    } catch (err) {
      insight.trackException(new Error('Prompting Error: ' + err));
    }
  },

  writing: function () {
    const done = this.async();
    this._copyProjectFiles()
    .then(() => {
      done();
    })
    .catch((err) => {
      insight.trackException(new Error('Installation Error: ' + err));
      process.exitCode = 1;
    });
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

  _configureProject: function(answerForProjectType, answerForScriptType, answerForHost, answerForName, isManifestProject, isExcelFunctionsProject)
  {
    try
    {
      this.project = {
        folder: this.options.output || answerForName.name || this.options.name,
        name: this.options.name || answerForName.name,
        host: this.options.host || answerForHost.host,
        projectType: _.toLower(this.options.projectType) || _.toLower(answerForProjectType.projectType),
        isManifestOnly: isManifestProject,
        isExcelFunctionsProject: isExcelFunctionsProject,
        scriptType: answerForScriptType.scriptType ? answerForScriptType.scriptType : this.options.ts ? typescript : javascript 
      };

      /* Set folder if to output param  if specified */
      if (this.options.output != null) {
        this.project.folder = this.options.output; }
      
      /* Set language variable */
      language = this.project.scriptType === typescript ? 'ts' : 'js';

      this.project.projectInternalName = _.kebabCase(this.project.name);
      this.project.projectDisplayName = this.project.name;
      this.project.projectId = uuid();
      if (this.project.projectType === excelCustomFunctions) {
        this.project.host = 'Excel';
        this.project.hostInternalName = 'Excel';
      }
      else {
        this.project.hostInternalName = this.project.host;
      }
      this.destinationRoot(this.project.folder);

      /* Check to to see if destination folder already exists. If so, we will exit and prompt the user to provide
      a different project name or output folder */
      this._exitYoOfficeIfProjectFolderExists();

      let duration = this.project.duration;
      insight.trackEvent('App_Data', { AppID: this.project.projectId, Host: this.project.host, ProjectType: this.project.projectType, isTypeScript: (this.project.scriptType === typescript).toString() }, { duration });
    }
    catch (err) {
      insight.trackException(new Error('Configuration Error: ' + err));
    }
  },
  
  _copyProjectFiles()
  {
    return new Promise(async (resolve, reject) => {
      try {        
        let jsonData = new projectsJsonData(this.templatePath());
        let projectRepoBranchInfo = jsonData.getProjectRepoAndBranch(this.project.projectType, language, this.options.prerelease);

        this._projectCreationMessage();

        // Copy project template files from project repository (currently only custom functions has its own separate repo)
        if (projectRepoBranchInfo.repo) {
          await helperMethods.downloadProjectTemplateZipFile(this.destinationPath(), projectRepoBranchInfo.repo, projectRepoBranchInfo.branch);
              
          // Call 'convert-to-single-host' npm script in generated project, passing in host parameter
          const cmdLine = `npm run convert-to-single-host --if-present -- ${_.toLower(this.project.hostInternalName)}`;
          await childProcessExec(cmdLine);

          // modify manifest guid and DisplayName
          await modifyManifestFile(`${this.destinationPath()}/manifest.xml`, 'random', `${this.project.name}`);

          return resolve()
        }
        else {
          // Manifest-only project
          const templateFills = Object.assign({}, this.project);
          this.fs.copyTpl(this.templatePath(`hosts/${_.toLower(this.project.hostInternalName)}/manifest.xml`), this.destinationPath('manifest.xml'), templateFills);
          this.fs.copyTpl(this.templatePath(`manifest-only/**`), this.destinationPath(), templateFills);
          return resolve();
        }
      }
      catch (err) {
        insight.trackException(new Error('File Copy Error: ' + err));
        return reject(err);
      }
    });
  },
  _postInstallHints: function () {
    /* Next steps and npm commands */
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    this.log(`      1. Go the directory where your project was created:\n`);
    this.log(`         ${chalk.bold('cd ' + this._destinationRoot)}\n`);
    this.log(`      2. Start the local web server and sideload the add-in:\n`);
    this.log(`         ${chalk.bold('npm start')}\n`);
    this.log(`      3. Open the project in VS Code:\n`);
    this.log(`         ${chalk.bold('code .')}\n`);
    this.log(`         For more information, visit http://code.visualstudio.com.\n`);
    this.log(`      Please visit https://docs.microsoft.com/office/dev/add-ins for more information about Office Add-ins.\n`);
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this._exitProcess();
  },

  _projectCreationMessage: function()
  {
    /* Log to console the type of project being created */
    if (this.project.isManifestOnly)
      {
        this.log('----------------------------------------------------------------------------------\n');
        this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} at ${chalk.bold.magenta(this._destinationRoot)}\n`);
        this.log('----------------------------------------------------------------------------------');
      }
    else
      {
        this.log('\n----------------------------------------------------------------------------------\n');
        this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in for ${chalk.bold.magenta(_.capitalize(this.project.host))} using ${chalk.bold.yellow(this.project.scriptType)} and ${chalk.bold.green(_.capitalize(this.project.projectType))} at ${chalk.bold.magenta(this._destinationRoot)}\n`);
        this.log('----------------------------------------------------------------------------------');
      }
  },

  _detailedHelp: function () {
    this.log(`\nYo Office ${chalk.bgGreen('Arguments')} and ${chalk.bgMagenta('Options.')}\n`);
    this.log(`NOTE: ${chalk.bgGreen('Arguments')} must be specified in the order below, and ${chalk.bgMagenta('Options')} must follow ${chalk.bgGreen('Arguments')}.\n`);
    this.log(`  ${chalk.bgGreen('projectType')}:Specifies the type of project to create. Valid project types include:`);
    this.log(`    ${chalk.yellow('angular:')}  Creates an Office add-in using Angular framework.`);
    this.log(`    ${chalk.yellow('excel-functions:')} Creates an Office add-in for Excel custom functions.  Must specify 'Excel' as host parameter.`);
    this.log(`    ${chalk.yellow('jquery:')} Creates an Office add-in using Jquery framework.`);
    this.log(`    ${chalk.yellow('manifest:')} Creates an only the manifest file for an Office add-in.`);
    this.log(`    ${chalk.yellow('react:')} Creates an Office add-in using React framework.\n`);
    this.log(`  ${chalk.bgGreen('name')}:Specifies the name for the project that will be created.\n`);
    this.log(`  ${chalk.bgGreen('host')}:Specifies the host app in the add-in manifest.`);
    this.log(`    ${chalk.yellow('excel:')}  Creates an Office add-in for Excel. Valid hosts include:`);
    this.log(`    ${chalk.yellow('onenote:')} Creates an Office add-in for OneNote.`);
    this.log(`    ${chalk.yellow('outlook:')} Creates an Office add-in for Outlook.`);
    this.log(`    ${chalk.yellow('powerpoint:')} Creates an Office add-in for PowerPoint.`);
    this.log(`    ${chalk.yellow('project:')} Creates an Office add-in for Project.`);
    this.log(`    ${chalk.yellow('word:')} Creates an Office add-in for Word.\n`);
    this.log(`  ${chalk.bgMagenta('--output')}:Specifies the location in the file system where the project will be created.`);
    this.log(`    ${chalk.yellow('If the option is not specified, the project will be created in the current folder')}\n`);
    this.log(`  ${chalk.bgMagenta('--js')}:Specifies that the project will use JavaScript instead of TypeScript.`);
    this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
    this.log(`  ${chalk.bgMagenta('--ts')}:Specifies that the project will use TypeScript instead of JavaScript.`);
    this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
    this._exitProcess();
  },

_exitYoOfficeIfProjectFolderExists: function ()
  {
    if (helperMethods.doesProjectFolderExist(this._destinationRoot))
      {
          this.log(`${chalk.bold.red(`\nFolder already exists at ${chalk.bold.green(this._destinationRoot)} and is not empty. To avoid accidentally overwriting any files, please start over and choose a different project name or destination folder via the ${chalk.bold.magenta(`--output`)} parameter`)}\n`);
          this._exitProcess();
      }
      return false;
  },

  _exitProcess: function () {
    process.exit();
  }
} as any);
