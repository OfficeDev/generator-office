/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
import _ from 'lodash';
import chalk from 'chalk';
import * as defaults from "./defaults.js";
import { helperMethods } from './helpers/helperMethods.js';
import projectsJsonData from './config/projectsJsonData.js';
import * as usageData from "office-addin-usage-data";
import { v4 as uuidv4 } from 'uuid';
import yosay from 'yosay';
import Generator, { PromptQuestion } from 'yeoman-generator';
import * as fs from "fs";
import * as path from "path";
import debug from "debug";

const log = debug("genOffice");
const trace = log.extend("trace");

const excelCustomFunctions = `excel-functions`;
let isSsoProject = false;
const javascript = `JavaScript`;
let language: string;
const manifest = 'manifest';
const sso = 'single-sign-on';
const typescript = `TypeScript`;
let jsonData: projectsJsonData;

let usageDataObject: usageData.OfficeAddinUsageData;
const usageDataOptions: usageData.IUsageDataOptions = {
  groupName: usageData.groupName,
  projectName: defaults.usageDataProjectName,
  raisePrompt: false,
  instrumentationKey: usageData.instrumentationKeyForOfficeAddinCLITools,
  promptQuestion: defaults.usageDataPromptMessage,
  usageDataLevel: usageData.UsageDataLevel.off,
  method: usageData.UsageDataReportingMethod.applicationInsights,
  isForTesting: false
}

export default class extends Generator {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  project: any;

  /*  Setup the generator */
  constructor(args, opts) {
    super(args, opts);

    if (parseInt(process.version.slice(1, process.version.indexOf('.'))) % 2 == 1) {
      console.log(yosay('generator-office does not support your version of Node. Please switch to the latest LTS version of Node.'));
      this._exitProcess();
    }

    this.argument('projectType', { type: String, required: false });
    this.argument('name', { type: String, required: false });
    this.argument('host', { type: String, required: false });
    this.argument('manifestType', { type: String, required: false });

    this.option('skip-install', {
      type: Boolean,
      description: 'Skip running `npm install` post scaffolding.'
    });

    this.option('js', {
      type: Boolean,
      description: 'Project uses JavaScript instead of TypeScript.'
    });

    this.option('ts', {
      type: Boolean,
      description: 'Project uses TypeScript instead of JavaScript.'
    });

    this.option('output', {
      alias: 'o',
      type: String,
      description: 'Project folder name if different from project name.'
    });

    this.option('prerelease', {
      type: String,
      description: 'Use the prerelease version of the project template.'
    });

    this.option('test', {
      type: String,
      description: 'Project is created in the context of unit tests.'
    });

    this.option('details', {
      alias: 'd',
      type: Boolean,
      description: 'Get more details on Yo Office arguments.'
    });
  }

  /* Generator initalization */
  async initializing(): Promise<void> {
    if (this.options["details"]) {
      await this._detailedHelp();
    }
    if (this.options['test']) {
      usageDataOptions.isForTesting = true;
    }
    if (this.options['skip-install']) {
      this.options.skipInstall = true;
    }
    const message = `Welcome to the ${chalk.bold.green('Office Add-in')} generator, by ${chalk.bold.green('@OfficeDev')}! Let\'s create a project together!`;
    await this.log(yosay(message));
    jsonData = new projectsJsonData(this.templatePath());
  }

  /* Prompt user for project options */
  async prompting(): Promise<void> {
    try {
      if (usageData.needToPromptForUsageData(usageDataOptions.groupName||"")) {
        const promptForUsageData: PromptQuestion[] = [
          {
            name: 'usageDataPromptAnswer',
            message: usageDataOptions.promptQuestion || defaults.usageDataPromptMessage,
            type: 'list',
            default: 'Continue',
            choices: ['Continue', 'Exit'],
          }
        ];
        const answerForUsageDataPrompt = await this.prompt(promptForUsageData);
        if (answerForUsageDataPrompt?.usageDataPromptAnswer === 'Continue') {
          usageDataOptions.usageDataLevel = usageData.UsageDataLevel.on;
        } else {
          process.exit();
        }
      } else {
        usageDataOptions.usageDataLevel = usageData.readUsageDataLevel(usageDataOptions.groupName||"");
      }

      let isManifestProject = false;
      let isExcelFunctionsProject = false;

      // Normalize host name if passed as a command line argument
      if (this.options["host"] != null) {
        this.options["host"] = jsonData.getHostDisplayName(this.options["host"]);
      }

      /* askForProjectType will only be triggered if no project type was specified via command line projectType argument,
       * and the projectType argument input was indeed valid */
      const startForProjectType = (new Date()).getTime();
      const askForProjectType: PromptQuestion[] = [
        {
          name: 'projectType',
          message: 'Choose a project type:',
          type: 'list',
          default: 'React',
          choices: jsonData.getProjectTemplateNames().map(template => ({ name: jsonData.getProjectDisplayName(template), value: template })),
          when: this.options["projectType"] == null || !jsonData.isValidProjectType(this.options["projectType"])
        }
      ];
      const answerForProjectType = await this.prompt(askForProjectType);
      const endForProjectType = (new Date()).getTime();
      const durationForProjectType = (endForProjectType - startForProjectType) / 1000;

      const projectType = _.toLower(this.options["projectType"]) || _.toLower(answerForProjectType.projectType);

      /* Set isManifestProject to true if Manifest project type selected from prompt or Manifest was specified via the command prompt */
      if ((answerForProjectType.projectType != null && _.toLower(answerForProjectType.projectType) === manifest)
        || (this.options["projectType"] != null && _.toLower(this.options["projectType"])) === manifest) {
        isManifestProject = true;
      }

      /* Set isExcelFunctionsProject to true if ExcelFunctions project type selected from prompt or Excel Functions was specified via the command prompt */
      if ((answerForProjectType.projectType != null && answerForProjectType.projectType) === excelCustomFunctions
        || (this.options["projectType"] != null && _.toLower(this.options["projectType"]) === excelCustomFunctions)) {
        isExcelFunctionsProject = true;
      }

      /* Set isSsoProject to true if SSO project type selected from prompt or Single Sign-On was specified via the command prompt */
      if ((answerForProjectType.projectType != null && answerForProjectType.projectType) === sso
        || (this.options["projectType"] != null && _.toLower(this.options["projectType"]) === sso)) {
        isSsoProject = true;
      }

      const getSupportedScriptTypes = jsonData.getScriptTypeOptions(projectType);
      const askForScriptType: PromptQuestion[] = [
        {
          name: 'scriptType',
          type: 'list',
          message: 'Choose a script type:',
          choices: getSupportedScriptTypes,
          default: getSupportedScriptTypes[0],
          when: !this.options["js"] && !this.options["ts"] && !isManifestProject && getSupportedScriptTypes.length > 1
        }
      ];
      const answerForScriptType = await this.prompt(askForScriptType);

      /* askforName will be triggered if no project name was specified via command line Name argument */
      const askForName: PromptQuestion[] = [{
        name: 'name',
        type: 'input',
        message: 'What do you want to name your add-in?',
        default: 'My Office Add-in',
        when: this.options["name"] == null
      }];
      const answerForName = await this.prompt(askForName);

      /* askForHost will be triggered if no project name was specified via the command line Host argument, and the Host argument
       * input was in fact valid, and the project type is not Excel-Functions */
      const startForHost = (new Date()).getTime();
      const supportedHosts = jsonData.getHostOptions(projectType);
      const askForHost: PromptQuestion[] = [{
        name: 'host',
        message: 'Which Office client application would you like to support?',
        type: 'list',
        default: supportedHosts[0],
        choices: supportedHosts.map(host => ({ name: host, value: host })),
        when: (this.options["host"] == null || this.options["host"] != null && !jsonData.isValidHost(this.options["host"]))
          && supportedHosts.length > 1
      }];
      const answerForHost = await this.prompt(askForHost);
      const endForHost = (new Date()).getTime();
      const durationForHost = (endForHost - startForHost) / 1000;

      const selectedHost = this.options["host"] || answerForHost.host || supportedHosts[0];

      usageDataObject = new usageData.OfficeAddinUsageData(usageDataOptions);

      /* aksForManifestType will be triggered if no type was specified via the command line manifestType argument */
      const startForManifestType = (new Date()).getTime();
      const manifestOptions = jsonData.getManifestOptions(projectType, selectedHost);
      const askForManifestType: PromptQuestion[] = [{
        name: 'manifestType',
        message: 'Which manifest type would you like to use?',
        type: 'list',
        default: manifestOptions[0],
        choices: manifestOptions.map(manifestType => ({ name: jsonData.getManifestDisplayName(manifestType), value: manifestType })),
        when: (this.options["manifestType"] == null || this.options["manifestType"] != null && !jsonData.isValidManifestType(this.options["manifestType"]))
          && jsonData.getManifestOptions(projectType, selectedHost).length > 1
      }];
      const answerForManifestType = await this.prompt(askForManifestType);
      const endForManifestType = (new Date()).getTime();
      const durationForManifestType = (endForManifestType - startForManifestType) / 1000;

      usageDataObject = new usageData.OfficeAddinUsageData(usageDataOptions);

      /* Configure project properties based on user input or answers to prompts */
      this._configureProject(answerForProjectType, answerForManifestType, answerForScriptType, answerForHost, answerForName, isManifestProject, isExcelFunctionsProject);
      const projectInfo = {
        Host: [this.project.host, durationForHost],
        ScriptType: [this.project.scriptType],
        IsManifestOnly: [this.project.isManifestOnly.toString()],
        ProjectType: [this.project.projectType, durationForProjectType],
        ManifestType: [this.project.manifestType, durationForManifestType],
        isForTesting: [usageDataOptions.isForTesting]
      };
      // Send usage data for project created
      usageDataObject.reportEvent(defaults.promptSelectionstEventName, projectInfo);
    } catch (err) {
      usageDataObject = new usageData.OfficeAddinUsageData(usageDataOptions);
      usageDataObject.reportError(defaults.promptSelectionsErrorEventName, new Error('Prompting Error: ' + err));
    }
  }

  async writing(): Promise<void> {
    await this._copyProjectFiles()
      .catch((err) => {
        usageDataObject.reportError(defaults.copyFilesErrorEventName, new Error('Installation Error: ' + err));
        process.exitCode = 1;
      });
  }

  async install(): Promise<void> {
    // Call 'convert-to-single-host' npm script in generated project, passing in host parameter
    // Need to call this here after package.json was written to disk, but before npm install is called
    await this.spawn("npm", ["run", "convert-to-single-host", "--if-present", "--", _.toLower(this.project.hostInternalName), this.project.manifestType, this.project.name]);
  }

  async end(): Promise<void> {
    if (!this.options['test']) {
      try {
        await this._postInstallHints();
      } catch (err) {
        usageDataObject.reportError(defaults.postInstallHintsErrorEventName, new Error('Exit Error: ' + err));
      }
    }
  }

  _configureProject(answerForProjectType, answerForManifestType, answerForScriptType, answerForHost, answerForName, isManifestProject, isExcelFunctionsProject): void {
    try {
      const projType = _.toLower(this.options["projectType"]) || _.toLower(answerForProjectType.projectType);
      const supportedHosts = jsonData.getHostOptions(projType);
      const selectedHost = this.options["host"] || answerForHost.host  || supportedHosts[0];

      this.project = {
        folder: this.options["output"] || answerForName.name || this.options["name"],
        host: answerForHost.host
          ? answerForHost.host
          : this.options["host"]
          ? this.options["host"]
          : jsonData?.getHostOptions(projType)[0],
        manifestType: answerForManifestType.manifestType
          ? answerForManifestType.manifestType
          : this.options["manifestType"]
          ? this.options["manifestType"]
          : jsonData?.getManifestOptions(projType, selectedHost)[0],
        name: this.options["name"] || answerForName.name,
        projectType: projType,
        scriptType: answerForScriptType.scriptType
          ? answerForScriptType.scriptType
          : this.options["ts"]
          ? typescript
          : this.options["js"]
          ? javascript
          : jsonData?.getScriptTypeOptions(projType)[0],
        isManifestOnly: isManifestProject,
        isExcelFunctionsProject: isExcelFunctionsProject,
      };

      /* Set folder if to output param  if specified */
      if (this.options["output"] != null) {
        this.project.folder = this.options["output"];
      }

      /* Set language variable */
      language = this.project.scriptType === typescript ? 'ts' : 'js';

      this.project.projectInternalName = _.kebabCase(this.project.name);
      this.project.projectDisplayName = this.project.name;
      this.project.projectId = uuidv4();
      this.project.hostInternalName = this.project.host == "All" ? "wxpo" : this.project.host;

      this.destinationRoot(this.project.folder);
      process.chdir(this.destinationRoot());
      this.env.cwd = this.destinationRoot();

      /* Check to to see if destination folder already exists. If so, we will exit and prompt the user to provide
      a different project name or output folder */
      this._exitYoOfficeIfProjectFolderExists();
    }
    catch (err) {
      usageDataObject.reportError(defaults.configurationErrorEventName, new Error('Configuration Error: ' + err));

    }
  }

  async _copyProjectFiles(): Promise<void> {
    return new Promise(async (resolve, reject) => {
      try {
        const projectRepoBranchInfo = jsonData.getProjectRepoAndBranch(this.project.projectType, language, this.options["prerelease"]);

        this._projectCreationMessage();

        // Copy project template files from project repository (currently only custom functions has its own separate repo)
        if (projectRepoBranchInfo.repo && projectRepoBranchInfo.branch) {
          const projectFolder: string = this.destinationPath();
          const zipFile: string = await helperMethods.downloadProjectTemplateZipFile(projectFolder, projectRepoBranchInfo.repo, projectRepoBranchInfo.branch);
          log("Zip File %s",zipFile);
          const unzippedFolder: string = await helperMethods.unzipProjectTemplate(projectFolder);
          log("Unzipped Folder: %s",unzippedFolder);
          const moveFromFolder: string = this.destinationPath(unzippedFolder);
          log("Moved From Folder: %s",moveFromFolder);

          // delete original zip file
          if (fs.existsSync(zipFile)) {
              log("Deleting original zip file %s",zipFile)
              fs.unlinkSync(zipFile);
          }

          // loop through all the files and folders in the unzipped folder and move them to project root
          log("Moving files from %s to %s",moveFromFolder,projectFolder);
          fs.readdirSync(moveFromFolder)
            .filter((file) => {
              let check = !file.includes(".gitignore") && !file.includes("package.json");
              if(!check) {
                log("Excluding %s",file);
              }
              return check;
            })
            .forEach(function (file) {
              trace("Moving %s",file);
              const fromPath = path.join(moveFromFolder, file);
              const toPath = path.join(projectFolder, file);
              if(fs.existsSync(fromPath)) {
                fs.renameSync(fromPath, toPath);
              }
            });

          // copy package.json file to new project directory and trigger npm install
          log("Copying package.json from %s to %s",moveFromFolder,projectFolder);
          this.fs.copyTpl(path.join(moveFromFolder, "package.json"), path.join(projectFolder, "package.json"));
      
          // delete project zipped folder
          log("Deleting folder: %s",unzippedFolder)
          helperMethods.deleteFolderRecursively(this.destinationPath(unzippedFolder));
 
        }
        else {
          // Manifest-only project
          const templateFills = Object.assign({}, this.project);
          this.fs.copyTpl(this.templatePath(`hosts/${_.toLower(this.project.hostInternalName)}/manifest.xml`), this.destinationPath('manifest.xml'), templateFills);
          this.fs.copyTpl(this.templatePath(`manifest-only/**`), this.destinationPath(), templateFills);
        }

        return resolve();
      }
      catch (err) {
        usageDataObject.reportError(defaults.copyFilesErrorEventName, new Error("File Copy Error: " + err));
        return reject(err);
      }
    });
  }

  async _postInstallHints(): Promise<void> {
    const projFolder: string = /\s/.test(this.destinationRoot()) ? "\"" + this.destinationRoot() + "\"" : this.destinationRoot();
    let stepNumber = 1;

    /* Next steps and npm commands */
    await this.log('----------------------------------------------------------------------------------------------------------\n');
    await this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    await this.log(`      ${stepNumber++}. Go the directory where your project was created:\n`);
    await this.log(`         ${chalk.bold('cd ' + projFolder)}\n`);
    
    if (isSsoProject) {
      await this.log(`      ${stepNumber++}. Configure your SSO taskpane add-in:\n`);
      await this.log(`         ${chalk.bold('npm run configure-sso')}\n`);
    } else if (this.project.isExcelFunctionsProject) {
      await this.log(`      ${stepNumber++}. Build your Excel Custom Functions taskpane add-in:\n`);
      await this.log(`         ${chalk.bold('npm run build')}\n`);
    }
    
    if (!this.project.isManifestOnly) {
      if (this.project.host === "Excel" || this.project.host === "Word" || this.project.host === "Powerpoint" || this.project.host === "Outlook" || this.project.host === "All") {
        await this.log(`      ${stepNumber++}. Start the local web server and sideload the add-in:\n`);
        await this.log(`         ${chalk.bold('npm start')}\n`);
      } else {
        await this.log(`      ${stepNumber++}. Start the local web server:\n`);
        await this.log(`         ${chalk.bold('npm run dev-server')}\n`);
        await this.log(`      ${stepNumber++}. Sideload the the add-in:\n`);
        await this.log(`         ${chalk.bold('Follow these instructions:')}`);
        await this.log(`         ${defaults.networkShareSideloadingSteps}\n`);
      }
    }

    await this.log(`      ${stepNumber++}. Open the project in VS Code:\n`);
    await this.log(`         ${chalk.bold('code .')}\n`);
    await this.log(`         For more information, visit http://code.visualstudio.com.\n`);
    await this.log(`      Please visit https://learn.microsoft.com/office/dev/add-ins for more information about Office Add-ins.\n`);
    if(this.project.host === "Outlook") {
      await this.log(`      Please visit ${defaults.outlookSideloadingSteps} for more information about Outlook sideloading.\n`);
    }
    await this.log('----------------------------------------------------------------------------------------------------------\n');
    this._exitProcess();
  }

  async _projectCreationMessage(): Promise<void> {
    /* Log to console the type of project being created */
    if (this.project.isManifestOnly) {
      await this.log('----------------------------------------------------------------------------------\n');
      await this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} at ${chalk.bold.magenta(this.destinationRoot())}\n`);
      await this.log('----------------------------------------------------------------------------------');
    }
    else {
      await this.log('\n----------------------------------------------------------------------------------\n');
      await this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in for ${chalk.bold.magenta(_.capitalize(this.project.host))}`);
      await this.log(`      using ${chalk.bold.yellow(this.project.scriptType)} and ${chalk.bold.magenta(jsonData.getProjectDisplayName(this.project.projectType))} and ${chalk.bold.yellow(jsonData.getManifestDisplayName(this.project.manifestType))}`);
      await this.log(`      at ${chalk.bold.magenta(this.destinationRoot())}\n`);
      await this.log('----------------------------------------------------------------------------------');
    }
  }

  async _detailedHelp(): Promise<void> {
    await this.log(`\nYo Office ${chalk.bgGreen('Arguments')} and ${chalk.bgMagenta('Options.')}\n`);
    await this.log(`NOTE: ${chalk.bgGreen('Arguments')} must be specified in the order below, and ${chalk.bgMagenta('Options')} must follow ${chalk.bgGreen('Arguments')}.\n`);
    await this.log(`  ${chalk.bgGreen('projectType')}:Specifies the type of project to create. Valid project types include:`);
    await this.log(`    ${chalk.yellow('taskpane:')} Creates an 'Office Add-in Task Pane project' project.`);
    await this.log(`    ${chalk.yellow('react:')} Creates an 'Office add-in using React framework' project.`);
    await this.log(`    ${chalk.yellow('excel-functions-shared:')} Creates an 'Office add-in for Excel custom functions using a Shared Runtime' project.`);
    await this.log(`    ${chalk.yellow('excel-functions:')} Creates an 'Office add-in for Excel custom functions using a JavaScript-only Runtime' project.`);
    await this.log(`    ${chalk.yellow('single-sign-on:')} Creates an 'Office Add-in Task Pane project supporting single sign-on' project.`);
    await this.log(`    ${chalk.yellow('nested-app-auth:')} Creates an 'Office Add-in Task Pane project supporting Nested App Auth single sign-on (preview)' project.`);
    await this.log(`    ${chalk.yellow('manifest:')} Creates an only the manifest file for an Office add-in project.\n`);
    await this.log(`  ${chalk.bgGreen('name')}:Specifies the name for the project that will be created.\n`);
    await this.log(`  ${chalk.bgGreen('host')}:Specifies the host app in the add-in manifest. Valid hosts include:`);
    await this.log(`    ${chalk.yellow('excel:')}  Creates an Office add-in for Excel.`);
    await this.log(`    ${chalk.yellow('onenote:')} Creates an Office add-in for OneNote.`);
    await this.log(`    ${chalk.yellow('outlook:')} Creates an Office add-in for Outlook.`);
    await this.log(`    ${chalk.yellow('powerpoint:')} Creates an Office add-in for PowerPoint.`);
    await this.log(`    ${chalk.yellow('project:')} Creates an Office add-in for Project.`);
    await this.log(`    ${chalk.yellow('word:')} Creates an Office add-in for Word.\n`);
    await this.log(`  ${chalk.bgGreen('manifestType')}:Specifies the manifest type to use for the add-in. Valid types include:`);
    await this.log(`    ${chalk.yellow('xml:')}  Creates a Add-in only manifest`);
    await this.log(`    ${chalk.yellow('json:')} Creates a unified manifest for Microsoft 365.\n`);
    await this.log(`  ${chalk.bgMagenta('--output')}:Specifies the location in the file system where the project will be created.`);
    await this.log(`    ${chalk.yellow('If the option is not specified, the project will be created in the current folder')}\n`);
    await this.log(`  ${chalk.bgMagenta('--js')}:Specifies that the project will use JavaScript instead of TypeScript.`);
    await this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
    await this.log(`  ${chalk.bgMagenta('--ts')}:Specifies that the project will use TypeScript instead of JavaScript.`);
    await this.log(`    ${chalk.yellow('If the option is not specified, Yo Office will prompt for TypeScript or JavaScript')}\n`);
    this._exitProcess();
  }

  async _exitYoOfficeIfProjectFolderExists(): Promise<boolean> {
    if (helperMethods.doesProjectFolderExist(this.destinationRoot())) {
      await this.log(`${chalk.bold.red(`\nFolder already exists at ${chalk.bold.green(this.destinationRoot())} and is not empty. To avoid accidentally overwriting any files, please start over and choose a different project name or destination folder via the ${chalk.bold.magenta(`--output`)} parameter`)}\n`);
      this._exitProcess();
    }
    return false;
  }

  _exitProcess(): void {
    process.exit();
  }
};
