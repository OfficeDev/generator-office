import * as fs from 'fs';
import * as path from 'path';
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';

let uuid = require('uuid/v4');
let yosay = require('yosay');
let yo = require('yeoman-generator');

// TODO: waiting for app insight data pipeline followup
//let insight = appInsights.getClient('1fd62c46-f0ef-4cfb-9560-448c857ab690');

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    this.argument('host', { type: String, required: false });
    this.argument('name', { type: String, required: false });
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
    let jsTemplates = getDirectories(this.templatePath('js')).concat('manifest');
    let tsTemplates = getDirectories(this.templatePath('ts')).concat('manifest');
    let manifests = getFiles(this.templatePath('manifest')).map(manifest => manifest.replace('.xml', ''));

    let prompts = [
      /** allow user to create new project or update existing project */
      {
        name: 'new',
        message: 'Would you like to create a new add-in?',
        type: 'confirm',
        default: true
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
        default: false
      },

      /** office client application that can host the addin */
      {
        name: 'host',
        message: 'Create the add-in for:',
        type: 'list',
        default: 'excel',
        choices: manifests.map(manifest => ({ name: manifest, value: manifest })),
        when: (this.options.host == null)
      },

      /** use TypeScript for the project */
      {
        name: 'ts',
        type: 'confirm',
        message: 'Would you like to use TypeScript?',
        default: true,
        when: (this.options.js == null)
      }
    ];

    // insight.trackTrace('User begins to choose options');
    // let start = (new Date()).getTime();

    // trigger prompts and store user input
    let answers = await this.prompt(prompts);

    let frameworkPrompts = [
      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: tsTemplates.map(template => ({ name: template, value: template })),
        when: (this.options.framework == null) && answers.ts
      },

      /** technology used to create the addin (html / angular / etc) */
      {
        name: 'framework',
        message: 'Choose a framework:',
        type: 'list',
        default: 'jquery',
        choices: jsTemplates.map(template => ({ name: template, value: template })),
        when: (this.options.framework == null) && !answers.ts
      }
    ];

    let frameworkAnswers = await this.prompt(frameworkPrompts);

    // let end = (new Date()).getTime();
    // let duration = (end - start) / 1000;
    // insight.trackEvent('WHYME', { Project_Type: this.project.host }, { duration });

    this.project = {
      new: answers.new,
      name: this.options.name || answers.name,
      host: this.options.host || answers.host,
      ts: answers.ts,
      folder: answers.folder,
      framework: frameworkAnswers.framework || 'jquery'
    };

    if (!(this.options.js == null)) {
      this.project.ts = !this.options.js;
    }
    else {
      this.project.ts = answers.ts;
    }

    if (answers.folder == null) {
      this.project.folder = true;
    }

    if (answers.new == null) {
      this.project.new = true;
    }
  },

  /**
   * save configs & config project
   */
  configuring: function () {
    this.project.projectInternalName = _.kebabCase(this.project.name);
    this.project.projectDisplayName = _.capitalize(this.project.name);
    this.project.manifest = this.project.host + '-' + this.project.projectInternalName;
    this.project.new = this.project.new;
    this.project.projectId = uuid();
    if (this.project.folder) {
      this.destinationRoot(this.project.projectInternalName);
    }
  },

  writing: {
    copyFiles: function () {
      let language = this.project.ts ? 'ts' : 'js';

      console.log('----------------------------------------------------------------------------------\n');
      console.log(`Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in using ${chalk.bold.magenta(language)} and ${chalk.bold.cyan(this.project.framework)}\n`);
      console.log('----------------------------------------------------------------------------------\n\n');

      /** Copy the manifest */
      this.fs.copyTpl(this.templatePath(`manifest/${this.project.host}.xml`), this.destinationPath(`manifest-${this.project.manifest}.xml`), this.project);
    
      if (this.project.new === true) {
        /** Copy the base template */
        this.fs.copy(this.templatePath(`${language}/base/**`), this.destinationPath());

        /** Copy the framework specific overrides */
        this.fs.copyTpl(this.templatePath(`${language}/${this.project.framework}/**`), this.destinationPath(), this.project);
      }
    }
  },

  install: function () {
    if (!this.options['skip-install'] && this.project.framework !== 'manifest-only') {
      this.npmInstall();
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
