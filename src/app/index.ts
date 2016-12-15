'use strict'

//http://mammal.io/articles/yeoman-generators-es6/

const guid = require('uuid');
import yo = require('yeoman-generator');
import chalk = require('chalk');
import yosay = require('yosay');
import ncp = require('ncp');
import Xml2Js = require('xml2js');
import * as path from 'path';

module.exports = yo.Base.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.Base.apply(this, arguments);

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

    this.option('client', {
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
      // friendly name of the generator
      {
        name: 'name',
        message: 'Project name (display name):',
        default: 'My Office Project',
        when: this.options.name === undefined
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
        filter: /* istanbul ignore next */ function (response) {
          if (response === 'current folder') {
            return '.';
          } else {
            return response;
          }
        }
      },
      // office client application that can host the addin
      {
        name: 'client',
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
        when: this.options.client === undefined
      }];

    // trigger prompts and store user input
    await this.prompt(prompts).then(function (responses) {
      this.genConfig = {
        name: responses.name,
        tech: responses.tech,
        'root-path': responses['root-path'],
        client: responses.client
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

    this.genConfig.projectId = guid.v4();
  }, // configuring()

  writing: {
    copyFiles: function () {
      /**
       * Output files
       */
      var manifestFilename = 'manifest-' + this.genConfig.client + '.xml';

      ncp.ncp(this.templatePath('common-static'), this.destinationPath(), err => console.log(err));
      this.fs.copyTpl(this.templatePath('common-dynamic/package.json'), 
            this.destinationPath('package.json'),
            this.genConfig);

      switch (this.genConfig.tech) {
        case 'html':
          ncp.ncp(this.templatePath('tech/html'), this.destinationPath(), err => console.log(err));
          break;
        case 'ng':
          ncp.ncp(this.templatePath('tech/ng'), this.destinationPath(), err => console.log(err));
          break;
      };

      switch (this.genConfig.client) {
        case 'document':
          this.fs.copyTpl(this.templatePath('hosts/word/' + manifestFilename), 
                          this.destinationPath(manifestFilename),
                          this.genConfig);
          break;
          //  case 'workbook':
          // this.fs.copyTpl(this.templatePath('hosts/workbook/' + manifestFilename), 
          //                 this.destinationPath(manifestFilename),
          //                 this.genConfig);
          // break;
      };
    },

    updateXml: function () {
      /**
       * Update the manifest.xml elements with the client input.
       */

      // manifest filename
      var manifestFilename = 'manifest-' + this.genConfig.client + '.xml';

      // workaround to 'this' context issue... I know it's hacky. Don't judge.
      var self = this;

      // load manifest.xml
      var manifestXml = self.fs.read(self.destinationPath(manifestFilename));

      // convert it to JSON
      var parser = new Xml2Js.Parser();
      parser.parseString(manifestXml, function (err, manifestJson) {
        manifestJson.OfficeApp.Id = self.genConfig.projectId;
        manifestJson.OfficeApp.DisplayName[0].$['DefaultValue'] = self.genConfig.projectDisplayName;

        // convert JSON => XML
        var xmlBuilder = new Xml2Js.Builder();
        var updatedManifestXml = xmlBuilder.buildObject(manifestJson);

        // write updated manifest
        self.fs.write(self.destinationPath(manifestFilename), updatedManifestXml);
      });
    }
  },

  install: function () {
    this.installDependencies();
  }
} as any);
