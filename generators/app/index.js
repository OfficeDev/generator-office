"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments)).next());
    });
};
const yo = require("yeoman-generator");
const chalk = require("chalk");
const yosay = require("yosay");
const ncp = require("ncp");
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
    },
    /**
     * Generator initalization
     */
    initializing: function () {
        return __awaiter(this, void 0, void 0, function* () {
            this.log(yosay('Welcome to the ' +
                chalk.red('Office Project') +
                ' generator, by ' +
                chalk.red('@OfficeDev') +
                '! Let\'s create a project together!'));
            // create global config object on this generator
            this.genConfig = {};
        });
    },
    /**
     * Prompt users for options
     */
    prompting: function () {
        return __awaiter(this, void 0, void 0, function* () {
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
                        }
                    ]
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
                        }
                        else {
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
                            value: 'Document'
                        },
                        {
                            name: 'Excel',
                            value: 'Workbook'
                        },
                        {
                            name: 'PowerPoint',
                            value: 'Presentation'
                        },
                        {
                            name: 'OneNote',
                            value: 'Notebook'
                        },
                        {
                            name: 'Project',
                            value: 'Project'
                        }
                    ],
                    when: this.options.client === undefined
                }
            ];
            // trigger prompts and store user input
            yield this.prompt(prompts).then(function (responses) {
                this.genConfig = {
                    name: responses.name,
                    tech: responses.tech,
                    'root-path': responses['root-path'],
                    client: responses.client
                };
            }.bind(this));
        });
    },
    writing: function () {
        ncp.ncp(this.templatePath('common'), this.destinationPath(), err => console.log(err));
        switch (this.genConfig.tech) {
            case 'html':
                ncp.ncp(this.templatePath('html'), this.destinationPath(), err => console.log(err));
                break;
            case 'ng':
                ncp.ncp(this.templatePath('ng'), this.destinationPath(), err => console.log(err));
                break;
        }
    },
    install: function () {
        this.installDependencies();
    }
});
//# sourceMappingURL=index.js.map