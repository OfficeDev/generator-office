import * as fs from 'fs';
import * as path from 'path';
import projectsJsonData from './../app/config/projectsJsonData'
import { helperMethods } from './../app/helpers/helperMethods';
import * as _ from 'lodash';

const assert = require('assert');
const jsonData = new projectsJsonData(process.cwd() + '/generators/test');
const js = '--js';
const ts = '--ts';
const parsedProjectJsonData = jsonData.getParsedProjectJsonData();
const projectTemplates = jsonData.getProjectTemplateNames();
const javascript = 'javascript';
const shell = require('shelljs');
const typescript = 'typescript';

describe('Setup test environment for Yo Office build tests', () => {
    it('Install Yeoman Generator, install local instance of Yo Office and link', function (done) {
        _setupTestEnvironment();
        done();
    });
});

describe('Install and build projects', () => {
});

// Install and build all supported projects for typescript and javascript
for (var j = 0; j < projectTemplates.length; j++) {
    let projectType = projectTemplates[j].toLowerCase();

    // If projectType is manifest, only install the project.  Building the project is not applicable
    if (projectType == 'manifest') {
        describe('Install ' + projectType, () => {
            it('Install and build succeeds', function (done) {
                let projectName = projectType;
                let projectFolder = path.join(__dirname, '/', projectName);
                _installProject(projectType, projectName, projectFolder, undefined);
                done();
            });
        });
    }
    else {
        if (parsedProjectJsonData.projectTypes[projectTemplates[j]].templates.typescript != undefined) {
            describe(`Install and build ${projectType} using typescript`, () => {
                it('Install and build succeeds', function (done) {
                    _installBuildProject(projectType, typescript);
                    done();
                });
            });
        }
        if (parsedProjectJsonData.projectTypes[projectTemplates[j]].templates.javascript != undefined) {
            describe(`Install and build ${projectType} using javascript`, () => {
                it('Install and build succeeds', function (done) {
                    _installBuildProject(projectType, javascript);
                    done();
                });
            });
        }
    }
}

function _installBuildProject(projectType: string, scriptType: string) {
    let projectName = `${projectType}-${scriptType}`;
    let projectFolder = path.join(__dirname, '/', projectName);
    _installProject(projectType, projectName, projectFolder, scriptType);
    _buildProject(projectFolder);
}

function _setupTestEnvironment() {
    shell.exec('npm install -g yo', { silent: true });
    shell.exec('npm install', { silent: true });
    shell.exec('npm link', { silent: true });
}

function _installProject(projectType: string, projectName: string, projectFolder: string, scriptType: string) {
    let language = scriptType == javascript ? js : ts;
    let cmdLine = `yo office --projectType ${projectType} --name ${projectName} ${language} --output ${projectFolder}`;
    shell.exec(cmdLine, { silent: true });
}

function _buildProject(projectFolder: string) {
    if (helperMethods.doesProjectFolderExists(projectFolder)) {
        const failure = 'error';
        shell.cd(projectFolder);
        let buildOutput = shell.exec('npm run build', { silent: true }).stdout;
        assert.equal(buildOutput.toLowerCase().indexOf(failure), -1, "Build output contained errors");
        shell.cd(__dirname);

        // do clean-up after test runs
        helperMethods.deleteFolderRecursively(projectFolder);
    }
    else {
        assert(false, projectFolder + " doesn't exist");
    }
}