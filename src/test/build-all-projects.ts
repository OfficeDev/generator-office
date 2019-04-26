import * as fs from 'fs';
import * as path from 'path';
import projectsJsonData from './../app/config/projectsJsonData'
import { helperMethods } from './../app/helpers/helperMethods';
import * as _ from 'lodash';

let shell = require('shelljs');
let assert = require('assert');
let jsonData = new projectsJsonData(process.cwd() + '/generators/test');

const stringBuildStart = 'Install and build ';
const stringBuildSucceeds = 'Install and build succeeds';
const yoOffice = 'yo office';
const output = '--output';
const js = '--js';
const ts = '--ts';
const javascript = 'javascript';
const typescript = 'typescript';
const space = ' ';
let parsedProjectJsonData = jsonData.getParsedProjectJsonData();
let projectTemplates = jsonData.getProjectTemplateNames();
let hostsTemplates = jsonData.getHostTemplateNames();

describe('Setup test environment for Yo Office build tests', () => {
    it ('Install Yeoman Generator, install local instance of Yo Office and link', function(done){
        _setupTestEnvironment();
        done();
    });
});

describe('Install and build projects', () => {        
});

// Build Typescript project types for all supported hosts
for (var i = 0; i < hostsTemplates.length; i++)
{
    let host = hostsTemplates[i].toLowerCase();
    
    for (var j = 0; j < projectTemplates.length; j++)
    {        
        let projectType = projectTemplates[j].toLowerCase();

        // The excel-functions project is only relevant if the host is Excel
        if (projectType == 'excel-functions' && host != 'excel'){
            continue;
        }

        // If projectType is manifest, only install the project.  Building the project is not applicable
        if (projectType == 'manifest')
        {
            describe('Install ' +  host + space + projectType, () => {
                it(stringBuildSucceeds,function(done){
                    let projectName = projectType + host;
                    let projectFolder = path.join(__dirname, '/', projectName);
                    _installProject(projectType, projectName, host, projectFolder, undefined);
                    done();                    
                });
            });
        }
        else
        {
            if (parsedProjectJsonData.projectTypes[projectTemplates[j]].templates.typescript != undefined)
            {
                describe(stringBuildStart +  host + space + projectType + space + typescript, () => {
                    it(stringBuildSucceeds,function(done){
                        _installBuildProject(host, projectType, typescript);
                        done();                    
                    });
                });
            }
            if (parsedProjectJsonData.projectTypes[projectTemplates[j]].templates.javascript != undefined)
            {
                describe(stringBuildStart +  host + space + projectType + space + javascript, () => {
                    it(stringBuildSucceeds,function(done){
                        _installBuildProject(host, projectType, javascript);
                        done();                    
                    });
                });
            }
        }
    }
}

function _installBuildProject(host: string, projectType: string, scriptType: string)
{
    let projectName = projectType + host + scriptType;
    let projectFolder = path.join(__dirname, '/', projectName);
    _installProject(projectType, projectName, host, projectFolder, scriptType);
    _buildProject(projectFolder);
}

function _setupTestEnvironment()
{
    shell.exec('npm install -g yo', {silent: true}); 
    shell.exec('npm install', {silent: true});
    shell.exec('npm link', {silent: true});
}

function _installProject(projectType: string, projectName: string, host: string, projectFolder: string, scriptType: string)
{
    let language = scriptType == javascript ? js : ts;
    let cmdLine = yoOffice + space + projectType + space + projectName + space + host + space + output + space + projectFolder + space + language;
    shell.exec(cmdLine, {silent: true});
}

function _buildProject(projectFolder: string)
{
    if (helperMethods.doesProjectFolderExist(projectFolder))
    {
        const failure = 'error';
        shell.cd(projectFolder);
        let buildOutput = shell.exec('npm run build', {silent: true}).stdout;
        assert.equal(buildOutput.toLowerCase().indexOf(failure), -1, "Build output contained errors");
        shell.cd(__dirname);
        
        // do clean-up after test runs
        helperMethods.deleteFolderRecursively(projectFolder);
    }
    else
    {
        assert(false, projectFolder + " doesn't exist");
    }
}

function _projectFolderExists (projectFolder: string)
 {      
   if (fs.existsSync(projectFolder))
     {
       if (fs.readdirSync(projectFolder).length > 0)
       {          
         return true;
       }
     }
     return false;
 };