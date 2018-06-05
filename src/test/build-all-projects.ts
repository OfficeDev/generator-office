import * as fs from 'fs';
import * as path from 'path';

var shell = require('shelljs');
var assert = require('assert');

const stringBuildStart = 'Generate and build ';
const stringBuildSucceeds = 'Install and build succeeds';
const yoOffice = 'yo office';
const output = '--output';
const js = '--js';
const ts = '--ts';
const javascript = 'javascript';
const typescript = 'typescript';
const space = ' ';
let jsTemplates = getDirectories('src/app/templates/js');
let tsTemplates = getDirectories('src/app/templates/js');
let hostsTemplates = getDirectories('src/app/templates/hosts');

describe('Setup test environment for Yo Office build tests', () => {
    it ('Install Yeoman Generator and Install local install of Yo Office and link', function(done){
        _setupTestEnvironment();
        done();
    });
}); 

// Build Typescript project types for all supported hosts
for (var i = 0; i < hostsTemplates.length; i++)
{
    for (var j = 0; j < tsTemplates.length; j++)
    {
        describe('Install and build projects', () => {
            let projectType = tsTemplates[j];        
            let host = hostsTemplates[i];
            let scriptType = typescript;
            let projectName = projectType + host + typescript;
            let projectFolder = path.join(__dirname, '/', projectName);          
        
            describe(stringBuildStart +  host + space + projectType + space + typescript, () => {
                it(stringBuildSucceeds,function(done){  
                    _generateProject(projectType, projectName, host, projectFolder, scriptType);
                    _buildProject(projectFolder, projectType);
                    done();                    
                  });
              }); 
            });
        }
    }

// Build Javascript project types for all supported hosts
for (var i = 0; i < hostsTemplates.length; i++)
{
    for (var j = 0; j < jsTemplates.length; j++)
    {
        describe('Install and build projects', () => {
            let projectType = jsTemplates[j];        
            let host = hostsTemplates[i];
            let scriptType = javascript;
            let projectName = projectType + host + javascript;
            let projectFolder = path.join(__dirname, '/', projectName);  
        
            describe(stringBuildStart +  host + space + projectType + space + javascript, () => {
                it(stringBuildSucceeds,function(done){         
                    _generateProject(projectType, projectName, host, projectFolder, scriptType);
                    _buildProject(projectFolder, projectType);
                    done();
                  });
              }); 
            });
        }
    }    
    
function _setupTestEnvironment()
{
    shell.exec('npm install -g yo', {silent: true}); 
    shell.exec('npm install', {silent: true});
    shell.exec('npm link', {silent: true});
}

function _generateProject(projectType, projectName, host, projectFolder, scriptType)
{
    let language = scriptType == javascript ? js : ts;
    let cmdLine = yoOffice + space + projectType + space + projectName + space + host + space + output + space + projectFolder + space + language;
    shell.exec(cmdLine, {silent: true});
}

function _buildProject(projectFolder, projectType)
{
    if (_projectFolderExists(projectFolder))
    {
        const failure = 'error';
        shell.cd(projectFolder);
        let buildOutput = shell.exec('npm run build', {silent: true}).stdout;
        assert.equal(buildOutput.toLowerCase().indexOf(failure), -1, "Build output contained errors");
        shell.cd(__dirname);
        
        // do clean-up after test runs
        _deleteFolderRecursively(projectFolder);
    }
    else
    {
        assert(false, projectFolder + " doesn't exist");
    }
}

function _projectFolderExists (projectFolder)
 {      
   if (fs.existsSync(projectFolder))
     {
       if (fs.readdirSync(projectFolder).length > 0)
       {          
         return true;
       }
     }
     return false;
 }

 function getDirectories(root)
 {
  return fs.readdirSync(root).filter(file => {
    if (file === 'base') {
      return false;
    }
    return fs.statSync(path.join(root, file)).isDirectory();
  });
}

function _deleteFolderRecursively(projectFolder) 
{
    if(fs.existsSync(projectFolder))
    {
        fs.readdirSync(projectFolder).forEach(function(file,index){ 
        var curPath = projectFolder + "/" + file; 
        
        if(fs.lstatSync(curPath).isDirectory())
        {
            _deleteFolderRecursively(curPath);
        }
        else
        {
            fs.unlinkSync(curPath);
        }
    }); 
    fs.rmdirSync(projectFolder); 
    }
};