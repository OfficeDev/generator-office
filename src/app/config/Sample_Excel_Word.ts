const shell = require('shelljs');
const ProgressBar = require('progress');
const log = require('single-line-log').stdout;
const Spinner = require('cli-spinner').Spinner;
const { spawn } = require('child_process');
const { exec } = require('child_process');
const { execSync } = require('child_process');
const readline = require('readline');
const open = require('open');

import * as os from 'os';
import * as fs from 'fs';

shell.config.silent = true;

async function exec_script_Excel_Mail(){
  // shell.exec('code .');
  console.log('Welcome to experience this Office add-in sample: Excel Mail merge Add-in project!');

  return new Promise<boolean>((resolve, reject) => {
    
        let is_vscode_installed = false;

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples.git', {async:true}, (code, stdout, stderr) => {
            shell.cd('./Excel-Scenario-based-Add-in-Samples/Mail-Merge-Sample-Add-in');
            // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {

            spinner.stop(true);
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0);

            replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-add-in-mail-merge-script', './src/taskpane/taskpane.html');

            console.log('Step [1/3] completed!');
    
            // Ask user if sample Add-in automatic launch is needed
            let rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            let auto_launch_answer = false;
            rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                if (answer.trim().toLowerCase() == 'y') {
                    auto_launch_answer = true;
                }

                rl.close();

                // Step 2: Check if VSCode is installed
                console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                if (shell.which('code')) {
                    console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                    is_vscode_installed = true;
                    shell.exec('code -n . ./README.md');
                } else {
                    console.log('Visual Studio Code is not installed on your machine.');
                    shell.exec('start Mail-Merge-Sample-Add-in');
                }

                console.log('Step [2/3] completed!');
                if (auto_launch_answer) {
                    // Continue with the operations
                    // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    console.log('It may take longer time to complete the process. Please wait patiently...');
                    spinner.start();

                    shell.cd('./Mail-Merge-Sample-Add-in');
                    shell.exec('npm install', {async:true}, (code, stdout, stderr) => {
                        shell.exec('npm run start', {async:true}, (code, stdout, stderr) => {

                        spinner.stop(true);
                        readline.clearLine(process.stdout, 0);
                        readline.cursorTo(process.stdout, 0);

                        console.log('Step [3/3] completed!');
                        console.log('Finished!');
                        // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                        // open('https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/tree/main/Mail-Merge-Sample-Add-in');
                        resolve(is_vscode_installed);
                        });
                    });
                }
                else{
                    // Don't continue with the operations
                    console.log('Step [3/3] skipped. You decided not to auto-launch the sample.')
                    console.log('No problem. You can always launch the sample add-in by running the following commands:');
                    console.log('--------------------------------------------');
                    console.log('npm install');
                    console.log('npm run start');
                    console.log('--------------------------------------------');    
                    console.log('Finished!');   
                    // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
                    // open('https://github.com/OfficeDev/Excel-Scenario-based-Add-in-Samples/tree/main/Mail-Merge-Sample-Add-in');
                    resolve(is_vscode_installed);
                }
            });
        });
    });
}

async function exec_script_Word_AIGC(){
    // shell.exec('code .');
    console.log('Welcome to experience this Office add-in sample: Word AIGC Add-in project!');
  
    return new Promise<boolean>((resolve, reject) => {
      
          let is_vscode_installed = false;
  
          // Step 1: Get sample code
          console.log('Step [1/3]: Getting sample code...');
          let spinner = new Spinner('Processing.. %s');
          spinner.setSpinnerString('|/-\\');
          spinner.start();
  
          shell.exec('git clone https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples.git', {async:true}, (code, stdout, stderr) => {
            shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
            // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
            replaceUrl('https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-localhost', 'https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc-script', './src/taskpane/taskpane-localhost.html');

            spinner.stop(true);
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0);

            // Step 2: Check if VSCode is installed
            console.log('Step [1/3] completed!');
              // Ask user if sample Add-in automatic launch is needed
            let rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            let auto_launch_answer = false;
            rl.question('Proceed to launch Office with the sample add-in? (Y/N)\n', (answer) => {
                if (answer.trim().toLowerCase() == 'y') {
                    auto_launch_answer = true;
                }

                rl.close();

                // Step 2: Check if VSCode is installed
                console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
                if (shell.which('code')) {
                    console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                    is_vscode_installed = true;
                    shell.exec('code -n . ./README.md');
                } else {
                    console.log('Visual Studio Code is not installed on your machine.');
                    shell.exec('start Word-Add-in-AIGC');
                }

                console.log('Step [2/3] completed!');
                if (auto_launch_answer) {
                    // Continue with the operations
                    // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    spinner.start();

                    shell.cd('./Word-Add-in-AIGC');
                    shell.exec('npm install', {async:true}, (code, stdout, stderr) => {
                        shell.exec('npm run start', {async:true}, (code, stdout, stderr) => {

                        spinner.stop(true);
                        readline.clearLine(process.stdout, 0);
                        readline.cursorTo(process.stdout, 0);

                        console.log('Step [3/3] completed!');
                        console.log('Finished!');
                        resolve(is_vscode_installed);
                        });
                    });
                }
                else{
                    // Don't continue with the operations
                    console.log('Step [3/3] skipped. You decided not to auto-launch the sample.')
                    console.log('No problem. You can always launch the sample add-in by running the following commands:');
                    console.log('--------------------------------------------');
                    console.log('npm install');
                    console.log('npm run start');
                    console.log('--------------------------------------------');    
                    console.log('Finished!');             
                    resolve(is_vscode_installed);
                }
            });
          });
      });
  }

function replaceUrl(url: string, newUrl: string, filePath: string) {
    fs.readFile(filePath, 'utf8', (err, data) => {
        if (err) {
            console.error(err);
            return;
        }

        const result = data.replace(url, newUrl);

        fs.writeFile(filePath, result, 'utf8', (err) => {
            if (err) {
                console.error(err);
                return;
            }
        });
    });
}

module.exports = { exec_script_Excel_Mail, exec_script_Word_AIGC };


