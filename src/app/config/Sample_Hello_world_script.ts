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


shell.config.silent = true;

async function exec_script_Excel_Hello_World(){
  // shell.exec('code .');
  console.log('Welcome to experience this Office add-in sample: Excel Hello World Add-in!');

  return new Promise<boolean>((resolve, reject) => {
    
        let is_vscode_installed = false;

        console.log('Welcome to experience this Office add-in sample!');

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', {async:true}, (code, stdout, stderr) => {
            shell.cd('./Office-Addin-TaskPane-React');
            shell.exec('npm run convert-to-single-host --if-present --Excel', {async:true}, (code, stdout, stderr) => {

            spinner.stop(true);
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0);

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
                    shell.exec('start .');
                }

                console.log('Step [2/3] completed!');
                if (auto_launch_answer) {
                    // Continue with the operations
                    // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    console.log('It may take longer time to complete the process. Please wait patiently...');
                    spinner.start();

                    // shell.cd('./Mail-Merge-Sample-Add-in');
                    shell.exec('npm install', {async:true}, (code, stdout, stderr) => {
                        shell.exec('npm run start', {async:true}, (code, stdout, stderr) => {

                        spinner.stop(true);
                        readline.clearLine(process.stdout, 0);
                        readline.cursorTo(process.stdout, 0);

                        console.log('Step [3/3] completed!');
                        console.log('Finished!');
                        // console.log('Hint: To try out the full functionality, please follow the instruction in the opening web page: Register a web application with the Azure Active Directory admin center.');
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
                    resolve(is_vscode_installed);
                }
            });
        });
        });
    });
}

async function exec_script_Word_Hello_World(){
    // shell.exec('code .');
    console.log('Welcome to experience this Office add-in sample: Word Hello World Add-in project!');
  
    return new Promise<boolean>((resolve, reject) => {
      
          let is_vscode_installed = false;
  
          console.log('Welcome to experience this Office add-in sample!');
  
          // Step 1: Get sample code
          console.log('Step [1/3]: Getting sample code...');
          let spinner = new Spinner('Processing.. %s');
          spinner.setSpinnerString('|/-\\');
          spinner.start();
  
          shell.exec('git clone git clone https://github.com/OfficeDev/Office-Addin-TaskPane-React.git', {async:true}, (code, stdout, stderr) => {
              shell.cd('./Office-Addin-TaskPane-React');
              // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
  
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
                    shell.exec('start .');
                }

                console.log('Step [2/3] completed!');
                if (auto_launch_answer) {
                    // Continue with the operations
                    // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    spinner.start();

                    // shell.cd('./Word-Add-in-AIGC');
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

// exec_script();

module.exports = { exec_script_Excel_Hello_World, exec_script_Word_Hello_World };