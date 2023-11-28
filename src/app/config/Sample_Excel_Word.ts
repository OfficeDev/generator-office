const shell = require('shelljs');
const ProgressBar = require('progress');
const log = require('single-line-log').stdout;
const Spinner = require('cli-spinner').Spinner;
const { spawn } = require('child_process');
const { exec } = require('child_process');
const { execSync } = require('child_process');
const readline = require('readline');

import * as os from 'os';


shell.config.silent = true;

async function exec_script_Excel_Mail(){
  // shell.exec('code .');
  console.log('Welcome to experience this Office add-in sample: Excel Mail merge Add-in project!');

  return new Promise<boolean>((resolve, reject) => {
    
        let is_vscode_installed = false;

        console.log('Welcome to experience this Office add-in sample!');

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

            // Step 2: Check if VSCode is installed
            console.log('Step [1/3] completed!');
            console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
            if (shell.which('code')) {
                console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                is_vscode_installed = true;
                shell.exec('code -g README.md:22');
            } else {
                console.log('Visual Studio Code is not installed on your machine.');
                shell.exec('start Mail-Merge-Sample-Add-in');
            }

            console.log('Step [2/3] completed!');
            
            // Ask user if sample Add-in automatic launch is needed
            let rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            let auto_launch_answer = false;
            rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                if (answer.trim().toLowerCase() == 'y') {
                  // Continue with the operations
                  // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    spinner.start();

                    shell.cd('./Mail-Merge-Sample-Add-in');
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

                  
                } else {
                  // Don't continue with the operations
                    console.log('No problem. You can always launch the sample add-in by running the following commands:');
                    
                    resolve(is_vscode_installed);
                }
                rl.close();
            });
        });
    });
}

async function exec_script_Word_AIGC(){
    // shell.exec('code .');
    console.log('Welcome to experience this Office add-in sample: Word AIGC Add-in project!');
  
    return new Promise<boolean>((resolve, reject) => {
      
          let is_vscode_installed = false;
  
          console.log('Welcome to experience this Office add-in sample!');
  
          // Step 1: Get sample code
          console.log('Step [1/3]: Getting sample code...');
          let spinner = new Spinner('Processing.. %s');
          spinner.setSpinnerString('|/-\\');
          spinner.start();
  
          shell.exec('git clone https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples.git', {async:true}, (code, stdout, stderr) => {
              shell.cd('./Word-Scenario-based-Add-in-Samples/Word-Add-in-AIGC');
              // shell.exec('git sparse-checkout set Mail-Merge-Sample-Add-in/', {async:true}, (code, stdout, stderr) => {
  
              spinner.stop(true);
              readline.clearLine(process.stdout, 0);
              readline.cursorTo(process.stdout, 0);
  
              // Step 2: Check if VSCode is installed
              console.log('Step [1/3] completed!');
              console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
              if (shell.which('code')) {
                  console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                  is_vscode_installed = true;
                  shell.exec('code -g README.md:22');
              } else {
                  console.log('Visual Studio Code is not installed on your machine.');
                  shell.exec('start Word-Add-in-AIGC');
              }
  
              console.log('Step [2/3] completed!');
              
              // Ask user if sample Add-in automatic launch is needed
              let rl = readline.createInterface({
                  input: process.stdin,
                  output: process.stdout
              });

              rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                  if (answer.trim().toLowerCase() == 'y') {
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

                    
                  } else {
                    // Don't continue with the operations
                      console.log('No problem. You can always launch the sample add-in by running the following commands:');
                      
                      resolve(is_vscode_installed);
                  }
                  rl.close();
              });
          });
      });
  }

// exec_script();

module.exports = { exec_script_Excel_Mail, exec_script_Word_AIGC };


