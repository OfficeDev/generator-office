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

async function exec_script(){
  // shell.exec('code .');
  console.log('Welcome to experience this Office add-in sample!');

  return new Promise<boolean>((resolve, reject) => {
    
        let is_vscode_installed = false;

        console.log('Welcome to experience this Office add-in sample!');

        // Step 1: Get sample code
        console.log('Step [1/3]: Getting sample code...');
        let spinner = new Spinner('Processing.. %s');
        spinner.setSpinnerString('|/-\\');
        spinner.start();

        shell.exec('git clone --depth 1 --filter=blob:none --sparse https://github.com/OfficeDev/Office-Add-in-samples.git ./Office_add_in_sample', {async:true}, (code, stdout, stderr) => {
            shell.cd('./Office_add_in_sample');
            shell.exec('git sparse-checkout set Samples/Excel.OfflineStorageAddin/', {async:true}, (code, stdout, stderr) => {

            spinner.stop(true);
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0);

            // Step 2: Check if VSCode is installed
            console.log('Step [1/3] completed!');
            console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
            if (shell.which('code')) {
                console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                is_vscode_installed = true;
                shell.exec('code ./Samples/Excel.OfflineStorageAddin README.md');
            } else {
                console.log('Visual Studio Code is not installed on your machine.');
                shell.exec('start Samples\\Excel.OfflineStorageAddin');
            }

            console.log('Step [2/3] completed!');
            
            // Ask user if sample Add-in automatic launch is needed
            let rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            let auto_launch_answer = false;
            rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                console.log(`Your input was: ${answer}`);
                if (answer.trim().toLowerCase() == 'y') {
                  // Continue with the operations
                  // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically side-load add-in directly...');
                    spinner.start();

                    shell.cd('./Samples/Excel.OfflineStorageAddin');
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
                

            // resolve(is_vscode_installed);

            // if (!auto_launch_answer) {
            //     resolve(is_vscode_installed);
            // }
            });
        });
    });
}

// exec_script();

module.exports = { exec_script };


