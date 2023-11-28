const shell = require('shelljs');
const ProgressBar = require('progress');
const log = require('single-line-log').stdout;
const Spinner = require('cli-spinner').Spinner;
const { spawn } = require('child_process');
const { exec } = require('child_process');
const { execSync } = require('child_process');
const readline = require('readline');
const os = require('os');
const fs = require('fs');

const homeDirectory = os.homedir();

shell.config.silent = true;

async function exec_script_hello_world_excel(): Promise<boolean>{
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

        shell.exec('git clone --depth 1 --filter=blob:none --sparse https://github.com/OfficeDev/Office-Add-in-samples.git ./Hello_World_sample_Excel', {async:true}, (code, stdout, stderr) => {
            shell.cd('./Hello_World_sample_Excel');
            shell.exec('git sparse-checkout set Samples/hello-world/excel-hello-world/', {async:true}, (code, stdout, stderr) => {

            spinner.stop(true);
            readline.clearLine(process.stdout, 0);
            readline.cursorTo(process.stdout, 0);

            // Step 2: Check if VSCode is installed
            console.log('Step [1/3] completed!');
            console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
            if (shell.which('code')) {
                console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                is_vscode_installed = true;
                shell.exec('code ./Samples/hello-world/excel-hello-world README.md');
            } else {
                console.log('Visual Studio Code is not installed on your machine.');
                shell.exec('start Samples\\hello-world\\excel-hello-world');
            }

            console.log('Step [2/3] completed!');
            
            // Ask user if sample Add-in automatic launch is needed
            let rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            console.log(homeDirectory);
            rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                if (answer.trim().toLowerCase() == 'y') {
                  // Continue with the operations
                  // Step 3: Provide user the command to side-load add-in directly 
                    console.log('Step [3/3]: Automatically launch add-in directly...');
                    spinner.start();
                    const crt_path = '.office-addin-dev-certs\\localhost.crt';
                    const key_path = '.office-addin-dev-certs\\localhost.key';
                    const full_path_crt = homeDirectory + '\\' + crt_path;
                    const full_path_key = homeDirectory + '\\' + key_path;

                    shell.cd('./Samples/hello-world/excel-hello-world');
                    shell.exec('npm install -g office-addin-debugging', {async:true}, (code, stdout, stderr) => {
                        shell.exec('office-addin-debugging start manifest.xml', {async:true}, (code, stdout, stderr) => {

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
    });
}

async function exec_script_hello_world_word(): Promise<boolean>{
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
  
          shell.exec('git clone --depth 1 --filter=blob:none --sparse https://github.com/OfficeDev/Office-Add-in-samples.git ./Hello_World_sample_Word', {async:true}, (code, stdout, stderr) => {
              shell.cd('./Hello_World_sample_Word');
              shell.exec('git sparse-checkout set Samples/hello-world/word-hello-world/', {async:true}, (code, stdout, stderr) => {
  
              spinner.stop(true);
              readline.clearLine(process.stdout, 0);
              readline.cursorTo(process.stdout, 0);
  
              // Step 2: Check if VSCode is installed
              console.log('Step [1/3] completed!');
              console.log('Step [2/3]: Checking if Visual Studio Code is installed...');
              if (shell.which('code')) {
                  console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                  is_vscode_installed = true;
                  shell.exec('code ./Samples/hello-world/word-hello-world README.md');
              } else {
                  console.log('Visual Studio Code is not installed on your machine.');
                  shell.exec('start Samples\\hello-world\\word-hello-world');
              }
  
              console.log('Step [2/3] completed!');
              
              // Ask user if sample Add-in automatic launch is needed
              let rl = readline.createInterface({
                  input: process.stdin,
                  output: process.stdout
              });
  
              console.log(homeDirectory);
              rl.question('Do you want to continue with some operations? (Y/N)\n', (answer) => {
                  if (answer.trim().toLowerCase() == 'y') {
                    // Continue with the operations
                    // Step 3: Provide user the command to side-load add-in directly 
                      console.log('Step [3/3]: Automatically launch add-in directly...');
                      spinner.start();
                      const crt_path = '.office-addin-dev-certs\\localhost.crt';
                      const key_path = '.office-addin-dev-certs\\localhost.key';
                      const full_path_crt = homeDirectory + '\\' + crt_path;
                      const full_path_key = homeDirectory + '\\' + key_path;
  
                      shell.cd('./Samples/hello-world/word-hello-world');
                      shell.exec('npm install -g office-addin-debugging', {async:true}, (code, stdout, stderr) => {
                          shell.exec('office-addin-debugging start manifest.xml', {async:true}, (code, stdout, stderr) => {
  
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
      });
  }
  

// exec_script();

module.exports = { exec_script_hello_world_excel , exec_script_hello_world_word };


