const shell = require('shelljs');
const ProgressBar = require('progress');
const log = require('single-line-log').stdout;
const Spinner = require('cli-spinner').Spinner;
const { spawn } = require('child_process');
const { exec } = require('child_process');
const { execSync } = require('child_process');
const readline = require('readline');


shell.config.silent = true;

async function exec_script(){
  // shell.exec('code .');
  console.log('Welcome to experience this Office add-in sample!');

  return new Promise<void>((resolve, reject) => {

        console.log('Welcome to experience this Office add-in sample!');

        // Step 1: Get sample code
        console.log('Step 1: Getting sample code...');
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
            console.log('Step 1 completed!');
            console.log('Step 2: Checking if Visual Studio Code is installed...');
            if (shell.which('code')) {
                console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
                shell.exec('code ./Samples/Excel.OfflineStorageAddin README.md');
            } else {
                console.log('Visual Studio Code is not installed on your machine.');
                shell.exec('start Samples\\Excel.OfflineStorageAddin');
            }

            console.log('Step 2 completed!');

            // Step 3: Provide user the command to side-load add-in directly 
            console.log('Step 3: Automatically side-load add-in directly...');
            spinner.start();

            shell.cd('./Samples/Excel.OfflineStorageAddin');
            shell.exec('npm install', {async:true}, (code, stdout, stderr) => {
                shell.exec('npm run start', {async:true}, (code, stdout, stderr) => {

                spinner.stop(true);
                readline.clearLine(process.stdout, 0);
                readline.cursorTo(process.stdout, 0);

                console.log('Step 3 completed!');
                console.log('Finished!');
                resolve();
                });
            });
            });
        });

        // //Step 1: Get sample code
        // console.log('Step 1: Getting sample code...');

        // let spinner = new Spinner('Processing.. %s');
        // spinner.setSpinnerString('|/-\\');
        // spinner.start();

        // const gitClone = spawn('git', ['clone', '--depth', '1', '--filter=blob:none', '--sparse', 'https://github.com/OfficeDev/Office-Add-in-samples.git', './Office_add_in_sample'], { stdio: 'ignore' });

        // gitClone.on('close', (code) => {
        // const gitSparseCheckout = spawn('git', ['sparse-checkout', 'set', 'Samples/Excel.OfflineStorageAddin/'], { cwd: './Office_add_in_sample', stdio: 'ignore' });

        // gitSparseCheckout.on('close', (code) => {
        //     spinner.stop(true);
        //     readline.clearLine(process.stdout, 0);
        //     readline.cursorTo(process.stdout, 0);

        //     // Step 2: Check if VSCode is installed
        //     console.log('Step 1 completed!');
        //     console.log('Step 2: Checking if Visual Studio Code is installed...');
        //     shell.cd('./Office_add_in_sample');
        //     const vscodeCheck = spawn('code', ['--version']);

        //     vscodeCheck.stdout.on('data', (data) => {
        //     console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
        //     shell.exec('code ./Samples/Excel.OfflineStorageAddin README.md');
        //     });

        //     vscodeCheck.stderr.on('data', (data) => {
        //     console.log('Visual Studio Code is not installed on your machine.');
        //     shell.exec('start Samples\\Excel.OfflineStorageAddin');
        //     });

        //     vscodeCheck.on('close', (code) => {
        //     console.log('Step 2 completed!');

        //     // Step 3: Provide user the command to side-load add-in directly 
        //     console.log('Step 3: Automatically side-load add-in directly...');
        //     spinner.start();

        //     const npmInstall = spawn('npm', ['install'], { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore' });

        //     npmInstall.on('close', (code) => {
        //         const npmStart = spawn('npm', ['run', 'start'], { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore' });

        //         npmStart.on('close', (code) => {
        //         spinner.stop(true);
        //         readline.clearLine(process.stdout, 0);
        //         readline.cursorTo(process.stdout, 0);

        //         console.log('Step 3 completed!');
        //         console.log('Finished!');

        //         resolve();
        //         });
        //     });
        //     });
        // });
        // });

        //     exec('git init ./Office_add_in_sample', (error, stdout, stderr) => {
                
        //         exec('git remote add -f origin https://github.com/OfficeDev/Office-Add-in-samples.git && git sparse-checkout init && git sparse-checkout set Samples/Excel.OfflineStorageAddin/ && git pull origin main', { cwd: './Office_add_in_sample', stdio: 'ignore'}, (error, stdout, stderr) => {
        //             spinner.stop(true);
        //             readline.clearLine(process.stdout, 0);
        //             readline.cursorTo(process.stdout, 0);
        //             //Step 2: Check if VSCode is installed
        //             console.log('Step 1 completed!');
        //             console.log('Step 2: Checking if Visual Studio Code is installed...');
        //             shell.cd('./Office_add_in_sample');
        //             shell.exec('code --version', (error, stdout, stderr) => {
        //                 if (error) {
        //                 console.error(`exec error: ${error}`);
        //                 return;
        //                 }
                    
        //                 if (stdout) {
        //                 console.log('Visual Studio Code is installed on your machine. Would open in VSCode for exploring the code.');
        //                 shell.exec('code ./Samples/Excel.OfflineStorageAddin README.md');
        //                 // shell.exec('start Samples\\Excel.OfflineStorageAddin');
        //                 } else {
        //                 console.log('Visual Studio Code is not installed on your machine.');
        //                 shell.exec('start Samples\\Excel.OfflineStorageAddin');
        //                 }

        //                 console.log('Step 2 completed!');
                    
        //                 //Step 3: Provide user the command to side-load add-in directly 
        //                 console.log('Step 3: Automatically side-load add-in directly...');

        //                 let spinner = new Spinner('Processing.. %s');
        //                 spinner.setSpinnerString('|/-\\');
        //                 spinner.start();

        //                 const npmInstall = spawn('npm', ['install'], { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore' });

        //                 npmInstall.on('close', (code) => {
        //                     const npmStart = spawn('npm', ['run', 'start'], { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore' });

        //                     npmStart.on('close', (code) => {
        //                         spinner.stop(true);
        //                         readline.clearLine(process.stdout, 0);
        //                         readline.cursorTo(process.stdout, 0);

        //                         console.log('Step 3 completed!');
        //                         console.log('Finished!');
        //                     }); 
        //                 });

        //                 // exec('npm install', { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore'},(error, stdout, stderr) => {

        //                 //     execSync('npm run start', { cwd: './Office_add_in_sample/Samples/Excel.OfflineStorageAddin', stdio: 'ignore'},(error, stdout, stderr) => {

        //                 //         spinner.stop(true);
        //                 //         readline.clearLine(process.stdout, 0);
        //                 //         readline.cursorTo(process.stdout, 0);

        //                 //         resolve();
        //                 //         console.log('Step 3 completed!');
        //                 //         console.log('Finished!');
        //                 //     });

                            
        //                 // })

        //                 // shell.cd("Samples/Excel.OfflineStorageAddin");
        //                 // shell.exec('npm install',(code, stdout, stderr) => {
        //                 //     shell.exec('npm run start');
        //                 //     resolve();

        //                 //     console.log('Step 3 completed!');
        //                 //     console.log('Finished!');
        //                 // });
        //             });
        //         }) 

            
        // });
    });
}

// exec_script();

module.exports = { exec_script };


