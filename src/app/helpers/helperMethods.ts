const path = require('path');
const fs = require('fs');
const hosts = [
    "excel",
    "onenote",
    "outlook",
    "powerpoint",
    "project",
    "word"
];

export namespace helperMethods {
    export function deleteFolderRecursively(projectFolder: string) 
    {
        if(fs.existsSync(projectFolder))
        {
            fs.readdirSync(projectFolder).forEach(function(file,index){ 
            var curPath = projectFolder + "/" + file; 
            
            if(fs.lstatSync(curPath).isDirectory())
            {
                deleteFolderRecursively(curPath);
            }
            else
            {
                fs.unlinkSync(curPath);
            }
        }); 
        fs.rmdirSync(projectFolder); 
        }
    }

    export function doesProjectFolderExists(projectFolder: string)
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

    export function cleanupProjectFolder(projectFolder: string, host: string, typescript: boolean) {
        return new Promise(async (resolve, reject) => {
            try {
                await _updateMoveProjectFiles(projectFolder, host, typescript);
                await _modifyPackageJsonFile(projectFolder, host);
                return resolve();
            } catch (err){
                return reject(err);
            }
        });
    }

    function _updateMoveProjectFiles(projectFolder: string, host: string, typescript: boolean) {
        return new Promise(async (resolve, reject) => {
            // copy host-specific manifest over manifest.xml
            fs.readFile(path.resolve(`${projectFolder}/manifest.${host}.xml`), 'utf8', (err, contents) => {
                if (err) return reject(err);
                fs.writeFile(path.resolve(`${projectFolder}/manifest.xml`), contents, (err) => {
                    if (err) return reject(err);
                });
            });

            // copy host-specific taskpane.ts over src/taskpane/taskpane.ts[js]
            fs.readFile(path.resolve(`${projectFolder}/src/taskpane/${host}.${typescript ? 'ts' : 'js'}`), 'utf8', (err, contents) => {
                if (err) return reject(err);
                fs.writeFile(path.resolve(`${projectFolder}/src/taskpane/taskpane.${typescript ? 'ts' : 'js'}`), contents, (err) => {
                    if (err) return reject(err);
                });
            });

            // delete all host specific files
            hosts.forEach(function (host) {
                fs.unlink(path.resolve(`${projectFolder}/manifest.${host}.xml`), (err) => {
                    if (err) return reject(err);
                });
                fs.unlink(path.resolve(`${projectFolder}/src/taskpane/${host}.${typescript ? 'ts' : 'js'}`), (err) => {
                    if (err) return reject(err);
                });
            });
            return resolve();
        });
    }

    function _modifyPackageJsonFile(projectFolder:string, host: string) {
        return new Promise(async (resolve, reject) => {
            // update package.json to reflect selected host
            const packageJson = path.resolve(`${projectFolder}/package.json`);
            fs.readFile(packageJson, 'utf8', (err, data) => {
                if (err) return reject(err);
                let content = JSON.parse(data);

                // update 'config' section in package.json to use selected host
                content.config["app-to-debug"] = host;

                // remove scripts from package.json that are unrelated to selected host,
                // and update sideload and unload scripts to use selected host.
                Object.keys(content.scripts).forEach(function (key) {
                    if (key.includes("sideload:") || key.includes("unload:")) {
                        delete content.scripts[key];
                    }

                    if (key == "sideload") {
                        content.scripts[key] = `office-toolbox sideload -m manifest.xml -a ${host}`
                    }

                    if (key == "unload") {
                        content.scripts[key] = `office-toolbox remove -m manifest.xml -a ${host}`
                    }
                });

                // write updated json to file
                fs.writeFile(packageJson, JSON.stringify(content, null, 4), (err) => {
                    if (err) return reject(err);
                    return resolve();
                });
            });
        });
    }
}