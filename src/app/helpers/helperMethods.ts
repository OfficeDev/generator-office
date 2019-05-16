import * as _ from 'lodash';
import * as fsextra from "fs-extra";
import * as unzip from "unzip";
import * as util from "util";
const fs = require('fs');
const readFileAsync = util.promisify(fs.readFile);
const unlinkFileAsync = util.promisify(fs.unlink);
const writeFileAsync = util.promisify(fs.writeFile);
const path = require('path');
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

    export function doesProjectFolderExist(projectFolder: string)
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

    export async function unzipProjectTemplate(projectFolder: string) {
        return new Promise(async (resolve, reject) => {
            const zipFile = 'project.zip';
            const readStream = fs.createReadStream(`${projectFolder}/${zipFile}`);
            readStream.pipe(unzip.Extract({ path: projectFolder }))
                .on('close', async () => {
                    await moveProjectFiles(projectFolder);
                    return resolve();
                });
            readStream.on('error', function (err) {
                return reject(err);
            })
        });
    }

    export async function moveProjectFiles(projectFolder: string) {
        return new Promise(async (resolve, reject) => {

            // delete original zip file
            const zipFile = path.join(`${projectFolder}/project.zip`);
            if (fs.existsSync(zipFile)) {
                await unlinkFileAsync(zipFile);
            }

            // get the name of the unzipped project folder
            fs.readdir(projectFolder, function (err, folder) {
                const unzippedFolder = folder;
            });

            const moveFrom = `${projectFolder}/Office-Addin-Taskpane-master`;
            var moveTo = projectFolder;

            // Loop through all the files in the temp directory
            fs.readdir(moveFrom, function (err, files) {
                if (err) {
                    reject(err);
                }

                files.forEach(function (file, index) {
                    // Make one pass and make the file complete
                    var fromPath = path.join(moveFrom, file);
                    var toPath = path.join(moveTo, file);

                    fs.exists(fromPath, (exists) => {
                        if (exists) {
                            fs.rename(fromPath, toPath, function (error) {
                                if (error) {
                                    reject(error);
                                }
                            });
                        }
                        else reject();
                    });
                    resolve();
                });
            });
        });
    }

    export function modifyProjectForSingleHost(projectFolder: string, projectType: string, host: string, typescript: boolean) {
        return new Promise(async (resolve, reject) => {
            try {
                await convertProjectToSingleHost(projectFolder, projectType, host, typescript);
                await updatePackageJsonForSingleHost(projectFolder, host);
                return resolve();
            } catch (err){
                return reject(err);
            }
        });
    }

    async function convertProjectToSingleHost(projectFolder: string, projectType: string, host: string, typescript: boolean): Promise<void> {        
        try {
            let extension = typescript ? "ts" : "js";
            // copy host-specific manifest over manifest.xml
            const manifestContent: any = await readFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`), 'utf8');
            await writeFileAsync(path.resolve(`${projectFolder}/manifest.xml`), manifestContent);

            switch (projectType) {
                case "taskpane":
                {
                    // copy host-specific taskpane.ts[js] over src/taskpane/taskpane.ts[js]
                    const srcContent = await readFileAsync(path.resolve(`${projectFolder}/src/taskpane/${host}.${extension}`), 'utf8');
                    await writeFileAsync(path.resolve(`${projectFolder}/src/taskpane/taskpane.${extension}`), srcContent);

                    // delete all host specific files
                    hosts.forEach(async function (host) {
                        await unlinkFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`));
                        await unlinkFileAsync(path.resolve(`${projectFolder}/src/taskpane/${host}.${extension}`));
                    });
                    break;
                }
                case "angular":
                {
                    // copy host-specific app.component.ts[js] over src/taskpane/app/app.component.ts[js]
                    const srcContent = await readFileAsync(path.resolve(`${projectFolder}/src/taskpane/app/${host}.app.component.${extension}`), 'utf8');
                    await writeFileAsync(path.resolve(`${projectFolder}/src/taskpane/app/app.component.${extension}`), srcContent);

                    // delete all host specific files
                    hosts.forEach(async function (host) {
                        await unlinkFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`));
                        await unlinkFileAsync(path.resolve(`${projectFolder}/src/taskpane/app/${host}.app.component.${extension}`));
                    });
                    break;
                }
                case "react":
                {
                    // copy host-specific App.tsx[js] over src/taskpane/app/components/App.tsx[js]
                    extension = typescript ? "tsx" : "js";
                    const srcContent = await readFileAsync(path.resolve(`${projectFolder}/src/taskpane/components/${_.upperFirst(host)}.App.${extension}`), 'utf8');
                    await writeFileAsync(path.resolve(`${projectFolder}/src/taskpane/components/App.${extension}`), srcContent);

                    // delete all host specific files
                    hosts.forEach(async function (host) {
                        await unlinkFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`));
                        await unlinkFileAsync(path.resolve(`${projectFolder}/src/taskpane/components/${_.upperFirst(host)}.App.${extension}`));
                    });
                    break;
                }
                default:
                    throw new Error("Invalid project type");
            }
        } catch(err) {
            throw new Error(err);
        }
    }

    async function updatePackageJsonForSingleHost(projectFolder:string, host: string): Promise<void> {
        try {
            // update package.json to reflect selected host
            const packageJson = path.resolve(`${projectFolder}/package.json`);
            const data: any = await readFileAsync(packageJson, 'utf8');
            let content = JSON.parse(data);

            // update 'config' section in package.json to use selected host
            content.config["app-to-debug"] = host;

            // remove scripts from package.json that are unrelated to selected host,
            // and update sideload and unload scripts to use selected host.
            Object.keys(content.scripts).forEach(function (key) {
                if (key.includes("sideload:") || key.includes("unload:")) {
                    delete content.scripts[key];
                }
                switch (key) {
                    case "sideload":
                    case "unload":
                        content.scripts[key] = content.scripts[`${key}:${host}`];
                        break;
                }
            });

            // write updated json to file
            await writeFileAsync(packageJson, JSON.stringify(content, null, 4));
        } catch (err) {
            throw new Error(err);
        }
    }
}