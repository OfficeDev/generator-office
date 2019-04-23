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

    export function modifyProjectForSingleHost(projectFolder: string, host: string, typescript: boolean) {
        return new Promise(async (resolve, reject) => {
            try {
                await convertProjectToSingleHost(projectFolder, host, typescript);
                await updatePackageJsonForSingleHost(projectFolder, host);
                return resolve();
            } catch (err){
                return reject(err);
            }
        });
    }

    async function convertProjectToSingleHost(projectFolder: string, host: string, typescript: boolean): Promise<void> {        
        try {
            // copy host-specific manifest over manifest.xml
            const manifestContent: any = await readFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`), 'utf8');
            await writeFileAsync(path.resolve(`${projectFolder}/manifest.xml`), manifestContent);

            // copy host-specific taskpane.ts over src/taskpane/taskpane.ts[js]
            const srcContent = await readFileAsync(path.resolve(`${projectFolder}/src/taskpane/${host}.${typescript ? 'ts' : 'js'}`), 'utf8');
            await writeFileAsync(path.resolve(`${projectFolder}/src/taskpane/taskpane.${typescript ? 'ts' : 'js'}`), srcContent);

            // delete all host specific files
            hosts.forEach(async function (host) {
                await unlinkFileAsync(path.resolve(`${projectFolder}/manifest.${host}.xml`));
                await unlinkFileAsync(path.resolve(`${projectFolder}/src/taskpane/${host}.${typescript ? 'ts' : 'js'}`));
            });
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