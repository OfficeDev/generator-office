import * as path from "path";
import * as request from "request";
import * as unzip from "unzipper";
const fs = require('fs');
const zipFile = 'project.zip';

export namespace helperMethods {
    function deleteFolderRecursively(projectFolder: string) {
        try {
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
        } catch (err) {
            throw new Error(err);
        }
    }

    export function doesProjectFolderExist(projectFolder: string) {      
    if (fs.existsSync(projectFolder))
        {
            if (fs.readdirSync(projectFolder).length > 0)
            {          
                return true;
            }
        }
        return false;
    };

    export async function downloadProjectTemplate(projectFolder: string, projectRepo: string, projectBranch: string): Promise<void> {
        return new Promise(async (resolve, reject) => {
            try {
                await request(`${projectRepo}/archive/${projectBranch}.zip`)
                .pipe(fs.createWriteStream(zipFile))
                .on('error', function () {
                    throw new Error("unable to download project zip file")
                })
                .on('close', async (err) => {
                  await unzipProjectTemplate(projectFolder);
                  resolve();
            });
        } catch (err) {
            reject(err);
        }
    });
    }

    async function unzipProjectTemplate(projectFolder: string): Promise<void> {
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

    async function moveProjectFiles(projectFolder: string): Promise<void> {
        // delete original zip file
        const zipFilePath = path.resolve(`${projectFolder}/${zipFile}`);
        if (fs.existsSync(zipFilePath)) {
            fs.unlinkSync(zipFilePath);
        }

        // get path to unzipped folder
        const unzippedFolder = fs.readdirSync(projectFolder).filter(function (file) {
            return fs.statSync(`${projectFolder}/${file}`).isDirectory();
          });

        // construct paths to move files out of unzipped folder into project root folder
        const moveFromFolder = path.resolve(`${projectFolder}/${unzippedFolder[0]}`);
        const moveToFolder = projectFolder;

        // loop through all the files and folders in the unzipped folder and move them to project root
        fs.readdirSync(moveFromFolder).forEach(function(file) { 
            var fromPath = path.join(moveFromFolder, file);
            var toPath = path.join(moveToFolder, file);

            if (fs.existsSync(fromPath) && !fromPath.includes("gitignore")) {
                fs.renameSync(fromPath, toPath);
            }
        });

        // delete project zipped folder
        deleteFolderRecursively(moveFromFolder);
    }
}