import axios from "axios"
import * as fs from "fs";
import * as path from "path";
import * as unzip from "unzipper";

const zipFile = 'project.zip';

// eslint-disable-next-line @typescript-eslint/no-namespace
export namespace helperMethods {
    function deleteFolderRecursively(projectFolder: string) {
        try {
            if (fs.existsSync(projectFolder)) {
                fs.readdirSync(projectFolder).forEach(function (file) {
                    const curPath = `${projectFolder}/${file}`;

                    if (fs.lstatSync(curPath).isDirectory()) {
                        deleteFolderRecursively(curPath);
                    }
                    else {
                        fs.unlinkSync(curPath);
                    }
                });
                fs.rmdirSync(projectFolder);
            }
        } catch (err) {
            throw new Error(`Unable to delete folder "${projectFolder}".\n${err}`);
        }
    }

    export function doesProjectFolderExist(projectFolder: string) {
        if (fs.existsSync(projectFolder)) {
            return fs.readdirSync(projectFolder).length > 0;
        }
        return false;
    };

    export async function downloadProjectTemplateZipFile(projectFolder: string, projectRepo: string, projectBranch: string): Promise<void> {
        const projectTemplateZipFile = `${projectRepo}/archive/${projectBranch}.zip`;
        return axios({
            method: 'get',
            url: projectTemplateZipFile,
            responseType: 'stream',
        }).then(response => {
            return new Promise<void>((resolve, reject) => {
                response.data.pipe(fs.createWriteStream(zipFile))
                .on('error', function (err) {
                    reject(`Unable to download project zip file for "${projectTemplateZipFile}".\n${err}`);
                })
                .on('close', async () => {
                    await unzipProjectTemplate(projectFolder);
                    resolve();
                });
            });
        }).catch(err => { console.log(`Unable to download project zip file for "${projectTemplateZipFile}".\n${err}`); });
    }

    async function unzipProjectTemplate(projectFolder: string): Promise<void> {
        return new Promise(async (resolve, reject) => {
            const zipFile = 'project.zip';
            const readStream = fs.createReadStream(`${projectFolder}/${zipFile}`);
            readStream.pipe(unzip.Extract({ path: projectFolder }))
                .on('error', function (err) {
                    reject(`Unable to unzip project zip file for "${projectFolder}".\n${err}`);
                })
                .on('close', async () => {
                    moveProjectFiles(projectFolder);
                    resolve();
                });
        });
    }

    function moveProjectFiles(projectFolder: string): void {
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

        // loop through all the files and folders in the unzipped folder and move them to project root
        fs.readdirSync(moveFromFolder).forEach(function (file) {
            const fromPath = path.join(moveFromFolder, file);
            const toPath = path.join(projectFolder, file);

            if (fs.existsSync(fromPath) && !fromPath.includes(".gitignore")) {
                fs.renameSync(fromPath, toPath);
            }
        });

        // delete project zipped folder
        deleteFolderRecursively(moveFromFolder);
    }
}