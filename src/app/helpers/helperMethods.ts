import axios from "axios"
import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";

const zipFile = 'project.zip';

// eslint-disable-next-line @typescript-eslint/no-namespace
export namespace helperMethods {
    export function deleteFolderRecursively(projectFolder: string) {
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

    export async function downloadProjectTemplateZipFile(projectFolder: string, projectRepo: string, projectBranch: string): Promise<string> {
        const projectTemplateZipFile = `${projectRepo}/archive/${projectBranch}.zip`;
        return axios({
            method: 'get',
            url: projectTemplateZipFile,
            responseType: 'stream',
        }).then(response => {
            return new Promise<string>((resolve, reject) => {
                response.data.pipe(fs.createWriteStream(zipFile))
                .on('error', function (err) {
                    reject(`Unable to download project zip file for "${projectTemplateZipFile}".\n${err}`);
                })
                .on('close', async () => {
                    resolve(path.resolve(`${projectFolder}/project.zip`));
                });
            });
        }).catch(err => {
            const error: string = `Unable to download project zip file for "${projectTemplateZipFile}".\n${err}`;
            console.log(error)
            return Promise.reject(error); 
        });
    }

    export async function unzipProjectTemplate(projectFolder: string): Promise<string> {
        return new Promise(async (resolve, reject) => {
            const zipFile = 'project.zip';
            const zip = new AdmZip(`${projectFolder}/${zipFile}`);
            try {
                zip.extractAllTo(/*target path*/projectFolder, /*overwrite*/true);
                // get path to unzipped folder
                const unzippedFolder = fs.readdirSync(projectFolder).filter(function (file) {
                    return fs.statSync(`${projectFolder}/${file}`).isDirectory();
                });
                resolve(unzippedFolder[0]);
            } catch (err) {
                reject(`Unable to unzip project zip file for "${projectFolder}".\n${err}`);
            }
        });
    }
}