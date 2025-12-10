import axios from "axios"
import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";
import debug from "debug";
import { AttemptAwareConfig, hasProxy, addProxy, addLogger, addInterceptor } from "./requestHelpers.js";

const log = debug("genOffice").extend("helper");

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
    }


    export async function downloadProjectTemplateZipFile(projectFolder: string, projectRepo: string, projectBranch: string): Promise<string> {
        const useProxyFirst = process.env.GENERATOR_OFFICE_USE_PROXY === "true";
        const projectTemplateZipFile = `${projectRepo}/archive/${projectBranch}.zip`;
        log("Setting up config for %s", projectTemplateZipFile);
        const config : AttemptAwareConfig ={
            method: 'get', 
            url: projectTemplateZipFile, 
            responseType: 'stream'
        };
        if(useProxyFirst && hasProxy()) {
            config.useProxyFirst=true;
            addProxy(config);
        }
        addLogger(config);
        log("Creating axios instance with config %s", config);
        let instance = axios.create(config);
        await addInterceptor(instance);

        log("Instance details %o", instance);
        log("Downloading %s", projectTemplateZipFile);
        return await instance(config).then(response => {
            log("Finished Downloading %s", projectTemplateZipFile);
            console.log("Downloaded Successfully!");
            return new Promise<string>((resolve, reject) => {
                response.data.pipe(fs.createWriteStream(zipFile))
                    .on('error', function (err) {
                        reject(`Unable to download project zip file for "${config.url}".\n${err}`);
                    })
                    .on('close', async () => {
                        resolve(path.resolve(`${projectFolder}/project.zip`));
                    });
            });
        }).catch(err => {
            log("Failed Downloading %s. Error: %o", projectTemplateZipFile, err);
            const error: string = `Unable to download project zip file for "${config.url}".\n${err}`;
            console.log(error);
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
