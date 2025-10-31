import axios, {AxiosError, AxiosRequestConfig} from "axios"
import * as fs from "fs";
import * as path from "path";
import AdmZip from "adm-zip";
import {HttpsProxyAgent} from "https-proxy-agent";
import {Agent} from "node:https"
import debug from "debug"

const log = debug("genOffice").extend("helper")

const zipFile = 'project.zip';
const httpsProxy = process.env.HTTPS_PROXY || process.env.https_proxy || `http:\\${process.env['https.proxyHost']}:${process.env['https.proxyPort']}`;
log("Proxy %s",httpsProxy)
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

    function hasProxy() {
        return httpsProxy.search("undefined") === -1;
    }

    function removeProxy(config: AxiosRequestConfig) {
        log("Removing Proxy")
        config.proxy = false
        config.httpsAgent = new Agent();
    }

    function addProxy(config: AxiosRequestConfig) {
        log("Adding Proxy")
        config.proxy = false;
        config.httpsAgent = new HttpsProxyAgent(httpsProxy,{keepAlive:false});
    }

    export async function downloadProjectTemplateZipFile(projectFolder: string, projectRepo: string, projectBranch: string): Promise<string> {
        const useProxyFirst = process.env.GENERATOR_OFFICE_USE_PROXY === "true";
        const projectTemplateZipFile = `${projectRepo}/archive/${projectBranch}.zip`;
        log("Setting up config for %s",projectBranch)
        const config : AxiosRequestConfig ={
            method: 'get',
            url: projectTemplateZipFile,
            responseType: 'stream'
        };
        if(useProxyFirst && hasProxy()) {
            addProxy(config);
        }
        let instance = axios.create(config);
        instance.interceptors.response.use(undefined, async (err) => {
            if(hasProxy() && err instanceof AxiosError && (err.code === "ECONNRESET" || err.code === "ECONNREFUSED" || err.code === "ENOENT")) {
                console.log(`Download failed for file ${projectTemplateZipFile}. Attempting ${useProxyFirst ? 'without' : 'with'} proxy. Previous Error: ${err}`)
                if(useProxyFirst) {
                    removeProxy(err.config!);
                } else {
                    addProxy(err.config!);
                }
                return await instance(err.config!);
            }
            throw err;
        })
        log("Downloading %s",projectTemplateZipFile)
        return await instance(config).then(response => {
            log("Finished Downloading %s",projectTemplateZipFile)
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
            log("Failed Downloading %s. Error: %o",projectTemplateZipFile,err);
            const error: string = `Unable to download project zip file for "${config.url}".\n${err}`;
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