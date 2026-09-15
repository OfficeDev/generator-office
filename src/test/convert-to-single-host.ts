/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import assert from 'yeoman-assert';
import * as fs from "fs";
import helpers from 'yeoman-test';
import { OfficeAddinManifest, ManifestInfo } from "office-addin-manifest";
import * as path from 'path';
import { promisify } from "util";
import { __dirname } from './utils.js';
import debug from 'debug';
import { getProxyURL } from '../app/helpers/requestHelpers.js';
import type { BaseOptions } from 'yeoman-generator';

const log = debug("genOffice").extend("test");

type TestGeneratorOptions = Partial<BaseOptions> & {
    host?: string;
    manifestType?: string;
    name?: string;
    projectType?: string;
    test: boolean;
    ts?: boolean;
};

const testOptions: TestGeneratorOptions = { test: true };

const hosts = ["Excel", "Onenote", "Outlook", "Powerpoint", "Project", "Word"];
const manifestXmlFile = "manifest.xml";
const manifestJsonFile = "manifest.json";
const packageJsonFile = "package.json";
const readFileAsync = promisify(fs.readFile);
const unexpectedManifestFiles = [
    'manifest.excel.xml', 
    'manifest.onenote.xml', 
    'manifest.outlook.xml', 
    'manifest.powerpoint.xml', 
    'manifest.project.xml', 
    'manifest.word.xml', 
]

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Excel host
describe('Projects configured with proxy', () => {

    const testProjectName = "TaskpaneProject";
    const useProxyFirst = process.env.GENERATOR_OFFICE_USE_PROXY === "true";
    const proxyBackup = getProxyURL();
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/excel.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/onenote.ts', 
        'src/taskpane/outlook.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane", 
        scriptType: "TypeScript", 
        name: testProjectName, 
        host: hosts[0],
        manifestType: "xml"
    };

    before((done) => {
        if(proxyBackup === undefined) {
            process.env.HTTPS_PROXY='http://localhost:83/';
        }
        process.env.GENERATOR_OFFICE_USE_PROXY = `${!useProxyFirst}`;
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });
    after((done) => {
        if(proxyBackup === undefined) {
            delete process.env['HTTPS_PROXY'];
        }
        process.env.GENERATOR_OFFICE_USE_PROXY = `${useProxyFirst}`;
        done();
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[0].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Excel host
describe('Office-Addin-Taskpane-Ts projects', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/excel.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/onenote.ts', 
        'src/taskpane/outlook.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane", 
        scriptType: "TypeScript", 
        name: testProjectName, 
        host: hosts[0],
        manifestType: "xml"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[0].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Excel host and prerelease flag
describe('Office-Addin-Taskpane-Ts prerelease projects', () => {
    const testProjectName = "Taskpane Project"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/excel.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/onenote.ts', 
        'src/taskpane/outlook.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane", 
        scriptType: "TypeScript", 
        name: testProjectName, 
        host: hosts[0],
        manifestType: "xml"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[0].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });
    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName); // TODO: update when new convert script is in yo-office template branches
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Outlook host and a json manifest
describe('Office-Addin-Taskpane-Ts Outlook json project', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestJsonFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/outlook.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts', 
        'src/taskpane/onenote.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane", 
        scriptType: "TypeScript", 
        name: testProjectName, 
        host: hosts[2], 
        manifestType: "json"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[2].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.json is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestJsonFile);
        assert.equal(manifestInfo.hosts?.[0], "mail");
        assert.equal(manifestInfo.displayName, testProjectName); // TODO: update when new convert script is in yo-office template branches
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Excel host and a json manifest
describe('Office-Addin-Taskpane-Ts Excel json project', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestJsonFile,
        'src/taskpane/taskpane.ts',
        'src/taskpane/excel.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/onenote.ts',
        'src/taskpane/outlook.ts',
        'src/taskpane/powerpoint.ts',
        'src/taskpane/project.ts',
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane",
        scriptType: "TypeScript",
        name: testProjectName,
        host: hosts[0],
        manifestType: "json"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[0].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.json is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestJsonFile);
        assert.equal(manifestInfo.hosts?.[0], "workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Word host and a json manifest
describe('Office-Addin-Taskpane-Ts Word json project', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestJsonFile,
        'src/taskpane/taskpane.ts',
        'src/taskpane/word.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
        'src/taskpane/onenote.ts',
        'src/taskpane/outlook.ts',
        'src/taskpane/powerpoint.ts',
        'src/taskpane/project.ts',
    ]
    const answers = {
        projectType: "taskpane",
        scriptType: "TypeScript",
        name: testProjectName,
        host: hosts[5],
        manifestType: "json"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[5].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.json is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestJsonFile);
        assert.equal(manifestInfo.hosts?.[0], "document");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Powerpoint host and a json manifest
describe('Office-Addin-Taskpane-Ts Powerpoint json project', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestJsonFile,
        'src/taskpane/taskpane.ts',
        'src/taskpane/powerpoint.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
        'src/taskpane/onenote.ts',
        'src/taskpane/outlook.ts',
        'src/taskpane/project.ts',
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane",
        scriptType: "TypeScript",
        name: testProjectName,
        host: hosts[3],
        manifestType: "json"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[3].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.json is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestJsonFile);
        assert.equal(manifestInfo.hosts?.[0], "presentation");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for Office-Addin-Taskpane Typescript project using Outlook host and a xml manifest
describe('Office-Addin-Taskpane-Ts Outlook xml project', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/outlook.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts', 
        'src/taskpane/onenote.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const answers = {
        projectType: "taskpane", 
        scriptType: "TypeScript", 
        name: testProjectName, 
        host: hosts[2], 
        manifestType: "xml"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[2].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Mailbox");
        assert.equal(manifestInfo.displayName, testProjectName); // TODO: update when new convert script is in yo-office template branches
    });
});

// Test to verify converting a project to a single host
// for React Typescript project using PowerPoint host
describe('Office-Addin-Taskpane-React-Ts project', () => {
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/components/App.tsx', 
    ]
    const unexpectedFiles = [
        'src/taskpane/components/Excel.App.tsx', 
        'src/taskpane/components/Onenote.App.tsx', 
        'src/taskpane/components/Outlook.App.tsx', 
        'src/taskpane/components/PowerPoint.App.tsx', 
        'src/taskpane/components/Project.App.tsx', 
        'src/taskpane/components/Word.App.tsx', 
    ]
    const answers = {
        projectType: "react", 
        scriptType: "TypeScript", 
        name: "ReactProject", 
        host: hosts[3],
        manifestType: "xml"
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[3].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Presentation");
    });
});

// Test to verify converting a project to a single host using the cli
// for Office-Addin-Taskpane Typescript project using Excel host
describe('Office-Addin-Taskpane-Ts projects via cli', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/taskpane/taskpane.ts', 
        'src/taskpane/excel.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/onenote.ts', 
        'src/taskpane/outlook.ts', 
        'src/taskpane/powerpoint.ts', 
        'src/taskpane/project.ts', 
        'src/taskpane/word.ts'
    ]
    const options: TestGeneratorOptions = {
        projectType: "taskpane", 
        name: testProjectName, 
        host: hosts[0], 
        ts: true,
        manifestType: "xml",
        test: true
    };
    const answers = {};

    before((done) => {
        log("Running helper for %s", options.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(options).withAnswers(answers).on('end', () => { log("Finished helper for %s", options.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });

    it('Package.json is updated properly', async () => {
        const data: string = await readFileAsync(packageJsonFile, 'utf8');
        const content = JSON.parse(data);
        assert.equal(content.config["app_to_debug"], hosts[0].toLowerCase());

        // Verify host-specific sideload and unload scripts have been removed
        let unexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexpectedScriptsFound = true;
            }
        });
        assert.equal(unexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for custom function Typescript project
describe('Custom-Functions-Shared-TS project', () => {
    const testProjectName = "CFTypeScriptProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/functions/functions.ts', 
    ]
    const unexpectedFiles = [
        'manifest.json', 
    ]
    const answers = {
        projectType: "excel-functions-shared", 
        scriptType: "TypeScript", 
        name: testProjectName
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });
    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for custom functions JavaScript project
describe('Custom-Functions-Shared-JS project', () => {
    const testProjectName = "CFJavaScriptProject"
    const expectedFiles = [
        packageJsonFile, 
        manifestXmlFile, 
        'src/functions/functions.js', 
    ]
    const unexpectedFiles = [
        'manifest.json', 
    ]
    const answers = {
        projectType: "excel-functions-shared", 
        scriptType: "JavaScript", 
        name: testProjectName
    };

    before((done) => {
        log("Running helper for %s", answers.name);
        helpers.run(path.join(__dirname, '../app')).withOptions(testOptions).withAnswers(answers).on('end', () => { log("Finished helper for %s", answers.name);done();});
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });
    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});
