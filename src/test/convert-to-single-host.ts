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
describe('Office-Addin-Taskpane-Ts projects', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestXmlFile,
        'src/taskpane/taskpane.ts',
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
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
        host: hosts[0]
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
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
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
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
        host: hosts[0]
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true, 'prerelease': true } as any).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
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
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
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
        host: hosts[2],
        manifestType: "json"
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true, 'prerelease': true } as any).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
    });

    it('Manifest.json is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestJsonFile);
        assert.equal(manifestInfo.hosts?.[0], "mail");
        assert.equal(manifestInfo.displayName, testProjectName); // TODO: update when new convert script is in yo-office template branches
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
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
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
        host: hosts[2],
        manifestType: "xml"
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true, 'prerelease': true } as any).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
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
        'src/taskpane/components/App.tsx', ,
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
        host: hosts[3]
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
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
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
        'src/taskpane/onenote.ts',
        'src/taskpane/outlook.ts',
        'src/taskpane/powerpoint.ts',
        'src/taskpane/project.ts',
        'src/taskpane/word.ts'
    ]
    const options: any = {
        projectType: "taskpane",
        name: testProjectName,
        host: hosts[0],
        ts: true,
        test: true
    };
    const answers = {};

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions(options).withPrompts(answers).on('end', done);
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

        // Verify host-specific sideload and unload sripts have been removed
        let unexexpectedScriptsFound = false;
        Object.keys(content.scripts).forEach(function (key) {
            if (key.includes("sideload:") || key.includes("unload:")) {
                unexexpectedScriptsFound = true;
            }
        });
        assert.equal(unexexpectedScriptsFound, false);
    });

    it('Manifest.xml is updated appropriately', async () => {
        const manifestInfo : ManifestInfo = await OfficeAddinManifest.readManifestFile(manifestXmlFile);
        assert.equal(manifestInfo.hosts?.[0], "Workbook");
        assert.equal(manifestInfo.displayName, testProjectName);
    });
});

// Test to verify converting a project to a single host
// for SSO Typescript project using Excel host
describe('Office-Addin-Taskpane-SSO-TS project', () => {
    const expectedFiles = [
        packageJsonFile,
        manifestXmlFile,
        '.ENV',
        'src/taskpane/taskpane.ts',
        'src/taskpane/taskpane.html',
        'src/taskpane/taskpane.css',
        'src/helpers/fallbackauthdialog.html',
        'src/helpers/fallbackauthdialog.ts',
        'src/helpers/message-helper.ts',
        'src/helpers/middle-tier-calls.ts',
        'src/helpers/sso-helper.ts',
        'src/middle-tier/app.ts',
        'src/middle-tier/msgraph-helper.ts',
        'src/middle-tier/ssoauth-helper.ts'
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.ts',
        'src/taskpane/word.ts',
        'src/taskpane/powerpoint.ts',
        'manifest.excel.xml',
        'manifest.word.xml',
        'manifest.powerpoint.xml'
    ]
    const answers = {
        projectType: "single-sign-on",
        scriptType: "TypeScript",
        name: "SSOTypeScriptProject",
        host: hosts[0]
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
    });
});

// Test to verify converting a project to a single host
// for SSO JavaScript project using PowerPoint host
describe('Office-Addin-Taskpane-SSO-JS project', () => {
    const expectedFiles = [
        packageJsonFile,
        manifestXmlFile,
        '.ENV',
        'src/taskpane/taskpane.js',
        'src/taskpane/taskpane.html',
        'src/taskpane/taskpane.css',
        'src/helpers/documenthelper.js',
        'src/helpers/fallbackauthdialog.html',
        'src/helpers/fallbackauthdialog.js',
        'src/helpers/message-helper.js',
        'src/helpers/middle-tier-calls.js',
        'src/helpers/sso-helper.js',
        'src/middle-tier/app.js',
        'src/middle-tier/msgraph-helper.js',
        'src/middle-tier/ssoauth-helper.js'
    ]
    const unexpectedFiles = [
        'src/taskpane/excel.js',
        'src/taskpane/word.js',
        'src/taskpane/powerpoint.js',
        'manifest.excel.xml',
        'manifest.word.xml',
        'manifest.powerpoint.xml'
    ]
    const answers = {
        projectType: "single-sign-on",
        scriptType: "JavaScript",
        name: "SSOJavaScriptProject",
        host: hosts[3]
    };

    before((done) => {
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
        assert.file(expectedFiles);
        assert.noFile(unexpectedFiles);
        assert.noFile(unexpectedManifestFiles);
        done();
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
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
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
        helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any as any).withPrompts(answers).on('end', done);
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
