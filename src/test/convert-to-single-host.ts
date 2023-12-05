/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import * as assert from 'yeoman-assert';
import * as fs from "fs";
import * as helpers from 'yeoman-test';
import { OfficeAddinManifest } from "office-addin-manifest";
import * as path from 'path';
import { promisify } from "util";

const hosts = ["excel", "onenote", "outlook", "powerpoint", "project", "word"];
const manifestFile = "manifest.xml";
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
describe('Office-Add-Taskpane-Ts projects', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestFile,
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

    describe('Office-Add-Taskpane project', () => {
        before((done) => {
            helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true }).withPrompts(answers).on('end', done);
        });

        it('creates expected files', (done) => {
            assert.file(expectedFiles);
            assert.noFile(unexpectedFiles);
            assert.noFile(unexpectedManifestFiles);
            done();
        });
    });

    describe('Package.json is updated appropriately', () => {
        it('Package.json is updated properly', async () => {
            const data: string = await readFileAsync(packageJsonFile, 'utf8');
            const content = JSON.parse(data);
            assert.equal(content.config["app_to_debug"], hosts[0]);

            // Verify host-specific sideload and unload sripts have been removed
            let unexexpectedScriptsFound = false;
            Object.keys(content.scripts).forEach(function (key) {
                if (key.includes("sideload:") || key.includes("unload:")) {
                    unexexpectedScriptsFound = true;
                }
            });
            assert.equal(unexexpectedScriptsFound, false);
        });
    });

    describe('Manifest.xml is updated appropriately', () => {
        it('Manifest.xml is updated appropriately', async () => {
            const manifestInfo = await OfficeAddinManifest.readManifestFile(manifestFile);
            assert.equal(manifestInfo.hosts, "Workbook");
            assert.equal(manifestInfo.displayName, testProjectName);
        });
    });
});

// for Office-Addin-Taskpane Typescript project using Excel host
describe('Office-Add-Taskpane-Ts prerelease projects', () => {
    const testProjectName = "TaskpaneProject"
    const expectedFiles = [
        packageJsonFile,
        manifestFile,
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

     describe('Office-Add-Taskpane prerelease project', () => {
        before((done) => {
            helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true, 'prerelease': true }).withPrompts(answers).on('end', done);
        });

        it('creates expected files', (done) => {
            assert.file(expectedFiles);
            assert.noFile(unexpectedFiles);
            assert.noFile(unexpectedManifestFiles);
            done();
        });
    });

    describe('Package.json is updated appropriately', () => {
        it('Package.json is updated properly', async () => {
            const data: string = await readFileAsync(packageJsonFile, 'utf8');
            const content = JSON.parse(data);
            assert.equal(content.config["app_to_debug"], hosts[0]);

            // Verify host-specific sideload and unload sripts have been removed
            let unexexpectedScriptsFound = false;
            Object.keys(content.scripts).forEach(function (key) {
                if (key.includes("sideload:") || key.includes("unload:")) {
                    unexexpectedScriptsFound = true;
                }
            });
            assert.equal(unexexpectedScriptsFound, false);
        });
    });

    describe('Manifest.xml is updated appropriately', () => {
        it('Manifest.xml is updated appropriately', async () => {
            const manifestInfo = await OfficeAddinManifest.readManifestFile(manifestFile);
            assert.equal(manifestInfo.hosts, "Workbook");
            assert.equal(manifestInfo.displayName, testProjectName);
        });
    });
});

// Test to verify converting a project to a single host
// for React Typescript project using PowerPoint host
describe('Office-Add-Taskpane-React-Ts project', () => {
    const expectedFiles = [
        packageJsonFile,
        manifestFile,
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

    describe('Office-Add-Taskpane project', () => {
        before((done) => {
            helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true }).withPrompts(answers).on('end', done);
        });

        it('creates expected files', (done) => {
            assert.file(expectedFiles);
            assert.noFile(unexpectedFiles);
            assert.noFile(unexpectedManifestFiles);
            done();
        });
    });

    describe('Package.json is updated appropriately', () => {
        it('Package.json is updated properly', async () => {
            const data: string = await readFileAsync(packageJsonFile, 'utf8');
            const content = JSON.parse(data);
            assert.equal(content.config["app_to_debug"], hosts[3]);

            // Verify host-specific sideload and unload sripts have been removed
            let unexexpectedScriptsFound = false;
            Object.keys(content.scripts).forEach(function (key) {
                if (key.includes("sideload:") || key.includes("unload:")) {
                    unexexpectedScriptsFound = true;
                }
            });
            assert.equal(unexexpectedScriptsFound, false);
        });
    });

    describe('Manifest.xml is updated appropriately', () => {
        it('Manifest.xml is updated appropriately', async () => {
            const manifestInfo = await OfficeAddinManifest.readManifestFile(manifestFile);
            assert.equal(manifestInfo.hosts, "Presentation");
        });
    });
});

// // Test to verify converting a project to a single host
// // for SSO Typescript project using Excel host
// describe('Office-Add-Taskpane-SSO-TS project', () => {
//     const expectedFiles = [
//         packageJsonFile,
//         manifestFile,
//         '.ENV',
//         'src/taskpane/taskpane.ts',
//         'src/taskpane/taskpane.html',
//         'src/taskpane/taskpane.css',
//         'src/helpers/fallbackauthdialog.html',
//         'src/helpers/fallbackauthdialog.ts',
//         'src/helpers/fallbackauthhelper.ts',
//         'src/helpers/ssoauthhelper.ts'

//     ]
//     const unexpectedFiles = [
//         'src/taskpane/excel.ts',
//         'src/taskpane/word.ts',
//         'src/taskpane/powerpoint.ts',
//         'manifest.excel.xml',
//         'manifest.word.xml',
//         'manifest.powerpoint.xml'
//     ]
//     const answers = {
//         projectType: "single-sign-on",
//         scriptType: "TypeScript",
//         name: "SSOTypeScriptProject",
//         host: hosts[0]
//     };

//     describe('Office-Add-Taskpane-SSO-TS project', () => {
//         before((done) => {
//             helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true }).withPrompts(answers).on('end', done);
//         });

//         it('creates expected files', (done) => {
//             assert.file(expectedFiles);
//             assert.noFile(unexpectedFiles);
//             assert.noFile(unexpectedManifestFiles);
//             done();
//         });
//     });
// });

// // Test to verify converting a project to a single host
// // for SSO JavaScript project using PowerPoint host
// describe('Office-Add-Taskpane-SSO-JS project', () => {
//     const expectedFiles = [
//         packageJsonFile,
//         manifestFile,
//         '.ENV',
//         'src/taskpane/taskpane.js',
//         'src/taskpane/taskpane.html',
//         'src/taskpane/taskpane.css',
//         'src/helpers/documenthelper.js',
//         'src/helpers/fallbackauthdialog.html',
//         'src/helpers/fallbackauthdialog.js',
//         'src/helpers/fallbackauthhelper.js',
//         'src/helpers/ssoauthhelper.js'

//     ]
//     const unexpectedFiles = [
//         'src/taskpane/excel.js',
//         'src/taskpane/word.js',
//         'src/taskpane/powerpoint.js',
//         'manifest.excel.xml',
//         'manifest.word.xml',
//         'manifest.powerpoint.xml'
//     ]
//     const answers = {
//         projectType: "single-sign-on",
//         scriptType: "JavaScript",
//         name: "SSOJavaScriptProject",
//         host: hosts[3]
//     };

//     describe('Office-Add-Taskpane-SSO-JS project', () => {
//         before((done) => {
//             helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true }).withPrompts(answers).on('end', done);
//         });

//         it('creates expected files', (done) => {
//             assert.file(expectedFiles);
//             assert.noFile(unexpectedFiles);
//             assert.noFile(unexpectedManifestFiles);
//             done();
//         });
//     });
// });
