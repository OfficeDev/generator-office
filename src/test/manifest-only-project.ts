/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

const manifestProject = 'Manifest';
const projectDisplayName = 'My Office Add-in';
let projectEscapedName = 'my-office-add-in';

const expectedFiles = [
  'package.json',
  'manifest.xml',
  'assets/icon-16.png',
  'assets/icon-32.png',
  'assets/icon-80.png',
  'assets/logo-filled.png'
]

const unexpectedFiles = [
  'function-file/function-file.html',
  'function-file/function-file.js',
  'function-file/function-file.html',
  'function-file/function-file.ts',
  'certs/ca.crt',
  'certs/server.crt',
  'certs/server.key',
  'config/webpack.common.js',
  'config/webpack.dev.js',
  'config/webpack.prod.js'
]

// Test generation of manifest project when user provides answers to project type and project name
describe('answers: manifest project and project name', () => {
  describe('manifest project', () => {
    let answers = {
      projectType: manifestProject,
      name: projectDisplayName,
    };
    before((done) => {
      helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});

// Test generation of manifest project when user passes in manifest project and project name as arguments 
describe('arguments: manifest project and project name', () => {
  describe('manifest project', () => {
    let argument = [];
    argument[0] = manifestProject;
    argument[1] = projectEscapedName;
    before((done) => {
      helpers.run(path.join(__dirname, '../app')).withArguments(argument).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});

// Test generation of manifest project when user passes in manifest project as argument and project name as answer
describe('arguments: manifest project; answer: project name', () => {
  let argument = [];
  argument[0] = manifestProject;
  let answers = { name: projectDisplayName };
  describe('manifest project', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});

// Test addin when user passes in project name as argument and provides manifest project as answer
describe('argument: project name; answer: manifest project', () => {
  let argument = [];
  argument[1] = projectEscapedName;
  let answers = { projectType: manifestProject };
  describe('manifest project', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});

