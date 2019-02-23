/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

const manifestProject = 'Manifest';

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

/**
 * Test addin from user answers
 * manifest project, default folder, defaul host.
 */
describe('manifest project - answers', () => {
  let projectDisplayName = 'My Office Add-in';
  let answers = {
    projectType: manifestProject,
    name: projectDisplayName,
  };

	/** Test addin when user chooses jquery and typescript. */
  describe('manifest', () => {
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

/**
 * Test addin from user answers and arguments
 * manifest project, default folder, typescript, jquery.
 */
describe('manifest project - answers & args - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    name: projectEscapedName,
  };
  let argument = [];

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in"
	 */
  describe('argument: project', () => {
    before((done) => {
      argument[0] = manifestProject;
      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel"
	 */
  describe('arguments: project name', () => {
    before((done) => {
      let answers = {
        name: projectEscapedName,
      };
      argument[0] = manifestProject;
      argument[1] = projectEscapedName;

      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel manifest"
	 */
  describe('arguments: project name host', () => {
    before((done) => {
      argument[0] = manifestProject;
      argument[1] = projectEscapedName;
      argument[2] = 'excel';
      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      assert.file(expectedFiles);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});
