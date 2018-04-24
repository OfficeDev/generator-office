/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

const expectedFiles = [
  'package.json',
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
 * manifest-only project, default folder, defaul host.
 */
describe('manifest-only project - answers', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    name: projectDisplayName,
    host: 'excel',
    projectType: 'manifest-only',
  };
  let manifestFileName = projectEscapedName + '-manifest.xml';

	/** Test addin when user chooses jquery and typescript. */
  describe('manifest-only', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        ...expectedFiles
      ];

      assert.file(expected);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});

/**
 * Test addin from user answers and arguments
 * manifest-only project, default folder, typescript, jquery.
 */
describe('manifest-only project - answers & args - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    name: null,
    host: null,
    projectType: 'manifest-only',
  };
  let argument = [];

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in"
	 */
  describe('argument: name', () => {
    before((done) => {
      answers.host = 'excel';
      argument[0] = projectEscapedName;

      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name  + '-manifest.xml';

      let expected = [
        manifestFileName,
        ...expectedFiles
      ];

      assert.file(expected);
      assert.noFile(unexpectedFiles);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel"
	 */
  describe('arguments: name host', () => {
    before((done) => {
      argument[0] = projectEscapedName;
      argument[1] = 'excel';

      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name  + '-manifest.xml';

      let expected = [
        manifestFileName,
        ...expectedFiles
      ];

      assert.file(expected);
      assert.noFile(unexpectedFiles);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel manifest-only"
	 */
  describe('arguments: name host framework', () => {
    before((done) => {
      argument[0] = projectEscapedName;
      argument[1] = 'excel';
      argument[2] = 'manifest-only';

      helpers.run(path.join(__dirname, '../app')).withArguments(argument).withPrompts(answers).on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name  + '-manifest.xml';

      let expected = [
        manifestFileName,
        ...expectedFiles
      ];

      assert.file(expected);
      assert.noFile(unexpectedFiles);
      done();
    });
  });
});
