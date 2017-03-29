/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let helpers = require('yeoman-test');
let assert = require('yeoman-assert');
import * as path from 'path';

/**
 * Test addin from user answers
 * new project, default folder, defaul host.
 */
describe('new project - answers', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: projectDisplayName,
    host: 'excel',
    isManifestOnly: false,
    ts: null,
    framework: null,
    open: false
  };
  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user chooses jquery and typescript. */
  describe('jquery & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'jquery';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses jquery and javascript. */
  describe('jquery & javascript', () => {
    before((done) => {
      answers.ts = false;
      answers.framework = 'jquery';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'bsconfig.json',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses angular and typescript. */
  describe('angular & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'angular';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses angular and javascript. */
  describe('angular & javascript', () => {
    before((done) => {
      answers.ts = false;
      answers.framework = 'angular';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'bsconfig.json',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses react and typescript. */
  describe('react & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'react';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'config/webpack.common.js',
        'config/webpack.dev.js',
        'config/webpack.prod.js',
        'src/assets/styles/_flex.scss',
        'src/assets/styles/global.scss',
        'src/components/app.tsx',
        'src/components/header.tsx',
        'src/components/hero-list.tsx',
        'src/components/progress.tsx',
        'src/index.html',
        'src/main.tsx',
        'src/polyfills.ts',
        'src/vendor.ts',
        'tsconfig.webpack.json',
        'tslint.json',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
});

/**
 * Test addin from user answers and arguments
 * new project, default folder, typescript, jquery.
 */
describe('new project - answers & args - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: null,
    host: null,
    isManifestOnly: false,
    ts: true,
    framework: null,
    open: false
  };
  let argument = [];

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in"
	 */
  describe('argument: name', () => {
    before((done) => {
      answers.host = 'excel';
      answers.framework = 'jquery';
      argument[0] = projectEscapedName;

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel"
	 */
  describe('arguments: name host', () => {
    before((done) => {
      answers.framework = 'jquery';
      argument[0] = projectEscapedName;
      argument[1] = 'excel';

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel jquery"
	 */
  describe('arguments: name host framework', () => {
    before((done) => {
      argument[0] = projectEscapedName;
      argument[1] = 'excel';
      argument[2] = 'jquery';

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[1] ? argument[1] : answers.host;
      let name = argument[0] ? argument[0] : answers.name;
      let manifestFileName = name + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
});

/**
 * Test addin from user answers and options
 * new project, default folder, typescript, jquery.
 */
describe('new project - answers & opts - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: projectDisplayName,
    host: 'excel',
    isManifestOnly: false,
    ts: null,
    framework: 'jquery',
    open: false
  };

  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user pass in --js. */
  describe('options: --js', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withOptions({ js: true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user pass in --skip-install. */
  describe('options: --skip-install', () => {
    before((done) => {
      answers.ts = true;
      helpers.run(path.join(__dirname, '../app'))
        .withOptions({ 'skip-install': true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
})