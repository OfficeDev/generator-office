/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

const expectedAssets = [
  'assets/icon-16.png',
  'assets/icon-32.png',
  'assets/icon-80.png',
  'assets/logo-filled.png',
];

const expectedFunctionFilesJs = [
  'function-file/function-file.html',
  'function-file/function-file.js',
];

const expectedFunctionFilesTs = [
  'function-file/function-file.html',
  'function-file/function-file.ts',
];

const certificateFiles = [
  'certs/ca.crt',
  'certs/server.crt',
  'certs/server.key'
]

const configFiles = [
  'config/webpack.common.js',
  'config/webpack.dev.js',
  'config/webpack.prod.js'
]

const commonExpectedFiles = [
  '.gitignore',
  'package.json',
  'webpack.config.js',
  'resource.html',
];

/**
 * Test addin from user answers
 * new project, default folder, defaul host.
 */
describe('Create new project from prompts only', () => {
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses jquery and javascript. */
  describe('jquery & javascript', () => {
    before((done) => {
      answers.framework = 'jquery';
      helpers.run(path.join(__dirname, '../app'))
        .withOptions({ js: true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        ...expectedAssets,
        ...expectedFunctionFilesJs,
        ...commonExpectedFiles,
        '.babelrc',
        'app.css',
        'src/index.js',
        'index.html',
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
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
        .withOptions({ js: true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        ...expectedAssets,
        ...expectedFunctionFilesJs,
        ...commonExpectedFiles,
        '.babelrc',
        'app.css',
        'jsconfig.json',
        'index.html',
        'src/index.js',
        'src/app/app.component.html',
        'src/app/app.component.js',
        'src/app/app.module.js',
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'tsconfig.json',
        'config/webpack.common.js',
        'config/webpack.dev.js',
        'config/webpack.prod.js',
        'src/styles.less',
        'src/components/App.tsx',
        'src/components/Header.tsx',
        'src/components/HeroList.tsx',
        'src/components/Progress.tsx',
        'src/index.html',
        'src/index.tsx',
        'tslint.json',
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
describe('Create new project from prompts and command line overrides', () => {
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
      ];

      assert.file(expected);
      done();
    });
  });

  	/**
	 * Test addin when user pass in arguments
	 * "my-office-add-in; excel; manifest-only"
	 */
  describe('arguments: name host', () => {
    before((done) => {
      argument[0] = projectEscapedName;
      argument[1] = 'excel';
      argument[2] = 'manifest-only';

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
        ...expectedAssets,
        'package.json',
        'resource.html'
      ];

      let notExpected = [
        expectedFunctionFilesJs,
        expectedFunctionFilesTs,
        certificateFiles,
        configFiles
      ]

      assert.file(expected);
      assert.noFile(notExpected);
      done();
    });
  });
});

/**
 * Test addin from user answers and options
 * new project, default folder, typescript, jquery.
 */
describe('Create new project from prompts with command line options', () => {
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
        ...expectedAssets,
        ...expectedFunctionFilesJs,
        ...commonExpectedFiles,
        '.babelrc',
        'app.css',
        'src/index.js',
        'index.html',
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
        ...expectedAssets,
        ...expectedFunctionFilesTs,
        ...commonExpectedFiles,
        'app.css',
        'tsconfig.json',
        'src/index.ts',
        'index.html',
      ];

      assert.file(expected);
      done();
    });
  }); 

    /** Test addin when user passes in --output. */
    let folderName = 'testFolder';
    describe('options: --output', () => {
      before((done) => {
        answers.folder = true;
        helpers.run(path.join(__dirname, '../app'))
          .withOptions({ 'output': folderName })
          .withPrompts(answers)
          .on('end', done);
      });
  
      it('creates expected files', (done) => {
        let expected = [
           manifestFileName,
          ...expectedAssets,
          ...expectedFunctionFilesTs,
          ...commonExpectedFiles,
          'app.css',
          'tsconfig.json',
          'src/index.ts',
          'index.html',
        ];  

        // Ensure manifest is found in expected output folder
        assert.ok(path.win32.resolve(manifestFileName).toString().indexOf(folderName) >=0,
        'manifest file not found in specified output folder');

        // Verify expected files were created
        assert.file(expected);
        done();
      });
    });    
});
