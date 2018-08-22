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

const certificateFiles = [
  'certs/ca.crt',
  'certs/server.crt',
  'certs/server.key'
]

const expectedFunctionFilesJs = [
  'function-file/function-file.html',
  'function-file/function-file.js',
];

const expectedFunctionFilesTs = [
  'function-file/function-file.html',
  'function-file/function-file.ts',
];

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

const expectExcelCustomFunctionFiles = [
  ...certificateFiles,
  '.gitignore',
  'package.json',
  'webpack.config.js',
  'config/web.config',
  'index.html',
  'src/customfunctions.js',
  'config/customfunctions.json',
];

/**
 * Test addin from user answers
 * new project, default folder, default host.
 */
describe('Create new project from prompts only', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    projectType: null,
    scriptType: null,    
    name: projectDisplayName,
    host: 'Excel'    
  };
  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user chooses jquery and typescript. */
  describe('jquery & typescript', () => {
    before((done) => {      
      answers.projectType = 'Jquery';
      answers.scriptType = 'Typescript';
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
      answers.scriptType = 'Javascript';
      answers.projectType = 'Jquery';
      helpers.run(path.join(__dirname, '../app'))
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
      answers.scriptType = 'Typescript';
      answers.projectType = 'Angular';
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
        '.babelrc',
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
      answers.scriptType = 'Javascript';
      answers.projectType = 'Angular';
      helpers.run(path.join(__dirname, '../app'))
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
      answers.scriptType = 'Typescript';
      answers.projectType = 'React';
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
    scriptType: null,
    projectType: null,
    name: null,
    host: null  
  };
  let argument = [];

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in"
	 */
  describe('argument: project', () => {
    before((done) => {
      answers.name = projectEscapedName;
      answers.scriptType = 'Typescript';
      answers.host = 'Excel';      
      argument[0] = 'Jquery';

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[2] ? argument[2] : answers.host;
      let name = argument[1] ? argument[1] : answers.name;
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
  describe('arguments: project, name', () => {
    before((done) => {
      answers.scriptType = 'Typescript';
      answers.name = null;
      answers.host = 'Excel'
      argument[0] = 'Jquery';
      argument[1] = projectEscapedName;

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[2] ? argument[2] : answers.host;
      let name = argument[1] ? argument[1] : answers.name;
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
  describe('arguments: project name host', () => {
    before((done) => {
      argument[0] = 'Jquery';
      argument[1] = projectEscapedName;
      argument[2] = 'Excel';      

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let host = argument[2] ? argument[2] : answers.host;
      let name = argument[1] ? argument[1] : answers.name;
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

  /** Test addin when user passes in projectType: excel-functions. */
  describe('arguments: project: custom-functions', () => {
    before((done) => {
      answers.scriptType = null;
      answers.name = projectEscapedName;
      argument[0] = 'excel-functions';
      argument.splice(1, 2);

      helpers.run(path.join(__dirname, '../app'))
        .withArguments(argument)
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let name = argument[1] ? argument[1] : answers.name;
      let manifestFileName = 'manifest.xml';     

      let expected = [
        manifestFileName,
        ...expectExcelCustomFunctionFiles       
      ];

      // assert.file(expected);
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
    scriptType: null,
    projectType: 'Jquery',    
    name: projectDisplayName,
    host: 'Excel'   
  };

  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user pass in --js. */
  describe(' --js', () => {
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
      answers.scriptType = 'Typescript';
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
      let answers = {
        scriptType: 'Typescript',
        projectType: 'Manifest',
        name: projectDisplayName,
        host: 'Excel'   
      };
      before((done) => {
        helpers.run(path.join(__dirname, '../app'))
           .withOptions({ 'output': folderName })
          .withPrompts(answers)
          .on('end', done);
      });
  
      it('creates expected files', (done) => {
        let expected = [
           manifestFileName,
          ...expectedAssets,
          'package.json',
          'resource.html'
        ];  

        // Ensure manifest is found in expected output folder
        assert.ok(path.win32.resolve(manifestFileName).toString().indexOf(folderName) >=0, 'manifest file not found in specified output folder');

        // Verify expected files were created
        assert.file(expected);
        done();
      });
    });    
});