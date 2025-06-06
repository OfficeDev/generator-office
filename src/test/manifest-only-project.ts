/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import helpers from 'yeoman-test';
import assert from 'yeoman-assert';
import * as path from 'path';
import { __dirname } from './utils.js';

const manifestProject = 'manifest';
const projectDisplayName = 'My Office Add-in';
const projectEscapedName = 'my-office-add-in';

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
  'config/webpack.common.js',
  'config/webpack.dev.js',
  'config/webpack.prod.js'
]

/**
 * Test addin when user passes in answers
 * manifest project, default folder, default host.
 */
describe('manifest project - answers', () => {
  const answers = {
    projectType: manifestProject,
    name: projectDisplayName,
    host: 'Excel',
  };

  before((done) => {
    helpers.run(path.join(__dirname, '../app')).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
  });

  it('creates expected files', (done) => {
    assert.file(expectedFiles);
    assert.noFile(unexpectedFiles);
    done();
  });
});

/**
 * Test addin when user passes in answers and arguments
 * manifest project, default folder, default host.
 */
describe('manifest project - answers & args', () => {
  const answers = {
    name: projectDisplayName,
    host: 'Excel',
  };
  const argument: string[] = [];

  before((done) => {
    argument[0] = manifestProject;
    helpers.run(path.join(__dirname, '../app')).withArguments(argument).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
  });

  it('creates expected files', (done) => {
    assert.file(expectedFiles);
    assert.noFile(unexpectedFiles);
    done();
  });
});

/**
 * Test addin when user passes in answers and arguments
 * manifest project, default folder, default host.
 */
describe('manifest project - answers & args', () => {
  const answers = {
    host: 'Excel',
  };
  const argument: string[] = [];

  before((done) => {
    argument[0] = manifestProject;
    argument[1] = projectEscapedName;
    helpers.run(path.join(__dirname, '../app')).withArguments(argument).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
  });

  it('creates expected files', (done) => {
    assert.file(expectedFiles);
    assert.noFile(unexpectedFiles);
    done();
  });
});

/**
 * Test addin when user passes in arguments
 * manifest project, default folder, default host.
 */
describe('manifest project - answers & args', () => {
  const answers = {};
  const argument: string[] = [];

  before((done) => {
    argument[0] = manifestProject;
    argument[1] = projectEscapedName;
    argument[2] = 'Excel';
    helpers.run(path.join(__dirname, '../app')).withArguments(argument).withOptions({ 'test': true } as any).withPrompts(answers).on('end', done);
  });

  it('creates expected files', (done) => {
    assert.file(expectedFiles);
    assert.noFile(unexpectedFiles);
    done();
  });
});
