let helpers = require('yeoman-test');
let assert = require('yeoman-assert');
import * as path from 'path';

/**
 * Test addin from user answers - new project, default folder, defaul host.
 */
describe('office:app', () => {
	let projectDisplayName = 'My Office Add-in';
	let projectEscapedName = 'my-office-add-in';
	let answers = {
			name: projectDisplayName,
			new: true,
			folder: false,
			host: 'excel',
			ts: null,
			framework: null
		};
	let manifestFileName = 'manifest-' + answers.host + '-' + projectEscapedName + '.xml';

	/** Test addin when user chooses jquery and typescript. */
	describe('jquery + typescript', () => {

		before((done) => {
			answers.ts = true;
			answers.framework = 'jquery';
			helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
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
				'tsconfig.json',
				'src/app.ts',
				'src/index.html',
				'typings.json',
			];

			assert.file(expected);
			done();
		});
	});

	/** Test addin when user chooses jquery and javascript. */
	describe('jquery + javascript', () => {

		before((done) => {
			answers.ts = false;
			answers.framework = 'jquery';
			helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
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
				'bsconfig.json',
				'app.js',
				'index.html'
			];

			assert.file(expected);
			done();
		});
	});
	
	/** Test addin when user chooses angular and typescript. */
	describe('angular + typescript', () => {

		before((done) => {
			answers.ts = true;
			answers.framework = 'angular';
			helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
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
				'tsconfig.json',
				'src/app.ts',
				'src/data.service.ts',
				'src/index.controller.ts',
				'src/index.html',
				'typings.json'
			];

			assert.file(expected);
			done();
		});
	});

	/** Test addin when user chooses angular and javascript. */
	describe('angular + javascript', () => {

		before((done) => {
			answers.ts = false;
			answers.framework = 'angular';
			helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
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
				'bsconfig.json',
				'app.js',
				'controllers/home.controller.js',
				'index.html',
				'services/data.service.js'
			];

			assert.file(expected);
			done();
		});
	});
	
			
});

