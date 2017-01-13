let helpers = require('yeoman-test');
let assert = require('yeoman-assert');
import * as path from 'path';

/**
 * Test addin from user answers - existing project, default folder, defaul host.
 */
describe('existing project - office:app', () => {
	let projectDisplayName = 'My Office Add-in';
	let projectEscapedName = 'my-office-add-in';
	let answers = {
			name: projectDisplayName,
			new: 'manifest-only',
			folder: null,
			host: 'excel',
			ts: null,
			framework: null
		};
	let manifestFileName = 'manifest-' + answers.host + '-' + projectEscapedName + '.xml';

	/** Test addin when user chooses jquery and typescript. */
	describe('manifest-only', () => {
		before((done) => {
			helpers.run(path.join(__dirname, '../app')).withPrompts(answers).on('end', done);
		});

		it('creates expected files', (done) => {
			let expected = [
				manifestFileName,
				'package.json',
				'assets/icon-16.png',
				'assets/icon-32.png',
				'assets/icon-80.png',
				'assets/logo-filled.png'
			];

			assert.file(expected);
			done();
		});
	});
});

