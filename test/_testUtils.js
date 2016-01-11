var fs = require('fs');
var assert = require('yeoman-assert');

/**
 * Helper function to check contents of object.
 */
exports.assertObjectContains = _assertObjectContains;

function _assertObjectContains(obj, content) {
  Object.keys(content).forEach(function (key) {
    if (typeof content[key] === 'object') {
      _assertObjectContains(content[key], obj[key]);
      return;
    }
    assert.equal(content[key], obj[key]);
  });
};

/**
 * Helper function to check contents of JSON file.
 */
exports.assertJSONFileContains = function (filename, content) {
  var obj = JSON.parse(fs.readFileSync(filename, 'utf8'));
  _assertObjectContains(obj, content);
};

/**
 * Setup an existing project in the test folder.
 * @param {RunContext} generator - The generator being run.
 */
exports.setupExistingProject = function (generator) {
  // create existing package.json file
  var existingPackage = {
    name: 'ProjectName',
    description: 'HTTPS site using Express and Node.js',
    version: '0.1.0',
    main: 'src/server/server.js',
    dependencies: {
      express: '^4.12.2'
    }
  };
  // write the package.json file
  generator.fs.writeJSON(generator.destinationPath('package.json'), existingPackage);

  // create existing bower.json file
  var existingBower = {
    name: 'ProjectName',
    version: '0.1.0',
    dependencies: {
      'jquery': '~1.9.1'
    }
  };
  // write the bower.json file
  generator.fs.writeJSON(generator.destinationPath('bower.json'), existingBower);
  
  // create existing tsd.json file
  // lodash is added to just test to ensure the existing tsd.json file isn't overwritten
  var existingTsd = {
    version: 'v4',
    repo: 'borisyankov/DefinitelyTyped',
    ref: 'master',
    path: 'typings',
    bundle: 'typings/tsd.d.ts',
    installed: {
      'jquery/jquery.d.ts': {
        commit: '04a025ada3492a22df24ca2d8521c911697721b3'
      },
      'lodash/lodash.d.ts': {
        commit: '62eedc3121a5e28c50473d2e4a9cefbcb9c3957f'
      }
    }
  };
  // write the bower.json file
  generator.fs.writeJSON(generator.destinationPath('tsd.json'), existingTsd);

  // write out static content
  generator.fs.write(generator.destinationPath('public/index.html'), 'foo');
  generator.fs.write(generator.destinationPath('public/content/site.css'), 'foo');
  generator.fs.write(generator.destinationPath('server/server.js'), 'foo');
};
