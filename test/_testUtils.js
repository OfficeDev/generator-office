var fs = require('fs');
var assert = require('yeoman-generator').assert;

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
  // write out static content
  generator.fs.write(generator.destinationPath('public/index.html'), 'foo');
  generator.fs.write(generator.destinationPath('public/content/site.css'), 'foo');
  generator.fs.write(generator.destinationPath('server/server.js'), 'foo');
};