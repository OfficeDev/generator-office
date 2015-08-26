
'use strict';

var fs = require('fs');
var path = require('path');
var mockery = require('mockery');
var assert = require('yeoman-generator').assert;
var helpers = require('yeoman-generator').test;

var Xml2Js = require('xml2js');
var validator = require('validator');
var chai = require('chai'),
  expect = chai.expect;

/**
 * Helper function to check contents of object.
 */
function assertObjectContains(obj, content) {
  Object.keys(content).forEach(function (key) {
    if (typeof content[key] === 'object') {
      assertObjectContains(content[key], obj[key]);
      return;
    }

    assert.equal(content[key], obj[key]);
  });
}

/**
 * Helper function to check contents of JSON file.
 */
function assertJSONFileContains(filename, content) {
  var obj = JSON.parse(fs.readFileSync(filename, 'utf8'));
  assertObjectContains(obj, content);
}

/**
 * Setup an existing project in the test folder.
 * @generator {RunContext} - The generator being run.
 */
function setupExistingProject(generator) {
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
}



// sub:generator options 
var options = {};


/* +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+ */

describe('office:taskpane', function () {

  before(function (done) {
    options = {
      name: 'My Office Add-in'
    };
    done();
  });
  
  /**
   * Test addin when running on empty folder.
   */
  describe('run on new project (empty folder)', function () {

    before(function (done) {
      // set to current folder
      options.rootPath = '';
      done();
    });

    /**
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function () {
      before(function (done) {
        //set language to html
        options.tech = 'manifest-only';
        options.startPage = 'https://localhost:8443/manifest-only/index.html';

        // run the generator
        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        assert.file('manifest.xml');
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/manifest-only/index.html');
          done();
        });

      }); //describe('manifest.xml contents')

    }); // describe('technology:manifest-only')

    /**
     * Test addin when technology = html
     */
    describe('technology:html', function () {

      before(function (done) {
        //set language to html
        options.tech = 'html';

        // run the generator
        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        var expected = [
          '.bowerrc',
          'bower.json',
          'package.json',
          'gulpfile.js',
          'manifest.xml',
          'app/app.js',
          'app/app.css',
          'app/home/home.js',
          'app/home/home.html',
          'app/home/home.css',
          'content/Office.css',
          'images/close.png',
          'scripts/MicrosoftAjax.js'
        ];
        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1'
          }
        };

        assert.file('bower.json');
        assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          devDependencies: {
            gulp: '^3.9.0',
            'gulp-webserver': '^0.9.1'
          }
        };

        assert.file('package.json');
        assertJSONFileContains('package.json', expected);
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://{addin-host-site}/app/home/home.html');
          done();
        });

      }); //describe('manifest.xml contents')

      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function () {

        it('contains task \'serve-static\'', function (done) {

          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });

      }); // describe('gulpfile.js contents')

    }); // describe('technology:html')
      
      
    /**
     * Test addin when technology = angular
     */
    describe('addin technology:ng', function () {

      before(function (done) {
        //set language to html
        options.tech = 'ng';

        // run the generator
        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        var expected = [
          '.bowerrc',
          'bower.json',
          'package.json',
          'gulpfile.js',
          'manifest.xml',
          'index.html',
          'app/app.module.js',
          'app/app.routes.js',
          'app/home/home.controller.js',
          'app/home/home.html',
          'app/services/data.service.js',
          'content/Office.css',
          'images/close.png',
          'scripts/MicrosoftAjax.js'
        ];
        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1',
            angular: '~1.4.4',
            'angular-route': '~1.4.4',
            'angular-sanitize': '~1.4.4'
          }
        };

        assert.file('bower.json');
        assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function (done) {
        var expected = {
          name: 'my-office-add-in',
          version: '0.1.0',
          devDependencies: {
            gulp: '^3.9.0',
            'gulp-webserver': '^0.9.1'
          }
        };

        assert.file('package.json');
        assertJSONFileContains('package.json', expected);
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://{addin-host-site}/index.html');
          done();
        });

      }); // describe('manifest.xml contents')
      
      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function () {

        it('contains task \'serve-static\'', function (done) {

          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });

      }); // describe('gulpfile.js contents')

    }); // describe('technology:ng')
      
  }); // describe('run on new project (empty folder)')








  describe('run on existing project (non-empty folder)', function () {
    var addinRootPath = 'src/public';
    
    // generator ran at 'src/public' so for files
    //  in the root, need to back up to the root
    before(function (done) {
      // set to current folder
      options.rootPath = addinRootPath;
      done();
    });
    
    /**
     * Test addin when technology = manifest-only
     */
    describe('technology:manifest-only', function () {
      before(function (done) {
        //set language to html
        options.tech = 'manifest-only';
        options.startPage = 'https://localhost:8443/manifest-only/index.html';

        // run the generator
        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        assert.file('manifest.xml');
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://localhost:8443/manifest-only/index.html');
          done();
        });

      }); //describe('manifest.xml contents')

    }); // describe('technology:manifest-only')
    
    /**
     * Test addin when technology = html
     */
    describe('technology:html', function () {

      before(function (done) {
        //set language to html
        options.tech = 'html';

        // run the generator
        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('ready', function (gen) {
            setupExistingProject(gen);
          }.bind(this))
          .on('end', done);

      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        var expected = [
          '.bowerrc',
          'bower.json',
          'gulpfile.js',
          'package.json',
          'manifest.xml',
          addinRootPath + '/app/app.js',
          addinRootPath + '/app/app.css',
          addinRootPath + '/app/home/home.js',
          addinRootPath + '/app/home/home.html',
          addinRootPath + '/app/home/home.css',
          addinRootPath + '/content/Office.css',
          addinRootPath + '/images/close.png',
          addinRootPath + '/scripts/MicrosoftAjax.js'
        ];
        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'ProjectName',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1'
          }
        };

        assert.file('bower.json');
        assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function (done) {
        var expected = {
          name: 'ProjectName',
          description: 'HTTPS site using Express and Node.js',
          version: '0.1.0',
          main: 'src/server/server.js',
          dependencies: {
            express: '^4.12.2'
          },
          devDependencies: {
            gulp: '^3.9.0',
            'gulp-webserver': '^0.9.1'
          }
        };

        assert.file('package.json');
        assertJSONFileContains('package.json', expected);
        done();
      });

      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://{addin-host-site}/app/home/home.html');
          done();
        });

      }); //describe('manifest.xml contents')

      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function () {

        it('contains task \'serve-static\'', function (done) {

          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });

      }); // describe('gulpfile.js contents')

    }); // describe('technology:html')


    /**
     * Test addin when technology = ng
     */
    describe('technology:ng', function () {

      before(function (done) {
        //set language to html
        options.tech = 'ng';

        helpers.run(path.join(__dirname, '../generators/taskpane'))
          .withOptions(options)
          .on('ready', function (gen) {
            setupExistingProject(gen);
          }.bind(this))
          .on('end', done);
      });

      after(function () {
        mockery.disable();
      });

      /**
       * All expected files are created.
       */
      it('creates expected files', function (done) {
        var expected = [
          '.bowerrc',
          'bower.json',
          'gulpfile.js',
          'package.json',
          'manifest.xml',
          addinRootPath + '/index.html',
          addinRootPath + '/app/app.module.js',
          addinRootPath + '/app/app.routes.js',
          addinRootPath + '/app/home/home.controller.js',
          addinRootPath + '/app/home/home.html',
          addinRootPath + '/app/services/data.service.js',
          addinRootPath + '/content/Office.css',
          addinRootPath + '/images/close.png',
          addinRootPath + '/scripts/MicrosoftAjax.js'
        ];


        assert.file(expected);
        done();
      });

      /**
       * bower.json is good
       */
      it('bower.json contains correct values', function (done) {
        var expected = {
          name: 'ProjectName',
          version: '0.1.0',
          dependencies: {
            'microsoft.office.js': '*',
            jquery: '~1.9.1',
            angular: '~1.4.4',
            'angular-route': '~1.4.4',
            'angular-sanitize': '~1.4.4'
          }
        };

        assert.file('bower.json');
        assertJSONFileContains('bower.json', expected);
        done();
      });

      /**
       * package.json is good
       */
      it('package.json contains correct values', function (done) {
        var expected = {
          name: 'ProjectName',
          description: 'HTTPS site using Express and Node.js',
          version: '0.1.0',
          main: 'src/server/server.js',
          dependencies: {
            express: '^4.12.2'
          },
          devDependencies: {
            gulp: '^3.9.0',
            'gulp-webserver': '^0.9.1'
          }
        };

        assert.file('package.json');
        assertJSONFileContains('package.json', expected);
        done();
      });
      
      /**
       * manfiest.xml is good
       */
      describe('manifest.xml contents', function () {
        var manifest = {};

        before(function (done) {
          var parser = new Xml2Js.Parser();
          fs.readFile('manifest.xml', 'utf8', function (err, manifestContent) {
            parser.parseString(manifestContent, function (err, manifestJson) {
              manifest = manifestJson;

              done();
            });
          });
        });

        it('has valid ID', function (done) {
          expect(validator.isUUID(manifest.OfficeApp.Id)).to.be.true;;
          done();
        });

        it('has correct display name', function (done) {
          expect(manifest.OfficeApp.DisplayName[0].$.DefaultValue).to.equal('My Office Add-in');
          done();
        });

        it('has correct start page', function (done) {
          expect(manifest.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue).to.equal('https://{addin-host-site}/index.html');
          done();
        });

      }); // describe('manifest.xml contents')

      /**
       * gulpfile.js is good
       */
      describe('gulpfule.js contents', function () {

        it('contains task \'serve-static\'', function (done) {

          assert.file('gulpfile.js');
          assert.fileContent('gulpfile.js', 'gulp.task(\'serve-static\',');
          done();
        });

      }); // describe('gulpfile.js contents')

    }); // describe('technology:ng')
    
  }); // describe('run on existing project (non-empty folder)')

});