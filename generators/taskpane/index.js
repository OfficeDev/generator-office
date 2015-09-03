'use strict';

var generators = require('yeoman-generator');
var chalk = require('chalk');
var path = require('path');
var extend = require('deep-extend');
var guid = require('uuid');

module.exports = generators.Base.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    generators.Base.apply(this, arguments);

    this.option('skip-install', {
      type: Boolean,
      required: false,
      defaults: false,
      desc: 'Skip running package managers (NPM, bower, etc) post scaffolding'
    });

    this.option('name', {
      type: String,
      desc: 'Title of the Office Add-in',
      required: false
    });

    this.option('root-path', {
      type: String,
      desc: 'Relative path where the Add-in should be created (blank = current directory)',
      required: false
    });

    this.option('tech', {
      type: String,
      desc: 'Technology to use for the Add-in (html = HTML; ng = Angular)',
      required: false
    });
    
    // create global config object on this generator
    this.genConfig = {};
  }, // constructor()
  
  /**
   * Prompt users for options
   */
  prompting: {

    askFor: function () {
      var done = this.async();

      var prompts = [
        // friendly name of the generator
        {
          name: 'name',
          message: 'Project name (display name):',
          default: 'My Office Add-in',
          when: this.options.name === undefined
        },
        // root path where the addin should be created; should go in current folder where 
        //  generator is being executed, or within a subfolder?
        {
          name: 'root-path',
          message: 'Root folder of project?'
          + ' Default to current directory\n (' + this.destinationRoot() + '), or specify relative path\n'
          + '  from current (src / public): ',
          default: 'current folder',
          when: this.options['root-path'] === undefined,
          filter: /* istanbul ignore next */ function (response) {
            if (response === 'current folder')
              return '';
            else
              return response;
          }
        },
        // technology used to create the addin (html / angular / etc)
        {
          name: 'tech',
          message: 'Technology to use:',
          type: 'list',
          when: this.options.tech === undefined,
          choices: [
            {
              name: 'HTML, CSS & JavaScript',
              value: 'html'
            }, {
              name: 'Angular',
              value: 'ng'
            }, {
              name: 'Manifest.xml only (no application source files)',
              value: 'manifest-only'
            }]
        }];
        
      // trigger prompts
      this.prompt(prompts, function (responses) {
        this.genConfig = extend(this.genConfig, this.options);
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));

    }, // askFor()
    
    /**
     * If user specified tech:manifest-only, prompt for start page.
     */
    askForStartPage: function () {
      if (this.genConfig.tech !== 'manifest-only')
        return;

      var done = this.async();

      var prompts = [
        // if tech = manifest only, prompt for start page
        {
          name: 'startPage',
          message: 'Add-in start URL:',
          when: this.options.startPage === undefined,
        }];
        
      // trigger prompts
      this.prompt(prompts, function (responses) {
        this.genConfig = extend(this.genConfig, responses);
        done();
      }.bind(this));

    } // askForStartPage()
    
  }, // prompting()
  
  /**
   * save configurations & config project
   */
  configuring: function () {
    // take name submitted and strip everything out non-alphanumeric or space
    var projectName = this.genConfig.name;
    projectName = projectName.replace(/[^\w\s\-]/g, '');
    projectName = projectName.replace(/\s{2,}/g, ' ');
    projectName = projectName.trim();
    
    // add the result of the question to the generator configuration object
    this.genConfig.projectInternalName = projectName.toLowerCase().replace(/ /g, "-");
    this.genConfig.projectDisplayName = projectName;
    this.genConfig.rootPath = this.genConfig['root-path'];
  }, // configuring()
  
  /**
   * write generator specific files
   */
  writing: {
    /**
     * If there is already a package.json in the root of this project, 
     * get the name of the project from that file as that should be used
     * in bower.json & update packages.
     */
    upsertPackage: function () {
      if (this.genConfig.tech !== 'manifest-only') {
        var done = this.async();
      
        // default name for the root project = project
        this.genConfig.rootProjectName = this.genConfig.projectInternalName;

        // path to package.json
        var pathToPackageJson = this.destinationPath('package.json');
      
        // if package.json doesn't exist
        if (!this.fs.exists(pathToPackageJson)) {
          // copy package.json to target
          this.fs.copyTpl(this.templatePath('common/_package.json'),
            this.destinationPath('package.json'),
            this.genConfig);
        } else {
          // load package.json
          var packageJson = this.fs.readJSON(pathToPackageJson, 'utf8');
        
          // .. get it's name property
          this.genConfig.rootProjectName = packageJson.name;
        
          // update devDependencies
          /* istanbul ignore else */
          if (!packageJson.devDependencies) {
            packageJson.devDependencies = {}
          }
          /* istanbul ignore else */
          if (!packageJson.devDependencies['gulp']) {
            packageJson.devDependencies['gulp'] = "^3.9.0"
          }
          /* istanbul ignore else */
          if (!packageJson.devDependencies['gulp-webserver']) {
            packageJson.devDependencies['gulp-webserver'] = "^0.9.1"
          }

          // overwrite existing package.json
          this.log(chalk.yellow('Adding additional packages to package.json'));
          this.fs.writeJSON(pathToPackageJson, packageJson);
        }

        done();
      }
    }, // upsertPackage()
        
    /**
     * If bower.json already exists in the root of this project, update it
     * with the necessary packages.  
     */
    upsertBower: function () {
      if (this.genConfig.tech !== 'manifest-only') {
        var done = this.async();

        var pathToBowerJson = this.destinationPath('bower.json');
        // if doesn't exist...
        if (!this.fs.exists(pathToBowerJson)) {
          // copy bower.json => project
          switch (this.genConfig.tech) {
            case "ng":
              this.fs.copyTpl(this.templatePath('ng/_bower.json'),
                this.destinationPath('bower.json'),
                this.genConfig);
              break;
            case "html":
              this.fs.copyTpl(this.templatePath('html/_bower.json'),
                this.destinationPath('bower.json'),
                this.genConfig);
              break;
          }
        } else {
          // verify the necessary package references are present in bower.json...
          //  if not, add them
          var bowerJson = this.fs.readJSON(pathToBowerJson, 'utf8');

          // all addins need these
          if (!bowerJson.dependencies["microsoft.office.js"]) {
            bowerJson.dependencies["microsoft.office.js"] = "*";
          }
          if (!bowerJson.dependencies["jquery"]) {
            bowerJson.dependencies["jquery"] = "~1.9.1";
          }

          switch (this.genConfig.tech) {
            // if angular...
            case "ng":
              if (!bowerJson.dependencies["angular"]) {
                bowerJson.dependencies["angular"] = "~1.4.4";
              }
              if (!bowerJson.dependencies["angular-route"]) {
                bowerJson.dependencies["angular-route"] = "~1.4.4";
              }
              if (!bowerJson.dependencies["angular-sanitize"]) {
                bowerJson.dependencies["angular-sanitize"] = "~1.4.4";
              }
              break;
          }
        
          // overwrite existing bower.json
          this.log(chalk.yellow('Adding additional packages to bower.json'));
          this.fs.writeJSON(pathToBowerJson, bowerJson);
        }

        done();
      }
    }, // upsertBower()

    app: function () {
      // helper function to build path to the file off root path
      this._parseTargetPath = function (file) {
        return path.join(this.genConfig['root-path'], file);
      };

      var done = this.async();

      // create a new ID for the project
      this.genConfig.projectId = guid.v4();

      if (this.genConfig.tech === 'manifest-only') {
        // create the manifest file
        this.fs.copyTpl(this.templatePath('common/manifest.xml'), this.destinationPath('manifest.xml'), this.genConfig);
      } else {
        // copy .bowerrc => project
        this.fs.copyTpl(
          this.templatePath('common/_bowerrc'),
          this.destinationPath('.bowerrc'),
          this.genConfig);

        // create common assets
        this.fs.copy(this.templatePath('common/gulpfile.js'), this.destinationPath('gulpfile.js'));
        this.fs.copy(this.templatePath('common/content/Office.css'), this.destinationPath(this._parseTargetPath('content/Office.css')));
        this.fs.copy(this.templatePath('common/content/fabric.css'), this.destinationPath(this._parseTargetPath('content/fabric.css')));
        this.fs.copy(this.templatePath('common/content/fabric.min.css'), this.destinationPath(this._parseTargetPath('content/fabric.min.css')));
        this.fs.copy(this.templatePath('common/content/fabric.rtl.css'), this.destinationPath(this._parseTargetPath('content/fabric.rtl.css')));
        this.fs.copy(this.templatePath('common/content/fabric.rtl.min.css'), this.destinationPath(this._parseTargetPath('content/fabric.rtl.min.css')));
        this.fs.copy(this.templatePath('common/content/fabric.components.css'), this.destinationPath(this._parseTargetPath('content/fabric.components.css')));
        this.fs.copy(this.templatePath('common/content/fabric.components.min.css'), this.destinationPath(this._parseTargetPath('content/fabric.components.min.css')));
        this.fs.copy(this.templatePath('common/content/fabric.components.rtl.css'), this.destinationPath(this._parseTargetPath('content/fabric.components.rtl.css')));
        this.fs.copy(this.templatePath('common/content/fabric.components.rtl.min.css'), this.destinationPath(this._parseTargetPath('content/fabric.components.rtl.min.css')));     
        this.fs.copy(this.templatePath('common/images/close.png'), this.destinationPath(this._parseTargetPath('images/close.png')));
        this.fs.copy(this.templatePath('common/scripts/MicrosoftAjax.js'), this.destinationPath(this._parseTargetPath('scripts/MicrosoftAjax.js')));
        this.fs.copy(this.templatePath('common/scripts/jquery.fabric.js'), this.destinationPath(this._parseTargetPath('scripts/jquery.fabric.js')));
        this.fs.copy(this.templatePath('common/scripts/jquery.fabric.min.js'), this.destinationPath(this._parseTargetPath('scripts/jquery.fabric.min.js')));

        switch (this.genConfig.tech) {
          case 'html':
            // determine startpage for addin
            this.genConfig.startPage = 'https://localhost:8443/app/home/home.html';

            // copy tsd & jsconfig files
            this.fs.copy(this.templatePath('html/_tsd.json'), this.destinationPath('tsd.json'));
            this.fs.copy(this.templatePath('common/_jsconfig.json'), this.destinationPath('jsconfig.json'));

            // create the manifest file
            this.fs.copyTpl(this.templatePath('common/manifest.xml'), this.destinationPath('manifest.xml'), this.genConfig);

            // copy addin files
            this.fs.copy(this.templatePath('html/app.css'), this.destinationPath(this._parseTargetPath('app/app.css')));
            this.fs.copy(this.templatePath('html/app.js'), this.destinationPath(this._parseTargetPath('app/app.js')));
            this.fs.copy(this.templatePath('html/home/home.html'), this.destinationPath(this._parseTargetPath('app/home/home.html')));
            this.fs.copy(this.templatePath('html/home/home.css'), this.destinationPath(this._parseTargetPath('app/home/home.css')));
            this.fs.copy(this.templatePath('html/home/home.js'), this.destinationPath(this._parseTargetPath('app/home/home.js')));
            break;
          case 'ng':
            // determine startpage for addin
            this.genConfig.startPage = 'https://localhost:8443/index.html';

            // copy tsd & jsconfig files
            this.fs.copy(this.templatePath('html/_tsd.json'), this.destinationPath('tsd.json'));
            this.fs.copy(this.templatePath('common/_jsconfig.json'), this.destinationPath('jsconfig.json'));

            // create the manifest file
            this.fs.copyTpl(this.templatePath('common/manifest.xml'), this.destinationPath('manifest.xml'), this.genConfig);

            // copy addin files
            this.genConfig.startPage = '{https-addin-host-site}/index.html';
            this.fs.copy(this.templatePath('ng/index.html'), this.destinationPath(this._parseTargetPath('index.html')));
            this.fs.copy(this.templatePath('ng/app.module.js'), this.destinationPath(this._parseTargetPath('app/app.module.js')));
            this.fs.copy(this.templatePath('ng/app.routes.js'), this.destinationPath(this._parseTargetPath('app/app.routes.js')));
            this.fs.copy(this.templatePath('ng/home/home.controller.js'), this.destinationPath(this._parseTargetPath('app/home/home.controller.js')));
            this.fs.copy(this.templatePath('ng/home/home.html'), this.destinationPath(this._parseTargetPath('app/home/home.html')));
            this.fs.copy(this.templatePath('ng/services/data.service.js'), this.destinationPath(this._parseTargetPath('app/services/data.service.js')));
            break;
        }
      }
      done();
    } // app()
  }, // writing()
    
  /**
   * conflict resolution
   */
  // conflicts: { }, 

  /**
   * run installations (bower, npm, tsd, etc)
   */
  install: function () {

    if (!this.options['skip-install'] && this.genConfig.tech !== 'manifest-only') {
      this.npmInstall();
      this.bowerInstall();
    }

  } // install ()
    
  /**
   * last cleanup, goodbye, etc
   */
  // end: { }

});