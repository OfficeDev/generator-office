# Microsoft Office Project Generator - YO OFFICE!

[![npm version](https://badge.fury.io/js/generator-office.svg)](http://badge.fury.io/js/generator-office)
[![Downloads](http://img.shields.io/npm/dm/generator-office.svg)](https://npmjs.org/package/generator-office)
[![TravisCI Build Status](https://travis-ci.org/OfficeDev/generator-office.svg)](https://travis-ci.org/OfficeDev/generator-office)
[![AppVeyor Build status](https://ci.appveyor.com/api/projects/status/skat83si6pepm2vp?svg=true)](https://ci.appveyor.com/project/andrewconnell/generator-office)
[![Coverage Status](https://coveralls.io/repos/OfficeDev/generator-office/badge.svg?branch=master&service=github)](https://coveralls.io/github/OfficeDev/generator-office?branch=master)
[![Dependency Status](https://david-dm.org/officedev/generator-office.svg)](https://david-dm.org/officedev/generator-office)
[![devDependency Status](https://david-dm.org/officedev/generator-office/dev-status.svg)](https://david-dm.org/officedev/generator-office#info=devDependencies)
[![Slack Network](http://officedevslack.azurewebsites.com/badge.svg)](http://officedevslack.azurewebsites.com/)

[Yeoman](http://yeoman.io) generator for creating Microsoft Office projects using any text editor. Microsoft includes fantastic & [rich development tools for creating Office related projects using Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013) or [tools for Visual Studio 2015](http://aka.ms/OfficeDevToolsForVS2015). This generator is for those developers who:

- use a editor other than Visual Studio
- interested in using a technology other than plain HTML, CSS & JavaScript

Like other Yeoman generators, this simply creates the scaffolding of files for your Office project. It allows you to create:

- Office Mail Add-ins (both read & compose forms)
- Office Task Pane Add-ins
- Office Content Add-ins

Choose to create the Office projects using plain HTML, CSS & JavaScript (*mirroring the same projects that Visual Studio creates*) or create Angular-based projects.

Check out the announcement blog post: [Office Dev Center Blog - Creating Office Add-ins with any editor - Introducing YO OFFICE!](http://dev.office.com/blogs/creating-office-add-ins-with-any-editor-introducing-yo-office) 

Read up on [how to use the generator to create Office Add-ins with Visual Studio Code](https://code.visualstudio.com/Docs/runtimes/office).

If you are interested in contributing, read the the [Contributing Guidelines](docs/contributing.md). 

## YO Office Demo (screenshot & video)
![](docs/assets/generatoroffice.png)

<iframe width="560" height="315" src="https://www.youtube.com/embed/78b18BLVosM" frameborder="0" allowfullscreen></iframe>

***

## Install

> **Wait!** 

> Is this the first time you're using Yeoman or installing a Yeoman generator? When working with Yeoman there are a few common prerequisites. Ensure you have already have a copy of the popular source control solution [Git](https://git-scm.com/download) installed.

> If you don't have git installed, once you install it we recommend you restart your console (or if on Windows, restart your machine) as system environment variables are set/updated during this installation.

Install `yo` (Yeoman) and `generator-office` globally using NPM (this also requires [Node.js](https://nodejs.org). 

The project files created by the generator leverage client side packages in [bower](http://bower.io) so you will want to install that as well.

In addition, the task runner [gulp](https://www.npmjs.com/package/gulp) is used in the build process to assist developers in validating the Office Add-in manifest file as well as other tasks, so that is listed as part of the install below.

In the v0.5.1 release we added TypeScript type definitions for autocompletion / IntelliSense... for this you need to install the TSD utility before you install the generator.

```bash
$ npm install -g tsd bower gulp yo generator-office
```

## Usage

```bash
$ yo office [options]
```

The generator is intented to be run from within a folder where you want the project scaffolding created. This can be in the root of the current folder or within a subfolder.

> Note: Office Add-ins must be hosted, even in development, in a **HTTPS** site. Refer to the section [Running the Generated Site](/OfficeDev/generator-office#running-the-generated-site) below for details.

## Sub Generators

Running the main generator will prompt you for the type of Office project to create. This triggers the execution of one of the included sub generators. You can instead call one of these sub generators directly to bypass that question:

  - `office:mail` - creates a Mail Add-in
  - `office:taskpane` - creates a Task Pane Add-in
  - `office:content` - creates a Content Add-in

> Remember you can see the options of each sub generators by running `$ yo office:[sub] --help`

## Running the Generated Site

All generators create a `gulpfile.js`. This uses [BrowserSync.io](https://www.browsersync.io) to start a web server running on HTTPS. This server includes a self-signed SSL cert that your development enviroment must trust. 

> Using a self-signed certificate involves adding it to your trusted root certificates... see our doc [Adding Self-Signed Certificates as Trusted Root Certificate](docs/trust-self-signed-cert.md) for instructions on how to do this.

> Because the gulp plugin is added to all Office Add-ins created using this generator, you only need to setup the trust relationship with the self-signed cert it includes one time per developer workstation.

Start the local HTTPS site on `https://localhost:8443/` and launch a browser to this site using:

```bash
$ gulp serve-static
```

You can add the `open` property set to a URL to have your default browser open & navigate to when running this task.

## Examples

Refer to the [docs](docs) for example executions & output of the generator.## Command Line Options:

List of supported options. If these are not provided, the generator will prompt you for the values before scaffolding the project.

> For a full list of all options & descriptions, run `yo office --help`

### `--skip-install`

After scaffolding the project, the generator (and all sub generators) run all package management install commands such as `npm install` & `bower install`. Specifying `--skip-install` tells the generator to skip this step.

  - Type: Boolean
  - Default: False
  - Optional

### `--name:'..'`

Title of the project - this is the display name that is written the `manifest.xml` file.

  - Type: String
  - Default: undefined / null
  - Optional

### `--root-path:'..'`

Relative path where the project should be created (blank = current directory). If specifying a subfolder, use a relative path. For instance, if you are currently in the `MyProject` folder specify `src/public` to create the addin in `MyProject/src/public`.

  - Type: String
  - Default: undefined / null
  - Optional  

### `--tech:[ 'html' | 'ng' | 'ng-adal' | 'manifest-only' ]`

Technology to use for the project. The supported options include HTML (`html`), Angular (`ng`) or Angular ADAL (`ng-adal`). You can also use Manifest.xml only (`manifest-only`) which will create only the `manifest.xml` for an an Office addin.

  - Type: String
  - Default: undefined / null
  - Optional  

### `--clients: [ 'Document' | 'Workbook' | 'Presentation' | 'Project' ]`

The Microsoft Office client application that can host the add-in. 

> This applies only to task pane or content add-ins.

  - Type: String[]
  - Default: undefined / null
  - Optional  

### `--outlookForm: [ 'mail-read' | 'mail-compose' | 'appointment-read' | 'appointment-compose' ]`

The type of form within Outlook that can host the add-in. 

> This applies only to mail add-ins.

  - Type: String[]
  - Default: undefined / null
  - Optional  
