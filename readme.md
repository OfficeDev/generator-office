# Microsoft Office Add-in Project Generator - YO OFFICE!

[![npm version](https://badge.fury.io/js/generator-office.svg)](http://badge.fury.io/js/generator-office)
[![Downloads](http://img.shields.io/npm/dm/generator-office.svg)](https://npmjs.org/package/generator-office)
[![TravisCI Build Status](https://travis-ci.org/OfficeDev/generator-office.svg)](https://travis-ci.org/OfficeDev/generator-office)

[Yeoman](http://yeoman.io) generator for creating [Microsoft Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/) projects using any text editor. Microsoft includes fantastic & [rich development tools for creating Office related projects using Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013) or [tools for Visual Studio 2015](http://aka.ms/OfficeDevToolsForVS2015). This generator is for those developers who:

- use an editor other than Visual Studio
- are interested in using a technology other than plain HTML, CSS & JavaScript

> If you are building an Angular or React add-in and would like to learn more about using Yo Office specifically for those frameworks, see [Build an Add-in with React](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react) or [Build an Add-in with Angular](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-angular).

Like other Yeoman generators, this simply creates the scaffolding of files for your Office Add-in project. It allows you to create add-ins for:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word

Choose to create Office Add-in projects using plain HTML, CSS & JavaScript (*mirroring the same projects that Visual Studio creates*) or create Angular-based projects.

## YO Office Demo
![](src/docs/assets/gettingstarted-slow.gif)

## Install

> **Important:** If this is the first time you're using Yeoman or installing a Yeoman generator, first install [Git](https://git-scm.com/download) and [Node.js](https://nodejs.org). For developers on Mac, we recommend using [Node Version Manager](https://github.com/creationix/nvm) to install Node.js with the right permissions. When the installation completes, restart your console (or if you are using Windows, restart your machine) to ensure you use the updated system environment variables.

Install `yo` (Yeoman) and `generator-office` globally using NPM.

```bash
$ npm install -g yo generator-office
```

## Usage

```bash
$ yo office [arguments] [options]
```

The generator is intended to be run from within a folder where you want the project scaffolding created. This can be in the root of the current folder or within a subfolder.

### Command Line Arguments
The following command line arguments are supported. The generator will prompt you accordingly based upon the arguments that you specify.

#### `name`
Title of the project - this is the display name that is written the manifest.xml file.
  - Type: String
  - Optional
 
>**Note:** The Windows command prompt requires this argument to be in quotes (e.g. "My Office Add-in")

#### `host`
The Microsoft Office client application that can host the add-in. The supported arguments include Excel (`excel`), OneNote (`onenote`), Outlook (`outlook`), PowerPoint (`powerpoint`), Project (`project`), Word (`word`) and CustomFunctions (`customfunctions`).
  - Type: String
  - Optional

#### `projectType`
Specifies the type of project to create. The supported arguments include JQuery (`jquery`), Angular (`angular`), and React (`react`). You can also use Manifest (`manifest`) which will create only the `manifest.xml` for an Office Add-in.
  - Type: String
  - Optional

### Command Line Options
The following command line options are supported. If these are not specified, the generator will prompt you for the values before scaffolding the project.

#### `--skip-install`

After scaffolding the project, the generator (and all sub generators) run all package management install commands such as `npm install` & `typings install`. Specifying `--skip-install` tells the generator to skip this step.

  - Type: Boolean
  - Default: False
  - Optional

#### `--ts`

Specifying `--js` tells the generator to use TypeScript.

  - Type: Boolean
  - Default: False
  - Optional

#### `--js`

Specifying `--js` tells the generator to use JavaScript.

  - Type: Boolean
  - Default: False
  - Optional

>**Note:** Do not use this flag when you pass `react` as framework argument.

Specifying `--output` tells the generator to create a project folder with a different name than the add-in name.

  - Type: String
  - Default: Add-in Name
  - Optional

## Running the Generated Site

Office Add-ins must be hosted in an HTTPS site. Yo Office generates a self-signed certificate for use with the development environment. Your computer will need to trust the certificate before you can use the generated add-in.

**Important:** Follow the instructions in [Adding Self-Signed Certificates as Trusted Root Certificate](src/docs/ssl.md) before you start your web application.
  		  
Launch the local HTTPS site on `https://localhost:3000` by simply typing the following command in your console:

```bash
$ npm start
```

Browse to the 'External' IP address listed in your console to test your web app across multiple browsers and devices that are connected on your local network.

![](src/docs/assets/browsersync.gif)

## Validate manifest.xml

As you modify your `manifest.xml` file, use the included [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator) to ensure that your XML file is correct and complete. It will also give you information on against what platforms to test your add-ins before submitting to the store.

To run Office Add-in Validator, use the following command in your project directory:
```bash
$ npm run validate your_manifest.xml
```
![](src/docs/assets/validator.gif)

For more information on manifest validation, refer to our [add-in manifests documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).

## Contributing

### [Contributing Guidelines](CONTRIBUTING.md)

If you are interested in contributing, please start by reading the [Contributing Guidelines](CONTRIBUTING.md).

### Development

#### Prerequisites

Ensure you have [Node.js](https://nodejs.org/en/) installed.

Install [Yeoman](http://yeoman.io/).
```bash
$ npm install -g yo
```

#### Initialize the repo

```bash
$ git clone https://github.com/OfficeDev/generator-office.git
$ cd generator-office
$ npm install
```

#### Make your desired changes

  - Project templates can be found under [src/app/templates](src/app/templates/)
  - Generator script can be found at [src/app/index.ts](src/app/index.ts)

#### Build and link your changes

```bash
$ npm run build
$ npm link
$ cd ..
$ yo office
```

At this point, `yo office` will be running with your custom built `office-generator` changes.

---

Copyright (c) 2017 Microsoft Corporation. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

