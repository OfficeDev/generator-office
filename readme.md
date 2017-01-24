# Microsoft Office Project Generator - YO OFFICE!

[![npm version](https://badge.fury.io/js/generator-office.svg)](http://badge.fury.io/js/generator-office)
[![Downloads](http://img.shields.io/npm/dm/generator-office.svg)](https://npmjs.org/package/generator-office)
[![TravisCI Build Status](https://travis-ci.org/OfficeDev/generator-office.svg)](https://travis-ci.org/OfficeDev/generator-office)
[![Dependency Status](https://david-dm.org/officedev/generator-office.svg)](https://david-dm.org/officedev/generator-office)
[![devDependency Status](https://david-dm.org/officedev/generator-office/dev-status.svg)](https://david-dm.org/officedev/generator-office#info=devDependencies)

[Yeoman](http://yeoman.io) generator for creating Microsoft Office projects using any text editor. Microsoft includes fantastic & [rich development tools for creating Office related projects using Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013) or [tools for Visual Studio 2015](http://aka.ms/OfficeDevToolsForVS2015). This generator is for those developers who:

- use an editor other than Visual Studio
- are interested in using a technology other than plain HTML, CSS & JavaScript

Like other Yeoman generators, this simply creates the scaffolding of files for your Office project. It allows you to create Add-ins for:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word

Choose to create the Office projects using plain HTML, CSS & JavaScript (*mirroring the same projects that Visual Studio creates*) or create Angular-based projects.

Check out the announcement blog post: [Office Dev Center Blog - Creating Office Add-ins with any editor - Introducing YO OFFICE!](http://dev.office.com/blogs/creating-office-add-ins-with-any-editor-introducing-yo-office) 

Read up on [how to use the generator to create Office Add-ins with Visual Studio Code](https://code.visualstudio.com/Docs/runtimes/office).

If you are interested in contributing, read the the [Contributing Guidelines](CONTRIBUTING.md). 

## YO Office Demo (screenshot & video)
![](src/docs/assets/generatoroffice.png)

<iframe width="560" height="315" src="https://www.youtube.com/embed/78b18BLVosM" frameborder="0" allowfullscreen></iframe>

***

## Install

> **Wait!** 

> Is this the first time you're using Yeoman or installing a Yeoman generator? When working with Yeoman there are a few common prerequisites. Ensure you have already have a copy of the popular source control solution [Git](https://git-scm.com/download) installed.

> If you don't have git installed, once you install it we recommend you restart your console (or if on Windows, restart your machine) as system environment variables are set/updated during this installation.

Install `yo` (Yeoman) and `generator-office` globally using NPM (this also requires [Node.js](https://nodejs.org). 

In the v1.0.0 release we added TypeScript type definitions for autocompletion / IntelliSense... for this you need to install typings.

```bash
$ npm install -g typings yo generator-office
```

## Usage

```bash
$ yo office [arguments] [options]
```

The generator is intended to be run from within a folder where you want the project scaffolding created. This can be in the root of the current folder or within a subfolder.

> Note: Office Add-ins must be hosted, even in development, in a **HTTPS** site. Refer to the section [Running the Generated Site](/OfficeDev/generator-office#running-the-generated-site) below for details.

## Running the Generated Site

All generators create a `bsconfig.json`. This uses [Browsersync](https://browsersync.io/) to make your tweaking and testing faster by synchronising file changes and interactions across multiple devices. 

Start developing and launch a local HTTPS site on `https://localhost:3000` by simply typing the following command in your console:

```bash
$ npm start
```

Browsersync will start a HTTPS server, which includes a self-signed SSL cert that your development environment must trust. Refer to our doc [Adding Self-Signed Certificates as Tusted Root Certificate](src/docs/ssl.md) for instructions on how to do this.

Browse to the 'External' IP address listed in your console to test your web app across multiple browsers and devices that are connected on your local network.

## Validate manifest.xml

Refer to the docs on [Add-in manifests](https://dev.office.com/docs/add-ins/overview/add-in-manifests) for information of manifest validation.

## Examples

Refer to the [docs](src/docs) for example executions & output of the generator.

## Command Line Arguments:
List of supported arguments. The generator will prompt you accordingly based on the arguments you provided.

### `name`
Title of the project - this is the display name that is written the manifest.xml file.
  - Type: String
  - Optional

### `host`
The Microsoft Office client application that can host the add-in. The supported arguments include Excel (`excel`), OneNote (`onenote`), Outlook (`outlook`), PowerPoint (`powerpoint`), Project (`project`), and Word (`word`).
  - Type: String
  - Optional

### `framework`
Framework to use for the project. The supported arguments include JQuery (`jquery`), and Angular (`ng`). You can also use Manifest.xml only (`manifest-only`) which will create only the `manifest.xml` for an an Office addin.
  - Type: String
  - Optional

## Command Line Options:
List of supported options. If these are not provided, the generator will prompt you for the values before scaffolding the project.

### `--skip-install`

After scaffolding the project, the generator (and all sub generators) run all package management install commands such as `npm install` & `bower install`. Specifying `--skip-install` tells the generator to skip this step.

  - Type: Boolean
  - Default: False
  - Optional

### `--js`

Specifying `--js` tells the generator to use Javascript.

  - Type: Boolean
  - Default: False
  - Optional

Copyright (c) 2017 Microsoft Corporation. All rights reserved.