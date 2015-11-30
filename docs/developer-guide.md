# Developer Guide

This describes the guidelines developers should follow when making contributions to this repo.  

## Development Tasks

Gulp is used to automate some test and style checks within this repo. To view a list of all the tasks available, run the following from the root of the repo:

```bash
$ gulp
```

For details on each task, refer to the `gulpfile.js`.  Each task is commented with [jsdoc](http://usejsdoc.org/) a description of the tasks purpose.

## Coding Style

All code must pass the [JSHint](http://www.jshint.com) & [JSCS](http://jscs.info) rules defined in this project which are defined in the following config files:

- JSHint Settings: `.jshintrc`
- JSCS Settings: `.jscsrc`

Check all source files to ensure they meet these guidelines using the provided gulp task **vet**:

```
$ gulp vet
```

While coding, you can automatically run this using the **autovet** task:

```
$ gulp autovet
```

To simplify the task of formatting your code, you can use the **JSCS** NPM CLI provided in the developer dependencies. This does not apply all code fixes, just the things that can be automatically fixed by the JSCS CLI:

```
$ node_modules/jscs/bin/jscs [path-to-js-files] --fix
```

You can also pass multiple files in using the following:

```
$ node_modules/jscs/bin/jscs test/taskpane/*.js --fix
```

## Documentation

- All non-trival functions should have a [jsdoc](http://usejsdoc.org/) description