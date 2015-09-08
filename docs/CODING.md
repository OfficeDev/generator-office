# Coding Conventions & Gudelines

Refer to the existing code & tests for examples of coding guidelines for this project.

All code must pass the [JSHint](http://www.jshint.com) & [JSCS](http://jscs.info) rules defined in this project which are defined in the following config files:

- JSHint Settings: [.jshintrc](/OfficeDev/generator-office/.jshint)
- JSCS Settings: [.jscsrc](/OfficeDev/generator-office/.jscsrc)

## Coding Style

Check all source files to ensure they meet these guidelines using the provided gulp task **vet**:

```
$ gulp vet
```

While coding, you can automatically run this using the **autovet** task:

```
$ gulp autovet
```

To simplify the task of formatting your code, you can use the **JSCS** command line NPM module provided in the developer dependencies. This does not do all code fixes, just the things that can be automatically fixed by the JSCS CLI

```
$ node_modules/jscs/bin/jscs [path-to-js-files] --fix
```

You can also pass multiple files in using the following:

```
$ node_modules/jscs/bin/jscs test/taskpane/*.js --fix
```

## Testing

- All code must have valid, **passing** unit tests
- All code should have 100% test coverage or good reasons why it doesn't meet 100% coverage 

Run all tests using the provided gulp task **test**:

```
$ gulp test
```

While coding, you can automatically run this using the **autotest** task:

```
$ gulp autotest
```

To generate a code coverage report, after running tests, open the code coverage report found in `/coverage/lcov-report/index.html` in a browser.

## Documentation

- All non-trival functions should have a [jsdoc](http://usejsdoc.org/) description