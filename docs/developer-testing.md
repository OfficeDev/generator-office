# Developer Testing

This describes the guidelines developer should follow to run & create tests.

- All code must have valid, **passing** unit tests
- All code should have 100% test coverage or good reasons why it doesn't meet 100% coverage 

## Testing Overview

- [Mocha](http://mochajs.org) is used as the test framework for this project.
- [Chai](http://chaijs.com/) is used as the assertion library for all tests.
  - Tests are written using the [Expect BDD](http://chaijs.com/guide/styles/#expect) style, a chainable language.

Run all tests using the provided gulp task **test**:

```
$ gulp test
```

While coding, you can automatically run this using the **autotest** task:

```
$ gulp autotest
```

The gulp tasks referenced above use the **spec** [Mocha Reporter](https://mochajs.org/#reporters) which lists all tests as they are run. When developing you might want to eliminate the list of all the tests and use the **dot** reporter which can be done using the following command:

```bash
$ mocha -R dot test/**/*.js
```

To setup a watcher and automatically run all tests while coding using this reporter, add the `-w` flag to the line above.

## Test Validation

When PR's are submitted, all tests must pass before they will be reviewed. This is done automatically using [Travis CI](https://travis-ci.org/OfficeDev/generator-office). Once a PR is submitted, Travis kicks in and automatically runs all tests for the specified platforms. If there are any failures, you should address them and commit changes to your PR until your tests pass.

## Code Coverage

Any new or changed code should be adequately covered by unit tests. Like tests, code coverage is run automatically when a PR is submitted using [Coveralls](https://coveralls.io/github/OfficeDev/generator-office?branch=master). Any PR's that lower the code coverage % with new code submission will be scrutinized before being merged. Depending on the reason for the drop in code coverage %, the PR may be rejected or the submitter may be requested to address it with more tests.

Coverage reports are also generated every time tests are run locally. To view a code coverage report, after running tests, open the code coverage report found in `/coverage/lcov-report/index.html` in a browser.