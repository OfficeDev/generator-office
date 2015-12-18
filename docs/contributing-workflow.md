# Contribution Workflow

You can contrubute to this repo with issues & PRs. Simply filing issues for problems you encounter is a great way to contribute. Contributing bug fixes & features is greatly appreciated!

## Getting Started

For new ideas, always start with an issue before starting development on an implementation.

You do not need to file an issue for trival changes such as typo fixes. Just a create a PR for those changes.

## PR - CI Process

The continuous integration (CI) system will automatically perform the required builds and run tests (including the ones you are expected to run) for PRs. Builds and test runs must be clean.

If the CI build fails for any reason, the PR issue will be updated with a link that can be used to determine the cause of the failure. This is done using Travis CI.

## Suggested Workflow

The following is recommended for contributing to the repo:

* Create an issue for your work.
  - This can be skipped for trival changes.
  - Reuse an existing issue on the topic, if there is one.
  - Get an agreement from core team members before starting work.
  - Clearly state you are going to be implementing a fix, if that's the case.
* Create a personal fork of the repository on GitHub (if you don't already have one).
  - If you already have a fork, ensure you have an upstream remote for the **DEV** branch and it is up to date with the repo (this repo). Refer to this help doc for more information on syncing your repo: [GitHub Help: Syncing a Fork](https://help.github.com/articles/syncing-a-fork/)
* Create a branch off master: `git checkout -b mybranch`
  - Name the branch so that it clearly communicates your intentions such as **issue-123** or **githubhandle-issue123**
  - Branches are useful since they isolate your changes from incoming changes from upstream. They also enable you to create multiple PRs from the same fork.
* Make & commit your changes
* Add new tests corresponding to your change if applicable.
* Build the repo with your changes.
  - Make sure that the builds are clean.
  - Make sure that the tests are all passing, including your new tests.
  - Make sure the code coverage % is no lower than the upstream repo.
* Push changes to your fork on GitHub.
* Create a pull request (PR) against the upstream repository's **DEV** branch.
  - Include details in the PR comment indicating what the PR is for. Include a reference to the issue (by number & link).

**Note: It is OK to create your PR as a "WIP" (work in progress) on the upstream repo before the implementation is done. This can be useful if you'd like to start the feedback process concurrent with your implementation. State this this is the case in the initial PR comment.