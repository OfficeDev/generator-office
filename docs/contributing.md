# Contributing Guidlines
 
If you would like to become an active contributor to this project please read the following guidelines.

In general, if it's going to take a long time to review a PR, we'll likely request you break it up.

Please feel free to reach out to the team if you have any questions about contributing. You can create an issue if you are having problems and @ mention [@jthake](/jthake) & [@andrewconnell](/andrewconnell).

## Fixing Typos

Typos are embarrasing! Most PR's that fix typos will be accepted. In order to make it easier to review the PR, please narrow the focus instead of sending a huge PR of fixes.

## Reporting Bugs & Issues
If you have any bugs or issues with the generator. Please submit them in the [Issues](/OfficeDev/generator-office/issues) for this repo.

## Commit Messages

Please format commit messages as follows (based on this [excellent post](http://tbaggery.com/2008/04/19/a-note-about-git-commit-messages.html)):

```
Summarize change in 50 characters or less

Provide more detail after the first line. Leave one blank line below the
summary and wrap all lines at 72 characters or less.

If the change fixes an issue, leave another blank line after the final
paragraph and indicate which issue is fixed in the specific format
below.

Fix #42
```

Do your best to factor commits appropriately, i.e not too large with unrelated
things in the same commit, and not too small with the same small change applied N
times in N different commits. If there was some accidental reformatting or whitespace
changes during the course of your commits, please rebase them away before submitting
  the PR.

## DO's & DON'Ts

- **DO** follow the coding style described in the [Developer Guide](developer-guide.md)
- **DO** follow the same project and test structure as the existing project
- **DO** include tests when adding new functionality and features. When fixing bugs, start with adding a test that highlights how the current behavior is broken.
- **DO** keep discussions focused. When a new or related topic comes up it's often better to create new issue than to side track the conversation.
- **DO NOT** submit PR's for coding style changes.
- **DO NOT** surprise us with big PR's. Instead file an issue & start a discussion so we can agree on a direction before you invest a large amount of time.
- **DO NOT** commit code you didn't write.
- **DO NOT** submit PR's that refactor existing code without a discussion first. 

## Submitting Feature Requests & Design Change Requests
Feature requests and Design Change Requests (DCRs) are an important part of the lifecycle of any software project. Please log these as [Issues](/OfficeDev/generator-office/issues) in the repo. 

When opening any feature requests, consider including as much information as possible, including: 

- Detailed scenarios enabled by the feature or DCR.
- Information about your use case or additional value the feature would provide.
- Make note of whether you are opening an issue you would like the Microsoft team or another community member to work on or if you are looking to design & develop the feature yourself.
- Any potential caveats or concerns you may have already thought about.
- A miniature test plan or list of test scenarios is always helpful.