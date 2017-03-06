# Contribute to Microsoft Office Project Generator

Thank you for your interest in this library! Your contributions and improvements will help the developer community.
* [Ways to contribute](https://github.com/OfficeDev/generator-office/blob/master/CONTRIBUTING.md#ways-to-contribute)
* [Before we can accept your pull request](https://github.com/OfficeDev/generator-office/blob/master/CONTRIBUTING.md#before-we-can-accept-your-pull-request)
* [Use GitHub, Git, and this repository](https://github.com/OfficeDev/generator-office/blob/master/CONTRIBUTING.md#use-github-git-and-this-repository)
* [More resources](https://github.com/OfficeDev/generator-office/blob/master/CONTRIBUTING.md#more-resources)

## Ways to contribute
You can contribute to Office Project Generator in these ways:
* Provide feedback
  * Report bugs and suggest enhancements via [GitHub Issues](https://github.com/OfficeDev/generator-office/issues)
* Do it yourself
  * Fix [Issues](https://github.com/OfficeDev/generator-office/issues) yourself and submit the changes as a [Pull Request](https://github.com/OfficeDev/generator-office/pulls) for review.
  * You can submit [code comment contributions](https://github.com/OfficeDev/generator-office/blob/master/CONTRIBUTING.md#provide-better-code-comments) where you want a better explanation of the code.

## Before we can accept your pull request
If you are in one of the following groups, you need to send us a signed Contribution License Agreement (CLA) before we can accept your pull request:
* Members of the Microsoft Open Technologies group
* Contributors who don't work for Microsoft

As a community member, you must sign the Contribution License Agreement (CLA) before you can contribute large submissions to this project, but you need to complete and submit the documentation only once. The Office 365 organization on GitHub will send a link to the CLA that we want you to sign via email. By signing the CLA, you acknowledge the rights of the GitHub community to use any code that you submit. The intellectual property represented by the code contribution is licensed for use by Microsoft open source projects. Please carefully review the document, as you may also need to have your employer sign the document.

Signing the Contribution License Agreement (CLA) does not grant you rights to commit to the main repository, but it does mean that the Office Developer and Office Developer Content Publishing teams will be able to review and consider your contributions and you will get credit if we do.

Once we receive and process your CLA, we'll do our best to review your pull requests within 10 business days.

## Use GitHub, Git, and this repository
**Note**: Most of the information in this section can be found in [GitHub Help](https://help.github.com/) articles. If you're familiar with Git and GitHub, skip to the Contribute code section for the specifics of the code contributions for this repository.

### To set up your fork of the repository
1. Set up a GitHub account so you can contribute to this project. If you haven't done this, go to [GitHub](https://github.com/join) and do it now.
1. Install Git on your computer. Follow the steps in the Setting up Git Tutorial.
1. Create your own fork of this repository. To do this, at the top of the page, choose the Fork button.
1. Copy your fork to your computer. To do this, open Git Bash. At the command prompt enter:

```
git clone https://github.com/<your_user_name>/<repo_name>.git
```
Next, create a reference to the root repository by entering these commands:
```
cd <repo_name>
git remote add upstream https://github.com/OfficeDev/<repo_name>.git
git fetch upstream
```
Congratulations! You've now set up your repository. You won't need to repeat these steps again.

### Provide better code comments
Code comments make code samples even better by helping developers learn to use the code correctly in their own applications. If you spot a class, method, or section of code that you think could use better descriptions, then create a pull request with your code comments.
In general we want our code comments to follow these guidelines:

* Any code that has associated documentation displayed in an IDE (such as IntelliSense, or JavaDocs) has code comments.
* Classes, methods, parameters, and return values have clear descriptions.
* Exceptions and errors are documented.
* Remarks exist for anything special or notable about the code.
* Sections of code that have complex algorithms have appropriate comments describing what they do.
* Code added from Stack Overflow, or any other source, is clearly attributed.

### Contribute code
To make the contribution process as seamless as possible for you, follow this procedure.

1. Create a new branch.
1. Add new content or edit existing content.
1. Submit a pull request to the main repository.
1. Delete the branch.

Limit each branch to a single module to streamline the workflow and reduce the chance of merge conflicts. The following types of contribution are appropriate for a new branch:

* A correction to the slide deck
* Instruction step fixes or additional clarification in hands on labs
* Code fixes in sample starter or completed projects
* Spelling and grammar edits on a hands on lab

#### Create a new branch
1. Open GitBash.
1. Type `git pull upstream master:<new_branch_name>` at the prompt. This creates a new branch locally that's copied from the latest OfficeDev master branch. **Note**: For internal contributors, replace `master` in the command with the branch for the publishing date you're targeting.
1. Type `git push origin <new_branch_name>` at the prompt. This will alert GitHub to the new branch. You should now see the new branch in your fork of the repository on GitHub.
1. Type `git checkout <new_branch_name>` to switch to your new branch.

#### Add new content or edit existing content
Navigate to the repository on your computer. On a Windows PC, the repository files are in `C:\Users\<yourusername>\<repo_name>`.
Use the IDE of your choice to modify and build the library. Once you have completed your change, commented your code, and test, check the code into the remote branch on GitHub.

Be sure to satisfy all of the requirements in the following list before submitting a pull request:
* Follow the code style found in the cloned repository code.
* Code must be tested.
* Test the library UI thoroughly to be sure nothing has been broken by your change.
Keep the size of your code change reasonable. If the repository owner cannot review your code change in 4 hours or less, your pull request may not be reviewed and approved quickly.
* Avoid unnecessary changes to cloned or forked code. The reviewer will use a tool to find the differences between your code and the original code. Whitespace changes are called out along with your code. Be sure your changes will help improve the content.

#### Push your code to the remote GitHub branch
The files in `C:\Users\<yourusername>\<repo_name>` are a working copy of the new branch that you created in your local repository. Changing anything in this folder doesn't affect the local repository until you commit a change. To commit a change to the local repository, type the following commands in GitBash:
```
git add .
git commit -v -a -m "<Commit_description>"
```
The `add` command adds your changes to a staging area in preparation for committing them to the repository. The period after the `add` command specifies that you want to stage all of the files that you added or modified, checking subfolders recursively. (If you don't want to commit all of the changes, you can add specific files. You can also undo a commit. For help, type `git add -help` or `git status`.)

The `commit` command applies the staged changes to the repository. The switch `-m` means you are providing the commit comment in the command line. The `-v` and `-a` switches can be omitted. The `-v` switch is for verbose output from the command, and `-a` does what you already did with the add command.

You can commit multiple times while you are doing your work, or you can commit once when you're done.

#### Submit a pull request to the main repository

When you're finished with your work and are ready to have it merged into the central repository, follow these steps.

1. In GitBash, type `git push origin <new_branch_name>` at the command prompt. In your local repository, origin refers to your GitHub repository that you cloned the local repository from. This command pushes the current state of your new branch, including all commits made in the previous steps, to your GitHub fork.
1. On the GitHub site, navigate in your fork to the new branch.
1. Click the **Pull Request** button at the top of the page.
1. Ensure that the Base branch is `OfficeDev/<repo_name>@master` and the Head branch is `<yourusername>/<repo_name>@<branch_name>`.
1. Click the **Update Commit Range** button.
1. Give your pull request a Title, and describe all the changes you're making. If your bug fixes a UserVoice item or GitHub issue, be sure to reference that issue in the description.
1. Submit the pull request.

One of the site administrators will now process your pull request. Your pull request will surface on the `OfficeDev/<repo_name>` site under Issues. When the pull request is accepted, the issue will be resolved.

#### Create a new branch after merging
After a branch is successfully merged (i.e., your pull request is accepted), don't continue working in the local branch that was successfully merged upstream. This can lead to merge conflicts if you submit another pull request. Instead, if you want to do another update, create a new local branch from the successfully merged upstream branch.

For example, suppose your local branch X was successfully merged into the OfficeDev/generator-office master branch and you want to make additional updates to the content that was merged. Create a new local branch, X2, from the OfficeDev/generator-office master branch. To do this, open GitBash and execute the following commands:
```
cd <repo name>
git pull upstream master:X2
git push origin X2
```
You now have local copies (in a new local branch) of the work that you submitted in branch X. The X2 branch also contains all the work other developers have merged, so if your work depends on others' work (for example, a base class), it is available in the new branch. You can verify that your previous work (and others' work) is in the branch by checking out the new branch...
```
git checkout X2
```
...and verifying the code. (The `checkout` command updates the files in `C:\Users\<yourusername>\generator-office` to the current state of the X2 branch.) Once you check out the new branch, you can make updates to the code and commit them as usual. However, to avoid working in the merged branch (X) by mistake, it's best to delete it (see the following **Delete a branch** section).

#### Delete a branch
Once your changes are successfully merged into the central repository, you can delete the branch you used because you no longer need it. Any additional work requires a new branch.

To delete your branch follow these steps:

1. In GitBash type `git checkout master` at the command prompt. This ensures that you aren't in the branch to be deleted (which isn't allowed).
1. Next, type `git branch -d <branch_name>` at the command prompt. This deletes the branch on your local machine only if it has been successfully merged to the upstream repository. (You can override this behavior with the `–D` flag, but first be sure you want to do this.)
1. Finally, type `git push origin :<branch_name>` at the command prompt (a space before the colon and no space after it). This will delete the branch on your github fork.

Congratulations, you have successfully contributed to the project.

## More resources
* To learn more about Markdown, see [Daring Fireball](http://daringfireball.net/).
* To learn more about using Git and GitHub, check out the [GitHub Help section](http://help.github.com/).
