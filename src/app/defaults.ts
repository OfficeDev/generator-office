import * as chalk from "chalk";

export const configurationErrorEventName = "configuration-error-generator-office";
export const copyFilesErrorEventName = "copy-files-error-generator-office";
export const installDependenciesErrorEventName = "install-dependencies-error-generator-office";
export const networkShareSideloadingSteps = "https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins";
export const outlookSideloadingSteps = "https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing";
export const postInstallHintsErrorEventName = "post-install-hints-error-generator-office";
export const promptSelectionstEventName = "prompt-selections-generator-office";
export const promptSelectionsErrorEventName = "prompt-selections-error-generator-office";
export const usageDataProjectName = "generator-office";
export const usageDataPromptMessage = `Office Add-in CLI tools collect anonymized usage data which is sent to Microsoft to help improve our product. Please read our privacy notice at ${chalk.blue('https://aka.ms/OfficeAddInCLIPrivacy')}. ​To disable data collection, choose Exit and run ${chalk.green('“npx office-addin-usage-data off”')}.\n\n`;
