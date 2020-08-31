import * as chalk from "chalk";

export const configurationErrorEventName: string = "configuration-error-generator-office";
export const copyFilesErrorEventName: string = "copy-files-error-generator-office";
export const installDependenciesErrorEventName = "install-dependencies-error-generator-office";
export const networkShareSideloadingSteps: string = "https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins";
export const outlookSideloadingSteps = "https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing";
export const postInstallHintsErrorEventName = "post-install-hints-error-generator-office";
export const promptSelectionstEventName: string = "prompt-selections-generator-office";
export const promptSelectionsErrorEventName: string = "prompt-selections-error-generator-office";
export const usageDataProjectName: string = "generator-office";
export const usageDataPromptMessage: string = `Office Add-in CLI tools collect anonymized usage data which is sent to Microsoft to help improve our product. Please read our privacy notice at ${chalk.blue('https://aka.ms/OfficeAddInCLIPrivacy')}. ​To disable data collection, choose Exit and run ${chalk.green('“npx office-addin-usage-data off”')}.\n\n`;
