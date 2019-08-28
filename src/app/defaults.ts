import * as chalk from "chalk";

export const usageDataProjectName: string = "generator-office";
export const usageDataPromptMessage: string = `Office Add-in CLI tools collect anonymized usage data which is sent to Microsoft to help improve our product. Please read our privacy notice at ${chalk.blue('https://aka.ms/OfficeAddInCLIPrivacy')}. ​To disable data collection, choose Exit and run ${chalk.green('“npx office-addin-usage-data off”')}.\n\n`;
export const usageDataInstrumentationKey: string = "9fdc4e3b-fdd0-4c52-b640-d8e2ba5c9e59";
