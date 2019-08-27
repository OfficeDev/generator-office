import * as chalk from "chalk";

export const usageDataProjectName: string = "generator-office";
export const usageDataPromptMessage: string = `Office Add-in CLI tools collect anonymized usage data which is sent to Microsoft to help improve our product. Please read our privacy notice at ${chalk.blue('https://aka.ms/OfficeAddInCLIPrivacy')}. ​To disable data collection, choose Exit and run ${chalk.green('“npx office-addin-usage-data off”')}.\n\n`;
export const usageDataInstrumentationKey: string = "1ced6a2f-b3b2-4da5-a1b8-746512fbc840";
