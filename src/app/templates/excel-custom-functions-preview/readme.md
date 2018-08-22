# Custom functions in Excel (Preview)

Learn how to use custom functions in Excel (similar to user-defined functions, or UDFs). Custom functions are JavaScript functions that you can add to Excel, and then use them like any native Excel function (for example =Sum). This sample accompanies the [Custom Functions Overview](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview) topic.

## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [To use the project](#to-use-the-project)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

* Oct 27, 2017: Initial version.
* April 23, 2018: Revised and expanded.
* June 1, 2018: Bug fixes.

## Prerequisites

* Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program. You must have Office build number 8711 or later.

## To use the project

On a machine with a valid instance of an Excel Insider build installed, follow these instructions to use this custom function sample add-in:

1. On the machine where your custom functions project is installed, follow the instructions to install the self-signed certificates (https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) . 
2. From a command prompt from within your custom functions project directory, run `npm run start` to start a localhost server instance. 
3. Run `npm run sideload` to launch Excel and load the custom functions add-in. Additonal information on sideloading can be found at <https://aka.ms/sideload-addins>.
4. After Excel launches, you will need to register the custom-functions add-in to work around a bug:
    a. On the upper-left-hand side of Excel, there is a small hexagon icon with a dropdown arrow. The icon is to right of the Save icon.
    b. Click on this dropdown arrow and then click on the Custom Functions Sample add-in to register it.
5. Test a custom function by entering `=CONTOSO.ADD42(num)` in a cell.
6. Try the other functions in the sample: `=CONTOSO.ADD42ASYNC(num, num)`, `CONTOSO.ISPRIME(num)`, `CONTOSO.NTHPRIME(num)`, `CONTOSO.GETDAY()`, `CONTOSO.INCREMENTVALUE(increment)`, and `CONTOSO.SECONDHIGHEST(range)`.
7. If you make changes to the sample add-in, copy the updated files to your website, and then close and reopen Excel. If your functions are not available in Excel, re-insert the add-in using **Insert** > **My Add-ins**.
8. Follow @OfficeDev on Twitter for updates and send feedback to <excelcustomfunctions@microsoft.com>.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Custom Functions Overview](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview)
* [Office add-in documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright
Copyright (c) 2017 Microsoft Corporation. All rights reserved.