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

## Prerequisites

* Install Office 2016 for Windows and join the [Office Insider](https://products.office.com/en-us/office-insider) program. You must have Office build number 8711 or later.

## To use the project

Follow these instructions to use this custom function sample add-in:

1. Publish the code files (HTML, JavaScript) in the same folder on a website.
2. Replace `https://<INSERT-URL-HERE>` in the manifest file (there should be 3 occurrences) with the URL of your website. 
3. Sideload the manifest using the instructions found at <https://aka.ms/sideload-addins>.
4. Test a custom function by entering `=CONTOSO.ADD42(3)` in a cell.
5. Try the other functions in the sample: `CONTOSO.ISEVEN(num)`, `CONTOSO.GETDAY()`, `CONTOSO.INCREMENTVALUE(increment, caller)`, and `CONTOSO.SECONDHIGHEST(range)`.
6. If you make changes to the sample add-in, copy the updated files to your website, and then close and reopen Excel. If your functions are not available in Excel, re-insert the add-in using **Insert** > **My Add-ins**.
7. Follow @OfficeDev on Twitter for updates and send feedback to <excelcustomfunctions@microsoft.com>.

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