# Adding Self-Signed Certificates as Trusted Root Certificate

Office clients require add-ins and webpages to come from a trusted and secure location. This generator leverages [Browsersync](https://browsersync.io/) to start a web server, which requires a self-signed certificate. Your workstation will not trust this certificate and thus, the Office client, in which you are running your Office Add-in, will not load your add-in.

When you browse to a site that has an untrusted certificate, the browser will display an error with the certificate:
  		  
  ![](assets/ssl-chrome-error.png)
   
  ![](assets/ssl-edge-error.png)
   
To fix this, you need to configure your developer workstation to trust the self-signed certificate. The steps for this differ depending on your developer environment (OSX / Windows / Linux). Use these instructions to trust the certificate:

## Table of Contents

* [OS X](#os-x)
  * [Get certificate in Chrome](#get-certificate-in-chrome)
  * [Get certificate file from project directory](#get-certificate-file-from-project-directory)
  * [Add certification file to Key Chain Access](#add-certification-file-to-key-chain-access)
* [Windows](#windows)

## [OS X](https://support.apple.com/kb/PH18677)

#### Get certificate in Chrome

1. Start Chrome and do the following:
   1. Open Developer Tools window by using keyboard shortcuts: Cmd + Opt + I.
   1. Click to go to 'security' panel and 'overview' screen.
	 1. Click 'View certificate'. 

   ![](assets/ssl-chrome-devtool.png)

1. Click and drag the image to your desktop. It looks like a little certificate.
![](assets/ssl-chrome-getcert.png)

#### Get certificate file from project directory

You can locate the server.crt file at **~/your_yo_office_project/certs/server.crt**

#### Add certification file to Key Chain Access

1. Open the **Keychain Access** utility in OS X.
   1. Select the **System** option on the left.
   1. Click the lock icon in the upper-left corner to enable changes.
   ![](assets/ssl-keychain-01.png)

   1. Click the plus button at the bottom and select the **localhost.cer** file you copied to the desktop.
   1. In the dialog that comes up, click **Always Trust**.
   1. After **localhost** gets added to the **System** keychain, double-click it to open it again.
   1. Expand the **Trust** section and for the first option, pick **Always Trust**.

  ![](assets/ssl-keychain-02.png)
  
At this point everything has been configured. Quit all browsers, then reopen and try to navigate to the local HTTPS site. The browser should report it as a valid certificate:

  ![](assets/ssl-chrome-good.png)

## [Windows](https://technet.microsoft.com/en-us/library/cc754841.aspx)

Take the following steps to setup the certificate authority cert for localhost:

1.	Go to {project root}\certs.
2.	Double-click ca.crt, and select **Install Certificate**.
       
![](assets/ssl-ie-04.png)

3.	Select **Local Machine** and select **Next** to continue.

![](assets/ssl-ie-05.png)

4.	Select **Place all certificates in the following store** and then select **Browse**.
5.	Select **Trusted Root Certification Authorities** and then select **OK**.
6.	Select **Next** and then **Finish**.

You now have a self-signed certificate installed on your machine.

Copyright (c) 2017 Microsoft Corporation. All rights reserved.
