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
  * [Get certificate file from project directory](#get-certificate-file-from-project-directory-1)
  * [Add certification file through IE](#add-certification-file-through-ie)
  * [Add certification file through Management Console](#add-certification-file-through-management-console)

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

#### Get certificate file from project directory

You can locate the server.crt file at **~/your_yo_office_project/certs/server.crt**

#### Add certification file through IE

1. Start IE with your local host url.
1. Select **Continue to this website (not recommended).**
![](assets/ssl-ie-01.png)
1. Click **Certificate Error** at the end of the address bar.
![](assets/ssl-ie-02.png)
1. In the drop down window, click **View certificates** at the bottom.
![](assets/ssl-ie-03.png)
1. In the Certificate window, click **Install Certificate...** button then click **OK**.
![](assets/ssl-ie-04.png)
1. This will take you to Certificate Import Wizard. Select **Local Machine** for store location then click **Next**.
![](assets/ssl-ie-05.png)
1. System now will prompt you security information. Click **Yes** to allow changes.
![](assets/ssl-ie-06.png)
1. In the Wizard, select **Place all certificates in the following store** and then click **Browse...** button.
![](assets/ssl-ie-07.png)
1. Select **Trusted Root Certification Authorities** folder and click **OK**.
![](assets/ssl-ie-08.png)
1. Now make sure **Trusted Root Certification Authorities** is in place for **Certificate store** and click **Next**.
![](assets/ssl-ie-09.png)
1. Your cert is now ready to be imported. Click **Finish**.
![](assets/ssl-ie-10.png)

You now have a self-signed certificate installed on your machine.

#### Add certification file through Management Console

1. Click **Start**, go to **Search box**, type **run**, then type **mmc** in the input field, and then press **ENTER**.
2. On the **File** menu, click **Add/Remove Snap-in**.

![](assets/ssl-windows-01.png)

3. Under **Available snap-ins**, click **Certificates**, and then click **Add**.
![](assets/ssl-windows-02.png)

4. Under **This snap-in will always manage certificates for**, click **Computer account**, and then click **Next**.
![](assets/ssl-windows-03.png)

5. Click **Local computer**, and click **Finish**.
![](assets/ssl-windows-04.png)
6. If you have no more snap-ins to add to the console, click **OK**.
7. In the console tree, double-click **Certificates**.
8. Right-click the **Trusted Root Certification Authorities** store.
9. Click **All Tasks** then **Import** to import the certificates and follow the steps in the Certificate Import Wizard.
![](assets/ssl-windows-05.png)
10. Find server.crt file and click **Open**.
11. Follow the rest of the steps to finish Certificate import.

You now have a self-signed certificate installed on your machine.

Copyright (c) 2017 Microsoft Corporation. All rights reserved.
