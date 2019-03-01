# Adding Self-Signed Certificates as Trusted Root Certificate

Office clients require add-ins and webpages to come from a trusted and secure location. This generator leverages [Browsersync](https://browsersync.io/) to start a web server, which requires a self-signed certificate. Your workstation will not trust this certificate and thus, the Office client, in which you are running your Office Add-in, will not load your add-in.

When you browse to a site that has an untrusted certificate, the browser will display an error with the certificate:
  		  
  ![](assets/ssl-chrome-error.png)
   
  ![](assets/ssl-edge-error.png)
   
To fix this, you need to configure your developer workstation to trust the self-signed certificate. The steps for this differ depending on your developer environment (OSX / Windows / Linux). Use these instructions to trust the certificate:

## Table of Contents

* [macOS](#macOS)
* [Windows](#windows)

## macOS

1. In Chrome, open the [Settings](chrome://settings) window.
2. Expand the **Advanced** section and choose **Manage certficiates**.
    
    The **Keychain Access** utility will open.
    
3. Double-click the **localhost-ca** certificate.
4. In the **Trust** section, set the following value
    
    **When using this certificate**: **Always Trust**
    
5. Close the dialog.
6. You'll be prompted for your credentials and will need to enter them to enable the certificate
  
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
