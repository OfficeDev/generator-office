# Example - Office Mail Add-in in Angular ADAL

This document demonstrates creating a Mail Add-in first in an empty project as well as in an existing project using [Angular](https://www.angularjs.org) ADAL as the technology.

## Empty Project

This example creates an Office Mail Add-in as Angular ADAL within an empty project folder.

```bash
$ yo office --skip-install
```

### Prompt Responses:

- **Project name (the display name):** My Office Add-in
- **Root folder of the project:** {blank} 
- **Office project type:** Mail Add-in (read & compose forms)
- **Technology to use:** Angular ADAL
- **Application ID as registered in Azure AD:** 03ad2348-c459-4573-8f7d-0ca44d822e7c
- **Supported Outlook forms:** E-Mail message - read form, Appointment - read form

```
.
├── .bowerrc
├── bower.json
├── gulpfile.js
├── jsconfig.json
├── manifest.xml
├── manifest.xsd
├── tsd.json
├── appread
│   ├── app.adalconfig.js
│   ├── app.config.js
│   ├── app.module.js
│   ├── app.routes.js
│   ├── index.html
│   ├── home
│   │   ├── home.controller.js
│   │   └── home.html
│   └── services
│       └── data.service.js
├── content
│   └── Office.css
├── images
│   └── close.png
└── scripts
    └── MicrosoftAjax.js
```

## Existing Project

The generator [nodehttps](https://www.npmjs.com/package/generator-nodehttps) is first used to create a folder for a self-hosted HTTPS site on the local development system:

```bash
$ yo nodehttps
```

### Prompt Responses:

- **What is the name of this project:** Project Name
- **What port will the site run on?**: 8443

### Results:

```
.
├── package.json
└── src
    ├── public
    │   ├── content
    │   │   └── site.css
    │   └── index.html
    └── server
        └── server.js
```

Now run the Office Add-in generator:

```bash
$ yo office --skip-install
```
### Prompt Responses:

- **Project name (the display name):** My Office Add-in
- **Root folder of the project:** src/public 
- **Office project type:** Mail Add-in (read & compose forms)
- **Technology to use:** Angular ADAL
- **Application ID as registered in Azure AD:** 03ad2348-c459-4573-8f7d-0ca44d822e7c
- **Supported Outlook forms:** E-Mail message - read form, Appointment - read form

### Results:

```
.
├── .bowerrc
├── bower.json
├── gulpfile.js
├── jsconfig.json
├── manifest.xml
├── package.json
├── tsd.json
└── src
    ├── public
    │   ├── index.html
    │   ├── appread
    │   │   ├── app.adalconfig.js
    │   │   ├── app.config.js
    │   │   ├── app.module.js
    │   │   ├── app.routes.js
    │   │   ├── index.html
    │   │   ├── home
    │   │   │   ├── home.controller.js
    │   │   │   └── home.html
    │   │   └── services
    │   │       └── data.service.js
    │   ├── content
    │   │   ├── Office.css
    │   │   └── site.css
    │   ├── images
    │   │   └── close.png
    │   └── scripts
    │       └── MicrosoftAjax.js
    └── server
        └── server.js
```