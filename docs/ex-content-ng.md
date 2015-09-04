# Example - Office Content Add-in in Angular

This document demonstrates creating a Content Add-in first in an empty project as well as in an existing project using [Angular](https://www.angularjs.org) as the technology.

## Empty Project

This example creates an Office Content Add-in as Angular within an empty project folder.

```bash
$ yo office --skip-install
```

### Prompt Responses

- **Project name (the display name):** My Office Add-in
- **Root folder of the project:** {blank} 
- **Office project type:** Content Add-in
- **Technology to use:** Angular
- **Supported Office Applications:** Word, Excel, PowerPoint, Project

```
.
├── .bowerrc
├── bower.json
├── gulpfile.js
├── index.html
├── jsconfig.json
├── manifest.xml
├── tsd.json
├── app
│   ├── app.module.js
│   ├── app.routes.js
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

The generator [nodehttps](https://www.npmjs.com/package/generator-nodehttps) is first used to create a folder for a self-hosted HTTPS site on the local development system.

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
- **Office project type:** Content Add-in
- **Technology to use:** Angular
- **Supported Office Applications:** Word, Excel, PowerPoint, Project

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
    │   ├── app
    │   │   ├── app.module.js
    │   │   ├── app.routes.js
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