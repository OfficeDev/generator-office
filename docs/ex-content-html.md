# Example - Office Content Add-in in HTML

This document demonstrates creating a Content Add-in first in an empty project as well as in an existing project using HTML, CSS & JavaScript as the technology.

## Empty Project

This example creates an Office Content Add-in as HTML within an empty project folder.

```bash
$ yo office --skip-install
```

### Prompt Responses

- **What is the name of the Add-in (the display name)?** My Office Add-in
- **What is the root folder where this Add-in should be craeted?** {blank} 
- **What type of Office Add-in do you want to create?** Content Add-in
- **What technology do you want to buld your Add-in with?** Html, CSS & JavaScript

```
.
├── .bowerrc
├── bower.json
├── app
│   ├── app.css
│   ├── app.js
│   └── home
│       ├── home.css
│       ├── home.html
│       └── home.js
├── content
│   └── Office.css
├── images
│   └── close.png
├── manifest.xml
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

- **What is the name of the Add-in (the display name)?** My Office Add-in
- **What is the root folder where this Add-in should be craeted?** src/public 
- **What type of Office Add-in do you want to create?** Content Add-in
- **What technology do you want to buld your Add-in with?** Html, CSS & JavaScript

### Results:

```
.
├── .bowerrc
├── bower.json
├── manifest.xml
├── package.json
└── src
    ├── public
    │   ├── app
    │   │   ├── app.css
    │   │   ├── app.js
    │   │   └── home
    │   │       ├── home.css
    │   │       ├── home.html
    │   │       └── home.js
    │   ├── content
    │   │   ├── Office.css
    │   │   └── site.css
    │   ├── images
    │   │   └── close.png
    │   ├── index.html
    │   └── scripts
    │       └── MicrosoftAjax.js
    └── server
        └── server.js
```