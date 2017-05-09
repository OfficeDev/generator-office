import * as Generator from 'yeoman-generator';
import * as chalk from 'chalk';
import * as yosay from 'yosay';
import * as _ from 'lodash';
import 'isomorphic-fetch';
import * as uuid from 'uuid/v4';
import * as jsYaml from 'js-yaml';

/**
 * Script Lab yeoman companion
 * Generate add-in projects from the shared public gists from Script-Lab
 */
module.exports = class extends Generator {
  snippet = null;
  answers: any;

  constructor(args, options) {
    super(args, options);

    this.argument('id', {
      type: String,
      required: false,
      description: 'ID of public gist shared using Script Lab.',
      default: null
    });

    this.option('skip-install', {
      type: Boolean,
      default: false,
      description: 'Skip running `npm install` post scaffolding.'
    });
  }

  /**
   * Generator initalization
   */
  initializing() {
    let message = `Welcome to the ${chalk.bold.green('Office Add-in')} generator, by ${chalk.bold.green('@OfficeDev')}! Let\'s create a project together!`;
    this.log(yosay(message));
  }

  prompting() {
    const done = (this as any).async();
    (async () => {
      this.answers = await this.prompt({
        type: 'input',
        name: 'question-id',
        message: 'What is the id of the public gist?',
        validate: (input: string) => _.isString(input),
        when: !_.isString(this.options['id'])
      });

      done();
    })();
  }

  configuring() {
    const done = (this as any).async();
    (async () => {
      const snippet = await this._downloadGist(this.options['id'] || this.answers['question-id']);
      const processedData = this._processLibraries(snippet);
      this.snippet = { ...snippet, ...processedData };
      this.snippet.id = uuid();
      this.snippet.safeName = _.kebabCase(snippet.name);
      this.snippet.template = snippet.template.content;
      this.snippet.script = snippet.script.content;
      this.snippet.style = snippet.style.content;
      this.snippet.host = snippet.host && snippet.host.toLowerCase();
      done();
    })();
  }

  writing() {
    this.destinationRoot(this.snippet.safeName);
    if (this.snippet.host && this.snippet.host !== 'web') {
      this.fs.copyTpl(this.templatePath(`manifest/${this.snippet.host}.xml`), this.destinationPath(`manifest.xml`), this.snippet);
    }
    this.fs.copy(this.templatePath(`common/**`), this.destinationPath());
    this.fs.copyTpl(this.templatePath(`default/**`), this.destinationPath(), this.snippet);
  }

  install() {
    if (this.options['skip-install']) {
      this.installDependencies({
        npm: false,
        bower: false
      });
    }
    else {
      this.installDependencies({
        npm: true,
        bower: false,
      });
    }
  }

  private async _downloadGist(id: string): Promise<any> {
    if (!_.isString(id)) {
      throw new Error('Failed to load gist: No Id was received.');
    }
    else {
      const res = await fetch(`https://api.github.com/gists/${id}`);
      if (res.status === 200) {
        try {
          const gist = await res.json();
          let snippet = _.find<any>(gist.files, (value, key: string) => value ? /\.ya?ml$/gi.test(key) : false);
          return jsYaml.load(snippet.content);
        }
        catch (e) {
          console.log(e);
          throw new Error('Invalid gist. Make sure the gist was created using Script Lab');
        }
      }
      else {
        if (res.status === 401) {
          throw new Error('Failed to load: Make sure the gist is public.');
        }
        else if (res.status === 404) {
          throw new Error('Failed to load: gist could not be found.');
        }
        else {
          throw new Error('Failed to load: Unexpected error while downloading gist.');
        }
      }
    }
  }

  private _processLibraries(snippet: any) {
    let linkReferences: string[] = [];
    let scriptReferences: string[] = [];
    let types: string[] = [];
    let officeJS: string = null;

    snippet.libraries.split('\n').forEach(processLibrary);

    return { linkReferences, scriptReferences, officeJS, types };

    function processLibrary(text: string) {
      if (text == null || text.trim() === '') {
        return null;
      }

      text = text.trim();

      let isNotScriptOrStyle = /^#.*|^\/\/.*|^\/\*.*|.*\*\/$.*|^dt~|\.d\.ts$/im.test(text);
      if (isNotScriptOrStyle) {
        return null;
      }

      let isTypes = /^@types/i.test(text);
      if (isTypes) {
        const matches = text.match(/(^@types\/[\/\w-]*?)$/im);
        if (matches && matches[1]) {
          return types.push(matches[1]);
        }
      }

      let resolvedUrlPath = (/^https?:\/\/|^ftp? :\/\//i.test(text)) ? text : `https://unpkg.com/${text}`;

      if (/\.css$/i.test(resolvedUrlPath)) {
        return linkReferences.push(resolvedUrlPath);
      }

      if (/\.ts$|\.js$/i.test(resolvedUrlPath)) {
        /*
        * Don't add Office.js to the rest of the script references --
        * it is special because of how it needs to be *outside* of the iframe,
        * whereas the rest of the script references need to be inside the iframe.
        */
        if (/(?:office|office.debug).js$/.test(resolvedUrlPath.toLowerCase())) {
          officeJS = resolvedUrlPath;
          return null;
        }

        return scriptReferences.push(resolvedUrlPath);
      }

      return scriptReferences.push(resolvedUrlPath);
    }
  }
};
