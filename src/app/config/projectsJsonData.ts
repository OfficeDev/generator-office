import fs from 'fs';
import _ from 'lodash';

export default class projectsJsonData {
  m_projectJsonDataFile = '/projectProperties.json';
  m_projectJsonData;

  constructor(templatePath: string) {
    const jsonData = fs.readFileSync(templatePath + this.m_projectJsonDataFile);
    this.m_projectJsonData = JSON.parse(jsonData.toString());
  }

  isValidProjectType(input: string): boolean {
    for (const key in this.m_projectJsonData.projectTypes) {
      if (_.toLower(input) == key) {
        return true;
      }
    }
    return false;
  }

  isValidHost(input: string): boolean {
    for (const key in this.m_projectJsonData.hostTypes) {
      if (_.toLower(input) == key) {
        return true;
      }
    }
    return false;
  }

  isValidManifestType(input: string): boolean {
    for (const key in this.m_projectJsonData.manifestTypes) {
      if (_.toLower(input) == key) {
        return true;
      }
    }
    return false;
  }

  getProjectDisplayName(projectType: string): string {
    return this.m_projectJsonData.projectTypes[_.toLower(projectType)].displayname;
  }

  getParsedProjectJsonData() {
    return this.m_projectJsonData;
  }

  getProjectTemplateNames(): string[] {
    const projectTemplates: string[] = [];
    for (const key in this.m_projectJsonData.projectTypes) {
      projectTemplates.push(key);
    }
    return projectTemplates;
  }

  projectBothScriptTypes(projectType: string): boolean {
    return this.m_projectJsonData.projectTypes[_.toLower(projectType)].templates.javascript != undefined && this.m_projectJsonData.projectTypes[_.toLower(projectType)].templates.typescript != undefined;
  }

  getManifestOptions(projectType: string, host: string): string[] {
    const selectedHost = this.m_projectJsonData.projectTypes[projectType]?.supportedHosts[host];
    return selectedHost ? selectedHost.supportedManifestTypes : [];
  }

  getHostOptions(projectType: string): string[] {
    const selectedProjectType = this.m_projectJsonData.projectTypes[projectType];
    return selectedProjectType ? Object.keys(selectedProjectType.supportedHosts) : [];
  }

  getScriptTypeOptions(projectType: string): string[] {
    const scriptTypes: string[] = [];
    for (const template in this.m_projectJsonData.projectTypes[projectType].templates) {
      let scriptType: string = "unknown";
      if (template === "javascript") {
        scriptType = "JavaScript";
      } else if (template === "typescript") {
        scriptType = "TypeScript";
      }

      scriptTypes.push(scriptType);
    }
    return scriptTypes;
  }

  getHostDisplayName(hostKey: string): string {
    for (const key in this.m_projectJsonData.hostTypes) {
      if (_.toLower(hostKey) == key) {
        return this.m_projectJsonData.hostTypes[key].displayname;
      }
    }
    return "";
  }

  getManifestDisplayName(hostKey: string): string {
    return this.m_projectJsonData.manifestTypes[hostKey]?.displayname;
  }

  getProjectTemplateRepository(projectTypeKey: string, scriptType: string): string | undefined {
    for (const key in this.m_projectJsonData.projectTypes) {
      if (_.toLower(projectTypeKey) == key) {
        if (projectTypeKey == 'manifest') {
          return this.m_projectJsonData.projectTypes[key].templates.manifestonly.repository;
        }
        else {
          return this.m_projectJsonData.projectTypes[key].templates[scriptType].repository;
        }
      }
    }
    return undefined;
  }

  getProjectTemplateBranchName(projectTypeKey: string, scriptType: string, prerelease: boolean): string | undefined {
    for (const key in this.m_projectJsonData.projectTypes) {
      if (_.toLower(projectTypeKey) == key) {
        if (projectTypeKey == 'manifest') {
          return this.m_projectJsonData.projectTypes.manifest.templates.branch;
        }
        else {
          if (prerelease) {
            if (this.m_projectJsonData.projectTypes[key].templates[scriptType].prerelease) {
              return this.m_projectJsonData.projectTypes[key].templates[scriptType].prerelease
            }
            else {
              return "master";
            }
          } else {
            return this.m_projectJsonData.projectTypes[key].templates[scriptType].branch;
          }
        }
      }
    }
    return undefined;
  }

  getProjectRepoAndBranch(projectTypeKey: string, scriptType: string, prerelease: boolean) {
    scriptType = scriptType === 'ts' ? 'typescript' : 'javascript';
    const repoBranchInfo = { repo: <string|undefined>undefined, branch: <string|undefined>undefined };

    repoBranchInfo.repo = this.getProjectTemplateRepository(projectTypeKey, scriptType);
    repoBranchInfo.branch = (repoBranchInfo.repo) ? this.getProjectTemplateBranchName(projectTypeKey, scriptType, prerelease) : undefined;

    return repoBranchInfo;
  }
}