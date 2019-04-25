import * as fs from 'fs';
import * as _ from 'lodash';

export default class projectsJsonData{
    m_projectJsonDataFile:string = '/projectProperties.json';
    m_projectJsonData;

    constructor(templatePath:string){
        let jsonData = fs.readFileSync(templatePath + this.m_projectJsonDataFile);
        this.m_projectJsonData = JSON.parse(jsonData.toString());
    }

    isValidInput(input: string, isHostParam: boolean)
    {
      if (isHostParam)
      {
        for (let key in this.m_projectJsonData.hostTypes)
        {
           if (_.toLower(input) == key){
             return true;
           }
          }
          return false;
        }
        else{
          for (let key in this.m_projectJsonData.projectTypes)
          {
            if (_.toLower(input) == key){
              return true;
            }
          }
          return false;
        }
    }

    getProjectDisplayName(projectType: string){
      return this.m_projectJsonData.projectTypes[_.toLower(projectType)].displayname;
    }
    
    getParsedProjectJsonData()
    {
      return this.m_projectJsonData;
    }
    
    getProjectTemplateNames()
    {
      let projectTemplates : string[] = [];
      for (let key in this.m_projectJsonData.projectTypes)
      {
        projectTemplates.push(key);
      }
      return projectTemplates;
    }
    
    projectBothScriptTypes (projectType: string)
    {
      return this.m_projectJsonData.projectTypes[_.toLower(projectType)].templates.javascript != undefined && this.m_projectJsonData.projectTypes[_.toLower(projectType)].templates.typescript != undefined;
    }

    getHostTemplateNames()
    {
      let hosts : string[] = [];
      for (let key in this.m_projectJsonData.hostTypes)
      {
        hosts.push(this.m_projectJsonData.hostTypes[key].displayname);
      }
      return hosts;
    }

    getHostDisplayName(hostKey: string)
    {
      for (let key in this.m_projectJsonData.hostTypes)
      {
        if (_.toLower(hostKey) == key){
          return this.m_projectJsonData.hostTypes[key].displayname;
        }
      }
      return undefined;
    }

    getProjectTemplateRepository(projectTypeKey: string, scriptType: string)
    {
      for (let key in this.m_projectJsonData.projectTypes)
      {
        if (_.toLower(projectTypeKey) == key){
          if (projectTypeKey == 'manifest'){
           return this.m_projectJsonData.projectTypes[key].templates.manifestonly.repository;
          }
          else{
            return this.m_projectJsonData.projectTypes[key].templates[scriptType].repository;
          }          
        }
      }
      return undefined;
    }

  getProjectTemplateBranchName(projectTypeKey: string, scriptType: string, prerelease: boolean)
    {
      for (let key in this.m_projectJsonData.projectTypes)
      {
        if (_.toLower(projectTypeKey) == key){
          if (projectTypeKey == 'manifest')
          {
            return this.m_projectJsonData.projectTypes.manifest.templates.branch;
          }
          else{
            if (prerelease) {
              return this.m_projectJsonData.projectTypes[key].templates[scriptType].prerelease
            } else {
              return this.m_projectJsonData.projectTypes[key].templates[scriptType].branch;
            }
          }          
        }
      }
      return undefined;
    }

    getProjectRepoAndBranch(projectTypeKey: string, scriptType: string, prerelease: boolean)
    {
      scriptType =  scriptType === 'ts' ? 'typescript' : 'javascript';
      let repoBranchInfo = { repo: <string> null, branch: <string> null };

      repoBranchInfo.repo = this.getProjectTemplateRepository(projectTypeKey, scriptType);
      repoBranchInfo.branch = (repoBranchInfo.repo) ? this.getProjectTemplateBranchName(projectTypeKey, scriptType, prerelease) : undefined;
      
      return repoBranchInfo;
    }
  }