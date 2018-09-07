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
           return this.m_projectJsonData.manifest.templates.manifestonly.repository;
          }
          else{
            return this.m_projectJsonData.projectTypes[key].templates[scriptType].repository;
          }          
        }
      }
      return undefined;
    }

    getProjectTemplateBranchName(projectTypeKey: string, scriptType: string, branchIndex: number)
    {
      // Check to see if repository is defined. If not, then a branch won't be defined, so just return.
      if (this.getProjectTemplateRepository(projectTypeKey, scriptType) == ""){
        return undefined;
      }

      for (let key in this.m_projectJsonData.projectTypes)
      {
        if (_.toLower(projectTypeKey) == key){
          if (projectTypeKey == 'manifest')
          {
            if (this.m_projectJsonData.manifest.templates.manifestonly.branches == undefined){
              return undefined;
            }
            else{
              return this.m_projectJsonData.manifest.templates.manifestonly.branches[branchIndex].name;
            }
          }
          else{
            if (this.m_projectJsonData.projectTypes[key].templates[scriptType].branches == undefined){
              return undefined;
            }
            else{
              return this.m_projectJsonData.projectTypes[key].templates[scriptType].branches[branchIndex].name;
            }
          }          
        }
      }
      return undefined;
    }
  }