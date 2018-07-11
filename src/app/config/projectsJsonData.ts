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

    getProjectDisplayNames(projectType: string){
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
      return this.m_projectJsonData.projectTypes[_.toLower(projectType)].javascript && this.m_projectJsonData.projectTypes[_.toLower(projectType)].typescript;
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

    normalizeHostNameFromInput(input: string)
    {
      for (let key in this.m_projectJsonData.hostTypes)
      {
        if (_.toLower(input) == key){
          return this.m_projectJsonData.hostTypes[key].displayname;
        }
      }
      return undefined;
    }
  }