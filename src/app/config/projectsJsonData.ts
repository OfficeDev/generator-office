import * as fs from 'fs';
import * as _ from 'lodash';

export default class projectsJsonData{
    m_projectJsonDataFile:string = '/projectProperties.json';
    m_projectJsonData;

    constructor(templatePath:string){
        let jsonData = fs.readFileSync(templatePath + this.m_projectJsonDataFile);
        this.m_projectJsonData = JSON.parse(jsonData.toString());
    }
    
    isValidInput(input, isHostParam)
    {
      if (isHostParam)
      {
        for (let key in this.m_projectJsonData.hostTypes)
        {
          let hostType = this.m_projectJsonData.hostTypes[key].name;
           if (_.toLower(input) == _.toLower(hostType))
           {
             input = hostType;
             return true;
           }
          }
          return false;
        }
        else{
          for (let key in this.m_projectJsonData.projectTypes)
          {
            let projectType = this.m_projectJsonData.projectTypes[key].name;
            if (_.toLower(input) == _.toLower(projectType))
            {
              input= projectType;
              return true;
            }
          }
          return false;
        }
    }

    getProjectDisplayNames(projectTemplate){
      return this.m_projectJsonData.projectTypes[_.toLower(projectTemplate)].displayname;
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
        projectTemplates.push(this.m_projectJsonData.projectTypes[key].name);
      }
      return projectTemplates;
    }
    
    projectBothScriptTypes (projectTemplate)
    {
      return this.m_projectJsonData.projectTypes[_.toLower(projectTemplate)].javascript && this.m_projectJsonData.projectTypes[_.toLower(projectTemplate)].typescript;
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
  }