const path = require('path');
const fs = require('fs');

export namespace helperMethods {

export function deleteFolderRecursively(projectFolder: string) 
{
    if(fs.existsSync(projectFolder))
    {
        fs.readdirSync(projectFolder).forEach(function(file,index){ 
        var curPath = projectFolder + "/" + file; 
        
        if(fs.lstatSync(curPath).isDirectory())
        {
            deleteFolderRecursively(curPath);
        }
        else
        {
            fs.unlinkSync(curPath);
        }
    }); 
    fs.rmdirSync(projectFolder); 
    }
};

export function doesProjectFolderExists (projectFolder: string)
{      
  if (fs.existsSync(projectFolder))
    {
      if (fs.readdirSync(projectFolder).length > 0)
      {          
        return true;
      }
    }
    return false;
};
}