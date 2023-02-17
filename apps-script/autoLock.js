/*
* Run on time trigger to check email notifying overdue and move the project folder into locked folder. 
* - run getJDriveIdObj function to get drive ID of every year and every country (only 2023+) for drive J and drive Lock
* - run getProjectArr function to extract project number from body of email within parenthesis.
* - run loopProjectFolders function to loop drive J within the same year of the project (from project number prefix). 
*   Find the matching folder and move to locked drive. (Matching year and country.)
*
*/
function lockFolderOnTrigger() {
  try {
    let j = getJDriveIdObj("J Folder")
    let lock = getJDriveIdObj("Lock Folder")

    let foundProject = false
    let projectArr = getProjectArr();
    projectArr.forEach(project => {
      if(parseInt(project.year) > 2022){
        console.log('>22')
        let {year, projectno} = project
        let jIdKeys = Object.keys(j[year])
        //loop countries folder in the year 
        for(let i = 0; i < jIdKeys.length ; i++){
          let countryKey = jIdKeys[i]
          console.log('countryKey: ' + countryKey)

          foundProject = loopProjectFolders(j[year][countryKey], lock[year][countryKey], projectno)
          console.log('moved: ' + foundProject)

          if(foundProject) break; //break for loop (looping country folder) if project is already moved
        }

      }else{
        console.log('<=22')
        let foundProject = loopProjectFolders(j[project.year], lock['Lock 2'], project.projectno)
        console.log('moved: ' + foundProject)
      }
    })

  } catch(e) {console.log('moveFolder error : ' + e.message);}
}

/*
* get object from 'Rule' spreadsheet.
* @param {string} sheetname - sheet name of wanted drive containing drive ID
*
* @return {Object} obj - the key is in A column and value is ID in B column. column A format: driveKey (year/drive name)
* For year 2023+ the key is year and the value is object with countries as properties. column A format: year-country
*/
function getJDriveIdObj(sheetname){
  let jFolderIdData = SpreadsheetApp.openById(ruleSheetId).getSheetByName(sheetname).getDataRange().getValues()
  jFolderIdData.shift()
  let obj = {}
  jFolderIdData.forEach(row=>{
    let [driveKey, id] = row
    //if '-' is in column A, project year is 2023+ and folders are also separated by countries.
    if( driveKey.includes('-') ){
      let [year, country] = driveKey.split('-')
      if(!obj[year]) obj[year] = {}
      obj[year][country] = id
    }else{
      obj[driveKey] = id
    }
  })
  return obj
}

/*
* Get project data that needs to be locked from email body. 
* @return {Object[]} projArr 
*         {string} projArr.projectno
*         {string} projArr.year
*/
function getProjectArr(){
  try {
    var projectno = '',year = '', projArr = []; 
    var threads = GmailApp.search('from:dwpservice@dwp.com AND subject:AR Notification (45 Days Overdue) - - AND is:unread')[0];
    var body = threads.getMessages()[0].getBody();
    var messages = threads.getMessages();
    //Read all messages in the thread.
    for(var i in messages) {
      var thread = messages[i];
      GmailApp.markMessageRead(thread)
      var body = thread.getBody();
      //Extract project number in ( ) in the email body
      projectno = body.match(/\((.*)\)/).pop();

      let projNoExtra = projectno.search(/\d+-\d+-\d/); //if it is branched project such as 22-0000-1, get the number of main project 22-0000
      if(projNoExtra !== -1){projectno = projectno.match(/\d+-\d+/)[0];}

      var splitprojectno = projectno.split('-');
      
      if(projectno[0].toLowerCase() == projectno[0].toUpperCase()) {
        year = '20'+splitprojectno[0]; //normal project number format
      } else {
        year = '20'+splitprojectno[1]; //some old project has alphabets prefix before year prefix
      }
      projArr.push({projectno: projectno.trim() + " " , year: year.toString()}); 
    }
    console.log(projArr)
    return projArr
  } catch(e) {Logger.log('getProjectArr error : ' + e.message);}
}

/*
* Loop through folders within jFolderId and move the folder that matched projectno to lockFolderId. 
* @return {boolean} foundProject - return true if found and moved matching project.
*/
function loopProjectFolders(jFolderId, lockFolderId, projectno){
  let foundProject = false
  var folders_j = DriveApp.getFolderById(jFolderId).getFolders();
  var folders_Lock = DriveApp.getFolderById(lockFolderId);
  while(folders_j.hasNext()) {
    var folder = folders_j.next();
    var name = folder.getName();
    if(name.includes(projectno)) {
      folder.moveTo(folders_Lock);
      console.log(`folder ${name} - J drive = ${jFolderId} moveTo folders_Lock = ${lockFolderId}`)
      foundProject = true
      break; //break while loop (looping project folder)
    }
  }
  return foundProject
}