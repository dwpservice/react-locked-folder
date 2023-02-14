let ruleSheetId = '1xuMyOiwY4fgXHDSjqo4YX8XdTg094yHxbQ3m3o7ySbs'

/*
* Return html page
*
*/
function doGet(){
    let tmp = HtmlService.createTemplateFromFile("index")
    return tmp.evaluate().setTitle("Locked Drive").addMetaTag("viewport", "width-device-width, initial-scale=1.0")
}

/*
* Get all project folders in all locked drives and return list of data.
* @returns {Object[]} projNo (string), projName (string), id (string)
*
*/
function getLockedDriveProjects(){
  const ss = SpreadsheetApp.openById(ruleSheetId).getSheetByName("Lock Folder")
  let data = ss.getDataRange().getValues()
  let firstRow = data.shift()
  //get permittedEmails by firstRow[5] = cell F1
  let hasPermission = getPermission(firstRow[5])
  let lockedDriveArr = data.map(row=> row[1])
  let projData = []
  lockedDriveArr.forEach(lockedDriveIdEach => {
    let drive = DriveApp.getFolderById(lockedDriveIdEach)
    let folders = drive.getFolders()
    while(folders.hasNext()){
      let folder = folders.next()
      let [projNo, ...projName] = folder.getName().split(" ")
      let folderId = folder.getId()
      projData.push({projNo, projName: projName.join(" ").trim(), id: folderId})
    }
  })
  projData = projData.sort((a, b) => b.projNo.localeCompare(a.projNo))
  
  
  return { hasPermission, projData }
}

/*
* Get permission by sessionEmail return true if the email is in "Rule" sheet "Locked Drive" tab "F1" cell. 
* If permission is true then show unlock button. If not only shows the project list.
* @param {string} permittedEmails - list of permitted emails separated by comma
* @return {boolean} hasPermission
*/
function getPermission(permittedEmails) {
  let sessionEmail = Session.getActiveUser().getEmail()
  let hasPermission = false

  if(permittedEmails.includes(sessionEmail.toString().toLowerCase().trim())) hasPermission = true

  return hasPermission
}


/*
* Move folder to J drive according to year and region (if > 2023)
* @param {string} folderID - folder id of folder in locked drive to move to J drive
* @return {string} folderUrl
*/
function moveLockedToDriveJ(folderID){
  let originalJFolder
  let folder = DriveApp.getFolderById(folderID)
  let folderName = folder.getName()
  let projNoYear = folderName.split("-")[0]

  const jFolderIdData = SpreadsheetApp.openById(ruleSheetId).getSheetByName("J Folder").getDataRange().getValues()

  //project year 2023 onwards will move to separated folder for each region. Before 2023 only separated by year
  if(parseInt(projNoYear) > 22){
    let parents = folder.getParents()
    if(parents.hasNext()){
      let parent = parents.next()
      let parentName = parent.getName()
      console.log(parentName)
      let originalJFolderId = jFolderIdData.find(row=> row[0].toString().trim() === `20${projNoYear}-${parentName}`)
      if(originalJFolderId) originalJFolder = DriveApp.getFolderById(originalJFolderId[1])
    }
  }else{
    let originalJFolderId = jFolderIdData.find(row=> row[0].toString().trim() === ((projNoYear === "18" || projNoYear === "20" ) ? `20${projNoYear}-1` : `20${projNoYear}`))
    if(originalJFolderId) originalJFolder = DriveApp.getFolderById(originalJFolderId[1])
  }

  if (originalJFolder) {
    folder.moveTo(originalJFolder)
    let folderUrl = folder.getUrl()
    let sessionEmail = Session.getActiveUser().getEmail()
    MailApp.sendEmail(
      `benjaporn.k@dwp.com`,
      `Locked folder has been unlocked.`,
      `Unlocked ${folderName}. Url: ${folderUrl} by ${sessionEmail}`,
      {
        name: `Locked folders webapp`,
        htmlBody: `Unlocked <a href="${folderUrl}" target="_blank">${folderName}</a>. by ${sessionEmail}`
      }
    )
    return folderUrl
  }else{
    throw new Error(`Drive J not found. Please contact IT.`)
  }
}