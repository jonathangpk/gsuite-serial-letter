/*
  @return {users, folderId, templateId}
*/
function getUsersFolderAndTemplate(docUrl) {
  var template = DocumentApp.openByUrl(docUrl);
  var users = getUsers(SpreadsheetApp.getActiveSheet());
  // Folder for PDFs gets created in Template Folder
  var folder = DriveApp.getFileById(template.getId()).getParents().next().createFolder("PDF: " + template.getName());
  var ret = {
    users: users,
    templateId: template.getId(),
    folderId: folder.getId()
  };
  return JSON.stringify(ret);
}
// Called bei GUI

/*
@param: Sheet (the active sheet with user data)
@return: Array of User Objects
*/
function getUsers(ss) {
  var users = [];
  var data = ss.getDataRange().getValues();
  var keys = data[0];
  for (var i = 1; i<data.length; i++) {
    var user = {}
    for (var j = 0; j<keys.length;j++) {
      if (!keys[j] || !data[i]) // Probably wont happen
        alert("Fehler bei: ("+i+","+j+")");
      else
        user[keys[j]] = data[i][j];
    }
    users.push(user);
  }
  return users;
}

/*
@param: Targetfolder for pdfs
@param: Gdoc Template as Document
@param: user Object
*/
function createUserPdf(folderId, templateId, user) {
  var folder = DriveApp.getFolderById(folderId);
  var sp = '{{[^{]+}}'; // search pattern for template strings
  var docfile = DriveApp.getFileById(templateId).makeCopy();
  docfile.setName(docfile.getName().replace('Kopie von ', '') + "_" + user.id + "_" + user.name);
  var template = DocumentApp.openById(docfile.getId());
  
  //Replace all template Strings with user values
  var range = template.getBody().findText(sp);
  while(range) {
    var ts = range.getElement().asText().getText();
    ts = ts.substring(ts.indexOf('{'), ts.indexOf('}') + 2);
    var key = getKeyFromTemplateString(ts);
    var val = getTextFromKey(key, user);
    if (val) {
      template.getBody().replaceText(ts, val);
      range = template.getBody().findText(sp);
    } else {
      range = template.getBody().findText(sp, range);
    }
  }
  template.saveAndClose();
  var pdf = DriveApp.createFile(docfile.getAs('application/pdf'));
  moveFileToDir(folder, pdf);
  Drive.Files.remove(docfile.getId());
}

function moveFileToDir(folder, file) {
  folder.addFile(file);
  var iter = file.getParents();
  var next;
  while (iter.hasNext()) {
    next = iter.next();
  }
  next.removeFile(file);
}
// Get The user value for the template key
function getTextFromKey(key, user) {
  switch (key) {
    case "formalsalutation": 
      return getFormalSalutation(user);
    case "fullNameWithTitle":
      return getFullNameWithTitle(user);
    default: 
      return user[key];
  }
}

function getFullNameWithTitle (m) {
  return (m.title && m.title.trim().length !== 0 ? m.title.trim() + ' ' : '') + m.firstname.trim() + ' ' + m.name.trim();
}

function getFormalSalutation (m) {
    if (m.salutation === 'Herr') {
        return 'Sehr geehrter Herr ' + (m.title && m.title.trim().length !== 0 ? m.title.trim() + ' ' : '') + m.name.trim();
    } else if (m.salutation === 'Frau') {
        return 'Sehr geehrte Frau ' + (m.title && m.title.trim().length !== 0 ? m.title.trim() + ' ' : '') + m.name.trim();
    } else {
      // TODO: Gender neutral salutation
        return 'Sehr geehrte(r) Herr/Frau ' + (m.title && m.title.trim().length !== 0 ? m.title.trim() + ' ' : '') + m.name.trim();;
    }
}

function getKeyFromTemplateString (s) {
  return s.replace('{{', '').replace('}}', '').trim();
}

function alert(arg) {
  SpreadsheetApp.getUi().alert(arg);
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('template')
      .setTitle('Personalisierte Briefe');
  SpreadsheetApp.getUi().showSidebar(ui);
}
