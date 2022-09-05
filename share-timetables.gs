function shareTimetables(folderId, contactEmail) {

  var files = DriveApp.getFolderById(folderId).getFiles();
  var errorString = '';
  while (files.hasNext()) {
    var file = files.next();
    var ugc = file.getName().slice(0,-4) + '@ugcloud.ca';
    try {
      file.addViewer(ugc);
    }
    catch(err){
      Logger.log(ugc);
      errorString = errorString + '\nError sharing file with ' + ugc;
    }
  }

  if (errorString !== '') {
    GmailApp.createDraft(contactEmail, 'Timetable Sharing Errors', errorString);
  }
  else {
    GmailApp.createDraft(contactEmail, 'Timetable Sharing Complete!', 'Done!');
  }

}
