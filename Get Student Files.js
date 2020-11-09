function getAll () {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var s=ss.getActiveSheet();
    var c=s.getActiveCell();
    var folder=DriveApp.getFolderById("zAlUJg_12zXxPwdGwqaaPljhdO2cxw");
    var files=folder.getFiles();
    var names=[],f,str;
    while (files.hasNext()) {
      f=files.next();
      names.push([f.getUrl(), f.getName()]);
    }
    s.getRange(c.getRow(),c.getColumn(),names.length, 2).setValues(names); 
  }