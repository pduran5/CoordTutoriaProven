var docTemplate = "1D_eeJ6zyjdQkep7xrzB5hYjNAx6D17UMETiw3Derle8";
var docName = "Seguiment PAT 1r Trimestre";

function onFormSubmit(e) {
  var responses = FormApp.getActiveForm().getResponses();
  var length = responses.length;
  var lastResponse = responses[length-1];
  var nomtutor="", emailtutor="", cicle="", curs="", grup="", comunestxt="", especifiquescfgmtxt="", especifiquescfgstxt="", observacions="";
  
  var formValues = lastResponse.getItemResponses();
  
  for(var i=0; i<formValues.length; i++) {
    switch(formValues[i].getItem().getTitle()) {
      case "Nom i cognoms del tutor/a": nomtutor = formValues[i].getResponse(); break;
      case "Email del tutor/a": emailtutor = formValues[i].getResponse(); break;
      case "Nom del cicle": cicle = formValues[i].getResponse(); break;
      case "Curs": curs = formValues[i].getResponse(); break;
      case "Grup": grup = formValues[i].getResponse(); break;
      case "Activitats comunes CFGM i CFGS":
            var comunes = formValues[i].getResponse();
            for(var j in comunes) {
                comunestxt += comunes[j] + "\n";
            }
            break;
      case "Activitats específiques CFGM":
            var especifiquescfgm = formValues[i].getResponse();
            for(var j in especifiquescfgm) {
                especifiquescfgmtxt += especifiquescfgm[j] + "\n";
            }
            break;
      case "Activitats específiques CFGS":
            var especifiquescfgs = formValues[i].getResponse();
            for(var j in especifiquescfgs) {
                especifiquescfgstxt += especifiquescfgs[j] + "\n";
            }
            break;
      case "Observacions": observacions = formValues[i].getResponse(); break;
    }         
  }

  
  // Data
  var date = new Date();
  var d = date.getDate();
  var m = date.getMonth() + 1; //Month from 0 to 11
  var y = date.getFullYear();
  var now = '' +  (d <= 9 ? '0' + d : d) + "/" + (m<=9 ? '0' + m : m) + '/' + y;
    
  // Get document template, copy it as a new temp doc, and save the Doc’s id
  var copyId = DriveApp.getFileById(docTemplate).makeCopy(docName+' '+nomtutor).getId();
  
  // Open the temporary document
  var copyDoc = DocumentApp.openById(copyId);
  
  // Get the document’s body section
  var copyBody = copyDoc.getActiveSection();
  
  // Replace place holder keys,in our google doc template
  copyBody.replaceText('#CICLE#', cicle);
  copyBody.replaceText('#CURS#', curs);
  copyBody.replaceText('#GRUP#', grup);
  copyBody.replaceText('#COMUNES#', comunestxt);
  copyBody.replaceText('#ESPECIFIQUESCFGM#', especifiquescfgmtxt);
  copyBody.replaceText('#ESPECIFIQUESCFGS#', especifiquescfgstxt);
  copyBody.replaceText('#OBSERVACIONS#', observacions);
  copyBody.replaceText('#NOMTUTOR#', nomtutor);
  copyBody.replaceText('#DATA#', now);
  
  // Save and close the temporary document
  copyDoc.saveAndClose();
  
  // Convert temporary document to PDF
  var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + copyId + "&format=docx&access_token=" + ScriptApp.getOAuthToken();
  var docx = UrlFetchApp.fetch(url).getBlob();
  docx.setName(docName + ' ' + nomtutor + '.docx');
  
  // Attach PDF and send the email
  var subject = docName + ' ' + nomtutor;
  var body = "A continuació s'adjunta el document " + docName + "";
  MailApp.sendEmail(emailtutor, subject, body, {htmlBody: body, attachments: [pdf, docx]});
  
  // Delete temp file
  // DriveApp.getFileById(copyId).setTrashed(true);
}