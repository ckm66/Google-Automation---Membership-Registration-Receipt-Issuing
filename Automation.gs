// CREATE A EDIT TRIGGER 
function CreateonEditTrigger() {
  ScriptApp.newTrigger('printReceipt')
      .forSpreadsheet('1MlS3qv4J4q9nNtgbik14BqrrXlTMmM2P3lUykqxk2zo')
      .onEdit()
      .create();
}

function printReceipt (name, uid, referenceNumber) {
  const receiptTemplate = DriveApp.getFileById("1zanPZ_J_Nv7xmNvgGPChX5KqeZeTZTrSDId2iW3jayg");
  const receiptFolder = DriveApp.getFolderById("14U70t-8pXxsgnzpBNcCsKj1Gc6T39cp8");
  const memberReceipt = receiptTemplate.makeCopy(receiptFolder);
  const memberReceiptDoc = DocumentApp.openById(memberReceipt.getId());
  const body = memberReceiptDoc.getBody();
  body.replaceText("{{Name}}", name);
  body.replaceText("{{UID}}", uid);
  body.replaceText("{{Reference Number}}", referenceNumber);
  memberReceiptDoc.saveAndClose();
  const receipt = memberReceipt.getAs(MimeType.PDF);
  const member_receipt = receiptFolder.createFile(receipt).setName(name + " (" + uid +  ") Receipt");
  receiptFolder.removeFile(memberReceipt);
  checkbox.getRange(rowIndex, columnIndex + 1).setValue("COMPLETED");
  return member_receipt;
}

function sendMembershipEmail (emailAddress) {
  var emailTemp = HtmlService.createTemplateFromFile("email");
  var htmlMessage = emailTemp.evaluate().getContent();
  checkbox.getRange(rowIndex, columnIndex + 2).setValue("SENDING");
  if (checkbox.getRange(rowIndex, columnIndex, 1, 1).getValues()[0][1] != "SENT")
  {
  MailApp.sendEmail(emailAddress, 'LifePlanet (Session 2022-2023) Membership Confirmation & Receipt', "", {
    name: "LifePlanet, HKU",
    replyTo: "lifeplanet.hkusu@gmail.com",
    htmlBody: htmlMessage,
    attachments:  [member_receipt],
  });
  checkbox.getRange(rowIndex, columnIndex + 2).setValue("SENT");
  }
}

function receiptIssusing (event) {
  const rangeModified = event.range;
  const columnIndex = rangeModified.getColumn();
  const rowIndex = rangeModified.getRow();
  const checkbox = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ALERT USER IF LESS THAN 7 QUOTA LEFT
  if (MailApp.getRemainingDailyQuota() < 7) {
    SpreadsheetApp.getUi().alert("FATAL: Reaching Mail Quota");
  }

  // Confirmation of Valid Input
  if (columnIndex !== 16) return;
  if (rowIndex < 2) return;
  if (checkbox.getRange(rowIndex, columnIndex, 1, 1).getDisplayValue() !== "TRUE") return;

  checkbox.getRange(rowIndex, columnIndex + 1).setValue("IN PROGRESS");
  checkbox.getRange(rowIndex, columnIndex + 2).setValue("PENDING");

  const row = checkbox.getRange(rowIndex, 1, 1, 18).getValues();
  const name = row[0][1];
  const uid = row[0][4];
  const referenceNumber = rowIndex;
  const emailAddress = row[0][4].toString();

  // PRINT RECEIPT
  const member_receipt = printReceipt(name, uid, referenceNumber);

  // SEND EMAIL
  sendMembershipEmail(emailAddress);

  // ALERT USER IF LESS THAN 10 QUOTA LEFT
  if (MailApp.getRemainingDailyQuota() < 10)
  {
    SpreadsheetApp.getUi().alert("WARMING: Reaching Mail Quota");
  }
}