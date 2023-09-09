function onOpen() {
  let menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Get Skill Metrix', 'processPDFsInFolder');
  menu.addToUi();
}
