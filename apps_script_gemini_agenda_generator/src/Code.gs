function onOpen(){
 SpreadsheetApp.getUi()
 .createMenu("AI Tools")
 .addItem("AI Agenda Generator","showSidebar")
 .addToUi();
}

function showSidebar(){
 SpreadsheetApp.getUi().showSidebar(
 HtmlService.createHtmlOutputFromFile("Sidebar")
 );
}