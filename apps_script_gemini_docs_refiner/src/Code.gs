function onOpen() {
  DocumentApp.getUi()
    .createMenu("AI Tools")
    .addItem("AI Content Refiner", "showRefinerSidebar")
    .addToUi();
}

function showRefinerSidebar() {
  DocumentApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("Sidebar")
      .setTitle("AI Content Refiner")
  );
}