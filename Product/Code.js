function onOpen(e) {
  DocumentApp.getUi()
    .createMenu('Better Word Counter')
    .addItem('Open', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Better Word Counter');
  DocumentApp.getUi().showSidebar(html);
}

function getWordCount() {
  const doc = DocumentApp.getActiveDocument();
  const text = doc.getBody().getText();
  const words = text.trim()
    .split(/\s+/)
    .filter(word => word.length > 0 && !/^[-–—]+$/.test(word));
  return words.length;
}

