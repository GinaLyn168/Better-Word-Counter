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

function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return null;
  
  const elements = selection.getRangeElements();
  let selectedText = '';
  
  for (const element of elements) {
    if (element.getElement().editAsText) {
      const text = element.getElement().asText().getText();
      if (element.isPartial()) {
        selectedText += text.substring(element.getStartOffset(), element.getEndOffsetInclusive() + 1);
      } else {
        selectedText += text;
      }
    }
  }
  
  return selectedText.trim();
}

function saveExclusion(text) {
  const docProperties = PropertiesService.getDocumentProperties();
  const exclusions = JSON.parse(docProperties.getProperty('exclusions') || '[]');
  if (!exclusions.includes(text)) {
    exclusions.push(text);
    docProperties.setProperty('exclusions', JSON.stringify(exclusions));
  }
  return exclusions;
}