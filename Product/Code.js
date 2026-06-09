// this function was made by AI
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
  // The following parts were coded by Ranbir
  const doc = DocumentApp.getActiveDocument();
  const text = doc.getBody().getText().replace(/\n/g, ' ');
  
  const props = PropertiesService.getDocumentProperties();
  const exclusions = JSON.parse(props.getProperty('exclusions') || '[]');
  
  const totalWords = text.trim()
    .split(/\s+/)
    .filter(word => word.length > 0 && !/^[-–—]+$/.test(word))
    .length;
  
  let excludedCount = 0;
  // This algorithm for finding the exclusions in the text was aided by AI a lot
  for (const phrase of exclusions) {
    const trimmedPhrase = phrase.trim();
    if (trimmedPhrase.length === 0) continue;
    const escapedPhrase = trimmedPhrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escapedPhrase, 'gi');
    const matches = text.match(regex);
    if (matches) {
      const phraseWordCount = trimmedPhrase.split(/\s+/).filter(w => w.length > 0).length;
      excludedCount += matches.length * phraseWordCount;
    }
  }
  
  return Math.max(0, totalWords - excludedCount);
}

// This functions beginning was made by Ranbir, but the loop was made by AI
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

// This functionw as made by Ranbir, with use of minimal AI to understand how to save the exclusion and syntax
function saveExclusion(text) {
  const docProperties = PropertiesService.getDocumentProperties();
  const exclusions = JSON.parse(docProperties.getProperty('exclusions') || '[]');
  if (!exclusions.includes(text)) {
    exclusions.push(text);
    docProperties.setProperty('exclusions', JSON.stringify(exclusions));
  }
  return exclusions;
}

// Function made by Ranbir, no use of AI
function clearExclusions() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('exclusions', '[]');
  return [];
}

// Function made by AI
function removeExclusion(phrase) {
  const props = PropertiesService.getDocumentProperties();
  const existing = JSON.parse(props.getProperty('exclusions') || '[]');
  const updated = existing.filter(p => p !== phrase);
  props.setProperty('exclusions', JSON.stringify(updated));
  return updated;
}

// Function made by Ranbir 
function getExclusions() {
  const props = PropertiesService.getDocumentProperties();
  return JSON.parse(props.getProperty('exclusions') || '[]');
}