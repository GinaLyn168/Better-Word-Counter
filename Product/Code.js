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

function normalizeText(text) {
  return text ? text.replace(/\s+/g, ' ').trim() : '';
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function phraseMatches(outerText, innerText) {
  const escapedInner = escapeRegex(innerText);
  const regex = new RegExp('(^|\\s)' + escapedInner + '(?=\\s|$)', 'i');
  return regex.test(outerText);
}

function countWords(text) {
  return normalizeText(text).split(/\s+/)
    .filter(word => word.length > 0 && !/^[-–—]+$/.test(word))
    .length;
}

function getWordCount() {
  // The following parts were coded by Ranbir
  const doc = DocumentApp.getActiveDocument();
  const text = normalizeText(doc.getBody().getText());
  
  const props = PropertiesService.getDocumentProperties();
  const exclusions = JSON.parse(props.getProperty('exclusions') || '[]');
  
  const totalWords = countWords(text);
  
  let excludedCount = 0;
  // This algorithm for finding the exclusions in the text was aided by AI a lot
  for (const phrase of exclusions) {
    const normalizedPhrase = normalizeText(phrase);
    if (normalizedPhrase.length === 0) continue;
    const escapedPhrase = normalizedPhrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escapedPhrase, 'gi');
    const matches = text.match(regex);
    if (matches) {
      const phraseWordCount = countWords(normalizedPhrase);
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
  const pieces = [];

  for (const element of elements) {
    if (element.getElement().editAsText) {
      const text = element.getElement().asText().getText();
      if (element.isPartial()) {
        pieces.push(text.substring(element.getStartOffset(), element.getEndOffsetInclusive() + 1));
      } else {
        pieces.push(text);
      }
    }
  }

  return normalizeText(pieces.join(' '));
}

// This functionw as made by Ranbir, with use of minimal AI to understand how to save the exclusion and syntax
function saveExclusion(text) {
  const docProperties = PropertiesService.getDocumentProperties();
  const exclusions = JSON.parse(docProperties.getProperty('exclusions') || '[]');
  const normalizedText = normalizeText(text);
  if (normalizedText.length === 0) return exclusions;

  if (exclusions.includes(normalizedText)) {
    return exclusions;
  }

  // If the new excluded phrase is already contained inside an existing excluded phrase,
  // it is redundant and should not be added.
  const alreadyCovered = exclusions.some(existing =>
    phraseMatches(existing, normalizedText)
  );
  if (alreadyCovered) {
    return exclusions;
  }

  // If an existing exclusion is fully contained inside the new phrase, remove the smaller one.
  const filtered = exclusions.filter(existing =>
    !phraseMatches(normalizedText, existing)
  );

  filtered.push(normalizedText);
  docProperties.setProperty('exclusions', JSON.stringify(filtered));
  return filtered;
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
  const exclusions = JSON.parse(props.getProperty('exclusions') || '[]');
  const text = normalizeText(DocumentApp.getActiveDocument().getBody().getText());

  return exclusions.map(phrase => {
    const normalizedPhrase = normalizeText(phrase);
    const escapedPhrase = normalizedPhrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escapedPhrase, 'gi');
    const found = Boolean(text.match(regex));
    return {
      phrase: normalizedPhrase,
      missing: !found,
    };
  });
}