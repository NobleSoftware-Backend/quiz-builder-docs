/**
 * Code.js - Controller and Entry Points
 */

/**
 * Runs when the add-on is installed.
 * @param {Object} e
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Quiz Builder')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Wrap as Question Block', 'wrapAsQuestionBlock')
    .addItem('Wrap as Options', 'wrapAsOptions')
    .addItem('Renumber Question Labels', 'renumberQuestionLabels')
    .addSeparator()
    .addItem('Validate Format', 'validateAndShowResults')
    .addItem('Export Quiz', 'exportQuiz')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Quiz Builder')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Returns the current quiz type header if present.
 * Header must be the first non-empty line: [BEGIN#MCQ] or [BEGIN#ESSAY]
 */
function getQuizTypeHeader() {
  const body = DocumentApp.getActiveDocument().getBody();
  const info = _getQuizTypeHeaderInfo_(body);
  return { type: info ? info.type : null, line: info ? info.line : 0 };
}

/**
 * Clears the document and writes a new quiz type header as the first non-empty line.
 * This is the supported way to change quiz type.
 */
function resetQuizAndSetType(type) {
  const quizType = String(type || '').toUpperCase();
  if (quizType !== 'MCQ' && quizType !== 'ESSAY') throw new Error('Invalid quiz type. Use MCQ or ESSAY.');

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // Non-destructive: insert header at the very top.
  // This guarantees the first non-empty line is the BEGIN tag.
  const header = body.insertParagraph(0, `[BEGIN#${quizType}]`);
  header.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  header.editAsText().setBold(true).setForegroundColor('#444444');

  // Add a spacer after the header for readability.
  body.insertParagraph(1, '');

  // Strict parser forbids BEGIN anywhere else.
  // Clear any other BEGIN tags that might exist further down.
  const n = body.getNumChildren();
  for (let i = 2; i < n; i++) {
    const el = body.getChild(i);
    const t = el.getType();
    if (t !== DocumentApp.ElementType.PARAGRAPH && t !== DocumentApp.ElementType.LIST_ITEM) continue;
    const container = t === DocumentApp.ElementType.PARAGRAPH ? el.asParagraph() : el.asListItem();
    const raw = String(container.getText() || '').trim();
    if (!raw) continue;
    if (/^\[BEGIN#(MCQ|ESSAY)\]$/.test(raw) || /^\[BEGIN#/.test(raw)) {
      container.setText('');
    }
  }

  return true;
}

function _getQuizTypeHeaderInfo_(body) {
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    const t = el.getType();
    if (t === DocumentApp.ElementType.PARAGRAPH) {
      const text = el.asParagraph().getText().trim();
      if (!text) continue;
      const m = text.match(/^\[BEGIN#(MCQ|ESSAY)\]$/);
      if (!m) return null;
      return { type: m[1], line: i + 1 };
    }
    if (t === DocumentApp.ElementType.LIST_ITEM) {
      const text = el.asListItem().getText().trim();
      if (!text) continue;
      const m = text.match(/^\[BEGIN#(MCQ|ESSAY)\]$/);
      if (!m) return null;
      return { type: m[1], line: i + 1 };
    }
    // Any other element means header is missing/invalid
    return null;
  }
  return null;
}

/**
 * Validates the document and shows a report.
 * Orchestrates calls to QuizParser (Core) and QuizValidator (Services).
 */
function validateAndShowResults() {
  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi();

  try {
    // 1. Parse (Core)
    const parseResult = QuizParser.parse(doc.getBody());
    
    // 2. Validate (Service)
    const validation = QuizValidator.validate(parseResult);

    // 3. Report Results
    let message = `📊 Validation Results:\n\n`;
    message += `Questions found: ${parseResult.questions.length}\n`;
    message += `Total images: ${Object.keys(parseResult.images).length}\n\n`;

    if (validation.errors.length > 0) {
      message += `❌ ERRORS (${validation.errors.length}):\n`;
      validation.errors.forEach(err => message += `  • ${err.message}\n`);
      message += `\n`;
    }

    if (validation.warnings.length > 0) {
      message += `⚠️ WARNINGS (${validation.warnings.length}):\n`;
      validation.warnings.forEach(warn => message += `  • ${warn.message}\n`);
      message += `\n`;
    }

    if (validation.isValid) {
      const status = validation.warnings.length > 0 ? "Passed with Warnings" : "Passed";
      ui.alert(status, message, ui.ButtonSet.OK);
    } else {
      ui.alert('Validation Failed', message + "Please fix errors before exporting.", ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('System Error', e.toString(), ui.ButtonSet.OK);
    console.error(e);
  }
}

/**
 * Exports the quiz to a ZIP file.
 */
function exportQuiz() {
  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi();

  try {
    // 1. Parse
    const parseResult = QuizParser.parse(doc.getBody());

    // 2. Validate
    const validation = QuizValidator.validate(parseResult);
    if (!validation.isValid) {
      ui.alert('❌ Export Blocked', 'Validation failed. Please run "Validate Format" to see errors.', ui.ButtonSet.OK);
      return;
    }

    // 3. Generate ZIP (Service)
    const zipBlob = QuizExporter.generateZip(parseResult);
    const sizeMB = (zipBlob.getBytes().length / (1024 * 1024)).toFixed(2);

    // 4. Handle Download/Drive Save logic
    if (zipBlob.getBytes().length > 50 * 1024 * 1024) {
      ui.alert('❌ Export Failed', `ZIP too large (${sizeMB}MB). Max 50MB.`, ui.ButtonSet.OK);
    } else if (zipBlob.getBytes().length > 10 * 1024 * 1024) {
      const fileName = `quiz_export_${Date.now()}.zip`;
      const file = DriveApp.createFile(zipBlob.setName(fileName));
      ui.alert('✅ Export Saved to Drive', `File: ${fileName}\nSize: ${sizeMB}MB\nLink: ${file.getUrl()}`, ui.ButtonSet.OK);
    } else {
      // Direct Download - Using a workaround for Apps Script dialogs to trigger download
      const fileName = `quiz_export_${Date.now()}.zip`;
      const downloadUrl = 'data:application/zip;base64,' + Utilities.base64Encode(zipBlob.getBytes());
      const html = HtmlService.createHtmlOutput(`
        <html>
          <body style="font-family: Arial, sans-serif; text-align: center; padding: 20px;">
            <h3>Download Ready</h3>
            <p>Filename: ${fileName}</p>
            <p>Size: ${sizeMB} MB</p>
            <a href="${downloadUrl}" download="${fileName}" class="btn" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; display: inline-block; margin-top: 10px;">Download Now</a>
            <script>
               // Auto-click attempt
               setTimeout(function() { document.querySelector('a').click(); }, 1000);
            </script>
          </body>
        </html>
      `).setWidth(400).setHeight(200);
      ui.showModalDialog(html, 'Export Complete');
    }

  } catch (e) {
    ui.alert('Export Error', e.toString(), ui.ButtonSet.OK);
    console.error(e);
  }
}

/**
 * Returns a list of questions for the Preview dropdown.
 */
function getPreviewQuestionList() {
  const doc = DocumentApp.getActiveDocument();
  const parseResult = QuizParser.parse(doc.getBody());
  return parseResult.questions.map(q => ({ value: q.id, label: `Question ${q.num} (${q.id})` }));
}

/**
 * Opens a modal dialog showing a MathJax-rendered preview.
 * @param {string} questionIdOrAll - 'all' or specific question id like 'q1'
 */
function openPreviewDialog(questionIdOrAll) {
  const doc = DocumentApp.getActiveDocument();
  const parseResult = QuizParser.parse(doc.getBody());

  // Build image data-uri map for preview
  const imageData = {};
  for (const [filename, blob] of Object.entries(parseResult.images || {})) {
    const mime = blob.getContentType ? blob.getContentType() : 'image/png';
    const b64 = Utilities.base64Encode(blob.getBytes());
    imageData[`images/${filename}`] = `data:${mime};base64,${b64}`;
  }

  const replaceImages = (html) => {
    if (!html) return '';
    return String(html).replace(/src\s*=\s*"(images\/[^"]+)"/g, function(_m, path) {
      const dataUrl = imageData[path];
      if (!dataUrl) return `src="${path}"`;
      return `src="${dataUrl}"`;
    });
  };

  const wanted = String(questionIdOrAll || 'all');
  const questions = wanted === 'all'
    ? parseResult.questions
    : parseResult.questions.filter(q => q.id === wanted);

  const fragmentParts = [];
  questions.forEach(q => {
    const content = replaceImages(q.content);
    const desc = replaceImages(q.descriptions);
    const opts = (q.options || []).map(o => ({
      html: replaceImages(o.text),
      isCorrect: !!o.isCorrect
    }));

    fragmentParts.push(
      `<div class="question">` +
        `<div class="q-header">` +
          `<div class="q-title">Question ${q.num}</div>` +
          `<div class="q-id">${q.id}</div>` +
        `</div>` +
        `<div class="q-content">${content}</div>` +
        `<ol class="options" type="A">` +
          opts.map(o => `<li${o.isCorrect ? ' class="is-correct"' : ''}>${o.html}</li>`).join('') +
        `</ol>` +
        (desc ? `<div class="q-content q-section q-descriptions"><div class="q-section-title">Descriptions</div>${desc}</div>` : ``) +
      `</div>`
    );
  });

  const template = HtmlService.createTemplateFromFile('Preview');
  template.previewHtml = fragmentParts.join('');
  const html = template.evaluate().setWidth(980).setHeight(720);
  DocumentApp.getUi().showModalDialog(html, 'Quiz Preview');
}

// -- Existing UI Helper Functions (Refactored) --

function wrapAsQuestionBlock() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return DocumentApp.getUi().alert('Select text first!');
  
  const elements = selection.getRangeElements();
  const body = doc.getBody();
  const headerInfo = _getQuizTypeHeaderInfo_(body);
  if (!headerInfo) throw new Error('Missing quiz type header. Add [BEGIN#MCQ] or [BEGIN#ESSAY] as the first non-empty line (use the sidebar buttons).');
  const quizType = headerInfo.type;
  let firstContainer = null;

  // Search logic
  for (let i = 0; i < elements.length; i++) {
    let el = elements[i].getElement();
    while (el && el.getData && !el.getText) el = el.getParent();
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH || el.getType() === DocumentApp.ElementType.LIST_ITEM) {
      firstContainer = el;
      break;
    }
  }

  if (!firstContainer) return DocumentApp.getUi().alert('Could not determine insertion point.');

  const index = body.getChildIndex(firstContainer);
  const nextLabel = _getNextQuestionLabelNumber_(body);
  const qTag = body.insertParagraph(index, `[QUESTION#${nextLabel}]`);
  qTag.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  qTag.editAsText().setBold(true).setForegroundColor('#0066CC');

  if (quizType === 'ESSAY') {
    // ESSAY: no [OPTIONS]
    return true;
  }

  // Attempt to find where options might start in selection
  let addedOptions = false;
  elements.forEach(e => {
    const el = e.getElement();
    if (el.getType() === DocumentApp.ElementType.LIST_ITEM && !addedOptions) {
      const idx = body.getChildIndex(el);
      if (idx > index) { // Ensure it's after question
          const oTag = body.insertParagraph(idx, '[OPTIONS]');
          oTag.editAsText().setBold(true).setForegroundColor('#009900');
          addedOptions = true;
      }
    }
  });

  if (!addedOptions) {
     const lastEl = elements[elements.length-1].getElement();
     let p = lastEl;
     while(p.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) p = p.getParent();
     const idx = body.getChildIndex(p);
     const oTag = body.insertParagraph(idx + 1, '[OPTIONS]');
     oTag.editAsText().setBold(true).setForegroundColor('#009900');
  }

  // Final cursor reset: Ensure a clean paragraph exists at the end of selection or after inserted block
  // This is tricky with selection, but for now, let's just ensure we didn't leave the document in a state 
  // where the next enter key inherits style. 
  // Best we can do is ensuring the tags themselves don't bleed style into next lines if the user types inside them.
}

function wrapAsOptions() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return DocumentApp.getUi().alert('Select content first.');

  const body = doc.getBody();
  const headerInfo = _getQuizTypeHeaderInfo_(body);
  if (!headerInfo) throw new Error('Missing quiz type header. Add [BEGIN#MCQ] or [BEGIN#ESSAY] as the first non-empty line (use the sidebar buttons).');
  if (headerInfo.type !== 'MCQ') throw new Error('This document is set to ESSAY. [OPTIONS] is not allowed.');
  
  const elements = selection.getRangeElements();
  const firstEl = elements[0].getElement();
  
  let p = firstEl;
  while(p.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) p = p.getParent();
  
  const idx = body.getChildIndex(p);
  const tag = body.insertParagraph(idx, '[OPTIONS]');
  tag.editAsText().setBold(true).setForegroundColor('#009900');
  
  elements.forEach(r => {
    const el = r.getElement();
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = el.getText();
      if (text.trim()) {
        const i = body.getChildIndex(el);
        body.removeChild(el);
        body.insertListItem(i, text).setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      }
    } else if (el.getType() === DocumentApp.ElementType.LIST_ITEM) {
        el.asListItem().setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
    }
  });
}

function insertNewQuestion() { return _insertTemplateAtCursor(true); }
function insertNewOption() { return _insertTemplateAtCursor(false); }
function addSeparator() { DocumentApp.getActiveDocument().getBody().appendHorizontalRule(); }

function insertDescriptionsSection() {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) throw new Error('Place the cursor where you want to insert [DESCRIPTIONS].');

  const body = doc.getBody();
  let parent = body;
  let index = body.getNumChildren();

  const el = cursor.getElement();
  let c = el;
  while (c.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) c = c.getParent();
  parent = c.getParent();
  index = parent.getChildIndex(c) + 1;

  const tag = parent.insertParagraph(index++, '[DESCRIPTIONS]');
  tag.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  tag.editAsText().setBold(true).setForegroundColor('#6A1B9A');

  const p = parent.insertParagraph(index++, 'Description text...');
  p.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  p.editAsText().setBold(false).setForegroundColor('#000000');

  parent.insertParagraph(index++, '');
  return true;
}

/**
 * Toggles the correct answer mark (<>) on selected list items.
 * Only one option per question can be marked as correct.
 * If marking a new option as correct, removes the mark from any previously marked option.
 * Throws error if selection is not a valid list item.
 */
function toggleCorrectAnswer() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const cursor = doc.getCursor();
  let modified = false;

  /**
   * Finds all list items that belong to the same [OPTIONS] section as the given list item.
   * Returns an array of list item elements.
   */
  const findSiblingOptions = (listItem) => {
    const body = doc.getBody();
    const numChildren = body.getNumChildren();
    
    // Find the index of the current list item in the body
    let currentIndex = -1;
    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
        if (child.asListItem().getText() === listItem.getText() && 
            child.asListItem().getListId() === listItem.getListId()) {
          currentIndex = i;
          break;
        }
      }
    }
    
    if (currentIndex === -1) return [];
    
    // Find the [OPTIONS] tag before this list item
    let optionsStart = -1;
    for (let i = currentIndex - 1; i >= 0; i--) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const text = child.asParagraph().getText().trim();
        if (text.startsWith('[OPTIONS]')) {
          optionsStart = i;
          break;
        }
        // Stop if we hit a [QUESTION] or [DESCRIPTIONS] tag
        if (text.startsWith('[QUESTION') || text.startsWith('[DESCRIPTIONS]')) {
          break;
        }
      }
    }
    
    if (optionsStart === -1) return [];
    
    // Collect all list items after [OPTIONS] until we hit a non-list-item
    const siblings = [];
    for (let i = optionsStart + 1; i < numChildren; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
        siblings.push(child.asListItem());
      } else if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const text = child.asParagraph().getText().trim();
        // Stop at next section tag
        if (text.startsWith('[QUESTION') || text.startsWith('[OPTIONS]') || 
            text.startsWith('[DESCRIPTIONS]') || text.startsWith('[BEGIN#')) {
          break;
        }
        // Empty paragraphs are ok, continue
        if (text) break;
      } else {
        break;
      }
    }
    
    return siblings;
  };

  /**
   * Removes the correct answer mark from a list item if it has one.
   */
  const removeMarkFromItem = (listItem) => {
    const text = listItem.getText();
    const match = text.match(/^\s*<>\s?/);
    if (match) {
      listItem.editAsText().deleteText(0, match[0].length - 1);
    }
  };

  const processElement = (el) => {
    // Navigate up to find List Item
    let current = el;
    while (current && current.getType() !== DocumentApp.ElementType.LIST_ITEM && current.getType() !== DocumentApp.ElementType.BODY_SECTION) {
      current = current.getParent();
    }

    if (current && current.getType() === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = current.asListItem();
      const text = listItem.getText();
      const textObj = listItem.editAsText();
      
      // Check for existing mark (start of string)
      // Regex: Optional whitespace, <>, optional space
      const match = text.match(/^\s*<>\s?/);
      
      if (match) {
        // Remove mark (toggle off)
        textObj.deleteText(0, match[0].length - 1);
      } else {
        // Add mark - but first remove from any sibling options
        const siblings = findSiblingOptions(listItem);
        siblings.forEach(sibling => {
          if (sibling !== listItem) {
            removeMarkFromItem(sibling);
          }
        });
        
        // Now add the mark to this item
        textObj.insertText(0, '<> ');
      }
      modified = true;
    }
  };

  if (selection) {
    selection.getRangeElements().forEach(r => processElement(r.getElement()));
  } else if (cursor) {
    processElement(cursor.getElement());
  } else {
    throw new Error('Please place cursor in an option or select options.');
  }

  if (!modified) {
    throw new Error('Selection is not a bullet list option. convert it first!');
  }
}

function _insertTemplateAtCursor(isQuestion) {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  const body = doc.getBody();
  let parent = body;
  let index = body.getNumChildren();

  if (cursor) {
    const el = cursor.getElement();
    let c = el;
    while (c.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) c = c.getParent();
    parent = c.getParent();
    index = parent.getChildIndex(c) + 1;
  }

  const headerInfo = _getQuizTypeHeaderInfo_(body);
  const quizType = headerInfo ? headerInfo.type : null;

  if (isQuestion) {
    if (!quizType) {
      throw new Error('Missing quiz type header. Add [BEGIN#MCQ] or [BEGIN#ESSAY] as the first non-empty line (use the sidebar buttons).');
    }
    const nextLabel = _getNextQuestionLabelNumber_(body);
    const q = parent.insertParagraph(index++, `[QUESTION#${nextLabel}]`);
    q.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    q.editAsText().setBold(true).setForegroundColor('#0066CC');
    
    // Explicitly reset style for content
    const content = parent.insertParagraph(index++, 'Question text...');
    content.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    content.editAsText().setBold(false).setForegroundColor('#000000');

    // Add empty line for spacing
    parent.insertParagraph(index++, '');

    if (quizType === 'MCQ') {
      const o = parent.insertParagraph(index++, '[OPTIONS]');
      o.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      o.editAsText().setBold(true).setForegroundColor('#009900');
      
      // Explicitly reset style for options
      const opt1 = parent.insertListItem(index++, '<> Option A (Correct)');
      opt1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      opt1.editAsText().setBold(false).setForegroundColor('#000000');

      const opt2 = parent.insertListItem(index++, 'Option B');
      opt2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      opt2.editAsText().setBold(false).setForegroundColor('#000000');
    } else {
      // ESSAY: no [OPTIONS]
      const d = parent.insertParagraph(index++, '[DESCRIPTIONS]');
      d.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      d.editAsText().setBold(true).setForegroundColor('#6A1B9A');
      const dp = parent.insertParagraph(index++, 'Optional explanation...');
      dp.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      dp.editAsText().setBold(false).setForegroundColor('#000000');
    }
  } else {
    if (!quizType) {
      throw new Error('Missing quiz type header. Add [BEGIN#MCQ] or [BEGIN#ESSAY] as the first non-empty line (use the sidebar buttons).');
    }
    if (quizType !== 'MCQ') {
      throw new Error('Options are only allowed for MCQ. This document is set to ESSAY.');
    }
    const item = parent.insertListItem(index, 'New Option');
    item.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
    item.editAsText().setBold(false).setForegroundColor('#000000');
  }
}

function _insertTemplateAtBodyEnd_(body, isQuestion, quizType) {
  // Helper used by resetQuizAndSetType
  // Append using cursor-less logic
  if (isQuestion) {
    const nextLabel = _getNextQuestionLabelNumber_(body);
    const q = body.appendParagraph(`[QUESTION#${nextLabel}]`);
    q.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    q.editAsText().setBold(true).setForegroundColor('#0066CC');
    const content = body.appendParagraph('Question text...');
    content.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    content.editAsText().setBold(false).setForegroundColor('#000000');
    body.appendParagraph('');
    if (quizType === 'MCQ') {
      const o = body.appendParagraph('[OPTIONS]');
      o.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      o.editAsText().setBold(true).setForegroundColor('#009900');
      const opt1 = body.appendListItem('<> Option A (Correct)');
      opt1.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      opt1.editAsText().setBold(false).setForegroundColor('#000000');
      const opt2 = body.appendListItem('Option B');
      opt2.setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
      opt2.editAsText().setBold(false).setForegroundColor('#000000');
    } else {
      const d = body.appendParagraph('[DESCRIPTIONS]');
      d.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      d.editAsText().setBold(true).setForegroundColor('#6A1B9A');
      const dp = body.appendParagraph('Optional explanation...');
      dp.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      dp.editAsText().setBold(false).setForegroundColor('#000000');
    }
  }
}

/**
 * Renumbers all question label tags in the document, rewriting them as:
 *   [QUESTION#1], [QUESTION#2], ... in document order.
 *
 * Notes:
 * - Converts legacy [QUESTION] tags to numbered form.
 * - Rejects invalid tags like [QUESTION#] or other [QUESTION...] variants.
 */
function renumberQuestionLabels() {
  const ui = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  try {
    const numChildren = body.getNumChildren();
    let counter = 0;

    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const type = element.getType();
      if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) continue;

      const container = type === DocumentApp.ElementType.PARAGRAPH ? element.asParagraph() : element.asListItem();
      const raw = String(container.getText() || '');
      const trimmed = raw.trim();
      if (!trimmed.startsWith('[QUESTION')) continue;

      const info = _getQuestionTagInfoFromTrimmed_(trimmed, i + 1);
      counter++;
      const newTag = `[QUESTION#${counter}]`;

      const leadingWsLen = raw.match(/^\s*/)[0].length;
      const textObj = container.editAsText();
      textObj.deleteText(leadingWsLen, leadingWsLen + info.tagLength - 1);
      textObj.insertText(leadingWsLen, newTag);

      // Re-apply styling to the tag only
      textObj.setBold(leadingWsLen, leadingWsLen + newTag.length - 1, true);
      textObj.setForegroundColor(leadingWsLen, leadingWsLen + newTag.length - 1, '#0066CC');
    }

    if (counter === 0) {
      ui.alert('Renumber Question Labels', 'No [QUESTION] / [QUESTION#n] tags found.', ui.ButtonSet.OK);
      return;
    }

    ui.alert('Renumber Question Labels', `Renumbered ${counter} question tag(s) to [QUESTION#1..#${counter}].`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Renumber Failed', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
    console.error(e);
  }
}

/**
 * Counts question tags in the body and returns the next label number.
 * Label numbers are treated as cosmetic; the next label is based on count,
 * not on the largest existing #n.
 */
function _getNextQuestionLabelNumber_(body) {
  const numChildren = body.getNumChildren();
  let count = 0;

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    const type = element.getType();
    if (type !== DocumentApp.ElementType.PARAGRAPH && type !== DocumentApp.ElementType.LIST_ITEM) continue;

    const container = type === DocumentApp.ElementType.PARAGRAPH ? element.asParagraph() : element.asListItem();
    const trimmed = String(container.getText() || '').trim();
    if (!trimmed.startsWith('[QUESTION')) continue;

    // Validate tag format (reject [QUESTION#] etc.) while counting.
    _getQuestionTagInfoFromTrimmed_(trimmed, i + 1);
    count++;
  }

  return count + 1;
}

/**
 * Parses a trimmed text line that begins with [QUESTION...]
 * and returns tag length.
 *
 * Supported:
 *  - [QUESTION]
 *  - [QUESTION#<digits>]
 */
function _getQuestionTagInfoFromTrimmed_(trimmed, lineNumber) {
  const t = String(trimmed || '');
  if (!t.startsWith('[QUESTION')) return null;

  if (t.startsWith('[QUESTION]')) {
    return { tagLength: '[QUESTION]'.length };
  }

  const m = t.match(/^\[QUESTION#\d+\]/);
  if (m) {
    return { tagLength: m[0].length };
  }

  throw new Error(`Invalid question tag on line ${lineNumber}. Use [QUESTION] or [QUESTION#<number>] (e.g., [QUESTION#1]). Found: ${t}`);
}
