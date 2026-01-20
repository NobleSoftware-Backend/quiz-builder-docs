/**
 * Core.js - Parser and Model
 */

class QuizParser {
  /**
   * Parses document body into a Quiz Model.
   * @param {GoogleAppsScript.Document.Body} body 
   * @returns {{quizType: 'MCQ'|'ESSAY', questions: Array, images: Object}}
   */
  static parse(body) {
    const result = {
      quizType: null,
      questions: [],
      images: {} // { filename: blob }
    };

    // Strict: BEGIN must be the first non-empty line.
    // We treat the first non-empty PARAGRAPH/LIST_ITEM text as the line.
    // Any other element (table/hr/image) before BEGIN is an error.
    const numChildren = body.getNumChildren();
    let beginIndex = -1;
    let beginLine = 0;
    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const type = element.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        const t = element.asParagraph().getText().trim();
        if (!t) continue;
        const begin = this._parseBeginTag(t);
        if (!begin) {
          throw new Error(`Missing quiz type header. The first non-empty line must be [BEGIN#MCQ] or [BEGIN#ESSAY]. (Found: "${t}")`);
        }
        result.quizType = begin.quizType;
        beginIndex = i;
        beginLine = i + 1;
        break;
      }

      if (type === DocumentApp.ElementType.LIST_ITEM) {
        const t = element.asListItem().getText().trim();
        if (!t) continue;
        const begin = this._parseBeginTag(t);
        if (!begin) {
          throw new Error(`Missing quiz type header. The first non-empty line must be [BEGIN#MCQ] or [BEGIN#ESSAY]. (Found: "${t}")`);
        }
        result.quizType = begin.quizType;
        beginIndex = i;
        beginLine = i + 1;
        break;
      }

      // Any other element before BEGIN is considered non-empty structure.
      throw new Error('Missing quiz type header. The first non-empty line must be [BEGIN#MCQ] or [BEGIN#ESSAY].');
    }

    if (!result.quizType) {
      throw new Error('Missing quiz type header. The first non-empty line must be [BEGIN#MCQ] or [BEGIN#ESSAY].');
    }
    
    let currentQuestion = null;
    let state = 'WAITING'; // WAITING, READING_QUESTION, READING_OPTIONS, READING_DESCRIPTIONS

    const closeQuestionListIfOpen = () => {
      if (!currentQuestion || !currentQuestion._listOpen) return;
      currentQuestion.contentHtmlParts.push(`</${currentQuestion._listTag}>`);
      currentQuestion._listOpen = false;
      currentQuestion._listTag = null;
      currentQuestion._listAttrs = '';
    };

    const closeDescriptionListIfOpen = () => {
      if (!currentQuestion || !currentQuestion._descListOpen) return;
      currentQuestion.descriptionHtmlParts.push(`</${currentQuestion._descListTag}>`);
      currentQuestion._descListOpen = false;
      currentQuestion._descListTag = null;
      currentQuestion._descListAttrs = '';
    };
    
    for (let i = 0; i < numChildren; i++) {
      if (i === beginIndex) continue;
      const element = body.getChild(i);
      const type = element.getType();
      let text = '';

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        text = element.asParagraph().getText().trim();
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        text = element.asListItem().getText().trim();
      }

      // Strict: BEGIN must not appear anywhere else.
      if (text && /^\[BEGIN#/.test(text)) {
        throw new Error(`Invalid BEGIN header placement at line ${i + 1}. [BEGIN#${result.quizType}] must be the first non-empty line (currently at line ${beginLine}).`);
      }

      // 1. State Transitions
      const questionTag = this._parseQuestionTag(text);
      if (questionTag) {
        // Close previous
        if (currentQuestion) {
          closeQuestionListIfOpen();
          closeDescriptionListIfOpen();
          result.questions.push(currentQuestion);
        }

        // Start new
        currentQuestion = this._createQuestionObj(result.questions.length + 1, i + 1);
        state = 'READING_QUESTION';
        
        // Handle inline text: "[QUESTION] What is...?" or "[QUESTION#7] What is...?"
        const inlineText = questionTag.inlineText;
        if (inlineText) {
            currentQuestion.contentHtmlParts.push(`<p>${this._escapeHtml(inlineText)}</p>`);
        }
        continue;
      }
      
      if (text.startsWith('[OPTIONS]')) {
        if (!currentQuestion) {
          throw new Error(`Invalid [OPTIONS] placement at line ${i + 1}. [OPTIONS] must appear after a [QUESTION#n] tag.`);
        }
        if (result.quizType === 'ESSAY') {
          throw new Error(`Invalid [OPTIONS] for ESSAY at line ${i + 1}. Essay questions must not contain [OPTIONS].`);
        }
        if (currentQuestion.hasOptionsTag) {
          throw new Error(`Duplicate [OPTIONS] tag for Question ${currentQuestion.num} at line ${i + 1}. Only one [OPTIONS] section is allowed per question.`);
        }
        currentQuestion.hasOptionsTag = true;
        currentQuestion.optionsTagLine = i + 1;
        closeQuestionListIfOpen();
        closeDescriptionListIfOpen();
        state = 'READING_OPTIONS';
        continue;
      }

      if (text.startsWith('[DESCRIPTIONS]')) {
        if (!currentQuestion) {
          throw new Error(`Invalid [DESCRIPTIONS] placement at line ${i + 1}. [DESCRIPTIONS] must appear after a [QUESTION#n] tag.`);
        }
        // For MCQ: DESCRIPTIONS must come after OPTIONS
        // For ESSAY: DESCRIPTIONS comes after question content
        if (result.quizType === 'MCQ' && state !== 'READING_OPTIONS') {
          throw new Error(`Invalid [DESCRIPTIONS] placement at line ${i + 1}. For MCQ, [DESCRIPTIONS] must appear after [OPTIONS].`);
        }
        if (result.quizType === 'ESSAY' && state !== 'READING_QUESTION') {
          throw new Error(`Invalid [DESCRIPTIONS] placement at line ${i + 1}. For ESSAY, [DESCRIPTIONS] must appear after the question content.`);
        }
        if (currentQuestion.hasDescriptionsTag) {
          throw new Error(`Duplicate [DESCRIPTIONS] tag for Question ${currentQuestion.num} at line ${i + 1}. Only one [DESCRIPTIONS] section is allowed per question.`);
        }
        currentQuestion.hasDescriptionsTag = true;
        currentQuestion.descriptionsTagLine = i + 1;
        closeQuestionListIfOpen();
        closeDescriptionListIfOpen();
        state = 'READING_DESCRIPTIONS';
        continue;
      }

      // 2. Content Processing based on State
      if (!currentQuestion) continue;

      if (state === 'READING_QUESTION') {
        if (type === DocumentApp.ElementType.PARAGRAPH) {
          closeQuestionListIfOpen();
           const html = this._extractHtml(element.asParagraph(), result.images, currentQuestion.id);
           if (html) currentQuestion.contentHtmlParts.push(`<p>${html}</p>`);
        } else if (type === DocumentApp.ElementType.TABLE) {
          closeQuestionListIfOpen();
           const html = this._extractTableHtml(element.asTable(), result.images, currentQuestion.id);
           currentQuestion.contentHtmlParts.push(html);
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          // Treated as part of question description, not options
           const html = this._extractHtml(element.asListItem(), result.images, currentQuestion.id);

          const info = this._htmlListInfoFromGlyph(element.asListItem().getGlyphType());
          const nextTag = info.tag;
          const nextAttrs = info.attrs;

          if (!currentQuestion._listOpen || currentQuestion._listTag !== nextTag || currentQuestion._listAttrs !== nextAttrs) {
           closeQuestionListIfOpen();
           currentQuestion.contentHtmlParts.push(`<${nextTag}${nextAttrs}>`);
           currentQuestion._listOpen = true;
           currentQuestion._listTag = nextTag;
           currentQuestion._listAttrs = nextAttrs;
          }

          currentQuestion.contentHtmlParts.push(`<li>${html}</li>`);
        }
      } 
      else if (state === 'READING_OPTIONS') {
        if (type === DocumentApp.ElementType.LIST_ITEM) {
          const item = element.asListItem();
          const rawText = item.getText().trim();
          let isCorrect = false;
          const choiceText = rawText.replace(/^\s*<>\s*/, '').trim().replace(/\s+/g, ' ');

          // Check if it marks Correct Answer (<> at start)
          // Supported: "<> Option text" (leading whitespace ok)
          // NOTE: We use getText() for detection so we don't depend on HTML escaping.
          if (/^\s*<>\s*/.test(rawText)) {
            isCorrect = true;
          }

          // Extract HTML. If it's marked correct, strip the leading "<>" marker
          // from the generated HTML (so preview/export never shows the marker).
          const htmlContext = isCorrect ? { stripLeadingCorrectMarker: true, _markerStripped: false } : undefined;
          const html = this._extractHtml(item, result.images, currentQuestion.id, htmlContext);

          currentQuestion.options.push({
            id: `opt_${currentQuestion.options.length}`,
            text: html,
            choiceText: choiceText,
            isCorrect: isCorrect,
            line: i + 1
          });
        }
      }
      else if (state === 'READING_DESCRIPTIONS') {
        if (type === DocumentApp.ElementType.PARAGRAPH) {
          closeDescriptionListIfOpen();
          const html = this._extractHtml(element.asParagraph(), result.images, currentQuestion.id);
          if (html) currentQuestion.descriptionHtmlParts.push(`<p>${html}</p>`);
        } else if (type === DocumentApp.ElementType.TABLE) {
          closeDescriptionListIfOpen();
          const html = this._extractTableHtml(element.asTable(), result.images, currentQuestion.id);
          currentQuestion.descriptionHtmlParts.push(html);
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          const html = this._extractHtml(element.asListItem(), result.images, currentQuestion.id);

          const info = this._htmlListInfoFromGlyph(element.asListItem().getGlyphType());
          const nextTag = info.tag;
          const nextAttrs = info.attrs;

          if (!currentQuestion._descListOpen || currentQuestion._descListTag !== nextTag || currentQuestion._descListAttrs !== nextAttrs) {
            closeDescriptionListIfOpen();
            currentQuestion.descriptionHtmlParts.push(`<${nextTag}${nextAttrs}>`);
            currentQuestion._descListOpen = true;
            currentQuestion._descListTag = nextTag;
            currentQuestion._descListAttrs = nextAttrs;
          }

          currentQuestion.descriptionHtmlParts.push(`<li>${html}</li>`);
        }
      }
    }

    // Push final question
    if (currentQuestion) {
      closeQuestionListIfOpen();
      closeDescriptionListIfOpen();
      result.questions.push(currentQuestion);
    }

    // Finalize content strings
    result.questions.forEach(q => {
      q.content = q.contentHtmlParts.join('');
      const desc = q.descriptionHtmlParts.join('');
      q.descriptions = desc ? desc : null;
    });

    return result;
  }

  /**
   * Parses a question tag marker.
   * Supported:
   *  - [QUESTION]
   *  - [QUESTION#<digits>]
   *
   * Any other [QUESTION...] format is rejected.
   *
   * @param {string} text
   * @returns {{inlineText: string, labelNumber?: number} | null}
   */
  static _parseQuestionTag(text) {
    const t = String(text || '').trim();
    if (!t.startsWith('[QUESTION')) return null;

    if (t.startsWith('[QUESTION]')) {
      return { inlineText: t.slice('[QUESTION]'.length).trim() };
    }

    const m = t.match(/^\[QUESTION#(\d+)\](.*)$/);
    if (m) {
      return { inlineText: String(m[2] || '').trim(), labelNumber: Number(m[1]) };
    }

    throw new Error(`Invalid question tag "${t}". Use [QUESTION] or [QUESTION#<number>] (e.g., [QUESTION#1]).`);
  }

  /**
   * Parses the required document header.
   * Supported:
   *  - [BEGIN#MCQ]
   *  - [BEGIN#ESSAY]
   *
   * Must be on its own line (no trailing text).
   * @param {string} text
   * @returns {{quizType: 'MCQ'|'ESSAY'} | null}
   */
  static _parseBeginTag(text) {
    const t = String(text || '').trim();
    const m = t.match(/^\[BEGIN#(MCQ|ESSAY)\]$/);
    if (!m) return null;
    return { quizType: m[1] };
  }

  static _createQuestionObj(num, line) {
    return {
      num: num,
      id: `q${num}`,
      line: line,
      content: '', // Final HTML string
      contentHtmlParts: [],
      descriptions: null,
      descriptionHtmlParts: [],
      hasOptionsTag: false,
      optionsTagLine: 0,
      hasDescriptionsTag: false,
      descriptionsTagLine: 0,
      options: []  // { text, isCorrect, line }
      ,_listOpen: false
      ,_listTag: null
      ,_listAttrs: ''
      ,_descListOpen: false
      ,_descListTag: null
      ,_descListAttrs: ''
    };
  }

  static _extractHtml(element, imageStore, questionId, context) {
    const parts = [];
    const num = element.getNumChildren();
    
    for (let i = 0; i < num; i++) {
        parts.push(this._extractNode(element.getChild(i), imageStore, questionId, context));
    }
    return parts.join('');
  }

  static _extractNode(child, imageStore, questionId, context) {
    const type = child.getType();

    if (type === DocumentApp.ElementType.TEXT) {
        const textEl = child.asText ? child.asText() : child;
        if (context && context.inEquation) return this._escapeHtml(textEl.getText());
      return this._extractStyledText(textEl, context);

    } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
        const existingCount = Object.keys(imageStore).filter(k => k.startsWith(questionId + '_')).length;
        const localCount = existingCount + 1;
        
      const inlineImg = child.asInlineImage();
      const blob = inlineImg.getBlob();
        const ext = this._getExtension(blob.getContentType());
        const filename = `${questionId}_img_${String(localCount).padStart(3, '0')}.${ext}`;
        
        imageStore[filename] = blob;

      const w = inlineImg.getWidth ? inlineImg.getWidth() : null;
      const h = inlineImg.getHeight ? inlineImg.getHeight() : null;
      const widthAttr = (typeof w === 'number' && isFinite(w) && w > 0) ? ` width="${Math.round(w)}"` : '';
      const heightAttr = (typeof h === 'number' && isFinite(h) && h > 0) ? ` height="${Math.round(h)}"` : '';

      // width/height preserve the user's resize in Google Docs; max-width makes it responsive in preview.
      return `<img src="images/${filename}"${widthAttr}${heightAttr} style="max-width: 100%; height: auto;" />`;

    } else if (type === DocumentApp.ElementType.EQUATION) {
        // Wrap entire equation in MathJax inline delimiters
      return '$' + this._extractHtml(child, imageStore, questionId, Object.assign({}, context, { inEquation: true })) + '$';

    } else if (type === DocumentApp.ElementType.EQUATION_FUNCTION) {
      const funcCodeRaw = child.getCode ? child.getCode() : '';
      const funcCode = String(funcCodeRaw || '').trim().replace(/^\\+/, '');
        const args = [];
        const numKids = child.getNumChildren();
        for (let j = 0; j < numKids; j++) {
          args.push(this._extractNode(child.getChild(j), imageStore, questionId, Object.assign({}, context, { inEquation: true })));
        }

        if (funcCode === 'frac') {
            // Robustly handle cases with >2 children (e.g. ["100", "", "\\times"])
            // by treating the first child as numerator and combining the rest as denominator.
            const numerator = args[0] || '';
            const denominator = args.slice(1).join('');
            return `\\frac{${numerator}}{${denominator}}`;
        } else if (funcCode === 'root') {
            // root(base, index) -> \sqrt[index]{base}
            // Note: Docs API might put index at args[1], base at args[0].
            if (args.length === 2 && args[1]) {
                return `\\sqrt[${args[1]}]{${args[0]}}`;
            }
            return `\\sqrt{${args[0]}}`;
        } else if (funcCode === 'super' || funcCode === 'superscript') {
            return `^{${args[0] || ''}}`;
        } else if (funcCode === 'sub' || funcCode === 'subscript') {
            return `_{${args[0] || ''}}`;
        }
        
        // Default generic function handling
        // Attempt to produce \func{arg1}{arg2}...
        if (!funcCode) return args.join('');
        return `\\${funcCode}` + args.map(a => `{${a}}`).join('');

    } else if (type === DocumentApp.ElementType.EQUATION_SYMBOL) {
        const code = child.getCode ? child.getCode() : '';
        if (code) return this._latexForEquationSymbolCode(code);

        // Fallback if no code
        return child.getText ? this._escapeHtml(child.getText()) : '';
    }

    return '';
  }

  static _latexForEquationSymbolCode(code) {
    // Google Docs equation editor returns internal codes, typically without leading backslash.
    // For variables like x, y, X, 2 we should return the literal.
    // For known operators/greek/etc we return a LaTeX command.
    const trimmed = String(code || '').trim().replace(/^\\+/, '');
    if (!trimmed) return '';

    // Some Docs equation editor tokens come through as words like "superscript"/"subscript".
    // These are not valid TeX commands in MathJax, so map them.
    if (trimmed === 'superscript' || trimmed === 'super') return '^';
    if (trimmed === 'subscript' || trimmed === 'sub') return '_';

    // Single alphanumeric: treat as literal variable/number
    if (/^[A-Za-z0-9]$/.test(trimmed)) return trimmed;

    // Some users type letters; Docs might return multi-letter variables as-is.
    // If it's purely letters but not a known command, treat it as literal text.
    const knownCommands = new Set([
      // Greek (common)
      'alpha','beta','gamma','delta','epsilon','zeta','eta','theta','iota','kappa','lambda','mu','nu','xi','omicron','pi','rho','sigma','tau','upsilon','phi','chi','psi','omega',
      'Gamma','Delta','Theta','Lambda','Xi','Pi','Sigma','Upsilon','Phi','Psi','Omega',

      // Operators/relations (common)
      'pm','mp','times','div','cdot','circ','ast','star','oplus','otimes',
      'leq','geq','neq','approx','sim','equiv','propto',
      'in','notin','subset','subseteq','supset','supseteq','cup','cap',
      'forall','exists','neg','land','lor',
      'rightarrow','leftarrow','leftrightarrow','Rightarrow','Leftarrow','Leftrightarrow',
      'infty'
    ]);

    if (/^[A-Za-z]+$/.test(trimmed) && !knownCommands.has(trimmed)) {
      return trimmed;
    }

    return `\\${trimmed}`;
  }

  static _extractTableHtml(table, imageStore, questionId) {
    let html = '<table>';
    for(let r=0; r<table.getNumRows(); r++) {
        html += '<tr>';
        const row = table.getRow(r);
        for(let c=0; c<row.getNumCells(); c++) {
            html += '<td>';
            const cell = row.getCell(c);
            const numChildren = cell.getNumChildren();

            let listOpen = false;
            let listTag = null;
            let listAttrs = '';
            const closeList = () => {
              if (!listOpen) return;
              html += `</${listTag}>`;
              listOpen = false;
              listTag = null;
              listAttrs = '';
            };
            
            for (let k = 0; k < numChildren; k++) {
                const block = cell.getChild(k);
                const type = block.getType();
                
                if (type === DocumentApp.ElementType.PARAGRAPH) {
                     closeList();
                     html += this._extractHtml(block.asParagraph(), imageStore, questionId) + '<br/>';
                } else if (type === DocumentApp.ElementType.LIST_ITEM) {
                     const li = block.asListItem();
                     const info = this._htmlListInfoFromGlyph(li.getGlyphType());
                     if (!listOpen || listTag !== info.tag || listAttrs !== info.attrs) {
                       closeList();
                       html += `<${info.tag}${info.attrs}>`;
                       listOpen = true;
                       listTag = info.tag;
                       listAttrs = info.attrs;
                     }
                     html += '<li>' + this._extractHtml(li, imageStore, questionId) + '</li>';
                } else if (type === DocumentApp.ElementType.TABLE) {
                     closeList();
                     html += this._extractTableHtml(block.asTable(), imageStore, questionId);
                }
            }
            closeList();
            html += '</td>';
        }
        html += '</tr>';
    }
    html += '</table>';
    return html;
  }

  static _htmlListInfoFromGlyph(glyphType) {
    // Default to UL for common bullet glyphs; otherwise use OL with an HTML type if possible.
    const g = glyphType;

    // Bullet glyphs
    if (
      g === DocumentApp.GlyphType.BULLET ||
      g === DocumentApp.GlyphType.HOLLOW_BULLET ||
      g === DocumentApp.GlyphType.SQUARE_BULLET
    ) {
      return { tag: 'ul', attrs: '' };
    }

    // Numbered/lettered lists
    const map = {
      [DocumentApp.GlyphType.NUMBER]: '1',
      [DocumentApp.GlyphType.LATIN_UPPER]: 'A',
      [DocumentApp.GlyphType.LATIN_LOWER]: 'a',
      [DocumentApp.GlyphType.ROMAN_UPPER]: 'I',
      [DocumentApp.GlyphType.ROMAN_LOWER]: 'i'
    };

    const typeAttr = map[g];
    if (typeAttr) return { tag: 'ol', attrs: ` type="${typeAttr}"` };

    // Fallback
    return { tag: 'ul', attrs: '' };
  }

  static _extractStyledText(textEl, context) {
    const text = textEl && textEl.getText ? String(textEl.getText() || '') : '';
    if (!text) return '';

    const indices = (textEl.getTextAttributeIndices && textEl.getTextAttributeIndices()) || [0];
    const parts = [];

    for (let i = 0; i < indices.length; i++) {
      const start = indices[i];
      const end = (i + 1 < indices.length) ? indices[i + 1] : text.length;
      let chunk = text.substring(start, end);
      if (!chunk) continue;

      // Strip leading correct marker "<>" only once, at the very start of an option.
      if (context && context.stripLeadingCorrectMarker && !context._markerStripped) {
        const before = chunk;
        chunk = chunk.replace(/^\s*<>\s*/, '');
        if (chunk !== before) {
          context._markerStripped = true;
        }
      }

      // If the first chunk didn't contain the marker, do not keep trying forever.
      // This prevents stripping "<>" that appears later in the option text.
      if (context && context.stripLeadingCorrectMarker && !context._markerStripped) {
        context._markerStripped = true;
      }

      const attrs = (textEl.getAttributes && textEl.getAttributes(start)) || {};
      let html = this._escapeHtml(chunk).replace(/\n/g, '<br/>');

      const valign = attrs[DocumentApp.Attribute.VERTICAL_ALIGNMENT];
      if (valign === DocumentApp.VerticalAlignment.SUPERSCRIPT) html = `<sup>${html}</sup>`;
      else if (valign === DocumentApp.VerticalAlignment.SUBSCRIPT) html = `<sub>${html}</sub>`;

      if (attrs[DocumentApp.Attribute.STRIKETHROUGH]) html = `<s>${html}</s>`;
      if (attrs[DocumentApp.Attribute.UNDERLINE]) html = `<u>${html}</u>`;
      if (attrs[DocumentApp.Attribute.ITALIC]) html = `<i>${html}</i>`;
      if (attrs[DocumentApp.Attribute.BOLD]) html = `<b>${html}</b>`;

      const linkUrl = attrs[DocumentApp.Attribute.LINK_URL];
      if (linkUrl) {
        const safeHref = this._escapeHtmlAttr(String(linkUrl));
        html = `<a href="${safeHref}" target="_blank" rel="noopener noreferrer">${html}</a>`;
      }

      parts.push(html);
    }

    return parts.join('');
  }

  static _escapeHtmlAttr(text) {
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  static _escapeHtml(text) {
    if (!text) return '';
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  static _getExtension(mime) {
    const map = {'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif'};
    return map[mime] || 'png';
  }
}
