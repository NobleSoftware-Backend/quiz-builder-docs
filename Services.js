/**
 * Services.js - Logic
 */

class QuizValidator {
  /**
   * Validates the Question Model.
   * @param {{quizType: 'MCQ'|'ESSAY', questions: Array}} parseResult
   * @returns {{isValid: boolean, errors: Array, warnings: Array}}
   */
  static validate(parseResult) {
    const errors = [];
    const warnings = [];

    const quizType = parseResult && parseResult.quizType;
    const questions = parseResult && parseResult.questions;

    if (quizType !== 'MCQ' && quizType !== 'ESSAY') {
      errors.push({ message: 'Missing or invalid quiz type header. The first non-empty line must be [BEGIN#MCQ] or [BEGIN#ESSAY].', line: 0 });
      return { isValid: false, errors, warnings };
    }

    if (!questions || questions.length === 0) {
      errors.push({ message: 'No questions found. Please add at least one [QUESTION] / [QUESTION#n] block.', line: 0 });
      return { isValid: false, errors, warnings };
    }

    questions.forEach(q => {
      // 1. Check Content
      if (!q.content || q.content.trim() === '') {
        errors.push({ message: `Question ${q.num}: Content is empty.`, line: q.line });
      }

      if (quizType === 'ESSAY') {
        if (q.hasOptionsTag || (q.options && q.options.length > 0)) {
          const line = q.optionsTagLine || (q.options && q.options[0] && q.options[0].line) || q.line;
          errors.push({ message: `Question ${q.num}: [OPTIONS] is not allowed for ESSAY. Remove [OPTIONS] and the option list.`, line: line });
        }
        return;
      }

      // MCQ rules
      if (!q.hasOptionsTag) {
        errors.push({ message: `Question ${q.num}: Missing [OPTIONS]. MCQ questions must have an [OPTIONS] section.`, line: q.line });
      }

      // 2. Check Options Count
      if (!q.options || q.options.length < 2) {
        const count = q.options ? q.options.length : 0;
        errors.push({ message: `Question ${q.num}: Needs at least 2 options (found ${count}).`, line: q.optionsTagLine || q.line });
      }

      // 3. Check Correct Answer (Strict: Exactly one *)
      const correctCount = q.options.filter(o => o.isCorrect).length;
      if (correctCount === 0) {
        errors.push({ message: `Question ${q.num}: No correct answer marked. Add '<>' to one option.`, line: q.line });
      } else if (correctCount > 1) {
        errors.push({ message: `Question ${q.num}: Multiple correct answers found (${correctCount}). Only one allowed.`, line: q.line });
      }

      // 4. Duplicate options (compare full line text)
      const seen = {};
      q.options.forEach((o, idx) => {
        const key = QuizValidator._normalizeOptionLine_(o);
        if (!key) return;
        if (Object.prototype.hasOwnProperty.call(seen, key)) {
          errors.push({ message: `Question ${q.num}: Duplicate options found: "${key}". Options must be unique.`, line: o.line || q.line });
        } else {
          seen[key] = idx;
        }
      });

      // 5. Check Empty Options
      q.options.forEach((o, idx) => {
          if (!o.text || o.text.trim() === '') {
              warnings.push({ message: `Question ${q.num}, Option ${idx+1}: Option text is empty.`, line: o.line });
          }
      });
    });

    return {
      isValid: errors.length === 0,
      errors: errors,
      warnings: warnings
    };
  }

  static _normalizeOptionLine_(option) {
    // Full-line comparison: use the author's visible line text, ignoring styling.
    // We DO NOT lowercase and we keep punctuation; we only normalize whitespace.
    const base = (option && typeof option.choiceText === 'string')
      ? option.choiceText
      : (option && typeof option.text === 'string')
        ? QuizValidator._stripHtml_(option.text)
        : '';
    return String(base || '').trim().replace(/\s+/g, ' ');
  }

  static _stripHtml_(html) {
    // Minimal HTML-to-text for validation comparisons.
    return String(html || '')
      .replace(/<\s*br\s*\/?\s*>/gi, ' ')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }
}

class QuizExporter {
  /**
   * Generates ZIP blob from parsed data.
   * @param {Object} parseResult - {questions, images}
   * @returns {GoogleAppsScript.Base.Blob}
   */
  static generateZip(parseResult) {
    const blobs = [];

    const exportJson = {
      metadata: {
        generated_at: new Date().toISOString(),
        quiz_type: parseResult.quizType,
        question_count: parseResult.questions.length
      },
      questions: parseResult.questions.map(q => ({
        id: q.id,
        content: q.content,
        descriptions: (typeof q.descriptions === 'string' && q.descriptions.trim() !== '') ? q.descriptions : null,
        options: q.options.map(o => ({
           content: o.text,
           is_correct: o.isCorrect
        }))
      }))
    };

    blobs.push(Utilities.newBlob(JSON.stringify(exportJson, null, 2), 'application/json', 'quiz.json'));

    for (const [filename, blob] of Object.entries(parseResult.images)) {
      blob.setName(`images/${filename}`);
      blobs.push(blob);
    }

    return Utilities.zip(blobs, 'quiz_export.zip');
  }
}
