/**
 * Services.js - Logic
 */

class QuizValidator {
  /**
   * Validates the Question Model.
   * @param {Array} questions 
   * @returns {{isValid: boolean, errors: Array, warnings: Array}}
   */
  static validate(questions) {
    const errors = [];
    const warnings = [];

    if (!questions || questions.length === 0) {
      errors.push({ message: "No questions found. Please check [QUESTION] / [QUESTION#n] tags.", line: 0 });
      return { isValid: false, errors, warnings };
    }

    questions.forEach(q => {
      // 1. Check Content
      if (!q.content || q.content.trim() === '') {
        errors.push({ message: `Question ${q.num}: Content is empty.`, line: q.line });
      }

      // 2. Check Options Count
      if (q.options.length < 2) {
        errors.push({ message: `Question ${q.num}: Needs at least 2 options (found ${q.options.length}).`, line: q.line });
      }

      // 3. Check Correct Answer (Strict: Exactly one *)
      const correctCount = q.options.filter(o => o.isCorrect).length;
      if (correctCount === 0) {
        errors.push({ message: `Question ${q.num}: No correct answer marked. Add '<>' to one option.`, line: q.line });
      } else if (correctCount > 1) {
        errors.push({ message: `Question ${q.num}: Multiple correct answers found (${correctCount}). Only one allowed.`, line: q.line });
      }
      
      // 4. Check Empty Options
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
        question_count: parseResult.questions.length
      },
      questions: parseResult.questions.map(q => ({
        id: q.id,
        content: q.content,
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
