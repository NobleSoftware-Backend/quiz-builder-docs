/**
 * Tests.js
 * Run the 'runTests' function manually in Apps Script editor to verify logic.
 */
function runTests() {
  console.log('🧪 Starting Quiz Tests...');

  // MCQ validator tests
  const mcqParseResult = {
    quizType: 'MCQ',
    questions: [
      {
        num: 1,
        line: 10,
        content: '<p>Valid Question</p>',
        hasOptionsTag: true,
        optionsTagLine: 11,
        options: [
          { text: '<p>blue</p>', choiceText: 'blue', isCorrect: true, line: 12 },
          { text: '<p>red</p>', choiceText: 'red', isCorrect: false, line: 13 }
        ]
      },
      {
        num: 2,
        line: 20,
        content: '', // Error: Empty content
        hasOptionsTag: true,
        optionsTagLine: 21,
        options: [
          { text: '<p>A</p>', choiceText: 'A', isCorrect: false, line: 22 }, // Error: No correct answer
          { text: '<p>B</p>', choiceText: 'B', isCorrect: false, line: 23 }
        ]
      },
      {
        num: 3,
        line: 30,
        content: '<p>Multi Correct</p>',
        hasOptionsTag: true,
        optionsTagLine: 31,
        options: [
          { text: '<p>A</p>', choiceText: 'A', isCorrect: true, line: 32 }, // Error: Multi correct
          { text: '<p>B</p>', choiceText: 'B', isCorrect: true, line: 33 }
        ]
      },
      {
        num: 4,
        line: 40,
        content: '<p>Duplicates</p>',
        hasOptionsTag: true,
        optionsTagLine: 41,
        options: [
          { text: '<p>blue</p>', choiceText: 'blue', isCorrect: true, line: 42 },
          { text: '<p>blue</p>', choiceText: 'blue', isCorrect: false, line: 43 }
        ]
      }
    ]
  };

  const mcqResult = QuizValidator.validate(mcqParseResult);

  const asserts = [
    { name: 'MCQ: Should have validation errors', pass: mcqResult.isValid === false },
    { name: 'MCQ: Empty Content', pass: mcqResult.errors.some(e => e.message.includes('Question 2') && e.message.includes('Content is empty')) },
    { name: 'MCQ: No correct answer', pass: mcqResult.errors.some(e => e.message.includes('Question 2') && e.message.includes('No correct answer')) },
    { name: 'MCQ: Multiple correct', pass: mcqResult.errors.some(e => e.message.includes('Question 3') && e.message.includes('Multiple correct')) },
    { name: 'MCQ: Duplicate options', pass: mcqResult.errors.some(e => e.message.includes('Question 4') && e.message.includes('Duplicate options')) }
  ];

  // ESSAY validator tests
  const essayParseResult = {
    quizType: 'ESSAY',
    questions: [
      {
        num: 1,
        line: 10,
        content: '<p>Essay Question</p>',
        hasOptionsTag: true,
        optionsTagLine: 12,
        options: [
          { text: '<p>Not allowed</p>', choiceText: 'Not allowed', isCorrect: false, line: 13 }
        ]
      }
    ]
  };

  const essayResult = QuizValidator.validate(essayParseResult);
  asserts.push({ name: 'ESSAY: OPTIONS forbidden', pass: essayResult.errors.some(e => e.message.includes('[OPTIONS] is not allowed for ESSAY')) });

  // Report
  asserts.forEach(a => {
    if (a.pass) console.log(`✅ PASS: ${a.name}`);
    else console.error(`❌ FAIL: ${a.name}`);
  });
  
  console.log('Done.');
}
