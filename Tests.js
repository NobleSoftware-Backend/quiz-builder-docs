/**
 * Tests.js
 * Run the 'runTests' function manually in Apps Script editor to verify logic.
 */
function runTests() {
  console.log('🧪 Starting Quiz Tests...');

  // Mock Data (Simulates what Parser would return)
  const mockModel = [
    {
      num: 1,
      line: 10,
      content: '<p>Valid Question</p>',
      options: [
        { text: 'A', isCorrect: true, line: 11 },
        { text: 'B', isCorrect: false, line: 12 }
      ]
    },
    {
      num: 2,
      line: 20,
      content: '', // Error: Empty content
      options: [
        { text: 'A', isCorrect: false, line: 21 }, // Error: No correct answer
        { text: 'B', isCorrect: false, line: 22 }
      ]
    },
    {
      num: 3,
      line: 30,
      content: 'Multi Correct',
      options: [
        { text: 'A', isCorrect: true, line: 31 }, // Error: Multi correct
        { text: 'B', isCorrect: true, line: 32 }
      ]
    }
  ];

  // Run Validator
  const result = QuizValidator.validate(mockModel);

  // Assertions
  const asserts = [
    { name: 'Should have validation errors', pass: result.isValid === false },
    { name: 'Should find 3 errors', pass: result.errors.length === 3 },
    { name: 'Error 1: Empty Content', pass: result.errors.some(e => e.message.includes('Question 2') && e.message.includes('Content is empty')) },
    { name: 'Error 2: No correct answer', pass: result.errors.some(e => e.message.includes('Question 2') && e.message.includes('No correct answer')) },
    { name: 'Error 3: Multiple correct', pass: result.errors.some(e => e.message.includes('Question 3') && e.message.includes('Multiple correct')) }
  ];

  // Report
  asserts.forEach(a => {
    if (a.pass) console.log(`✅ PASS: ${a.name}`);
    else console.error(`❌ FAIL: ${a.name}`);
  });
  
  console.log('Done.');
}
