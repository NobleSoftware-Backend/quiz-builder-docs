# Project Quiz Builder

This Google Apps Script project allows you to build quizzes directly in Google Docs and export them to JSON format. It features robust validation, LaTeX equation support (MathJax), and a live visual preview.

## Features

- **Strict Validation:** Ensures every question has valid options and exactly one correct answer.
- **Rich Text & Images:** Extracts formatting, lists, tables, and images.
- **Equation Support:** Automatically converts Google Docs equations into standard LaTeX for the web.
- **Live Preview:** Preview how your questions will look rendered with MathJax before exporting.
- **ZIP Export:** Downloads a clean package containing `quiz.json` and all extracted images/assets.

## Installation

This script is designed to be deployed as a Google Docs Editor Add-on.

1.  **Deploy**: Use `clasp push` to upload code.
2.  **Test**: In the Apps Script Editor, go to **Deploy > Test Deployments**.
3.  **Select**: Choose "Editor Add-on", select a document, and click "Execute".

## Format Guide

The Quiz Builder relies on strict formatting tags in the document:

0. **[BEGIN#MCQ]** or **[BEGIN#ESSAY]**: Required quiz type header.
  - Must be the **first non-empty line** in the document.
  - Example: `[BEGIN#MCQ]`
  - If you need to change type, reset the document and set the type again.

1.  **[QUESTION#n]**: Marks the start of a new question (recommended).
  - Example: `[QUESTION#7]`
  - The `#n` is a human label only; the exporter still uses the document order internally.
  - Invalid tags like `[QUESTION#]` are rejected.
2.  **[QUESTION]**: Legacy format (still supported) for older documents.
3.  **Question Text**: The text immediately following the question tag (until the options).
4.  **[DESCRIPTIONS]** (optional): Extra explanation/notes for the question, exported as HTML (or `null` if omitted).
  - Must appear **after** `[QUESTION...]` and **before** `[OPTIONS]`.
5.  **[OPTIONS]**: Marks the start of the options list.
6.  **List**: The options must be a list (e.g., A, B, C).

### Optional Descriptions

If you want an optional description/explanation per question, insert a `[DESCRIPTIONS]` section between the question content and `[OPTIONS]`.

```text
[BEGIN#MCQ]

[QUESTION#1]
What is the capital of Indonesia?

[DESCRIPTIONS]
This is a geography question.

[OPTIONS]
A. <> Jakarta
B. Bandung
```

### MCQ vs ESSAY Rules

- **MCQ** (`[BEGIN#MCQ]`)
  - `[OPTIONS]` is mandatory for every question.
  - Options must be unique (no duplicate option lines like "blue" and "blue").
  - Options must not contain images (styling/formatting is OK).

- **ESSAY** (`[BEGIN#ESSAY]`)
  - `[OPTIONS]` is not allowed anywhere.

### Marking the Correct Answer

To mark an option as correct, place `<>` at the very beginning of the list item line.

**Example:**

```text
[BEGIN#MCQ]

[QUESTION#1]
What is the capital of Indonesia?

[OPTIONS]
A. <> Jakarta
B. Bandung
C. Surabaya
D. Bali
```

### Renumbering Labels

If you insert/reorder questions and want clean sequential labels again:

- Use the menu: `Quiz Builder > Renumber Question Labels`
- This rewrites all question tags in document order to `[QUESTION#1]`, `[QUESTION#2]`, ...

### Math Equations

You can use the built-in Google Docs equation editor (Insert > Equation).
- The add-on will automatically converting them to LaTeX (e.g., `\frac{a}{b}`).
- Click **"Preview Question"** in the sidebar to verify the rendering.

## Usage

1.  **Open Sidebar**: Go to `Quiz Builder > Open Sidebar`.
2.  **Wrap Question**: Select your question text and options list, then click `Wrap as Question Block`.
3.  **Preview**: Place cursor inside a question block and click `Preview Question` to see the rendered HTML/MathJax.
4.  **Validate**: Click `Check Format` to scan the document for errors.
5.  **Export**: Click `Export Quiz` to generate a ZIP file.

## Output Format (quiz.json)

```json
{
  "metadata": { ... },
  "questions": [
    {
      "id": "q1",
      "content": "<p>Question text...</p>",
      "descriptions": "<p>Optional description...</p>",
      "options": [
        {
          "content": "Jakarta",
          "is_correct": true
        },
        {
          "content": "Bandung",
          "is_correct": false
        }
      ]
    }
  ]
}
```
