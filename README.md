# CoolWords

A Google Apps Script project for collecting interesting words I come across and getting one back by email each morning.

## What it does

**Collecting words.** I add a word through a Google Form (word, meaning, context, source, source type). On submit, the script writes it to the top of a `CoolWords` sheet. If the word is already there, it doesn't make a duplicate row. It fills in a missing meaning and adds any new context or source to the existing entry. Each source is colour-coded by type (book, podcast, article, film/TV/game, internet, conversation, other).

**Word of the Day.** A daily time-based trigger picks a random word from the sheet and emails it to the script owner. The email includes:

- everything from the sheet row, with how long ago the word was added
- pronunciation, part of speech, example sentences, synonyms and antonyms from the [Free Dictionary API](https://dictionaryapi.dev/)
- an **AI word coach** section written by Claude: two fresh example sentences, a memory tip, and a quick fill-in-the-blank quiz
- a small challenge to use the word that day

It remembers which words were sent recently (in Script Properties), so words don't repeat until the whole list has gone round.

## How the AI part works

`askWordCoach()` sends the word, the meaning I noted and the context I found it in to the Claude API (`UrlFetchApp`, no extra hosting). It asks for structured JSON output against a fixed schema, so the reply always has the same shape and can go straight into the email template.

- The API key lives in **Script Properties**, never in the code.
- If the key is missing, the request fails or the model doesn't finish, the function returns `null` and the email goes out without the AI section. The daily email never breaks because of the AI call.
- Quiz options are shuffled in code, so the right answer isn't always in the same place.
- It runs once a day at low effort, so the cost is very small.

## Built with

- Google Apps Script (V8 runtime)
- SpreadsheetApp, MailApp, UrlFetchApp, PropertiesService, ScriptApp triggers
- Free Dictionary API
- Claude API (Anthropic) with structured outputs
- [clasp](https://github.com/google/clasp) for local development and version control

## Setup

1. Create a Google Form with the questions **Word**, **Meaning**, **Context**, **Source** and **Source Type**, and link it to a Google Sheet.
2. Add a sheet tab called `CoolWords` with the headers: Word, Meaning, Context, Source, Date added.
3. Open **Extensions > Apps Script** from the sheet and add `Code.js`.
4. Add an installable trigger: `onFormSubmit`, *From spreadsheet*, *On form submit*.
5. Optional, for the AI section: get an API key from [console.anthropic.com](https://console.anthropic.com) and add it under **Project Settings > Script Properties** as `ANTHROPIC_API_KEY`. Run `testWordCoach` to check it works.
6. Run `createWordOfTheDayTrigger` once to schedule the daily email (default 8:00 in the script's time zone; change `WOTD_SEND_HOUR` to adjust).

## Working on it locally

```bash
npm install -g @google/clasp
clasp login
clasp pull   # get the latest from Apps Script
clasp push   # send local changes back
```
