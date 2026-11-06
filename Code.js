function onFormSubmit(e) {
  const vocabulary = e.source.getSheetByName("CoolWords");
  const answers = e.namedValues;

  const getAnswer = (...names) => {
    for (const name of names) {
      if (answers[name]?.[0]) return answers[name][0].trim();
    }
    return "";
  };

  const word = getAnswer("Word");
  const definition = getAnswer("Meaning");
  const context = getAnswer("Context");
  const source = getAnswer("Source");
  const sourceType = getAnswer("Source Type");
  const date = new Date();

  if (!word) return;

  const lastRow = vocabulary.getLastRow();
  const rows =
    lastRow > 1
      ? vocabulary.getRange(2, 1, lastRow - 1, 5).getValues()
      : [];

  const existingIndex = rows.findIndex(
    row => String(row[0]).trim().toLowerCase() === word.toLowerCase()
  );

if (existingIndex === -1) {
  vocabulary.insertRowBefore(2);

  vocabulary
    .getRange(2, 1, 1, 5)
    .setValues([[
      word,
      definition,
      context,
      source,
      date
    ]]);

  vocabulary.getRange(2, 5).setNumberFormat("d-M-yy");

  const colour = getSourceColour(sourceType);

  if (colour) {
    vocabulary.getRange(2, 4).setBackground(colour);
  }

  return;
}

  const rowNumber = existingIndex + 2;
  const existing = rows[existingIndex];

  if (!existing[1] && definition) {
    vocabulary.getRange(rowNumber, 2).setValue(definition);
  }

  appendUnique(vocabulary.getRange(rowNumber, 3), context);
  appendUnique(vocabulary.getRange(rowNumber, 4), source);
}

function appendUnique(cell, newText) {
  if (!newText) return;

  const currentText = String(cell.getValue()).trim();
  const existingItems = currentText
    ? currentText.split("\n").map(item => item.trim())
    : [];

  if (!existingItems.includes(newText)) {
    cell.setValue([...existingItems, newText].join("\n"));
  }
}

function getSourceColour(sourceType) {
  const normalized = String(sourceType)
    .trim()
    .toLowerCase()
    .replace(/\s*\/\s*/g, "/");

  const colours = {
    "book": "#d9ead3",
    "podcast": "#fff2cc",
    "article/news": "#d0e8e6",
    "film/tv show/video game": "#eadcf8",
    "internet": "#fce5cd",
    "conversation/everyday life": "#f4cccc",
    "other": "#d9d9d9"
  };

  return colours[normalized] || "#ffffff";
}

// ---------- Word of the Day ----------

const WOTD_SEND_HOUR = 8; // 24h clock, in the script's time zone
const WOTD_SENT_KEY = "WOTD_RECENTLY_SENT";
const WOTD_DEFAULT_LABELS = ["Word", "Meaning", "Context", "Source", "Date added"];

function sendWordOfTheDay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CoolWords");
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const sourceColours = sheet.getRange(2, 4, lastRow - 1, 1).getBackgrounds();

  const entries = rows
    .map((row, i) => ({ row, sourceColour: sourceColours[i][0] }))
    .filter(entry => String(entry.row[0]).trim());

  if (!entries.length) return;

  const entry = pickWordOfTheDay(entries);
  const word = String(entry.row[0]).trim();
  const dictionary = lookupWord(word);
  const coach = askWordCoach(word, String(entry.row[1] || "").trim(), String(entry.row[2] || "").trim());

  MailApp.sendEmail({
    // Sends to whoever owns the script, so no address is hard-coded
    to: Session.getEffectiveUser().getEmail(),
    subject: "Cool Word of the Day",
    name: "CoolWords",
    htmlBody: buildWordOfTheDayHtml(entry, headers, dictionary, coach),
    body: buildWordOfTheDayText(entry, headers, dictionary, coach)
  });
}

function pickWordOfTheDay(entries) {
  const props = PropertiesService.getScriptProperties();
  const key = entry => String(entry.row[0]).trim().toLowerCase();
  let recent = JSON.parse(props.getProperty(WOTD_SENT_KEY) || "[]");

  let pool = entries.filter(entry => !recent.includes(key(entry)));

  if (!pool.length) {
    recent = [];
    pool = entries;
  }

  const entry = pool[Math.floor(Math.random() * pool.length)];
  recent.push(key(entry));

  // Script properties max out around 9KB, so forget the oldest words if needed
  while (JSON.stringify(recent).length > 8000) recent.shift();

  props.setProperty(WOTD_SENT_KEY, JSON.stringify(recent));
  return entry;
}

function lookupWord(word) {
  try {
    const response = UrlFetchApp.fetch(
      "https://api.dictionaryapi.dev/api/v2/entries/en/" + encodeURIComponent(word.toLowerCase()),
      { muteHttpExceptions: true }
    );

    if (response.getResponseCode() !== 200) return null;

    const results = JSON.parse(response.getContentText());
    const phonetic = results
      .flatMap(result => [result.phonetic, ...(result.phonetics || []).map(p => p.text)])
      .find(Boolean) || "";

    const partsOfSpeech = new Set();
    const examples = [];
    const synonyms = new Set();
    const antonyms = new Set();

    results.forEach(result => {
      (result.meanings || []).forEach(meaning => {
        partsOfSpeech.add(meaning.partOfSpeech);
        (meaning.synonyms || []).forEach(s => synonyms.add(s));
        (meaning.antonyms || []).forEach(a => antonyms.add(a));

        (meaning.definitions || []).forEach(def => {
          if (def.example) {
            examples.push({
              partOfSpeech: meaning.partOfSpeech,
              definition: def.definition,
              example: def.example
            });
          }
          (def.synonyms || []).forEach(s => synonyms.add(s));
          (def.antonyms || []).forEach(a => antonyms.add(a));
        });
      });
    });

    return {
      phonetic,
      partsOfSpeech: [...partsOfSpeech].filter(Boolean),
      examples: examples.slice(0, 3),
      synonyms: [...synonyms].slice(0, 8),
      antonyms: [...antonyms].slice(0, 5)
    };
  } catch (err) {
    console.warn("Dictionary lookup failed for " + word + ": " + err);
    return null;
  }
}

// ---------- AI word coach (Claude) ----------
//
// Set the API key once in Project Settings > Script Properties as
// ANTHROPIC_API_KEY. It is never stored in the code. If the key is missing
// or the call fails, the email still goes out without the AI section.

const CLAUDE_MODEL = "claude-opus-5-5";

const WORD_COACH_SCHEMA = {
  type: "object",
  properties: {
    sentences: {
      type: "array",
      items: { type: "string" },
      description: "Two fresh, natural example sentences using the word"
    },
    memoryTip: {
      type: "string",
      description: "One short, memorable way to remember the meaning"
    },
    quizSentence: {
      type: "string",
      description: "A sentence with the word replaced by ____"
    },
    quizOptions: {
      type: "array",
      items: { type: "string" },
      description: "Four options, one of which is the word"
    }
  },
  required: ["sentences", "memoryTip", "quizSentence", "quizOptions"],
  additionalProperties: false
};

function askWordCoach(word, meaning, context) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) return null;

  const prompt = [
    `Word: ${word}`,
    meaning ? `Meaning I noted: ${meaning}` : "",
    context ? `Where I came across it:\n${context}` : "",
    "",
    "I'm building my vocabulary and this is today's word. Write:",
    "- two new example sentences in everyday British English, different in tone from where I found it",
    "- one short memory tip for the meaning (an image, a word link, or a contrast)",
    "- a fill-in-the-blank quiz sentence with four single-word options: the word itself and three plausible but wrong ones",
    "Keep everything short enough to read in a morning email."
  ].filter(Boolean).join("\n");

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      contentType: "application/json",
      muteHttpExceptions: true,
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2026-10-21",
        "anthropic-beta": "server-side-fallback-2026-07-01"
      },
      payload: JSON.stringify({
        model: CLAUDE_MODEL,
        max_tokens: 4000,
        fallbacks: "default",
        output_config: {
          effort: "low",
          format: { type: "json_schema", schema: WORD_COACH_SCHEMA }
        },
        messages: [{ role: "user", content: prompt }]
      })
    });

    if (response.getResponseCode() !== 200) {
      console.warn("Word coach HTTP " + response.getResponseCode() + ": " + response.getContentText());
      return null;
    }

    const message = JSON.parse(response.getContentText());
    if (message.stop_reason !== "end_turn") {
      console.warn("Word coach stopped early: " + message.stop_reason);
      return null;
    }

    const textBlock = (message.content || []).find(block => block.type === "text");
    if (!textBlock) return null;

    const coach = JSON.parse(textBlock.text);
    coach.quizOptions = shuffle(coach.quizOptions);
    return coach;
  } catch (err) {
    console.warn("Word coach failed for " + word + ": " + err);
    return null;
  }
}

function shuffle(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

// Run by hand from the Apps Script editor to check the AI section without sending an email
function testWordCoach() {
  console.log(JSON.stringify(askWordCoach("sonder", "the realisation that every passer-by has a life as vivid as your own", ""), null, 2));
}

function getColumnLabel(headers, i) {
  return String(headers[i] || WOTD_DEFAULT_LABELS[i] || `Column ${i + 1}`).trim();
}

function buildWordOfTheDayHtml(entry, headers, dictionary, coach) {
  const word = String(entry.row[0]).trim();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE d MMMM");
  const lines = value => String(value || "").split("\n").map(s => s.trim()).filter(Boolean);

  const label = text =>
    `<div style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:#8a8a8a;font-weight:bold;margin-bottom:10px;">${text}</div>`;

  const section = (title, content) => content ? `
    <tr><td style="padding:24px 32px 0;">
      ${label(title)}
      ${content}
    </td></tr>` : "";

  const chip = (text, colour) =>
    `<span style="display:inline-block;padding:4px 12px;margin:0 6px 6px 0;border-radius:14px;border:1px solid #dddddd;background:${colour};font-family:Arial,sans-serif;font-size:13px;color:#333333;">${escapeHtml(text)}</span>`;

  // Columns: 0 Word, 1 Meaning, 2 Context, 3 Source, 4 Date added, then any extras
  const formatCell = (value, i) => {
    if (!String(value).trim()) return `<span style="color:#bbbbbb;">—</span>`;
    if (value instanceof Date) return escapeHtml(describeDateAdded(value));
    if (i === 0) return `<b style="font-size:18px;">${escapeHtml(value)}</b>`;
    if (i === 2) {
      return lines(value).map(c => `
        <div style="border-left:3px solid #3d3a8c;padding:6px 12px;margin:0 0 6px;background:#f6f5ff;font-style:italic;">“${escapeHtml(c)}”</div>`).join("");
    }
    if (i === 3) return lines(value).map(s => chip(s, entry.sourceColour || "#eeeeee")).join("");
    return escapeWithBreaks(value);
  };

  const detailRows = entry.row.map((value, i) => `
    <tr>
      <td style="padding:12px 16px 12px 0;vertical-align:top;width:100px;border-top:1px solid #eeeeee;font-family:Arial,sans-serif;font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#8a8a8a;font-weight:bold;">${escapeHtml(getColumnLabel(headers, i))}</td>
      <td style="padding:12px 0;vertical-align:top;border-top:1px solid #eeeeee;font-size:16px;line-height:1.5;color:#222222;">${formatCell(value, i)}</td>
    </tr>`).join("");

  const detailsHtml = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${detailRows}</table>`;

  const subtitle = dictionary
    ? [dictionary.phonetic, dictionary.partsOfSpeech.join(" · ")].filter(Boolean).join("   ")
    : "";

  const examplesHtml = dictionary?.examples.length
    ? dictionary.examples.map(ex => `
      <div style="margin-bottom:14px;">
        <div style="font-size:16px;line-height:1.5;color:#222222;">“${escapeHtml(ex.example)}”</div>
        <div style="font-family:Arial,sans-serif;font-size:12px;color:#888888;margin-top:3px;"><b>${escapeHtml(ex.partOfSpeech)}</b> — ${escapeHtml(ex.definition)}</div>
      </div>`).join("")
    : "";

  const synonymsHtml = dictionary?.synonyms.length
    ? dictionary.synonyms.map(s => chip(s, "#eef7ee")).join("")
    : "";

  const antonymsHtml = dictionary?.antonyms.length
    ? dictionary.antonyms.map(a => chip(a, "#fbeeee")).join("")
    : "";

  const coachHtml = coach ? `
    ${coach.sentences.map(s => `
      <div style="border-left:3px solid #3d3a8c;padding:6px 12px;margin:0 0 8px;font-size:16px;line-height:1.5;color:#222222;">${escapeHtml(s)}</div>`).join("")}
    <div style="background:#f6f5ff;border-radius:8px;padding:12px 16px;margin:12px 0;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#333333;">
      🧠 <b>Remember it:</b> ${escapeHtml(coach.memoryTip)}
    </div>
    <div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#333333;">
      <b>Quick quiz:</b> ${escapeHtml(coach.quizSentence)}
      <div style="margin-top:8px;">${coach.quizOptions.map(o => chip(o, "#ffffff")).join("")}</div>
      <div style="font-size:12px;color:#999999;">(The answer is today's word.)</div>
    </div>` : "";

  const lookupUrl = "https://www.google.com/search?q=define+" + encodeURIComponent(word);

  return `
<div style="background:#f2f2f7;padding:24px 8px;font-family:Georgia,'Times New Roman',serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;">
    <tr><td style="background:#3d3a8c;padding:28px 32px;color:#ffffff;">
      <div style="font-family:Arial,sans-serif;font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#c9c7f2;">Cool Word of the Day · ${escapeHtml(today)}</div>
      <div style="font-size:42px;font-weight:bold;margin-top:8px;color:#ffffff;">${escapeHtml(word)}</div>
      ${subtitle ? `<div style="font-size:16px;margin-top:4px;color:#dcdaf7;">${escapeHtml(subtitle)}</div>` : ""}
    </td></tr>
    ${section("From your CoolWords sheet", detailsHtml)}
    ${section("Your AI word coach", coachHtml)}
    ${section("More ways to use it", examplesHtml)}
    ${section("Similar words", synonymsHtml)}
    ${section("Opposites", antonymsHtml)}
    <tr><td style="padding:24px 32px 28px;">
      <div style="background:#fff8e1;border-radius:8px;padding:16px 18px;font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#5d4a00;">
        💡 <b>Today's challenge:</b> use <b>${escapeHtml(word)}</b> in a conversation, message or sentence before the day is out.
      </div>
      <div style="font-family:Arial,sans-serif;font-size:12px;color:#999999;margin-top:16px;text-align:center;">
        <a href="${lookupUrl}" style="color:#3d3a8c;">Look it up in more detail</a> · Sent from your CoolWords sheet
      </div>
    </td></tr>
  </table>
</div>`;
}


function buildWordOfTheDayText(entry, headers, dictionary, coach) {
  const parts = ["COOL WORD OF THE DAY: " + String(entry.row[0]).trim(), ""];

  entry.row.forEach((value, i) => {
    const text = value instanceof Date ? describeDateAdded(value) : String(value).trim();
    parts.push(getColumnLabel(headers, i) + ":\n" + (text || "—") + "\n");
  });

  if (coach) {
    parts.push("AI word coach:\n" + coach.sentences.map(s => "- " + s).join("\n"));
    parts.push("Remember it: " + coach.memoryTip);
    parts.push("Quick quiz: " + coach.quizSentence + "\nOptions: " + coach.quizOptions.join(" / ") + "\n");
  }
  if (dictionary?.examples.length) {
    parts.push("More examples:\n" + dictionary.examples.map(ex => "- " + ex.example).join("\n"));
  }
  if (dictionary?.synonyms.length) {
    parts.push("\nSimilar words: " + dictionary.synonyms.join(", "));
  }

  return parts.join("\n");
}

function describeDateAdded(value) {
  if (!(value instanceof Date) || isNaN(value)) return value ? String(value) : "";

  const formatted = Utilities.formatDate(value, Session.getScriptTimeZone(), "d MMMM yyyy");
  const days = Math.floor((Date.now() - value.getTime()) / 86400000);
  const ago =
    days <= 0 ? "today" :
    days === 1 ? "yesterday" :
    days < 60 ? `${days} days ago` :
    days < 730 ? `about ${Math.round(days / 30)} months ago` :
    `about ${Math.round(days / 365)} years ago`;

  return `${formatted} (${ago})`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeWithBreaks(value) {
  return escapeHtml(value).replace(/\n/g, "<br>");
}

// Run this once by hand to schedule the daily email
function createWordOfTheDayTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === "sendWordOfTheDay")
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("sendWordOfTheDay")
    .timeBased()
    .everyDays(1)
    .atHour(WOTD_SEND_HOUR)
    .create();
}