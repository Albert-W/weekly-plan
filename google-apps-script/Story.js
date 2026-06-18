/**
 * Meal-time habit stories (Google Sheets edition).
 *
 * Three times a day (CONFIG.STORY.MEAL_HOURS — breakfast, lunch, dinner)
 * a random habit is picked and Google Gemini writes a short, original
 * story whose heart is that habit. The story is insightful, meaningful,
 * joyful and interesting; it ends with the narrator realizing how much
 * the habit matters and resolving to work on it today. The result is
 * sent to Telegram (plain text) to keep renewing the user's drive to
 * build the habit.
 *
 * Variety: a rotating set of narrative "angles" (CONFIG.STORY.ANGLES) is
 * used, and the recently-used angles are remembered and fed back to the
 * model as things to avoid — so stories stay fresh even when the same
 * habit is picked again.
 *
 * Driven by the time-based `mealStory` trigger (see Triggers.js), deduped
 * per meal slot per day. Requires a Gemini API key in Script Properties
 * (set via setUpGemini) and the script.external_request OAuth scope.
 */

// ----------------------------------------------------------------------
// Gemini credentials
// ----------------------------------------------------------------------

/**
 * Read the Gemini API key from Script Properties.
 * @returns {string}
 */
function geminiApiKey_() {
  return (
    PropertiesService.getScriptProperties().getProperty(CONFIG.GEMINI.API_KEY_PROP) || ''
  ).trim();
}

/**
 * True when a Gemini API key is present.
 * @returns {boolean}
 */
function geminiConfigured_() {
  return !!geminiApiKey_();
}

/**
 * Persist the Gemini API key to Script Properties.
 * @param {string} key
 */
function setGeminiApiKey_(key) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.GEMINI.API_KEY_PROP, String(key));
}

/**
 * Menu-callable setup: prompt for the Gemini API key and store it.
 */
function setUpGemini() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Set up Gemini',
    'Paste your Google Gemini API key (from aistudio.google.com → Get API key):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const key = resp.getResponseText().trim();
  if (!key) {
    ui.alert('API key is required.');
    return;
  }
  setGeminiApiKey_(key);
  ui.alert('Gemini key saved. Use "Send meal story now" to test it.');
}

// ----------------------------------------------------------------------
// Habit pick + angle rotation
// ----------------------------------------------------------------------

/**
 * Pick a random habit name from the Habits sheet, or '' when none.
 * Reuses the quest habit candidate scan (non-empty habit names).
 * @returns {string}
 */
function pickRandomHabit_() {
  const habits = questHabitCandidates_();
  if (!habits.length) return '';
  return habits[Math.floor(Math.random() * habits.length)].name;
}

/**
 * The recently-used story angles (newest first), for avoidance.
 * @returns {string[]}
 */
function getRecentAngles_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty(
      CONFIG.STORY.RECENT_ANGLES_PROP
    );
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

/**
 * Pick a narrative angle, preferring ones not used recently.
 * @returns {string}
 */
function pickStoryAngle_() {
  const all = CONFIG.STORY.ANGLES;
  if (!all || !all.length) return 'a fresh, original framing';
  const recent = getRecentAngles_();
  const fresh = all.filter((a) => recent.indexOf(a) === -1);
  const pool = fresh.length ? fresh : all;
  return pool[Math.floor(Math.random() * pool.length)];
}

/**
 * Remember an angle as recently used (capped at CONFIG.STORY.ANGLE_HISTORY).
 * @param {string} angle
 */
function rememberAngle_(angle) {
  let list = getRecentAngles_();
  list.unshift(angle);
  if (list.length > CONFIG.STORY.ANGLE_HISTORY) {
    list = list.slice(0, CONFIG.STORY.ANGLE_HISTORY);
  }
  PropertiesService.getDocumentProperties().setProperty(
    CONFIG.STORY.RECENT_ANGLES_PROP,
    JSON.stringify(list)
  );
}

// ----------------------------------------------------------------------
// Gemini story generation
// ----------------------------------------------------------------------

/**
 * Build the story prompt for a habit + chosen angle, instructing the
 * model to avoid recently used angles so stories never repeat.
 * @param {string} habitName
 * @param {string} angle
 * @returns {string}
 */
function buildStoryPrompt_(habitName, angle) {
  const recent = getRecentAngles_();
  const avoid = recent.length ? recent.join('; ') : '(none yet)';
  return [
    'You are a gifted storyteller and a warm, encouraging coach.',
    '',
    'Write a short, self-contained story whose heart is the habit of "' + habitName + '".',
    'It may be a true story or fiction. Make it insightful, meaningful, joyful, and genuinely interesting.',
    'Tell it in the first person ("I"). By the end, I clearly realize how important the habit of "' +
      habitName +
      '" is to the life I want, and I resolve to start working on it right now, today.',
    '',
    'Use this fresh angle for the framing: ' + angle + '.',
    'Do NOT reuse any of these recently used angles, and do not retell a story I have likely already seen: ' +
      avoid +
      '.',
    'Even if this habit has come up before, choose a different situation, character, setting, and insight.',
    '',
    'IMPORTANT: Write the entire story in Simplified Chinese (简体中文). Use natural, fluent Chinese — do not include any English except, if unavoidable, the habit name itself.',
    'Style: vivid but concise, short paragraphs, no markdown, no headings, no hashtags, no emojis.',
    'End with a single uplifting sentence on its own line that nudges me to act today.',
  ].join('\n');
}

/**
 * Pull the generated text out of a Gemini generateContent response body.
 * @param {string} body raw JSON
 * @returns {string}
 */
function extractGeminiText_(body) {
  try {
    const data = JSON.parse(body);
    const cand = data && data.candidates && data.candidates[0];
    if (!cand || !cand.content || !cand.content.parts) return '';
    return cand.content.parts
      .map((p) => (p && p.text ? p.text : ''))
      .join('')
      .trim();
  } catch (e) {
    return '';
  }
}

/**
 * Generate a habit story via the Gemini generateContent API.
 * @param {string} habitName
 * @param {string} angle
 * @returns {string} the story text
 * @throws {Error} when the key is missing or the API fails / returns empty
 */
function generateStory_(habitName, angle) {
  const key = geminiApiKey_();
  if (!key) throw new Error('Gemini API key not set — run "Set up Gemini…" first.');

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(CONFIG.GEMINI.MODEL) +
    ':generateContent?key=' +
    encodeURIComponent(key);

  const payload = {
    contents: [{ parts: [{ text: buildStoryPrompt_(habitName, angle) }] }],
    generationConfig: {
      temperature: CONFIG.GEMINI.TEMPERATURE,
      maxOutputTokens: CONFIG.GEMINI.MAX_OUTPUT_TOKENS,
    },
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Gemini API error ' + code + ': ' + body);
  }
  const text = extractGeminiText_(body);
  if (!text) throw new Error('Gemini returned no story text: ' + body);
  return text;
}

// ----------------------------------------------------------------------
// Send
// ----------------------------------------------------------------------

/**
 * Pick a habit, generate a story, and send it to Telegram. Enable-gated
 * and deduped per meal slot per day (each hour in CONFIG.STORY.MEAL_HOURS
 * sends once). `force` bypasses both for manual previews.
 * @param {boolean} [force]
 * @returns {string} status message
 */
function sendMealStory_(force) {
  if (!force && !CONFIG.STORY.ENABLED) return 'Meal story disabled (Config.STORY.ENABLED).';
  if (!geminiConfigured_()) return 'Gemini not configured — run "Set up Gemini…" first.';
  if (!telegramConfigured_()) return 'Telegram not configured — run "Set up Telegram…" first.';

  const props = PropertiesService.getDocumentProperties();
  const today = formatDateYYYYMMDD(new Date());
  const hour = new Date().getHours();

  // Per-slot dedup: { date, hours:[...] }, reset when the date rolls over.
  let stamp = { date: today, hours: [] };
  try {
    const parsed = JSON.parse(props.getProperty(CONFIG.STORY.LAST_SENT_PROP) || '{}');
    if (parsed && parsed.date === today && Array.isArray(parsed.hours)) stamp = parsed;
  } catch (e) {
    // keep default
  }
  if (!force && stamp.hours.indexOf(hour) !== -1) {
    return 'Meal story already sent for this slot today.';
  }

  const habit = pickRandomHabit_();
  if (!habit) return 'No habits available to feature — add some on the Habits sheet.';

  const angle = pickStoryAngle_();
  let story = generateStory_(habit, angle);
  rememberAngle_(angle);

  // Telegram's hard message limit is 4096 chars; keep headroom for the title.
  const message = '📖 关于「' + habit + '」的小故事\n\n' + story;
  const safe = message.length > 4000 ? message.slice(0, 3999) + '…' : message;

  // Plain text (no parse mode) — LLM prose may contain characters that
  // would break Telegram's HTML parser.
  sendTelegramMessage_(safe, null);

  if (stamp.hours.indexOf(hour) === -1) stamp.hours.push(hour);
  props.setProperty(CONFIG.STORY.LAST_SENT_PROP, JSON.stringify(stamp));
  return 'Meal story sent (habit: ' + habit + ').';
}

/**
 * Menu/sidebar-callable manual send (forces past the dedup + enable flag)
 * so the user can preview a story any time.
 * @returns {string} status message
 */
function sendMealStoryNowFromUI() {
  let msg;
  try {
    msg = sendMealStory_(true);
  } catch (e) {
    msg = 'Meal story failed: ' + (e && e.message ? e.message : e);
  }
  toast_(msg, 'Weekly Plan');
  return msg;
}
