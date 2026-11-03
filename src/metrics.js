/*
 * Voice Drift Tracker: pure text metrics.
 * No DOM and no Office.js here, so the same file runs in Node tests and in the task pane.
 */
(function (root) {
  'use strict';

  var HEDGES = ['might', 'may', 'perhaps', 'seem', 'seems', 'seemed', 'possibly',
    'probably', 'arguably', 'apparently', 'somewhat', 'suggests'];
  var PERSONAL = ['i', 'me', 'my', 'mine', 'we', 'us', 'our', 'ours', 'you', 'your', 'yours'];
  var THIRD = ['he', 'him', 'his', 'she', 'her', 'hers', 'they', 'them', 'their', 'theirs', 'it', 'its'];

  var METRICS = [
    { key: 'sentenceLength', label: 'Sentence length', unit: 'Average words per sentence', floor: 1.5 },
    { key: 'wordLength', label: 'Word length', unit: 'Average letters per word (vocabulary difficulty)', floor: 0.3 },
    { key: 'hedging', label: 'Hedging', unit: 'Hedge words per 100 words (might, perhaps, seems)', floor: 1 },
    { key: 'personal', label: 'Personal voice', unit: 'Share of I/we/you vs. he/she/they (%)', floor: 10 },
    { key: 'punctuation', label: 'Punctuation mood', unit: '?, !, and dashes per 100 words', floor: 1 }
  ];

  function toSet(list) {
    var s = {};
    list.forEach(function (w) { s[w] = true; });
    return s;
  }
  var HEDGE_SET = toSet(HEDGES);
  var PERSONAL_SET = toSet(PERSONAL);
  var THIRD_SET = toSet(THIRD);

  function words(text) {
    var matches = text.match(/[A-Za-z]+(?:['’-][A-Za-z]+)*/g) || [];
    return matches.map(function (w) { return w.toLowerCase().replace(/’/g, "'"); });
  }

  function sentences(text) {
    return text
      .replace(/\b(Mr|Mrs|Ms|Dr|St|vs|etc)\./g, '$1\u0000')
      // Mark sentence ends, then split. Avoids regex lookbehind, which older Safari
      // (and so Word on older macOS) cannot parse.
      .replace(/([.!?])\s+/g, '$1\u0001')
      .split('\u0001')
      .map(function (s) { return s.replace(/\u0000/g, '.').trim(); })
      .filter(function (s) { return /[A-Za-z0-9]/.test(s); });
  }

  function avg(list) {
    return list.reduce(function (a, b) { return a + b; }, 0) / list.length;
  }

  // Returns metric values for one paragraph, or null if it has no words.
  function analyzeParagraph(text) {
    var w = words(text);
    var n = w.length;
    if (n === 0) return null;

    var sentenceCount = Math.max(1, sentences(text).length);
    var letters = w.reduce(function (a, x) { return a + x.replace(/[^a-z]/g, '').length; }, 0);
    var hedges = w.filter(function (x) { return HEDGE_SET[x]; }).length;
    var personal = w.filter(function (x) { return PERSONAL_SET[x]; }).length;
    var third = w.filter(function (x) { return THIRD_SET[x]; }).length;
    var marks = (text.match(/[?!]|—|–|--/g) || []).length;

    return {
      words: n,
      sentenceLength: n / sentenceCount,
      wordLength: letters / n,
      hedging: (hedges * 100) / n,
      personal: (personal + third) ? (personal / (personal + third)) * 100 : null,
      punctuation: (marks * 100) / n
    };
  }

  // Keeps each row's index so the task pane can jump back to that paragraph.
  function analyzeDocument(paragraphs) {
    var rows = [];
    paragraphs.forEach(function (text, index) {
      var m = analyzeParagraph(text);
      if (m) rows.push(Object.assign({ index: index }, m));
    });
    return rows;
  }

  // Target band per metric: mean +/- half-width. Half-width is the spread of the
  // target paragraphs, but never narrower than a per-metric floor.
  function buildProfile(paragraphs) {
    var rows = analyzeDocument(paragraphs);
    if (!rows.length) return null;
    var profile = {};
    METRICS.forEach(function (m) {
      var vals = rows.map(function (r) { return r[m.key]; }).filter(function (v) { return v != null; });
      if (!vals.length) { profile[m.key] = null; return; }
      var mean = avg(vals);
      var sd = Math.sqrt(avg(vals.map(function (v) { return (v - mean) * (v - mean); })));
      profile[m.key] = { mean: mean, half: Math.max(sd, 0.1 * Math.abs(mean), m.floor) };
    });
    return profile;
  }

  // 'in' = within the band, 'near' = within twice the band, 'out' = further away.
  function classify(value, p) {
    if (value == null || !p) return 'none';
    var d = Math.abs(value - p.mean) / p.half;
    if (d <= 1) return 'in';
    if (d <= 2) return 'near';
    return 'out';
  }

  var api = {
    METRICS: METRICS,
    words: words,
    sentences: sentences,
    analyzeParagraph: analyzeParagraph,
    analyzeDocument: analyzeDocument,
    buildProfile: buildProfile,
    classify: classify
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    root.VoiceDrift = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
