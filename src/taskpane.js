/* Voice Drift Tracker: task pane UI and Word integration. */
(function () {
  'use strict';

  var M = window.VoiceDrift;
  // True only when the pane is running inside Word. Office.js also loads in a normal
  // browser, so this is set from Office.onReady's host, not from Office existing.
  var inWord = false;
  var state = { metric: 'sentenceLength', paras: [], rows: [], orig: [], profile: null };

  // Plain-language help for each measure. ends[0] is the low side, ends[1] the high side.
  // empty is shown on a row the measure can't score.
  var HELP = {
    sentenceLength: { text: 'Average words per sentence.', ends: ['Shorter sentences', 'Longer sentences'], empty: 'No sentences to measure here' },
    wordLength: { text: 'Average letters per word. Longer words usually sound more formal.', ends: ['Simpler words', 'Longer words'], empty: 'No words to measure here' },
    hedging: { text: 'How often you use words like "might", "perhaps", and "seems". More hedging sounds more tentative.', ends: ['More certain', 'More tentative'], empty: 'No words to measure here' },
    personal: { text: 'How often you use I, we, and you compared with he, she, and they. Higher means more personal.', ends: ['Impersonal', 'Personal'], empty: 'No I, you, he, she, or they here to compare' },
    punctuation: { text: 'How often you use question marks, exclamation points, and dashes. Higher means more expressive.', ends: ['Plain', 'Expressive'], empty: 'No words to measure here' }
  };

  // Highest possible value for measures that are percentages of something.
  var MAX_VALUE = { personal: 100 };

  // Used only when the page runs outside Word (for previewing the layout).
  var DEMO_DOC = [
    "Okay, so here's the thing. Most people think budgeting is boring. I did too, honestly, until I saw my bank statement.",
    'Right? We all say that. But look at what happens next.',
    'The data suggests that small habits might matter more than big plans, perhaps because they are easier to keep.',
    'Consider a student who tracked every coffee for a month. The total was staggering, and the lesson was clear: patterns beat willpower.',
    "So let's talk about it. What would you change first?"
  ];
  var DEMO_TARGET = 'The analysis establishes that consistent saving depends on structural incentives. ' +
    'Individuals who automate transfers demonstrate higher retention rates across multiple longitudinal studies. ' +
    'Consequently, policy designers should prioritize default enrollment mechanisms.';
  var DEMO_ORIGINAL = 'Budgeting is something most people neglect.\n\n' +
    'Small habits are important for financial health and should be considered carefully by individuals.\n\n' +
    'Many people believe that large plans are more effective than small daily changes.\n\n' +
    'A student who tracks spending may find that the total is surprising, which could encourage changes in behavior.\n\n' +
    'It is therefore recommended that readers review their spending on a regular basis.';

  var $ = function (id) { return document.getElementById(id); };

  function metricDef(key) {
    return M.METRICS.filter(function (m) { return m.key === key; })[0];
  }

  // Each non-empty line is a paragraph, matching how Word counts paragraphs.
  function paragraphsFrom(text) {
    return text.split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean);
  }

  // On the website, a paragraph number selects that paragraph in the text box.
  function selectInBox(index) {
    var box = $('source'), text = box.value, paras = paragraphsFrom(text), pos = 0;
    for (var i = 0; i < paras.length; i++) {
      var at = text.indexOf(paras[i], pos);
      if (at < 0) return;
      if (i === index) {
        box.scrollIntoView({ block: 'center', behavior: 'smooth' });
        box.focus();
        box.setSelectionRange(at, at + paras[i].length);
        return;
      }
      pos = at + paras[i].length;
    }
  }

  function setStatus(msg) { $('status').textContent = msg; }

  function debounce(fn, ms) {
    var t;
    return function () {
      clearTimeout(t);
      t = setTimeout(fn, ms);
    };
  }

  // ---------- Word integration ----------

  function readDocument() {
    return Word.run(async function (ctx) {
      var paras = ctx.document.body.paragraphs;
      paras.load('items/text');
      await ctx.sync();
      return paras.items.map(function (p) { return p.text; });
    });
  }

  async function jump(index) {
    if (!inWord) {
      selectInBox(index);
      return;
    }
    try {
      await Word.run(async function (ctx) {
        var paras = ctx.document.body.paragraphs;
        paras.load('items');
        await ctx.sync();
        paras.items[index].select();
        await ctx.sync();
      });
    } catch (err) {
      setStatus('Could not jump: ' + err.message);
    }
  }

  // ---------- Data flow ----------

  async function refresh() {
    if (inWord) setStatus('Reading document...');
    try {
      state.paras = inWord ? await readDocument() : paragraphsFrom($('source').value);
      recompute();
      setStatus(state.rows.length + ' paragraphs measured');
    } catch (err) {
      setStatus('Could not read the document: ' + err.message);
    }
  }

  function recompute() {
    state.rows = M.analyzeDocument(state.paras);
    state.profile = M.buildProfile(paragraphsFrom($('target').value));
    var origParas = paragraphsFrom($('orig').value);
    state.orig = origParas.length ? M.analyzeDocument(origParas) : [];
    render();
  }

  // ---------- Rendering ----------

  function buildMetricButtons() {
    var box = $('metrics');
    M.METRICS.forEach(function (m) {
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'seg';
      b.textContent = m.label;
      b.setAttribute('data-key', m.key);
      b.setAttribute('aria-pressed', 'false');
      b.addEventListener('click', function () {
        state.metric = m.key;
        render();
      });
      box.appendChild(b);
    });
  }

  function renderMetricButtons() {
    Array.prototype.forEach.call($('metrics').children, function (b) {
      b.setAttribute('aria-pressed', b.getAttribute('data-key') === state.metric ? 'true' : 'false');
    });
  }

  // Words, not numbers, describe where a paragraph sits relative to the target.
  function describe(cls, value, prof) {
    if (!prof) return 'No target set yet';
    var high = value > prof.mean;
    if (cls === 'in') return 'In your target range';
    if (cls === 'near') return high ? 'Slightly above target' : 'Slightly below target';
    return high ? 'Well above target' : 'Well below target';
  }

  function renderDrift() {
    var key = state.metric, prof = state.profile && state.profile[key];
    var rows = state.rows, orig = state.orig;
    var el = $('drift'), summary = $('summary');

    if (!rows.length) {
      el.innerHTML = '<p class="empty">Read the document to see each paragraph here.</p>';
      summary.textContent = '';
      return;
    }

    var vals = [];
    rows.forEach(function (r) { if (r[key] != null) vals.push(r[key]); });
    orig.forEach(function (r) { if (r[key] != null) vals.push(r[key]); });
    if (!vals.length) {
      el.innerHTML = '<p class="empty">No measurable words for this measure yet.</p>';
      summary.textContent = '';
      return;
    }

    // The target band is clamped to values the measure can actually take (never below 0,
    // and never above 100 for percentages), so it always fits on the track.
    var maxVal = MAX_VALUE[key] != null ? MAX_VALUE[key] : Infinity;
    var bandLo = prof ? Math.max(0, prof.mean - prof.half) : null;
    var bandHi = prof ? Math.min(maxVal, prof.mean + prof.half) : null;

    var lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    if (prof) {
      lo = Math.min(lo, bandLo);
      hi = Math.max(hi, bandHi);
    }
    var span = (hi - lo) || 1;
    lo = Math.max(0, lo - span * 0.08);
    hi = Math.min(maxVal, hi + span * 0.08);
    var range = (hi - lo) || 1;
    function frac(v) { return Math.min(1, Math.max(0, (v - lo) / range)); }
    // Dots are 14px wide, so positions are inset 7px from each end of the track.
    function pos(v) { return 'calc(7px + (100% - 14px) * ' + frac(v).toFixed(4) + ')'; }

    var outCount = 0, nearCount = 0;
    var html = rows.map(function (r, i) {
      var v = r[key];
      var label = '<button type="button" class="drift-label" data-jump="' + r.index +
        '" aria-label="Go to paragraph ' + (r.index + 1) + '">' + (r.index + 1) + '</button>';

      if (v == null) {
        return '<div class="drift-row">' + label +
          '<div class="drift-body"><div class="track"></div><div class="drift-note">' + HELP[key].empty + '</div></div></div>';
      }

      var cls = M.classify(v, prof);
      if (cls === 'out') outCount++;
      if (cls === 'near') nearCount++;

      var o = orig[i] ? orig[i][key] : null;
      var band = '';
      if (prof) {
        // Keep very narrow ranges visible: widen to at least 5% of the track, staying inside it.
        var a = frac(bandLo), z = frac(bandHi), MIN = 0.05;
        if (z - a < MIN) {
          var mid = (a + z) / 2;
          a = Math.min(Math.max(0, mid - MIN / 2), 1 - MIN);
          z = a + MIN;
        }
        band = '<span class="target" style="left:calc(7px + (100% - 14px) * ' + a.toFixed(4) +
          ');width:calc((100% - 14px) * ' + (z - a).toFixed(4) + ')"></span>';
      }
      var origDot = o != null ? '<span class="dot orig-dot" style="left:' + pos(o) +
        '" title="Original draft, paragraph ' + (i + 1) + '"></span>' : '';
      var dot = '<span class="dot ' + cls + '" style="left:' + pos(v) +
        '" title="Paragraph ' + (r.index + 1) + '"></span>';

      return '<div class="drift-row">' + label +
        '<div class="drift-body"><div class="track">' + band + origDot + dot + '</div>' +
        '<div class="drift-note">' + describe(cls, v, prof) + '</div></div></div>';
    }).join('');

    html += '<div class="scale"><span>' + HELP[key].ends[0] + '</span><span>' + HELP[key].ends[1] + '</span></div>';
    el.innerHTML = html;

    if (!prof) {
      summary.textContent = 'Paste a target style sample to see which paragraphs match it.';
    } else if (outCount === 0 && nearCount === 0) {
      summary.textContent = 'Every paragraph sits inside your target range on this measure.';
    } else {
      summary.textContent = outCount + ' of ' + rows.length + ' paragraphs are well outside your target range' +
        (nearCount ? ', and ' + nearCount + (nearCount === 1 ? ' is' : ' are') + ' close to its edge.' : '.') +
        ' Those are the places to revisit.';
    }
  }

  function renderStrip() {
    var key = state.metric, prof = state.profile && state.profile[key];
    var el = $('strip');
    if (!state.rows.length) {
      el.innerHTML = '<p class="empty">No paragraphs yet.</p>';
      return;
    }
    el.innerHTML = state.rows.map(function (r) {
      var cls = M.classify(r[key], prof);
      var word = cls === 'in' ? 'in range' : cls === 'near' ? 'near edge' : cls === 'out' ? 'drifted' : 'no words';
      return '<button type="button" class="cell ' + cls + '" data-jump="' + r.index +
        '" title="Paragraph ' + (r.index + 1) + ': ' + word + '" aria-label="Paragraph ' +
        (r.index + 1) + ', ' + word + '">' + (r.index + 1) + '</button>';
    }).join('');
  }

  function render() {
    var key = state.metric;
    var prof = state.profile && state.profile[key];
    renderMetricButtons();
    $('metric-hint').textContent = HELP[key].text +
      (prof ? '' : ' Paste a target style sample to add the shaded target range.');
    renderDrift();
    renderStrip();
    $('legend-orig').hidden = !state.orig.length;
  }

  // ---------- Setup ----------

  function init() {
    buildMetricButtons();
    $('refresh').addEventListener('click', refresh);
    $('target').addEventListener('input', debounce(recompute, 250));
    $('orig').addEventListener('input', debounce(recompute, 250));
    document.addEventListener('click', function (e) {
      var t = e.target && e.target.closest ? e.target.closest('[data-jump]') : null;
      if (t) jump(Number(t.getAttribute('data-jump')));
    });
    if (!inWord) setUpWebMode();
    refresh();
  }

  // Outside Word: show the text box, hide "Read document", and update as the user types.
  function loadSample() {
    $('source').value = DEMO_DOC.join('\n\n');
    $('target').value = DEMO_TARGET;
    $('orig').value = DEMO_ORIGINAL;
  }

  function setUpWebMode() {
    $('source-card').hidden = false;
    $('refresh').hidden = true;
    $('source').addEventListener('input', debounce(refresh, 250));
    $('load-sample').addEventListener('click', function () { loadSample(); refresh(); });
    $('clear-source').addEventListener('click', function () {
      $('source').value = '';
      refresh();
      $('source').focus();
    });
    loadSample();
  }

  var started = false;
  function start(info) {
    var word = !!(info && typeof Office !== 'undefined' && info.host === Office.HostType.Word);
    if (started) {
      // Word reported in after the fallback started website mode: switch to Word mode.
      if (word && !inWord) {
        inWord = true;
        $('source-card').hidden = true;
        $('refresh').hidden = false;
        refresh();
      }
      return;
    }
    started = true;
    inWord = word;
    init();
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(start);
    // If Office.js never reports ready (blocked or offline), fall back to the website mode.
    // A slow Word start still switches back to Word mode when onReady finally fires.
    setTimeout(function () { start(null); }, 3000);
  } else {
    start(null);
  }
})();
