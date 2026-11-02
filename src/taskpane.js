/* Voice Drift Tracker: task pane UI and Word integration. */
(function () {
  'use strict';

  var M = window.VoiceDrift;
  var hasOffice = typeof Office !== 'undefined';
  var state = { metric: 'sentenceLength', paras: [], rows: [], orig: [], profile: null };

  // Plain-language help for each measure. ends[0] is the low side, ends[1] the high side.
  var HELP = {
    sentenceLength: { text: 'Average words per sentence.', ends: ['Shorter sentences', 'Longer sentences'] },
    wordLength: { text: 'Average letters per word. Longer words usually sound more formal.', ends: ['Simpler words', 'Longer words'] },
    hedging: { text: 'How often you use words like "might", "perhaps", and "seems". More hedging sounds more tentative.', ends: ['More certain', 'More tentative'] },
    personal: { text: 'How often you use I, we, and you compared with he, she, and they. Higher means more personal.', ends: ['Impersonal', 'Personal'] },
    punctuation: { text: 'How often you use question marks, exclamation points, and dashes. Higher means more expressive.', ends: ['Plain', 'Expressive'] }
  };

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

  function paragraphsFrom(text) {
    return text.split(/\n\s*\n/).map(function (s) { return s.trim(); }).filter(Boolean);
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
    if (!hasOffice) {
      setStatus('Jumping to paragraph ' + (index + 1) + ' works inside Word.');
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
    setStatus('Reading document...');
    try {
      state.paras = hasOffice ? await readDocument() : DEMO_DOC;
      recompute();
      setStatus(state.rows.length + ' paragraphs measured' + (hasOffice ? '' : ' (demo text)'));
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

    var lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    if (prof) {
      lo = Math.min(lo, prof.mean - prof.half);
      hi = Math.max(hi, prof.mean + prof.half);
    }
    var span = (hi - lo) || 1;
    lo = Math.max(0, lo - span * 0.08);
    hi = hi + span * 0.08;
    var range = hi - lo;
    function pos(v) { return ((v - lo) / range * 100).toFixed(2) + '%'; }

    var outCount = 0, nearCount = 0;
    var html = rows.map(function (r, i) {
      var v = r[key];
      var label = '<button type="button" class="drift-label" data-jump="' + r.index +
        '" aria-label="Go to paragraph ' + (r.index + 1) + '">' + (r.index + 1) + '</button>';

      if (v == null) {
        return '<div class="drift-row">' + label +
          '<div class="drift-body"><div class="track"></div><div class="drift-note">No words to measure here</div></div></div>';
      }

      var cls = M.classify(v, prof);
      if (cls === 'out') outCount++;
      if (cls === 'near') nearCount++;

      var o = orig[i] ? orig[i][key] : null;
      var band = prof ? '<span class="target" style="left:' + pos(prof.mean - prof.half) +
        ';width:' + ((2 * prof.half) / range * 100).toFixed(2) + '%"></span>' : '';
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
    if (!hasOffice) {
      $('target').value = DEMO_TARGET;
      $('orig').value = DEMO_ORIGINAL;
    }
    refresh();
  }

  if (hasOffice) {
    Office.onReady(init);
  } else {
    init();
  }
})();
