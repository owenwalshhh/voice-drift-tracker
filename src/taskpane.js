/* Voice Drift Tracker: task pane UI and Word integration. */
(function () {
  'use strict';

  var M = window.VoiceDrift;
  var hasOffice = typeof Office !== 'undefined';
  var state = { metric: 'sentenceLength', paras: [], rows: [], orig: [], profile: null };

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
      var on = b.getAttribute('data-key') === state.metric;
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
  }

  function path(list, key, x, y) {
    var d = '', pen = false;
    list.forEach(function (r, i) {
      var v = r[key];
      if (v == null) { pen = false; return; }
      d += (pen ? ' L' : ' M') + x(i).toFixed(1) + ' ' + y(v).toFixed(1);
      pen = true;
    });
    return d.trim();
  }

  function chartSVG() {
    var key = state.metric, rows = state.rows, orig = state.orig;
    var prof = state.profile && state.profile[key];
    if (!rows.length) return '<p class="empty">Nothing to measure yet. Write a few paragraphs, then read the document.</p>';

    var vals = [];
    rows.concat(orig).forEach(function (r) { if (r[key] != null) vals.push(r[key]); });
    if (!vals.length) return '<p class="empty">No measurable words for this metric yet.</p>';

    var lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    if (prof) {
      lo = Math.min(lo, prof.mean - prof.half);
      hi = Math.max(hi, prof.mean + prof.half);
    }
    // Every metric is zero or above, so the axis never goes negative.
    var span = (hi - lo) || 1;
    lo = Math.max(0, lo - span * 0.12);
    hi += span * 0.12;

    var W = 300, H = 160, L = 36, R = 12, T = 10, B = 22;
    var n = Math.max(rows.length, orig.length);
    var x = function (i) { return n <= 1 ? (L + W - R) / 2 : L + i * (W - L - R) / (n - 1); };
    var y = function (v) { return T + (1 - (v - lo) / (hi - lo)) * (H - T - B); };

    var out = [];
    out.push('<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="' +
      metricDef(key).label + ' across paragraphs">');

    for (var k = 0; k <= 2; k++) {
      var t = lo + k * (hi - lo) / 2;
      out.push('<line class="grid" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t).toFixed(1) + '" y2="' + y(t).toFixed(1) + '"/>');
      out.push('<text class="axis" x="' + (L - 6) + '" y="' + (y(t) + 3).toFixed(1) + '" text-anchor="end">' + t.toFixed(1) + '</text>');
    }

    if (prof) {
      var top = y(prof.mean + prof.half), bottom = y(prof.mean - prof.half);
      out.push('<rect class="band" x="' + L + '" y="' + top.toFixed(1) + '" width="' + (W - L - R) +
        '" height="' + Math.max(0, bottom - top).toFixed(1) + '"/>');
    }

    if (orig.length) {
      out.push('<path class="orig" d="' + path(orig, key, x, y) + '"/>');
      orig.forEach(function (r, i) {
        if (r[key] == null) return;
        out.push('<circle class="orig-pt" cx="' + x(i).toFixed(1) + '" cy="' + y(r[key]).toFixed(1) +
          '" r="2.5"><title>Original, paragraph ' + (i + 1) + ': ' + r[key].toFixed(1) + '</title></circle>');
      });
    }
    out.push('<path class="line" d="' + path(rows, key, x, y) + '"/>');

    rows.forEach(function (r, i) {
      if (r[key] == null) return;
      out.push('<circle class="pt" data-jump="' + r.index + '" cx="' + x(i).toFixed(1) +
        '" cy="' + y(r[key]).toFixed(1) + '" r="3.5"><title>Paragraph ' + (r.index + 1) + ': ' +
        r[key].toFixed(1) + '</title></circle>');
    });

    out.push('</svg>');
    return out.join('');
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
      var v = r[key] == null ? 'no value' : r[key].toFixed(1);
      return '<button type="button" class="cell ' + cls + '" data-jump="' + r.index +
        '" title="Paragraph ' + (r.index + 1) + ': ' + v + '" aria-label="Paragraph ' +
        (r.index + 1) + ', ' + v + '">' + (r.index + 1) + '</button>';
    }).join('');
  }

  function renderSummary() {
    var key = state.metric, prof = state.profile && state.profile[key];
    var el = $('summary');
    if (!prof) {
      el.textContent = 'Paste a target sample to see which paragraphs drift from it.';
      return;
    }
    var out = 0, near = 0;
    state.rows.forEach(function (r) {
      var c = M.classify(r[key], prof);
      if (c === 'out') out++;
      if (c === 'near') near++;
    });
    el.textContent = out + ' of ' + state.rows.length + ' paragraphs drift outside the target band; ' +
      near + ' sit near its edge.';
  }

  function render() {
    renderMetricButtons();
    $('chart').innerHTML = chartSVG();
    renderStrip();
    renderSummary();
    var def = metricDef(state.metric);
    $('metric-hint').textContent = def.unit + (state.profile ? '' : '. Add a target sample to see the band.');
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
