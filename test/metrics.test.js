const test = require('node:test');
const assert = require('node:assert/strict');
const M = require('../src/metrics.js');

test('words ignores punctuation and keeps contractions', () => {
  assert.deepEqual(M.words("Don't stop, it's fine."), ["don't", 'stop', "it's", 'fine']);
});

test('sentences splits on terminal punctuation', () => {
  assert.equal(M.sentences('Hello there. How are you? Fine!').length, 3);
});

test('sentences does not split on common abbreviations', () => {
  assert.equal(M.sentences('Dr. Smith arrived.').length, 1);
});

test('analyzeParagraph returns null for empty text', () => {
  assert.equal(M.analyzeParagraph('   '), null);
});

test('sentence length is words divided by sentences', () => {
  const m = M.analyzeParagraph('One two three four. Five six.');
  assert.equal(m.sentenceLength, 3);
});

test('hedging counts hedge words per 100 words', () => {
  // words: it, might, seem, possible, perhaps -> 3 hedges out of 5 words
  const m = M.analyzeParagraph('It might seem possible, perhaps.');
  assert.equal(m.hedging, 60);
});

test('personal voice is the share of I/we/you among person words', () => {
  assert.equal(M.analyzeParagraph('I think we can.').personal, 100);
  assert.equal(M.analyzeParagraph('They said he would.').personal, 0);
});

test('personal voice is null when the paragraph has no person words', () => {
  assert.equal(M.analyzeParagraph('Budgets matter greatly.').personal, null);
});

test('punctuation mood counts question marks, exclamations and dashes', () => {
  // two marks over two words
  assert.equal(M.analyzeParagraph('Why? Really!').punctuation, 100);
  // one dash over four words
  assert.equal(M.analyzeParagraph('Good — very good stuff.').punctuation, 25);
});

test('word length averages letters per word', () => {
  assert.equal(M.analyzeParagraph('cat dog').wordLength, 3);
});

test('analyzeDocument keeps original paragraph indexes and skips empty ones', () => {
  const rows = M.analyzeDocument(['First one.', '', 'Third here.']);
  assert.deepEqual(rows.map((r) => r.index), [0, 2]);
});

test('buildProfile returns null with no target text', () => {
  assert.equal(M.buildProfile([]), null);
});

test('buildProfile gives a band that is never narrower than the metric floor', () => {
  const text = 'One sentence with six words right here.';
  const p = M.buildProfile([text]);
  const sl = p.sentenceLength;
  assert.equal(sl.mean, M.analyzeParagraph(text).sentenceLength);
  assert.ok(sl.half >= 1.5, 'half-width should be at least the floor');
});

test('classify marks values in, near, or out of the band', () => {
  const p = { mean: 10, half: 2 };
  assert.equal(M.classify(11, p), 'in');
  assert.equal(M.classify(13, p), 'near');
  assert.equal(M.classify(20, p), 'out');
  assert.equal(M.classify(null, p), 'none');
  assert.equal(M.classify(10, null), 'none');
});
