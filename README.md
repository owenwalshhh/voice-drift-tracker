<<<<<<< HEAD
# voice-drift-tracker
=======
# Voice Drift Tracker

A Microsoft Word task pane that charts how your writing voice shifts from paragraph to paragraph, and compares it with a target style.

## What it shows

- **Measure:** switch between sentence length, word length (vocabulary difficulty), hedging (might, perhaps, seems), personal voice (I/we/you vs. he/she/they), and punctuation mood (?, !, dashes).
- **Drift chart:** one point per paragraph, with a green target band from your style sample. Click a point to jump to that paragraph in Word.
- **Paragraph strip:** each paragraph is coloured in band, near the edge, or drifted. Click one to jump to it.
- **Original overlay:** paste the draft from before your rewrite to see both on the same chart (paragraphs are matched by position).

## Run it

Requires Node.js (for the certificate and local server) and desktop Word on Windows or Mac.

```bash
npx office-addin-dev-certs install    # one-time: trusts a local HTTPS certificate
npx http-server src -p 3000 -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -c-1
```

Leave the server running. Then in Word:

1. Open **Insert > Add-ins > My Add-ins > Upload My Add-in** (on some versions: **Insert > Get Add-ins > My Add-ins**).
2. Choose `manifest.xml` from this folder.
3. Click **Voice Drift Tracker** in the ribbon. Click **Read document**.

## Test the metrics

```bash
npm test
```

The metric logic lives in `src/metrics.js` with no Office dependencies, so it is covered by Node tests in `test/`.

## Notes

- Paragraph counts ignore empty paragraphs, but the Word paragraph number is kept so clicks jump to the right place.
- Word's sentence splitting is approximate. It handles common abbreviations (Dr., Mr., etc.) but not every edge case.
- The hedge, pronoun, and punctuation lists are short on purpose. Edit `src/metrics.js` to change them.
- Outside Word, `src/taskpane.html` opens with demo text so you can preview the layout in a browser.
>>>>>>> 2d7e681 (Add Voice Drift Tracker Word task pane)
