# Voice Drift Tracker

A Microsoft Word add-in that shows where your writing changes voice. Each paragraph is compared with a target style you choose, so you can see which parts drift.

Try it in your browser, no Word needed: https://owenwalshhh.github.io/voice-drift-tracker/taskpane.html. Paste your writing into the **Your text** box, one paragraph per line, and the results update as you type.

## Install in Word (no setup needed)

You need desktop Word (Microsoft 365 for Windows or Mac).

1. Download the manifest file: open https://github.com/owenwalshhh/voice-drift-tracker/blob/main/manifest.xml, click the **Download raw file** button (or right-click the **Raw** link and choose **Save link as**), and save it as `manifest.xml`.
2. Open Word.
3. **Windows:** Insert > Get Add-ins > My Add-ins > Upload My Add-in, then choose `manifest.xml`.
   **Mac:** Insert > Add-ins > My Add-ins > Upload My Add-in, then choose `manifest.xml`.
   If Mac Word doesn't show Upload My Add-in, quit Word, then copy `manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the folder if it's missing) and reopen Word.
4. Click **Voice Drift** in the Home tab (or under My Add-ins) to open the pane.

## Use it

1. Write or paste your document in Word. Put a blank line between paragraphs.
2. Click **Read document**.
3. Under **Target style and original draft**, paste a passage in the style you want. The shaded bar in each row is built from it.
4. Pick a measure, such as sentence length or hedging. Each row is one paragraph. A dot outside the shaded bar is a paragraph that drifts from the target.
5. Click a paragraph number to jump to it in Word.

To compare with a rewrite, paste the draft from before the rewrite under **Original draft**. It appears as a hollow dot in each row.

## Measures

- **Sentence length:** average words per sentence.
- **Word length:** average letters per word.
- **Hedging:** words like "might", "perhaps", and "seems", per 100 words.
- **Personal voice:** I, we, and you compared with he, she, and they.
- **Punctuation mood:** question marks, exclamation points, and dashes, per 100 words.

## Developer setup

Run the tests:

```bash
npm test
```

The metric logic is in `src/metrics.js` and has no Office dependencies, so it runs under Node.

To edit the task pane and test it in Word before publishing, serve `src` over HTTPS on port 3000:

1. Run `npx office-addin-dev-certs install` once. Enter your Mac password when asked.
2. Run `npx http-server src -p 3000 -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -c-1`.
3. Upload `manifest.local.xml` in Word instead of `manifest.xml`.

Pushes to `main` publish the task pane automatically through GitHub Pages.

## Publishing (one-time setup)

1. On GitHub, open **Settings > Pages** and set **Source** to **GitHub Actions**.
2. Push to `main`. The **Deploy task pane to GitHub Pages** workflow publishes `src/` to the live URL above.

On a free GitHub account, Pages requires a public repository.

## Notes

- Word's sentence splitting is approximate. It handles common abbreviations such as "Dr." but not every edge case.
- The hedge, pronoun, and punctuation word lists are short on purpose. Edit `src/metrics.js` to change them.
- Paragraphs in the original draft are matched to rows by position, so a long insertion can shift the comparison.
