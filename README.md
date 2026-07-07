# Unani Terminology Converter — Word Add-in (v1 / Phase 3 MVP)

## What this actually does (and doesn't) right now

**Included in this build:**
- Deterministic dictionary lookup against WHO IUMT 2022 (4,502 terms) + CCRUM 2012 (2,388 terms) — 5,461 usable entries after removing flagged/ambiguous noise
- Convert Selected Text / Convert Entire Document
- High-confidence matches applied automatically; ambiguous common-English words (pain, chronic, etc.) flagged in the log but *not* silently changed unless you tick "Also apply REVIEW-tier matches"
- Highlighting of replaced terms (green = high confidence, yellow = review-tier)
- Search box to look up any term against the database directly
- A running conversion log showing original → Unani term, source, and Term ID for every match
- Undo: this uses normal Word document edits, so **Ctrl+Z undoes conversions exactly like any other edit** — no custom undo logic needed

**Not yet built (be aware before you rely on this for the dissertation):**
- Hover tooltips showing full definitions on the replaced term itself (Word's comment API can do this — it's a clean next step, just not in this pass)
- Batch/multi-document conversion
- The English↔Arabic/Urdu direction, and any AI-assisted disambiguation for terms with multiple accepted equivalents beyond "prefer WHO over CCRUM"
- I have **not been able to test this inside actual Word** — I don't have Word available in my environment. The matching logic itself is tested and verified (see our Phase 2 conversation), but the Office.js wiring (search/replace, highlighting, comments) should be treated as a first draft that needs a real test pass on your machine.

---

## Files

```
word_addin/
├── manifest.xml              ← Add-in manifest (needs your hosting URL filled in)
├── taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js
└── assets/
    └── terminology.json      ← the actual database, ~1.2 MB, bundled client-side
```

## Step 1 — Host the files somewhere HTTPS

Word add-ins **require HTTPS** to load the task pane — `file://` won't work for sideloading. Two practical options:

**Option A — GitHub Pages (simplest, free, persistent)**
1. Create a GitHub repo, add all files under `word_addin/` to it
2. Enable GitHub Pages on the repo (Settings → Pages → deploy from main branch)
3. Your files will be at `https://<yourusername>.github.io/<repo>/...`
4. Open `manifest.xml` and replace every `https://REPLACE_WITH_YOUR_HOST` with that URL

**Option B — Local dev server (for quick testing today)**
```bash
npm install -g office-addin-dev-certs
npx office-addin-dev-certs install     # trusts a local HTTPS cert
npx http-server ./word_addin -p 3000 --ssl --cert <cert path> --key <key path>
```
Then use `https://localhost:3000` in the manifest instead.

You'll also need three small icon PNGs (16x16, 32x32, 80x80) in `assets/` — any simple icon works; Word just needs the files to exist at the URLs referenced in the manifest.

## Step 2 — Sideload into Word

**Word desktop (Windows/Mac):**
1. Home tab → Add-ins → "More Add-ins" → "My Add-ins" → gear/upload icon → "Upload My Add-in"
2. Browse to your edited `manifest.xml`
3. The "Unani Converter" button appears on the Home ribbon

**Word Online:**
1. Insert tab → Add-ins → Upload My Add-in → select `manifest.xml`

## Step 3 — Use it

1. Open the task pane (Home → Unani Converter)
2. Select a paragraph → "Convert Selected Text", or run "Convert Entire Document"
3. Check the Conversion Log — every change is listed with its source (WHO/CCRUM) and Term ID, so you can verify against the originals before it goes into a submitted chapter
4. Ctrl+Z to undo any conversion you don't want

## Recommended first real test

Don't run this on a live dissertation chapter first. Copy a paragraph from your Discussion into a throwaway Word doc, run the conversion there, and check:
- Did it correctly skip GERD/atony-type phrases with no exact WHO/CCRUM entry? (it should — see our test)
- Are the REVIEW-tier flags (pain, chronic, etc.) behaving the way you want, or is the ambiguous-word list too aggressive/too lax for your writing style? That list is easy to edit — it's a plain array in `taskpane.js`.
