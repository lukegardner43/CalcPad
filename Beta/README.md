# CalcPad Beta

Live at **https://lukegardner43.github.io/CalcPad/Beta**

A place to try AI features on real calculations before they go anywhere near
the production sheet.

## Why this is not a second copy of the app

The obvious way to build a beta is to copy `index.html` into `Beta/`. Don't —
the copy stops tracking main the first time a fix lands in one and not the
other, and after a few weeks nobody can say which behaviour the beta is
actually testing.

So there is only ever one app:

```
index.html              the app — the only copy, unchanged except for ~15 lines
                        at the end that load the files below when ?beta=1 is set
Beta/index.html         a splash page that redirects to ../index.html?beta=1
Beta/ai-features.js     the AI layer
Beta/ai-features.css    its styles
```

Opening `/CalcPad/Beta` lands on the splash page, which sends the browser to
`/CalcPad/index.html?beta=1`. The app sees the flag and pulls in the two files
above. **Beta is therefore always today's main plus the AI layer** — a fix
pushed to `index.html` is in the beta the moment it is live, with nothing to
port and nothing to keep in step.

Other ways in, all equivalent: `?beta=1`, `#beta`, or any URL with `/Beta/` in
the path.

The trade is that the address bar reads `/CalcPad/?beta=1` rather than
`/CalcPad/Beta`. That is the whole cost, and it buys away the divergence.

## What is in the beta

### AI QA
Reads the sheet the same way the Word export does — chips come through as
`M_Ed = w*L^2/8 = 45 kN·m`, not as markup — and reports on:

- **flow** — does the argument hold together, is each quantity introduced
  before it is used, does it reach a conclusion
- **spelling** — British English house style
- **grammar** — structure, agreement, punctuation

plus units, clarity and obvious gaps where it sees them. Each finding carries
a specific correction and a line reference; clicking one highlights the text on
the page. **Nothing is ever changed for you** — it reports, you decide.

### Assistant
A chat box that drafts calculations. It is given the CalcPad syntax, the title
block, every variable currently defined and the document text, so it builds on
the sheet rather than starting from scratch. Replies come back with the lines
in a fenced block and an **Insert at cursor** button, which routes through the
app's own paste path: real chips, evaluated on insert, with the usual
duplicate-name prompt if it would redefine something.

## The API key

Testers use their own Anthropic key. It is entered once per browser tab and:

- lives in `sessionStorage` — gone when the tab closes
- is never written into the document, an autosave, or a `.perega` file
- goes straight from the browser to `api.anthropic.com` and nowhere else

The browser can only call the API at all because of the
`anthropic-dangerous-direct-browser-access` header. That is fine for testing
with a key you can rotate — set a spend limit on it. It would not be fine for
a public release, where the call belongs behind a small server endpoint that
holds one key. That is the main thing to change before any of this ships.

## Adding another beta feature

Put it in `Beta/`. Add the file to `loadBetaFeatures()` in `index.html` if it
needs its own script. Nothing beta-specific belongs in the core app beyond
that loader — delete this folder and the app carries on as if it never existed.
