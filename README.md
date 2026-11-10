# AI Email Assistant — Outlook Add-in (multi-provider)

One shared task pane (`taskpane.html`) + one manifest per AI provider, in
`manifests/`. All ten manifests point at the same `taskpane.html`, just with
a different `?provider=` flag — so the pane opens pre-selected to that model,
but you can still switch providers from the dropdown inside the pane.

Providers included: OpenAI (GPT), Claude (Anthropic), Gemini (Google), Qwen
(Alibaba DashScope), DeepSeek, Grok (xAI), Mistral, Groq, Together AI,
Perplexity.

Note on "Copilot": Microsoft Copilot has no public direct-chat API for
individual developers to call like this — Microsoft's equivalent is **Azure
OpenAI Service**, which uses the OpenAI-compatible format already built into
`taskpane.html`. If you have an Azure OpenAI deployment, duplicate
`manifests/manifest-openai.xml` as `manifest-azure.xml`, and change the
`provider` query param's target endpoint in `taskpane.html` (`OPENAI_COMPAT`
base URL) to your Azure endpoint, or just reuse the OpenAI entry with your
Azure key/endpoint swapped in.

## 1. Host the files

Outlook add-ins must be loaded over HTTPS from a real URL — they can't be
loaded from your local disk. The easiest free option is GitHub Pages:

1. Create a public GitHub repo, e.g. `ai-email-addin`.
2. Upload `taskpane.html`, the `assets/` folder, and (optionally) this
   README to the repo root.
3. In the repo settings, enable **GitHub Pages** for the `main` branch /
   root folder. You'll get a URL like
   `https://YOUR-USERNAME.github.io/ai-email-addin`.
4. Open `generate_manifests.py`, set `BASE_URL` to that exact URL (no
   trailing slash), and re-run it:
   ```
   python3 generate_manifests.py
   ```
   This regenerates all 10 manifests in `manifests/` with the correct URLs
   baked in. (The manifests included in this package already have
   `https://YOUR-USERNAME.github.io/ai-email-addin` as a placeholder — you
   must replace it before sideloading, either by editing the script and
   rerunning it, or with find-and-replace across the XML files.)

Any other static HTTPS host works too (Netlify, Vercel, Azure Static Web
Apps, S3 + CloudFront, your own server) — just point `BASE_URL` at it.

## 2. Sideload a manifest into Outlook

**Outlook on the web / new Outlook (Windows):**
Settings → General → Manage add-ins → *My add-ins* → **Add a custom add-in**
→ **Add from file** → pick e.g. `manifest-claude.xml`.

**Classic Outlook desktop (Windows):**
File → Manage Add-ins (opens Outlook on the web) → same steps as above, or
Home → Get Add-ins → My add-ins → Add Custom Add-in → Add from File.

**Outlook on Mac:**
Get Add-ins → My add-ins → Add Custom Add-in → Add from File.

Repeat for as many provider manifests as you want — each shows up as its
own button in the ribbon (e.g. "OpenAI", "Claude", "Gemini", ...), all
backed by the same task pane.

## 3. Use it

1. Open an email (read or compose/reply) and click the provider's ribbon
   button.
2. Paste your API key for that provider into the task pane. It's saved in
   the browser's local storage, scoped per provider — never sent anywhere
   except directly to that provider's API.
3. Click **Summarize email** or **Draft reply**. The current email's text
   is sent straight from your browser to the provider's API over HTTPS.
4. In a reply/compose window, **Insert into reply body** drops the result
   straight into the message.

## 4. Important: CORS / direct-browser-call limits

This add-in calls each provider's API **directly from the browser** — there
is no backend server. Most providers' chat-completion endpoints are
designed for direct browser use or at least don't block it, but policies
can change, and a few notes:

- **Claude**: the request includes
  `anthropic-dangerous-direct-browser-access: true`, which Anthropic
  requires for browser-origin calls. This is intended for
  prototyping/personal use — for production or shared use, Anthropic
  recommends calling the API from a backend instead.
- **OpenAI-compatible providers** (OpenAI, DeepSeek, Grok, Mistral, Groq,
  Together, Perplexity, Qwen/DashScope): generally reachable directly, but
  if you see a network error with no response body, it's almost always a
  CORS block — the fix is to proxy the request through a small backend
  (e.g. a Cloudflare Worker or a one-route Node/Express server) that adds
  your key server-side instead of in the browser.
- **Gemini**: the API key is passed as a URL query parameter per Google's
  REST format; treat your hosted task pane URL as effectively public, since
  anyone who can see network requests from your own browser session can
  see the key you typed in (this is true of all providers here — the key
  only ever leaves your machine, but it's visible in your own browser's
  dev tools).

If you want a hardened, production version (key never touches the browser,
usage logging, per-user access control), the next step is adding a thin
backend proxy — happy to build that if you want it.

## Files in this package

```
taskpane.html           shared task pane UI + logic for all providers
generate_manifests.py   regenerates manifests/*.xml from BASE_URL
manifests/
  manifest-openai.xml
  manifest-claude.xml
  manifest-gemini.xml
  manifest-qwen.xml
  manifest-deepseek.xml
  manifest-grok.xml
  manifest-mistral.xml
  manifest-groq.xml
  manifest-together.xml
  manifest-perplexity.xml
assets/
  icon-<provider>-16.png / -32.png / -80.png   (placeholder icons, swap anytime)
```
