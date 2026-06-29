# 📐 Math Image Generator

Generates math graph images (line graphs & scatter plots) from text prompts, embeds them into `.docx` placeholders, and supports batch processing via Google Drive.

---

## Features

- **Single document mode** — upload a `.docx`, select a placeholder, paste a prompt, download the updated doc with the image embedded
- **Batch mode** — point at a Google Drive folder, process every `.docx` automatically, output goes to a `Generated_Images` subfolder
- **AI parsing** (optional, recommended) — uses Claude to parse any prompt phrasing reliably
- **Regex fallback** — works without an API key for known prompt formats

---

## Deploying to Streamlit Community Cloud (free permanent URL)

### Step 1 — Push this repo to GitHub

```bash
# On your machine (one time):
git init
git add .
git commit -m "Initial commit"
# Create a new repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/math-image-generator.git
git push -u origin main
```

### Step 2 — Connect to Streamlit Community Cloud

1. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub
2. Click **"New app"**
3. Select your repo → branch: `main` → main file: `app.py`
4. Click **"Deploy"**

### Step 3 — Add Secrets (API keys)

In your app dashboard → **⚙️ Settings → Secrets**, paste:

```toml
ANTHROPIC_API_KEY = "sk-ant-YOUR_KEY_HERE"

GOOGLE_CREDS_JSON = '''
{ ... paste your service account JSON here ... }
'''
```

That's it — your coworkers get a permanent URL like `https://your-app-name.streamlit.app`

---

## Google Drive Setup (for Batch Mode)

### Option A — Service Account (recommended for teams)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project (or select existing)
3. Enable **Google Drive API** (APIs & Services → Enable APIs)
4. Go to **IAM & Admin → Service Accounts → Create Service Account**
5. Give it a name, click **Create**
6. Click the service account → **Keys → Add Key → JSON** → download
7. **Share your Drive folder** with the service account email (`xxx@your-project.iam.gserviceaccount.com`) — give it **Editor** access
8. Paste the JSON content into Streamlit secrets as `GOOGLE_CREDS_JSON`

### Option B — Paste JSON in the sidebar

If you don't want to store credentials in secrets, paste the service account JSON directly in the sidebar each session.

---

## Placeholder Format

The app looks for paragraphs containing:

```
Image (if any): <your prompt here>
```

The prompt text after the colon is what gets sent to the parser.

---

## Prompt Formats Supported

**Line graphs:**
- `the line y=2x+1 (solid) passes through (0,1) and (2,5)`
- `Line A: y=2x+1, passing through points (0,1) and (2,5)`
- `Dataset 1 — Elena's Savings (solid line): Point 1: x=0, y=10 → labeled (0,10)...`

**Scatter grids:**
- `Plot A points: (1,2), (2,3), (3,5)`
- `Plot A shows points 3,9,5,4,6,16`

---

## Local Development

```bash
pip install -r requirements.txt
streamlit run app.py
```
