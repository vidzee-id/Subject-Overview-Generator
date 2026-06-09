# Subject Overview Generator

Upload any subject syllabus → get a beautifully designed PNG overview image, powered by Claude AI.

---

## What you need before starting

- A GitHub account ✓ (you have this)
- A free Streamlit Cloud account → sign up at https://streamlit.io (use your GitHub login)
- An Anthropic API key → get one at https://console.anthropic.com

---

## Step 1 — Create a GitHub repository

1. Go to https://github.com/new
2. Name it `subject-overview-generator` (or anything you like)
3. Set it to **Public**
4. Click **Create repository**

---

## Step 2 — Add the files

In your new repo, create these four files by clicking **Add file → Create new file**:

| File | What it does |
|------|-------------|
| `app.py` | The main application code |
| `requirements.txt` | Python packages to install |
| `packages.txt` | System packages to install |

Copy the contents of each file from this folder into GitHub. Commit each one.

---

## Step 3 — Deploy on Streamlit Cloud

1. Go to https://share.streamlit.io
2. Click **New app**
3. Connect your GitHub account if prompted
4. Select your repository (`subject-overview-generator`)
5. Set **Main file path** to `app.py`
6. Click **Deploy**

Streamlit will build and launch your app. It takes about 2 minutes.

---

## Step 4 — Add your API key

1. In the Streamlit Cloud dashboard, find your deployed app
2. Click the **⋮ menu** → **Settings** → **Secrets**
3. Add this exactly (replace with your real key):

```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
```

4. Click **Save** — the app restarts automatically

---

## Step 5 — Share the link

Your app is now live at a URL like:
```
https://your-name-subject-overview-generator-app-xxxx.streamlit.app
```

Copy it from the Streamlit Cloud dashboard and share it with anyone.
Anyone with the link can upload a syllabus and download an overview image.
No Claude account needed — it uses your API key.

---

## Cost

Each poster generation costs approximately **$0.01–0.03** in API usage.
You can set a monthly spending limit at https://console.anthropic.com to stay in control.

---

## Supported file types

- PDF (`.pdf`)
- Word document (`.docx`)
- Plain text (`.txt`)
