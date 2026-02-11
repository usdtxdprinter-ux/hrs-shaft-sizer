# HRS Shaft Sizer — Deployment Guide
## GitHub: usdtxdprinter-ux → Streamlit Cloud

---

## Validation Summary

| Category | Tests | Result |
|---|---|---|
| Core Engineering (friction, VP, Darcy, areas) | 24 | ✅ ALL PASSED |
| Integration (full sizing engine, 10 scenarios) | 61 | ✅ ALL PASSED |
| App Structure (functions, steps, state, branding) | 22 | ✅ ALL PASSED |
| **TOTAL** | **107** | **✅ ALL PASSED** |

---

## Files You Need

You only need **two files** to deploy:

```
hrs_shaft_sizer.py      ← The app (main file)
requirements.txt        ← Dependencies
```

---

## Step-by-Step Deployment

### STEP 1 — Create the GitHub Repository

1. Go to **https://github.com/new**
2. Fill in:
   - **Repository name:** `hrs-shaft-sizer`
   - **Description:** `HRS Exhaust Shaft Sizing Calculator — LF Systems`
   - **Visibility:** Public *(Streamlit Cloud free tier requires public repos)*
   - ☑️ Check **"Add a README file"**
3. Click **Create repository**

Your repo URL will be:
```
https://github.com/usdtxdprinter-ux/hrs-shaft-sizer
```

---

### STEP 2 — Upload the Files

**Option A — GitHub Web UI (easiest):**

1. Go to your new repo: `https://github.com/usdtxdprinter-ux/hrs-shaft-sizer`
2. Click **"Add file"** → **"Upload files"**
3. Drag and drop both files:
   - `hrs_shaft_sizer.py`
   - `requirements.txt`
4. Type a commit message: `Initial upload — HRS shaft sizer app`
5. Click **"Commit changes"**

**Option B — Git Command Line:**

```bash
# Clone the repo
git clone https://github.com/usdtxdprinter-ux/hrs-shaft-sizer.git
cd hrs-shaft-sizer

# Copy your files into the repo folder
# (put hrs_shaft_sizer.py and requirements.txt here)

# Commit and push
git add .
git commit -m "Initial upload — HRS shaft sizer app"
git push origin main
```

---

### STEP 3 — Deploy to Streamlit Cloud

1. Go to **https://share.streamlit.io**
2. Click **"Sign in with GitHub"** and authorize with your `usdtxdprinter-ux` account
3. Click **"New app"** (top right)
4. Fill in the deployment form:

   | Field | Value |
   |---|---|
   | **Repository** | `usdtxdprinter-ux/hrs-shaft-sizer` |
   | **Branch** | `main` |
   | **Main file path** | `hrs_shaft_sizer.py` |

5. Click **"Deploy!"**

Streamlit will install the dependencies from `requirements.txt` and launch your app.
This typically takes 2-3 minutes on first deploy.

---

### STEP 4 — Access Your App

Once deployed, your app URL will be:
```
https://usdtxdprinter-ux-hrs-shaft-sizer-hrs-shaft-sizer-XXXXX.streamlit.app
```

You can also set a **custom subdomain** in Streamlit Cloud settings:
1. Click the **⋮** menu on your app in the Streamlit dashboard
2. Click **"Settings"**
3. Under **"General"** → set custom URL to something like:
```
https://hrs-shaft-sizer.streamlit.app
```

---

## Sharing the App

Once deployed, you can share the URL with anyone. No login required for viewers.

To share with your engineering team or customers:
- Send them the direct Streamlit URL
- Embed it in an iframe on your website:
  ```html
  <iframe
    src="https://hrs-shaft-sizer.streamlit.app"
    width="100%"
    height="800"
    frameborder="0">
  </iframe>
  ```

---

## Updating the App

Any time you push changes to the `main` branch on GitHub, Streamlit Cloud
will automatically redeploy within ~60 seconds.

**To update via GitHub web UI:**
1. Go to `https://github.com/usdtxdprinter-ux/hrs-shaft-sizer`
2. Click on `hrs_shaft_sizer.py`
3. Click the **pencil icon** (✏️) to edit
4. Make your changes
5. Click **"Commit changes"**
6. Streamlit will auto-redeploy

---

## Making the Repo Private (Optional)

Streamlit Community Cloud free tier requires **public** repos.
If you need a private repo, you have two options:

1. **Streamlit Teams plan** ($250/month) — supports private repos
2. **Self-host** — run on your own server:
   ```bash
   pip install streamlit pandas
   streamlit run hrs_shaft_sizer.py --server.port 8501
   ```

---

## Local Testing (Before Deploying)

To test locally on your machine first:

```bash
# Install Python 3.9+ if not already installed
# https://www.python.org/downloads/

# Install dependencies
pip install streamlit pandas

# Run the app
streamlit run hrs_shaft_sizer.py
```

This will open `http://localhost:8501` in your browser.

---

## Troubleshooting

| Issue | Solution |
|---|---|
| "Module not found" error | Make sure `requirements.txt` is in the repo root |
| App won't load | Check Streamlit Cloud logs (click "Manage app" → "Logs") |
| Blank page | Clear browser cache, or try incognito mode |
| "st.rerun" error | Make sure you're using Streamlit ≥ 1.32.0 (in requirements.txt) |
| Changes not showing | Push to `main` branch; Streamlit auto-deploys from `main` only |

---

## Repository Structure

```
hrs-shaft-sizer/
├── hrs_shaft_sizer.py      ← Main Streamlit app (1016 lines)
├── requirements.txt         ← Python dependencies
└── README.md                ← Optional project description
```

That's it — two files is all Streamlit Cloud needs.
