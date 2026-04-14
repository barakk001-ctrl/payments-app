# paymentsProject

A web app for visualizing Israeli credit card statements (Cal / Isracard `.xlsx` files).

## Features
- Upload monthly `.xlsx` or saved `.json` files
- Interactive dashboard: charts, category breakdown, top merchants, duplicates, subscriptions, installments
- Two-month comparison view
- Smart insights
- Save as HTML or JSON for offline use
- Yearly summary mode

---

## Deploy to Railway (recommended)

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
3. Select your repo — Railway auto-detects Python and uses the `Procfile`
4. Your app will be live at a `*.railway.app` URL in ~2 minutes

---

## Deploy to Render

1. Push to GitHub
2. Go to [render.com](https://render.com) → **New Web Service** → connect your repo
3. Render picks up `render.yaml` automatically
4. Free tier spins down after 15 min of inactivity (first request after sleep is slow)

---

## Deploy to PythonAnywhere

1. Sign up at [pythonanywhere.com](https://pythonanywhere.com)
2. Open a Bash console and run:
   ```bash
   git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
   cd YOUR_REPO
   pip install -r requirements.txt
   ```
3. Go to **Web** tab → **Add a new web app** → **Manual configuration** → Python 3.11
4. Set the WSGI file source path to your project folder and add:
   ```python
   import sys
   sys.path.insert(0, '/home/YOUR_USERNAME/YOUR_REPO')
   from payments_server import app as application
   ```
5. Click **Reload**

---

## Run locally

```bash
pip install -r requirements.txt
python payments_server.py --open    # opens browser automatically
```

For the yearly summary:
```bash
python payments_yearly.py ./folder-with-xlsx-files/ --open
```
