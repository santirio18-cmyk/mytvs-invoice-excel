# myTVS — Invoice PDF → Excel

Public tool (no login). Upload up to **10 invoices** (PDF or clear scan) and download **CSV + Excel**.

Line items: **Part Number, Description, HSN/SAC, Qty, Rate, Amount**  
Sheet names: `Invoice No · Supplier · Date`

---

## Local development

### 1. Backend (port 8000)

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# macOS: brew install tesseract
uvicorn main:app --reload --port 8000
```

### 2. Frontend (port 3000)

```bash
cd frontend
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

---

## Deploy publicly (Option B)

### A. Backend → [Railway](https://railway.app)

1. Push this repo to GitHub (or deploy from local).
2. Railway → **New Project** → **Deploy from GitHub**.
3. Set **Root Directory** to `backend`.
4. Railway will use `backend/Dockerfile` (includes Tesseract OCR).
5. After deploy, copy the public URL, e.g.  
   `https://mytvs-invoice-api.up.railway.app`
6. Check health: open `https://YOUR-RAILWAY-URL/api/health` → should show `{"status":"ok"}`.

Optional Railway variables:

| Variable | Value |
|----------|--------|
| `CORS_ORIGINS` | `https://YOUR-VERCEL-APP.vercel.app,http://localhost:3000` |

### B. Frontend → [Vercel](https://vercel.com)

1. Vercel → **Add New Project** → import the same GitHub repo.
2. Set **Root Directory** to `frontend`.
3. Add environment variable:

| Name | Value |
|------|--------|
| `NEXT_PUBLIC_API_URL` | `https://YOUR-RAILWAY-URL` (no trailing slash) |

4. Deploy. Share the Vercel URL with your team.

### C. Connect them

1. Deploy backend first → get Railway URL.  
2. Set `NEXT_PUBLIC_API_URL` on Vercel → redeploy frontend.  
3. Optionally set `CORS_ORIGINS` on Railway to your Vercel domain → redeploy backend.  
4. Test: open Vercel site → upload a sample invoice → Generate Excel / CSV.

---

## Notes

- Prefer clear digital PDFs or straight scans (see Do’s / Don’ts in the UI).
- Max 10 files · 20 MB each.
- Blurry phone photos may miss line items.
