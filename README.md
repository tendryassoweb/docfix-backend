# DocFix Backend — Impulse AI
# Déploiement sur Render

FastAPI + python-docx + LibreOffice + Gemini Flash

---

## Structure

```
backend/
├── app/
│   ├── __init__.py
│   ├── main.py        ← FastAPI app + routes
│   ├── config.py      ← Variables d'environnement
│   ├── jobs.py        ← Gestion des jobs en mémoire
│   └── processor.py   ← Moteur DOCX (12 étapes)
├── requirements.txt
├── Dockerfile
├── render.yaml
└── .env.example
```

## Endpoints

| Méthode | Route | Description |
|---------|-------|-------------|
| GET  | `/` | Info API |
| GET  | `/health` | Health check |
| POST | `/api/v1/process` | Upload + lancer traitement |
| POST | `/api/v1/webhook/process` | Idem (entrée n8n) |
| GET  | `/api/v1/status/{job_id}` | Polling du statut |
| GET  | `/api/v1/download/{job_id}/{format}` | Télécharger docx ou pdf |

---

## Test local

### Prérequis
- Python 3.11+
- LibreOffice installé (`brew install libreoffice` sur macOS)
- Clé API Gemini (https://aistudio.google.com/app/apikey)

```bash
# 1. Créer l'environnement virtuel
cd backend
python -m venv venv
source venv/bin/activate        # macOS/Linux
# venv\Scripts\activate         # Windows

# 2. Installer les dépendances
pip install -r requirements.txt

# 3. Configurer l'environnement
cp .env.example .env
# Éditer .env :
#   GEMINI_API_KEY=votre_cle
#   LIBREOFFICE_PATH=/Applications/LibreOffice.app/Contents/MacOS/soffice  # macOS
#   ALLOWED_ORIGINS=http://localhost:3000

# 4. Lancer le serveur
uvicorn app.main:app --reload --port 8000
```

L'API est disponible sur `http://localhost:8000`
Documentation Swagger : `http://localhost:8000/docs`

### Test rapide avec curl
```bash
# Upload un fichier
curl -X POST http://localhost:8000/api/v1/process \
  -F "file=@mon_document.docx"
# → {"job_id": "uuid-xxx", "status": "queued"}

# Polling
curl http://localhost:8000/api/v1/status/uuid-xxx

# Télécharger le résultat
curl -O http://localhost:8000/api/v1/download/uuid-xxx/docx
curl -O http://localhost:8000/api/v1/download/uuid-xxx/pdf
```

---

## Déploiement sur Render

### Option 1 — Via GitHub (recommandé)

1. **Pousser le code sur GitHub**
   ```bash
   git init
   git add .
   git commit -m "feat: docfix backend impulse ai"
   git remote add origin https://github.com/votre-org/docfix-backend.git
   git push -u origin main
   ```

2. **Créer un service sur Render**
   - Aller sur https://dashboard.render.com/new/web
   - Cliquer **New Web Service**
   - Connecter votre repo GitHub
   - Sélectionner le dossier `backend/` comme **Root Directory**
   - Render détecte automatiquement le `Dockerfile`

3. **Configurer les paramètres**
   ```
   Name     : docfix-api
   Region   : Frankfurt (EU) — recommandé pour la Madagascar
   Branch   : main
   Runtime  : Docker
   Plan     : Starter (7$/mois) ou Standard (25$/mois)
   ```

4. **Ajouter les variables d'environnement**
   Dans **Environment > Environment Variables** :
   ```
   GEMINI_API_KEY        = [votre clé Gemini]
   ALLOWED_ORIGINS       = https://votre-app.vercel.app,http://localhost:3000
   GEMINI_MODEL          = gemini-1.5-flash
   MAX_FILE_SIZE_MB      = 50
   JOB_EXPIRY_SECONDS    = 7200
   TEMP_DIR              = /tmp/docfix
   LIBREOFFICE_PATH      = /usr/bin/soffice
   DEBUG                 = false
   ```

5. **Cliquer "Create Web Service"**
   - Le premier build prend ~5 min (installation LibreOffice)
   - URL finale : `https://docfix-api.onrender.com`

6. **Mettre à jour le frontend**
   Dans Vercel, mettre à jour :
   ```
   NEXT_PUBLIC_API_URL = https://docfix-api.onrender.com
   ```

### Option 2 — Via render.yaml (Infrastructure as Code)

```bash
# Installer le CLI Render
npm install -g @render-com/cli

# Déployer depuis le render.yaml
render deploy
```

---

## Choix du plan Render

| Plan | RAM | Prix | Recommandé pour |
|------|-----|------|-----------------|
| Free | 512 Mo | 0$ | Tests (dort après 15min) |
| Starter | 512 Mo | 7$/mois | MVP, faible trafic |
| Standard | 2 Go | 25$/mois | **Production — LibreOffice a besoin de RAM** |

> ⚠️ LibreOffice consomme ~300-500 Mo lors d'une conversion PDF.
> Le plan **Standard** est fortement recommandé en production.

---

## Logs & Monitoring

```bash
# Voir les logs en temps réel (CLI Render)
render logs --service docfix-api --tail

# Ou dans le dashboard : Service → Logs
```

---

## Architecture complète

```
[Frontend Vercel]
      │ POST /api/v1/process (multipart)
      ▼
[FastAPI sur Render]
      │ BackgroundTask
      ▼
[processor.py]
   ├── Gemini Flash API (analyse IA)
   ├── python-docx (12 corrections)
   └── LibreOffice headless (→ PDF)
      │
      ▼ polling GET /api/v1/status/{id}
[Frontend] ← résultat
      │ GET /api/v1/download/{id}/docx|pdf
      ▼
[Téléchargement navigateur]
```
