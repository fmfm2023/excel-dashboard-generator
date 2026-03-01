# 📊 Excel Dashboard Generator — Guide d'installation complet

> Génère automatiquement des dashboards Excel professionnels (KPIs, graphiques natifs, TCD)
> à partir de fichiers CSV/XLSX uploadés via un formulaire web ou email.

---

## 🗂️ Structure du projet

```
excel-dashboard-generator/
├── blueprint-make.json       ← Scénario Make.com (importable)
├── generate-dashboard.py     ← API Flask + génération Excel (openpyxl)
├── requirements.txt          ← Dépendances Python
├── upload-form.html          ← Formulaire web drag & drop
├── email-template.html       ← Template email de livraison
└── README.md                 ← Ce guide
```

---

## 🏗️ Architecture globale

```
Utilisateur
    │
    ▼ Upload fichier + email
┌─────────────────┐
│  upload-form.html│  (hébergé sur GitHub Pages / Netlify / votre serveur)
└────────┬────────┘
         │ POST JSON {filename, file_data (base64), email}
         ▼
┌─────────────────┐
│   Make.com      │  Webhook → HTTP call → Router → Gmail
│   Scénario      │
└────────┬────────┘
         │ POST /generate-dashboard
         ▼
┌─────────────────────────────┐
│  generate-dashboard.py      │  Flask API (Railway / Render / VPS)
│  ─ Charge CSV ou XLSX       │
│  ─ Détecte colonnes auto    │
│  ─ Calcule KPIs             │
│  ─ Génère Excel (openpyxl)  │
│  ─ Retourne base64          │
└─────────────────────────────┘
         │ {status, excel_base64, kpis}
         ▼
┌─────────────────┐
│   Gmail/Outlook │  Envoie email + fichier Excel en pièce jointe
└─────────────────┘
         │
         ▼
    Utilisateur reçoit son Dashboard Excel 📊
```

---

## ⚙️ ÉTAPE 1 — Déployer l'API Python

### Option A : Railway (recommandé, gratuit)

1. Créez un compte sur [railway.app](https://railway.app)
2. Cliquez **New Project → Deploy from GitHub Repo**
3. Uploadez les fichiers `generate-dashboard.py` et `requirements.txt`
4. Ajoutez le fichier `Procfile` avec ce contenu :
   ```
   web: gunicorn generate-dashboard:app --bind 0.0.0.0:$PORT
   ```
5. Railway détecte Python automatiquement et installe les dépendances
6. Notez l'URL générée, ex : `https://excel-dashboard.up.railway.app`
7. Testez : `GET https://excel-dashboard.up.railway.app/health` → `{"status": "ok"}`

### Option B : Render (gratuit)

1. Créez un compte sur [render.com](https://render.com)
2. **New → Web Service → Upload files**
3. Start command : `gunicorn generate-dashboard:app`
4. Notez l'URL publique

### Option C : VPS / Serveur local

```bash
# 1. Cloner les fichiers
mkdir dashboard-api && cd dashboard-api
# Copiez generate-dashboard.py et requirements.txt ici

# 2. Créer environnement virtuel
python3 -m venv venv
source venv/bin/activate       # Linux/Mac
# venv\Scripts\activate        # Windows

# 3. Installer les dépendances
pip install -r requirements.txt

# 4. Lancer en production
gunicorn generate-dashboard:app --bind 0.0.0.0:5000 --workers 2 --timeout 120

# 5. (Optionnel) Lancer en développement
python generate-dashboard.py
```

> Sur **Windows** en production, remplacez `gunicorn` par :
> ```bash
> waitress-serve --host=0.0.0.0 --port=5000 generate-dashboard:app
> ```

### Test rapide de l'API

```bash
curl -X POST https://VOTRE-URL/generate-dashboard \
  -H "Content-Type: application/json" \
  -d '{
    "filename": "test.csv",
    "file_data": "Q29sdW1uMSxDb2x1bW4yLENvbHVtbjMKQSw0MSwyMDI0LTAxLTAxCkIsMzIsMjAyNC0wMS0wMgpDLDU1LDIwMjQtMDEtMDM=",
    "email": "test@example.com"
  }'
```

Réponse attendue : `{"status": "success", "excel_base64": "...", "kpis": {...}}`

---

## ⚙️ ÉTAPE 2 — Importer le scénario Make.com

1. Connectez-vous à [make.com](https://make.com)
2. Allez dans **Scenarios → Create a new scenario**
3. Cliquez sur les **trois points (...)** en bas → **Import Blueprint**
4. Uploadez le fichier `blueprint-make.json`
5. Le scénario s'ouvre avec tous les modules préconfigurés

### Modules du scénario dans l'ordre :

| # | Module | Rôle |
|---|--------|------|
| 1 | **Webhook** (CustomWebHook) | Reçoit le fichier de l'utilisateur |
| 2 | **Basic Feeder** | Extrait le corps de la requête |
| 3 | **HTTP: Make a Request** | Appelle l'API Python pour générer l'Excel |
| 4 | **Set Variable** | Stocke le statut de la réponse |
| 5 | **Router** | Branche selon succès ou erreur |
| 6 | **Gmail: Send Email** | Envoie le dashboard (succès) |
| 7 | **Gmail: Send Email** | Envoie email d'erreur (échec) |

---

## ⚙️ ÉTAPE 3 — Configurer le Webhook

1. Dans Make.com, cliquez sur le **module Webhook** (module 1)
2. Cliquez **Add** → nommez-le `Excel Dashboard Webhook`
3. Copiez l'URL générée, ex :
   ```
   https://hook.eu1.make.com/abc123xyz789
   ```
4. **Gardez cette URL** — vous en aurez besoin à l'étape 5

---

## ⚙️ ÉTAPE 4 — Configurer le module HTTP (module 3)

1. Cliquez sur le **module HTTP: Make a Request** (module 3)
2. Remplacez `{{YOUR_PYTHON_API_URL}}` par l'URL de votre API Python :
   ```
   https://excel-dashboard.up.railway.app/generate-dashboard
   ```
3. Vérifiez que la méthode est **POST** et le type de body **JSON (Raw)**
4. Le body est déjà mappé pour transmettre `filename`, `file_data`, `email`

---

## ⚙️ ÉTAPE 5 — Connecter Gmail

1. Cliquez sur le **module Gmail: Send Email** (module 6)
2. Cliquez sur **Add** à côté de "Connection"
3. Choisissez **Gmail** et connectez votre compte Google
4. Accordez les permissions demandées
5. Faites de même pour le module 7 (email d'erreur)

> 💡 **Outlook** : Remplacez les modules Gmail par **Microsoft 365 Email: Send an Email**
> et reconnectez votre compte Microsoft.

---

## ⚙️ ÉTAPE 6 — Activer le scénario

1. Cliquez sur le bouton **ON/OFF** en bas à gauche pour activer le scénario
2. Vérifiez que le statut passe à **Active**
3. Définissez le scheduling sur **Instant** (le webhook le déclenche à la demande)

---

## ⚙️ ÉTAPE 7 — Configurer le formulaire d'upload

Ouvrez `upload-form.html` dans un éditeur et modifiez la ligne suivante :

```javascript
// Ligne ~200 du fichier upload-form.html
const WEBHOOK_URL = 'https://hook.eu1.make.com/VOTRE_WEBHOOK_ID';
//                                              ^^^^^^^^^^^^^^^^
//                   Remplacez par l'URL copiée à l'étape 3
```

---

## ⚙️ ÉTAPE 8 — Héberger le formulaire d'upload

### Option A : GitHub Pages (gratuit, recommandé)

```bash
# 1. Créez un repo GitHub (ex: "dashboard-upload")
# 2. Uploadez upload-form.html à la racine
# 3. Settings → Pages → Source: "main branch / root"
# 4. Votre URL : https://VOTRE_USERNAME.github.io/dashboard-upload/upload-form.html
```

### Option B : Netlify (gratuit, drag & drop)

1. Allez sur [netlify.com](https://netlify.com)
2. Faites glisser le fichier `upload-form.html` sur le dashboard
3. Netlify génère une URL en quelques secondes

### Option C : Serveur local (test)

```bash
# Python
python3 -m http.server 8080
# Puis ouvrez : http://localhost:8080/upload-form.html
```

---

## 🧪 Test complet end-to-end

1. **Ouvrez** le formulaire hébergé dans votre navigateur
2. **Glissez** un fichier CSV ou Excel (exemple : `ventes.csv`)
3. **Entrez** votre adresse email
4. **Cliquez** sur "Générer mon Dashboard"
5. **Attendez** 1-3 minutes
6. **Vérifiez** votre boîte email → vous devez recevoir `Dashboard_ventes.xlsx`

### Fichier CSV de test

Copiez ce contenu dans un fichier `test.csv` :

```csv
Date,Produit,Catégorie,Vendeur,Montant,Quantité
2024-01-05,Laptop Pro,Informatique,Alice,1299.99,2
2024-01-08,Souris Ergonomique,Accessoires,Bob,49.90,5
2024-01-12,Écran 4K,Informatique,Alice,599.00,1
2024-02-03,Clavier Mécanique,Accessoires,Charlie,129.00,3
2024-02-14,Laptop Pro,Informatique,Bob,1299.99,1
2024-02-20,Webcam HD,Accessoires,Alice,89.00,4
2024-03-01,Serveur NAS,Stockage,Charlie,799.00,1
2024-03-15,SSD 2To,Stockage,Bob,219.00,6
2024-03-22,Laptop Pro,Informatique,Charlie,1299.99,2
```

---

## 🛠️ Dépannage

### Le webhook ne répond pas
- Vérifiez que le scénario est **actif** dans Make.com
- Vérifiez que l'URL du webhook est correcte dans `upload-form.html`
- Testez via Postman ou curl directement

### L'API Python retourne une erreur 500
- Consultez les logs Railway/Render
- Vérifiez que toutes les dépendances sont installées
- Testez la route `/health` pour confirmer que le serveur tourne

### Email non reçu
- Vérifiez le dossier **Spam/Indésirables**
- Dans Make.com, ouvrez le scénario et regardez les **Executions** pour voir si l'email a été envoyé
- Vérifiez que la connexion Gmail est active (non expirée)

### Graphiques vides dans Excel
- Les graphiques nécessitent **Microsoft Excel 2016+**
- Dans Google Sheets, les graphiques natifs Excel s'affichent en lecture seule
- Vérifiez que votre fichier contient au moins une colonne numérique et une colonne catégorielle

### Fichier trop lourd
- Le formulaire limite à 50 Mo
- Make.com limite les webhooks à 5 Mo par défaut (plan gratuit)
- Solution : activez la compression ou utilisez la route `/generate-from-upload` avec multipart

---

## 🔐 Sécurité (production)

```python
# Ajoutez dans generate-dashboard.py pour sécuriser l'API :
import hmac, hashlib, os

API_KEY = os.environ.get('API_KEY', 'changez-moi')

@app.before_request
def check_api_key():
    if request.path == '/health':
        return  # Route publique
    key = request.headers.get('X-API-Key', '')
    if not hmac.compare_digest(key, API_KEY):
        return jsonify({'error': 'Unauthorized'}), 401
```

Puis dans Make.com, ajoutez le header `X-API-Key: VOTRE_CLE` dans le module HTTP.

---

## 📦 Variables d'environnement (optionnel)

Créez un fichier `.env` à la racine du projet :

```env
PORT=5000
DEBUG=false
API_KEY=votre_cle_secrete_ici
```

Et chargez-le en ajoutant en début de `generate-dashboard.py` :
```python
from dotenv import load_dotenv
load_dotenv()
```

---

## 🚀 Extensions possibles

| Idée | Comment |
|------|---------|
| **Slack notification** | Ajoutez un module Slack dans Make.com après le Router |
| **Sauvegarde Google Drive** | Ajoutez module "Google Drive: Upload a File" |
| **Trigger par email** | Remplacez le Webhook par "Gmail: Watch Emails" |
| **Planification** | Ajoutez un module "Google Sheets: Watch Rows" comme trigger |
| **Multi-langues** | Passez le paramètre `lang` dans le payload JSON |
| **Logo personnalisé** | Ajoutez votre logo en base64 dans le script Python |

---

## 📋 Checklist de déploiement

- [ ] API Python déployée et `/health` répond 200
- [ ] URL de l'API notée
- [ ] Scénario Make.com importé depuis `blueprint-make.json`
- [ ] Webhook créé et URL copiée
- [ ] Module HTTP configuré avec l'URL de l'API
- [ ] Gmail connecté sur les deux modules email
- [ ] Scénario activé (statut ON)
- [ ] `upload-form.html` mis à jour avec l'URL du webhook
- [ ] Formulaire hébergé (GitHub Pages / Netlify)
- [ ] Test end-to-end réussi avec `test.csv`
- [ ] Email reçu avec dashboard Excel en pièce jointe ✅

---

*Généré le 28/02/2026 — Excel Dashboard Generator v1.0.0*
