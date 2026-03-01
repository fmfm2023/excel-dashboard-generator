# 📊 Excel Dashboard Generator — Résumé du Projet

## 🗓️ Date de création
01 Mars 2026

---

## 🌐 URLs importantes

| Ressource | URL |
|-----------|-----|
| **Formulaire client** | https://fmfm2023.github.io/excel-dashboard-generator/ |
| **API Railway** | https://web-production-81cbc.up.railway.app |
| **Health check API** | https://web-production-81cbc.up.railway.app/health |
| **GitHub repo** | https://github.com/fmfm2023/excel-dashboard-generator |
| **Make.com** | https://make.com |
| **Railway dashboard** | https://railway.app |

---

## 🔑 Clés & IDs

| Élément | Valeur |
|---------|--------|
| **Webhook Make.com** | https://hook.eu1.make.com/3s5l7wakmtiq16aem4ftdpt4u276fzxj |
| **Email Gmail** | faycel.mastour@gmail.com |
| **GitHub user** | fmfm2023 |

---

## 🏗️ Architecture du système

```
Client
  │
  ▼
📄 Formulaire HTML (GitHub Pages)
  │  Upload fichier CSV/Excel + email
  ▼
🔗 Webhook Make.com
  │  Reçoit {filename, file_data, email, file_type}
  ▼
🌐 HTTP Module → API Python (Railway)
  │  POST /generate-dashboard
  ▼
🐍 Flask API (generate-dashboard.py)
  │  - Lit le fichier (CSV ou Excel)
  │  - Détecte les colonnes automatiquement
  │  - Génère dashboard Excel (3 onglets)
  │  - Retourne {status, excel_base64, kpis}
  ▼
🔀 Router Make.com
  ├── ✅ Succès → Gmail 6 : envoie dashboard Excel
  └── ❌ Erreur  → Gmail 7 : envoie email d'erreur
```

---

## 📁 Fichiers du projet

| Fichier | Rôle |
|---------|------|
| `generate-dashboard.py` | API Flask Python (895 lignes) |
| `requirements.txt` | Dépendances Python |
| `Procfile` | Config Railway/gunicorn |
| `blueprint-make.json` | Scénario Make.com (à importer) |
| `index.html` / `upload-form.html` | Formulaire client |
| `email-template.html` | Template email HTML |
| `test-api.ps1` | Script test PowerShell local |
| `README.md` | Guide installation complet |

---

## 🚀 Comment tester

### Test local (API Python)
```powershell
# Démarrer le serveur
C:\Users\fayce\AppData\Local\Python\bin\python.exe generate-dashboard.py

# Tester avec PowerShell
.\test-api.ps1 -ApiUrl "http://localhost:5000" -FilePath ".\ventes_janvier.xlsx"
```

### Test production
```powershell
# Vérifier que l'API Railway tourne
curl https://web-production-81cbc.up.railway.app/health
```

### Test complet end-to-end
1. Aller sur https://fmfm2023.github.io/excel-dashboard-generator/
2. Uploader un fichier CSV ou Excel
3. Entrer son email
4. Cliquer Envoyer
5. Recevoir le dashboard par email en ~60 secondes

---

## ⚙️ Make.com — Scénario

**Nom :** Excel Dashboard Generator
**Modules :**
1. **Webhooks** — Reçoit les fichiers uploadés
2. **HTTP (legacy)** — Appelle l'API Railway
3. **Router** — Route selon succès/erreur
4. **Gmail 6** — Envoie le dashboard (succès)
5. **Gmail 7** — Envoie l'erreur (échec)

**Important :** Garder le toggle "Immediately as data arrives" sur **ON**

---

## 💰 Coûts mensuels

| Service | Gratuit jusqu'à | Coût après |
|---------|----------------|------------|
| GitHub Pages | Illimité | 0€ |
| Make.com | 1 000 opérations/mois | ~9€/mois |
| Railway | ~500h/mois | ~5$/mois |
| Gmail | Illimité | 0€ |

---

## 🔧 Modifications courantes

### Changer l'email destinataire par défaut
→ Modifier les modules Gmail 6 et 7 dans Make.com

### Changer l'URL de l'API
→ Module HTTP dans Make.com → champ URL

### Modifier le style du dashboard Excel
→ `generate-dashboard.py` → fonctions `build_dashboard_sheet()` et `build_charts()`

### Modifier le formulaire client
→ `index.html` → commiter et pusher sur GitHub → GitHub Pages se met à jour automatiquement

---

## 📞 Reprendre avec Claude

Pour reprendre ce projet avec Claude Code :
1. Ouvrir un terminal dans `C:\Users\fayce\Desktop\test-claude`
2. Lancer `claude`
3. Dire : *"Reprends le projet Excel Dashboard Generator, voir PROJET.md"*
