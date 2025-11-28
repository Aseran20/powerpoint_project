# Office Add-in Development - Best Practices

## Problème Rencontré (Excel Add-in)

### ❌ Ce qui ne marchait pas
- **Manifest fait manuellement** : Configuration Shared Runtime incomplète
- **Webpack personnalisé** : Ne suivait pas les conventions Microsoft
- **React + Office.js** : Mélange non-officiel causant des conflits
- **Résultat** : `#NAME?` dans les Custom Functions

### ✅ Solution
Utiliser le **template officiel** `yo office` qui génère :
- Le manifest exact avec toutes les conventions Microsoft
- La configuration webpack optimale
- Le wiring correct entre taskpane et custom functions

## Guide : Créer un Office Add-in (Excel, PowerPoint, etc.)

### Prérequis
```powershell
# Installer Yeoman et le générateur Office
npm install -g yo generator-office
```

### Commandes pour Différents Types de Projets

#### 1. Excel Add-in avec Custom Functions (Shared Runtime)
```powershell
# Navigation
cd mon-projet

# Génération (interactif)
yo office

# Sélections :
# - Project type: Excel Custom Functions using a Shared Runtime
# - Script type: TypeScript (ou JavaScript)
# - Name: mon-addin-excel
```

**OU en non-interactif :**
```powershell
npx --package yo --package generator-office yo office \
  --projectType excel-functions-shared \
  --name mon-addin-excel \
  --host excel \
  --ts \
  --skip-install
```

#### 2. PowerPoint Add-in (Taskpane)
```powershell
# Interactif
yo office

# Sélections :
# - Project type: Office Add-in Task Pane project
# - Host: PowerPoint
# - Script type: TypeScript
# - Name: mon-addin-powerpoint
```

**OU en non-interactif :**
```powershell
npx --package yo --package generator-office yo office \
  --projectType taskpane \
  --name mon-addin-powerpoint \
  --host powerpoint \
  --ts \
  --skip-install
```

#### 3. Word Add-in (Taskpane)
```powershell
npx --package yo --package generator-office yo office \
  --projectType taskpane \
  --name mon-addin-word \
  --host word \
  --ts \
  --skip-install
```

### Types de Projets Disponibles

| `--projectType` | Description | Hosts compatibles |
|-----------------|-------------|-------------------|
| `taskpane` | Volet latéral standard | Excel, Word, PowerPoint, Outlook |
| `excel-functions` | Custom Functions (runtime séparé) | Excel uniquement |
| `excel-functions-shared` | Custom Functions + Taskpane (Shared Runtime) | Excel uniquement |
| `react` | Taskpane avec React | Excel, Word, PowerPoint |
| `manifest-only` | Juste le manifest (pour projet existant) | Tous |

### Structure Générée (Exemple Excel Shared Runtime)

```
mon-addin-excel/
├── manifest.xml          # Configuration de l'add-in
├── webpack.config.js     # Build configuration
├── package.json
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html
│   │   ├── taskpane.ts
│   │   └── taskpane.css
│   ├── functions/
│   │   └── functions.ts  # Custom functions
│   └── commands/
│       └── commands.ts   # Ribbon commands
└── assets/
    └── icon-*.png
```

## Workflow de Développement

### 1. Génération
```powershell
# Créer le dossier
mkdir mon-projet-office
cd mon-projet-office

# Générer avec yo office
yo office
# OU
npx --package yo --package generator-office yo office --projectType taskpane --host powerpoint --ts

# Installer les dépendances
npm install
```

### 2. Personnalisation

#### A. Modifier le Manifest
Éditez `manifest.xml` pour changer :
- `<Id>` : GUID unique (générez-en un nouveau avec `uuidgen` ou en ligne)
- `<DisplayName>` : Nom affiché dans Office
- `<Description>` : Description de votre add-in
- URLs de production (remplacer `https://www.contoso.com/`)

#### B. Ajouter Votre Logique
- **Taskpane** : Modifiez `src/taskpane/taskpane.ts` et `taskpane.html`
- **Custom Functions** : Ajoutez vos fonctions dans `src/functions/functions.ts`
- **Backend** : Configurez les URLs backend via webpack (voir exemple AlgoSheet)

### 3. Développement Local
```powershell
# Terminal 1: Backend (si applicable)
cd ../backend
npm run dev

# Terminal 2: Add-in
npm run dev-server  # Démarre sur https://localhost:3000

# Terminal 3: Lancer Office (optionnel)
npm start  # Lance Excel/PowerPoint/Word avec l'add-in
```

### 4. Test sur Office Web
1. Ouvrir [office.com](https://office.com)
2. Créer/ouvrir un document
3. **Insertion** > **Compléments** > **Gérer mes compléments** > **Télécharger mon complément**
4. Sélectionner `manifest.xml`
5. **Important** : Accepter les certificats localhost (`https://localhost:3000/taskpane.html`)

### 5. Build Production
```powershell
npm run build

# Le dossier dist/ contient :
# - manifest.xml (avec URLs de prod)
# - Tous les fichiers JS/HTML/CSS optimisés
```

## Configuration Backend URL (Pattern AlgoSheet)

### webpack.config.js
```javascript
const webpack = require("webpack");

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const urlProd = "https://mon-domaine.com/";
  
  return {
    // ... autres configs
    plugins: [
      // ... autres plugins
      new webpack.DefinePlugin({
        "process.env.BACKEND_URL": JSON.stringify(
          dev ? "https://localhost:3100/api" : "https://api.mon-domaine.com/api"
        ),
      }),
    ],
  };
};
```

### Dans votre code TypeScript
```typescript
async function callBackend(data: any) {
  // @ts-ignore
  const backendUrl = process.env.BACKEND_URL || "https://api.mon-domaine.com/api";
  
  const response = await fetch(backendUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  
  return await response.json();
}
```

## Déploiement

### VPS (Nginx)
```bash
# Sur le serveur
cd /var/www/mon-addin

# Copier les fichiers de dist/
# Assurer que manifest.xml pointe vers https://mon-domaine.com/

# Configuration Nginx
server {
    listen 443 ssl;
    server_name mon-domaine.com;
    
    ssl_certificate /etc/letsencrypt/live/mon-domaine.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/mon-domaine.com/privkey.pem;
    
    root /var/www/mon-addin;
    
    location / {
        try_files $uri $uri/ =404;
        add_header Access-Control-Allow-Origin *;
    }
}
```

## Règles d'Or ✨

1. **Toujours partir du template officiel** (`yo office`)
2. **Ne pas modifier la structure webpack** sauf si nécessaire
3. **Tester sur Excel/PowerPoint Web** pour un débogage facile (F12)
4. **Utiliser Shared Runtime** pour Excel si vous avez Taskpane + Custom Functions
5. **HTTPS obligatoire** en production
6. **Garder le même GUID** entre dev et prod (dans `<Id>`)

## Troubleshooting Commun

| Problème | Cause | Solution |
|----------|-------|----------|
| `#NAME?` dans Excel | Custom Functions non chargées | Vérifier Shared Runtime, vider cache Excel, tester sur Web |
| CORS errors | Backend refuse les requêtes | Ajouter headers CORS sur le backend |
| Certificat non valide | Localhost certificate | Ouvrir `https://localhost:3000/taskpane.html` et accepter |
| Add-in ne se charge pas | Manifest invalide | Valider avec `npm run validate` |
| Port 3000 utilisé | Autre serveur actif | `netstat -ano \| findstr :3000` puis tuer le process |

## Ressources

- [Documentation Microsoft](https://learn.microsoft.com/office/dev/add-ins/)
- [Yo Office GitHub](https://github.com/OfficeDev/generator-office)
- [Samples Officiels](https://github.com/OfficeDev/Office-Add-in-samples)

🛠️ Comment Faire
PowerPoint Add-in :

powershell
# Navigation vers votre projet
cd mon-projet-powerpoint

# Génération (choisir les options interactivement)
yo office

# OU directement en CLI
npx --package yo --package generator-office yo office \
  --projectType taskpane \
  --name mon-addin-powerpoint \
  --host powerpoint \
  --ts \
  --skip-install

# Installer et démarrer
npm install
npm run dev-server