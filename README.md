# CEN-Mail

Extension Thunderbird pour la gestion des emails — CEN Nouvelle-Aquitaine

## Fonctionnalités

| Module | Description |
|--------|-------------|
| **M365** | Authentification OAuth2 Microsoft, connexion Graph API |
| **Étiquettes** | Mapping catégories Outlook ↔ labels Thunderbird (couleurs) |
| **Migration** | Migration batch d'emails entre dossiers (checkpoints, reprise) |
| **Synchronisation** | Analyse et synchro source/destination en 4 phases |
| **Export** | Téléchargement et export de messages |
| **Tags** | Gestion des tags par message |

## Prérequis

- Mozilla Thunderbird 128.0+ (Manifest v3)
- Compte Microsoft 365 (pour le module M365)
- Windows 11 / Linux / macOS

## Installation

1. Ouvrir Thunderbird
2. Menu **Outils → Modules complémentaires**
3. Roue dentée → **Installer un module depuis un fichier**
4. Sélectionner `mail-cen-v5.3.xpi`

## Stack technique

| Composant | Technologie |
|-----------|-------------|
| Type | Extension Thunderbird (Manifest v3) |
| Langage | JavaScript ES2020+ |
| APIs | Thunderbird Messenger API, Microsoft Graph API |
| Auth | OAuth2 (Azure / Entra ID) |
| Stockage | browser.storage.local |
| Dépendances externes | Aucune |

## Structure du projet

```
CEN-Mail/
├── src/                        # Sources décompressées
│   ├── manifest.json           # Métadonnées extension + permissions (MV3)
│   ├── background.js           # Logique principale (53 KB)
│   │                           #   - Config (CFG)
│   │                           #   - Migration batch avec checkpoints
│   │                           #   - Synchro 4 phases
│   │                           #   - Gestion dossiers/tags
│   │                           #   - Microsoft Graph (OAuth2 + catégories)
│   ├── popup/
│   │   ├── popup.html          # Interface 7 onglets (39 KB)
│   │   └── popup.js            # Logique UI (55 KB)
│   ├── token-exchange.html     # Handler OAuth silencieux
│   ├── token-exchange.js       # Échange de token Microsoft
│   └── icons/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-64.png
└── mail-cen-v5.3.xpi          # Extension compilée (prête à installer)
```

## Configuration migration

```javascript
BATCH_SIZE    = 20      // Messages par batch
BATCH_DELAY   = 600     // ms entre chaque batch
POLL_RETRIES  = 30      // Tentatives de retry
POLL_INTERVAL = 500     // ms entre les polls
TEMP_FOLDER   = "Mail-CEN-Temp"  // Dossier de staging
```

## Permissions requises

- `storage` — Sauvegarde état/config locale
- `identity` — OAuth2 (launchWebAuthFlow)
- `messagesRead`, `messagesMove`, `messagesImport`, `messagesDelete` — Lecture, déplacement, import, suppression de messages
- `messagesTags`, `messagesTagsList`, `messagesUpdate` — Gestion des étiquettes/tags
- `accountsRead`, `accountsFolders` — Accès comptes et dossiers
- `addressBooks` — Export des contacts
- `notifications`, `menus`, `tabs` — UI (notifications, menu contextuel, onglets)
- Accès réseau (host_permissions) : `login.microsoftonline.com`, `graph.microsoft.com`, `dns.google`

## Build

Pour recompiler le XPI depuis les sources :

```bash
cd src
zip -r ../mail-cen-v5.3.xpi . -x ".*"
```
