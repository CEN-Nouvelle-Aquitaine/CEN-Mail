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
4. Sélectionner `mail-cen-v6.0.xpi`

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
└── mail-cen-v6.0.xpi          # Extension compilée (prête à installer)
```

## Configuration migration (v6.0)

```javascript
BATCH_SIZE       = 5       // Petits batchs pour éviter le throttle Outlook IMAP
BATCH_DELAY      = 1500    // ms entre chaque batch
MSG_DELAY        = 200     // ms entre messages individuels
RETRY_MAX        = 4       // Tentatives par opération
RETRY_BACKOFF    = 2000    // Délai initial du retry (×attempt)
HEALTH_THRESHOLD = 5       // Erreurs consécutives avant mode dégradé
HEALTH_COOLDOWN  = 30000   // Pause de récupération en mode dégradé (ms)
HEALTH_DELAY_MULT= 3       // Multiplicateur des délais en mode dégradé
TEMP_FOLDER      = "Mail-CEN-Temp"
```

### Robustesse v6.0

- **3 stratégies en cascade** par message : `move/copy direct` → `copy+delete` → `import via raw eml`
- **Classification d'erreurs** : transitoires (Aborted, timeout…) retentées, permanentes (doublon, quota…) skippées
- **Health monitor** : détecte les cascades d'erreurs et passe en mode dégradé (×3 délais + cooldown 30s)
- **Option "Forcer (ignorer doublons)"** : désactive la détection pré-import des doublons
- **Conformité TB MV3** : utilise `{ deletePermanently: true }` (et non `skipTrash`)

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
zip -r ../mail-cen-v6.0.xpi . -x ".*"
```

## Changelog

### v6.0.0 — Robustesse migration IMAP + conformité MV3 stricte

- **Migration cascade** : 3 stratégies de fallback (direct → copy+delete → import raw)
- **Health monitor** : détection des cascades d'erreurs + mode dégradé automatique
- **Classification d'erreurs** : transient/permanent/unknown avec retry adaptatif
- **Fix MV3** : `deletePermanently` au lieu de `skipTrash` (conforme doc officielle TB)
- **Outlook IMAP** : batchs réduits (20→5), délais augmentés, retry exponentiel ×4
- **UI mode dégradé** : avertissement utilisateur quand la connexion sature

### v5.3 — Compat Manifest v3 + APIs TB 128+

- Migration Manifest v2 → v3
- Async iterators / paginated lists
- `getRaw()` retourne File/Blob (TB 117+)
- `MailFolder.type` → `specialUse`, `MessageHeader.folder` → `folderId`
