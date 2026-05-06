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
4. Sélectionner `mail-cen-v7.1.xpi`

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
│   ├── background.js           # Logique principale
│   │                           #   - Config + classification d'erreurs
│   │                           #   - Migration cascade 3 stratégies + retry
│   │                           #   - Health monitor (mode dégradé adaptatif)
│   │                           #   - Synchro étiquettes (analyse + apply)
│   │                           #   - Microsoft Graph (OAuth2 + catégories)
│   │                           #   - Export contacts (API contacts MV3)
│   ├── popup/
│   │   ├── popup.html          # Interface 7 onglets
│   │   └── popup.js            # Logique UI + reprise d'état
│   └── icons/
│       ├── icon-16.png
│       ├── icon-32.png
│       └── icon-64.png
└── mail-cen-v7.1.xpi          # Extension compilée (prête à installer)
```

## Configuration migration (v7.0)

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
zip -r ../mail-cen-v7.1.xpi . -x ".*"
```

## Changelog

### v7.1.0 — Fix encodage UTF-8 sur la migration

- **Caractères français corrompus à la copie** (é, à, ç, etc. → `?`) : la stratégie 3 décodait le raw email en UTF-8 puis utilisait `charCodeAt(i) & 0xff`, ce qui tronquait les caractères multi-octets. Désormais le `File` retourné par `messages.getRaw()` est passé **tel quel** à `messages.import()`, sans aucune conversion intermédiaire.
- **`applyTagsToSubject`** : même bug, corrigé via `TextEncoder` qui produit des octets UTF-8 propres.
- Nouveau helper `getRawFile()` (octets bruts) distinct de `getRawString()` (texte UTF-8 décodé).

### v7.0.0 — Audit complet + corrections critiques

**Bugs corrigés :**
- **Export contacts** : `addressBooks.list()` en MV3 ne renvoie plus `contacts[]`. Migré vers `messenger.contacts.list(bookId)` qui retourne effectivement les contacts. `count` et `.vcf` fonctionnent enfin.
- **Doublons silencieux mode "Déplacer"** : si `delete` source échoue après `copy`, un `warning` explicite est désormais propagé à la progression au lieu d'une fausse confirmation OK.
- **Reprise sync après fermeture popup** : `restoreState()` reconnaît tous les types `SYNC_APPLY_DONE`, `SYNC_ANALYSE_DONE`, `GRAPH_APPLY_DONE` et restaure correctement l'état.
- **Bouton Annuler interruptible** : `cancellableSleep()` remplace les `setTimeout` bloquants. Le retry, le cooldown du mode dégradé et les batchs sync respectent désormais `mig.cancel`.
- **Récursion sous-dossiers complète** : `analyseBoxes` (scanSrc/scanDst) et `runSubjectTagAll` re-fetchent via `getSubFolders` quand le cache est vide (cohérent avec `migrateFolderRecursive`).
- **Crash sur message corrompu** : `full.headers?.subject` au lieu de `full.headers.subject`.
- **Double listener tab Migration** : suppression du `addEventListener` redondant.

**Nettoyage code mort :**
- Suppression de `src/token-exchange.html` + `.js` (orphelins, le flux PKCE a été remplacé par implicit).
- Suppression de `OL_CATEGORIES`, `CHECKPOINT_KEY`, `_deviceCodeCancel`, `waitForNewAccount`, action `graphCancelAuth`, broadcast `M365_ACCOUNT_DETECTED`.
- `default` du switch de messages renvoie une erreur claire au lieu de `false`.

**Conformité TB MV3** : audit final confirmé sur webextension-api.thunderbird.net.

### v6.2.0 — Création des sous-dossiers robuste

- **Récursion sous-dossiers** : re-fetch via `getSubFolders` si le cache est vide (certaines configurations IMAP ne peuplent pas `subFolders` à la 1re passe)
- **`ensureFolder`** : retry exponentiel sur `folders.create()` + revérification post-erreur (Outlook peut throw alors que la création a réussi)
- **Logs explicites** : nombre de sous-dossiers détectés + statut de création de chaque dossier dans la console

### v6.1.0 — Import IMAP direct (suppression du détour temp local)

- **Stratégie 3 réécrite** : `messages.import()` direct vers le dossier IMAP destination (TB 128+ utilise APPEND, INTERNALDATE préservée)
- Le dossier `Mail-CEN-Temp` n'est créé **que si l'import direct échoue** (rare)
- Migration cross-account 2× plus rapide en cas nominal (un seul transfert IMAP au lieu de import-local + move-to-IMAP)

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
