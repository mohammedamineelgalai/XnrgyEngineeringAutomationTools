# Guide de Test Firebase - XNRGY Engineering Automation Tools

## Structure Firebase Complete (avec controle hierarchique)

```
xeat-remote-control-default-rtdb/
├── appConfig/
│   ├── currentVersion: "1.0.0"
│   ├── minVersion: "1.0.0"
│   ├── maintenanceMode: false          ← TESTER: mettre true
│   ├── maintenanceMessage: "..."
│   ├── forceUpdate: false              ← TESTER: mettre true
│   └── updateUrl: "..."
│
├── commands/
│   └── global/
│       ├── killSwitch: false           ← TESTER: mettre true
│       ├── killSwitchMessage: "..."
│       └── forceUpdate: false
│
├── devices/                            ← CONTROLE HIERARCHIQUE COMPLET
│   └── LAPTOP-MOHAMMED_mohammedamine_elgala/
│       ├── machineName: "LAPTOP-MOHAMMED"
│       ├── userName: "mohammedamine.elgala"
│       ├── appVersion: "1.0.0"
│       ├── status: "online"
│       ├── enabled: true               ← false = SUSPENDRE LE POSTE ENTIER
│       ├── disabledMessage: "..."
│       ├── disabledReason: "suspended"
│       ├── heartbeat/...
│       │
│       └── users/                      ← CONTROLE PAR UTILISATEUR SUR CE POSTE
│           ├── jean_dupont/
│           │   ├── enabled: true       ← false = SUSPENDRE CET USER SUR CE POSTE
│           │   ├── disabledMessage: "..."
│           │   └── disabledReason: "suspended"
│           │
│           └── marie_martin/
│               ├── enabled: false      ← Cette utilisatrice est bloquee ici
│               ├── disabledMessage: "Acces revoque suite a changement de departement"
│               └── disabledReason: "revoked"
│
├── users/                              ← (Optionnel - controle GLOBAL par compte)
│   └── user_mohammedamine_elgalai/
│       ├── email: "mohammedamine.elgalai@xnrgy.com"
│       ├── enabled: true
│       └── ...
│
├── versionInfo/
│   └── latest/
│       ├── version: "1.0.0"
│       └── ...
│
└── broadcasts/
    └── welcome_message/
        └── ...
```

---

## Hierarchie de Controle (du plus restrictif au moins restrictif)

```
┌─────────────────────────────────────────────────────────────────────┐
│  NIVEAU 1: Kill Switch (commands/global/killSwitch)                 │
│  → Bloque TOUS les postes, TOUS les utilisateurs                    │
├─────────────────────────────────────────────────────────────────────┤
│  NIVEAU 2a: Device (devices/[ID]/enabled)                           │
│  → Bloque UN POSTE entier, peu importe qui l'utilise                │
├─────────────────────────────────────────────────────────────────────┤
│  NIVEAU 2b: Device/User (devices/[ID]/users/[USER]/enabled)         │
│  → Bloque UN UTILISATEUR sur UN POSTE specifique                    │
├─────────────────────────────────────────────────────────────────────┤
│  NIVEAU 3: User Global (users/[USER]/enabled)                       │
│  → Bloque un utilisateur sur TOUS les postes (optionnel)            │
├─────────────────────────────────────────────────────────────────────┤
│  NIVEAU 4: Maintenance (appConfig/maintenanceMode)                  │
│  → Bloque temporairement tous                                        │
├─────────────────────────────────────────────────────────────────────┤
│  NIVEAU 5: Force Update (appConfig/forceUpdate)                     │
│  → Bloque jusqu'a mise a jour                                        │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Tests a Effectuer

### 1. Kill Switch (Bloquer TOUS les utilisateurs)

Dans Firebase Console, modifier:
```
commands/global/killSwitch: true
commands/global/killSwitchMessage: "Application suspendue pour maintenance critique"
```

**Resultat attendu**: Au prochain demarrage, l'application affiche un message rouge et se ferme.

**Retablir**: Remettre `killSwitch: false`

---

### 2. Suspendre UN POSTE DE TRAVAIL ENTIER

Dans Firebase Console, trouver le device dans `devices/` et modifier:
```
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/enabled: false
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/disabledMessage: "Ce poste est suspendu pour verification."
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/disabledReason: "suspended"
```

**Raisons disponibles pour DEVICE**:
| Raison | Icone | Couleur | Description |
|--------|-------|---------|-------------|
| `"suspended"` | 🖥️ | Orange | Suspension generale |
| `"maintenance"` | 🔧 | Jaune | Poste en maintenance |
| `"unauthorized"` | ⛔ | Rouge | Poste non autorise |

**Resultat attendu**: 
- Ce poste est bloque pour TOUS les utilisateurs Windows
- Les autres postes continuent de fonctionner

**Retablir**: Remettre `enabled: true`

---

### 3. Suspendre UN UTILISATEUR sur UN POSTE (Nouveau!)

Dans Firebase Console, creer/modifier sous `devices/[ID]/users/`:
```
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/users/jean_dupont/enabled: false
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/users/jean_dupont/disabledMessage: "Votre acces a ce poste a ete revoque."
devices/LAPTOP-MOHAMMED_mohammedamine_elgala/users/jean_dupont/disabledReason: "revoked"
```

**Raisons disponibles pour USER sur DEVICE**:
| Raison | Icone | Couleur | Description |
|--------|-------|---------|-------------|
| `"suspended"` | 👤 | Orange | Utilisateur suspendu |
| `"unauthorized"` | 🚷 | Rouge | Non autorise sur ce poste |
| `"revoked"` | 🔐 | Rouge fonce | Acces revoque |

**Resultat attendu**: 
- `jean_dupont` est bloque sur CE poste uniquement
- `jean_dupont` peut encore utiliser d'autres postes
- Les autres utilisateurs sur ce poste fonctionnent normalement

**Retablir**: Remettre `enabled: true` ou supprimer l'entree

---

### 4. Desactiver UN Utilisateur GLOBALEMENT (Optionnel)

Dans Firebase Console, modifier:
```
users/user_mohammedamine_elgalai/enabled: false
users/user_mohammedamine_elgalai/disabledMessage: "Votre acces a ete temporairement suspendu"
```

**Resultat attendu**: Cet utilisateur est bloque sur TOUS les postes.

**Note**: Preferer la suspension par DEVICE/USER (test 3) pour un controle plus precis.

**Retablir**: Remettre `enabled: true`

---

### 5. Mode Maintenance

Dans Firebase Console, modifier:
```
appConfig/maintenanceMode: true
appConfig/maintenanceMessage: "Mise a jour du serveur en cours. Retour prevu a 14h00."
```

**Resultat attendu**: Message jaune de maintenance, application bloquee.

**Retablir**: Remettre `maintenanceMode: false`

---

### 6. Mise a Jour Optionnelle

Dans Firebase Console, modifier:
```
versionInfo/latest/version: "1.1.0"
versionInfo/latest/changelog: "- Nouvelle fonctionnalite X\n- Correction bug Y"
appConfig/forceUpdate: false
```

**Resultat attendu**: Popup "Mise a jour disponible" avec boutons "Telecharger" et "Plus tard".

---

### 7. Mise a Jour FORCEE

Dans Firebase Console, modifier:
```
versionInfo/latest/version: "2.0.0"
appConfig/forceUpdate: true
```
OU
```
commands/global/forceUpdate: true
```

**Resultat attendu**: Popup rouge "Mise a jour requise", seul "Telecharger" disponible, app bloquee.

**Retablir**: Remettre `forceUpdate: false` et version a "1.0.0"

---

### 8. Message Broadcast (Information)

Ajouter dans Firebase Console sous `broadcasts/`:
```json
{
  "test_info": {
    "active": true,
    "title": "Information importante",
    "message": "Une reunion est prevue demain a 10h00 en salle A.",
    "type": "info",
    "createdAt": 1737100000000,
    "expiresAt": 0,
    "targetUser": "",
    "targetDevice": ""
  }
}
```

**Resultat attendu**: Popup bleu d'information, bouton "OK", l'app continue.

---

### 9. Message Broadcast (Avertissement)

```json
{
  "test_warning": {
    "active": true,
    "title": "Attention",
    "message": "Le serveur Vault sera redémarre ce soir a 22h00.",
    "type": "warning",
    "targetUser": "",
    "targetDevice": ""
  }
}
```

**Resultat attendu**: Popup jaune d'avertissement, bouton "Compris", l'app continue.

---

### 10. Message Broadcast BLOQUANT (Erreur)

```json
{
  "test_error": {
    "active": true,
    "title": "Erreur critique",
    "message": "Une erreur critique a ete detectee. Veuillez contacter le support.",
    "type": "error",
    "targetUser": "",
    "targetDevice": ""
  }
}
```

**Resultat attendu**: Popup rouge, bouton "Fermer", l'app est BLOQUEE.

---

### 11. Message Cible (Un seul utilisateur)

```json
{
  "target_user_message": {
    "active": true,
    "title": "Message personnel",
    "message": "Bonjour Mohammed, n'oublie pas de valider le projet 10359.",
    "type": "info",
    "targetUser": "mohammedamine",
    "targetDevice": ""
  }
}
```

**Resultat attendu**: Seul l'utilisateur "mohammedamine" voit ce message.

---

### 12. Message Cible (Un seul poste)

```json
{
  "target_device_message": {
    "active": true,
    "title": "Message pour ce poste",
    "message": "Ce poste sera mis a jour ce week-end.",
    "type": "warning",
    "targetUser": "",
    "targetDevice": "LAPTOP-MOHAMMED"
  }
}
```

**Resultat attendu**: Seul le poste "LAPTOP-MOHAMMED" voit ce message.

---

## Regles Firebase (a copier dans Console > Regles)

```json
{
  "rules": {
    "appConfig": { ".read": true, ".write": false },
    "commands": { ".read": true, ".write": false },
    "versionInfo": { ".read": true, ".write": false },
    "users": { ".read": true, ".write": false },
    "devices": { ".read": true, ".write": true },
    "broadcasts": { ".read": true, ".write": false }
  }
}
```

---

## Ordre de Priorite des Verifications

1. **Kill Switch** → Bloque TOUT (tous les postes, tous les utilisateurs)
2. **Device suspendu** → Bloque CE POSTE entier (peu importe l'utilisateur Windows)
3. **Device/User suspendu** → Bloque UN UTILISATEUR sur CE POSTE (nouveau!)
4. **Utilisateur global desactive** → Bloque cet utilisateur sur tous les postes
5. **Mode Maintenance** → Bloque temporairement tous
6. **Force Update** → Bloque jusqu'a mise a jour
7. **Update Optionnel** → Propose, ne bloque pas
8. **Message Broadcast** → Affiche, bloque seulement si type="error"

---

## Surveillance des Postes Actifs

Dans Firebase Console > Realtime Database > devices:

| Propriete | Description |
|-----------|-------------|
| `status: "online"` | Application en cours d'execution |
| `status: "offline"` | Application fermee proprement |
| `enabled: true` | Poste autorise |
| `enabled: false` | **POSTE SUSPENDU** - App bloquee |
| `disabledReason` | "suspended", "maintenance", "unauthorized" |
| `lastHeartbeat` | Dernier signe de vie (toutes les 60s) |
| `users/[USER]/enabled` | Controle par utilisateur sur ce poste |

### Suspendre un poste rapidement:
```
devices/[DEVICE_ID]/enabled: false
devices/[DEVICE_ID]/disabledMessage: "Votre message ici"
devices/[DEVICE_ID]/disabledReason: "suspended"
```

### Suspendre un utilisateur sur un poste:
```
devices/[DEVICE_ID]/users/[USERNAME]/enabled: false
devices/[DEVICE_ID]/users/[USERNAME]/disabledMessage: "Votre message ici"
devices/[DEVICE_ID]/users/[USERNAME]/disabledReason: "revoked"
```

---

## Cas d'Usage Typiques

### Scenario 1: Employe qui quitte l'entreprise
```
# Bloquer sur tous les postes où il etait connecte
devices/POSTE-RECEPTION/users/ancien_employe/enabled: false
devices/POSTE-RECEPTION/users/ancien_employe/disabledReason: "revoked"
```

### Scenario 2: Poste compromis ou vol
```
# Bloquer le poste entier immediatement
devices/LAPTOP-PERDU_user/enabled: false
devices/LAPTOP-PERDU_user/disabledReason: "unauthorized"
devices/LAPTOP-PERDU_user/disabledMessage: "Poste signale comme perdu. Contactez IT."
```

### Scenario 3: Maintenance planifiee
```
# Mettre le poste en maintenance
devices/PC-ATELIER-01_technicien/enabled: false
devices/PC-ATELIER-01_technicien/disabledReason: "maintenance"
devices/PC-ATELIER-01_technicien/disabledMessage: "Mise a jour materiel en cours."
```

Si un poste a un `lastHeartbeat` vieux de plus de 2 minutes et `status: "online"`, 
l'application a probablement crashe ou ete fermee de force.
