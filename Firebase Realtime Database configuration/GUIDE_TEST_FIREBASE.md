# Guide de Test Firebase - XNRGY Engineering Automation Tools

## Structure Firebase Complete

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
├── users/
│   └── user_mohammedamine_elgalai/
│       ├── email: "mohammedamine.elgalai@xnrgy.com"
│       ├── displayName: "Mohammed Amine Elgalai"
│       ├── enabled: true               ← TESTER: mettre false
│       ├── disabledMessage: "..."
│       ├── role: "admin"
│       └── site: "Laval"
│
├── devices/                            ← Rempli automatiquement
│   └── LAPTOP-MOHAMMED_mohammedamine_elgala/
│       ├── machineName: "LAPTOP-MOHAMMED"
│       ├── userName: "mohammedamine.elgala"
│       ├── appVersion: "1.0.0"
│       ├── status: "online"
│       └── heartbeat/...
│
├── versionInfo/
│   └── latest/
│       ├── version: "1.0.0"            ← TESTER: mettre "1.1.0"
│       ├── downloadUrl: "..."
│       ├── releaseDate: "2026-01-16"
│       └── changelog: "..."
│
└── broadcasts/
    └── welcome_message/
        ├── active: false               ← TESTER: mettre true
        ├── title: "Bienvenue"
        ├── message: "..."
        ├── type: "info"                ← Valeurs: "info", "warning", "error"
        ├── targetUser: ""              ← Vide = tous, sinon username
        └── targetDevice: ""            ← Vide = tous, sinon deviceId
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

### 2. Desactiver UN Utilisateur Specifique

Dans Firebase Console, modifier:
```
users/user_mohammedamine_elgalai/enabled: false
users/user_mohammedamine_elgalai/disabledMessage: "Votre acces a ete temporairement suspendu"
```

**Resultat attendu**: Seul cet utilisateur est bloque, les autres continuent.

**Retablir**: Remettre `enabled: true`

---

### 3. Mode Maintenance

Dans Firebase Console, modifier:
```
appConfig/maintenanceMode: true
appConfig/maintenanceMessage: "Mise a jour du serveur en cours. Retour prevu a 14h00."
```

**Resultat attendu**: Message jaune de maintenance, application bloquee.

**Retablir**: Remettre `maintenanceMode: false`

---

### 4. Mise a Jour Optionnelle

Dans Firebase Console, modifier:
```
versionInfo/latest/version: "1.1.0"
versionInfo/latest/changelog: "- Nouvelle fonctionnalite X\n- Correction bug Y"
appConfig/forceUpdate: false
```

**Resultat attendu**: Popup "Mise a jour disponible" avec boutons "Telecharger" et "Plus tard".

---

### 5. Mise a Jour FORCEE

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

### 6. Message Broadcast (Information)

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

### 7. Message Broadcast (Avertissement)

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

### 8. Message Broadcast BLOQUANT (Erreur)

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

### 9. Message Cible (Un seul utilisateur)

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

### 10. Message Cible (Un seul poste)

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

1. **Kill Switch** → Bloque TOUT
2. **Utilisateur desactive** → Bloque cet utilisateur
3. **Mode Maintenance** → Bloque temporairement
4. **Force Update** → Bloque jusqu'a mise a jour
5. **Update Optionnel** → Propose, ne bloque pas
6. **Message Broadcast** → Affiche, bloque seulement si type="error"

---

## Surveillance des Postes Actifs

Dans Firebase Console > Realtime Database > devices:

- **status: "online"** = Application en cours
- **status: "offline"** = Application fermee
- **lastHeartbeat** = Dernier signe de vie (mis a jour toutes les 60s)

Si un poste a un `lastHeartbeat` vieux de plus de 2 minutes et `status: "online"`, 
l'application a probablement crashe ou ete fermee de force.
