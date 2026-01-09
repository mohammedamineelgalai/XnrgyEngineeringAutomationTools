# Synchronisation Checklist HVAC avec Vault

## 📋 Vue d'ensemble

La Checklist HVAC synchronise maintenant automatiquement avec Vault toutes les **4-5 minutes** de façon bidirectionnelle :
- **Upload/Écraser** : Les modifications locales sont envoyées vers Vault
- **Téléchargement** : Les changements des autres utilisateurs sont récupérés
- **Résolution de conflits** : Dernier modifié gagne (basé sur `LastModifiedDate`)

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    ChecklistHVACWindow                       │
│  ┌───────────────────────────────────────────────────────┐  │
│  │  WebView2 (HTML React Checklist)                      │  │
│  │  ┌─────────────────────────────────────────────────┐  │  │
│  │  │ JavaScript Bridge (window.checklistSync)        │  │  │
│  │  │  - saveData()                                   │  │  │
│  │  │  - loadData()                                   │  │  │
│  │  │  - syncNow()                                    │  │  │
│  │  └─────────────────────────────────────────────────┘  │  │
│  └───────────────────────────────────────────────────────┘  │
│                           ↕                                  │
│  ┌───────────────────────────────────────────────────────┐  │
│  │  ChecklistSyncService                                 │  │
│  │  - Sync automatique (timer 4 min)                    │  │
│  │  - Upload vers Vault                                  │  │
│  │  - Download depuis Vault                              │  │
│  │  - Merge données (conflits)                           │  │
│  └───────────────────────────────────────────────────────┘  │
│                           ↕                                  │
│  ┌───────────────────────────────────────────────────────┐  │
│  │  VaultSDKService                                      │  │
│  │  - Connexion Vault                                    │  │
│  │  - UploadFile / GetFolder                             │  │
│  └───────────────────────────────────────────────────────┘  │
│                           ↕                                  │
│                    Vault Professional 2026                   │
│  $/Engineering/Inventor_Standards/.../Checklist_HVAC_Data/  │
└─────────────────────────────────────────────────────────────┘
```

## 📁 Structure des fichiers

```
Modules/ChecklistHVAC/
├── Models/
│   └── ChecklistDataModel.cs              # Modèle JSON des données
├── Services/
│   └── ChecklistSyncService.cs            # Service de synchronisation
├── Resources/
│   ├── Checklist HVACAHU.html             # Fichier HTML principal (À MIGRER)
│   └── ChecklistSyncBridge.js             # Script JavaScript bridge
└── Views/
    └── ChecklistHVACWindow.xaml(.cs)      # Fenêtre principale
```

## 🔄 Workflow de synchronisation

### 1. Sauvegarde depuis le HTML (instantanée)
```javascript
// Depuis votre code React
window.saveChecklistResponse(
    "25001-01-01",  // moduleId
    "25001",        // projectNumber
    "01",           // reference
    "01",           // module
    checkpointId,   // ID du checkpoint
    "fait",         // status: "fait", "non_applicable", "pas_fait"
    "Commentaire",  // commentaire
    "MAE"           // initiales utilisateur
);
```

### 2. Synchronisation automatique (toutes les 4 minutes)
- `ChecklistSyncService` scanne tous les fichiers JSON locaux
- Télécharge les versions Vault (si plus récentes)
- Fusionne avec les données locales
- Upload vers Vault (écrase si nécessaire)

### 3. Synchronisation manuelle
- Bouton "🔄 Sync Vault" dans l'interface
- Force une synchronisation immédiate du module actuel

## 📦 Format des données JSON

Stocké dans Vault : `$/Engineering/Inventor_Standards/Automation_Standard/Checklist_HVAC_Data/Checklist_[PROJECT]-[REF]-[MODULE].json`

```json
{
  "moduleId": "25001-01-01",
  "projectNumber": "25001",
  "reference": "01",
  "module": "01",
  "lastModifiedBy": "MAE",
  "lastModifiedDate": "2026-01-15T10:30:00Z",
  "version": 5,
  "responses": {
    "1": {
      "checkpointId": 1,
      "status": "fait",
      "comment": "Vérifié avec succès",
      "userInitials": "MAE",
      "modifiedDate": "2026-01-15T10:25:00Z"
    },
    "2": {
      "checkpointId": 2,
      "status": "non_applicable",
      "comment": "",
      "userInitials": "AC",
      "modifiedDate": "2026-01-15T09:15:00Z"
    }
  }
}
```

## 🚀 Migration du HTML

### Étape 1: Copier le fichier HTML
```powershell
# Copier depuis le projet de démonstration vers le projet principal
Copy-Item "C:\Users\mohammedamine.elgala\source\repos\ChecklistHVAC\Checklist HVACAHU - By Mohammed Amine Elgalai.html" `
    "C:\Users\mohammedamine.elgala\source\repos\XnrgyEngineeringAutomationTools\Modules\ChecklistHVAC\Resources\ChecklistHVAC.html"
```

### Étape 2: Modifier MainWindow.xaml.cs
```csharp
// Remplacer le chemin hardcodé par le chemin dans le projet
private string ChecklistHVACPath => Path.Combine(
    AppDomain.CurrentDomain.BaseDirectory, 
    "Modules", "ChecklistHVAC", "Resources", "ChecklistHVAC.html"
);
```

### Étape 3: Ajouter le script bridge dans le HTML
Ajouter avant `</body>` dans le HTML :
```html
<script type="text/javascript">
    // Le script bridge sera injecté automatiquement par ChecklistHVACWindow
    // Mais vous pouvez aussi l'inclure manuellement :
    // <script src="ChecklistSyncBridge.js"></script>
</script>
```

### Étape 4: Modifier le code React pour utiliser le bridge
Dans votre code React, remplacer `localStorage` par `window.checklistSync` :

```javascript
// AVANT (localStorage uniquement)
const saveResponse = (checkpointId, status, comment) => {
    const key = `checklist_${moduleId}_${checkpointId}`;
    localStorage.setItem(key, JSON.stringify({ status, comment }));
};

// APRÈS (avec synchronisation Vault)
const saveResponse = (checkpointId, status, comment) => {
    const success = window.saveChecklistResponse(
        moduleId,           // "25001-01-01"
        projectNumber,      // "25001"
        reference,          // "01"
        module,             // "01"
        checkpointId,       // 1, 2, 3, ...
        status,             // "fait", "non_applicable", "pas_fait"
        comment,            // "Commentaire..."
        currentUser.initials // "MAE"
    );
    
    if (success) {
        setNotification('Réponse sauvegardée et synchronisée avec Vault');
    }
};

// Charger les réponses au démarrage
useEffect(() => {
    const savedResponses = window.loadChecklistData(moduleId);
    setResponses(savedResponses);
}, [moduleId]);
```

## 🔧 Configuration

### Chemin Vault (dans ChecklistSyncService.cs)
```csharp
private const string VAULT_CHECKLIST_FOLDER = 
    "$/Engineering/Inventor_Standards/Automation_Standard/Checklist_HVAC_Data";
```

### Intervalle de synchronisation (modifiable)
```csharp
private readonly int _syncIntervalMinutes = 4;  // Modifier ici (4-5 minutes recommandé)
```

### Cache local
```
C:\Users\[USER]\AppData\Local\XnrgyEngineeringAutomationTools\ChecklistHVAC\
└── Checklist_[MODULE_ID].json
```

## 📊 Résolution de conflits

**Stratégie : Dernier modifié gagne**

1. Comparer `LastModifiedDate` entre version locale et Vault
2. Si Vault plus récent : utiliser Vault comme base, fusionner nouvelles réponses locales
3. Si local plus récent : utiliser local comme base
4. Incrémenter `version` à chaque modification

**Important** : Les réponses individuelles sont fusionnées intelligemment :
- Si un checkpoint a été modifié après la dernière sync, il est conservé
- Les réponses de différents utilisateurs peuvent coexister

## 🌐 Export Word (Optionnel - À implémenter)

Pour l'instant, `ExportToWordAsync` exporte en JSON. Pour exporter en Word réel :

1. Installer NuGet package : `DocX` ou utiliser `iTextSharp` (déjà dans le projet)
2. Implémenter génération de document Word avec formatage
3. Exemple de structure Word :
   - En-tête avec Module ID, date, utilisateur
   - Tableau des checkpoints avec statuts
   - Commentaires formatés

## ✅ Checklist de déploiement

- [x] ChecklistSyncService créé
- [x] ChecklistDataModel créé
- [x] ChecklistHVACWindow modifié avec bridge JavaScript
- [x] Script JavaScript bridge créé
- [x] Bouton synchronisation manuelle ajouté
- [ ] HTML migré dans `Modules/ChecklistHVAC/Resources/`
- [ ] MainWindow.xaml.cs mis à jour avec nouveau chemin
- [ ] Code React modifié pour utiliser `window.checklistSync`
- [ ] Test synchronisation avec plusieurs utilisateurs
- [ ] Documentation utilisateur créée

## 🐛 Dépannage

### Synchronisation ne fonctionne pas
- Vérifier que Vault est connecté : `_vaultService.IsConnected == true`
- Vérifier les logs : `bin\Release\Logs\VaultSDK_*.log`
- Vérifier les permissions Vault sur le dossier `Checklist_HVAC_Data`

### Données non synchronisées
- Vérifier que le fichier JSON local existe dans `AppData\Local\...`
- Vérifier que le module ID est correct (format: "PROJECT-REF-MODULE")
- Vérifier les erreurs dans la console JavaScript (F12 dans WebView2)

### Pont JavaScript non disponible
- Vérifier que WebView2 est initialisé : `WebViewControl.CoreWebView2 != null`
- Vérifier que `SetupJavaScriptBridge()` est appelé après `NavigationCompleted`
- Consulter les logs C# pour erreurs `AddHostObjectToScript`

## 📝 Notes importantes

1. **Performance** : La synchronisation se fait en arrière-plan (non-bloquante)
2. **Réseau** : Nécessite connexion Vault active pour fonctionner
3. **Offline** : Mode fallback avec localStorage si Vault non connecté
4. **Multi-utilisateurs** : Les modifications sont visibles après la prochaine sync (4-5 min)
5. **Version** : Le numéro de version est incrémenté automatiquement à chaque modification

## 🔐 Sécurité

- Les fichiers JSON sont stockés en clair dans Vault (pas de données sensibles)
- Les permissions Vault s'appliquent normalement (droit d'écriture requis)
- Le cache local est dans `AppData\Local` (protégé par Windows)

---

**Dernière mise à jour** : 2026-01-15  
**Auteur** : Mohammed Amine Elgalai - XNRGY Climate Systems ULC

