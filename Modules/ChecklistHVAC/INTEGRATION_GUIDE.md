# Guide d'Intégration - Checklist HVAC Synchronisation Vault

## ✅ Ce qui a été fait

### 1. Services créés
- ✅ `ChecklistSyncService.cs` : Service de synchronisation bidirectionnelle avec Vault
- ✅ `ChecklistDataModel.cs` : Modèle de données JSON pour les réponses

### 2. Interface modifiée
- ✅ `ChecklistHVACWindow.xaml.cs` : Intégration synchronisation automatique + bridge JavaScript
- ✅ `ChecklistHVACWindow.xaml` : Bouton "Sync Vault" ajouté
- ✅ `MainWindow.xaml.cs` : Chemin HTML mis à jour avec fallback

### 3. Documentation
- ✅ `SYNCHRONISATION_VAULT.md` : Documentation complète
- ✅ `ChecklistSyncBridge.js` : Script JavaScript helper (optionnel)
- ✅ `INTEGRATION_GUIDE.md` : Ce guide

## 🔧 Ce qu'il reste à faire

### Étape 1: Migrer le fichier HTML (REQUIS)

Copier le fichier HTML depuis le projet de démonstration vers le projet principal :

```powershell
# Créer le dossier Resources s'il n'existe pas
New-Item -ItemType Directory -Force -Path "Modules\ChecklistHVAC\Resources"

# Copier le fichier HTML
Copy-Item `
    "C:\Users\mohammedamine.elgala\source\repos\ChecklistHVAC\Checklist HVACAHU - By Mohammed Amine Elgalai.html" `
    "C:\Users\mohammedamine.elgala\source\repos\XnrgyEngineeringAutomationTools\Modules\ChecklistHVAC\Resources\ChecklistHVAC.html"
```

**Important** : Renommer le fichier en `ChecklistHVAC.html` (sans espaces, plus simple).

### Étape 2: Ajouter le HTML au projet .csproj (REQUIS)

Ouvrir `XnrgyEngineeringAutomationTools.csproj` et ajouter :

```xml
<!-- Dans la section ItemGroup des Resources -->
<Content Include="Modules\ChecklistHVAC\Resources\ChecklistHVAC.html">
  <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
</Content>
```

### Étape 3: Modifier le code React dans le HTML (REQUIS)

Dans votre fichier HTML `ChecklistHVAC.html`, modifier le code React pour utiliser le bridge :

#### A. Remplacer localStorage par window.checklistSync

**AVANT** :
```javascript
// Sauvegarder dans localStorage
localStorage.setItem(`checklist_${moduleId}_${checkpointId}`, JSON.stringify({
    status: status,
    comment: comment
}));

// Charger depuis localStorage
const saved = localStorage.getItem(`checklist_${moduleId}_${checkpointId}`);
```

**APRÈS** :
```javascript
// Sauvegarder via le bridge (sync avec Vault automatique)
window.saveChecklistResponse(
    moduleId,        // "25001-01-01"
    projectNumber,   // "25001"
    reference,       // "01"
    module,          // "01"
    checkpointId,    // 1, 2, 3, ...
    status,          // "fait", "non_applicable", "pas_fait"
    comment,         // "Commentaire..."
    currentUser.initials // "MAE"
);

// Charger depuis le bridge (depuis Vault si disponible)
const savedResponses = window.loadChecklistData(moduleId);
const saved = savedResponses[checkpointId];
```

#### B. Ajouter un bouton "Synchroniser maintenant"

Dans votre interface React, ajouter un bouton :

```javascript
<button onClick={() => window.syncChecklistWithVault(moduleId, projectNumber, reference, module)}>
    🔄 Synchroniser avec Vault
</button>
```

#### C. Charger les données au démarrage

Dans votre `useEffect` d'initialisation :

```javascript
useEffect(() => {
    // Charger les réponses sauvegardées depuis Vault/localStorage
    const savedResponses = window.loadChecklistData(moduleId);
    
    if (savedResponses && Object.keys(savedResponses).length > 0) {
        // Restaurer les réponses dans l'état React
        setResponses(savedResponses);
        console.log(`[Checklist] ${Object.keys(savedResponses).length} réponses chargées`);
    }
}, [moduleId]);
```

### Étape 4: Tester la synchronisation

1. **Test local** :
   - Ouvrir Checklist HVAC
   - Remplir quelques checkpoints
   - Vérifier que les fichiers JSON sont créés dans `AppData\Local\...\ChecklistHVAC\`

2. **Test avec Vault** :
   - Connecter à Vault
   - Remplir des checkpoints
   - Attendre 4-5 minutes (ou cliquer "Sync Vault")
   - Vérifier dans Vault : `$/Engineering/Inventor_Standards/.../Checklist_HVAC_Data/`

3. **Test multi-utilisateurs** :
   - Utilisateur 1 : Remplir des checkpoints → Attendre sync
   - Utilisateur 2 : Ouvrir le même module → Vérifier que les réponses apparaissent après sync

## 📝 Exemple d'intégration React complète

```javascript
// Dans votre composant React principal
const ChecklistApp = () => {
    const [moduleId] = useState("25001-01-01");
    const [projectNumber] = useState("25001");
    const [reference] = useState("01");
    const [module] = useState("01");
    const [responses, setResponses] = useState({});
    const [currentUser] = useState({ initials: "MAE" });

    // Charger les données au démarrage
    useEffect(() => {
        const saved = window.loadChecklistData(moduleId);
        if (saved) {
            setResponses(saved);
        }
    }, [moduleId]);

    // Sauvegarder une réponse
    const handleSaveResponse = (checkpointId, status, comment) => {
        // Mettre à jour l'état React
        setResponses(prev => ({
            ...prev,
            [checkpointId]: {
                checkpointId,
                status,
                comment,
                userInitials: currentUser.initials,
                modifiedDate: new Date().toISOString()
            }
        }));

        // Sauvegarder via le bridge (sync avec Vault)
        const success = window.saveChecklistResponse(
            moduleId, projectNumber, reference, module,
            checkpointId, status, comment, currentUser.initials
        );

        if (success) {
            showNotification('Réponse sauvegardée et synchronisée avec Vault', 'success');
        }
    };

    // Synchroniser manuellement
    const handleSyncNow = () => {
        window.syncChecklistWithVault(moduleId, projectNumber, reference, module);
        showNotification('Synchronisation avec Vault en cours...', 'info');
    };

    return (
        <div>
            {/* Votre interface de checklist ici */}
            <button onClick={handleSyncNow}>
                🔄 Synchroniser maintenant
            </button>
        </div>
    );
};
```

## 🎯 Avantages de cette solution

1. ✅ **Pas de serveur requis** : Utilise Vault comme backend (déjà disponible)
2. ✅ **Synchronisation automatique** : Toutes les 4-5 minutes en arrière-plan
3. ✅ **Mode offline** : Fallback localStorage si Vault non connecté
4. ✅ **Multi-utilisateurs** : Les modifications sont partagées via Vault
5. ✅ **Résolution de conflits** : Dernier modifié gagne (simple et efficace)
6. ✅ **Intégration native** : Pas besoin de modifier l'infrastructure réseau

## 🔄 Alternative : Serveur web (Optionnel - Plus tard)

Si vous voulez vraiment un serveur web plus tard :

**Avantages** :
- Synchronisation en temps réel (WebSocket)
- API REST plus flexible
- Base de données SQL pour requêtes complexes

**Inconvénients** :
- Infrastructure supplémentaire à maintenir
- Coûts serveur
- Complexité accrue

**Recommandation** : Utiliser Vault d'abord, migrer vers serveur seulement si besoin de fonctionnalités avancées (temps réel, analytics, etc.)

## 📞 Support

Pour toute question ou problème :
- Consulter les logs : `bin\Release\Logs\VaultSDK_*.log`
- Vérifier `SYNCHRONISATION_VAULT.md` pour le dépannage
- Contact : mohammedamine.elgalai@xnrgy.com

---

**Version** : 1.0.0  
**Date** : 2026-01-15  
**Auteur** : Mohammed Amine Elgalai - XNRGY Climate Systems ULC


