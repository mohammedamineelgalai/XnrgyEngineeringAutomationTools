# XNRGY Engineering Automation Tools

> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2
>
> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC

---

## Description

**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

### Objectif Principal

Remplacer les multiples applications standalone par une **plateforme unique** avec :
- Connexion centralisée à Vault & Inventor
- Interface utilisateur moderne et cohérente (thèmes sombre/clair)
- Partage de services communs (logging, configuration chiffrée AES-256)
- Déploiement multi-sites et maintenance simplifiée
- Paramètres centralisés via Vault (50+ utilisateurs, 3 sites)

---

## Modules Intégrés

| Module | Description | Statut |
|--------|-------------|--------|
| **Upload Module** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | [+] 100% |
| **Créer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |
| **Réglages Admin** | Configuration centralisée et synchronisée via Vault (AES-256) | [+] 100% |
| **Upload Template** | Upload templates vers Vault (réservé Admin) | [+] 100% |
| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |
| **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | [~] Planifié |
| **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |
| **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |
| **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifié |

---

## Architecture du Projet

```
XnrgyEngineeringAutomationTools/
├── App.xaml(.cs)                    # Point d'entrée application
├── MainWindow.xaml(.cs)             # Dashboard principal (hub)
├── appsettings.json                 # Configuration utilisateur
│
├── Modules/                         # Modules fonctionnels isolés
│   ├── UploadModule/                # Module Upload Vault
│   │   ├── Models/
│   │   │   ├── VaultUploadFileItem.cs
│   │   │   └── VaultUploadModels.cs
│   │   └── Views/
│   │       └── UploadModuleWindow.xaml(.cs)
│   │
│   ├── CreateModule/                # Module Pack & Go / Copy Design
│   │   ├── Models/
│   │   │   ├── CreateModuleRequest.cs
│   │   │   └── CreateModuleSettings.cs
│   │   ├── Services/
│   │   │   ├── InventorCopyDesignService.cs
│   │   │   └── ModuleCopyService.cs
│   │   └── Views/
│   │       ├── CreateModuleWindow.xaml(.cs)
│   │       └── CreateModuleSettingsWindow.xaml(.cs)
│   │
│   ├── UploadTemplate/              # Module Upload Template
│   │   └── Views/
│   │       └── UploadTemplateWindow.xaml(.cs)
│   │
│   └── ChecklistHVAC/              # Module Checklist HVAC
│       └── Views/
│           └── ChecklistHVACWindow.xaml(.cs)
│
├── Shared/                          # Composants partagés
│   ├── Views/                       # Fenêtres partagées
│   │   ├── LoginWindow.xaml(.cs)    # Connexion Vault
│   │   ├── ModuleSelectionWindow.xaml(.cs)  # Sélection module
│   │   ├── PreviewWindow.xaml(.cs)  # Prévisualisation
│   │   └── XnrgyMessageBox.xaml(.cs)  # MessageBox moderne
│   ├── Models/                      # Modèles partagés (vide pour l'instant)
│   └── Services/                    # Services partagés (vide pour l'instant)
│
├── Services/                        # Services métier partagés
│   ├── VaultSDKService.cs           # SDK Vault v31.0.84 (~1200 lignes)
│   ├── VaultSettingsService.cs      # Config chiffrée + sync Vault
│   ├── InventorService.cs           # Connexion Inventor COM
│   ├── InventorPropertyService.cs   # iProperties via Inventor COM
│   ├── ApprenticePropertyService.cs # Lecture iProperties via Apprentice
│   ├── OlePropertyService.cs         # Lecture OLE via OpenMCDF
│   ├── NativeOlePropertyService.cs  # Lecture OLE native
│   ├── WindowsPropertyService.cs    # Lecture propriétés Windows Shell
│   ├── Logger.cs                    # Logging NLog UTF-8
│   ├── SimpleLogger.cs              # Logger simple
│   ├── JournalColorService.cs       # Couleurs journal (succès/erreur)
│   ├── ThemeHelper.cs               # Gestion thèmes sombre/clair
│   ├── UserPreferencesManager.cs   # Préférences utilisateur
│   ├── SettingsService.cs           # Gestion settings app
│   └── CredentialsManager.cs        # Gestion credentials chiffrées
│
├── Models/                          # Modèles de données partagés
│   ├── ApplicationConfiguration.cs  # Configuration application
│   ├── CategoryItem.cs              # Item catégorie pour ComboBox
│   ├── FileItem.cs                  # Item fichier pour DataGrid
│   ├── FileToUpload.cs              # Fichier à uploader
│   ├── LifecycleDefinitionItem.cs  # Lifecycle Definition
│   ├── LifecycleStateItem.cs        # Lifecycle State
│   ├── ModuleInfo.cs                # Informations module
│   ├── ProjectInfo.cs               # Informations projet
│   ├── ProjectProperties.cs         # Propriétés Project/Ref/Module
│   └── VaultConfiguration.cs        # Configuration Vault
│
├── ViewModels/                      # MVVM ViewModels
│   └── AppMainViewModel.cs          # ViewModel principal
│
├── Converters/                      # Convertisseurs WPF
│   ├── AnyOperationActiveConverter.cs
│   ├── BooleanToColorConverter.cs
│   ├── BooleanToTextConverter.cs
│   ├── InverseBooleanConverter.cs
│   ├── InverseBooleanToVisibilityConverter.cs
│   ├── NullToVisibilityConverter.cs
│   └── ProgressToWidthConverter.cs
│
├── Styles/                          # Styles centralisés
│   └── XnrgyStyles.xaml             # 50+ styles partagés (couleurs, boutons, inputs)
│
├── Resources/                       # Images et icônes
│   ├── VaultAutomationTool.ico
│   └── VaultAutomationTool.png
│
├── Assets/                          # Assets visuels
│   └── Icons/
│       ├── ChecklistHVAC.png
│       ├── DXFVerifier.ico
│       ├── DXFVerifier.png
│       ├── VaultUpload.ico
│       ├── VaultUpload.png
│       └── XnrgyTools.ico
│
├── Tools/                           # Outils utilitaires
│   └── VaultBulkUploader/           # Outil console upload massif
│
├── Scripts/                          # Scripts PowerShell
│   ├── CleanInventor2023Registry.ps1
│   ├── Prepare-TemplateFiles.ps1
│   └── Upload-ToVaultProd.ps1
│
├── Temp&Test/                       # Fichiers temporaires/test (exclus du build)
│   ├── DiagnoseOleProperties.cs
│   └── TestWindowsPropertyService.cs
│
├── Backups/                         # Sauvegardes locales
│   └── BACKUP_YYYYMMDD_HHMMSS/
│
└── build-and-run.ps1                # Script compilation MSBuild automatique
```

---

## Fonctionnalités Implémentées

### 1. Upload Module (100%)

Module intégré pour l'upload de fichiers vers Vault avec gestion complète des propriétés :

- **Connexion centralisée** - Utilise la connexion Vault de l'app principale
- **Scan automatique** des modules engineering avec extraction propriétés
- **Séparation Inventor/Non-Inventor** dans deux DataGrids avec headers visibles
- **Application automatique** des propriétés métier:
  - Project (ID=112)
  - Reference (ID=121)
  - Module (ID=122)
- **Assignation complète**:
  - Catégories Vault
  - Lifecycle Definitions et States
  - Révisions
- **Synchronisation Vault vers iProperties** via `IExplorerUtil`
- **Journal des opérations** avec barre de progression style Créer Module
- **Contrôles**: Pause/Stop/Annuler pendant l'upload
- **Styles DataGrid** avec headers fond sombre et texte bleu XNRGY
- **Interface moderne** avec thèmes sombre/clair

### 2. Créer Module - Copy Design (100%)

**Sources disponibles :**
- Depuis Template : `$/Engineering/Library/Xnrgy_Module` (1083 fichiers Inventor)
- Depuis Projet Existant : Sélection d'un projet local ou Vault

**Workflow automatisé :**
1. Switch vers projet source (IPJ)
2. Ouverture Top Assembly (Module_.iam)
3. Application iProperties sur le template
4. Collecte de toutes les références (bottom-up)
5. Copy Design natif avec SaveAs (IPT -> IAM -> Top Assembly)
6. Traitement des dessins (.idw) avec mise à jour des références
7. **Mise à jour des références des composants suppressed** (v1.1)
8. Copie des fichiers orphelins (1059 fichiers non-référencés)
9. Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)
10. Renommage du fichier .ipj
11. Switch vers le nouveau projet
12. Application des iProperties finales et paramètres Inventor
13. Design View -> "Default", masquage Workfeatures
14. Vue ISO + Zoom All (Fit)
15. Update All (rebuild) + Save All
16. Module reste ouvert pour le dessinateur

**Gestion intelligente des références :**
- Fichiers Library (IPT_Typical_Drawing) : Liens préservés
- Fichiers Module : Copies avec références mises à jour
- Fichiers IDW : Références corrigées via `PutLogicalFileNameUsingFull`
- **Composants suppressed** : Références mises à jour même si supprimés

**Options de renommage (v1.1) :**
- Rechercher/Remplacer (cumulatif sur NewFileName)
- Préfixe/Suffixe (appliqué sur OriginalFileName)
- **Checkbox "Inclure fichiers non-Inventor"**

### 3. Réglages Admin (100%)

**Système de configuration centralisée :**
- Chiffrement AES-256 des fichiers de configuration
- Synchronisation automatique via Vault au démarrage
- Accès restreint aux administrateurs (Role "Administrator" ou Groupe "Admin_Designer")
- Déploiement multi-sites : Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs

**Chemin Vault :**
```
$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/
```

**Sections configurables :**
- Liste des initiales designers (26 entrées + "Autre...")
- Chemins templates et projets
- Extensions Inventor supportées
- Dossiers/fichiers exclus
- Noms des iProperties

**Interface moderne :**
- Styles uniformisés avec effets glow sur boutons
- GroupBox avec titres orange (#FF8C00)
- Thèmes sombre/clair cohérents

### 4. Upload Template (100%)

- **Réservé aux administrateurs** - Message XnrgyMessageBox si non-admin
- **Upload templates** depuis Library vers Vault
- **Utilise la connexion partagée** de l'app principale
- **Journal intégré** avec barre de progression
- **Interface alignée** avec Upload Module (hauteurs, styles, thèmes)

### 5. Checklist HVAC (100%)

- Validation des modules AHU
- Checklist interactive avec critères XNRGY
- Stockage des validations dans Vault
- Interface WebView2 pour affichage HTML

### 6. Connexions Automatiques

- **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique
- **Inventor Professional 2026.2** - COM avec détection d'instance active
- **Throttling intelligent** pour éviter spam logs (v1.1)
- **Vérification fenêtre Inventor** prête avant connexion COM
- **Update Workspace** - Synchronisation dossiers au démarrage :
  - `$/Content Center Files`
  - `$/Engineering/Inventor_Standards`
  - `$/Engineering/Library/Cabinet`
  - `$/Engineering/Library/Xnrgy_M99`
  - `$/Engineering/Library/Xnrgy_Module`

---

## Services Principaux

### VaultSDKService.cs

Service principal pour l'interaction avec Vault SDK (~1200 lignes).

**Responsabilités :**
- Connexion/déconnexion Vault
- Chargement des Property Definitions
- Chargement des Catégories
- Chargement des Lifecycle Definitions
- Upload de fichiers avec `FileManager.AddFile`
- Application des propriétés via `UpdateFileProperties`
- Synchronisation Vault -> iProperties via `IExplorerUtil.UpdateFileProperties`
- Assignation de catégories via `UpdateFileCategories`
- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (reflection)
- Assignation de révisions via `UpdateFileRevisionNumbers`
- Gestion des erreurs Vault (1003, 1013, 1136, etc.)

### InventorService.cs

Service pour la connexion COM à Inventor.

**Améliorations v1.1 :**
- Throttling intelligent (minimum 2 sec entre tentatives)
- Vérification fenêtre Inventor prête (MainWindowHandle != IntPtr.Zero)
- Logs silencieux pour COMException 0x800401E3
- Compteur d'échecs consécutifs avec log périodique

### InventorCopyDesignService.cs

Service pour Copy Design natif avec gestion des références.

**Méthode principale :**
```csharp
Task<bool> ExecuteRealPackAndGoAsync(
    string templatePath,
    string destinationPath,
    string projectNumber,
    string reference,
    string module,
    IProgress<string> progress
)
```

### Services de Propriétés

- **InventorPropertyService** : iProperties via Inventor COM
- **ApprenticePropertyService** : Lecture iProperties via Apprentice (sans Inventor ouvert)
- **OlePropertyService** : Lecture OLE via OpenMCDF
- **NativeOlePropertyService** : Lecture OLE native (Windows API)
- **WindowsPropertyService** : Lecture propriétés Windows Shell

### Autres Services

- **Logger.cs** : Logging NLog avec fichiers UTF-8
- **ThemeHelper.cs** : Gestion thèmes sombre/clair
- **UserPreferencesManager.cs** : Persistance préférences utilisateur
- **VaultSettingsService.cs** : Configuration chiffrée + sync Vault
- **JournalColorService.cs** : Couleurs uniformes pour journal
- **CredentialsManager.cs** : Gestion credentials chiffrées AES-256

---

## Propriétés XNRGY

Le système extrait automatiquement les propriétés depuis le chemin de fichier:

```
C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]
                              |         |       |
Vault Property IDs:        ID=112    ID=121  ID=122
```

| Propriété | ID Vault | Description |
|-----------|----------|-------------|
| Project | 112 | Numéro de projet (5 chiffres) |
| Reference | 121 | Numéro de référence (2 chiffres) |
| Module | 122 | Numéro de module (2 chiffres) |

### Mapping Catégorie -> Lifecycle Definition

| Catégorie | Lifecycle Definition |
|-----------|---------------------|
| Engineering | Flexible Release Process |
| Office | Simple Release Process |
| Standard | Basic Release Process |
| Base | (aucun) |

---

## Interface Utilisateur

### Thèmes

L'application supporte deux thèmes avec propagation automatique vers toutes les fenêtres :

- **Thème Sombre** (défaut) : Fond #1E1E2E, panneaux #252536
- **Thème Clair** : Fond #F5F7FA, panneaux #FCFDFF

**Éléments à fond FIXE** (même en thème clair) :
- Journal : #1A1A28 (noir)
- Panneaux statistiques : #1A1A28 (noir)
- Headers GroupBox : #2A4A6F (bleu marine) avec texte blanc

### Styles Centralisés

Le fichier `Styles/XnrgyStyles.xaml` contient 50+ styles partagés :

- **Couleurs** : Palette XNRGY officielle
- **Boutons** : PrimaryButton, SecondaryButton, SuccessButton, WarningButton, DangerButton avec effets glow
- **Inputs** : ModernTextBox, ModernComboBox avec bordures #4A7FBF
- **DataGrid** : Headers fond sombre #1A1A28, texte bleu #0078D4
- **Containers** : XnrgyGroupBox avec header bleu marine
- **ProgressBar** : Style personnalisé avec fill et shine
- **Labels** : ModernLabel avec tailles et poids uniformisés

### Effets Visuels

- **Glow Effects** : DropShadowEffect bleu brillant (#00D4FF) sur hover des boutons
- **Animations** : Scale animations sur press des boutons
- **Bordures** : #4A7FBF (bleu brillant) sur tous les éléments interactifs

---

## Prérequis

- **Windows 10/11 x64**
- **.NET Framework 4.8**
- **Autodesk Vault Professional 2026** (SDK v31.0.84)
- **Autodesk Inventor Professional 2026.2**
- **Visual Studio 2022** (pour compilation)
- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)

---

## Compilation et Exécution

### Script automatique (RECOMMANDÉ)

```powershell
cd XnrgyEngineeringAutomationTools
.\build-and-run.ps1
```

**Fonctionnalités du script :**
- [+] Compilation automatique en mode Release
- [+] Détection automatique de MSBuild (VS 2022 Enterprise/Professional/Community)
- [+] Arrêt automatique de l'instance existante (taskkill /F)
- [+] Lancement automatique après compilation réussie
- [+] Affichage des erreurs de compilation si présentes

**Options disponibles :**
```powershell
.\build-and-run.ps1              # Build Release + Run
.\build-and-run.ps1 -Debug       # Build Debug + Run
.\build-and-run.ps1 -Clean       # Clean + Build Release + Run
.\build-and-run.ps1 -BuildOnly   # Build sans lancer
.\build-and-run.ps1 -KillOnly    # Tuer les instances existantes
```

### MSBuild manuel

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `
  XnrgyEngineeringAutomationTools.csproj /t:Rebuild /p:Configuration=Release /m /v:minimal /nologo
```

> **[!] IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne génère pas les fichiers .g.cs pour WPF .NET Framework 4.8.

---

## Exclusions de fichiers

**Extensions exclues:**
- `.v`, `.bak`, `.old` (Backup Vault)
- `.tmp`, `.temp` (Temporaires)
- `.ipj` (Projet Inventor)
- `.lck`, `.lock`, `.log` (Système/logs)
- `.dwl`, `.dwl2` (AutoCAD locks)

**Préfixes exclus:**
- `~$` (Office temporaire)
- `._` (macOS temporaire)
- `Backup_` (Backup générique)
- `.~` (Temporaire générique)

**Dossiers exclus:**
- `OldVersions`, `oldversions`
- `Backup`, `backup`
- `.vault`, `.git`, `.vs`

---

## Logs et Debugging

### Emplacement des logs

```
bin\Release\Logs\VaultSDK_POC_YYYYMMDD_HHMMSS.log
```

### Format des logs

```
[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] [+] Message
```

**Niveaux:** INFO, DEBUG, SUCCESS, WARN, ERROR

**Icônes textuelles utilisées (pas d'emoji dans les logs):**
- `[+]` = Succès
- `[-]` = Erreur
- `[!]` = Avertissement
- `[>]` = Action en cours
- `[i]` = Information
- `[~]` = Attente/Polling
- `[#]` = Liste/Propriétés
- `[?]` = Vérification

---

## Dépendances NuGet

| Package | Version | Usage |
|---------|---------|-------|
| CommunityToolkit.Mvvm | 8.2.2 | MVVM helpers |
| Microsoft.Extensions.DependencyInjection | 7.0.0 | DI container |
| Microsoft.Extensions.Logging | 7.0.0 | Logging abstractions |
| Newtonsoft.Json | 13.0.3 | Sérialisation JSON |
| NLog | 5.2.8 | Logging framework |
| OpenMcdf | 2.3.1 | OLE Compound Documents |
| Microsoft.Web.WebView2 | 1.0.2478.35 | WebView2 pour HTML |

---

## Dépendances Autodesk

| DLL | Chemin | Usage |
|-----|--------|-------|
| Autodesk.Connectivity.WebServices | Vault 2026 SDK | Vault API |
| Autodesk.DataManagement.Client.Framework | Vault 2026 SDK | Vault Framework |
| Autodesk.Inventor.Interop | Inventor 2026 | Inventor COM API |
| Autodesk.Connectivity.Explorer.ExtensibilityTools | Vault 2026 SDK | IExplorerUtil |

---

## Dépanage

### L'application ne démarre pas
- Vérifier .NET Framework 4.8 installé
- Vérifier Vault Professional 2026 installé
- Vérifier les dépendances NuGet restaurées

### Erreur de connexion Vault
- Vérifier serveur accessible
- Vérifier vault existe
- Vérifier identifiants
- Voir logs dans `bin\Release\Logs\`

### Erreur connexion Inventor (0x800401E3)
- Inventor doit être **complètement démarré** (fenêtre principale visible)
- L'app attend que Inventor s'enregistre dans la Running Object Table (ROT)
- Le timer de reconnexion réessaie automatiquement toutes les 3 secondes

### Propriétés non appliquées
- Vérifier logs : rechercher "Application des propriétés"
- Si erreur 1003 : Fichier en traitement par Job Processor (normal)
- Si erreur 1013 : CheckOut nécessaire (automatique)
- Vérifier que les Property Definitions sont chargées (Project, Reference, Module)
- Pour fichiers Inventor : Vérifier que `IExplorerUtil` est chargé
- Pour writeback iProperties : Vérifier que le writeback est activé dans Vault

### Headers DataGrid invisibles
- Les styles DataGrid sont définis dans Window.Resources
- Fond sombre (#1A1A28) avec texte bleu XNRGY (#0078D4)
- Style appliqué globalement via `<Style TargetType="DataGridColumnHeader">`

### Titres GroupBox ne changent pas avec le thème
- Les titres restent blancs même en thème clair (fond header bleu marine fixe)
- Correction appliquée dans `MainWindow.xaml.cs` et `XnrgyStyles.xaml`

---

## Configuration

### appsettings.json (local)

```json
{
  "VaultConfig": {
    "Server": "VAULTPOC",
    "Vault": "PROD_XNGRY",
    "User": "username",
    "Password": ""
  },
  "Paths": {
    "DefaultLibrary": "$/Engineering/Library",
    "DefaultTemplate": "$/Engineering/Library/Xnrgy_Module",
    "ProjectsRoot": "C:\\Vault\\Engineering\\Projects"
  }
}
```

---

## Changelog

### v1.0.0 (31 Décembre 2025) - RELEASE OFFICIELLE ACTUELLE

**[+] Upload Module intégré:**
- Module VaultAutomationTool intégré dans l'app principale (`Modules/UploadModule/`)
- Interface avec deux DataGrids (Inventor/Non-Inventor)
- Styles DataGrid avec headers visibles (fond sombre #1A1A28, texte bleu #0078D4)
- Barre de progression et journal des opérations style Créer Module
- Utilise la connexion Vault partagée (pas de login séparé)
- Contrôles Pause/Stop/Annuler

**[+] Upload Template:**
- Nouvelle fenêtre pour upload templates (réservé Admin)
- Utilise connexion partagée de l'app principale
- XnrgyMessageBox si utilisateur non-admin

**[+] Corrections Inventor:**
- Throttling intelligent pour éviter spam logs
- Vérification fenêtre Inventor prête avant connexion COM
- Logs silencieux pour COMException 0x800401E3
- Timer de reconnexion optimisé

**[+] VaultBulkUploader:**
- Outil console pour upload massif (6152 fichiers uploadés vers PROD_XNGRY)
- Situé dans `Tools/VaultBulkUploader/`

### v1.0.1 (01 Janvier 2026) - Organisation Structure (Développement)

**[+] Organisation professionnelle:**
- Nouveau fichier ARCHITECTURE.md avec plan de structure
- Création dossiers modules: Modules/UploadTemplate, Modules/CreateModule, Modules/ChecklistHVAC
- Création dossier Shared/ pour composants partagés
- Dossier Temp&Test/ pour fichiers de debug/test
- Documentation structure actuelle et cible

**[+] Styles centralisés:**
- XnrgyStyles.xaml: 50+ styles collectés de toutes les fenêtres
- 11 sections: Colors, Buttons, Inputs, DataGrid, Containers, Labels, ProgressBar, DatePicker, ListBox, Special, Responsive
- Effets glow sur boutons au hover
- Aliases pour compatibilité (XnrgyPrimaryButton = PrimaryButton)

**[+] Backup automatique:**
- Dossier Backups/ avec sauvegardes horodatées
- Format: BACKUP_YYYYMMDD_HHMMSS/

**[+] Système de Thèmes Amélioré:**
- Propagation thème light/dark vers TOUTES les sous-fenêtres
- Event MainWindow.ThemeChanged déclenche ApplyTheme() partout
- Éléments à fond FIXE noir (#1A1A28): journal, panneaux stats
- Texte journal toujours blanc (#DCDCDC) quel que soit le thème
- UserPreferencesManager pour persistance locale du thème

**[+] Standardisation Titres Fenêtres:**
- Format uniforme: [Nom] - v1.0.0 - Released on 2026-01-02 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC
- Appliqué à toutes les 9 fenêtres de l'application

**[+] Couleurs Thème:**
| Élément | Dark | Light |
|---------|------|-------|
| MainGrid | #1E1E2E | #F5F7FA |
| Stats/Log | #1A1A28 (FIXE) | #1A1A28 (FIXE) |
| Headers | #2A4A6F | #2A4A6F |

**[+] Créer Module - Copy Design:**
- Copy Design natif avec 1133 fichiers
- Gestion des fichiers orphelins (1059 fichiers)
- Mise à jour références IDW
- Switch IPJ automatique
- Application iProperties et paramètres Inventor
- Design View "Default" + Workfeatures cachés
- Vue ISO + Zoom All + Save All
- Module reste ouvert pour le dessinateur

**[+] Vault Upload:**
- Upload complet avec propriétés automatiques
- Gestion Inventor et non-Inventor séparée
- Catégories, lifecycle et révisions
- Synchronisation Vault -> iProperties via IExplorerUtil

**[+] Réglages Admin:**
- Chiffrement AES-256
- Synchronisation automatique via Vault
- Interface graphique avec validation temps réel

**[+] Connexions automatiques:**
- Vault SDK v31.0.84
- Inventor COM 2026.2
- Update Workspace au démarrage

### v0.9.0 (15 Décembre 2025)

- Release initiale beta
- Dashboard principal avec boutons modules
- Connexion Vault centralisée
- Thèmes sombre/clair

---

## Documentation

- **README.md** - Ce fichier (documentation utilisateur)
- **ARCHITECTURE.md** - Structure projet et plan de migration
- **.github/instructions/XnrgyEngineeringAutomationTools.instructions.md** - Instructions pour Copilot

---

## Auteur

**Mohammed Amine Elgalai**  
Engineering Automation Developer  
XNRGY Climate Systems ULC  
Email: mohammedamine.elgalai@xnrgy.com

---

## Licence

Propriétaire - XNRGY Climate Systems ULC (c) 2025-2026

---

**Dernière mise à jour**: 02 Janvier 2026 - v1.0.0 (Release actuelle)

---

## Versions en Développement (Non publiées)

### v1.1.0 (En développement) - Améliorations UI/UX

**[+] Uniformisation Interface:**
- Hauteurs de boutons alignées avec barre de progression (34px)
- Tailles de police uniformisées (13px) avec padding ajusté
- Cases commentaire alignées entre UploadModule et UploadTemplate (36px)
- Titres GroupBox restent blancs même en thème clair (fond bleu marine fixe)
- Effets glow appliqués à tous les boutons (CreateModuleSettingsWindow)
- Styles parent appliqués à CreateModuleSettingsWindow (ModernTextBox, ModernLabel, ModernGroupBox)
- Titres GroupBox orange (#FF8C00) dans CreateModuleSettingsWindow

**[+] Corrections Styles:**
- Bordures GroupBox changées de White à #4A7FBF
- VerticalContentAlignment="Center" sur tous les boutons
- ProgressBar hauteur réduite à 34px pour alignement
