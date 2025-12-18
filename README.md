# XNRGY Engineering Automation Tools# VaultAutomationTool



🏭 **Suite d'outils d'automatisation engineering unifiée** pour piloter Autodesk Vault Professional 2026 et Inventor Professional 2026.2Application WPF pour l'upload automatisé de fichiers vers Autodesk Vault Professional 2026 avec application automatique des propriétés métier (Project, Reference, Module), catégories, lifecycle et révisions.



## 📋 Description## 📋 Description



Application hub centralisée qui regroupe tous les outils d'automatisation engineering XNRGY :Cette application permet de :

- Scanner automatiquement les modules engineering (structure `Projects\[NUMERO]\REF[NUM]\M[NUM]`)

- **Vault Upload** - Upload automatisé vers Vault avec propriétés (Project/Reference/Module)- Uploader des fichiers vers Vault avec création automatique de l'arborescence

- **Pack & Go** - GET depuis Vault, insertion dans assemblages, Copy Design- Appliquer automatiquement les propriétés métier (Project, Reference, Module)

- **Smart Tools** - Création IPT/STEP, génération PDF, iLogic Forms- Assigner des catégories, lifecycle definitions/states et révisions

- **DXF Verifier** - Validation des fichiers DXF avant envoi- Gérer l'upload de fichiers Inventor et non-Inventor séparément

- **Checklist HVAC** - Validation modules AHU avec stockage Vault

- **Update Workspace** - Synchronisation des librairies depuis Vault## 🎯 Caractéristiques



## 🎯 Fonctionnalités- ✅ Connexion directe à Vault via SDK (VaultSDKService.cs)

- ✅ Scan automatique des modules engineering (FileScanner.cs)

### Connexions Automatiques- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

- ✅ Connexion centralisée à **Vault Professional 2026** (SDK v31.0.84)- ✅ Application automatique des propriétés métier extraites du chemin

- ✅ Connexion COM à **Inventor Professional 2026.2**- ✅ Assignation de catégories, lifecycle definitions/states et révisions

- ✅ Détection automatique d'Inventor s'il est en cours d'exécution- ✅ Gestion de la progression et pause/reprise

- ✅ Logs détaillés UTF-8 avec emoji (Logger.cs)

### Update Workspace (GET automatique)- ✅ Exclusion automatique des fichiers temporaires (.bak, .dwl, .log, OldVersions, ~$)

Au démarrage ou sur demande, synchronisation des dossiers essentiels :- ✅ Sauvegarde configuration (appsettings.json)

- `$/Content Center Files` → `C:\Vault\Content Center Files`- ✅ Interface MVVM avec séparation Inventor/Non-Inventor

- `$/Engineering/Inventor_Standards` → `C:\Vault\Engineering\Inventor_Standards`

- `$/Engineering/Library/Cabinet` → `C:\Vault\Engineering\Library\Cabinet`## 📦 Prérequis

- `$/Engineering/Library/Xnrgy_M99` → `C:\Vault\Engineering\Library\Xnrgy_M99`

- `$/Engineering/Library/Xnrgy_Module` → `C:\Vault\Engineering\Library\Xnrgy_Module`- Windows 10/11 x64

- .NET Framework 4.8

## 📦 Modules Intégrés- Autodesk Vault Professional 2026

- Visual Studio 2022 ou supérieur (pour compilation)

| Module | Description | Status |- MSBuild 18.0.0+ (REQUIS - dotnet build ne fonctionne PAS pour WPF)

|--------|-------------|--------|

| 📤 Vault Upload | Upload avec propriétés automatiques | ✅ Intégré |## 🏗️ Structure du projet

| 📦 Pack & Go | GET Vault + Copy Design | 🚧 En développement |

| ⚡ Smart Tools | IPT, STEP, PDF, iLogic | 🚧 En développement |```

| 📐 DXF Verifier | Validation fichiers DXF | 🚧 Migration |VaultAutomationTool/

| ✅ Checklist HVAC | Validation AHU + Vault | 🚧 Migration |├── Models/                          # Modèles de données (10 fichiers)

│   ├── ApplicationConfiguration.cs  # Configuration application

## 📦 Prérequis│   ├── CategoryItem.cs             # Item catégorie pour ComboBox

│   ├── FileItem.cs                 # Item fichier pour DataGrid

- **Windows 10/11 x64**│   ├── FileToUpload.cs             # Fichier à uploader

- **.NET Framework 4.8**│   ├── LifecycleDefinitionItem.cs  # Lifecycle Definition

- **Autodesk Vault Professional 2026** (SDK v31.0.84)│   ├── LifecycleStateItem.cs       # Lifecycle State

- **Autodesk Inventor Professional 2026.2**│   ├── ModuleInfo.cs               # Informations module

- **Visual Studio 2022** (pour compilation)│   ├── ProjectInfo.cs              # Informations projet

│   ├── ProjectProperties.cs        # Propriétés Project/Ref/Module

## 🏗️ Structure du Projet│   └── VaultConfiguration.cs       # Configuration Vault

├── Services/                        # Services métier (2 fichiers)

```│   ├── VaultSDKService.cs         # Service principal Vault SDK

XnrgyEngineeringAutomationTools/│   └── Logger.cs                   # Système logging UTF-8

├── MainWindow.xaml              # Dashboard principal├── ViewModels/                      # ViewModels MVVM (1 fichier)

├── App.xaml                     # Configuration WPF│   ├── AppMainViewModel.cs         # ViewModel principal

├── Assets/│   └── RelayCommand.cs             # Implémentation ICommand

│   └── Icons/                   # Icônes des modules├── Properties/

├── Modules/│   └── AssemblyInfo.cs             # Informations assembly

│   ├── VaultUpload/            # Module upload Vault├── App.xaml(.cs)                   # Point d'entrée application

│   ├── PackAndGo/              # Module Pack & Go├── MainWindow.xaml(.cs)            # Fenêtre principale

│   ├── SmartTools/             # Module Smart Tools├── appsettings.json                # Configuration sauvegardée

│   ├── DXFVerifier/            # Module DXF Verifier├── README.md                        # Ce fichier

│   └── ChecklistHVAC/          # Module Checklist HVAC└── bin/Release/                     # Exécutable compilé

├── Services/    ├── VaultAutomationTool.exe     # Application

│   ├── VaultSdkService.cs      # Service Vault SDK    └── Logs/                       # Logs d'exécution UTF-8

│   ├── InventorService.cs      # Service Inventor COM        └── VaultSDK_POC_YYYYMMDD_HHMMSS.log

│   └── Logger.cs               # Système de logs```

├── Views/

│   ├── LoginWindow.xaml        # Fenêtre connexion## 🔧 Architecture

│   └── VaultUploadWindow.xaml  # Fenêtre upload Vault

└── ViewModels/                  # MVVM ViewModels### Pattern MVVM (Model-View-ViewModel)

```

- **Models** : Données et configuration

## 🚀 Compilation et Lancement- **Views** : Interface utilisateur XAML (MainWindow.xaml)

- **ViewModels** : Logique métier et binding (AppMainViewModel.cs)

### Script automatique- **Services** : Accès aux données Vault (VaultSDKService.cs)

```powershell

cd XnrgyEngineeringAutomationTools### Services principaux

.\build-and-run.ps1

```#### 1. VaultSDKService.cs



### MSBuild manuelService principal pour l'interaction avec Vault SDK.

```powershell

& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `**Responsabilités** :

  XnrgyEngineeringAutomationTools.csproj /p:Configuration=Release /t:Rebuild- Connexion/déconnexion Vault

```- Chargement des Property Definitions

- Chargement des Catégories

## 🔌 APIs Utilisées- Chargement des Lifecycle Definitions

- Upload de fichiers avec `FileManager.AddFile`

### Vault SDK 2026- Application des propriétés via `UpdateFileProperties`

- `VDF.Vault.Library.ConnectionManager` - Connexion- Synchronisation des propriétés Vault → iProperties via `IExplorerUtil.UpdateFileProperties` (pour fichiers Inventor)

- `VDF.Vault.Currency.Connections.Connection` - Session- Assignation de catégories via `UpdateFileCategories`

- `FileManager.AddFile()` - Upload- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (via reflection)

- `FileManager.AcquireFiles()` - Download (GET)- Assignation de révisions via `UpdateFileRevisionNumbers`

- `DocumentService.UpdateFileProperties()` - Propriétés- Gestion des erreurs Vault (1003, 1013, 1136, etc.)



### Inventor 2026.2 COM**Méthodes principales** :

- `Inventor.Application` via `GetActiveObject()````csharp

- `Application.ActiveDocument` - Document actifbool Connect(string server, string vaultName, string username, string password)

- iProperties via `Document.PropertySets`void Disconnect()

List<(long Id, string Name)> GetAvailableCategories()

## 👤 AuteurList<LifecycleDefinitionItem> GetAvailableLifecycleDefinitions()

long? GetLifecycleDefinitionIdByCategory(string categoryName)

**Mohammed Amine Elgalai**  long? GetWorkInProgressStateId(long lifecycleDefinitionId)

Smart Tools Amine - XNRGY Climate Systems ULC  bool UploadFile(string filePath, string vaultFolderPath, 

Email: mohammedamine.elgalai@xnrgy.com    string? projectNumber = null, string? reference = null, string? module = null,

    long? categoryId = null, string? categoryName = null,

## 📄 Version    long? lifecycleDefinitionId = null, long? lifecycleStateId = null, string? revision = null)

```

**v1.0.0** - Décembre 2025

**Gestion des propriétés** :

### Historique- Propriétés XNRGY : Project (ID: 112), Reference (ID: 121), Module (ID: 122)

- **v1.0.0** (17 Décembre 2025) : Version initiale- Chargement automatique des Property Definitions au démarrage

  - Dashboard principal avec modules- Application via `UpdateFileProperties` (nécessite CheckOut pour fichiers existants)

  - Connexion Vault & Inventor centralisée- **Synchronisation Vault → iProperties** : Utilisation de `IExplorerUtil.UpdateFileProperties` pour les fichiers Inventor

  - Update Workspace automatique  - Chargement lazy d'ExplorerUtil si nécessaire

  - Module Vault Upload intégré  - Writeback automatique des propriétés Vault vers les iProperties Inventor

  - Nécessite que le writeback soit activé dans Vault (`GetEnableItemPropertyWritebackToFiles`)

## 📜 Licence

**Gestion du lifecycle** :

Propriétaire - XNRGY Climate Systems ULC- Utilisation de `DocumentServiceExtensions.UpdateFileLifeCycleDefinitions` via reflection

- Support de différentes signatures de SDK

---- Assignation directe sans CheckOut pour nouveaux fichiers

**Dernière mise à jour** : 17 Décembre 2025

#### 2. Logger.cs

Système de logging UTF-8 avec emoji.

**Niveaux de log** :
- **TRACE** : Détails techniques très fins
- **DEBUG** : Informations de débogage détaillées
- **INFO** : Opérations importantes (connexion, upload, succès)
- **WARNING** : Avertissements non bloquants
- **ERROR** : Erreurs bloquantes
- **CRITICAL** : Erreurs critiques système

**Format des logs** :
```
[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] Message avec emoji
```

**Emoji utilisés** :
- 🔌 = Connexion
- ✅ = Succès
- ❌ = Erreur
- ⚠️ = Avertissement
- 📋 = Liste/Propriétés
- 📊 = Statistiques
- ⏳ = Attente/Polling
- 🔍 = Vérification
- 📄 = Fichier
- 🔓 = CheckOut
- 💾 = Mise à jour
- 🔒 = CheckIn
- 📤 = Upload
- 🔄 = Traitement
- 💡 = Info

### ViewModels

#### AppMainViewModel.cs

ViewModel principal avec toutes les propriétés et commandes.

**Propriétés principales** :
- `IsConnected` : État de connexion Vault
- `IsProcessing` : État de traitement (scan/upload)
- `StatusMessage` : Message de statut
- `ProgressValue` : Valeur de progression (0-100)
- `InventorFiles` : Collection fichiers Inventor
- `NonInventorFiles` : Collection fichiers non-Inventor
- `AvailableCategories` : Catégories disponibles
- `SelectedCategoryInventor` / `SelectedCategoryNonInventor` : Catégories sélectionnées
- `AvailableLifecycleDefinitions` : Lifecycle Definitions disponibles
- `SelectedLifecycleDefinitionInventor` / `SelectedLifecycleDefinitionNonInventor` : Lifecycle Definitions sélectionnées
- `AvailableStatesInventor` / `AvailableStatesNonInventor` : États disponibles
- `SelectedLifecycleStateInventor` / `SelectedLifecycleStateNonInventor` : États sélectionnés
- `RevisionInventor` / `RevisionNonInventor` : Révisions saisies

**Commandes** :
- `ToggleConnectionCommand` : Connexion/déconnexion Vault
- `ScanProjectCommand` : Scan d'un module
- `AutoCheckInCommand` : Upload des fichiers sélectionnés
- `PauseCommand` : Pause/reprise du traitement

**Méthodes principales** :
```csharp
void ToggleConnection()
void ScanProject(string projectPath)
async Task AutoCheckInAsync()
void UpdateAvailableStates() // Met à jour les états selon la Lifecycle Definition sélectionnée
```

### Models

#### FileItem.cs
Représente un fichier à uploader avec :
- `IsChecked` : Sélectionné pour upload
- `FullPath` : Chemin complet
- `FileName` : Nom du fichier
- `Extension` : Extension
- `Category` : Catégorie (Inventor/Non-Inventor)

#### ProjectProperties.cs
Propriétés extraites du chemin :
- `Project` : Numéro de projet
- `Reference` : Numéro de référence
- `Module` : Numéro de module

#### CategoryItem.cs
Catégorie Vault avec :
- `Id` : ID de la catégorie
- `Name` : Nom de la catégorie

#### LifecycleDefinitionItem.cs
Lifecycle Definition avec :
- `Id` : ID de la définition
- `Name` : Nom de la définition
- `States` : Collection des états disponibles

#### LifecycleStateItem.cs
Lifecycle State avec :
- `Id` : ID de l'état
- `Name` : Nom de l'état
- `IsDefault` : État par défaut

## 🔌 API Vault SDK utilisées

### Connexion
```csharp
VDF.Vault.Library.ConnectionManager.LogIn(
    server, vaultName, username, password,
    VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, null
)
```

### Upload de fichiers
```csharp
_connection.FileManager.AddFile(
    targetFolder, fileName, null, lastWriteTime, null, null,
    fileClassification, false, fileStream
)
```

### Application des propriétés
```csharp
// Pour nouveaux fichiers (sans CheckOut)
_connection.WebServiceManager.DocumentService.UpdateFileProperties(
    new[] { file.Id }, new[] { propArray }
)

// Pour fichiers existants (nécessite CheckOut)
_connection.WebServiceManager.DocumentService.CheckoutFile(...)
_connection.WebServiceManager.DocumentService.UpdateFileProperties(...)
_connection.FileManager.CheckinFile(...)
```

### Assignation de catégories
```csharp
// Via DocumentServiceExtensions (via reflection)
var documentServiceExtensions = _connection.WebServiceManager.DocumentServiceExtensions;
documentServiceExtensions.UpdateFileCategories(
    new[] { file.Id }, new[] { categoryId }
)
```

### Assignation de lifecycle
```csharp
// Via DocumentServiceExtensions (via reflection)
var documentServiceExtensions = _connection.WebServiceManager.DocumentServiceExtensions;
documentServiceExtensions.UpdateFileLifeCycleDefinitions(
    new[] { file.Id },
    new[] { lifecycleDefinitionId },
    new[] { lifecycleStateId },
    "Commentaire"
)
```

### Gestion des erreurs Vault

**Erreur 1003** : Fichier en traitement par Job Processor
- **Solution** : Retour immédiat sans attente (pas de polling)

**Erreur 1013** : Fichier doit être checké out pour modification
- **Solution** : CheckOut → Update → CheckIn

**Erreur 1008** : Fichier existe déjà
- **Solution** : Récupérer le fichier existant et appliquer les modifications

**Erreur 1136** : Restriction lifecycle
- **Solution** : Vérifier les permissions et l'état du fichier

## 📝 Flux d'upload

### 1. Scan du module
- Chemin attendu : `...\Engineering\Projects\[NUMERO]\REF[NUM]\M[NUM]`
- Extraction automatique : Project, Reference, Module
- Scan récursif avec exclusions (fichiers temporaires, dossiers système)

### 2. Sélection des fichiers
- Séparation Inventor / Non-Inventor
- Sélection par défaut de tous les fichiers
- Filtres de recherche disponibles

### 3. Configuration
- Sélection de la catégorie (Base par défaut)
- Sélection de la Lifecycle Definition (selon catégorie)
- Sélection de l'état (selon Lifecycle Definition)
- Saisie de la révision (manuel pour l'instant)

### 4. Upload
- Création de l'arborescence Vault si nécessaire
- Upload du fichier avec `FileManager.AddFile` (commentaire personnalisé pour la version 1)
- Assignation de la catégorie (si spécifiée)
- Assignation du lifecycle (si spécifié)
- Assignation de la révision (si spécifiée) via `UpdateFileRevisionNumbers`
- Application des propriétés (Project, Reference, Module)
- Synchronisation Vault → iProperties pour fichiers Inventor (si `IExplorerUtil` disponible)

### 5. Gestion des fichiers existants
- Détection du fichier existant
- CheckOut si nécessaire
- Application des modifications
- CheckIn pour valider

## ⚙️ Configuration

### appsettings.json
```json
{
  "VaultConfig": {
    "Server": "VAULTPOC",
    "Vault": "TestXNRGY",
    "User": "mohammedamine.elgalai",
    "Password": ""  // Sauvegardé si "Sauvegarder identifications" coché
  }
}
```

### Mapping Catégorie → Lifecycle Definition

Dans `VaultSDKService.cs`, méthode `GetLifecycleDefinitionIdByCategory` :
- **Engineering** → Flexible Release Process
- **Office** → Simple Release Process
- **Standard** → Basic Release Process
- **Base** → Aucun mapping par défaut

### Exclusions de fichiers

**Extensions exclues** :
- `.v`, `.bak`, `.old` (Backup Vault)
- `.tmp`, `.temp` (Temporaires)
- `.ipj` (Projet Inventor)
- `.lck`, `.lock`, `.log` (Système/logs)
- `.dwl`, `.dwl2` (AutoCAD locks)

**Préfixes exclus** :
- `~$` (Office temporaire)
- `._` (macOS temporaire)
- `Backup_` (Backup générique)
- `.~` (Temporaire générique)

**Dossiers exclus** :
- `OldVersions`, `oldversions`
- `Backup`, `backup`
- `.vault`, `.git`, `.vs`

## 🚀 Compilation

### Avec MSBuild (REQUIS pour WPF)
```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe' `
  'VaultAutomationTool.csproj' `
  /t:Build `
  /p:Configuration=Release `
  /p:Platform=x64
```

### Avec Visual Studio
1. Ouvrir `VaultAutomationTool.sln`
2. Build > Build Solution (Ctrl+Shift+B)
3. Vérifier dans Output pour erreurs

**⚠️ IMPORTANT** : 
- **NE PAS utiliser `dotnet build`** - il ne génère pas correctement les fichiers .g.cs depuis XAML pour WPF
- Seul MSBuild supporte complètement la génération de code WPF

## 📦 Dépendances NuGet

```xml
<PackageReference Include="Autodesk.Connectivity.WebServices" Version="31.0.0" />
<PackageReference Include="Autodesk.DataManagement.Client.Framework" Version="31.0.0" />
```

## 🔍 Détails techniques

### Gestion du FileClassification

Le `FileClassification` est déterminé selon la catégorie sélectionnée :
- **Base** → `FileClassification.None`
- **Engineering** → `FileClassification.None`
- **Design Representation** → `FileClassification.DesignRepresentation`
- Mapping automatique via `DetermineFileClassificationByCategory`

### Application des propriétés

**Pour les fichiers Inventor** :
1. Upload avec `FileManager.AddFile` (avec commentaire personnalisé pour la version 1)
2. GET (téléchargement réel du fichier)
3. CheckOut
4. `UpdateFileProperties` (UDP Vault)
5. `IExplorerUtil.UpdateFileProperties` (writeback Vault → iProperties, si disponible)
6. CheckIn pour persister les propriétés
7. GET final pour mettre à jour le statut du fichier dans Vault Client

**Pour les fichiers non-Inventor** :
1. Upload avec `FileManager.AddFile` (avec commentaire personnalisé pour la version 1)
2. CheckOut
3. `UpdateFileProperties` (UDP Vault)
4. CheckIn pour persister les propriétés
5. GET final pour mettre à jour le statut du fichier dans Vault Client

**Note** : La synchronisation des propriétés (Property Compliance) se fait automatiquement via le Job Processor de Vault après le CheckIn. Le writeback vers iProperties pour les fichiers Inventor nécessite `IExplorerUtil` qui est chargé automatiquement si disponible.

### Assignation du lifecycle via reflection

Le SDK peut avoir différentes signatures pour `UpdateFileLifeCycleDefinitions`. Le code utilise la reflection pour supporter :
- `(long[] fileIds, long[] lifecycleDefIds, long[] lifecycleStateIds, string comment)`
- Autres variantes possibles selon la version SDK

### Synchronisation des propriétés Vault → iProperties (Inventor)

**Stratégie implémentée** :

1. **Upload fichier vers Vault** avec `FileManager.AddFile`
2. **GET** : Téléchargement réel du fichier
3. **CheckOut** : Verrouillage du fichier pour modification
4. **UpdateFileProperties** : Application des UDP (User-Defined Properties) dans Vault
5. **IExplorerUtil.UpdateFileProperties** : Writeback automatique Vault → iProperties (si disponible)
6. **CheckIn** : Persistance des modifications
7. **GET final** : Mise à jour du statut du fichier dans Vault Client

**Avantages** :
- ✅ **UDP Vault correctes** (via UpdateFileProperties)
- ✅ **iProperties Inventor synchronisées** (via IExplorerUtil si disponible)
- ✅ **Statut fichier à jour** dans Vault Client (via GET final)
- ✅ **Pas de rond rouge de synchronisation** après le GET final

**Prérequis** :
- Writeback activé dans Vault (`GetEnableItemPropertyWritebackToFiles` doit retourner `true`)
- `IExplorerUtil` disponible (chargé automatiquement via `ExplorerLoader.LoadExplorerUtil`)

**Note** : Si `IExplorerUtil` n'est pas disponible, les UDP Vault sont toujours appliquées, mais le writeback vers iProperties ne se fait pas automatiquement. La synchronisation se fera via le Job Processor de Vault après le CheckIn.

### Construction du chemin Vault

Le chemin Vault est construit avec les préfixes "REF" et "M" :
- Chemin attendu : `$/Engineering/Projects/12345/REF01/M01`
- Pas : `$/Engineering/Projects/12345/01/01`

## 🐛 Dépannage

### L'application ne démarre pas
- Vérifier .NET Framework 4.8 installé
- Vérifier Vault Professional 2026 installé
- Vérifier les dépendances NuGet restaurées

### Erreur de connexion Vault
- Vérifier serveur accessible
- Vérifier vault existe
- Vérifier identifiants
- Voir logs dans `bin/Release/Logs/` pour détails

### Propriétés non appliquées
- Vérifier logs : rechercher "Application des propriétés"
- Si erreur 1003 : Fichier en traitement par Job Processor (normal pour nouveaux fichiers)
- Si erreur 1013 : CheckOut nécessaire (automatique pour fichiers existants)
- Vérifier que les Property Definitions sont chargées (Project, Reference, Module)
- Pour fichiers Inventor : Vérifier que `IExplorerUtil` est chargé (voir logs "ExplorerUtil chargé")
- Pour writeback iProperties : Vérifier que le writeback est activé dans Vault (`GetEnableItemPropertyWritebackToFiles`)

### Lifecycle non assigné
- Vérifier que la Lifecycle Definition est sélectionnée
- Vérifier que l'état est sélectionné
- Vérifier logs pour erreurs de reflection
- Vérifier permissions Vault

### Catégories non chargées
- Vérifier connexion Vault active
- Vérifier logs pour erreurs `GetCategoriesByEntityClassId`
- "Base" devrait être sélectionnée par défaut

### États non chargés
- Vérifier qu'une Lifecycle Definition est sélectionnée
- `UpdateAvailableStates` est appelé automatiquement lors du changement de Lifecycle Definition

## 📚 Références

- [Autodesk Vault API Documentation](https://www.autodesk.com/developer-network/platform-technologies/vault)
- [MVVM Pattern](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/data/data-binding-overview)
- [WPF Documentation](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)

## 👤 Auteur

**Mohammed Amine Elgalai**  
Smart Tools Amine - XNRGY Climate Systems ULC  
Email: mohammedamine.elgalai@xnrgy.com

## 📄 Version

**v1.0.0** - Décembre 2025 (En développement)

### Historique des versions

- **v1.0.0** (17 Décembre 2025) - Version actuelle en développement :
  
  **🔧 Corrections et stabilisation (17 Décembre 2025)** :
  - ✅ Suppression du listing des jobs Vault historiques qui polluait les logs
  - ✅ Correction de la connexion à Inventor via P/Invoke (`oleaut32.dll` + `ole32.dll`)
  - ✅ Bouton "🔧 Depuis Inventor" : récupère le chemin du document actif dans Inventor
  - ✅ Extraction automatique des propriétés (Project/Reference/Module) depuis le chemin pour les boutons "Depuis Inventor" et "Parcourir"
  - ✅ Les propriétés extraites sont SANS préfixes (ex: `01` au lieu de `REF01` ou `M01`)
  - ✅ Amélioration des scripts `build-and-run.ps1` et `build-and-run.bat` :
    - Force l'arrêt de l'application si elle est en cours (`taskkill /F`)
    - Détection automatique de MSBuild VS 2022 (Enterprise/Professional/Community)
    - Lancement automatique de l'application après compilation
    - Affichage propre des étapes et messages
  
  **📋 Fonctionnalités principales** :
  - Upload automatisé avec propriétés via Vault SDK
  - Scan modules avec exclusion fichiers temporaires
  - Support catégories, lifecycle definitions/states et révisions
  - Séparation Inventor/Non-Inventor dans l'interface
  - Application des propriétés avec CheckOut/CheckIn pour garantir la persistance
  - Synchronisation Vault → iProperties via `IExplorerUtil.UpdateFileProperties` pour fichiers Inventor
  - Commentaire personnalisé pour le premier check-in
  - Assignation de révision via `UpdateFileRevisionNumbers`
  - GET final pour mettre à jour le statut des fichiers dans Vault Client
  - Mapping automatique catégorie → lifecycle definition
  - Gestion améliorée des fichiers existants
  - Logs UTF-8 avec emoji

## 🚀 Compilation et lancement rapide

### Script automatique

Un script PowerShell `build-and-run.ps1` est fourni pour compiler et lancer l'application automatiquement :

```powershell
# Double-clic sur build-and-run.bat ou exécuter dans PowerShell:
.\build-and-run.ps1
```

**Fonctionnalités** :
- ✅ Compilation automatique en mode Release
- ✅ Détection automatique de MSBuild (VS 2022 Professional/Community/Enterprise/Insiders)
- ✅ Arrêt automatique de l'instance existante si déjà en cours
- ✅ Lancement automatique de l'application après compilation réussie
- ✅ Affichage des erreurs de compilation si présentes

**Alternative** : Double-clic sur `build-and-run.bat` (plus simple pour Windows)

## 📜 Licence

Propriétaire - XNRGY Climate Systems ULC

---

**Dernière mise à jour** : 17 Décembre 2025  
**Documentation complète** : Toutes informations projet, architecture, API, dépannage

## 🔄 Changelog détaillé

### v1.0.0 (17 Décembre 2025) - En développement

**🔧 Corrections et stabilisation** :
- ✅ Suppression du listing des jobs Vault historiques (évite les `[WARNING]` inutiles dans les logs)
- ✅ Correction de la connexion à Inventor via P/Invoke natif (`oleaut32.dll` + `ole32.dll`)
- ✅ Bouton "🔧 Depuis Inventor" fonctionne maintenant correctement
- ✅ Extraction automatique des propriétés depuis le chemin pour tous les boutons de sélection
- ✅ Propriétés extraites SANS préfixes (`01` au lieu de `REF01` ou `M01`)

**📝 Scripts de build améliorés** (`build-and-run.ps1` / `build-and-run.bat`) :
- ✅ Force l'arrêt de l'application si elle est en cours d'exécution
- ✅ Détection automatique de MSBuild VS 2022 (Enterprise/Professional/Community)
- ✅ Compilation en mode Release
- ✅ Lancement automatique après compilation réussie
- ✅ Messages clairs et structurés

**🎯 Fonctionnalités validées** :
- ✅ Connexion/Déconnexion Vault SDK
- ✅ Scan des modules engineering
- ✅ Upload fichiers vers Vault avec arborescence automatique
- ✅ Application des propriétés Project/Reference/Module
- ✅ Assignation catégories, lifecycle et révisions
- ✅ Synchronisation Vault → iProperties pour fichiers Inventor
- ✅ GET final pour enlever le rond rouge de synchronisation

---

**Dernière mise à jour** : 17 Décembre 2025  
**Auteur** : Mohammed Amine Elgalai - Smart Tools Amine - XNRGY Climate Systems ULC
