# XNRGY Engineering Automation Tools

> **Suite d'outils d'automatisation engineering unifiee** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2
>
> Developpe par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC

---

## Description

**XNRGY Engineering Automation Tools** est une application hub centralisee (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering developpes pour XNRGY Climate Systems. Cette suite vise a simplifier et accelerer les workflows des equipes de design en integrant la gestion Vault, les manipulations Inventor, et les validations qualite dans une interface unifiee.

### Objectif Principal

Remplacer les multiples applications standalone par une **plateforme unique** avec :
- Connexion centralisee a Vault & Inventor
- Interface utilisateur moderne et coherente (themes sombre/clair)
- Partage de services communs (logging, configuration chiffree AES-256)
- Deploiement multi-sites et maintenance simplifies
- Parametres centralises via Vault (50+ utilisateurs, 3 sites)

---

## Modules Integres

| Module | Description | Statut |
|--------|-------------|--------|
| **Upload Module** | Upload automatise vers Vault avec proprietes (Project/Ref/Module) | [+] 100% |
| **Creer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |
| **Reglages Admin** | Configuration centralisee et synchronisee via Vault (AES-256) | [+] 100% |
| **Upload Template** | Upload templates vers Vault (reserve Admin) | [+] 100% |
| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |
| **Smart Tools** | Creation IPT/STEP, generation PDF, iLogic Forms | [~] Planifie |
| **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |
| **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |
| **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifie |

---

## Fonctionnalites Implementees

### 1. Upload Module (100%) - NOUVEAU v1.1

Module integre (ex-VaultAutomationTool) pour l'upload de fichiers vers Vault:

- **Connexion centralisee** - Utilise la connexion Vault de l'app principale
- **Scan automatique** des modules engineering avec extraction proprietes
- **Separation Inventor/Non-Inventor** dans deux DataGrids avec headers visibles
- **Application automatique** des proprietes metier:
  - Project (ID=112)
  - Reference (ID=121)
  - Module (ID=122)
- **Assignation complete**:

  - Categories Vault
  - Lifecycle Definitions et States
  - Revisions
- **Synchronisation Vault vers iProperties** via `IExplorerUtil`
- **Journal des operations** avec barre de progression style Creer Module
- **Controles**: Pause/Stop/Annuler pendant l'upload
- **Styles DataGrid** avec headers fond sombre et texte bleu XNRGY

### 2. Creer Module - Copy Design (100%)

**Sources disponibles :**
- Depuis Template : `$/Engineering/Library/Xnrgy_Module` (1083 fichiers Inventor)
- Depuis Projet Existant : Selection d'un projet local ou Vault

**Workflow automatise :**
1. Switch vers projet source (IPJ)
2. Ouverture Top Assembly (Module_.iam)
3. Application iProperties sur le template
4. Collecte de toutes les references (bottom-up)
5. Copy Design natif avec SaveAs (IPT -> IAM -> Top Assembly)
6. Traitement des dessins (.idw) avec mise a jour des references
7. **Mise a jour des references des composants suppressed** (v1.1)
8. Copie des fichiers orphelins (1059 fichiers non-references)
9. Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)
10. Renommage du fichier .ipj
11. Switch vers le nouveau projet
12. Application des iProperties finales et parametres Inventor
13. Design View -> "Default", masquage Workfeatures
14. Vue ISO + Zoom All (Fit)
15. Update All (rebuild) + Save All
16. Module reste ouvert pour le dessinateur

**Gestion intelligente des references :**
- Fichiers Library (IPT_Typical_Drawing) : Liens preserves
- Fichiers Module : Copies avec references mises a jour
- Fichiers IDW : References corrigees via `PutLogicalFileNameUsingFull`
- **Composants suppressed** : References mises a jour meme si supprimes

**Options de renommage (v1.1) :**
- Rechercher/Remplacer (cumulatif sur NewFileName)
- Prefixe/Suffixe (applique sur OriginalFileName)
- **Checkbox "Inclure fichiers non-Inventor"**

### 3. Reglages Admin (100%)

**Systeme de configuration centralisee :**
- Chiffrement AES-256 des fichiers de configuration
- Synchronisation automatique via Vault au demarrage
- Acces restreint aux administrateurs (Role "Administrator" ou Groupe "Admin_Designer")
- Deploiement multi-sites : Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs

**Chemin Vault :**
```
$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/
```

**Sections configurables :**
- Liste des initiales designers (26 entrees + "Autre...")
- Chemins templates et projets
- Extensions Inventor supportees
- Dossiers/fichiers exclus
- Noms des iProperties

### 4. Upload Template (100%)

- **Reserve aux administrateurs** - Message XnrgyMessageBox si non-admin
- **Upload templates** depuis Library vers Vault
- **Utilise la connexion partagee** de l'app principale
- **Journal integre** avec barre de progression

### 5. Checklist HVAC (100%)

- Validation des modules AHU
- Checklist interactive avec criteres XNRGY
- Stockage des validations dans Vault

### 6. Connexions Automatiques

- **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique
- **Inventor Professional 2026.2** - COM avec detection d'instance active
- **Throttling intelligent** pour eviter spam logs (v1.1)
- **Verification fenetre Inventor** prete avant connexion COM
- **Update Workspace** - Synchronisation dossiers au demarrage :
  - `$/Content Center Files`
  - `$/Engineering/Inventor_Standards`
  - `$/Engineering/Library/Cabinet`
  - `$/Engineering/Library/Xnrgy_M99`
  - `$/Engineering/Library/Xnrgy_Module`

---

## Architecture

```
XnrgyEngineeringAutomationTools/
+-- App.xaml(.cs)                    # Point d'entree application
+-- MainWindow.xaml(.cs)             # Dashboard principal
+-- appsettings.json                 # Configuration sauvegardee
|
+-- Models/                          # Modeles de donnees
|   +-- ApplicationConfiguration.cs  # Configuration application
|   +-- CategoryItem.cs              # Item categorie pour ComboBox
|   +-- FileItem.cs                  # Item fichier pour DataGrid
|   +-- FileToUpload.cs              # Fichier a uploader
|   +-- LifecycleDefinitionItem.cs   # Lifecycle Definition
|   +-- LifecycleStateItem.cs        # Lifecycle State
|   +-- ModuleInfo.cs                # Informations module
|   +-- ProjectInfo.cs               # Informations projet
|   +-- ProjectProperties.cs         # Proprietes Project/Ref/Module
|   +-- VaultConfiguration.cs        # Configuration Vault
|   +-- CreateModuleRequest.cs       # Requete creation module
|
+-- Services/                        # Services metier
|   +-- VaultSdkService.cs           # SDK Vault v31.0.84
|   +-- VaultSettingsService.cs      # Config chiffree + sync Vault
|   +-- InventorService.cs           # Connexion Inventor COM
|   +-- InventorCopyDesignService.cs # Copy Design natif
|   +-- Logger.cs                    # Logging UTF-8
|
+-- Views/                           # Fenetres et dialogues
|   +-- LoginWindow.xaml(.cs)        # Connexion Vault
|   +-- CreateModuleWindow.xaml(.cs) # Creer Module
|   +-- CreateModuleSettingsWindow.xaml(.cs) # Reglages Admin
|   +-- UploadTemplateWindow.xaml(.cs)       # Upload Template
|   +-- ChecklistHVACWindow.xaml(.cs)        # Checklist HVAC
|   +-- ModuleSelectionWindow.xaml(.cs)      # Selection module
|   +-- PreviewWindow.xaml(.cs)              # Previsualisation
|   +-- XnrgyMessageBox.xaml(.cs)            # MessageBox moderne
|
+-- Modules/                         # Modules integres
|   +-- VaultUpload/
|       +-- Models/
|       |   +-- VaultUploadFileItem.cs
|       |   +-- VaultUploadModels.cs
|       +-- Views/
|           +-- VaultUploadModuleWindow.xaml(.cs)
|
+-- ViewModels/                      # MVVM ViewModels
|   +-- AppMainViewModel.cs          # ViewModel principal
|   +-- RelayCommand.cs              # Implementation ICommand
|
+-- Converters/                      # Convertisseurs WPF
+-- Resources/                       # Images et icones
+-- Logs/                            # Fichiers logs
+-- build-and-run.ps1                # Script compilation MSBuild
```

---

## Proprietes XNRGY

Le systeme extrait automatiquement les proprietes depuis le chemin de fichier:

```
C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]
                              |         |       |
Vault Property IDs:        ID=112    ID=121  ID=122
```

| Propriete | ID Vault | Description |
|-----------|----------|-------------|
| Project | 112 | Numero de projet (5 chiffres) |
| Reference | 121 | Numero de reference (2 chiffres) |
| Module | 122 | Numero de module (2 chiffres) |

### Mapping Categorie -> Lifecycle Definition

| Categorie | Lifecycle Definition |
|-----------|---------------------|
| Engineering | Flexible Release Process |
| Office | Simple Release Process |
| Standard | Basic Release Process |
| Base | (aucun) |

---

## Prerequis

- **Windows 10/11 x64**
- **.NET Framework 4.8**
- **Autodesk Vault Professional 2026** (SDK v31.0.84)
- **Autodesk Inventor Professional 2026.2**
- **Visual Studio 2022** (pour compilation)
- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)

---

## Compilation et Execution

### Script automatique (RECOMMANDE)

```powershell
cd XnrgyEngineeringAutomationTools
.\build-and-run.ps1
```

**Fonctionnalites du script :**
- [+] Compilation automatique en mode Release
- [+] Detection automatique de MSBuild (VS 2022 Enterprise/Professional/Community)
- [+] Arret automatique de l'instance existante (taskkill /F)
- [+] Lancement automatique apres compilation reussie
- [+] Affichage des erreurs de compilation si presentes

### MSBuild manuel

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `
  XnrgyEngineeringAutomationTools.csproj /t:Rebuild /p:Configuration=Release /m /v:minimal /nologo
```

> **[!] IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne genere pas les fichiers .g.cs pour WPF .NET Framework 4.8.

---

## Exclusions de fichiers

**Extensions exclues:**
- `.v`, `.bak`, `.old` (Backup Vault)
- `.tmp`, `.temp` (Temporaires)
- `.ipj` (Projet Inventor)
- `.lck`, `.lock`, `.log` (Systeme/logs)
- `.dwl`, `.dwl2` (AutoCAD locks)

**Prefixes exclus:**
- `~$` (Office temporaire)
- `._` (macOS temporaire)
- `Backup_` (Backup generique)
- `.~` (Temporaire generique)

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

**Icones textuelles utilisees (pas d'emoji dans les logs):**
- `[+]` = Succes
- `[-]` = Erreur
- `[!]` = Avertissement
- `[>]` = Action en cours
- `[i]` = Information
- `[~]` = Attente/Polling
- `[#]` = Liste/Proprietes
- `[?]` = Verification

---

## Services Principaux

### VaultSdkService.cs

Service principal pour l'interaction avec Vault SDK.

**Responsabilites :**
- Connexion/deconnexion Vault
- Chargement des Property Definitions
- Chargement des Categories
- Chargement des Lifecycle Definitions
- Upload de fichiers avec `FileManager.AddFile`
- Application des proprietes via `UpdateFileProperties`
- Synchronisation Vault -> iProperties via `IExplorerUtil.UpdateFileProperties`
- Assignation de categories via `UpdateFileCategories`
- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (reflection)
- Assignation de revisions via `UpdateFileRevisionNumbers`
- Gestion des erreurs Vault (1003, 1013, 1136, etc.)

### InventorService.cs

Service pour la connexion COM a Inventor.

**Ameliorations v1.1 :**
- Throttling intelligent (minimum 2 sec entre tentatives)
- Verification fenetre Inventor prete (MainWindowHandle != IntPtr.Zero)
- Logs silencieux pour COMException 0x800401E3
- Compteur d'echecs consecutifs avec log periodique

### InventorCopyDesignService.cs

Service pour Copy Design natif avec gestion des references.

**Methode principale :**
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

---

## Depannage

### L'application ne demarre pas
- Verifier .NET Framework 4.8 installe
- Verifier Vault Professional 2026 installe
- Verifier les dependances NuGet restaurees

### Erreur de connexion Vault
- Verifier serveur accessible
- Verifier vault existe
- Verifier identifiants
- Voir logs dans `bin\Release\Logs\`

### Erreur connexion Inventor (0x800401E3)
- Inventor doit etre **completement demarre** (fenetre principale visible)
- L'app attend que Inventor s'enregistre dans la Running Object Table (ROT)
- Le timer de reconnexion reessaie automatiquement toutes les 3 secondes

### Proprietes non appliquees
- Verifier logs : rechercher "Application des proprietes"
- Si erreur 1003 : Fichier en traitement par Job Processor (normal)
- Si erreur 1013 : CheckOut necessaire (automatique)
- Verifier que les Property Definitions sont chargees (Project, Reference, Module)
- Pour fichiers Inventor : Verifier que `IExplorerUtil` est charge
- Pour writeback iProperties : Verifier que le writeback est active dans Vault

### Headers DataGrid invisibles
- Les styles DataGrid sont definis dans Window.Resources
- Fond sombre (#1A1A28) avec texte bleu XNRGY (#0078D4)
- Style applique globalement via `<Style TargetType="DataGridColumnHeader">`

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

### v1.1.0 (30 Decembre 2025)

**[+] Upload Module integre:**
- Module VaultAutomationTool integre dans l'app principale (`Modules/VaultUpload/`)
- Interface avec deux DataGrids (Inventor/Non-Inventor)
- Styles DataGrid avec headers visibles (fond sombre #1A1A28, texte bleu #0078D4)
- Barre de progression et journal des operations style Creer Module
- Utilise la connexion Vault partagee (pas de login separe)
- Controles Pause/Stop/Annuler

**[+] Upload Template:**
- Nouvelle fenetre pour upload templates (reserve Admin)
- Utilise connexion partagee de l'app principale
- XnrgyMessageBox si utilisateur non-admin

**[+] Corrections Inventor:**
- Throttling intelligent pour eviter spam logs
- Verification fenetre Inventor prete avant connexion COM
- Logs silencieux pour COMException 0x800401E3
- Timer de reconnexion optimise

**[+] VaultBulkUploader:**
- Outil console pour upload massif (6152 fichiers uploades vers PROD_XNGRY)
- Situe dans `Tools/VaultBulkUploader/`

### v1.0.0 (17 Decembre 2025)

**[+] Creer Module - Copy Design:**
- Copy Design natif avec 1133 fichiers
- Gestion des fichiers orphelins (1059 fichiers)
- Mise a jour references IDW
- Switch IPJ automatique
- Application iProperties et parametres Inventor
- Design View "Default" + Workfeatures caches
- Vue ISO + Zoom All + Save All
- Module reste ouvert pour le dessinateur

**[+] Vault Upload:**
- Upload complet avec proprietes automatiques
- Gestion Inventor et non-Inventor separee
- Categories, lifecycle et revisions
- Synchronisation Vault -> iProperties via IExplorerUtil

**[+] Reglages Admin:**
- Chiffrement AES-256
- Synchronisation automatique via Vault
- Interface graphique avec validation temps reel

**[+] Connexions automatiques:**
- Vault SDK v31.0.84
- Inventor COM 2026.2
- Update Workspace au demarrage

### v1.0.0 (31 Decembre 2025) - RELEASE OFFICIELLE

**[+] Systeme de Themes Ameliore:**
- Propagation theme light/dark vers TOUTES les sous-fenetres
- Event MainWindow.ThemeChanged declenche ApplyTheme() partout
- Elements a fond FIXE noir (#1A1A28): journal, panneaux stats
- Texte journal toujours blanc (#DCDCDC) quel que soit le theme
- UserPreferencesManager pour persistance locale du theme

**[+] Standardisation Titres Fenetres:**
- Format uniforme: [Nom] - v1.0.0 - Released on 2025-12-31 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC
- Applique a toutes les 9 fenetres de l'application

**[+] Couleurs Theme:**
| Element | Dark | Light |
|---------|------|-------|
| MainGrid | #1E1E2E | #F5F7FA |
| Stats/Log | #1A1A28 (FIXE) | #1A1A28 (FIXE) |
| Headers | #2A4A6F | #2A4A6F |

### v0.9.0 (15 Decembre 2025)

- Release initiale beta
- Dashboard principal avec boutons modules
- Connexion Vault centralisee
- Themes sombre/clair

---

## Auteur

**Mohammed Amine Elgalai**  
Engineering Automation Developer  
XNRGY Climate Systems ULC  
Email: mohammedamine.elgalai@xnrgy.com

---

## Licence

Proprietaire - XNRGY Climate Systems ULC (c) 2025

---

**Derniere mise a jour**: 31 Decembre 2025 - v1.0.0 Release