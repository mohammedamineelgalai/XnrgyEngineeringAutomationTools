# 🏭 XNRGY Engineering Automation Tools# 🏭 XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# VaultAutomationTool



> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

>

> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2



--->



## 📋 Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC🏭 **Suite d'outils d'automatisation engineering unifiée** pour piloter Autodesk Vault Professional 2026 et Inventor Professional 2026.2Application WPF pour l'upload automatisé de fichiers vers Autodesk Vault Professional 2026 avec application automatique des propriétés métier (Project, Reference, Module), catégories, lifecycle et révisions.



**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.



### 🎯 Objectif Principal---



Remplacer les multiples applications standalone par une **plateforme unique** avec :

- Connexion centralisée à Vault & Inventor

- Interface utilisateur moderne et cohérente## 📋 Description## 📋 Description## 📋 Description

- Partage de services communs (logging, configuration, etc.)

- Déploiement et maintenance simplifiés



---**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.



## 📦 Modules Intégrés



| Module | Description | Statut |### 🎯 Objectif PrincipalApplication hub centralisée qui regroupe tous les outils d'automatisation engineering XNRGY :Cette application permet de :

|--------|-------------|--------|

| 📤 **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ **100%** |

| 📦 **Créer Module** | Copy Design natif depuis template Library vers Projects | ✅ **95%** |

| ⚡ **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | 📋 Planifié |Remplacer les multiples applications standalone par une **plateforme unique** avec :- Scanner automatiquement les modules engineering (structure `Projects\[NUMERO]\REF[NUM]\M[NUM]`)

| 📐 **DXF Verifier** | Validation des fichiers DXF avant envoi | 📋 Migration |

| ✅ **Checklist HVAC** | Validation modules AHU avec stockage Vault | 📋 Migration |- Connexion centralisée à Vault & Inventor

| ⏱️ **Time Tracker** | Analyse temps de travail modules HVAC | 📋 Migration |

| 🔄 **Update Workspace** | Synchronisation librairies depuis Vault | 📋 Planifié |- Interface utilisateur moderne et cohérente- **Vault Upload** - Upload automatisé vers Vault avec propriétés (Project/Reference/Module)- Uploader des fichiers vers Vault avec création automatique de l'arborescence



---- Partage de services communs (logging, configuration, etc.)



## ✅ Fonctionnalités Implémentées- Déploiement et maintenance simplifiés- **Pack & Go** - GET depuis Vault, insertion dans assemblages, Copy Design- Appliquer automatiquement les propriétés métier (Project, Reference, Module)



### 1. Vault Upload (100% ✅)



Module complet pour l'upload automatisé vers Autodesk Vault Professional 2026.---- **Smart Tools** - Création IPT/STEP, génération PDF, iLogic Forms- Assigner des catégories, lifecycle definitions/states et révisions



**Caractéristiques :**

- ✅ Connexion directe via SDK Vault v31.0.84

- ✅ Scan automatique des modules engineering (`Projects\[NUM]\REF[XX]\M[XX]`)## 📦 Modules Intégrés- **DXF Verifier** - Validation des fichiers DXF avant envoi- Gérer l'upload de fichiers Inventor et non-Inventor séparément

- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

- ✅ Application automatique des propriétés métier extraites du chemin

- ✅ Assignation de catégories, lifecycle definitions/states et révisions

- ✅ Synchronisation Vault → iProperties via `IExplorerUtil`| Module | Description | Statut | Source |- **Checklist HVAC** - Validation modules AHU avec stockage Vault

- ✅ Gestion séparée Inventor / Non-Inventor

- ✅ Logs détaillés UTF-8 avec emojis|--------|-------------|--------|--------|



### 2. Créer Module - Copy Design (95% ✅)| 📤 **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ **Fonctionnel** | Natif |- **Update Workspace** - Synchronisation des librairies depuis Vault## 🎯 Caractéristiques



Module pour créer de nouveaux modules depuis le template Library avec Copy Design natif.| 📦 **Pack & Go** | GET depuis Vault + Copy Design natif | 🚧 **En cours** | Natif |



**Workflow complet :**| ⚡ **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | 📋 **Planifié** | Nouveau |

```

📁 Template: $/Engineering/Library/Xnrgy_Module| 📐 **DXF Verifier** | Validation DXF/CSV vs PDF Cut Lists | 📋 **Migration** | `DXFVerifier/` |

    ↓

📦 Copy Design Natif (1083 fichiers Inventor)| ✅ **Checklist HVAC** | Validation modules AHU avec stockage Vault | 📋 **Migration** | `ChecklistHVAC/` |## 🎯 Fonctionnalités- ✅ Connexion directe à Vault via SDK (VaultSDKService.cs)

    ↓

📂 Destination: C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]| ⏱️ **Time Tracker** | Analyse temps de travail modules HVAC | 📋 **Migration** | `HVACTimeTracker/` |

```

| 🔄 **Update Workspace** | Synchronisation librairies depuis Vault | 📋 **Planifié** | Nouveau |- ✅ Scan automatique des modules engineering (FileScanner.cs)

**Étapes automatisées :**

1. ✅ Switch vers projet template (IPJ)

2. ✅ Ouverture Top Assembly (Module_.iam)

3. ✅ Application iProperties sur le template---### Connexions Automatiques- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

4. ✅ Collecte de toutes les références (bottom-up)

5. ✅ Copy Design natif avec SaveAs (IPT → IAM → Top Assembly)

6. ✅ Traitement des dessins (.idw) avec mise à jour des références

7. ✅ Copie des fichiers orphelins (1059 fichiers non-référencés)## ✅ Fonctionnalités Implémentées- ✅ Connexion centralisée à **Vault Professional 2026** (SDK v31.0.84)- ✅ Application automatique des propriétés métier extraites du chemin

8. ✅ Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)

9. ✅ Renommage du fichier .ipj (XXXXX-XX-XX_2026.ipj → 123450101.ipj)

10. ✅ Switch vers le nouveau projet

11. ✅ Ouverture du nouveau Top Assembly### 1. Vault Upload (100%)- ✅ Connexion COM à **Inventor Professional 2026.2**- ✅ Assignation de catégories, lifecycle definitions/states et révisions

12. ✅ Application des iProperties finales

13. ✅ Application des paramètres Inventor

14. ✅ Design View → "Default"

15. ✅ Masquage des Workfeatures (plans, axes, points)Module complet pour l'upload automatisé vers Autodesk Vault Professional 2026.- ✅ Détection automatique d'Inventor s'il est en cours d'exécution- ✅ Gestion de la progression et pause/reprise

16. ✅ Vue ISO + Zoom All (Fit)

17. ✅ Update All (rebuild)

18. ✅ Save All

19. ✅ Module reste ouvert pour le dessinateur**Caractéristiques :**- ✅ Logs détaillés UTF-8 avec emoji (Logger.cs)



**Gestion intelligente des références :**- ✅ Connexion directe via SDK Vault v31.0.84

- 🔗 Fichiers Library (IPT_Typical_Drawing) : Liens préservés

- 📁 Fichiers Module : Copiés avec références mises à jour- ✅ Scan automatique des modules engineering (`Projects\[NUM]\REF[XX]\M[XX]`)### Update Workspace (GET automatique)- ✅ Exclusion automatique des fichiers temporaires (.bak, .dwl, .log, OldVersions, ~$)

- 📄 Fichiers IDW : Références corrigées via `PutLogicalFileNameUsingFull`

- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

### 3. Connexions Automatiques

- ✅ Application automatique des propriétés métier extraites du cheminAu démarrage ou sur demande, synchronisation des dossiers essentiels :- ✅ Sauvegarde configuration (appsettings.json)

- ✅ **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique

- ✅ **Inventor Professional 2026.2** - COM avec détection d'instance active- ✅ Assignation de catégories, lifecycle definitions/states et révisions

- ✅ **Update Workspace** - Synchronisation dossiers au démarrage :

  - `$/Content Center Files`- ✅ Synchronisation Vault → iProperties via `IExplorerUtil`- `$/Content Center Files` → `C:\Vault\Content Center Files`- ✅ Interface MVVM avec séparation Inventor/Non-Inventor

  - `$/Engineering/Inventor_Standards`

  - `$/Engineering/Library/Cabinet`- ✅ Gestion séparée Inventor / Non-Inventor

  - `$/Engineering/Library/Xnrgy_M99`

  - `$/Engineering/Library/Xnrgy_Module`- ✅ Logs détaillés UTF-8 avec emojis- `$/Engineering/Inventor_Standards` → `C:\Vault\Engineering\Inventor_Standards`



---



## 📦 Prérequis### 2. Pack & Go (70%)- `$/Engineering/Library/Cabinet` → `C:\Vault\Engineering\Library\Cabinet`## 📦 Prérequis



- **Windows 10/11 x64**

- **.NET Framework 4.8**

- **Autodesk Vault Professional 2026** (SDK v31.0.84)Module pour extraire depuis Vault et créer des copies de modules avec références mises à jour.- `$/Engineering/Library/Xnrgy_M99` → `C:\Vault\Engineering\Library\Xnrgy_M99`

- **Autodesk Inventor Professional 2026.2**

- **Visual Studio 2022** (pour compilation)

- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)

**Implémenté :**- `$/Engineering/Library/Xnrgy_Module` → `C:\Vault\Engineering\Library\Xnrgy_Module`- Windows 10/11 x64

---

- ✅ GET automatique depuis Vault avec dépendances

## 🏗️ Architecture Technique

- ✅ Extraction vers dossier temporaire- .NET Framework 4.8

### Stack Technologique

- ✅ Interface de sélection de destination

```

┌─────────────────────────────────────────────────────────┐- 🚧 Copy Design natif (bottom-up SaveAs avec références)## 📦 Modules Intégrés- Autodesk Vault Professional 2026

│                    Présentation (WPF)                   │

│  MainWindow.xaml │ Views/*.xaml │ MVVM Pattern          │

├─────────────────────────────────────────────────────────┤

│                   ViewModels (MVVM)                     │**En cours :**- Visual Studio 2022 ou supérieur (pour compilation)

│  AppMainViewModel.cs │ RelayCommand │ INotifyProperty   │

├─────────────────────────────────────────────────────────┤- 🔄 Correction des références croisées entre assemblages siblings

│                    Services Layer                       │

│  VaultSDKService │ InventorService │ Logger             │- 🔄 Gestion OldVersions et fichiers obsolètes| Module | Description | Status |- MSBuild 18.0.0+ (REQUIS - dotnet build ne fonctionne PAS pour WPF)

│  InventorCopyDesignService │ ModuleCopyService          │

├─────────────────────────────────────────────────────────┤

│                    Models (Data)                        │

│  FileItem │ ModuleInfo │ ProjectProperties │ Config     │### 3. Connexions Automatiques|--------|-------------|--------|

├─────────────────────────────────────────────────────────┤

│                   External APIs                         │

│  Vault SDK v31.0.84 │ Inventor COM 2026.2               │

└─────────────────────────────────────────────────────────┘- ✅ **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique| 📤 Vault Upload | Upload avec propriétés automatiques | ✅ Intégré |## 🏗️ Structure du projet

```

- ✅ **Inventor Professional 2026.2** - COM avec détection d'instance active

### Structure du Projet

- ✅ **Update Workspace** - Synchronisation dossiers au démarrage :| 📦 Pack & Go | GET Vault + Copy Design | 🚧 En développement |

```

XnrgyEngineeringAutomationTools/  - `$/Content Center Files`

├── App.xaml(.cs)                    # Point d'entrée application

├── MainWindow.xaml(.cs)             # Dashboard principal avec boutons modules  - `$/Engineering/Inventor_Standards`| ⚡ Smart Tools | IPT, STEP, PDF, iLogic | 🚧 En développement |```

├── appsettings.json                 # Configuration sauvegardée

│  - `$/Engineering/Library/Cabinet`

├── Models/                          # Modèles de données

│   ├── ApplicationConfiguration.cs  # Configuration application  - `$/Engineering/Library/Xnrgy_M99`| 📐 DXF Verifier | Validation fichiers DXF | 🚧 Migration |VaultAutomationTool/

│   ├── FileItem.cs                  # Item fichier pour DataGrid

│   ├── ModuleInfo.cs                # Informations module  - `$/Engineering/Library/Xnrgy_Module`

│   ├── ProjectProperties.cs         # Propriétés Project/Ref/Module

│   └── VaultConfiguration.cs        # Configuration Vault| ✅ Checklist HVAC | Validation AHU + Vault | 🚧 Migration |├── Models/                          # Modèles de données (10 fichiers)

│

├── Services/                        # Services métier---

│   ├── VaultSDKService.cs           # Service principal Vault SDK

│   ├── InventorService.cs           # Service Inventor COM│   ├── ApplicationConfiguration.cs  # Configuration application

│   ├── InventorCopyDesignService.cs # Copy Design natif

│   ├── ModuleCopyService.cs         # Service copie module## 📋 Fonctionnalités Planifiées

│   └── Logger.cs                    # Système logging UTF-8

│## 📦 Prérequis│   ├── CategoryItem.cs             # Item catégorie pour ComboBox

├── ViewModels/                      # ViewModels MVVM

│   ├── AppMainViewModel.cs          # ViewModel principal### Smart Tools (À développer)

│   └── RelayCommand.cs              # Implémentation ICommand

││   ├── FileItem.cs                 # Item fichier pour DataGrid

├── Views/                           # Fenêtres et contrôles

│   ├── CreateModuleWindow.xaml      # Fenêtre création module| Outil | Description | Priorité |

│   └── VaultUploadWindow.xaml       # Fenêtre upload Vault

│|-------|-------------|----------|- **Windows 10/11 x64**│   ├── FileToUpload.cs             # Fichier à uploader

└── bin/Release/                     # Exécutable compilé

    ├── XnrgyEngineeringAutomationTools.exe| **IPT Creator** | Création rapide de pièces avec templates XNRGY | Haute |

    └── Logs/                        # Logs d'exécution

        └── VaultSDK_POC_YYYYMMDD_HHMMSS.log| **STEP Exporter** | Export batch STEP avec options | Moyenne |- **.NET Framework 4.8**│   ├── LifecycleDefinitionItem.cs  # Lifecycle Definition

```

| **PDF Generator** | Génération PDF depuis IDW avec watermarks | Haute |

---

| **iLogic Forms** | Formulaires personnalisés pour iLogic | Moyenne |- **Autodesk Vault Professional 2026** (SDK v31.0.84)│   ├── LifecycleStateItem.cs       # Lifecycle State

## 🚀 Compilation et Lancement

| **BOM Exporter** | Export nomenclatures vers Excel | Haute |

### Script automatique (Recommandé)

- **Autodesk Inventor Professional 2026.2**│   ├── ModuleInfo.cs               # Informations module

```powershell

# Compiler et lancer (mode Release)### DXF Verifier Migration (À migrer)

.\build-and-run.ps1

- **Visual Studio 2022** (pour compilation)│   ├── ProjectInfo.cs              # Informations projet

# Mode Debug

.\build-and-run.ps1 -Debug> Source : `DXFVerifier/` - VB.NET → C# WPF



# Clean + Build│   ├── ProjectProperties.cs        # Propriétés Project/Ref/Module

.\build-and-run.ps1 -Clean

**Fonctionnalités à migrer :**

# Build seulement (sans lancer)

.\build-and-run.ps1 -BuildOnly- Double stratégie extraction PDF (tableaux + ballons)## 🏗️ Structure du Projet│   └── VaultConfiguration.cs       # Configuration Vault



# Kill les instances en cours- Comparaison DXF/CSV vs Cut Lists PDF

.\build-and-run.ps1 -KillOnly

```- Génération rapports Excel avec templates XNRGY├── Services/                        # Services métier (2 fichiers)



### Compilation manuelle- ~97% précision extraction



```powershell```│   ├── VaultSDKService.cs         # Service principal Vault SDK

# ⚠️ IMPORTANT: Utiliser MSBuild, pas dotnet build

$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe"### HVAC Time Tracker Migration (À migrer)

& $msbuild XnrgyEngineeringAutomationTools.csproj /t:Rebuild /p:Configuration=Release /m /v:minimal

```XnrgyEngineeringAutomationTools/│   └── Logger.cs                   # Système logging UTF-8



---> Source : `HVACTimeTracker/` - VB.NET → C# WPF



## 📋 Fonctionnalités Planifiées├── MainWindow.xaml              # Dashboard principal├── ViewModels/                      # ViewModels MVVM (1 fichier)



### Smart Tools (À développer)**Fonctionnalités à migrer :**



| Outil | Description | Priorité |- Analyse hybride API Inventor + estimation calibrée├── App.xaml                     # Configuration WPF│   ├── AppMainViewModel.cs         # ViewModel principal

|-------|-------------|----------|

| **IPT Creator** | Création rapide de pièces avec templates XNRGY | Haute |- Catégorisation automatique (3D/2D Equipment/Cabinet)

| **STEP Exporter** | Export batch STEP avec options | Moyenne |

| **PDF Generator** | Génération PDF depuis IDW avec watermarks | Haute |- Statistiques temps réel (9 cartes Σ)├── Assets/│   └── RelayCommand.cs             # Implémentation ICommand

| **iLogic Forms** | Formulaires personnalisés pour iLogic | Moyenne |

| **BOM Exporter** | Export nomenclatures vers Excel | Haute |- Export Excel professionnel



### DXF Verifier Migration (À migrer)│   └── Icons/                   # Icônes des modules├── Properties/



> Source : `DXFVerifier/` - VB.NET → C# WPF### Checklist HVAC Migration (À migrer)



- Double stratégie extraction PDF (tableaux + ballons)├── Modules/│   └── AssemblyInfo.cs             # Informations assembly

- Comparaison DXF/CSV vs Cut Lists PDF

- Génération rapports Excel avec templates XNRGY> Source : `ChecklistHVAC/` - HTML/JS → WPF avec stockage Vault

- ~97% précision extraction

│   ├── VaultUpload/            # Module upload Vault├── App.xaml(.cs)                   # Point d'entrée application

### HVAC Time Tracker Migration (À migrer)

**Fonctionnalités à migrer :**

> Source : `HVACTimeTracker/` - VB.NET → C# WPF

- Checklist validation modules AHU│   ├── PackAndGo/              # Module Pack & Go├── MainWindow.xaml(.cs)            # Fenêtre principale

- Analyse hybride API Inventor + estimation calibrée

- Catégorisation automatique (3D/2D Equipment/Cabinet)- Stockage état dans Vault

- Statistiques temps réel (9 cartes Σ)

- Export Excel professionnel- Génération PDF rapport│   ├── SmartTools/             # Module Smart Tools├── appsettings.json                # Configuration sauvegardée



### Checklist HVAC Migration (À migrer)- Historique par module



> Source : `ChecklistHVAC/` - HTML/JS → WPF avec stockage Vault│   ├── DXFVerifier/            # Module DXF Verifier├── README.md                        # Ce fichier



- Checklist validation modules AHU### Update Workspace (À développer)

- Stockage état dans Vault

- Génération PDF rapport│   └── ChecklistHVAC/          # Module Checklist HVAC└── bin/Release/                     # Exécutable compilé

- Historique par module

| Fonctionnalité | Description |

---

|----------------|-------------|├── Services/    ├── VaultAutomationTool.exe     # Application

## 📊 Logs et Debugging

| **Sync Sélectif** | Choisir quels dossiers synchroniser |

### Emplacement des logs

| **Sync Programmé** | Planification automatique |│   ├── VaultSdkService.cs      # Service Vault SDK    └── Logs/                       # Logs d'exécution UTF-8

```

bin\Release\Logs\VaultSDK_POC_YYYYMMDD_HHMMSS.log| **Diff Visuel** | Voir les différences avant sync |

```

| **Rollback** | Restaurer version précédente |│   ├── InventorService.cs      # Service Inventor COM        └── VaultSDK_POC_YYYYMMDD_HHMMSS.log

### Format des logs



```

[2025-12-26 21:42:24.123] [INFO   ] ✅ Module prêt pour le dessinateur: 123450101.iam---│   └── Logger.cs               # Système de logs```

[2025-12-26 21:42:24.456] [DEBUG  ] 📐 Traitement de 9 fichiers de dessins...

[2025-12-26 21:42:24.789] [SUCCESS] ✅ COPY DESIGN TERMINÉ: 1133 fichiers copiés

```

## 🏗️ Architecture Technique├── Views/

### Niveaux de log



- `INFO` - Informations générales

- `DEBUG` - Détails techniques### Stack Technologique│   ├── LoginWindow.xaml        # Fenêtre connexion## 🔧 Architecture

- `SUCCESS` - Opérations réussies ✅

- `WARN` - Avertissements ⚠️

- `ERROR` - Erreurs ❌

```│   └── VaultUploadWindow.xaml  # Fenêtre upload Vault

---

┌─────────────────────────────────────────────────────────┐

## 📁 Chemins Importants

│                    Présentation (WPF)                   │└── ViewModels/                  # MVVM ViewModels### Pattern MVVM (Model-View-ViewModel)

| Chemin | Description |

|--------|-------------|│  MainWindow.xaml │ Views/*.xaml │ MVVM Pattern          │

| `C:\Vault\Engineering\Library\Xnrgy_Module` | Template source pour Copy Design |

| `C:\Vault\Engineering\Library\Cabinet\IPT_Typical_Drawing` | Fichiers partagés (liens préservés) |├─────────────────────────────────────────────────────────┤```

| `C:\Vault\Engineering\Projects\[NUM]\REF[XX]\M[XX]` | Destination des modules créés |

| `$/Engineering/Projects/` | Racine Vault des projets |│                   ViewModels (MVVM)                     │



---│  AppMainViewModel.cs │ RelayCommand │ INotifyProperty   │- **Models** : Données et configuration



## 🔄 Changelog├─────────────────────────────────────────────────────────┤



### v1.0 (2025-12-26)│                    Services Layer                       │## 🚀 Compilation et Lancement- **Views** : Interface utilisateur XAML (MainWindow.xaml)



**Créer Module - Copy Design :**│  VaultSDKService │ InventorService │ Logger             │

- ✅ Copy Design natif avec 1133 fichiers

- ✅ Gestion des fichiers orphelins (1059 fichiers)│  InventorCopyDesignService │ ModuleCopyService          │- **ViewModels** : Logique métier et binding (AppMainViewModel.cs)

- ✅ Mise à jour références IDW

- ✅ Switch IPJ automatique├─────────────────────────────────────────────────────────┤

- ✅ Application iProperties et paramètres Inventor

- ✅ Design View "Default" + Workfeatures cachés│                    Models (Data)                        │### Script automatique- **Services** : Accès aux données Vault (VaultSDKService.cs)

- ✅ Vue ISO + Zoom All + Save All

- ✅ Module reste ouvert pour le dessinateur│  FileItem │ ModuleInfo │ ProjectProperties │ Config     │



**Vault Upload :**├─────────────────────────────────────────────────────────┤```powershell

- ✅ Upload complet avec propriétés automatiques

- ✅ Gestion Inventor et non-Inventor séparée│                  External APIs                          │

- ✅ Catégories, lifecycle et révisions

│  Vault SDK 2026 (v31.0.84) │ Inventor COM 2026.2        │cd XnrgyEngineeringAutomationTools### Services principaux

---

└─────────────────────────────────────────────────────────┘

## 👨‍💻 Auteur

```.\build-and-run.ps1

**Mohammed Amine Elgalai**  

Engineering Automation Developer  

XNRGY Climate Systems ULC

### Structure des Fichiers```#### 1. VaultSDKService.cs

---



## 📄 Licence

```

Propriétaire - XNRGY Climate Systems ULC © 2025

XnrgyEngineeringAutomationTools/

├── 📁 Assets/                      # Ressources graphiques### MSBuild manuelService principal pour l'interaction avec Vault SDK.

│   └── Icons/                      # Icônes des modules

├── 📁 Converters/                  # Convertisseurs XAML```powershell

├── 📁 Models/                      # Modèles de données (11 fichiers)

│   ├── ApplicationConfiguration.cs& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `**Responsabilités** :

│   ├── CategoryItem.cs

│   ├── CreateModuleRequest.cs  XnrgyEngineeringAutomationTools.csproj /p:Configuration=Release /t:Rebuild- Connexion/déconnexion Vault

│   ├── FileItem.cs

│   ├── FileToUpload.cs```- Chargement des Property Definitions

│   ├── LifecycleDefinitionItem.cs

│   ├── LifecycleStateItem.cs- Chargement des Catégories

│   ├── ModuleInfo.cs

│   ├── ProjectInfo.cs## 🔌 APIs Utilisées- Chargement des Lifecycle Definitions

│   ├── ProjectProperties.cs

│   └── VaultConfiguration.cs- Upload de fichiers avec `FileManager.AddFile`

├── 📁 Modules/                     # Modules (dossiers réservés)

│   ├── ChecklistHVAC/              # 📋 À migrer### Vault SDK 2026- Application des propriétés via `UpdateFileProperties`

│   ├── DXFVerifier/                # 📋 À migrer

│   ├── PackAndGo/                  # 🚧 En cours- `VDF.Vault.Library.ConnectionManager` - Connexion- Synchronisation des propriétés Vault → iProperties via `IExplorerUtil.UpdateFileProperties` (pour fichiers Inventor)

│   ├── SmartTools/                 # 📋 Planifié

│   └── VaultUpload/                # ✅ Intégré dans Views- `VDF.Vault.Currency.Connections.Connection` - Session- Assignation de catégories via `UpdateFileCategories`

├── 📁 Services/                    # Services métier (11 fichiers)

│   ├── ApprenticePropertyService.cs- `FileManager.AddFile()` - Upload- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (via reflection)

│   ├── InventorCopyDesignService.cs

│   ├── InventorPropertyService.cs- `FileManager.AcquireFiles()` - Download (GET)- Assignation de révisions via `UpdateFileRevisionNumbers`

│   ├── InventorService.cs

│   ├── Logger.cs- `DocumentService.UpdateFileProperties()` - Propriétés- Gestion des erreurs Vault (1003, 1013, 1136, etc.)

│   ├── ModuleCopyService.cs

│   ├── NativeOlePropertyService.cs

│   ├── OlePropertyService.cs

│   ├── SimpleLogger.cs### Inventor 2026.2 COM**Méthodes principales** :

│   ├── VaultSDKService.cs

│   └── WindowsPropertyService.cs- `Inventor.Application` via `GetActiveObject()````csharp

├── 📁 ViewModels/                  # ViewModels MVVM

│   └── AppMainViewModel.cs- `Application.ActiveDocument` - Document actifbool Connect(string server, string vaultName, string username, string password)

├── 📁 Views/                       # Fenêtres et dialogues (6 fenêtres)

│   ├── ChecklistHVACWindow.xaml- iProperties via `Document.PropertySets`void Disconnect()

│   ├── CreateModuleWindow.xaml

│   ├── LoginWindow.xamlList<(long Id, string Name)> GetAvailableCategories()

│   ├── ModuleSelectionWindow.xaml

│   ├── PreviewWindow.xaml## 👤 AuteurList<LifecycleDefinitionItem> GetAvailableLifecycleDefinitions()

│   └── VaultUploadWindow.xaml

├── 📄 App.xaml / App.xaml.cs       # Point d'entrée WPFlong? GetLifecycleDefinitionIdByCategory(string categoryName)

├── 📄 MainWindow.xaml              # Dashboard principal

├── 📄 appsettings.json             # Configuration persistante**Mohammed Amine Elgalai**  long? GetWorkInProgressStateId(long lifecycleDefinitionId)

├── 📄 build-and-run.ps1            # Script compilation + lancement

└── 📄 README.md                    # Ce fichierSmart Tools Amine - XNRGY Climate Systems ULC  bool UploadFile(string filePath, string vaultFolderPath, 

```

Email: mohammedamine.elgalai@xnrgy.com    string? projectNumber = null, string? reference = null, string? module = null,

### Services Principaux

    long? categoryId = null, string? categoryName = null,

#### VaultSDKService.cs

Service principal pour l'interaction avec Vault SDK.## 📄 Version    long? lifecycleDefinitionId = null, long? lifecycleStateId = null, string? revision = null)



```csharp```

// Connexion

bool Connect(string server, string vaultName, string username, string password)**v1.0.0** - Décembre 2025

void Disconnect()

**Gestion des propriétés** :

// Chargement données

List<PropertyDefinition> GetPropertyDefinitions()### Historique- Propriétés XNRGY : Project (ID: 112), Reference (ID: 121), Module (ID: 122)

List<Category> GetAvailableCategories()

List<LifecycleDefinition> GetLifecycleDefinitions()- **v1.0.0** (17 Décembre 2025) : Version initiale- Chargement automatique des Property Definitions au démarrage



// Upload  - Dashboard principal avec modules- Application via `UpdateFileProperties` (nécessite CheckOut pour fichiers existants)

FileUploadResult AddFile(string localPath, string vaultPath, ...)

void UpdateFileProperties(long fileId, Dictionary<string, object> properties)  - Connexion Vault & Inventor centralisée- **Synchronisation Vault → iProperties** : Utilisation de `IExplorerUtil.UpdateFileProperties` pour les fichiers Inventor

void UpdateFileCategories(long fileId, long categoryId)

```  - Update Workspace automatique  - Chargement lazy d'ExplorerUtil si nécessaire



#### InventorService.cs  - Module Vault Upload intégré  - Writeback automatique des propriétés Vault vers les iProperties Inventor

Service pour l'interaction avec Inventor COM API.

  - Nécessite que le writeback soit activé dans Vault (`GetEnableItemPropertyWritebackToFiles`)

```csharp

// Connexion## 📜 Licence

bool Connect()                    // Connexion à instance existante

bool StartInventor()              // Démarrer nouvelle instance**Gestion du lifecycle** :

void Disconnect()

Propriétaire - XNRGY Climate Systems ULC- Utilisation de `DocumentServiceExtensions.UpdateFileLifeCycleDefinitions` via reflection

// Documents

Document OpenDocument(string path)- Support de différentes signatures de SDK

void SaveDocument(Document doc)

AssemblyDocument GetAssemblyDocument(string path)---- Assignation directe sans CheckOut pour nouveaux fichiers

```

**Dernière mise à jour** : 17 Décembre 2025

#### InventorCopyDesignService.cs

Service pour Copy Design natif avec gestion des références.#### 2. Logger.cs



```csharpSystème de logging UTF-8 avec emoji.

// Copy Design

Task<bool> ExecuteRealPackAndGoAsync(**Niveaux de log** :

    string sourceRoot,           // Dossier source (module template)- **TRACE** : Détails techniques très fins

    string destinationRoot,      // Dossier destination- **DEBUG** : Informations de débogage détaillées

    string topAssemblyPath,      // Assemblage principal- **INFO** : Opérations importantes (connexion, upload, succès)

    IProgress<string> progress- **WARNING** : Avertissements non bloquants

)- **ERROR** : Erreurs bloquantes

```- **CRITICAL** : Erreurs critiques système



---**Format des logs** :

```

## 📦 Prérequis[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] Message avec emoji

```

### Logiciels Requis

**Emoji utilisés** :

| Logiciel | Version | Notes |- 🔌 = Connexion

|----------|---------|-------|- ✅ = Succès

| Windows | 10/11 x64 | |- ❌ = Erreur

| .NET Framework | 4.8 | Inclus dans Windows 10+ |- ⚠️ = Avertissement

| Visual Studio | 2022 | Pour compilation |- 📋 = Liste/Propriétés

| MSBuild | 18.0.0+ | **REQUIS** - `dotnet build` ne fonctionne PAS |- 📊 = Statistiques

| Autodesk Vault Professional | 2026 | SDK v31.0.84 |- ⏳ = Attente/Polling

| Autodesk Inventor Professional | 2026.2 | COM Interop |- 🔍 = Vérification

- 📄 = Fichier

### Configuration Vault- 🔓 = CheckOut

- 💾 = Mise à jour

```- 🔒 = CheckIn

Serveur: vaultpro2026.yourcompany.com- 📤 = Upload

Vault: XNRGY_Engineering- 🔄 = Traitement

Utilisateur: [Active Directory]- 💡 = Info

```

### ViewModels

---

#### AppMainViewModel.cs

## 🚀 Compilation et Lancement

ViewModel principal avec toutes les propriétés et commandes.

### Script Automatique (Recommandé)

**Propriétés principales** :

```powershell- `IsConnected` : État de connexion Vault

cd XnrgyEngineeringAutomationTools- `IsProcessing` : État de traitement (scan/upload)

- `StatusMessage` : Message de statut

# Build Release + Run- `ProgressValue` : Valeur de progression (0-100)

.\build-and-run.ps1- `InventorFiles` : Collection fichiers Inventor

- `NonInventorFiles` : Collection fichiers non-Inventor

# Options disponibles- `AvailableCategories` : Catégories disponibles

.\build-and-run.ps1 -Debug       # Build Debug + Run- `SelectedCategoryInventor` / `SelectedCategoryNonInventor` : Catégories sélectionnées

.\build-and-run.ps1 -Clean       # Clean + Build Release + Run- `AvailableLifecycleDefinitions` : Lifecycle Definitions disponibles

.\build-and-run.ps1 -BuildOnly   # Build sans lancer- `SelectedLifecycleDefinitionInventor` / `SelectedLifecycleDefinitionNonInventor` : Lifecycle Definitions sélectionnées

.\build-and-run.ps1 -KillOnly    # Tuer les instances existantes- `AvailableStatesInventor` / `AvailableStatesNonInventor` : États disponibles

```- `SelectedLifecycleStateInventor` / `SelectedLifecycleStateNonInventor` : États sélectionnés

- `RevisionInventor` / `RevisionNonInventor` : Révisions saisies

### MSBuild Manuel

**Commandes** :

```powershell- `ToggleConnectionCommand` : Connexion/déconnexion Vault

# Release- `ScanProjectCommand` : Scan d'un module

& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `- `AutoCheckInCommand` : Upload des fichiers sélectionnés

  XnrgyEngineeringAutomationTools.csproj /p:Configuration=Release /t:Rebuild /v:minimal /nologo- `PauseCommand` : Pause/reprise du traitement



# Debug**Méthodes principales** :

& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' ````csharp

  XnrgyEngineeringAutomationTools.csproj /p:Configuration=Debug /t:Rebuild /v:minimal /nologovoid ToggleConnection()

```void ScanProject(string projectPath)

async Task AutoCheckInAsync()

### ⚠️ Importantvoid UpdateAvailableStates() // Met à jour les états selon la Lifecycle Definition sélectionnée

```

**NE PAS utiliser `dotnet build`** - Ce projet est WPF .NET Framework 4.8, pas .NET Core/5+.

### Models

---

#### FileItem.cs

## 📁 Projets Sources (À Intégrer)Représente un fichier à uploader avec :

- `IsChecked` : Sélectionné pour upload

Ces projets existants dans le repo doivent être migrés vers cette application hub :- `FullPath` : Chemin complet

- `FileName` : Nom du fichier

### DXFVerifier/- `Extension` : Extension

- **Langage** : VB.NET / .NET 9- `Category` : Catégorie (Inventor/Non-Inventor)

- **Type** : Windows Forms

- **Fonction** : Validation DXF/CSV vs PDF Cut Lists#### ProjectProperties.cs

- **Statut migration** : 📋 PlanifiéPropriétés extraites du chemin :

- **Priorité** : Haute (usage quotidien)- `Project` : Numéro de projet

- `Reference` : Numéro de référence

### HVACTimeTracker/- `Module` : Numéro de module

- **Langage** : VB.NET / .NET 9

- **Type** : Windows Forms#### CategoryItem.cs

- **Fonction** : Analyse temps de travail modules HVACCatégorie Vault avec :

- **Statut migration** : 📋 Planifié- `Id` : ID de la catégorie

- **Priorité** : Moyenne- `Name` : Nom de la catégorie



### ChecklistHVAC/#### LifecycleDefinitionItem.cs

- **Langage** : HTML/JavaScriptLifecycle Definition avec :

- **Type** : Application Web locale- `Id` : ID de la définition

- **Fonction** : Checklist validation modules AHU- `Name` : Nom de la définition

- **Statut migration** : 📋 Planifié- `States` : Collection des états disponibles

- **Priorité** : Moyenne

#### LifecycleStateItem.cs

### InventorVaultIntegration/Lifecycle State avec :

- **Langage** : C# / .NET 8 WPF- `Id` : ID de l'état

- **Type** : WPF MVVM- `Name` : Nom de l'état

- **Fonction** : Upload Vault avec batch scripts- `IsDefault` : État par défaut

- **Statut** : ✅ Code source de référence pour ce projet

## 🔌 API Vault SDK utilisées

---

### Connexion

## 🔧 Configuration```csharp

VDF.Vault.Library.ConnectionManager.LogIn(

### appsettings.json    server, vaultName, username, password,

    VDF.Vault.Currency.Connections.AuthenticationFlags.Standard, null

```json)

{```

  "VaultServer": "vaultpro2026.yourcompany.com",

  "VaultName": "XNRGY_Engineering",### Upload de fichiers

  "LastUsername": "",```csharp

  "RememberCredentials": false,_connection.FileManager.AddFile(

  "AutoConnectInventor": true,    targetFolder, fileName, null, lastWriteTime, null, null,

  "UpdateWorkspaceOnStartup": true,    fileClassification, false, fileStream

  "LogLevel": "Info",)

  "Paths": {```

    "ProjectsRoot": "C:\\Vault\\Engineering\\Projects",

    "LibraryRoot": "C:\\Vault\\Engineering\\Library",### Application des propriétés

    "TemplatesRoot": "C:\\Vault\\Engineering\\Library\\Xnrgy_Module"```csharp

  }// Pour nouveaux fichiers (sans CheckOut)

}_connection.WebServiceManager.DocumentService.UpdateFileProperties(

```    new[] { file.Id }, new[] { propArray }

)

---

// Pour fichiers existants (nécessite CheckOut)

## 📊 Logs_connection.WebServiceManager.DocumentService.CheckoutFile(...)

_connection.WebServiceManager.DocumentService.UpdateFileProperties(...)

Les logs sont générés dans `bin\Release\Logs\` avec le format :_connection.FileManager.CheckinFile(...)

``````

VaultCheckIn_YYYYMMDD_HHMMSS.log

```### Assignation de catégories

```csharp

### Niveaux de Log// Via DocumentServiceExtensions (via reflection)

var documentServiceExtensions = _connection.WebServiceManager.DocumentServiceExtensions;

| Niveau | Emoji | Usage |documentServiceExtensions.UpdateFileCategories(

|--------|-------|-------|    new[] { file.Id }, new[] { categoryId }

| TRACE | 🔍 | Détails techniques |)

| DEBUG | 🐛 | Informations debug |```

| INFO | ℹ️ | Opérations normales |

| WARN | ⚠️ | Avertissements |### Assignation de lifecycle

| ERROR | ❌ | Erreurs récupérables |```csharp

| FATAL | 💀 | Erreurs critiques |// Via DocumentServiceExtensions (via reflection)

var documentServiceExtensions = _connection.WebServiceManager.DocumentServiceExtensions;

---documentServiceExtensions.UpdateFileLifeCycleDefinitions(

    new[] { file.Id },

## 🛣️ Roadmap    new[] { lifecycleDefinitionId },

    new[] { lifecycleStateId },

### Phase 1 - Consolidation (En cours)    "Commentaire"

- [x] Vault Upload fonctionnel)

- [x] Connexions Vault & Inventor```

- [ ] Pack & Go - Copy Design stable

- [ ] Tests complets sur modules réels### Gestion des erreurs Vault



### Phase 2 - Smart Tools (Q1 2026)**Erreur 1003** : Fichier en traitement par Job Processor

- [ ] IPT Creator avec templates- **Solution** : Retour immédiat sans attente (pas de polling)

- [ ] PDF Generator batch

- [ ] BOM Exporter vers Excel**Erreur 1013** : Fichier doit être checké out pour modification

- [ ] STEP Exporter avec options- **Solution** : CheckOut → Update → CheckIn



### Phase 3 - Migrations (Q2 2026)**Erreur 1008** : Fichier existe déjà

- [ ] DXF Verifier (VB.NET → C# WPF)- **Solution** : Récupérer le fichier existant et appliquer les modifications

- [ ] HVAC Time Tracker (VB.NET → C# WPF)

- [ ] Checklist HVAC (HTML → WPF + Vault)**Erreur 1136** : Restriction lifecycle

- **Solution** : Vérifier les permissions et l'état du fichier

### Phase 4 - Avancé (Q3 2026)

- [ ] Update Workspace avec diff visuel## 📝 Flux d'upload

- [ ] Notifications temps réel

- [ ] Dashboard statistiques### 1. Scan du module

- [ ] Plugin Inventor (bouton dans ruban)- Chemin attendu : `...\Engineering\Projects\[NUMERO]\REF[NUM]\M[NUM]`

- Extraction automatique : Project, Reference, Module

---- Scan récursif avec exclusions (fichiers temporaires, dossiers système)



## 📝 Changelog### 2. Sélection des fichiers

- Séparation Inventor / Non-Inventor

### v1.0.0 (2025-12-26)- Sélection par défaut de tous les fichiers

- 🎉 Version initiale- Filtres de recherche disponibles

- ✅ Module Vault Upload complet

- ✅ Connexions Vault & Inventor### 3. Configuration

- 🚧 Module Pack & Go en développement- Sélection de la catégorie (Base par défaut)

- 📁 Structure modulaire préparée- Sélection de la Lifecycle Definition (selon catégorie)

- Sélection de l'état (selon Lifecycle Definition)

---- Saisie de la révision (manuel pour l'instant)



## 👤 Auteur### 4. Upload

- Création de l'arborescence Vault si nécessaire

**Mohammed Amine Elgalai**  - Upload du fichier avec `FileManager.AddFile` (commentaire personnalisé pour la version 1)

Design Engineer - XNRGY Climate Systems ULC  - Assignation de la catégorie (si spécifiée)

📧 mohammedamine.elgalai@xnrgy.com- Assignation du lifecycle (si spécifié)

- Assignation de la révision (si spécifiée) via `UpdateFileRevisionNumbers`

---- Application des propriétés (Project, Reference, Module)

- Synchronisation Vault → iProperties pour fichiers Inventor (si `IExplorerUtil` disponible)

## 📄 Licence

### 5. Gestion des fichiers existants

Propriétaire - XNRGY Climate Systems ULC © 2025- Détection du fichier existant

- CheckOut si nécessaire

---- Application des modifications

- CheckIn pour valider

## 🔗 Références

## ⚙️ Configuration

- [Autodesk Vault SDK 2026](https://www.autodesk.com/developer-network/platform-technologies/vault)

- [Autodesk Inventor API 2026](https://www.autodesk.com/developer-network/platform-technologies/inventor)### appsettings.json

- [WPF Documentation](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/)```json

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
