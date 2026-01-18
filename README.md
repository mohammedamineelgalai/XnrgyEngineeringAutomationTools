# XNRGY Engineering Automation Tools (XEAT)# XNRGY Engineering Automation Tools (XEAT)# XNRGY Engineering Automation Tools (XEAT)# XNRGY Engineering Automation Tools



> **Suite d'outils d'automatisation engineering unifiee** pour Autodesk Vault Professional 2026 et Inventor Professional 2026.2

>

> Developpe par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiee** pour Autodesk Vault Professional 2026 et Inventor Professional 2026.2

>

> **Version**: v1.0.0 | **Release**: R 2026-01-15>



---> Developpe par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2



## Description>



**XNRGY Engineering Automation Tools** (XEAT) est une application hub centralisee (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering developpes pour XNRGY Climate Systems. Cette suite vise a simplifier et accelerer les workflows des equipes de design en integrant la gestion Vault, les manipulations Inventor, et les validations qualite dans une interface unifiee.> **Version**: v1.0.0 | **Release**: R 2026-01-10>>



### Objectif Principal



Remplacer les multiples applications standalone par une **plateforme unique** avec :---> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC

- Connexion centralisee a Vault et Inventor

- Interface utilisateur moderne et coherente (themes sombre/clair)

- Partage de services communs (logging, configuration chiffree AES-256)

- Deploiement multi-sites et maintenance simplifiee## Description>

- Parametres centralises via Vault (50+ utilisateurs, 3 sites)

- **Controle a distance via Firebase** (kill switch, maintenance, mises a jour)



---**XNRGY Engineering Automation Tools** (XEAT) est une application hub centralisee (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering developpes pour XNRGY Climate Systems. Cette suite vise a simplifier et accelerer les workflows des equipes de design en integrant la gestion Vault, les manipulations Inventor, et les validations qualite dans une interface unifiee.> **Version**: v1.0.0 | **Release**: R 2026-01-10---



## Modules Integres (9 Modules)



| Module | Description | Statut |### Objectif Principal

|--------|-------------|--------|

| **Upload Module** | Upload automatise vers Vault avec proprietes (Project/Ref/Module) | [+] 100% |

| **Creer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |

| **Place Equipment** | Placement equipements avec Copy Design automatise | [+] 100% |Remplacer les multiples applications standalone par une **plateforme unique** avec :---## Description

| **Smart Tools** | 25+ outils: creation IPT/STEP, generation PDF, iLogic Forms, BOM, etc. | [+] 100% |

| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |- Connexion centralisee a Vault et Inventor

| **ACP** | Assistant Conception Projet - validation points critiques | [+] 100% |

| **Upload Template** | Upload templates vers Vault (reserve Admin) | [+] 100% |- Interface utilisateur moderne et coherente (themes sombre/clair)

| **DXF Verifier** | Validation des fichiers DXF/CSV vs PDF (~97% precision) | [+] 100% |

| **Update Workspace** | Synchronisation librairies et outils depuis Vault | [+] 100% |- Partage de services communs (logging, configuration chiffree AES-256)



---- Deploiement multi-sites et maintenance simplifiee## Description**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.



## Fonctionnalites Cles- Parametres centralises via Vault (50+ utilisateurs, 3 sites)



### Centralisation UI

Toutes les metadonnees de l'application sont centralisees dans `Styles/XnrgyStyles.xaml`:

```xml---

<sys:String x:Key="AppVersion">v1.0.0</sys:String>

<sys:String x:Key="AppReleaseDate">R 2026-01-15</sys:String>**XNRGY Engineering Automation Tools** (XEAT) est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.### Objectif Principal

<sys:String x:Key="AppAuthor">Mohammed Amine Elgalai</sys:String>

<sys:String x:Key="AppCompany">XNRGY CLIMATE SYSTEMS ULC</sys:String>## Modules Integres (9 Modules)

```



### Configuration Centralisee via Vault

- Chiffrement AES-256 des fichiers de configuration| Module | Description | Statut |

- Synchronisation automatique au demarrage

- Acces administrateur verifie via Vault API|--------|-------------|--------|### Objectif PrincipalRemplacer les multiples applications standalone par une **plateforme unique** avec :

- Support 50+ utilisateurs sur 3 sites

| **Upload Module** | Upload automatise vers Vault avec proprietes (Project/Ref/Module) | 100% |

### Firebase Remote Control (NOUVEAU - Janvier 2026)

Integration Firebase Realtime Database pour le controle a distance:| **Creer Module** | Copy Design natif depuis template Library ou projet existant | 100% |- Connexion centralisée à Vault & Inventor

- **Kill Switch**: Blocage instantane de l'application (global, par site, par device)

- **Maintenance Mode**: Desactivation temporaire avec message personnalise| **Place Equipment** | Placement equipements avec Copy Design automatise | 100% |

- **Force Update**: Mise a jour obligatoire vers une version minimum

- **Broadcast Messages**: Diffusion d'annonces a tous les utilisateurs| **Smart Tools** | 25+ outils: creation IPT/STEP, generation PDF, iLogic Forms, BOM, etc. | 100% |Remplacer les multiples applications standalone par une **plateforme unique** avec :- Interface utilisateur moderne et cohérente (thèmes sombre/clair)

- **Device Tracking**: Suivi des postes de travail en temps reel

- **Audit Logs**: Enregistrement des erreurs et sessions dans Firebase| **Checklist HVAC** | Validation modules AHU avec stockage Vault | 100% |

- **Admin Console**: Interface web complete pour la gestion

| **ACP** | Assistant Conception Projet - validation points critiques | 100% |- Connexion centralisée à Vault & Inventor- Partage de services communs (logging, configuration chiffrée AES-256)

---

| **Upload Template** | Upload templates vers Vault (reserve Admin) | 100% |

## Firebase Integration

| **DXF Verifier** | Validation des fichiers DXF/CSV vs PDF (~97% precision) | 100% |- Interface utilisateur moderne et cohérente (thèmes sombre/clair)- Déploiement multi-sites et maintenance simplifiée

### Architecture

```| **Update Workspace** | Synchronisation librairies et outils depuis Vault | 100% |

Firebase Realtime Database

├── appConfig/           # Configuration application- Partage de services communs (logging, configuration chiffrée AES-256)- Paramètres centralisés via Vault (50+ utilisateurs, 3 sites)

├── broadcasts/          # Messages de diffusion

├── devices/             # Registre des postes de travail---

├── featureFlags/        # Feature toggles

├── forceUpdate/         # Configuration mise a jour forcee- Déploiement multi-sites et maintenance simplifiée

├── killSwitch/          # Configuration kill switch

├── maintenance/         # Mode maintenance## Fonctionnalites Cles

├── security/            # Parametres securite

├── statistics/          # Statistiques d'utilisation- Paramètres centralisés via Vault (50+ utilisateurs, 3 sites)---

├── auditLog/            # Journal des actions et erreurs

├── users/               # Gestion utilisateurs### Centralisation UI (NOUVEAU - Janvier 2026)

└── welcomeMessages/     # Messages de bienvenue

```



### Console AdminToutes les metadonnees de l'application sont centralisees dans `Styles/XnrgyStyles.xaml` :

Interface web complete dans `Firebase Realtime Database configuration/admin-panel/`:

- **Dashboard**: Vue d'ensemble en temps reel---## Modules Intégrés

- **Utilisateurs**: Gestion des utilisateurs et permissions

- **Appareils**: Suivi des postes de travail```xml

- **Securite**: Kill Switch, Maintenance, Mise a jour forcee

- **Broadcasts**: Gestion des annonces<!-- Variables globales - MODIFIER ICI UNIQUEMENT -->

- **Telemetrie**: Statistiques d'utilisation

- **Audit Logs**: Historique des actions<sys:String x:Key="AppVersion">v1.0.0</sys:String>



### Scripts de Deploiement<sys:String x:Key="AppReleaseDate">R 2026-01-10</sys:String>## Modules Intégrés (9 Modules)| Module | Description | Statut |



#### build-and-run.ps1 (v2.2.0)<sys:String x:Key="AppAuthor">Mohammed Amine Elgalai</sys:String>

Script principal de build et deploiement:

```powershell<sys:String x:Key="AppCompany">XNRGY CLIMATE SYSTEMS ULC</sys:String>|--------|-------------|--------|

# Build simple

.\build-and-run.ps1<sys:String x:Key="AppShortName">XEAT</sys:String>



# Build + Deploy Firebase<sys:String x:Key="AppFullName">XNRGY Engineering Automation Tools</sys:String>| Module | Description | Statut || **Upload Module** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | [+] 100% |

.\build-and-run.ps1 -Deploy

```

# Build + Deploy + Publish GitHub

.\build-and-run.ps1 -Deploy -Publish|--------|-------------|--------|| **Créer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |



# Sync donnees avant deploy (preserve devices/users/auditLog)**Avantages:**

.\build-and-run.ps1 -SyncFirebase -Deploy

```- Mise a jour de version/date en **un seul endroit**| **Upload Module** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ 100% || **Réglages Admin** | Configuration centralisée et synchronisée via Vault (AES-256) | [+] 100% |



#### Sync-Firebase.ps1- Tous les titres de fenetres utilisent `StaticResource`

Script de synchronisation pour preserver les donnees dynamiques avant deploiement:

```powershell- Toutes les barres copyright utilisent `StaticResource`| **Créer Module** | Copy Design natif depuis template Library ou projet existant | ✅ 100% || **Upload Template** | Upload templates vers Vault (réservé Admin) | [+] 100% |

# Mode interactif

.\Sync-Firebase.ps1- 22+ fenetres synchronisees automatiquement



# Mode pre-deployment (automatique)| **Place Equipment** | Placement équipements avec Copy Design automatisé | ✅ 100% || **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |

.\Sync-Firebase.ps1 -PreDeploy

```### Configuration Centralisee via Vault



**Donnees preservees:**| **Smart Tools** | 25+ outils: création IPT/STEP, génération PDF, iLogic Forms, BOM, etc. | ✅ 100% || **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | [~] Planifié |

- devices (postes enregistres)

- users (utilisateurs)- Chiffrement AES-256 des fichiers de configuration

- userActiveSessions (sessions actives)

- auditLog (journal des actions)- Synchronisation automatique au demarrage| **Checklist HVAC** | Validation modules AHU avec stockage Vault | ✅ 100% || **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |

- statistics (statistiques)

- telemetryEvents (telemetrie)- Acces administrateur verifie via Vault API

- errorReports (rapports d'erreurs)

- broadcasts (messages actifs)- Support 50+ utilisateurs sur 3 sites| **ACP** | Assistant Conception Projet - validation points critiques | ✅ 100% || **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |



### Services C#



#### FirebaseAuditService.cs### Smart Tools (25+ Outils)| **Upload Template** | Upload templates vers Vault (réservé Admin) | ✅ 100% || **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifié |

Service singleton pour l'envoi des erreurs et sessions a Firebase:

```csharp

// Initialisation au demarrage

await FirebaseAuditService.Instance.InitializeAsync();| Categorie | Outils || **DXF Verifier** | Validation des fichiers DXF/CSV vs PDF (~97% précision) | ✅ 100% |



// Envoi automatique des erreurs (via Logger)|-----------|--------|

Logger.LogException("Context", exception);  // Envoie automatiquement a Firebase

Logger.Error("Message");                     // Envoie automatiquement a Firebase| **Creation** | Creer IPT, Creer STEP/IPT, Creer PDF, Creer PNG || **Update Workspace** | Synchronisation librairies et outils depuis Vault | ✅ 100% |---



// Fin de session| **Export** | Export IAM, Export Drawing, Export BOM |

await FirebaseAuditService.Instance.RegisterSessionEndAsync();

```| **iProperties** | Gestionnaire Proprietes Batch, iProperties Summary |



#### FirebaseRemoteConfigService.cs| **Analyse** | Rapport Contraintes, Analyse Assembly |

Verification de la configuration au demarrage:

- Kill Switch (global, par site, par device, par utilisateur)| **Utilitaires** | Rename Files, Copy Files, Delete Backup Files |---## Architecture du Projet

- Mode Maintenance

- Mise a jour disponible/forcee

- Messages broadcast

---

---



## Prerequis

## Architecture du Projet## Fonctionnalités Clés```

### Logiciels Requis

- **Windows 10/11** (x64)

- **Autodesk Inventor Professional 2026.2**

- **Autodesk Vault Professional 2026** (SDK v31.0.84)```XnrgyEngineeringAutomationTools/

- **.NET Framework 4.8**

- **Visual Studio 2022** (pour le developpement)XnrgyEngineeringAutomationTools/



### Acces Reseau+-- App.xaml(.cs)                    # Point d'entree application### Centralisation UI (NOUVEAU - Janvier 2026)├── App.xaml(.cs)                    # Point d'entrée application

- Serveur Vault: `XNRGY-VAULT` / `XNRGY-VAULT2`

- Firebase: `xeat-remote-control-default-rtdb.firebaseio.com`+-- MainWindow.xaml(.cs)             # Dashboard principal (hub)



---+-- appsettings.json                 # Configuration utilisateur├── MainWindow.xaml(.cs)             # Dashboard principal (hub)



## Installation|



### Pour les Utilisateurs+-- Modules/                         # === MODULES FONCTIONNELS ===Toutes les métadonnées de l'application sont centralisées dans `Styles/XnrgyStyles.xaml` :├── appsettings.json                 # Configuration utilisateur

1. Telecharger depuis GitHub Releases ou Vault

2. Executer `XnrgyInstaller.exe`|   +-- UploadModule/                # Upload vers Vault

3. L'installeur enregistre automatiquement le poste dans Firebase

|   |   +-- Models/│

### Pour les Developpeurs

```powershell|   |   |   +-- VaultUploadFileItem.cs

# Cloner le depot

git clone https://github.com/MohammedAmineElgalai/XnrgyEngineeringAutomationTools.git|   |   |   +-- VaultUploadModels.cs```xml├── Modules/                         # Modules fonctionnels isolés



# Build|   |   +-- Views/

cd XnrgyEngineeringAutomationTools

.\build-and-run.ps1|   |       +-- UploadModuleWindow.xaml(.cs)<!-- Variables globales - MODIFIER ICI UNIQUEMENT -->│   ├── UploadModule/                # Module Upload Vault



# Build + Deploy complet|   |

.\build-and-run.ps1 -Deploy -Publish

```|   +-- CreateModule/                # Pack and Go / Copy Design<sys:String x:Key="AppVersion">v1.0.0</sys:String>│   │   ├── Models/



---|   |   +-- Models/



## Configuration|   |   |   +-- CreateModuleRequest.cs<sys:String x:Key="AppReleaseDate">R 2026-01-10</sys:String>│   │   │   ├── VaultUploadFileItem.cs



### Fichiers de Configuration|   |   |   +-- CreateModuleSettings.cs

| Fichier | Description |

|---------|-------------||   |   +-- Services/<sys:String x:Key="AppAuthor">Mohammed Amine Elgalai</sys:String>│   │   │   └── VaultUploadModels.cs

| `appsettings.json` | Configuration locale (Vault server, chemins) |

| `firebase-init.json` | Structure Firebase initiale ||   |   |   +-- InventorCopyDesignService.cs  # ~3000 lignes

| `XnrgyStyles.xaml` | Metadonnees application |

|   |   |   +-- ModuleCopyService.cs<sys:String x:Key="AppCompany">XNRGY CLIMATE SYSTEMS ULC</sys:String>│   │   └── Views/

### Variables d'Environnement (optionnel)

```|   |   |   +-- PdfCoverService.cs

XEAT_VAULT_SERVER=XNRGY-VAULT

XEAT_DEBUG_MODE=false|   |   +-- Views/<sys:String x:Key="AppShortName">XEAT</sys:String>│   │       └── UploadModuleWindow.xaml(.cs)

```

|   |       +-- CreateModuleWindow.xaml(.cs)

---

|   |       +-- CreateModuleSettingsWindow.xaml(.cs)<sys:String x:Key="AppFullName">XNRGY Engineering Automation Tools</sys:String>│   │

## Logs

|   |

### Emplacements

- **Application**: `bin\Release\Logs\VaultSDK_POC_*.log`|   +-- PlaceEquipment/              # Placement equipements (NOUVEAU)```│   ├── CreateModule/                # Module Pack & Go / Copy Design

- **Installer**: `%AppData%\XnrgyInstaller\Logs\`

- **Firebase**: Console Admin > Audit Logs|   |   +-- Models/



### Format|   |   |   +-- EquipmentPlacementModels.cs│   │   ├── Models/

```

[2026-01-15 14:30:22.123] [INFO   ] [+] Message|   |   +-- Services/

[2026-01-15 14:30:23.456] [ERROR  ] [-] Erreur message

```|   |   |   +-- EquipmentPlacementService.cs**Avantages:**│   │   │   ├── CreateModuleRequest.cs



---|   |   |   +-- EquipmentCopyDesignService.cs



## Regles de Developpement|   |   +-- Views/- Mise à jour de version/date en **un seul endroit**│   │   │   └── CreateModuleSettings.cs



### Build Tools|   |       +-- PlaceEquipmentWindow.xaml(.cs)  # ~4200 lignes

- **NET Framework 4.8**: MSBuild uniquement (pas `dotnet build`)

- Utiliser toujours `.\build-and-run.ps1`|   |- Tous les titres de fenêtres utilisent `StaticResource`│   │   ├── Services/



### Emojis dans le Code|   +-- SmartTools/                  # 25+ outils Inventor (COMPLET)

**INTERDITS** dans le code C#, logs, messages:

```csharp|   |   +-- Resources/- Toutes les barres copyright utilisent `StaticResource`│   │   │   ├── InventorCopyDesignService.cs

// INTERDIT

Logger.Info("✅ Connexion etablie");|   |   +-- Services/



// CORRECT|   |   |   +-- SmartToolsService.cs  # ~5700 lignes- 22+ fenêtres synchronisées automatiquement│   │   │   └── ModuleCopyService.cs

Logger.Info("[+] Connexion etablie");

```|   |   +-- Views/



**AUTORISES** dans les interfaces XAML:|   |       +-- SmartToolsWindow.xaml(.cs)│   │   └── Views/

```xml

<!-- OK pour XAML -->|   |       +-- ConstraintReportWindow.xaml(.cs)

<TextBlock Text="✅ Succes"/>

```|   |       +-- CustomPropertyBatchWindow.xaml(.cs)### Configuration Centralisée via Vault│   │       ├── CreateModuleWindow.xaml(.cs)



### Marqueurs ASCII|   |       +-- ExportOptionsWindow.xaml(.cs)

| Interdit | Remplacement | Usage |

|----------|--------------|-------||   |       +-- HtmlPopupWindow.xaml(.cs)│   │       └── CreateModuleSettingsWindow.xaml(.cs)

| ❌ | [-] | Erreur |

| ✅ | [+] | Succes ||   |       +-- IPropertiesWindow.xaml(.cs)

| ⚠️ | [!] | Avertissement |

| 🔄 | [>] | Traitement ||   |       +-- ProgressWindow.xaml(.cs)- Chiffrement AES-256 des fichiers de configuration│   │

| 📁 | [i] | Info |

| ⏳ | [~] | Attente ||   |       +-- SmartProgressWindow.xaml(.cs)



---|   |- Synchronisation automatique au démarrage│   ├── UploadTemplate/              # Module Upload Template



## Architecture du Projet|   +-- ChecklistHVAC/               # Validation AHU



```|   |   +-- Services/- Accès administrateur vérifié via Vault API│   │   └── Views/

XnrgyEngineeringAutomationTools/

├── App.xaml(.cs)                    # Point d'entree + Firebase check|   |   |   +-- ChecklistSyncService.cs

├── MainWindow.xaml(.cs)             # Fenetre principale hub

├── Modules/                         # Modules fonctionnels|   |   +-- Views/- Support 50+ utilisateurs sur 3 sites│   │       └── UploadTemplateWindow.xaml(.cs)

│   ├── UploadVault/                 # Upload vers Vault

│   ├── CreateModule/                # Copy Design|   |       +-- ChecklistHVACWindow.xaml(.cs)

│   ├── PlaceEquipment/              # Placement equipements

│   ├── SmartTools/                  # Outils Inventor|   |│   │

│   ├── ChecklistHVAC/               # Validation AHU

│   ├── ACP/                         # Assistant Conception|   +-- ACP/                         # Assistant Conception Projet

│   ├── UploadTemplates/             # Upload templates

│   ├── DXFVerifier/                 # Validation DXF|   |   +-- Services/### Smart Tools (25+ Outils)│   └── ChecklistHVAC/              # Module Checklist HVAC

│   └── UpdateWorkspace/             # Sync Workspace

├── Services/                        # Services partages|   |   |   +-- ACPExcelService.cs

│   ├── VaultSDKService.cs           # API Vault SDK

│   ├── Logger.cs                    # Logging + Firebase|   |   |   +-- ACPSyncService.cs│       └── Views/

│   ├── FirebaseAuditService.cs      # Audit Firebase

│   ├── FirebaseRemoteConfigService.cs # Config distante|   |   +-- Views/

│   ├── DeviceTrackingService.cs     # Tracking device

│   └── AutoUpdateService.cs         # Mises a jour|   |       +-- ACPWindow.xaml(.cs)| Catégorie | Outils |│           └── ChecklistHVACWindow.xaml(.cs)

├── Installer/                       # Installeur

│   ├── InstallationService.cs       # Logique installation|   |

│   └── FirebaseService.cs           # Firebase pour installer

├── Firebase Realtime Database configuration/|   +-- UploadTemplate/              # Upload Template (Admin)|-----------|--------|│

│   ├── firebase-init.json           # Structure initiale

│   └── admin-panel/                 # Console admin web|   |   +-- Views/

│       └── index.html               # Interface admin

├── Styles/|   |       +-- UploadTemplateWindow.xaml(.cs)| **Création** | Créer IPT, Créer STEP/IPT, Créer PDF, Créer PNG |├── Shared/                          # Composants partagés

│   └── XnrgyStyles.xaml             # Styles + metadonnees

├── build-and-run.ps1                # Script build (v2.2.0)|   |

└── Sync-Firebase.ps1                # Script sync Firebase

```|   +-- DXFVerifier/                 # Validation DXF (MIGRE)| **Export** | Export IAM, Export Drawing, Export BOM |│   ├── Views/                       # Fenêtres partagées



---|   |   +-- Services/



## Changelog Recent|   |   |   +-- PdfAnalyzer.cs| **iProperties** | Gestionnaire Propriétés Batch, iProperties Summary |│   │   ├── LoginWindow.xaml(.cs)    # Connexion Vault



### v1.0.0 (2026-01-15)|   |   |   +-- ExcelManager.cs

- **Firebase Integration**: Kill switch, maintenance, force update, broadcasts

- **Audit Logs Firebase**: Erreurs et sessions envoyees automatiquement|   |   +-- Views/| **Analyse** | Rapport Contraintes, Analyse Assembly |│   │   ├── ModuleSelectionWindow.xaml(.cs)  # Sélection module

- **Installer Firebase**: Enregistrement install/uninstall dans Firebase

- **Sync-Firebase.ps1**: Preservation des donnees dynamiques lors des deploiements|   |       +-- DXFVerifierWindow.xaml(.cs)

- **Admin Console**: Interface web complete pour gestion Firebase

- **build-and-run.ps1 v2.2.0**: Nouveaux parametres -Deploy, -SyncFirebase, -Publish|   || **Utilitaires** | Rename Files, Copy Files, Delete Backup Files |│   │   ├── PreviewWindow.xaml(.cs)  # Prévisualisation

- 9 modules integres et fonctionnels

- Code signing avec certificat XNRGY|   +-- UpdateWorkspace/             # Sync Workspace



---|       +-- Views/│   │   └── XnrgyMessageBox.xaml(.cs)  # MessageBox moderne



## Auteur|           +-- UpdateWorkspaceWindow.xaml(.cs)



**Mohammed Amine Elgalai**  |---│   ├── Models/                      # Modèles partagés (vide pour l'instant)

Vault/Inventor Software Developer  

XNRGY Climate Systems ULC+-- Shared/                          # === COMPOSANTS PARTAGES ===



- GitHub: [@MohammedAmineElgalai](https://github.com/MohammedAmineElgalai)|   +-- Views/│   └── Services/                    # Services partagés (vide pour l'instant)

- Email: melgalai@xnrgy.com

|       +-- LoginWindow.xaml(.cs)         # Connexion Vault

---

|       +-- ModuleSelectionWindow.xaml(.cs)## Architecture du Projet│

## Licence

|       +-- PreviewWindow.xaml(.cs)

Copyright (c) 2024-2026 XNRGY Climate Systems ULC. Tous droits reserves.

Usage interne uniquement.|       +-- XnrgyMessageBox.xaml(.cs)     # MessageBox moderne├── Services/                        # Services métier partagés


|

+-- Services/                        # === SERVICES METIER ===```│   ├── VaultSDKService.cs           # SDK Vault v31.0.84 (~1200 lignes)

|   +-- VaultSDKService.cs           # SDK Vault v31.0.84 (~3200 lignes)

|   +-- VaultSettingsService.cs      # Config chiffree AES-256XnrgyEngineeringAutomationTools/│   ├── VaultSettingsService.cs      # Config chiffrée + sync Vault

|   +-- InventorService.cs           # Connexion Inventor COM

|   +-- InventorPropertyService.cs   # iProperties via Inventor├── App.xaml(.cs)                    # Point d'entrée application│   ├── InventorService.cs           # Connexion Inventor COM

|   +-- ApprenticePropertyService.cs # iProperties via Apprentice

|   +-- UpdateWorkspaceService.cs    # Service sync workspace├── MainWindow.xaml(.cs)             # Dashboard principal (hub)│   ├── InventorPropertyService.cs   # iProperties via Inventor COM

|   +-- Logger.cs                    # Logging NLog UTF-8

|   +-- ThemeHelper.cs               # Gestion themes├── appsettings.json                 # Configuration utilisateur│   ├── ApprenticePropertyService.cs # Lecture iProperties via Apprentice

|

+-- Models/                          # === MODELES PARTAGES ===││   ├── OlePropertyService.cs         # Lecture OLE via OpenMCDF

|   +-- ApplicationConfiguration.cs

|   +-- FileItem.cs├── Modules/                         # ═══ MODULES FONCTIONNELS ═══│   ├── NativeOlePropertyService.cs  # Lecture OLE native

|   +-- ModuleInfo.cs

|   +-- ProjectInfo.cs│   ├── UploadModule/                # Upload vers Vault│   ├── WindowsPropertyService.cs    # Lecture propriétés Windows Shell

|   +-- ProjectProperties.cs

|   +-- VaultConfiguration.cs│   │   ├── Models/│   ├── Logger.cs                    # Logging NLog UTF-8

|

+-- Styles/                          # === STYLES CENTRALISES ===│   │   │   ├── VaultUploadFileItem.cs│   ├── SimpleLogger.cs              # Logger simple

|   +-- XnrgyStyles.xaml             # 1350+ lignes (12 sections)

|       +-- Section 1: Palettes couleurs│   │   │   └── VaultUploadModels.cs│   ├── JournalColorService.cs       # Couleurs journal (succès/erreur)

|       +-- Section 1.5: Version/Copyright (CENTRALISES)

|       +-- Section 2: Styles boutons│   │   └── Views/│   ├── ThemeHelper.cs               # Gestion thèmes sombre/clair

|       +-- Section 3: Styles inputs

|       +-- Section 4: Styles DataGrid│   │       └── UploadModuleWindow.xaml(.cs)│   ├── UserPreferencesManager.cs   # Préférences utilisateur

|       +-- Section 5: Styles GroupBox

|       +-- Section 6: Styles ComboBox│   ││   ├── SettingsService.cs           # Gestion settings app

|       +-- Section 7: Styles ScrollBar

|       +-- Section 8: Styles CheckBox/Radio│   ├── CreateModule/                # Pack & Go / Copy Design│   └── CredentialsManager.cs        # Gestion credentials chiffrées

|       +-- Section 9: Styles Progress

|       +-- Section 10: Styles Tooltip│   │   ├── Models/│

|       +-- Section 11: Animations

|│   │   │   ├── CreateModuleRequest.cs├── Models/                          # Modèles de données partagés

+-- Converters/                      # Convertisseurs WPF

+-- Resources/                       # Images et icones│   │   │   └── CreateModuleSettings.cs│   ├── ApplicationConfiguration.cs  # Configuration application

+-- Logs/                            # Fichiers logs

```│   │   ├── Services/│   ├── CategoryItem.cs              # Item catégorie pour ComboBox



---│   │   │   ├── InventorCopyDesignService.cs  # ~3000 lignes│   ├── FileItem.cs                  # Item fichier pour DataGrid



## Prerequis│   │   │   ├── ModuleCopyService.cs│   ├── FileToUpload.cs              # Fichier à uploader



| Composant | Version |│   │   │   └── PdfCoverService.cs│   ├── LifecycleDefinitionItem.cs  # Lifecycle Definition

|-----------|---------|

| Autodesk Inventor Professional | 2026.2 |│   │   └── Views/│   ├── LifecycleStateItem.cs        # Lifecycle State

| Autodesk Vault Professional | 2026 |

| Vault SDK | v31.0.84 |│   │       ├── CreateModuleWindow.xaml(.cs)│   ├── ModuleInfo.cs                # Informations module

| .NET Framework | 4.8 |

| Visual Studio | 2022 (MSBuild 18.0.0+) |│   │       └── CreateModuleSettingsWindow.xaml(.cs)│   ├── ProjectInfo.cs               # Informations projet

| Windows | 10/11 x64 |

│   ││   ├── ProjectProperties.cs         # Propriétés Project/Ref/Module

---

│   ├── PlaceEquipment/              # Placement équipements (NOUVEAU)│   └── VaultConfiguration.cs        # Configuration Vault

## Installation

│   │   ├── Models/│

### Build depuis les sources

│   │   │   └── EquipmentPlacementModels.cs├── ViewModels/                      # MVVM ViewModels

```powershell

# Cloner le repository│   │   ├── Services/│   └── AppMainViewModel.cs          # ViewModel principal

git clone https://github.com/mohammedamineelgalai/VaultAutomationTool.git

cd XnrgyEngineeringAutomationTools│   │   │   ├── EquipmentPlacementService.cs│



# METHODE OBLIGATOIRE - Script build-and-run.ps1│   │   │   └── EquipmentCopyDesignService.cs├── Converters/                      # Convertisseurs WPF

.\build-and-run.ps1

│   │   └── Views/│   ├── AnyOperationActiveConverter.cs

# Options disponibles:

.\build-and-run.ps1 -Clean      # Nettoyer avant compilation│   │       └── PlaceEquipmentWindow.xaml(.cs)  # ~4200 lignes│   ├── BooleanToColorConverter.cs

.\build-and-run.ps1 -NoBuild    # Lancer sans recompiler

```│   ││   ├── BooleanToTextConverter.cs



> **IMPORTANT**: Toujours utiliser `build-and-run.ps1`. Ne JAMAIS utiliser `dotnet build` pour les projets .NET Framework 4.8 WPF.│   ├── SmartTools/                  # 25+ outils Inventor (COMPLET)│   ├── InverseBooleanConverter.cs



### Executable│   │   ├── Resources/│   ├── InverseBooleanToVisibilityConverter.cs



L'application compilee se trouve dans:│   │   ├── Services/│   ├── NullToVisibilityConverter.cs

```

bin\Release\XnrgyEngineeringAutomationTools.exe│   │   │   └── SmartToolsService.cs  # ~5700 lignes│   └── ProgressToWidthConverter.cs

```

│   │   └── Views/│

---

│   │       ├── SmartToolsWindow.xaml(.cs)├── Styles/                          # Styles centralisés

## Centralisation des Metadonnees

│   │       ├── ConstraintReportWindow.xaml(.cs)│   └── XnrgyStyles.xaml             # 50+ styles partagés (couleurs, boutons, inputs)

### Fichier: `Styles/XnrgyStyles.xaml`

│   │       ├── CustomPropertyBatchWindow.xaml(.cs)│

#### Variables Globales (modifier ici uniquement)

```xml│   │       ├── ExportOptionsWindow.xaml(.cs)├── Resources/                       # Images et icônes

<sys:String x:Key="AppVersion">v1.0.0</sys:String>

<sys:String x:Key="AppReleaseDate">R 2026-01-10</sys:String>│   │       ├── HtmlPopupWindow.xaml(.cs)│   ├── XnrgyEngineeringAutomationTools.ico

<sys:String x:Key="AppAuthor">Mohammed Amine Elgalai</sys:String>

<sys:String x:Key="AppCompany">XNRGY CLIMATE SYSTEMS ULC</sys:String>│   │       ├── IPropertiesWindow.xaml(.cs)│   └── XnrgyEngineeringAutomationTools.png

```

│   │       ├── ProgressWindow.xaml(.cs)│

#### Titres de Fenetres (22 ressources)

```xml│   │       └── SmartProgressWindow.xaml(.cs)├── Assets/                          # Assets visuels

<!-- Utilisation dans XAML: Title="{StaticResource WindowTitleXXX}" -->

<sys:String x:Key="WindowTitleMain">XNRGY Engineering Automation Tools - v1.0.0 - R 2026-01-10 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC</sys:String>│   ││   └── Icons/

<sys:String x:Key="WindowTitleUploadModule">Upload Module - XEAT - v1.0.0 - ...</sys:String>

<sys:String x:Key="WindowTitleSmartTools">Smart Tools - XEAT - v1.0.0 - ...</sys:String>│   ├── ChecklistHVAC/               # Validation AHU│       ├── ChecklistHVAC.png

<!-- ... 19 autres titres -->

```│   │   ├── Services/│       ├── DXFVerifier.ico



#### Barres Copyright (16 ressources)│   │   │   └── ChecklistSyncService.cs│       ├── DXFVerifier.png

```xml

<!-- Utilisation dans XAML: Text="{StaticResource CopyrightXXX}" -->│   │   └── Views/│       ├── VaultUpload.ico

<sys:String x:Key="CopyrightMain">XNRGY Engineering Automation Tools - v1.0.0 - R 2026-01-10 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC</sys:String>

<sys:String x:Key="CopyrightSmartTools">Smart Tools - XEAT - v1.0.0 - ...</sys:String>│   │       └── ChecklistHVACWindow.xaml(.cs)│       ├── VaultUpload.png

<!-- ... 14 autres copyrights -->

```│   ││       └── XnrgyTools.ico



---│   ├── ACP/                         # Assistant Conception Projet│



## Configuration│   │   ├── Services/├── Tools/                           # Outils utilitaires



### Configuration Vault via fichier centralise│   │   │   ├── ACPExcelService.cs│   └── VaultBulkUploader/           # Outil console upload massif



**Chemin Vault:**│   │   │   └── ACPSyncService.cs│

```

$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/│   │   └── Views/├── Scripts/                          # Scripts PowerShell

  XnrgyEngineeringAutomationToolsApp/XnrgyEngineeringAutomationToolsSettings.config

```│   │       └── ACPWindow.xaml(.cs)│   ├── CleanInventor2023Registry.ps1



**Chemin Local:**│   ││   ├── Prepare-TemplateFiles.ps1

```

C:\Vault\Engineering\Inventor_Standards\Automation_Standard\Configuration_Files\│   ├── UploadTemplate/              # Upload Template (Admin)│   └── Upload-ToVaultProd.ps1

  XnrgyEngineeringAutomationToolsApp\XnrgyEngineeringAutomationToolsSettings.config

```│   │   └── Views/│



### Proprietes Vault (IDs fixes)│   │       └── UploadTemplateWindow.xaml(.cs)├── Temp&Test/                       # Fichiers temporaires/test (exclus du build)



| Propriete | ID | Description |│   ││   ├── DiagnoseOleProperties.cs

|-----------|-----|-------------|

| Project | 112 | Numero de projet (ex: 10359) |│   ├── DXFVerifier/                 # Validation DXF (MIGRÉ)│   └── TestWindowsPropertyService.cs

| Reference | 121 | Numero de reference (ex: 09) |

| Module | 122 | Numero de module (ex: 03) |│   │   ├── Services/│



### Pattern de chemin standard│   │   │   ├── PdfAnalyzer.cs├── Backups/                         # Sauvegardes locales



```│   │   │   └── ExcelManager.cs│   └── BACKUP_YYYYMMDD_HHMMSS/

C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]

                              |         |       |│   │   └── Views/│

Exemple:                   10359      REF09     M03

```│   │       └── DXFVerifierWindow.xaml(.cs)└── build-and-run.ps1                # Script compilation MSBuild automatique



---│   │```



## Logs│   └── UpdateWorkspace/             # Sync Workspace



| Module | Emplacement |│       └── Views/---

|--------|-------------|

| Application | `bin\Release\Logs\XEAT_*.log` |│           └── UpdateWorkspaceWindow.xaml(.cs)

| Vault Upload | `bin\Release\Logs\VaultUpload_*.log` |

| Copy Design | `bin\Release\Logs\CopyDesign_*.log` |│## Fonctionnalités Implémentées

| Smart Tools | `bin\Release\Logs\SmartTools_*.log` |

├── Shared/                          # ═══ COMPOSANTS PARTAGÉS ═══

### Format de log

```│   └── Views/### 1. Upload Module (100%)

[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] [+] Message

```│       ├── LoginWindow.xaml(.cs)         # Connexion Vault



---│       ├── ModuleSelectionWindow.xaml(.cs)Module intégré pour l'upload de fichiers vers Vault avec gestion complète des propriétés :



## Regles de Developpement│       ├── PreviewWindow.xaml(.cs)



### Emojis dans le Code│       └── XnrgyMessageBox.xaml(.cs)     # MessageBox moderne- **Connexion centralisée** - Utilise la connexion Vault de l'app principale



#### INTERDITS (code C#, logs):│- **Scan automatique** des modules engineering avec extraction propriétés

Tous les emojis sont interdits dans le code source et les logs.

├── Services/                        # ═══ SERVICES MÉTIER ═══- **Séparation Inventor/Non-Inventor** dans deux DataGrids avec headers visibles

#### AUTORISES (interfaces XAML uniquement):

checkmark, X, warning, info, refresh, file, folder, search, list, hourglass│   ├── VaultSDKService.cs           # SDK Vault v31.0.84 (~3200 lignes)- **Application automatique** des propriétés métier:



#### Marqueurs ASCII pour logs:│   ├── VaultSettingsService.cs      # Config chiffrée AES-256  - Project (ID=112)

| Symbole | Remplacement | Usage |

|---------|--------------|-------|│   ├── InventorService.cs           # Connexion Inventor COM  - Reference (ID=121)

| X | [-] | Erreur |

| checkmark | [+] | Succes |│   ├── InventorPropertyService.cs   # iProperties via Inventor  - Module (ID=122)

| warning | [!] | Avertissement |

| refresh | [>] | Traitement |│   ├── ApprenticePropertyService.cs # iProperties via Apprentice- **Assignation complète**:



### Build Tool│   ├── UpdateWorkspaceService.cs    # Service sync workspace  - Catégories Vault



```powershell│   ├── Logger.cs                    # Logging NLog UTF-8  - Lifecycle Definitions et States

# TOUJOURS utiliser:

.\build-and-run.ps1│   └── ThemeHelper.cs               # Gestion thèmes  - Révisions



# JAMAIS utiliser:│- **Synchronisation Vault vers iProperties** via `IExplorerUtil`

dotnet build  # Ne genere pas les .g.cs pour WPF

```├── Models/                          # ═══ MODÈLES PARTAGÉS ═══- **Journal des opérations** avec barre de progression style Créer Module



---│   ├── ApplicationConfiguration.cs- **Contrôles**: Pause/Stop/Annuler pendant l'upload



## Changelog Recent│   ├── FileItem.cs- **Styles DataGrid** avec headers fond sombre et texte bleu XNRGY



### v1.0.0 (R 2026-01-10)│   ├── ModuleInfo.cs- **Interface moderne** avec thèmes sombre/clair



#### Centralisation UI│   ├── ProjectInfo.cs

- [+] Toutes les metadonnees centralisees dans `XnrgyStyles.xaml`

- [+] 22 titres de fenetres via `StaticResource`│   ├── ProjectProperties.cs### 2. Créer Module - Copy Design (100%)

- [+] 16 textes copyright via `StaticResource`

- [+] Grand titre MainWindow centralise│   └── VaultConfiguration.cs

- [+] Suppression de tous les `MultiBinding` redondants

│**Sources disponibles :**

#### Smart Tools

- [+] 25+ outils Inventor complets├── Styles/                          # ═══ STYLES CENTRALISÉS ═══- Depuis Template : `$/Engineering/Library/Xnrgy_Module` (1083 fichiers Inventor)

- [+] Service SmartToolsService (~5700 lignes)

- [+] Fenetres: ConstraintReport, CustomPropertyBatch, IProperties, etc.│   └── XnrgyStyles.xaml             # 1350+ lignes (12 sections)- Depuis Projet Existant : Sélection d'un projet local ou Vault



#### Place Equipment│       ├── Section 1: Palettes couleurs

- [+] Module complet avec Copy Design automatise

- [+] PlaceEquipmentWindow (~4200 lignes)│       ├── Section 1.5: Version/Copyright (CENTRALISÉS)**Workflow automatisé :**

- [+] Services: EquipmentPlacementService, EquipmentCopyDesignService

│       ├── Section 2: Styles boutons1. Switch vers projet source (IPJ)

#### DXF Verifier

- [+] Migration complete vers XEAT│       ├── Section 3: Styles inputs2. Ouverture Top Assembly (Module_.iam)

- [+] Validation DXF/CSV vs PDF (~97% precision)

- [+] Version specifique: v1.2.0│       ├── Section 4: Styles DataGrid3. Application iProperties sur le template



#### Architecture│       ├── Section 5: Styles GroupBox4. Collecte de toutes les références (bottom-up)

- [+] 9 modules fonctionnels complets

- [+] Structure modulaire Modules/{Module}/Views|Services|Models│       ├── Section 6: Styles ComboBox5. Copy Design natif avec SaveAs (IPT -> IAM -> Top Assembly)

- [+] Styles centralises (1350+ lignes, 12 sections)

│       ├── Section 7: Styles ScrollBar6. Traitement des dessins (.idw) avec mise à jour des références

---

│       ├── Section 8: Styles CheckBox/Radio7. **Mise à jour des références des composants suppressed** (v1.1)

## Auteur

│       ├── Section 9: Styles Progress8. Copie des fichiers orphelins (1059 fichiers non-référencés)

**Mohammed Amine Elgalai**  

XNRGY Climate Systems ULC  │       ├── Section 10: Styles Tooltip9. Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)

2025-2026

│       └── Section 11: Animations10. Renommage du fichier .ipj

---

│11. Switch vers le nouveau projet

## Licence

├── Converters/                      # Convertisseurs WPF12. Application des iProperties finales et paramètres Inventor

Proprietaire - XNRGY Climate Systems ULC. Tous droits reserves.

├── Resources/                       # Images et icônes13. Design View -> "Default", masquage Workfeatures

└── Logs/                            # Fichiers logs14. Vue ISO + Zoom All (Fit)

```15. Update All (rebuild) + Save All

16. Module reste ouvert pour le dessinateur

---

**Gestion intelligente des références :**

## Prérequis- Fichiers Library (IPT_Typical_Drawing) : Liens préservés

- Fichiers Module : Copies avec références mises à jour

| Composant | Version |- Fichiers IDW : Références corrigées via `PutLogicalFileNameUsingFull`

|-----------|---------|- **Composants suppressed** : Références mises à jour même si supprimés

| Autodesk Inventor Professional | 2026.2 |

| Autodesk Vault Professional | 2026 |**Options de renommage (v1.1) :**

| Vault SDK | v31.0.84 |- Rechercher/Remplacer (cumulatif sur NewFileName)

| .NET Framework | 4.8 |- Préfixe/Suffixe (appliqué sur OriginalFileName)

| Visual Studio | 2022 (MSBuild 18.0.0+) |- **Checkbox "Inclure fichiers non-Inventor"**

| Windows | 10/11 x64 |

### 3. Réglages Admin (100%)

---

**Système de configuration centralisée :**

## Installation- Chiffrement AES-256 des fichiers de configuration

- Synchronisation automatique via Vault au démarrage

### Build depuis les sources- Accès restreint aux administrateurs (Role "Administrator" ou Groupe "Admin_Designer")

- Déploiement multi-sites : Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs

```powershell

# Cloner le repository**Chemin Vault :**

git clone https://github.com/mohammedamineelgalai/VaultAutomationTool.git```

cd XnrgyEngineeringAutomationTools$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/

```

# MÉTHODE OBLIGATOIRE - Script build-and-run.ps1

.\build-and-run.ps1**Sections configurables :**

- Liste des initiales designers (26 entrées + "Autre...")

# Options disponibles:- Chemins templates et projets

.\build-and-run.ps1 -Clean      # Nettoyer avant compilation- Extensions Inventor supportées

.\build-and-run.ps1 -NoBuild    # Lancer sans recompiler- Dossiers/fichiers exclus

```- Noms des iProperties



> ⚠️ **IMPORTANT**: Toujours utiliser `build-and-run.ps1`. Ne JAMAIS utiliser `dotnet build` pour les projets .NET Framework 4.8 WPF.**Interface moderne :**

- Styles uniformisés avec effets glow sur boutons

### Exécutable- GroupBox avec titres orange (#FF8C00)

- Thèmes sombre/clair cohérents

L'application compilée se trouve dans:

```### 4. Upload Template (100%)

bin\Release\XnrgyEngineeringAutomationTools.exe

```- **Réservé aux administrateurs** - Message XnrgyMessageBox si non-admin

- **Upload templates** depuis Library vers Vault

---- **Utilise la connexion partagée** de l'app principale

- **Journal intégré** avec barre de progression

## Centralisation des Métadonnées- **Interface alignée** avec Upload Module (hauteurs, styles, thèmes)



### Fichier: `Styles/XnrgyStyles.xaml`### 5. Checklist HVAC (100%)



#### Variables Globales (modifier ici uniquement)- Validation des modules AHU

```xml- Checklist interactive avec critères XNRGY

<sys:String x:Key="AppVersion">v1.0.0</sys:String>- Stockage des validations dans Vault

<sys:String x:Key="AppReleaseDate">R 2026-01-10</sys:String>- Interface WebView2 pour affichage HTML

<sys:String x:Key="AppAuthor">Mohammed Amine Elgalai</sys:String>

<sys:String x:Key="AppCompany">XNRGY CLIMATE SYSTEMS ULC</sys:String>### 6. Connexions Automatiques

```

- **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique

#### Titres de Fenêtres (22 ressources)- **Inventor Professional 2026.2** - COM avec détection d'instance active

```xml- **Throttling intelligent** pour éviter spam logs (v1.1)

<!-- Utilisation dans XAML: Title="{StaticResource WindowTitleXXX}" -->- **Vérification fenêtre Inventor** prête avant connexion COM

<sys:String x:Key="WindowTitleMain">XNRGY Engineering Automation Tools - v1.0.0 - R 2026-01-10 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC</sys:String>- **Update Workspace** - Synchronisation dossiers au démarrage :

<sys:String x:Key="WindowTitleUploadModule">Upload Module - XEAT - v1.0.0 - ...</sys:String>  - `$/Content Center Files`

<sys:String x:Key="WindowTitleSmartTools">Smart Tools - XEAT - v1.0.0 - ...</sys:String>  - `$/Engineering/Inventor_Standards`

<!-- ... 19 autres titres -->  - `$/Engineering/Library/Cabinet`

```  - `$/Engineering/Library/Xnrgy_M99`

  - `$/Engineering/Library/Xnrgy_Module`

#### Barres Copyright (16 ressources)

```xml---

<!-- Utilisation dans XAML: Text="{StaticResource CopyrightXXX}" -->

<sys:String x:Key="CopyrightMain">XNRGY Engineering Automation Tools - v1.0.0 - R 2026-01-10 - By Mohammed Amine Elgalai - XNRGY CLIMATE SYSTEMS ULC</sys:String>## Services Principaux

<sys:String x:Key="CopyrightSmartTools">Smart Tools - XEAT - v1.0.0 - ...</sys:String>

<!-- ... 14 autres copyrights -->### VaultSDKService.cs

```

Service principal pour l'interaction avec Vault SDK (~1200 lignes).

---

**Responsabilités :**

## Configuration- Connexion/déconnexion Vault

- Chargement des Property Definitions

### Configuration Vault via fichier centralisé- Chargement des Catégories

- Chargement des Lifecycle Definitions

**Chemin Vault:**- Upload de fichiers avec `FileManager.AddFile`

```- Application des propriétés via `UpdateFileProperties`

$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/- Synchronisation Vault -> iProperties via `IExplorerUtil.UpdateFileProperties`

  XnrgyEngineeringAutomationToolsApp/XnrgyEngineeringAutomationToolsSettings.config- Assignation de catégories via `UpdateFileCategories`

```- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (reflection)

- Assignation de révisions via `UpdateFileRevisionNumbers`

**Chemin Local:**- Gestion des erreurs Vault (1003, 1013, 1136, etc.)

```

C:\Vault\Engineering\Inventor_Standards\Automation_Standard\Configuration_Files\### InventorService.cs

  XnrgyEngineeringAutomationToolsApp\XnrgyEngineeringAutomationToolsSettings.config

```Service pour la connexion COM à Inventor.



### Propriétés Vault (IDs fixes)**Améliorations v1.1 :**

- Throttling intelligent (minimum 2 sec entre tentatives)

| Propriété | ID | Description |- Vérification fenêtre Inventor prête (MainWindowHandle != IntPtr.Zero)

|-----------|-----|-------------|- Logs silencieux pour COMException 0x800401E3

| Project | 112 | Numéro de projet (ex: 10359) |- Compteur d'échecs consécutifs avec log périodique

| Reference | 121 | Numéro de référence (ex: 09) |

| Module | 122 | Numéro de module (ex: 03) |### InventorCopyDesignService.cs



### Pattern de chemin standardService pour Copy Design natif avec gestion des références.



```**Méthode principale :**

C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]```csharp

                              ↓         ↓       ↓Task<bool> ExecuteRealPackAndGoAsync(

Exemple:                   10359      REF09     M03    string templatePath,

```    string destinationPath,

    string projectNumber,

---    string reference,

    string module,

## Logs    IProgress<string> progress

)

| Module | Emplacement |```

|--------|-------------|

| Application | `bin\Release\Logs\XEAT_*.log` |### Services de Propriétés

| Vault Upload | `bin\Release\Logs\VaultUpload_*.log` |

| Copy Design | `bin\Release\Logs\CopyDesign_*.log` |- **InventorPropertyService** : iProperties via Inventor COM

| Smart Tools | `bin\Release\Logs\SmartTools_*.log` |- **ApprenticePropertyService** : Lecture iProperties via Apprentice (sans Inventor ouvert)

- **OlePropertyService** : Lecture OLE via OpenMCDF

### Format de log- **NativeOlePropertyService** : Lecture OLE native (Windows API)

```- **WindowsPropertyService** : Lecture propriétés Windows Shell

[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] [+] Message

```### Autres Services



---- **Logger.cs** : Logging NLog avec fichiers UTF-8

- **ThemeHelper.cs** : Gestion thèmes sombre/clair

## Règles de Développement- **UserPreferencesManager.cs** : Persistance préférences utilisateur

- **VaultSettingsService.cs** : Configuration chiffrée + sync Vault

### Emojis dans le Code- **JournalColorService.cs** : Couleurs uniformes pour journal

- **CredentialsManager.cs** : Gestion credentials chiffrées AES-256

#### INTERDITS (code C#, logs):

```---

😊 🙂 😄 👍 ❤️ 🔥 💯 ⭐ 🚀 etc.

```## Propriétés XNRGY



#### AUTORISÉS (interfaces XAML):Le système extrait automatiquement les propriétés depuis le chemin de fichier:

```

✅ ❌ ⚠️ ℹ️ 🔄 📄 📁 📊 🔍 📋 ⏳```

```C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]

                              |         |       |

#### Marqueurs ASCII pour logs:Vault Property IDs:        ID=112    ID=121  ID=122

| Emoji | Remplacement | Usage |```

|-------|--------------|-------|

| ❌ | [-] | Erreur || Propriété | ID Vault | Description |

| ✅ | [+] | Succès ||-----------|----------|-------------|

| ⚠️ | [!] | Avertissement || Project | 112 | Numéro de projet (5 chiffres) |

| 🔄 | [>] | Traitement || Reference | 121 | Numéro de référence (2 chiffres) |

| Module | 122 | Numéro de module (2 chiffres) |

### Build Tool

### Mapping Catégorie -> Lifecycle Definition

```powershell

# TOUJOURS utiliser:| Catégorie | Lifecycle Definition |

.\build-and-run.ps1|-----------|---------------------|

| Engineering | Flexible Release Process |

# JAMAIS utiliser:| Office | Simple Release Process |

dotnet build  # Ne génère pas les .g.cs pour WPF| Standard | Basic Release Process |

```| Base | (aucun) |



------



## Changelog Récent## Interface Utilisateur



### v1.0.0 (R 2026-01-10)### Thèmes



#### Centralisation UIL'application supporte deux thèmes avec propagation automatique vers toutes les fenêtres :

- [+] Toutes les métadonnées centralisées dans `XnrgyStyles.xaml`

- [+] 22 titres de fenêtres via `StaticResource`- **Thème Sombre** (défaut) : Fond #1E1E2E, panneaux #252536

- [+] 16 textes copyright via `StaticResource`- **Thème Clair** : Fond #F5F7FA, panneaux #FCFDFF

- [+] Grand titre MainWindow centralisé

- [+] Suppression de tous les `MultiBinding` redondants**Éléments à fond FIXE** (même en thème clair) :

- Journal : #1A1A28 (noir)

#### Smart Tools- Panneaux statistiques : #1A1A28 (noir)

- [+] 25+ outils Inventor complets- Headers GroupBox : #2A4A6F (bleu marine) avec texte blanc

- [+] Service SmartToolsService (~5700 lignes)

- [+] Fenêtres: ConstraintReport, CustomPropertyBatch, IProperties, etc.### Styles Centralisés



#### Place EquipmentLe fichier `Styles/XnrgyStyles.xaml` contient 50+ styles partagés :

- [+] Module complet avec Copy Design automatisé

- [+] PlaceEquipmentWindow (~4200 lignes)- **Couleurs** : Palette XNRGY officielle

- [+] Services: EquipmentPlacementService, EquipmentCopyDesignService- **Boutons** : PrimaryButton, SecondaryButton, SuccessButton, WarningButton, DangerButton avec effets glow

- **Inputs** : ModernTextBox, ModernComboBox avec bordures #4A7FBF

#### DXF Verifier- **DataGrid** : Headers fond sombre #1A1A28, texte bleu #0078D4

- [+] Migration complète vers XEAT- **Containers** : XnrgyGroupBox avec header bleu marine

- [+] Validation DXF/CSV vs PDF (~97% précision)- **ProgressBar** : Style personnalisé avec fill et shine

- [+] Version spécifique: v1.2.0- **Labels** : ModernLabel avec tailles et poids uniformisés



#### Architecture### Effets Visuels

- [+] 9 modules fonctionnels complets

- [+] Structure modulaire Modules/{Module}/Views|Services|Models- **Glow Effects** : DropShadowEffect bleu brillant (#00D4FF) sur hover des boutons

- [+] Styles centralisés (1350+ lignes, 12 sections)- **Animations** : Scale animations sur press des boutons

- **Bordures** : #4A7FBF (bleu brillant) sur tous les éléments interactifs

---

---

## Auteur

## Prérequis

**Mohammed Amine Elgalai**  

XNRGY Climate Systems ULC  - **Windows 10/11 x64**

2025-2026- **.NET Framework 4.8**

- **Autodesk Vault Professional 2026** (SDK v31.0.84)

---- **Autodesk Inventor Professional 2026.2**

- **Visual Studio 2022** (pour compilation)

## Licence- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)



Propriétaire - XNRGY Climate Systems ULC. Tous droits réservés.---


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

**Emojis dans les interfaces (XAML) :**
- ✅ = Succès/Sélection
- ❌ = Erreur/Annulation
- ⚠️ = Avertissement
- ℹ️ = Information
- ❓ = Question
- 🔄 = Statut/Refresh
- 📄 = Fichier
- 📐 = Extension
- 📊 = Taille
- 📁 = Chemin/Dossier
- ⏸️ = Pause/Ignoré

**Emojis INTERDITS (ne pas utiliser) :**
😊 🙂 😄 😁 😆 😂 🤣 🥲 😅 😍 🥰 😢 😭 😔 😒 😠 😡 😎 🥳 🤗 🤔 🙄 😴 📢 📣 🗣️ 🧠 🤖 🔥 💯 ⭐ ✨ 🌟 🎉 🎊 👍 👎 👏 🤝 ✌️ 🙌 ❤️ 🧡 💛 💚 💙 💜 🖤 💞 💓 💗 💕 💖 💻 🖥️ 📱 📷 🎧 🕹️ 🍎 🍌 🍕 🍔 🍟 🍿 🍣 🍩 🍪 🍫 ⏰ 🌧️ ☀️ ⛅ 🌙 📍🚀

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

## Services Principaux - Détails Techniques

### VaultSDKService.cs (~3000 lignes)

**Méthodes principales :**

#### Connexion et Authentification
- `Connect(string server, string vaultName, string username, string password)` : Connexion Vault avec gestion d'erreurs
- `Disconnect()` : Déconnexion propre avec libération ressources
- `IsConnected` : Propriété booléenne pour vérifier l'état de connexion
- `Reconnect()` : Reconnexion automatique en cas de perte de connexion

#### Gestion des Propriétés
- `LoadPropertyDefinitions()` : Chargement de toutes les Property Definitions au démarrage
- `GetPropertyDefinitionId(string propertyName)` : Récupération ID d'une propriété par nom
- `UpdateFileProperties(long fileId, Dictionary<string, object> properties)` : Application propriétés UDP
- `GetFileProperties(long fileId)` : Lecture propriétés d'un fichier
- **Propriétés XNRGY** : Project (ID=112), Reference (ID=121), Module (ID=122)

#### Gestion des Catégories
- `GetAvailableCategories()` : Liste toutes les catégories Vault disponibles
- `GetCategoryIdByName(string categoryName)` : Récupération ID catégorie
- `UpdateFileCategories(long fileId, long categoryId)` : Assignation catégorie à un fichier
- **Mapping automatique** : Catégorie → Lifecycle Definition (Engineering → Flexible Release Process)

#### Gestion du Lifecycle
- `GetAvailableLifecycleDefinitions()` : Liste toutes les Lifecycle Definitions
- `GetLifecycleDefinitionIdByCategory(string categoryName)` : Mapping catégorie → lifecycle
- `GetWorkInProgressStateId(long lifecycleDefinitionId)` : Récupération état "Work In Progress"
- `UpdateFileLifeCycleDefinitions(long fileId, long lifecycleDefinitionId, long lifecycleStateId)` : Assignation lifecycle (via reflection pour compatibilité SDK)

#### Gestion des Révisions
- `UpdateFileRevisionNumbers(long fileId, string revision)` : Assignation numéro de révision
- Support des formats de révision standards (A, B, C, 1, 2, 3, etc.)

#### Upload de Fichiers
- `UploadFile(string filePath, string vaultFolderPath, ...)` : Upload fichier avec propriétés complètes
- **Workflow complet** :
  1. Vérification existence fichier dans Vault
  2. CheckOut si fichier existe (nécessaire pour UpdateFileProperties)
  3. Upload via `FileManager.AddFile` (nouveaux fichiers) ou `FileManager.CheckoutFile` (existants)
  4. Application propriétés UDP via `UpdateFileProperties`
  5. **Synchronisation Vault → iProperties** via `IExplorerUtil.UpdateFileProperties` (fichiers Inventor uniquement)
  6. Assignation catégorie, lifecycle, révision
  7. CheckIn si fichier existait
  8. GET final pour mettre à jour le statut dans Vault Client

#### Synchronisation Vault → iProperties (Inventor)
- **IExplorerUtil** : Chargement lazy via `ExplorerLoader.LoadExplorerUtil`
- **Writeback automatique** : Les propriétés UDP Vault sont synchronisées vers les iProperties Inventor
- **Prérequis** : Writeback activé dans Vault (`GetEnableItemPropertyWritebackToFiles` doit retourner `true`)
- **Avantage** : Pas besoin d'ouvrir Inventor pour modifier les iProperties, Vault le fait automatiquement

#### Gestion des Erreurs Vault
- **Erreur 1003** : Job Processor actif (normal pour nouveaux fichiers) → Retour immédiat
- **Erreur 1013** : Fichier verrouillé → CheckOut automatique puis retry
- **Erreur 1136** : Fichier déjà existe → CheckOut puis UpdateFileProperties
- **Erreur 1001** : Permission insuffisante → Log et skip
- **Retry automatique** : 3 tentatives avec délai exponentiel

#### Opérations sur Dossiers
- `GetFolderAsync(string vaultFolderPath)` : Téléchargement dossier complet (GET)
- `CreateFolder(string parentPath, string folderName)` : Création dossier
- `FolderExists(string vaultFolderPath)` : Vérification existence dossier
- **Update Workspace** : Synchronisation automatique de 5 dossiers au démarrage

### InventorService.cs

**Fonctionnalités :**
- **Connexion COM** : Détection instance Inventor active ou démarrage invisible
- **Throttling intelligent** : Minimum 2 secondes entre tentatives de connexion
- **Vérification fenêtre** : Attente que `MainWindowHandle != IntPtr.Zero` avant connexion
- **Méthodes multiples** : `Marshal.GetActiveObject` puis P/Invoke `GetActiveObject` en fallback
- **Logs silencieux** : COMException 0x800401E3 (ROT non prêt) loggée en DEBUG uniquement
- **Compteur échecs** : Log périodique toutes les 5 tentatives pour éviter spam
- **Timer reconnexion** : Tentative automatique toutes les 3 secondes si Inventor non connecté

**Méthodes principales :**
- `TryConnect()` : Tentative connexion avec throttling
- `TryConnectViaMarshall()` : Méthode 1 - Marshal.GetActiveObject
- `TryConnectViaPInvoke()` : Méthode 2 - P/Invoke GetActiveObject
- `IsConnected` : Propriété booléenne
- `GetApplication()` : Récupération instance Inventor.Application

### InventorCopyDesignService.cs (~2700 lignes)

**Workflow Copy Design complet :**

1. **Préparation** :
   - Switch vers projet source (IPJ)
   - Ouverture Top Assembly (Module_.iam) en mode invisible
   - Application iProperties sur le template (Project, Reference, Module)

2. **Collecte références** :
   - Scan bottom-up de toutes les références (IPT, IAM, IDW)
   - Détection fichiers Library vs Module
   - Identification fichiers orphelins (non-référencés)

3. **Copy Design natif** :
   - `SaveAs` pour chaque fichier (IPT → IAM → Top Assembly)
   - Préservation liens Library (IPT_Typical_Drawing)
   - Copie fichiers Module avec références mises à jour

4. **Traitement dessins (.idw)** :
   - Mise à jour références via `PutLogicalFileNameUsingFull`
   - Correction chemins relatifs
   - **Composants suppressed** : Références mises à jour même si supprimés (v1.1)

5. **Fichiers orphelins** :
   - Copie des 1059 fichiers non-référencés
   - Préservation structure dossiers

6. **Fichiers non-Inventor** :
   - Copie Excel, PDF, Word, etc. (option "Inclure fichiers non-Inventor")

7. **Finalisation** :
   - Renommage fichier .ipj
   - Switch vers nouveau projet
   - Application iProperties finales
   - Paramètres Inventor (Design View "Default", Workfeatures cachés)
   - Vue ISO + Zoom All + Update All + Save All
- Module reste ouvert pour le dessinateur

**Méthodes principales :**
- `ExecuteRealPackAndGoAsync(...)` : Méthode principale orchestrant tout le workflow
- `CollectAllReferences(...)` : Collecte récursive des références
- `CopyFileWithReferences(...)` : Copie fichier avec mise à jour références
- `UpdateDrawingReferences(...)` : Mise à jour références dans les dessins
- `ApplyIProperties(...)` : Application iProperties via InventorPropertyService

### Services de Propriétés (5 implémentations)

#### 1. InventorPropertyService.cs
- **Méthode** : Inventor COM API (`Application.Documents.Open`)
- **Performance** : 3-15 secondes par fichier
- **Usage** : Modification iProperties AVANT upload vers Vault
- **Fonctionnalités** :
  - `SetIProperties(filePath, projectNumber, reference, module)` : Propriétés de base
  - `SetAllModuleProperties(...)` : Toutes propriétés XNRGY (Project, Reference, Module, Initiale_du_Dessinateur, Initiale_du_Co_Dessinateur, Creation_Date, Numero_de_Projet)
  - `SetOrCreateProperty(...)` : Création propriété si n'existe pas
- **Mode invisible** : Documents ouverts sans fenêtre visible
- **Auto-close** : Fermeture automatique après modification

#### 2. ApprenticePropertyService.cs
- **Méthode** : Autodesk Inventor Apprentice API
- **Performance** : ~1-2 secondes par fichier
- **Usage** : Lecture iProperties SANS ouvrir Inventor
- **Avantage** : Pas besoin d'Inventor installé (Apprentice suffit)

#### 3. OlePropertyService.cs
- **Méthode** : OpenMCDF (NuGet package)
- **Performance** : ~100-200ms par fichier
- **Usage** : Lecture propriétés OLE Compound Documents
- **Limitation** : Lecture seule, pas de modification

#### 4. NativeOlePropertyService.cs (~700 lignes)
- **Méthode** : Windows API native (P/Invoke ole32.dll)
- **Performance** : ~50-100ms par fichier (LE PLUS RAPIDE)
- **Usage** : Modification propriétés OLE haute performance
- **Fonctionnalités** :
  - `StgOpenStorageEx` : Ouverture fichier OLE
  - `IPropertySetStorage` : Accès Property Sets
  - `IPropertyStorage` : Lecture/écriture propriétés
  - Support FMTID_UserDefinedProperties (GUID standard OLE)
- **Avantage** : 10-30x plus rapide que Inventor COM API

#### 5. WindowsPropertyService.cs
- **Méthode** : Windows Shell API
- **Performance** : ~200-500ms par fichier
- **Usage** : Lecture propriétés Windows standard (Title, Author, Subject, etc.)
- **Limitation** : Propriétés Windows uniquement, pas iProperties Inventor

### VaultSettingsService.cs

**Fonctionnalités :**
- **Chiffrement AES-256** : Tous les fichiers de configuration sont chiffrés
- **Synchronisation Vault** : Téléchargement automatique au démarrage depuis `$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/`
- **Accès restreint** : Vérification rôle "Administrator" ou groupe "Admin_Designer"
- **Déploiement multi-sites** : Saint-Hubert QC + Arizona US (50+ utilisateurs)
- **Cache local** : Fichiers chiffrés stockés localement pour accès hors ligne

**Méthodes principales :**
- `LoadSettingsFromVault()` : Téléchargement depuis Vault
- `SaveSettingsToVault(settings)` : Upload vers Vault (Admin uniquement)
- `LoadSettingsFromLocal()` : Chargement depuis cache local
- `EncryptSettings(settings)` : Chiffrement AES-256
- `DecryptSettings(encryptedData)` : Déchiffrement

### UserPreferencesManager.cs

**Gestion préférences utilisateur :**
- **Persistance locale** : Fichier JSON chiffré dans `AppData\Local\XnrgyEngineeringAutomationTools\`
- **Préférences stockées** :
  - `IsDarkTheme` : Thème sombre/clair (défaut: sombre)
  - `IsMaximized` : Fenêtre maximisée au démarrage
  - `AutoConnectVault` : Connexion automatique Vault (défaut: true)
  - `AutoConnectInventor` : Connexion automatique Inventor (défaut: true)
  - `ShowStartupChecklist` : Afficher checklist au démarrage
  - `AppVersion` : Version application
  - `LastUser` : Dernier utilisateur connecté

**Méthodes principales :**
- `Load()` : Chargement préférences
- `Save(preferences)` : Sauvegarde préférences
- `SaveTheme(isDarkTheme)` : Sauvegarde thème uniquement
- `LoadTheme()` : Chargement thème uniquement
- `Reset()` : Réinitialisation aux valeurs par défaut

### CredentialsManager.cs

**Gestion credentials chiffrées :**
- **Chiffrement AES-256** : Mots de passe stockés chiffrés
- **Emplacement** : `AppData\Local\XnrgyEngineeringAutomationTools\credentials.encrypted`
- **Sécurité** : Clé dérivée depuis machine ID + user SID

**Méthodes principales :**
- `SaveCredentials(server, vault, username, password)` : Sauvegarde chiffrée
- `LoadCredentials()` : Chargement et déchiffrement
- `ClearCredentials()` : Suppression credentials

### Logger.cs

**Système de logging NLog :**
- **Format** : `[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] Message`
- **Niveaux** : TRACE, DEBUG, INFO, WARN, ERROR, FATAL
- **Fichiers** : Rotation quotidienne dans `bin\Release\Logs\`
- **Encodage** : UTF-8 pour support caractères spéciaux
- **Icônes textuelles** : `[+]`, `[-]`, `[!]`, `[>]`, `[i]`, `[~]`, `[#]`, `[?]`

**Méthodes principales :**
- `Log(string message, LogLevel level)` : Log message
- `LogException(string context, Exception ex, LogLevel level)` : Log exception avec stack trace

### JournalColorService.cs

**Couleurs uniformes pour journaux UI :**
- **Palette standard XNRGY** :
  - SUCCESS : #00FF7F (vert SpringGreen brillant)
  - ERROR : #FF4444 (rouge vif)
  - WARNING : #FFD700 (jaune or brillant)
  - INFO : #FFFFFF (blanc pur)
  - DEBUG : #00FFFF (cyan brillant)
  - TIMESTAMP : #888888 (gris moyen)
- **Brushes pré-créés** : Performance optimisée (singleton)
- **Méthodes utilitaires** :
  - `GetBrushForLevel(LogLevel)` : Brush selon niveau
  - `GetColorForLevel(LogLevel)` : Couleur selon niveau
  - `GetHexColorForLevel(LogLevel)` : Code hex selon niveau
  - `GetPrefixForLevel(LogLevel)` : Préfixe standard (`[+]`, `[-]`, etc.)

### ThemeHelper.cs

**Gestion thèmes sombre/clair :**
- **Couleurs thème sombre** :
  - Background : #1E1E2E
  - Panel : #252536
  - Input : #2D2D44
  - Border : #404060
- **Couleurs thème clair** :
  - Background : #F5F7FA
  - Panel : #FCFDFF
  - Input : #F0F5FC
  - Border : #C8D2E1
- **Couleurs fixes** (ne changent pas) :
  - Bleu Marine : #2A4A6F (headers GroupBox)
  - StatusBar : #1A1A28 (journal, panneaux stats)
- **Méthodes utilitaires** :
  - `ApplyThemeToWindow(Window)` : Application thème à une fenêtre
  - `ApplyThemeToGroupBox(GroupBox)` : Application thème à un GroupBox
  - `ApplyThemeToTextBox(TextBox)` : Application thème à un TextBox
  - `ApplyThemeToDataGrid(DataGrid)` : Application thème à un DataGrid

### SettingsService.cs

**Gestion paramètres application :**
- **Fichier** : `modulesettings.json` dans le répertoire de l'application
- **Format** : JSON avec camelCase
- **Structure** : `ModuleSettings` contient `CreateModuleSettings`
- **Cache** : Singleton avec lock thread-safe
- **Méthodes principales** :
  - `Load()` : Chargement depuis fichier
  - `Save(settings)` : Sauvegarde vers fichier
  - `Reload()` : Rechargement depuis fichier
  - `ResetToDefaults()` : Réinitialisation valeurs par défaut

## Modèles de Données

### VaultUploadFileItem.cs
- **Propriétés** :
  - `FileName`, `FilePath`, `FileExtension`
  - `ProjectNumber`, `Reference`, `Module`
  - `IsSelected` (bool)
  - `Status` (string)
  - `CategoryId`, `LifecycleDefinitionId`, `LifecycleStateId`
  - `Revision`
- **Usage** : Item pour DataGrid dans UploadModuleWindow

### CreateModuleRequest.cs
- **Propriétés** :
  - `SourceType` (Template ou ExistingProject)
  - `TemplatePath`, `ExistingProjectPath`
  - `DestinationPath`
  - `ProjectNumber`, `Reference`, `Module`
  - `FullProjectNumber` (25001REF1M1)
- **Usage** : Paramètres pour Copy Design

### CreateModuleSettings.cs
- **Propriétés** :
  - `DefaultTemplatePath`
  - `DefaultDestinationPath`
  - `DefaultProjectNumber`
  - `IncludeNonInventorFiles` (bool)
  - `RenameOptions` (Rechercher/Remplacer, Préfixe/Suffixe)
- **Usage** : Paramètres sauvegardés pour Créer Module

## Fenêtres et Interfaces

### MainWindow.xaml(.cs)
**Dashboard principal (hub) :**
- **Fonctionnalités** :
  - Connexion Vault/Inventor centralisée
  - Boutons modules (Upload Module, Créer Module, Upload Template, Checklist HVAC, DXF Verifier, Update Workspace)
  - Journal des opérations avec couleurs uniformes
  - Panneaux statistiques (fond noir fixe)
  - Thème sombre/clair avec propagation automatique
  - Barre de statut avec informations connexion
- **Méthodes principales** :
  - `ConnectVault_Click()` : Connexion Vault
  - `ConnectInventor_Click()` : Connexion Inventor
  - `OpenUploadModule_Click()` : Ouverture Upload Module
  - `OpenCreateModule_Click()` : Ouverture Créer Module
  - `ApplyTheme(bool isDark)` : Application thème à toutes les sous-fenêtres
  - `UpdateWorkspace_Click()` : Synchronisation dossiers Vault

### UploadModuleWindow.xaml(.cs) (~1200 lignes)
**Module upload Vault :**
- **Fonctionnalités** :
  - Scan automatique modules engineering
  - Extraction propriétés depuis chemin (Project/Ref/Module)
  - Deux DataGrids séparés (Inventor/Non-Inventor)
  - Sélection multiple avec checkboxes
  - Filtres (recherche texte, extension, statut)
  - Barre de progression avec glow brillant (#00FF7F)
  - Journal des opérations avec couleurs
  - Contrôles Pause/Stop/Annuler
  - Statistiques (total, sélectionnés, uploadés, erreurs)
- **Workflow upload** :
  1. Scan dossier sélectionné
  2. Extraction propriétés depuis chemin
  3. Séparation Inventor/Non-Inventor
  4. Upload batch avec gestion erreurs
  5. Application propriétés UDP
  6. Synchronisation Vault → iProperties (Inventor)
  7. Assignation catégorie, lifecycle, révision
- **Méthodes principales** :
  - `SelectModule_Click()` : Sélection dossier module
  - `ScanModule_Click()` : Scan automatique
  - `UploadSelected_Click()` : Upload fichiers sélectionnés
  - `UpdateProgress(int current, int total)` : Mise à jour barre progression
  - `UpdateStatistics()` : Mise à jour statistiques

### CreateModuleWindow.xaml(.cs) (~1800 lignes)
**Module Copy Design :**
- **Fonctionnalités** :
  - Sélection source (Template ou Projet Existant)
  - Champs Project/Reference/Module avec validation
  - Prévisualisation fichiers avant copie
  - DataGrid avec fichiers et références
  - Barre de progression avec glow brillant
  - Journal détaillé des opérations
  - Boutons Preview/Cancel/Create Module
  - Fenêtre réglages (CreateModuleSettingsWindow)
- **Workflow Copy Design** : Voir section InventorCopyDesignService
- **Méthodes principales** :
  - `SelectTemplate_Click()` : Sélection template
  - `SelectExistingProject_Click()` : Sélection projet existant
  - `Preview_Click()` : Prévisualisation fichiers
  - `CreateModule_Click()` : Démarrage Copy Design
  - `UpdateProgress(int current, int total, string message)` : Mise à jour progression

### UploadTemplateWindow.xaml(.cs) (~1100 lignes)
**Module upload templates (Admin) :**
- **Fonctionnalités** :
  - Vérification rôle administrateur
  - Scan templates depuis Library
  - Filtres (recherche, extension, statut)
  - Sélection multiple
  - Upload batch vers Vault
  - Journal et barre de progression
- **Méthodes principales** :
  - `ScanTemplates_Click()` : Scan templates
  - `UploadSelected_Click()` : Upload sélectionnés
  - `ApplyFilters()` : Application filtres

### CreateModuleSettingsWindow.xaml(.cs)
**Fenêtre réglages Créer Module :**
- **Fonctionnalités** :
  - Configuration chemins templates
  - Liste initiales designers (26 + "Autre...")
  - Options renommage (Rechercher/Remplacer, Préfixe/Suffixe)
  - Checkbox "Inclure fichiers non-Inventor"
  - Styles uniformisés avec effets glow
  - Titres GroupBox orange (#FF8C00)
- **Méthodes principales** :
  - `SaveSettings_Click()` : Sauvegarde paramètres
  - `LoadSettings()` : Chargement paramètres

### ChecklistHVACWindow.xaml(.cs)
**Module validation HVAC :**
- **Fonctionnalités** :
  - Checklist interactive avec critères XNRGY
  - Validation modules AHU
  - Stockage validations dans Vault
  - Interface WebView2 pour affichage HTML

### LoginWindow.xaml(.cs)
**Fenêtre connexion Vault :**
- **Fonctionnalités** :
  - Champs serveur, vault, utilisateur, mot de passe
  - Checkbox "Se souvenir" (CredentialsManager)
  - Validation connexion
  - Styles uniformisés

### ModuleSelectionWindow.xaml(.cs)
**Fenêtre sélection module :**
- **Fonctionnalités** :
  - Liste modules disponibles
  - Description chaque module
  - Navigation vers module sélectionné

### PreviewWindow.xaml(.cs)
**Fenêtre prévisualisation :**
- **Fonctionnalités** :
  - Affichage liste fichiers
  - Informations détaillées (chemin, taille, type)
  - Bouton fermer

### XnrgyMessageBox.xaml(.cs)
**MessageBox moderne XNRGY :**
- **Types** : Info, Success, Warning, Error, Question
- **Boutons** : OK, OKCancel, YesNo, YesNoCancel
- **Styles** : Thème XNRGY avec icônes colorées
- **Méthodes statiques** :
  - `Show(message, title, type, buttons, owner)`
  - `ShowSuccess(message, title, owner)`
  - `ShowError(message, title, owner)`
  - `ShowInfo(message, title, owner)`
  - `ShowWarning(message, title, owner)`
  - `Confirm(message, title, owner)` : Retourne bool

## Gestion d'Erreurs et Retry

### Stratégies de Retry
- **Vault Upload** : 3 tentatives avec délai exponentiel (1s, 2s, 4s)
- **Inventor COM** : Timer reconnexion toutes les 3 secondes avec throttling (min 2s)
- **Vault Settings** : Retry automatique en cas d'échec synchronisation

### Gestion Erreurs Vault
- **1003** : Job Processor actif → Retour immédiat (normal)
- **1013** : Fichier verrouillé → CheckOut automatique puis retry
- **1136** : Fichier existe → CheckOut puis UpdateFileProperties
- **1001** : Permission insuffisante → Log et skip
- **Timeout** : 30 secondes par défaut

### Gestion Erreurs Inventor
- **COMException 0x800401E3** : ROT non prêt → Log DEBUG silencieux, retry automatique
- **Fenêtre non prête** : Attente `MainWindowHandle != IntPtr.Zero`
- **Instance non trouvée** : Démarrage instance invisible

## Structure Shared/

### Composants Partagés

Le dossier `Shared/` contient tous les composants réutilisables entre modules :

#### Views/
- **LoginWindow.xaml(.cs)** : Fenêtre connexion Vault réutilisable
- **ModuleSelectionWindow.xaml(.cs)** : Sélection module avec navigation
- **PreviewWindow.xaml(.cs)** : Prévisualisation fichiers/listes
- **XnrgyMessageBox.xaml(.cs)** : MessageBox moderne avec thème XNRGY

#### Models/ (vide actuellement)
- Réservé pour modèles de données partagés

#### Services/ (vide actuellement)
- Réservé pour services partagés entre modules

## Outils et Scripts

### build-and-run.ps1
**Script PowerShell compilation automatique :**
- **Détection MSBuild** : VS 2022 Enterprise/Professional/Community
- **Compilation Release/Debug** : Mode configurable
- **Arrêt instances** : `taskkill /F` automatique
- **Lancement automatique** : Après compilation réussie
- **Options** :
  - `-BuildOnly` : Compilation sans lancer
  - `-Debug` : Mode Debug
  - `-Clean` : Clean + Build
  - `-KillOnly` : Tuer instances uniquement

### build-and-run.bat
**Wrapper batch pour PowerShell :**
- Appelle `build-and-run.ps1` avec paramètres

### Scripts PowerShell (Scripts/)
- **CleanInventor2023Registry.ps1** : Nettoyage registre Inventor 2023
- **Prepare-TemplateFiles.ps1** : Préparation fichiers templates
- **Upload-ToVaultProd.ps1** : Upload vers Vault Production

### Tools/VaultBulkUploader/
**Outil console upload massif :**
- **Usage** : Upload batch de milliers de fichiers
- **Performance** : 6152 fichiers uploadés vers PROD_XNGRY
- **Fonctionnalités** :
  - Scan récursif dossiers
  - Upload parallèle avec gestion erreurs
  - Journal détaillé
  - Statistiques finales

## Versions en Développement (Non publiées)

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
- Effets glow brillants sur toutes les barres de progression (#00FF7F, BlurRadius=20, Opacity=0.85)

---

## Firebase Remote Control System

L'application integre **Firebase Realtime Database** pour le controle a distance complet:

### Fonctionnalites
- **Kill Switch**: Desactivation globale instantanee de l'application
- **Maintenance Mode**: Message de maintenance avec compte a rebours
- **Force Update**: Mise a jour obligatoire avec blocage optionnel
- **Device Tracking**: Suivi temps reel des postes (CPU, RAM, heartbeat 30s)
- **Broadcasts**: Messages push vers les utilisateurs (info/warning/error)
- **User/Device Blocking**: Blocage granulaire par utilisateur ou poste

### Firebase URL
```
https://xeat-remote-control-default-rtdb.firebaseio.com
```

### Admin Panel
- **Chemin**: `Firebase Realtime Database configuration/admin-panel/index.html`
- **Auth**: Firebase Authentication (admin@xnrgy.com)
- **Dashboard**: Stats temps reel, postes en ligne, activite
- **Controles**: Kill Switch, Maintenance, Force Update
- **Gestion**: Utilisateurs, Postes (details hardware), Sites
- **Communication**: Broadcasts, Audit logs

---

## Custom Installer

Installateur WPF personnalise multi-etapes remplacant les installateurs standards.

### Structure
```
Installer/
 XnrgyInstaller.csproj
 InstallerWindow.xaml(.cs)
 InstallerService.cs
 Build-Installer.ps1
```

### Build Commands
```powershell
.\build-and-run.ps1 -WithInstaller       # App + Installer
.\build-and-run.ps1 -InstallerOnly       # Installer seul
.\build-and-run.ps1 -WithInstaller -CreatePackage  # + ZIP distribution
```

### Pages Wizard
1. Welcome - Logo et version
2. License - Acceptation EULA
3. Destination - Choix dossier + espace requis
4. Progress - Installation avec log
5. Complete - Succes/Echec + Lancer app

---

*Derniere mise a jour : 17 Janvier 2026*
