# XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# 🏭 XNRGY Engineering Automation Tools# 🏭 XNRGY Engineering Automation Tools# XNRGY Engineering Automation Tools# VaultAutomationTool



> **Suite d'outils d'automatisation engineering unifiee** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

>

> Developpe par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2



--->



## Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2



**XNRGY Engineering Automation Tools** est une application hub centralisee (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering developpes pour XNRGY Climate Systems. Cette suite vise a simplifier et accelerer les workflows des equipes de design en integrant la gestion Vault, les manipulations Inventor, et les validations qualite dans une interface unifiee.



### Objectif Principal--->



Remplacer les multiples applications standalone par une **plateforme unique** avec :

- Connexion centralisee a Vault & Inventor

- Interface utilisateur moderne et coherente (themes sombre/clair)## Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

- Partage de services communs (logging, configuration chiffree AES-256)

- Deploiement multi-sites et maintenance simplifies

- Parametres centralises via Vault (50+ utilisateurs, 3 sites)

**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

---



## Modules Integres

### Objectif Principal--->

| Module | Description | Statut |

|--------|-------------|--------|

| **Upload Module** | Upload automatise vers Vault avec proprietes (Project/Ref/Module) | [+] 100% |

| **Creer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |Remplacer les multiples applications standalone par une **plateforme unique** avec :

| **Reglages Admin** | Configuration centralisee et synchronisee via Vault (AES-256) | [+] 100% |

| **Upload Template** | Upload templates vers Vault (reserve Admin) | [+] 100% |- Connexion centralisée à Vault & Inventor

| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |

| **Smart Tools** | Creation IPT/STEP, generation PDF, iLogic Forms | [~] Planifie |- Interface utilisateur moderne et cohérente (thèmes sombre/clair)## Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

| **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |

| **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |- Partage de services communs (logging, configuration chiffrée AES-256)

| **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifie |

- Déploiement multi-sites et maintenance simplifiés

---

- Paramètres centralisés via Vault (50+ utilisateurs, 3 sites)

## Fonctionnalites Implementees

**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

### 1. Upload Module (100%) - NOUVEAU v1.1

---

Module integre (ex-VaultAutomationTool) pour l'upload de fichiers vers Vault:



- **Connexion centralisee** - Utilise la connexion Vault de l'app principale

- **Scan automatique** des modules engineering avec extraction proprietes## Modules Intégrés

- **Separation Inventor/Non-Inventor** dans deux DataGrids avec headers visibles

- **Application automatique** des proprietes metier:### Objectif Principal--->

  - Project (ID=112)

  - Reference (ID=121)| Module | Description | Statut |

  - Module (ID=122)

- **Assignation complete**:|--------|-------------|--------|

  - Categories Vault

  - Lifecycle Definitions et States| **Upload Module** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | [+] 100% |

  - Revisions

- **Synchronisation Vault vers iProperties** via `IExplorerUtil`| **Créer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |Remplacer les multiples applications standalone par une **plateforme unique** avec :

- **Journal des operations** avec barre de progression style Creer Module

- **Controles**: Pause/Stop/Annuler pendant l'upload| **Réglages Admin** | Configuration centralisée et synchronisée via Vault (AES-256) | [+] 100% |

- **Styles DataGrid** avec headers fond sombre et texte bleu XNRGY

| **Upload Template** | Upload templates vers Vault (réservé Admin) | [+] 100% |- Connexion centralisée à Vault & Inventor

### 2. Creer Module - Copy Design (100%)

| **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | [~] Planifié |

**Sources disponibles :**

- Depuis Template : `$/Engineering/Library/Xnrgy_Module` (1083 fichiers Inventor)| **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |- Interface utilisateur moderne et cohérente (thème sombre)## Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

- Depuis Projet Existant : Selection d'un projet local ou Vault

| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [+] 100% |

**Workflow automatise :**

1. Switch vers projet source (IPJ)| **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |- Partage de services communs (logging, configuration chiffrée)

2. Ouverture Top Assembly (Module_.iam)

3. Application iProperties sur le template| **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifié |

4. Collecte de toutes les references (bottom-up)

5. Copy Design natif avec SaveAs (IPT -> IAM -> Top Assembly)- Déploiement multi-sites et maintenance simplifiés

6. Traitement des dessins (.idw) avec mise a jour des references

7. **Mise a jour des references des composants suppressed** (v1.1)---

8. Copie des fichiers orphelins (1059 fichiers non-references)

9. Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)- Paramètres centralisés via Vault (50+ utilisateurs, 3 sites)

10. Renommage du fichier .ipj

11. Switch vers le nouveau projet## Fonctionnalités Implémentées

12. Application des iProperties finales et parametres Inventor

13. Design View -> "Default", masquage Workfeatures**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems.

14. Vue ISO + Zoom All (Fit)

15. Update All (rebuild) + Save All### 1. Upload Module (100%) - NOUVEAU

16. Module reste ouvert pour le dessinateur

---

**Gestion intelligente des references :**

- Fichiers Library (IPT_Typical_Drawing) : Liens preservesModule intégré (ex-VaultAutomationTool) pour l'upload de fichiers vers Vault:

- Fichiers Module : Copies avec references mises a jour

- Fichiers IDW : References corrigees via `PutLogicalFileNameUsingFull`

- **Composants suppressed** : References mises a jour meme si supprimes

- **Connexion centralisée** - Utilise la connexion Vault de l'app principale

**Options de renommage (v1.1) :**

- Rechercher/Remplacer (cumulatif sur NewFileName)- **Scan automatique** des modules engineering avec extraction propriétés## Modules Intégrés

- Prefixe/Suffixe (applique sur OriginalFileName)

- **Checkbox "Inclure fichiers non-Inventor"**- **Séparation Inventor/Non-Inventor** dans deux DataGrids



### 3. Reglages Admin (100%)- **Application automatique** des propriétés métier:### Objectif Principal--->



**Systeme de configuration centralisee :**  - Project (ID=112)

- Chiffrement AES-256 des fichiers de configuration

- Synchronisation automatique via Vault au demarrage  - Reference (ID=121) | Module | Description | Statut |

- Acces restreint aux administrateurs (Role "Administrator" ou Groupe "Admin_Designer")

- Deploiement multi-sites : Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs  - Module (ID=122)



**Chemin Vault :**- **Assignation complète**:|--------|-------------|--------|

```

$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/  - Catégories Vault

```

  - Lifecycle Definitions et States| **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | [+] 100% |

**Sections configurables :**

- Liste des initiales designers (26 entrees + "Autre...")  - Révisions

- Chemins templates et projets

- Extensions Inventor supportees- **Synchronisation Vault vers iProperties** via `IExplorerUtil`| **Créer Module** | Copy Design natif depuis template Library ou projet existant | [+] 100% |Remplacer les multiples applications standalone par une **plateforme unique** avec :

- Dossiers/fichiers exclus

- Noms des iProperties- **Journal des opérations** avec barre de progression



### 4. Upload Template (100%)- **Contrôles**: Pause/Stop/Annuler pendant l'upload| **Réglages Admin** | Configuration centralisée et synchronisée via Vault (AES-256) | [+] 100% |



- **Reserve aux administrateurs** - Message XnrgyMessageBox si non-admin

- **Upload templates** depuis Library vers Vault

- **Utilise la connexion partagee** de l'app principale### 2. Créer Module - Copy Design (100%)| **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | [~] Planifié |- Connexion centralisée à Vault & Inventor

- **Journal integre** avec barre de progression



### 5. Checklist HVAC (100%)

**Sources disponibles :**| **DXF Verifier** | Validation des fichiers DXF avant envoi | [~] Migration |

- Validation des modules AHU

- Checklist interactive avec criteres XNRGY- Depuis Template : `$/Engineering/Library/Xnrgy_Module`

- Stockage des validations dans Vault

- Depuis Projet Existant : Sélection d'un projet local ou Vault| **Checklist HVAC** | Validation modules AHU avec stockage Vault | [~] Migration |- Interface utilisateur moderne et cohérente## Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC> **Suite d'outils d'automatisation engineering unifiée** pour Autodesk Vault Professional 2026 & Inventor Professional 2026.2

### 6. Connexions Automatiques



- **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique

- **Inventor Professional 2026.2** - COM avec detection d'instance active**Workflow automatisé :**| **Time Tracker** | Analyse temps de travail modules HVAC | [~] Migration |

- **Throttling intelligent** pour eviter spam logs (v1.1)

- **Verification fenetre Inventor** prete avant connexion COM1. Switch vers projet source (IPJ)

- **Update Workspace** - Synchronisation dossiers au demarrage :

  - `$/Content Center Files`2. Ouverture Top Assembly| **Update Workspace** | Synchronisation librairies depuis Vault | [~] Planifié |- Partage de services communs (logging, configuration, etc.)

  - `$/Engineering/Inventor_Standards`

  - `$/Engineering/Library/Cabinet`3. Application iProperties

  - `$/Engineering/Library/Xnrgy_M99`

  - `$/Engineering/Library/Xnrgy_Module`4. Collecte références (bottom-up)



---5. Copy Design natif avec SaveAs



## Architecture6. Traitement dessins (.idw) avec mise à jour références---- Déploiement et maintenance simplifiés



```7. Mise à jour références des composants suppressed

XnrgyEngineeringAutomationTools/

+-- App.xaml(.cs)                    # Point d'entree application8. Copie fichiers orphelins et non-Inventor

+-- MainWindow.xaml(.cs)             # Dashboard principal

+-- appsettings.json                 # Configuration sauvegardee9. Renommage fichier .ipj

|

+-- Models/                          # Modeles de donnees10. Switch vers nouveau projet## Fonctionnalités Implémentées

|   +-- ApplicationConfiguration.cs  # Configuration application

|   +-- CategoryItem.cs              # Item categorie pour ComboBox11. Application iProperties finales et paramètres Inventor

|   +-- FileItem.cs                  # Item fichier pour DataGrid

|   +-- FileToUpload.cs              # Fichier a uploader12. Module reste ouvert pour le dessinateur

|   +-- LifecycleDefinitionItem.cs   # Lifecycle Definition

|   +-- LifecycleStateItem.cs        # Lifecycle State

|   +-- ModuleInfo.cs                # Informations module

|   +-- ProjectInfo.cs               # Informations projet**Options de renommage :**### 1. Vault Upload (100%)---**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

|   +-- ProjectProperties.cs         # Proprietes Project/Ref/Module

|   +-- VaultConfiguration.cs        # Configuration Vault- Rechercher/Remplacer (cumulatif)

|   +-- CreateModuleRequest.cs       # Requete creation module

|- Préfixe/Suffixe

+-- Services/                        # Services metier

|   +-- VaultSdkService.cs           # SDK Vault v31.0.84- Checkbox "Inclure fichiers non-Inventor"

|   +-- VaultSettingsService.cs      # Config chiffree + sync Vault

|   +-- InventorService.cs           # Connexion Inventor COM- Connexion directe via SDK Vault v31.0.84

|   +-- InventorCopyDesignService.cs # Copy Design natif

|   +-- Logger.cs                    # Logging UTF-8### 3. Réglages Admin (100%)

|

+-- Views/                           # Fenetres et dialogues- Scan automatique des modules engineering

|   +-- LoginWindow.xaml(.cs)        # Connexion Vault

|   +-- CreateModuleWindow.xaml(.cs) # Creer Module- **Configuration centralisée** stockée dans Vault (`$/Admin/Config/app_settings.json`)

|   +-- CreateModuleSettingsWindow.xaml(.cs) # Reglages Admin

|   +-- UploadTemplateWindow.xaml(.cs)       # Upload Template- **Chiffrement AES-256** pour les données sensibles (mots de passe)- Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)## Modules Intégrés

|   +-- ChecklistHVACWindow.xaml(.cs)        # Checklist HVAC

|   +-- ModuleSelectionWindow.xaml(.cs)      # Selection module- **Interface graphique** avec validation temps réel

|   +-- PreviewWindow.xaml(.cs)              # Previsualisation

|   +-- XnrgyMessageBox.xaml(.cs)            # MessageBox moderne- **Synchronisation automatique** au démarrage si connecté à Vault- Application automatique des propriétés métier extraites du chemin

|

+-- Modules/                         # Modules integres- **Sections configurables**:

|   +-- VaultUpload/

|       +-- Models/  - Paramètres Vault (server, vault, credentials)- Assignation de catégories, lifecycle definitions/states et révisions

|       |   +-- VaultUploadFileItem.cs

|       |   +-- VaultUploadModels.cs  - Chemins par défaut (Library, Templates, Projects)

|       +-- Views/

|           +-- VaultUploadModuleWindow.xaml(.cs)  - Options Copy Design- Synchronisation Vault vers iProperties via `IExplorerUtil`

|

+-- ViewModels/                      # MVVM ViewModels  - Paramètres généraux

|   +-- AppMainViewModel.cs          # ViewModel principal

|   +-- RelayCommand.cs              # Implementation ICommand| Module | Description | Statut |### Objectif Principal--->

|

+-- Converters/                      # Convertisseurs WPF### 4. Upload Template (100%)

+-- Resources/                       # Images et icones

+-- Logs/                            # Fichiers logs### 2. Créer Module - Copy Design (100%)

+-- build-and-run.ps1                # Script compilation MSBuild

```- **Réservé aux administrateurs** - Message XnrgyMessageBox si non-admin



---- **Upload templates** depuis Library vers Vault|--------|-------------|--------|



## Proprietes XNRGY- **Utilise la connexion partagée** de l'app principale



Le systeme extrait automatiquement les proprietes depuis le chemin de fichier:- **Journal intégré** avec barre de progression**Sources disponibles :**



```

C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]

                              |         |       |### 5. Checklist HVAC (100%)- Depuis Template : `$/Engineering/Library/Xnrgy_Module`| **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ 100% |

Vault Property IDs:        ID=112    ID=121  ID=122

```



| Propriete | ID Vault | Description |- Validation des modules AHU- Depuis Projet Existant : Sélection d'un projet local ou Vault

|-----------|----------|-------------|

| Project | 112 | Numero de projet (5 chiffres) |- Checklist interactive avec critères XNRGY

| Reference | 121 | Numero de reference (2 chiffres) |

| Module | 122 | Numero de module (2 chiffres) |- Stockage des validations dans Vault| **Créer Module** | Copy Design natif depuis template Library ou projet existant | ✅ 100% |



### Mapping Categorie -> Lifecycle Definition



| Categorie | Lifecycle Definition |---**Workflow automatisé :**

|-----------|---------------------|

| Engineering | Flexible Release Process |

| Office | Simple Release Process |

| Standard | Basic Release Process |## Architecture1. Switch vers projet source (IPJ)| **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | Planifié |Remplacer les multiples applications standalone par une **plateforme unique** avec :

| Base | (aucun) |



---

```2. Ouverture Top Assembly

## Prerequis

XnrgyEngineeringAutomationTools/

- **Windows 10/11 x64**

- **.NET Framework 4.8**├── App.xaml(.cs)                    # Point d'entrée application3. Application iProperties| **DXF Verifier** | Validation des fichiers DXF avant envoi | Migration |

- **Autodesk Vault Professional 2026** (SDK v31.0.84)

- **Autodesk Inventor Professional 2026.2**├── MainWindow.xaml(.cs)             # Fenêtre principale hub

- **Visual Studio 2022** (pour compilation)

- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)├── Models/                          # Modèles de données4. Collecte références (bottom-up)



---│   ├── ModuleInfo.cs



## Compilation et Execution│   ├── FileEntry.cs5. Copy Design natif avec SaveAs| **Checklist HVAC** | Validation modules AHU avec stockage Vault | Migration |- Connexion centralisée à Vault & Inventor



### Script automatique (RECOMMANDE)│   ├── CopyDesignOptions.cs



```powershell│   └── AppSettings.cs6. Traitement dessins (.idw) avec mise à jour références

cd XnrgyEngineeringAutomationTools

.\build-and-run.ps1├── Services/                        # Services métier

```

│   ├── VaultSdkService.cs           # SDK Vault v31.0.847. Mise à jour références des composants suppressed| **Time Tracker** | Analyse temps de travail modules HVAC | Migration |

**Fonctionnalites du script :**

- [+] Compilation automatique en mode Release│   ├── InventorService.cs           # COM Inventor

- [+] Detection automatique de MSBuild (VS 2022 Enterprise/Professional/Community)

- [+] Arret automatique de l'instance existante (taskkill /F)│   ├── InventorCopyDesignService.cs # Copy Design natif8. Copie fichiers orphelins et non-Inventor

- [+] Lancement automatique apres compilation reussie

- [+] Affichage des erreurs de compilation si presentes│   ├── VaultSettingsService.cs      # Config centralisée



### MSBuild manuel│   └── Logger.cs                    # Logging UTF-89. Renommage fichier .ipj| **Update Workspace** | Synchronisation librairies depuis Vault | Planifié |- Interface utilisateur moderne et cohérente## 📋 Description> Développé par **Mohammed Amine Elgalai** - XNRGY Climate Systems ULC🏭 **Suite d'outils d'automatisation engineering unifiée** pour piloter Autodesk Vault Professional 2026 et Inventor Professional 2026.2Application WPF pour l'upload automatisé de fichiers vers Autodesk Vault Professional 2026 avec application automatique des propriétés métier (Project, Reference, Module), catégories, lifecycle et révisions.



```powershell├── Views/                           # Fenêtres et dialogues

& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `

  XnrgyEngineeringAutomationTools.csproj /t:Rebuild /p:Configuration=Release /m /v:minimal /nologo│   ├── LoginWindow.xaml(.cs)10. Switch vers nouveau projet

```

│   ├── CreateModuleWindow.xaml(.cs)

> **[!] IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne genere pas les fichiers .g.cs pour WPF .NET Framework 4.8.

│   ├── CreateModuleSettingsWindow.xaml(.cs)11. Application iProperties finales et paramètres Inventor

---

│   ├── UploadTemplateWindow.xaml(.cs)

## Exclusions de fichiers

│   ├── ChecklistHVACWindow.xaml(.cs)12. Module reste ouvert pour le dessinateur

**Extensions exclues:**

- `.v`, `.bak`, `.old` (Backup Vault)│   ├── ModuleSelectionWindow.xaml(.cs)

- `.tmp`, `.temp` (Temporaires)

- `.ipj` (Projet Inventor)│   ├── PreviewWindow.xaml(.cs)---- Partage de services communs (logging, configuration, etc.)

- `.lck`, `.lock`, `.log` (Systeme/logs)

- `.dwl`, `.dwl2` (AutoCAD locks)│   └── XnrgyMessageBox.xaml(.cs)    # MessageBox custom



**Prefixes exclus:**├── Modules/                         # Modules intégrés**Options de renommage :**

- `~$` (Office temporaire)

- `._` (macOS temporaire)│   └── VaultUpload/

- `Backup_` (Backup generique)

- `.~` (Temporaire generique)│       ├── Models/- Rechercher/Remplacer (cumulatif)



**Dossiers exclus:**│       │   ├── VaultUploadFileItem.cs

- `OldVersions`, `oldversions`

- `Backup`, `backup`│       │   └── VaultUploadModels.cs- Préfixe/Suffixe

- `.vault`, `.git`, `.vs`

│       └── Views/

---

│           ├── VaultUploadModuleWindow.xaml- Checkbox "Inclure fichiers non-Inventor"## Fonctionnalités Implémentées- Déploiement et maintenance simplifiés

## Logs et Debugging

│           └── VaultUploadModuleWindow.xaml.cs

### Emplacement des logs

├── Converters/                      # Convertisseurs WPF

```

bin\Release\Logs\VaultSDK_POC_YYYYMMDD_HHMMSS.log├── Resources/                       # Images et icônes

```

└── Logs/                            # Fichiers logs### 3. Réglages Admin (100%) - NOUVEAU

### Format des logs

```

```

[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] [+] Message

```

---

**Niveaux:** INFO, DEBUG, SUCCESS, WARN, ERROR

**Système de configuration centralisée :**### 1. Vault Upload

**Icones textuelles utilisees (pas d'emoji dans les logs):**

- `[+]` = Succes## Compilation

- `[-]` = Erreur

- `[!]` = Avertissement- Chiffrement AES-256 des fichiers de configuration

- `[>]` = Action en cours

- `[i]` = Information### Script automatique (RECOMMANDÉ)

- `[~]` = Attente/Polling

- `[#]` = Liste/Proprietes- Synchronisation automatique via Vault au démarrage

- `[?]` = Verification

```powershell

---

cd XnrgyEngineeringAutomationTools- Accès restreint aux administrateurs (Role "Administrator" ou Groupe "Admin_Designer")

## Services Principaux

.\build-and-run.ps1

### VaultSdkService.cs

```- Déploiement multi-sites : Saint-Hubert QC + Arizona US (2 usines) = 50+ utilisateurs- Connexion directe via SDK Vault v31.0.84---**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

Service principal pour l'interaction avec Vault SDK.



**Responsabilites :**

- Connexion/deconnexion Vault**Fonctionnalités du script :**

- Chargement des Property Definitions

- Chargement des Categories- [+] Compilation automatique en mode Release

- Chargement des Lifecycle Definitions

- Upload de fichiers avec `FileManager.AddFile`- [+] Détection automatique de MSBuild (VS 2022)**Chemin Vault :**- Scan automatique des modules engineering

- Application des proprietes via `UpdateFileProperties`

- Synchronisation Vault -> iProperties via `IExplorerUtil.UpdateFileProperties`- [+] Arrêt automatique de l'instance existante

- Assignation de categories via `UpdateFileCategories`

- Assignation de lifecycle via `UpdateFileLifeCycleDefinitions` (reflection)- [+] Lancement automatique après compilation réussie```

- Assignation de revisions via `UpdateFileRevisionNumbers`

- Gestion des erreurs Vault (1003, 1013, 1136, etc.)- [+] Affichage des erreurs de compilation si présentes



### InventorService.cs$/Engineering/Inventor_Standards/Automation_Standard/Configuration_Files/XnrgyEngineeringAutomationToolsApp/- Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)



Service pour la connexion COM a Inventor.### Avec MSBuild directement



**Ameliorations v1.1 :**```

- Throttling intelligent (minimum 2 sec entre tentatives)

- Verification fenetre Inventor prete (MainWindowHandle != IntPtr.Zero)```powershell

- Logs silencieux pour COMException 0x800401E3

- Compteur d'echecs consecutifs avec log periodique& 'C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe' `- Application automatique des propriétés métier extraites du chemin



### InventorCopyDesignService.cs  XnrgyEngineeringAutomationTools.csproj `



Service pour Copy Design natif avec gestion des references.  /t:Rebuild `**Paramètres configurables :**



**Methode principale :**  /p:Configuration=Release `

```csharp

Task<bool> ExecuteRealPackAndGoAsync(  /m /v:minimal /nologo- Liste des initiales designers (26 entrées + "Autre...")- Assignation de catégories, lifecycle definitions/states et révisions## Modules Intégrés

    string templatePath,

    string destinationPath,```

    string projectNumber,

    string reference,- Chemins templates et projets

    string module,

    IProgress<string> progress**[!] IMPORTANT**: 

)

```- **NE PAS utiliser `dotnet build`** - il ne génère pas correctement les fichiers `.g.cs` depuis XAML pour WPF .NET Framework 4.8- Extensions Inventor supportées- Synchronisation Vault vers iProperties via `IExplorerUtil`



---- Seul **MSBuild** supporte complètement la génération de code WPF



## Depannage- Dossiers/fichiers exclus



### L'application ne demarre pas---

- Verifier .NET Framework 4.8 installe

- Verifier Vault Professional 2026 installe- Noms des iProperties

- Verifier les dependances NuGet restaurees

## Configuration

### Erreur de connexion Vault

- Verifier serveur accessible

- Verifier vault existe

- Verifier identifiants### appsettings.json (local)

- Voir logs dans `bin\Release\Logs\`

### 4. Interface Utilisateur Moderne### 2. Créer Module - Copy Design

### Erreur connexion Inventor (0x800401E3)

- Inventor doit etre **completement demarre** (fenetre principale visible)```json

- L'app attend que Inventor s'enregistre dans la Running Object Table (ROT)

- Le timer de reconnexion reessaie automatiquement toutes les 3 secondes{



### Proprietes non appliquees  "VaultConfig": {

- Verifier logs : rechercher "Application des proprietes"

- Si erreur 1003 : Fichier en traitement par Job Processor (normal)    "Server": "VAULTPOC",**XnrgyMessageBox :**| Module | Description | Statut |### 🎯 Objectif Principal---

- Si erreur 1013 : CheckOut necessaire (automatique)

- Verifier que les Property Definitions sont chargees (Project, Reference, Module)    "Vault": "PROD_XNGRY",

- Pour fichiers Inventor : Verifier que `IExplorerUtil` est charge

- Pour writeback iProperties : Verifier que le writeback est active dans Vault    "User": "username",- MessageBox personnalisé avec thème sombre XNRGY



### Headers DataGrid invisibles    "Password": ""

- Les styles DataGrid sont definis dans Window.Resources

- Fond sombre (#1A1A28) avec texte bleu XNRGY (#0078D4)  },- Logo et icônes ASCII ([+], [-], [!], [?], [i])**Sources disponibles :**

- Style applique globalement via `<Style TargetType="DataGridColumnHeader">`

  "Paths": {

---

    "DefaultLibrary": "$/Engineering/Library",- Types : Success, Error, Warning, Info, Question

## Changelog

    "DefaultTemplate": "$/Engineering/Library/Xnrgy_Module",

### v1.1.0 (30 Decembre 2025)

    "ProjectsRoot": "C:\\Vault\\Engineering\\Projects"- Boutons : OK, OKCancel, YesNo, YesNoCancel- Depuis Template : `$/Engineering/Library/Xnrgy_Module`|--------|-------------|--------|

**[+] Upload Module integre:**

- Module VaultAutomationTool integre dans l'app principale (`Modules/VaultUpload/`)  }

- Interface avec deux DataGrids (Inventor/Non-Inventor)

- Styles DataGrid avec headers visibles (fond sombre #1A1A28, texte bleu #0078D4)}

- Barre de progression et journal des operations style Creer Module

- Utilise la connexion Vault partagee (pas de login separe)```

- Controles Pause/Stop/Annuler

### 5. Connexions Automatiques- Depuis Projet Existant : Sélection d'un projet local ou Vault

**[+] Upload Template:**

- Nouvelle fenetre pour upload templates (reserve Admin)### Configuration Vault centralisée

- Utilise connexion partagee de l'app principale

- XnrgyMessageBox si utilisateur non-admin



**[+] Corrections Inventor:**Fichier stocké dans Vault: `$/Admin/Config/app_settings.json`

- Throttling intelligent pour eviter spam logs

- Verification fenetre Inventor prete avant connexion COM- Synchronisé automatiquement au démarrage- **Vault Professional 2026** - SDK v31.0.84| **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ **100%** |

- Logs silencieux pour COMException 0x800401E3

- Timer de reconnexion optimise- Chiffrement AES-256 pour les mots de passe



**[+] VaultBulkUploader:**- Accessible via le module "Réglages Admin"- **Inventor Professional 2026.2** - COM

- Outil console pour upload massif (6152 fichiers uploades vers PROD_XNGRY)

- Situe dans `Tools/VaultBulkUploader/`



### v1.0.0 (17 Decembre 2025)---- **Update Workspace** - Synchronisation dossiers au démarrage**Workflow automatisé :**



**[+] Creer Module - Copy Design:**

- Copy Design natif avec 1133 fichiers

- Gestion des fichiers orphelins (1059 fichiers)## Mapping Propriétés

- Mise a jour references IDW

- Switch IPJ automatique

- Application iProperties et parametres Inventor

- Design View "Default" + Workfeatures caches### Extraction depuis le chemin---1. Switch vers projet source (IPJ)| **Créer Module** | Copy Design natif depuis template Library ou projet existant | ✅ **100%** |

- Vue ISO + Zoom All + Save All

- Module reste ouvert pour le dessinateur



**[+] Vault Upload:**```

- Upload complet avec proprietes automatiques

- Gestion Inventor et non-Inventor separeeC:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]\fichier.ipt

- Categories, lifecycle et revisions

- Synchronisation Vault -> iProperties via IExplorerUtil                              ↓         ↓       ↓## Prérequis2. Ouverture Top Assembly



**[+] Reglages Admin:**Vault Property IDs:        ID=112    ID=121  ID=122

- Chiffrement AES-256

- Synchronisation automatique via Vault```

- Interface graphique avec validation temps reel



**[+] Connexions automatiques:**

- Vault SDK v31.0.84### Catégorie vers Lifecycle Definition- Windows 10/11 x643. Application iProperties| **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | 📋 Planifié |Remplacer les multiples applications standalone par une **plateforme unique** avec :

- Inventor COM 2026.2

- Update Workspace au demarrage



### v0.9.0 (15 Decembre 2025)| Catégorie | Lifecycle Definition |- .NET Framework 4.8



- Release initiale beta|-----------|---------------------|

- Dashboard principal avec boutons modules

- Connexion Vault centralisee| Engineering | Flexible Release Process |- Autodesk Vault Professional 2026 (SDK v31.0.84)4. Collecte références (bottom-up)

- Themes sombre/clair

| Office | Simple Release Process |

---

| Standard | Basic Release Process |- Autodesk Inventor Professional 2026.2

## Auteur

| Base | (aucun) |

**Mohammed Amine Elgalai**  

Engineering Automation Developer  - Visual Studio 2022 (pour compilation)5. Copy Design natif avec SaveAs| **DXF Verifier** | Validation des fichiers DXF avant envoi | 📋 Migration |

XNRGY Climate Systems ULC  

Email: mohammedamine.elgalai@xnrgy.com---



---- MSBuild 18.0.0+ (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)



## Licence## Exclusions de fichiers



Proprietaire - XNRGY Climate Systems ULC (c) 20256. Traitement dessins (.idw) avec mise à jour références



---**Extensions exclues:**



**Derniere mise a jour**: 30 Decembre 2025- `.v`, `.bak`, `.old` (Backup Vault)---


- `.tmp`, `.temp` (Temporaires)

- `.ipj` (Projet Inventor)7. Mise à jour références des composants suppressed| **Checklist HVAC** | Validation modules AHU avec stockage Vault | 📋 Migration |- Connexion centralisée à Vault & Inventor

- `.lck`, `.lock`, `.log` (Système/logs)

- `.dwl`, `.dwl2` (AutoCAD locks)## Compilation et Exécution



**Préfixes exclus:**8. Copie fichiers orphelins et non-Inventor

- `~$` (Office temporaire)

- `._` (macOS temporaire)```powershell

- `Backup_` (Backup générique)

# Utiliser le script build-and-run.ps19. Renommage fichier .ipj| **Time Tracker** | Analyse temps de travail modules HVAC | 📋 Migration |

**Dossiers exclus:**

- `OldVersions`, `Backup`cd XnrgyEngineeringAutomationTools

- `.vault`, `.git`, `.vs`

.\build-and-run.ps110. Switch vers nouveau projet

---

```

## Dépannage

11. Application iProperties finales et paramètres Inventor| **Update Workspace** | Synchronisation librairies depuis Vault | 📋 Planifié |- Interface utilisateur moderne et cohérente## 📋 Description## 📋 Description## 📋 Description

### L'application ne démarre pas

- Vérifier .NET Framework 4.8 installé> **IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne génère pas les fichiers .g.cs pour WPF.

- Vérifier Vault Professional 2026 installé

- Vérifier les dépendances NuGet restaurées12. Module reste ouvert pour le dessinateur



### Erreur de connexion Vault---

- Vérifier serveur accessible

- Vérifier vault existe

- Vérifier identifiants

- Voir logs dans `bin\Release\Logs\`## Propriétés XNRGY



### Erreur connexion Inventor (0x800401E3)**Options de renommage :**

- Inventor doit être **complètement démarré** (fenêtre principale visible)

- L'app attend que Inventor s'enregistre dans la Running Object Table (ROT)Le système extrait automatiquement les propriétés depuis le chemin de fichier:

- Le timer de reconnexion réessaie automatiquement toutes les 3 secondes

- Rechercher/Remplacer (cumulatif)---- Partage de services communs (logging, configuration, etc.)

### Propriétés non appliquées

- Vérifier logs : rechercher "Application des propriétés"```

- Si erreur 1003 : Fichier en traitement par Job Processor (normal)

- Si erreur 1013 : CheckOut nécessaire (automatique)C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]- Préfixe/Suffixe



---```



## Dépendances- Checkbox "Inclure fichiers non-Inventor"



```xml| Propriété | ID Vault | Description |

<PackageReference Include="Autodesk.Connectivity.WebServices" Version="31.0.0" />

<PackageReference Include="Autodesk.DataManagement.Client.Framework" Version="31.0.0" />|-----------|----------|-------------|

```

| Project | 112 | Numéro de projet (5 chiffres) |

**Logiciels requis:**

- Autodesk Inventor Professional 2026.2| Reference | 121 | Numéro de référence (2 chiffres) |### 3. Connexions Automatiques## Fonctionnalités Implémentées- Déploiement et maintenance simplifiés

- Autodesk Vault Professional 2026 (SDK v31.0.84)

- Visual Studio 2022 (MSBuild 18.0.0+)| Module | 122 | Numéro de module (2 chiffres) |

- .NET Framework 4.8



---

---

## Auteur

- **Vault Professional 2026** - SDK v31.0.84

**Mohammed Amine Elgalai**  

Design Engineer - XNRGY Climate Systems ULC  ## Architecture

Email: mohammedamine.elgalai@xnrgy.com

- **Inventor Professional 2026.2** - COM

---

```

## Licence

XnrgyEngineeringAutomationTools/- **Update Workspace** - Synchronisation dossiers au démarrage### 1. Vault Upload (100%)

Propriétaire - XNRGY Climate Systems ULC © 2025

├── Views/

---

│   ├── MainWindow.xaml                    # Fenêtre principale hub

## Changelog

│   ├── LoginWindow.xaml                   # Connexion Vault

### v1.1.0 (30 Décembre 2025)

│   ├── CreateModuleWindow.xaml            # Créer Module---

**[+] Upload Module intégré:**

- Module VaultAutomationTool intégré dans l'app principale│   ├── CreateModuleSettingsWindow.xaml    # Réglages Admin

- Interface avec deux DataGrids (Inventor/Non-Inventor)

- Styles DataGrid avec headers visibles (fond sombre, texte bleu XNRGY)│   ├── PreviewWindow.xaml                 # Prévisualisation

- Barre de progression et journal des opérations

- Utilise la connexion Vault partagée│   └── XnrgyMessageBox.xaml               # MessageBox moderne



**[+] Upload Template:**├── Services/## PrérequisModule complet pour l'upload automatisé vers Autodesk Vault Professional 2026.---**XNRGY Engineering Automation Tools** est une application hub centralisée (WPF/.NET Framework 4.8) qui regroupe tous les outils d'automatisation engineering développés pour XNRGY Climate Systems. Cette suite vise à simplifier et accélérer les workflows des équipes de design en intégrant la gestion Vault, les manipulations Inventor, et les validations qualité dans une interface unifiée.

- Nouvelle fenêtre pour upload templates (réservé Admin)

- Utilise connexion partagée (pas de login séparé)│   ├── VaultSdkService.cs                 # Connexion Vault SDK

- XnrgyMessageBox si utilisateur non-admin

│   ├── VaultSettingsService.cs            # Config chiffrée + sync Vault

**[+] Corrections Inventor:**

- Throttling intelligent pour éviter spam logs│   ├── InventorService.cs                 # Connexion Inventor COM

- Vérification fenêtre Inventor prête avant connexion COM

- Logs silencieux pour COMException 0x800401E3│   ├── InventorCopyDesignService.cs       # Copy Design natif- Windows 10/11 x64



### v1.0.0 (17 Décembre 2025)│   └── SettingsService.cs                 # Configuration locale



- Version initiale avec Créer Module et Réglages Admin├── Models/- .NET Framework 4.8

- Connexion centralisée Vault/Inventor

- Thèmes sombre/clair│   ├── ModuleSettings.cs                  # Modèle settings global



---│   └── CreateModuleSettings.cs            # Settings Créer Module- Autodesk Vault Professional 2026 (SDK v31.0.84)**Caractéristiques :**



**Dernière mise à jour**: 30 Décembre 2025├── ViewModels/


│   └── AppMainViewModel.cs                # MVVM ViewModel- Autodesk Inventor Professional 2026.2

├── Resources/

│   └── xnrgy_logo.png                     # Logo XNRGY- Visual Studio 2022 (pour compilation)- Connexion directe via SDK Vault v31.0.84

└── build-and-run.ps1                      # Script compilation MSBuild

```- MSBuild 18.0.0+ (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)



---- Scan automatique des modules engineering (`Projects\[NUM]\REF[XX]\M[XX]`)## 📦 Modules Intégrés



## Changelog---



### v1.1.0 (30 Décembre 2025)- Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

- [+] Système de réglages Admin avec chiffrement AES-256

- [+] Synchronisation automatique des paramètres via Vault## Compilation et Exécution

- [+] XnrgyMessageBox moderne avec thème XNRGY

- [+] Liste des initiales designers configurable (26 entrées)- Application automatique des propriétés métier extraites du chemin

- [+] Vérification admin via Vault API (Roles + Groups)

- [+] Création récursive des dossiers Vault```powershell



### v1.0.0 (29 Décembre 2025)# Utiliser le script build-and-run.ps1- Assignation de catégories, lifecycle definitions/states et révisions

- [+] Vault Upload complet avec propriétés automatiques

- [+] Copy Design depuis template Library ou projet existantcd XnrgyEngineeringAutomationTools

- [+] Connexions automatiques Vault/Inventor

- [+] Mise à jour des références des composants suppressed.\build-and-run.ps1- Synchronisation Vault → iProperties via `IExplorerUtil`| Module | Description | Statut |### 🎯 Objectif PrincipalApplication hub centralisée qui regroupe tous les outils d'automatisation engineering XNRGY :Cette application permet de :

- [+] Support des fichiers .idw dans la mise à jour des références

- [+] Checkbox "Inclure fichiers non-Inventor"```



---- Gestion séparée Inventor / Non-Inventor



## Auteur> **IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne génère pas les fichiers .g.cs pour WPF.



**Mohammed Amine Elgalai**  - Logs détaillés UTF-8|--------|-------------|--------|

Engineering Automation Developer  

XNRGY Climate Systems ULC---



---



*Dernière mise à jour: 30 décembre 2025*## Propriétés XNRGY


### 2. Créer Module - Copy Design (100%)| 📤 **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ **100%** |

Le système extrait automatiquement les propriétés depuis le chemin de fichier:



```

C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]Module pour créer de nouveaux modules depuis le template Library ou un projet existant.| 📦 **Créer Module** | Copy Design natif depuis template Library vers Projects | ✅ **95%** |

```



| Propriété | ID Vault | Description |

|-----------|----------|-------------|**Sources disponibles :**| ⚡ **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | 📋 Planifié |Remplacer les multiples applications standalone par une **plateforme unique** avec :- Scanner automatiquement les modules engineering (structure `Projects\[NUMERO]\REF[NUM]\M[NUM]`)

| Project | 112 | Numéro de projet (5 chiffres) |

| Reference | 121 | Numéro de référence (2 chiffres) |- **Depuis Template** : `$/Engineering/Library/Xnrgy_Module` (1083 fichiers Inventor)

| Module | 122 | Numéro de module (2 chiffres) |

- **Depuis Projet Existant** : Sélection d'un projet local ou Vault existant| 📐 **DXF Verifier** | Validation des fichiers DXF avant envoi | 📋 Migration |

---



## Changelog

**Workflow complet :**| ✅ **Checklist HVAC** | Validation modules AHU avec stockage Vault | 📋 Migration |- Connexion centralisée à Vault & Inventor

### v1.0.0 (En développement)

- Vault Upload complet```

- Copy Design depuis template Library ou projet existant

- Connexions automatiques Vault/Inventor📁 Source: Template Library OU Projet Existant| ⏱️ **Time Tracker** | Analyse temps de travail modules HVAC | 📋 Migration |

- Mise à jour des références des composants suppressed

- Renommage prefix/suffix conservé correctement    ↓

- Checkbox "Inclure fichiers non-Inventor"

- Support des fichiers .idw dans la mise à jour des références📦 Copy Design Natif (bottom-up SaveAs)| 🔄 **Update Workspace** | Synchronisation librairies depuis Vault | 📋 Planifié |- Interface utilisateur moderne et cohérente- **Vault Upload** - Upload automatisé vers Vault avec propriétés (Project/Reference/Module)- Uploader des fichiers vers Vault avec création automatique de l'arborescence



---    ↓



## Auteur📂 Destination: C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]



**Mohammed Amine Elgalai**  ```

Engineering Automation Developer  

XNRGY Climate Systems ULC---- Partage de services communs (logging, configuration, etc.)



---**Étapes automatisées :**



*Dernière mise à jour: 29 décembre 2025*1. Switch vers projet source (IPJ)


2. Ouverture Top Assembly

3. Application iProperties sur le template## ✅ Fonctionnalités Implémentées- Déploiement et maintenance simplifiés- **Pack & Go** - GET depuis Vault, insertion dans assemblages, Copy Design- Appliquer automatiquement les propriétés métier (Project, Reference, Module)

4. Collecte de toutes les références (bottom-up)

5. Copy Design natif avec SaveAs (IPT → IAM → Top Assembly)

6. Traitement des dessins (.idw) avec mise à jour des références

7. **Mise à jour des références des composants suppressed** (nouveauté v1.1)### 1. Vault Upload (100% ✅)

8. Copie des fichiers orphelins (non-référencés dans les assemblages)

9. Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)

10. Renommage du fichier .ipj

11. Switch vers le nouveau projetModule complet pour l'upload automatisé vers Autodesk Vault Professional 2026.---- **Smart Tools** - Création IPT/STEP, génération PDF, iLogic Forms- Assigner des catégories, lifecycle definitions/states et révisions

12. Ouverture du nouveau Top Assembly

13. Application des iProperties finales et paramètres Inventor

14. Design View → "Default", masquage Workfeatures

15. Vue ISO + Zoom All (Fit)**Caractéristiques :**

16. Update All (rebuild) + Save All

17. Module reste ouvert pour le dessinateur- ✅ Connexion directe via SDK Vault v31.0.84



**Gestion intelligente des références :**- ✅ Scan automatique des modules engineering (`Projects\[NUM]\REF[XX]\M[XX]`)## 📦 Modules Intégrés- **DXF Verifier** - Validation des fichiers DXF avant envoi- Gérer l'upload de fichiers Inventor et non-Inventor séparément

- 🔗 Fichiers Library (IPT_Typical_Drawing) : Liens préservés

- 📁 Fichiers Module : Copiés avec références mises à jour- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

- 📄 Fichiers IDW : Références corrigées via `PutLogicalFileNameUsingFull`

- 🔧 **Composants suppressed** : Références mises à jour même si supprimés dans l'assemblage- ✅ Application automatique des propriétés métier extraites du chemin



**Options de renommage (v1.1) :**- ✅ Assignation de catégories, lifecycle definitions/states et révisions

- Rechercher/Remplacer (cumulatif sur NewFileName)

- Préfixe/Suffixe (appliqué sur OriginalFileName pour éviter doublons)- ✅ Synchronisation Vault → iProperties via `IExplorerUtil`| Module | Description | Statut | Source |- **Checklist HVAC** - Validation modules AHU avec stockage Vault

- **Checkbox "Inclure fichiers non-Inventor"** : Contrôle si le renommage s'applique aux fichiers non-Inventor

- ✅ Gestion séparée Inventor / Non-Inventor

### 3. Connexions Automatiques

- ✅ Logs détaillés UTF-8 avec emojis|--------|-------------|--------|--------|

- **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique

- **Inventor Professional 2026.2** - COM avec détection d'instance active

- **Update Workspace** - Synchronisation dossiers au démarrage :

  - `$/Content Center Files`### 2. Créer Module - Copy Design (95% ✅)| 📤 **Vault Upload** | Upload automatisé vers Vault avec propriétés (Project/Ref/Module) | ✅ **Fonctionnel** | Natif |- **Update Workspace** - Synchronisation des librairies depuis Vault## 🎯 Caractéristiques

  - `$/Engineering/Inventor_Standards`

  - `$/Engineering/Library/Cabinet`

  - `$/Engineering/Library/Xnrgy_M99`

  - `$/Engineering/Library/Xnrgy_Module`Module pour créer de nouveaux modules depuis le template Library avec Copy Design natif.| 📦 **Pack & Go** | GET depuis Vault + Copy Design natif | 🚧 **En cours** | Natif |



---



## Prérequis**Workflow complet :**| ⚡ **Smart Tools** | Création IPT/STEP, génération PDF, iLogic Forms | 📋 **Planifié** | Nouveau |



- **Windows 10/11 x64**```

- **.NET Framework 4.8**

- **Autodesk Vault Professional 2026** (SDK v31.0.84)📁 Template: $/Engineering/Library/Xnrgy_Module| 📐 **DXF Verifier** | Validation DXF/CSV vs PDF Cut Lists | 📋 **Migration** | `DXFVerifier/` |

- **Autodesk Inventor Professional 2026.2**

- **Visual Studio 2022** (pour compilation)    ↓

- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)

📦 Copy Design Natif (1083 fichiers Inventor)| ✅ **Checklist HVAC** | Validation modules AHU avec stockage Vault | 📋 **Migration** | `ChecklistHVAC/` |## 🎯 Fonctionnalités- ✅ Connexion directe à Vault via SDK (VaultSDKService.cs)

---

    ↓

## Architecture Technique

📂 Destination: C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]| ⏱️ **Time Tracker** | Analyse temps de travail modules HVAC | 📋 **Migration** | `HVACTimeTracker/` |

### Stack Technologique

```

```

┌─────────────────────────────────────────────────────────┐| 🔄 **Update Workspace** | Synchronisation librairies depuis Vault | 📋 **Planifié** | Nouveau |- ✅ Scan automatique des modules engineering (FileScanner.cs)

│                    Présentation (WPF)                   │

│  MainWindow.xaml │ Views/*.xaml │ MVVM Pattern          │**Étapes automatisées :**

├─────────────────────────────────────────────────────────┤

│                   ViewModels (MVVM)                     │1. ✅ Switch vers projet template (IPJ)

│  AppMainViewModel.cs │ RelayCommand │ INotifyProperty   │

├─────────────────────────────────────────────────────────┤2. ✅ Ouverture Top Assembly (Module_.iam)

│                    Services Layer                       │

│  VaultSDKService │ InventorService │ Logger             │3. ✅ Application iProperties sur le template---### Connexions Automatiques- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

│  InventorCopyDesignService │ ModuleCopyService          │

├─────────────────────────────────────────────────────────┤4. ✅ Collecte de toutes les références (bottom-up)

│                    Models (Data)                        │

│  FileItem │ ModuleInfo │ ProjectProperties │ Config     │5. ✅ Copy Design natif avec SaveAs (IPT → IAM → Top Assembly)

├─────────────────────────────────────────────────────────┤

│                   External APIs                         │6. ✅ Traitement des dessins (.idw) avec mise à jour des références

│  Vault SDK v31.0.84 │ Inventor COM 2026.2               │

└─────────────────────────────────────────────────────────┘7. ✅ Copie des fichiers orphelins (1059 fichiers non-référencés)## ✅ Fonctionnalités Implémentées- ✅ Connexion centralisée à **Vault Professional 2026** (SDK v31.0.84)- ✅ Application automatique des propriétés métier extraites du chemin

```

8. ✅ Copie des fichiers non-Inventor (Excel, PDF, Word, etc.)

### Structure du Projet

9. ✅ Renommage du fichier .ipj (XXXXX-XX-XX_2026.ipj → 123450101.ipj)

```

XnrgyEngineeringAutomationTools/10. ✅ Switch vers le nouveau projet

├── App.xaml(.cs)                    # Point d'entrée application

├── MainWindow.xaml(.cs)             # Dashboard principal11. ✅ Ouverture du nouveau Top Assembly### 1. Vault Upload (100%)- ✅ Connexion COM à **Inventor Professional 2026.2**- ✅ Assignation de catégories, lifecycle definitions/states et révisions

├── appsettings.json                 # Configuration sauvegardée

│12. ✅ Application des iProperties finales

├── Models/                          # Modèles de données

│   ├── ApplicationConfiguration.cs  # Configuration application13. ✅ Application des paramètres Inventor

│   ├── FileItem.cs                  # Item fichier pour DataGrid

│   ├── ModuleInfo.cs                # Informations module14. ✅ Design View → "Default"

│   └── CreateModuleRequest.cs       # Requête création module

│15. ✅ Masquage des Workfeatures (plans, axes, points)Module complet pour l'upload automatisé vers Autodesk Vault Professional 2026.- ✅ Détection automatique d'Inventor s'il est en cours d'exécution- ✅ Gestion de la progression et pause/reprise

├── ViewModels/                      # MVVM ViewModels

│   └── AppMainViewModel.cs          # ViewModel principal (1758L)16. ✅ Vue ISO + Zoom All (Fit)

│

├── Views/                           # Fenêtres et dialogues17. ✅ Update All (rebuild)

│   ├── CreateModuleWindow.xaml(.cs) # Fenêtre création module

│   └── VaultConnectionDialog.xaml   # Dialogue connexion Vault18. ✅ Save All

│

├── Services/                        # Services métier19. ✅ Module reste ouvert pour le dessinateur**Caractéristiques :**- ✅ Logs détaillés UTF-8 avec emoji (Logger.cs)

│   ├── VaultSDKService.cs           # API Vault SDK (3224L)

│   ├── InventorCopyDesignService.cs # Copy Design natif (2298L)

│   └── Logger.cs                    # Logging UTF-8

│**Gestion intelligente des références :**- ✅ Connexion directe via SDK Vault v31.0.84

└── Logs/                            # Fichiers log

    └── VaultSDK_POC_*.log- 🔗 Fichiers Library (IPT_Typical_Drawing) : Liens préservés

```

- 📁 Fichiers Module : Copiés avec références mises à jour- ✅ Scan automatique des modules engineering (`Projects\[NUM]\REF[XX]\M[XX]`)### Update Workspace (GET automatique)- ✅ Exclusion automatique des fichiers temporaires (.bak, .dwl, .log, OldVersions, ~$)

---

- 📄 Fichiers IDW : Références corrigées via `PutLogicalFileNameUsingFull`

## Compilation et Exécution

- ✅ Upload de tous types de fichiers (Inventor, PDF, Excel, Word, images)

### Build (OBLIGATOIRE: MSBuild)

### 3. Connexions Automatiques

```powershell

# Utiliser le script build-and-run.ps1- ✅ Application automatique des propriétés métier extraites du cheminAu démarrage ou sur demande, synchronisation des dossiers essentiels :- ✅ Sauvegarde configuration (appsettings.json)

cd XnrgyEngineeringAutomationTools

.\build-and-run.ps1- ✅ **Vault Professional 2026** - SDK v31.0.84 avec reconnexion automatique



# OU manuellement avec MSBuild- ✅ **Inventor Professional 2026.2** - COM avec détection d'instance active- ✅ Assignation de catégories, lifecycle definitions/states et révisions

& "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" `

    XnrgyEngineeringAutomationTools.csproj /t:Rebuild /p:Configuration=Release- ✅ **Update Workspace** - Synchronisation dossiers au démarrage :

```

  - `$/Content Center Files`- ✅ Synchronisation Vault → iProperties via `IExplorerUtil`- `$/Content Center Files` → `C:\Vault\Content Center Files`- ✅ Interface MVVM avec séparation Inventor/Non-Inventor

> **IMPORTANT**: Ne PAS utiliser `dotnet build` - il ne génère pas les fichiers .g.cs pour WPF.

  - `$/Engineering/Inventor_Standards`

### Exécution

  - `$/Engineering/Library/Cabinet`- ✅ Gestion séparée Inventor / Non-Inventor

```powershell

.\bin\Release\XnrgyEngineeringAutomationTools.exe  - `$/Engineering/Library/Xnrgy_M99`

```

  - `$/Engineering/Library/Xnrgy_Module`- ✅ Logs détaillés UTF-8 avec emojis- `$/Engineering/Inventor_Standards` → `C:\Vault\Engineering\Inventor_Standards`

---



## Logs

---

Les logs sont générés dans `bin\Release\Logs\VaultSDK_POC_*.log`



**Format:**

```## 📦 Prérequis### 2. Pack & Go (70%)- `$/Engineering/Library/Cabinet` → `C:\Vault\Engineering\Library\Cabinet`## 📦 Prérequis

[YYYY-MM-DD HH:MM:SS.mmm] [LEVEL] [+] Message

```



**Niveaux:** TRACE, DEBUG, INFO, WARNING, ERROR- **Windows 10/11 x64**



**Consulter les derniers logs:**- **.NET Framework 4.8**

```powershell

Get-Content "bin\Release\Logs\VaultSDK_POC_*.log" | Select-Object -Last 100- **Autodesk Vault Professional 2026** (SDK v31.0.84)Module pour extraire depuis Vault et créer des copies de modules avec références mises à jour.- `$/Engineering/Library/Xnrgy_M99` → `C:\Vault\Engineering\Library\Xnrgy_M99`

```

- **Autodesk Inventor Professional 2026.2**

---

- **Visual Studio 2022** (pour compilation)

## Propriétés XNRGY

- **MSBuild 18.0.0+** (REQUIS - `dotnet build` ne fonctionne PAS pour WPF)

Le système extrait automatiquement les propriétés depuis le chemin de fichier:

**Implémenté :**- `$/Engineering/Library/Xnrgy_Module` → `C:\Vault\Engineering\Library\Xnrgy_Module`- Windows 10/11 x64

```

C:\Vault\Engineering\Projects\[PROJECT]\REF[XX]\M[XX]---

                              ↓         ↓       ↓

Exemple:                   10359      REF09     M03- ✅ GET automatique depuis Vault avec dépendances

Vault Property IDs:      ID=112     ID=121   ID=122

```## 🏗️ Architecture Technique



| Propriété | ID Vault | Description |- ✅ Extraction vers dossier temporaire- .NET Framework 4.8

|-----------|----------|-------------|

| Project | 112 | Numéro de projet (5 chiffres) |### Stack Technologique

| Reference | 121 | Numéro de référence (2 chiffres) |

| Module | 122 | Numéro de module (2 chiffres) |- ✅ Interface de sélection de destination



---```



## Changelog┌─────────────────────────────────────────────────────────┐- 🚧 Copy Design natif (bottom-up SaveAs avec références)## 📦 Modules Intégrés- Autodesk Vault Professional 2026



### v1.1 (2025-12-29)│                    Présentation (WPF)                   │

- **Fix** : Mise à jour des références des composants suppressed dans les assemblages

- **Fix** : Renommage prefix/suffix ne se réinitialise plus│  MainWindow.xaml │ Views/*.xaml │ MVVM Pattern          │

- **Ajout** : Checkbox "Inclure fichiers non-Inventor" pour contrôler le renommage

- **Ajout** : Support des fichiers .idw dans la mise à jour des références├─────────────────────────────────────────────────────────┤



### v1.0 (2025-12-15)│                   ViewModels (MVVM)                     │**En cours :**- Visual Studio 2022 ou supérieur (pour compilation)

- Release initiale

- Vault Upload complet│  AppMainViewModel.cs │ RelayCommand │ INotifyProperty   │

- Copy Design depuis template Library

- Connexions automatiques Vault/Inventor├─────────────────────────────────────────────────────────┤- 🔄 Correction des références croisées entre assemblages siblings



---│                    Services Layer                       │



## Auteur│  VaultSDKService │ InventorService │ Logger             │- 🔄 Gestion OldVersions et fichiers obsolètes| Module | Description | Status |- MSBuild 18.0.0+ (REQUIS - dotnet build ne fonctionne PAS pour WPF)



**Mohammed Amine Elgalai**  │  InventorCopyDesignService │ ModuleCopyService          │

Engineering Automation Developer  

XNRGY Climate Systems ULC├─────────────────────────────────────────────────────────┤



---│                    Models (Data)                        │



*Dernière mise à jour: 29 décembre 2025*│  FileItem │ ModuleInfo │ ProjectProperties │ Config     │### 3. Connexions Automatiques|--------|-------------|--------|


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
