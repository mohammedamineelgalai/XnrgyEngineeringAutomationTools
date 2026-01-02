# Architecture - XNRGY Engineering Automation Tools

> **Document technique** - Structure du projet et plan d'organisation
> 
> **Version**: 1.0.0 | **Date**: 2026-01-02
> **Auteur**: Mohammed Amine Elgalai - XNRGY Climate Systems ULC

---

## Vue d'Ensemble

L'application est structuree selon le pattern **MVVM** (Model-View-ViewModel) avec une architecture modulaire. Cette documentation decrit la structure actuelle et le plan d'organisation cible.

---

## Structure Actuelle (v1.0.0)

```
XnrgyEngineeringAutomationTools/
|
|-- App.xaml / App.xaml.cs              # Point d'entree, ResourceDictionary
|-- MainWindow.xaml / MainWindow.xaml.cs # Hub principal avec tous les boutons modules
|-- appsettings.json                     # Configuration utilisateur sauvegardee
|
|-- Models/                              # Modeles de donnees partages
|   |-- ApplicationConfiguration.cs      # Config app (chemins, extensions)
|   |-- CategoryItem.cs                  # Categorie Vault pour ComboBox
|   |-- CreateModuleRequest.cs           # Requete Pack&Go
|   |-- CreateModuleSettings.cs          # Reglages admin
|   |-- FileItem.cs                      # Item fichier generique
|   |-- FileToUpload.cs                  # Fichier a uploader vers Vault
|   |-- LifecycleDefinitionItem.cs       # Lifecycle Definition Vault
|   |-- LifecycleStateItem.cs            # Lifecycle State Vault
|   |-- ModuleInfo.cs                    # Info module (Project/Ref/Module)
|   |-- ProjectInfo.cs                   # Info projet
|   |-- ProjectProperties.cs             # Proprietes iProperties
|   |-- VaultConfiguration.cs            # Config connexion Vault
|
|-- Services/                            # Services metier partages
|   |-- ApprenticePropertyService.cs     # Lecture iProperties via Apprentice
|   |-- CredentialsManager.cs            # Gestion credentials chiffrees
|   |-- InventorCopyDesignService.cs     # Copy Design natif (~900 lignes)
|   |-- InventorPropertyService.cs       # iProperties via Inventor COM
|   |-- InventorService.cs               # Connexion Inventor COM
|   |-- JournalColorService.cs           # Couleurs journal (succes/erreur/etc)
|   |-- Logger.cs                        # Logging NLog
|   |-- ModuleCopyService.cs             # Copie fichiers module
|   |-- NativeOlePropertyService.cs      # Lecture OLE native
|   |-- OlePropertyService.cs            # Lecture OLE via OpenMCDF
|   |-- SettingsService.cs               # Gestion settings app
|   |-- SimpleLogger.cs                  # Logger simple
|   |-- ThemeHelper.cs                   # Gestion themes sombre/clair
|   |-- UserPreferencesManager.cs        # Preferences utilisateur
|   |-- VaultSDKService.cs               # SDK Vault v31.0.84 (~1200 lignes)
|   |-- VaultSettingsService.cs          # Config chiffree + sync Vault
|   |-- WindowsPropertyService.cs        # Lecture proprietes Windows Shell
|
|-- ViewModels/                          # ViewModels MVVM
|   |-- AppMainViewModel.cs              # VM principal (MainWindow)
|
|-- Views/                               # Toutes les fenetres
|   |-- ChecklistHVACWindow.xaml(.cs)    # Module: Checklist HVAC
|   |-- CreateModuleSettingsWindow.xaml(.cs) # Module: Reglages Admin
|   |-- CreateModuleWindow.xaml(.cs)     # Module: Pack & Go
|   |-- LoginWindow.xaml(.cs)            # Partage: Connexion Vault
|   |-- ModuleSelectionWindow.xaml(.cs)  # Partage: Selection module
|   |-- PreviewWindow.xaml(.cs)          # Partage: Previsualisation
|   |-- UploadTemplateWindow.xaml(.cs)   # Module: Upload Template
|   |-- VaultUploadWindow.xaml(.cs)      # (Legacy - remplace par VaultUploadModuleWindow)
|   |-- XnrgyMessageBox.xaml(.cs)        # Partage: MessageBox moderne
|
|-- Modules/                             # Modules isoles
|   |-- VaultUpload/                     # Module Upload Vault
|       |-- Models/
|       |   |-- VaultUploadFileItem.cs
|       |   |-- VaultUploadModels.cs
|       |-- Views/
|           |-- VaultUploadModuleWindow.xaml(.cs)
|
|-- Styles/                              # Styles centralises
|   |-- XnrgyStyles.xaml                 # 50+ styles partages (couleurs, boutons, inputs)
|
|-- Converters/                          # Convertisseurs WPF
|   |-- BoolToVisibilityConverter.cs
|   |-- etc.
|
|-- Resources/                           # Images et icones
|   |-- VaultAutomationTool.ico
|   |-- VaultAutomationTool.png
|
|-- Assets/                              # Assets visuels
|   |-- Icons/
|       |-- DXFVerifier.png
|       |-- XnrgyTools.ico
|
|-- Temp&Test/                           # Fichiers temporaires/test (exclus du build)
|   |-- TestWindowsPropertyService.cs
|   |-- DiagnoseOleProperties.cs
|
|-- Backups/                             # Sauvegardes locales
|   |-- BACKUP_YYYYMMDD_HHMMSS/
|
|-- Logs/                                # Fichiers logs
|   |-- VaultSDK_POC_*.log
```

---

## Modules Fonctionnels

### 1. Upload Module (VaultUpload)
**Statut**: [+] Isole dans Modules/VaultUpload

| Fichier | Localisation |
|---------|--------------|
| VaultUploadModuleWindow.xaml(.cs) | Modules/VaultUpload/Views/ |
| VaultUploadFileItem.cs | Modules/VaultUpload/Models/ |
| VaultUploadModels.cs | Modules/VaultUpload/Models/ |

**Services utilises**: VaultSDKService (partage)

---

### 2. Pack & Go (CreateModule)
**Statut**: [~] Dans Views/ (a migrer vers Modules/)

| Fichier | Localisation Actuelle | Cible |
|---------|----------------------|-------|
| CreateModuleWindow.xaml(.cs) | Views/ | Modules/CreateModule/Views/ |
| CreateModuleRequest.cs | Models/ | Modules/CreateModule/Models/ |
| CreateModuleSettings.cs | Models/ | Modules/CreateModule/Models/ |
| InventorCopyDesignService.cs | Services/ | Modules/CreateModule/Services/ |

**Services utilises**: VaultSDKService, InventorService (partages)

---

### 3. Upload Template
**Statut**: [~] Dans Views/ (a migrer vers Modules/)

| Fichier | Localisation Actuelle | Cible |
|---------|----------------------|-------|
| UploadTemplateWindow.xaml(.cs) | Views/ | Modules/UploadTemplate/Views/ |

**Services utilises**: VaultSDKService (partage)

---

### 4. Checklist HVAC
**Statut**: [~] Dans Views/ (a migrer vers Modules/)

| Fichier | Localisation Actuelle | Cible |
|---------|----------------------|-------|
| ChecklistHVACWindow.xaml(.cs) | Views/ | Modules/ChecklistHVAC/Views/ |

---

### 5. Reglages Admin
**Statut**: [~] Integre dans CreateModule (CreateModuleSettingsWindow)

---

## Composants Partages (Shared)

Ces composants sont utilises par plusieurs modules et doivent rester dans un emplacement partage:

### Views Partagees
| Fichier | Description |
|---------|-------------|
| LoginWindow.xaml(.cs) | Connexion Vault |
| ModuleSelectionWindow.xaml(.cs) | Selection module a uploader |
| PreviewWindow.xaml(.cs) | Previsualisation fichiers |
| XnrgyMessageBox.xaml(.cs) | MessageBox stylise XNRGY |

### Services Partages
| Fichier | Description |
|---------|-------------|
| VaultSDKService.cs | API Vault SDK v31.0.84 |
| InventorService.cs | Connexion Inventor COM |
| Logger.cs | Logging application |
| ThemeHelper.cs | Gestion themes |
| UserPreferencesManager.cs | Preferences utilisateur |
| VaultSettingsService.cs | Config chiffree |

### Models Partages
| Fichier | Description |
|---------|-------------|
| CategoryItem.cs | Categories Vault |
| LifecycleDefinitionItem.cs | Lifecycle Vault |
| ModuleInfo.cs | Info module |
| ProjectInfo.cs | Info projet |

---

## Structure Cible (Plan de Migration)

```
XnrgyEngineeringAutomationTools/
|
|-- App.xaml / App.xaml.cs
|-- MainWindow.xaml / MainWindow.xaml.cs
|-- appsettings.json
|
|-- Shared/                              # [NOUVEAU] Composants partages
|   |-- Views/
|   |   |-- LoginWindow.xaml(.cs)
|   |   |-- ModuleSelectionWindow.xaml(.cs)
|   |   |-- PreviewWindow.xaml(.cs)
|   |   |-- XnrgyMessageBox.xaml(.cs)
|   |-- Services/
|   |   |-- VaultSDKService.cs
|   |   |-- InventorService.cs
|   |   |-- Logger.cs
|   |   |-- (autres services partages)
|   |-- Models/
|       |-- CategoryItem.cs
|       |-- ModuleInfo.cs
|       |-- ProjectInfo.cs
|       |-- (autres models partages)
|
|-- Modules/
|   |-- VaultUpload/                     # [EXISTE] Module Upload Vault
|   |   |-- Views/
|   |   |-- Models/
|   |   |-- Services/
|   |
|   |-- CreateModule/                    # [A CREER] Module Pack & Go
|   |   |-- Views/
|   |   |   |-- CreateModuleWindow.xaml(.cs)
|   |   |   |-- CreateModuleSettingsWindow.xaml(.cs)
|   |   |-- Models/
|   |   |   |-- CreateModuleRequest.cs
|   |   |   |-- CreateModuleSettings.cs
|   |   |-- Services/
|   |       |-- InventorCopyDesignService.cs
|   |
|   |-- UploadTemplate/                  # [A CREER] Module Upload Template
|   |   |-- Views/
|   |       |-- UploadTemplateWindow.xaml(.cs)
|   |
|   |-- ChecklistHVAC/                   # [A CREER] Module Checklist
|       |-- Views/
|           |-- ChecklistHVACWindow.xaml(.cs)
|
|-- Styles/
|-- Converters/
|-- Resources/
|-- Assets/
|-- Temp&Test/
|-- Backups/
|-- Logs/
```

---

## Namespaces

### Convention Actuelle
```csharp
namespace XnrgyEngineeringAutomationTools.Views
namespace XnrgyEngineeringAutomationTools.Services
namespace XnrgyEngineeringAutomationTools.Models
namespace XnrgyEngineeringAutomationTools.Modules.VaultUpload.Views
```

### Convention Cible
```csharp
// Partages
namespace XnrgyEngineeringAutomationTools.Shared.Views
namespace XnrgyEngineeringAutomationTools.Shared.Services
namespace XnrgyEngineeringAutomationTools.Shared.Models

// Modules
namespace XnrgyEngineeringAutomationTools.Modules.VaultUpload.Views
namespace XnrgyEngineeringAutomationTools.Modules.CreateModule.Views
namespace XnrgyEngineeringAutomationTools.Modules.CreateModule.Services
```

---

## Notes de Migration

**[!] IMPORTANT**: La migration des fichiers vers la structure cible doit etre faite incrementalement pour eviter de casser le build:

1. **Phase 1**: Creer la structure de dossiers (FAIT)
2. **Phase 2**: Migrer un module a la fois
3. **Phase 3**: Mettre a jour les namespaces
4. **Phase 4**: Mettre a jour le .csproj
5. **Phase 5**: Tester le build et l'execution

**Fichiers a NE PAS deplacer** (risque de casser des references):
- MainWindow.xaml(.cs) - Hub principal
- App.xaml(.cs) - Point d'entree
- VaultSDKService.cs - Reference par tous les modules
- InventorService.cs - Reference par plusieurs modules

---

## Build Configuration

### MSBuild (REQUIS pour WPF .NET Framework 4.8)
```powershell
.\build-and-run.ps1   # Script automatique
```

### Exclusions du Build
- `Temp&Test/**` - Fichiers de test/debug
- `Backups/**` - Sauvegardes locales
- `Logs/**` - Fichiers logs

---

## Changelog Structure

| Date | Version | Modification |
|------|---------|--------------|
| 2026-01-02 | 1.0.0 | Documentation initiale |
| 2026-01-02 | 1.0.0 | Creation Temp&Test avec fichiers de test deplaces |
| 2026-01-02 | 1.0.0 | Creation dossiers modules vides (UploadTemplate, CreateModule, ChecklistHVAC, Shared) |

---

*Document maintenu par Mohammed Amine Elgalai - XNRGY Climate Systems ULC*
