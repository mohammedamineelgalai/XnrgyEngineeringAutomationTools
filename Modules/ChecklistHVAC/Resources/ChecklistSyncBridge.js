/**
 * Script JavaScript pour intégration Checklist HVAC avec synchronisation Vault
 * Ce script doit être injecté dans le HTML de la checklist pour activer la synchronisation
 * 
 * Usage dans le HTML:
 * <script src="checklistSyncBridge.js"></script>
 * 
 * OU injecter directement via WebView2 (recommandé - déjà fait dans ChecklistHVACWindow.xaml.cs)
 */

(function() {
    'use strict';

    // Vérifier si window.checklistSync existe (injecté par C#)
    if (typeof window.checklistSync === 'undefined') {
        console.warn('[ChecklistSync] Bridge C# non disponible - synchronisation désactivée');
        
        // Fallback: utiliser localStorage uniquement
        window.checklistSync = {
            saveData: function(moduleId, projectNumber, reference, module, data) {
                try {
                    localStorage.setItem(`checklist_${moduleId}`, JSON.stringify(data));
                    console.log(`[ChecklistSync] Données sauvegardées en localStorage: ${moduleId}`);
                    return "OK";
                } catch (e) {
                    console.error('[ChecklistSync] Erreur sauvegarde localStorage:', e);
                    return "ERROR: " + e.message;
                }
            },
            loadData: function(moduleId) {
                try {
                    const json = localStorage.getItem(`checklist_${moduleId}`);
                    return json ? JSON.parse(json) : null;
                } catch (e) {
                    console.error('[ChecklistSync] Erreur chargement localStorage:', e);
                    return null;
                }
            },
            syncNow: function(moduleId, projectNumber, reference, module) {
                console.warn('[ChecklistSync] Synchronisation Vault non disponible - mode localStorage uniquement');
            },
            getSyncStatus: function() {
                return JSON.stringify({ connected: false, autoSync: false, mode: 'localStorage' });
            }
        };
    }

    /**
     * Fonction helper pour sauvegarder automatiquement après modification
     * À appeler depuis votre code React après chaque modification de checkpoint
     */
    window.saveChecklistResponse = function(moduleId, projectNumber, reference, module, checkpointId, status, comment, userInitials) {
        try {
            // Charger les données existantes
            let data = window.checklistSync.loadData(moduleId);
            
            if (!data) {
                // Créer une nouvelle structure
                data = {
                    moduleId: moduleId,
                    projectNumber: projectNumber,
                    reference: reference,
                    module: module,
                    lastModifiedBy: userInitials,
                    lastModifiedDate: new Date().toISOString(),
                    version: 1,
                    responses: {}
                };
            }

            // Mettre à jour la réponse du checkpoint
            data.responses[checkpointId] = {
                checkpointId: checkpointId,
                status: status,
                comment: comment || '',
                userInitials: userInitials,
                modifiedDate: new Date().toISOString()
            };

            // Mettre à jour les métadonnées
            data.lastModifiedBy = userInitials;
            data.lastModifiedDate = new Date().toISOString();
            data.version = (data.version || 0) + 1;

            // Sauvegarder (va aussi sync avec Vault si disponible)
            const result = window.checklistSync.saveData(moduleId, projectNumber, reference, module, data);
            
            if (result === "OK") {
                console.log(`[ChecklistSync] Réponse checkpoint ${checkpointId} sauvegardée`);
                
                // Afficher notification visuelle (optionnel)
                if (window.showNotification) {
                    window.showNotification('Réponse sauvegardée', 'success');
                }
            } else {
                console.error('[ChecklistSync] Erreur sauvegarde:', result);
                if (window.showNotification) {
                    window.showNotification('Erreur sauvegarde: ' + result, 'error');
                }
            }

            return result === "OK";
        } catch (e) {
            console.error('[ChecklistSync] Erreur saveChecklistResponse:', e);
            return false;
        }
    };

    /**
     * Fonction helper pour charger toutes les réponses d'un module
     * À appeler au chargement du module dans votre code React
     */
    window.loadChecklistData = function(moduleId) {
        try {
            const data = window.checklistSync.loadData(moduleId);
            if (data && data.responses) {
                console.log(`[ChecklistSync] ${Object.keys(data.responses).length} réponses chargées pour ${moduleId}`);
                return data.responses;
            }
            return {};
        } catch (e) {
            console.error('[ChecklistSync] Erreur loadChecklistData:', e);
            return {};
        }
    };

    /**
     * Fonction helper pour synchroniser manuellement avec Vault
     * À appeler depuis un bouton "Synchroniser maintenant"
     */
    window.syncChecklistWithVault = function(moduleId, projectNumber, reference, module) {
        try {
            console.log(`[ChecklistSync] Démarrage synchronisation manuelle: ${moduleId}`);
            
            // Charger les données locales
            const data = window.checklistSync.loadData(moduleId);
            if (!data) {
                console.warn('[ChecklistSync] Aucune donnée locale à synchroniser');
                return false;
            }

            // Sauvegarder (ce qui déclenchera la sync avec Vault)
            const result = window.checklistSync.saveData(moduleId, projectNumber, reference, module, data);
            
            // Forcer une synchronisation immédiate
            window.checklistSync.syncNow(moduleId, projectNumber, reference, module);

            if (result === "OK") {
                console.log('[ChecklistSync] Synchronisation manuelle déclenchée');
                if (window.showNotification) {
                    window.showNotification('Synchronisation avec Vault en cours...', 'info');
                }
                return true;
            }

            return false;
        } catch (e) {
            console.error('[ChecklistSync] Erreur syncChecklistWithVault:', e);
            return false;
        }
    };

    /**
     * Fonction helper pour obtenir le statut de synchronisation
     * À utiliser pour afficher un indicateur dans l'UI
     */
    window.getChecklistSyncStatus = function() {
        try {
            const statusJson = window.checklistSync.getSyncStatus();
            return JSON.parse(statusJson);
        } catch (e) {
            return { connected: false, autoSync: false, error: e.message };
        }
    };

    // Auto-sauvegarde toutes les 30 secondes (backup)
    let autoSaveInterval = null;
    window.startAutoSave = function(moduleId, projectNumber, reference, module, getCurrentDataCallback) {
        if (autoSaveInterval) {
            clearInterval(autoSaveInterval);
        }

        autoSaveInterval = setInterval(() => {
            try {
                const currentData = getCurrentDataCallback();
                if (currentData) {
                    window.checklistSync.saveData(moduleId, projectNumber, reference, module, currentData);
                    console.log('[ChecklistSync] Auto-sauvegarde effectuée');
                }
            } catch (e) {
                console.error('[ChecklistSync] Erreur auto-sauvegarde:', e);
            }
        }, 30000); // Toutes les 30 secondes
    };

    window.stopAutoSave = function() {
        if (autoSaveInterval) {
            clearInterval(autoSaveInterval);
            autoSaveInterval = null;
        }
    };

    console.log('[ChecklistSync] Bridge JavaScript initialisé');
})();



