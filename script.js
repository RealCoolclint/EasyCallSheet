document.addEventListener('DOMContentLoaded', function() {
    // Initialisation des éléments DOM
    const form = document.getElementById('callSheetForm');
    const previewButton = document.getElementById('previewButton');
    const generatePdfButton = document.getElementById('generatePdfButton');
    const previewContent = document.getElementById('previewContent');
    const managerNameInput = document.getElementById('managerName');
    const managerPhoneInput = document.getElementById('managerPhone');
    const managerSuggestions = document.getElementById('managerSuggestions');
    const manageManagersButton = document.getElementById('manageManagersButton');
    const managersModal = document.getElementById('managersModal');
    const modalClose = document.querySelector('.modal-close');
    const addManagerButton = document.getElementById('addManagerButton');
    const newManagerNameInput = document.getElementById('newManagerName');
    const newManagerPhoneInput = document.getElementById('newManagerPhone');
    const managersList = document.getElementById('managersList');
    
    // Précharger l'image du logo
    const logoImg = new Image();
    logoImg.src = 'logo_etudiant.png';
    
    // Définir la date du jour par défaut
    document.getElementById('date').valueAsDate = new Date();
    
    // ===== NOUVEAU : GESTION DES PARAMÈTRES URL DEPUIS MONDAY =====
    
    /**
     * Extraire intelligemment le nom de l'invité et l'école depuis un titre
     * Exemples de formats gérés :
     * - "Interview Jean Dupont - Sciences Po"
     * - "L'interro Sophie Martin (HEC)"
     * - "C'est quoi Pierre Durand"
     * - "Marc Lambert - École 42"
     */
    function extractGuestAndSchool(title) {
        if (!title) return { guest: '', school: '' };
        
        title = title.trim();
        
        // Patterns de recherche
        const patterns = [
            // Format: "Nom - École" ou "Nom (École)"
            { regex: /^(?:.*?\s)?([A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+(?:\s+[A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+)+)\s*[-–(]\s*([^)]+)/, guest: 1, school: 2 },
            
            // Format: "Format Nom École" (ex: "Interview Jean Dupont Sciences Po")
            { regex: /^(?:Interview|L'interview|L'interro|C'est quoi|Audrey t'explique)\s+([A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+(?:\s+[A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+)+)\s+(.+)$/i, guest: 1, school: 2 },
            
            // Format: "Nom seulement" (sans école)
            { regex: /([A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+(?:\s+[A-ZÀÂÄÇÉÈÊËÏÎÔÙÛÜ][a-zàâäçéèêëïîôùûü]+)+)/, guest: 1, school: null }
        ];
        
        for (const pattern of patterns) {
            const match = title.match(pattern.regex);
            if (match) {
                const guest = match[pattern.guest] ? match[pattern.guest].trim() : '';
                const school = pattern.school && match[pattern.school] ? match[pattern.school].trim() : '';
                
                // Nettoyer l'école (enlever les parenthèses éventuelles)
                const cleanSchool = school.replace(/[()]/g, '').trim();
                
                return { guest, school: cleanSchool };
            }
        }
        
        // Si aucun pattern ne correspond, retourner le titre comme nom d'invité
        return { guest: title, school: '' };
    }
    
    /**
     * Lire les paramètres URL et pré-remplir le formulaire
     * Paramètres attendus depuis Monday :
     * - titre : Name (titre de la ligne Monday)
     * - format : FORMATS 2026
     * - date : Date de tournage (format YYYY-MM-DD)
     * - responsable : Auteurs
     * - tel : Téléphone du responsable (optionnel)
     * - lieu : Lieu (optionnel)
     */
    function loadFromUrlParams() {
        const urlParams = new URLSearchParams(window.location.search);
        
        // Titre (pour extraire nom et école)
        const titre = urlParams.get('titre');
        if (titre) {
            const { guest, school } = extractGuestAndSchool(decodeURIComponent(titre));
            if (guest) {
                document.getElementById('guestName').value = guest;
            }
            if (school) {
                document.getElementById('schoolName').value = school;
            }
        }
        
        // Format
        const format = urlParams.get('format');
        if (format) {
            const formatSelect = document.getElementById('formatType');
            const decodedFormat = decodeURIComponent(format);
            
            // Mapper les formats Monday vers les options du formulaire
            const formatMapping = {
                "L'interview": "L'interview",
                "Interview": "L'interview",
                "L'interro": "L'interro",
                "Interro": "L'interro",
                "C'est quoi?": "C'est quoi",
                "C'est quoi": "C'est quoi",
                "Audrey t'explique": "Audrey T'explique",
                "Audrey T'explique": "Audrey T'explique"
            };
            
            const mappedFormat = formatMapping[decodedFormat] || decodedFormat;
            
            // Chercher l'option correspondante
            for (let option of formatSelect.options) {
                if (option.value === mappedFormat || option.text === mappedFormat) {
                    formatSelect.value = option.value;
                    break;
                }
            }
        }
        
        // Date
        const date = urlParams.get('date');
        if (date) {
            try {
                const dateInput = document.getElementById('date');
                // Monday envoie la date au format ISO ou français
                let parsedDate;
                
                if (date.includes('-')) {
                    // Format ISO: YYYY-MM-DD
                    parsedDate = date;
                } else if (date.includes('/')) {
                    // Format français: DD/MM/YYYY
                    const parts = date.split('/');
                    parsedDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                }
                
                if (parsedDate) {
                    dateInput.value = parsedDate;
                }
            } catch (e) {
                console.warn('Impossible de parser la date:', date);
            }
        }
        
        // Responsable
        const responsable = urlParams.get('responsable');
        if (responsable) {
            const decodedResponsable = decodeURIComponent(responsable);
            document.getElementById('managerName').value = decodedResponsable;
            
            // Essayer de retrouver le téléphone si le responsable existe déjà
            const managers = getManagers();
            const existingManager = managers.find(m => 
                m.name.toLowerCase() === decodedResponsable.toLowerCase()
            );
            if (existingManager) {
                document.getElementById('managerPhone').value = existingManager.phone;
            }
        }
        
        // Téléphone (si fourni explicitement)
        const tel = urlParams.get('tel');
        if (tel) {
            document.getElementById('managerPhone').value = decodeURIComponent(tel);
        }
        
        // Lieu
        const lieu = urlParams.get('lieu');
        if (lieu && lieu.trim() !== '') {
            const decodedLieu = decodeURIComponent(lieu);
            document.getElementById('isExterior').checked = true;
            document.getElementById('addressField').style.display = 'block';
            document.getElementById('exteriorAddress').value = decodedLieu;
        }
        
        // Heure PAT
        const heure = urlParams.get('heure');
        if (heure) {
            document.getElementById('patTime').value = decodeURIComponent(heure);
        }
    }
    
    // Charger les paramètres URL au démarrage
    loadFromUrlParams();
    
    // ===== FIN DU CODE MONDAY - REPRISE DU CODE ORIGINAL =====
    
    // ===== GESTION DES RESPONSABLES DE PROJET =====
    
    // Clé pour le localStorage
    const STORAGE_KEY = 'easyCallSheets_managers';
    
    // Récupérer les responsables depuis le localStorage
    function getManagers() {
        const stored = localStorage.getItem(STORAGE_KEY);
        return stored ? JSON.parse(stored) : [];
    }
    
    // Sauvegarder les responsables dans le localStorage
    function saveManagers(managers) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(managers));
    }
    
    // Ajouter un responsable
    function addManager(name, phone) {
        if (!name || !phone) {
            alert('Veuillez remplir tous les champs');
            return;
        }
        
        const managers = getManagers();
        
        // Vérifier si le responsable existe déjà
        const existingIndex = managers.findIndex(m => m.name.toLowerCase() === name.toLowerCase());
        if (existingIndex !== -1) {
            if (confirm('Ce responsable existe déjà. Voulez-vous mettre à jour son numéro ?')) {
                managers[existingIndex].phone = phone;
                saveManagers(managers);
                refreshManagersList();
                return;
            } else {
                return;
            }
        }
        
        managers.push({ name: name.trim(), phone: phone.trim() });
        // Trier par nom
        managers.sort((a, b) => a.name.localeCompare(b.name));
        saveManagers(managers);
        refreshManagersList();
        
        // Vider les champs
        newManagerNameInput.value = '';
        newManagerPhoneInput.value = '';
    }
    
    // Supprimer un responsable
    function deleteManager(index) {
        if (confirm('Êtes-vous sûr de vouloir supprimer ce responsable ?')) {
            const managers = getManagers();
            managers.splice(index, 1);
            saveManagers(managers);
            refreshManagersList();
        }
    }
    
    // Afficher la liste des responsables dans la modal
    function refreshManagersList() {
        const managers = getManagers();
        managersList.innerHTML = '';
        
        if (managers.length === 0) {
            managersList.innerHTML = '<li class="empty-message">Aucun responsable sauvegardé</li>';
            return;
        }
        
        managers.forEach((manager, index) => {
            const li = document.createElement('li');
            li.innerHTML = `
                <span class="manager-info">
                    <strong>${manager.name}</strong> - ${manager.phone}
                </span>
                <button type="button" class="delete-button" data-index="${index}">Supprimer</button>
            `;
            managersList.appendChild(li);
        });
        
        // Ajouter les événements de suppression
        managersList.querySelectorAll('.delete-button').forEach(button => {
            button.addEventListener('click', function() {
                const index = parseInt(this.getAttribute('data-index'));
                deleteManager(index);
            });
        });
    }
    
    // Autocomplétion pour le nom du responsable
    function showSuggestions(query) {
        if (!query || query.length < 1) {
            managerSuggestions.innerHTML = '';
            managerSuggestions.style.display = 'none';
            return;
        }
        
        const managers = getManagers();
        const filtered = managers.filter(manager => 
            manager.name.toLowerCase().includes(query.toLowerCase())
        );
        
        if (filtered.length === 0) {
            managerSuggestions.innerHTML = '';
            managerSuggestions.style.display = 'none';
            return;
        }
        
        managerSuggestions.innerHTML = '';
        filtered.forEach(manager => {
            const div = document.createElement('div');
            div.className = 'suggestion-item';
            div.textContent = `${manager.name} - ${manager.phone}`;
            div.addEventListener('click', function() {
                managerNameInput.value = manager.name;
                managerPhoneInput.value = manager.phone;
                managerSuggestions.innerHTML = '';
                managerSuggestions.style.display = 'none';
            });
            managerSuggestions.appendChild(div);
        });
        
        managerSuggestions.style.display = 'block';
    }
    
    // Sauvegarder automatiquement quand le formulaire est soumis
    function autoSaveManager() {
        const name = managerNameInput.value.trim();
        const phone = managerPhoneInput.value.trim();
        
        if (name && phone) {
            const managers = getManagers();
            const exists = managers.some(m => 
                m.name.toLowerCase() === name.toLowerCase() && m.phone === phone
            );
            
            if (!exists) {
                // Vérifier si le nom existe déjà avec un autre numéro
                const existingIndex = managers.findIndex(m => m.name.toLowerCase() === name.toLowerCase());
                if (existingIndex === -1) {
                    // Nouveau responsable, l'ajouter
                    managers.push({ name, phone });
                    managers.sort((a, b) => a.name.localeCompare(b.name));
                    saveManagers(managers);
                }
            }
        }
    }
    
    // Événements pour l'autocomplétion
    managerNameInput.addEventListener('input', function() {
        showSuggestions(this.value);
    });
    
    // Fermer les suggestions quand on clique ailleurs
    document.addEventListener('click', function(e) {
        if (!managerNameInput.contains(e.target) && !managerSuggestions.contains(e.target)) {
            managerSuggestions.style.display = 'none';
        }
    });
    
    // Événements pour la modal
    manageManagersButton.addEventListener('click', function() {
        refreshManagersList();
        managersModal.style.display = 'block';
    });
    
    modalClose.addEventListener('click', function() {
        managersModal.style.display = 'none';
    });
    
    window.addEventListener('click', function(event) {
        if (event.target === managersModal) {
            managersModal.style.display = 'none';
        }
    });
    
    addManagerButton.addEventListener('click', function() {
        const name = newManagerNameInput.value.trim();
        const phone = newManagerPhoneInput.value.trim();
        addManager(name, phone);
    });
    
    // Import CSV
    const importCsvButton = document.getElementById('importCsvButton');
    const csvFileInput = document.getElementById('csvFile');
    
    importCsvButton.addEventListener('click', function() {
        const file = csvFileInput.files[0];
        if (!file) {
            alert('Veuillez sélectionner un fichier CSV');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const text = e.target.result;
                const lines = text.split('\n');
                let imported = 0;
                
                const managers = getManagers();
                
                lines.forEach(line => {
                    const parts = line.split(',').map(p => p.trim());
                    if (parts.length >= 2 && parts[0] && parts[1]) {
                        const name = parts[0];
                        const phone = parts[1];
                        
                        // Vérifier si existe déjà
                        const existingIndex = managers.findIndex(m => 
                            m.name.toLowerCase() === name.toLowerCase()
                        );
                        
                        if (existingIndex === -1) {
                            managers.push({ name, phone });
                            imported++;
                        }
                    }
                });
                
                if (imported > 0) {
                    managers.sort((a, b) => a.name.localeCompare(b.name));
                    saveManagers(managers);
                    refreshManagersList();
                    alert(`${imported} responsable(s) importé(s) avec succès !`);
                } else {
                    alert('Aucun nouveau responsable à importer.');
                }
                
                csvFileInput.value = '';
                
            } catch (error) {
                alert('Erreur lors de la lecture du fichier CSV. Vérifiez le format.');
                console.error(error);
            }
        };
        
        reader.readAsText(file);
    });
    
    // Export CSV
    const exportCsvButton = document.getElementById('exportCsvButton');
    exportCsvButton.addEventListener('click', function() {
        const managers = getManagers();
        
        if (managers.length === 0) {
            alert('Aucun responsable à exporter.');
            return;
        }
        
        // Créer le CSV
        let csvContent = '';
        managers.forEach(manager => {
            csvContent += `${manager.name},${manager.phone}\n`;
        });
        
        // Créer un blob et télécharger
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', 'responsables_easycallsheets.csv');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
    
    // ===== GESTION DU FORMULAIRE =====
    
    // Gestion du checkbox tournage extérieur
    const isExteriorCheckbox = document.getElementById('isExterior');
    const addressField = document.getElementById('addressField');
    
    isExteriorCheckbox.addEventListener('change', function() {
        if (this.checked) {
            addressField.style.display = 'block';
        } else {
            addressField.style.display = 'none';
        }
    });
    
    // Gestion du checkbox horaires personnalisés
    const customScheduleCheckbox = document.getElementById('customSchedule');
    const customScheduleFields = document.getElementById('customScheduleFields');
    
    customScheduleCheckbox.addEventListener('change', function() {
        if (this.checked) {
            customScheduleFields.style.display = 'block';
        } else {
            customScheduleFields.style.display = 'none';
        }
    });
    
    // Fonction pour formater la date
    function formatDate(dateString) {
        if (!dateString) return '';
        
        const date = new Date(dateString);
        const days = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
        const months = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
        
        const dayName = days[date.getDay()];
        const day = date.getDate();
        const month = months[date.getMonth()];
        const year = date.getFullYear();
        
        return `${dayName} ${day} ${month} ${year}`.toUpperCase();
    }
    
    // Fonction pour formater l'heure
    function formatTime(timeString) {
        if (!timeString) return '';
        return timeString;
    }
    
    // Fonction pour calculer les horaires par défaut
    function calculateSchedule(patTime, isCustom, customTimes) {
        if (isCustom && customTimes.install) {
            return {
                install: customTimes.install,
                hmcStart: customTimes.hmcStart,
                hmcEnd: customTimes.hmcEnd,
                rdv: customTimes.rdv,
                end: customTimes.end
            };
        }
        
        // Horaires par défaut basés sur l'heure PAT
        const [hours, minutes] = patTime.split(':').map(Number);
        
        // Installation : PAT - 1h30
        const installDate = new Date();
        installDate.setHours(hours, minutes);
        installDate.setMinutes(installDate.getMinutes() - 90);
        const install = `${String(installDate.getHours()).padStart(2, '0')}:${String(installDate.getMinutes()).padStart(2, '0')}`;
        
        // HMC Début : PAT - 30min
        const hmcStartDate = new Date();
        hmcStartDate.setHours(hours, minutes);
        hmcStartDate.setMinutes(hmcStartDate.getMinutes() - 30);
        const hmcStart = `${String(hmcStartDate.getHours()).padStart(2, '0')}:${String(hmcStartDate.getMinutes()).padStart(2, '0')}`;
        
        // HMC Fin : PAT + 15min
        const hmcEndDate = new Date();
        hmcEndDate.setHours(hours, minutes);
        hmcEndDate.setMinutes(hmcEndDate.getMinutes() + 15);
        const hmcEnd = `${String(hmcEndDate.getHours()).padStart(2, '0')}:${String(hmcEndDate.getMinutes()).padStart(2, '0')}`;
        
        // RDV : PAT
        const rdv = patTime;
        
        // Fin : PAT + 2h30
        const endDate = new Date();
        endDate.setHours(hours, minutes);
        endDate.setMinutes(endDate.getMinutes() + 150);
        const end = `${String(endDate.getHours()).padStart(2, '0')}:${String(endDate.getMinutes()).padStart(2, '0')}`;
        
        return { install, hmcStart, hmcEnd, rdv, end };
    }
    
    // Fonction pour générer le texte du déroulé
    function generateScheduleText(schedule, patTime, guestName, isExterior) {
        const locationNote = isExterior ? '(sur place)' : '(à L\'Etudiant)';
        
        return `
            <p><strong>${schedule.install || '08:00'}</strong> : Installation du matériel ${locationNote}</p>
            <p><strong>${schedule.hmcStart || '08:30'}</strong> : Début HMC</p>
            <p><strong>${schedule.hmcEnd || '09:15'}</strong> : Fin HMC</p>
            <p><strong>${schedule.rdv || patTime}</strong> : Arrivée de ${guestName || 'l\'invité'}</p>
            <p style="margin-left: 80px;">- Accueil / installation / passage en loges</p>
            <p style="margin-left: 80px;">- Tournage (entre 15 et 30 minutes)</p>
            <p><strong>${schedule.end || '11:30'}</strong> : Fin de tournage et rangement</p>
        `.trim();
    }
    
    // Fonction pour générer l'aperçu
    function generatePreview() {
        const formatType = document.getElementById('formatType').value;
        const date = document.getElementById('date').value;
        const guestName = document.getElementById('guestName').value;
        const schoolName = document.getElementById('schoolName').value;
        const patTime = document.getElementById('patTime').value;
        const managerName = document.getElementById('managerName').value;
        const managerPhone = document.getElementById('managerPhone').value;
        const isCustomSchedule = document.getElementById('customSchedule').checked;
        const isExterior = document.getElementById('isExterior').checked;
        const exteriorAddress = document.getElementById('exteriorAddress').value;
        
        // Récupérer les horaires personnalisés si activés
        const customTimes = isCustomSchedule ? {
            install: document.getElementById('customInstall').value,
            hmcStart: document.getElementById('customHmcStart').value,
            hmcEnd: document.getElementById('customHmcEnd').value,
            rdv: document.getElementById('customRdv').value,
            end: document.getElementById('customEnd').value
        } : {};
        
        const schedule = calculateSchedule(patTime, isCustomSchedule, customTimes);
        
        // Générer le HTML de l'aperçu
        const previewHTML = `
            <div class="pdf-content">
                <div class="pdf-header">
                    <div class="pdf-logo-text">L'Étudiant</div>
                    <h2>FEUILLE DE SERVICE - ${formatDate(date)}</h2>
                </div>
                
                <br><br>
                
                <div class="pdf-section">
                    <h3>FORMAT : "${formatType.toUpperCase()}"</h3>
                    <p>INVITÉ : ${guestName}${schoolName ? ' - ' + schoolName : ''}</p>
                </div>
                
                <div class="pdf-section">
                    <h3>LIEU DE RDV :</h3>
                    ${isExterior && exteriorAddress ? 
                        exteriorAddress.split('\n').map(line => `<p>${line}</p>`).join('') :
                        `<p>L'Etudiant, Carré Daumesnil</p>
                        <p>52, rue Jacques-Hillairet - 75012 PARIS</p>`
                    }
                </div>
                
                <div class="pdf-section">
                    <h3>HEURE DE RDV : ${schedule.rdv ? formatTime(schedule.rdv) : formatTime(patTime)}</h3>
                </div>
                
                <div class="pdf-section">
                    <h3>CONTACTS :</h3>
                    <p>${managerName} (Responsable projet) - ${managerPhone}</p>
                    <p>Martin Pavloff (Responsable vidéo) - 06 12 52 85 69</p>
                </div>
                
                <div class="pdf-section">
                    <h3>DÉROULÉ DE LA JOURNÉE :</h3>
                    ${generateScheduleText(schedule, patTime, guestName, isExterior)}
                </div>
                
                <div class="pdf-section">
                    <h3>NOTE AUX INTERVENANTS :</h3>
                    <p>• Évitez les vêtements avec marques apparentes, les logos, les carreaux et les rayures</p>
                    <p>• Nous tournons parfois sur fond vert (qui est notre couleur d'incrustation), merci donc de ne pas porter de vert (au risque de vous fondre dans le décor)</p>
                    <p>• Si vous portez des lunettes, dans les mesures du possible, merci de privilégier les lentilles de contact pour notre tournage</p>
                </div>
            </div>
        `;
        
        previewContent.innerHTML = previewHTML;
    }
    
    // Fonction pour générer le PDF
    function generatePDF() {
        const formatType = document.getElementById('formatType').value;
        const date = document.getElementById('date').value;
        const guestName = document.getElementById('guestName').value;
        const schoolName = document.getElementById('schoolName').value;
        const patTime = document.getElementById('patTime').value;
        const managerName = document.getElementById('managerName').value;
        const managerPhone = document.getElementById('managerPhone').value;
        const isCustomSchedule = document.getElementById('customSchedule').checked;
        const isExterior = document.getElementById('isExterior').checked;
        const exteriorAddress = document.getElementById('exteriorAddress').value;
        
        // Validation
        if (!date || !guestName || !patTime || !managerName || !managerPhone) {
            alert('Veuillez remplir tous les champs obligatoires.');
            return;
        }
        
        // Récupérer les horaires personnalisés si activés
        const customTimes = isCustomSchedule ? {
            install: document.getElementById('customInstall').value,
            hmcStart: document.getElementById('customHmcStart').value,
            hmcEnd: document.getElementById('customHmcEnd').value,
            rdv: document.getElementById('customRdv').value,
            end: document.getElementById('customEnd').value
        } : {};
        
        const schedule = calculateSchedule(patTime, isCustomSchedule, customTimes);
        
        // Nom du fichier
        const fileName = `CallSheet_${formatType.replace(/'/g, '')}_${guestName.replace(/\s+/g, '_')}_${date}.pdf`;
        
        // Afficher un message de chargement
        const loadingMessage = document.createElement('div');
        loadingMessage.textContent = 'Génération du PDF en cours...';
        loadingMessage.style.position = 'fixed';
        loadingMessage.style.top = '50%';
        loadingMessage.style.left = '50%';
        loadingMessage.style.transform = 'translate(-50%, -50%)';
        loadingMessage.style.padding = '20px';
        loadingMessage.style.background = 'rgba(0,0,0,0.7)';
        loadingMessage.style.color = 'white';
        loadingMessage.style.borderRadius = '5px';
        loadingMessage.style.zIndex = '9999';
        document.body.appendChild(loadingMessage);
        
        try {
            // Créer une nouvelle fenêtre pour l'impression
            const printWindow = window.open('', '_blank');
            
            if (!printWindow) {
                alert("Votre navigateur a bloqué l'ouverture d'une nouvelle fenêtre. Veuillez autoriser les popups pour ce site.");
                document.body.removeChild(loadingMessage);
                return;
            }
            
            // Obtenir le chemin absolu du logo
            const logoPath = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1) + 'logo_etudiant.png';
            
            // Générer le HTML complet avec les styles intégrés
            const printContent = `
                <!DOCTYPE html>
                <html lang="fr">
                <head>
                    <meta charset="UTF-8">
                    <title>${fileName}</title>
                    <style>
                        @page {
                            size: A4;
                            margin: 20mm;
                        }
                        * {
                            margin: 0;
                            padding: 0;
                            box-sizing: border-box;
                        }
                        body {
                            font-family: 'Nunito', Arial, sans-serif;
                            margin: 0;
                            padding: 0;
                            color: #000;
                            line-height: 1.6;
                            background-color: white;
                        }
                        .container {
                            width: 100%;
                            max-width: 210mm;
                            margin: 0 auto;
                            padding: 0;
                        }
                        .header {
                            text-align: center;
                            margin-bottom: 20px;
                        }
                        .logo {
                            max-width: 210px;
                            height: auto;
                            display: block;
                            margin: 0 auto 15px;
                        }
                        h1 {
                            font-size: 20px;
                            margin: 10px 0;
                            font-weight: bold;
                        }
                        h2 {
                            font-size: 16px;
                            margin: 15px 0 5px;
                            font-weight: bold;
                        }
                        p {
                            margin: 5px 0;
                        }
                        .section {
                            margin-bottom: 20px;
                            page-break-inside: avoid;
                        }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="header">
                            <img src="${logoPath}" alt="Logo L'Etudiant" class="logo" onerror="this.style.display='none'">
                            <h1>FEUILLE DE SERVICE - ${formatDate(date)}</h1>
                        </div>
                        <br><br>
                        <div class="section">
                            <h2>FORMAT : "${formatType.toUpperCase()}"</h2>
                            <p>INVITÉ : ${guestName}${schoolName ? ' - ' + schoolName : ''}</p>
                        </div>
                        
                        <div class="section">
                            <h2>LIEU DE RDV :</h2>
                            ${isExterior && exteriorAddress ? 
                                exteriorAddress.split('\n').map(line => `<p>${line}</p>`).join('') :
                                `<p>L'Etudiant, Carré Daumesnil</p>
                                <p>52, rue Jacques-Hillairet - 75012 PARIS</p>`
                            }
                        </div>
                        
                        <div class="section">
                            <h2>HEURE DE RDV : ${schedule.rdv ? formatTime(schedule.rdv) : formatTime(patTime)}</h2>
                        </div>
                        
                        <div class="section">
                            <h2>CONTACTS :</h2>
                            <p>${managerName} (Responsable projet) - ${managerPhone}</p>
                            <p>Martin Pavloff (Responsable vidéo) - 06 12 52 85 69</p>
                        </div>
                        
                        <div class="section">
                            <h2>DÉROULÉ DE LA JOURNÉE :</h2>
                            ${generateScheduleText(schedule, patTime, guestName, isExterior)}
                        </div>
                        
                        <div class="section">
                            <h2>NOTE AUX INTERVENANTS :</h2>
                            <p>• Évitez les vêtements avec marques apparentes, les logos, les carreaux et les rayures</p>
                            <p>• Nous tournons parfois sur fond vert (qui est notre couleur d'incrustation), merci donc de ne pas porter de vert (au risque de vous fondre dans le décor)</p>
                            <p>• Si vous portez des lunettes, dans les mesures du possible, merci de privilégier les lentilles de contact pour notre tournage</p>
                        </div>
                    </div>

                    <script>
                        // Attendre que l'image soit chargée avant d'imprimer
                        var logo = document.querySelector('.logo');
                        var printed = false;
                        
                        function printAndClose() {
                            if (!printed) {
                                printed = true;
                                window.print();
                                setTimeout(function() {
                                    window.close();
                                }, 1000);
                            }
                        }
                        
                        // Si l'image est déjà chargée
                        if (logo && logo.complete) {
                            setTimeout(printAndClose, 500);
                        } else if (logo) {
                            // Sinon attendre le chargement de l'image
                            logo.onload = function() {
                                setTimeout(printAndClose, 500);
                            };
                            
                            // En cas d'erreur de chargement de l'image, continuer quand même
                            logo.onerror = function() {
                                setTimeout(printAndClose, 500);
                            };
                        } else {
                            // Pas de logo, imprimer directement
                            setTimeout(printAndClose, 500);
                        }
                        
                        // Fallback en cas de problème
                        setTimeout(printAndClose, 2000);
                    </script>
                </body>
                </html>
            `;
            
            // Écrire le contenu dans la nouvelle fenêtre
            printWindow.document.open();
            printWindow.document.write(printContent);
            printWindow.document.close();
            
            // Supprimer le message de chargement
            document.body.removeChild(loadingMessage);
            
            // Sauvegarder automatiquement le responsable
            autoSaveManager();
            
        } catch (error) {
            console.error('Erreur lors de la génération du PDF:', error);
            document.body.removeChild(loadingMessage);
            alert('Une erreur est survenue lors de la génération du PDF. Veuillez réessayer.');
        }
    }
    
    // Événements
    previewButton.addEventListener('click', generatePreview);
    generatePdfButton.addEventListener('click', generatePDF);
});
