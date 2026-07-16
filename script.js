// ============================================================
//  script.js - Charge et affiche les données du fichier Excel
// ============================================================

// Configuration : nom du fichier Excel dans le même dossier
const EXCEL_FILE = 'donnees.xlsx';

// Éléments DOM
const tableBody = document.getElementById('tableBody');
const loadingMsg = document.getElementById('loadingMessage');
const totalBalanceEl = document.getElementById('totalBalance');
const totalIncomeEl = document.getElementById('totalIncome');
const totalExpenseEl = document.getElementById('totalExpense');
const totalCountEl = document.getElementById('totalCount');
const lastUpdateEl = document.getElementById('lastUpdate');

// Éléments des filtres
const searchInput = document.getElementById('searchInput');
const dateFrom = document.getElementById('dateFrom');
const dateTo = document.getElementById('dateTo');
const resetBtn = document.getElementById('resetFilters');

// ============================================================
//  1.  Chargement du fichier Excel
// ============================================================

async function loadExcelData() {
    try {
        const response = await fetch(EXCEL_FILE);
        if (!response.ok) {
            throw new Error(`Impossible de charger ${EXCEL_FILE} (HTTP ${response.status})`);
        }
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // On prend la première feuille
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        // Conversion en tableau d'objets (première ligne = en-têtes)
        const data = XLSX.utils.sheet_to_json(firstSheet);

        if (!data || data.length === 0) {
            throw new Error('Le fichier Excel est vide ou mal formaté.');
        }

        return data;
    } catch (error) {
        console.error('Erreur de chargement :', error);
        loadingMsg.innerHTML = `<i class="fas fa-exclamation-triangle" style="color:#ef4444;"></i> 
                                Erreur : ${error.message}. 
                                Vérifiez que le fichier "donnees.xlsx" est bien dans le même dossier.`;
        return null;
    }
}

// ============================================================
//  2.  Nettoyage et transformation des données
// ============================================================

function cleanData(rawData) {
    return rawData.map(row => {
        // On récupère les valeurs en s'adaptant aux noms de colonnes possibles
        const date = row['A'] || row['Date'] || row['date'] || '';
        const designation = row['B'] || row['Désignation'] || row['designation'] || '';
        const entrees = parseFloat(row['C'] || row['Entrées'] || row['entrees'] || 0) || 0;
        const sorties = parseFloat(row['D'] || row['Sorties'] || row['sorties'] || 0) || 0;
        const solde = parseFloat(row['E'] || row['Solde'] || row['solde'] || 0) || 0;
        const observations = row['F'] || row['Observations'] || row['observations'] || '';

        // Formatage de la date
        let formattedDate = date;
        if (date && typeof date === 'string') {
            // Si c'est une chaîne, on essaie de la nettoyer
            const parts = date.split(/[-/.]/);
            if (parts.length === 3) {
                // Supposons que le format est JJ/MM/AAAA ou AAAA-MM-JJ
                if (parts[0].length === 4) {
                    formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                } else {
                    formattedDate = `${parts[0]}/${parts[1]}/${parts[2]}`;
                }
            }
        } else if (date && typeof date === 'number') {
            // Si c'est un nombre (timestamp Excel), on le convertit
            const d = new Date((date - 25569) * 86400 * 1000);
            formattedDate = d.toLocaleDateString('fr-FR');
        }

        return {
            date: formattedDate,
            designation: designation || 'Sans désignation',
            entrees: entrees,
            sorties: sorties,
            solde: solde,
            observations: observations || '-'
        };
    });
}

// ============================================================
//  3.  Affichage du tableau
// ============================================================

function renderTable(data) {
    if (!data || data.length === 0) {
        tableBody.innerHTML = `
            <tr>
                <td colspan="6" style="text-align:center; padding:40px; color:#94a3b8;">
                    <i class="fas fa-inbox" style="font-size:2rem; display:block; margin-bottom:10px;"></i>
                    Aucune donnée à afficher.
                </td>
            </tr>
        `;
        return;
    }

    let html = '';
    data.forEach(row => {
        const soldeClass = row.solde >= 0 ? 'balance-positive' : 'balance-negative';
        html += `
            <tr>
                <td><i class="fas fa-calendar-alt" style="color:#94a3b8; margin-right:6px;"></i>${row.date}</td>
                <td>${escapeHtml(row.designation)}</td>
                <td class="entry-amount">${row.entrees > 0 ? formatPrice(row.entrees) : '-'}</td>
                <td class="exit-amount">${row.sorties > 0 ? formatPrice(row.sorties) : '-'}</td>
                <td class="balance-amount ${soldeClass}">${formatPrice(row.solde)}</td>
                <td style="font-size:0.9rem; color:#64748b;">${escapeHtml(row.observations)}</td>
            </tr>
        `;
    });

    tableBody.innerHTML = html;
}

// ============================================================
//  4.  Mise à jour du résumé
// ============================================================

function updateSummary(data) {
    if (!data || data.length === 0) {
        totalBalanceEl.textContent = '0 €';
        totalIncomeEl.textContent = '0 €';
        totalExpenseEl.textContent = '0 €';
        totalCountEl.textContent = '0';
        return;
    }

    let totalIncome = 0;
    let totalExpense = 0;
    let totalBalance = 0;

    data.forEach(row => {
        totalIncome += row.entrees;
        totalExpense += row.sorties;
        totalBalance += row.solde;
    });

    totalIncomeEl.textContent = formatPrice(totalIncome);
    totalExpenseEl.textContent = formatPrice(totalExpense);
    totalBalanceEl.textContent = formatPrice(totalBalance);
    totalCountEl.textContent = data.length;
}

// ============================================================
//  5.  Filtres (recherche + période)
// ============================================================

function applyFilters(data) {
    const searchTerm = searchInput.value.toLowerCase().trim();
    const fromDate = dateFrom.value ? new Date(dateFrom.value + 'T00:00:00') : null;
    const toDate = dateTo.value ? new Date(dateTo.value + 'T23:59:59') : null;

    return data.filter(row => {
        // Filtre texte sur la désignation + observations
        let matchText = true;
        if (searchTerm) {
            const text = (row.designation + ' ' + row.observations).toLowerCase();
            matchText = text.includes(searchTerm);
        }

        // Filtre date (on parse la date stockée en JJ/MM/AAAA)
        let matchDate = true;
        if (fromDate || toDate) {
            const parts = row.date.split('/');
            if (parts.length === 3) {
                const rowDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
                if (fromDate && rowDate < fromDate) matchDate = false;
                if (toDate && rowDate > toDate) matchDate = false;
            } else {
                matchDate = false; // date non reconnue
            }
        }

        return matchText && matchDate;
    });
}

// ============================================================
//  6.  Fonctions utilitaires
// ============================================================

function formatPrice(value) {
    return new Intl.NumberFormat('fr-FR', {
        style: 'currency',
        currency: 'EUR',
        minimumFractionDigits: 2
    }).format(value);
}

// Échapper les caractères HTML pour éviter les injections
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============================================================
//  7.  Initialisation
// ============================================================

let allData = [];

async function init() {
    loadingMsg.style.display = 'block';
    const rawData = await loadExcelData();

    if (!rawData) {
        loadingMsg.style.display = 'none';
        return;
    }

    // Nettoyer les données
    allData = cleanData(rawData);

    // Mettre à jour la date de dernière mise à jour
    const now = new Date();
    lastUpdateEl.textContent = now.toLocaleString('fr-FR');

    // Afficher les données (avec filtres appliqués)
    const filtered = applyFilters(allData);
    renderTable(filtered);
    updateSummary(filtered);

    loadingMsg.style.display = 'none';

    // Activer les filtres
    searchInput.addEventListener('input', () => refreshDisplay());
    dateFrom.addEventListener('change', () => refreshDisplay());
    dateTo.addEventListener('change', () => refreshDisplay());
    resetBtn.addEventListener('click', resetFilters);
}

function refreshDisplay() {
    const filtered = applyFilters(allData);
    renderTable(filtered);
    updateSummary(filtered);
}

function resetFilters() {
    searchInput.value = '';
    dateFrom.value = '';
    dateTo.value = '';
    refreshDisplay();
}

// Lancer l'application
init();