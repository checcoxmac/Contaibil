// 🚀 VERSIONE LOGIC.JS - BUILD 2025-12-10 17:30 🚀
console.log("🚀🚀🚀 LOGIC.JS CARICATO - BUILD 10/12/2025 17:30 🚀🚀🚀");
console.log("✅ Fix ADE indexOf() attivo");
console.log("✅ Fix date normalizeDateItalyStrict attivo");
console.log("✅ Fix CA Bank quote handling attivo");
console.log("✅ Debug logging completo attivo");

window.addEventListener("DOMContentLoaded", () => {
  // --- WIDE MODE TOGGLE ---
  const btnWideMode = document.getElementById('btnWideMode');
  if (btnWideMode) {
    // Carica preferenza salvata
    const savedWideMode = localStorage.getItem('contaibl-wide-mode') === 'true';
    if (savedWideMode) {
      document.body.classList.add('wide-mode');
      btnWideMode.classList.add('active');
    }
    
    btnWideMode.addEventListener('click', () => {
      document.body.classList.toggle('wide-mode');
      btnWideMode.classList.toggle('active');
      localStorage.setItem('contaibl-wide-mode', document.body.classList.contains('wide-mode'));
    });
  }
  // --- COMPACT MODE TOGGLE (Excel-like) ---
  const btnCompactMode = document.getElementById('btnCompactMode');
  if (btnCompactMode) {
    const savedCompact = localStorage.getItem('contaibl-compact-mode') === 'true';
    if (savedCompact) {
      document.body.classList.add('compact-mode');
      btnCompactMode.classList.add('active');
      // Aggiungi classe excel-view alle tabelle
      document.querySelectorAll('table').forEach(t => t.classList.add('excel-view'));
    }
    btnCompactMode.addEventListener('click', () => {
      const enabled = document.body.classList.toggle('compact-mode');
      btnCompactMode.classList.toggle('active');
      localStorage.setItem('contaibl-compact-mode', enabled);
      // Aggiungi/rimuovi classe excel-view dalle tabelle
      document.querySelectorAll('table').forEach(t => {
        if (enabled) t.classList.add('excel-view');
        else t.classList.remove('excel-view');
      });
    });
  }
  // ============================================================
  // 🔧 CONFIGURAZIONE TOLLERANZE MATCHING (Dic 2025)
  // ============================================================
  const MATCH_CONFIG = {
    // Tolleranze importi
    TOLERANCE_EXACT: 0.01,           // ±1 centesimo
    TOLERANCE_ROUNDING: 0.05,        // ±5 centesimi (arrotondamenti)
    TOLERANCE_L2_AMOUNT: 1.00,       // ±1 euro (L2 match)
    TOLERANCE_L4_AMOUNT: 0.50,       // ±50 centesimi (L4 match)
    TOLERANCE_L5_AMOUNT: 50.00,      // ±50 euro (L5 match rilassato - DA REVISIONARE)
    TOLERANCE_STAMP_DUTY: 2.00,      // Marca da bollo

    // Tolleranza forte per match NUM+DATA+TOT
    TOL_TOT: 0.05,                   // ±5 centesimi
    
    // Soglie per classificazione differenze
    THRESHOLD_SIGNIFICANT_DIFF: 100, // Oltre 100€ = differenza significativa
    THRESHOLD_LOW_TOTAL: 1.00,       // Sotto 1€ = totale quasi zero
    
    // Tolleranze date e numeri
    DATE_TOLERANCE_DAYS: 3,          // ±3 giorni per match
    INVOICE_MIN_DIGITS: 4,           // Minimo cifre per fuzzy match numeri
    
    // Fuzzy matching nomi
    FUZZY_SAME_PIVA: 2,              // Levenshtein con stessa P.IVA
    FUZZY_DIFF_PIVA: 3,              // Levenshtein con P.IVA diverse
    
    // Lunghezza P.IVA italiana
    PIVA_IT_LENGTH: 11,
    PIVA_IT_MIN_VALUE: 10000000000   // 10 miliardi (per riconoscere P.IVA come numero)
  };

  // ============================================================
  // 📊 STATO GLOBALE APPLICAZIONE
  // ============================================================
  let lastAdeRecords = [];
  let lastGestRecords = [];
  let lastResults = [];
  let currentFilter = "ALL";       // filtro per stato (Tutti / MATCH_OK / ...)
  let totalAdeCount = 0;
  let totalGestCount = 0;
  let currentMonthFilter = null;   // filtro per mese ADE (es. "2025-03" oppure null = tutti)
  let currentSupplierFilter = "";  // filtro per fornitore (denominazione / P.IVA)
  let currentInvoiceFilter = "";   // nuovo filtro per NUM fattura
  let filterInvoiceInput = null;   // riferimento all’input HTML del numero fattura
  let pendingMatchSourceId = null;
  let pendingMatchSourceSide = null;
  let nextResultId = 1;
  let originalResults = null;
  let undoStack = [];

  // Stato flusso guidato / stepper
  const flowState = {
    files: { ade: false, gest: false, nc: false },
    mapping: { gestBConfirmed: false, gestCConfirmed: false },
    matchDone: false,
    correctionsDone: false,
    exportDone: false
  };
  const mappingSession = {
    GEST_B: { confirmed: false, headerSig: null },
    GEST_C: { confirmed: false, headerSig: null }
  };
  const cachedHeaders = { GEST_B: null, GEST_C: null };

  const DEFAULT_PROCEDURE_FILENAME = "procedure_STANDARD.json";

  // analisi fornitori
  let analysisMinDate = null;
  let analysisMaxDate = null;
  let lastSupplierAnalysis = [];

  // editor
  let currentEditRowId = null;

  // ---------- utils ----------
  // modules/ThemeManager.js
const ThemeManager = {
    THEME_KEY: 'user-theme-preference',
    DARK_CLASS: 'dark-mode',
    
    init() {
        // 1. Controlla LocalStorage
        const storedTheme = localStorage.getItem(this.THEME_KEY);
        
        // 2. Controlla preferenza OS
        const sysPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        
        if (storedTheme === 'dark' || (!storedTheme && sysPrefersDark)) {
            this.enableDarkMode();
        } else {
            this.disableDarkMode();
        }
        
        // 3. Listener attivo per cambi real-time del sistema
        window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', e => {
            if (!localStorage.getItem(this.THEME_KEY)) { // Solo se l'utente non ha forzato
                e.matches ? this.enableDarkMode() : this.disableDarkMode();
            }
        });
    },
    
    enableDarkMode() {
        document.documentElement.setAttribute('data-theme', 'dark');
        // Importante per componenti Apple-style: aggiorna meta tag status bar
        document.querySelector('meta[name="theme-color"]')?.setAttribute('content', '#000000');
    },
    
    disableDarkMode() {
        document.documentElement.removeAttribute('data-theme');
        document.querySelector('meta[name="theme-color"]')?.setAttribute('content', '#FFFFFF');
    },
    
    toggle() {
        if (document.documentElement.getAttribute('data-theme') === 'dark') {
            this.disableDarkMode();
            localStorage.setItem(this.THEME_KEY, 'light');
        } else {
            this.enableDarkMode();
            localStorage.setItem(this.THEME_KEY, 'dark');
        }
    }
};
  function getAdePart(r) {
    return r?.ADE || r?.a || null;
  }

  function getGestPart(r) {
    return r?.GEST || r?.g || null;
  }

  function refreshResultsView() {
    if (typeof applyFilterAndRender === "function") {
      applyFilterAndRender();
    } else if (typeof renderResultsTable === "function") {
      renderResultsTable();
    } else {
      console.error("Funzione di refresh della tabella non trovata.");
    }
  }

  function parseCSV(text) {
    const lines = text.split(/\r?\n/).filter(l => l.trim() !== "");
    if (!lines.length) return { headers: [], rows: [] };
    const sep = lines[0].includes(";") ? ";" : ",";
    const headers = lines[0].split(sep).map(h => {
      return h
        .replace(/^\uFEFF/, "")     // rimuove BOM UTF-8
        .replace(/\u00A0/g, " ")     // NBSP -> spazio
        .replace(/"/g, '')
        .trim();
    });
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split(sep);
      const row = {};
      headers.forEach((h, idx) => {
        row[h] = (parts[idx] || "").trim();
      });
      rows.push(row);
    }
    return { headers, rows };
  }

  // ============================================================
  // NORMALIZZAZIONE DENOMINAZIONI (Miglioramento Dic 2025)
  // ============================================================
  // Obiettivo: rendere confrontabili denominazioni scritte in modo diverso
  // Esempi: "Wind Tre S.p.A." ↔ "H3G S.p.A." (tramite P.IVA)
  //         "Over Security srl" ↔ "OVER SECURITY SRL"
  // Nota: la P.IVA lato gestionale NON è disponibile, quindi questa
  //       normalizzazione è fondamentale per i match basati su denominazione
  // ============================================================
  function normalizeName(s) {
  if (!s) return "";

  // 1) MAIUSCOLO + TRIM (standardizza case)
  let t = String(s).toUpperCase().trim();

  // 2) Rimuovi accenti (à -> A, è -> E, ecc.)
  try {
    t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  } catch (e) {}

  // 3) Uniforma le forme societarie con tutti i pattern possibili
  //    SRLS / S.R.L.S. / S.R.L. / SRL. / srl -> SRL
  t = t.replace(/S\.?\s?R\.?\s?L\.?\s?S?\.?/gi, " SRL ");
  //    SPA / S.P.A. / S P A / SpA -> SPA
  t = t.replace(/S\.?\s?P\.?\s?A\.?/gi, " SPA ");
  //    SAS / S.A.S. -> SAS
  t = t.replace(/S\.?\s?A\.?\s?S\.?/gi, " SAS ");
  //    SNC / S.N.C. -> SNC
  t = t.replace(/S\.?\s?N\.?\s?C\.?/gi, " SNC ");
  //    SS / S.S. -> SS
  t = t.replace(/\bS\.?\s?S\.?\b/gi, " SS ");

  // 4) Rimuovi parole "giuridiche" poco significative per il match
  t = t.replace(/\bSOCIETA['`]?\b/g, " ");
  t = t.replace(/\bUNIPERSONALE\b/g, " ");
  t = t.replace(/\bCOOP(ERATIVA)?\b/g, " ");
  t = t.replace(/\bA SOCIO UNICO\b/g, " ");
  t = t.replace(/\bIN LIQUIDAZIONE\b/g, " ");

  // 5) Rimuovi apostrofi, punteggiatura comune, trattini, barre
  t = t.replace(/['".,;:()\-\/\\]/g, " ");

  // 6) Rimuovi punti finali e doppi spazi residui
  t = t.replace(/\.$/g, "");

  // 7) Mantieni solo lettere, numeri e spazi
  t = t.replace(/[^A-Z0-9 ]/g, " ");

  // 8) Compatta spazi multipli in uno solo
  t = t.replace(/\s+/g, " ").trim();

  return t;
}

  function namesSimilar(a, b) {
    if (!a || !b) return false;
    const toParts = (s) => normalizeName(s).split(" ").filter(Boolean);
    const stopWords = new Set([
      "SRL","SRLS","SAS","SPA","SNC","COOP","COOPERATIVA","SOCIETA",
      "ARL","RL","SA","SCPA",
      "DI","DEI","DE","DEL","DELLA","DELLE","LO","LA","IL","THE"
    ]);
    const pa = toParts(a).filter(t => !stopWords.has(t));
    const pb = toParts(b).filter(t => !stopWords.has(t));
    if (!pa.length || !pb.length) return false;

    const setB = new Set(pb);
    let common = 0;
    let commonLong = 0;
    for (const t of pa) {
      if (setB.has(t)) {
        common++;
        if (t.length >= 4) commonLong++;
      }
    }

    if (commonLong >= 1) return true;
    if (common >= 2) return true;
    if (pa[0] === pb[0] && pa[0].length >= 4) return true;

    return false;
  }

  // Calcola quanto sono diverse due stringhe (0 = identiche)
  function levenshteinDistance(a, b) {
    const tmp = [];
    if (a.length === 0) { return b.length; }
    if (b.length === 0) { return a.length; }

    for (let i = 0; i <= b.length; i++) { tmp[i] = [i]; }
    for (let j = 0; j <= a.length; j++) { tmp[0][j] = j; }

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          tmp[i][j] = tmp[i - 1][j - 1];
        } else {
          tmp[i][j] = Math.min(tmp[i - 1][j - 1] + 1, Math.min(tmp[i][j - 1] + 1, tmp[i - 1][j] + 1));
        }
      }
    }
    return tmp[b.length][a.length];
  }

  // Funzione helper per decidere se due nomi sono "abbastanza simili"
  function isFuzzyMatch(name1, name2, tolerance = 3) {
    if (!name1 || !name2) return false;
    // NORMALIZZAZIONE AVANZATA: usiamo la funzione che pulisce SRL, SPA, punteggiatura, etc.
    const n1 = normalizeName(name1);
    const n2 = normalizeName(name2);
    
    // Se uno contiene l'altro (es. "HOTEL ROMA" in "HOTEL ROMA SRL") è match
    if (n1.includes(n2) || n2.includes(n1)) return true;
    
    // Altrimenti usa la distanza matematica sulla stringa pulita
    const dist = levenshteinDistance(n1, n2);
    return dist <= tolerance;
  }

  function onlyDigits(s) {
  return String(s || "").replace(/[^0-9]/g, "");
}

// ============================================================
// GESTIONE DATE FORMATO ITALIANO (Refactoring Dic 2025)
// ============================================================
// Obiettivo: TUTTE le date devono essere gestite in formato italiano DD/MM/YYYY
// Questa è la funzione centrale per parsing flessibile e conversione
// ============================================================

// ============================================================
// 📅 UTILITIES DATE
// ============================================================

/**
 * Parsing flessibile di date in qualsiasi formato comune
 * @param {*} input - Può essere: stringa, numero (seriale Excel), Date
 * @returns {Date|null} - Oggetto Date valido o null
 */
function parseDateFlexible(input) {
  if (!input) return null;

  // Se è già un oggetto Date valido
  if (input instanceof Date && !isNaN(input.getTime())) {
    return input;
  }

  // Se è un numero (seriale Excel: giorni da 1/1/1900)
  if (typeof input === 'number' && !isNaN(input)) {
    // Excel epoch: 1 gennaio 1900 = seriale 1 (con bug leap year)
    const excelEpoch = new Date(1899, 11, 30); // 30 dic 1899
    const msPerDay = 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + input * msPerDay);
  }

  // Se è una stringa, usa normalizeDate esistente
  const normalized = normalizeDate(input);
  if (!normalized) return null;

  // normalizeDate restituisce "YYYY-MM-DD"
  const parts = normalized.split("-");
  if (parts.length !== 3) return null;

  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  const d = parseInt(parts[2], 10);

  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;

  return new Date(y, m - 1, d);
}

/**
 * Formatta una data in formato italiano DD/MM/YYYY
 * @param {*} input - Qualsiasi formato di input (stringa ISO, Date, numero)
 * @returns {string} - Data in formato "DD/MM/YYYY" o "" se non valida
 */
function formatDateIT(input) {
  if (!input) return "";
  
  // Se l'input è già una stringa ISO "YYYY-MM-DD", converte direttamente
  if (typeof input === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(input)) {
    const [y, m, d] = input.split("-").map(x => parseInt(x, 10));
    if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      const dd = String(d).padStart(2, '0');
      const mm = String(m).padStart(2, '0');
      const yyyy = String(y);
      return `${dd}/${mm}/${yyyy}`;
    }
  }
  
  // Altrimenti usa il parser flessibile per Date objects
  const date = parseDateFlexible(input);
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
    return "";
  }

  const d = date.getDate();
  const m = date.getMonth() + 1;
  const y = date.getFullYear();

  const dd = String(d).padStart(2, '0');
  const mm = String(m).padStart(2, '0');
  const yyyy = String(y);

  return `${dd}/${mm}/${yyyy}`;
}

// parseDate: usa normalizeDate per capire il formato,
// poi restituisce SEMPRE un oggetto Date (oppure null)
// DEPRECATO: usare parseDateFlexible al suo posto
function parseDate(d) {
  return parseDateFlexible(d);
}

// Converte una stringa data "strana" in formato italiano DD/MM/YYYY
// DEPRECATO: usare formatDateIT al suo posto
function formatDateForUIRaw(raw) {
  return formatDateIT(raw);
}

/**
 * Verifica se due date sono vicine entro una tolleranza in giorni
 * @param {Date} d1 - Prima data
 * @param {Date} d2 - Seconda data
 * @param {number} maxDiffDays - Massima differenza in giorni
 * @returns {boolean}
 */
function datesClose(d1, d2, maxDiffDays) {
  // Guard clause: valida input
  if (!d1 || !d2 || !(d1 instanceof Date) || !(d2 instanceof Date)) {
    return false;
  }
  if (isNaN(d1.getTime()) || isNaN(d2.getTime())) {
    return false;
  }
  
  const diffMs = Math.abs(d1.getTime() - d2.getTime());
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  return diffDays <= maxDiffDays;
}

// rimane uguale
function maxCommonDigitRun(a, b, minLen) {
  const da = onlyDigits(a);
  const db = onlyDigits(b);
  if (!da || !db) return 0;
  let best = 0;
  for (let i = 0; i < da.length; i++) {
    for (let j = i + minLen; j <= da.length; j++) {
      const sub = da.slice(i, j);
      if (db.includes(sub)) {
        const len = sub.length;
        if (len > best) best = len;
      }
    }
  }
  return best;
}

// per mostrare in tabella in formato italiano (da Date)
// DEPRECATO: usare formatDateIT al suo posto
function formatDateForUI(dateObj) {
  return formatDateIT(dateObj);
}

    // relazione “forte” tra cifre dei numeri documento
  // true se uno è contenuto nell'altro (prefisso/suffisso/substring) con almeno N cifre
  function invoiceDigitsRelated(aNum, gNum, minLen = 3) {
    const aD = normalizeInvoiceNumber(aNum, false); // solo cifre
    const gD = normalizeInvoiceNumber(gNum, false);
    if (!aD || !gD) return false;

    if (aD === gD) return true;

    let shorter = aD.length <= gD.length ? aD : gD;
    let longer  = aD.length <= gD.length ? gD : aD;

    if (shorter.length < minLen) return false;

    return longer.includes(shorter);
  }

/**
 * Verifica uguaglianza importi con tolleranza
 * @param {number} val1 - Primo importo
 * @param {number} val2 - Secondo importo
 * @param {number} tolerance - Tolleranza (default: TOLERANCE_ROUNDING)
 * @returns {boolean}
 */
  function amountsMatch(val1, val2, tolerance = MATCH_CONFIG.TOLERANCE_ROUNDING) {
      // Guard clause: valida input
      if (val1 === null || val1 === undefined || isNaN(Number(val1))) return false;
      if (val2 === null || val2 === undefined || isNaN(Number(val2))) return false;
      
      const v1 = Number(val1);
      const v2 = Number(val2);
      return Math.abs(v1 - v2) <= tolerance;
  }

  // Helper legacy per compatibilità con vecchi pezzi di codice
  function numbersAlmostEqual(a, b, tolerance = 0.05) {
      const v1 = Number(a ?? 0.0);
      const v2 = Number(b ?? 0.0);
      if (isNaN(v1) || isNaN(v2)) return false;
      return Math.abs(v1 - v2) <= tolerance;
  }

  // ============================================================
  // NORMALIZZAZIONE DATE (Refactoring Dic 2025)
  // ============================================================
  // Due funzioni separate per garantire parsing corretto:
  // 1. normalizeDateItalyStrict → SOLO per ADE (SEMPRE DD/MM/YYYY)
  // 2. normalizeDateFlexible → per Gestionale (DD/MM, MM/DD, ISO)
  // ============================================================

  /**
   * Parser RIGIDO per date ADE - SEMPRE formato italiano DD/MM/YYYY
   * @param {*} raw - Stringa data da ADE (es. "08/02/2025", "8/2/25")
   * @returns {string} - Formato ISO "YYYY-MM-DD" o "" se non valida
   * 
   * IMPORTANTE: Questa funzione NON interpreta mai il formato americano.
   * Assume SEMPRE che il primo numero sia il GIORNO e il secondo il MESE.
   */
  function normalizeDateItalyStrict(raw) {
    if (!raw) return "";
    let v = String(raw).trim();

    // Rimuovi eventuale componente ora (es. "08/02/2025 00:00:00" -> "08/02/2025")
    v = v.split(" ")[0];

    // 🔍 DEBUG: Log per vedere cosa arriva dal file ADE
    console.log(`📅 ADE Date Input: "${raw}" → dopo pulizia: "${v}"`);

    // Se è già in formato ISO "2025-02-08", valida e restituisci
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      const [y, m, d] = v.split("-").map(x => parseInt(x, 10));
      if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
        return v;
      }
    }

    // Pattern SOLO per date con separatori / o - in formato DD/MM/YYYY
    const m = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (!m) {
      console.warn(`⚠️ ADE: Data non riconosciuta (formato atteso DD/MM/YYYY): ${raw}`);
      return "";
    }

    // INTERPRETAZIONE RIGIDA ITALIANA: primo numero = giorno, secondo = mese
    let d = parseInt(m[1], 10);
    let mth = parseInt(m[2], 10);
    let y = parseInt(m[3], 10);

    // Anno a 2 cifre → aggiungi 2000 (es. 25 → 2025)
    if (y < 100) y += 2000;

    // Validazione: verifica che giorno e mese siano nei range corretti
    if (mth < 1 || mth > 12) {
      console.warn(`⚠️ ADE: Mese non valido (${mth}) in data: ${raw}`);
      return "";
    }
    if (d < 1 || d > 31) {
      console.warn(`⚠️ ADE: Giorno non valido (${d}) in data: ${raw}`);
      return "";
    }

    const mm = String(mth).padStart(2, "0");
    const dd = String(d).padStart(2, "0");
    const result = `${y}-${mm}-${dd}`;
    
    // 🔍 DEBUG: Uncomment per vedere la conversione
    console.log(`✅ ADE: "${raw}" → ISO: "${result}" (giorno=${d}, mese=${mth})`);
    
    return result;
  }

  /**
   * Parser FLESSIBILE per date Gestionale - gestisce vari formati
   * @param {*} raw - Stringa data da Gestionale (vari formati possibili)
   * @returns {string} - Formato ISO "YYYY-MM-DD" o "" se non valida
   * 
   * Gestisce:
   * - DD/MM/YYYY (italiano)
   * - MM/DD/YYYY (americano - se giorno > 12)
   * - YYYY-MM-DD (ISO)
   * - DD-MM-YY (italiano breve)
   */
  function normalizeDateFlexible(raw) {
    if (!raw) return "";
    let v = String(raw).trim();

    // Rimuovi eventuale componente ora
    v = v.split(" ")[0];

    // Se è già in formato ISO "2025-03-17", valida e restituisci
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      const [y, m, d] = v.split("-").map(x => parseInt(x, 10));
      if (y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
        return v;
      }
    }

    // Pattern per date con separatori / o -
    const m = v.match(/^(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (!m) {
      console.warn(`⚠️ GEST: Data non riconosciuta: ${raw}`);
      return "";
    }

    let n1 = parseInt(m[1], 10);
    let n2 = parseInt(m[2], 10);
    let n3 = parseInt(m[3], 10);

    // Anno a 2 cifre → aggiungi 2000
    if (n3 < 100) n3 += 2000;
    if (n1 < 100 && n1 > 31) n1 += 2000;

    let y, mth, d;

    // Logica di disambiguazione basata sui valori
    if (n1 > 31) {
      // Formato YYYY/MM/DD o YYYY-MM-DD
      y = n1;
      mth = n2;
      d = n3;
    } else if (n1 > 12 && n2 <= 12) {
      // Formato DD/MM/YYYY (giorno > 12 quindi non può essere mese)
      d = n1;
      mth = n2;
      y = n3;
    } else if (n2 > 12 && n1 <= 12) {
      // Formato MM/DD/YYYY (stile USA) → mese/giorno/anno
      mth = n1;
      d = n2;
      y = n3;
    } else {
      // ⚠️ CASI AMBIGUI: preferenza per formato italiano DD/MM/YYYY
      d = n1;
      mth = n2;
      y = n3;
    }

    // Validazione finale
    if (mth < 1 || mth > 12 || d < 1 || d > 31) {
      console.warn(`⚠️ GEST: Data non valida dopo normalizzazione: ${raw} -> ${y}-${mth}-${d}`);
      return "";
    }

    const mm = String(mth).padStart(2, "0");
    const dd = String(d).padStart(2, "0");
    return `${y}-${mm}-${dd}`;
  }

  // Funzione legacy per compatibilità - ora usa normalizeDateFlexible
  function normalizeDate(raw) {
    return normalizeDateFlexible(raw);
  }

  // ============================================================
  // NORMALIZZAZIONE NUMERI FATTURA (Miglioramento Dic 2025)
  // ============================================================
  // Obiettivo: creare un "numero confrontabile" stabile tra ADE e gestionale
  // Esempi: "5901499012/LC" ↔ "5901499012"
  //         "73/2025" ↔ "73"
  //         "F2503919205" ↔ "2503919205"
  // Parametro keepLetters:
  //   true  = mantiene lettere (es. "F2503919205" -> "F2503919205")
  //   false = solo cifre (es. "F2503919205" -> "2503919205")
  // ============================================================
    function normalizeInvoiceNumber(val, keepLetters = true) {
      if (!val) return "";

      // 1) Rimuovi apici, spazi extra, converti in maiuscolo, trim
      let s = String(val)
        .replace(/^['"]|['"]$/g, "") // rimuovi apici all'inizio e alla fine
        .replace(/['"]/g, "")        // rimuovi apici residui nel mezzo
        .replace(/\s+/g, "")         // spazi multipli -> nessuno spazio
        .toUpperCase()
        .trim();

      // 2) GESTIONE NOTAZIONE SCIENTIFICA (es. 1,00918E+11)
      const sciMatch = s.match(/^(\d+(?:[.,]\d+)?)E\+?(\d+)$/i);
      if (sciMatch) {
          try {
              const baseStr = sciMatch[1].replace(",", ".");
              const exp = parseInt(sciMatch[2], 10);
              const num = Number(baseStr + "e" + exp);
              if (!isNaN(num) && isFinite(num)) {
                  s = Math.round(num).toString();
              }
          } catch (e) {
              // fallback: lascio s così com'è
          }
      }

      // 3) Rimuovi caratteri speciali comuni nei numeri fattura
      //    Manteniamo solo alfanumerici (togliamo /, -, _, ecc.)
      s = s.replace(/[^A-Z0-9]/g, "");

      // 4) Se keepLetters è false, rimuovi anche le lettere
      if (!keepLetters) {
          s = s.replace(/[A-Z]/g, "");
      }

      // 5) Rimuovi zeri iniziali (es. "00123" -> "123")
      //    Ma solo se rimane qualcosa dopo la rimozione
      const withoutLeadingZeros = s.replace(/^0+/, "");
      if (withoutLeadingZeros.length > 0) {
          s = withoutLeadingZeros;
      }

      return s;
  }

    // Applica una regola di trasformazione numero fattura (apprendimento)
  function applyTransformRule(ruleCode, rawNum) {
      let s = String(rawNum || "").trim().toUpperCase();
      if (!s) return "";

      switch (ruleCode) {
          case "REMOVE_ZEROS":
              // es: "000123" -> "123"
              return s.replace(/^0+/, "");
          case "REMOVE_YEAR":
              // es: "123/2024" -> "123"
              return s.split("/")[0];
          case "REMOVE_FIRST_CHAR":
              // es: "F123" -> "123"
              return s.length > 1 ? s.substring(1) : s;
          case "GEST_CONTAINS_ADE":
              // qui il numero ADE resta com'è, verrà usato come "pezzo" da cercare nel GEST
              return s;
          default:
              return s;
      }
  }

  // ============================================================
  // PARSING NUMERI FORMATO ITALIANO/AMERICANO (Miglioramento Dic 2025)
  // ============================================================
  // Obiettivo: convertire correttamente numeri in formato italiano E americano
  // 
  // Formato ITALIANO: punto=migliaia, virgola=decimali
  //   "4.717,00" → 4717.00
  //   "1.234.567,89" → 1234567.89
  // 
  // Formato AMERICANO: virgola=migliaia, punto=decimali  
  //   "4,717.00" → 4717.00
  //   "1,298.17" → 1298.17
  // 
  // Logica di riconoscimento:
  // - Se ha SOLO punto → americano (punto=decimale)
  // - Se ha SOLO virgola → italiano (virgola=decimale)
  // - Se ha ENTRAMBI → guarda la posizione:
  //   * Punto DOPO virgola (x,xxx.yy) = americano
  //   * Virgola DOPO punto (x.xxx,yy) = italiano
  // ============================================================

  // ============================================================
  // 💰 UTILITIES IMPORTI E NUMERI
  // ============================================================

  function parseNumberIT(val) {
      if (val == null) return 0.0;

      let s = String(val)
          .replace(/['"]/g, '')
          .replace(/€/g, '')
          .replace(/EUR/gi, '')
          .replace(/\s+/g, '')
          .trim();

      if (s === "" || s === "-") return 0.0;

      // CASO ADE: numero con solo punto decimale e nessuna virgola → formato americano
      // Es: 195.14  /  24838.49  /  887.00
      if (/^\d+\.\d{2}$/.test(s)) {
          return parseFloat(s);
      }

      // CASO ITALIANO (Gestionale): "4.717,00" → rimuovi punti e cambia virgola
      if (s.includes(",") && s.includes(".")) {
          return parseFloat(s.replace(/\./g, "").replace(",", "."));
      }

      // SOLO VIRGOLA → decimali italiani
      if (s.includes(",")) {
          return parseFloat(s.replace(",", "."));
      }

      // SOLO CIFRE → intero
      if (/^\d+$/.test(s)) {
          return parseFloat(s);
      }

      // fallback
      const num = parseFloat(s);
      return isNaN(num) ? 0.0 : num;
  }
  // Esempi:
  //   "9,24"    → 9.24
  //   "-18,09"  → -18.09
  //   "1.234,50" → 1234.5

  // ============================================================
  // 🔢 UTILITIES P.IVA E NUMERI FATTURA
  // ============================================================

  // ============================================================
  // VALIDAZIONE E FIX IMPORTI (Correzione Dic 2025)
  // ============================================================
  /**
   * Valida e corregge la coerenza tra imponibile, IVA e totale
   * @param {Object} doc - Documento con campi imp, iva, tot
   * @param {string} source - "ADE" o "GEST" per logging
   * @param {string} docId - Identificativo documento per logging
   * @returns {Object} - Documento con importi validati/corretti
   */
  function fixAndValidateAmounts(doc, source, docId) {
    const { imp, iva, tot } = doc;
    let fixed = { ...doc };
    
    // 1. Calcola il totale atteso
    const expectedTot = +(imp + iva).toFixed(2);
    
    // 2. Se il totale è mancante o zero, calcolalo
    if (!tot || tot === 0) {
      if (imp > 0 || iva > 0) {
        console.warn(`⚠️ ${source} [${docId}]: Totale mancante - calcolato da imp+iva: ${expectedTot}`);
        fixed.tot = expectedTot;
      }
    }
    // 3. Se il totale esiste ma è inconsistente (diff > 0.5€)
    else if (Math.abs(tot - expectedTot) > 0.50) {
      console.warn(`⚠️ ${source} [${docId}]: INCONSISTENZA IMPORTI!`);
      console.warn(`   Imponibile: ${imp}`);
      console.warn(`   IVA: ${iva}`);
      console.warn(`   Totale dichiarato: ${tot}`);
      console.warn(`   Totale atteso (imp+iva): ${expectedTot}`);
      console.warn(`   Differenza: ${(tot - expectedTot).toFixed(2)}€`);
      
      // Se l'imponibile è sospettosamente basso rispetto al totale
      if (imp <= 10 && tot > 100) {
        console.error(`   🚨 PROBABILE ERRORE MAPPING: Imponibile troppo basso (${imp}) per totale ${tot}`);
      }
    }
    // 4. Validazione positiva se tutto torna
    else {
      // Tutto ok - differenza entro ±0.50€
    }
    
    return fixed;
  }

  // ============================================================
  // VALIDAZIONE COLONNE IVA vs P.IVA (Correzione Dic 2025)
  // ============================================================
  // Problema: la colonna IVA contiene P.IVA per errore di mapping
  // Esempio: GEST_Iva = 278730825,00 (è una P.IVA, non un importo IVA!)
  // ============================================================

  /**
   * Verifica se un valore sembra essere una P.IVA invece di un importo IVA
   * @param {*} value - Valore da verificare
   * @returns {boolean} - true se sembra una P.IVA
   */
  function looksLikePIVA(value) {
    if (!value) return false;
    
    const str = String(value).replace(/[^0-9]/g, '');
    
    // P.IVA italiana: 11 cifre
    if (str.length === MATCH_CONFIG.PIVA_IT_LENGTH) {
      const asNumber = parseFloat(str);
      // Se è un numero molto grande (> 10 miliardi), probabilmente è P.IVA
      if (asNumber > MATCH_CONFIG.PIVA_IT_MIN_VALUE) return true;
    }
    
    // P.IVA estera con prefisso (es. DE123456789)
    const original = String(value).toUpperCase().trim();
    if (/^[A-Z]{2}\d{7,}/.test(original)) return true;
    
    return false;
  }

  /**
   * Pulisce un valore IVA da possibili contaminazioni con P.IVA
   * @param {*} ivaValue - Valore dalla colonna IVA
   * @param {*} pivaValue - Valore dalla colonna P.IVA (per confronto)
   * @returns {number} - Importo IVA pulito o 0.0
   */
  function cleanIVAValue(ivaValue, pivaValue) {
    if (!ivaValue) return 0.0;
    
    // Se il valore IVA sembra una P.IVA, restituisci 0
    if (looksLikePIVA(ivaValue)) {
      console.warn(`⚠️ Valore P.IVA trovato in colonna IVA: ${ivaValue} - impostato a 0`);
      return 0.0;
    }
    
    // Se IVA e P.IVA sono identici (errore di mapping), restituisci 0
    if (pivaValue && String(ivaValue) === String(pivaValue)) {
      console.warn(`⚠️ IVA uguale a P.IVA: ${ivaValue} - impostato a 0`);
      return 0.0;
    }
    
    // Altrimenti parse normale
    return parseNumberIT(ivaValue);
  }

  // ============================================================
  // RILEVAMENTO NOTE CREDITO GESTITE AL CONTRARIO (Dic 2025)
  // ============================================================
  /**
   * Rileva se una fattura è probabilmente una nota di credito gestita al contrario nel gestionale.
   * Pattern tipico:
   * - ADE_Iva negativa (nota di credito)
   * - GEST_Iva = 0,00
   * - ADE_Imponibile = GEST_Imponibile
   * - DIFF_TOTALE ≈ -2 × ADE_Iva (in valore assoluto)
   * 
   * @param {Object} adeRecord - Record ADE con campi: imp, iva, tot
   * @param {Object} gestRecord - Record Gestionale con campi: imp, iva, tot
   * @returns {Object} - { isNCInvertita: boolean, diff: number, formula: string }
   */
  function detectNCInvertita(adeRecord, gestRecord) {
    const result = {
      isNCInvertita: false,
      diff: 0,
      formula: "",
      kind: "" // "SEGNO_OPPOSTO" oppure "IVA_ZERO_GEST"
    };

    if (!adeRecord || !gestRecord) return result;

    const adeIvaNum = adeRecord.iva || 0;
    const adeImpNum = adeRecord.imp || 0;
    const adeTotNum = adeRecord.tot || 0;

    const gestImpNum = gestRecord.imp || 0;
    const gestIvaNum = gestRecord.iva || 0;
    const gestTotNum = gestRecord.tot || 0;

    const diff = adeTotNum - gestTotNum;
    result.diff = diff;

    // ============================================================
    // ✅ NUOVO PATTERN: SEGNO OPPOSTO (ADE positivo, GEST negativo o viceversa)
    // Se |imp|, |iva|, |tot| coincidono ma il totale ha segno opposto => probabile NC
    // ============================================================
    const segnoOpposto =
      (adeTotNum * gestTotNum < 0) && // segni opposti
      (Math.abs(Math.abs(adeImpNum) - Math.abs(gestImpNum)) <= 0.01) &&
      (Math.abs(Math.abs(adeIvaNum) - Math.abs(gestIvaNum)) <= 0.01) &&
      (Math.abs(Math.abs(adeTotNum) - Math.abs(gestTotNum)) <= 0.02);

    if (segnoOpposto) {
      result.isNCInvertita = true;
      result.kind = "SEGNO_OPPOSTO";
      result.formula = `|ADE|≈|GEST| ma segno opposto (tot ADE=${adeTotNum.toFixed(2)} / tot GEST=${gestTotNum.toFixed(2)})`;
      return result;
    }

    // ============================================================
    // ✅ PATTERN VECCHIO: NC “invertita” nel gestionale (IVA=0 ecc.)
    // ============================================================
    const isNC = adeIvaNum < 0;
    const imponibileUguale = Math.abs(adeImpNum - gestImpNum) <= 0.01;
    const ivaGestionaleZero = Math.abs(gestIvaNum) <= 0.01;
    const expectedDiff = -2 * Math.abs(adeIvaNum);

    const patternNCInvertita =
      isNC &&
      imponibileUguale &&
      ivaGestionaleZero &&
      Math.abs(diff - expectedDiff) <= 0.05;

    result.isNCInvertita = patternNCInvertita;

    if (patternNCInvertita) {
      result.kind = "IVA_ZERO_GEST";
      result.formula = `diff=${diff.toFixed(2)} ≈ -2×|${adeIvaNum.toFixed(2)}| = ${expectedDiff.toFixed(2)}`;
    }

    return result;
  }

    function isForeignPiva(raw) {
    if (!raw) return false;
    let v = String(raw).toUpperCase().trim();

    // Se inizia con un prefisso paese diverso da IT (es. DE, FR, ES...) la considero estera
    const m = v.match(/^([A-Z]{2})/);
    if (m && m[1] !== "IT") {
      return true;
    }

    // Tutto il resto (anche CF, codici strani, ecc.) resta NEL perimetro
    return false;
  }

  function normalizePIVA(raw) {
    if (!raw) return "";
    let v = String(raw).toUpperCase().trim();
    // Rimuove il prefisso IT se presente
    v = v.replace(/^IT/g, "");
    // Rimuove eventuali spazi o altri caratteri non numerici
    return v.replace(/[^0-9]/g, "");
  }
  // ===============================
  // APPRENDIMENTO CONTAIBIL (ESTESO)
  // ===============================

  // stato unico per apprendimento: regole + mapping colonne + dizionario fornitori
  let learningData = {
    versione: 2,
    azienda: "STANDARD",
    columns: {
      ADE: {},
      GEST: {},       // legacy key (compatibilità)
      GEST_B: {},     // File B (Gestionale Fatture)
      GEST_C: {}      // File C (Note Credito)
    },
    rules: [],
    supplierPivaMap: {} // Nuovo: per il recupero P.IVA
  };

  function ensureLearningShape() {
    if (!learningData || typeof learningData !== "object") {
      learningData = {
        versione: 2,
        azienda: "STANDARD",
        columns: { ADE: {}, GEST: {}, GEST_B: {}, GEST_C: {} },
        rules: []
      };
      if (!learningData.supplierPivaMap) learningData.supplierPivaMap = {}; // Assicurati che la mappa esista
      return;
    }
    if (!learningData.columns) learningData.columns = { ADE: {}, GEST: {}, GEST_B: {}, GEST_C: {} };
    if (!learningData.columns.ADE) learningData.columns.ADE = {};
    if (!learningData.columns.GEST) learningData.columns.GEST = {};
    if (!learningData.columns.GEST_B) learningData.columns.GEST_B = {};
    if (!learningData.columns.GEST_C) learningData.columns.GEST_C = {};
    // migrazione soft: se GEST_B è vuoto ma GEST ha dati, copiali su B per compatibilità
    if (Object.keys(learningData.columns.GEST_B).length === 0 && Object.keys(learningData.columns.GEST).length > 0) {
      learningData.columns.GEST_B = { ...learningData.columns.GEST };
    }
    if (!Array.isArray(learningData.rules)) learningData.rules = [];
    if (!learningData.versione || learningData.versione < 2) {
      learningData.versione = 2;
    }
  }

  // restituisce la config colonne logica per "ADE" o "GEST"
  function getColumnConfig(tipo) {
    ensureLearningShape();
    const t = tipo.toUpperCase();
    const base = {
      num: null, den: null, piva: null,
      data: null, dataReg: null, // <--- NUOVO CAMPO
      imp: null, iva: null, tot: null
    };
    // Priorità: chiave specifica (GEST_B / GEST_C) → legacy GEST → vuoto
    const columns = learningData.columns || {};
    const fromJson = columns[t] || (t.startsWith("GEST") ? (columns.GEST_B || columns.GEST || {}) : {});
    return { ...base, ...fromJson };
  }

  // aggiorna (o crea) la config colonne per ADE/GEST
  function updateColumnConfig(tipo, partialCfg) {
    ensureLearningShape();
    const t = tipo.toUpperCase();
    const prev = learningData.columns[t] || {};
    learningData.columns[t] = { ...prev, ...partialCfg };
    // Mantieni anche la chiave legacy GEST se stai aggiornando B e quella legacy è vuota
    if (t === "GEST_B") {
      const legacy = learningData.columns.GEST || {};
      learningData.columns.GEST = { ...legacy, ...partialCfg };
    }
    refreshLearningPanel();
  }

  // legge eventuali modifiche fatte a mano nel textarea
  function readLearningFromTextarea() {
    const box = document.getElementById("learningCodeBox");
    if (!box) return;
    const txt = box.value.trim();
    if (!txt) return;
    try {
      const parsed = JSON.parse(txt);
      if (parsed && typeof parsed === "object") {
        learningData = parsed;
        // Assicurati che la struttura sia completa anche se il JSON è parziale
        ensureLearningShape();
      }
    } catch (e) {
      console.warn("JSON procedure non valido, uso configurazione interna.", e);
    }
  }

  // riscrive il pannello con lo stato corrente
  function refreshLearningPanel() {
    ensureLearningShape();
    const box = document.getElementById("learningCodeBox");
    if (!box) return;
    box.value = JSON.stringify(learningData, null, 2);
  }

  // costruisce una regola a partire da una riga ADE/GEST
  function buildLearningRule(adeRow, gestRow, tipo) {
    ensureLearningShape();
    const rule = {
      tipo: tipo || "MATCH_GENERICO",
      azienda: learningData.azienda || "STANDARD",
      
      // chiavi lato ADE
      adePiva: (adeRow?.ADE_PIVA || "").toString().trim(),
      adeNum: (adeRow?.ADE_Num || "").toString().trim(),
      adeNumNorm: normalizeInvoiceNumber(adeRow?.ADE_Num || "", false),
      adeDen: (adeRow?.ADE_Den || adeRow?.ADE_Denominazione || "").toString().trim(),
      adeDenNorm: normalizeName(adeRow?.ADE_Den || adeRow?.ADE_Denominazione || ""), // Già normalizzato
      adeTot: parseNumberIT(adeRow?.ADE_Totale ?? adeRow?.ADE_Imponibile ?? 0.0),

      // chiavi lato GEST
      gestNum: (gestRow.GEST_Num || "").toString().trim(),
      gestNumNorm: normalizeInvoiceNumber(gestRow.GEST_Num || ""),
      gestDen: (gestRow.GEST_Den || gestRow.GEST_Denominazione || "").toString().trim(),
      gestDenNorm: normalizeName(gestRow.GEST_Den || gestRow.GEST_Denominazione || ""),
      gestTot: parseNumberIT(gestRow?.GEST_Totale ?? gestRow?.GEST_Imponibile ?? 0.0),

      createdAt: new Date().toISOString(),
    };

    return rule;
  }

  // registra una nuova regola evitando duplicati
  function registerLearningRule(adeRow, gestRow, tipo) {
    const rule = buildLearningRule(adeRow, gestRow, tipo);
    ensureLearningShape();
    const rulesArr = learningData.rules;

    const exists = rulesArr.some(r =>
      r.tipo === rule.tipo &&
      r.azienda === rule.azienda &&
      r.adePiva === rule.adePiva &&
      r.adeNumNorm === rule.adeNumNorm &&
      r.gestNumNorm === rule.gestNumNorm &&
      r.gestDenNorm === rule.gestDenNorm &&
      r.adeTot === rule.adeTot &&
      r.gestTot === rule.gestTot
    );

    if (!exists) { // Aggiungi solo se non esiste già una regola identica
      rulesArr.push(rule);
      refreshLearningPanel();
    }
  }

  // ---------- Flow helpers (stepper + mapping) ----------

  function headersSignature(headers = []) {
    return headers.map(h => String(h || "").toLowerCase().trim()).join("|");
  }

  function markMappingConfirmed(key, headers) {
    const k = key || "GEST_B";
    if (!mappingSession[k]) mappingSession[k] = { confirmed: false, headerSig: null };
    mappingSession[k].confirmed = true;
    mappingSession[k].headerSig = headersSignature(headers);
    flowState.mapping[k === "GEST_C" ? "gestCConfirmed" : "gestBConfirmed"] = true;
  }

  function resetMappingConfirmation(key) {
    const k = key || "GEST_B";
    if (!mappingSession[k]) mappingSession[k] = { confirmed: false, headerSig: null };
    mappingSession[k].confirmed = false;
    mappingSession[k].headerSig = null;
    if (k === "GEST_C") {
      flowState.mapping.gestCConfirmed = false;
    } else {
      flowState.mapping.gestBConfirmed = false;
    }
  }

  function isMappingConfirmed(key, headers) {
    const k = key || "GEST_B";
    const info = mappingSession[k];
    if (!info || !info.confirmed) return false;
    if (!headers || !headers.length) return info.confirmed; // se non abbiamo header, fidati dello stato
    return info.headerSig === headersSignature(headers);
  }

  // ---------- MAPPING MANUALE COLONNE GESTIONALE ----------

  function fillSelectWithHeaders(selectEl, headers, preselect) {
    if (selectEl) {
      selectEl.innerHTML = "";
      const optEmpty = document.createElement("option");
      optEmpty.value = "";
      optEmpty.textContent = "— nessuna —";
      selectEl.appendChild(optEmpty);

      headers.forEach(h => {
        const opt = document.createElement("option");
        opt.value = h;
        opt.textContent = h;
        if (preselect && preselect === h) opt.selected = true;
        selectEl.appendChild(opt);
      });
    }
  }

  function openGestMappingModal(headers, fileDescription, mappingKey = "GEST_B") {
    return new Promise((resolve, reject) => {
      const backdrop = document.getElementById("gestMapBackdrop");
      if (!backdrop) { resolve({}); return; } // Sicurezza

      const modalTitle = document.getElementById("gestMapTitle");
      const selNum  = document.getElementById("selGestNum");
      const selPiva = document.getElementById("selGestPiva");
      const selDen  = document.getElementById("selGestDen");
      const selData = document.getElementById("selGestData");
      const selDataReg = document.getElementById("selGestDataReg");
      const selImp  = document.getElementById("selGestImp");
      const selIva  = document.getElementById("selGestIva");
      const selTot  = document.getElementById("selGestTot");

      const helper  = document.getElementById("gestMapHelper");

      const btnOk     = document.getElementById("btnGestMapOk");
      const btnCancel = document.getElementById("btnGestMapCancel");

      // Aggiorna titolo e helper in base a mappingKey
      if (mappingKey === "GEST_C") {
        if (modalTitle) {
          modalTitle.textContent = "Aggancio colonne Note di Credito (File C)";
        }
        if (helper) {
          helper.textContent = "File Note di Credito (NC): mapping separato dal File B. Obbligatori: Numero documento, Data, Totale oppure (Imponibile + IVA).";
        }
      } else {
        if (modalTitle) {
          modalTitle.textContent = "Aggancio colonne Gestionale (File B)";
        }
        if (helper) {
          helper.textContent = "Obbligatori: Numero documento, Data, Totale oppure (Imponibile + IVA).";
        }
      }

      // config attuale (se già salvata in learningData.columns.<KEY>)
      const cfg = getColumnConfig(mappingKey || "GEST_B");

      fillSelectWithHeaders(selNum,  headers, cfg.num);
      fillSelectWithHeaders(selPiva, headers, cfg.piva);
      fillSelectWithHeaders(selDen,  headers, cfg.den);
      fillSelectWithHeaders(selData, headers, cfg.data);
      fillSelectWithHeaders(selDataReg, headers, cfg.dataReg);
      fillSelectWithHeaders(selImp,  headers, cfg.imp);
      fillSelectWithHeaders(selIva,  headers, cfg.iva);
      fillSelectWithHeaders(selTot,  headers, cfg.tot);

      function closeModal() {
        backdrop.classList.add("hidden");
        btnOk.onclick = null;
        btnCancel.onclick = null;
      }

      function validateAndToggle() {
        const hasNum  = !!selNum.value;
        const hasData = !!selData.value;
        const hasTot  = !!selTot.value;
        const hasImpIva = !!selImp.value && !!selIva.value;
        const amountsOk = hasTot || hasImpIva;
        const valid = hasNum && hasData && amountsOk;

        if (btnOk) btnOk.disabled = !valid;
        if (helper) {
          helper.textContent = valid
            ? "Pronto: i campi minimi sono selezionati."
            : "Obbligatori: Numero documento, Data documento, e Totale oppure (Imponibile + IVA).";
          helper.classList.toggle("error", !valid);
        }
        return valid;
      }

      [selNum, selData, selImp, selIva, selTot].forEach(sel => {
        if (sel) sel.addEventListener("change", validateAndToggle);
      });
      validateAndToggle();

      btnOk.onclick = () => {
        if (!validateAndToggle()) return;

        const newCfg = {
          num:  selNum.value  || null,
          piva: selPiva.value || null,
          den:  selDen.value  || null,
          data: selData.value || null,
          dataReg: selDataReg ? (selDataReg.value || null) : null,
          imp:  selImp.value  || null,
          iva:  selIva.value  || null,
          tot:  selTot.value  || null,
        };
        updateColumnConfig(mappingKey || "GEST_B", newCfg);
        refreshLearningPanel(); // aggiornare il JSON nel box
        markMappingConfirmed(mappingKey || "GEST_B", headers);
        closeModal();
        resolve(newCfg);
      };

      btnCancel.onclick = () => {
        closeModal();
        resetMappingConfirmation(mappingKey || "GEST_B");
        // se annulli, non cambio nulla
        resolve(null);
      };

      backdrop.classList.remove("hidden");
    });
  }

  // chiede sempre il mapping prima di usare il gestionale
  function hasValidGestConfig(headers, fileDescription, mappingKey = "GEST_B") {
    // controlla se le colonne salvate in learningData.columns.GEST
    // esistono davvero nell'header del file caricato
    const cfg = getColumnConfig(mappingKey || "GEST_B");
    if (!cfg) return false;
    
    // Se non ci sono colonne configurate, la config non è valida.
    const configuredCols = Object.values(cfg).filter(Boolean);
    if (configuredCols.length === 0) return false;

    const headerSet = new Set(headers || []);

    // Campi minimi che vogliamo sicuri per il match. Per le note di credito, la denominazione è meno critica.
    const requiredKeys = (fileDescription && fileDescription.includes("Note Credito")) ? ["num", "data", "imp"] : ["num", "data"];

    // Campi obbligatori: num + data + (tot OR imp+iva)
    const hasNum = cfg.num && headerSet.has(cfg.num);
    const hasData = cfg.data && headerSet.has(cfg.data);
    const hasTot = cfg.tot && headerSet.has(cfg.tot);
    const hasImpIva = cfg.imp && headerSet.has(cfg.imp) && cfg.iva && headerSet.has(cfg.iva);
    const amountsOk = hasTot || hasImpIva;

    if (!hasNum || !hasData || !amountsOk) return false;

    return true; // configurazione ok per questi headers
  }

  // ---------- GESTIONE PERSISTENZA PROCEDURE (localStorage) ----------

    const LEARNING_STORAGE_KEY = "contaibil_learning_data";

  // prova prima da localStorage, se non trova niente tenta il file JSON accanto all'HTML
  async function loadProceduresAtStartup() {
    let loaded = false;

    // 1) Tenta dal localStorage
    try {
      const storedData = localStorage.getItem(LEARNING_STORAGE_KEY);
      if (storedData) {
        const parsed = JSON.parse(storedData);
        if (parsed && typeof parsed === "object" && parsed.versione >= 2) {
          learningData = parsed;
          ensureLearningShape();
          refreshLearningPanel();
          console.log("Procedure caricate automaticamente dal localStorage.");
          loaded = true;
        }
      }
    } catch (e) {
      console.error("Errore nel caricamento delle procedure dal localStorage:", e);
    }

    if (loaded) return;

    // 2) Se non ho trovato nulla nel localStorage, provo a leggere il file JSON accanto all'HTML
    try {
      const resp = await fetch(DEFAULT_PROCEDURE_FILENAME, { cache: "no-store" });
      if (!resp.ok) {
        console.log("Nessun file procedure JSON trovato accanto all'HTML:", DEFAULT_PROCEDURE_FILENAME);
        return;
      }
      const content = await resp.json();
      if (content && typeof content === "object" && Array.isArray(content.rules)) {
        learningData = content;
        ensureLearningShape();
        refreshLearningPanel();
        console.log("Procedure caricate dal file JSON:", DEFAULT_PROCEDURE_FILENAME);
        // salvo anche nel localStorage per le aperture successive
        try {
          localStorage.setItem(LEARNING_STORAGE_KEY, JSON.stringify(learningData));
        } catch (e2) {
          console.warn("Impossibile salvare le procedure nel localStorage:", e2);
        }
      } else {
        console.warn("File procedure JSON non valido:", DEFAULT_PROCEDURE_FILENAME);
      }
    } catch (e) {
      console.log(
        "Impossibile leggere il file procedure JSON (se stai aprendo il file con file:// il browser potrebbe bloccare fetch).",
        e
      );
    }
  }

    // salvataggio / caricamento file JSON
  function setupLearningButtons() {
    const btnSave = document.getElementById("btnSaveProcedures");
    const btnLoad = document.getElementById("btnLoadProcedures");
    const fileInput = document.getElementById("fileProcedures");

    // --- SALVA PROCEDURE ---
    if (btnSave) {
      btnSave.addEventListener("click", () => {
        ensureLearningShape();

        const dataStr = JSON.stringify(learningData, null, 2);

        // 1) Salva nel localStorage per la persistenza automatica
        try {
          localStorage.setItem(LEARNING_STORAGE_KEY, dataStr);
          console.log("Procedure salvate nel localStorage.");
        } catch (e) {
          console.error("Errore nel salvataggio su localStorage:", e);
        }

        // 2) Forza download del JSON con nome fisso
        const blob = new Blob([dataStr], {
          type: "application/json;charset=utf-8"
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = DEFAULT_PROCEDURE_FILENAME;   // es. "procedure_STANDARD.json"
        a.style.display = "none";

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);

        URL.revokeObjectURL(url);
      });
    }

    // --- CARICA PROCEDURE DA FILE ---
    if (btnLoad && fileInput) {
      btnLoad.addEventListener("click", () => {
        fileInput.value = "";
        fileInput.click();
      });

      fileInput.addEventListener("change", () => {
        const file = fileInput.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const content = JSON.parse(e.target.result); // Il JSON può contenere anche la mappa dei fornitori
            if (content && Array.isArray(content.rules)) {
              learningData = content;
              ensureLearningShape();
              refreshLearningPanel();

              // aggiorno anche il localStorage per coerenza
              try {
                localStorage.setItem(LEARNING_STORAGE_KEY, JSON.stringify(learningData));
              } catch (e2) {
                console.warn("Impossibile salvare le procedure nel localStorage dopo il load:", e2);
              }

              alert("Procedure caricate correttamente.");
            } else {
              alert("File procedure non valido.");
            }
          } catch (err) {
            console.error(err);
            alert("Errore nella lettura del file JSON.");
          }
        };
        reader.readAsText(file);
      });
    }
  }

  // Questa funzione è stata semplificata e non è più usata direttamente per il matching
  // usa le regole apprese per proporre un match GEST per una riga ADE
  function applyLearningRules(adeRow, gestRows) {
    ensureLearningShape();
    const rulesArr = learningData.rules;
    if (!rulesArr.length) return null;

    const adePivaNorm = normalizePIVA(adeRow.ADE_PIVA || "");
    const adeNumNorm = normalizeInvoiceNumber(adeRow.ADE_Num || "", false); // Usa la normalizzazione solo cifre
    const adeDenNorm = normalizeName(adeRow?.ADE_Den || adeRow?.ADE_Denominazione || "");
    const adeTot = parseNumberIT(adeRow?.ADE_Totale ?? adeRow?.ADE_Imponibile ?? 0.0);

    // 1. filtro le regole per P.IVA
    const candidateRules = rulesArr.filter(r => {
      const rPiva = normalizePIVA(r?.adePiva || "");
      return rPiva === adePivaNorm;
    });

    if (!candidateRules.length) return null;

    // 2. scelgo la regola più simile alla riga ADE
    let bestRule = null;
    let bestRuleScore = 0;

    candidateRules.forEach(rule => {
      let s = 0;
      if (rule?.adeNumNorm && adeNumNorm && normalizeInvoiceNumber(rule.adeNumNorm, false) === adeNumNorm) s += 50; // Confronta numeri normalizzati solo cifre
      if (rule?.adeDenNorm && adeDenNorm && rule.adeDenNorm === adeDenNorm) s += 30;
      if (numbersAlmostEqual(adeTot, rule.adeTot, 1)) s += 20;
      if (amountsMatch(adeTot, rule.adeTot, 1)) s += 20; // Sostituisce numbersAlmostEqual
      if (s > bestRuleScore) {
        bestRuleScore = s;
        bestRule = rule;
      }
    });

    if (!bestRule || bestRuleScore < 60) return null;

    // 3. tra le righe GEST disponibili scelgo quella più vicina alla regola
    let bestGest = null;
    let bestGestScore = 0;

    gestRows.forEach(gestRow => {
      let s = 0;

      const gNumNorm = normalizeInvoiceNumber(gestRow?.GEST_Num || "", false); // Usa la normalizzazione solo cifre
      const gDenNorm = normalizeName(gestRow?.GEST_Den || gestRow?.GEST_Denominazione || "");
      const gTot = Number(gestRow?.GEST_Totale ?? gestRow?.GEST_Imponibile ?? 0.0);

      if (bestRule?.gestNumNorm && gNumNorm && bestRule.gestNumNorm === gNumNorm) s += 60;
      if (bestRule?.gestDenNorm && gDenNorm && bestRule.gestDenNorm === gDenNorm) s += 30;
      if (amountsMatch(bestRule.gestTot, gTot, 1)) s += 10; // Sostituisce numbersAlmostEqual

      if (s > bestGestScore) {
        bestGestScore = s;
        bestGest = gestRow;
      }
    });

    if (!bestGest || bestGestScore < 70) return null;
    return bestGest;
  }

  // Nuova funzione più robusta per il GESTIONALE
  function findColGest(headers, lowerHeaders, aliases) {
    // normalizzo togliendo spazi, underscore, ecc.
    const normHeaders = headers.map(h =>
      String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "")
    );

    // 1) match esatto sul nome "ripulito"
    for (const alias of aliases) {
      const normAlias = String(alias || "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
      const idx = normHeaders.findIndex(h => h === normAlias);
      if (idx >= 0) return headers[idx];
    }

    // 2) match "contiene" sul nome ripulito
    for (const alias of aliases) {
      const normAlias = String(alias || "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
      const idx = normHeaders.findIndex(h => h.includes(normAlias));
      if (idx >= 0) return headers[idx];
    }

    return null;
  }

  function worksheetToCsv(ws, isADE = false) {
    // 🔧 CONVERSIONE MANUALE DATE E IMPORTI: processa ogni cella del worksheet
    // File ADE: mantiene punto decimale (formato americano)
    // File GEST: converte virgola decimale (formato italiano)
    const range = XLSX.utils.decode_range(ws['!ref']);

    // ✅ FIX: converti seriali Excel in DATE solo nelle colonne "Data..."
    const headerRow = range.s.r; // prima riga del foglio (header)
    const dateCols = new Set();

    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({ r: headerRow, c: C });
      const cell = ws[addr];
      const header = String(cell?.v ?? cell?.w ?? "").toLowerCase().trim();
      if (header.includes("data")) dateCols.add(C);
    }

    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[cellAddress];

        if (cell && cell.t === 'n' && cell.v != null) {
          const numValue = cell.v;

          // CASO 1: Seriale Excel per DATE (tra 1 e 80000 E intero)
          // Seriali Excel: 1 = 1/1/1900, 45000 ≈ 2023
          // Importante: deve essere intero per distinguerlo da importi come 610.75
          if (numValue >= 1 && numValue <= 80000 && Number.isInteger(numValue) && dateCols.has(C)) {
            // Converti seriale Excel in Date JavaScript
            const excelEpoch = new Date(1899, 11, 30);
            const msPerDay = 24 * 60 * 60 * 1000;
            const dateObj = new Date(excelEpoch.getTime() + numValue * msPerDay);

            // Formatta in DD/MM/YYYY italiano
            const d = String(dateObj.getDate()).padStart(2, '0');
            const m = String(dateObj.getMonth() + 1).padStart(2, '0');
            const y = dateObj.getFullYear();

            // Sovrascrivi il valore della cella con la stringa formattata
            cell.t = 's'; // cambia tipo a stringa
            cell.v = `${d}/${m}/${y}`;
            cell.w = `${d}/${m}/${y}`;

            delete cell.z; // rimuovi formato numero
          }
          // CASO 2: TUTTI gli altri numeri (importi, IVA, totali, quantità)
          // Include sia decimali (610.75) che interi (887)
          else {
            let strValue;

            if (isADE) {
              // ⚠️ FILE ADE: mantieni formato AMERICANO con PUNTO decimale
              // I file ADE dall'Agenzia delle Entrate hanno già il punto decimale
              // Es: 195.14 → "195.14" (NO conversione)
              strValue = numValue.toFixed(2);
            } else {
              // ✅ FILE GESTIONALE: converti in formato ITALIANO con VIRGOLA decimale
              // Es: 887 → 887.00 → "887,00"
              strValue = numValue.toFixed(2).replace('.', ',');
            }

            cell.t = 's'; // cambia tipo a stringa
            cell.v = strValue;
            cell.w = strValue;

            delete cell.z; // rimuovi formato numero
          }
        }
      }
    }

    // 🔧 FIX DEFINITIVO: Converti direttamente in JSON invece di CSV
    // Evita problemi di shift colonne causati da delimitatori CSV malformati
    console.log(`✅ Excel caricato - conversione diretta a JSON (bypass CSV)`);

    const jsonData = XLSX.utils.sheet_to_json(ws, {
      header: 1,        // Array di array (non oggetti)
      raw: false,       // Usa valori formattati (stringhe)
      defval: ""        // Celle vuote = stringa vuota
    });

    if (jsonData.length === 0) {
      throw new Error("File Excel vuoto");
    }

    // Prima riga = headers, resto = dati
    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    // 🔍 DEBUG: Mostra headers e prima riga dati
    console.log(`📋 HEADERS (${headers.length} colonne):`, headers);
    if (rows.length > 0) {
      console.log(`📋 PRIMA RIGA DATI:`, rows[0]);
    }

    // Cerca la riga con fattura 2528012466 per debug
    if (isADE) {
      const targetRow = rows.find(row => {
        return row.some(cell => String(cell).includes('2528012466'));
      });

      if (targetRow) {
        console.log("🔍 RIGA FATTURA 2528012466 DA EXCEL:");
        headers.forEach((header, idx) => {
          console.log(`   [${idx}] "${header}" = "${targetRow[idx] || ''}"`);
        });
      }
    }

    // Converti in formato CSV compatibile per il parsing esistente
    const csvLines = [];
    csvLines.push(headers.join(';'));
    rows.forEach(row => {
      csvLines.push(row.join(';'));
    });

    const csv = csvLines.join('\n');
    const formatDesc = isADE ? "punto decimale (ADE)" : "virgola decimale (GEST)";
    console.log(`📊 ${formatDesc} - Righe totali: ${rows.length}`);
    return csv;
  }

  function loadTableFromFile(file, isADE = false) {
    const name = file.name.toLowerCase();
    const isXmlSpreadsheet = name.endsWith(".xml");

    if (name.endsWith(".csv")) {
      return file.text();
    }

    if (name.endsWith(".xlsx") || isXmlSpreadsheet) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
          try {
            let wb;

            if (isXmlSpreadsheet) {
              // I file ADE originali arrivano spesso come SpreadsheetML (.xml)
              const xmlText = typeof e.target.result === "string"
                ? e.target.result
                : new TextDecoder("utf-8").decode(e.target.result);

              wb = XLSX.read(xmlText, {
                type: "string",
                cellDates: false,
                cellNF: false
              });
            } else {
              const data = new Uint8Array(e.target.result);
              wb = XLSX.read(data, {
                type: "array",
                cellDates: false,
                cellNF: false
              });
            }

            const sheetName = wb.SheetNames[0];
            const ws = wb.Sheets[sheetName];
            const csv = worksheetToCsv(ws, isADE);
            resolve(csv);
          } catch (err) {
            reject(err);
          }
        };

        reader.onerror = reject;
        if (isXmlSpreadsheet) {
          reader.readAsText(file);
        } else {
          reader.readAsArrayBuffer(file);
        }
      });
    }

    return file.text();
  }

  function getMinMaxDates(records) {
    let min = null;
    let max = null;
    for (const r of (records || [])) {
      if (!r || !r.data || !(r.data instanceof Date)) continue;
      if (!min || r.data < min) min = r.data;
      if (!max || r.data > max) max = r.data;
    }
    return { min, max };
  }

  function dateToInputValue(date) {
    if (!date) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function parseInputDate(value) {
    if (!value) return null;
    const parts = value.split("-");
    if (parts.length !== 3) return null;
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
    return new Date(y, m, d);
  }

  function formatNumberITDisplay(value) {
    if (value == null || value === "") return "0,00";
    const num = Number(value);
    if (isNaN(num)) return "0,00";
    
    // Usa le impostazioni locali italiane per mettere i punti
    return num.toLocaleString('it-IT', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
  }

  // ---------- copertura periodo riconciliazione ----------
  function updatePeriodCoverage(adeRecords, gestRecords) {
    const adeEl = document.getElementById("periodAde");
    const gestEl = document.getElementById("periodGest");
    const badgeEl = document.getElementById("periodStatusBadge");
    const textEl = document.getElementById("periodStatusText");

    if (!adeEl || !gestEl || !badgeEl || !textEl) return;

    const adeDates = getMinMaxDates(adeRecords);
    const gestDates = getMinMaxDates(gestRecords);

    if (!adeDates.min && !gestDates.min) {
      adeEl.textContent = "n/d";
      gestEl.textContent = "n/d";
      badgeEl.className = "period-badge period-badge-soft";
      badgeEl.textContent = "ℹ️ Nessuna data rilevata";
      textEl.textContent = "Carica i file ADE e gestionale e rilancia il confronto per vedere la copertura periodo.";
      return;
    }

    const adeRange = adeDates.min
      ? `dal ${formatDateIT(adeDates.min)} al ${formatDateIT(adeDates.max)}`
      : "n/d";
    const gestRange = gestDates.min
      ? `dal ${formatDateIT(gestDates.min)} al ${formatDateIT(gestDates.max)}`
      : "n/d";

    adeEl.textContent = adeRange;
    gestEl.textContent = gestRange;

    let badgeClass = "period-badge-ok";
    let badgeLabel = "✅ Copertura allineata";
    let statusText = "Le date di caricamento ADE e gestionale risultano allineate.";
    let isPartial = false;

    const aMax = adeDates.max;
    const gMax = gestDates.max;

    if (aMax && gMax) {
      const diffMs = aMax.getTime() - gMax.getTime();
      const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
      const absDiff = Math.abs(diffDays);

      if (absDiff <= 3) {
        badgeClass = "period-badge-ok";
        badgeLabel = "✅ Copertura allineata";
        statusText = "Le date finali di ADE e gestionale sono allineate (differenza entro 3 giorni).";
      } else if (absDiff <= 7) {
        badgeClass = "period-badge-soft";
        badgeLabel = "⚠️ Quasi allineata";
        isPartial = true;
        if (diffDays > 0) {
          statusText = `ADE arriva al ${formatDateIT(aMax)}, il gestionale risulta caricato solo fino al ${formatDateIT(gMax)} (differenza ~${absDiff} giorni).`;
        } else {
          statusText = `Il gestionale ha registrazioni fino al ${formatDateIT(gMax)}, ma ADE è fermo al ${formatDateIT(aMax)} (differenza ~${absDiff} giorni).`;
        }
      } else if (absDiff <= 30) {
        badgeClass = "period-badge-warn";
        badgeLabel = "⚠️ Copertura non allineata";
        isPartial = true;
        if (diffDays > 0) {
          statusText = `Attenzione: ADE arriva al ${formatDateIT(aMax)}, il gestionale risulta caricato solo fino al ${formatDateIT(gMax)}. Differenza: circa ${absDiff} giorni non coperti lato gestionale.`;
        } else {
          statusText = `Attenzione: il gestionale ha registrazioni fino al ${formatDateIT(gMax)}, ma ADE è fermo al ${formatDateIT(aMax)}. Differenza: circa ${absDiff} giorni non coperti lato ADE.`;
        }
      } else {
        badgeClass = "period-badge-critical";
        badgeLabel = "⛔ Copertura molto disallineata";
        isPartial = true;
        if (diffDays > 0) {
          statusText = `Grave disallineamento: ADE arriva al ${formatDateIT(aMax)}, mentre il gestionale è fermo al ${formatDateIT(gMax)}. Differenza stimata: ~${absDiff} giorni non coperti lato gestionale.`;
        } else {
          statusText = `Grave disallineamento: il gestionale arriva al ${formatDateIT(gMax)}, mentre ADE è fermo al ${formatDateIT(aMax)}. Differenza stimata: ~${absDiff} giorni non coperti lato ADE.`;
        }
      }

      if (isPartial) {
        statusText += " Il confronto di match potrebbe quindi essere solo parziale.";
      }
    } else {
      badgeClass = "period-badge-warn";
      badgeLabel = "⚠️ Copertura incompleta";
      statusText = "Solo uno tra ADE e gestionale ha date leggibili: il confronto potrebbe essere parziale.";
    }

    badgeEl.className = "period-badge " + badgeClass;
    badgeEl.textContent = badgeLabel;
    textEl.textContent = statusText;
  }

  // Per ADE (come prima)
  function findCol(headers, lowerHeaders, ...keys) {
    for (const k of keys) {
      const lk = k.toLowerCase();
      const idx = lowerHeaders.findIndex(t => t.includes(lk));
      if (idx >= 0) return headers[idx];
    }
    return null;
  }

  // ---------- build records ----------

  function buildAdeRecords(parsed) {
    const normalizeAdeHeader = (name) =>
      String(name || "")
        .replace(/^\uFEFF/, "")    // BOM UTF-8
        .replace(/\u00A0/g, " ")    // NBSP
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();

    const H = parsed.headers;
    const h = H.map(x => x.toLowerCase());
    const normH = H.map(normalizeAdeHeader);

  // 🔍 DEBUG: Stampa TUTTE le colonne del file ADE
  console.log("🔥🔥🔥 ===== INIZIO DEBUG FILE ADE ===== 🔥🔥🔥");
  console.log("📋 TUTTE LE COLONNE DEL FILE ADE (totale:", H.length, "colonne):");
  H.forEach((col, idx) => {
    console.log(`   [${idx}] "${col}"`);
  });
  console.log("🔥🔥🔥 ===== FINE LISTA COLONNE ===== 🔥🔥🔥");

  const cfg = getColumnConfig("ADE");

  const colNum   = cfg.num  || findCol(H, h, "numero fattura / documento", "numero fattura", "numero documento");

  // Date ADE: priorità Data ricezione, fallback Data emissione
  const colDataRicez = findCol(H, h, "data ricezione", "data di ricezione", "data ricez");
  const colDataEmiss = findCol(H, h, "data emissione", "data fattura", "data registrazione");
  const colData  = colDataRicez || colDataEmiss;
  if (!colData) {
    const msg = "❌ Colonna DATA ADE non trovata: atteso 'Data ricezione' o 'Data emissione'";
    console.error(msg, H);
    throw new Error(msg);
  }
  // Denominazione/P.IVA possono usare config se presente, ma gli importi NO
  const colPivaFor = cfg.piva || findCol(H, h, "partita iva fornitore", "partita iva cedente", "partita iva cedente / prestatore");
  const colPivaCli = findCol(H, h, "partita iva cliente");
  const colDenFor  = cfg.den || findCol(H, h, "denominazione fornitore", "denominazione cedente", "denominazione cedente / prestatore");
  
  // ============================================================
  // 🔧 FIX MAPPING IMPORTI ADE - indexOf() con nomi esatti
  // ============================================================
  // Usa SEMPRE i nomi ESATTI delle colonne ADE ufficiali
  // e IGNORA le vecchie config salvate per IMP e IVA
  // (così non rischiamo più di leggere colonne tipo "Aliquota IVA")
  // ============================================================

  // ✅ MAPPING IMPORTI ADE: IGNORA config salvate, usa SOLO header ufficiali
  const findAdeCol = (targetLabel) => {
    const target = normalizeAdeHeader(targetLabel);

    // 1) match esatto normalizzato
    let idx = normH.indexOf(target);
    if (idx !== -1) return idx;

    // 2) match parziale
    idx = normH.findIndex(nh => nh.includes(target));
    return idx;
  };

  let idxImpAde = findAdeCol("Imponibile/Importo (totale in euro)");
  let idxIvaAde = findAdeCol("Imposta (totale in euro)");

  if (idxImpAde === -1 || idxIvaAde === -1) {
    const missing = [];
    if (idxImpAde === -1) missing.push("Imponibile/Importo (totale in euro)");
    if (idxIvaAde === -1) missing.push("Imposta (totale in euro)");
    const msg = `❌ Colonne ADE mancanti: ${missing.join(", ")}. Verifica l'export ADE.`;
    console.error(msg, H);
    throw new Error(msg);
  }

  // Provo a individuare anche il TOTALE documento per un controllo di coerenza
  const idxTotAde = normH.findIndex((nh) => {
    if (!nh) return false;
    if (!nh.includes("totale")) return false;
    if (nh.includes("imponibile")) return false;
    if (nh.includes("imposta")) return false;
    if (nh.includes("iva")) return false;
    return true;
  });
  const colTotAde = idxTotAde !== -1 ? H[idxTotAde] : null;
  if (colTotAde) {
    console.log(`📌 Colonna TOT rilevata in ADE: "${colTotAde}"`);
  }

  const looksLikeAliquota = (idx) => {
    const sample = [];
    for (let i = 0; i < parsed.rows.length && sample.length < 50; i++) {
      const val = parsed.rows[i][H[idx]];
      if (val === undefined || val === null || val === "") continue;
      const n = parseNumberIT(val);
      if (!isNaN(n)) sample.push(n);
    }
    if (!sample.length) return false;

    const distinct = Array.from(new Set(sample.map(x => +x.toFixed(2))));
    const allSmall = sample.every(x => x > 0 && x <= 30);
    const commonAliq = new Set([0,4,5,10,18,20,21,22,5.5,4.5,15,16,6,7,8,9]);
    const mostlyAliq = sample.filter(x => commonAliq.has(+x.toFixed(1)) || commonAliq.has(+x.toFixed(0))).length >= Math.max(10, sample.length * 0.7);
    return allSmall && mostlyAliq && distinct.length <= 12;
  };

  const sampleRows = parsed.rows.slice(0, Math.min(200, parsed.rows.length));
  const calcMismatchWithTotal = (impIdx) => {
    if (!colTotAde) return { rate: 0, count: 0, mismatch: 0 };
    let count = 0;
    let mismatch = 0;
    const colImpName = H[impIdx];
    sampleRows.forEach(r => {
      const tot = parseNumberIT(r[colTotAde]);
      if (isNaN(tot)) return;
      const imp = parseNumberIT(r[colImpName]);
      const iva = parseNumberIT(r[H[idxIvaAde]]);
      if (isNaN(imp) || isNaN(iva)) return;
      count++;
      if (Math.abs((imp + iva) - tot) > 0.99) mismatch++;
    });
    return {
      rate: count ? mismatch / count : 0,
      count,
      mismatch
    };
  };

  const colNameIsAmountCandidate = (nh) => {
    if (!nh) return false;
    if (nh.includes("aliquota") || nh.includes("iva")) return false;
    if (nh.includes("imposta")) return false;
    if (nh.includes("piva") || nh.includes("p.iva") || nh.includes("partita")) return false;
    if (nh.includes("data") || nh.includes("numero") || nh.includes("protocol") || nh.includes("trasm")) return false;
    if (nh.includes("denomin") || nh.includes("ragione") || nh.includes("codice")) return false;
    return true;
  };

  // Se imp+iva non torna con il totale ADE su molte righe, prova a rimappare l'imponibile
  const currentMismatch = calcMismatchWithTotal(idxImpAde);
  if (colTotAde && currentMismatch.count >= 5 && currentMismatch.rate > 0.30) {
    console.warn(`⚠️ ADE: imp+iva non coerenti con il Totale documento (${(currentMismatch.rate * 100).toFixed(1)}% di righe con scarto > 1€). Provo a rimappare l'Imponibile.`);

    let bestIdx = idxImpAde;
    let bestScore = currentMismatch;

    normH.forEach((nh, idx) => {
      if (idx === idxImpAde || idx === idxIvaAde || idx === idxTotAde) return;
      if (!colNameIsAmountCandidate(nh)) return;

      const score = calcMismatchWithTotal(idx);
      if (score.count >= 5 && score.rate + 0.0001 < bestScore.rate) {
        bestIdx = idx;
        bestScore = score;
      }
    });

    if (bestIdx !== idxImpAde && bestScore.rate + 0.0001 < currentMismatch.rate) {
      console.warn(`✅ ADE: rimappata colonna Imponibile su "${H[bestIdx]}" usando il Totale documento (righe incoerenti: ${(bestScore.rate * 100).toFixed(1)}% vs ${(currentMismatch.rate * 100).toFixed(1)}%).`);
      idxImpAde = bestIdx;
    } else {
      const msg = "❌ ADE: Imponibile+IVA non coerenti con il Totale documento. Interrompo per evitare import errati.";
      console.error(msg, { currentMismatch, colImp: H[idxImpAde], colIva: H[idxIvaAde], colTot: colTotAde });
      throw new Error(msg);
    }
  }

  const colImp = H[idxImpAde];
  const colIva = H[idxIvaAde];

  if (looksLikeAliquota(idxImpAde)) {
    const msg = "❌ ADE: la colonna selezionata per Imponibile sembra un'ALIQUOTA IVA (valori tutti piccoli).";
    console.error(msg, { col: colImp });
    throw new Error(msg);
  }
  if (looksLikeAliquota(idxIvaAde)) {
    const msg = "❌ ADE: la colonna selezionata per IVA sembra un'ALIQUOTA IVA (valori tutti piccoli).";
    console.error(msg, { col: colIva });
    throw new Error(msg);
  }
  
  const colTipo  = findCol(H, h, "tipo documento");

  // 🔍 DEBUG: Log mapping colonne per diagnostica
  console.log("📋 MAPPING FINALE COLONNE ADE:");
  console.log(`   Numero Fattura: "${colNum}"`);
  console.log(`   Data (priorità ricezione): "${colDataRicez || colDataEmiss}"`);
  console.log(`   P.IVA Fornitore: "${colPivaFor}"`);
  console.log(`   Denominazione: "${colDenFor}"`);
  console.log(`   ⭐ Imponibile: "${colImp}"`);
  console.log(`   ⭐ IVA: "${colIva}"`);
  console.log(`   Tipo Documento: "${colTipo}"`);

  updateColumnConfig("ADE", {
    num: colNum,
    data: colData,
    piva: colPivaFor || colPivaCli,
    den: colDenFor,
    imp: colImp,
    iva: colIva
  });

  const recs = [];
  for (const r of parsed.rows) {
    const num  = r[colNum]    || "";
    const den  = r[colDenFor] || "";
    const piva = r[colPivaFor] || r[colPivaCli] || "";

    // Tipo documento ADE normalizzato (serve per gestione segno NC)
    const tipoDocRaw = (colTipo ? (r[colTipo] || "") : "").toString().toUpperCase();

    // 🔍 DEBUG FATTURA SPECIFICA - TRACCIA TUTTO IL PROCESSO
    if (num === "2528012466" || String(num).includes("2528012466")) {
      console.log("🔥🔥🔥 INIZIO PARSING FATTURA 2528012466 🔥🔥🔥");
      console.log("   Riga grezza completa:", r);
      console.log("   Numero fattura:", num);
      console.log("   Denominazione:", den);
      console.log("   P.IVA:", piva);
    }

    // ===== GESTIONE DATE ADE - FORMATO ITALIANO RIGIDO =====
    // Le date ADE arrivano SEMPRE in formato italiano DD/MM/YYYY
    // Usiamo normalizeDateItalyStrict per parsing rigido senza ambiguità
    const dataRawRicez = colDataRicez ? (r[colDataRicez] || "") : "";
    const dataRawEmiss = colDataEmiss ? (r[colDataEmiss] || "") : "";

    // ISO separati (servono per scelta base-data IVA)
    const dataRicezIso = dataRawRicez ? normalizeDateFlexible(dataRawRicez) : "";
    const dataEmissIso = dataRawEmiss ? normalizeDateFlexible(dataRawEmiss) : "";

    // Data per MATCH (priorità emissione, fallback ricezione)
    const dataRaw = dataRawEmiss || dataRawRicez || "";
    const dataIso = dataEmissIso || dataRicezIso || "";
    const data = dataIso ? parseDateFlexible(dataIso) : null;
    const dataStr = dataIso ? formatDateIT(dataIso) : "";
    const dateSource = dataEmissIso ? "emissione" : (dataRicezIso ? "ricezione" : "nessuna");

    // Data per IVA (default ricezione, fallback emissione)
    const vatDateIso = dataRicezIso || dataEmissIso || "";
    const vatDate = vatDateIso ? parseDateFlexible(vatDateIso) : null;


    // ===== GESTIONE IMPORTI (NESSUN AUTO-FIX) =====
    const impRaw = r[colImp] || "";
    const ivaRaw = r[colIva] || "";

    // 🔍 DEBUG: Log valori grezzi
    if (num === "2528012466" || String(num).includes("2528012466")) {
      console.log(`🔍 DEBUG ADE Fattura ${num} - LETTURA COLONNE:`);
      console.log(`   📊 Colonna Imponibile: "${colImp}"`);
      console.log(`   📊 Valore grezzo: "${impRaw}"`);
      console.log(`   📊 Colonna IVA: "${colIva}"`);
      console.log(`   📊 Valore grezzo IVA: "${ivaRaw}"`);
      
      // Mostra TUTTA la riga
      console.log(`   📋 RIGA COMPLETA:`, r);
      H.forEach((colName, idx) => {
        console.log(`      [${idx}] "${colName}" = "${r[colName]}"`);
      });
    }
    
    // ORA fai il parsing con il valore corretto
    const impParsed = parseNumberIT(impRaw);
    const ivaParsed = cleanIVAValue(ivaRaw, piva);

    // ============================================================
    // ✅ NORMALIZZAZIONE SEGNO (ULTRA FIX)
    // - ADE è source of truth (mai ricalcolare IVA)
    // - Evita double-negation su TD04
    // - Se ADE arriva già negativo (imp/iva), rispettalo SEMPRE anche se tipo documento è sbagliato/vuoto
    // ============================================================

    const tipoDoc = (tipoDocRaw || "").trim().toUpperCase();
    const isTD04 = (tipoDoc === "TD04");

    // segnali “nota credito” anche senza TD04
    const isNCByTipo = tipoDoc.includes("TD04") || tipoDoc.includes("NOTA") || tipoDoc.includes("NC");
    const isNCBySign = (impParsed < 0) || (ivaParsed < 0);

    let imp = impParsed;
    let iva = ivaParsed;
    let adeTechNote = "";

    // 1) Se è TD04 (o nota credito per tipo) e gli importi sono positivi, rendili negativi
    //    Se sono già negativi, NON toccare (no double-negation)
    if (isTD04 || isNCByTipo) {
      if (imp > 0) imp = -imp;
      if (iva > 0) iva = -iva;
    }

    // 2) Se ADE arriva già con segno negativo (anche senza tipoDoc), lo rispettiamo SEMPRE
    //    (questa è la parte che ti salva TIM/ENI e tutte le righe con tipo documento ambiguo)
    if (isNCBySign) {
      imp = -Math.abs(imp);
      iva = -Math.abs(iva);
    }

    // LOG DIAGNOSTICO (solo per righe potenzialmente NC)
    if (isTD04 || isNCByTipo || isNCBySign) {
      console.log(
        `🧾 NC_CHECK [${num}] tipo="${tipoDocRaw}" impRaw=${impRaw} ivaRaw=${ivaRaw} -> imp=${imp} iva=${iva}`
      );
    }

    // Totale coerente con i valori ADE normalizzati
    let tot = +(imp + iva).toFixed(2);
    
    // 🔍 DEBUG: Valori dopo parsing
    if (num === "2528012466" || String(num).includes("2528012466")) {
      console.log(`   ✅ Dopo parsing: imp=${imp}, iva=${iva}, tot=${tot}`);
    }
    
    // ✅ VALIDAZIONE IMPORTI
    const validated = fixAndValidateAmounts({ imp, iva, tot }, "ADE", num);
    tot = validated.tot;

    // 🚨 Validazione forte per imponibili sospettosamente bassi
    if (imp <= 30 && (iva >= 100 || (imp + iva) >= 100)) {
      console.error(`🚨 ADE [${num}] imponibile sospetto troppo basso rispetto all'IVA/totale`, {
        num,
        den,
        piva,
        imp,
        iva,
        tot,
        colImp,
        colIva,
        tipoDoc: tipoDocRaw,
        row: r
      });
    }
    
    // 🔍 DEBUG FINALE: Valori nel record
    if (num === "2528012466" || String(num).includes("2528012466")) {
      console.log(`   ✅ RECORD FINALE: imp=${imp}, iva=${iva}, tot=${tot}`);
      console.log("🔥🔥🔥 FINE PARSING FATTURA 2528012466 🔥🔥🔥");
    }
    
    recs.push({
      src: "ADE",
      num,
      numDigits: normalizeInvoiceNumber(num, false),
      den,
      denNorm: normalizeName(den),
      piva,
      pivaDigits: onlyDigits(piva),
      dataRawRicez,
      dataRawEmiss,
      dataRicezIso,
      dataEmissIso,

      // Matching (data fattura): emissione -> fallback ricezione
      dataRaw,
      dataIso,
      data,
      dataStr,
      dateSource,

      // IVA (periodo): ricezione -> fallback emissione
      vatDateIso,
      vatDate,
      imp,
      iva,
      tot,
      tipoDoc: tipoDocRaw,
      isForeign: isForeignPiva(piva),
      techNote: adeTechNote
    });
  }

  // Check finale coerenza importi ADE (con NC negative)
  const sumImp = recs.reduce((acc, r) => acc + (r.imp || 0), 0);
  const sumIva = recs.reduce((acc, r) => acc + (r.iva || 0), 0);
  console.log("✅ ADE check interno - somma imponibile (con NC negative):", sumImp.toFixed(2));
  console.log("✅ ADE check interno - somma IVA (con NC negative):", sumIva.toFixed(2));

  return recs;
}

   // ---------- GESTIONALE FATTURE ACQUISTO (FILE B) ----------
function buildGestAcqRecords(parsed, options = {}) {
  const { isNotesCredit = false, fileName = "" } = options;
  const inferredNC = /nota|note|nc|credito/i.test(fileName || "");
  const isGestNotesCredit = isNotesCredit || inferredNC;
  const H = parsed.headers;
  const h = H.map(x => x.toLowerCase());

  // mapping base da JSON (columns.GEST_B.*)
  const cfgGest = getColumnConfig("GEST_B");

  // ----------------------------------------------------
  // NUMERO DOCUMENTO / NUMERO FATTURA (GESTIONALE)
  // ----------------------------------------------------

  // colonne speciali
  const colNumDocEsteso = H.find(hdr =>
    String(hdr || "").toLowerCase().replace(/[^a-z0-9]/g, "") === "numerodocesteso"
  );

  // colonna "NUMEROALFA" molto usata in CONTEL
  const colNumeroAlfa = H.find(hdr =>
    String(hdr || "").toLowerCase().replace(/[^a-z0-9]/g, "") === "numeroalfa"
  );

  const colDocumento = H.find(hdr =>
    String(hdr || "").toLowerCase().replace(/[^a-z0-9]/g, "") === "documento"
  );

  let colNumConfig = cfgGest.num || null;

  // se la config punta a una DATA, la scarto
  if (colNumConfig && isDateHeaderName(colNumConfig)) {
    colNumConfig = null;
  }

  // Ordine di scelta per il NUMERO GESTIONALE:
  // 1) NUMEROALFA
  // 2) Numero doc.esteso
  // 3) Config salvata
  // 4) Documento
  // 5) altre colonne candidate
  if (colNumeroAlfa) {
    colNumConfig = colNumeroAlfa;
  } else if (colNumDocEsteso) {
    colNumConfig = colNumDocEsteso;
  } else if (!colNumConfig) {
    if (colDocumento) {
      colNumConfig = colDocumento;
    } else {
      colNumConfig = findColGest(H, h, [
        "numerofattura",
        "numero fattura",
        "nreggestionale",
        "numeroregistrazione",
        "numreg",
        "num",
        "doc"
      ]);
    }
  }

 // P.IVA fornitore
const colPiva = cfgGest.piva || findColGest(H, h, [
  "partita iva",
  "p.iva",
  "piva",
  "p iva",
  "pi",
  "cfpiva",
  "p. iva fornitore",
  "partitaiva"
]);

const colData = cfgGest.data || findColGest(H, h, [
  "data",
  "data fattura",
  "data doc",
  "data documento",
  "data registrazione",
  "datareg"
]);

const colDen = cfgGest.den || findColGest(H, h, [
  "ragione sociale",
  "ragionesociale",
  "fornitore",
  "descrizione conto",
  "descrizione conto dare",
  "descrizione conto avere",
  "descr"
]);

const colImp = cfgGest.imp || findColGest(H, h, [
  "tot_imp",
  "imponibile",
  "totale imponibile",
  "totale",
  "totale fattura",
  "importo documento",
  "importo",
  "totale doc"
]);

const colIva = cfgGest.iva || findColGest(H, h, [
  "iva",
  "imposta",
  "importoiva",
  "tot_iva"
]);

const colTot = cfgGest.tot || findColGest(H, h, [
  "totale",
  "totalefattura",
  "tot_fattura",
  "importo documento",
  "importo",
  "totale doc",
  "totdoc"
]);

const colDoc = findColGest(H, h, [
  "documento",
  "tipo documento",
  "doc"
]);

const recs = [];
for (const r of parsed.rows) {

  // numero documento “intelligente” da colNumConfig / Documento
  let num = "";
  if (colNumConfig) {
    const raw = r[colNumConfig] || "";
    const s = String(raw).trim();

    // se è evidentemente una data (tipo 21/03/2025), la scarto
    if (!/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s)) {
      num = s;
    }
  }

  // se ancora vuoto, provo a ripiegare su "Documento"
  if (!num && colDocumento) {
    const s2 = String(r[colDocumento] || "").trim();
    if (s2 && !/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s2)) {
      num = s2;
    }
  }

  const den  = colDen  ? (r[colDen]  || "") : "";
  const piva = colPiva ? (r[colPiva] || "") : "";

  // ===== GESTIONE DATE GESTIONALE - FORMATO FLESSIBILE =====
  // Il gestionale può avere vari formati: DD/MM/YYYY, MM/DD/YYYY, ISO
  // Usiamo normalizeDateFlexible per gestire l'ambiguità
  const dataRaw = colData ? (r[colData] || "") : "";       // Valore originale dal CSV
  const dataIso = normalizeDateFlexible(dataRaw);            // "YYYY-MM-DD" con parsing flessibile
  const data = parseDateFlexible(dataIso);                   // Oggetto Date per confronti (usa ISO già normalizzato)
  const dataStr = dataIso ? formatDateIT(dataIso) : "";     // "DD/MM/YYYY" per display

  // ===== GESTIONE IMPORTI CON VALIDAZIONE IVA =====
  let imp = colImp ? parseNumberIT(r[colImp]) : 0;
  let iva = colIva ? cleanIVAValue(r[colIva], piva) : 0;  // Valida IVA vs P.IVA

  let tot = 0;
  if (colTot) {
    tot = parseNumberIT(r[colTot]);
  } else {
    tot = +(imp + iva).toFixed(2);
  }

  // ✅ FIX NC Gestionale: se il file è una Nota Credito, forziamo segno negativo
  // (Contel talvolta esporta NC con importi positivi)
  if (isGestNotesCredit === true) {
    imp = -Math.abs(imp);
    iva = -Math.abs(iva);
    tot = -Math.abs(tot);
  }
  
  // 🚨 VALIDAZIONE MAPPING ERRATO GESTIONALE - Rileva casi impossibili
  // Se imponibile < 10 ma totale > 500, probabilmente la colonna è sbagliata
  if (imp > 0 && iva > 0 && imp < 10 && tot > 500) {
    console.error(`🚨 GEST [${num}]: MAPPING ERRATO IMPONIBILE!`);
    console.error(`   Imponibile: ${imp} € (troppo basso!)`);
    console.error(`   IVA: ${iva} €`);
    console.error(`   Totale: ${tot} €`);
    console.error(`   ⚠️ Probabilmente stai leggendo da colonna sbagliata`);
    console.error(`   Colonna Imponibile usata: "${colImp}"`);
    console.error(`   Colonna IVA usata: "${colIva}"`);
    console.error(`   Riga grezza:`, r);
  }
  
  // Se IVA è zero ma imponibile e totale esistono, potrebbe essere mapping errato
  if (imp > 0 && iva === 0 && tot > imp + 10) {
    console.warn(`⚠️ GEST [${num}]: IVA = 0 ma totale (${tot}) >> imponibile (${imp})`);
    console.warn(`   Possibile mapping errato colonna IVA: "${colIva}"`);
  }
  
  // ✅ VALIDAZIONE E FIX IMPORTI GESTIONALE
  const validated = fixAndValidateAmounts({ imp, iva, tot }, "GEST", num);
  tot = validated.tot;  // Usa il totale validato/corretto

  // per le fatture acquisto standard consideriamo NON NC
  const isNC = false;

  recs.push({
    src: "GEST",
    num,
    numDigits: normalizeInvoiceNumber(num, false),
    den,
    denNorm: normalizeName(den),
    piva,
    pivaDigits: onlyDigits(piva),
    dataRaw,   // valore originale dal file
    dataIso,   // "YYYY-MM-DD" normalizzato
    data,      // oggetto Date per confronti
    dataStr,   // "DD/MM/YYYY" per display
    imp,
    iva,
    tot,
    isNC,
    isForeign: isForeignPiva(piva)
  });
}

return recs;
}

// Dice se un header sembra una colonna di DATA (per non usarla come numero doc)
function isDateHeaderName(hdr) {
  const n = String(hdr || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

  return (
    n.startsWith("data") ||         // data, datareg, datadoc, dataregistrazione...
    n === "datadoc" ||
    n === "datareg" ||
    n === "dataregistrazione"
  );
}

 // ---------- GESTIONALE NOTE CREDITO (FILE C) ----------
function buildGestNcRecords(parsed, options = {}) {
  const { isNotesCredit = true, fileName = "" } = options;
  const inferredNC = /nota|note|nc|credito/i.test(fileName || "");
  const isGestNotesCredit = isNotesCredit || inferredNC;
  const H = parsed.headers;
  const h = H.map(x => x.toLowerCase());

  const cfgGest = getColumnConfig("GEST_C");

  // 🔢 NUMERO DOCUMENTO / NUMERO FATTURA NC
  const colNumEst = cfgGest.num || findColGest(H, h, [
    "numero fattura",
    "nreggestionale",
    "numeroregistrazione",
    "numero doc.esteso",
    "numero documento",
    "numero doc",
    "n.doc",
    "doc",
    "num"
  ]);

  // P.IVA fornitore
  const colPiva = cfgGest.piva || findColGest(H, h, [
    "partita iva",
    "p.iva",
    "piva",
    "p iva",
    "pi",
    "cfpiva",
    "p. iva fornitore",
    "partitaiva"
  ]);

  // DATA documento / registrazione
  const colData = cfgGest.data || findColGest(H, h, [
    "data",
    "data fattura",
    "data doc",
    "data documento",
    "data registrazione",
    "datareg"
  ]);

  // DENOMINAZIONE fornitore
  const colDen = cfgGest.den || findColGest(H, h, [
    "ragione sociale",
    "ragionesociale",
    "fornitore",
    "descrizione conto",
    "descrizione conto dare",
    "descrizione conto avere",
    "descr"
  ]);

  // IMPONIBILE
  const colImp = cfgGest.imp || findColGest(H, h, [
    "tot_imp",
    "imponibile",
    "totale imponibile",
    "totale",
    "totale fattura",
    "importo documento",
    "importo",
    "totale doc"
  ]);

  // IVA
  const colIva = cfgGest.iva || findColGest(H, h, [
    "iva",
    "imposta",
    "importoiva",
    "tot_iva"
  ]);

  // TOTALE documento
  const colTot = cfgGest.tot || findColGest(H, h, [
    "totale",
    "totalefattura",
    "tot_fattura",
    "importo documento",
    "importo",
    "totale doc",
    "totdoc"
  ]);

  // TIPO DOCUMENTO
  const colDoc = findColGest(H, h, [
    "documento",
    "tipo documento",
    "doc"
  ]);

  // DESCRIZIONE OPERAZIONE
  const colDescrOp = findColGest(H, h, [
    "descrizione operazione",
    "descrizione",
    "descr"
  ]);

  const recs = [];
  for (const r of parsed.rows) {

    // 🔢 numero documento NC “intelligente” (non deve essere una data!)
    let num = "";
    if (colNumEst) {
      const raw = r[colNumEst] || "";
      const s = String(raw).trim();

      // se è evidentemente una data (tipo 21/03/2025), la scarto
      if (!/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s)) {
        num = s;
      }
    }

    // se ancora vuoto, provo a ripiegare su "Documento"
    if (!num && colDoc) {
      const s2 = String(r[colDoc] || "").trim();
      if (s2 && !/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s2)) {
        num = s2;
      }
    }

    const den  = colDen  ? (r[colDen]  || "") : "";
    const piva = colPiva ? (r[colPiva] || "") : "";

    // ===== GESTIONE DATE GESTIONALE NC - FORMATO FLESSIBILE =====
    // Le note credito gestionale possono avere vari formati
    // Usiamo normalizeDateFlexible per gestire l'ambiguità
    const dataRaw = colData ? (r[colData] || "") : "";       // Valore originale dal CSV
    const dataIso = normalizeDateFlexible(dataRaw);            // "YYYY-MM-DD" con parsing flessibile
    const data = parseDateFlexible(dataIso);                   // Oggetto Date per confronti (usa ISO già normalizzato)
    const dataStr = dataIso ? formatDateIT(dataIso) : "";     // "DD/MM/YYYY" per display

    // ===== GESTIONE IMPORTI CON VALIDAZIONE IVA =====
    let imp = colImp ? parseNumberIT(r[colImp]) : 0;
    let iva = colIva ? cleanIVAValue(r[colIva], piva) : 0;  // Valida IVA vs P.IVA

    let tot = 0;
    if (colTot) {
      tot = parseNumberIT(r[colTot]);
    } else {
      tot = +(imp + iva).toFixed(2);
    }

    // ✅ FIX NC Gestionale: se sto caricando il file "NOTE CREDITO", il segno deve essere negativo
    // (in Contel spesso le NC devono essere negative ma possono arrivare positive dal file)
    if (isGestNotesCredit === true) {
      imp = -Math.abs(imp);
      iva = -Math.abs(iva);
      tot = -Math.abs(tot);
    }

    const docField = colDoc     ? String(r[colDoc] || "").toUpperCase()     : "";
    const descrOp  = colDescrOp ? String(r[colDescrOp] || "").toUpperCase() : "";

    const isNC = (docField && docField.startsWith("NR")) ||
      descrOp.includes("NOTA") ||
      descrOp.includes("NC");

    recs.push({
      src: "GEST",
      num,
      numDigits: normalizeInvoiceNumber(num, false),
      den,
      denNorm: normalizeName(den),
      piva,
      pivaDigits: onlyDigits(piva),
      dataRaw,   // valore originale dal file
      dataIso,   // "YYYY-MM-DD" normalizzato
      data,      // Date per il match
      dataStr,   // "DD/MM/YYYY" per display
      imp,
      iva,
      tot,
      isNC,
      isForeign: isForeignPiva(piva)
    });
  }

  return recs;
}

  // verifica che ADE e GEST siano lo stesso fornitore
    function sameSupplier(a, g) {
      const aPiva = onlyDigits(a.piva || "");
      const gPiva = onlyDigits(g.piva || "");
  
      // Caso 1: Entrambe le P.IVA sono presenti e valide
      if (aPiva && gPiva) {
        // Se le P.IVA coincidono, è lo stesso fornitore. Fine.
        if (aPiva === gPiva) return true;
        
        // Se le P.IVA sono DIVERSE, siamo più severi.
        // Richiediamo che i nomi siano MOLTO simili per considerarlo un match.
        // Uniamo il controllo su parole comuni (namesSimilar) e la distanza di Levenshtein (isFuzzyMatch).
        if (namesSimilar(a.den, g.den) || isFuzzyMatch(a.den, g.den, MATCH_CONFIG.FUZZY_SAME_PIVA)) {
          return true;
        }
        return false; // P.IVA e nomi sono diversi -> non è lo stesso fornitore.
      }
  
      // Caso 2: Almeno una P.IVA manca. Ci affidiamo completamente ai nomi.
      // Usiamo una logica combinata più permissiva per recuperare più match possibili.
      return namesSimilar(a.den, g.den) || isFuzzyMatch(a.den, g.den, MATCH_CONFIG.FUZZY_DIFF_PIVA);
    }

    // ---------- matching ----------

// ============================================================
// LOGICA DI MATCH MULTILIVELLO (Refactoring Dic 2025)
// ============================================================
// Implementa i nuovi livelli richiesti per il matching
// ============================================================

/**
 * Normalizza numero fattura per confronti fuzzy
 * Rimuove: zeri iniziali, spazi, caratteri speciali comuni
 * @param {string} num - Numero fattura grezzo
 * @returns {string} - Numero normalizzato
 */
function normalizeInvoiceForMatch(num) {
  if (!num) return "";

  // 🔧 FIX: rimuove apici/virgolette (ADE spesso ha 2523335617' o '8U00126007')
  let s = String(num)
    .replace(/["']/g, "")
    .toUpperCase()
    .trim();

  // Togli separatori e schifezze comuni
  s = s.replace(/[\/\-_\s.*@#]/g, "");

  // 🔧 EXTRA SAFE: se rimane qualcosa di non alfanumerico, lo elimino
  s = s.replace(/[^A-Z0-9]/g, "");

  // Rimuovi zeri iniziali dopo eventuale prefisso lettere
  // Es: "F000123" -> "F123", "000586" -> "586"
  s = s.replace(/^([A-Z]*)0+/, "$1");

  return s;
}

// ===============================
// 🔐 STEP ULTRA: MATCH per chiavi forti (NUM + DATA + TOTALE)
// ===============================
function strongKeyMatch(a, g) {
  if (!a || !g) return false;

  const normNum = (v) => normalizeInvoiceForMatch(v || "");
  const numOk =
    normNum(a.num) === normNum(g.num) ||
    (String(a.numDigits || "") && String(g.numDigits || "") && String(a.numDigits) === String(g.numDigits));

  // Data: confronta oggetto Date se disponibile, altrimenti stringhe normalizzate
  const sameDate = () => {
    if (a.data && g.data) return a.data.getTime() === g.data.getTime();
    if (a.dataIso && g.dataIso) return a.dataIso === g.dataIso;
    if (a.dataStr && g.dataStr) return a.dataStr === g.dataStr;
    return false;
  };
  const dateOk = sameDate();

  const totTol = MATCH_CONFIG.TOL_TOT || MATCH_CONFIG.TOLERANCE_ROUNDING || 0.05;
  const totOk = Math.abs((a.tot || 0) - (g.tot || 0)) <= totTol;

  return numOk && dateOk && totOk;
}

// ============================================================
// 🎯 MATCHING LEVELS (L1-L4) - Dic 2025
// ============================================================

/**
 * Verifica se due numeri fattura sono "fuzzy match"
 * Gestisce: prefissi, suffissi, zeri iniziali
 */
function invoiceNumbersFuzzy(num1, num2) {
  if (!num1 || !num2) return false;
  
  const n1 = normalizeInvoiceForMatch(num1);
  const n2 = normalizeInvoiceForMatch(num2);
  
  if (n1 === n2) return true;
  
  // Estrai solo le cifre
  const digits1 = onlyDigits(n1);
  const digits2 = onlyDigits(n2);
  
  if (!digits1 || !digits2) return false;
  
  // Se le cifre sono identiche (ignorando prefissi/suffissi)
  if (digits1 === digits2) return true;
  
  // Se uno contiene l'altro (min 4 cifre per evitare falsi positivi)
  if (digits1.length >= 4 && digits2.length >= 4) {
    if (digits1.includes(digits2) || digits2.includes(digits1)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Verifica match livello L1: esatto forte
 */
function isL1Match(a, g) {
  // Numero esatto (dopo normalizzazione base)
  const aNum = normalizeInvoiceForMatch(a.num);
  const gNum = normalizeInvoiceForMatch(g.num);
  if (!aNum || !gNum || aNum !== gNum) return false;
  
  // P.IVA uguale
  const aPiva = normalizePIVA(a.piva);
  const gPiva = normalizePIVA(g.piva);
  if (!aPiva || !gPiva || aPiva !== gPiva) return false;
  
  // Totale esatto
  if (Math.abs(a.tot - g.tot) >= MATCH_CONFIG.TOLERANCE_EXACT) return false;
  
  return true;
}

/**
 * Verifica match livello L2: esatto con tolleranza importo
 */
function isL2Match(a, g) {
  // Numero esatto
  const aNum = normalizeInvoiceForMatch(a.num);
  const gNum = normalizeInvoiceForMatch(g.num);
  if (!aNum || !gNum || aNum !== gNum) return false;
  
  // P.IVA uguale
  const aPiva = normalizePIVA(a.piva);
  const gPiva = normalizePIVA(g.piva);
  if (!aPiva || !gPiva || aPiva !== gPiva) return false;
  
  // Totale con tolleranza ±1 euro
  if (Math.abs(a.tot - g.tot) > MATCH_CONFIG.TOLERANCE_L2_AMOUNT) return false;
  
  return true;
}

/**
 * Verifica match livello L3: numero fuzzy
 */
function isL3Match(a, g) {
  // Numero fuzzy (gestisce prefissi, zeri iniziali, ecc.)
  if (!invoiceNumbersFuzzy(a.num, g.num)) return false;
  
  // P.IVA uguale
  const aPiva = normalizePIVA(a.piva);
  const gPiva = normalizePIVA(g.piva);
  if (!aPiva || !gPiva || aPiva !== gPiva) return false;
  
  // Totale con tolleranza ±1 euro
  if (Math.abs(a.tot - g.tot) > MATCH_CONFIG.TOLERANCE_L2_AMOUNT) return false;
  
  return true;
}

/**
 * Verifica match livello L4: P.IVA + totale + data
 */
function isL4Match(a, g) {
  // P.IVA uguale
  const aPiva = normalizePIVA(a.piva);
  const gPiva = normalizePIVA(g.piva);
  if (!aPiva || !gPiva || aPiva !== gPiva) return false;
  
  // Totale esatto o molto vicino
  if (Math.abs(a.tot - g.tot) > MATCH_CONFIG.TOLERANCE_L4_AMOUNT) return false;
  
  // Date vicine (±3 giorni)
  if (!a.data || !g.data) return false;
  if (!datesClose(a.data, g.data, MATCH_CONFIG.DATE_TOLERANCE_DAYS)) return false;
  
  return true;
}

// ============================================================
// 🔧 HELPER FUNCTIONS PER MATCHING (Dic 2025)
// ============================================================

/**
 * Corregge automaticamente date invertite (ADE sempre corretto)
 * @param {Object} adeRecord - Record ADE
 * @param {Object} gestRecord - Record Gestionale
 */
function correctInvertedDates(adeRecord, gestRecord) {
  if (!adeRecord?.data || !gestRecord?.data) return;
  
  const dayA = adeRecord.data.getDate();
  const monthA = adeRecord.data.getMonth() + 1;
  const dayG = gestRecord.data.getDate();
  const monthG = gestRecord.data.getMonth() + 1;
  
  // Rileva inversione (evita palindromi come 07/07)
  if (dayA === monthG && monthA === dayG && dayA !== monthA) {
    console.warn(`🔄 CORREZIONE DATA INVERTITA: GEST aveva ${formatDateIT(gestRecord.data)}, corretto in ${formatDateIT(adeRecord.data)} (come ADE)`);
    gestRecord.data = new Date(adeRecord.data);
    gestRecord.dataStr = formatDateIT(adeRecord.data);
  }
}

/**
 * Classifica la nota di match in base alla differenza importi
 * @param {number} diffTotale - Differenza tra ADE e GEST
 * @param {Object} adeRecord - Record ADE
 * @param {Object} gestRecord - Record Gestionale
 * @param {Object} ncInfo - Info nota di credito invertita
 * @param {string} criterio - Criterio di matching usato
 * @returns {Object} - { status: string, note: string }
 */
function classifyMatchNote(diffTotale, adeRecord, gestRecord, ncInfo, criterio) {
  const diffAssoluta = Math.abs(diffTotale);
  let status = "MATCH_FIX";
  let note = "";
  
  // 1. Totali coincidenti
  if (diffAssoluta <= MATCH_CONFIG.TOLERANCE_EXACT) {
    // Controllo speciale per totali bassi
    if (Math.abs(adeRecord.tot) < MATCH_CONFIG.THRESHOLD_LOW_TOTAL) {
      const ivaDiff = Math.abs((adeRecord.iva || 0) - (gestRecord.iva || 0));
      if (ivaDiff > MATCH_CONFIG.TOLERANCE_EXACT) {
        status = "MATCH_FIX";
        note = `⚠️ Totale nullo o molto basso (${adeRecord.tot.toFixed(2)}) ma IVA diverse ` +
               `(ADE: ${(adeRecord.iva || 0).toFixed(2)}, GEST: ${(gestRecord.iva || 0).toFixed(2)}). Verificare registrazione.`;
      } else {
        status = "MATCH_OK";
      }
    } else {
      status = "MATCH_OK";
    }
  }
  // 2. Segno opposto (probabile Nota di Credito)
  else if (ncInfo.isNCInvertita && ncInfo.kind === "SEGNO_OPPOSTO") {
    status = "MATCH_OK";
    note = `⚠ Probabile NOTA DI CREDITO: ADE e Gestionale hanno stessi importi in valore assoluto ma segno opposto. ${ncInfo.formula}`;
  }
  // 3. Nota di credito invertita (pattern storico IVA=0 nel gestionale)
  else if (ncInfo.isNCInvertita) {
    note = `⚠ POSSIBILE NOTA DI CREDITO GESTITA AL CONTRARIO NEL GESTIONALE: ${ncInfo.formula}`;
    status = "MATCH_FIX";
  }
  // 3. Arrotondamenti
  else if (diffAssoluta <= MATCH_CONFIG.TOLERANCE_ROUNDING) {
    note = `🟡 Differenza minima dovuta ad arrotondamenti (≤ 0,05 €). Scarto: ${diffTotale >= 0 ? '+' : ''}${diffTotale.toFixed(2)} €.`;
  }
  // 4. Differenza significativa
  else if (diffAssoluta > MATCH_CONFIG.THRESHOLD_SIGNIFICANT_DIFF) {
    note = `⚠ Differenza significativa tra ADE e gestionale: scarto di ${diffTotale >= 0 ? '+' : ''}${diffTotale.toFixed(2)} €. Verificare la registrazione contabile.`;
  }
  // 5. Altri casi
  else {
    note = `❌ IMPORTO DIVERSO: Totale ADE vs GEST non coincidono (scarto di ${diffTotale >= 0 ? '+' : ''}${diffTotale.toFixed(2)} €).`;
  }
  
  // Contesto criterio specifico
  if (criterio?.includes("2_NUM_COMPATIBILE+IMPORTO")) {
    note += " Numero fattura compatibile, ma importi diversi.";
  }
  
  return { status, note };
}

function matchRecords(adeList, gestList) {
  const ade = adeList
    .filter(r => !r.isForeign)    // <-- niente P.IVA estera ADE
    .map((r, idx) => ({ idx, ...r }));

  const gest = gestList
    .filter(r => !r.isForeign)    // <-- niente P.IVA estera Gest
    .map((r, idx) => ({ idx, ...r }));

  const matchedAde = new Set();
  const matchedGest = new Set();
  const results = [];

    function addMatch(a, g, type, criterio) {
      matchedAde.add(a.idx);
      matchedGest.add(g.idx);

      let status = (type === "OK" || type === "MATCH_OK") ? "MATCH_OK" : "MATCH_FIX";

      // Correzione date invertite
      correctInvertedDates(a, g);

      // Calcolo differenza
      const diffTotale = (a?.tot || 0) - (g?.tot || 0);
      const ncInfo = detectNCInvertita(a, g);
      
      // Classificazione
      const classification = classifyMatchNote(diffTotale, a, g, ncInfo, criterio);

      const record = {
        ADE: a,
        GEST: g,
        STATUS: classification.status,
        CRITERIO: criterio,
        DIFF_TOTALE: diffTotale,
        ORIGINE: "MATCH",
        NOTE_MATCH: classification.note,
        FLAG_REVISIONE: ""  // Inizializzato vuoto, sarà popolato dal post-processing
      };
      
      results.push(record);
      return record;  // Restituisce il record creato per permettere modifiche successive
    }

    function totEqual(aTot, gTot, toll = 0.05) {
      return Math.abs(aTot - gTot) < toll;
    }

    // ============================================================
    // 🎯 NUOVI LIVELLI DI MATCH (Dic 2025)
    // ============================================================
    // L1, L2, L3, L4 come richiesto nella specifica
    // Questi pass vengono eseguiti PRIMA dei vecchi per priorità
    // ============================================================

    // ==========================================
    // 🎯 LIVELLO L1: MATCH ESATTO FORTE
    // Numero esatto + P.IVA esatta + Totale esatto
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        const match = isL1Match(a, g);
        if (match && normalizePIVA(a.piva || "") === "8349560014") {
          console.log(`🔍 L1 MATCH CA AUTO BANK: ${a.num} - diff €${Math.abs(a.tot - g.tot).toFixed(2)}`);
        }
        return match;
      });

      if (matchGest) {
        addMatch(a, matchGest, "MATCH_OK", "L1_MATCH_ESATTO");
      }
    }

    // ==========================================
    // 🎯 LIVELLO L2: MATCH ESATTO CON TOLLERANZA IMPORTO
    // Numero esatto + P.IVA esatta + Totale ±1 euro
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        const match = isL2Match(a, g);
        if (match && normalizePIVA(a.piva || "") === "8349560014") {
          console.log(`🔍 L2 MATCH CA AUTO BANK: ${a.num} - diff €${Math.abs(a.tot - g.tot).toFixed(2)}`);
        }
        return match;
      });

      if (matchGest) {
        const diff = Math.abs(a.tot - matchGest.tot);
        const note = diff > 0.01 ? `Diff. €${diff.toFixed(2)}` : "";
        addMatch(a, matchGest, "MATCH_OK", `L2_IMPORTO_TOLLERATO ${note}`.trim());
      }
    }

    // ==========================================
    // 🎯 LIVELLO L3: MATCH FUZZY SU NUMERO FATTURA
    // Numero fuzzy (prefissi/zeri) + P.IVA esatta + Totale ±1 euro
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        const match = isL3Match(a, g);
        if (match && normalizePIVA(a.piva || "") === "8349560014") {
          console.log(`🔍 L3 MATCH CA AUTO BANK: ${a.num} - diff €${Math.abs(a.tot - g.tot).toFixed(2)}`);
        }
        return match;
      });

      if (matchGest) {
        addMatch(a, matchGest, "MATCH_OK", "L3_NUMERO_FUZZY");
      }
    }

    // ==========================================
    // 🎯 LIVELLO L4: MATCH PER P.IVA + TOTALE + DATA
    // Quando numero non affidabile: P.IVA + importo + data ±3gg
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        const match = isL4Match(a, g);
        if (match && normalizePIVA(a.piva || "") === "8349560014") {
          console.log(`🔍 L4 MATCH CA AUTO BANK: ${a.num} - diff €${Math.abs(a.tot - g.tot).toFixed(2)}`);
        }
        return match;
      });

      if (matchGest) {
        addMatch(a, matchGest, "MATCH_OK", "L4_FORNITORE_DATA_TOTALE");
      }
    }

    // ==========================================
    // 🔍 PASS L5: MATCH RILASSATO - Stesso numero fattura + P.IVA, importi/date diversi
    // PRIORITÀ ALTA: Cattura casi con stesso numero fattura e P.IVA ma importi/date differenti
    // Li classifica come MATCH con FLAG_REVISIONE invece di lasciarli come SOLO_ADE/SOLO_GEST
    // ESEGUITO PRIMA DEI PASS LEGACY per garantire che numero+PIVA abbiano priorità assoluta
    // ==========================================
    {
      const adeStillUnmatched = ade.filter(a => !matchedAde.has(a.idx));
      const gestStillUnmatched = gest.filter(g => !matchedGest.has(g.idx));
      
      console.log(`🔍 PASS L5: ${adeStillUnmatched.length} ADE non matchati, ${gestStillUnmatched.length} GEST non matchati`);

      let l5MatchCount = 0;
      
      for (const a of adeStillUnmatched) {
        // DEBUG: Log per ogni record ADE con P.IVA 8349560014
        if (normalizePIVA(a.piva || "").includes("834956")) {
          console.log(`🔍 PASS L5 - ADE record: num="${a.num}", piva="${a.piva}", tot=${a.tot}`);
        }
        
        for (const g of gestStillUnmatched) {
          if (matchedGest.has(g.idx)) continue;

          // Condizioni per L5 - PRIORITÀ ALTA: Stesso numero fattura + P.IVA
          // 1. Stessa P.IVA (normalizzata)
          const aPivaNorm = normalizePIVA(a.piva || "");
          const gPivaNorm = normalizePIVA(g.piva || "");
          const samePiva = aPivaNorm === gPivaNorm && aPivaNorm.length >= 11;
          
          // DEBUG: Log per CA AUTO BANK
          if (aPivaNorm.includes("834956") && gPivaNorm.includes("834956")) {
            console.log(`🔍 PASS L5 - PIVA match check: ADE="${aPivaNorm}" vs GEST="${gPivaNorm}" -> samePiva=${samePiva}`);
          }
          
          if (!samePiva) continue;

          // 2. Stesso numero fattura (normalizzato per confronto)
          const aNumNorm = normalizeInvoiceForMatch(a.num || "");
          const gNumNorm = normalizeInvoiceForMatch(g.num || "");
          const sameInvoiceNum = aNumNorm === gNumNorm && aNumNorm.length >= 3;
          
          // DEBUG: Log dettagliato per CA AUTO BANK
          if (aPivaNorm.includes("834956")) {
            console.log(`🔍 CA AUTO BANK check: ADE raw="${a.num}" norm="${aNumNorm}" vs GEST raw="${g.num}" norm="${gNumNorm}" -> sameNum=${sameInvoiceNum}, lenA=${aNumNorm.length}, lenG=${gNumNorm.length}`);
          }
          
          if (sameInvoiceNum) {
            // MATCH FORTE: Stesso numero + P.IVA (anche se date/importi diversi)
            const diffTot = Math.abs(a.tot - g.tot);
            
            console.log(`✅ PASS L5 MATCH: ${a.num} (P.IVA ${aPivaNorm}), diff €${diffTot.toFixed(2)}`);
            l5MatchCount++;
            
            const record = addMatch(
              a,
              g,
              "MATCH_FIX",
              `L5_NUM+PIVA_RILASSATO (stesso numero fattura e P.IVA, importi/date diversi €${diffTot.toFixed(2)})`
            );
            
            // Forza il flag di revisione
            if (record) {
              record.FLAG_REVISIONE = "DA_REVISIONARE";
              
              // Nota dettagliata
              const dateDiff = a.data && g.data ? Math.abs(Math.round((a.data - g.data) / (1000 * 60 * 60 * 24))) : 999;
              record.NOTE_MATCH = `⚠️ STESSO NUMERO FATTURA ma dati diversi: ` +
                                `Num: ${a.num || ''}, P.IVA: ${aPivaNorm}. ` +
                                `Differenza importo: €${(a.tot - g.tot).toFixed(2)} (ADE: €${a.tot.toFixed(2)}, GEST: €${g.tot.toFixed(2)}). ` +
                                `Differenza date: ${dateDiff} giorni. VERIFICARE DATI.`;
            }
            
            continue; // Passa al prossimo record ADE
          }

          // Condizioni per L5 - PRIORITÀ BASSA: Solo P.IVA + data vicina + importo simile
          // 3. Date entro tolleranza (±3 giorni)
          const dateClose = datesClose(a.data, g.data, MATCH_CONFIG.DATE_TOLERANCE_DAYS);
          
          if (!dateClose) continue;

          // 4. Importi diversi ma entro tolleranza L5 (±50€)
          const diffTot = Math.abs(a.tot - g.tot);
          
          if (diffTot <= MATCH_CONFIG.TOLERANCE_L5_AMOUNT) {
            // Match trovato! Lo aggiungiamo con flag di revisione
            const record = addMatch(
              a,
              g,
              "MATCH_FIX",
              `L5_PIVA+DATA (stessa P.IVA, data vicina, importi diversi €${diffTot.toFixed(2)})`
            );
            
            // Forza il flag di revisione per questo tipo di match
            if (record) {
              record.FLAG_REVISIONE = "DA_REVISIONARE";
              
              // Aggiorna la nota
              record.NOTE_MATCH = `⚠️ VERIFICA IMPORTI: Stessa P.IVA (${aPivaNorm}) e data simile, ` +
                                `ma differenza importo di €${(a.tot - g.tot).toFixed(2)}. ` +
                                `ADE: €${a.tot.toFixed(2)}, GEST: €${g.tot.toFixed(2)}. Controllare registrazione.`;
            }
          }
        }
      }
      
      console.log(`✅ PASS L5 completato: ${l5MatchCount} match trovati`);
    }

    // ============================================================
    // 🔧 PASS LEGACY (mantieni per retro compatibilità regole apprese)
    // ============================================================

    // ==========================================
    // 🧠 PASS 1: MATCH PERFETTO (Num strict + Importo)
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const aNumStrict = normalizeInvoiceNumber(a.num, true);
      if (!aNumStrict) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        const gNumStrict = normalizeInvoiceNumber(g.num, true);
        if (!gNumStrict) return false;

        // prima confronto strict
        if (aNumStrict !== gNumStrict) {
          // poi confronto solo cifre
          const aDigits = normalizeInvoiceNumber(a.num, false);
          const gDigits = normalizeInvoiceNumber(g.num, false);
          if (!aDigits || !gDigits || aDigits !== gDigits) return false;
        }

        return totEqual(a.tot, g.tot, 0.05);
      });

      if (matchGest) {
        addMatch(a, matchGest, "OK", "1_MATCH_ESATTO_STRICT");
        continue;
      }
    }

    // ==========================================
    // 🧠 PASS 1B: REGOLE IMPARATE (TIPO_TRASFORMAZIONE)
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const adePivaNorm = normalizePIVA(a.piva || "");
      if (!adePivaNorm) continue;

      const transformRules = (learningData.rules || []).filter(r =>
        r.tipo === "TIPO_TRASFORMAZIONE" &&
        normalizePIVA(r.piva || "") === adePivaNorm
      );
      if (!transformRules.length) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        const gStrict = normalizeInvoiceNumber(g.num, true);
        const gDigits = normalizeInvoiceNumber(g.num, false);
        if (!gStrict && !gDigits) return false;

        for (const rule of transformRules) {
          const transformed = applyTransformRule(rule.rule, a.num);

          const tStrict = normalizeInvoiceNumber(transformed, true);
          const tDigits = normalizeInvoiceNumber(transformed, false);

          let numOk = false;

          if (rule.rule === "GEST_CONTAINS_ADE") {
            if (tStrict && gStrict && gStrict.includes(tStrict)) {
              numOk = true;
            } else if (tDigits && gDigits && gDigits.includes(tDigits)) {
              numOk = true;
            }
          } else {
            if (tStrict && gStrict && tStrict === gStrict) {
              numOk = true;
            } else if (tDigits && gDigits && tDigits === gDigits) {
              numOk = true;
            }
          }

          if (!numOk) continue;

          const diffTot = Math.abs(a.tot - g.tot);
          const imponibileOk = amountsMatch(a.imp, g.imp);
          const bolloOk = Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING;
          const totOk = diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING;

          if (totOk || bolloOk || imponibileOk) {
            return true;
          }
        }

        return false;
      });

      if (matchGest) {
        const diffTot = Math.abs(a.tot - matchGest.tot);
        const isBollo = Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING;

        const adeDigits  = normalizeInvoiceNumber(a.num, false);
        const gestDigits = normalizeInvoiceNumber(matchGest.num, false);
        const sameDigits = adeDigits && gestDigits && adeDigits === gestDigits;

        const isStrongSafe = (diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING && sameDigits);

        let status;
        let note;

        if (isBollo) {
          status = "MATCH_OK";
          note = "Regola appresa + diff. bollo (2€)";
        } else if (isStrongSafe) {
          status = "MATCH_OK";
          note = "Regola appresa: Num (solo cifre) + Importo coincidono";
        } else {
          status = "MATCH_FIX";
          note = "Regola appresa su numero fattura";
        }

        addMatch(a, matchGest, status, `1B_LEARNING_${note}`);
      }
    }

    // ==========================================
    // 🧠 PASS 2: NUMERI COMPATIBILI (prefisso/suffisso) + IMPORTO
    //    Gestisce casi: V1-7309 ↔ 7309, 76/VF01 ↔ 76,
    //    8Z00166851 ↔ 00166851, 480/00/2025 ↔ 480/00,
    //    48000000000029733/V3 ↔ 0000000029733, 1/48 ↔ 48, ecc.
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const aNum = a.num || "";
      if (!aNum) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        const gNum = g.num || "";
        if (!gNum) return false;

        // relazione forte tra cifre
        if (!invoiceDigitsRelated(aNum, gNum, 2)) return false;

        const diff = Math.abs(a.tot - g.tot);

        if (diff < MATCH_CONFIG.TOLERANCE_ROUNDING) return true;
        if (Math.abs(diff - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING) return true;
        if (amountsMatch(a.imp, g.imp)) return true;

        return false;
      });

      if (matchGest) {
        const diffTot = Math.abs(a.tot - matchGest.tot);
        const sameDigits   = normalizeInvoiceNumber(a.num, false) === normalizeInvoiceNumber(matchGest.num, false);
        const digitsCompat = invoiceDigitsRelated(a.num, matchGest.num, 2);
        const datesOk      = datesClose(a.data, matchGest.data, 3);

        const isBollo      = Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING;
        const isStrongSafe = (diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING && (sameDigits || (digitsCompat && datesOk)));

        let status;
        let note;

        if (isBollo) {
          status = "MATCH_OK";
          note = "Diff. Bollo (2€)";
        } else if (isStrongSafe) {
          status = "MATCH_OK";
          note = "Num compatibile (prefisso/suffisso) + Importo coincidono";
        } else {
          status = "MATCH_FIX";
          note = "Match su numero compatibile / importi (controllare)";
        }

        addMatch(a, matchGest, status, "2_NUM_COMPATIBILE+IMPORTO");
      }
    }

    // ==========================================
    // 🧠 PASS 2C: NUMERI CORRELATI (solo cifre) + IMPORTO
    // Dà per buoni i casi tipo:
    //   ADE: 'V1-7309'   / '480/00/2025' / '01S620252800095557'
    //   GEST: '7309'     / '480/00'      / '620252800095557'
    // purché stesso fornitore + importo allineato
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        // numeri documento con cifre "correlate"
        if (!invoiceDigitsRelated(a.num, g.num, 3)) return false;

        const diffTot = Math.abs(a.tot - g.tot);

        // Totale uguale o differenza tipo bollo
        if (diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING) return true;
        if (Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING) return true;

        return false;
      });

      if (matchGest) {
        const diffTot = Math.abs(a.tot - matchGest.tot);
        const isBollo = Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING;

        let status, note;
        if (diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING) {
          status = "MATCH_OK";
          note   = "Num correlati (solo cifre) + Totale coincidente";
        } else if (isBollo) {
          status = "MATCH_OK";
          note   = "Num correlati (solo cifre) + diff. 2€ (bollo)";
        } else {
          status = "MATCH_FIX";
          note   = "Num correlati (solo cifre) + importo da verificare";
        }

        addMatch(a, matchGest, status, `2C_NUM_CORRELATO_${note}`);
      }
    }

        // ==========================================
    // 🧠 PASS 2D: DENOMINAZIONE + IMPORTO (+ DATA)
    //    Usa SOLO la denominazione (anche se P.IVA sporca)
    //    e l'importo. Serve per gestionali che usano numeri
    //    interni (es. "918") scollegati dal numero fattura.
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;

        // denominazioni simili
        if (!namesSimilar(a.den, g.den)) return false;

        // importo molto vicino (totale o imponibile)
        const diffTot = Math.abs(a.tot - g.tot);
        const totOk =
          diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING || Math.abs(diffTot - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING;
        const impOk = amountsMatch(a.imp, g.imp);

        if (!totOk && !impOk) return false;

        // date non troppo lontane (più larghi: ±5 giorni)
        if (!datesClose(a.data, g.data, 5)) return false;

        return true;
      });

      if (matchGest) {
        addMatch(
          a,
          matchGest,
          "MATCH_FIX",
          "2D_DENOMINAZIONE+IMPORTO (P.IVA / num. poco affidabili)"
        );
      }
    }

    // ==========================================
    // 🧠 PASS 2D: NOTE DI CREDITO con segno invertito
    //    Esempi tipo:
    //    ADE: -968,44  | GEST: +968,44
    //    Stesso fornitore, data vicina, numero correlato
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      // ci interessano solo righe ADE con totale negativo
      if (a.tot == null || a.tot >= 0) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        if (g.tot == null) return false;

        // totale di segno opposto ma modulo molto simile
        const sumAbs = Math.abs((a.tot || 0) + (g.tot || 0));
        if (sumAbs > 0.05) return false; // es: -968,44 + 968,44 ≈ 0

        // date vicine (fino a 5 giorni di differenza)
        if (!datesClose(a.data, g.data, 5)) return false;

        // numeri documento "correlati" sulle cifre
        if (!invoiceDigitsRelated(a.num, g.num, 3)) return false;

        return true;
      });

      if (matchGest) {
        addMatch(a, matchGest, "MATCH_FIX", "2D_NC_segno_invertito (controllare)");
      }
    }

    // ==========================================
    // 🧠 PASS 3: DATA + IMPORTO (fallback) – ignora numero
    //    SOLO SE i numeri NON risultano compatibili
    // ==========================================
    for (const a of ade) {
      if (matchedAde.has(a.idx)) continue;

      const matchGest = gest.find(g => {
        if (matchedGest.has(g.idx)) return false;
        if (!sameSupplier(a, g)) return false;

        // se i numeri sono compatibili, avrebbero dovuto essere presi dal PASS 2
        if (invoiceDigitsRelated(a.num, g.num, 2)) return false;

        // Date vicine (±2 gg)
        if (!datesClose(a.data, g.data, 2)) return false;

        const diff = Math.abs(a.tot - g.tot);
        return (diff < MATCH_CONFIG.TOLERANCE_ROUNDING || Math.abs(diff - MATCH_CONFIG.TOLERANCE_STAMP_DUTY) < MATCH_CONFIG.TOLERANCE_ROUNDING);
      });

      if (matchGest) {
        addMatch(a, matchGest, "MATCH_FIX", "3_DATA_IMPORTO (Num ignorato / illeggibile)");
      }
    }

    // ==========================================
    // 🧠 PASS 4: MATCH PER PUNTEGGIO (scansione residui)
    //    Prova a recuperare i casi rimasti SOLO_ADE / SOLO_GEST
    //    usando un punteggio basato su: fornitore, importi,
    //    date e numeri documento.
    // ==========================================
        function computeScoreFallback(a, g) {
      // se non sono lo stesso fornitore in base alle nostre regole -> 0
      if (!sameSupplier(a, g)) return 0;

      let s = 0;

      // 1) Totale documento
      const diffTot = Math.abs((a.tot || 0) - (g.tot || 0));

      if (diffTot < MATCH_CONFIG.TOLERANCE_ROUNDING) {
        // praticamente uguale
        s += 55;
      } else if (diffTot < MATCH_CONFIG.TOLERANCE_L4_AMOUNT) {
        // differenza piccola
        s += 35;
      } else if (diffTot < 2.50) {
        // possibile bollo / arrotondamenti
        s += 15;
      } else if (
        diffTot <= 5.00 &&              // differenza massima 5 €
        a.tot != null && a.tot >= 1000  // fatture medio/grandi
      ) {
        // caso come SHELL: stesso fornitore, stessa data, stesso numero,
        // ma differenza fissa di qualche euro su importi alti
        s += 30;
      }

      // 2) Imponibile
      const diffImp = Math.abs((a.imp || 0) - (g.imp || 0));
      if (diffImp < MATCH_CONFIG.TOLERANCE_ROUNDING) {
        s += 25;
      } else if (diffImp < MATCH_CONFIG.TOLERANCE_L4_AMOUNT) {
        s += 15;
      }

      // 3) Date vicine
      if (datesClose(a.data, g.data, 2)) {
        s += 20;
      } else if (datesClose(a.data, g.data, 5)) {
        s += 10;
      }

      // 4) Numeri documento “correlati” sulle cifre
      if (invoiceDigitsRelated(a.num, g.num, 3)) {
        s += 20;
      }

      // 5) Denominazioni simili
      if (namesSimilar(a.den, g.den)) {
        s += 15;
      }

      return s;
    }

    // prendiamo solo le righe ancora non matchate dopo i primi pass
    const adeUnmatched = ade.filter(a => !matchedAde.has(a.idx));
    const gestUnmatched = gest.filter(g => !matchedGest.has(g.idx));

    for (const a of adeUnmatched) {
      let bestGest = null;
      let bestScore = 0;

      for (const g of gestUnmatched) {
        if (matchedGest.has(g.idx)) continue;

        const score = computeScoreFallback(a, g);
        if (score > bestScore) {
          bestScore = score;
          bestGest = g;
        }
      }

      // soglia di sicurezza: solo match con punteggio alto
      if (bestGest && bestScore >= 70) {
        addMatch(
          a,
          bestGest,
          "MATCH_FIX",
          `4_SCORE_MATCH (punteggio=${bestScore})`
        );
      }
    }

    // ===============================
    // 🚀 STEP ULTRA – MATCH_FIX per chiavi forti (NUM+DATA+TOTALE)
    // ===============================
    {
      const adeStill = ade.filter(a => !matchedAde.has(a.idx));
      const gestStill = gest.filter(g => !matchedGest.has(g.idx));

      for (const a of adeStill) {
        if (matchedAde.has(a.idx)) continue;

        const candidate = gestStill.find(g => !matchedGest.has(g.idx) && strongKeyMatch(a, g));
        if (candidate) {
          matchedAde.add(a.idx);
          matchedGest.add(candidate.idx);

          const diffTot = (a.tot || 0) - (candidate.tot || 0);
          const note = "⚠ Match per chiavi forti (NUM+DATA+TOTALE). Fornitore non agganciato automaticamente (P.IVA mancante o denominazione diversa). Verifica consigliata.";

          results.push({
            ADE: a,
            GEST: candidate,
            STATUS: "MATCH_FIX",
            CRITERIO: "NUM_DATA_TOT",
            NOTE_MATCH: note,
            ORIGINE: "MATCH",
            DIFF_TOTALE: diffTot,
            FLAG_REVISIONE: "SI"
          });
        }
      }
    }

    // 🧠 PASS 5: POPOLA NOTE DIAGNOSTICHE per i MATCH_FIX
    // Aggiunge una spiegazione automatica nel campo note per i match incerti
    for (const res of results) {
      if (res.STATUS === "MATCH_FIX" && res.ADE && res.GEST) {
        // Se non c'è già una nota, genera quella automatica
        if (!res.NOTE_MATCH) {
          const report = generateDiagnosticReport(res.ADE, res.GEST, res.CRITERIO);
          // Prendiamo solo la prima riga del report per brevità nell'export
          res.NOTE_MATCH = report.split('\n')[0];  // ✅ Usa NOTE_MATCH, non matchNote
        }
      }
    }

    // ==========================================
    // 🧠 GESTIONE RIGHE NON MATCHATE (SOLO_ADE / SOLO_GEST)
    // Aggiungi NOTE_MATCH descrittive per i casi non associabili
    // ==========================================
    for (const a of ade) {
      if (!matchedAde.has(a.idx)) {
        // Riga presente SOLO in ADE
        const origine = "SOLO_ADE";
        const note = "Presente solo in ADE (manca nel gestionale).";
        
        results.push({ 
          ADE: a, 
          GEST: null, 
          STATUS: "SOLO_ADE",  // 🔴 STATUS rimane SOLO_ADE per compatibilità con filtri
          CRITERIO: "NO_MATCH",
          NOTE_MATCH: note,
          ORIGINE: origine,
          DIFF_TOTALE: 0  // Non applicabile per righe non matchate
        });
      }
    }
    for (const g of gest) {
      if (!matchedGest.has(g.idx)) {
        // Riga presente SOLO nel Gestionale
        const origine = "SOLO_GESTIONALE";
        const note = "Presente solo nel gestionale (manca in ADE).";
        
        results.push({ 
          ADE: null, 
          GEST: g, 
          STATUS: "SOLO_GEST",  // 🔴 STATUS rimane SOLO_GEST per compatibilità con filtri
          CRITERIO: "NO_MATCH",
          NOTE_MATCH: note,
          ORIGINE: origine,
          DIFF_TOTALE: 0  // Non applicabile per righe non matchate
        });
      }
    }

    // ============================================================
    // 🚀 STEP ULTRA (FIX DEFINITIVO): unisci SOLO_ADE + SOLO_GEST
    // Chiavi forti calcolate al volo: NUM(normalizzato) + DATA(normalizzata) + TOTALE
    // (non ci fidiamo di numDigits/dataIso se arrivano sporchi)
    // ============================================================
    (function ultraMergeSolo() {
      const TOL = MATCH_CONFIG.TOL_TOT ?? MATCH_CONFIG.TOLERANCE_ROUNDING ?? 0.05;

      const isActive = (r) => r && !r.deleted;

      // 🔑 chiave forte (sempre ricalcolata)
      const keyOf = (rec) => {
        const numKey = normalizeInvoiceNumber(rec?.num || "", false);     // solo cifre
        const dateKey = normalizeDateFlexible(rec?.dataRaw || rec?.dataStr || rec?.dataIso || ""); // YYYY-MM-DD
        return `${numKey}|${dateKey}`;
      };

      const totVal = (rec) => Number(rec?.tot ?? 0);

      const soloAde = results.filter(r => isActive(r) && r.STATUS === "SOLO_ADE" && r.ADE);
      const soloGest = results.filter(r => isActive(r) && r.STATUS === "SOLO_GEST" && r.GEST);

      // indicizza SOLO_GEST per chiave
      const map = new Map();
      for (const gr of soloGest) {
        const k = keyOf(gr.GEST);
        if (!map.has(k)) map.set(k, []);
        map.get(k).push(gr);
      }

      let mergedCount = 0;

      for (const ar of soloAde) {
        const a = ar.ADE;
        const k = keyOf(a);
        const candidates = map.get(k);
        if (!candidates || !candidates.length) continue;

        const aTot = totVal(a);
        let pickIndex = -1;

        for (let i = 0; i < candidates.length; i++) {
          const gTot = totVal(candidates[i].GEST);
          if (Math.abs(aTot - gTot) <= TOL) {
            pickIndex = i;
            break;
          }
        }
        if (pickIndex < 0) continue;

        const gr = candidates.splice(pickIndex, 1)[0];
        const g = gr.GEST;

        // ✅ unisco nella riga ADE
        ar.GEST = g;
        ar.STATUS = "MATCH_FIX";
        ar.ORIGINE = "MATCH";
        ar.CRITERIO = "NUM_DATA_TOT";
        ar.NOTE_MATCH =
          "⚠ Match ULTRA (NUM+DATA+TOTALE). Fornitore/P.IVA non coerenti o mancanti nel gestionale. Verifica consigliata.";
        ar.FLAG_REVISIONE = "DA_REVISIONARE";
        ar.DIFF_TOTALE = (aTot - totVal(g));

        // ❌ nascondo la riga SOLO_GEST duplicata
        gr.deleted = true;
        gr.STATUS = "MERGED_ULTRA";

        mergedCount++;
      }

      console.log(`🚀 ULTRA MERGE completato: ${mergedCount} righe unite`);
    })();

    // ============================================================
    // NORMALIZZAZIONE FINALE: Coerenza STATUS / ORIGINE
    // ============================================================
    for (const r of results) {
      // Assicurati che NOTE_MATCH esista sempre
      if (!r.NOTE_MATCH) {
        r.NOTE_MATCH = "";
      }
      
      // Assicurati che DIFF_TOTALE esista sempre e sia valido
      if (r.DIFF_TOTALE === undefined || r.DIFF_TOTALE === null || isNaN(r.DIFF_TOTALE)) {
        r.DIFF_TOTALE = 0;
      }

      if (r.FLAG_REVISIONE === undefined || r.FLAG_REVISIONE === null) {
        r.FLAG_REVISIONE = "";
      }
      
      // 🎯 CONTROLLO COERENZA STATUS / ORIGINE
      // Garantisce che ORIGINE rispecchi sempre STATUS
      if (r.STATUS === "MATCH_OK" || r.STATUS === "MATCH_FIX") {
        r.ORIGINE = "MATCH";
      } else if (r.STATUS === "SOLO_GEST") {
        r.ORIGINE = "SOLO_GESTIONALE";
      } else if (r.STATUS === "SOLO_ADE") {
        r.ORIGINE = "SOLO_ADE";
      } else if (r.STATUS === "MERGED_ULTRA") {
        r.ORIGINE = "MERGED_ULTRA";
      } else if (!r.ORIGINE) {
        // Caso anomalo: STATUS non riconosciuto
        console.warn(`⚠️ Record con STATUS sconosciuto: "${r.STATUS}" - possibile bug nel matching`, r);
        r.ORIGINE = "UNKNOWN";
      }
    }

    return results;
  }

  // ============================================================
  // 🔍 POST-PROCESSING: FLAGGING CASI SOSPETTI (Dic 2025)
  // ============================================================
  /**
   * Identifica coppie SOLO_ADE / SOLO_GEST con stesso numero e P.IVA
   * ma importi/date diverse. Aggiunge FLAG_REVISIONE per segnalare
   * possibili abbinamenti manuali da verificare.
   * 
   * NON modifica STATUS, ORIGINE o CRITERIO.
   * NON crea match automatici.
   * 
   * @param {Array} rows - Array completo dei risultati da matchRecords
   * @returns {Array} - Stesso array con FLAG_REVISIONE e NOTE_MATCH aggiornati
   */
  function flagSuspiciousSoloAdeSoloGest(rows) {
    // Configurazione soglie
    const AMOUNT_TOLERANCE = 0.50; // Differenza importi oltre 50¢ = sospetto
    const DATE_DIFF_DAYS = 0;      // Date diverse = sospetto (anche 1 giorno)
    
    // 1. Filtra solo SOLO_ADE e SOLO_GEST
    const soloRows = rows.filter(r => 
      r.STATUS === "SOLO_ADE" || r.STATUS === "SOLO_GEST"
    );
    
    // 2. Raggruppa per chiave (PIVA normalizzata + Numero normalizzato)
    const groups = new Map();
    
    for (const row of soloRows) {
      // Estrai dati da ADE o GEST
      const record = row.ADE || row.GEST;
      if (!record) continue;
      
      // Normalizza P.IVA e numero
      const pivaNorm = normalizePIVA(record.piva || "");
      const numNorm = normalizeInvoiceNumber(record.num, false); // Solo cifre
      
      if (!pivaNorm || !numNorm) continue;
      
      const key = `${pivaNorm}|${numNorm}`;
      
      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key).push(row);
    }
    
    // 3. Identifica gruppi sospetti
    const suspiciousGroups = [];
    
    for (const [key, groupRows] of groups.entries()) {
      const adeRows = groupRows.filter(r => r.STATUS === "SOLO_ADE");
      const gestRows = groupRows.filter(r => r.STATUS === "SOLO_GEST");
      
      // Serve almeno 1 ADE e 1 GEST
      if (adeRows.length === 0 || gestRows.length === 0) continue;
      
      // Confronta importi e date
      let isSuspicious = false;
      let reason = [];
      
      // Prendi il primo di ogni tipo per confronto
      const adeRecord = adeRows[0].ADE;
      const gestRecord = gestRows[0].GEST;
      
      // Confronta totali
      const adeTot = adeRecord?.tot || 0;
      const gestTot = gestRecord?.tot || 0;
      const diffAmount = Math.abs(adeTot - gestTot);
      
      if (diffAmount > AMOUNT_TOLERANCE) {
        isSuspicious = true;
        reason.push(`importi diversi (ADE: €${adeTot.toFixed(2)}, GEST: €${gestTot.toFixed(2)}, diff: €${diffAmount.toFixed(2)})`);
      }
      
      // Confronta date
      const adeDate = adeRecord?.data;
      const gestDate = gestRecord?.data;
      
      if (adeDate && gestDate) {
        const adeDateStr = formatDateIT(adeDate);
        const gestDateStr = formatDateIT(gestDate);
        
        if (adeDateStr !== gestDateStr) {
          const diffMs = Math.abs(adeDate.getTime() - gestDate.getTime());
          const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
          
          if (diffDays > DATE_DIFF_DAYS) {
            isSuspicious = true;
            reason.push(`date diverse (ADE: ${adeDateStr}, GEST: ${gestDateStr}, diff: ${diffDays} gg)`);
          }
        }
      }
      
      if (isSuspicious) {
        suspiciousGroups.push({
          key,
          adeRows,
          gestRows,
          reason: reason.join("; ")
        });
      }
    }
    
    // 4. Marca le righe sospette
    for (const group of suspiciousGroups) {
      // Marca righe ADE
      for (const row of group.adeRows) {
        row.FLAG_REVISIONE = "DA_REVISIONARE";
        const currentNote = row.NOTE_MATCH || "";
        const newNote = "⚠️ Possibile abbinamento con riga gestionale ma " + group.reason + ". Verificare manualmente.";
        row.NOTE_MATCH = currentNote ? `${currentNote} | ${newNote}` : newNote;
      }
      
      // Marca righe GEST
      for (const row of group.gestRows) {
        row.FLAG_REVISIONE = "DA_REVISIONARE";
        const currentNote = row.NOTE_MATCH || "";
        const newNote = "⚠️ Possibile abbinamento con fattura ADE ma " + group.reason + ". Verificare manualmente.";
        row.NOTE_MATCH = currentNote ? `${currentNote} | ${newNote}` : newNote;
      }
    }
    
    // 5. Aggiungi FLAG_REVISIONE = "" a tutte le righe non marcate
    for (const row of rows) {
      if (!row.FLAG_REVISIONE) {
        row.FLAG_REVISIONE = "";
      }
    }
    
    console.log(`🔍 Post-processing: trovati ${suspiciousGroups.length} gruppi sospetti SOLO_ADE/SOLO_GEST da revisionare`);
    
    return rows;
  }

  // ============================================================
  // ANALISI "SOLO GESTIONALE" (Miglioramento Dic 2025)
  // ============================================================
  // Suddivide le fatture "Solo Gestionale" in:
  // - Fisiologiche: successive alla data massima ADE (disallineamento periodo)
  // - Anomalie: all'interno del periodo coperto da ADE (possibili errori)
  // ============================================================
  function analyzeSoloGest(results) {
    const soloGest = results.filter(r => r.STATUS === "SOLO_GEST" && !r.deleted && r.GEST);
    
    if (soloGest.length === 0) {
      return {
        total: 0,
        fisiologiche: 0,
        anomalie: 0,
        maxDateADE: null,
        minDateADE: null,
        detailsFisiologiche: [],
        detailsAnomalie: []
      };
    }

    // Trova range date ADE
    const adeRecords = results
      .filter(r => r.ADE && !r.deleted)
      .map(r => r.ADE.data)
      .filter(d => d);

    let maxDateADE = null;
    let minDateADE = null;

    if (adeRecords.length > 0) {
      maxDateADE = new Date(Math.max(...adeRecords.map(d => d.getTime())));
      minDateADE = new Date(Math.min(...adeRecords.map(d => d.getTime())));
    }

    // Classifica Solo Gestionale
    const fisiologiche = [];
    const anomalie = [];

    for (const r of soloGest) {
      const gestDate = r.GEST.data;
      
      if (!gestDate || !maxDateADE) {
        // Se non abbiamo date, consideriamo anomalia
        anomalie.push(r);
      } else if (gestDate > maxDateADE) {
        // Data successiva al massimo ADE = fisiologica
        fisiologiche.push(r);
      } else if (gestDate < minDateADE) {
        // Data precedente al minimo ADE = fisiologica
        fisiologiche.push(r);
      } else {
        // Data all'interno del periodo ADE = anomalia
        anomalie.push(r);
      }
    }

    return {
      total: soloGest.length,
      fisiologiche: fisiologiche.length,
      anomalie: anomalie.length,
      maxDateADE,
      minDateADE,
      detailsFisiologiche: fisiologiche,
      detailsAnomalie: anomalie
    };
  }

  // --- RINOMINA MASSIVA FORNITORE E RICALCOLO ---
  function bulkRenameAndRematch(badName, goodName) {
    if (!badName || !goodName) return;

    // 1. Chiedo conferma per sicurezza
    if (!confirm(`Sei sicuro? \n\nTutte le righe del Gestionale con nome:\n"${badName}"\n\nverranno rinominate in:\n"${goodName}"\n\nE il confronto verrà ricalcolato da zero.`)) {
      return;
    }

    // 2. Aggiorno i dati grezzi in memoria (lastGestRecords)
    let count = 0;
    for (const r of lastGestRecords) {
      // Controllo se il nome è quello "sbagliato" (normalizzato o esatto)
      if (r.den === badName || normalizeName(r.den) === normalizeName(badName)) {
        r.den = goodName; // Sostituisco col nome ADE
        r.denNorm = normalizeName(goodName); // Aggiorno anche la versione normalizzata
        count++;
      }
    }

    if (count === 0) {
      alert("Nessuna riga trovata con quel nome nel gestionale.");
      return;
    }

    // 3. Ricalcolo tutto il Match!
    // Nota: Dobbiamo rifare il matching completo perché ora i nomi coincidono
    lastResults = matchRecords(lastAdeRecords, lastGestRecords);
    
    // 🔍 Post-processing: flagga casi sospetti SOLO_ADE/SOLO_GEST
    lastResults = flagSuspiciousSoloAdeSoloGest(lastResults);
    
    // Ripristino ID e pulizia
    nextResultId = 1;
    lastResults.forEach(r => {
      r.id = nextResultId++;
      r.rowId = r.id;
      r.deleted = false;
      r.matchNote = r.matchNote || "";
    });

    // 4. Aggiorno la vista
    currentEditRowId = null; // Chiudo l'editor se aperto
    document.getElementById("editPanelBackdrop").classList.add("hidden");
    
    applyFilterAndRender(); // Ridisegno tabella e KPI
    
    alert(`Fatto! ${count} righe aggiornate e riconciliazione ricalcolata.`);
  }

  // --- GESTIONE MATCH MANUALE (INIZIO) ---

  function startManualMatch(rowId, sourceSide) {
    pendingMatchSourceId = rowId;
    pendingMatchSourceSide = sourceSide;
    const sideName = sourceSide === "ADE" ? "GESTIONALE" : "ADE";
    alert(`Hai selezionato una riga ${sourceSide}.\nOra clicca sulla riga ${sideName} con cui vuoi unirla.`);
    refreshResultsView(); // Ridisegna per mostrare il giallo
  }

  function completeManualMatch(targetRowId) {
    const sourceRow = lastResults.find(r => r.id === pendingMatchSourceId);
    const targetRow = lastResults.find(r => r.id === targetRowId);

    if (!sourceRow || !targetRow) {
      alert("Errore: righe non trovate.");
      resetManualMatchState();
      return;
    }

    // Evita di abbinare la riga con se stessa
    if (sourceRow.id === targetRow.id) return;

    // Verifiche di coerenza (ADE con GEST)
    const sourceIsAde = pendingMatchSourceSide === "ADE";
    const targetIsGest = targetRow.STATUS === "SOLO_GEST";
    const targetIsAde = targetRow.STATUS === "SOLO_ADE";

    if (sourceIsAde && !targetIsGest) {
      alert("Devi selezionare una riga SOLO GESTIONALE.");
      return;
    }
    if (!sourceIsAde && !targetIsAde) {
      alert("Devi selezionare una riga SOLO ADE.");
      return;
    }

    // Creazione match manuale
    undoStack.push(JSON.parse(JSON.stringify(lastResults)));
    btnUndoDelete.disabled = false;

    const newRow = {
      id: nextResultId++,
      rowId: nextResultId,
      ADE: sourceIsAde ? sourceRow.ADE : targetRow.ADE,
      GEST: sourceIsAde ? targetRow.GEST : sourceRow.GEST,
      STATUS: "MANUAL_MATCH",
      CRITERIO: "Abbinamento Manuale Utente",
      matchNote: "Unione manuale",
      deleted: false
    };

    // Nascondo le vecchie righe e aggiungo la nuova
    sourceRow.deleted = true;
    targetRow.deleted = true;
    lastResults.unshift(newRow);

    resetManualMatchState();
    alert("Match completato con successo!");
  }

  function resetManualMatchState() {
    pendingMatchSourceId = null;
    pendingMatchSourceSide = null;
    refreshResultsView();
  }
  // --- GESTIONE MATCH MANUALE (FINE) ---

  // ---------- rendering riconciliazione ----------

  function renderResults(resultsToRender, fullResultsForCounts) {
    const tbody = document.querySelector("#resultTable tbody");
    tbody.innerHTML = "";

    const full = fullResultsForCounts || resultsToRender;

    let cntOk = 0, cntFix = 0, cntAde = 0, cntGest = 0;
    for (const row of full) {
      if (row.deleted) continue;
      if (row.STATUS === "MATCH_OK") cntOk++;
      else if (row.STATUS === "MATCH_FIX") cntFix++;
      else if (row.STATUS === "SOLO_ADE") cntAde++;
      else if (row.STATUS === "SOLO_GEST") cntGest++;
    }

    for (const row of resultsToRender) {
      if (row.deleted) continue;

      const tr = document.createElement("tr");
      let cls = "", pillClass = "", pillText = "";

      if (row.STATUS === "MATCH_OK") {
        cls = "status-match-ok";
        pillClass = "pill-ok";
        pillText = "MATCH OK";
      } else if (row.STATUS === "MATCH_FIX") {
        cls = "status-match-fix";
        pillClass = "pill-fix";
        pillText = "MATCH DA CORREGGERE";
      } else if (row.STATUS === "SOLO_ADE") {
        cls = "status-solo-ade";
        pillClass = "pill-ade";
        pillText = "SOLO ADE";
      } else if (row.STATUS === "SOLO_GEST") {
        cls = "status-solo-gest";
        pillClass = "pill-gest";
        pillText = "SOLO GESTIONALE";
      } else if (row.STATUS === "MANUAL_MATCH") {
        cls = "status-manual-match";
        pillClass = "pill-manual";
        pillText = "MATCH MANUALE";
      }

      tr.className = cls;
      if (row.id === pendingMatchSourceId) {
        tr.classList.add("match-source-row");
      }
      if (typeof row.id !== "undefined") {
        tr.dataset.id = String(row.id);
      }

      const a = row.ADE;
      const g = row.GEST;

      function tdText(text, extraClass = "") {
        const td = document.createElement("td");
        if (extraClass) td.className = extraClass;
        td.textContent = text ?? "";
        return td;
      }

          // ADE
tr.appendChild(tdText(a ? a.num : "", "mono"));
tr.appendChild(tdText(a ? a.den : ""));
tr.appendChild(tdText(a ? a.piva : "", "mono"));

// 🌟 Conversione data ADE: usa dataIso normalizzato in formato italiano DD/MM/YYYY
const adeDateDisplay = a
  ? (a.dataIso ? formatDateIT(a.dataIso) : (a.dataStr || ""))
  : "";

tr.appendChild(tdText(adeDateDisplay, "mono"));
tr.appendChild(tdText(a ? a.imp.toFixed(2) : "", "mono"));
tr.appendChild(tdText(a ? a.iva.toFixed(2) : "", "mono"));
tr.appendChild(tdText(a ? a.tot.toFixed(2) : "", "mono"));

// GEST
tr.appendChild(tdText(g ? g.num : "", "mono"));
tr.appendChild(tdText(g ? g.den : ""));

// 🌟 Conversione data Gestionale: usa dataIso normalizzato in formato italiano DD/MM/YYYY
const gestDateDisplay = g
  ? (g.dataIso ? formatDateIT(g.dataIso) : (g.dataStr || ""))
  : "";

tr.appendChild(tdText(gestDateDisplay, "mono"));
tr.appendChild(tdText(g ? g.imp.toFixed(2) : "", "mono"));
tr.appendChild(tdText(g ? g.iva.toFixed(2) : "", "mono"));
tr.appendChild(tdText(g ? g.tot.toFixed(2) : "", "mono"));

      // Stato
      const tdStatus = document.createElement("td");
      const span = document.createElement("span");
      span.className = "status-pill " + pillClass;
      span.textContent = pillText;
      tdStatus.appendChild(span);
      tr.appendChild(tdStatus);

      // Criterio
      tr.appendChild(tdText(row.CRITERIO || "", "muted criterio-cell"));

      // Azioni
      const tdActions = document.createElement("td");
      tdActions.className = "actions-cell";

      if (row.STATUS === "SOLO_ADE") {
        const btn = document.createElement("button");
        btn.className = "btn-link-small";
        btn.textContent = "Abbina a Gest.";
        btn.onclick = (e) => { e.stopPropagation(); startManualMatch(row.id, "ADE"); };
        tdActions.appendChild(btn);
      } else if (row.STATUS === "SOLO_GEST") {
        const btn = document.createElement("button");
        btn.className = "btn-link-small";
        btn.textContent = "Abbina ad ADE";
        btn.onclick = (e) => { e.stopPropagation(); startManualMatch(row.id, "GEST"); };
        tdActions.appendChild(btn);
      }

      // Bottone elimina per righe che hanno una parte gestionale
      if (row.GEST) {
        const btnDel = document.createElement("button");
        btnDel.type = "button";
        btnDel.className = "row-delete-btn";
        btnDel.textContent = "🗑";
        btnDel.title = "Elimina parte Gestionale";
        tdActions.appendChild(btnDel);
      }
      tr.appendChild(tdActions);

      tbody.appendChild(tr);
    }

    document.getElementById("cntOk").textContent = cntOk;
    document.getElementById("cntFix").textContent = cntFix;
    document.getElementById("cntAde").textContent = cntAde;
    document.getElementById("cntGest").textContent = cntGest;

    // ============================================================
    // ANALISI GUIDATA SOLO GESTIONALE (Miglioramento Dic 2025)
    // ============================================================
    // Mostra quanti "Solo Gestionale" sono fisiologici (fuori periodo ADE)
    // e quanti sono anomalie da verificare (dentro periodo ADE)
    // ============================================================
    if (cntGest > 0) {
      const analysis = analyzeSoloGest(full);
      const existingAnalysis = document.getElementById("soloGestAnalysis");
      
      if (existingAnalysis) {
        existingAnalysis.remove();
      }

      const analysisDiv = document.createElement("div");
      analysisDiv.id = "soloGestAnalysis";
      analysisDiv.className = "solo-gest-analysis";
      analysisDiv.style.cssText = "margin: 10px 0; padding: 12px 14px; background: var(--sys-bg-content-alt); border: 1px solid var(--sys-separator); border-radius: 12px;";
      
      let html = `<strong style="color: var(--sys-text-primary);">📊 Analisi "Solo Gestionale" (${cntGest} totali):</strong><br>`;
      
      if (analysis.maxDateADE) {
        const formatDate = (d) => d ? `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}` : "N/D";
        html += `<span style="color: var(--sys-text-secondary); font-size: 0.9rem;">Periodo ADE: dal ${formatDate(analysis.minDateADE)} al ${formatDate(analysis.maxDateADE)}</span><br>`;
        html += `• <span style="color: #28a745; font-weight: 600;">${analysis.fisiologiche} fisiologiche</span> <span style="color: var(--sys-text-secondary); font-size: 0.85rem;">(fuori periodo ADE)</span><br>`;
        html += `• <span style="color: #dc3545; font-weight: 600;">${analysis.anomalie} da verificare</span> <span style="color: var(--sys-text-secondary); font-size: 0.85rem;">(dentro periodo ADE ma non matchate)</span>`;
      } else {
        html += `<em style="color: var(--sys-text-secondary);">Impossibile determinare periodo ADE (nessuna data disponibile)</em>`;
      }
      
      analysisDiv.innerHTML = html;
      
      const summaryRow = document.getElementById("summaryRow");
      if (summaryRow && summaryRow.parentNode) {
        summaryRow.parentNode.insertBefore(analysisDiv, summaryRow.nextSibling);
      }
    } else {
      const existingAnalysis = document.getElementById("soloGestAnalysis");
      if (existingAnalysis) {
        existingAnalysis.remove();
      }
    }

    document.getElementById("summaryRow").style.display = "flex";
    document.getElementById("tableToolbar").style.display = "block";
    document.getElementById("tableContainer").style.display = "block";

        const total = full.filter(r => !r.deleted).length;

    // se per qualche motivo i contatori globali non sono stati valorizzati,
    // uso come fallback la lunghezza degli array in memoria
    const adeCount = totalAdeCount || (lastAdeRecords?.length ?? 0);
    const gestCount = totalGestCount || (lastGestRecords?.length ?? 0);

    document.getElementById("statusText").textContent =
      `Confronto completato: ${total} righe attive ` +
      `(OK: ${cntOk}, Da correggere: ${cntFix}, Solo ADE: ${cntAde}, Solo Gest.: ${cntGest}) ` +
      `| Fatture ADE caricate: ${adeCount} | Righe gestionali caricate: ${gestCount}`;
    
    // Attiva il ridimensionamento sulla tabella appena renderizzata
    const tableEl = document.getElementById("resultTable");
    if (tableEl) createResizableTable(tableEl);
  }

  // ---------- export riconciliazione ----------

  function formatNumberForExcel(value) {
    if (value == null || value === "" || isNaN(value)) return "";
    let num = typeof value === "number" ? value : parseNumberIT(value);
    if (isNaN(num)) return "";
    let s = num.toFixed(2);
    s = s.replace(".", ",");
    return s;
  }

  function exportToCSV(results, fileLabel) {
    if (!results || !results.length) return;

    const active = results.filter(r => !r.deleted && r.STATUS !== "MERGED_ULTRA");

    if (!active.length) return;

    const headers = [
      "ADE_Num","ADE_Den","ADE_PIVA","ADE_Data","ADE_Imponibile","ADE_Iva","ADE_Totale",
      "GEST_Num","GEST_Den","GEST_PIVA","GEST_Data","GEST_Imponibile","GEST_Iva","GEST_Totale",
      "DIFF_TOTALE","STATUS","ORIGINE","CRITERIO","NOTE_MATCH","FLAG_REVISIONE"
    ];
    const rows = [headers.join(";")];

    for (const row of active) {
      const a = row.ADE;
      const g = row.GEST;

      const cleanStr = (s) => String(s ?? "").replace(/"/g, '""').replace(/^'|'$/g, '');

      const lineParts = [];

            // ADE - testuali
      lineParts.push(`"${cleanStr(a?.num)}"`);
      lineParts.push(`"${cleanStr(a?.den)}"`);
      lineParts.push(`"${cleanStr(a?.piva)}"`);

      // 🔹 Data ADE sempre normalizzata in formato italiano (DD/MM/YYYY)
      const adeDataOut = a
        ? (a.dataIso ? formatDateIT(a.dataIso) : (a.dataStr || ""))
        : "";
      lineParts.push(`"${cleanStr(adeDataOut)}"`);

      // ADE - numerici
      lineParts.push(formatNumberForExcel(a?.imp));
      lineParts.push(formatNumberForExcel(a?.iva));
      lineParts.push(formatNumberForExcel(a?.tot));

      // GEST - testuali
      lineParts.push(`"${cleanStr(g?.num)}"`);
      lineParts.push(`"${cleanStr(g?.den)}"`);
      lineParts.push(`"${cleanStr(g?.piva)}"`);

      // 🔹 Data Gestionale anch'essa normalizzata in DD/MM/YYYY
      const gestDataOut = g
        ? (g.dataIso ? formatDateIT(g.dataIso) : (g.dataStr || ""))
        : "";
      lineParts.push(`"${cleanStr(gestDataOut)}"`);

      // GEST - numerici
      lineParts.push(formatNumberForExcel(g?.imp));
      lineParts.push(formatNumberForExcel(g?.iva));
      lineParts.push(formatNumberForExcel(g?.tot));

      // DIFF_TOTALE (ADE_Totale - GEST_Totale)
      lineParts.push(formatNumberForExcel(row.DIFF_TOTALE || 0));

      // Stato, origine, criterio, nota
      const criterioOut = row.CRITERIO || row.criterio || "";
      const flagOut = row.FLAG_REVISIONE || (row.flagRevisione ? "SI" : "");

      lineParts.push(`"${cleanStr(row.STATUS)}"`);
      lineParts.push(`"${cleanStr(row.ORIGINE || "")}"`);
      lineParts.push(`"${cleanStr(criterioOut)}"`);

      const noteOut = [
        (row.NOTE_MATCH || row.matchNote || row.noteMatch || ""),
        (a && a.techNote) ? a.techNote : ""
      ].filter(Boolean).join(" | ");

      lineParts.push(`"${cleanStr(noteOut)}"`);  // Usa NOTE_MATCH + ade.techNote se presente
      lineParts.push(`"${cleanStr(flagOut)}"`);
      rows.push(lineParts.join(";"));
    }

    const csvContent = "\uFEFF" + rows.join("\r\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const ts = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `match_fatture_${fileLabel || "tutte"}_${ts}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ---------- ANALISI FORNITORI (ADE) ----------

  const anStartInput = document.getElementById("anStart");
  const anEndInput = document.getElementById("anEnd");
  const anApplyBtn = document.getElementById("btnAnApply");
  const anAllBtn = document.getElementById("btnAnAll");
  const anYearBtn = document.getElementById("btnAnYear");
  const anLast12Btn = document.getElementById("btnAnLast12");
  const anSearchInput = document.getElementById("anSearch");
  const analysisStatusEl = document.getElementById("analysisStatus");
  const analysisCountEl = document.getElementById("analysisCount");
  const analysisTableContainer = document.getElementById("analysisTableContainer");
  const analysisTbody = document.getElementById("analysisTbody");

  function initAnalysisPeriodFromAde() {
    if (!lastAdeRecords || !lastAdeRecords.length) {
      analysisStatusEl.textContent = "In attesa dati ADE: carica il file nella sezione Riconciliazione.";
      return;
    }
    const { min, max } = getMinMaxDates(lastAdeRecords);
    if (!min || !max) {
      analysisStatusEl.textContent = "Non sono riuscito a leggere le date dal file ADE.";
      return;
    }
    analysisMinDate = min;
    analysisMaxDate = max;
    if (anStartInput && anEndInput) {
      anStartInput.value = dateToInputValue(min);
      anEndInput.value = dateToInputValue(max);
    }
    computeSupplierAnalysis();
  }

  function computeSupplierAnalysis() {
    if (!lastAdeRecords || !lastAdeRecords.length) {
      analysisStatusEl.textContent = "Nessun dato ADE caricato. Usa la sezione Riconciliazione per importare il file.";
      analysisTableContainer.style.display = "none";
      analysisCountEl.textContent = "Fornitori nel periodo: 0";
      analysisTbody.innerHTML = "";
      return;
    }

    let startDate = parseInputDate(anStartInput.value);
    let endDate = parseInputDate(anEndInput.value);

    if (!startDate && analysisMinDate) startDate = analysisMinDate;
    if (!endDate && analysisMaxDate) endDate = analysisMaxDate;

    if (startDate && endDate && endDate < startDate) {
      const tmp = startDate;
      startDate = endDate;
      endDate = tmp;
      if (anStartInput && anEndInput) {
        anStartInput.value = dateToInputValue(startDate);
        anEndInput.value = dateToInputValue(endDate);
      }
    }

    const suppliersMap = new Map();

    for (const r of lastAdeRecords) {
      if (!r.data) continue;
      if (startDate && r.data < startDate) continue;
      if (endDate && r.data > endDate) continue;

      const key = (r.pivaDigits || "") + "|" + (r.denNorm || "");

      if (!suppliersMap.has(key)) {
        suppliersMap.set(key, {
          den: r.den || "",
          denNorm: r.denNorm || "",
          piva: r.piva || "",
          totImp: 0,
          totIva: 0,
          totTot: 0,
          count: 0,
          minDate: null,
          maxDate: null
        });
      }

      const s = suppliersMap.get(key);
      s.totImp += r.imp || 0;
      s.totIva += r.iva || 0;
      s.totTot += r.tot || 0;
      s.count += 1;
      if (!s.minDate || r.data < s.minDate) s.minDate = r.data;
      if (!s.maxDate || r.data > s.maxDate) s.maxDate = r.data;
    }

    const arr = Array.from(suppliersMap.values());
    arr.sort((a, b) => b.totTot - a.totTot);

    lastSupplierAnalysis = arr;
    renderSupplierAnalysis();
  }

  function getFilteredSuppliers() {
    if (!lastSupplierAnalysis) return [];
    const term = (anSearchInput?.value || "").trim().toUpperCase();
    if (!term) return lastSupplierAnalysis;

    return lastSupplierAnalysis.filter(s => {
      const den = (s.denNorm || "").toUpperCase();
      const piva = (s.piva || "").toUpperCase();
      return den.includes(term) || piva.includes(term);
    });
  }

  function renderSupplierAnalysis() {
    const list = getFilteredSuppliers();
    analysisTbody.innerHTML = "";

    if (!list.length) {
      analysisTableContainer.style.display = "none";
      analysisCountEl.textContent = "Fornitori nel periodo: 0";
      analysisStatusEl.textContent = "Nessun fornitore trovato nel periodo selezionato.";
      return;
    }

    for (const s of list) {
      const tr = document.createElement("tr");

      function td(text, extraClass) {
        const td = document.createElement("td");
        if (extraClass) td.className = extraClass;
        td.textContent = text ?? "";
        return td;
      }

      tr.appendChild(td(s.den));
      tr.appendChild(td(s.piva, "mono"));
      tr.appendChild(td(formatNumberITDisplay(s.totImp), "mono"));
      tr.appendChild(td(formatNumberITDisplay(s.totIva), "mono"));
      tr.appendChild(td(formatNumberITDisplay(s.totTot), "mono"));
      tr.appendChild(td(String(s.count), "mono"));

      const periodoStr = (s.minDate && s.maxDate)
        ? `${formatDateIT(s.minDate)} – ${formatDateIT(s.maxDate)}`
        : "n/d";
      tr.appendChild(td(periodoStr, "mono"));

      analysisTbody.appendChild(tr);
    }

    analysisTableContainer.style.display = "block";
    analysisCountEl.textContent = `Fornitori nel periodo: ${list.length}`;
    analysisStatusEl.textContent = "Analisi calcolata sui dati ADE nel periodo selezionato.";
    
    const tableEl = document.getElementById("analysisTable");
    if (tableEl) createResizableTable(tableEl);
  }

  // ---------- event handlers riconciliazione ----------
  const btnMatch = document.getElementById("btnMatch");
  const btnExportAll = document.getElementById("btnExportAll");
  const btnExportAde = document.getElementById("btnExportAde");
  const btnExportGest = document.getElementById("btnExportGest");
  const btnExportFix = document.getElementById("btnExportFix");
  const fileInputA = document.getElementById("fileA");
  const fileInputB = document.getElementById("fileB");
  const fileInputC = document.getElementById("fileC");
  const statusEl = document.getElementById("statusText");
  const filterButtons = Array.from(document.querySelectorAll(".filter-btn"));
  const btnUndoDelete = document.getElementById("btnUndoDelete");
  const btnResetAll = document.getElementById("btnResetAll");
  const monthFilterSelect = document.getElementById("monthFilter");
  const supplierFilterInput = document.getElementById("supplierFilterInput");
  filterInvoiceInput = document.getElementById("filterInvoice");
  const btnClearFilters = document.getElementById("btnClearFilters");
  const advancedFilterRow = document.getElementById("advancedFilterRow");

  const stepperEl = document.getElementById("flowStepper");
  const stepNodes = stepperEl ? Array.from(stepperEl.querySelectorAll(".flow-step")) : [];

  function setStepState(stepNumber, state, descText) {
    const node = stepNodes.find(n => n.dataset.step === String(stepNumber));
    if (!node) return;
    node.classList.remove("disabled", "active", "completed");
    node.classList.add(state);
    const descEl = node.querySelector(".flow-step-desc");
    if (descEl && descText) descEl.textContent = descText;
  }

  function computeBlockingReason() {
    if (!flowState.files.ade) return "Carica il File A (ADE)";
    if (!flowState.files.gest) return "Carica il File B (Gestionale)";
    if (!flowState.mapping.gestBConfirmed) return "Mapping colonne File B obbligatorio";
    if (flowState.files.nc && !flowState.mapping.gestCConfirmed) return "Mapping colonne File C obbligatorio";
    return null;
  }

  // Scroll smooth to element
  function scrollToEl(el) {
    if (!el) return;
    el.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // Setup stepper navigation
  function setupStepperNavigation() {
    if (!stepperEl) return;
    stepNodes.forEach((node, idx) => {
      node.style.cursor = "pointer";
      node.addEventListener("click", () => {
        const stepNum = parseInt(node.dataset.step);
        if (stepNum === 1) scrollToEl(document.getElementById("fileA"));
        else if (stepNum === 2) scrollToEl(document.getElementById("gestMapBackdrop")?.parentElement || document.querySelector(".card"));
        else if (stepNum === 3) scrollToEl(btnMatch);
        else if (stepNum === 4) scrollToEl(document.getElementById("resultTable"));
        else if (stepNum === 5) scrollToEl(document.getElementById("btnExportFiltered") || document.getElementById("btnExportAll"));
      });
    });
  }

  function refreshMatchButtonState() {
    if (!btnMatch) return;
    const reason = computeBlockingReason();
    const isBlocked = Boolean(reason);
    btnMatch.disabled = isBlocked;
    if (reason) {
      btnMatch.setAttribute("data-disabled-reason", reason);
      btnMatch.title = reason;
      if (statusEl && !flowState.matchDone) statusEl.textContent = reason;
    } else {
      btnMatch.removeAttribute("data-disabled-reason");
      btnMatch.title = "Confronta fatture";
      if (statusEl && !flowState.matchDone) {
        const current = statusEl.textContent || "";
        const isPlaceholder = current.trim() === "" || /Carica|Mapping|attesa/i.test(current);
        if (isPlaceholder) {
          statusEl.textContent = "Pronto a confrontare le fatture.";
        }
      }
    }
  }

  function refreshStepper() {
    const step1Done = flowState.files.ade && flowState.files.gest;
    const step2Done = flowState.mapping.gestBConfirmed && (!flowState.files.nc || flowState.mapping.gestCConfirmed);
    const step3Done = flowState.matchDone;
    const step4Done = flowState.correctionsDone;
    const step5Done = flowState.exportDone;

    setStepState(1, step1Done ? "completed" : (flowState.files.ade || flowState.files.gest ? "active" : "disabled"),
      step1Done ? "File ADE e Gestionale caricati" : "Carica ADE (A) e Gestionale (B)");
    setStepState(2, step2Done ? "completed" : (step1Done ? "active" : "disabled"),
      step2Done ? "Mapping colonne confermato" : "Conferma mapping colonne obbligatorio");
    setStepState(3, step3Done ? "completed" : (step2Done ? "active" : "disabled"),
      step3Done ? "Confronto eseguito" : "Esegui il confronto fatture");
    setStepState(4, step4Done ? "completed" : (step3Done ? "active" : "disabled"),
      step4Done ? "Correzioni salvate" : "Correggi / conferma i match");
    setStepState(5, step5Done ? "completed" : (step3Done ? "active" : "disabled"),
      step5Done ? "Esportazione completata" : "Esporta i risultati");

    // Highlight mapping requirement if files loaded but mapping not confirmed
    const mappingSection = document.getElementById("gestMapBackdrop")?.parentElement;
    if (mappingSection) {
      const needsAttention = step1Done && !step2Done;
      if (needsAttention) {
        mappingSection.classList.add("needs-attention");
      } else {
        mappingSection.classList.remove("needs-attention");
      }
    }
  }

  function refreshFlowUI() {
    refreshMatchButtonState();
    refreshStepper();
  }

  async function handleGestFileSelection(inputEl, mappingKey, description) {
    if (!inputEl) return;
    const file = inputEl.files && inputEl.files[0];
    if (mappingKey === "GEST_B") {
      flowState.files.gest = Boolean(file);
      resetMappingConfirmation("GEST_B");
    } else {
      flowState.files.nc = Boolean(file);
      resetMappingConfirmation("GEST_C");
    }
    flowState.matchDone = false;
    flowState.correctionsDone = false;
    flowState.exportDone = false;

    if (!file) {
      refreshFlowUI();
      return;
    }

    try {
      statusEl.textContent = `Lettura colonne da ${description}...`;
      const content = await loadTableFromFile(file, false);
      const parsed = parseCSV(content);
      cachedHeaders[mappingKey] = parsed.headers;
      await ensureGestConfigForHeaders(parsed.headers, description, true, mappingKey);
    } catch (e) {
      console.error("Errore durante la lettura del file", e);
      alert(`Errore lettura ${description}: ${e?.message || e}`);
    }

    refreshFlowUI();
  }

  // Gestione change input file (step guidati)
  if (fileInputA) {
    fileInputA.addEventListener("change", () => {
      flowState.files.ade = Boolean(fileInputA.files && fileInputA.files[0]);
      flowState.matchDone = false;
      flowState.correctionsDone = false;
      flowState.exportDone = false;
      refreshFlowUI();
    });
  }
  if (fileInputB) {
    fileInputB.addEventListener("change", () => {
      handleGestFileSelection(fileInputB, "GEST_B", "Gestionale Fatture (File B)");
    });
  }
  if (fileInputC) {
    fileInputC.addEventListener("change", () => {
      handleGestFileSelection(fileInputC, "GEST_C", "Note Credito (File C)");
    });
  }

    // --- Eventi filtri avanzati mese / fornitore ---

  if (monthFilterSelect) {
    monthFilterSelect.addEventListener("change", () => {
      const val = monthFilterSelect.value;
      currentMonthFilter = val || null;
      applyFilterAndRender();
    });
  }

  if (supplierFilterInput) {
    supplierFilterInput.addEventListener("input", () => {
      currentSupplierFilter = supplierFilterInput.value || "";
      applyFilterAndRender();
    });
  }

  if (filterInvoiceInput) {
    filterInvoiceInput.addEventListener("input", (e) => {
      currentInvoiceFilter = (e.target.value || "").trim();
      applyFilterAndRender();
    });
  }

  if (btnClearFilters) {
    btnClearFilters.addEventListener("click", () => {
      // reset filtro stato
      currentFilter = "ALL";
      filterButtons.forEach(b => {
        b.classList.toggle("active", b.dataset.filter === "ALL");
      });

      // reset mese
      currentMonthFilter = null;
      if (monthFilterSelect) monthFilterSelect.value = "";

      // reset fornitore
      currentSupplierFilter = "";
      if (supplierFilterInput) supplierFilterInput.value = "";

      // reset filtro numero fattura
      currentInvoiceFilter = "";
      if (filterInvoiceInput) filterInvoiceInput.value = "";

      applyFilterAndRender();
    });
  }

  // ==========================================
  // 🧠 FUNZIONI FILTRO E RENDER (CORRETTE)
  // ==========================================

  function getGlobalFilteredResults() {
    // Se non ci sono risultati, ritorna array vuoto
    if (!lastResults || !lastResults.length) return [];

    // 1. Base: righe non eliminate
    let filtered = lastResults.filter(r => !r.deleted);

    // 2. Filtro Stato (Bottoni in alto: Tutti, Match OK, ecc.)
    if (currentFilter !== "ALL") {
      filtered = filtered.filter(r => r.STATUS === currentFilter);
    }

    // 3. Filtro Mese
    if (currentMonthFilter) {
      filtered = filtered.filter(r => {
        if (!r.ADE || !r.ADE.data) return false;
        return monthKeyFromDate(r.ADE.data) === currentMonthFilter;
      });
    }

    // 4. Filtro Fornitore (Input testo)
    if (currentSupplierFilter && currentSupplierFilter.trim() !== "") {
      const term = currentSupplierFilter.trim().toUpperCase();
      filtered = filtered.filter(r => {
        const aDenNorm = (r.ADE?.denNorm || "").toUpperCase();
        const aDen = (r.ADE?.den || "").toUpperCase();
        const aPiva = (r.ADE?.piva || "").toUpperCase();
        const gDenNorm = (r.GEST?.denNorm || "").toUpperCase();
        const gDen = (r.GEST?.den || "").toUpperCase();
        const gPiva = (r.GEST?.piva || "").toUpperCase();
        return (
          aDenNorm.includes(term) || aDen.includes(term) || aPiva.includes(term) ||
          gDenNorm.includes(term) || gDen.includes(term) || gPiva.includes(term)
        );
      });
    }

    // 5. Filtro Numero Fattura (Input testo)
    if (currentInvoiceFilter && currentInvoiceFilter.trim() !== "") {
      const needle = currentInvoiceFilter.trim().toLowerCase();
      filtered = filtered.filter(r => {
        const adeNum = String(r.ADE?.num || "").toLowerCase();
        const gestNum = String(r.GEST?.num || "").toLowerCase();
        return adeNum.includes(needle) || gestNum.includes(needle);
      });
    }

    return filtered;
  }

  function applyFilterAndRender() {
    // 1. Ottieni i dati filtrati
    const filteredResults = getGlobalFilteredResults();

    // 2. Definisci su cosa calcolare i contatori (totali o filtrati)
    // Se stiamo usando filtri di ricerca (mese, fornitore, fattura), i contatori si adattano alla ricerca.
    // Altrimenti (solo filtro stato), i contatori mostrano il totale generale.
    const isSearching = currentMonthFilter || (currentSupplierFilter && currentSupplierFilter.trim() !== "") || (currentInvoiceFilter && currentInvoiceFilter.trim() !== "");
    
    // Per i contatori usiamo filteredResults se c'è una ricerca attiva, altrimenti usiamo lastResults pulito dai cancellati
    const forCounts = isSearching ? filteredResults : lastResults.filter(r => !r.deleted);

    // 3. Renderizza la tabella
    renderResults(filteredResults, forCounts);

    // 4. Aggiorna KPI dashboard
    renderPeriodKpi(lastResults);
    
    // 5. Aggiorna stato bottone "Export Vista Attuale"
    const btnExpFilt = document.getElementById("btnExportFiltered");
    if(btnExpFilt) {
        btnExpFilt.disabled = (filteredResults.length === 0);
        // Aggiorna il testo per mostrare quante righe esporterai
        const iconSpan = btnExpFilt.querySelector(".btn-icon"); 
        const icon = iconSpan ? iconSpan.outerHTML : "👁️";
        btnExpFilt.innerHTML = `${icon} VISTA ATTUALE (${filteredResults.length})`;
    }
  }

  // ==========================================
  // EVENTI NUOVI BOTTONI (DA INSERIRE QUI O DOPO)
  // ==========================================
  
  // 1. Export VISTA ATTUALE (Filtrati)
  const btnExportFiltered = document.getElementById("btnExportFiltered");
  if (btnExportFiltered) {
    btnExportFiltered.addEventListener("click", () => {
      const dataToExport = getGlobalFilteredResults();
      if (dataToExport.length === 0) {
        alert("Nessuna riga da esportare con i filtri attuali.");
        return;
      }
      exportToCSV(dataToExport, "vista_filtrata");
    });
  }

  // 2. Toggle Espandi / Stringi Tabella (Posizionalo nel main scope, NON dentro altre funzioni)
  const btnToggleExpand = document.getElementById("btnToggleExpand");
  if (btnToggleExpand) {
    btnToggleExpand.addEventListener("click", () => {
      document.body.classList.toggle("wide-mode");
      
      // Piccola finezza: forziamo il ricalcolo layout se necessario
      setTimeout(() => {
          const isWide = document.body.classList.contains("wide-mode");
          btnToggleExpand.innerHTML = isWide ? "⤡ Vista Normale" : "⤢ Modalità Estesa";
      }, 50);
    });
  }

  // ---------- MAPPING AUTOMATICO / MANUALE GESTIONALE ----------

  async function ensureGestConfigForHeaders(headers, fileDescription, forcePopup = false, mappingKey = "GEST_B") {
    const key = mappingKey || "GEST_B";
    cachedHeaders[key] = headers || [];

    // Se forcePopup è true, apri direttamente il modale guidato
    if (forcePopup) {
      const userConfig = await openGestMappingModal(headers, fileDescription, key);
      if (userConfig) {
        markMappingConfirmed(key, headers);
      } else {
        resetMappingConfirmation(key);
      }
      return;
    }

    // Verifica se esiste una config valida e già confermata per queste colonne
    if (hasValidGestConfig(headers, fileDescription, key) && isMappingConfirmed(key, headers)) {
      console.log("✅ Configurazione colonne confermata e valida, procedo automaticamente");
      return; // già ok, proseguo
    }

    // Tentativo di rilevamento automatico per precompilare il modale
    console.log("🔍 Tentativo rilevamento automatico colonne...");
    const autoConfig = tryAutoDetectGestColumns(headers, fileDescription);
    updateColumnConfig(key, autoConfig);

    // Mostra comunque il modale per conferma esplicita (flow guidato)
    console.log("⚠️ Configurazione da confermare: apro il modale mapping");
    const userConfig = await openGestMappingModal(headers, fileDescription, key);
    if (userConfig) {
      markMappingConfirmed(key, headers);
    } else {
      resetMappingConfirmation(key);
    }
  }

  // Funzione helper per rilevamento automatico colonne
  function tryAutoDetectGestColumns(headers, fileDescription) {
    const H = headers;
    const h = H.map(x => x.toLowerCase());
    
    const config = {
      num: null,
      piva: null,
      den: null,
      data: null,
      imp: null,
      iva: null,
      tot: null
    };

    // Numero documento
    const colNumeroAlfa = H.find(hdr =>
      String(hdr || "").toLowerCase().replace(/[^a-z0-9]/g, "") === "numeroalfa"
    );
    const colNumDocEsteso = H.find(hdr =>
      String(hdr || "").toLowerCase().replace(/[^a-z0-9]/g, "") === "numerodocesteso"
    );
    
    if (colNumeroAlfa) {
      config.num = colNumeroAlfa;
    } else if (colNumDocEsteso) {
      config.num = colNumDocEsteso;
    } else {
      config.num = findColGest(H, h, [
        "numerofattura", "numero fattura", "nreggestionale",
        "numeroregistrazione", "numreg", "num", "doc", "documento"
      ]);
    }

    // P.IVA
    config.piva = findColGest(H, h, [
      "partita iva", "p.iva", "piva", "p iva", "pi", "cfpiva",
      "p. iva fornitore", "partitaiva"
    ]);

    // Denominazione
    config.den = findColGest(H, h, [
      "ragione sociale", "ragionesociale", "fornitore",
      "descrizione conto", "descrizione conto dare", "descrizione conto avere", "descr"
    ]);

    // Data
    config.data = findColGest(H, h, [
      "data", "data fattura", "data doc", "data documento",
      "data registrazione", "datareg"
    ]);

    // Imponibile
    config.imp = findColGest(H, h, [
      "tot_imp", "imponibile", "totale imponibile", "totale",
      "totale fattura", "importo documento", "importo", "totale doc"
    ]);

    // IVA
    config.iva = findColGest(H, h, [
      "iva", "imposta", "importoiva", "tot_iva"
    ]);

    // Totale
    config.tot = findColGest(H, h, [
      "totale", "totalefattura", "tot_fattura", "importo documento",
      "importo", "totale doc", "totdoc"
    ]);

    return config;
  }

  function initMonthFilterOptions(results) {
  if (!monthFilterSelect) return;

  const monthSet = new Set();

  for (const r of results || []) {
    if (r.deleted) continue;
    if (!r.ADE || !r.ADE.data) continue;
    const key = monthKeyFromDate(r.ADE.data);
    monthSet.add(key);
  }

  // Svuota e ricostruisci le opzioni
  monthFilterSelect.innerHTML = "";

  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "Tutti i mesi ADE";
  monthFilterSelect.appendChild(optAll);

  const months = Array.from(monthSet).sort(); // "2025-01", "2025-02", ...
  months.forEach(key => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = monthLabelFromKey(key);
    monthFilterSelect.appendChild(opt);
  });

  // Abilito/disabilito a seconda che ci siano mesi
  monthFilterSelect.disabled = months.length === 0;

  // reset filtro mese
  currentMonthFilter = null;
  monthFilterSelect.value = "";
}

  btnMatch.addEventListener("click", async () => {
    const blockReason = computeBlockingReason();
    if (blockReason) {
      statusEl.textContent = blockReason + " (step bloccato)";
      refreshFlowUI();
      return;
    }

    statusEl.textContent = "Click rilevato, inizio lettura file...";
    const fA = document.getElementById("fileA").files[0];
    const fB = document.getElementById("fileB").files[0];
    const fC = document.getElementById("fileC").files[0];

    if (!fA || !fB) {
      statusEl.textContent = "Carica almeno il file ADE (A) e il file fatture gestionale (B). Il file C (note credito) è opzionale.";
      return;
    }

    // leggo eventuali modifiche manuali nel box JSON (columns / rules)
    readLearningFromTextarea();

    btnMatch.disabled = true;
    statusEl.textContent = "Elaborazione in corso...";

    try {
      const [txtA, txtB, txtC] = await Promise.all([
        loadTableFromFile(fA, true),   // File A = ADE → mantieni punto decimale
        loadTableFromFile(fB, false),  // File B = GEST → converti virgola decimale
        fC ? loadTableFromFile(fC, false) : Promise.resolve("")  // File C = Note Credito GEST → virgola decimale
      ]);

      const parsedA = parseCSV(txtA);
      const parsedB = parseCSV(txtB);
      const parsedC = fC ? parseCSV(txtC) : { headers: [], rows: [] };

      // 🟦 1) CHIEDO MAPPING PER FILE B (bloccante)
      await ensureGestConfigForHeaders(parsedB.headers, "Gestionale Fatture (File B)", false, "GEST_B");

      // 🟪 1-bis) SE ESISTE FILE C, CHIEDO MAPPING ANCHE PER C (indipendente)
      if (parsedC && parsedC.headers && parsedC.headers.length > 0) {
        await ensureGestConfigForHeaders(parsedC.headers, "Note Credito (File C)", false, "GEST_C");
      }

      // 🟩 2) SOLO ORA costruisco i record
      const adeRecords = buildAdeRecords(parsedA);
      const gestAcq = buildGestAcqRecords(parsedB, { fileName: fB?.name, isNotesCredit: false });
      const gestNc  = parsedC.rows.length ? buildGestNcRecords(parsedC, { fileName: fC?.name, isNotesCredit: true }) : [];

      const gestRecords = gestAcq.concat(gestNc);

      // contatori totali di righe caricate dai file
      totalAdeCount = adeRecords.length;
      totalGestCount = gestRecords.length;

      // salva globalmente
      // --- STEP DI PRE-PROCESSING: RECUPERO P.IVA ---
      // Creiamo una mappa: Nome Fornitore Normalizzato -> P.IVA (dal file ADE)
      const supplierMap = new Map();

      // 1. Popola la mappa usando ADE (che ha i dati sicuri)
      adeRecords.forEach(row => {
          // Normalizziamo il nome rimuovendo SPA, SRL, spazi, simboli
          const cleanName = normalizeName(row.den); 
          const piva = normalizePIVA(row.piva || ""); // Usa la nuova normalizePIVA
          if (cleanName && piva) {
              supplierMap.set(cleanName, piva);
          }
      });

      // 2. Applica la P.IVA ai record del Gestionale che non ce l'hanno
      gestRecords.forEach(row => {
          if (!row.piva || normalizePIVA(row.piva) === "") { // Se non ha piva o è vuota dopo normalizzazione
              const cleanName = normalizeName(row.den);
              if (supplierMap.has(cleanName)) {
                  row.piva = supplierMap.get(cleanName);
                  row.pivaDigits = onlyDigits(row.piva); // Aggiorna anche la versione solo cifre
              }
          }
      });
      lastAdeRecords = adeRecords;
      lastGestRecords = gestRecords;

      // aggiorna copertura periodo
      updatePeriodCoverage(adeRecords, gestRecords);

      // aggiorna periodo analisi fornitori se non ancora inizializzato
      if (!analysisMinDate || !analysisMaxDate) {
        initAnalysisPeriodFromAde();
      }

      lastResults = matchRecords(adeRecords, gestRecords);
      
      // 🔍 Post-processing: flagga casi sospetti SOLO_ADE/SOLO_GEST
      lastResults = flagSuspiciousSoloAdeSoloGest(lastResults);
      
      nextResultId = 1;
      lastResults.forEach(r => {
        r.id = nextResultId++;
        r.rowId = r.id; // Mantiene la compatibilità con il pannello di modifica
        r.deleted = false;
        r.matchNote = r.matchNote || "";
      });

      // salvo lo stato iniziale dopo il match (serve per "Ripristina eliminazioni")
      originalResults = JSON.parse(JSON.stringify(lastResults));
      undoStack = [];
      btnResetAll.disabled = false;
      btnUndoDelete.disabled = true;

      currentFilter = "ALL";
      filterButtons.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.filter === "ALL");
      });

      // aggiorna tabella riconciliazione
      applyFilterAndRender();
      // aggiorna dashboard KPI
      renderPeriodKpi(lastResults);

      // 🔹 inizializza i filtri avanzati dopo aver calcolato i risultati
      initMonthFilterOptions(lastResults);
      currentMonthFilter = null;
      currentSupplierFilter = "";
      if (monthFilterSelect) monthFilterSelect.value = "";
      if (supplierFilterInput) supplierFilterInput.value = "";
      if (advancedFilterRow) {
        advancedFilterRow.style.display = "flex";
      }

      btnExportAll.disabled = false;
      btnExportAde.disabled = false;
      btnExportGest.disabled = false;
      btnExportFix.disabled = false;

      flowState.matchDone = true;
      flowState.correctionsDone = false;
      flowState.exportDone = false;
      refreshFlowUI();
    } catch (e) {
      console.error(e);
      statusEl.textContent = "Errore durante il confronto: " + (e && e.message ? e.message : e);
    } finally {
      btnMatch.disabled = false;
    }
  });

  filterButtons.forEach(btn => {
    btn.addEventListener("click", () => {
      const filter = btn.dataset.filter;
      currentFilter = filter;
      filterButtons.forEach(b => b.classList.toggle("active", b === btn));
      applyFilterAndRender();
    });
  });

  btnExportAll.addEventListener("click", () => {
    if (lastResults && lastResults.length) {
      exportToCSV(lastResults, "tutte");
      flowState.exportDone = true;
      refreshFlowUI();
    }
  });

  btnExportAde.addEventListener("click", () => {
    const subset = lastResults.filter(r => !r.deleted && r.STATUS === "SOLO_ADE");
    exportToCSV(subset, "solo_ADE");
    flowState.exportDone = true;
    refreshFlowUI();
  });

  btnExportGest.addEventListener("click", () => {
    const subset = lastResults.filter(r => !r.deleted && r.STATUS === "SOLO_GEST");
    exportToCSV(subset, "solo_GEST");
    flowState.exportDone = true;
    refreshFlowUI();
  });

  btnExportFix.addEventListener("click", () => {
    const subset = lastResults.filter(r => !r.deleted && r.STATUS === "MATCH_FIX");
    exportToCSV(subset, "solo_FIX");
    flowState.exportDone = true;
    refreshFlowUI();
  });

  // Annulla ultimo elimina
  btnUndoDelete.addEventListener("click", () => {
    if (!undoStack.length) return;
    lastResults = undoStack.pop();
    applyFilterAndRender();
    if (!undoStack.length) {
      btnUndoDelete.disabled = true;
    }
  });

  // Ripristina tutte le eliminazioni
  btnResetAll.addEventListener("click", () => {
    if (!originalResults) return;
    lastResults = JSON.parse(JSON.stringify(originalResults));
    undoStack = [];
    btnUndoDelete.disabled = true;
    flowState.correctionsDone = false;
    flowState.exportDone = false;
    refreshFlowUI();
    applyFilterAndRender();
  });

  // Pulsanti per forzare il mapping
  const btnForceMapB = document.getElementById("btnForceMapB");
  const btnForceMapC = document.getElementById("btnForceMapC");

  async function handleForceMap(fileInputId, fileDescription, mappingKey = "GEST_B") {
      const fileInput = document.getElementById(fileInputId);
      if (!fileInput || !fileInput.files[0]) {
          alert(`Per forzare il mapping, prima seleziona un file per "${fileDescription}".`);
          return;
      }
      
      try {
          statusEl.textContent = `Lettura colonne dal file ${fileDescription}...`;
          // File B e C sono sempre GESTIONALE → isADE = false
          const fileContent = await loadTableFromFile(fileInput.files[0], false);
          const parsed = parseCSV(fileContent);
          
          if (!parsed.headers || parsed.headers.length === 0) {
              alert("Impossibile leggere le colonne da questo file.");
              return;
          }

          // Chiamiamo la funzione di mapping forzando la comparsa del popup
          await ensureGestConfigForHeaders(parsed.headers, fileDescription, true, mappingKey);
          statusEl.textContent = `Mapping per "${fileDescription}" aggiornato. Ora puoi premere "Confronta fatture".`;
          flowState.matchDone = false;
          flowState.correctionsDone = false;
          flowState.exportDone = false;
          refreshFlowUI();

      } catch (e) {
          console.error("Errore forzando il mapping:", e);
          alert("Si è verificato un errore durante la lettura del file.");
      }
  }

  btnForceMapB.addEventListener("click", () => handleForceMap("fileB", "Gestionale Fatture (File B)", "GEST_B"));
  btnForceMapC.addEventListener("click", () => handleForceMap("fileC", "Note Credito (File C)", "GEST_C"));

  // ---------- editor: apertura / chiusura / salvataggio ----------

  const editBackdrop = document.getElementById("editPanelBackdrop");
  const editCloseBtn = document.getElementById("editCloseBtn");
  const editCancelBtn = document.getElementById("editCancelBtn");
  const editSaveBtn = document.getElementById("editSaveBtn");
  const editDeleteGestBtn = document.getElementById("editDeleteGestBtn");
  const editRowInfo = document.getElementById("editRowInfo");
    const editBulkOkSupplier = document.getElementById("editBulkOkSupplier");

  const editAdeNum = document.getElementById("editAdeNum");
  const editAdeDen = document.getElementById("editAdeDen");
  const editAdePiva = document.getElementById("editAdePiva");
  const editAdeData = document.getElementById("editAdeData");
  const editAdeImpIvaTot = document.getElementById("editAdeImpIvaTot");

  const editGestSection = document.getElementById("editGestSection");
  const editGestNum = document.getElementById("editGestNum");
  const editGestDen = document.getElementById("editGestDen");
  const editGestData = document.getElementById("editGestData");
  const editGestImp = document.getElementById("editGestImp");
  const editGestIva = document.getElementById("editGestIva");
  const editGestTot = document.getElementById("editGestTot");

  const editStatusSelect = document.getElementById("editStatus");
  const editMatchNote = document.getElementById("editMatchNote");

  // --- FUNZIONE DIAGNOSTICA PER CAPIRE PERCHÉ NON MATCHA ---
  function generateDiagnosticReport(a, g, criterio) {
      if (!a || !g) return "Manca una delle due parti (Solo ADE o Solo GEST).";
  
      let report = [];
      
      // 1. Analisi Importi
      const diff = a.tot - g.tot;
      if (Math.abs(diff) > 0.01) {
          report.push(`❌ IMPORTO DIVERSO: Scarto di ${diff.toFixed(2)} € (ADE: ${a.tot} vs GEST: ${g.tot})`);
          if (Math.abs(Math.abs(diff) - 2.00) < 0.01) report.push("   -> Probabile Marca da Bollo mancante.");
          if (Math.abs(a.imp - g.imp) < 0.01) report.push("   -> L'Imponibile però coincide! (Problema IVA o Esente).");
      } else {
          report.push("✅ IMPORTO: Coincide.");
      }
  
        // 2. Analisi Date
      if (a.data && g.data) {
          const diffTime = Math.abs(a.data - g.data);
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
          if (diffDays > 0) {
              let msg = `⚠️ DATA DIVERSA: Distanza di ${diffDays} giorni (ADE: ${formatDateIT(a.data)} vs GEST: ${formatDateIT(g.data)})`;

              // check inversione giorno/mese (es: 03/12 vs 12/03)
              const dayA   = a.data.getDate();
              const monthA = a.data.getMonth() + 1;
              const dayG   = g.data.getDate();
              const monthG = g.data.getMonth() + 1;

              if (dayA === monthG && monthA === dayG) {
                  msg += " -> possibile inversione giorno/mese tra i due sistemi.";
              }

              report.push(msg);
          } else {
              report.push("✅ DATA: Coincide.");
          }
      } else {
          report.push("⚠️ DATA: Una delle date non è valida.");
      }

           // 3. Analisi Numero Fattura
                  const nA  = normalizeInvoiceNumber(a.num, true);   // con lettere
      const nG  = normalizeInvoiceNumber(g.num, true);
      const nAd = normalizeInvoiceNumber(a.num, false);  // solo cifre
      const nGd = normalizeInvoiceNumber(g.num, false);

      const digitsCompat = invoiceDigitsRelated(a.num, g.num, 2);

      if (nAd && nGd && nAd === nGd) {
        report.push("✅ NUMERO: Coincide sulle cifre (differenza solo lettere/prefissi).");
        report.push(`   ADE (Grezz): "${a.num}" -> Cifre: "${nAd}"`);
        report.push(`   GEST (Grezz): "${g.num}" -> Cifre: "${nGd}"`);
      } else if (digitsCompat) {
        report.push("🟡 NUMERO: Compatibile (prefisso/suffisso o parte del numero).");
        report.push(`   ADE (Grezz): "${a.num}" -> Cifre: "${nAd}"`);
        report.push(`   GEST (Grezz): "${g.num}" -> Cifre: "${nGd}"`);
      } else if (nA === nG) {
        report.push("✅ NUMERO: Coincide.");
      } else {
        report.push(`❌ NUMERO DIVERSO:`);
        report.push(`   ADE (Grezz): "${a.num}" -> Norm: "${nA}"`);
        report.push(`   GEST (Grezz): "${g.num}" -> Norm: "${nG}"`);
      }
  
      // 4. Criterio usato dal sistema
      report.push(`\nℹ️ MOTIVO SISTEMA: ${criterio || "Nessuno"}`);
  
      return report.join("\n");
  }

  function getRowById(rowId) {
    return lastResults.find(r => r.rowId === rowId && !r.deleted);
  }

  function openEditPanel(rowId) {
    const row = getRowById(rowId);
    if (!row) return;

    currentEditRowId = rowId;

    const a = row.ADE;
    const g = row.GEST;

    // ADE
    editAdeNum.textContent = a ? a.num || "" : "";
    editAdeDen.textContent = a ? a.den || "" : "";
    editAdePiva.textContent = a ? a.piva || "" : "";
    editAdeData.textContent = a && a.data
      ? formatDateIT(a.data)
      : (a ? a.dataStr || "" : "");
    editAdeImpIvaTot.textContent = a
      ? `${formatNumberITDisplay(a.imp)} / ${formatNumberITDisplay(a.iva)} / ${formatNumberITDisplay(a.tot)}`
      : "";

    // GEST
    if (g) {
      editGestSection.style.display = "flex";
      editGestNum.value = g.num || "";

      editGestDen.value = g.den || "";

      // --- NUOVO: Aggiunta Checkbox "Applica a tutti" ---
      const denContainer = document.getElementById("editGestDen").parentNode;
      
      // Pulizia vecchi elementi aggiunti dinamicamente
      const oldChk = denContainer.querySelector(".bulk-rename-wrapper");
      if(oldChk) oldChk.remove();
      const oldBtn = denContainer.querySelector(".btn-fix-name"); // Rimuovo anche il bottone magico se c'era
      if(oldBtn) oldBtn.remove();

      // Creo il contenitore per la checkbox
      const wrapper = document.createElement("div");
      wrapper.className = "bulk-rename-wrapper";
      wrapper.style.marginTop = "6px";
      wrapper.style.display = "flex";
      wrapper.style.alignItems = "center";
      wrapper.style.gap = "6px";

      // Creo la checkbox
      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.id = "chkApplyToAll";
      chk.style.cursor = "pointer";

      // Creo l'etichetta
      const lbl = document.createElement("label");
      lbl.htmlFor = "chkApplyToAll";
      lbl.textContent = "Applica cambio nome a tutte le righe di questo fornitore";
      lbl.style.fontSize = "0.7rem";
      lbl.style.color = "#fcd34d"; // Un colore giallo/oro per evidenziarlo
      lbl.style.cursor = "pointer";

      wrapper.appendChild(chk);
      wrapper.appendChild(lbl);
      denContainer.appendChild(wrapper);
      // --------------------------------------------------
            editGestData.value = g.data
        ? formatDateIT(g.data)     // sempre DD/MM/YYYY
        : (g.dataStr || "");
      editGestImp.value = g.imp != null ? formatNumberITDisplay(g.imp) : "";
      editGestIva.value = g.iva != null ? formatNumberITDisplay(g.iva) : "";
      editGestTot.value = g.tot != null ? formatNumberITDisplay(g.tot) : "";
      editDeleteGestBtn.disabled = false;
    } else {
      editGestSection.style.display = "none";
      editGestNum.value = "";
      editGestDen.value = "";
      editGestData.value = "";
      editGestImp.value = "";
      editGestIva.value = "";
      editGestTot.value = "";
      editDeleteGestBtn.disabled = true;
    }

    // --- MODIFICA DA QUI IN GIU ---
    // Stato: abilito solo quelli coerenti
    const options = Array.from(editStatusSelect.options);
    const allowed = new Set();

    if (a && g) {
      // riga con ADE + GEST: posso decidere se è un match oppure
      // staccare una delle due parti
      allowed.add("MATCH_OK");
      allowed.add("MATCH_FIX");
      allowed.add("SOLO_ADE");
      allowed.add("SOLO_GEST");
    }
    if (a && !g) {
      allowed.add("SOLO_ADE");
    }
    if (!a && g) {
      allowed.add("SOLO_GEST");
    }

    options.forEach(opt => {
      opt.disabled = !allowed.has(opt.value);
    });

    let targetStatus = row.STATUS;
    if (!allowed.has(targetStatus) && allowed.size > 0) {
      targetStatus = Array.from(allowed)[0];
    }
    editStatusSelect.value = targetStatus;

    // >>> LOGICA DIAGNOSTICA AUTOMATICA <<<
    // Se c'è già una nota manuale dell'utente, la manteniamo.
    // Altrimenti, generiamo il report diagnostico.
    const adeExtra = (a && a.techNote) ? `\n${a.techNote}\n` : "";
    if (row.matchNote && row.matchNote.trim() !== "") {
      editMatchNote.value = row.matchNote + adeExtra;
    } else {
      // Genera diagnosi automatica se siamo in FIX o se vogliamo info
      if (a && g) {
        editMatchNote.value = adeExtra + generateDiagnosticReport(a, g, row.CRITERIO).trim();
      } else {
        editMatchNote.value = adeExtra + (row.CRITERIO || "").trim();
      }
    }

    editRowInfo.textContent = `Riga #${row.rowId} – Stato attuale: ${row.STATUS}`;

    editBackdrop.classList.remove("hidden");
  }

  function closeEditPanel() {
    currentEditRowId = null;
    editBackdrop.classList.add("hidden");
  }

  function saveEditPanel() {
    if (currentEditRowId == null) return;
    const row = getRowById(currentEditRowId);
    if (!row) {
      closeEditPanel();
      return;
    }

    // 1) Gestione GESTIONALE e RINOMINA MASSIVA
    if (row.GEST) {
      const g = row.GEST;
      const oldDen = g.den; // Ci salviamo il vecchio nome prima di cambiarlo
      const newDen = editGestDen.value.trim(); // Prendo il nuovo nome
      
      // Controllo se l'utente vuole applicare a tutti
      const applyToAll = document.getElementById("chkApplyToAll")?.checked;

      // Aggiornamento standard della riga corrente
      g.num = editGestNum.value.trim();
      g.den = newDen;
      g.dataStr = editGestData.value.trim();
      g.data = parseDate(g.dataStr);
      g.imp = parseNumberIT(editGestImp.value);
      g.iva = parseNumberIT(editGestIva.value);
      g.tot = parseNumberIT(editGestTot.value);

      // --- LOGICA BULK UPDATE ---
      if (applyToAll && oldDen !== newDen) {
          if (confirm(`Attenzione: stai per rinominare TUTTE le righe del gestionale da:\n"${oldDen}"\na:\n"${newDen}"\n\nConfermi e ricalcoli i match?`)) {
              
              let count = 0;
              // Aggiorno direttamente la fonte dati globale
              for (const r of lastGestRecords) {
                  if (r.den === oldDen) {
                      r.den = newDen;
                      r.denNorm = normalizeName(newDen);
                      count++;
                  }
              }

              // Ricalcolo tutto da zero con i nuovi nomi
              lastResults = matchRecords(lastAdeRecords, lastGestRecords);
              
              // 🔍 Post-processing: flagga casi sospetti SOLO_ADE/SOLO_GEST
              lastResults = flagSuspiciousSoloAdeSoloGest(lastResults);
              
              // Ripristino ID univoci
              nextResultId = 1;
              lastResults.forEach(r => {
                r.id = nextResultId++;
                r.rowId = r.id;
                r.deleted = false;
                r.matchNote = r.matchNote || "";
              });

              alert(`Aggiornate ${count} righe e ricalcolato il confronto!`);
              
              closeEditPanel();
              applyFilterAndRender();
              return; // Esco subito perché ho rigenerato tutto
          }
      }
    }

    // 2) Aggiornamento stato
    const newStatus = editStatusSelect.value;

    // 3) Gestione separazione righe
    if (newStatus === "SOLO_GEST" && row.GEST && row.ADE) {
        const newSoloAde = {
          id: nextResultId++,
          rowId: nextResultId, // per compatibilità
          ADE: row.ADE,
          GEST: null,
          STATUS: "SOLO_ADE",
          CRITERIO: row.CRITERIO || "Distaccata manualmente dal match",
          deleted: false,
          matchNote: row.matchNote || ""
        };
        lastResults.push(newSoloAde);
      
      row.ADE = null; // La riga corrente diventa SOLO_GEST
      row.STATUS = "SOLO_GEST";
    } else {
      row.STATUS = newStatus;
    }

    // 4) Aggiornamento note e apprendimento (uguale a prima)
    if (row.ADE && row.GEST && (row.STATUS === "MATCH_OK" || row.STATUS === "MATCH_FIX")) {      
        const cleanStr = (s) => String(s || "").trim().replace(/^['"]|['"]$/g, "");
        const adeNum = cleanStr(row.ADE.num).toUpperCase();
        const gestNum = cleanStr(row.GEST.num).toUpperCase();
        let detectedRule = null;

        if (adeNum.replace(/^0+/, '') === gestNum) {
            detectedRule = "REMOVE_ZEROS";
        } else if (adeNum.split('/')[0] === gestNum) {
            detectedRule = "REMOVE_YEAR";
        } else if (adeNum.length > 1 && adeNum.substring(1) === gestNum) {
            detectedRule = "REMOVE_FIRST_CHAR";
        } else if (gestNum.includes(adeNum)) {
            detectedRule = "GEST_CONTAINS_ADE";
        }

        if (detectedRule) {
            const piva = cleanStr(row.ADE.piva || "ND");
            const den = cleanStr(row.ADE.den || "Fornitore");
            const newRule = {
                tipo: "TIPO_TRASFORMAZIONE",
                piva: piva,
                rule: detectedRule,
                note: `Auto-generata da match: ${adeNum} -> ${gestNum}`,
                createdAt: new Date().toISOString()
            };
            const ruleExists = learningData.rules.some(r => 
                r.tipo === newRule.tipo && 
                r.piva === newRule.piva && 
                r.rule === newRule.rule
            );

            if (!ruleExists) {
                learningData.rules.unshift(newRule); // Aggiungi in cima per visibilità
                refreshLearningPanel();
                const learningBox = document.getElementById("learningCodeBox");
                if (learningBox) {
                    learningBox.style.borderColor = "#22c55e";
                    learningBox.style.boxShadow = "0 0 15px rgba(34, 197, 94, 0.5)";
                    setTimeout(() => {
                        learningBox.style.borderColor = "";
                        learningBox.style.boxShadow = "";
                    }, 2500);
                }
                alert(`✅ ContAIbil ha imparato una nuova regola per ${den}!\n\nRegola: ${detectedRule}\nLa procedura è stata salvata automaticamente. Ricorda di usare "Salva Procedure" per mantenerla per le prossime sessioni.`);
            }
        } else {
            const adeRowLearn = { ADE_PIVA: row.ADE.piva, ADE_Num: row.ADE.num, ADE_Den: row.ADE.den, ADE_Totale: row.ADE.tot };
            const gestRowLearn = { GEST_PIVA: row.GEST.piva, GEST_Num: row.GEST.num, GEST_Den: row.GEST.den, GEST_Totale: row.GEST.tot };
            registerLearningRule(adeRowLearn, gestRowLearn, "MATCH_ESTESO");
        }
    }

    row.matchNote = (editMatchNote.value || "").trim();
    flowState.correctionsDone = true;
    flowState.exportDone = false;
    refreshFlowUI();
    applyFilterAndRender();
    closeEditPanel();
}

  function deleteGestPart(rowId) {
    const row = lastResults.find(r => r.rowId === rowId);
    if (!row || !row.GEST) return;

    // salvo lo stato corrente per poter fare "Annulla ultimo elimina"
    undoStack.push(JSON.parse(JSON.stringify(lastResults)));
    btnUndoDelete.disabled = false;

    if (row.ADE) {
      // da match a SOLO_ADE
      row.GEST = null;
      row.STATUS = "SOLO_ADE";
      row.CRITERIO = row.CRITERIO || "GESTIONALE eliminato manualmente";
    } else {
      // SOLO_GEST → riga eliminata del tutto
      row.deleted = true;
    }
    flowState.correctionsDone = true;
    flowState.exportDone = false;
    refreshFlowUI();
    applyFilterAndRender();
  }

    function bulkApproveMatchesForSupplier(baseRow, applyOnlyFix = true) {
    if (!baseRow || !baseRow.ADE || !baseRow.GEST) return;

    // chiave fornitore: P.IVA (se c'è) + denominazione normalizzata
    const baseAde = baseRow.ADE;
    const supplierKey = (baseAde.pivaDigits || "") + "|" + (baseAde.denNorm || "");

    if (!supplierKey || supplierKey === "|") {
      alert("Non riesco a identificare in modo univoco il fornitore (manca P.IVA / denominazione normalizzata).");
      return;
    }

    const candidates = lastResults.filter(r => {
      if (!r || r.deleted) return false;
      if (!r.ADE || !r.GEST) return false;
      const key = (r.ADE.pivaDigits || "") + "|" + (r.ADE.denNorm || "");
      return key === supplierKey;
    });

    const matchFixCount = candidates.filter(r => r.STATUS === "MATCH_FIX").length;
    const messageLines = [
      `Fornitore: ${baseAde.den || "n/d"} (${baseAde.piva || "n/d"})`,
      `Righe totali coinvolte: ${candidates.length}`,
      `Righe in stato MATCH_FIX: ${matchFixCount}`,
      `Modalità: ${applyOnlyFix ? "Solo righe MATCH_FIX (consigliato)" : "Tutte le righe coerenti"}`
    ];

    const proceed = confirm(`Conferma azione massiva?\n\n${messageLines.join("\n")}\n\nVuoi procedere con l'approvazione?`);
    if (!proceed) return;

    // Opzione: includere anche righe non MATCH_FIX (richiesta esplicita)
    const includeAlsoNonFix = confirm("Vuoi includere anche le righe NON in stato MATCH_FIX? (OK = includi, Annulla = solo MATCH_FIX)");
    applyOnlyFix = !includeAlsoNonFix;

    // salvo lo stato per poter fare "Annulla ultimo" anche dopo bulk
    undoStack.push(JSON.parse(JSON.stringify(lastResults)));
    btnUndoDelete.disabled = false;

    let changedCount = 0;

    for (const r of candidates) {
      if (!r || r.deleted) continue;
      if (applyOnlyFix && r.STATUS !== "MATCH_FIX") continue;

      const ade = r.ADE;
      const gest = r.GEST;

      const diffTot = Math.abs(ade.tot - gest.tot);
      const totOk = diffTot < 0.05;

      const dateOk = datesClose(ade.data, gest.data, 2);

      const adeDigits  = normalizeInvoiceNumber(ade.num, false);
      const gestDigits = normalizeInvoiceNumber(gest.num, false);
      const rawGestNum = String(gest.num || "").toUpperCase().trim();
      const isSci = /E\+?\d+$/i.test(rawGestNum);

      const numOk =
        (adeDigits && gestDigits && adeDigits === gestDigits) ||
        isSci;  // se è scientifico, numero gestionale inaffidabile ma accettato nel bulk

      if (!totOk || !dateOk || !numOk) continue;

      r.STATUS = "MATCH_OK";
      r.CRITERIO = (r.CRITERIO || "") + " | BulkOK fornitore";
      r.matchNote = (r.matchNote || "") + "\n[APPROVATO MASSIVO] Confermato dall'utente per questo fornitore.";

      changedCount++;
    }

    applyFilterAndRender();
    flowState.correctionsDone = true;
    flowState.exportDone = false;
    refreshFlowUI();

    if (changedCount > 0) {
      alert(`Sono stati confermati automaticamente ${changedCount} match per questo fornitore.\nControlla comunque un campione per sicurezza.`);
    } else {
      alert("Non ho trovato righe sicure da confermare per questo fornitore con i criteri attuali.");
    }
  }

  // click su tabella: row click / delete click
  const resultTableBody = document.querySelector("#resultTable tbody");
  if (resultTableBody) {
    resultTableBody.addEventListener("click", (e) => {
      const delBtn = e.target.closest(".row-delete-btn");
      const tr = e.target.closest("tr");
      if (!tr) return;
      const idAttr = tr.dataset.id || tr.dataset.rowId;
      if (!idAttr) return;
      const id = parseInt(idAttr, 10);
      if (isNaN(id)) return;

      // 1. Gestione eliminazione (rimane uguale)
      if (delBtn) {
        // elimina parte gestionale
        deleteGestPart(id);
        e.stopPropagation();
        return;
      }

      // 2. NUOVO: Gestione click per completare match manuale
      if (pendingMatchSourceId !== null) {
        // Se siamo in modalità "attesa secondo click", questo click serve a chiudere il match
        completeManualMatch(id);
        return; // Fermiamo qui, non aprire l'editor
      }

      // apertura editor solo per MATCH_FIX e SOLO_GEST (e, se vuoi, anche per altri)
      // 3. Gestione apertura editor (rimane uguale)
      const row = getRowById(id);
      if (!row) return;
      // Ora apriamo l'editor anche per MANUAL_MATCH per permettere di modificare le note
      if (row.STATUS === "MATCH_FIX" || row.STATUS === "SOLO_GEST" || row.STATUS === "MANUAL_MATCH" || row.STATUS === "SOLO_ADE") {
        openEditPanel(id);
      }
    });
  }

  // eventi editor
  editCloseBtn.addEventListener("click", () => {
    closeEditPanel();
  });
  editCancelBtn.addEventListener("click", () => {
    closeEditPanel();
  });
  editBackdrop.addEventListener("click", (e) => {
    if (e.target === editBackdrop) {
      closeEditPanel();
    }
  });
  editSaveBtn.addEventListener("click", () => {
    saveEditPanel();
  });
  editDeleteGestBtn.addEventListener("click", () => {
    if (currentEditRowId != null) {
      deleteGestPart(currentEditRowId);
    }
    closeEditPanel();
  });
    if (editBulkOkSupplier) {
    editBulkOkSupplier.addEventListener("click", () => {
      if (currentEditRowId == null) return;
      const row = getRowById(currentEditRowId);
      if (!row) return;

      if (!row.ADE || !row.GEST) {
        alert("Serve una riga con ADE e Gestionale per identificare il fornitore.");
        return;
      }

      bulkApproveMatchesForSupplier(row, true);
    });
  }

  // ---------- event handlers analisi fornitori ----------

  anApplyBtn.addEventListener("click", () => {
    computeSupplierAnalysis();
  });

  anAllBtn.addEventListener("click", () => {
    if (!analysisMinDate || !analysisMaxDate) {
      initAnalysisPeriodFromAde();
      return;
    }
    anStartInput.value = dateToInputValue(analysisMinDate);
    anEndInput.value = dateToInputValue(analysisMaxDate);
    computeSupplierAnalysis();
  });

  anYearBtn.addEventListener("click", () => {
    const today = new Date();
    const year = today.getFullYear();
    const start = new Date(year, 0, 1);
    const end = new Date(year, 11, 31);
    anStartInput.value = dateToInputValue(start);
    anEndInput.value = dateToInputValue(end);
    computeSupplierAnalysis();
  });

  anLast12Btn.addEventListener("click", () => {
    let end = analysisMaxDate || new Date();
    let start = new Date(end.getTime());
    start.setDate(start.getDate() - 365);
    anStartInput.value = dateToInputValue(start);
    anEndInput.value = dateToInputValue(end);
    computeSupplierAnalysis();
  });

  anSearchInput.addEventListener("input", () => {
    renderSupplierAnalysis();
  });

  // ---------- nav tab grafica ----------
  const navTabs = Array.from(document.querySelectorAll(".nav-tab"));
  const sections = {
    "section-recon": document.getElementById("section-recon"),
    "section-analysis": document.getElementById("section-analysis"),
    "section-dashboard": document.getElementById("section-dashboard"),
    "section-vat": document.getElementById("section-vat"), // <--- AGGIUNTA
  };

  navTabs.forEach(tab => {
    tab.addEventListener("click", () => {
      const targetId = tab.dataset.section;
      if (!targetId) return;

      navTabs.forEach(t => t.classList.toggle("active", t === tab));
      Object.entries(sections).forEach(([id, el]) => {
        el.classList.toggle("active", id === targetId);
      });
    });
  });

    // inizializzo pannello apprendimento
  setupLearningButtons();

  // Carica automaticamente le procedure salvate all'avvio
  // 1) localStorage, 2) file JSON accanto all'HTML (se disponibile)
  loadProceduresAtStartup();

  // Inizializza il gestore del tema (Dark/Light mode)
  ThemeManager.init();

  refreshLearningPanel();
  refreshFlowUI();

  function buildAiCopyPayload(row) {
      if (!row) return null;

      const fmtMoney = (v) => {
        const num = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
        if (isNaN(num)) return "n/d";
        return formatNumberITDisplay(num);
      };

      const sideBlock = (label, rec) => {
        if (!rec) return [`${label}: assente`];
        const tipo = rec.isNC ? "NC" : label;
        const dataStr = rec.data ? formatDateIT(rec.data) : (rec.dataStr || rec.dataIso || "n/d");
        return [
          `${label} | Tipo record: ${tipo}`,
          `Numero documento: ${rec.num || "n/d"}`,
          `Data: ${dataStr || "n/d"}`,
          `Fornitore: ${rec.den || "n/d"} (PIVA/CF: ${rec.piva || "n/d"})`,
          `Imponibile / IVA / Totale: ${fmtMoney(rec.imp)} / ${fmtMoney(rec.iva)} / ${fmtMoney(rec.tot)}`
        ];
      };

      const lines = [];
      lines.push(...sideBlock("ADE", row.ADE));
      lines.push("---");
      lines.push(...sideBlock(row.GEST?.isNC ? "NC" : "GEST", row.GEST));
      lines.push("---");
      lines.push(`Stato match: ${row.STATUS} | Criterio: ${row.CRITERIO || "n/d"}`);

      const mapKey = row.GEST?.isNC ? "GEST_C" : "GEST_B";
      const cfg = getColumnConfig(mapKey);
      lines.push(`Mapping ${mapKey}: num=${cfg.num || "n/d"}, data=${cfg.data || "n/d"}, imp=${cfg.imp || "n/d"}, iva=${cfg.iva || "n/d"}, tot=${cfg.tot || "n/d"}`);

      return lines.join("\n");
  }

  // --- FUNZIONE PER IL BOTTONE "COPIA REPORT" ---
  const btnCopyReport = document.getElementById("btnCopyReport");
  if (btnCopyReport) {
      btnCopyReport.addEventListener("click", () => {
          const txtArea = document.getElementById("editMatchNote");
          const row = currentEditRowId != null ? getRowById(currentEditRowId) : null;
          const payload = buildAiCopyPayload(row) || (txtArea ? txtArea.value : "");
          if (!payload) return;

          navigator.clipboard.writeText(payload).then(() => {
              // Feedback visivo temporaneo
              const originalText = btnCopyReport.textContent;
              btnCopyReport.textContent = "✅ COPIATO!";
              btnCopyReport.style.backgroundColor = "#d4edda";
              setTimeout(() => {
                  btnCopyReport.textContent = originalText;
                  btnCopyReport.style.backgroundColor = "";
              }, 1500);
          }).catch(err => {
              console.error("Errore copia:", err);
              alert("Impossibile copiare automaticamente. Seleziona il testo e premi Ctrl+C.");
          });
      });
  }

  // ===============================================================
//   KPI ALLINEAMENTO PERIODI (ADE vs Gestionale)
// ===============================================================

function monthKeyFromDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function monthLabelFromKey(key) {
  const [y, m] = key.split("-");
  const mesi = [
    "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
    "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
  ];
  const idx = parseInt(m, 10) - 1;
  return `${mesi[idx]} ${y}`;
}

function computePeriodKpi(results) {
  const map = new Map();

  for (const r of results || []) {
    if (r.deleted) continue;

    // 🔹 GESTIONE SOLO_GEST: usa la data GEST invece della data ADE
    if (r.STATUS === "SOLO_GEST") {
      const g = r.GEST;
      if (g?.data) {
        const key = monthKeyFromDate(g.data);
        if (!map.has(key)) {
          map.set(key, {
            key,
            adeCount: 0,
            matchOk: 0,
            matchFix: 0,
            soloAde: 0,
            soloGest: 0
          });
        }
        const agg = map.get(key);
        agg.soloGest++;
      }
      continue; // salta il resto del loop per SOLO_GEST
    }

    // 🔹 Per tutti gli altri record, usa la data ADE
    const a = r.ADE;
    if (!a?.data) continue;

    // 🔹 fuori perimetro: fatture ADE senza P.IVA italiana (es. estero)
    const pivaDigits = (a.piva ? onlyDigits(a.piva) : "");
    const isItalian = pivaDigits.length === 11;

    const key = monthKeyFromDate(a.data);
    if (!map.has(key)) {
      map.set(key, {
        key,
        adeCount: 0,
        matchOk: 0,
        matchFix: 0,
        soloAde: 0,
        soloGest: 0
      });
    }

    const agg = map.get(key);
    agg.adeCount++;

    // MODIFICA: Ora includiamo MANUAL_MATCH nel conteggio dei successi (matchOk)
    if (r.STATUS === "MATCH_OK" || r.STATUS === "MANUAL_MATCH") {
      agg.matchOk++;
    } else if (r.STATUS === "MATCH_FIX") {
      agg.matchFix++;
    } else if (r.STATUS === "SOLO_ADE") {
      // 🔸 SOLO le SOLO_ADE con P.IVA italiana le consideriamo "buco vero"
      if (isItalian) {
        agg.soloAde++;
      }
    }
  }

  const arr = Array.from(map.values());
  arr.sort((a, b) => a.key.localeCompare(b.key));
  return arr;
}

/**
 * Factory function per creare oggetti status periodo
 * @param {string} icon - Icona emoji
 * @param {string} label - Etichetta descrittiva
 * @returns {Object} - { icon, label }
 */
function createPeriodStatus(icon, label) {
  return { icon, label };
}

function classifyPeriodRow(row) {
  const { adeCount, matchOk, matchFix, soloAde } = row;

  if (adeCount === 0)
    return createPeriodStatus("⚪", "Nessuna fattura ADE nel periodo");

  const matched = matchOk + matchFix;

  if (matched === 0)
    return createPeriodStatus("🔴", "Non allineato: nessun match con il gestionale");

  if (soloAde === 0 && matchOk > 0 && matchFix === 0)
    return createPeriodStatus("✅", "Allineato: tutti i match sono OK");

  if (soloAde === 0 && matchOk === 0 && matchFix > 0)
    return createPeriodStatus("🟡", "Allineato ma con soli MATCH DA CORREGGERE");

  if (soloAde > 0)
    return createPeriodStatus("🟠", `Parzialmente allineato: ${soloAde} fatture ADE senza match`);

  return createPeriodStatus("🟡", "Allineamento da verificare");
}

function renderPeriodKpi(results) {
  const tbody = document.getElementById("kpiPeriodTbody");
  const summaryEl = document.getElementById("kpiPeriodSummary");
  if (!tbody || !summaryEl) return;

  const data = computePeriodKpi(results);
  tbody.innerHTML = "";

  if (!data.length) {
    summaryEl.textContent = "Nessun dato disponibile. Esegui un confronto fatture.";
    return;
  }

  let ok = 0, warn = 0, bad = 0;

  // 🔹 qui calcoliamo anche l’ULTIMO MESE COMPLETAMENTE ALLINEATO
  let lastFullyAlignedKey = null;

  for (const row of data) {
    const status = classifyPeriodRow(row);

    // mese "pienamente allineato" = nessuna SOLO_ADE
    if (row.soloAde === 0) {
      lastFullyAlignedKey = row.key;
    }

    const tr = document.createElement("tr");
    tr.dataset.monthKey = row.key;
    tr.style.cursor = "pointer";
    tr.title = "Clicca per vedere le righe di questo mese nella riconciliazione";

    const tdMonth = document.createElement("td");
    tdMonth.textContent = monthLabelFromKey(row.key);
    tr.appendChild(tdMonth);

    const tdAde = document.createElement("td");
    tdAde.className = "mono";
    tdAde.textContent = row.adeCount;
    tr.appendChild(tdAde);

    const tdOk = document.createElement("td");
    tdOk.className = "mono";
    tdOk.textContent = row.matchOk;
    tr.appendChild(tdOk);

    const tdFix = document.createElement("td");
    tdFix.className = "mono";
    tdFix.textContent = row.matchFix;
    tr.appendChild(tdFix);

    const tdSoloAde = document.createElement("td");
    tdSoloAde.className = "mono";
    tdSoloAde.textContent = row.soloAde;
    tr.appendChild(tdSoloAde);

    const tdSoloGest = document.createElement("td");
    tdSoloGest.className = "mono";
    tdSoloGest.textContent = row.soloGest; // resterà quasi sempre 0 qui
    tr.appendChild(tdSoloGest);

    const tdStatus = document.createElement("td");
    tdStatus.textContent = `${status.icon} ${status.label}`;
    tr.appendChild(tdStatus);

    // 👉 click: filtra il mese in riconciliazione
    tr.addEventListener("click", () => {
      currentMonthFilter = row.key;
      const tabRecon = document.querySelector('.nav-tab[data-section="section-recon"]');
      if (tabRecon) tabRecon.click();

      currentFilter = "ALL";
      if (filterButtons && filterButtons.length) {
        filterButtons.forEach(btn => {
          btn.classList.toggle("active", btn.dataset.filter === "ALL");
        });
      }

      applyFilterAndRender();
    });

    tbody.appendChild(tr);

    if (status.icon === "✅") ok++;
    else if (status.icon === "🔴") bad++;
    else warn++;
  }

  // 🔹 testo riassuntivo "intelligente"
  let baseSummary =
    `Mesi analizzati: ${data.length} — ` +
    `✅ OK: ${ok}, 🟡/🟠 Da verificare: ${warn}, 🔴 Non allineati: ${bad}.`;

  if (lastFullyAlignedKey) {
    const lastLabel = monthLabelFromKey(lastFullyAlignedKey);
    baseSummary += ` Ultimo mese completamente allineato (nessuna fattura ADE mancante in perimetro): ${lastLabel}.`;
  }

  summaryEl.textContent = baseSummary;
}

  // ===============================================================
  // SEZIONE LIQUIDAZIONE IVA (ADE)
  // ===============================================================

  function getYearFromIso(iso) {
    if (!iso || typeof iso !== "string") return null;
    const m = iso.match(/^(\d{4})-\d{2}-\d{2}$/);
    return m ? parseInt(m[1], 10) : null;
  }

  function getVatIsoFromRecord(r, basis) {
    if (!r) return "";
    const b = String(basis || "RICEZIONE").toUpperCase();
    if (b === "EMISSIONE") return r.dataEmissIso || "";
    if (b === "AUTO") return r.vatDateIso || r.dataRicezIso || r.dataEmissIso || "";
    // RICEZIONE (default): ricezione -> fallback emissione
    return r.dataRicezIso || r.vatDateIso || r.dataEmissIso || "";
  }

  function detectVatTargetYear(records, basis) {
    const freq = new Map();
    for (const r of records || []) {
      const y = getYearFromIso(getVatIsoFromRecord(r, basis));
      if (!y) continue;
      freq.set(y, (freq.get(y) || 0) + 1);
    }
    if (freq.size === 0) return new Date().getFullYear();
    let bestYear = null;
    let bestCount = -1;
    for (const [y, c] of freq.entries()) {
      if (c > bestCount) { bestYear = y; bestCount = c; }
    }
    return bestYear;
  }

  function shouldIncludeInVatYear(r, targetYear, includePrevYear, basis) {
    const y = getYearFromIso(getVatIsoFromRecord(r, basis));
    if (!y) return false;
    if (y === targetYear) return true;
    if (includePrevYear && y === (targetYear - 1)) return true;
    return false;
  }

  function buildVatGroups(records, basis, includePrevYear) {
    const targetYear = detectVatTargetYear(records, basis);
    const groups = {};
    let grandImp = 0;
    let grandIva = 0;
    let grandTot = 0;
    let excludedNoDate = 0;
    let excludedOutOfYear = 0;

    (records || []).forEach(r => {
      const inScope = shouldIncludeInVatYear(r, targetYear, includePrevYear, basis);
      if (!inScope) { excludedOutOfYear++; return; }

      const iso = getVatIsoFromRecord(r, basis);
      if (!iso || iso.length < 7) { excludedNoDate++; return; }

      const key = iso.slice(0, 7); // YYYY-MM
      if (!groups[key]) groups[key] = { count: 0, imp: 0, iva: 0, tot: 0, label: monthLabelFromKey(key) };

      const imp = r.imp || 0;
      const iva = r.iva || 0;
      const tot = r.tot || 0;

      groups[key].count++;
      groups[key].imp += imp;
      groups[key].iva += iva;
      groups[key].tot += tot;

      grandImp += imp;
      grandIva += iva;
      grandTot += tot;
    });

    const keys = Object.keys(groups).sort((a, b) => a.localeCompare(b));

    return { groups, keys, grandImp, grandIva, grandTot, targetYear, excludedNoDate, excludedOutOfYear };
  }


  function renderVatLiquidation() {
    const tbody = document.getElementById("vatTbody");
    const elImp = document.getElementById("vatTotImp");
    const elIva = document.getElementById("vatTotIva");
    const elDoc = document.getElementById("vatTotDoc");

    if (!tbody) return;
    tbody.innerHTML = "";

    if (!lastAdeRecords || lastAdeRecords.length === 0) {
      tbody.innerHTML = "<tr><td colspan='5' style='text-align:center; padding:20px;'>Nessun dato ADE caricato.</td></tr>";
      if (elImp) elImp.textContent = "€ 0,00";
      if (elIva) elIva.textContent = "€ 0,00";
      if (elDoc) elDoc.textContent = "€ 0,00";
      return;
    }

    // UI state (persistito)
    const lsBasisKey = "contaibl-vat-date-basis";
    const lsPrevKey  = "contaibl-vat-include-prev-year";

    let basis = (localStorage.getItem(lsBasisKey) || "RICEZIONE").toUpperCase();
    const selBasis = document.getElementById("vatDateBasis");
    if (selBasis) {
      selBasis.value = basis;
      selBasis.onchange = () => {
        localStorage.setItem(lsBasisKey, selBasis.value);
        renderVatLiquidation();
      };
      basis = String(selBasis.value || basis).toUpperCase();
    }

    let includePrevYear = localStorage.getItem(lsPrevKey) === "true";
    const cb = document.getElementById("vatIncludePrevYear");
    if (cb) {
      cb.checked = includePrevYear;
      cb.onchange = () => {
        localStorage.setItem(lsPrevKey, cb.checked ? "true" : "false");
        renderVatLiquidation();
      };
      includePrevYear = cb.checked;
    }

    const vat = buildVatGroups(lastAdeRecords, basis, includePrevYear);

    // KPI (stesso perimetro della tabella)
    if (elImp) elImp.textContent = "€ " + formatNumberITDisplay(vat.grandImp);
    if (elIva) elIva.textContent = "€ " + formatNumberITDisplay(vat.grandIva);
    if (elDoc) elDoc.textContent = "€ " + formatNumberITDisplay(vat.grandTot);

    if (!vat.keys.length) {
      const msg = "Nessun record nel perimetro selezionato.";
      tbody.innerHTML = `<tr><td colspan='5' style='text-align:center; padding:20px;'>${msg}</td></tr>`;
      return;
    }

    vat.keys.forEach(key => {
      const g = vat.groups[key];
      const tr = document.createElement("tr");

      const tdP = document.createElement("td");
      tdP.textContent = g.label || monthLabelFromKey(key);

      const tdN = document.createElement("td");
      tdN.textContent = String(g.count || 0);

      const tdImp = document.createElement("td");
      tdImp.className = "num";
      tdImp.textContent = formatNumberITDisplay(g.imp || 0);

      const tdIva = document.createElement("td");
      tdIva.className = "num";
      tdIva.textContent = formatNumberITDisplay(g.iva || 0);

      const tdTot = document.createElement("td");
      tdTot.className = "num";
      tdTot.textContent = formatNumberITDisplay(g.tot || 0);

      tr.appendChild(tdP);
      tr.appendChild(tdN);
      tr.appendChild(tdImp);
      tr.appendChild(tdIva);
      tr.appendChild(tdTot);

      tbody.appendChild(tr);
    });
  }

  // Listener per aggiornare la tabella quando si clicca sul TAB
  const btnVatTab = document.querySelector('.nav-tab[data-section="section-vat"]');
  if (btnVatTab) {
    btnVatTab.addEventListener("click", () => {
       renderVatLiquidation();
    });
  }

  // Export CSV Liquidazione IVA
  const btnExportVat = document.getElementById("btnExportVat");
  if (btnExportVat) {
    btnExportVat.addEventListener("click", () => {
      if (!lastAdeRecords || !lastAdeRecords.length) {
        alert("Nessun dato da esportare.");
        return;
      }

      const lsBasisKey = "contaibl-vat-date-basis";
      const lsPrevKey  = "contaibl-vat-include-prev-year";

      let basis = (localStorage.getItem(lsBasisKey) || "RICEZIONE").toUpperCase();
      const selBasis = document.getElementById("vatDateBasis");
      if (selBasis) basis = String(selBasis.value || basis).toUpperCase();

      let includePrevYear = localStorage.getItem(lsPrevKey) === "true";
      const cb = document.getElementById("vatIncludePrevYear");
      if (cb) includePrevYear = cb.checked;

      const vat = buildVatGroups(lastAdeRecords, basis, includePrevYear);

      if (!vat.keys.length) {
        alert("Nessun record nel perimetro selezionato.");
        return;
      }

      // CSV
      let csv = "Periodo (Mese);N. Fatture;Imponibile;IVA;Totale Fattura\n";
      vat.keys.forEach(key => {
        const g = vat.groups[key];
        csv += `${monthLabelFromKey(key)};${g.count};${formatNumberITDisplay(g.imp)};${formatNumberITDisplay(g.iva)};${formatNumberITDisplay(g.tot)}\n`;
      });

      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const yearSuffix = vat.targetYear ? String(vat.targetYear) : "export";
      a.download = `liquidazione_iva_${yearSuffix}.csv`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });
  }

  // ---------- RESIZE COLONNE TABELLA (VERSIONE CORRETTA) ----------
  function createResizableTable(table) {
    if (!table) return;

    // Decidi se mantenere table-layout: fixed permanentemente per questa tabella
    const keepFixed = (table.id === 'resultTable') || table.classList.contains('keep-fixed');
    // Salva valore corrente per ripristinarlo dopo il drag (se non vogliamo mantenere fixed)
    let prevTableLayout = table.style.tableLayout || window.getComputedStyle(table).tableLayout;
    if (keepFixed) prevTableLayout = 'fixed';

    // Determina il numero effettivo di colonne espandendo i colspan nelle righe di thead
    const theadRows = Array.from(table.querySelectorAll('thead tr'));
    if (!theadRows.length) return;

    const lastHeaderCells = Array.from(table.querySelectorAll('thead tr:last-child th'));
    // calcola il numero totale di colonne (espandendo colspan) prendendo il massimo tra le righe
    const totalCols = theadRows.reduce((max, row) => {
      const cells = Array.from(row.querySelectorAll('th'));
      const count = cells.reduce((s, c) => s + (parseInt(c.getAttribute('colspan') || '1', 10) || 1), 0);
      return Math.max(max, count);
    }, 0);
    if (!totalCols) return;

    // Crea o aggiorna <colgroup> per controllare le larghezze delle colonne in modo affidabile
    let colgroup = table.querySelector('colgroup');
    if (!colgroup) {
      colgroup = document.createElement('colgroup');
      for (let i = 0; i < totalCols; i++) {
        colgroup.appendChild(document.createElement('col'));
      }
      table.insertBefore(colgroup, table.firstChild);
    } else {
      // assicurati che il numero di <col> corrisponda al totale espanso
      const existing = colgroup.querySelectorAll('col').length;
      if (existing < totalCols) {
        for (let i = existing; i < totalCols; i++) colgroup.appendChild(document.createElement('col'));
      } else if (existing > totalCols) {
        // rimuovi quelli in eccesso
        const colsToRemove = Array.from(colgroup.querySelectorAll('col')).slice(totalCols);
        colsToRemove.forEach(c => c.parentNode.removeChild(c));
      }
    }

    const cols = Array.from(colgroup.querySelectorAll('col'));

    // Assicuriamoci che il numero di <col> non superi il numero totale di colonne calcolate
    if (cols.length > totalCols) {
      for (let i = totalCols; i < cols.length; i++) cols[i].parentNode.removeChild(cols[i]);
    }

    // Per misurazioni più affidabili, forziamo table-layout: fixed; se keepFixed lo manteniamo
    try { table.style.tableLayout = 'fixed'; } catch (err) {}

    // inizializza larghezze se non presenti (misuriamo con getBoundingClientRect per precisione)
    // Distribuiamo le larghezze dei th dell'ultima riga sui <col>, rispettando eventuali colspan
    try {
      // array di larghezze per ciascuna col (index 0..totalCols-1)
      const colWidths = new Array(totalCols).fill(0);
      let writeIdx = 0;
      const thToColIndex = [];
      lastHeaderCells.forEach(th => {
        const span = Math.max(1, parseInt(th.getAttribute('colspan') || '1', 10) || 1);
        const rect = th.getBoundingClientRect();
        const w = Math.max(30, Math.round(rect.width)) || 100;
        // registra l'indice di col corrispondente al primo slot di questa th
        thToColIndex.push(writeIdx);
        // assegna larghezza divisa uniformemente tra le colonne coperte dal colspan
        const per = Math.round(w / span);
        for (let k = 0; k < span; k++) {
          if (writeIdx + k < totalCols) colWidths[writeIdx + k] = per;
        }
        writeIdx += span;
      });

      // fallback: se alcune colonne risultano 0, assegna un minimo
      for (let i = 0; i < totalCols; i++) if (!colWidths[i] || colWidths[i] < 30) colWidths[i] = 80;

      // Applica le larghezze ai col corrispondenti
      cols.forEach((col, i) => {
        const w = colWidths[i] || 80;
        if (!col.style.width) col.style.width = Math.round(w) + 'px';
      });

      // Aggiorna larghezza delle celle di tutte le righe di thead rispettando colspan
      try {
        theadRows.forEach(row => {
          let ci = 0;
          const cells = Array.from(row.querySelectorAll('th'));
          cells.forEach(th => {
            const span = Math.max(1, parseInt(th.getAttribute('colspan') || '1', 10) || 1);
            let sum = 0;
            for (let j = ci; j < Math.min(ci + span, cols.length); j++) {
              const val = cols[j].style.width || window.getComputedStyle(cols[j]).width || '0px';
              sum += parseFloat(String(val).replace('px','')) || 0;
            }
            if (sum > 0) {
              th.style.width = Math.round(sum) + 'px';
              th.style.minWidth = Math.round(sum) + 'px';
            }
            ci += span;
          });
        });
      } catch (e) { /* ignore */ }
      // salva la mappatura per l'uso successivo (resizer -> col index)
      table.__thToColIndex = thToColIndex;
    } catch (e) {
      // Se qualcosa va storto, fallback semplice: imposta larghezza basata su offsetWidth dei th monocolonna
      lastHeaderCells.forEach((th, idx) => {
        const col = cols[idx];
        const w = th.offsetWidth || 100;
        if (col && !col.style.width) col.style.width = w + 'px';
        if (th) { th.style.width = (col && col.style.width) || (w + 'px'); th.style.minWidth = th.style.width; }
      });
    }

    // Calcola la somma delle larghezze e forza la larghezza complessiva della tabella
    try {
      const totalWidth = Array.from(colgroup.querySelectorAll('col')).reduce((sum, c) => {
        const val = c.style.width || window.getComputedStyle(c).width || '0px';
        return sum + (parseFloat(String(val).replace('px','')) || 0);
      }, 0);
      if (totalWidth > 0) table.style.width = Math.round(totalWidth) + 'px';
    } catch (e) { /* ignore */ }

    // Ripristina il table-layout originale subito dopo l'inizializzazione, a meno che non vogliamo mantenerlo fixed
    try {
      if (!keepFixed) {
        if (prevTableLayout) table.style.tableLayout = prevTableLayout;
        else table.style.removeProperty('table-layout');
      } else {
        table.style.tableLayout = 'fixed';
      }
    } catch (err) { /* ignore */ }

    // --- DEBUG OVERLAY (opzionale) ---
    const debugEnabled = (window.__contaibl_resizer_debug === true) || localStorage.getItem('contaibl-resizer-debug') === 'true';
    if (debugEnabled) {
      try {
        // assicurati che il wrapper sia posizionato relativamente per overlay assoluto
        const wrapper = table.parentElement || table;
        if (wrapper && getComputedStyle(wrapper).position === 'static') wrapper.style.position = 'relative';

        let dbg = wrapper.querySelector('.resizer-debug-overlay');
        if (!dbg) {
          dbg = document.createElement('div');
          dbg.className = 'resizer-debug-overlay';
          dbg.style.position = 'absolute';
          dbg.style.left = '0';
          dbg.style.top = '0';
          dbg.style.pointerEvents = 'none';
          dbg.style.zIndex = 9999;
          wrapper.appendChild(dbg);
        }

        // pulisce eventuali badge precedenti
        dbg.innerHTML = '';
        const tableRect = table.getBoundingClientRect();
        const debugBadges = [];
        lastHeaderCells.forEach((th, i) => {
          const rect = th.getBoundingClientRect();
          const badge = document.createElement('div');
          badge.className = 'resizer-debug-badge';
          badge.style.position = 'absolute';
          badge.style.left = (rect.left - tableRect.left) + 'px';
          badge.style.top = (rect.top - tableRect.top) + 'px';
          badge.style.pointerEvents = 'none';
          badge.style.zIndex = 10000;
          badge.textContent = Math.round(rect.width) + 'px';
          dbg.appendChild(badge);
          debugBadges.push(badge);
        });

        // memorizza per aggiornamenti durante il resize
        table.__resizerDebug = { container: dbg, badges: debugBadges };
        const firstBodyRow = table.querySelector('tbody tr');
        const firstBodyCells = firstBodyRow ? firstBodyRow.children.length : 0;
        console.log('Resizer DEBUG: attivato', { prevTableLayout, colCount: totalCols, headerCells: lastHeaderCells.length, firstBodyCells });
      } catch (e) { console.warn('Resizer DEBUG init failed', e); }
    }

    // Attacca resizer a ciascun th dell'ultima riga dell'intestazione
    lastHeaderCells.forEach((th, thIdx) => {
      if (th.querySelector('.resizer')) return; // evita duplicati
      const resizer = document.createElement('div');
      resizer.className = 'resizer';
      // assicurati che il th sia posizionato relativamente
      th.style.position = th.style.position || 'relative';
      th.appendChild(resizer);

      let isResizing = false;
      let startX = 0;
      let startWidth = 0;

      const mouseMoveHandler = function(e) {
        if (!isResizing) return;
        const dx = e.clientX - startX;
        const newWidth = Math.max(30, startWidth + dx);
        // trova l'indice di col corrispondente a questo th
        const mapping = table.__thToColIndex || [];
        const colIndex = mapping[thIdx] != null ? mapping[thIdx] : thIdx;
        // Applica la larghezza al <col> corrispondente (più affidabile)
        cols[colIndex].style.width = newWidth + 'px';
        // Mantieni il th coerente (solo se monocolonna)
        const span = Math.max(1, parseInt(th.getAttribute('colspan') || '1', 10) || 1);
        if (span === 1) { th.style.width = newWidth + 'px'; th.style.minWidth = newWidth + 'px'; }
        // Aggiorna larghezza complessiva della tabella per mantenere allineamento
        try {
          const total = Array.from(colgroup.querySelectorAll('col')).reduce((s, c) => s + (parseFloat((c.style.width||window.getComputedStyle(c).width||'0').replace('px',''))||0), 0);
          if (total > 0) table.style.width = Math.round(total) + 'px';
        } catch (err) { /* ignore */ }
        // Aggiorna badge debug se presente
        try {
          const dbg = table.__resizerDebug;
          if (dbg && dbg.badges && dbg.badges[thIdx]) {
            const badge = dbg.badges[thIdx];
            badge.textContent = newWidth + 'px';
            // riposiziona il badge in caso di spostamento
            const tr = th.getBoundingClientRect();
            const tb = table.getBoundingClientRect();
            badge.style.left = (tr.left - tb.left) + 'px';
            badge.style.top = (tr.top - tb.top) + 'px';
          }
        } catch (err) { /* ignore */ }
      };

      const mouseUpHandler = function() {
        if (!isResizing) return;
        isResizing = false;
        document.body.style.cursor = '';
        resizer.classList.remove('resizing');
        document.removeEventListener('mousemove', mouseMoveHandler);
        document.removeEventListener('mouseup', mouseUpHandler);
        document.body.style.userSelect = '';
        // Ripristina il table-layout originale (hybrid behavior)
        try {
          // Se abbiamo deciso di mantenere fixed per questa tabella, non ripristinare
          if (!keepFixed) {
            if (prevTableLayout) table.style.tableLayout = prevTableLayout;
            else table.style.removeProperty('table-layout');
          } else {
            table.style.tableLayout = 'fixed';
          }
        } catch (err) { /* ignore */ }
        // Log diagnostico finale
        try {
          const colWidths = Array.from(colgroup.querySelectorAll('col')).map(c => c.style.width || window.getComputedStyle(c).width);
          console.log('Resizer mouseup - final widths:', colWidths);
          const dbg = table.__resizerDebug;
          if (dbg && dbg.badges) {
            dbg.badges.forEach((b, i) => {
              // Aggiorna testo con larghezza effettiva
              b.textContent = colWidths[i] || b.textContent;
            });
          }
        } catch (e) { /* ignore */ }
      };

      resizer.addEventListener('mousedown', function(e) {
        e.preventDefault();
        e.stopPropagation();
        isResizing = true;
        startX = e.clientX;
        // leggi larghezza corrente del <col>
        const mapping = table.__thToColIndex || [];
        const startCol = mapping[thIdx] != null ? mapping[thIdx] : thIdx;
        startWidth = parseInt(window.getComputedStyle(cols[startCol]).width, 10) || cols[startCol].offsetWidth || th.offsetWidth;
        // forziamo la tabella in px per evitare conflitti con %
        try { table.style.width = table.offsetWidth + 'px'; } catch (err) {}
        // Durante il drag forziamo table-layout: fixed per avere un resize prevedibile,
        // poi lo ripristiniamo nel mouseUpHandler
        try { table.style.tableLayout = 'fixed'; } catch (err) {}
        // Log diagnostico iniziale
        try {
          const thRects = lastHeaderCells.map(t => t.getBoundingClientRect());
          const colRects = Array.from(colgroup.querySelectorAll('col')).map(c => ({ w: c.style.width || window.getComputedStyle(c).width }));
          console.log('Resizer mousedown thIdx=' + thIdx, { startWidth, thRects, colRects, prevTableLayout });
        } catch (e) { /* ignore */ }
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
        resizer.classList.add('resizing');
        document.addEventListener('mousemove', mouseMoveHandler);
        document.addEventListener('mouseup', mouseUpHandler);
      });
    });
  }

  // Setup stepper navigation (click to scroll)
  setupStepperNavigation();

});