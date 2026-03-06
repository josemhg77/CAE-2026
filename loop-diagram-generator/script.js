// LOOP CAD 2026 - Multi-sheet, full DXF export
document.addEventListener('DOMContentLoaded', () => {
    init();
});

const canvas = document.getElementById('loopCanvas');
const ctx = canvas.getContext('2d');

const CONFIG = {
    canvasWidth: 1800,
    canvasHeight: 3000, // Increased to accommodate 350px rows
    colors: {
        bg: '#1a1a1a',
        text: '#ffffff',
        line: '#ffffff',
        cable: '#ffffff',
        wireText: '#ffffff',
        shield: '#cccccc',
        grid: '#333333'
    },
    fonts: {
        header: 'bold 16px "Consolas", monospace',
        subHeader: 'bold 12px "Consolas", monospace',
        tag: 'bold 14px "Consolas", monospace',
        label: '12px "Consolas", monospace',
        small: '12px "Consolas", monospace',
        mono: '12px "Consolas", monospace'
    }
};

// ── Multi-sheet state ──────────────────────────────────────────────────────────
let sheets = [];          // sheets[sheetIdx] = array of 8 loops
let currentSheetIndex = 0;
let currentLoopIndex = 0;
let cableNamingMode = 'TAG';
let projectName = '';
let canvasScale = 1.0;

let duplicateTags = new Set();
let duplicateCables = new Set();

function getCurrentLoops() { return sheets[currentSheetIndex]; }

function createEmptyLoop(id) {
    const startTerm = id * 2 + 1;
    return {
        id,
        instTag: `LT-${100 + id}`,
        instType: 'Nivel',
        instTerms: ['+', '-', '3', '4', '5', '6', '7', '8'],
        instCategory: 'STD', // STD, 3W, VALVE, MOTOR
        signalNames: ['SIG1', 'SIG2', 'SIG3', 'SIG4'],

        cable1Tag: `WLT-${104 + id}`,
        cable1Section: '2x1.0mm²',
        cable1Shield: true,
        cable1Armor: false,
        cable1Len: 0,

        jbEnabled: true,
        marshEnabled: true,

        jbTag: 'JB-1',
        jbModule: 'JX1',
        jbTerms: [`${startTerm}`, `${startTerm + 1}`, '3', '4', '5', '6', '7', '8'],
        jbTermSh: 'PE',

        cable2Tag: `JB-1`,
        cable2Section: '12x1.0mm²',
        cable2Shield: true,
        cable2Armor: true,
        cable2Len: 0,

        marshTag: 'BME-101',
        marshTerms: [`${startTerm}`, `${startTerm + 1}`, '3', '4', '5', '6', '7', '8'],
        marshTermInSh: 'PE',
        marshModule: 'X12',

        cableSysTag: 'W-SYS',
        cableSysSection: '2x0.5mm²',
        cableSysShield: true,
        cableSysArmor: false,
        cableSysLen: 0,

        plcTag: 'SCP',
        plcModule: 'X1',
        plcTerms: ['1', '2', '3', '4', '5', '6', '7', '8'],
        plcTermPe: 'PE',
        plcDetails: `Señal: LT-${104 + id}`
    };
}

function createBlankLoop(id) {
    return {
        id,
        instTag: '', instType: 'Nivel', instTerms: ['', '', '', '', '', '', '', ''], instCategory: 'STD',
        signalNames: ['', '', '', ''],
        cable1Tag: '', cable1Section: '', cable1Shield: false, cable1Armor: false, cable1Len: 0,
        jbEnabled: false, marshEnabled: false,
        jbTag: '', jbModule: '', jbTerms: ['', '', '', '', '', '', '', ''], jbTermSh: '',
        cable2Tag: '', cable2Section: '', cable2Shield: false, cable2Armor: false, cable2Len: 0,
        marshTag: '', marshTerms: ['', '', '', '', '', '', '', ''], marshTermInSh: '', marshModule: '',
        cableSysTag: '', cableSysSection: '', cableSysShield: false, cableSysArmor: false, cableSysLen: 0,
        plcTag: '', plcModule: '', plcTerms: ['', '', '', '', '', '', '', ''], plcTermPe: '', plcDetails: ''
    };
}

function createEmptySheet() {
    const s = [];
    for (let i = 0; i < 8; i++) s.push(createEmptyLoop(i)); // Default populated loops
    return s;
}

function createBlankSheet() {
    const s = [];
    for (let i = 0; i < 8; i++) s.push(createBlankLoop(i)); // Truly empty loops
    return s;
}

// ── Init ───────────────────────────────────────────────────────────────────────
function init() {
    canvas.width = CONFIG.canvasWidth;
    canvas.height = CONFIG.canvasHeight;

    const savedData = localStorage.getItem('loopCAD_autosave');
    if (savedData) {
        try {
            const d = JSON.parse(savedData);
            if (d.sheets) {
                sheets = d.sheets;
            } else if (d.loops) {
                sheets = [d.loops];
            }
            if (d.cableNamingMode) {
                cableNamingMode = d.cableNamingMode;
                const cnEl = document.getElementById('cableNaming');
                if (cnEl) cnEl.value = cableNamingMode;
            }
            if (d.projectName) {
                projectName = d.projectName;
                document.getElementById('projectName').value = projectName;
            }
            if (d.colors) {
                CONFIG.colors = { ...CONFIG.colors, ...d.colors };
                const cct = document.getElementById('colorCableTag');
                if (cct) cct.value = CONFIG.colors.cable;
                const cwt = document.getElementById('colorWireText');
                if (cwt) cwt.value = CONFIG.colors.wireText;
                const cs = document.getElementById('colorShield');
                if (cs) cs.value = CONFIG.colors.shield;
            }
        } catch (e) {
            console.warn('Error loading autosave', e);
            sheets = [createEmptySheet()];
        }
    } else {
        sheets = [createEmptySheet()];
    }

    buildLoopButtons();
    updateSheetUI();

    const inputs = document.querySelectorAll('.equipment-bar input, .equipment-bar textarea, .equipment-bar select, .properties-panel input, .properties-panel textarea, .properties-panel select');
    inputs.forEach(el => {
        const handler = (e) => {
            updateCurrentLoopFromUI(e);
            renderScene();
        };
        el.addEventListener('input', handler);
        el.addEventListener('change', handler);
    });

    document.getElementById('jbEnabled').addEventListener('change', updateUIState);
    document.getElementById('marshEnabled').addEventListener('change', updateUIState);
    document.getElementById('instCategory').addEventListener('change', updateUIState);

    document.getElementById('btnSave').addEventListener('click', saveProject);
    document.getElementById('fileLoad').addEventListener('change', loadProject);
    document.getElementById('btnExportPDF').addEventListener('click', exportPDF);
    document.getElementById('btnExportDXF').addEventListener('click', () => {
        const fn = (projectName || 'proyecto_lazos') + '.dxf';
        exportDXF(fn);
    });

    document.getElementById('btnExportCSV').addEventListener('click', exportCSVTemplate);
    document.getElementById('fileLoadCSV').addEventListener('change', importDataFromCSV);




    // Project Name specifically
    // Color pickers and project name handled by general inputs above if they have the class
    // But let's double check CONFIG updates
    document.getElementById('projectName').addEventListener('input', (e) => { projectName = e.target.value; autoSaveToBrowser(); });

    if (document.getElementById('cableNaming')) {
        document.getElementById('cableNaming').addEventListener('change', (e) => { cableNamingMode = e.target.value; autoSaveToBrowser(); });
    }

    canvas.addEventListener('wheel', (e) => {
        if (e.ctrlKey || e.metaKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? 0.9 : 1.1;
            canvasScale *= delta;
            canvasScale = Math.max(0.2, Math.min(5, canvasScale));
            renderScene();
        }
    }, { passive: false });

    document.getElementById('btnNew').addEventListener('click', () => {
        if (confirm('¿Desea crear un NUEVO PROYECTO? Se borrarán todas las hojas.')) {
            sheets = [createEmptySheet()];
            currentSheetIndex = 0;
            currentLoopIndex = 0;
            projectName = '';
            document.getElementById('projectName').value = '';
            buildLoopButtons();
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
        }
    });

    // Add a specific listener for a separate "New Sheet" button if we add one
    // or keep btnNextSheet as the primary way.

    document.getElementById('btnPrevSheet').addEventListener('click', () => {
        if (currentSheetIndex > 0) {
            updateCurrentLoopFromUI();
            currentSheetIndex--;
            currentLoopIndex = 0;
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
        }
    });

    document.getElementById('btnNextSheet').addEventListener('click', () => {
        updateCurrentLoopFromUI();
        if (currentSheetIndex < sheets.length - 1) {
            currentSheetIndex++;
            currentLoopIndex = 0;
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
        }
    });

    document.getElementById('btnAddSheet').addEventListener('click', () => {
        updateCurrentLoopFromUI();
        sheets.push(createBlankSheet());
        currentSheetIndex = sheets.length - 1;
        currentLoopIndex = 0;
        updateSheetUI();
        loadLoopToUI(0);
        renderScene();
    });

    document.getElementById('btnDeleteSheet').addEventListener('click', () => {
        if (sheets.length <= 1) {
            alert('No se puede borrar la única hoja del proyecto.');
            return;
        }
        if (confirm(`¿Está seguro de que desea BORRAR COMPLETAMENTE la Hoja ${currentSheetIndex + 1}? Esta acción no se puede deshacer.`)) {
            sheets.splice(currentSheetIndex, 1);
            if (currentSheetIndex >= sheets.length) {
                currentSheetIndex = sheets.length - 1;
            }
            currentLoopIndex = 0;
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
            autoSaveToBrowser();
        }
    });

    document.getElementById('btnClearLoop').addEventListener('click', clearLoopData);
    document.getElementById('btnDuplicateLoop').addEventListener('click', duplicateLoopData);

    const btnDiag = document.getElementById('btnViewDiagram');
    const btnCables = document.getElementById('btnViewCableList');
    const diagView = document.getElementById('diagramView');
    const cableView = document.getElementById('cableListView');

    btnDiag.addEventListener('click', () => {
        updateCurrentLoopFromUI();
        btnDiag.classList.add('active-view');
        btnCables.classList.remove('active-view');
        diagView.classList.add('active');
        cableView.classList.remove('active');
        renderScene();
    });

    btnCables.addEventListener('click', () => {
        updateCurrentLoopFromUI();
        btnCables.classList.add('active-view');
        btnDiag.classList.remove('active-view');
        cableView.classList.add('active');
        diagView.classList.remove('active');
        showCableList();
    });

    document.getElementById('btnExportListCSV_New').addEventListener('click', exportCableList);

    loadLoopToUI(0);
    renderScene();
}

function buildLoopButtons() {
    const btnContainer = document.getElementById('loopButtons');
    btnContainer.innerHTML = '';
    for (let i = 0; i < 8; i++) {
        const btn = document.createElement('button');
        btn.textContent = `${i + 1}`;
        btn.className = (i === 0) ? 'loop-btn active' : 'loop-btn';
        btn.onclick = () => selectLoop(i);
        btnContainer.appendChild(btn);
    }
}

function updateSheetUI() {
    const el = document.getElementById('sheetLabel');
    if (el) el.textContent = `Hoja ${currentSheetIndex + 1} / ${sheets.length}`;
    // highlight active loop button
    const btns = document.querySelectorAll('.loop-btn');
    const loops = getCurrentLoops();
    btns.forEach((b, i) => {
        let stateClass = 'empty';
        if (loops[i] && loops[i].instTag && loops[i].instTag.trim() !== '') {
            stateClass = 'populated';
        }

        b.className = (i === currentLoopIndex) ? 'loop-btn active' : `loop-btn ${stateClass}`;
    });
    const titleEl = document.getElementById('currentLoopTitle');
    if (titleEl) {
        titleEl.textContent = `PROPIEDADES LAZO ${currentLoopIndex + 1} — HOJA ${currentSheetIndex + 1}`;
    }

    const activeLabelEl = document.getElementById('activeLoopLabel');
    if (activeLabelEl) {
        activeLabelEl.textContent = `Lazo: ${currentLoopIndex + 1} | Hoja: ${currentSheetIndex + 1}`;
    }
}

// ── Loop selection ─────────────────────────────────────────────────────────────
function selectLoop(index) {
    updateCurrentLoopFromUI();
    currentLoopIndex = index;
    updateSheetUI();
    loadLoopToUI(index);
}

function loadLoopToUI(index) {
    const data = getCurrentLoops()[index];
    if (!data) return;
    const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v ?? ''; };
    const setChk = (id, v) => { const el = document.getElementById(id); if (el) el.checked = !!v; };

    setVal('projectName', projectName);

    setVal('instTag', data.instTag);
    setVal('instType', data.instType);
    setVal('instCategory', data.instCategory || 'STD');
    const iterms = data.instTerms || ['', '', '', '', '', '', '', ''];
    for (let i = 0; i < 8; i++) setVal(`instTerm${i + 1}`, iterms[i]);

    const signalNames = data.signalNames || ['', '', '', ''];
    for (let i = 0; i < 4; i++) setVal(`sigName${i + 1}`, signalNames[i]);

    setVal('cable1Tag', data.cable1Tag);
    setVal('cable1Section', data.cable1Section);
    setVal('cable1Len', data.cable1Len || 0);
    setChk('cable1Shield', data.cable1Shield);
    setChk('cable1Armor', data.cable1Armor);

    setChk('jbEnabled', data.jbEnabled !== false);
    setVal('jbTag', data.jbTag);
    setVal('jbModule', data.jbModule);
    setVal('jbTermPos', (data.jbTerms && data.jbTerms[0]) || data.jbTermPos || '1');
    setVal('jbTermNeg', (data.jbTerms && data.jbTerms[1]) || data.jbTermNeg || '2');
    setVal('jbTermSh', data.jbTermSh || 'PE');

    setVal('cable2Tag', data.cable2Tag);
    setVal('cable2Section', data.cable2Section);
    setVal('cable2Len', data.cable2Len || 0);
    setChk('cable2Shield', data.cable2Shield);
    setChk('cable2Armor', data.cable2Armor);

    setChk('marshEnabled', data.marshEnabled !== false);
    setVal('marshTag', data.marshTag);
    setVal('marshModule', data.marshModule);
    setVal('marshTermInPos', (data.marshTerms && data.marshTerms[0]) || data.marshTermInPos || '1');
    setVal('marshTermInNeg', (data.marshTerms && data.marshTerms[1]) || data.marshTermInNeg || '2');
    setVal('marshTermInSh', data.marshTermInSh || 'PE');

    setVal('cableSysTag', data.cableSysTag);
    setVal('cableSysSection', data.cableSysSection);
    setVal('cableSysLen', data.cableSysLen || 0);
    setChk('cableSysShield', data.cableSysShield);
    setChk('cableSysArmor', data.cableSysArmor);

    setVal('plcTag', data.plcTag);
    for (let i = 0; i < 8; i++) {
        setVal(`plcTerm${i + 1}`, (data.plcTerms && data.plcTerms[i]) || '');
    }
    setVal('plcTermPe', data.plcTermPe || 'PE');
    setVal('plcDetails', data.plcDetails);

    updateUIState();
}

function updateUIState() {
    const loop = getCurrentLoops()[currentLoopIndex];
    const category = document.getElementById('instCategory').value;
    const jbEn = document.getElementById('jbEnabled').checked;
    const marshEn = document.getElementById('marshEnabled').checked;

    // Show/Hide instrument terminals
    const tCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    for (let i = 0; i < 8; i++) {
        const el = document.getElementById(`instTerm${i + 1}`);
        if (el) el.style.display = i < tCount ? 'block' : 'none';
    }

    // Show/Hide signal descriptions based on category
    const sigNamesContainer = document.querySelector('.inner-row-grid.four-cols')?.parentElement?.parentElement;
    if (sigNamesContainer) {
        sigNamesContainer.style.display = 'flex'; // Always show now

        const s1 = document.getElementById('sigName1');
        const s2 = document.getElementById('sigName2');
        const s3 = document.getElementById('sigName3');
        const s4 = document.getElementById('sigName4');

        const sigsToShow = (category === 'VALVE') ? 3 : (category === 'MOTOR' ? 4 : 1);

        if (s1) s1.style.display = sigsToShow >= 1 ? 'block' : 'none';
        if (s2) s2.style.display = sigsToShow >= 2 ? 'block' : 'none';
        if (s3) s3.style.display = sigsToShow >= 3 ? 'block' : 'none';
        if (s4) s4.style.display = sigsToShow >= 4 ? 'block' : 'none';

        // Show/Hide PLC terminals (Phase 11)
        for (let i = 0; i < 8; i++) {
            const el = document.getElementById(`plcTerm${i + 1}`);
            if (el) el.style.display = i < tCount ? 'block' : 'none';
        }
    }

    const jbFields = document.getElementById('jbFields');
    const marshFields = document.getElementById('marshFields');
    const cable2Group = document.getElementById('cable2Group');
    const cableSysGroup = document.getElementById('cableSysGroup');

    if (jbFields) {
        jbFields.querySelectorAll('input:not([type="checkbox"]), textarea, select').forEach(el => el.disabled = !jbEn);
        jbFields.style.opacity = jbEn ? '1' : '0.5';
    }
    if (marshFields) {
        marshFields.querySelectorAll('input:not([type="checkbox"]), textarea, select').forEach(el => el.disabled = !marshEn);
        marshFields.style.opacity = marshEn ? '1' : '0.5';
    }

    if (cable2Group) cable2Group.style.visibility = jbEn ? 'visible' : 'hidden';
    if (cableSysGroup) cableSysGroup.style.visibility = marshEn ? 'visible' : 'hidden';
}

function clearLoopData() {
    if (confirm('¿Borrar todos los datos del lazo actual?')) {
        getCurrentLoops()[currentLoopIndex] = createBlankLoop(currentLoopIndex);
        loadLoopToUI(currentLoopIndex);
        renderScene();
        autoSaveToBrowser();
    }
}

function duplicateLoopData() {
    if (currentLoopIndex < 7) {
        const currentData = JSON.parse(JSON.stringify(getCurrentLoops()[currentLoopIndex]));
        currentData.id = currentLoopIndex + 1;
        getCurrentLoops()[currentLoopIndex + 1] = currentData;

        // Auto-select the next loop to show it was copied
        selectLoop(currentLoopIndex + 1);
        renderScene();
        autoSaveToBrowser();
    } else {
        alert('No se puede duplicar en el último lazo de la hoja.');
    }
}

function updateCurrentLoopFromUI(e) {
    const loop = getCurrentLoops()[currentLoopIndex];
    if (!loop) return;

    const getVal = (id) => { const el = document.getElementById(id); return el ? el.value : ''; };
    const getChk = (id) => { const el = document.getElementById(id); return el ? el.checked : false; };

    // Auto-update logic for Instrument Tag
    if (e && e.target.id === 'instTag') {
        const newTag = e.target.value;
        // Auto-update Cable 1 tag: W + InstrumentTag (as requested previously)
        loop.cable1Tag = `W${newTag}`;
        const c1el = document.getElementById('cable1Tag');
        if (c1el) c1el.value = loop.cable1Tag;

        // Update Signal in PLC Details
        if (loop.plcDetails && loop.plcDetails.includes('Señal:')) {
            loop.plcDetails = loop.plcDetails.replace(/Señal:.*?\n/, `Señal: ${newTag}\n`);
            const pdel = document.getElementById('plcDetails');
            if (pdel) pdel.value = loop.plcDetails;
        }
    }

    loop.instTag = getVal('instTag');
    loop.instType = getVal('instType');
    loop.instCategory = getVal('instCategory');

    // Sync instTerms
    if (!loop.instTerms) loop.instTerms = ['', '', '', '', '', '', '', ''];
    for (let i = 0; i < 8; i++) {
        loop.instTerms[i] = getVal(`instTerm${i + 1}`);
    }

    // Sync signalNames
    if (!loop.signalNames) loop.signalNames = ['', '', '', ''];
    for (let i = 0; i < 4; i++) {
        loop.signalNames[i] = getVal(`sigName${i + 1}`);
    }

    loop.cable1Tag = getVal('cable1Tag');
    loop.cable1Section = getVal('cable1Section');
    loop.cable1Len = parseFloat(getVal('cable1Len')) || 0;
    loop.cable1Shield = getChk('cable1Shield');
    loop.cable1Armor = getChk('cable1Armor');

    loop.jbEnabled = getChk('jbEnabled');
    loop.jbTag = getVal('jbTag');
    loop.jbModule = getVal('jbModule');
    loop.cable2Tag = getVal('cable2Tag');
    loop.cable2Section = getVal('cable2Section');
    loop.cable2Len = parseFloat(getVal('cable2Len')) || 0;
    loop.cable2Shield = getChk('cable2Shield');
    loop.cable2Armor = getChk('cable2Armor');

    loop.marshEnabled = getChk('marshEnabled');
    loop.marshTag = getVal('marshTag');
    loop.marshModule = getVal('marshModule');

    loop.plcTag = getVal('plcTag');
    loop.plcTermPe = getVal('plcTermPe');
    loop.plcDetails = getVal('plcDetails');

    // Save PLC terms (Phase 11)
    loop.plcTerms = [];
    for (let i = 1; i <= 8; i++) {
        loop.plcTerms.push(getVal(`plcTerm${i}`));
    }
    loop.marshTag = getVal('marshTag');
    loop.marshModule = getVal('marshModule');
    loop.cableSysTag = getVal('cableSysTag');
    loop.cableSysSection = getVal('cableSysSection');
    loop.cableSysLen = parseFloat(getVal('cableSysLen')) || 0;
    loop.cableSysShield = getChk('cableSysShield');
    loop.cableSysArmor = getChk('cableSysArmor');

    loop.plcTag = getVal('plcTag');
    loop.plcModule = getVal('plcModule');
    loop.plcTermPe = getVal('plcTermPe');
    loop.plcDetails = getVal('plcDetails');

    if (loop.jbEnabled) {
        loop.jbTerms[0] = getVal('jbTermPos') || '1';
        loop.jbTerms[1] = getVal('jbTermNeg') || '2';
    }
    if (loop.marshEnabled) {
        loop.marshTerms[0] = getVal('marshTermInPos') || '1';
        loop.marshTerms[1] = getVal('marshTermInNeg') || '2';
    }

    findDuplicates();
    markDuplicatesUI();
    autoSaveToBrowser();
}

function findDuplicates() {
    const instCount = {};
    const cableCountMap = {};
    duplicateTags.clear();
    duplicateCables.clear();

    sheets.forEach(sheet => {
        sheet.forEach(loop => {
            // Check Instrument Tags
            if (loop.instTag && loop.instTag.trim() !== '') {
                const t = loop.instTag.trim().toUpperCase();
                instCount[t] = (instCount[t] || 0) + 1;
            }
            // Check ONLY Instrument Cable (Cable 1)
            if (cableNamingMode === 'TAG') {
                const cNames = buildCableNames(loop, 1, () => { });
                // We only care about cNames.c1 for duplicates as per user request
                const c1 = cNames.c1;
                if (c1 && c1.trim() !== '' && c1 !== 'N/A') {
                    const tc = c1.trim().toUpperCase();
                    cableCountMap[tc] = (cableCountMap[tc] || 0) + 1;
                }
            }
        });
    });

    for (const t in instCount) { if (instCount[t] > 1) duplicateTags.add(t); }
    for (const c in cableCountMap) { if (cableCountMap[c] > 1) duplicateCables.add(c); }
}

function markDuplicatesUI() {
    const it = document.getElementById('instTag');
    if (it) {
        const val = it.value.trim().toUpperCase();
        it.style.color = duplicateTags.has(val) ? '#ff5555' : '';
        it.style.fontWeight = duplicateTags.has(val) ? 'bold' : '';
    }
}

// ── AutoSave ─────────────────────────────────────────────────────────────────
function autoSaveToBrowser() {
    const data = JSON.stringify({ sheets, cableNamingMode, projectName, colors: CONFIG.colors });
    localStorage.setItem('loopCAD_autosave', data);
}

// ── Render ─────────────────────────────────────────────────────────────────────
function renderScene() {
    const wrapper = document.querySelector('.canvas-wrapper');
    const bar = document.querySelector('.equipment-bar');

    const scaledW = CONFIG.canvasWidth * canvasScale;
    const scaledH = CONFIG.canvasHeight * canvasScale;

    if (wrapper) {
        wrapper.style.width = scaledW + 'px';
        wrapper.style.height = scaledH + 'px';
        wrapper.style.position = 'relative';
        wrapper.style.display = 'block';
    }

    if (canvas) {
        canvas.style.width = scaledW + 'px';
        canvas.style.height = scaledH + 'px';
        canvas.style.position = 'relative';
        canvas.style.top = '0';
    }

    // Reset bar styles
    if (bar) {
        bar.style.width = '1800px';
        bar.style.height = '450px';
        bar.style.transform = 'none';
        bar.style.position = 'relative';
        bar.style.top = '0';
        bar.style.left = '0';
    }

    findDuplicates(); // v2.0.5: Call validation
    markDuplicatesUI(); // v2.0.5: Call validation UI update

    ctx.fillStyle = CONFIG.colors.bg;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    drawGrid();
    drawHeaderTable();

    const startY = 440; // Subido de 480 para aprovechar mejor el espacio superior
    const rowH = 350; // Increased to 350px to ensure no overlaps even with 8 wires

    let cableCount = getGlobalCableCount(currentSheetIndex);
    const loops = getCurrentLoops();
    const cableWireMap = {}; // Phase 8: track wire index per cable tag

    loops.forEach((loop, i) => {
        const y = startY + (i * rowH);

        ctx.fillStyle = CONFIG.colors.text;
        ctx.font = "bold 24px Arial";
        ctx.textAlign = 'center';
        ctx.fillText(i + 1, 40, y);

        if (loop.instTag && loop.instTag.trim() !== '') {
            drawRow(loop, y, buildCableNames(loop, cableCount, (n) => { cableCount = n; }), cableWireMap);
        }
    });

    markDuplicatesUI();
}

function buildCableNames(loop, cableCount, setCableCount) {
    let dest1 = loop.plcTag || 'PLC';
    if (loop.marshEnabled) dest1 = loop.marshTag || 'MRSH';
    if (loop.jbEnabled) dest1 = loop.jbTag || 'JB';

    let dest2 = loop.plcTag || 'PLC';
    if (loop.marshEnabled) dest2 = loop.marshTag || 'MRSH';

    const dest3 = loop.plcTag || 'PLC';

    let c1 = '', c2 = '', c3 = '';
    let cnt = cableCount;
    const pad = (n) => n.toString().padStart(3, '0');

    if (cableNamingMode === 'AUTO') {
        c1 = `W-${pad(cnt++)}/${dest1}`;
        if (loop.jbEnabled) c2 = `W-${pad(cnt++)}/${dest2}`;
        if (loop.marshEnabled) c3 = `W-${pad(cnt++)}/${dest3}`;
    } else {
        c1 = `W${loop.instTag || 'TAG'}/${dest1}`;
        if (loop.jbEnabled) c2 = `${loop.jbTag || 'JB'}/${dest2}`;
        if (loop.marshEnabled) c3 = `${loop.marshTag || 'MRSH'}/${dest3}`;
    }

    if (setCableCount) setCableCount(cnt);
    return { c1, c2, c3 };
}

function getGlobalCableCount(uptoSheetIndex) {
    let count = 1;
    for (let s = 0; s < uptoSheetIndex; s++) {
        sheets[s].forEach(loop => {
            const hasData = loop.instTag && loop.instTag.trim() !== '';
            if (!hasData) return;
            // Count cables for this loop:
            // Always at least 1 cable (INST→first_stop or INST→PLC)
            // +1 if JB enabled (JB→MARSH or JB→PLC)
            // +1 if MARSH enabled (MARSH→PLC)
            let cablesInLoop = 1;
            if (loop.jbEnabled) cablesInLoop++;
            if (loop.marshEnabled) cablesInLoop++;
            count += cablesInLoop;
        });
    }
    return count;
}

function drawGrid() { /* Optional */ }

function drawHeaderTable() {
    const W = CONFIG.canvasWidth;
    const H = 80;
    ctx.lineWidth = 2;
    ctx.strokeStyle = CONFIG.colors.line;
    ctx.strokeRect(0, 0, W, H);

    ctx.beginPath(); ctx.moveTo(0, 40); ctx.lineTo(W, 40); ctx.stroke();
    const splitX = W * 0.45;
    ctx.beginPath(); ctx.moveTo(splitX, 0); ctx.lineTo(splitX, H); ctx.stroke();

    ctx.fillStyle = CONFIG.colors.text;
    ctx.font = CONFIG.fonts.header;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`CAMPO — Hoja ${currentSheetIndex + 1}`, splitX / 2, 20);
    ctx.fillText("EDIFICIO DE CONTROL", splitX + (W - splitX) / 2, 20);

    // Project Name in the middle bottom of header or similar?
    // Let's put it in a new row or top left?
    // User asked for it as a parameter, common to show in header or footer.
    // Let's add it to the header area.
    ctx.font = CONFIG.fonts.subHeader;
    ctx.textAlign = 'left';
    ctx.fillText(`PROYECTO: ${projectName}`, 10, 20);

    const cols = getCols();
    ctx.font = CONFIG.fonts.subHeader;
    ctx.textAlign = 'center';
    ctx.fillText("INSTRUMENTO", cols.inst, 60);
    ctx.fillText("CAJA DE AGRUPACIÓN", cols.jb, 60);
    ctx.fillText("MARSHALLING", cols.marsh, 60);
    ctx.fillText("ARMARIO DE CONTROL", cols.plc, 60);
}

function getCols() {
    const W = CONFIG.canvasWidth;
    return {
        inst: 250,
        jb: 700,
        marsh: 1150,
        plc: 1600
    };
}

function drawRow(loop, y, cableNames, wireMap) {
    const instTag = (loop.instTag || '').trim().toUpperCase();
    const isDup = duplicateTags.has(instTag);

    // Instrument
    const oldFill = ctx.fillStyle;
    if (isDup) ctx.fillStyle = '#ff5555';
    drawInstrument(loop, getCols().inst, y);
    ctx.fillStyle = oldFill;

    if (loop.jbEnabled) drawJB(loop, getCols().jb, y);
    if (loop.marshEnabled) drawMarsh(loop, getCols().marsh, y);
    drawPLC(loop, getCols().plc, y);

    // Cables
    let lc = 'INST';
    if (loop.jbEnabled) {
        const c1isDup = duplicateCables.has(cableNames.c1.trim().toUpperCase());
        drawCable(getCols().inst, getCols().jb, y,
            { tag: cableNames.c1, sec: loop.cable1Section, sh: loop.cable1Shield, arm: loop.cable1Armor, isDup: c1isDup },
            'INST-JB', loop, wireMap);
        lc = 'JB';
    }
    if (loop.marshEnabled) {
        const s = lc === 'INST' ? getCols().inst : getCols().jb;
        const tag = lc === 'INST' ? cableNames.c1 : cableNames.c2;
        const ctype = lc === 'INST' ? 'INST-MARSH' : 'JB-MARSH';
        const sec = lc === 'INST' ? loop.cable1Section : loop.cable2Section;
        const sh = lc === 'INST' ? loop.cable1Shield : loop.cable2Shield;
        const arm = lc === 'INST' ? loop.cable1Armor : loop.cable2Armor;
        const cDup = duplicateCables.has(tag.trim().toUpperCase());
        drawCable(s, getCols().marsh, y, { tag, sec, sh, arm, isDup: cDup }, ctype, loop, wireMap);
        lc = 'MARSH';
    }

    const s3 = lc === 'MARSH' ? getCols().marsh : (lc === 'JB' ? getCols().jb : getCols().inst);
    const t3 = lc === 'MARSH' ? cableNames.c3 : (lc === 'JB' ? cableNames.c2 : cableNames.c1);
    const ty3 = lc === 'MARSH' ? 'MARSH-PLC' : (lc === 'JB' ? 'JB-PLC' : 'INST-PLC');
    const sec3 = lc === 'MARSH' ? loop.cableSysSection : (lc === 'JB' ? loop.cable2Section : loop.cable1Section);
    const sh3 = lc === 'MARSH' ? loop.cableSysShield : (lc === 'JB' ? loop.cable2Shield : loop.cable1Shield);
    const arm3 = lc === 'MARSH' ? loop.cableSysArmor : (lc === 'JB' ? loop.cable2Armor : loop.cable1Armor);
    const c3Dup = duplicateCables.has(t3.trim().toUpperCase());
    drawCable(s3, getCols().plc, y, { tag: t3, sec: sec3, sh: sh3, arm: arm3, isDup: c3Dup }, ty3, loop, wireMap);
}

// ── Component drawing ──────────────────────────────────────────────────────────
function drawInstrument(loop, x, y) {
    const r = 30;
    ctx.beginPath();
    ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.lineWidth = 2;
    ctx.strokeStyle = CONFIG.colors.line;
    ctx.fillStyle = CONFIG.colors.text;
    ctx.stroke();

    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.font = CONFIG.fonts.tag;

    const parts = loop.instTag.split('-');
    if (parts.length > 1) {
        ctx.beginPath();
        ctx.moveTo(x - r, y); ctx.lineTo(x + r, y); ctx.stroke();
        // Lowered text slightly (y-5, y+11)
        ctx.fillText(parts[0], x, y - 5);
        ctx.fillText(parts[1], x, y + 11);
    } else {
        ctx.fillText(loop.instTag, x, y + 3);
    }

    // Borneras apiladas verticalmente, centradas en y
    const termSize = 15;
    const tx = x + r;
    const category = loop.instCategory || 'STD';

    if (!loop.instTerms) loop.instTerms = [loop.instTermPos || '+', loop.instTermNeg || '-', '3', '4', '5', '6', '7', '8'];

    const termCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    const startTermY = y - ((termCount - 1) * 20) / 2;

    for (let i = 0; i < termCount; i++) {
        const ty = startTermY + (i * 20);
        drawTermSquare(tx, ty - 7.5, loop.instTerms[i] || (i + 1).toString());
    }

    // Tipo de instrumento — bajado
    ctx.font = CONFIG.fonts.small;
    ctx.textAlign = 'center';
    ctx.fillText(loop.instType, x, y + r + 22);
}

function drawTermSquare(x, y, txt, width = 15) {
    const h = 15;
    const w = width;
    ctx.lineWidth = 1;
    ctx.strokeStyle = CONFIG.colors.line;
    ctx.strokeRect(x, y, w, h);
    ctx.fillStyle = CONFIG.colors.text;
    ctx.font = CONFIG.fonts.small;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(txt, x + w / 2, y + h / 2);
}

function drawJB(loop, x, y) {
    const category = loop.instCategory || 'STD';
    const termCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    const h = termCount * 20 + 115; // Increased to include ground symbol inside
    const w = 60;
    const top = y - h / 2;

    ctx.lineWidth = 2;
    ctx.strokeStyle = CONFIG.colors.line;
    ctx.strokeRect(x - w / 2, top, w, h);

    ctx.textAlign = 'center';
    ctx.font = CONFIG.fonts.tag;
    ctx.fillText(loop.jbTag, x, top - 15);

    ctx.fillStyle = CONFIG.colors.text;
    ctx.font = CONFIG.fonts.small;
    ctx.fillText(loop.jbModule || '', x, top + 15);
    ctx.beginPath();
    ctx.moveTo(x - w / 2, top + 25); ctx.lineTo(x + w / 2, top + 25);
    ctx.stroke();
    if (!loop.jbTerms) {
        loop.jbTerms = [loop.jbTermPos || '1', loop.jbTermNeg || '2', '3', '4', '5', '6', '7', '8'];
    }

    const startTermY = y - ((termCount - 1) * 20) / 2;
    for (let i = 0; i < termCount; i++) {
        const ty = startTermY + (i * 20);
        drawTermSquare(x - 20, ty - 7.5, loop.jbTerms[i] || (i + 1).toString(), 40);
    }

    const peLineY = startTermY + (termCount * 20);
    const peBoxY = peLineY + 5;
    drawTermSquare(x - 20, peBoxY, loop.jbTermSh || 'PE', 40);
    drawGroundSymbol(x, peBoxY + 15);
}

function drawMarsh(loop, x, y) {
    const category = loop.instCategory || 'STD';
    const termCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    const h = termCount * 20 + 120; // Increased to include ground symbol inside
    const w = 60;
    const top = y - h / 2;

    ctx.lineWidth = 2;
    ctx.strokeStyle = CONFIG.colors.line;
    ctx.strokeRect(x - w / 2, top, w, h);

    ctx.textAlign = 'center';
    ctx.font = CONFIG.fonts.tag;
    ctx.fillText(loop.marshTag, x, top - 15);

    ctx.fillStyle = CONFIG.colors.text;
    ctx.font = CONFIG.fonts.small;
    ctx.fillText(loop.marshModule || '', x, top + 15);
    ctx.beginPath();
    ctx.moveTo(x - w / 2, top + 25); ctx.lineTo(x + w / 2, top + 25);
    ctx.stroke();

    if (!loop.marshTerms) {
        loop.marshTerms = [loop.marshTermInPos || '1', loop.marshTermInNeg || '2', '3', '4', '5', '6', '7', '8'];
    }

    const startTermY = y - ((termCount - 1) * 20) / 2;
    for (let i = 0; i < termCount; i++) {
        const ty = startTermY + (i * 20);
        drawTermSquare(x - 20, ty - 7.5, loop.marshTerms[i] || (i + 1).toString(), 40);
    }

    const peLineY = startTermY + (termCount * 20);
    const peBoxY = peLineY + 5;
    drawTermSquare(x - 20, peBoxY, loop.marshTermInSh || 'PE', 40);
    drawGroundSymbol(x, peBoxY + 15);
}

function drawPLC(loop, x, y) {
    const category = loop.instCategory || 'STD';
    const w = 280;
    let h = 200;
    if (category === 'VALVE' || category === 'MOTOR') h = 260;

    ctx.strokeStyle = CONFIG.colors.line;
    ctx.lineWidth = 2;
    ctx.strokeRect(x - w / 2, y - h / 2, w, h);

    ctx.font = CONFIG.fonts.tag;
    ctx.textAlign = 'center';
    ctx.fillText(loop.plcTag, x, y - h / 2 - 10);

    const termX = x - w / 2 + 10;
    const termW = 35;
    const termCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    const startY = y - ((termCount - 1) * 20) / 2;


    if (!loop.plcTerms) {
        loop.plcTerms = [loop.plcTermPos || '1', loop.plcTermNeg || '2', '3', '4', '5', '6', '7', '8'];
    }

    for (let i = 0; i < termCount; i++) {
        const ty = startY + (i * 20) - 7.5;
        drawTermSquare(termX, ty, loop.plcTerms[i] || (i + 1).toString(), termW);
    }

    const peY = startY + (termCount * 20);
    drawTermSquare(termX, peY - 7.5, loop.plcTermPe || 'PE', termW);
    drawGroundSymbol(termX + termW / 2, peY + 7.5);

    // Right side: TAG label + signal names for VALVE/MOTOR, or plcDetails for others
    const labelX = x - w / 2 + 55;
    const signalNames = loop.signalNames || [];
    const sigGroups = (category === 'VALVE') ? 3 : (category === 'MOTOR' ? 4 : 1);

    ctx.textAlign = 'left';
    ctx.fillStyle = CONFIG.colors.text;

    // Position TAG slightly higher to leave room for SIG1 in STD/3W
    const tagYOffset = (sigGroups === 1) ? 22 : 30;
    ctx.font = 'bold 11px "Consolas", monospace';
    ctx.fillText('TAG: ' + (loop.instTag || ''), labelX, y - h / 2 + tagYOffset);

    ctx.font = '10px "Consolas", monospace';
    // One label per signal pair
    for (let g = 0; g < sigGroups; g++) {
        const pairTop = startY + (g * 2) * 20;
        const pairBot = startY + (g * 2 + 1) * 20;
        let pairCy = (pairTop + pairBot) / 2 + 4;

        const label = signalNames[g] || `SIG${g + 1}`;
        ctx.fillText(label, labelX, pairCy);
    }

    if (sigGroups === 1) {
        // Show detail lines below SIG1
        ctx.font = '11px "Consolas", monospace';
        const lines = (loop.plcDetails || '').split('\n').filter(l => l.trim() !== '');
        let ly = startY + 40; // Below SIG1
        lines.forEach(l => {
            ctx.fillText(l, labelX, ly);
            ly += 16;
        });
    }
}

// ── Cable drawing ──────────────────────────────────────────────────────────────
function drawCable(x1, x2, y, opts, type, loop, wireMap) {
    const tag = opts.tag || '';
    const sec = opts.sec || '';
    const sh = opts.sh || false;
    const arm = opts.arm || false;
    const isDup = opts.isDup || false;

    let sx = x1, dx = x2;
    if (type === 'INST-JB') { sx = x1 + 45; dx = x2 - 30; }
    else if (type === 'INST-MARSH') { sx = x1 + 45; dx = x2 - 30; }
    else if (type === 'INST-PLC') { sx = x1 + 45; dx = x2 - 145; }
    else if (type === 'JB-MARSH') { sx = x1 + 30; dx = x2 - 30; }
    else if (type === 'JB-PLC') { sx = x1 + 30; dx = x2 - 145; }
    else if (type === 'MARSH-PLC') { sx = x1 + 30; dx = x2 - 145; }

    const gap = type.endsWith('PLC') ? 45 : 22;
    const csx = sx + (type.startsWith('INST') ? 22 : gap);
    const cex = dx - gap;

    ctx.lineWidth = 1;
    ctx.strokeStyle = CONFIG.colors.cable; // Use cable color for main line

    // Main line
    ctx.beginPath(); ctx.moveTo(csx, y); ctx.lineTo(cex, y); ctx.stroke();

    // Fan out
    const category = loop.instCategory || 'STD';
    const tCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
    const startY_local = y - ((tCount - 1) * 20) / 2;

    const baseTag = tag.trim().toUpperCase();
    let wireStart = 1;
    if (wireMap) {
        if (wireMap[baseTag]) wireStart = wireMap[baseTag] + 1;
    }

    for (let i = 0; i < tCount; i++) {
        const wireNum = wireStart + i;
        const wy = startY_local + (i * 20);
        ctx.beginPath(); ctx.moveTo(sx, wy); ctx.lineTo(csx, wy); ctx.lineTo(csx, y + (i < tCount / 2 ? -2 : 2)); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(cex, y + (i < tCount / 2 ? -2 : 2)); ctx.lineTo(cex, wy);
        // Wires finish exactly at dx
        ctx.lineTo(dx, wy);
        ctx.stroke();
        ctx.fillStyle = CONFIG.colors.wireText; ctx.font = "8px Arial"; ctx.textAlign = 'left';
        ctx.fillText(wireNum, csx + 5, wy - 3); ctx.textAlign = 'right';
        ctx.fillText(wireNum, cex - 5, wy - 3);
    }

    if (wireMap) {
        wireMap[baseTag] = wireStart + tCount - 1;
    }

    if (sh) {
        const shieldRY = 20;
        ctx.beginPath(); ctx.fillStyle = CONFIG.colors.bg; ctx.strokeStyle = CONFIG.colors.shield;
        ctx.ellipse((csx + cex) / 2, y, 7, shieldRY, 0, 0, Math.PI * 2); ctx.fill(); ctx.stroke();

        // Target side connection (to component entrance)
        if (type.endsWith('JB') || type.endsWith('MARSH') || type.endsWith('PLC')) {
            const tTermCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
            const tStartTermY = y - ((tTermCount - 1) * 20) / 2;
            let sey = tStartTermY + (tTermCount * 20);

            // PE Border logic (Do NOT enter the box)
            let targetX = dx;
            if (type.endsWith('JB') || type.endsWith('MARSH')) {
                sey += 12.5; // Center of PE box vertically
                targetX = x2 - 20; // Edge of the 40w box
            } else if (type.endsWith('PLC')) {
                targetX = x2 - 140 + 10; // Left edge of PLC PE term (termX)
            }

            ctx.save();
            ctx.strokeStyle = CONFIG.colors.shield;
            ctx.lineWidth = 1;
            ctx.setLineDash([5, 4]);
            ctx.beginPath(); ctx.moveTo(cex, y); ctx.lineTo(cex, sey);
            ctx.lineTo(targetX, sey);
            ctx.stroke();
            ctx.setLineDash([]);
            ctx.restore();
        }

        // Source side connection (from component exit - right side of boxes)
        if (type.startsWith('JB') || type.startsWith('MARSH')) {
            const sTermCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
            const sStartTermY = y - ((sTermCount - 1) * 20) / 2;
            let ssy = sStartTermY + (sTermCount * 20) + 12.5; // Center of PE box vertically

            ctx.save();
            ctx.strokeStyle = CONFIG.colors.shield;
            ctx.lineWidth = 1;
            ctx.setLineDash([5, 4]);
            ctx.beginPath(); ctx.moveTo(csx, y); ctx.lineTo(csx, ssy); ctx.lineTo(x1 + 20, ssy); ctx.stroke();
            ctx.restore();
        }
    }

    ctx.fillStyle = isDup ? '#ff5555' : CONFIG.colors.text;
    ctx.font = isDup ? "bold 12px Arial" : "11px Arial"; ctx.textAlign = 'center';
    ctx.fillText(tag, (csx + cex) / 2, y - 40);
    ctx.fillStyle = CONFIG.colors.text; ctx.font = "10px Arial";
    ctx.fillText(sec, (csx + cex) / 2, y + 80);
    if (arm) ctx.fillText("(ARM)", (csx + cex) / 2, y + 92);
}

// ── Ground symbol ──────────────────────────────────────────────────────────────
function drawGroundSymbol(x, y) {
    // y = bottom edge of PE terminal (connection point)
    const stemLen = 10;
    const symY = y + stemLen;

    ctx.lineWidth = 1.5;
    ctx.strokeStyle = CONFIG.colors.shield;

    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x, symY);
    ctx.stroke();

    ctx.beginPath();
    ctx.moveTo(x - 8, symY); ctx.lineTo(x + 8, symY);
    ctx.moveTo(x - 5, symY + 3); ctx.lineTo(x + 5, symY + 3);
    ctx.moveTo(x - 2, symY + 6); ctx.lineTo(x + 2, symY + 6);
    ctx.stroke();
}

// ── Save / Load ────────────────────────────────────────────────────────────────
function saveProject() {
    autoSaveToBrowser();
    const data = JSON.stringify({ sheets, cableNamingMode, projectName, colors: CONFIG.colors }, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (projectName || 'proyecto_lazos') + '.json';
    a.click();
}

function loadProject(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const d = JSON.parse(e.target.result);
            // Support old format (single loops array) and new (sheets)
            if (d.sheets) {
                sheets = d.sheets;
            } else if (d.loops) {
                sheets = [d.loops];
            }
            if (d.cableNamingMode) cableNamingMode = d.cableNamingMode;
            if (d.projectName) {
                projectName = d.projectName;
                document.getElementById('projectName').value = projectName;
            }
            if (d.colors) {
                CONFIG.colors = { ...CONFIG.colors, ...d.colors };
                document.getElementById('colorCableTag').value = CONFIG.colors.cable;
                document.getElementById('colorWireText').value = CONFIG.colors.wireText;
                document.getElementById('colorShield').value = CONFIG.colors.shield;
            }
            currentSheetIndex = 0;
            currentLoopIndex = 0;
            buildLoopButtons();
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
        } catch (err) { alert('Archivo inválido'); }
    };
    reader.readAsText(file);
}

// ── ZIP Project Backup ────────────────────────────────────────────────────────
async function descargarZipProyecto() {
    if (typeof JSZip === 'undefined') {
        return alert('La librería JSZip no está cargada.');
    }
    const zip = new JSZip();
    try {
        // We include all essential project files
        const files = ['index.html', 'script.js', 'style.css'];
        for (const file of files) {
            // Append timestamp to prevent cached versions and ensure we get the latest saved code
            const resp = await fetch(file + '?t=' + Date.now());
            if (!resp.ok) throw new Error(`No se pudo leer ${file}: ${resp.statusText}`);
            const text = await resp.text();
            zip.file(file, text);
        }

        // Also include the current project state specifically as a JSON
        const projectData = JSON.stringify({
            sheets,
            cableNamingMode,
            projectName,
            colors: CONFIG.colors
        }, null, 2);
        zip.file('project_data_backup.json', projectData);

        const content = await zip.generateAsync({ type: 'blob' });

        // Generate a nice timestamp for the filename
        const now = new Date();
        const datePart = now.getFullYear() +
            String(now.getMonth() + 1).padStart(2, '0') +
            String(now.getDate()).padStart(2, '0');
        const timePart = String(now.getHours()).padStart(2, '0') +
            String(now.getMinutes()).padStart(2, '0');

        // Clean project name for valid filename
        const cleanName = (projectName || 'proyecto_lazos').replace(/[^a-z0-9]/gi, '_').toLowerCase();

        const a = document.createElement('a');
        a.href = URL.createObjectURL(content);
        a.download = `${cleanName}_backup_${datePart}_${timePart}.zip`;
        a.click();
    } catch (err) {
        console.error('ZIP Error:', err);
        alert('Error al generar el Backup ZIP: ' + err.message + '\n\nIMPORTANTE: Esta función requiere un servidor local para leer los archivos (ej: abrir con VS Code Live Server o http-server). Si abres el archivo directamente desde el explorador de Windows, el navegador bloqueará la lectura del código por seguridad.');
    }
}

// ── PDF Export ─────────────────────────────────────────────────────────────────
async function exportPDF() {
    if (!window.jspdf) return alert('Falta la librería PDF');
    const { jsPDF } = window.jspdf;
    // A3 Portrait (297x420mm) fits 8 loops (1400x2400 aspect ratio) much better
    const pdf = new jsPDF('p', 'mm', 'a3');
    const oldColors = { ...CONFIG.colors };
    const oldSheet = currentSheetIndex;

    // Force B&W for export
    CONFIG.colors = { bg: '#ffffff', text: '#000000', line: '#000000', cable: '#000000', wireText: '#000000', shield: '#000000', grid: '#ffffff' };

    for (let i = 0; i < sheets.length; i++) {
        currentSheetIndex = i;
        renderScene();
        await new Promise(r => setTimeout(r, 150)); // Wait for render
        const img = canvas.toDataURL('image/png');
        if (i > 0) pdf.addPage('a3', 'p');

        // A3 Portrait: 297x420mm. 
        // Logic: 400mm height (10mm margin top/bottom), calculate width based on aspect ratio
        const imgH = 400;
        const imgW = imgH * (canvas.width / canvas.height); // 400 * (1400/2400) = ~233mm
        const xPos = (297 - imgW) / 2; // Center horizontally

        pdf.addImage(img, 'PNG', xPos, 10, imgW, imgH);
    }

    pdf.save((projectName || 'diagramas_lazo') + '.pdf');
    currentSheetIndex = oldSheet;
    CONFIG.colors = oldColors;
    renderScene();
}

// ── DXF Export (all sheets) ────────────────────────────────────────────────────
function exportDXF(filename = 'loop_sheet.dxf') {
    const sheetH = CONFIG.canvasHeight;
    const TOTAL_H = sheets.length * sheetH;
    const W = CONFIG.canvasWidth;
    const cols = getCols();
    const fy = (v) => TOTAL_H - v; // DXF Y-coordinates are typically inverted

    // ── DXF R12 minimal format (broadest compatibility) ──────────────────────────
    let dxf = `0\nSECTION\n2\nENTITIES\n`;

    // Per-sheet fy: flip Y within each sheet so first row is at top in CAD
    const sheetFy = (v, offY) => (sheetH - (v - offY));

    const DXF_COLOR = 7; // white (appears black on white bg in most CAD apps)

    // These helpers are created with closure per sheet (offY passed in)
    const makeDxfHelpers = (offY) => {
        const fy = (v) => sheetFy(v, offY);
        const line = (x1, y1, x2, y2, color = DXF_COLOR) => {
            dxf += `0\nLINE\n8\n0\n62\n${color}\n10\n${x1.toFixed(2)}\n20\n${fy(y1).toFixed(2)}\n11\n${x2.toFixed(2)}\n21\n${fy(y2).toFixed(2)}\n`;
        };
        const dashedLine = (x1, y1, x2, y2, color = DXF_COLOR) => {
            const ddx = x2 - x1, ddy = y2 - y1;
            const len = Math.sqrt(ddx * ddx + ddy * ddy);
            if (len === 0) return;
            const dashLen = 8, gapLen = 5, total = dashLen + gapLen;
            const steps = Math.floor(len / total);
            const nx = ddx / len, ny = ddy / len;
            for (let i = 0; i <= steps; i++) {
                const sx = x1 + nx * i * total, sy = y1 + ny * i * total;
                const ex = Math.min(1, (i * total + dashLen) / len);
                const ex2 = x1 + ddx * ex, ey2 = y1 + ddy * ex;
                dxf += `0\nLINE\n8\n0\n62\n${color}\n10\n${sx.toFixed(2)}\n20\n${fy(sy).toFixed(2)}\n11\n${ex2.toFixed(2)}\n21\n${fy(ey2).toFixed(2)}\n`;
            }
        };
        const circle = (x, y, r, color = DXF_COLOR) => {
            dxf += `0\nCIRCLE\n8\n0\n62\n${color}\n10\n${x.toFixed(2)}\n20\n${fy(y).toFixed(2)}\n40\n${r.toFixed(2)}\n`;
        };
        const ellipse = (x, y, rx, ry, color = DXF_COLOR) => {
            const steps = 64; // High resolution for perfect ellipse look
            for (let i = 0; i < steps; i++) {
                const a1 = (i / steps) * Math.PI * 2, a2 = ((i + 1) / steps) * Math.PI * 2;
                line(x + rx * Math.cos(a1), y + ry * Math.sin(a1), x + rx * Math.cos(a2), y + ry * Math.sin(a2), color);
            }
        };
        const text = (x, y, txt, h = 10, align = 'center', color = DXF_COLOR) => {
            if (!txt) return;
            const clean = String(txt).replace(/[\r\n]/g, ' ').replace(/²/g, '2'); // Replace mm² with mm2
            const xs = x.toFixed(2), ys = fy(y).toFixed(2);
            dxf += `0\nTEXT\n8\n0\n62\n${color}\n10\n${xs}\n20\n${ys}\n40\n${h.toFixed(2)}\n1\n${clean}\n`;
            if (align === 'center') {
                dxf += `72\n1\n11\n${xs}\n21\n${ys}\n`;
            } else if (align === 'right') {
                dxf += `72\n2\n11\n${xs}\n21\n${ys}\n`;
            }
        };
        const drawSq = (x, y, tt, w = 15) => {
            line(x, y, x + w, y);
            line(x + w, y, x + w, y + 15);
            line(x + w, y + 15, x, y + 15);
            line(x, y + 15, x, y);
            text(x + w / 2, y + 11, tt, 6, 'center');
        };
        const ground = (x, y) => {
            line(x, y, x, y + 10);
            line(x - 8, y + 10, x + 8, y + 10);
            line(x - 5, y + 13, x + 5, y + 13);
            line(x - 2, y + 16, x + 2, y + 16);
        };
        return { line, dashedLine, circle, ellipse, text, drawSq, ground };
    };

    const cableWireMap = {}; // Phase 8: track wire index per cable tag

    sheets.forEach((sheetLoops, sIdx) => {
        const offY = sIdx * sheetH;
        const { line, dashedLine, circle, ellipse, text, drawSq, ground } = makeDxfHelpers(offY);
        const baseY = offY; // local alias for clarity

        // Header Table (Y positions relative to sheet top = offY)
        const hTop = offY, hMid = offY + 40, hBot = offY + 80;
        line(0, hTop, W, hTop); line(0, hBot, W, hBot); line(0, hTop, 0, offY + sheetH); line(W, hTop, W, offY + sheetH); line(0, hMid, W, hMid);
        const splitX = W * 0.45; line(splitX, hTop, splitX, hBot);
        text(10, hTop + 25, `PROYECTO: ${projectName}`, 10, 'left');
        text(splitX / 2 - 80, hTop + 25, `CAMPO - Hoja ${sIdx + 1}`, 12, 'center');
        text(splitX + 20, hTop + 25, 'EDIFICIO DE CONTROL', 12, 'left');
        text(cols.inst, hBot - 10, 'INSTRUMENTO', 10, 'center');
        text(cols.jb, hBot - 10, 'CAJA DE AGRUPACION', 10, 'center');
        text(cols.marsh, hBot - 10, 'MARSHALING', 10, 'center');
        text(cols.plc, hBot - 10, 'ARMARIO DE CONTROL', 10, 'center');

        let cCount = getGlobalCableCount(sIdx);
        sheetLoops.forEach((loop, i) => {
            const y = offY + 240 + (i * 350);
            line(0, y + 175, W, y + 175); text(15, y + 5, (i + 1).toString(), 20);
            if (!loop.instTag || loop.instTag.trim() === '') return;
            const category = loop.instCategory || 'STD', tCount = category === 'STD' ? 2 : (category === '3W' ? 3 : (category === 'VALVE' ? 6 : 8));
            const cNames = buildCableNames(loop, cCount, (n) => { cCount = n; });
            const ix = cols.inst, jx = cols.jb, mx = cols.marsh, px = cols.plc;

            // Instrument
            circle(ix, y, 30);
            const p = loop.instTag.split('-');
            if (p.length > 1) {
                line(ix - 30, y, ix + 30, y);
                text(ix, y - 5, p[0], 8, 'center');
                text(ix, y + 13, p[1], 8, 'center');
            } else {
                text(ix, y + 5, loop.instTag, 8, 'center');
            }
            text(ix, y + 45, loop.instType, 8, 'center');

            const itx = ix + 40; // Matching Canvas offset
            line(ix + 30, y - 5, itx, y - 5);
            line(ix + 30, y + 5, itx, y + 5);

            const sTermYI = y - ((tCount - 1) * 20) / 2;
            for (let j = 0; j < tCount; j++) {
                const ty = sTermYI + (j * 20);
                drawSq(itx, ty - 7.5, loop.instTerms ? loop.instTerms[j] : (j + 1).toString(), 15);
            }

            // JB
            if (loop.jbEnabled) {
                const jh = tCount * 20 + 115;
                const jw = 60;
                const jTop = y - jh / 2;
                line(jx - jw / 2, jTop, jx + jw / 2, jTop); line(jx + jw / 2, jTop, jx + jw / 2, jTop + jh);
                line(jx + jw / 2, jTop + jh, jx - jw / 2, jTop + jh); line(jx - jw / 2, jTop + jh, jx - jw / 2, jTop);
                text(jx, jTop + 15, loop.jbModule || '', 9, 'center'); line(jx - jw / 2, jTop + 25, jx + jw / 2, jTop + 25);

                const sTermYJ = y - ((tCount - 1) * 20) / 2;
                for (let j = 0; j < tCount; j++) {
                    const ty = sTermYJ + (j * 20);
                    drawSq(jx - 20, ty - 7.5, category === 'STD' ? (loop.jbTerms ? loop.jbTerms[j] : (j + 1).toString()) : (j + 1).toString(), 40);
                }
                const peLineY = sTermYJ + (tCount * 20);
                const peBoxY = peLineY + 5;
                drawSq(jx - 20, peBoxY, loop.jbTermSh || 'PE', 40);
                ground(jx, peBoxY + 15);
            }

            // Marshalling
            if (loop.marshEnabled) {
                const mh = tCount * 20 + 120;
                const mw = 60;
                const mTop = y - mh / 2;
                line(mx - mw / 2, mTop, mx + mw / 2, mTop); line(mx + mw / 2, mTop, mx + mw / 2, mTop + mh);
                line(mx + mw / 2, mTop + mh, mx - mw / 2, mTop + mh); line(mx - mw / 2, mTop + mh, mx - mw / 2, mTop);
                text(mx, mTop + 15, loop.marshModule || '', 9, 'center'); line(mx - mw / 2, mTop + 25, mx + mw / 2, mTop + 25);

                const sTermYM = y - ((tCount - 1) * 20) / 2;
                for (let j = 0; j < tCount; j++) {
                    const ty = sTermYM + (j * 20);
                    drawSq(mx - 20, ty - 7.5, category === 'STD' ? (loop.marshTerms ? loop.marshTerms[j] : (j + 1).toString()) : (j + 1).toString(), 40);
                }
                const peLineY = sTermYM + (tCount * 20);
                const peBoxY = peLineY + 5;
                drawSq(mx - 20, peBoxY, loop.marshTermInSh || 'PE', 40);
                ground(mx, peBoxY + 15);
            }

            // PLC Cabinet Layout
            let ph = (category === 'VALVE' || category === 'MOTOR') ? 260 : 200, pw = 280;
            const pcx = px; // Center of PLC column
            line(pcx - pw / 2, y - ph / 2, pcx + pw / 2, y - ph / 2);
            line(pcx + pw / 2, y - ph / 2, pcx + pw / 2, y + ph / 2);
            line(pcx + pw / 2, y + ph / 2, pcx - pw / 2, y + ph / 2);
            line(pcx - pw / 2, y + ph / 2, pcx - pw / 2, y - ph / 2);
            text(pcx, y - ph / 2 - 10, loop.plcTag, 10, 'center');

            const termX = pcx - pw / 2 + 10; // Exactly 10px inside
            const ptw = 35;
            const sTermYP = y - ((tCount - 1) * 20) / 2;
            for (let j = 0; j < tCount; j++) {
                const ty = sTermYP + (j * 20);
                const t = (loop.plcTerms && loop.plcTerms[j]) || (j + 1).toString();
                drawSq(termX, ty - 7.5, t, ptw);
            }
            // PE terminal for PLC
            const pey = sTermYP + (tCount * 20);
            drawSq(termX, pey - 7.5, loop.plcTermPe || 'PE', ptw);
            ground(termX + ptw / 2, pey + 7.5);

            // TAG + signal labels on right side
            const signalNames = loop.signalNames || [];
            const sigCount = (category === 'VALVE') ? 3 : (category === 'MOTOR' ? 4 : 1);
            const tagYOff = (sigCount === 1) ? 22 : 30;
            text(pcx - pw / 2 + 55, y - ph / 2 + tagYOff, 'TAG: ' + (loop.instTag || ''), 10, 'left');

            for (let g = 0; g < sigCount; g++) {
                const pairCy = sTermYP + (g * 2) * 20 + 10;
                const label = signalNames[g] || `SIG${g + 1}`;
                if (label) text(pcx - pw / 2 + 55, pairCy + 3.5, label, 8, 'left');
            }

            if (sigCount === 1) {
                const dets = (loop.plcDetails || '').split('\n').filter(l => l.trim() !== '');
                let dy = sTermYP + 40;
                dets.forEach(dl => { text(pcx - pw / 2 + 55, dy, dl, 8, 'left'); dy += 16; });
            }

            // Cables
            const drawCab = (tag_in, specs, sh, arm, ct) => {
                let sx, dx;
                if (ct === 'INST-JB') { sx = ix + 55; dx = jx - 30; }
                else if (ct === 'INST-MARSH') { sx = ix + 55; dx = mx - 30; }
                else if (ct === 'INST-PLC') { sx = ix + 55; dx = px - 145; }
                else if (ct === 'JB-MARSH') { sx = jx + 30; dx = mx - 30; }
                else if (ct === 'JB-PLC') { sx = jx + 30; dx = px - 145; }
                else if (ct === 'MARSH-PLC') { sx = mx + 30; dx = px - 145; }

                const gap = ct.endsWith('PLC') ? 45 : 22;
                const csx = sx + (ct.startsWith('INST') ? 22 : gap);
                const cex = dx - gap;

                line(csx, y, cex, y, DXF_COLOR);
                const csy = y - ((tCount - 1) * 20) / 2;

                const baseTag = (tag_in || '').trim().toUpperCase();
                let wireStart = 1;
                if (cableWireMap[baseTag]) wireStart = cableWireMap[baseTag] + 1;

                for (let j = 0; j < tCount; j++) {
                    const wireNum = wireStart + j;
                    const wy = csy + (j * 20);
                    line(sx, wy, csx, wy); line(csx, wy, csx, y + (j < tCount / 2 ? -2 : 2));
                    line(cex, y + (j < tCount / 2 ? -2 : 2), cex, wy); line(cex, wy, dx, wy);
                    text(csx + 10, wy + 2.5, wireNum.toString(), 6, 'center');
                    text(cex - 10, wy + 2.5, wireNum.toString(), 6, 'center');
                }
                cableWireMap[baseTag] = wireStart + tCount - 1;

                if (sh) {
                    ellipse((csx + cex) / 2, y, 7, 20, DXF_COLOR);
                    // Target side
                    if (ct.endsWith('JB') || ct.endsWith('MARSH') || ct.endsWith('PLC')) {
                        let sey = csy + (tCount * 20);
                        let targetX = dx;
                        if (ct.endsWith('JB')) {
                            sey += 12.5; targetX = jx - 20;
                        } else if (ct.endsWith('MARSH')) {
                            sey += 12.5; targetX = mx - 20;
                        } else if (ct.endsWith('PLC')) {
                            targetX = px - 130; // PLC termX
                        }
                        dashedLine(cex, y, cex, sey, DXF_COLOR);
                        dashedLine(cex, sey, targetX, sey, DXF_COLOR);
                    }

                    // Source side
                    if (ct.startsWith('JB') || ct.startsWith('MARSH')) {
                        let sy2 = csy + (tCount * 20);
                        let targetX = sx;
                        if (ct.startsWith('JB')) {
                            sy2 += 12.5; targetX = jx + 20;
                        } else if (ct.startsWith('MARSH')) {
                            sy2 += 12.5; targetX = mx + 20;
                        }
                        dashedLine(csx, y, csx, sy2, DXF_COLOR);
                        dashedLine(csx, sy2, targetX, sy2, DXF_COLOR);
                    }
                }
                text((csx + cex) / 2, y - 45, tag_in, 10, 'center', DXF_COLOR);
                text((csx + cex) / 2, y + 45, specs, 8, 'center', DXF_COLOR);
            };

            let currentComp = 'INST';
            if (loop.jbEnabled) {
                drawCab(cNames.c1, loop.cable1Section, loop.cable1Shield, loop.cable1Armor, 'INST-JB');
                currentComp = 'JB';
            }
            if (loop.marshEnabled) {
                const sTag = (currentComp === 'INST' ? cNames.c1 : cNames.c2);
                const sType = (currentComp === 'INST' ? 'INST-MARSH' : 'JB-MARSH');
                const sSec = (currentComp === 'INST' ? loop.cable1Section : loop.cable2Section);
                const sSh = (currentComp === 'INST' ? loop.cable1Shield : loop.cable2Shield);
                const sAr = (currentComp === 'INST' ? loop.cable1Armor : loop.cable2Armor);
                drawCab(sTag, sSec, sSh, sAr, sType);
                currentComp = 'MARSH';
            }
            const fTag = (currentComp === 'MARSH' ? cNames.c3 : (currentComp === 'JB' ? cNames.c2 : cNames.c1));
            const fType = (currentComp === 'MARSH' ? 'MARSH-PLC' : (currentComp === 'JB' ? 'JB-PLC' : 'INST-PLC'));
            const fSec = (currentComp === 'MARSH' ? loop.cableSysSection : (currentComp === 'JB' ? loop.cable2Section : loop.cable1Section));
            const fSh = (currentComp === 'MARSH' ? loop.cableSysShield : (currentComp === 'JB' ? loop.cable2Shield : loop.cable1Shield));
            const fAr = (currentComp === 'MARSH' ? loop.cableSysArmor : (currentComp === 'JB' ? loop.cable2Armor : loop.cable1Armor));
            drawCab(fTag, fSec, fSh, fAr, fType);
        });
    });

    dxf += `0\nENDSEC\n0\nEOF\n`;
    const blob = new Blob([dxf], { type: 'application/dxf' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click();
}

// ── Excel (CSV) Import/Export ──────────────────────────────────────────────────
function exportCSVTemplate() {
    let headers = ["Sheet", "LoopIdx", "ProjName", "InstTag", "InstType", "InstCategory"];
    for (let i = 1; i <= 8; i++) headers.push(`Inst_T${i}`);
    headers.push("C1_Tag", "C1_Sec", "C1_Sh", "C1_Ar", "C1_Len");
    headers.push("JB_En", "JB_Tag", "JB_Mod");
    for (let i = 1; i <= 8; i++) headers.push(`JB_T${i}`);
    headers.push("JB_PE");
    headers.push("C2_Tag", "C2_Sec", "C2_Sh", "C2_Ar", "C2_Len");
    headers.push("Marsh_En", "Marsh_Tag", "Marsh_Mod");
    for (let i = 1; i <= 8; i++) headers.push(`Marsh_T${i}`);
    headers.push("Marsh_PE");
    headers.push("C3_Tag", "C3_Sec", "C3_Sh", "C3_Ar", "C3_Len");
    headers.push("PLC_Tag", "PLC_Mod");
    for (let i = 1; i <= 8; i++) headers.push(`PLC_T${i}`);
    headers.push("PLC_PE", "PLC_Det");

    let csv = headers.join(",") + "\n";

    sheets.forEach((sheet, sIdx) => {
        sheet.forEach((loop, lIdx) => {
            if (!loop.instTag) return;
            const row = [
                sIdx + 1, lIdx + 1, `"${projectName}"`, `"${loop.instTag}"`, `"${loop.instType}"`, loop.instCategory || 'STD',
                ...(loop.instTerms || ['', '', '', '', '', '', '', '']).map(t => `"${t}"`),
                `"${loop.cable1Tag}"`, loop.cable1Section, loop.cable1Shield ? 1 : 0, loop.cable1Armor ? 1 : 0, loop.cable1Len || 0,
                loop.jbEnabled ? 1 : 0, `"${loop.jbTag}"`, `"${loop.jbModule}"`,
                ...(loop.jbTerms || ['', '', '', '', '', '', '', '']).map(t => `"${t}"`),
                `"${loop.jbTermSh || 'PE'}"`,
                `"${loop.cable2Tag}"`, loop.cable2Section, loop.cable2Shield ? 1 : 0, loop.cable2Armor ? 1 : 0, loop.cable2Len || 0,
                loop.marshEnabled ? 1 : 0, `"${loop.marshTag}"`, `"${loop.marshModule}"`,
                ...(loop.marshTerms || ['', '', '', '', '', '', '', '']).map(t => `"${t}"`),
                `"${loop.marshTermInSh || 'PE'}"`,
                `"${loop.cableSysTag}"`, loop.cableSysSection, loop.cableSysShield ? 1 : 0, loop.cableSysArmor ? 1 : 0, loop.cableSysLen || 0,
                `"${loop.plcTag}"`, `"${loop.plcModule}"`,
                ...(loop.plcTerms || ['', '', '', '', '', '', '', '']).map(t => `"${t}"`),
                `"${loop.plcTermPe || 'PE'}"`,
                `"${(loop.plcDetails || '').replace(/"/g, '""')}"`
            ];
            csv += row.join(',') + "\n";
        });
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (projectName || 'lazos') + '_plantilla.csv';
    a.click();
}

function importDataFromCSV(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
        const text = event.target.result;
        const lines = text.split('\n');
        const newSheets = [createBlankSheet()];

        lines.slice(1).forEach(line => {
            if (!line.trim()) return;
            // Enhanced CSV parsing for quoted values
            const parts = [];
            let current = '';
            let inQuotes = false;
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                if (char === '"' && line[i + 1] === '"') { current += '"'; i++; }
                else if (char === '"') inQuotes = !inQuotes;
                else if (char === ',' && !inQuotes) { parts.push(current); current = ''; }
                else current += char;
            }
            parts.push(current);

            if (parts.length < 10) return;

            const sIdx = parseInt(parts[0]) - 1;
            const lIdx = parseInt(parts[1]) - 1;
            if (isNaN(sIdx) || isNaN(lIdx)) return;

            while (newSheets.length <= sIdx) newSheets.push(createBlankSheet());

            const loop = newSheets[sIdx][lIdx];
            if (!loop) return;

            if (parts[2]) projectName = parts[2].replace(/^"|"$/g, '');
            loop.instTag = parts[3].replace(/^"|"$/g, '');
            loop.instType = parts[4].replace(/^"|"$/g, '');
            loop.instCategory = parts[5];

            loop.instTerms = parts.slice(6, 14).map(p => p.replace(/^"|"$/g, ''));

            loop.cable1Tag = parts[14].replace(/^"|"$/g, '');
            loop.cable1Section = parts[15];
            loop.cable1Shield = parts[16] === "1";
            loop.cable1Armor = parts[17] === "1";
            loop.cable1Len = parseFloat(parts[18]) || 0;

            loop.jbEnabled = parts[19] === "1";
            loop.jbTag = parts[20].replace(/^"|"$/g, '');
            loop.jbModule = parts[21].replace(/^"|"$/g, '');
            loop.jbTerms = parts.slice(22, 30).map(p => p.replace(/^"|"$/g, ''));
            loop.jbTermSh = parts[30].replace(/^"|"$/g, '');

            loop.cable2Tag = parts[31].replace(/^"|"$/g, '');
            loop.cable2Section = parts[32];
            loop.cable2Shield = parts[33] === "1";
            loop.cable2Armor = parts[34] === "1";
            loop.cable2Len = parseFloat(parts[35]) || 0;

            loop.marshEnabled = parts[36] === "1";
            loop.marshTag = parts[37].replace(/^"|"$/g, '');
            loop.marshModule = parts[38].replace(/^"|"$/g, '');
            loop.marshTerms = parts.slice(39, 47).map(p => p.replace(/^"|"$/g, ''));
            loop.marshTermInSh = parts[47].replace(/^"|"$/g, '');

            loop.cableSysTag = parts[48].replace(/^"|"$/g, '');
            loop.cableSysSection = parts[49];
            loop.cableSysShield = parts[50] === "1";
            loop.cableSysArmor = parts[51] === "1";
            loop.cableSysLen = parseFloat(parts[52]) || 0;

            loop.plcTag = parts[53].replace(/^"|"$/g, '');
            loop.plcModule = parts[54].replace(/^"|"$/g, '');
            loop.plcTerms = parts.slice(55, 63).map(p => p.replace(/^"|"$/g, ''));
            loop.plcTermPe = parts[63].replace(/^"|"$/g, '');
            loop.plcDetails = parts[64] ? parts[64].replace(/^"|"$/g, '').replace(/""/g, '"') : '';
        });

        sheets = newSheets;
        currentSheetIndex = 0;
        currentLoopIndex = 0;
        document.getElementById('projectName').value = projectName;
        buildLoopButtons();
        updateSheetUI();
        loadLoopToUI(0);
        renderScene();
        alert('Datos importados correctamente');
    };
    reader.readAsText(file);
}
async function uploadToGitHub() {
    const token = prompt("Por favor, introduce tu GitHub Personal Access Token (PAT) con permisos de 'repo':");
    if (!token) return;

    const repoName = prompt("Introduce el nombre para el repositorio:", (projectName || "loop-diagram-generator").replace(/\s+/g, '-').toLowerCase());
    if (!repoName) return;

    const headers = {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
    };

    try {
        // 1. Check if user exists and create Repo
        const userRes = await fetch('https://api.github.com/user', { headers });
        if (!userRes.ok) throw new Error("Token inválido o error de conexión.");
        const userData = await userRes.json();
        const username = userData.login;

        let repoRes = await fetch(`https://api.github.com/user/repos`, {
            method: 'POST',
            headers,
            body: JSON.stringify({ name: repoName, private: false })
        });

        if (repoRes.status !== 201 && repoRes.status !== 422) {
            throw new Error("Error al crear el repositorio.");
        }

        alert(`Repositorio '${repoName}' listo. Subiendo archivos (index.html, script.js, style.css)...`);

        const files = ['index.html', 'script.js', 'style.css'];
        for (const fileName of files) {
            try {
                const fileResp = await fetch(fileName);
                if (!fileResp.ok) throw new Error(`No se pudo leer localmente ${fileName}. Asegúrate de estar ejecutando la app desde un servidor (ej: http-server).`);
                const content = await fileResp.text();
                // Base64 encoding that handles UTF-8 correctly
                const base64Content = btoa(unescape(encodeURIComponent(content)));

                const fileUrl = `https://api.github.com/repos/${username}/${repoName}/contents/${fileName}`;
                const checkFile = await fetch(fileUrl, { headers });
                let sha = null;
                if (checkFile.ok) {
                    const checkData = await checkFile.json();
                    sha = checkData.sha;
                }

                await fetch(fileUrl, {
                    method: 'PUT',
                    headers,
                    body: JSON.stringify({
                        message: `Upload ${fileName}`,
                        content: base64Content,
                        sha: sha
                    })
                });
            } catch (err) {
                console.error(`Error subiendo ${fileName}:`, err);
                alert(`Error subiendo ${fileName}: ${err.message}`);
            }
        }

        alert(`¡Éxito! Aplicación subida a: https://github.com/${username}/${repoName}`);
        window.open(`https://github.com/${username}/${repoName}`, '_blank');
    } catch (err) {
        console.error(err);
        alert("Integración con GitHub: " + err.message);
    }
}
async function downloadFromGitHub() {
    const token = prompt("Por favor, introduce tu GitHub Personal Access Token (PAT) con permisos de 'repo':");
    if (!token) return;

    const userRepo = prompt("Introduce el nombre de usuario y el repositorio (ej: usuario/repositorio):");
    if (!userRepo || !userRepo.includes('/')) return;

    const [username, repoName] = userRepo.split('/');
    const headers = {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3+json'
    };

    try {
        const fileRes = await fetch(`https://api.github.com/repos/${username}/${repoName}/contents/project.json`, { headers });
        if (!fileRes.ok) throw new Error("No se pudo encontrar 'project.json' en el repositorio.");

        const fileData = await fileRes.json();
        const content = decodeURIComponent(escape(atob(fileData.content)));
        const data = JSON.parse(content);

        if (data.sheets) {
            sheets = data.sheets;
            cableNamingMode = data.cableNamingMode || 'TAG';
            projectName = data.projectName || '';
            if (data.colors) Object.assign(CONFIG.colors, data.colors);

            currentSheetIndex = 0;
            currentLoopIndex = 0;
            document.getElementById('projectName').value = projectName;
            buildLoopButtons();
            updateSheetUI();
            loadLoopToUI(0);
            renderScene();
            alert("Proyecto descargado correctamente desde GitHub.");
        }
    } catch (err) {
        console.error(err);
        alert("Error al descargar de GitHub: " + err.message);
    }
}

// ── Cable List ────────────────────────────────────────────────────────────────
function showCableList() {
    const sumTotals = {}; // { section: totalLength }
    const cableDetails = []; // [{ loop, tag, type, origin, dest, len }]
    let globalCableCount = 1;

    sheets.forEach((sheet, sIdx) => {
        let sheetCableCount = 1; // RESET COUNT PER SHEET TO MATCH DIAGRAM
        sheet.forEach((loop, lIdx) => {
            const hasData = (loop.instTag && loop.instTag.trim() !== '') ||
                (loop.jbEnabled && loop.jbTag && loop.jbTag.trim() !== '') ||
                (loop.marshEnabled && loop.marshTag && loop.marshTag.trim() !== '');

            if (!hasData) return;

            const cNames = buildCableNames(loop, sheetCableCount, (n) => { sheetCableCount = n; });
            const loopTag = (loop.instTag && loop.instTag.trim() !== '') ? loop.instTag : `Lazo ${lIdx + 1}`;

            const addCable = (tag, type, origin, dest, len) => {
                const l = parseFloat(len) || 0;
                cableDetails.push({
                    loopTag: loopTag,
                    tag: tag || 'N/A',
                    type: type || 'S/D',
                    origin: origin || 'S/D',
                    dest: dest || 'S/D',
                    length: l
                });
                if (l > 0) {
                    const sec = (type || 'S/D').trim();
                    sumTotals[sec] = (sumTotals[sec] || 0) + l;
                }
            };

            // Cable 1
            let c1Origin = loopTag;
            let c1Dest = loop.jbEnabled ? loop.jbTag : (loop.marshEnabled ? loop.marshTag : loop.plcTag);
            addCable(cNames.c1, loop.cable1Section, c1Origin, c1Dest, loop.cable1Len);

            // Cable 2
            if (loop.jbEnabled) {
                let c2Origin = loop.jbTag;
                let c2Dest = loop.marshEnabled ? loop.marshTag : loop.plcTag;
                addCable(cNames.c2, loop.cable2Section, c2Origin, c2Dest, loop.cable2Len);
            }

            // Cable 3 (System)
            if (loop.marshEnabled) {
                let c3Origin = loop.marshTag;
                let c3Dest = loop.plcTag;
                addCable(cNames.c3, loop.cableSysSection, c3Origin, c3Dest, loop.cableSysLen);
            }
        });
    });

    // Populate Detailed Table
    const tbodyDetail = document.querySelector('#cableTableNew tbody');
    if (tbodyDetail) {
        tbodyDetail.innerHTML = '';
        cableDetails.forEach(c => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${c.loopTag}</td>
                <td>${c.tag}</td>
                <td>${c.type}</td>
                <td>${c.origin}</td>
                <td>${c.dest}</td>
                <td style="text-align:right;">${c.length.toFixed(1)} m</td>
            `;
            tbodyDetail.appendChild(tr);
        });
    }

    // Populate Summary Table
    const tbodySum = document.querySelector('#summaryTableNew tbody');
    tbodySum.innerHTML = '';
    const sections = Object.keys(sumTotals).sort();
    if (sections.length === 0) {
        tbodySum.innerHTML = '<tr><td colspan="2" style="text-align:center;">No hay cables con longitud definida > 0.</td></tr>';
    } else {
        sections.forEach(s => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${s}</td><td style="text-align:right;"><strong>${sumTotals[s].toFixed(2)} m</strong></td>`;
            tbodySum.appendChild(tr);
        });
    }

    document.getElementById('cableListView').classList.add('active');
}

function exportCableList() {
    const list = [];
    sheets.forEach((sheet, sIdx) => {
        let sheetCableCount = 1;
        sheet.forEach((loop, lIdx) => {
            const hasData = (loop.instTag && loop.instTag.trim() !== '') ||
                (loop.jbEnabled && loop.jbTag && loop.jbTag.trim() !== '') ||
                (loop.marshEnabled && loop.marshTag && loop.marshTag.trim() !== '');
            if (!hasData) return;

            const cNames = buildCableNames(loop, sheetCableCount, (n) => { sheetCableCount = n; });
            const loopTag = (loop.instTag && loop.instTag.trim() !== '') ? loop.instTag : `Lazo ${lIdx + 1}`;

            const pushCable = (tag, type, origin, dest, len) => {
                list.push([loopTag, tag || 'N/A', type || 'S/D', origin || 'S/D', dest || 'S/D', parseFloat(len) || 0]);
            };

            let c1Dest = loop.jbEnabled ? loop.jbTag : (loop.marshEnabled ? loop.marshTag : loop.plcTag);
            pushCable(cNames.c1, loop.cable1Section, loopTag, c1Dest, loop.cable1Len);

            if (loop.jbEnabled) {
                let c2Dest = loop.marshEnabled ? loop.marshTag : loop.plcTag;
                pushCable(cNames.c2, loop.cable2Section, loop.jbTag, c2Dest, loop.cable2Len);
            }
            if (loop.marshEnabled) {
                pushCable(cNames.c3, loop.cableSysSection, loop.marshTag, loop.plcTag, loop.cableSysLen);
            }
        });
    });

    // Forced sep=; for Excel column recognition
    let csv = "sep=;\n";
    csv += "Lazo;Tag Cable;Tipo;Origen;Destino;Metros\n";
    list.forEach(row => {
        // Use semicolon (;) for better compatibility with Excel
        // Sanitize '²' to '2' to prevent encoding issues (e.g. mmÂ²)
        csv += row.map(v => {
            let s = (v || '').toString();
            s = s.replace(/²/g, '2');
            return `"${s.replace(/"/g, '""')}"`;
        }).join(";") + "\n";
    });

    const totals = {};
    list.forEach(r => { if (r[5] > 0) totals[r[2]] = (totals[r[2]] || 0) + r[5]; });
    csv += "\nRESUMEN POR TIPOS\nTipo;Total(m)\n";
    Object.keys(totals).sort().forEach(s => {
        const cleanS = s.replace(/²/g, '2');
        csv += `"${cleanS}";${totals[s].toFixed(2)}\n`;
    });

    const blob = new Blob(["\ufeff" + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (projectName || 'Lista_Cables') + '.csv';
    a.click();
}
