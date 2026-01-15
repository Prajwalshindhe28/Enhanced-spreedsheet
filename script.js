const ROWS = 100;
const COLS = 26;

let sheetData = [];
let dependencyGraph = [];
let currentCell = { row: 0, col: 0 };

for (let i = 0; i < ROWS; i++) {
    sheetData[i] = [];
    dependencyGraph[i] = [];
    for (let j = 0; j < COLS; j++) {
        sheetData[i][j] = { value: '', formula: '', children: [] };
        dependencyGraph[i][j] = [];
    }
}

const addressDisplay = document.querySelector('.address-display');
const formulaInput = document.querySelector('.formula-input');
const columnHeaders = document.querySelector('.column-headers');
const rowsContainer = document.querySelector('.rows-container');

function getCellAddress(r, c) {
    return String.fromCharCode(65 + c) + (r + 1);
}

function parseCellAddress(addr) {
    const col = addr.charCodeAt(0) - 65;
    const row = parseInt(addr.slice(1)) - 1;
    return { row, col };
}

for (let i = 0; i < COLS; i++) {
    const h = document.createElement('div');
    h.className = 'column-header';
    h.textContent = String.fromCharCode(65 + i);
    columnHeaders.appendChild(h);
}

for (let i = 0; i < ROWS; i++) {
    const row = document.createElement('div');
    row.className = 'row-container';

    const rh = document.createElement('div');
    rh.className = 'row-header';
    rh.textContent = i + 1;
    row.appendChild(rh);

    for (let j = 0; j < COLS; j++) {
        const cell = document.createElement('div');
        cell.className = 'cell';
        cell.contentEditable = true;
        cell.dataset.row = i;
        cell.dataset.col = j;

        cell.onclick = () => selectCell(i, j);

        cell.onblur = () => {
            if (!sheetData[i][j].formula) {
                sheetData[i][j].value = cell.textContent;
                updateDependentCells(i, j);
            }
        };

        row.appendChild(cell);
    }
    rowsContainer.appendChild(row);
}

function selectCell(r, c) {
    currentCell = { row: r, col: c };
    addressDisplay.textContent = getCellAddress(r, c);
    formulaInput.value = sheetData[r][c].formula || sheetData[r][c].value;
}

formulaInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
        const { row, col } = currentCell;
        const input = formulaInput.value.trim();
        const cell = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);

        dependencyGraph[row][col] = [];
        sheetData[row][col].children = [];

        const refs = input.match(/[A-Z]\d+/g);
        if (refs) {
            sheetData[row][col].formula = input;
            refs.forEach(ref => {
                const p = parseCellAddress(ref);
                dependencyGraph[row][col].push(p);
                sheetData[p.row][p.col].children.push(getCellAddress(row, col));
            });

            if (detectCycle()) {
                alert("Cycle detected!");
                sheetData[row][col].formula = '';
                cell.textContent = '';
                return;
            }

            const val = evalFormula(input);
            sheetData[row][col].value = val;
            cell.textContent = val;
        } else {
            sheetData[row][col].value = input;
            cell.textContent = input;
        }
        updateDependentCells(row, col);
    }
});

function detectCycle() {
    const vis = Array.from({ length: ROWS }, () => Array(COLS).fill(false));
    const stack = Array.from({ length: ROWS }, () => Array(COLS).fill(false));

    function dfs(r, c) {
        vis[r][c] = stack[r][c] = true;
        for (let p of dependencyGraph[r][c]) {
            if (!vis[p.row][p.col] && dfs(p.row, p.col)) return true;
            if (stack[p.row][p.col]) return true;
        }
        stack[r][c] = false;
        return false;
    }

    for (let i = 0; i < ROWS; i++)
        for (let j = 0; j < COLS; j++)
            if (!vis[i][j] && dfs(i, j)) return true;

    return false;
}

function evalFormula(f) {
    return eval(f.replace(/[A-Z]\d+/g, m => {
        const { row, col } = parseCellAddress(m);
        return sheetData[row][col].value || 0;
    }));
}

function updateDependentCells(r, c) {
    for (let child of sheetData[r][c].children) {
        const { row, col } = parseCellAddress(child);
        const val = evalFormula(sheetData[row][col].formula);
        sheetData[row][col].value = val;
        document.querySelector(`[data-row="${row}"][data-col="${col}"]`).textContent = val;
        updateDependentCells(row, col);
    }
}

selectCell(0, 0);
