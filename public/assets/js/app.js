let rawRows = [];
let currentFileName = '';
let compareMode = null;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const filebar = document.getElementById('filebar');
const fileNameEl = document.getElementById('fileName');
const clearBtn = document.getElementById('clearBtn');
const toleranceEl = document.getElementById('tolerance');
const tolValEl = document.getElementById('tolVal');
const errorBox = document.getElementById('errorBox');
const resultsArea = document.getElementById('resultsArea');
const resultsBody = document.getElementById('resultsBody');
const emptyState = document.getElementById('emptyState');
const downloadBtn = document.getElementById('downloadBtn');
const btnProfimetrics = document.getElementById('btnProfimetrics');
const btnC5 = document.getElementById('btnC5');
const modeSwitch = document.getElementById('modeSwitch');
const modeHint = document.getElementById('modeHint');
const rulesContainer = document.getElementById('rulesContainer');
const introText = document.getElementById('introText');

const MODE_CONFIG = {
  profimetrics: {
    label: 'Profimetrics',
    extraFields: [],
    rules: [
      { k: '1º', v: 'Seção' },
      { k: '2º', v: 'Grupo' },
      { k: '3º', v: 'Marca' },
      { k: '4º', v: 'Custo de Tabela' },
      { k: '5º', v: 'Gramagem' }
    ],
    columns: [
      { key: 'id', label: 'ID Família' },
      { key: 'familia', label: 'Família', isFamilia: true },
      { key: 'descritivo', label: 'Descritivo' },
      { key: 'marca', label: 'Marcas' },
      { key: 'secao', label: 'Seção' },
      { key: 'grupo', label: 'Grupo' },
      { key: 'custo', label: 'Custo de Tabela', isCusto: true },
      { key: 'gramagem', label: 'Gramagem', isGramagem: true }
    ]
  },
  c5: {
    label: 'C5',
    extraFields: ['subgrupo', 'cest', 'ncm', 'ipi', 'origem'],
    rules: [
      { k: '1º', v: 'Seção' },
      { k: '2º', v: 'Grupo' },
      { k: '3º', v: 'Subgrupo' },
      { k: '4º', v: 'Marca' },
      { k: '5º', v: 'CEST' },
      { k: '6º', v: 'NCM' },
      { k: '7º', v: 'IPI' },
      { k: '8º', v: 'Origem Tributária' },
      { k: '9º', v: 'Custo de Tabela' },
      { k: '10º', v: 'Gramagem' }
    ],
    columns: [
      { key: 'id', label: 'ID Família' },
      { key: 'descritivo', label: 'Descritivo' },
      { key: 'marca', label: 'Marcas' },
      { key: 'secao', label: 'Seção' },
      { key: 'grupo', label: 'Grupo' },
      { key: 'subgrupo', label: 'Subgrupo' },
      { key: 'cest', label: 'CEST' },
      { key: 'ncm', label: 'NCM' },
      { key: 'ipi', label: 'IPI' },
      { key: 'origem', label: 'Origem Tributária' },
      { key: 'custo', label: 'Custo de Tabela', isCusto: true },
      { key: 'gramagem', label: 'Gramagem', isGramagem: true }
    ]
  }
};

const FIELD_LABELS = {
  id: 'ID Família', descritivo: 'Descritivo', marca: 'Marcas', custo: 'Custo de Tabela',
  secao: 'Seção', grupo: 'Grupo', subgrupo: 'Subgrupo', cest: 'CEST', ncm: 'NCM', ipi: 'IPI', origem: 'Origem Tributária'
};

// Descritivos que começam com um código de 4 a 7 dígitos seguido de "-" já pertencem a uma família existente
function detectFamiliaFlag(descritivo){
  if (!descritivo) return '';
  return /^\s*\d{4,7}\s*-/.test(descritivo.toString()) ? 'F' : '';
}

function renderRules(mode){
  rulesContainer.innerHTML = MODE_CONFIG[mode].rules
    .map(r => '<div class="rule"><div class="k">' + r.k + '</div><div class="v">' + r.v + '</div></div>')
    .join('');
}

function selectMode(mode){
  compareMode = mode;
  btnProfimetrics.classList.toggle('active', mode === 'profimetrics');
  btnC5.classList.toggle('active', mode === 'c5');
  modeSwitch.classList.remove('required');
  modeHint.style.display = 'none';
  dropzone.classList.remove('disabled');
  renderRules(mode);
  introText.textContent = 'Suba sua planilha sempre que precisar rodar a comparação. Modo atual: ' + MODE_CONFIG[mode].label + '.';
  errorBox.innerHTML = '';
  if (rawRows.length) process();
}

btnProfimetrics.addEventListener('click', () => selectMode('profimetrics'));
btnC5.addEventListener('click', () => selectMode('c5'));

let lastGroups = [];

dropzone.addEventListener('click', () => { if (compareMode) fileInput.click(); });
dropzone.addEventListener('dragover', e => { if (!compareMode) return; e.preventDefault(); dropzone.classList.add('drag'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e => {
  if (!compareMode) return;
  e.preventDefault();
  dropzone.classList.remove('drag');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});
clearBtn.addEventListener('click', () => {
  rawRows = [];
  currentFileName = '';
  fileInput.value = '';
  filebar.style.display = 'none';
  dropzone.style.display = 'block';
  resultsArea.style.display = 'none';
  errorBox.innerHTML = '';
});
toleranceEl.addEventListener('input', () => {
  tolValEl.textContent = toleranceEl.value + '%';
  if (rawRows.length) process();
});

function normalize(s){
  return (s === null || s === undefined ? '' : s.toString())
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();
}

function handleFile(file){
  if (!compareMode){
    showError('Selecione o modo de comparação (Profimetrics ou C5) antes de enviar a planilha.');
    return;
  }
  currentFileName = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      rawRows = aoa;
      fileNameEl.textContent = file.name;
      filebar.style.display = 'flex';
      dropzone.style.display = 'none';
      process();
    } catch(err){
      showError('Não foi possível ler o arquivo. Verifique se é um .xlsx, .xls ou .csv válido.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function showError(html){
  errorBox.innerHTML = '<div class="error">' + html + '</div>';
  resultsArea.style.display = 'none';
}

const FIELD_MATCHERS = {
  id:        h => h.includes('id') && h.includes('famili'),
  descritivo:h => h.includes('descri'),
  marca:     h => h.includes('marca'),
  custo:     h => h.includes('custo'),
  secao:     h => h.includes('sec') && !h.includes('descri'),
  grupo:     h => h.includes('grupo') && !h.includes('sub'),
  subgrupo:  h => h.includes('subgrupo') || (h.includes('sub') && h.includes('grupo')),
  cest:      h => h.includes('cest'),
  ncm:       h => h.includes('ncm'),
  ipi:       h => h.includes('ipi'),
  origem:    h => h.includes('origem')
};

function detectColumns(headerRow){
  const normalized = headerRow.map(normalize);
  const map = {};
  for (const field in FIELD_MATCHERS){
    const idx = normalized.findIndex(FIELD_MATCHERS[field]);
    if (idx !== -1) map[field] = idx;
  }
  return map;
}

function parseCusto(val){
  if (typeof val === 'number') return val;
  if (val === null || val === undefined) return NaN;
  let s = val.toString().trim().replace(/[^\d,.\-]/g, '');
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else if (s.includes(',')) s = s.replace(',', '.');
  return parseFloat(s);
}

function normalizeUnit(val, unit){
  unit = unit.toLowerCase();
  if (unit === 'kg') return { value: val * 1000, type: 'g' };
  if (unit === 'g' || unit === 'gr') return { value: val, type: 'g' };
  if (unit === 'l' || unit === 'lt' || unit === 'litro' || unit === 'litros') return { value: val * 1000, type: 'ml' };
  if (unit === 'ml') return { value: val, type: 'ml' };
  return null;
}

// Extrai a gramagem/volume de um texto de descrição, ex: "10g", "500ml", "1kg", "6x100g"
function extractGramagem(desc){
  if (!desc) return null;
  const s = desc.toString();

  // Padrão com multiplicador: "6X50G", "6 x 50 g"
  const multiRegex = /(\d+)\s*[xX]\s*(\d+(?:[.,]\d+)?)\s*(kg|gr|g|ml|lt|l)\b/gi;
  const multiMatches = [...s.matchAll(multiRegex)];
  if (multiMatches.length){
    const m = multiMatches[multiMatches.length - 1];
    const qty = parseFloat(m[1]);
    const val = parseFloat(m[2].replace(',', '.'));
    return normalizeUnit(qty * val, m[3]);
  }

  // Padrão simples: "10g", "500 ml", "1kg"
  const simpleRegex = /(\d+(?:[.,]\d+)?)\s*(kg|gr|g|ml|lt|l)\b/gi;
  const simpleMatches = [...s.matchAll(simpleRegex)];
  if (simpleMatches.length){
    const m = simpleMatches[simpleMatches.length - 1];
    const val = parseFloat(m[1].replace(',', '.'));
    return normalizeUnit(val, m[2]);
  }

  return null;
}

function formatGramagem(g){
  if (!g) return '—';
  const val = Number.isInteger(g.value) ? g.value : g.value.toFixed(2).replace('.', ',');
  return val + g.type;
}

function process(){
  errorBox.innerHTML = '';
  if (!rawRows.length) return;
  if (!compareMode){
    showError('Selecione o modo de comparação (Profimetrics ou C5) antes de continuar.');
    return;
  }

  const cfg = MODE_CONFIG[compareMode];
  const headerRow = rawRows[0];
  const colMap = detectColumns(headerRow);
  const required = ['id','descritivo','marca','custo','secao','grupo'].concat(cfg.extraFields);
  const missing = required.filter(f => !(f in colMap));

  if (missing.length){
    showError(
      'Não encontrei estas colunas na planilha para o modo ' + cfg.label + ': ' +
      missing.map(m => '<code>' + FIELD_LABELS[m] + '</code>').join(', ') +
      '.<br>Cabeçalhos encontrados: <code>' + headerRow.filter(Boolean).join(', ') + '</code>'
    );
    return;
  }

  const extraFieldsAll = ['subgrupo','cest','ncm','ipi','origem'];
  const items = [];
  for (let i = 1; i < rawRows.length; i++){
    const row = rawRows[i];
    if (!row || row.every(c => c === null || c === '')) continue;
    const custo = parseCusto(row[colMap.custo]);
    const secao = row[colMap.secao];
    const marca = row[colMap.marca];
    const grupo = row[colMap.grupo];
    const descritivo = row[colMap.descritivo];

    const extra = {};
    let missingExtra = false;
    for (const f of extraFieldsAll){
      const v = (f in colMap) ? row[colMap[f]] : null;
      extra[f] = v;
      if (cfg.extraFields.includes(f) && (v === null || v === '')) missingExtra = true;
    }

    if (secao === null || marca === null || grupo === null || isNaN(custo) || missingExtra) continue;
    const gramagem = extractGramagem(descritivo);
    items.push({
      id: row[colMap.id],
      descritivo: descritivo,
      marca: marca,
      custo: custo,
      secao: secao,
      grupo: grupo,
      gramagem: gramagem,
      familia: compareMode === 'profimetrics' ? detectFamiliaFlag(descritivo) : '',
      subgrupo: extra.subgrupo, cest: extra.cest, ncm: extra.ncm, ipi: extra.ipi, origem: extra.origem,
      marcaNorm: normalize(marca),
      secaoNorm: normalize(secao),
      grupoNorm: normalize(grupo),
      subgrupoNorm: normalize(extra.subgrupo),
      cestNorm: normalize(extra.cest),
      ncmNorm: normalize(extra.ncm),
      ipiNorm: normalize(extra.ipi),
      origemNorm: normalize(extra.origem),
      // itens sem gramagem detectada recebem chave única (nunca agrupam entre si por essa regra)
      gramagemKey: gramagem ? (gramagem.type + ':' + gramagem.value) : ('sem-gramagem-' + i)
    });
  }

  const tolerance = parseFloat(toleranceEl.value);

  const base = new Map();
  for (const item of items){
    let key = item.secaoNorm + '||' + item.grupoNorm + '||' + item.marcaNorm;
    if (compareMode === 'c5'){
      key += '||' + item.subgrupoNorm + '||' + item.cestNorm + '||' + item.ncmNorm + '||' + item.ipiNorm + '||' + item.origemNorm;
    }
    key += '||' + item.gramagemKey;
    if (!base.has(key)) base.set(key, []);
    base.get(key).push(item);
  }

  const groups = [];
  for (const [, groupItems] of base){
    groupItems.sort((a,b) => a.custo - b.custo);
    let current = null;
    for (const it of groupItems){
      if (!current){
        current = { anchor: it.custo, items: [it] };
      } else {
        const diffPct = current.anchor === 0
          ? (it.custo === 0 ? 0 : Infinity)
          : Math.abs(it.custo - current.anchor) / Math.abs(current.anchor) * 100;
        if (diffPct <= tolerance){
          current.items.push(it);
        } else {
          groups.push(current);
          current = { anchor: it.custo, items: [it] };
        }
      }
    }
    if (current) groups.push(current);
  }

  const similarGroups = groups
    .filter(g => g.items.length > 1)
    .sort((a,b) => {
      const s = a.items[0].secaoNorm.localeCompare(b.items[0].secaoNorm);
      if (s !== 0) return s;
      const gr1 = a.items[0].grupoNorm.localeCompare(b.items[0].grupoNorm);
      if (gr1 !== 0) return gr1;
      const m = a.items[0].marcaNorm.localeCompare(b.items[0].marcaNorm);
      if (m !== 0) return m;
      return a.anchor - b.anchor;
    });

  lastGroups = similarGroups;
  render(items.length, similarGroups);
}

function render(totalItems, groups){
  const cfg = MODE_CONFIG[compareMode];
  resultsArea.style.display = 'block';
  document.getElementById('statTotal').textContent = totalItems;
  document.getElementById('statGroups').textContent = groups.length;
  document.getElementById('statItems').textContent = groups.reduce((s,g) => s + g.items.length, 0);

  const headRow = document.getElementById('resultsHeadRow');
  headRow.innerHTML = '<th>Nº</th>' + cfg.columns.map(c =>
    '<th' + ((c.isCusto || c.isGramagem || c.isFamilia) ? ' class="custo"' : '') + '>' + c.label + '</th>'
  ).join('');

  resultsBody.innerHTML = '';
  if (!groups.length){
    document.getElementById('resultsTable').style.display = 'none';
    emptyState.style.display = 'block';
    downloadBtn.disabled = true;
    return;
  }
  document.getElementById('resultsTable').style.display = 'table';
  emptyState.style.display = 'none';
  downloadBtn.disabled = false;

  groups.forEach((g, gi) => {
    const borderClass = 'group-border-' + (gi % 2);
    g.items.forEach(it => {
      const tr = document.createElement('tr');
      tr.className = borderClass;
      let cells = '<td class="grupo">#' + (gi+1) + '</td>';
      cfg.columns.forEach(c => {
        let val;
        if (c.isCusto) val = it.custo.toFixed(4).replace('.', ',');
        else if (c.isGramagem) val = formatGramagem(it.gramagem);
        else if (c.isFamilia) val = it.familia ? '<span class="familia-flag">F</span>' : '';
        else val = it[c.key] ?? '';
        cells += '<td' + ((c.isCusto || c.isGramagem || c.isFamilia) ? ' class="custo"' : '') + '>' + val + '</td>';
      });
      tr.innerHTML = cells;
      resultsBody.appendChild(tr);
    });
  });
}

downloadBtn.addEventListener('click', () => {
  if (!lastGroups.length) return;
  const cfg = MODE_CONFIG[compareMode];
  const header = ['Nº'].concat(cfg.columns.map(c => c.label));
  const aoa = [header];
  lastGroups.forEach((g, gi) => {
    g.items.forEach(it => {
      const row = [gi+1];
      cfg.columns.forEach(c => {
        if (c.isCusto) row.push(it.custo);
        else if (c.isGramagem) row.push(formatGramagem(it.gramagem));
        else row.push(it[c.key] ?? '');
      });
      aoa.push(row);
    });
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = header.map(() => ({ wch: 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Itens Semelhantes');
  const baseName = currentFileName.replace(/\.[^.]+$/, '') || 'planilha';
  XLSX.writeFile(wb, 'itens_semelhantes_' + cfg.label.toLowerCase() + '_' + baseName + '.xlsx');
});
