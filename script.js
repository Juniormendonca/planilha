
let workbook = null;

function parseNumber(v){
  if(v === undefined || v === null) return NaN;
  if(typeof v === 'number') return v;
  // trata 1.234,56 e 1234.56
  let s = String(v).replace(/\./g, '').replace(/,/g, '.');
  let n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

function findCell(ws, needle){
  const range = XLSX.utils.decode_range(ws['!ref']);
  const matches = [];
  for(let r=range.s.r; r<=range.e.r; r++){
    for(let c=range.s.c; c<=range.e.c; c++){
      const addr = XLSX.utils.encode_cell({r, c});
      const cell = ws[addr];
      if(cell && typeof cell.v === 'string' && cell.v.toLowerCase().includes(needle.toLowerCase())){
        matches.push({r, c, v: cell.v});
      }
    }
  }
  return matches;
}

function getRowValues(ws, r, cStart, cEnd){
  const out = [];
  for(let c=cStart; c<=cEnd; c++){
    const addr = XLSX.utils.encode_cell({r, c});
    out.push(ws[addr]?.v ?? null);
  }
  return out;
}

function findRowByLabel(ws, label, colIndex, rStart, rEnd){
  const range = XLSX.utils.decode_range(ws['!ref']);
  const end = rEnd ?? range.e.r;
  for(let r=rStart; r<=end; r++){
    const addr = XLSX.utils.encode_cell({r, c: colIndex});
    const cell = ws[addr];
    if(cell && String(cell.v).trim().toLowerCase() === label.toLowerCase()){
      return r;
    }
  }
  return -1;
}

async function readFile(file){
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data);
  document.getElementById('status').textContent = `Planilha carregada, abas: ${workbook.SheetNames.join(', ')}`;
}

document.getElementById('file').addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(file) readFile(file);
});

document.getElementById('btn').addEventListener('click', ()=>{
  if(!workbook){ 
    document.getElementById('status').textContent = "Envie a planilha primeiro";
    return;
  }
  const pergunta = document.getElementById('pergunta').value.toLowerCase();
  const sheetName = workbook.SheetNames.includes("Base Dados Linedata") ? "Base Dados Linedata" : workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];

  if(pergunta.includes("orçado") && pergunta.includes("realizado") && pergunta.includes("receita")){
    plotOrcadoVsReal(ws);
  }else if(pergunta.includes("receita") && pergunta.includes("mês")){
    plotReceitaPorMes(ws);
  }else{
    document.getElementById('status').textContent = "Pergunta não reconhecida ainda. Exemplos, Receita Líquida por mês. Receita Líquida orçado vs realizado.";
  }
});

let chart = null;
function drawChart(cfg){
  const ctx = document.getElementById('chart');
  if(chart){ chart.destroy(); }
  chart = new Chart(ctx, cfg);
}

function extractMeses(ws, headerRow){
  // meses nas colunas 2..13 zero-based, ou seja C..N
  const meses = getRowValues(ws, headerRow, 2, 13).map(v=>String(v||"").toUpperCase());
  return meses;
}

function plotReceitaPorMes(ws){
  // localizar header com "Orçado"
  const orcados = findCell(ws, "Orçado");
  if(!orcados.length){
    document.getElementById('status').textContent = "Não achei cabeçalho Orçado";
    return;
  }
  const headerRow = orcados[0].r;
  const meses = extractMeses(ws, headerRow);
  const rowReceita = findRowByLabel(ws, "Receita Liquida", 1, 0, headerRow+30);
  if(rowReceita < 0){
    document.getElementById('status').textContent = "Não encontrei a linha Receita Liquida";
    return;
  }
  const valoresRaw = getRowValues(ws, rowReceita, 2, 13).map(parseNumber);

  drawChart({
    type: 'line',
    data: {
      labels: meses,
      datasets: [{
        label: 'Receita Líquida',
        data: valoresRaw,
        borderWidth: 2,
        pointRadius: 4,
        tension: 0.25
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { labels: { color: '#fff' } }
      },
      scales: {
        x: { ticks: { color: '#fff' }, grid: { color: 'rgba(255,255,255,0.1)' } },
        y: { ticks: { color: '#fff' }, grid: { color: 'rgba(255,255,255,0.1)' } }
      }
    }
  });
  document.getElementById('status').textContent = "Gráfico gerado, Receita Líquida por mês";
}

function plotOrcadoVsReal(ws){
  const orcados = findCell(ws, "Orçado");
  const realizados = findCell(ws, "Realizado");
  if(!orcados.length || !realizados.length){
    document.getElementById('status').textContent = "Não achei blocos Orçado e Realizado";
    return;
  }
  const headerOrcado = orcados[0].r;
  const headerReal = realizados[0].r;
  const meses = extractMeses(ws, headerOrcado);

  const rowOrcado = findRowByLabel(ws, "Receita Liquida", 1, headerOrcado, headerOrcado+30);
  const rowReal = findRowByLabel(ws, "Receita Liquida", 1, headerReal, headerReal+30);
  if(rowOrcado < 0 || rowReal < 0){
    document.getElementById('status').textContent = "Não encontrei Receita Liquida nos blocos";
    return;
  }

  const orcadoVals = getRowValues(ws, rowOrcado, 2, 13).map(parseNumber);
  const realVals = getRowValues(ws, rowReal, 2, 13).map(parseNumber);

  const xlabels = meses;
  drawChart({
    type: 'bar',
    data: {
      labels: xlabels,
      datasets: [
        { label: 'Orçado', data: orcadoVals },
        { label: 'Realizado', data: realVals }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { labels: { color: '#fff' } }
      },
      scales: {
        x: { ticks: { color: '#fff' }, grid: { color: 'rgba(255,255,255,0.1)' } },
        y: { ticks: { color: '#fff' }, grid: { color: 'rgba(255,255,255,0.1)' } }
      }
    }
  });
  document.getElementById('status').textContent = "Gráfico gerado, Orçado x Realizado";
}
