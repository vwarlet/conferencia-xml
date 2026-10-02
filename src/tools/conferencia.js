import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { parseXml, analisar } from '../lib/nfe.js';

// Colunas da aba "Itens". O usuário escolhe quais entram na planilha.
const COLUNAS = [
  ['Dados da nota', 'numero', 'Número'], ['Dados da nota', 'serie', 'Série'], ['Dados da nota', 'dataEmissao', 'Data de emissão'],
  ['Dados da nota', 'chave', 'Chave'], ['Dados da nota', 'emitCnpj', 'CNPJ emitente'], ['Dados da nota', 'emitNome', 'Razão social emitente'],
  ['Dados da nota', 'destDoc', 'CNPJ/CPF destinatário'], ['Dados da nota', 'destNome', 'Nome destinatário'], ['Dados da nota', 'status', 'Situação'],
  ['Item', 'item', 'Item'], ['Item', 'codigo', 'Código'], ['Item', 'descricao', 'Descrição'], ['Item', 'ncm', 'NCM'], ['Item', 'cfop', 'CFOP'],
  ['Item', 'unidade', 'Unidade'], ['Item', 'quantidade', 'Quantidade'], ['Item', 'valorUnit', 'Valor unitário'], ['Item', 'valorItem', 'Valor do item'],
  ['Impostos', 'cst', 'CST/CSOSN'], ['Impostos', 'bcIcms', 'Base ICMS'], ['Impostos', 'vIcms', 'Valor ICMS'], ['Impostos', 'bcSt', 'Base ST'],
  ['Impostos', 'vSt', 'Valor ST'], ['Impostos', 'vPis', 'PIS'], ['Impostos', 'vCofins', 'COFINS'], ['Impostos', 'vIpi', 'IPI'],
];
const COLUNAS_PREVIA = new Set(['numero', 'dataEmissao', 'emitNome', 'descricao', 'cfop', 'valorItem']);

const esc = (s) => String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
const brl = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const MOEDA = new Set(['valorUnit', 'valorItem', 'bcIcms', 'vIcms', 'bcSt', 'vSt', 'vPis', 'vCofins', 'vIpi']);
const LARGURA = { descricao: 36, emitNome: 30, destNome: 30, chave: 48, destDoc: 20, emitCnpj: 18, dataEmissao: 14 };

const DICAS = {
  'Numeração pulada': 'Pode ser nota não emitida, inutilizada, ou um arquivo que faltou no lote. Vale conferir no sistema.',
  Duplicada: 'O mesmo arquivo foi enviado duas vezes. Só a primeira cópia entrou nos totais.',
  Cancelada: 'A nota foi cancelada e por isso não entra nos totais por CFOP.',
};

async function lerArquivos(files) {
  const docs = [];
  for (const f of files) {
    const nome = f.name.toLowerCase();
    if (nome.endsWith('.zip')) {
      const zip = await JSZip.loadAsync(f);
      for (const entry of Object.values(zip.files)) {
        if (!entry.dir && entry.name.toLowerCase().endsWith('.xml')) {
          docs.push({ ...parseXml(await entry.async('string')), arquivo: entry.name });
        }
      }
    } else if (nome.endsWith('.xml')) {
      docs.push({ ...parseXml(await f.text()), arquivo: f.name });
    }
  }
  return docs;
}


function somar(notas, k) { return notas.filter((n) => n.status !== 'cancelada').reduce((s, n) => s + n.itens.reduce((a, i) => a + i[k], 0), 0); }

async function gerarPlanilha(r, escolhidas) {
  const wb = new ExcelJS.Workbook();
  const ativas = r.notas.filter((n) => n.status !== 'cancelada');
  const datas = r.notas.map((n) => n.dataEmissao).filter(Boolean).sort();
  const nAtencao = r.problemas.length + r.erros.length;

  const tabela = (nome, cols, linhas, destaque) => {
    const ws = wb.addWorksheet(nome, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = cols.map((c) => ({ header: c.h, key: c.k, width: c.w || 16, style: c.m ? { numFmt: '#,##0.00' } : {} }));
    linhas.forEach((l) => ws.addRow(l));
    const cab = ws.getRow(1);
    cab.height = 22;
    cab.eachCell((c) => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A6B4D' } }; c.alignment = { vertical: 'middle' }; });
    if (cols.length && linhas.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
    if (destaque) linhas.forEach((l, i) => { const cor = destaque(l); if (cor) ws.getRow(i + 2).eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor } }; }); });
    return ws;
  };

  // 1. Resumo
  const rs = wb.addWorksheet('Resumo', { views: [{ state: 'frozen', ySplit: 1 }] });
  rs.columns = [{ width: 38 }, { width: 24 }, { width: 3 }, { width: 3 }, { width: 3 }, { width: 24 }, { width: 72 }];
  rs.properties.tabColor = { argb: 'FF126B50' };
  rs.addTable({
    name: 'ResumoLote',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium4', showRowStripes: true, showColumnStripes: false },
    columns: [{ name: 'Indicador do lote' }, { name: 'Valor ou resultado' }],
    rows: [
      ['Período das notas', datas.length ? `${datas[0]} a ${datas[datas.length - 1]}` : '-'],
      ['Notas válidas', ativas.length],
      ['Notas canceladas', r.notas.length - ativas.length],
      ['Valor total das notas válidas', ativas.reduce((s, n) => s + n.valorNota, 0)],
      ['ICMS', somar(r.notas, 'vIcms')],
      ['ICMS ST', somar(r.notas, 'vSt')],
      ['PIS', somar(r.notas, 'vPis')],
      ['COFINS', somar(r.notas, 'vCofins')],
      ['IPI', somar(r.notas, 'vIpi')],
      ['Pontos de atenção', nAtencao],
    ],
  });
  rs.getRow(1).height = 26;
  [1, 2].forEach((column) => {
    const cell = rs.getRow(1).getCell(column);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF126B50' } };
    cell.alignment = { vertical: 'middle' };
  });
  for (let n = 2; n <= 11; n++) {
    rs.getRow(n).height = 22;
    rs.getCell(`A${n}`).alignment = { vertical: 'middle' };
    rs.getCell(`B${n}`).alignment = { horizontal: n === 2 ? 'left' : 'right', vertical: 'middle' };
  }
  [5, 6, 7, 8, 9, 10].forEach((n) => { rs.getCell(`B${n}`).numFmt = '"R$" #,##0.00'; });
  rs.getRow(5).height = 26;
  [1, 2].forEach((column) => {
    const cell = rs.getRow(5).getCell(column);
    cell.font = { bold: true, color: { argb: 'FF126B50' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF2ED' } };
  });
  [1, 2].forEach((column) => {
    rs.getRow(11).getCell(column).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: nAtencao ? 'FFFDE8CF' : 'FFE4EFE9' } };
  });

  // 2. Notas
  const dn = (k, h, w, m) => ({ k, h, w, m });
  tabela('Notas', [dn('numero', 'Número', 10), dn('serie', 'Série', 8), dn('dataEmissao', 'Data de emissão', 14), dn('emitNome', 'Emitente', 30), dn('destNome', 'Destinatário', 30),
    dn('valorNota', 'Valor da nota', 16, 1), dn('status', 'Situação', 14), dn('chave', 'Chave', 48)], r.notas, (l) => (l.status === 'cancelada' ? 'FFFDE2E2' : null));

  // 3. Itens (colunas escolhidas)
  const cols = COLUNAS.filter((c) => escolhidas.has(c[1]));
  const itens = r.notas.flatMap((n) => n.itens.map((it) => Object.fromEntries(cols.map(([, k]) => [k, k in it ? it[k] : n[k]]))));
  tabela('Itens', cols.map(([, k, t]) => dn(k, t, LARGURA[k], MOEDA.has(k))), itens);

  // 4. Totais por CFOP
  tabela('Totais por CFOP', [dn('mes', 'Mês', 10), dn('cfop', 'CFOP', 8), dn('itens', 'Itens', 8), dn('valor', 'Valor dos itens', 18, 1), dn('bcIcms', 'Base ICMS', 16, 1), dn('vIcms', 'ICMS', 14, 1),
    dn('vSt', 'ST', 14, 1), dn('vPis', 'PIS', 14, 1), dn('vCofins', 'COFINS', 14, 1), dn('vIpi', 'IPI', 14, 1)], r.totais);

  // 5. Conferência
  const prob = r.problemas.map((p) => ({ tipo: p.tipo, detalhe: p.detalhe, dica: DICAS[p.tipo] || '' }))
    .concat(r.erros.map((e) => ({ tipo: 'Arquivo ignorado', detalhe: `${e.arquivo}: ${e.erro}`, dica: 'O arquivo não foi lido.' })));
  tabela('Conferência', [dn('tipo', 'Tipo', 20), dn('detalhe', 'O que foi encontrado', 70), dn('dica', 'O que significa', 80)],
    prob.length ? prob : [{ tipo: 'OK', detalhe: 'Nenhum problema encontrado no lote.', dica: '' }], (l) => (l.tipo === 'OK' ? 'FFE4EFE9' : 'FFFDE8CF'));

  // Tabela auxiliar da legenda ao lado da tabela principal do resumo.
  const legendaRows = [
    ['Resumo', 'Indicadores do lote, valores e pontos de atenção.'],
    ['Notas', 'Uma linha por nota, com emitente, destinatário, valor e situação.'],
    ['Itens', 'Produtos, CFOP, quantidades, valores e impostos.'],
    ['Totais por CFOP', 'Valores agrupados por mês e natureza da operação.'],
    ['Conferência', 'Duplicidades, cancelamentos, saltos de numeração e arquivos com erro.'],
  ];
  rs.addTable({
    name: 'LegendaAbas',
    ref: 'F1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true, showColumnStripes: false },
    columns: [{ name: 'Legenda' }, { name: 'Conteúdo' }],
    rows: legendaRows,
  });
  [6, 7].forEach((column) => {
    const cell = rs.getRow(1).getCell(column);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF536B72' } };
    cell.alignment = { vertical: 'middle' };
  });
  for (let n = 2; n <= legendaRows.length + 1; n++) {
    rs.getRow(n).height = Math.max(28, Math.ceil(String(rs.getCell(`G${n}`).value).length / 72) * 15);
    [6, 7].forEach((column) => {
      const cell = rs.getRow(n).getCell(column);
      cell.alignment = { vertical: 'top', wrapText: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: n % 2 === 0 ? 'FFF0F4F4' : 'FFFFFFFF' } };
    });
    rs.getCell(`F${n}`).font = { bold: true, color: { argb: 'FF40575D' } };
  }
  rs.getRow(1).height = 26;

  const buf = await wb.xlsx.writeBuffer();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = 'conferencia-xml.xlsx';
  a.click();
  URL.revokeObjectURL(a.href);
}

export function render(box) {
  box.innerHTML = `
    <div class="app-shell">
      <section class="hero">
        <div class="hero-copy">
          <span class="eyebrow">CONFERÊNCIA FISCAL · NF-e e NFC-e</span>
          <h1>Conferência fiscal de NF-e e NFC-e</h1>
          <p class="lead">Converta lotes de NF-e e NFC-e em uma conferência clara: totais, itens, alertas e uma planilha Excel para revisar. Tudo é processado no navegador.</p>
          <div class="hero-badges" aria-label="Principais diferenciais">
            <span>Leitura local</span>
            <span>Sem cadastro</span>
            <span>Arquivo .xlsx</span>
          </div>
        </div>
      </section>

      <section class="workspace">
        <div class="drop" tabindex="0" role="button" aria-label="Escolher arquivos XML ou ZIP">
          <div class="drop-inner">
            <span class="drop-icon">⇩</span>
            <strong>Selecione ou arraste seus arquivos</strong>
            <span>XML ou ZIP · processamento local · exportação Excel</span>
            <input type="file" accept=".xml,.zip" multiple hidden />
          </div>
        </div>
        <aside class="input-guide" aria-label="Formato de arquivo aceito">
          <strong>Formato esperado</strong>
          <p>XML padrão de NF-e (modelo 55) ou NFC-e (modelo 65), individualmente ou dentro de um ZIP. O documento precisa conter chave de acesso, dados do emitente, número, série, data de emissão, valor total e ao menos um item.</p>
          <p>CT-e, NFS-e, PDF/DANFE e planilhas não são aceitos. XML sem protocolo é identificado como “sem protocolo”; inclua o evento XML junto da nota para reconhecer cancelamentos.</p>
        </aside>
        <div id="saida" aria-live="polite"></div>
      </section>

      <section class="feature-grid" aria-label="Recursos da ferramenta">
        <h2>Mais clareza para revisar cada lote</h2>
        <article class="feature">
          <div class="icon">01</div>
          <h3>Análise em lote</h3>
          <p>Leia arquivos XML individualmente ou reúna tudo em um ZIP.</p>
        </article>
        <article class="feature">
          <div class="icon">02</div>
          <h3>Inconsistências destacadas</h3>
          <p>Encontre duplicadas, canceladas, saltos na numeração e erros de leitura.</p>
        </article>
        <article class="feature">
          <div class="icon">03</div>
          <h3>Planilha pronta para revisar</h3>
          <p>Abas de Resumo, Notas, Itens, Totais por CFOP e Conferência, com colunas de itens configuráveis.</p>
        </article>
        <article class="feature">
          <div class="icon">04</div>
          <h3>Privacidade local</h3>
          <p>Os arquivos são processados no navegador e não são enviados ao servidor.</p>
        </article>
      </section>

    </div>

    <footer class="site-footer">
      <p>Ferramenta de apoio para conferência fiscal. Sempre revise os resultados antes de usar em processos contábeis ou fiscais. Os arquivos XML são processados localmente; métricas de navegação são coletadas pelo Google Analytics.</p>
    </footer>`;

  const drop = box.querySelector('.drop');
  const input = box.querySelector('input[type=file]');
  const saida = box.querySelector('#saida');
  let ultimo = null;
  let marcadas = new Set(COLUNAS.map((c) => c[1]));
  const fmt = (k, v) => (MOEDA.has(k) && typeof v === 'number' ? v.toLocaleString('pt-BR', { minimumFractionDigits: 2 }) : v ?? '');

  function previa(r, el) {
    const cols = COLUNAS.filter((c) => COLUNAS_PREVIA.has(c[1]));
    const linhas = r.notas.flatMap((n) => n.itens.map((it) => ({ ...n, ...it }))).slice(0, 3);
    el.innerHTML = `<table><thead><tr>${cols.map((c) => `<th>${c[2]}</th>`).join('')}</tr></thead><tbody>${linhas.map((l) => `<tr>${cols.map((c) => `<td>${esc(fmt(c[1], l[c[1]]))}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  }

  async function processar(files) {
    saida.innerHTML = '<p class="dica">Lendo arquivos...</p>';
    try {
      const docs = await lerArquivos(files);
      if (!docs.length) { saida.innerHTML = '<div class="prob"><b>Nenhum XML encontrado.</b><small>Envie arquivos .xml ou um .zip que contenha XMLs de nota fiscal.</small></div>'; return; }
      ultimo = analisar(docs);
      const r = ultimo;
      const ativas = r.notas.filter((n) => n.status !== 'cancelada');
      const lista = r.problemas.map((p) => `<div class="prob"><b>${esc(p.tipo)}:</b> ${esc(p.detalhe)}<small>${esc(DICAS[p.tipo] || '')}</small></div>`).join('')
        + r.erros.map((e) => `<div class="prob"><b>Arquivo ignorado:</b> ${esc(e.arquivo)}<small>${esc(e.erro)}</small></div>`).join('');
      saida.innerHTML = `
        <div class="resumo">
          <div><b>${r.notas.length}</b><span>notas lidas</span></div>
          <div><b>${r.notas.reduce((s, n) => s + n.itens.length, 0)}</b><span>itens</span></div>
          <div><b>${brl(ativas.reduce((s, n) => s + n.valorNota, 0))}</b><span>valor das notas válidas</span></div>
          <div><b>${r.problemas.length + r.erros.length}</b><span>pontos de atenção</span></div>
        </div>
        ${lista || '<div class="ok"><b>Tudo certo.</b> Nenhum problema encontrado no lote.</div>'}
        <details open>
          <summary>Escolher colunas da aba Itens</summary>
          <p class="dica">Escolha as colunas do Excel. A prévia compacta mostra até 3 linhas com os principais dados.</p>
          <div class="colunas" id="colunas"></div>
          <div class="previa" id="previa"></div>
        </details>
        <div class="download-row">
          <button id="baixar" class="download-btn" type="button">Baixar planilha (.xlsx)</button>
        </div>`;
      const colunasEl = saida.querySelector('#colunas');
      let grupo = '';
      for (const [g, k, t] of COLUNAS) {
        if (g !== grupo) { grupo = g; colunasEl.insertAdjacentHTML('beforeend', `<div class="grupo">${g}</div>`); }
        colunasEl.insertAdjacentHTML('beforeend', `<label><input type="checkbox" data-k="${k}" ${marcadas.has(k) ? 'checked' : ''} /> ${t}</label>`);
      }
      const pv = saida.querySelector('#previa');
      previa(r, pv);
      colunasEl.onchange = () => { marcadas = new Set([...colunasEl.querySelectorAll('input:checked')].map((i) => i.dataset.k)); previa(r, pv); };
      saida.querySelector('#baixar').onclick = () => gerarPlanilha(ultimo, marcadas);
    } catch (e) {
      saida.innerHTML = `<div class="prob"><b>Não foi possível ler os arquivos.</b><small>${esc(e.message)}</small></div>`;
    }
  }

  drop.onclick = () => input.click();
  drop.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } };
  input.onchange = () => input.files.length && processar([...input.files]);
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove('over'); processar([...e.dataTransfer.files]); };
}
