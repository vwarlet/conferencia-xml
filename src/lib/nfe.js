// Leitura de XML de NF-e / NFC-e (modelos 55 e 65) e de eventos de cancelamento.
const txt = (el, tag) => {
  const n = el ? el.getElementsByTagName(tag)[0] : null;
  return n ? n.textContent.trim() : '';
};
const num = (v) => {
  const n = parseFloat(v);
  return Number.isNaN(n) ? 0 : n;
};

export function parseXml(texto) {
  const doc = new DOMParser().parseFromString(texto, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return { erro: 'XML inválido ou corrompido' };

  if (doc.getElementsByTagName('infEvento').length) {
    return { tipo: 'evento', chave: txt(doc, 'chNFe'), evento: txt(doc, 'tpEvento') };
  }

  const inf = doc.getElementsByTagName('infNFe')[0];
  if (!inf) return { erro: 'Não é NF-e/NFC-e (CT-e e NFS-e ainda não são suportados)' };

  const ide = inf.getElementsByTagName('ide')[0];
  const emit = inf.getElementsByTagName('emit')[0];
  const dest = inf.getElementsByTagName('dest')[0];
  const total = inf.getElementsByTagName('total')[0];
  const detalhes = [...inf.getElementsByTagName('det')];
  const obrigatorios = [
    ['modelo', txt(ide, 'mod')],
    ['número', txt(ide, 'nNF')],
    ['série', txt(ide, 'serie')],
    ['data de emissão', txt(ide, 'dhEmi') || txt(ide, 'dEmi')],
    ['identificação do emitente', txt(emit, 'CNPJ') || txt(emit, 'CPF')],
    ['nome do emitente', txt(emit, 'xNome')],
    ['valor total da nota', txt(total, 'vNF')],
  ];
  const faltantes = obrigatorios.filter(([, valor]) => !valor).map(([nome]) => nome);
  if (!detalhes.length) faltantes.push('ao menos um item');
  if (!inf.getAttribute('Id')) faltantes.push('chave de acesso');
  if (faltantes.length) return { erro: `XML de NF-e/NFC-e incompleto: faltam ${faltantes.join(', ')}.` };

  const modelo = txt(ide, 'mod');
  if (!['55', '65'].includes(modelo)) return { erro: `Modelo ${modelo} não suportado. Use NF-e modelo 55 ou NFC-e modelo 65.` };

  const prot = doc.getElementsByTagName('protNFe')[0];
  const cStat = txt(prot, 'cStat');

  const itens = detalhes.map((det) => {
    const icms = det.getElementsByTagName('ICMS')[0];
    return {
      item: det.getAttribute('nItem'),
      codigo: txt(det, 'cProd'),
      descricao: txt(det, 'xProd'),
      ncm: txt(det, 'NCM'),
      cfop: txt(det, 'CFOP'),
      unidade: txt(det, 'uCom'),
      quantidade: num(txt(det, 'qCom')),
      valorUnit: num(txt(det, 'vUnCom')),
      valorItem: num(txt(det, 'vProd')),
      cst: txt(icms, 'CST') || txt(icms, 'CSOSN'),
      bcIcms: num(txt(icms, 'vBC')),
      vIcms: num(txt(icms, 'vICMS')),
      bcSt: num(txt(icms, 'vBCST')),
      vSt: num(txt(icms, 'vICMSST')),
      vPis: num(txt(det.getElementsByTagName('PIS')[0], 'vPIS')),
      vCofins: num(txt(det.getElementsByTagName('COFINS')[0], 'vCOFINS')),
      vIpi: num(txt(det.getElementsByTagName('IPI')[0], 'vIPI')),
    };
  });

  return {
    tipo: 'nota',
    chave: (inf.getAttribute('Id') || '').replace('NFe', ''),
    modelo,
    serie: txt(ide, 'serie'),
    numero: txt(ide, 'nNF'),
    dataEmissao: (txt(ide, 'dhEmi') || txt(ide, 'dEmi')).slice(0, 10),
    emitCnpj: txt(emit, 'CNPJ'),
    emitNome: txt(emit, 'xNome'),
    destDoc: txt(dest, 'CNPJ') || txt(dest, 'CPF'),
    destNome: txt(dest, 'xNome'),
    valorNota: num(txt(inf.getElementsByTagName('total')[0], 'vNF')),
    status: !prot ? 'sem protocolo' : cStat === '100' ? 'autorizada' : `cStat ${cStat}`,
    itens,
  };
}

// Confere o lote: duplicadas, canceladas, buracos de numeração e totais por mês e CFOP.
export function analisar(docs) {
  const notas = [], eventos = [], erros = [], problemas = [];
  const vistas = new Map();

  for (const d of docs) {
    if (d.erro) erros.push(d);
    else if (d.tipo === 'evento') eventos.push(d);
    else if (vistas.has(d.chave)) {
      problemas.push({ tipo: 'Duplicada', detalhe: `Nota ${d.numero} aparece em "${vistas.get(d.chave)}" e em "${d.arquivo}"` });
    } else {
      vistas.set(d.chave, d.arquivo);
      notas.push(d);
    }
  }

  const canceladas = new Set(eventos.filter((e) => e.evento === '110111').map((e) => e.chave));
  for (const n of notas) {
    if (canceladas.has(n.chave)) {
      n.status = 'cancelada';
      problemas.push({ tipo: 'Cancelada', detalhe: `Nota ${n.numero} (${n.dataEmissao}) tem evento de cancelamento` });
    }
  }

  const grupos = new Map();
  for (const n of notas) {
    const k = `${n.emitCnpj}|${n.modelo}|${n.serie}`;
    if (!grupos.has(k)) grupos.set(k, { nome: n.emitNome, modelo: n.modelo, serie: n.serie, nums: new Set() });
    grupos.get(k).nums.add(Number(n.numero));
  }
  for (const g of grupos.values()) {
    const nums = [...g.nums].sort((a, b) => a - b);
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] - nums[i - 1] > 1) {
        const de = nums[i - 1] + 1, ate = nums[i] - 1;
        problemas.push({ tipo: 'Numeração pulada', detalhe: `${g.nome}, modelo ${g.modelo}, série ${g.serie}: faltam ${de === ate ? de : `${de} a ${ate}`}` });
      }
    }
  }

  const totais = new Map();
  for (const n of notas.filter((x) => x.status !== 'cancelada')) {
    for (const it of n.itens) {
      const k = `${n.dataEmissao.slice(0, 7)}|${it.cfop}`;
      const t = totais.get(k) || { mes: n.dataEmissao.slice(0, 7), cfop: it.cfop, itens: 0, valor: 0, bcIcms: 0, vIcms: 0, vSt: 0, vPis: 0, vCofins: 0, vIpi: 0 };
      t.itens += 1; t.valor += it.valorItem; t.bcIcms += it.bcIcms; t.vIcms += it.vIcms;
      t.vSt += it.vSt; t.vPis += it.vPis; t.vCofins += it.vCofins; t.vIpi += it.vIpi;
      totais.set(k, t);
    }
  }

  return { notas, erros, problemas, totais: [...totais.values()].sort((a, b) => (a.mes + a.cfop).localeCompare(b.mes + b.cfop)) };
}
