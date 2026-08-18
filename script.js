/* DASHBOARD CONCRESUPER 2026 */
const ARQUIVO_DADOS = 'dados.xlsx';
const ARQUIVO_FALLBACK = 'Banco_de_Dados_ATUALIZADO.xlsx';
const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const DIAS_UTEIS_MES = 22;
const MOTIVOS = ['Concorrência','Crédito','Disponibilidade de agenda','Erro de sistema','Falta de Bomba Lança','Falta de Bomba Mangote','Falta de Caminhão Betoneira','Fora do escopo de serviço','Preço','Sem retorno','Permuta','Orçamento'];
const FCKS = ['FCK 25 CV','FCK 25 BB','FCK 30 CV','FCK 30 BB'];
const FCK_CHAVES = ['fck_25_cv','fck_25_bb','fck_30_cv','fck_30_bb'];

let DADOS = null;
let FILTROS = { mes:'todos', filial:'todos', vendedor:'todos', cidade:'todos', fck:'todos' };
let GRAFICOS = {};
let dataReferencia = new Date();

const FILTROS_POR_ABA = { visao:['mes','filial','vendedor'], propostas:['mes','vendedor','cidade'], perdas:['mes','vendedor'], concorrencia:['mes','filial','fck'], projecao:['mes','filial','vendedor'], filiais:['mes','filial'], posicionamento:['filial'] };
const FILTRO_IDS = { mes:'filtro-mes', filial:'filtro-filial', vendedor:'filtro-vendedor', cidade:'filtro-cidade', fck:'filtro-fck' };

function limpar(t){ return String(t==null?'':t).trim(); }
function normalizarNome(s){ s = limpar(s).replace(/\s+/g,' '); if(!s) return s; return s.split(' ').map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' '); }
function num(v){ if(v == null) return null; if(typeof v === 'number') return isNaN(v) ? null : v; const s = String(v).trim(); if(!s) return null; const n = parseFloat(s.replace(/\./g,'').replace(',','.')); return isNaN(n) ? null : n; }
function fmt(v, dec=0){ if(v==null || isNaN(v)) return '—'; return v.toLocaleString('pt-BR', {minimumFractionDigits: dec, maximumFractionDigits: dec}); }
function fmtPct(v, dec=1){ if(v==null || isNaN(v)) return '—'; return v.toLocaleString('pt-BR', {minimumFractionDigits: dec, maximumFractionDigits: dec}) + '%'; }
function soma(lista){ return lista.reduce((a,b)=>a+(isNaN(b)?0:b),0); }
function esc(t){ return String(t==null?'':t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

async function carregarArquivo(){
  const info = document.getElementById('infoArquivo');
  try{
    let buf = null;
    for(const nome of [ARQUIVO_DADOS, ARQUIVO_FALLBACK]){
      try{
        const res = await fetch(nome + '?t=' + Date.now(), {cache:'no-store'});
        if(res.ok){ buf = await res.arrayBuffer(); break; }
      }catch(e){}
    }
    if(!buf){
      info.textContent = '⚠️ Arquivo de dados não encontrado. Verifique se dados.xlsx está na raiz do repositório.';
      document.querySelectorAll('.painel').forEach(p=>p.innerHTML='<div class="vazio">Sem dados disponíveis</div>');
      return;
    }
    const wb = XLSX.read(buf, {type:'array'});
    DADOS = processarWorkbook(wb);
    preencherFiltros();
    renderizarTudo();
    const agora = new Date();
    document.getElementById('ultimaAtualizacao').textContent = 'Última atualização: ' + agora.toLocaleDateString('pt-BR');
    info.textContent = ARQUIVO_DADOS + ' · ' + DADOS.nPropostas + ' propostas · ' + DADOS.nPerdas + ' perdas · ' + DADOS.nProjecao + ' projeções · ' + DADOS.nMercado + ' levantamentos';
  }catch(err){
    info.textContent = '❌ Erro ao carregar: ' + err.message;
    console.error(err);
  }
}

function acharAba(wb, chave){
  const nome = wb.SheetNames.find(n => n.toLowerCase().includes(chave.toLowerCase()));
  return nome ? wb.Sheets[nome] : null;
}

function processarWorkbook(wb){
  const out = { propostas:[], perdas:[], projecao:[], mercado:[], posicionamento:{filiais:[], concorrencia:{}} };
  const wsProp = acharAba(wb, 'Propostas');
  if(wsProp){
    XLSX.utils.sheet_to_json(wsProp, {header:1, defval:''}).forEach(l => {
      const mes = normalizarMes(limpar(l[0]));
      const vendedor = normalizarNome(l[1]);
      const cidade = normalizarNome(l[2]);
      const propostas = num(l[3]);
      if(mes && vendedor && propostas != null) out.propostas.push({mes, vendedor, cidade, propostas});
    });
  }
  const wsPerdas = acharAba(wb, 'Motivos');
  if(wsPerdas){
    XLSX.utils.sheet_to_json(wsPerdas, {header:1, defval:''}).forEach(l => {
      const mes = normalizarMes(limpar(l[0]));
      const vendedor = normalizarNome(l[1]);
      if(!mes || !vendedor) return;
      const motivo = {};
      let total = 0;
      MOTIVOS.forEach((m,i) => { const v = num(l[i+2]) || 0; motivo[m] = v; total += v; });
      if(total > 0) out.perdas.push({mes, vendedor, motivo, total});
    });
  }
  const wsProj = acharAba(wb, 'Projeção');
  if(wsProj){
    XLSX.utils.sheet_to_json(wsProj, {header:1, defval:''}).forEach(l => {
      const mes = normalizarMes(limpar(l[0]));
      const filial = normalizarNome(l[1]);
      const vendedor = normalizarNome(l[2]);
      const produzido = num(l[3]);
      const meta = num(l[4]);
      if(mes && filial && produzido != null && meta != null) out.projecao.push({mes, filial, vendedor, produzido, meta});
    });
  }
  const wsMerc = acharAba(wb, 'Levantamento');
  if(wsMerc){
    XLSX.utils.sheet_to_json(wsMerc, {header:1, defval:''}).forEach(l => {
      const mes = normalizarMes(limpar(l[0]));
      const filial = normalizarNome(l[1]);
      const empresa = normalizarNome(l[2]);
      if(!mes || !filial || !empresa) return;
      const precos = {};
      FCK_CHAVES.forEach((ch,i) => precos[FCKS[i]] = num(l[i+3]));
      out.mercado.push({mes, filial, empresa, precos});
    });
  }
  const wsPos = acharAba(wb, 'Posicionamento');
  if(wsPos) out.posicionamento = parsePosicionamento(wsPos);
  out.nPropostas = out.propostas.length;
  out.nPerdas = out.perdas.length;
  out.nProjecao = out.projecao.length;
  out.nMercado = out.mercado.length;
  return out;
}

function normalizarMes(m){
  if(!m) return null;
  m = limpar(m).toLowerCase().replace(/ç/g,'c').replace(/[^a-z]/gi,'');
  const mapa = {'janeiro':'Janeiro','fevereiro':'Fevereiro','marco':'Março','abril':'Abril','maio':'Maio','junho':'Junho','julho':'Julho','agosto':'Agosto','setembro':'Setembro','outubro':'Outubro','novembro':'Novembro','dezembro':'Dezembro'};
  return mapa[m] || null;
}

/* 
   PARSER ABA POSICIONAMENTO COMERCIAL — CORRIGIDO
   A tabela resumo (colunas J=9 nome, K=10 MarketSize, L=11 Volume, M=12 Share)
   fica ALINHADA nas mesmas linhas dos blocos de concorrentes.
   Por isso a leitura do resumo é feita ANTES do processamento dos blocos,
   em TODAS as linhas — sem early-return que a impeça.
    */
function parsePosicionamento(ws){
  const linhas = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  const filiais = [];
  const vistos = new Set();
  const blocosConc = [];

  function normalizarShare(raw, ms, vol){
    let sh = raw;
    while(sh > 1.5 && sh < 100000) sh = sh / 100;
    if(ms > 0 && vol > 0 && Math.abs(sh - vol/ms) > 0.05) sh = vol / ms;
    return sh;
  }

  linhas.forEach(row => {
    const cells = (row||[]).map(c => limpar(c));
    const A = cells[0] || '';

    // 1) TABELA RESUMO — lê em TODA linha, ANTES dos returns
    const nome = cells[9];
    const ms = num(cells[10]);
    const vol = num(cells[11]);
    const shRaw = num(cells[12]);
    if(nome && !/^(volume|market|—|-|média|filial)/i.test(nome) && ms != null && vol != null && shRaw != null && ms > 0 && vol > 0){
      const chave = nome.toLowerCase();
      if(!vistos.has(chave)){
        vistos.add(chave);
        filiais.push({filial: normalizarNome(nome), marketSize: ms, volume: vol, marketShare: normalizarShare(shRaw, ms, vol)});
      }
    }

    // 2) BLOCOS DE CONCORRENTES (lado esquerdo, colunas B-H)
    if(/^filial:\s*(.+)/i.test(A)) return;
    if(/^concorr[eê]ncia/i.test(A)){
      blocosConc.push({concorrentes: cells.slice(1, 8).filter(c => c && c !== '—' && c !== '-').map(n => ({nome: normalizarNome(n), volume: null, share: null, capacidade: null}))});
      return;
    }
    if(/^volume mensal/i.test(A) && blocosConc.length){
      const vols = cells.slice(1, 8).map(num);
      const bloco = blocosConc[blocosConc.length-1];
      bloco.concorrentes.forEach((c, i) => { c.volume = vols[i] != null ? vols[i] : null; });
      return;
    }
    if(/^market share \(%\)/i.test(A) && blocosConc.length){
      const shares = cells.slice(1, 8).map(num);
      const bloco = blocosConc[blocosConc.length-1];
      bloco.concorrentes.forEach((c, i) => { c.share = shares[i] != null ? shares[i] : null; });
      return;
    }
    if(/^capacidade/i.test(A) && blocosConc.length){
      const caps = cells.slice(1, 8);
      const bloco = blocosConc[blocosConc.length-1];
      bloco.concorrentes.forEach((c, i) => { c.capacidade = caps[i] && caps[i] !== 'Dados não preenchidos' ? num(caps[i]) : null; });
      return;
    }
  });

  // Pareia blocos com filiais por ORDEM (Cascavel, Toledo, Matelândia, Guaíra, Palotina, Rondon)
  const concorrencia = {};
  const filiaisParaBlocos = filiais.filter(f => !/total/i.test(f.filial));
  blocosConc.forEach((bloco, i) => {
    if(i >= filiaisParaBlocos.length) return;
    const lista = bloco.concorrentes.filter(c => c.nome && (c.volume != null || c.share != null));
    if(lista.length) concorrencia[filiaisParaBlocos[i].filial] = lista;
  });

  return {filiais, concorrencia};
}

function preencherFiltros(){
  const uniq = arr => [...new Set(arr)];
  const meses = uniq(DADOS.propostas.map(p=>p.mes).concat(DADOS.perdas.map(p=>p.mes)).concat(DADOS.projecao.map(p=>p.mes))).filter(m => MESES.includes(m)).sort((a,b) => MESES.indexOf(a) - MESES.indexOf(b));
  preencherSelect('fMes', meses);
  preencherSelect('fFilial', uniq(DADOS.projecao.map(p=>p.filial).concat(DADOS.mercado.map(p=>p.filial)).concat(DADOS.posicionamento.filiais.map(p=>p.filial))).sort((a,b)=>a.localeCompare(b,'pt-BR')));
  preencherSelect('fVendedor', uniq(DADOS.propostas.map(p=>p.vendedor).concat(DADOS.perdas.map(p=>p.vendedor)).concat(DADOS.projecao.map(p=>p.vendedor))).sort((a,b)=>a.localeCompare(b,'pt-BR')));
  preencherSelect('fCidade', uniq(DADOS.propostas.map(p=>p.cidade)).sort((a,b)=>a.localeCompare(b,'pt-BR')));
  preencherSelect('fFck', FCKS);
}
function preencherSelect(id, valores){
  const sel = document.getElementById(id);
  sel.innerHTML = '<option value="todos">' + (id==='fMes'?'Todos os meses':id==='fFilial'?'Todas as filiais':id==='fVendedor'?'Todos os vendedores':id==='fCidade'?'Todas as cidades':'Todos os FCKs') + '</option>';
  valores.forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
}

function atualizarFiltrosVisiveis(aba){
  const ativos = FILTROS_POR_ABA[aba] || [];
  Object.keys(FILTRO_IDS).forEach(chave => { const el = document.getElementById(FILTRO_IDS[chave]); if(el) el.style.display = ativos.includes(chave) ? '' : 'none'; });
}

function aplicarFiltros(item){
  if(FILTROS.mes !== 'todos' && item.mes && item.mes !== FILTROS.mes) return false;
  if(FILTROS.filial !== 'todos' && item.filial && item.filial !== FILTROS.filial) return false;
  if(FILTROS.vendedor !== 'todos' && item.vendedor && item.vendedor !== FILTROS.vendedor) return false;
  if(FILTROS.cidade !== 'todos' && item.cidade && item.cidade !== FILTROS.cidade) return false;
  if(FILTROS.fck !== 'todos' && item.fck && item.fck !== FILTROS.fck) return false;
  return true;
}

function renderizarTudo(){ limparGraficos(); renderVisaoGeral(); renderPropostas(); renderPerdas(); renderConcorrencia(); renderProjecao(); renderFiliais(); renderPosicionamento(); }
function limparGraficos(){ Object.values(GRAFICOS).forEach(g => { if(g) g.destroy(); }); GRAFICOS = {}; }
function novoGrafico(id, config){ const ctx = document.getElementById(id); if(!ctx) return null; if(GRAFICOS[id]) GRAFICOS[id].destroy(); GRAFICOS[id] = new Chart(ctx, config); return GRAFICOS[id]; }
function kpiHTML(rotulo, valor, detalhe='', status=''){ return '<div class="kpi ' + status + '"><div class="rotulo">' + rotulo + '</div><div class="valor">' + valor + '</div>' + (detalhe ? '<div class="detalhe">' + detalhe + '</div>' : '') + '</div>'; }
function vazioHTML(msg){ return '<div class="vazio">' + (msg || 'Sem dados disponíveis') + '</div>'; }

function renderVisaoGeral(){
  const prop = DADOS.propostas.filter(aplicarFiltros);
  const perdas = DADOS.perdas.filter(aplicarFiltros);
  const proj = DADOS.projecao.filter(aplicarFiltros);
  const pos = DADOS.posicionamento;
  const totalPropostas = soma(prop.map(p=>p.propostas));
  const mesesComDados = new Set(prop.map(p=>p.mes)).size;
  const mediaDiaria = mesesComDados ? totalPropostas / (mesesComDados * DIAS_UTEIS_MES) : 0;
  const totalPerdas = soma(perdas.map(p=>p.total));
  const motivoRank = rankingMotivos(perdas);
  const principalMotivo = motivoRank[0] || null;
  const produzido = soma(proj.map(p=>p.produzido));
  const meta = soma(proj.map(p=>p.meta));
  const atingimento = meta ? (produzido/meta*100) : null;
  const marketSize = soma(pos.filiais.map(f=>f.marketSize));
  const volumeCS = soma(pos.filiais.map(f=>f.volume));
  const marketShare = marketSize ? (volumeCS/marketSize*100) : null;
  document.getElementById('kpisVisao').innerHTML = [
    kpiHTML('Propostas', fmt(totalPropostas)),
    kpiHTML('Média diária de propostas', fmt(mediaDiaria,1), mesesComDados + ' ' + (mesesComDados===1?'mês':'meses') + ' com dados (' + DIAS_UTEIS_MES + ' dias úteis)'),
    kpiHTML('Pedidos perdidos', fmt(totalPerdas)),
    kpiHTML('Principal motivo de perda', principalMotivo ? principalMotivo.motivo : '—', principalMotivo ? fmt(principalMotivo.total) + ' perdas' : ''),
    kpiHTML('Produzido', fmt(produzido,1) + ' m³', 'projeção real'),
    kpiHTML('Atingimento da meta', atingimento!=null ? fmtPct(atingimento) : '—', atingimento!=null ? 'meta ' + fmt(meta,1) + ' m³' : ''),
    kpiHTML('Market Share', marketShare!=null ? fmtPct(marketShare) : '—', 'Concresuper'),
    kpiHTML('Market Size', fmt(marketSize) + ' m³')
  ].join('');
  renderGraficoEvolucao(prop, perdas, proj);
  const porFilial = agregarPorFilial(proj);
  const filiais = Object.keys(porFilial);
  novoGrafico('grafDesempenhoFilial', { type:'bar', data:{ labels: filiais, datasets:[ { label:'Produzido', data: filiais.map(f=>porFilial[f].produzido), backgroundColor:'#1d6fa5' }, { label:'Meta', data: filiais.map(f=>porFilial[f].meta), backgroundColor:'#f59e0b' } ]}, options: opcoesBase('Projetado vs Meta (m³)') });
  renderCorrelacoes(prop, perdas, proj, pos);
}

function renderGraficoEvolucao(prop, perdas, proj){
  const btn = document.querySelector('#segEvolucao button.ativa');
  const ind = btn ? btn.dataset.ind : 'propostas';
  let dados;
  if(ind === 'propostas') dados = MESES.map(m => soma(prop.filter(p=>p.mes===m).map(p=>p.propostas)));
  else if(ind === 'perdas') dados = MESES.map(m => soma(perdas.filter(p=>p.mes===m).map(p=>p.total)));
  else dados = MESES.map(m => soma(proj.filter(p=>p.mes===m).map(p=>p.produzido)));
  const cor = ind==='propostas' ? '#1d6fa5' : ind==='perdas' ? '#dc2626' : '#16a34a';
  novoGrafico('grafEvolucao', { type:'bar', data:{ labels: MESES, datasets:[{ label: ind, data: dados, backgroundColor: cor }] }, options: opcoesBase('Evolução mensal — ' + ind) });
}

function agregarPorFilial(proj){
  const out = {};
  proj.forEach(p => { if(!out[p.filial]) out[p.filial] = {filial:p.filial, produzido:0, meta:0, propostas:0, perdas:0}; out[p.filial].produzido += p.produzido; out[p.filial].meta += p.meta; });
  return out;
}

function renderCorrelacoes(prop, perdas, proj, pos){
  const el = document.getElementById('correlacoes');
  const itens = [];
  const precoPorMes = precoMedioPorMes(DADOS.mercado, FILTROS);
  const perdaPorMes = MESES.map(m => soma(perdas.filter(p=>p.mes===m).map(p=>p.total)));
  const corrPrecoPerda = correlacao(precoPorMes, perdaPorMes);
  itens.push(corrPrecoPerda != null ? '<div class="correlacao">📊 <b>Preço × Perdas:</b> possível relação ' + descreverCorrelacao(corrPrecoPerda) + ' (correlação ' + corrPrecoPerda.toFixed(2) + ').</div>' : '<div class="correlacao">📊 <b>Preço × Perdas:</b> dados insuficientes.</div>');
  const propPorMes = MESES.map(m => soma(prop.filter(p=>p.mes===m).map(p=>p.propostas)));
  const prodPorMes = MESES.map(m => soma(proj.filter(p=>p.mes===m).map(p=>p.produzido)));
  const corrPropProd = correlacao(propPorMes, prodPorMes);
  itens.push(corrPropProd != null ? '<div class="correlacao">📦 <b>Propostas × Produção:</b> ' + (corrPropProd > 0.3 ? 'tendência positiva observada' : 'sem relação clara') + ' (correlação ' + corrPropProd.toFixed(2) + ').</div>' : '<div class="correlacao">📦 <b>Propostas × Produção:</b> dados insuficientes.</div>');
  el.innerHTML = itens.join('') || vazioHTML();
}

function precoMedioPorMes(mercado, filtros){
  return MESES.map(m => {
    const rows = mercado.filter(r => r.mes===m && (filtros.filial==='todos'||r.filial===filtros.filial) && (filtros.fck==='todos'||r.fck===filtros.fck));
    if(!rows.length) return null;
    const vals = rows.flatMap(r => FCKS.map(k=>r.precos[k])).filter(v=>v!=null);
    return vals.length ? soma(vals)/vals.length : null;
  });
}
function correlacao(xs, ys){
  const pares = xs.map((x,i)=>[x,ys[i]]).filter(p=>p[0]!=null && p[1]!=null);
  if(pares.length < 3) return null;
  const n = pares.length;
  const mx = soma(pares.map(p=>p[0]))/n, my = soma(pares.map(p=>p[1]))/n;
  let num = 0, dx = 0, dy = 0;
  pares.forEach(p=>{ num += (p[0]-mx)*(p[1]-my); dx += (p[0]-mx)*(p[0]-mx); dy += (p[1]-my)*(p[1]-my); });
  if(!dx || !dy) return null;
  return num / Math.sqrt(dx*dy);
}
function descreverCorrelacao(r){ const a = Math.abs(r); if(a > 0.7) return r>0 ? 'positiva forte' : 'negativa forte'; if(a > 0.4) return r>0 ? 'positiva moderada' : 'negativa moderada'; if(a > 0.2) return r>0 ? 'positiva fraca' : 'negativa fraca'; return 'muito fraca'; }

function renderPropostas(){
  const prop = DADOS.propostas.filter(aplicarFiltros);
  const total = soma(prop.map(p=>p.propostas));
  const meses = new Set(prop.map(p=>p.mes)).size;
  const mediaDiaria = meses ? total/(meses*DIAS_UTEIS_MES) : 0;
  const porVendedor = {};
  prop.forEach(p => { if(!porVendedor[p.vendedor]) porVendedor[p.vendedor] = {vendedor:p.vendedor, propostas:0, cidades:new Set()}; porVendedor[p.vendedor].propostas += p.propostas; porVendedor[p.vendedor].cidades.add(p.cidade); });
  const rankingV = Object.values(porVendedor).sort((a,b)=>b.propostas-a.propostas);
  const melhorV = rankingV[0] || null;
  const porCidade = {};
  prop.forEach(p => { porCidade[p.cidade] = (porCidade[p.cidade]||0) + p.propostas; });
  const melhorCidade = Object.entries(porCidade).sort((a,b)=>b[1]-a[1])[0] || null;
  document.getElementById('kpisPropostas').innerHTML = [
    kpiHTML('Total de propostas', fmt(total)),
    kpiHTML('Média diária', fmt(mediaDiaria,1), meses + ' ' + (meses===1?'mês':'meses') + ' com dados'),
    kpiHTML('Melhor vendedor', melhorV ? melhorV.vendedor : '—', melhorV ? fmt(melhorV.propostas) + ' propostas' : ''),
    kpiHTML('Melhor cidade', melhorCidade ? melhorCidade[0] : '—', melhorCidade ? fmt(melhorCidade[1]) + ' propostas' : '')
  ].join('');
  novoGrafico('grafPropMes', { type:'bar', data:{ labels:MESES, datasets:[{ label:'Propostas', data:MESES.map(m=>soma(prop.filter(p=>p.mes===m).map(p=>p.propostas))), backgroundColor:'#1d6fa5' }] }, options: opcoesBase('Janeiro → Dezembro') });
  novoGrafico('grafPropVendedor', { type:'bar', data:{ labels: rankingV.map(v=>v.vendedor), datasets:[{ label:'Propostas', data: rankingV.map(v=>v.propostas), backgroundColor:'#164e7a' }] }, options: Object.assign(opcoesBase('Vendedor | Propostas'), {indexAxis:'y'}) });
  novoGrafico('grafMediaVendedor', { type:'bar', data:{ labels: rankingV.map(v=>v.vendedor), datasets:[{ label:'Média diária', data: rankingV.map(v=>v.propostas/(meses*DIAS_UTEIS_MES)), backgroundColor:'#f59e0b' }] }, options: Object.assign(opcoesBase('Média diária = total ÷ (meses × 22 dias úteis)'), {indexAxis:'y'}) });
  const rows = rankingV.map(v => { const part = total ? (v.propostas/total*100) : 0; return '<tr><td>' + esc(v.vendedor) + '</td><td>' + esc([...v.cidades].join(', ')) + '</td><td class="num">' + fmt(v.propostas) + '</td><td class="num">' + fmtPct(part) + '</td></tr>'; }).join('');
  document.getElementById('tabelaPropostas').innerHTML = '<table><thead><tr><th>Vendedor</th><th>Cidade</th><th class="num">Propostas</th><th class="num">Participação</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

function rankingMotivos(perdas){ const acc = {}; perdas.forEach(p => MOTIVOS.forEach(m => { acc[m] = (acc[m]||0) + (p.motivo[m]||0); })); return Object.entries(acc).map(([motivo,total])=>({motivo,total})).sort((a,b)=>b.total-a.total); }
function renderPerdas(){
  const perdas = DADOS.perdas.filter(aplicarFiltros);
  const total = soma(perdas.map(p=>p.total));
  const rank = rankingMotivos(perdas);
  const principal = rank[0] || null;
  const porVendedor = {};
  perdas.forEach(p => { porVendedor[p.vendedor] = (porVendedor[p.vendedor]||0) + p.total; });
  const rankV = Object.entries(porVendedor).sort((a,b)=>b[1]-a[1]);
  document.getElementById('kpisPerdas').innerHTML = [
    kpiHTML('Total de perdas', fmt(total)),
    kpiHTML('Principal motivo', principal ? principal.motivo : '—', principal ? fmt(principal.total) + ' perdas (' + fmtPct(total?principal.total/total*100:0) + ')' : ''),
    kpiHTML('Vendedor com mais perdas', rankV[0] ? rankV[0][0] : '—', rankV[0] ? fmt(rankV[0][1]) + ' perdas' : ''),
    kpiHTML('Motivos com perdas', fmt(rank.filter(r=>r.total>0).length) + ' de ' + MOTIVOS.length)
  ].join('');
  novoGrafico('grafMotivos', { type:'bar', data:{ labels: rank.map(r=>r.motivo), datasets:[{ label:'Perdas', data: rank.map(r=>r.total), backgroundColor:'#dc2626' }] }, options: Object.assign(opcoesBase('Motivos das perdas'), {indexAxis:'y'}) });
  novoGrafico('grafPerdasMes', { type:'line', data:{ labels:MESES, datasets:[{ label:'Perdas', data:MESES.map(m=>soma(perdas.filter(p=>p.mes===m).map(p=>p.total))), borderColor:'#dc2626', backgroundColor:'rgba(220,38,38,.1)', fill:true, tension:.3 }] }, options: opcoesBase('Janeiro → Dezembro') });
  novoGrafico('grafPerdasVendedor', { type:'bar', data:{ labels: rankV.map(v=>v[0]), datasets:[{ label:'Perdas', data: rankV.map(v=>v[1]), backgroundColor:'#9f1239' }] }, options: Object.assign(opcoesBase('Vendedor | Total de perdas'), {indexAxis:'y'}) });
  novoGrafico('grafRankingMotivos', { type:'pie', data:{ labels: rank.map(r=>r.motivo), datasets:[{ data: rank.map(r=>r.total), backgroundColor: paleta(rank.length) }] }, options:{ responsive:true, plugins:{ legend:{ position:'right' } } } });
  const insights = [];
  if(principal && total){
    insights.push('💡 <b>' + principal.motivo + '</b> representa <b>' + fmtPct(principal.total/total*100) + '</b> das perdas (' + fmt(principal.total) + ' de ' + fmt(total) + ').');
    if(rank[1]) insights.push('💡 <b>' + rank[1].motivo + '</b> é o segundo maior motivo, com ' + fmt(rank[1].total) + ' perdas (' + fmtPct(rank[1].total/total*100) + ').');
    const preco = rank.find(r=>r.motivo==='Preço');
    if(preco && preco.total > 0 && preco.total/total >= 0.3) insights.push('⚠️ <b>Preço</b> concentra ' + fmtPct(preco.total/total*100) + ' das perdas — revisar posicionamento de preço.');
  }
  if(rankV[0]) insights.push('💡 <b>' + rankV[0][0] + '</b> concentra o maior volume de perdas (' + fmt(rankV[0][1]) + ').');
  document.getElementById('insightsPerdas').innerHTML = insights.map(i=>'<div class="insight">' + i + '</div>').join('') || vazioHTML();
}

function renderConcorrencia(){
  let mercado = DADOS.mercado;
  if(FILTROS.filial !== 'todos') mercado = mercado.filter(r=>r.filial===FILTROS.filial);
  if(FILTROS.fck !== 'todos') mercado = mercado.filter(r=>r.fck===FILTROS.fck);
  if(FILTROS.mes !== 'todos') mercado = mercado.filter(r=>r.mes===FILTROS.mes);
  const empresas = [...new Set(mercado.map(r=>r.empresa))].sort();
  const fckAtivo = FILTROS.fck !== 'todos' ? FILTROS.fck : 'FCK 25 CV';
  const kpis = FCKS.map(fck => { const vals = mercado.flatMap(r => { const v = r.precos[fck]; return v!=null ? [v] : []; }); const media = vals.length ? soma(vals)/vals.length : null; return kpiHTML(fck, media!=null ? 'R$ ' + fmt(media) : '—', vals.length + ' ' + (vals.length===1?'registro':'registros')); });
  document.getElementById('kpisConcorrencia').innerHTML = kpis.join('');
  const precoPorEmpresa = empresas.map(e => { const vals = mercado.filter(r=>r.empresa===e).map(r=>r.precos[fckAtivo]).filter(v=>v!=null); return vals.length ? soma(vals)/vals.length : null; });
  novoGrafico('grafPrecoConcorrente', { type:'bar', data:{ labels: empresas, datasets:[{ label: fckAtivo + ' (média)', data: precoPorEmpresa, backgroundColor:'#164e7a' }] }, options: Object.assign(opcoesBase('Preço médio por empresa — ' + fckAtivo), {indexAxis:'y'}) });
  const datasetsEvol = empresas.map((e,i) => ({ label: e, data: MESES.map(m => { const vals = mercado.filter(r=>r.mes===m && r.empresa===e).map(r=>r.precos[fckAtivo]).filter(v=>v!=null); return vals.length ? soma(vals)/vals.length : null; }), borderColor: paleta(empresas.length)[i % paleta(empresas.length).length], tension:.3, spanGaps:true }));
  novoGrafico('grafEvolucaoPrecos', { type:'line', data:{ labels:MESES, datasets:datasetsEvol }, options: opcoesBase('Evolução mensal — ' + fckAtivo) });
  const tabelaRows = [...new Set(mercado.map(r=>r.mes))].sort((a,b)=>MESES.indexOf(a)-MESES.indexOf(b)).map(mes => { return [...new Set(mercado.filter(r=>r.mes===mes).map(r=>r.filial))].map(filial => { return mercado.filter(r=>r.mes===mes && r.filial===filial).map(r => { const precos = FCKS.map(fck => r.precos[fck]!=null ? fmt(r.precos[fck]) : '—'); return '<tr><td>' + mes + '</td><td>' + esc(filial) + '</td><td>' + esc(r.empresa) + '</td>' + precos.map(p=>'<td class="num">' + p + '</td>').join('') + '</tr>'; }).join(''); }).join(''); }).join('');
  document.getElementById('tabelaPrecos').innerHTML = '<table><thead><tr><th>Mês</th><th>Filial</th><th>Empresa</th>' + FCKS.map(f=>'<th class="num">' + f + '</th>').join('') + '</tr></thead><tbody>' + tabelaRows + '</tbody></table>';
  const temConcresuper = mercado.some(r=>/concresuper/i.test(r.empresa));
  const elPos = document.getElementById('posicionamentoPreco');
  if(!temConcresuper){ elPos.innerHTML = '<div class="vazio">Preço da Concresuper não disponível no levantamento atual.</div>'; return; }
  const conc = mercado.filter(r=>!/concresuper/i.test(r.empresa));
  const cs = mercado.filter(r=>/concresuper/i.test(r.empresa));
  const html = FCKS.map(fck => {
    const valsConc = conc.flatMap(r=>{const v=r.precos[fck]; return v!=null?[v]:[];});
    const mediaConc = valsConc.length ? soma(valsConc)/valsConc.length : null;
    const valsCS = cs.map(r=>r.precos[fck]).filter(v=>v!=null);
    const precoCS = valsCS.length ? soma(valsCS)/valsCS.length : null;
    if(mediaConc==null || precoCS==null) return '<div class="correlacao">' + fck + ': dados insuficientes.</div>';
    const dif = precoCS - mediaConc;
    const difPct = (dif/mediaConc)*100;
    const posicao = Math.abs(difPct) < 1 ? 'próximo da média' : dif > 0 ? 'acima da média' : 'abaixo da média';
    return '<div class="correlacao"><b>' + fck + ':</b> Concresuper R$ ' + fmt(precoCS) + ' × média concorrência R$ ' + fmt(mediaConc) + ' — <b>' + posicao + '</b> (' + (dif>0?'+':'') + 'R$ ' + fmt(dif) + ', ' + (dif>0?'+':'') + fmtPct(difPct) + ').</div>';
  }).join('');
  elPos.innerHTML = html;
}

function renderProjecao(){
  const proj = DADOS.projecao.filter(aplicarFiltros);
  const produzido = soma(proj.map(p=>p.produzido));
  const meta = soma(proj.map(p=>p.meta));
  const atingimento = meta ? produzido/meta*100 : null;
  const diferenca = produzido - meta;
  document.getElementById('kpisProjecao').innerHTML = [
    kpiHTML('Meta total', fmt(meta,1) + ' m³'),
    kpiHTML('Produzido', fmt(produzido,1) + ' m³'),
    kpiHTML('Atingimento', atingimento!=null ? fmtPct(atingimento) : '—', statusAtingimento(atingimento)),
    kpiHTML('Diferença para meta', (diferenca>=0?'+':'') + fmt(diferenca,1) + ' m³', diferenca>=0?'acima da meta':'abaixo da meta', diferenca>=0?'ok':'ruim')
  ].join('');
  const porV = {};
  proj.forEach(p => { if(!porV[p.vendedor]) porV[p.vendedor] = {vendedor:p.vendedor, produzido:0, meta:0}; porV[p.vendedor].produzido += p.produzido; porV[p.vendedor].meta += p.meta; });
  const listaV = Object.values(porV).sort((a,b)=>(b.produzido/b.meta)-(a.produzido/a.meta));
  novoGrafico('grafMetaProduzido', { type:'bar', data:{ labels: listaV.map(v=>v.vendedor), datasets:[ { label:'Produzido', data: listaV.map(v=>v.produzido), backgroundColor:'#16a34a' }, { label:'Meta', data: listaV.map(v=>v.meta), backgroundColor:'#f59e0b' } ]}, options: Object.assign(opcoesBase('Produzido vs Meta (m³)'), {indexAxis:'y'}) });
  novoGrafico('grafAtingimento', { type:'bar', data:{ labels: listaV.map(v=>v.vendedor), datasets:[{ label:'% atingimento', data: listaV.map(v=>v.meta?v.produzido/v.meta*100:0), backgroundColor:'#164e7a' }] }, options: Object.assign(opcoesBase('Vendedor | Meta | Produzido | % atingimento'), {indexAxis:'y'}) });
  const hoje = dataReferencia;
  const mesAtual = MESES[hoje.getMonth()];
  const rowsMes = proj.filter(p=>p.mes===mesAtual);
  const produzidoMes = soma(rowsMes.map(p=>p.produzido));
  const metaMes = soma(rowsMes.map(p=>p.meta));
  const diaAtual = hoje.getDate();
  const diasNoMes = new Date(hoje.getFullYear(), hoje.getMonth()+1, 0).getDate();
  let htmlProj = '<div class="projecao">';
  htmlProj += '<div class="item"><div class="rotulo">Mês de referência</div><div class="valor">' + mesAtual + ' / ' + hoje.getFullYear() + '</div><div class="nota">data atual do navegador</div></div>';
  if(rowsMes.length && diaAtual > 0){
    const ritmoDia = produzidoMes / diaAtual;
    const projecaoFinal = ritmoDia * diasNoMes;
    htmlProj += '<div class="item"><div class="rotulo">Produzido até hoje</div><div class="valor">' + fmt(produzidoMes,1) + ' m³</div><div class="nota">' + diaAtual + ' de ' + diasNoMes + ' dias</div></div>';
    htmlProj += '<div class="item"><div class="rotulo">Ritmo atual</div><div class="valor">' + fmt(ritmoDia,1) + ' m³/dia</div><div class="nota">produzido ÷ dias transcorridos</div></div>';
    htmlProj += '<div class="item"><div class="rotulo">Projeção estimada</div><div class="valor">' + fmt(projecaoFinal,1) + ' m³</div><div class="nota">ritmo × ' + diasNoMes + ' dias — estimativa calculada</div></div>';
    htmlProj += '<div class="item"><div class="rotulo">Meta do mês</div><div class="valor">' + fmt(metaMes,1) + ' m³</div><div class="nota">' + (metaMes?fmtPct(projecaoFinal/metaMes*100):'—') + ' da meta se mantido o ritmo</div></div>';
  } else {
    htmlProj += '<div class="item"><div class="rotulo">Sem dados</div><div class="valor">Sem dados disponíveis</div><div class="nota">não há projeção para o mês atual</div></div>';
  }
  htmlProj += '</div>';
  document.getElementById('projecaoEstimada').innerHTML = htmlProj;
  const rows = listaV.map(v => { const pct = v.meta ? v.produzido/v.meta*100 : 0; return '<tr><td>' + esc(v.vendedor) + '</td><td class="num">' + fmt(v.meta,1) + '</td><td class="num">' + fmt(v.produzido,1) + '</td><td class="num">' + fmtPct(pct) + '</td></tr>'; }).join('');
  document.getElementById('tabelaProjecao').innerHTML = '<table><thead><tr><th>Vendedor</th><th class="num">Meta</th><th class="num">Produzido</th><th class="num">% Atingimento</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

function statusAtingimento(pct){ if(pct==null) return ''; if(pct >= 80) return 'rumo à meta'; if(pct >= 50) return 'em evolução'; return 'abaixo do esperado'; }

function renderFiliais(){
  const proj = DADOS.projecao.filter(aplicarFiltros);
  const prop = DADOS.propostas.filter(aplicarFiltros);
  const pos = DADOS.posicionamento;
  const mapa = {};
  proj.forEach(p => { if(!mapa[p.filial]) mapa[p.filial] = {filial:p.filial, produzido:0, meta:0, propostas:0, perdas:0, marketSize:null, marketShare:null}; mapa[p.filial].produzido += p.produzido; mapa[p.filial].meta += p.meta; });
  prop.forEach(p => { if(mapa[p.cidade]) mapa[p.cidade].propostas += p.propostas; });
  pos.filiais.forEach(f => { if(mapa[f.filial]) { mapa[f.filial].marketSize = f.marketSize; mapa[f.filial].marketShare = f.marketShare; } });
  const filiais = Object.values(mapa);
  filiais.forEach(f => { f.atingimento = f.meta ? f.produzido/f.meta*100 : null; });
  const nFiliais = filiais.length;
  const maiorProd = [...filiais].sort((a,b)=>b.produzido-a.produzido)[0];
  const maiorAting = [...filiais].sort((a,b)=>(b.atingimento==null?-1:b.atingimento)-(a.atingimento==null?-1:a.atingimento))[0];
  document.getElementById('kpisFiliais').innerHTML = [
    kpiHTML('Filiais com dados', fmt(nFiliais)),
    kpiHTML('Maior produção', maiorProd ? maiorProd.filial : '—', maiorProd ? fmt(maiorProd.produzido,1) + ' m³' : ''),
    kpiHTML('Maior atingimento', maiorAting ? maiorAting.filial : '—', maiorAting && maiorAting.atingimento!=null ? fmtPct(maiorAting.atingimento) : ''),
    kpiHTML('Produzido total', fmt(soma(filiais.map(f=>f.produzido)),1) + ' m³')
  ].join('');
  const rows = filiais.map(f => { const share = f.marketShare != null ? fmtPct(f.marketShare*100) : '—'; const size = f.marketSize != null ? fmt(f.marketSize) : '—'; return '<tr><td>' + esc(f.filial) + '</td><td class="num">' + fmt(f.propostas) + '</td><td class="num">' + fmt(f.perdas) + '</td><td class="num">' + fmt(f.produzido,1) + '</td><td class="num">' + fmt(f.meta,1) + '</td><td class="num">' + (f.atingimento!=null?fmtPct(f.atingimento):'—') + '</td><td class="num">' + size + '</td><td class="num">' + share + '</td></tr>'; }).join('');
  document.getElementById('tabelaFiliais').innerHTML = '<table><thead><tr><th>Filial</th><th class="num">Propostas</th><th class="num">Perdas</th><th class="num">Produzido</th><th class="num">Meta</th><th class="num">Atingimento</th><th class="num">Market Size</th><th class="num">Market Share</th></tr></thead><tbody>' + rows + '</tbody></table>';
  novoGrafico('grafComparativoFilial', { type:'bar', data:{ labels: filiais.map(f=>f.filial), datasets:[ { label:'Produzido', data: filiais.map(f=>f.produzido), backgroundColor:'#16a34a' }, { label:'Meta', data: filiais.map(f=>f.meta), backgroundColor:'#f59e0b' } ]}, options: opcoesBase('Filial × Produzido × Meta (m³)') });
  document.getElementById('rankProducao').innerHTML = rankingHTML(medalhas([...filiais].sort((a,b)=>b.produzido-a.produzido).map(f=>({nome:f.filial, valor:fmt(f.produzido,1)+' m³'}))));
  document.getElementById('rankAtingimento').innerHTML = rankingHTML(medalhas([...filiais].sort((a,b)=>(b.atingimento==null?-1:b.atingimento)-(a.atingimento==null?-1:a.atingimento)).map(f=>({nome:f.filial, valor:f.atingimento!=null?fmtPct(f.atingimento):'—'}))));
  document.getElementById('rankPropostas').innerHTML = rankingHTML(medalhas([...filiais].sort((a,b)=>b.propostas-a.propostas).map(f=>({nome:f.filial, valor:fmt(f.propostas)}))));
  document.getElementById('rankPerdas').innerHTML = rankingHTML(medalhas([...filiais].sort((a,b)=>a.perdas-b.perdas).map(f=>({nome:f.filial, valor:fmt(f.perdas)}))));
}

function medalhas(lista){ const m = ['🥇','🥈','🥉']; return lista.map((item,i)=>({...item, medalha: i<3 ? m[i] : (i+1)+'º'})); }
function rankingHTML(lista){ if(!lista.length) return vazioHTML(); return lista.map(item => '<div class="item"><span class="medalha">' + item.medalha + '</span><span class="nome">' + esc(item.nome) + '</span><span class="valor">' + item.valor + '</span></div>').join(''); }

function acharConcorrencia(concorrencia, filial){
  if(concorrencia[filial]) return concorrencia[filial];
  const f = filial.toLowerCase();
  const chave = Object.keys(concorrencia).find(k => { const kk = k.toLowerCase(); return kk === f || kk.includes(f) || f.includes(kk) || kk.split(' ').some(p => f.split(' ').includes(p) && p.length > 3); });
  return chave ? concorrencia[chave] : [];
}

function renderPosicionamento(){
  const pos = DADOS.posicionamento;
  const filiais = pos.filiais;
  const concorrencia = pos.concorrencia;
  const marketSize = soma(filiais.map(f=>f.marketSize));
  const volumeCS = soma(filiais.map(f=>f.volume));
  const marketShare = marketSize ? (volumeCS/marketSize*100) : null;
  let principalConc = null;
  Object.values(concorrencia).forEach(lista => { lista.forEach(c => { if(c.volume != null && (!principalConc || c.volume > principalConc.volume)) principalConc = {nome:c.nome, volume:c.volume}; }); });
  document.getElementById('kpisPosicionamento').innerHTML = [
    kpiHTML('Market Size', fmt(marketSize) + ' m³'),
    kpiHTML('Volume Concresuper', fmt(volumeCS,1) + ' m³', 'volume médio mensal'),
    kpiHTML('Market Share Concresuper', marketShare!=null ? fmtPct(marketShare) : '—'),
    kpiHTML('Principal concorrente', principalConc ? principalConc.nome : '—', principalConc ? fmt(principalConc.volume) + ' m³/mês' : '')
  ].join('');
  novoGrafico('grafMarketShare', { type:'doughnut', data:{ labels: filiais.map(f=>f.filial), datasets:[{ data: filiais.map(f=>f.marketShare*100), backgroundColor: paleta(filiais.length) }] }, options:{ responsive:true, plugins:{ legend:{ position:'right' } } } });
  const todosConc = [];
  Object.values(concorrencia).forEach(lista => lista.forEach(c => { if(c.volume != null) todosConc.push({nome:c.nome, volume:c.volume}); }));
  const agregadoConc = {};
  todosConc.forEach(c => { agregadoConc[c.nome] = (agregadoConc[c.nome]||0) + c.volume; });
  const listaConc = Object.entries(agregadoConc).sort((a,b)=>b[1]-a[1]);
  novoGrafico('grafVolumeConcorrentes', { type:'bar', data:{ labels: listaConc.map(c=>c[0]), datasets:[{ label:'Volume mensal (m³)', data: listaConc.map(c=>c[1]), backgroundColor:'#164e7a' }] }, options: Object.assign(opcoesBase('Volume médio mensal por concorrente (m³)'), {indexAxis:'y'}) });
  novoGrafico('grafComparativoMercado', { type:'bar', data:{ labels: filiais.map(f=>f.filial), datasets:[ { label:'Market Size', data: filiais.map(f=>f.marketSize), backgroundColor:'#94a3b8' }, { label:'Concresuper', data: filiais.map(f=>f.volume), backgroundColor:'#f59e0b' } ]}, options: opcoesBase('Market Size × Volume Concresuper (m³)') });
  const oportunidade = filiais.map(f => ({...f, oportunidade: f.marketSize - f.volume}));
  const maiorOport = [...oportunidade].sort((a,b)=>b.oportunidade-a.oportunidade)[0];
  const htmlOport = oportunidade.map(f => { const pct = f.marketSize ? f.volume/f.marketSize*100 : null; return '<div class="correlacao"><b>' + esc(f.filial) + ':</b> Market Size ' + fmt(f.marketSize) + ' m³ · Concresuper ' + fmt(f.volume) + ' m³ · <b>Mercado potencial ' + fmt(f.oportunidade) + ' m³</b> (' + (pct!=null?fmtPct(pct):'—') + ' de participação).</div>'; }).join('');
  document.getElementById('oportunidadeMercado').innerHTML = htmlOport || vazioHTML();
  const rows = oportunidade.map(f => { const concs = acharConcorrencia(concorrencia, f.filial).filter(c=>c.volume!=null).sort((a,b)=>b.volume-a.volume); const conc = concs[0] || null; return '<tr><td>' + esc(f.filial) + '</td><td class="num">' + fmt(f.marketSize) + '</td><td class="num">' + fmt(f.volume,1) + '</td><td class="num">' + fmtPct(f.marketShare*100) + '</td><td>' + (conc?esc(conc.nome):'—') + '</td><td class="num">' + (conc&&conc.volume!=null?fmt(conc.volume):'—') + '</td><td class="num">' + fmt(f.oportunidade) + '</td></tr>'; }).join('');
  document.getElementById('tabelaEstrategica').innerHTML = '<table><thead><tr><th>Filial</th><th class="num">Market Size</th><th class="num">Concresuper</th><th class="num">Market Share</th><th>Principal Concorrente</th><th class="num">Volume Concorrente</th><th class="num">Oportunidade</th></tr></thead><tbody>' + rows + '</tbody></table>';
  const insights = [];
  filiais.forEach(f => { const concs = acharConcorrencia(concorrencia, f.filial).filter(c=>c.volume!=null).sort((a,b)=>b.volume-a.volume); const top = concs[0] || null; const linha = '📍 <b>' + f.filial + '</b> possui Market Size estimado de <b>' + fmt(f.marketSize) + ' m³</b> e Market Share Concresuper de <b>' + fmtPct(f.marketShare*100) + '</b>.'; insights.push(top ? linha + ' Principal concorrente: <b>' + esc(top.nome) + '</b> (' + fmt(top.volume) + ' m³/mês).' : linha); });
  const sharePorFilial = filiais.map(f=>f.marketShare);
  const maiorShare = sharePorFilial.length ? Math.max(...sharePorFilial) : null;
  const filialMaiorShare = maiorShare!=null ? filiais.find(f=>f.marketShare===maiorShare) : null;
  if(filialMaiorShare) insights.push('🏆 <b>' + filialMaiorShare.filial + '</b> apresenta o maior Market Share Concresuper (' + fmtPct(filialMaiorShare.marketShare*100) + ').');
  if(maiorOport) insights.push('🎯 <b>' + maiorOport.filial + '</b> apresenta a maior oportunidade potencial de mercado (' + fmt(maiorOport.oportunidade) + ' m³).');
  document.getElementById('insightsPosicionamento').innerHTML = insights.map(i=>'<div class="insight">' + i + '</div>').join('') || vazioHTML();
}

function opcoesBase(titulo){
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position:'bottom' }, title: { display:false } },
    scales: { y: { beginAtZero: true } },
    tooltip: { callbacks: { label: c => ' ' + c.dataset.label + ': ' + fmt(c.parsed.y != null ? c.parsed.y : c.parsed) } }
  };
}
function paleta(n){
  const base = ['#1d6fa5','#f59e0b','#16a34a','#dc2626','#7c3aed','#0891b2','#db2777','#65a30d','#ea580c','#475569','#0f766e','#b91c1c'];
  return Array.from({length:n}, (_,i)=>base[i%base.length]);
}

document.getElementById('btnAtualizar').addEventListener('click', () => { carregarArquivo(); });
document.getElementById('btnLimpar').addEventListener('click', () => {
  FILTROS = { mes:'todos', filial:'todos', vendedor:'todos', cidade:'todos', fck:'todos' };
  document.querySelectorAll('.filtros select').forEach(s => s.value = 'todos');
  renderizarTudo();
});
['fMes','fFilial','fVendedor','fCidade','fFck'].forEach(id => {
  document.getElementById(id).addEventListener('change', e => {
    const chave = id.replace('f','').toLowerCase();
    FILTROS[chave] = e.target.value;
    renderizarTudo();
  });
});
document.getElementById('segEvolucao').addEventListener('click', e => {
  const btn = e.target.closest('button');
  if(!btn) return;
  document.querySelectorAll('#segEvolucao button').forEach(b=>b.classList.remove('ativa'));
  btn.classList.add('ativa');
  renderGraficoEvolucao(DADOS.propostas.filter(aplicarFiltros), DADOS.perdas.filter(aplicarFiltros), DADOS.projecao.filter(aplicarFiltros));
});
document.getElementById('abas').addEventListener('click', e => {
  const btn = e.target.closest('.aba');
  if(!btn) return;
  document.querySelectorAll('.aba').forEach(a=>a.classList.remove('ativa'));
  btn.classList.add('ativa');
  document.querySelectorAll('.painel').forEach(p=>p.classList.remove('ativo'));
  document.getElementById('aba-'+btn.dataset.aba).classList.add('ativo');
  atualizarFiltrosVisiveis(btn.dataset.aba);
});

window.addEventListener('error', function(e){
  const info = document.getElementById('infoArquivo');
  if(info) info.textContent = '❌ Erro de script: ' + e.message;
});
carregarArquivo();