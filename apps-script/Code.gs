// Onion Chips · Estoque v12 — Apps Script
// Cole em: planilha > Extensões > Apps Script. Ajuste CHAVE e ADMIN e implante uma NOVA VERSÃO.
var CHAVE = 'troque-esta-senha';
var ADMIN = 'oninsadm123';   // senha do painel administrador (fica só aqui)
var MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
var CAB = ['Data','Hora do envio','Responsável','Dia da contagem','Categoria','Produto','Campo','Quantidade','Unidade','Observação'];

// id da lista -> prefixo da aba mensal, loja e nome amigável.
// As listas da Praça Nova mantêm os ids antigos (cozinha / frente), então o histórico continua valendo.
var LISTAS = {
  'cozinha':    {pre:'Cozinha ',         loja:'pn', nome:'Praça Nova · Cozinha'},
  'frente':     {pre:'Frente ',          loja:'pn', nome:'Praça Nova · Frente'},
  'av-cozinha': {pre:'Aviação Cozinha ', loja:'av', nome:'Aviação · Cozinha'},
  'av-frente':  {pre:'Aviação Frente ',  loja:'av', nome:'Aviação · Frente'}
};
// Linhas da aba Config
var CFG_LIN = {ativo:1, msg:2, pct:3, pn:4, av:5};

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function proteger(s){ try{ var p = s.protect().setDescription('Somente o dono'); p.removeEditors(p.getEditors()); p.setDomainEdit(false); }catch(e){} }
function protegerAbas(){ ['Config','Mínimos','Painel','Listas'].forEach(function(n){ var s = ss().getSheetByName(n); if(s) proteger(s); }); }

// ---------- Config ----------
function cfg(){
  var s = ss().getSheetByName('Config');
  if(!s){
    s = ss().insertSheet('Config');
    s.getRange('A1:B5').setValues([['ativo','SIM'],['mensagem','Acesso encerrado. Fale com a gerência.'],['alerta_pct',25],['senha_pn','123'],['senha_av','123']]);
    proteger(s);
  }
  var v = s.getRange('B1:B3').getValues(), p = Number(v[2][0]);
  return { ativo: String(v[0][0]).trim().toUpperCase() === 'SIM', msg: String(v[1][0]), pct: (v[2][0]==='' || isNaN(p)) ? 25 : p };
}
function senhaDe(loja){
  cfg();
  var s = ss().getSheetByName('Config');
  var v = String(s.getRange(loja==='av' ? CFG_LIN.av : CFG_LIN.pn, 2).getValue()).trim();
  return v || '123';
}
function setCfg(linha, rotulo, valor){
  cfg();
  var s = ss().getSheetByName('Config');
  s.getRange(linha, 1, 1, 2).setValues([[rotulo, valor]]);
}

// ---------- Listas editáveis (aba Listas) ----------
function listasSalvas(){
  var s = ss().getSheetByName('Listas'), o = {};
  if(!s || s.getLastRow() < 1) return o;
  s.getRange(1,1,s.getLastRow(),2).getValues().forEach(function(r){ if(r[0]) o[String(r[0])] = String(r[1]); });
  return o;
}
function salvarLista(id, texto){
  var s = ss().getSheetByName('Listas');
  if(!s){ s = ss().insertSheet('Listas'); proteger(s); }
  var n = s.getLastRow(), linha = 0;
  if(n > 0){
    var ids = s.getRange(1,1,n,1).getValues();
    for(var i = 0; i < n; i++){ if(String(ids[i][0]) === id){ linha = i + 1; break; } }
  }
  if(linha) s.getRange(linha, 2).setValue(texto); else s.appendRow([id, texto]);
}

// ---------- Rotas ----------
function doGet(e){
  var p = e.parameter || {};
  if(p.chave !== CHAVE) return out({ativo:false, msg:'Chave inválida'});
  var c = cfg();
  switch(p.acao){
    case 'painel':
      if(!c.ativo) return out({ok:false, erro:'bloqueado'});
      return adm(p, function(){ return dadosPainel(calcular()); });
    case 'login':
      if(!c.ativo) return out({ok:false, erro:'bloqueado'});
      return out({ok: (p.loja === 'pn' || p.loja === 'av') && senhaDe(p.loja) === String(p.senha || '')});
    case 'auth':
      return out({ok: p.adm === ADMIN});
    case 'datas':
      return adm(p, function(){ return {contagens: contagens(lst(p)).map(resumo)}; });
    case 'contagem':
      return adm(p, function(){
        var x = achar(lst(p), p.id);
        return {info: x.info, linhas: x.rows.map(function(r){ return {cat:r[4], produto:r[5], campo:r[6], qtd:r[7], un:r[8]}; })};
      });
    case 'consumo':
      return adm(p, function(){ return {linhas: consumo(lst(p), p.de || '0000-00-00', p.ate || '9999-12-31')}; });
    default:
      return out({ativo:c.ativo, msg:c.msg, listas: listasSalvas()});
  }
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    var d = JSON.parse(e.postData.contents);
    if(d.chave !== CHAVE) return out({ok:false, erro:'chave'});
    if(!cfg().ativo) return out({ok:false, erro:'bloqueado'});
    lock.waitLock(20000);
    switch(d.acao){
      case 'minimos':
        return out(salvarMin(d));
      case 'salvar_lista':
        return adm(d, function(){
          if(!LISTAS[d.lista]) throw new Error('lista inválida');
          if(typeof d.texto !== 'string' || d.texto.length < 3) throw new Error('texto inválido');
          salvarLista(d.lista, d.texto); return {};
        });
      case 'editar_contagem':
        return adm(d, function(){ return editarContagem(d); });
      case 'excluir_contagem':
        return adm(d, function(){ excluirContagem(lst(d), d.id); return {}; });
      case 'senha':
        return adm(d, function(){
          if(d.loja !== 'pn' && d.loja !== 'av') throw new Error('loja inválida');
          var n = String(d.nova || '').trim();
          if(n.length < 3) throw new Error('senha muito curta');
          setCfg(d.loja === 'av' ? CFG_LIN.av : CFG_LIN.pn, 'senha_' + d.loja, n);
          return {};
        });
      case 'alerta':
        return adm(d, function(){
          var n = Number(d.pct);
          if(isNaN(n) || n < 0 || n > 100) throw new Error('percentual inválido');
          setCfg(CFG_LIN.pct, 'alerta_pct', n); calcularEscrever(); return {};
        });
    }
    // Envio de contagem (padrão, usado pelos aparelhos)
    if(!LISTAS[d.lista]) return out({ok:false, erro:'lista'});
    var p = d.data.split('-'), dt = p[2] + '/' + p[1] + '/' + p[0];
    var nome = LISTAS[d.lista].pre + MESES[+p[1] - 1] + ' ' + p[0];
    var s = ss().getSheetByName(nome);
    if(!s){
      s = ss().insertSheet(nome); s.appendRow(CAB);
      s.getRange(1,1,1,CAB.length).setFontWeight('bold').setBackground('#1F6B4A').setFontColor('#ffffff');
      s.setFrozenRows(1); s.setColumnWidth(6, 240);
    }
    var n = s.getLastRow();   // reenvio da mesma pessoa, data e dia substitui o anterior
    if(n > 1){
      var v = s.getRange(2,1,n-1,4).getDisplayValues(), run = 0;
      for(var i = v.length - 1; i >= -1; i--){
        var m = i >= 0 && v[i][0] === dt && v[i][2] === d.resp && v[i][3] === d.dia;
        if(m) run++; else if(run){ s.deleteRows(i + 3, run); run = 0; }
      }
    }
    var rows = d.linhas.map(function(l){ return [dt, d.hora, d.resp, d.dia, l[0], l[1], l[2], l[3], l[4], l[5]]; });
    s.getRange(s.getLastRow() + 1, 1, rows.length, 4).setNumberFormat('@');
    s.getRange(s.getLastRow() + 1, 1, rows.length, CAB.length).setValues(rows);
    calcularEscrever();
    return out({ok:true, linhas:rows.length, aba:nome});
  }catch(err){ return out({ok:false, erro:String(err)}); }
  finally{ try{ lock.releaseLock(); }catch(x){} }
}

function adm(p, fn){
  if(p.adm !== ADMIN) return out({ok:false, erro:'senha'});
  try{ var r = fn() || {}; r.ok = true; return out(r); }
  catch(err){ return out({ok:false, erro:String(err)}); }
}
function lst(p){ if(!LISTAS[p.lista]) throw new Error('lista inválida'); return p.lista; }
function achar(l, id){
  var c = contagens(l).filter(function(x){ return x.id === id; })[0];
  if(!c) throw new Error('contagem não encontrada');
  return c;
}
function resumo(c){ return {id:c.id, data:c.info.data, iso:c.info.iso, hora:c.info.hora, resp:c.info.resp, dia:c.info.dia, itens:c.rows.length}; }
function calcularEscrever(){ try{ escreverPainel(calcular()); }catch(x){} }

// ---------- Contagens ----------
// Todas as contagens de uma lista, da mais antiga para a mais nova.
function contagens(l){
  var m = {}, pre = LISTAS[l].pre;
  ss().getSheets().forEach(function(s){
    var nome = s.getName();
    if(nome.indexOf(pre) !== 0 || s.getLastRow() < 2) return;
    var n = s.getLastRow() - 1;
    var disp = s.getRange(2,1,n,10).getDisplayValues(), vals = s.getRange(2,1,n,10).getValues();
    disp.forEach(function(r, i){
      var p = String(r[0]).split('/');
      if(p.length !== 3) return;
      var id = p[2] + p[1] + p[0] + '|' + r[1] + '|' + r[2] + '|' + r[3];
      var row = [r[0], r[1], r[2], r[3], r[4], r[5], r[6], vals[i][7], r[8], r[9]];
      if(!m[id]) m[id] = {id:id, info:{data:String(r[0]), iso:p[2]+'-'+p[1]+'-'+p[0], hora:String(r[1]), resp:String(r[2]), dia:String(r[3])}, rows:[], refs:[]};
      m[id].rows.push(row);
      m[id].refs.push({sh:nome, r:i + 2});
    });
  });
  return Object.keys(m).sort().map(function(k){ return m[k]; });
}

function editarContagem(d){
  var c = achar(lst(d), d.id), mp = {}, n = 0;
  (d.valores || []).forEach(function(v){ mp[v[0] + '|' + v[1]] = Number(v[2]); });
  c.refs.forEach(function(rf, i){
    var k = c.rows[i][5] + '|' + c.rows[i][6];
    if(mp.hasOwnProperty(k) && typeof c.rows[i][7] === 'number' && !isNaN(mp[k])){
      ss().getSheetByName(rf.sh).getRange(rf.r, 8).setValue(mp[k]);
      n++;
    }
  });
  calcularEscrever();
  return {atualizados: n};
}

function excluirContagem(l, id){
  var c = achar(l, id), por = {};
  c.refs.forEach(function(rf){ (por[rf.sh] = por[rf.sh] || []).push(rf.r); });
  Object.keys(por).forEach(function(sh){
    var rs = por[sh].sort(function(a, b){ return b - a; }), s = ss().getSheetByName(sh), i = 0;
    while(i < rs.length){
      var top = rs[i], cnt = 1;
      while(i + cnt < rs.length && rs[i + cnt] === top - cnt) cnt++;
      s.deleteRows(top - cnt + 1, cnt);
      i += cnt;
    }
  });
  calcularEscrever();
}

// Consumo (saídas) no período: soma das quedas entre contagens consecutivas.
function consumo(l, de, ate){
  var prev = {}, acc = {};
  contagens(l).forEach(function(c){
    if(c.info.iso < de || c.info.iso > ate) return;
    c.rows.forEach(function(r){
      var k = r[5] + '|' + r[6], q = r[7];
      if(typeof q !== 'number') return;
      var a = acc[k] = acc[k] || {produto:r[5], campo:r[6], un:r[8], consumo:0, reposicao:0, contagens:0};
      if(prev[k] !== undefined){
        var df = q - prev[k];
        if(df < 0) a.consumo += -df; else a.reposicao += df;
      }
      a.contagens++; prev[k] = q;
    });
  });
  return Object.keys(acc).map(function(k){
    var a = acc[k];
    a.consumo = Math.round(a.consumo * 1000) / 1000;
    a.reposicao = Math.round(a.reposicao * 1000) / 1000;
    return a;
  }).sort(function(x, y){ return y.consumo - x.consumo; });
}

// ---------- Painel, mínimos ----------
function lerMin(){ var s = ss().getSheetByName('Mínimos'); return (!s || s.getLastRow() < 2) ? [] : s.getRange(2,1,s.getLastRow()-1,5).getValues(); }
function norm(q, u){ var f = {g:['g',1], kg:['g',1000], ml:['ml',1], l:['ml',1000]}[String(u || '').toLowerCase()]; return f ? [q * f[1], f[0]] : [q, 'un.']; }

function calcular(){
  var mins = {}, pct = cfg().pct, res = [], ult = {};
  lerMin().forEach(function(m){ if(m[3] !== '') mins[m[0] + '|' + m[1] + '|' + m[2]] = [Number(m[3]), m[4]]; });
  Object.keys(LISTAS).forEach(function(l){
    var c = contagens(l); if(!c.length) return;
    var u = c[c.length - 1], prev = {};
    ult[l] = u.info;
    if(c.length > 1) c[c.length - 2].rows.forEach(function(r){ prev[r[5] + '|' + r[6]] = r[7]; });
    u.rows.forEach(function(r){
      var k = r[5] + '|' + r[6], q = r[7];
      if(typeof q !== 'number') return;
      var it = {lista:l, produto:r[5], campo:r[6], qtd:q, un:r[8], min:null, minun:'', status:'sem', falta:0,
                ant:(typeof prev[k] === 'number') ? prev[k] : null};
      var mn = mins[l + '|' + k];
      if(mn){
        var A = norm(q, r[8]), B = norm(mn[0], mn[1]);
        it.min = mn[0]; it.minun = mn[1];
        if(A[1] !== B[1]) it.status = 'unidade';
        else{
          it.status = A[0] <= B[0] ? 'abaixo' : (A[0] <= B[0] * (1 + pct / 100) ? 'proximo' : 'ok');
          it.falta = Math.max(0, Math.round((B[0] - A[0]) / norm(1, mn[1])[0] * 1000) / 1000);
        }
      }
      res.push(it);
    });
  });
  return {res:res, ult:ult};
}
function dadosPainel(r){
  return {ok:true, ult:r.ult, minimos:lerMin(), pct:cfg().pct,
          itens:r.res.filter(function(x){ return x.status === 'abaixo' || x.status === 'proximo'; })};
}
function escreverPainel(r){
  var s = ss().getSheetByName('Painel'); if(!s){ s = ss().insertSheet('Painel'); proteger(s); }
  s.clear();
  var ord = {abaixo:0, proximo:1, unidade:2, ok:3, sem:4};
  var nm = {abaixo:'NO MÍNIMO/ABAIXO', proximo:'PERTO DO MÍNIMO', ok:'OK', sem:'sem mínimo', unidade:'conferir unidade'};
  var cor = {abaixo:'#F4C7C3', proximo:'#FCE8B2', ok:'#D9EAD3', sem:'#FFFFFF', unidade:'#EEEEEE'};
  var res = r.res.slice().sort(function(a, b){ return ord[a.status] - ord[b.status]; });
  var H = ['Loja · Lista','Produto','Campo','Qtd atual','Unidade','Mínimo','Un. mín.','Status','Falta p/ mínimo','Variação desde a contagem anterior (negativo = consumo)'];
  var rows = res.map(function(x){
    return [LISTAS[x.lista].nome, x.produto, x.campo, x.qtd, x.un, x.min === null ? '' : x.min, x.minun, nm[x.status], x.falta || '', x.ant === null ? '' : Math.round((x.qtd - x.ant) * 1000) / 1000];
  });
  s.getRange(1,1,rows.length+1,H.length).setValues([H].concat(rows));
  s.getRange(1,1,1,H.length).setFontWeight('bold').setBackground('#1F6B4A').setFontColor('#ffffff');
  if(rows.length) s.getRange(2,8,rows.length,1).setBackgrounds(res.map(function(x){ return [cor[x.status]]; }));
  s.setFrozenRows(1);
  var linhas = [['Última contagem']].concat(Object.keys(LISTAS).map(function(l){
    var u = r.ult[l];
    return [LISTAS[l].nome + ': ' + (u ? u.data + ' ' + u.hora + ' ' + u.resp : '-')];
  }));
  s.getRange(1,12,linhas.length,1).setValues(linhas);
}
function salvarMin(d){
  if(d.adm !== ADMIN) return {ok:false, erro:'senha'};
  var s = ss().getSheetByName('Mínimos'); if(!s){ s = ss().insertSheet('Mínimos'); proteger(s); }
  s.clear();
  var rows = [['Lista','Produto','Campo','Mínimo','Unidade']].concat((d.lista || []).filter(function(x){
    return Array.isArray(x) && x.length === 5 && LISTAS[x[0]];
  }));
  s.getRange(1,1,rows.length,5).setValues(rows); s.getRange(1,1,1,5).setFontWeight('bold'); s.setFrozenRows(1);
  var r = calcular(); escreverPainel(r); return dadosPainel(r);
}
