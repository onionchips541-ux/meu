// Cole em: planilha > Extensões > Apps Script. Depois de colar, ajuste CHAVE (igual à do app) e, se quiser, ADMIN.
var CHAVE = 'troque-esta-senha';
var ADMIN = 'oninsadm123';   // senha do painel administrador (fica só aqui, não vai dentro do app)
var MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
var CAB = ['Data','Hora do envio','Responsável','Dia da contagem','Categoria','Produto','Campo','Quantidade','Unidade','Observação'];

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }

// Deixa a aba editável só pelo dono. Rode protegerAbas() uma vez no editor se já criou as abas antes.
function proteger(s){ try{ var p = s.protect().setDescription('Somente o dono'); p.removeEditors(p.getEditors()); p.setDomainEdit(false); }catch(e){} }
function protegerAbas(){ ['Config','Mínimos','Painel'].forEach(function(n){ var s = ss().getSheetByName(n); if(s) proteger(s); }); }

// Aba Config: B1 = SIM/NAO (liga/desliga o app), B2 = mensagem de bloqueio, B3 = % para "perto do mínimo"
function cfg(){
  var s = ss().getSheetByName('Config');
  if(!s){ s = ss().insertSheet('Config');
    s.getRange('A1:B3').setValues([['ativo','SIM'],['mensagem','Acesso encerrado. Fale com a gerência.'],['alerta_pct',25]]); proteger(s); }
  var v = s.getRange('B1:B3').getValues(), p = Number(v[2][0]);
  return { ativo: String(v[0][0]).trim().toUpperCase() === 'SIM', msg: String(v[1][0]), pct: (v[2][0]===''||isNaN(p)) ? 25 : p };
}

function doGet(e){
  var p = e.parameter || {};
  if(p.chave !== CHAVE) return out({ativo:false, msg:'Chave inválida'});
  var c = cfg();
  if(p.acao === 'painel'){
    if(!c.ativo) return out({ok:false, erro:'bloqueado'});
    if(p.adm !== ADMIN) return out({ok:false, erro:'senha'});
    return out(dadosPainel(calcular()));
  }
  return out({ativo:c.ativo, msg:c.msg});
}

function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    var d = JSON.parse(e.postData.contents);
    if(d.chave !== CHAVE) return out({ok:false, erro:'chave'});
    if(!cfg().ativo) return out({ok:false, erro:'bloqueado'});
    lock.waitLock(20000);
    if(d.acao === 'minimos') return out(salvarMin(d));
    var p = d.data.split('-'), dt = p[2]+'/'+p[1]+'/'+p[0];
    var nome = (d.lista==='frente' ? 'Frente ' : 'Cozinha ') + MESES[+p[1]-1] + ' ' + p[0];
    var s = ss().getSheetByName(nome);
    if(!s){ s = ss().insertSheet(nome); s.appendRow(CAB);
      s.getRange(1,1,1,CAB.length).setFontWeight('bold').setBackground('#1F6B4A').setFontColor('#ffffff');
      s.setFrozenRows(1); s.setColumnWidth(6,240); }
    var n = s.getLastRow();   // reenvio da mesma pessoa, data e dia substitui o anterior
    if(n > 1){
      var v = s.getRange(2,1,n-1,4).getDisplayValues(), run = 0;
      for(var i=v.length-1; i>=-1; i--){
        var m = i>=0 && v[i][0]===dt && v[i][2]===d.resp && v[i][3]===d.dia;
        if(m) run++; else if(run){ s.deleteRows(i+3, run); run = 0; }
      }
    }
    var rows = d.linhas.map(function(l){ return [dt,d.hora,d.resp,d.dia,l[0],l[1],l[2],l[3],l[4],l[5]]; });
    s.getRange(s.getLastRow()+1, 1, rows.length, 4).setNumberFormat('@');
    s.getRange(s.getLastRow()+1, 1, rows.length, CAB.length).setValues(rows);
    try{ escreverPainel(calcular()); }catch(x){}
    return out({ok:true, linhas:rows.length, aba:nome});
  }catch(err){ return out({ok:false, erro:String(err)}); }
  finally{ try{ lock.releaseLock(); }catch(x){} }
}

// ---------- Painel, mínimos e consumo ----------
function contagens(lista){   // todas as contagens da lista, da mais antiga para a mais nova
  var m = {}, pre = lista==='frente' ? 'Frente ' : 'Cozinha ';
  ss().getSheets().forEach(function(s){
    if(s.getName().indexOf(pre)!==0 || s.getLastRow()<2) return;
    s.getRange(2,1,s.getLastRow()-1,10).getValues().forEach(function(r){
      var p = String(r[0]).split('/'), k = p[2]+p[1]+p[0]+' '+r[1]+' '+r[2]+' '+r[3];
      (m[k] = m[k] || {info:{data:String(r[0]),hora:String(r[1]),resp:r[2],dia:r[3]}, rows:[]}).rows.push(r);
    });
  });
  return Object.keys(m).sort().map(function(k){ return m[k]; });
}
function lerMin(){ var s = ss().getSheetByName('Mínimos'); return (!s || s.getLastRow()<2) ? [] : s.getRange(2,1,s.getLastRow()-1,5).getValues(); }
function norm(q,u){ var f = {g:['g',1],kg:['g',1000],ml:['ml',1],l:['ml',1000]}[String(u||'').toLowerCase()]; return f ? [q*f[1],f[0]] : [q,'un.']; }

function calcular(){
  var mins = {}, pct = cfg().pct, res = [], ult = {};
  lerMin().forEach(function(m){ if(m[3]!=='') mins[m[0]+'|'+m[1]+'|'+m[2]] = [Number(m[3]), m[4]]; });
  ['cozinha','frente'].forEach(function(l){
    var c = contagens(l); if(!c.length) return;
    var u = c[c.length-1], prev = {}; ult[l] = u.info;
    if(c.length>1) c[c.length-2].rows.forEach(function(r){ prev[r[5]+'|'+r[6]] = r[7]; });
    u.rows.forEach(function(r){
      var k = r[5]+'|'+r[6], q = r[7]; if(typeof q !== 'number') return;
      var it = {lista:l, produto:r[5], campo:r[6], qtd:q, un:r[8], min:null, minun:'', status:'sem', falta:0,
                ant:(typeof prev[k]==='number') ? prev[k] : null};
      var mn = mins[l+'|'+k];
      if(mn){
        var A = norm(q,r[8]), B = norm(mn[0],mn[1]);
        it.min = mn[0]; it.minun = mn[1];
        if(A[1] !== B[1]) it.status = 'unidade';
        else{ it.status = A[0] <= B[0] ? 'abaixo' : (A[0] <= B[0]*(1+pct/100) ? 'proximo' : 'ok');
              it.falta = Math.max(0, Math.round((B[0]-A[0])/norm(1,mn[1])[0]*1000)/1000); }
      }
      res.push(it);
    });
  });
  return {res:res, ult:ult};
}
function dadosPainel(r){
  return {ok:true, ult:r.ult, minimos:lerMin(),
          itens:r.res.filter(function(x){ return x.status==='abaixo' || x.status==='proximo'; })};
}
function escreverPainel(r){
  var s = ss().getSheetByName('Painel'); if(!s){ s = ss().insertSheet('Painel'); proteger(s); }
  s.clear();
  var ord = {abaixo:0,proximo:1,unidade:2,ok:3,sem:4}, nm = {abaixo:'NO MÍNIMO/ABAIXO',proximo:'PERTO DO MÍNIMO',ok:'OK',sem:'sem mínimo',unidade:'conferir unidade'};
  var cor = {abaixo:'#F4C7C3',proximo:'#FCE8B2',ok:'#D9EAD3',sem:'#FFFFFF',unidade:'#EEEEEE'};
  var res = r.res.slice().sort(function(a,b){ return ord[a.status]-ord[b.status]; });
  var H = ['Lista','Produto','Campo','Qtd atual','Unidade','Mínimo','Un. mín.','Status','Falta p/ mínimo','Variação desde a contagem anterior (negativo = consumo)'];
  var rows = res.map(function(x){ return [x.lista,x.produto,x.campo,x.qtd,x.un,x.min===null?'':x.min,x.minun,nm[x.status],x.falta||'',x.ant===null?'':Math.round((x.qtd-x.ant)*1000)/1000]; });
  s.getRange(1,1,rows.length+1,H.length).setValues([H].concat(rows));
  s.getRange(1,1,1,H.length).setFontWeight('bold').setBackground('#1F6B4A').setFontColor('#ffffff');
  if(rows.length) s.getRange(2,8,rows.length,1).setBackgrounds(res.map(function(x){ return [cor[x.status]]; }));
  s.setFrozenRows(1);
  s.getRange('L1:L3').setValues([['Última contagem'],['Cozinha: '+(r.ult.cozinha ? r.ult.cozinha.data+' '+r.ult.cozinha.hora+' '+r.ult.cozinha.resp : '-')],['Frente: '+(r.ult.frente ? r.ult.frente.data+' '+r.ult.frente.hora+' '+r.ult.frente.resp : '-')]]);
}
function salvarMin(d){
  if(d.adm !== ADMIN) return {ok:false, erro:'senha'};
  var s = ss().getSheetByName('Mínimos'); if(!s){ s = ss().insertSheet('Mínimos'); proteger(s); }
  s.clear();
  var rows = [['Lista','Produto','Campo','Mínimo','Unidade']].concat((d.lista||[]).filter(function(x){ return Array.isArray(x) && x.length===5; }));
  s.getRange(1,1,rows.length,5).setValues(rows); s.getRange(1,1,1,5).setFontWeight('bold'); s.setFrozenRows(1);
  var r = calcular(); escreverPainel(r); return dadosPainel(r);
}
