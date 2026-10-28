const APP_NAME = '4ª Jornada Tecnológica 2026 — SKA TECH HUB';
const CRITERIOS = [
  ['Ideia','Objetivo claro e alcançável, proposta criativa e coerência entre a ideia e o resultado.'],
  ['Desenvolvimento','Funcionamento das interações, consistência visual, respeito ao espaço e recurso diferenciado.'],
  ['Aprendizado','Domínio do projeto e capacidade de explicar as soluções e os conhecimentos utilizados.'],
  ['Material','Organização e qualidade do PPT e dos recursos de apoio à apresentação.'],
  ['Postura','Organização da apresentação e participação dos integrantes do grupo.'],
  ['Clareza','Explicação compreensível do objetivo, do funcionamento e das soluções do projeto.']
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function setupSistema() {
  const ss = SpreadsheetApp.getActive();
  let cfg = ss.getSheetByName('Config');
  if (!cfg) cfg = ss.insertSheet('Config');
  cfg.clear();
  cfg.getRange('A1:B1').setValues([['CHAVE','VALOR']]).setFontWeight('bold');
  cfg.getRange('A2:B8').setValues([
    ['AVALIADORES_ATIVOS','4'],
    ['AVALIADOR_1','SKA'],
    ['AVALIADOR_2','SENAI'],
    ['AVALIADOR_3','EDUCAÇÃO'],
    ['AVALIADOR_4','CMAIT'],
    ['RESULTADOS_PUBLICOS','NAO']
  ]);
  cfg.autoResizeColumns(1,2);

  let av = ss.getSheetByName('Avaliacoes_Web');
  if (!av) av = ss.insertSheet('Avaliacoes_Web');
  av.clear();
  av.getRange(1,1,1,13).setValues([['Timestamp','Avaliador','ProjetoID','Ideia','Desenvolvimento','Aprendizado','Material','Postura','Clareza','Media','Observacoes','Token','AtualizadoEm']]).setFontWeight('bold');
  av.setFrozenRows(1);

  let fotos = ss.getSheetByName('Fotos_Equipes');
  if (!fotos) fotos = ss.insertSheet('Fotos_Equipes');
  fotos.clear();
  fotos.getRange('A1:C1').setValues([['ProjetoID','URL','FileID']]).setFontWeight('bold');
  fotos.setFrozenRows(1);

  PropertiesService.getDocumentProperties().setProperty('SETUP_OK','1');
  return {ok:true, message:'Sistema configurado. Avaliadores e Organização acessam apenas informando o nome.'};
}

function getBootstrap() {
  ensureSetup_();
  return {
    appName: APP_NAME,
    criterios: CRITERIOS,
    avaliadores: getAvaliadores_().map(x => x.nome),
    projetos: getProjetos_(),
    publicResults: getConfig_('RESULTADOS_PUBLICOS') === 'SIM'
  };
}

function loginAvaliador(nome) {
  ensureSetup_();
  const avaliador = String(nome||'').trim().replace(/\s+/g,' ');
  if (!avaliador) throw new Error('Informe seu nome como avaliador.');
  if (avaliador.length > 80) throw new Error('O nome do avaliador é muito longo.');
  return {ok:true, nome:avaliador, token: makeToken_(avaliador)};
}

function loginAdmin() {
  ensureSetup_();
  return {ok:true, nome:'Organização', token:makeToken_('ADMIN')};
}

function getPainelAvaliador(token) {
  const who = verifyToken_(token);
  if (!who || who === 'ADMIN') throw new Error('Sessão inválida.');
  const projetos = getProjetos_();
  const avaliacoes = getAvaliacoes_().filter(x => x.avaliador === who);
  const map = {};
  avaliacoes.forEach(x => map[x.projetoId] = x);
  return {avaliador:who, projetos:projetos.map(p => ({...p, avaliacao:map[p.id] || null}))};
}

function getAvaliacao(token, projetoId) {
  const who = verifyToken_(token);
  if (!who || who === 'ADMIN') throw new Error('Sessão inválida.');
  const p = getProjetos_().find(x => x.id === projetoId);
  if (!p) throw new Error('Projeto não encontrado.');
  const a = getAvaliacoes_().find(x => x.avaliador === who && x.projetoId === projetoId) || null;
  return {projeto:p, avaliacao:a, criterios:CRITERIOS};
}

function salvarAvaliacao(token, payload) {
  const who = verifyToken_(token);
  if (!who || who === 'ADMIN') throw new Error('Sessão inválida.');
  const id = String(payload.projetoId||'').trim();
  if (!getProjetos_().some(x => x.id === id)) throw new Error('Projeto não encontrado.');
  const notas = [payload.ideia,payload.desenvolvimento,payload.aprendizado,payload.material,payload.postura,payload.clareza].map(Number);
  if (notas.some(n => !Number.isFinite(n) || n < 0 || n > 10)) throw new Error('Todas as notas devem estar entre 0 e 10.');
  const media = notas.reduce((a,b)=>a+b,0)/notas.length;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Avaliacoes_Web');
  const data = sh.getDataRange().getValues();
  let row = -1;
  for (let i=1;i<data.length;i++) if (String(data[i][1])===who && String(data[i][2])===id) { row=i+1; break; }
  const now = new Date();
  const values = [now, who, id, ...notas, media, String(payload.observacoes||'').trim(), token, now];
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    if (row > 0) sh.getRange(row,1,1,13).setValues([values]);
    else sh.appendRow(values);
  } finally { lock.releaseLock(); }
  return {ok:true, media:Number(media.toFixed(2))};
}

function getAdminDashboard(token) {
  if (verifyToken_(token) !== 'ADMIN') throw new Error('Sessão de administrador inválida.');
  return buildDashboard_();
}

function excluirAvaliacao(token, projetoId, avaliador) {
  if (verifyToken_(token) !== 'ADMIN') throw new Error('Sessão de administrador inválida.');
  const id = String(projetoId || '').trim();
  const nome = String(avaliador || '').trim();
  if (!id || !nome) throw new Error('Avaliação inválida.');
  const sh = SpreadsheetApp.getActive().getSheetByName('Avaliacoes_Web');
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === nome && String(data[i][2]).trim() === id) {
      sh.deleteRow(i + 1);
      return {ok:true};
    }
  }
  throw new Error('Avaliação não encontrada.');
}

function getResultadosPublicos() {
  if (getConfig_('RESULTADOS_PUBLICOS') !== 'SIM') throw new Error('Os resultados públicos ainda não foram liberados.');
  const d = buildDashboard_();
  return {ranking:d.ranking, concluido:d.concluido, total:d.total, podium:d.podium};
}

function setResultadosPublicos(token, liberar) {
  if (verifyToken_(token) !== 'ADMIN') throw new Error('Sessão de administrador inválida.');
  setConfig_('RESULTADOS_PUBLICOS', liberar ? 'SIM':'NAO');
  return {ok:true};
}

function atualizarConfig(token, data) {
  if (verifyToken_(token) !== 'ADMIN') throw new Error('Sessão de administrador inválida.');
  const avaliadores = Array.isArray(data.avaliadores) ? data.avaliadores : [];
  avaliadores.slice(0,10).forEach((a,i) => {
    const nome=String((a && a.nome) || a || '').trim();
    if (nome) setConfig_('AVALIADOR_'+(i+1), nome);
  });
  if (avaliadores.length) setConfig_('AVALIADORES_ATIVOS', String(Math.min(10,avaliadores.length)));
  return {ok:true};
}

function uploadFotoEquipe(token, projetoId, fileName, mimeType, base64Data) {
  if (verifyToken_(token) !== 'ADMIN') throw new Error('Sessão de administrador inválida.');
  const p = getProjetos_().find(x => x.id === projetoId);
  if (!p) throw new Error('Projeto não encontrado.');
  if (!/^image\//.test(mimeType||'')) throw new Error('Envie uma imagem válida.');
  const bytes = Utilities.base64Decode(String(base64Data).replace(/^data:[^;]+;base64,/,''));
  if (bytes.length > 5 * 1024 * 1024) throw new Error('A imagem deve ter no máximo 5 MB.');
  const props = PropertiesService.getDocumentProperties();
  let folderId = props.getProperty('FOTOS_FOLDER_ID');
  let folder;
  if (folderId) { try { folder=DriveApp.getFolderById(folderId); } catch(e){} }
  if (!folder) {
    folder = DriveApp.createFolder('SKA TECH HUB - Fotos Equipes - '+new Date().getFullYear());
    props.setProperty('FOTOS_FOLDER_ID', folder.getId());
  }
  const blob = Utilities.newBlob(bytes, mimeType, projetoId+'_'+String(fileName||'foto').replace(/[^\w.\-]+/g,'_'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url='https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w1200';
  const sh=SpreadsheetApp.getActive().getSheetByName('Fotos_Equipes');
  const data=sh.getDataRange().getValues(); let row=-1;
  for(let i=1;i<data.length;i++) if(String(data[i][0])===projetoId){row=i+1;break;}
  const vals=[projetoId,url,file.getId()];
  if(row>0) sh.getRange(row,1,1,3).setValues([vals]); else sh.appendRow(vals);
  return {ok:true,url:url};
}

function buildDashboard_() {
  const projetos=getProjetos_(); const avs=getAvaliacoes_();
  const esperadas=Math.max(1,Math.min(10,Number(getConfig_('AVALIADORES_ATIVOS')||4)));
  const nomesAtivos=[...new Set(avs.map(a=>String(a.avaliador||'').trim()).filter(Boolean))];
  const rows=projetos.map((p,idx)=>{
    const porAvaliador={};
    avs.filter(a=>a.projetoId===p.id).forEach(a=>porAvaliador[a.avaliador]=a);
    const aa=Object.values(porAvaliador);
    const medias=aa.map(a=>Number(a.media)).filter(Number.isFinite);
    const parcial=medias.length ? medias.reduce((a,b)=>a+b,0)/medias.length : null;
    const final=medias.length>=esperadas ? medias.slice(0,esperadas).reduce((a,b)=>a+b,0)/esperadas : null;
    const detalhes=aa.map(a=>({
      avaliador:a.avaliador, ideia:a.ideia, desenvolvimento:a.desenvolvimento, aprendizado:a.aprendizado,
      material:a.material, postura:a.postura, clareza:a.clareza, media:Number(a.media), observacoes:a.observacoes||''
    }));
    return {...p, recebidas:medias.length, esperadas:esperadas, mediaParcial:parcial===null?null:Number(parcial.toFixed(4)), mediaFinal:final===null?null:Number(final.toFixed(4)), avaliacoes:detalhes, _idx:idx};
  });
  const completos=rows.filter(x=>x.mediaFinal!==null).sort((a,b)=> b.mediaFinal-a.mediaFinal || a._idx-b._idx);
  let prev=null, pos=0;
  completos.forEach((x,i)=>{ if(prev===null || x.mediaFinal!==prev) pos=i+1; x.colocacao=pos; prev=x.mediaFinal; });
  const ranking=completos.map(({_idx,...x})=>x);
  const atuais=rows.filter(x=>x.mediaParcial!==null).sort((a,b)=> b.mediaParcial-a.mediaParcial || a._idx-b._idx);
  let prevAtual=null, posAtual=0;
  atuais.forEach((x,i)=>{ if(prevAtual===null || x.mediaParcial!==prevAtual) posAtual=i+1; x.colocacaoAtual=posAtual; prevAtual=x.mediaParcial; });
  const rankingAtual=atuais.map(({_idx,...x})=>x);
  return {total:projetos.length, concluido:completos.length, pendente:projetos.length-completos.length, avaliadores:nomesAtivos, esperadas:esperadas, projetos:rows.map(({_idx,...x})=>x), ranking, podium:ranking.filter(x=>x.colocacao<=3), rankingAtual, podiumAtual:rankingAtual.slice(0,3)};
}

function getProjetos_() {
  const sh=SpreadsheetApp.getActive().getSheetByName('Projetos');
  if(!sh) throw new Error('A aba Projetos não foi encontrada.');
  const fotos=getFotos_();
  const data=sh.getRange(9,2,Math.max(1,sh.getLastRow()-8),4).getValues();
  return data.filter(r=>String(r[1]||'').trim()).map(r=>({id:String(r[0]||'').trim(), nome:String(r[1]||'').trim(), alunos:String(r[2]||'').trim(), turma:String(r[3]||'').trim(), foto:fotos[String(r[0]||'').trim()]||''}));
}

function getFotos_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Fotos_Equipes'); if(!sh || sh.getLastRow()<2) return {};
  const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues(); const o={}; vals.forEach(r=>{if(r[0]&&r[1])o[String(r[0])]=String(r[1]);}); return o;
}
function getAvaliacoes_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Avaliacoes_Web'); if(!sh || sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,13).getValues().filter(r=>r[1]&&r[2]).map(r=>({timestamp:r[0],avaliador:String(r[1]),projetoId:String(r[2]),ideia:Number(r[3]),desenvolvimento:Number(r[4]),aprendizado:Number(r[5]),material:Number(r[6]),postura:Number(r[7]),clareza:Number(r[8]),media:Number(r[9]),observacoes:String(r[10]||'')}));
}
function getAvaliadores_(){
  const n=Math.max(1,Math.min(10,Number(getConfig_('AVALIADORES_ATIVOS')||4))); const arr=[];
  for(let i=1;i<=n;i++){ const raw=getConfig_('AVALIADOR_'+i); if(raw){const nome=String(raw).split('|')[0].trim(); if(nome) arr.push({nome:nome});}}
  return arr;
}
function ensureSetup_(){
  if(!SpreadsheetApp.getActive().getSheetByName('Config') || !SpreadsheetApp.getActive().getSheetByName('Avaliacoes_Web')) throw new Error('Execute a função setupSistema() uma vez no Apps Script antes de publicar.');
}
function getConfig_(key){ const sh=SpreadsheetApp.getActive().getSheetByName('Config'); if(!sh)return ''; const d=sh.getDataRange().getValues(); for(let i=1;i<d.length;i++)if(String(d[i][0])===key)return String(d[i][1]||''); return ''; }
function setConfig_(key,val){ const sh=SpreadsheetApp.getActive().getSheetByName('Config'); const d=sh.getDataRange().getValues(); for(let i=1;i<d.length;i++)if(String(d[i][0])===key){sh.getRange(i+1,2).setValue(val);return;} sh.appendRow([key,val]); }
function makeToken_(who){ const raw=who+'|'+Date.now()+'|'+Utilities.getUuid(); CacheService.getUserCache().put('t_'+Utilities.base64EncodeWebSafe(raw),who,21600); return Utilities.base64EncodeWebSafe(raw); }
function verifyToken_(token){ if(!token)return null; return CacheService.getUserCache().get('t_'+token); }
