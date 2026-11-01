// ============ PORTAL DE ESPORTES - SERVIDOR (Google Apps Script) ============
// Instalação / atualização:
// 1) Planilha principal > Extensões > Apps Script > substitua TODO o código por este.
// 2) Crie uma planilha NOVA e PRIVADA só para inscrições (ex.: "Inscrições SMEL"). Copie o ID dela (trecho da URL entre /d/ e /edit).
// 3) Configurações do projeto > Propriedades do script (opcionais):
//      INSCRICOES_ID = (ID da planilha de inscrições)   [se vazio, usa a planilha principal]
//      PASTA_HOME    = (link ou ID da pasta "Página principal" do Drive)
//      PASTA_EVENTOS = (link ou ID da pasta mãe onde serão criadas as pastas dos eventos; se vazio, usa a pasta da planilha)
// 4) Execute a função "autorizar" (aceite as permissões). Ela cria as abas e configura as pastas.
// 5) Implantar > Gerenciar implantações > lápis > Nova versão > Implantar. (A URL /exec continua a mesma.)
// Planilha principal: eventos | feed | colocacoes | usuarios | auditoria
// Planilha de inscrições: "indice" + uma aba por evento ("i_<id do evento>"), criadas sozinhas.
const ABAS = {
  eventos: ["id","nome","modalidade","categoria","local_evento","data_inicio","data_fim","prazo_inscricao","status","min_atletas","max_atletas","max_vagas","descricao","regulamento_url","link_challonge_embed","pasta_fotos_url","imagem_url","aceita_individual","pasta_evento_url"],
  feed: ["data","titulo","legenda","imagem_url","categoria"],
  colocacoes: ["evento_id","colocacao","nome_equipe"],
  usuarios: ["login","senha"],
  auditoria: ["data_hora","servidor","acao","detalhes"]
};
const INSC = ["data_envio","codigo","protocolo","time","instagram_time","nome","idade","telefone","cpf","email","rede_social","status_aprovacao","origem","termo","evento_id","evento"];
const IDX = ["codigo","evento_id","protocolo"];

function autorizar() {
  DriveApp.getRootFolder();
  Object.keys(ABAS).forEach(aba);
  indice();
  eventos().forEach(e => abaInsc(e.id, true));
}

// ---- Trilha de Auditoria (Logs de Atividade Oficial) ----
function registrarAuditoria(servidor, acao, detalhes) {
  try {
    const sh = aba("auditoria");
    const tz = Session.getScriptTimeZone();
    const dh = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy 'às' HH:mm:ss");
    sh.appendRow([dh, cel(servidor || "sistema"), cel(acao || ""), cel(detalhes || "")]);
  } catch (e) {
    console.error("Erro ao gravar auditoria:", e);
  }
}

// ---- Rate Limit (Proteção contra Brute Force e Spam) ----
function verificarRateLimit(chave, limite, janelaSeg) {
  const c = CacheService.getScriptCache();
  const atual = +(c.get(chave) || 0);
  if (atual >= limite) return false;
  return true;
}
function registrarFalhaRateLimit(chave, janelaSeg) {
  const c = CacheService.getScriptCache();
  const atual = +(c.get(chave) || 0) + 1;
  c.put(chave, String(atual), janelaSeg || 900);
  return atual;
}
function limparRateLimit(chave) {
  CacheService.getScriptCache().remove(chave);
}

function doPost(e) {
  let out;
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.acao === "publico") out = publico();
    else if (p.acao === "fotos") out = {ok:1, fotos: pastaDoEvento(p.pasta) ? fotosPasta(p.pasta, 60) : []};
    else if (p.acao === "inscrever") out = inscrever(p);
    else if (p.acao === "minha") out = minha(p);
    else if (p.acao === "editar") out = editar(p);
    else if (p.acao === "login") out = login(p);
    else out = admin(p);
  } catch (err) { out = {ok:0, msg:"Erro interno. Tente novamente."}; console.error(err); }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---- planilhas ----
function abaSh(ss, n, cols) {
  let sh = ss.getSheetByName(n);
  if (!sh) { sh = ss.insertSheet(n); sh.appendRow(cols); return sh; }
  const h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  const falta = cols.filter(k => h.indexOf(k) < 0);
  if (falta.length && h.some(Boolean)) sh.getRange(1, h.length + 1, 1, falta.length).setValues([falta]);
  return sh;
}
const aba = n => abaSh(SpreadsheetApp.getActive(), n, ABAS[n]);
function ssInsc() { const id = PropertiesService.getScriptProperties().getProperty("INSCRICOES_ID"); return id ? SpreadsheetApp.openById(id.trim()) : SpreadsheetApp.getActive(); }
const nomeAba = id => "i_" + String(id).replace(/[\[\]*?:\/\\]/g, "-").slice(0, 90);
function abaInsc(id, criar) { const ss = ssInsc(), n = nomeAba(id); return (criar || ss.getSheetByName(n)) ? abaSh(ss, n, INSC) : null; }
const indice = () => abaSh(ssInsc(), "indice", IDX);
const cabecalho = sh => sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
function lerSh(sh, cols) {
  const v = sh.getDataRange().getValues(), h = v.shift().map(String), tz = Session.getScriptTimeZone();
  return v.map((r,i) => { const o = {_row:i+2};
    h.forEach((k,j) => { let x = r[j]; if (x instanceof Date) x = Utilities.formatDate(x, tz, "yyyy-MM-dd"); o[k] = String(x); });
    cols.forEach(k => { if (o[k] === undefined) o[k] = ""; });
    return o; }).filter(o => cols.some(k => o[k] !== ""));
}
const ler = n => lerSh(aba(n), ABAS[n]);
function lerInsc(id) { const sh = abaInsc(id, false); return sh ? lerSh(sh, INSC) : []; }
const cel = x => { x = x == null ? "" : String(x); return /^[=+@-]/.test(x) ? "'" + x : x; };
const hoje = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
const slug = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
function idUnico(nome, data, usados) {
  const b = slug(nome) + (data ? "-" + String(data).slice(0,4) : ""); let id = b, n = 2;
  while (usados[id]) id = b + "-" + n++; usados[id] = 1; return id;
}

// Eventos: gera o "id" se faltar e encerra automaticamente eventos cuja data já passou.
function eventos() {
  const sh = aba("eventos"), ev = ler("eventos"), h = cabecalho(sh);
  const colId = h.indexOf("id") + 1, colSt = h.indexOf("status") + 1, u = {}, hj = hoje();
  ev.forEach(e => { if (e.id) u[e.id] = 1; });
  ev.forEach(e => {
    if (!e.id && e.nome) {
      e.id = idUnico(e.nome, e.data_inicio, u);
      if (colId > 0) sh.getRange(e._row, colId).setValue(e.id);
    }
    const fim = String(e.data_fim || e.data_inicio || "").trim().slice(0, 10);
    if (e.aceita_individual && (/no[ _]?evento/i.test(e.aceita_individual) || /esgotad/i.test(e.aceita_individual)) && !/^encerr/i.test(e.status || "")) {
      e.status = /no[ _]?evento/i.test(e.aceita_individual) ? "Inscrição feita apenas no evento" : "Vagas esgotadas";
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(fim) && fim < hj && !/^encerr/i.test(e.status || "")) {
      e.status = "Encerrado";
      if (colSt > 0 && e._row) {
        try { sh.getRange(e._row, colSt).setValue("Encerrado"); } catch (err) {}
      }
    }
    e.max_atletas = e.max_atletas || e.max_por_time || "";
    e.max_vagas = e.max_vagas || e.vagas_max || e.max_inscritos || "";
  });
  return ev.filter(e => e.id);
}

// Vaga = 1 time (evento por equipe) ou 1 atleta (evento individual, máx. 1 atleta). "r" = inscrições do evento, sem as recusadas.
const vagas = (ev, r) => +ev.max_atletas > 1 ? new Set(r.map(i => (i.time || "").trim().toLowerCase())).size : r.length;
const ativas = id => lerInsc(id).filter(i => i.status_aprovacao !== "Recusada");
const aberto = ev => {
  if (!ev || !ev.id || !/^aberto/i.test(ev.status)) return false;
  const hj = hoje(), fim = String(ev.data_fim || ev.data_inicio || "").trim().slice(0, 10);
  if (fim && fim < hj) return false;
  if (ev.prazo_inscricao && hj > ev.prazo_inscricao) return false;
  return true;
};
function cpfOk(c) {
  c = String(c).replace(/\D/g, "");
  if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return false;
  for (let t = 9; t < 11; t++) { let s = 0; for (let i = 0; i < t; i++) s += +c[i] * (t + 1 - i); if ((s * 10) % 11 % 10 !== +c[t]) return false; }
  return true;
}
function protocolo() {
  const pr = PropertiesService.getScriptProperties(), a = new Date().getFullYear(), k = "SEQ_" + a, n = (+pr.getProperty(k) || 0) + 1;
  pr.setProperty(k, String(n)); return "PROTOCOLO-" + a + "-" + String(n).padStart(4, "0");
}
// Código de 6 caracteres (letras e números). Começa com letra para a planilha não virar número.
function novoCodigo() {
  const u = {}; indice().getDataRange().getValues().forEach(r => { u[String(r[0]).toUpperCase()] = 1; });
  const L = "ABCDFGHJKLMNPQRSTUVWXYZ", A = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let c;
  do { c = L[Math.floor(Math.random() * L.length)]; for (let i = 1; i < 6; i++) c += A[Math.floor(Math.random() * A.length)]; } while (u[c]);
  return c;
}

// ---- Pastas e Fotos no Google Drive ----
function idPasta(u) { const m = String(u || "").match(/folders\/([\w-]+)|[?&]id=([\w-]+)|^([\w-]{15,})$/); return m ? (m[1] || m[2] || m[3]) : ""; }
function pastaDoEvento(u) { const id = idPasta(u); return id && eventos().some(e => idPasta(e.pasta_fotos_url) === id); }

function pastaMaeEventos() {
  try {
    const prop = PropertiesService.getScriptProperties().getProperty("PASTA_EVENTOS");
    const pid = idPasta(prop);
    if (pid) return DriveApp.getFolderById(pid);
    const pais = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
    if (pais.hasNext()) return pais.next();
  } catch (e) {}
  return DriveApp.getRootFolder();
}

function obterOuCriarSubpasta(pai, nome) {
  const limpo = String(nome || "Evento").trim().replace(/[\/\\]/g, "-") || "Evento";
  const it = pai.getFoldersByName(limpo);
  if (it.hasNext()) return it.next();
  return pai.createFolder(limpo);
}

// Cria automaticamente a pasta do evento e dentro dela as subpastas "capa evento" e "Fotos"
function garantirPastasEvento(nomeEvento, pastaEventoUrlExistente) {
  let pastaEv = null;
  const idEv = idPasta(pastaEventoUrlExistente);
  if (idEv) {
    try { pastaEv = DriveApp.getFolderById(idEv); } catch (e) { pastaEv = null; }
  }
  if (!pastaEv) {
    pastaEv = obterOuCriarSubpasta(pastaMaeEventos(), nomeEvento || "Evento SMEL");
  } else if (nomeEvento && pastaEv.getName() !== nomeEvento) {
    try { pastaEv.setName(nomeEvento); } catch (e) {}
  }
  const pastaCapa = obterOuCriarSubpasta(pastaEv, "capa evento");
  const pastaFotos = obterOuCriarSubpasta(pastaEv, "Fotos");
  try { pastaCapa.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  try { pastaFotos.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return {
    pastaEv,
    pastaCapa,
    pastaFotos,
    eventoUrl: "https://drive.google.com/drive/folders/" + pastaEv.getId(),
    fotosUrl: "https://drive.google.com/drive/folders/" + pastaFotos.getId()
  };
}

function salvarBase64NoDrive(pasta, dataUrl, nomePadrao, substituirExistentes) {
  const m = String(dataUrl || "").match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) return null;
  const mime = m[1], bytes = Utilities.base64Decode(m[2]);
  const ext = (mime.split("/")[1] || "jpg").replace("jpeg", "jpg");
  const nomeArq = (nomePadrao ? String(nomePadrao).replace(/\.[^.]+$/, "") : "imagem-" + Date.now()) + "." + ext;
  if (substituirExistentes) {
    const ant = pasta.getFiles();
    while (ant.hasNext()) {
      try { ant.next().setTrashed(true); } catch (e) {}
    }
  }
  const f = pasta.createFile(Utilities.newBlob(bytes, mime, nomeArq));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return {
    id: f.getId(),
    thumb: "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w1200",
    link: "https://drive.google.com/file/d/" + f.getId() + "/view"
  };
}

function buscarCapaNaPasta(pastaCapa) {
  try {
    const it = pastaCapa.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/^image\//.test(f.getMimeType())) {
        return "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w1200";
      }
    }
  } catch (e) {}
  return "";
}

function fotosPasta(u, limite) {
  const id = idPasta(u); if (!id) return [];
  const max = limite || 4;
  const c = CacheService.getScriptCache(), k = "f" + id + "_" + max, h = c.get(k); if (h) return JSON.parse(h);
  let r = [];
  try {
    const arq = [], it = DriveApp.getFolderById(id).getFiles();
    while (it.hasNext()) { const f = it.next(); if (/^image\//.test(f.getMimeType())) arq.push({nome:f.getName(), id:f.getId()}); }
    arq.sort((a,b) => a.nome.localeCompare(b.nome, undefined, {numeric:true}));
    r = arq.slice(0, max).map(f => ({nome:f.nome, thumb:"https://drive.google.com/thumbnail?id=" + f.id + "&sz=w800", link:"https://drive.google.com/file/d/" + f.id + "/view"}));
  } catch (e) { return []; }
  try { c.put(k, JSON.stringify(r), 180); } catch (e) {}
  return r;
}

// ---- leitura pública (sem dados pessoais; só quantas vagas já foram ocupadas) ----
function publico() {
  const c = CacheService.getScriptCache(), k = c.get("pub"); if (k) return JSON.parse(k);
  const lim = r => { delete r._row; return r; };
  const o = {
    eventos: eventos().map(e => { e.inscritos = String(aberto(e) && +e.max_vagas > 0 ? vagas(e, ativas(e.id)) : 0); return lim(e); }),
    feed: ler("feed").map(lim), colocacoes: ler("colocacoes").map(lim),
    home_fotos: fotosPasta(PropertiesService.getScriptProperties().getProperty("PASTA_HOME"), 4)
  };
  try { c.put("pub", JSON.stringify(o), 60); } catch (e) {}
  return o;
}

// ---- validação oficial (servidor) ----
function checar(ev, p, ins, o) {
  o = o || {};
  const mx = Math.max(1, +ev.max_atletas || 1), mn = Math.min(mx, Math.max(1, +ev.min_atletas || 1));
  const at = (p.atletas || []).filter(a => a && a.nome), time = String(p.time || "").trim(), t = time.toLowerCase();
  if (mx > 1 && !time) return {erro:"Informe o nome do time."};
  if (at.length < mn || at.length > mx) return {erro:"Este evento exige de " + mn + " a " + mx + " atleta(s) por inscrição."};
  for (const a of at) {
    const fone = String(a.telefone).replace(/\D/g, ""), idade = +a.idade;
    if (!cpfOk(a.cpf)) return {erro:"CPF inválido: " + a.cpf};
    if (!(idade > 0 && idade < 100) || fone.length < 10 || fone.length > 11 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a.email)) return {erro:"Dados incompletos ou inválidos para " + a.nome + "."};
  }
  const cpfs = new Set(ins.map(i => i.cpf.replace(/\D/g, ""))), novos = at.map(a => String(a.cpf).replace(/\D/g, ""));
  if (new Set(novos).size !== novos.length) return {erro:"Há CPF repetido na lista de atletas."};
  const dup = at.find((a,i) => cpfs.has(novos[i]));
  if (dup) return {erro:"O CPF " + dup.cpf + " já está inscrito neste evento."};
  const doTime = mx > 1 ? ins.filter(i => (i.time || "").trim().toLowerCase() === t) : [];
  if (mx > 1 && doTime.length + at.length > mx) return {erro:"Esse time já tem atletas inscritos e não comporta mais " + at.length + "."};
  if (!o.editando && !o.admin && +ev.max_vagas > 0 && vagas(ev, ins) + (mx > 1 ? (doTime.length ? 0 : 1) : at.length) > +ev.max_vagas) return {erro:"Vagas esgotadas."};
  return {at, time};
}

function gravar(ev, v, o) {
  const sh = abaInsc(ev.id, true), h = cabecalho(sh), col = k => h.indexOf(k) + 1, rows = o.rows || [];
  const campos = ["nome","idade","telefone","cpf","email","rede_social"];
  v.at.forEach((a, i) => { if (!rows[i]) return; const r = rows[i]._row;
    campos.forEach(k => { if (col(k)) sh.getRange(r, col(k)).setValue(cel(a[k])); });
    if (col("time")) sh.getRange(r, col("time")).setValue(cel(v.time));
    if (col("instagram_time")) sh.getRange(r, col("instagram_time")).setValue(cel(v.igt));
    if (o.status && col("status_aprovacao")) sh.getRange(r, col("status_aprovacao")).setValue(o.status); });
  for (let i = rows.length - 1; i >= v.at.length; i--) sh.deleteRow(rows[i]._row);
  if (v.at.length > rows.length) {
    const b = rows[0] ? {data_envio:rows[0].data_envio, codigo:rows[0].codigo, protocolo:rows[0].protocolo, status_aprovacao:o.status || rows[0].status_aprovacao, origem:rows[0].origem, termo:rows[0].termo}
                      : {data_envio:new Date(), codigo:o.codigo, protocolo:o.protocolo, status_aprovacao:o.status || "Pendente", origem:o.origem, termo:o.termo};
    b.evento_id = ev.id; b.evento = ev.nome; b.time = cel(v.time); b.instagram_time = cel(v.igt);
    const novas = v.at.slice(rows.length).map(a => { const d = Object.assign({}, b); campos.forEach(k => { d[k] = cel(a[k]); }); return h.map(k => d[k] === undefined ? "" : d[k]); });
    sh.getRange(sh.getLastRow() + 1, 1, novas.length, h.length).setValues(novas);
  }
}

// ---- inscrição pelo site ----
function inscrever(p) {
  if (!verificarRateLimit("rl_insc_global", 30, 300)) return {ok:0, msg:"Muitas inscrições simultâneas. Aguarde alguns instantes e tente novamente."};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ev = eventos().find(e => e.id === p.evento_id);
    if (!ev || /^inscri.*no evento/i.test(ev.status || "")) return {ok:0, msg:"Inscrição feita no dia do evento, com os organizadores."};
    if (!/^aberto/i.test(ev.status)) return {ok:0, msg:"As inscrições deste evento não estão abertas."};
    if (!aberto(ev)) return {ok:0, msg:"O prazo de inscrição deste evento já terminou."};
    if (p.termo !== true) return {ok:0, msg:"É necessário aceitar o Termo de Responsabilidade."};
    const v = checar(ev, p, ativas(ev.id)); if (v.erro) return {ok:0, msg:v.erro};
    registrarFalhaRateLimit("rl_insc_global", 300);
    const codigo = novoCodigo(), proto = protocolo();
    gravar(ev, {at:v.at, time:v.time, igt:p.instagram_time}, {codigo, protocolo:proto, origem:"site", termo:"Aceito em " + hoje()});
    indice().appendRow([codigo, ev.id, proto]);
    CacheService.getScriptCache().remove("pub");
    return {ok:1, protocolo:proto, codigo};
  } finally { lock.releaseLock(); }
}

// ---- Minha inscrição: consultar e editar pelo código ----
function achar(codigo) {
  const c = String(codigo || "").trim().toUpperCase(); if (!/^[A-Z][A-Z0-9]{5}$/.test(c)) return null;
  const l = indice().getDataRange().getValues().find(r => String(r[0]).toUpperCase() === c); if (!l) return null;
  const rows = lerInsc(String(l[1])).filter(i => i.codigo.trim().toUpperCase() === c);
  return rows.length ? {rows, evId:String(l[1]), codigo:c} : null;
}
function minha(p) {
  if (!verificarRateLimit("falha_cod", 10, 900)) return {ok:0, msg:"Muitas tentativas inválidas. Por segurança, aguarde 15 minutos e tente de novo."};
  const f = achar(p.codigo);
  if (!f) {
    Utilities.sleep(350);
    registrarFalhaRateLimit("falha_cod", 900);
    return {ok:0, msg:"Código não encontrado. Confira as 6 letras e números."};
  }
  const ev = eventos().find(e => e.id === f.evId) || {}, r = f.rows;
  return {ok:1, codigo:f.codigo, protocolo:r[0].protocolo, situacao:r[0].status_aprovacao, editavel:aberto(ev),
    evento:{id:ev.id, nome:ev.nome || r[0].evento, data_inicio:ev.data_inicio, data_fim:ev.data_fim, status:ev.status, min_atletas:ev.min_atletas, max_atletas:ev.max_atletas},
    time:r[0].time, instagram_time:r[0].instagram_time,
    atletas:r.map(i => ({nome:i.nome, idade:i.idade, telefone:i.telefone, cpf:i.cpf, email:i.email, rede_social:i.rede_social}))};
}
function editar(p) {
  if (!verificarRateLimit("falha_cod", 10, 900)) return {ok:0, msg:"Muitas tentativas inválidas. Por segurança, aguarde 15 minutos e tente de novo."};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const f = achar(p.codigo);
    if (!f) {
      Utilities.sleep(350);
      registrarFalhaRateLimit("falha_cod", 900);
      return {ok:0, msg:"Código não encontrado."};
    }
    const ev = eventos().find(e => e.id === f.evId);
    if (!aberto(ev)) return {ok:0, msg:"A edição desta inscrição está encerrada. Fale com a Secretaria."};
    const ins = ativas(ev.id).filter(i => i.codigo.trim().toUpperCase() !== f.codigo);
    const v = checar(ev, {time:f.rows[0].time, atletas:p.atletas}, ins, {editando:true}); if (v.erro) return {ok:0, msg:v.erro};
    gravar(ev, {at:v.at, time:v.time, igt:p.instagram_time}, {rows:f.rows});
    CacheService.getScriptCache().remove("pub"); return {ok:1};
  } finally { lock.releaseLock(); }
}

// ---- Acesso Servidores: login pela aba "usuarios" com Rate Limit ----
function login(p) {
  const c = CacheService.getScriptCache(), lg = String(p.login || "").trim().toLowerCase(), kf = "falha_" + lg, kg = "falha_geral";
  if (!lg || !p.senha) return {ok:0, msg:"Informe login e senha."};
  if (!verificarRateLimit(kf, 5, 900) || !verificarRateLimit(kg, 15, 900)) {
    return {ok:0, msg:"Muitas tentativas de login. Aguarde 15 minutos e tente novamente."};
  }
  const u = ler("usuarios").find(x => x.login.trim().toLowerCase() === lg && x.senha.trim() === String(p.senha).trim());
  if (!u) {
    Utilities.sleep(500);
    const tent = registrarFalhaRateLimit(kf, 900);
    registrarFalhaRateLimit(kg, 900);
    const rest = Math.max(0, 5 - tent);
    return {ok:0, msg:"Login ou senha incorretos." + (rest > 0 ? " Restam " + rest + " tentativa(s)." : " Bloqueado por 15 minutos.")};
  }
  limparRateLimit(kf);
  const t = Utilities.getUuid() + Utilities.getUuid();
  c.put("tok_" + t, lg, 21600); // sessão máxima de 6 horas no servidor (e 15 min de inatividade no cliente)
  registrarAuditoria(lg, "Login", lg + " iniciou sessão no painel administrativo");
  return {ok:1, token:t, login:lg};
}

function admin(p) {
  const c = CacheService.getScriptCache();
  if (!verificarRateLimit("falha_tok_admin", 15, 900)) {
    return {ok:0, msg:"Muitas requisições não autorizadas. Aguarde 15 minutos.", expirou:1};
  }
  const usr = p.token ? c.get("tok_" + p.token) : "";
  if (!usr) {
    Utilities.sleep(300);
    registrarFalhaRateLimit("falha_tok_admin", 900);
    return {ok:0, msg:"Sessão expirada. Entre novamente.", expirou:1};
  }
  const ab = p.aba, limpa = () => c.remove("pub"), evl = () => eventos().map(e => ({id:e.id, nome:e.nome}));
  if (p.acao === "logout") {
    c.remove("tok_" + p.token);
    registrarAuditoria(usr, "Logout", usr + " encerrou a sessão administrativa");
    return {ok:1};
  }
  if (p.acao === "listar") {
    if (ab === "auditoria") return {ok:1, rows:ler("auditoria").reverse().slice(0, 200)};
    if (ab === "inscricoes") { const ev = p.evento_id ? eventos().find(e => e.id === p.evento_id) : null; return {ok:1, rows:ev ? lerInsc(ev.id) : [], eventos:evl(), evento:ev || null}; }
    if (["eventos","feed","colocacoes"].indexOf(ab) < 0) return {ok:0, msg:"Aba inválida."};
    return {ok:1, rows:ler(ab), eventos:evl()};
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (p.acao === "excluir" && ["eventos","feed","colocacoes"].indexOf(ab) >= 0) {
      if (+p.row < 2) return {ok:0};
      const sh = aba(ab), h = cabecalho(sh), vals = sh.getRange(+p.row, 1, 1, h.length).getValues()[0];
      const nomeItem = vals[h.indexOf("nome")] || vals[h.indexOf("titulo")] || vals[h.indexOf("nome_equipe")] || ("linha " + p.row);
      sh.deleteRow(+p.row);
      registrarAuditoria(usr, "Exclusão (" + ab + ")", usr + " excluiu '" + nomeItem + "' da aba " + ab);
      limpa(); return {ok:1};
    }
    if (p.acao === "salvar" && ["eventos","feed","colocacoes"].indexOf(ab) >= 0) {
      const sh = aba(ab), h = cabecalho(sh), d = p.dados || {};
      const old = d._row ? sh.getRange(+d._row, 1, 1, h.length).getValues()[0] : [];
      if (ab === "eventos") {
        if (!d.id && !d._row) {
          const u = {}; ler("eventos").forEach(e => { if (e.id) u[e.id] = 1; });
          d.id = idUnico(d.nome, d.data_inicio, u);
        }
        const fim = String(d.data_fim || d.data_inicio || "").trim().slice(0, 10);
        if (/^\d{4}-\d{2}-\d{2}$/.test(fim) && fim < hoje()) {
          d.status = "Encerrado";
        }
        try {
          const idxPastaEv = h.indexOf("pasta_evento_url");
          const pastaEvAntiga = d.pasta_evento_url || (idxPastaEv >= 0 ? old[idxPastaEv] : "");
          const pst = garantirPastasEvento(d.nome || "Evento " + (d.id || ""), pastaEvAntiga);
          d.pasta_evento_url = pst.eventoUrl;
          if (!d.pasta_fotos_url) d.pasta_fotos_url = pst.fotosUrl;
          if (p.capa_b64) {
            const img = salvarBase64NoDrive(pst.pastaCapa, p.capa_b64, p.capa_nome || ("capa-" + (d.id || "evento")), true);
            if (img) d.imagem_url = img.thumb;
          } else if (!d.imagem_url) {
            const capaDrive = buscarCapaNaPasta(pst.pastaCapa);
            if (capaDrive) d.imagem_url = capaDrive;
          }
        } catch (errDrive) {
          console.error("Aviso Drive:", errDrive);
        }
      }
      const row = h.map((k,i) => d[k] !== undefined && k !== "_row" ? cel(d[k]) : (old[i] === undefined ? "" : old[i]));
      const targetRow = d._row ? +d._row : sh.getLastRow() + 1;
      try { sh.getRange(targetRow, 1, 1, h.length).clearDataValidations(); } catch (eVal) {}
      sh.getRange(targetRow, 1, 1, h.length).setValues([row]);
      if (ab === "eventos") { const id = d.id || row[h.indexOf("id")]; if (id) abaInsc(id, true); }
      const descItem = d.nome || d.titulo || d.nome_equipe || d.id || "";
      registrarAuditoria(usr, (d._row ? "Edição (" : "Criação (") + ab + ")", usr + (d._row ? " atualizou '" : " cadastrou '") + descItem + "'" + (d.status ? " [Status: " + d.status + ", Vagas: " + (d.max_vagas || "sem limite") + "]" : ""));
      limpa(); return {ok:1};
    }
    if (p.acao === "uploadFotos") {
      const sh = aba("eventos"), h = cabecalho(sh), evs = ler("eventos");
      const ev = evs.find(e => e.id === p.evento_id || String(e._row) === String(p.row));
      if (!ev) return {ok:0, msg:"Evento não encontrado."};
      const pst = garantirPastasEvento(ev.nome, ev.pasta_evento_url);
      let pastaAlvo = pst.pastaFotos;
      const idF = idPasta(ev.pasta_fotos_url);
      if (idF) {
        try { pastaAlvo = DriveApp.getFolderById(idF); } catch (e) { pastaAlvo = pst.pastaFotos; }
      }
      const arqs = Array.isArray(p.arquivos) ? p.arquivos : [];
      if (!arqs.length) return {ok:0, msg:"Nenhuma foto selecionada."};
      let enviadas = 0;
      arqs.forEach((arq, idx) => {
        const salvo = salvarBase64NoDrive(pastaAlvo, arq.b64, arq.nome || ("foto-" + (idx + 1)), false);
        if (salvo) enviadas++;
      });
      const colFotos = h.indexOf("pasta_fotos_url") + 1, colEvUrl = h.indexOf("pasta_evento_url") + 1;
      const fotosUrlFinal = ev.pasta_fotos_url || pst.fotosUrl;
      if (!ev.pasta_fotos_url && colFotos > 0) sh.getRange(ev._row, colFotos).setValue(fotosUrlFinal);
      if (!ev.pasta_evento_url && colEvUrl > 0) sh.getRange(ev._row, colEvUrl).setValue(pst.eventoUrl);
      const idLimpar = idPasta(fotosUrlFinal);
      if (idLimpar) {
        c.remove("f" + idLimpar + "_4");
        c.remove("f" + idLimpar + "_60");
      }
      registrarAuditoria(usr, "Upload de Fotos", usr + " enviou " + enviadas + " foto(s) para o evento '" + ev.nome + "'");
      limpa();
      return {ok:1, enviadas, pasta_fotos_url: fotosUrlFinal};
    }
    // ----- inscrições (por evento) -----
    const ev = eventos().find(e => e.id === p.evento_id), cod = String(p.codigo || "").trim().toUpperCase();
    if (!ev) return {ok:0, msg:"Evento não encontrado."};
    const todas = lerInsc(ev.id), rows = cod ? todas.filter(i => i.codigo.trim().toUpperCase() === cod) : [];
    if (p.acao === "inscAdmin") {
      if (cod && !rows.length) return {ok:0, msg:"Inscrição não encontrada."};
      const st = ["Pendente","Aprovada","Recusada"].indexOf(p.status) >= 0 ? p.status : "", basica = p.modo === "basica";
      let v;
      if (basica) {
        const time = String(p.time || "").trim(), at = (p.atletas || []).filter(a => a && String(a.nome || "").trim());
        if (!time) return {ok:0, msg:"Informe o nome do time."};
        if (!at.length) return {ok:0, msg:"Informe ao menos um atleta."};
        v = {at, time};
      } else {
        v = checar(ev, {time:p.time, atletas:p.atletas}, todas.filter(i => i.status_aprovacao !== "Recusada" && i.codigo.trim().toUpperCase() !== cod), {admin:true, editando:!!cod});
        if (v.erro) return {ok:0, msg:v.erro};
      }
      v.igt = p.instagram_time;
      if (rows.length) {
        gravar(ev, v, {rows, status:st});
        registrarAuditoria(usr, "Edição de Inscrição", usr + " editou a inscrição " + cod + " (" + (v.time || v.at[0].nome) + ") no evento '" + ev.nome + "'");
      } else {
        const codigo = novoCodigo(), proto = protocolo();
        gravar(ev, v, {codigo, protocolo:proto, origem:basica ? "admin-basica" : "admin", termo:basica ? "Registro presencial (Secretaria)" : "Registrado pela Secretaria", status:st || "Aprovada"});
        indice().appendRow([codigo, ev.id, proto]);
        registrarAuditoria(usr, "Nova Inscrição (Admin)", usr + " cadastrou a inscrição " + codigo + " (" + (v.time || v.at[0].nome) + ") no evento '" + ev.nome + "'");
      }
      limpa(); return {ok:1};
    }
    if (p.acao === "inscStatus" && rows.length) {
      const sh = abaInsc(ev.id, false), col = cabecalho(sh).indexOf("status_aprovacao") + 1;
      if (["Pendente","Aprovada","Recusada"].indexOf(p.status) < 0) return {ok:0, msg:"Situação inválida."};
      rows.forEach(r => sh.getRange(r._row, col).setValue(p.status));
      const nomeEq = rows[0].time || rows[0].nome || cod;
      registrarAuditoria(usr, "Status de Inscrição", usr + " alterou a inscrição de '" + nomeEq + "' (" + cod + ") para " + p.status + " no evento '" + ev.nome + "'");
      limpa(); return {ok:1};
    }
    if (p.acao === "inscExcluir" && rows.length) {
      const nomeEq = rows[0].time || rows[0].nome || cod;
      const sh = abaInsc(ev.id, false); rows.slice().reverse().forEach(r => sh.deleteRow(r._row));
      registrarAuditoria(usr, "Exclusão de Inscrição", usr + " excluiu a inscrição de '" + nomeEq + "' (" + cod + ") do evento '" + ev.nome + "'");
      limpa(); return {ok:1};
    }
    return {ok:0, msg:"Ação desconhecida."};
  } finally { lock.releaseLock(); }
}
