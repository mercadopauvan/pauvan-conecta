/**
 * PAUVAN CONECTA - servidor (Google Apps Script)
 * Planilha: PAUVAN CONECTA (pasta ARQUIVOS CLAUDE)
 * Implantar como App da Web: executar como "Eu", acesso "Qualquer pessoa".
 * O aplicativo (GitHub Pages) conversa com este endereço por POST (texto JSON).
 */
var SS_ID = '1Gps3J5uFACgOT-icPzTE6zJqSdI607lrquiCvKrKSJo';
var TZ = 'America/Cuiaba';
var DIAS_HISTORICO = 75;
/* Consulta de PIN do App Pauvan: o mesmo PIN vale nos dois aplicativos. */
var PIN_URL = 'https://script.google.com/macros/s/AKfycby5E_gkrNjAOsBgpYbrfJvXaKcHqMbTwH8Jb5ma_U-ThOhAFnVstqabM8bbnbJOB4DA/exec';

function doGet() {
  return saida_({ ok: true, app: 'PAUVAN CONECTA' });
}

function doPost(e) {
  var r;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    r = rotear_(req);
  } catch (err) {
    r = { ok: false, erro: String((err && err.message) || err) };
  }
  return saida_(r);
}

function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- planilha ---------- */
function aba_(nome) {
  var s = SpreadsheetApp.openById(SS_ID).getSheetByName(nome);
  if (!s) throw new Error('Aba ' + nome + ' não encontrada');
  return s;
}
function txt_(v) {
  if (v instanceof Date) {
    var comHora = v.getHours() || v.getMinutes();
    return Utilities.formatDate(v, TZ, comHora ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  return String(v === null || v === undefined ? '' : v).trim();
}
function num_(v) {
  if (typeof v === 'number') return v;
  var s = String(v === null || v === undefined ? '' : v);
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function ler_(nome) {
  var s = aba_(nome), n = s.getLastRow();
  if (n < 2) return [];
  var v = s.getRange(1, 1, n, s.getLastColumn()).getValues(), cab = v[0], out = [];
  for (var i = 1; i < v.length; i++) {
    if (txt_(v[i][0]) === '') continue;
    var o = { _lin: i + 1 };
    for (var c = 0; c < cab.length; c++) o[txt_(cab[c])] = v[i][c];
    out.push(o);
  }
  return out;
}
function cab_(nome) {
  var s = aba_(nome);
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(txt_);
}
function incluir_(nome, objs) {
  if (!objs.length) return;
  var s = aba_(nome), cab = cab_(nome);
  var linhas = objs.map(function (o) { return cab.map(function (k) { return o[k] === undefined ? '' : o[k]; }); });
  s.getRange(s.getLastRow() + 1, 1, linhas.length, cab.length).setValues(linhas);
}
function gravar_(nome, lin, campos) {
  var s = aba_(nome), cab = cab_(nome);
  Object.keys(campos).forEach(function (k) {
    var c = cab.indexOf(k);
    if (c > -1) s.getRange(lin, c + 1).setValue(campos[k]);
  });
}
function agora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function sim_(v) { return txt_(v).toUpperCase() === 'SIM'; }

/* ---------- acesso ---------- */
function falhou_() {
  var c = CacheService.getScriptCache(), n = Number(c.get('falhas') || 0) + 1;
  c.put('falhas', String(n), 600);
}
function entrar_(req) {
  var c = CacheService.getScriptCache();
  if (Number(c.get('falhas') || 0) > 30) throw new Error('Muitas tentativas. Aguarde 10 minutos.');
  var pin = txt_(req.pin), token = txt_(req.token);
  if (token) {
    var fs = ler_('FORNECEDORES').filter(function (f) { return sim_(f.ativo) && txt_(f.token) === token; });
    if (!fs.length) { falhou_(); throw new Error('Link inválido.'); }
    var f = fs[0], fpin = txt_(f.pin);
    if (fpin && fpin !== pin) {
      if (pin) falhou_();
      return { precisaPin: true, nome: txt_(f.nome), erro: pin ? 'PIN incorreto.' : '' };
    }
    return { tipo: 'fornecedor', nome: txt_(f.nome), fornecedor: txt_(f.id) };
  }
  if (!/^\d{4}$/.test(pin)) throw new Error('Digite o PIN de 4 números.');
  var nome = nomeDoPin_(pin, txt_(req.aparelho).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 60));
  var ps = ler_('PESSOAS').filter(function (p) { return sim_(p.ativo) && chave_(p.nome) === chave_(nome); });
  if (!ps.length) throw new Error(nome + ' ainda não tem acesso ao Conecta.');
  return { tipo: 'mercado', nome: txt_(ps[0].nome), perfil: txt_(ps[0].perfil) };
}
function chave_(v) {
  return txt_(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}
/* Pergunta ao App Pauvan de quem é o PIN. Guarda a resposta certa por 6 horas (sem guardar o PIN). */
function nomeDoPin_(pin, aparelho) {
  var c = CacheService.getScriptCache();
  var k = 'pin_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'conecta|' + pin + '|' + aparelho));
  var nome = c.get(k);
  if (nome) return nome;
  var r;
  try {
    var resp = UrlFetchApp.fetch(PIN_URL, { method: 'post', contentType: 'text/plain', payload: JSON.stringify({ acao: 'conectaPin', pin: pin, aparelho: aparelho }), muteHttpExceptions: true });
    r = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('Não consegui conferir o PIN no App Pauvan. Tente de novo.');
  }
  if (!r || !r.ok || !r.nome) throw new Error((r && r.erro) || 'PIN incorreto.');
  c.put(k, String(r.nome), 21600);
  return String(r.nome);
}

/* ---------- dados ---------- */
function estado_(quem) {
  var corte = Utilities.formatDate(new Date(Date.now() - DIAS_HISTORICO * 86400000), TZ, 'yyyy-MM-dd');
  var meu = quem.tipo === 'fornecedor' ? quem.fornecedor : null;
  var forn = ler_('FORNECEDORES').filter(function (f) { return sim_(f.ativo) && (!meu || txt_(f.id) === meu); }).map(function (f) {
    var o = { id: txt_(f.id), nome: txt_(f.nome), grupo: txt_(f.grupo), tem_pedido: sim_(f.tem_pedido), taxa: num_(f.taxa_coleta), fechamento: txt_(f.fechamento) };
    if (!meu) { o.whatsapp = txt_(f.whatsapp).replace(/\D/g, ''); o.token = txt_(f.token); }
    return o;
  });
  var prod = ler_('PRODUTOS').filter(function (p) { return sim_(p.ativo) && (!meu || txt_(p.fornecedor) === meu); }).map(function (p) {
    return { codigo: txt_(p.codigo), fornecedor: txt_(p.fornecedor), nome: txt_(p.nome), desenho: txt_(p.desenho), custo: num_(p.custo), ordem: num_(p.ordem) };
  }).sort(function (a, b) { return a.ordem - b.ordem; });
  var itens = {};
  ler_('ITENS').forEach(function (i) {
    var k = txt_(i.pedido);
    (itens[k] = itens[k] || []).push({ codigo: txt_(i.codigo), nome: txt_(i.nome), custo: num_(i.custo), pedida: num_(i.pedida), enviada: txt_(i.enviada) === '' ? null : num_(i.enviada), recebida: txt_(i.recebida) === '' ? null : num_(i.recebida) });
  });
  var ped = ler_('PEDIDOS').filter(function (p) {
    return (!meu || txt_(p.fornecedor) === meu) && txt_(p.status) !== 'CANCELADO' && (txt_(p.dia) >= corte || txt_(p.status) !== 'RECEBIDO');
  }).map(function (p) {
    return { id: txt_(p.id), fornecedor: txt_(p.fornecedor), dia: txt_(p.dia), status: txt_(p.status), criado_em: txt_(p.criado_em), criado_por: txt_(p.criado_por), enviado_em: txt_(p.enviado_em), recebido_em: txt_(p.recebido_em), recebido_por: txt_(p.recebido_por), coleta: txt_(p.coleta), taxa: num_(p.taxa), itens: itens[txt_(p.id)] || [] };
  });
  var perc = ler_('PERCAS').filter(function (p) { return (!meu || txt_(p.fornecedor) === meu) && txt_(p.data) >= corte; }).map(function (p) {
    return { id: txt_(p.id), data: txt_(p.data), fornecedor: txt_(p.fornecedor), codigo: txt_(p.codigo), nome: txt_(p.nome), qtd: num_(p.qtd), custo: num_(p.custo), motivo: txt_(p.motivo), por: txt_(p.lancado_por) };
  });
  return { quem: quem, hoje: hoje_(), fornecedores: forn, produtos: prod, pedidos: ped, percas: perc };
}
var SEQ_ = 0;
function novoId_(pre) {
  return pre + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10) + (SEQ_++);
}
function qtds_(lista) {
  var m = {};
  (lista || []).forEach(function (i) {
    var q = Math.floor(Number(i.qtd));
    if (!(q >= 0) || q > 999) throw new Error('Quantidade inválida.');
    m[String(i.codigo)] = q;
  });
  return m;
}
function pedido_(id, quem) {
  var p = ler_('PEDIDOS').filter(function (x) { return txt_(x.id) === String(id); })[0];
  if (!p) throw new Error('Pedido não encontrado.');
  if (quem.tipo === 'fornecedor' && txt_(p.fornecedor) !== quem.fornecedor) throw new Error('Pedido de outro fornecedor.');
  return p;
}
function gravarItens_(id, campo, m) {
  ler_('ITENS').forEach(function (i) {
    if (txt_(i.pedido) !== String(id)) return;
    var c = {};
    c[campo] = m[txt_(i.codigo)] === undefined ? 0 : m[txt_(i.codigo)];
    gravar_('ITENS', i._lin, c);
  });
}

/* ---------- ações ---------- */
function rotear_(req) {
  var quem = entrar_(req);
  if (quem.precisaPin) return { ok: true, precisaPin: true, nome: quem.nome, erro: quem.erro };
  var a = req.acao || 'estado';
  if (a !== 'estado') {
    var trava = LockService.getScriptLock();
    trava.waitLock(20000);
    try { agir_(a, req, quem); } finally { trava.releaseLock(); }
  }
  return { ok: true, estado: estado_(quem) };
}
function soMercado_(quem) {
  if (quem.tipo !== 'mercado') throw new Error('Ação só do mercado.');
}
function agir_(a, req, quem) {
  if (a === 'pedido_salvar') {
    soMercado_(quem);
    var f = ler_('FORNECEDORES').filter(function (x) { return txt_(x.id) === String(req.fornecedor) && sim_(x.ativo); })[0];
    if (!f) throw new Error('Fornecedor não encontrado.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(req.dia)) || String(req.dia) < hoje_()) throw new Error('Dia do pedido inválido.');
    var m = qtds_(req.itens), id = novoId_('P'), linhas = [];
    ler_('PRODUTOS').forEach(function (p) {
      var q = m[txt_(p.codigo)];
      if (sim_(p.ativo) && txt_(p.fornecedor) === txt_(f.id) && q > 0) linhas.push({ pedido: id, codigo: txt_(p.codigo), nome: txt_(p.nome), custo: num_(p.custo), pedida: q, enviada: '', recebida: '' });
    });
    if (!linhas.length) throw new Error('Informe a quantidade de pelo menos um produto.');
    incluir_('PEDIDOS', [{ id: id, fornecedor: txt_(f.id), dia: String(req.dia), status: 'PEDIDO', criado_em: agora_(), criado_por: quem.nome, coleta: req.coleta === 'FORNECEDOR' ? 'FORNECEDOR' : 'MERCADO' }]);
    incluir_('ITENS', linhas);
  } else if (a === 'pedido_cancelar') {
    soMercado_(quem);
    var p1 = pedido_(req.id, quem);
    if (txt_(p1.status) === 'RECEBIDO') throw new Error('Pedido já recebido não pode ser cancelado.');
    gravar_('PEDIDOS', p1._lin, { status: 'CANCELADO' });
  } else if (a === 'pedido_enviar') {
    var p2 = pedido_(req.id, quem);
    if (txt_(p2.status) === 'RECEBIDO' || txt_(p2.status) === 'CANCELADO') throw new Error('Este pedido já foi encerrado.');
    gravarItens_(req.id, 'enviada', qtds_(req.itens));
    gravar_('PEDIDOS', p2._lin, { status: 'ENVIADO', enviado_em: agora_() });
  } else if (a === 'pedido_receber') {
    soMercado_(quem);
    var p3 = pedido_(req.id, quem);
    if (txt_(p3.status) === 'RECEBIDO' || txt_(p3.status) === 'CANCELADO') throw new Error('Este pedido já foi encerrado.');
    var f3 = ler_('FORNECEDORES').filter(function (x) { return txt_(x.id) === txt_(p3.fornecedor); })[0];
    var buscou = req.coleta === 'MERCADO';
    gravarItens_(req.id, 'recebida', qtds_(req.itens));
    gravar_('PEDIDOS', p3._lin, { status: 'RECEBIDO', recebido_em: agora_(), recebido_por: quem.nome, coleta: buscou ? 'MERCADO' : 'FORNECEDOR', taxa: buscou && f3 ? num_(f3.taxa_coleta) : 0 });
  } else if (a === 'perca_lancar') {
    soMercado_(quem);
    var m4 = qtds_(req.itens), motivo = String(req.motivo || '').slice(0, 30), novas = [];
    ler_('PRODUTOS').forEach(function (p) {
      var q = m4[txt_(p.codigo)];
      if (txt_(p.fornecedor) === String(req.fornecedor) && q > 0) novas.push({ id: novoId_('X'), data: hoje_(), fornecedor: txt_(p.fornecedor), codigo: txt_(p.codigo), nome: txt_(p.nome), qtd: q, custo: num_(p.custo), motivo: motivo, lancado_por: quem.nome, lancado_em: agora_() });
    });
    if (!novas.length) throw new Error('Informe a quantidade da perca.');
    incluir_('PERCAS', novas);
  } else if (a === 'perca_excluir') {
    soMercado_(quem);
    var x = ler_('PERCAS').filter(function (p) { return txt_(p.id) === String(req.id); })[0];
    if (!x) throw new Error('Perca não encontrada.');
    if (txt_(x.data) !== hoje_() && quem.perfil !== 'GESTAO') throw new Error('Só a gestão exclui perca de outro dia.');
    aba_('PERCAS').deleteRow(x._lin);
  } else {
    throw new Error('Ação desconhecida.');
  }
}

/* Teste: confere se o servidor consegue falar com o App Pauvan (nao usa PIN de verdade). */
function testePin() {
  var r = UrlFetchApp.fetch(PIN_URL, { method: 'post', contentType: 'text/plain', payload: JSON.stringify({ acao: 'conectaPin', pin: 'ab', aparelho: 'teste-conecta' }), muteHttpExceptions: true });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText().slice(0, 300));
}
