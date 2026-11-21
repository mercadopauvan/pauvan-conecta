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
  SS_ = null; ABAS_ = {}; LIDO_ = {}; CAB_ = {};
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

/* ---------- planilha ----------
 * Regras de velocidade (valem para todo código novo ou alterado):
 * 1) a planilha é aberta uma vez só e cada aba é lida uma vez só por chamada, em bloco (getValues);
 * 2) abas que mudam pouco (fornecedores, produtos, pessoas) ficam em cache por CACHE_SEG segundos;
 * 3) ao celular vai só o que a tela mostra (ver estado_).
 */
var CACHE_ABAS = { FORNECEDORES: true, PRODUTOS: true, PESSOAS: true };
var CACHE_SEG = 600;
var SS_ = null, ABAS_ = {}, LIDO_ = {}, CAB_ = {};
function ss_() {
  if (!SS_) SS_ = SpreadsheetApp.openById(SS_ID);
  return SS_;
}
function aba_(nome) {
  if (!ABAS_[nome]) {
    var s = ss_().getSheetByName(nome);
    if (!s) throw new Error('Aba ' + nome + ' não encontrada');
    ABAS_[nome] = s;
  }
  return ABAS_[nome];
}
/* Depois de gravar numa aba, a próxima leitura da mesma chamada busca de novo. */
function sujo_(nome) { delete LIDO_[nome]; }
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
  if (LIDO_[nome]) return LIDO_[nome];
  var cache = CACHE_ABAS[nome] ? CacheService.getScriptCache() : null, chave = 'aba_v2_' + nome;
  if (cache) {
    var j = cache.get(chave);
    if (j) return (LIDO_[nome] = JSON.parse(j));
  }
  var v = aba_(nome).getDataRange().getValues(), cab = v[0] || [], out = [];
  CAB_[nome] = cab.map(txt_);
  for (var i = 1; i < v.length; i++) {
    if (txt_(v[i][0]) === '') continue;
    var o = { _lin: i + 1 };
    for (var c = 0; c < cab.length; c++) o[txt_(cab[c])] = v[i][c];
    out.push(o);
  }
  LIDO_[nome] = out;
  if (cache) { try { cache.put(chave, JSON.stringify(out), CACHE_SEG); } catch (e) {} }
  return out;
}
function cab_(nome) {
  if (!CAB_[nome]) {
    var s = aba_(nome);
    CAB_[nome] = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(txt_);
  }
  return CAB_[nome];
}
function incluir_(nome, objs) {
  if (!objs.length) return;
  var s = aba_(nome), cab = cab_(nome);
  var linhas = objs.map(function (o) { return cab.map(function (k) { return o[k] === undefined ? '' : o[k]; }); });
  s.getRange(s.getLastRow() + 1, 1, linhas.length, cab.length).setValues(linhas);
  sujo_(nome);
}
function gravar_(nome, lin, campos) {
  var s = aba_(nome), cab = cab_(nome);
  Object.keys(campos).forEach(function (k) {
    var c = cab.indexOf(k);
    if (c > -1) s.getRange(lin, c + 1).setValue(campos[k]);
  });
  sujo_(nome);
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
    if (!meu) { o.whatsapp = txt_(f.whatsapp).replace(/\D/g, ''); o.contador = txt_(f.contador).replace(/\D/g, ''); o.token = txt_(f.token); }
    return o;
  });
  var prod = ler_('PRODUTOS').filter(function (p) { return sim_(p.ativo) && (!meu || txt_(p.fornecedor) === meu); }).map(function (p) {
    return { codigo: txt_(p.codigo), fornecedor: txt_(p.fornecedor), nome: txt_(p.nome), desenho: txt_(p.desenho), custo: num_(p.custo), ordem: num_(p.ordem) };
  }).sort(function (a, b) { return a.ordem - b.ordem; });
  var itens = {};
  ler_('ITENS').forEach(function (i) {
    var k = txt_(i.pedido);
    (itens[k] = itens[k] || []).push({ codigo: txt_(i.codigo), nome: txt_(i.nome), custo: num_(i.custo), pedida: num_(i.pedida), enviada: txt_(i.enviada) === '' ? null : num_(i.enviada), recebida: txt_(i.recebida) === '' ? null : num_(i.recebida), validade: txt_(i.validade) });
  });
  var ped = ler_('PEDIDOS').filter(function (p) {
    return (!meu || txt_(p.fornecedor) === meu) && txt_(p.status) !== 'CANCELADO' && (txt_(p.dia) >= corte || txt_(p.status) !== 'RECEBIDO');
  }).map(function (p) {
    return { id: txt_(p.id), fornecedor: txt_(p.fornecedor), dia: txt_(p.dia), status: txt_(p.status), criado_em: txt_(p.criado_em), criado_por: txt_(p.criado_por), enviado_em: txt_(p.enviado_em), recebido_em: txt_(p.recebido_em), recebido_por: txt_(p.recebido_por), coleta: txt_(p.coleta), taxa: num_(p.taxa), itens: itens[txt_(p.id)] || [] };
  });
  var perc = ler_('PERCAS').filter(function (p) { return (!meu || txt_(p.fornecedor) === meu) && (txt_(p.data) >= corte || txt_(p.troca) === 'PENDENTE'); }).map(function (p) {
    var tr = txt_(p.troca);
    return { id: txt_(p.id), data: txt_(p.data), fornecedor: txt_(p.fornecedor), codigo: txt_(p.codigo), nome: txt_(p.nome), qtd: num_(p.qtd), custo: num_(p.custo), motivo: txt_(p.motivo), por: txt_(p.lancado_por), troca: tr, pend: tr === 'PENDENTE' ? (txt_(p.pend) === '' ? num_(p.qtd) : num_(p.pend)) : 0, troca_em: txt_(p.troca_em) };
  });
  var intr = {}, props = PropertiesService.getScriptProperties();
  forn.forEach(function (f) { if (f.tem_pedido) return; var j = props.getProperty('interesse_' + f.id); if (j) { try { intr[f.id] = JSON.parse(j); } catch (e) {} } });
  return { quem: quem, hoje: hoje_(), fornecedores: forn, produtos: prod, pedidos: ped, percas: perc, interesses: intr };
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
function forn_(id) {
  return ler_('FORNECEDORES').filter(function (x) { return txt_(x.id) === String(id); })[0];
}
/* Fornecedor que entrega por conta própria, sem pedido do mercado (ex.: Iogurte Soares). */
function semPedido_(f) { return f && !sim_(f.tem_pedido); }
function data_(v) { var t = String(v || ''); return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : ''; }
function gravarValidades_(id, val) {
  if (!val) return;
  ler_('ITENS').forEach(function (i) {
    if (txt_(i.pedido) !== String(id)) return;
    var d = data_(val[txt_(i.codigo)]);
    if (d) gravar_('ITENS', i._lin, { validade: d });
  });
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
  } else if (a === 'entrega_lancar') {
    if (quem.tipo !== 'fornecedor') throw new Error('Só o fornecedor lança a entrega.');
    var fe = forn_(quem.fornecedor);
    if (!semPedido_(fe)) throw new Error('Este fornecedor trabalha com pedido do mercado.');
    var me = qtds_(req.itens), ide = novoId_('E'), le = [], vl = req.validades || {};
    ler_('PRODUTOS').forEach(function (p) {
      var q = me[txt_(p.codigo)];
      if (sim_(p.ativo) && txt_(p.fornecedor) === txt_(fe.id) && q > 0) le.push({ pedido: ide, codigo: txt_(p.codigo), nome: txt_(p.nome), custo: num_(p.custo), pedida: q, enviada: q, recebida: '', validade: data_(vl[txt_(p.codigo)]) });
    });
    if (!le.length) throw new Error('Informe a quantidade de pelo menos um produto.');
    var ag = agora_();
    incluir_('PEDIDOS', [{ id: ide, fornecedor: txt_(fe.id), dia: hoje_(), status: 'ENVIADO', criado_em: ag, criado_por: quem.nome, enviado_em: ag, coleta: 'FORNECEDOR', taxa: 0 }]);
    incluir_('ITENS', le);
    PropertiesService.getScriptProperties().deleteProperty('interesse_' + txt_(fe.id));
  } else if (a === 'interesse_salvar') {
    /* Pedido sem quantidade (quero / não quero) para fornecedor que entrega por conta própria. */
    soMercado_(quem);
    var fi = forn_(req.fornecedor);
    if (!semPedido_(fi) || !sim_(fi.ativo)) throw new Error('Fornecedor não encontrado.');
    var ok = {};
    ler_('PRODUTOS').forEach(function (p) { if (sim_(p.ativo) && txt_(p.fornecedor) === txt_(fi.id)) ok[txt_(p.codigo)] = 1; });
    var cs = (req.codigos || []).map(String).filter(function (c) { return ok[c]; });
    if (!cs.length) throw new Error('Marque pelo menos um produto.');
    PropertiesService.getScriptProperties().setProperty('interesse_' + txt_(fi.id), JSON.stringify({ codigos: cs, por: quem.nome, em: agora_() }));
  } else if (a === 'pedido_cancelar') {
    soMercado_(quem);
    var p1 = pedido_(req.id, quem);
    if (txt_(p1.status) === 'RECEBIDO') throw new Error('Pedido já recebido não pode ser cancelado.');
    gravar_('PEDIDOS', p1._lin, { status: 'CANCELADO' });
  } else if (a === 'pedido_enviar') {
    var p2 = pedido_(req.id, quem);
    if (txt_(p2.status) === 'RECEBIDO' || txt_(p2.status) === 'CANCELADO') throw new Error('Este pedido já foi encerrado.');
    var m2 = qtds_(req.itens);
    gravarItens_(req.id, 'enviada', m2);
    if (semPedido_(forn_(txt_(p2.fornecedor)))) { gravarItens_(req.id, 'pedida', m2); gravarValidades_(req.id, req.validades); }
    gravar_('PEDIDOS', p2._lin, { status: 'ENVIADO', enviado_em: agora_() });
  } else if (a === 'pedido_receber') {
    soMercado_(quem);
    var p3 = pedido_(req.id, quem);
    if (txt_(p3.status) === 'RECEBIDO' || txt_(p3.status) === 'CANCELADO') throw new Error('Este pedido já foi encerrado.');
    var f3 = ler_('FORNECEDORES').filter(function (x) { return txt_(x.id) === txt_(p3.fornecedor); })[0];
    var buscou = req.coleta === 'MERCADO' && !semPedido_(f3);
    gravarItens_(req.id, 'recebida', qtds_(req.itens));
    if (semPedido_(f3)) {
      gravarValidades_(req.id, req.validades);
      var tq = req.trocas || {}, ag3 = agora_();
      Object.keys(tq).forEach(function (c) {
        var n = Math.max(0, Math.min(999, Math.floor(Number(tq[c]) || 0)));
        ler_('PERCAS').filter(function (x) { return txt_(x.fornecedor) === txt_(f3.id) && txt_(x.troca) === 'PENDENTE' && txt_(x.codigo) === String(c); })
          .sort(function (x, y) { return txt_(x.data) < txt_(y.data) ? -1 : 1; })
          .forEach(function (x) {
            if (n <= 0) return;
            var pend = txt_(x.pend) === '' ? num_(x.qtd) : num_(x.pend), u = Math.min(n, pend);
            n -= u; pend -= u;
            gravar_('PERCAS', x._lin, pend ? { pend: pend } : { pend: 0, troca: 'ENTREGUE', troca_em: ag3 });
          });
      });
    }
    gravar_('PEDIDOS', p3._lin, { status: 'RECEBIDO', recebido_em: agora_(), recebido_por: quem.nome, coleta: buscou ? 'MERCADO' : 'FORNECEDOR', taxa: buscou && f3 ? num_(f3.taxa_coleta) : 0 });
  } else if (a === 'perca_lancar') {
    soMercado_(quem);
    var m4 = qtds_(req.itens), motivo = String(req.motivo || '').slice(0, 30), novas = [], troca4 = semPedido_(forn_(req.fornecedor));
    ler_('PRODUTOS').forEach(function (p) {
      var q = m4[txt_(p.codigo)];
      if (txt_(p.fornecedor) === String(req.fornecedor) && q > 0) novas.push({ id: novoId_('X'), data: hoje_(), fornecedor: txt_(p.fornecedor), codigo: txt_(p.codigo), nome: txt_(p.nome), qtd: q, custo: num_(p.custo), motivo: motivo, lancado_por: quem.nome, lancado_em: agora_(), troca: troca4 ? 'PENDENTE' : '', pend: troca4 ? q : '' });
    });
    if (!novas.length) throw new Error('Informe a quantidade da perca.');
    incluir_('PERCAS', novas);
  } else if (a === 'perca_excluir') {
    soMercado_(quem);
    var x = ler_('PERCAS').filter(function (p) { return txt_(p.id) === String(req.id); })[0];
    if (!x) throw new Error('Perca não encontrada.');
    if (txt_(x.data) !== hoje_() && quem.perfil !== 'GESTAO') throw new Error('Só a gestão exclui perca de outro dia.');
    aba_('PERCAS').deleteRow(x._lin);
    sujo_('PERCAS');
  } else {
    throw new Error('Ação desconhecida.');
  }
}

/* Teste: confere se o servidor consegue falar com o App Pauvan (nao usa PIN de verdade). */
function testePin() {
  var r = UrlFetchApp.fetch(PIN_URL, { method: 'post', contentType: 'text/plain', payload: JSON.stringify({ acao: 'conectaPin', pin: 'ab', aparelho: 'teste-conecta' }), muteHttpExceptions: true });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText().slice(0, 300));
}
