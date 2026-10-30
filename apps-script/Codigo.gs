/**
 * App de Requisição de Materiais — backend (Google Apps Script)
 * Cole este arquivo em: Planilha Google > Extensões > Apps Script.
 * 1) Rode a função "configurar" uma vez (cria as abas e o catálogo de exemplo).
 * 2) Implantar > Nova implantação > App da Web > Executar como: Eu / Quem pode acessar: Qualquer pessoa.
 */

const ABA = { CAT: 'Categorias', ITENS: 'Itens', FUNC: 'Funcionarios', PED: 'Pedidos' };
const CAB = {
  CAT:   ['Categoria', 'Foto', 'Ordem'],
  ITENS: ['Código', 'Ativo', 'Categoria', 'Item', 'Especificação', 'Unidade', 'Foto'],
  FUNC:  ['Nome', 'Ativo'],
  PED:   ['Nº Pedido', 'ID Linha', 'Data', 'Funcionário', 'Urgente', 'Categoria', 'Código', 'Item',
          'Especificação', 'Qtd', 'Unidade', 'Obs', 'Status', 'Atualizado em']
};
const STATUS = ['Novo', 'Comprado', 'Entregue', 'Cancelado'];
const PIN_PADRAO = '1234';

// ---------- Instalação ----------
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  criarAba_(ss, ABA.CAT, CAB.CAT, [
    ["Parafusos e fixação", "parafuso.svg", 1],
    ["Colas e fitas", "cola.svg", 2],
    ["Lixas", "lixa.svg", 3],
    ["Pintura e acabamento", "lata.svg", 4],
    ["Ferragens", "dobradica.svg", 5],
    ["Chapas", "chapa.svg", 6],
    ["Ferramentas e consumíveis", "broca.svg", 7],
    ["EPI", "luva.svg", 8]
  ]);
  criarAba_(ss, ABA.ITENS, CAB.ITENS, [
    ["FIX-001", "SIM", "Parafusos e fixação", "Parafuso chipboard", "3,5 x 16 mm, cabeça chata", "un", "parafuso.svg"],
    ["FIX-002", "SIM", "Parafusos e fixação", "Parafuso chipboard", "4,0 x 25 mm, cabeça chata", "un", "parafuso.svg"],
    ["FIX-003", "SIM", "Parafusos e fixação", "Parafuso chipboard", "4,0 x 40 mm, cabeça chata", "un", "parafuso.svg"],
    ["FIX-004", "SIM", "Parafusos e fixação", "Parafuso chipboard", "4,5 x 50 mm, cabeça chata", "un", "parafuso.svg"],
    ["FIX-005", "SIM", "Parafusos e fixação", "Cavilha de madeira", "8 x 30 mm", "un", "cavilha.svg"],
    ["FIX-006", "SIM", "Parafusos e fixação", "Minifix (tambor + pino)", "15 mm", "jogo", "minifix.svg"],
    ["FIX-007", "SIM", "Parafusos e fixação", "Prego com cabeça", "15 x 15", "kg", "prego.svg"],
    ["FIX-008", "SIM", "Parafusos e fixação", "Grampo para pinador", "10 mm", "caixa", "grampo.svg"],
    ["COL-001", "SIM", "Colas e fitas", "Cola PVA branca", "1 kg", "un", "cola.svg"],
    ["COL-002", "SIM", "Colas e fitas", "Cola de contato", "750 g", "lata", "lata.svg"],
    ["COL-003", "SIM", "Colas e fitas", "Fita de borda PVC", "22 mm, branca", "rolo", "fita-borda.svg"],
    ["COL-004", "SIM", "Colas e fitas", "Fita crepe", "48 mm x 50 m", "rolo", "fita-crepe.svg"],
    ["LIX-001", "SIM", "Lixas", "Lixa folha", "grão 80", "folha", "lixa.svg"],
    ["LIX-002", "SIM", "Lixas", "Lixa folha", "grão 120", "folha", "lixa.svg"],
    ["LIX-003", "SIM", "Lixas", "Lixa folha", "grão 220", "folha", "lixa.svg"],
    ["LIX-004", "SIM", "Lixas", "Disco de lixa", "5\", grão 120", "un", "disco-lixa.svg"],
    ["PIN-001", "SIM", "Pintura e acabamento", "Seladora", "3,6 L", "galão", "lata.svg"],
    ["PIN-002", "SIM", "Pintura e acabamento", "Verniz PU", "3,6 L", "galão", "lata.svg"],
    ["PIN-003", "SIM", "Pintura e acabamento", "Tinta PU", "informar cor na observação", "L", "lata.svg"],
    ["PIN-004", "SIM", "Pintura e acabamento", "Thinner", "5 L", "galão", "galao.svg"],
    ["PIN-005", "SIM", "Pintura e acabamento", "Massa para madeira", "400 g", "un", "pote.svg"],
    ["PIN-006", "SIM", "Pintura e acabamento", "Pincel", "2\"", "un", "pincel.svg"],
    ["PIN-007", "SIM", "Pintura e acabamento", "Rolo de espuma", "9 cm", "un", "rolo.svg"],
    ["FER-001", "SIM", "Ferragens", "Dobradiça caneco", "35 mm, reta", "un", "dobradica.svg"],
    ["FER-002", "SIM", "Ferragens", "Dobradiça caneco", "35 mm, curva", "un", "dobradica.svg"],
    ["FER-003", "SIM", "Ferragens", "Corrediça telescópica", "450 mm", "par", "corredica.svg"],
    ["FER-004", "SIM", "Ferragens", "Puxador", "informar modelo na observação", "un", "puxador.svg"],
    ["CHA-001", "SIM", "Chapas", "Chapa MDF", "15 mm, branco, 2750 x 1850 mm", "chapa", "chapa.svg"],
    ["CHA-002", "SIM", "Chapas", "Chapa MDF", "18 mm, cru, 2750 x 1850 mm", "chapa", "chapa.svg"],
    ["CHA-003", "SIM", "Chapas", "Chapa MDF", "6 mm, cru, 2750 x 1850 mm", "chapa", "chapa.svg"],
    ["FRM-001", "SIM", "Ferramentas e consumíveis", "Broca para madeira", "8 mm", "un", "broca.svg"],
    ["FRM-002", "SIM", "Ferramentas e consumíveis", "Broca Forstner (caneco)", "35 mm", "un", "broca.svg"],
    ["FRM-003", "SIM", "Ferramentas e consumíveis", "Lâmina de estilete", "18 mm", "caixa", "estilete.svg"],
    ["EPI-001", "SIM", "EPI", "Luva de proteção", "tamanho M", "par", "luva.svg"],
    ["EPI-002", "SIM", "EPI", "Máscara PFF2", "", "un", "mascara.svg"],
    ["EPI-003", "SIM", "EPI", "Óculos de proteção", "", "un", "oculos.svg"],
    ["EPI-004", "SIM", "EPI", "Protetor auricular tipo concha", "", "un", "abafador.svg"]
  ]);
  criarAba_(ss, ABA.FUNC, CAB.FUNC, ["Exemplo 1 (apagar)", "Exemplo 2 (apagar)"].map(n => [n, 'SIM']));
  const ped = criarAba_(ss, ABA.PED, CAB.PED, []);

  // Listas suspensas e cores
  const sim = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).build();
  ss.getSheetByName(ABA.ITENS).getRange('B2:B').setDataValidation(sim);
  ss.getSheetByName(ABA.FUNC).getRange('B2:B').setDataValidation(sim);
  ped.getRange('M2:M').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());
  ped.getRange('A2:B').setNumberFormat('@');   // texto: evita "0001" virar 1 ou "0001-2" virar data
  ped.getRange('C2:C').setNumberFormat('dd/MM/yyyy HH:mm');
  ped.getRange('N2:N').setNumberFormat('dd/MM/yyyy HH:mm');
  const r = ped.getRange('A2:N');
  const cor = (txt, bg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$M2="' + txt + '"').setBackground(bg).setRanges([r]).build();
  ped.setConditionalFormatRules([cor('Novo', '#fff4cc'), cor('Comprado', '#dbe9ff'),
                                 cor('Entregue', '#dcf5dc'), cor('Cancelado', '#eeeeee')]);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN')) props.setProperty('PIN', PIN_PADRAO);
  if (!props.getProperty('SEQ')) props.setProperty('SEQ', '0');
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);
  Logger.log('Configuração concluída. PIN do painel de compras: ' + props.getProperty('PIN'));
}

/** Para trocar o PIN do painel de compras: altere o valor abaixo e rode esta função. */
function trocarPin() {
  PropertiesService.getScriptProperties().setProperty('PIN', '1234');
}

function criarAba_(ss, nome, cab, linhas) {
  let sh = ss.getSheetByName(nome);
  if (sh) return sh;                       // nunca sobrescreve uma aba existente
  sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#2f3b4c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (linhas.length) sh.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
  sh.autoResizeColumns(1, cab.length);
  return sh;
}

// ---------- Leitura ----------
function tabela_(nome) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  const v = sh.getDataRange().getValues();
  const cab = v.shift();
  return v.filter(l => l.join('') !== '').map(l => {
    const o = {}; cab.forEach((c, i) => o[c] = l[i]); return o;
  });
}
const ativo_ = v => String(v).trim().toUpperCase() !== 'NÃO' && String(v).trim().toUpperCase() !== 'NAO';

function catalogo_() {
  const categorias = tabela_(ABA.CAT)
    .sort((a, b) => (Number(a['Ordem']) || 99) - (Number(b['Ordem']) || 99))
    .map(c => ({ nome: String(c['Categoria']), foto: String(c['Foto'] || '') }));
  const itens = tabela_(ABA.ITENS).filter(i => ativo_(i['Ativo']) && i['Código']).map(i => ({
    id: String(i['Código']), categoria: String(i['Categoria']), nome: String(i['Item']),
    espec: String(i['Especificação'] || ''), unidade: String(i['Unidade'] || 'un'), foto: String(i['Foto'] || '')
  }));
  const funcionarios = tabela_(ABA.FUNC).filter(f => ativo_(f['Ativo']) && f['Nome']).map(f => String(f['Nome']));
  return { categorias, itens, funcionarios };
}

function linhasPedidos_(filtro) {
  return tabela_(ABA.PED).filter(filtro).map(p => ({
    pedido: String(p['Nº Pedido']), linha: String(p['ID Linha']),
    data: p['Data'] instanceof Date ? p['Data'].toISOString() : String(p['Data']),
    funcionario: String(p['Funcionário']), urgente: String(p['Urgente']) === 'SIM',
    categoria: String(p['Categoria']), codigo: String(p['Código']), item: String(p['Item']),
    espec: String(p['Especificação']), qtd: Number(p['Qtd']), unidade: String(p['Unidade']),
    obs: String(p['Obs']), status: String(p['Status'] || 'Novo')
  }));
}

// ---------- API ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'catalogo') return json_(catalogo_());
    if (p.action === 'meus') {
      const limite = Date.now() - 30 * 864e5;
      return json_({ linhas: linhasPedidos_(l => String(l['Funcionário']) === p.funcionario &&
                                                  new Date(l['Data']).getTime() >= limite) });
    }
    if (p.action === 'pedidos') {
      if (!pinOk_(p.pin)) return json_({ erro: 'PIN inválido' });
      const dias = Number(p.dias) || 15, limite = Date.now() - dias * 864e5;
      return json_({ linhas: linhasPedidos_(l => ['Novo', 'Comprado'].includes(String(l['Status'])) ||
                                                  new Date(l['Data']).getTime() >= limite) });
    }
    return json_({ ok: true, app: 'requisicao-materiais' });
  } catch (err) { return json_({ erro: String(err) }); }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ erro: 'Requisição inválida' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'novoPedido') return json_(novoPedido_(body));
    if (body.action === 'status') return json_(mudarStatus_(body));
    return json_({ erro: 'Ação desconhecida' });
  } catch (err) {
    return json_({ erro: String(err) });
  } finally { lock.releaseLock(); }
}

function novoPedido_(b) {
  const cat = catalogo_();
  if (!cat.funcionarios.includes(b.funcionario)) throw new Error('Funcionário não cadastrado');
  if (!Array.isArray(b.itens) || !b.itens.length) throw new Error('Pedido vazio');
  if (b.chave) {                                   // evita pedido duplicado se o celular reenviar
    const c = CacheService.getScriptCache().get('k_' + b.chave);
    if (c) return { ok: true, pedido: c, repetido: true };
  }
  const mapa = {}; cat.itens.forEach(i => mapa[i.id] = i);
  const props = PropertiesService.getScriptProperties();
  const seq = Number(props.getProperty('SEQ') || 0) + 1;
  props.setProperty('SEQ', String(seq));
  const num = String(seq).padStart(4, '0');
  const agora = new Date();
  const linhas = b.itens.map((it, n) => {
    const qtd = Number(it.qtd);
    if (!(qtd > 0)) throw new Error('Quantidade inválida');
    const c = mapa[it.id];
    const nome = c ? c.nome : String(it.nome || '').slice(0, 120);
    if (!nome) throw new Error('Item sem nome');
    return [num, num + '-' + (n + 1), agora, b.funcionario, b.urgente ? 'SIM' : 'NÃO',
            c ? c.categoria : 'Outro (fora do catálogo)', c ? c.id : 'OUTRO', nome,
            c ? c.espec : '', qtd, c ? c.unidade : String(it.unidade || 'un').slice(0, 20),
            String(it.obs || b.obs || '').slice(0, 300), 'Novo', agora];
  });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA.PED);
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  if (b.chave) CacheService.getScriptCache().put('k_' + b.chave, num, 21600);
  return { ok: true, pedido: num };
}

function mudarStatus_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  if (!STATUS.includes(b.status)) throw new Error('Status inválido');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA.PED);
  const v = sh.getDataRange().getValues();
  const iPed = CAB.PED.indexOf('Nº Pedido'), iLin = CAB.PED.indexOf('ID Linha'),
        iSt = CAB.PED.indexOf('Status'), iAt = CAB.PED.indexOf('Atualizado em');
  const alvo = new Set((b.linhas || []).map(String));
  let n = 0;
  for (let r = 1; r < v.length; r++) {
    if (alvo.has(String(v[r][iLin])) || (b.pedido && String(v[r][iPed]) === String(b.pedido))) {
      sh.getRange(r + 1, iSt + 1).setValue(b.status);
      sh.getRange(r + 1, iAt + 1).setValue(new Date());
      n++;
    }
  }
  return { ok: true, alteradas: n };
}

function pinOk_(pin) {
  return String(pin || '') === String(PropertiesService.getScriptProperties().getProperty('PIN') || PIN_PADRAO);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
