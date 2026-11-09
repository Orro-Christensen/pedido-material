/**
 * App de Requisição de Materiais — backend (Google Apps Script)
 * Projeto Apps Script independente (script.google.com) ligado à planilha pelo PLANILHA_ID.
 * 1) Rode a função "configurar" uma vez (cria as abas e o catálogo de exemplo).
 * 2) Implantar > Nova implantação > App da Web > Executar como: Eu / Quem pode acessar: Qualquer pessoa.
 */

const ABA = { CAT: 'Categorias', ITENS: 'Itens', FUNC: 'Funcionarios', PED: 'Pedidos', FORN: 'Fornecedores' };
const CAB = {
  CAT:   ['Categoria', 'Foto', 'Ordem'],
  ITENS: ['Código', 'Ativo', 'Categoria', 'Item', 'Especificação', 'Unidade', 'Foto', 'Fornecedor'],
  FUNC:  ['Nome', 'Ativo', 'Senha'],
  FORN:  ['Código', 'Ativo', 'Nome', 'Contato', 'Telefone / WhatsApp', 'E-mail', 'CNPJ', 'Observações'],
  PED:   ['Nº Pedido', 'ID Linha', 'Data', 'Funcionário', 'Urgente', 'Categoria', 'Código', 'Item',
          'Especificação', 'Qtd', 'Unidade', 'Obs', 'Status', 'Atualizado em',
          'Data Compra', 'Qtd Comprada', 'Preço Unit.', 'Fornecedor',
          'Nº Ped. Fornecedor', 'Data Ped. Fornecedor', 'Visto em']
};
// ID da Planilha Google (o trecho entre /d/ e /edit no endereço da planilha).
const PLANILHA_ID = '12xfUT4sbeSyINRmJhTsSd3VkouEKVhEJYkCTwraNL5E';
function planilha_() { return SpreadsheetApp.openById(PLANILHA_ID); }

const STATUS = ['Novo', 'Pedido ao fornecedor', 'Comprado', 'Cancelado'];
const PIN_PADRAO = '1234';

// ---------- Catálogo (gerado a partir de "Lista de insumos.xlsx") ----------
const SEED_CATEGORIAS = [
  ["Parafusos & Fixação", "Parafusos.webp", 1],
  ["Colas", "Cola_de_contato.webp", 2],
  ["Lixas", "lixas.webp", 3],
  ["Pintura e acabamento", "Pintura.webp", 4],
  ["Chapas", "Chapas.jpg", 5],
  ["Ferramentas", "Ferramentas.webp", 6],
  ["EPI e Uniformes", "EPI.webp", 7],
  ["Maquinas", "maquina.webp", 8],
  ["Sapatas e Niveladores", "Pe_nivelador_essentra.jpg", 9],
  ["Ferragens", "Ferragens.jpg", 10]
];
const SEED_ITENS = [
  ["FIX-001", "SIM", "Parafusos & Fixação", "Allen Estrutural", "7 x 70 mm", "un", "Allen_Rosca_Parcial.webp"],
  ["FIX-002", "SIM", "Parafusos & Fixação", "Allen Estrutural", "7 x 60 mm", "un", "Allen_Rosca_Parcial.webp"],
  ["FIX-003", "SIM", "Parafusos & Fixação", "Chip Board", "4,5 x 45 mm", "un", "Chip_4-5x45mm.webp; ChipBoard.webp"],
  ["FIX-004", "SIM", "Parafusos & Fixação", "Chip Board", "6,0 x 70 mm Charllote", "un", "ChipBoard.webp"],
  ["FIX-005", "SIM", "Parafusos & Fixação", "Chip Board", "6,0 x 80 mm", "un", "ChipBoard.webp"],
  ["FIX-006", "SIM", "Parafusos & Fixação", "Chip Board", "4,0 x 35 mm", "un", "ChipBoard.webp"],
  ["FIX-007", "SIM", "Parafusos & Fixação", "Chip Board", "3,5 x 20 mm", "un", "ChipBoard.webp"],
  ["FIX-008", "SIM", "Parafusos & Fixação", "Chip Board", "5,0 x 35 mm", "un", "ChipBoard.webp"],
  ["FIX-009", "SIM", "Parafusos & Fixação", "Flangeado Chip Board", "5,0 x 35 mm", "un", "ChipBoard_Flangeado.webp"],
  ["FIX-010", "SIM", "Parafusos & Fixação", "Flangeado Chip Board", "4,5 x 25 mm", "un", "ChipBoard_Flangeado.webp"],
  ["FIX-011", "SIM", "Parafusos & Fixação", "Allen Conico Preto", "1/4 x 25 mm - mesas", "un", "Allen_Preto_conico.webp; Allen_Preto_conico_Perfil.webp"],
  ["FIX-012", "SIM", "Parafusos & Fixação", "Barra Roscada", "1/4 pol x 1 m", "un", "Barra_roscada.jpg"],
  ["FIX-013", "SIM", "Parafusos & Fixação", "Barra Roscada", "6 mm x 1 m", "un", "Barra_roscada.jpg"],
  ["FIX-014", "SIM", "Parafusos & Fixação", "Porca", "6 mm", "un", "Porca_Sextavada.webp"],
  ["FIX-015", "SIM", "Parafusos & Fixação", "Porca", "1/4 pol", "un", "Porca_Sextavada.webp"],
  ["FIX-016", "SIM", "Parafusos & Fixação", "Arruela", "Furo 6 mm", "un", "Arruela.webp"],
  ["FIX-017", "SIM", "Parafusos & Fixação", "Arruela", "Furo 1/4", "un", "Arruela.webp"],
  ["FIX-018", "SIM", "Parafusos & Fixação", "Bucha Americana", "1/4 polegada x 20 mm", "un", "Bucha_americana.webp"],
  ["FIX-019", "SIM", "Parafusos & Fixação", "Bucha Americana", "1/4 polegada x 25 mm", "un", "Bucha_americana.webp"],
  ["FIX-020", "SIM", "Parafusos & Fixação", "Bucha Americana", "6 mm x 20 mm", "un", "Bucha_americana.webp"],
  ["FIX-021", "SIM", "Parafusos & Fixação", "Bucha Americana", "6 mm x 25 mm", "un", "Bucha_americana.webp"],
  ["FIX-022", "SIM", "Parafusos & Fixação", "Bucha Americana", "3/8 Polegada", "un", "Bucha_americana.webp"],
  ["FIX-023", "SIM", "Parafusos & Fixação", "Bucha Americana", "10 mm", "un", "Bucha_americana.webp"],
  ["FIX-024", "SIM", "Parafusos & Fixação", "Cavilha", "10 x 30 mm", "un", "Cavilha.webp"],
  ["FIX-025", "SIM", "Parafusos & Fixação", "Cavilha", "10 x 40 mm", "un", "Cavilha.webp"],
  ["FIX-026", "SIM", "Parafusos & Fixação", "Cavilha", "10 x 50 mm", "un", "Cavilha.webp"],
  ["FIX-027", "SIM", "Parafusos & Fixação", "Cavilha", "8 x 30 mm", "un", "Cavilha.webp"],
  ["FIX-028", "SIM", "Parafusos & Fixação", "Cavilha", "8 x 40 mm", "un", "Cavilha.webp"],
  ["FIX-029", "SIM", "Parafusos & Fixação", "MiniFix Hafele - confirmar diametro e codigo", "Häfele MH 15 Cód. do Item 262.26.037", "un", "Minifix_Hafele.jpg; Minifix_Hafele_foto2.jpg"],
  ["FIX-030", "SIM", "Parafusos & Fixação", "MaxiFix Hafele 35 - 35 mm", "Conector Hafele Maxifix 35 262.87.003", "un", "Maxifix_Hafele35.jpg"],
  ["FIX-031", "SIM", "Parafusos & Fixação", "Tampa de acabamento conector Maxifix 15", "Branca", "un", "Tampa_maxifix15_foto2.jpg; Tampa_maxifix15.jpg"],
  ["FIX-032", "SIM", "Parafusos & Fixação", "Tampa de acabamento conector Maxifix 15", "Preto", "un", "Tampa_maxifix15.jpg; Tampa_maxifix15_foto2.jpg"],
  ["FIX-033", "SIM", "Parafusos & Fixação", "Tampa de acabamento conector Maxifix 15", "Cinza", "un", "Tampa_maxifix15.jpg; Tampa_maxifix15_foto2.jpg"],
  ["FIX-034", "SIM", "Parafusos & Fixação", "Tampa de acabamento pino conector Maxifix 15", "branco", "un", "Tampa_maxifix15_com_pino.jpg"],
  ["FIX-035", "SIM", "Parafusos & Fixação", "Tampa de acabamento pino conector Maxifix 15", "preto", "un", "Tampa_maxifix15_com_pino.jpg"],
  ["FIX-036", "SIM", "Parafusos & Fixação", "Tampa de acabamento pino conector Maxifix 15", "bege", "un", "Tampa_maxifix15_com_pino.jpg"],
  ["FIX-037", "SIM", "Parafusos & Fixação", "Pino Aço tipo T pinador", "40 mm F40", "caixa", "Grampo_T.webp"],
  ["FIX-038", "SIM", "Parafusos & Fixação", "Pino Aço tipo T pinador", "50 mm F50", "caixa", "Grampo_T.webp"],
  ["FIX-039", "SIM", "Parafusos & Fixação", "Pino Aço tipo T pinador", "20 mm F20", "caixa", "Grampo_T.webp"],
  ["FIX-040", "SIM", "Parafusos & Fixação", "Pino Aço tipo T pinador", "30 mm F30", "caixa", "Grampo_T.webp"],
  ["FIX-041", "SIM", "Parafusos & Fixação", "Grampo Grampeador Pneumático", "10 mm X112 mm X 0,6 mm", "caixa", "Grampo_Grampeador_pneumatico.webp"],
  ["FIX-042", "SIM", "Parafusos & Fixação", "Grampo Grampeador Pneumático", "20 mm X 112 x 1,00 mm", "caixa", "Grampo_Grampeador_pneumatico.webp"],
  ["FIX-043", "SIM", "Parafusos & Fixação", "Grampo Grampeador Pneumático", "8 mm 112 mm 0,6 mm", "caixa", "Grampo_Grampeador_pneumatico.webp"],
  ["PIN-001", "SIM", "Pintura e acabamento", "Fundo PU 6468 5 Litros", "Fundo Acabamento PU FL 20 6468", "un", "Sayerlack.png"],
  ["PIN-002", "SIM", "Pintura e acabamento", "Diluente PU 4068 5 litros", "Diluente para PU DF-4068 5,0 litros Sayerlack", "un", "Diluente4068.webp"],
  ["PIN-003", "SIM", "Pintura e acabamento", "Fundo PU 6673 5 litros", "Fundo Acabamento PU FL-6673-005", "un", "Sayerlack.png"],
  ["PIN-004", "SIM", "Pintura e acabamento", "Catalisador PU FC-6975 5 litros", "Catalisador PU FC-6975 5,0 litros", "un", "Sayerlack.png"],
  ["PIN-005", "SIM", "Pintura e acabamento", "Thinner Limpeza 5 litros", "Thinner Limpeza", "un", "Thinner_Limpeza.webp"],
  ["PIN-006", "SIM", "Pintura e acabamento", "Seladora Sayerlack 3,6 litros", "Seladora Sayerlack 3,6 litros", "un", "Seladora_sayerlack.webp"],
  ["PIN-007", "SIM", "Pintura e acabamento", "Polisten Sayerlack Nogueira 3,6 litros", "Polysten Nogueira 3,6 litros", "un", "Polisten_Sayerlack.webp"],
  ["PIN-008", "SIM", "Pintura e acabamento", "Pano Branco", "Pano para boneca", "un", "Pano_Branco.webp"],
  ["PIN-009", "SIM", "Pintura e acabamento", "Filtro Cabine de Pintura", "50 x 50 x 3-cm", "un", "Pintura.webp"],
  ["LIX-001", "SIM", "Lixas", "Banda Larga", "1250 x 2200 mm Grana 80", "un", "lixas.webp"],
  ["LIX-002", "SIM", "Lixas", "Banda Larga", "1250 x 2200 mm Grana 120", "un", "lixas.webp"],
  ["LIX-003", "SIM", "Lixas", "Lixadeira Cinta Grande", "7.195 x 150 mm Grana 60", "un", "lixas.webp"],
  ["LIX-004", "SIM", "Lixas", "Lixadeira Cinta Grande", "7.195 x 150 mm Grana 80", "un", "lixas.webp"],
  ["LIX-005", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 76 mm Grana 60", "un", "lixas.webp"],
  ["LIX-006", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 76 mm Grana 80", "un", "lixas.webp"],
  ["LIX-007", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 76 mm Grana 120", "un", "lixas.webp"],
  ["LIX-008", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 100 mm Grana 60", "un", "lixas.webp"],
  ["LIX-009", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 100 mm Grana 80", "un", "lixas.webp"],
  ["LIX-010", "SIM", "Lixas", "Lixadeira de Cinta Manual", "610 x 100 mm Grana 120", "un", "lixas.webp"],
  ["LIX-011", "SIM", "Lixas", "Lixadeira de Tubinhos", "2400 x 150 Grana 80", "un", "lixas.webp"],
  ["LIX-012", "SIM", "Lixas", "Lixadeira de Tubinhos", "2400 x 150 Grana 120", "un", "lixas.webp"],
  ["LIX-013", "SIM", "Lixas", "Lixa Folha 230 x 280 PP BA11", "Grana 100", "un", "lixas.webp"],
  ["LIX-014", "SIM", "Lixas", "Lixa Folha 230 x 280 PP BA11", "Grana 150", "un", "lixas.webp"],
  ["LIX-015", "SIM", "Lixas", "Lixa Folha 230 x 280 PP BA11", "Grana 220", "un", "lixas.webp"],
  ["LIX-016", "SIM", "Lixas", "Lixa Folha 230 x 280 PP BA11", "Grana 320", "un", "lixas.webp"],
  ["LIX-017", "SIM", "Lixas", "Disco Lixa com velcro 150 mm", "Grana 80", "un", "Disco-lixadeira-150mm-com-velcro.webp"],
  ["LIX-018", "SIM", "Lixas", "Disco Lixa com velcro 150 mm", "Grana 100", "un", "Disco-lixadeira-150mm-com-velcro.webp"],
  ["LIX-019", "SIM", "Lixas", "Disco Lixa com velcro 150 mm", "Grana 120", "un", "Disco-lixadeira-150mm-com-velcro.webp"],
  ["LIX-020", "SIM", "Lixas", "Disco Lixa com velcro 150 mm", "Grana 200", "un", "Disco-lixadeira-150mm-com-velcro.webp"],
  ["LIX-021", "SIM", "Lixas", "Disco Lixa com velcro 150 mm", "Grana 220", "un", "Disco-lixadeira-150mm-com-velcro.webp"],
  ["LIX-022", "SIM", "Lixas", "Lixa Rolinho tupia 100 x 220 mm", "Grana 120", "un", "lixas.webp"],
  ["LIX-023", "SIM", "Lixas", "Lixa Rolinho tupia 100 x 220 mm", "Grana 150", "un", "lixas.webp"],
  ["COL-001", "SIM", "Colas", "Cola de Contato Formica", "Formica Cola de contato 2,8 kg", "un", "Cola_de_contato.webp; formica.jpg"],
  ["COL-002", "SIM", "Colas", "Cola Branca Willbond 10 Normal", "ADESIVO FRANK WILLBOND 10 25 KG", "un", "Willbond-10.jpg"],
  ["COL-003", "SIM", "Colas", "Cola Branca Multibond X-080 - BiComponente", "Multibond X-080 25kg Catalizado", "un", "Multibond_X-080.jpg"],
  ["COL-004", "SIM", "Colas", "Cola Branca ADVANTAGE EP-950A", "ADVANTAGE EP-950A Catalizado 25kg", "un", "Advantage-EP-950A.jpg"],
  ["COL-005", "SIM", "Colas", "Cola TITEBOND ® II Premium", "Madeira de uso interno e externo 20kg", "un", "TiteBond_Premium.jpg"],
  ["COL-006", "SIM", "Colas", "Cascamite / Redemit 5kg", "Cascamite / Redemit 5kg", "un", "Cascamite.webp"],
  ["COL-007", "SIM", "Colas", "Araldite", "Adesivo Epoxi Kit 1,8Kg", "un", "Araldite.webp"],
  ["COL-008", "SIM", "Colas", "Cola Rapida Super Bonder", "Tubo 100Gramas", "un", "SuperBonder.webp"],
  ["FRM-001", "SIM", "Ferramentas", "Brocas HSS", "4 mm", "un", "Broca_HSS.webp"],
  ["FRM-002", "SIM", "Ferramentas", "Brocas HSS", "5 mm", "un", "Broca_HSS.webp"],
  ["FRM-003", "SIM", "Ferramentas", "Brocas HSS", "6 mm", "un", "Broca_HSS.webp"],
  ["FRM-004", "SIM", "Ferramentas", "Brocas HSS", "7 mm", "un", "Broca_HSS.webp"],
  ["FRM-005", "SIM", "Ferramentas", "Brocas HSS", "8 mm", "un", "Broca_HSS.webp"],
  ["FRM-006", "SIM", "Ferramentas", "Brocas HSS", "9 mm", "un", "Broca_HSS.webp"],
  ["FRM-007", "SIM", "Ferramentas", "Brocas HSS", "10 mm", "un", "Broca_HSS.webp"],
  ["FRM-008", "SIM", "Ferramentas", "Brocas HSS", "11 mm", "un", "Broca_HSS.webp"],
  ["FRM-009", "SIM", "Ferramentas", "Brocas HSS", "12 mm", "un", "Broca_HSS.webp"],
  ["FRM-010", "SIM", "Ferramentas", "Brocas de Videa", "6 mm", "un", "Broca_videa.webp"],
  ["FRM-011", "SIM", "Ferramentas", "Brocas de Videa", "8 mm", "un", "Broca_videa.webp"],
  ["FRM-012", "SIM", "Ferramentas", "Brocas de Videa", "10 mm", "un", "Broca_videa.webp"],
  ["FRM-013", "SIM", "Ferramentas", "Fresa CNC Router - Informar comprimento da Haste", "6 mm", "un", "Fresa_CNC.webp"],
  ["FRM-014", "SIM", "Ferramentas", "Fresa CNC Router - Informar comprimento da Haste", "7 mm", "un", "Fresa_CNC.webp"],
  ["FRM-015", "SIM", "Ferramentas", "Fresa CNC Router - Informar comprimento da Haste", "8 mm", "un", "Fresa_CNC.webp"],
  ["FRM-016", "SIM", "Ferramentas", "Fresa CNC Router - Informar comprimento da Haste", "9 mm", "un", "Fresa_CNC.webp"],
  ["FRM-017", "SIM", "Ferramentas", "Fresa CNC Router - Informar comprimento da Haste", "10 mm", "un", "Fresa_CNC.webp"],
  ["FRM-018", "SIM", "Ferramentas", "Fresa Tupia - Informar todas as medidas", "Informar medidas e modelo desejado em OBS", "un", "Fresa_tupia.webp; Fresa_CNC.webp"],
  ["FRM-019", "SIM", "Ferramentas", "Serra de fita Nova", "Informar medidas", "un", "Ferramentas.webp"],
  ["FRM-020", "SIM", "Ferramentas", "Serra de fita Soldar", "Informar quantidade", "un", "Ferramentas.webp"],
  ["FRM-021", "SIM", "Ferramentas", "Serra Circular - mandar medidas e quant de dente", "Especificar medidas", "un", "Ferramentas.webp"],
  ["FRM-022", "SIM", "Ferramentas", "Afiar Fresas", "Informar qual tipo de fresa", "un", "Fresa_tupia.webp; Fresa_CNC.webp"],
  ["EPI-001", "SIM", "EPI e Uniformes", "Luvas de proteção nitrílicas ou de PU", "", "un", "EPI.webp"],
  ["EPI-002", "SIM", "EPI e Uniformes", "Oculos de proteção", "", "un", "Oculos_protecao.jpeg"],
  ["EPI-003", "SIM", "EPI e Uniformes", "Protetor Auricular tipo concha", "", "un", "Protetor_concha.jpeg"],
  ["EPI-004", "SIM", "EPI e Uniformes", "Protetor Auricular tipo Plug", "", "un", "Protetor_plug.webp"],
  ["EPI-005", "SIM", "EPI e Uniformes", "Protetor facial (Face Shield)", "", "un", "protetor-facial-betaface-shield.webp"],
  ["EPI-006", "SIM", "EPI e Uniformes", "Máscara PFF2", "", "un", "Mascara_PFF2.jpeg"],
  ["EPI-007", "SIM", "EPI e Uniformes", "Respirador com filtro duplo facial", "", "un", "mascara-respirador-facial-2-filtros.webp"],
  ["EPI-008", "SIM", "EPI e Uniformes", "Calçado de segurança", "", "un", "Sapato_seguranca.jpeg"],
  ["EPI-009", "SIM", "EPI e Uniformes", "Calça Uniforme - Informar o Tamanho", "", "un", "Calca_uniforme.webp"],
  ["EPI-010", "SIM", "EPI e Uniformes", "Camisa ou Camiseta - Informar o Tamanho", "", "un", "Camiseta.jpeg"],
  ["CHA-001", "SIM", "Chapas", "MDF Cru", "30 mm", "un", "MDF-cru.webp"],
  ["CHA-002", "SIM", "Chapas", "MDF Cru", "25 mm", "un", "MDF-cru.webp"],
  ["CHA-003", "SIM", "Chapas", "MDF Cru", "18 mm", "un", "MDF-cru.webp"],
  ["CHA-004", "SIM", "Chapas", "MDF Cru", "15 mm", "un", "MDF-cru.webp"],
  ["CHA-005", "SIM", "Chapas", "MDF Cru", "12 mm", "un", "MDF-cru.webp"],
  ["CHA-006", "SIM", "Chapas", "MDF Cru", "9 mm", "un", "MDF-cru.webp"],
  ["CHA-007", "SIM", "Chapas", "MDF Cru", "6 mm", "un", "MDF-cru.webp"],
  ["CHA-008", "SIM", "Chapas", "MDF Cru", "3 mm", "un", "MDF-cru.webp"],
  ["CHA-009", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "25 mm", "un", "Compensado_parica.webp"],
  ["CHA-010", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "20 mm", "un", "Compensado_parica.webp"],
  ["CHA-011", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "18 mm", "un", "Compensado_parica.webp"],
  ["CHA-012", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "15 mm", "un", "Compensado_parica.webp"],
  ["CHA-013", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "12 mm", "un", "Compensado_parica.webp"],
  ["CHA-014", "SIM", "Chapas", "Compensado Parica 2,5 x 1,6 m", "10 mm", "un", "Compensado_parica.webp"],
  ["CHA-015", "SIM", "Chapas", "Compensado Flexivel Balcão 2,5 x 1,60 m", "6 mm", "un", "Compensado_flexivel.jpg"],
  ["CHA-016", "SIM", "Chapas", "Compensado Flexivel Balcão 2,5 x 1,60 m", "3 mm", "un", "Compensado_flexivel.jpg"],
  ["CHA-017", "SIM", "Chapas", "Compensado Flexivel Coluna 2,5 x 1,60 m", "6 mm", "un", "Compensado_flexivel.jpg"],
  ["CHA-018", "SIM", "Chapas", "Compensado Flexivel Coluna 2,5 x 1,60 m", "3 mm", "un", "Compensado_flexivel.jpg"],
  ["CHA-019", "SIM", "Chapas", "Formica PF", "Informar Referencia no Campo Observações", "un", "formica.jpg"],
  ["CHA-020", "SIM", "Chapas", "Formica TX", "Informar Referencia no Campo Observações", "un", "formica.jpg"],
  ["MAQ-001", "SIM", "Maquinas", "Furadeira", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-002", "SIM", "Maquinas", "Parafusadeira", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-003", "SIM", "Maquinas", "Lixadeira de Cinta Manual", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-004", "SIM", "Maquinas", "Pistola de Pintura", "Conserto ou troca ? Informar em Obs", "un", "Pintura.webp"],
  ["MAQ-005", "SIM", "Maquinas", "Lixadeira Orbital", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-006", "SIM", "Maquinas", "Lixadeira Oscilante", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-007", "SIM", "Maquinas", "Serra esquadrejadeira manual", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["MAQ-008", "SIM", "Maquinas", "Tupia Manual", "Conserto ou troca ? Informar em Obs", "un", "maquina.webp"],
  ["NIV-001", "SIM", "Sapatas e Niveladores", "Nivelador Nylon AFD1010A", "Diametro 33 mm x 3/8\"", "un", "Pe_nivelador_essentra.jpg; Sapata_AFD1010A.gif"],
  ["NIV-002", "SIM", "Sapatas e Niveladores", "Sapata Redonda Dominó", "BRR055C Preto 25,4 mm", "un", "SapataBRR055A.jpg; SapataBRR055A_2.gif"]
];

/** Substitui TODO o conteúdo das abas Categorias e Itens pelo catálogo acima.
 *  Não mexe em Funcionarios nem em Pedidos. */
function substituirCatalogo() {
  const ss = planilha_();
  const completa = (l, n) => l.concat(Array(Math.max(0, n - l.length)).fill('')).slice(0, n);
  [[ABA.CAT, CAB.CAT, SEED_CATEGORIAS], [ABA.ITENS, CAB.ITENS, SEED_ITENS.map(l => completa(l, CAB.ITENS.length))]].forEach(([nome, cab, linhas]) => {
    const sh = ss.getSheetByName(nome) || criarAba_(ss, nome, cab, []);
    if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
    sh.getRange(1, 1, 1, cab.length).setValues([cab]);
    sh.getRange(2, 1, linhas.length, cab.length).setValues(linhas);
    sh.autoResizeColumns(1, cab.length);
  });
  ss.getSheetByName(ABA.ITENS).getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).build());
  Logger.log('Catálogo substituído: ' + SEED_CATEGORIAS.length + ' categorias, ' + SEED_ITENS.length + ' itens.');
}

// ---------- Instalação ----------
function configurar() {
  const ss = planilha_();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');
  criarAba_(ss, ABA.CAT, CAB.CAT, SEED_CATEGORIAS);
  criarAba_(ss, ABA.ITENS, CAB.ITENS, SEED_ITENS.map(l => l.concat(Array(CAB.ITENS.length - l.length).fill(''))));
  criarAba_(ss, ABA.FUNC, CAB.FUNC, ["Exemplo 1 (apagar)", "Exemplo 2 (apagar)"].map(n => [n, 'SIM', '']));
  const ped = criarAba_(ss, ABA.PED, CAB.PED, []);

  // Listas suspensas e cores
  const sim = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).build();
  ss.getSheetByName(ABA.ITENS).getRange('B2:B').setDataValidation(sim);
  ss.getSheetByName(ABA.FUNC).getRange('B2:B').setDataValidation(sim);
  formatarPedidos_(ped);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN')) props.setProperty('PIN', PIN_PADRAO);
  if (!props.getProperty('SEQ')) props.setProperty('SEQ', '0');
  const padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);
  Logger.log('Configuração concluída. PIN do painel de compras: ' + props.getProperty('PIN'));
}

function formatarPedidos_(ped) {
  ped.getRange('M2:M').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).build());
  ped.getRange('A2:B').setNumberFormat('@');   // texto: evita "0001" virar 1 ou "0001-2" virar data
  ped.getRange('C2:C').setNumberFormat('dd/MM/yyyy HH:mm');
  ped.getRange('N2:N').setNumberFormat('dd/MM/yyyy HH:mm');
  ped.getRange('O2:O').setNumberFormat('dd/MM/yyyy');
  ped.getRange('Q2:Q').setNumberFormat('"R$" #,##0.00');
  ped.getRange('S2:S').setNumberFormat('@');
  ped.getRange('T2:T').setNumberFormat('dd/MM/yyyy');
  ped.getRange('U2:U').setNumberFormat('dd/MM/yyyy HH:mm');
  const r = ped.getRange('A2:U');
  const cor = (txt, bg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$M2="' + txt + '"').setBackground(bg).setRanges([r]).build();
  ped.setConditionalFormatRules([cor('Novo', '#fff4cc'), cor('Pedido ao fornecedor', '#f3e5ff'),
                                 cor('Comprado', '#dbe9ff'), cor('Cancelado', '#eeeeee')]);
}

/** Atualização do módulo de compras (rodar uma vez): cria as colunas Data Compra, Qtd Comprada,
 *  Preço Unit. e Fornecedor na aba Pedidos, troca "Entregue" por "Comprado" e ajusta listas e cores. */
function atualizarCompras() {
  const sh = planilha_().getSheetByName(ABA.PED);
  sh.getRange(1, 1, 1, CAB.PED.length).setValues([CAB.PED])
    .setFontWeight('bold').setBackground('#2f3b4c').setFontColor('#ffffff');
  const n = sh.getLastRow();
  let trocados = 0;
  if (n > 1) {
    const iSt = CAB.PED.indexOf('Status'), iAt = CAB.PED.indexOf('Atualizado em'), iDc = CAB.PED.indexOf('Data Compra');
    const v = sh.getRange(2, 1, n - 1, CAB.PED.length).getValues();
    v.forEach(l => {
      if (String(l[iSt]) === 'Entregue') { l[iSt] = 'Comprado'; trocados++; }
      if (String(l[iSt]) === 'Comprado' && !l[iDc] && l[iAt] instanceof Date) l[iDc] = dia_(l[iAt]);
    });
    sh.getRange(2, iSt + 1, n - 1, CAB.PED.length - iSt).setValues(v.map(l => l.slice(iSt)));
  }
  formatarPedidos_(sh);
  sh.autoResizeColumns(13, 6);
  Logger.log('Compras atualizado. Linhas "Entregue" convertidas: ' + trocados);
}

/** Garante que a aba Pedidos tem todas as colunas novas (roda sozinho; não precisa executar à mão). */
function garantirColunas_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('colunas_ok') === String(CAB.PED.length)) return;
  const sh = planilha_().getSheetByName(ABA.PED);
  const atual = sh.getRange(1, 1, 1, CAB.PED.length).getValues()[0];
  if (atual.some((v, i) => String(v) !== CAB.PED[i])) {
    sh.getRange(1, 1, 1, CAB.PED.length).setValues([CAB.PED])
      .setFontWeight('bold').setBackground('#2f3b4c').setFontColor('#ffffff');
    formatarPedidos_(sh);
  }
  cache.put('colunas_ok', String(CAB.PED.length), 21600);
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
  const sh = planilha_().getSheetByName(nome);
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
    obs: String(p['Obs']), status: String(p['Status'] || 'Novo'),
    dataCompra: p['Data Compra'] instanceof Date ? Utilities.formatDate(p['Data Compra'], TZ_, 'yyyy-MM-dd') : '',
    qtdComprada: p['Qtd Comprada'] === '' || p['Qtd Comprada'] == null ? null : Number(p['Qtd Comprada']),
    preco: p['Preço Unit.'] === '' || p['Preço Unit.'] == null ? null : Number(p['Preço Unit.']),
    fornecedor: String(p['Fornecedor'] ?? ''),
    pedFornecedor: String(p['Nº Ped. Fornecedor'] ?? ''),
    dataPedFornecedor: p['Data Ped. Fornecedor'] instanceof Date ? Utilities.formatDate(p['Data Ped. Fornecedor'], TZ_, 'yyyy-MM-dd') : '',
    vistoEm: p['Visto em'] instanceof Date ? p['Visto em'].toISOString() : ''
  }));
}
const TZ_ = 'America/Sao_Paulo';
const dia_ = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
function lerDia_(s) {                         // 'yyyy-mm-dd' -> Date (meia-noite local)
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function lerNumero_(x) {                       // aceita 12,50 ou 12.50; vazio = null
  if (x === null || x === undefined || String(x).trim() === '') return null;
  const n = Number(String(x).replace(/\./g, (m, i, t) => (t.indexOf(',') >= 0 ? '' : '.')).replace(',', '.'));
  return isFinite(n) ? n : null;
}

// ---------- API ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'catalogo') return json_(catalogo_());
    if (p.action === 'pedidos') {
      if (!pinOk_(p.pin)) return json_({ erro: 'PIN inválido' });
      garantirColunas_();
      marcarVistos_();                                   // compras abriu o painel = pedidos novos "recebidos"
      const dias = Number(p.dias) || 60, limite = Date.now() - dias * 864e5;
      const recente = l => new Date(l['Atualizado em'] || l['Data']).getTime() >= limite;
      const aberto = st => st === 'Novo' || st === 'Pedido ao fornecedor';
      const linhas = linhasPedidos_(l => aberto(String(l['Status'] || 'Novo')) || recente(l));
      return json_({ linhas, fornecedores: fornecedores_(), cadastro: cadastroFornecedores_() });
    }
    if (p.action === 'relatorio') {
      if (!pinOk_(p.pin)) return json_({ erro: 'PIN inválido' });
      const de = lerDia_(p.de), ate = lerDia_(p.ate);
      if (!de || !ate) return json_({ erro: 'Período inválido' });
      const fim = new Date(ate.getTime() + 864e5);
      return json_({ linhas: linhasPedidos_(l => {
        if (String(l['Status']) !== 'Comprado') return false;
        const d = l['Data Compra'] instanceof Date ? l['Data Compra'] : new Date(l['Atualizado em']);
        return d >= de && d < fim;
      }) });
    }
    return json_({ ok: true, app: 'requisicao-materiais' });
  } catch (err) { return json_({ erro: String(err.message || err) }); }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ erro: 'Requisição inválida' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { garantirColunas_(); } catch (err) { /* segue mesmo assim */ }
  try {
    if (body.action === 'login') { senhaOk_(body.funcionario, body.senha); return json_({ ok: true }); }
    if (body.action === 'meus') return json_(meusPedidos_(body));
    if (body.action === 'novoPedido') return json_(novoPedido_(body));
    if (body.action === 'status') return json_(mudarStatus_(body));
    if (body.action === 'comprar') return json_(comprar_(body));
    if (body.action === 'pedirFornecedor') return json_(pedirFornecedor_(body));
    if (body.action === 'editarCompra') return json_(editarCompra_(body));
    if (body.action === 'salvarFornecedor') return json_(salvarFornecedor_(body));
    return json_({ erro: 'Ação desconhecida' });
  } catch (err) {
    const m = String(err.message || err);
    if (m.indexOf('SENHA|') === 0) return json_({ erro: m.slice(6), senha: true });
    return json_({ erro: m });
  } finally { lock.releaseLock(); }
}

// ---------- Senha de 4 números por funcionário (coluna "Senha" da aba Funcionarios) ----------
const erroSenha_ = msg => new Error('SENHA|' + msg);
function senhaOk_(nome, senha) {
  nome = String(nome || '');
  const f = tabela_(ABA.FUNC).find(x => String(x['Nome']) === nome && ativo_(x['Ativo']));
  if (!f) throw erroSenha_('Funcionário não cadastrado');
  const cache = CacheService.getScriptCache(), k = 'tent_' + nome;
  const tent = Number(cache.get(k) || 0);
  if (tent >= 5) throw erroSenha_('Muitas tentativas erradas. Aguarde 15 minutos ou fale com o responsável.');
  let certa = String(f['Senha'] ?? '').trim();
  if (/^\d{1,4}$/.test(certa)) certa = certa.padStart(4, '0');
  if (!certa) throw erroSenha_('Você ainda não tem senha. Fale com o responsável.');
  if (String(senha || '').trim() !== certa) {
    cache.put(k, String(tent + 1), 900);
    throw erroSenha_(tent + 1 >= 5 ? 'Senha incorreta. Acesso bloqueado por 15 minutos.' : 'Senha incorreta');
  }
  cache.remove(k);
  return true;
}

function meusPedidos_(b) {
  senhaOk_(b.funcionario, b.senha);
  const limite = Date.now() - 30 * 864e5;
  return { linhas: linhasPedidos_(l => String(l['Funcionário']) === b.funcionario &&
                                       new Date(l['Data']).getTime() >= limite) };
}

/** Cria a coluna "Senha" na aba Funcionarios e gera senhas de 4 números para quem ainda não tem.
 *  Pode rodar de novo sempre que incluir funcionários: só preenche quem está sem senha. */
function prepararSenhas() {
  const sh = planilha_().getSheetByName(ABA.FUNC);
  sh.getRange(1, 3).setValue('Senha').setFontWeight('bold').setBackground('#2f3b4c').setFontColor('#ffffff');
  sh.getRange('C2:C').setNumberFormat('@');
  const n = sh.getLastRow(); if (n < 2) return;
  const v = sh.getRange(2, 1, n - 1, 3).getValues();
  const usadas = new Set(v.map(r => String(r[2]).trim()).filter(Boolean));
  let novas = 0;
  v.forEach(r => {
    if (String(r[0]).trim() && !String(r[2]).trim()) {
      let s;
      do { s = String(Math.floor(Math.random() * 10000)).padStart(4, '0'); }
      while (usadas.has(s) || /^(\d)\1{3}$/.test(s) || '0123456789 9876543210'.includes(s));
      usadas.add(s); r[2] = s; novas++;
    }
  });
  sh.getRange(2, 3, n - 1, 1).setValues(v.map(r => [String(r[2])]));
  Logger.log('Senhas geradas: ' + novas);
}

function novoPedido_(b) {
  senhaOk_(b.funcionario, b.senha);
  const cat = catalogo_();
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
  const sh = planilha_().getSheetByName(ABA.PED);
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  if (b.chave) CacheService.getScriptCache().put('k_' + b.chave, num, 21600);
  return { ok: true, pedido: num };
}

// ---------- Compras: altera as colunas Status..Fornecedor (M:R) de uma vez ----------
function alterarLinhas_(alteracao) {
  const sh = planilha_().getSheetByName(ABA.PED);
  const n = sh.getLastRow(); if (n < 2) return 0;
  const v = sh.getRange(1, 1, n, CAB.PED.length).getValues();
  const cab = v.shift(), c = {};
  CAB.PED.forEach(nome => { c[nome] = cab.indexOf(nome); if (c[nome] < 0) throw new Error('Coluna "' + nome + '" não encontrada. Rode atualizarCompras.'); });
  let alteradas = 0;
  v.forEach(l => { if (alteracao(l, c)) alteradas++; });
  if (alteradas) {
    const ini = c['Status'];
    sh.getRange(2, ini + 1, v.length, CAB.PED.length - ini).setValues(v.map(l => l.slice(ini)));
  }
  return alteradas;
}

function mudarStatus_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  if (!STATUS.includes(b.status)) throw new Error('Status inválido');
  const alvo = new Set((b.linhas || []).map(String));
  const n = alterarLinhas_((l, c) => {
    if (!(alvo.has(String(l[c['ID Linha']])) || (b.pedido && String(l[c['Nº Pedido']]) === String(b.pedido)))) return false;
    l[c['Status']] = b.status; l[c['Atualizado em']] = new Date();
    return true;
  });
  return { ok: true, alteradas: n };
}

/** Marca "Visto em" nos itens Novos que ainda não tinham (o funcionário vê "Recebido pelas compras"). */
function marcarVistos_() {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const agora = new Date();
    return alterarLinhas_((l, c) => {
      if (String(l[c['Status']] || 'Novo') !== 'Novo' || !l[c['Nº Pedido']] || l[c['Visto em']]) return false;
      l[c['Visto em']] = agora; return true;
    });
  } finally { lock.releaseLock(); }
}

/** Gera um pedido ao fornecedor (nº PF-0001…) para várias linhas. Tudo opcional, menos os itens:
 *  { itens: [{linha, qtd?, preco?}], fornecedor?, data?: 'yyyy-mm-dd' }  (a observação vai só na mensagem) */
function pedirFornecedor_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  if (!Array.isArray(b.itens) || !b.itens.length) throw new Error('Selecione ao menos um item');
  const props = PropertiesService.getScriptProperties();
  const seq = Number(props.getProperty('SEQ_PF') || 0) + 1;
  props.setProperty('SEQ_PF', String(seq));
  const num = 'PF-' + String(seq).padStart(4, '0');
  const porLinha = {}; b.itens.forEach(i => porLinha[String(i.linha)] = i);
  const data = lerDia_(b.data) || dia_(new Date());
  const forn = String(b.fornecedor || '').trim().slice(0, 80);
  const n = alterarLinhas_((l, c) => {
    const it = porLinha[String(l[c['ID Linha']])]; if (!it) return false;
    l[c['Status']] = 'Pedido ao fornecedor'; l[c['Atualizado em']] = new Date();
    l[c['Nº Ped. Fornecedor']] = num; l[c['Data Ped. Fornecedor']] = data;
    if (!l[c['Visto em']]) l[c['Visto em']] = new Date();
    const q = lerNumero_(it.qtd); l[c['Qtd Comprada']] = q === null ? l[c['Qtd']] : q;
    const pr = lerNumero_(it.preco); if (pr !== null) l[c['Preço Unit.']] = pr;
    if (forn) l[c['Fornecedor']] = forn;
    return true;
  });
  return { ok: true, pedidoFornecedor: num, alteradas: n };
}

/** Marca como Comprado várias linhas de uma vez. Tudo opcional:
 *  { itens: [{linha, qtd?, preco?}], data?: 'yyyy-mm-dd', fornecedor? } */
function comprar_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  const porLinha = {}; (b.itens || []).forEach(i => porLinha[String(i.linha)] = i);
  const data = lerDia_(b.data) || dia_(new Date());
  const forn = String(b.fornecedor || '').trim().slice(0, 80);
  const n = alterarLinhas_((l, c) => {
    const it = porLinha[String(l[c['ID Linha']])]; if (!it) return false;
    l[c['Status']] = 'Comprado'; l[c['Atualizado em']] = new Date(); l[c['Data Compra']] = data;
    const q = lerNumero_(it.qtd), jaTinha = l[c['Qtd Comprada']] !== '' && l[c['Qtd Comprada']] != null;
    l[c['Qtd Comprada']] = q !== null ? q : (jaTinha ? l[c['Qtd Comprada']] : l[c['Qtd']]);   // vazio = qtd do pedido ao fornecedor (ou a pedida)
    const pr = lerNumero_(it.preco); if (pr !== null) l[c['Preço Unit.']] = pr;
    if (forn) l[c['Fornecedor']] = forn;
    return true;
  });
  return { ok: true, alteradas: n };
}

/** Corrige/complementa os dados de compra de uma linha (campos vazios ficam vazios). */
function editarCompra_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  const n = alterarLinhas_((l, c) => {
    if (String(l[c['ID Linha']]) !== String(b.linha)) return false;
    if ('data' in b) l[c['Data Compra']] = lerDia_(b.data) || '';
    if ('qtd' in b) { const q = lerNumero_(b.qtd); l[c['Qtd Comprada']] = q === null ? '' : q; }
    if ('preco' in b) { const p = lerNumero_(b.preco); l[c['Preço Unit.']] = p === null ? '' : p; }
    if ('fornecedor' in b) l[c['Fornecedor']] = String(b.fornecedor || '').trim().slice(0, 80);
    l[c['Atualizado em']] = new Date();
    return true;
  });
  return { ok: true, alteradas: n };
}

/** Nomes para sugerir no campo Fornecedor: cadastro (ativos) + nomes já usados em compras. */
function fornecedores_() {
  const s = new Set();
  cadastroFornecedores_().filter(f => f.ativo).forEach(f => s.add(f.nome));
  tabela_(ABA.PED).forEach(l => { const f = String(l['Fornecedor'] ?? '').trim(); if (f) s.add(f); });
  return [...s].sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

// ---------- Cadastro de fornecedores (aba Fornecedores) ----------
function cadastroFornecedores_() {
  const sh = planilha_().getSheetByName(ABA.FORN); if (!sh) return [];
  return tabela_(ABA.FORN).filter(f => String(f['Nome'] || '').trim()).map(f => ({
    codigo: String(f['Código'] || ''), ativo: ativo_(f['Ativo']), nome: String(f['Nome']).trim(),
    contato: String(f['Contato'] ?? ''), telefone: String(f['Telefone / WhatsApp'] ?? ''),
    email: String(f['E-mail'] ?? ''), cnpj: String(f['CNPJ'] ?? ''), obs: String(f['Observações'] ?? '')
  })).sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

/** Inclui (sem código) ou altera (com código) um fornecedor. Só o nome é necessário. */
function salvarFornecedor_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  const nome = String(b.nome || '').trim().slice(0, 80);
  if (!nome) throw new Error('Informe ao menos o nome do fornecedor');
  const sh = planilha_().getSheetByName(ABA.FORN) || criarFornecedores();
  const txt = x => String(x ?? '').trim().slice(0, 200);
  const linha = [null, b.ativo === false ? 'NÃO' : 'SIM', nome, txt(b.contato), txt(b.telefone), txt(b.email), txt(b.cnpj), txt(b.obs)];
  const v = sh.getDataRange().getValues();
  const iCod = v[0].indexOf('Código');
  if (b.codigo) {
    for (let r = 1; r < v.length; r++) if (String(v[r][iCod]) === String(b.codigo)) {
      linha[0] = v[r][0]; sh.getRange(r + 1, 1, 1, linha.length).setValues([linha]); return { ok: true, codigo: String(linha[0]) };
    }
    throw new Error('Fornecedor não encontrado');
  }
  if (v.slice(1).some(l => String(l[2]).trim().toLowerCase() === nome.toLowerCase())) throw new Error('Já existe um fornecedor com esse nome');
  const maior = v.slice(1).reduce((m, l) => Math.max(m, Number(String(l[iCod]).replace(/\D/g, '')) || 0), 0);
  linha[0] = 'F' + String(maior + 1).padStart(3, '0');
  sh.getRange(sh.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);
  return { ok: true, codigo: linha[0] };
}

/** Cria a aba Fornecedores (rodar uma vez) já com os nomes usados nas compras. */
function criarFornecedores() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA.FORN);
  if (!sh) {
    const usados = new Set(); tabela_(ABA.PED).forEach(l => { const f = String(l['Fornecedor'] ?? '').trim(); if (f) usados.add(f); });
    const linhas = [...usados].sort().map((n, i) => ['F' + String(i + 1).padStart(3, '0'), 'SIM', n, '', '', '', '', '']);
    sh = criarAba_(ss, ABA.FORN, CAB.FORN, linhas);
    sh.getRange('A2:A').setNumberFormat('@'); sh.getRange('E2:G').setNumberFormat('@');
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NÃO'], true).build());
    Logger.log('Aba Fornecedores criada com ' + linhas.length + ' fornecedor(es) já usados nas compras.');
  } else Logger.log('A aba Fornecedores já existe.');
  return sh;
}

/** Cria a coluna "Fornecedor" na aba Itens (fornecedor habitual do item — uso interno,
 *  não aparece no app). Lista suspensa com os nomes da aba Fornecedores, mas aceita qualquer texto. */
function criarColunaFornecedorItens() {
  const ss = planilha_(), sh = ss.getSheetByName(ABA.ITENS);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let col = cab.indexOf('Fornecedor') + 1;
  if (!col) {
    col = cab.filter(String).length + 1;
    sh.getRange(1, col).setValue('Fornecedor').setFontWeight('bold').setBackground('#2f3b4c').setFontColor('#ffffff');
  }
  const forn = ss.getSheetByName(ABA.FORN) || criarFornecedores();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(forn.getRange('C2:C'), true).setAllowInvalid(true).build());
  sh.setColumnWidth(col, 200);
  Logger.log('Coluna "Fornecedor" pronta na aba Itens (coluna ' + col + ').');
}

/** BACKUP: cria uma cópia completa da planilha (todas as abas, com formatação) no Google Drive
 *  da conta dona, com nome "Backup Pedidos de Material dd-mm-aaaa hh'h'mm". Pode rodar quando quiser. */
function backupPlanilha() {
  const ss = planilha_();
  const nome = 'Backup ' + ss.getName() + ' ' + Utilities.formatDate(new Date(), TZ_, "dd-MM-yyyy HH'h'mm");
  const copia = ss.copy(nome);
  Logger.log('Backup criado no seu Google Drive: "' + nome + '"\n' + copia.getUrl());
  return copia.getUrl();
}

/** ⚠️ ZERAR PEDIDOS — para colocar o sistema em uso.
 *  1) Guarda uma cópia da aba Pedidos (ex.: "Pedidos (teste) 01-11-2026 10h40");
 *  2) Apaga todas as linhas de pedido (mantém o cabeçalho);
 *  3) Reinicia a numeração: o próximo pedido será o nº 0001.
 *  Não mexe em Itens, Categorias, Funcionarios (nem senhas) ou Fornecedores. */
function zerarPedidos() {
  const ss = planilha_(), sh = ss.getSheetByName(ABA.PED);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const nome = 'Pedidos (teste) ' + Utilities.formatDate(new Date(), TZ_, "dd-MM-yyyy HH'h'mm");
    sh.copyTo(ss).setName(nome);
    if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
    PropertiesService.getScriptProperties().setProperty('SEQ', '0');
    formatarPedidos_(sh);
    Logger.log('Pedidos zerados. Cópia guardada na aba "' + nome + '". Próximo pedido: nº 0001.');
  } finally { lock.releaseLock(); }
}

/** Confere o PIN de compras. Depois de 5 PINs errados seguidos, o painel fica bloqueado
 *  por 15 minutos para todo mundo (impede descobrir o PIN tentando combinações). */
function pinOk_(pin) {
  const cache = CacheService.getScriptCache(), k = 'tent_pin_compras';
  const tent = Number(cache.get(k) || 0);
  if (tent >= 5) throw new Error('PIN bloqueado: muitas tentativas erradas. Aguarde 15 minutos.');
  const certo = String(PropertiesService.getScriptProperties().getProperty('PIN') || PIN_PADRAO);
  if (String(pin || '') === certo) { if (tent) cache.remove(k); return true; }
  cache.put(k, String(tent + 1), 900);
  return false;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
