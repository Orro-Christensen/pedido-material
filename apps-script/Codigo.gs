/**
 * App de Requisição de Materiais — backend (Google Apps Script)
 * Projeto Apps Script independente (script.google.com) ligado à planilha pelo PLANILHA_ID.
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
// ID da Planilha Google (o trecho entre /d/ e /edit no endereço da planilha).
const PLANILHA_ID = '12xfUT4sbeSyINRmJhTsSd3VkouEKVhEJYkCTwraNL5E';
function planilha_() { return SpreadsheetApp.openById(PLANILHA_ID); }

const STATUS = ['Novo', 'Comprado', 'Entregue', 'Cancelado'];
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
  [[ABA.CAT, CAB.CAT, SEED_CATEGORIAS], [ABA.ITENS, CAB.ITENS, SEED_ITENS]].forEach(([nome, cab, linhas]) => {
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
  criarAba_(ss, ABA.ITENS, CAB.ITENS, SEED_ITENS);
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
  const sh = planilha_().getSheetByName(ABA.PED);
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  if (b.chave) CacheService.getScriptCache().put('k_' + b.chave, num, 21600);
  return { ok: true, pedido: num };
}

function mudarStatus_(b) {
  if (!pinOk_(b.pin)) throw new Error('PIN inválido');
  if (!STATUS.includes(b.status)) throw new Error('Status inválido');
  const sh = planilha_().getSheetByName(ABA.PED);
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
