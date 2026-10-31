# Pede Aí O&C — app de pedido de material

App web para a produção pedir matéria-prima e insumos pelo celular, com fotos, e painel para quem compra.

- **Telas** (`index.html`, `compras.html`) → hospedadas no **GitHub Pages**
- **Dados** (catálogo, funcionários, pedidos) → **Planilha Google**, gravada pelo **Apps Script** (`apps-script/Codigo.gs`)

Sem a URL configurada, o app abre em **modo demonstração**: dá para testar tudo, mas os pedidos ficam só no aparelho. O PIN do painel de compras na demonstração é `1234`.

---

## Instalação atual (já feita)

- Planilha: **Pedidos de Material** (Google Drive de marcos.markevich@gmail.com)
- Apps Script: projeto independente ("Projeto sem título") em script.google.com, ligado à planilha pelo `PLANILHA_ID`
- `config.js` já aponta para o App da Web publicado

As seções 1 e 2 abaixo ficam como referência para reinstalar do zero.

## 1. Planilha + Apps Script (cerca de 10 min)

1. Crie uma Planilha Google nova, por exemplo "Pedidos de Material".
2. Abra **Extensões → Apps Script**, apague o conteúdo do editor, cole todo o arquivo `apps-script/Codigo.gs` e clique em **Salvar**.
3. Na barra de cima, escolha a função **`configurar`** e clique em **Executar**.
   - O Google pede autorização. Se aparecer "O Google não verificou este app", clique em **Avançado → Acessar (não seguro)**. O app é seu, então isso é normal.
   - O script cria as abas **Categorias, Itens, Funcionarios e Pedidos**, já com o catálogo de exemplo.
4. Clique em **Implantar → Nova implantação**, no ícone de engrenagem escolha **App da Web** e configure:
   - **Executar como:** Eu
   - **Quem pode acessar:** Qualquer pessoa
   - Clique em **Implantar** e copie a **URL do app da Web** (termina em `/exec`).
5. **Troque o PIN de compras** (o padrão é 1234): na função `trocarPin`, troque `'1234'` pelo PIN novo, salve e execute `trocarPin`.

> Se alterar o código depois: **Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar**. Assim a URL continua a mesma.

## 2. GitHub Pages

1. Crie um repositório, por exemplo `pedido-material`. Na conta gratuita o GitHub Pages exige repositório **público**.
2. Envie os arquivos desta pasta: `index.html`, `compras.html`, `config.js`, `comum.js`, `demo.js`, `estilo.css` e a pasta `img/`. A pasta `apps-script/` é opcional, serve só como cópia de segurança.
3. Edite o `config.js`:
   ```js
   window.API_URL = 'https://script.google.com/macros/s/XXXX/exec';
   window.PLANILHA_URL = 'https://docs.google.com/spreadsheets/d/XXXX';
   window.NOME_EMPRESA = 'Pedido de Material';
   ```
4. Vá em **Settings → Pages → Source: Deploy from a branch → main / (root) → Save**. Em 1 a 2 minutos o link fica ativo:
   - Produção: `https://SEU-USUARIO.github.io/pedido-material/`
   - Compras: `https://SEU-USUARIO.github.io/pedido-material/compras.html`
5. Em cada celular: abra o link e use **Adicionar à tela inicial**. O app vira um ícone como os outros aplicativos.

## 3. Uso no dia a dia

**Funcionário:** toca no próprio nome e digita a senha de 4 números. Fica conectado por 12 horas (um turno); depois o app pede a senha de novo. Erro de senha 5 vezes bloqueia aquele nome por 15 minutos. Depois disso, escolhe a categoria, toca no item, ajusta a quantidade e envia. Também pode marcar o pedido como urgente, pedir um item que não está na lista e acompanhar o status em "Meus pedidos".

**Compras:** abre `compras.html` e entra com o PIN.
- **Novos:** pedidos a atender, com os urgentes no topo. Marque o pedido inteiro, ou só as linhas selecionadas, como Comprado, Entregue ou Cancelado.
- **Lista de compra:** soma de tudo que está pendente, agrupada por item, com botão para copiar e mandar no WhatsApp do fornecedor.
- O status também pode ser mudado direto na aba **Pedidos** da planilha (coluna Status, com lista suspensa).

## 4. Manutenção do catálogo (direto na planilha)

O catálogo foi carregado a partir de **Lista de insumos.xlsx** (145 itens). Para recarregar tudo de uma vez a partir do código, use a função `substituirCatalogo` no Apps Script. Ela sobrescreve as abas Categorias e Itens, mas não mexe em Funcionarios nem em Pedidos. No dia a dia, edite direto na planilha.


| Aba | O que fazer |
|---|---|
| **Funcionarios** | Nome, Ativo e **Senha** (4 números). Para esconder alguém, use Ativo = NÃO. Ao incluir alguém, deixe a Senha vazia e rode `prepararSenhas` no Apps Script, ou digite você mesmo os 4 números. |
| **Itens** | Uma linha por item. O Código precisa ser único (ex.: FIX-009). Para tirar um item do app, use Ativo = NÃO em vez de apagar a linha. |
| **Categorias** | Nome, foto e ordem em que aparecem na tela inicial. |

O app busca o catálogo atualizado toda vez que é aberto.

**Fotos (coluna Foto)** aceitam três formatos:
1. Nome de um arquivo da pasta `img/` do repositório, por exemplo `parafuso-4x40.jpg`. Tire a foto com o celular, envie para a pasta `img/` no GitHub e escreva o nome na planilha. Use nomes sem espaço nem acento.
2. Link de compartilhamento do **Google Drive**. O arquivo precisa estar compartilhado como "Qualquer pessoa com o link".
3. Qualquer link de imagem que comece com `https://`.

**Duas fotos (ou mais) no mesmo item:** escreva os nomes na mesma célula separados por ponto e vírgula, por exemplo `dobradica-reta.jpg; dobradica-reta-lado.jpg`. A primeira aparece no botão do item; na janela de quantidade aparecem todas lado a lado, e tocando nelas a foto amplia.

Dica: fotografe o item real sobre um fundo liso. Para o marceneiro, isso funciona melhor que o desenho.

## Observações

- O pedido grava **uma linha por item** na aba Pedidos, com Nº do pedido, funcionário, data, código, quantidade e status. Isso facilita filtrar, montar tabela dinâmica ou cruzar com outras planilhas.
- Se a internet cair na hora de enviar, o pedido continua salvo no celular e pode ser reenviado. O sistema não duplica o pedido.
- A URL do Apps Script fica visível no repositório público. Com ela é possível ver o catálogo e enviar pedidos com um nome cadastrado. Já mudar status e ver todos os pedidos exige o PIN.
