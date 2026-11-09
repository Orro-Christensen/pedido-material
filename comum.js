// Funções compartilhadas entre o app da produção (index.html) e o painel de compras (compras.html).
(function () {
  // Raiz do app = pasta onde está o comum.js (as páginas ficam em producao/ e compras/)
  const RAIZ = new URL('.', (document.currentScript && document.currentScript.src) || location.href).href;
  const DEMO = !window.API_URL;
  const ls = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} }
  };

  async function comTempo(p, ms) {
    let t; const limite = new Promise((_, rej) => t = setTimeout(() => rej(new Error('Sem resposta do servidor')), ms));
    try { return await Promise.race([p, limite]); } finally { clearTimeout(t); }
  }

  async function get(params) {
    if (DEMO) return demoGet(params);
    const r = await comTempo(fetch(window.API_URL + '?' + new URLSearchParams(params)), 20000);
    const j = await r.json(); if (j.erro) throw new Error(j.erro); return j;
  }
  async function post(body) {
    if (DEMO) return demoPost(body);
    // Sem cabeçalho Content-Type: o navegador envia text/plain e o Apps Script aceita sem bloqueio de CORS.
    const r = await comTempo(fetch(window.API_URL, { method: 'POST', body: JSON.stringify(body) }), 30000);
    const j = await r.json();
    if (j.erro) { const e = new Error(j.erro); e.senha = !!j.senha; throw e; }
    return j;
  }

  // ---------- modo demonstração (localStorage) ----------
  function demoGet(p) {
    const linhas = ls.get('demo_pedidos', []);
    if (p.action === 'catalogo') return Promise.resolve(window.DEMO);
    if (p.action === 'pedidos' || p.action === 'relatorio') {
      if (p.pin !== '1234') return Promise.reject(new Error('PIN inválido'));
      if (p.action === 'relatorio') return Promise.resolve({ linhas: linhas.filter(l => l.status === 'Comprado' && l.dataCompra >= p.de && l.dataCompra <= p.ate) });
      linhas.forEach(l => { if (l.status === 'Novo' && !l.vistoEm) l.vistoEm = new Date().toISOString(); }); ls.set('demo_pedidos', linhas);
      const cad = ls.get('demo_forn', []);
      return Promise.resolve({ linhas, cadastro: cad, fornecedores: [...new Set(cad.filter(f => f.ativo).map(f => f.nome).concat(linhas.map(l => l.fornecedor).filter(Boolean)))] });
    }
    return Promise.resolve({ ok: true });
  }
  function demoPost(b) {
    const linhas = ls.get('demo_pedidos', []);
    if (['login', 'meus', 'novoPedido'].includes(b.action) && b.senha !== '1234') {
      const e = new Error('Senha incorreta'); e.senha = true; return Promise.reject(e);
    }
    if (b.action === 'login') return new Promise(r => setTimeout(() => r({ ok: true }), 300));
    if (b.action === 'meus') return Promise.resolve({ linhas: linhas.filter(l => l.funcionario === b.funcionario) });
    if (b.action === 'novoPedido') {
      const seq = ls.get('demo_seq', 0) + 1; ls.set('demo_seq', seq);
      const num = String(seq).padStart(4, '0'); const mapa = {};
      window.DEMO.itens.forEach(i => mapa[i.id] = i);
      b.itens.forEach((it, n) => {
        const c = mapa[it.id];
        linhas.push({ pedido: num, linha: num + '-' + (n + 1), data: new Date().toISOString(), funcionario: b.funcionario,
          urgente: !!b.urgente, categoria: c ? c.categoria : 'Outro (fora do catálogo)', codigo: c ? c.id : 'OUTRO',
          item: c ? c.nome : it.nome, espec: c ? c.espec : '', qtd: it.qtd, unidade: c ? c.unidade : it.unidade,
          obs: it.obs || b.obs || '', status: 'Novo' });
      });
      ls.set('demo_pedidos', linhas);
      return new Promise(r => setTimeout(() => r({ ok: true, pedido: num }), 500));
    }
    const num_ = x => { x = String(x ?? '').trim(); if (!x) return null; const n = Number(x.includes(',') ? x.replace(/\./g, '').replace(',', '.') : x); return isFinite(n) ? n : null; };
    if (b.action === 'pedirFornecedor') {
      if (b.pin !== '1234') return Promise.reject(new Error('PIN inválido'));
      const seq = ls.get('demo_pf', 0) + 1; ls.set('demo_pf', seq); const num = 'PF-' + String(seq).padStart(4, '0');
      linhas.forEach(l => { const it = (b.itens || []).find(i => i.linha === l.linha); if (!it) return;
        Object.assign(l, { status: 'Pedido ao fornecedor', pedFornecedor: num, dataPedFornecedor: b.data || new Date().toISOString().slice(0, 10),
          qtdComprada: num_(it.qtd) ?? l.qtd, preco: num_(it.preco) ?? l.preco ?? null, fornecedor: b.fornecedor || l.fornecedor || '',
          vistoEm: l.vistoEm || new Date().toISOString() }); });
      ls.set('demo_pedidos', linhas); return Promise.resolve({ ok: true, pedidoFornecedor: num });
    }
    if (b.action === 'comprar' || b.action === 'editarCompra') {
      if (b.pin !== '1234') return Promise.reject(new Error('PIN inválido'));
      const hoje = new Date().toISOString().slice(0, 10);
      linhas.forEach(l => {
        if (b.action === 'comprar') { const it = (b.itens || []).find(i => i.linha === l.linha); if (!it) return;
          Object.assign(l, { status: 'Comprado', dataCompra: b.data || hoje, qtdComprada: num_(it.qtd) ?? l.qtdComprada ?? l.qtd,
            preco: num_(it.preco) ?? l.preco ?? null, fornecedor: b.fornecedor || l.fornecedor || '' }); }
        else if (l.linha === b.linha) Object.assign(l, { dataCompra: b.data || '', qtdComprada: num_(b.qtd), preco: num_(b.preco), fornecedor: b.fornecedor || '' });
      });
      ls.set('demo_pedidos', linhas); return Promise.resolve({ ok: true });
    }
    if (b.action === 'salvarFornecedor') {
      const cad = ls.get('demo_forn', []); const f = { ...b }; delete f.action; delete f.pin;
      if (!f.codigo) f.codigo = 'F' + String(cad.length + 1).padStart(3, '0');
      const i = cad.findIndex(x => x.codigo === f.codigo); if (i >= 0) cad[i] = f; else cad.push(f);
      ls.set('demo_forn', cad); return Promise.resolve({ ok: true, codigo: f.codigo });
    }
    if (b.action === 'status') {
      if (b.pin !== '1234') return Promise.reject(new Error('PIN inválido'));
      const alvo = new Set(b.linhas || []); let n = 0;
      linhas.forEach(l => { if (alvo.has(l.linha) || l.pedido === b.pedido) { l.status = b.status; n++; } });
      ls.set('demo_pedidos', linhas);
      return Promise.resolve({ ok: true, alteradas: n });
    }
    return Promise.reject(new Error('Ação desconhecida'));
  }

  // Foto: nome de arquivo em img/, URL completa, ou link de compartilhamento do Google Drive.
  function foto(f) {
    f = String(f || '').trim();
    if (!f) return RAIZ + 'img/outro.svg';
    const drive = f.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?id=)([\w-]+)/);
    if (drive) return 'https://drive.google.com/thumbnail?id=' + drive[1] + '&sz=w600';
    if (/^https?:\/\//.test(f)) return f;
    return RAIZ + 'img/' + f;
  }

  // Várias fotos na mesma célula, separadas por ";" → lista de endereços.
  function fotos(f) {
    const l = String(f || '').split(/\s*;\s*/).filter(Boolean).map(foto);
    return l.length ? l : [foto('')];
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function dataBR(iso) {
    const d = new Date(iso); if (isNaN(d)) return '';
    return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' }) + ' ' +
           d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }
  function qtdBR(n) { return Number(n).toLocaleString('pt-BR', { maximumFractionDigits: 2 }); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 10); }

  // ---------- instalação como app (Android e iPhone) ----------
  let promptInstalar = null;
  const ouvintes = [];
  const avisar = () => ouvintes.forEach(f => { try { f(); } catch (e) {} });
  const ua = navigator.userAgent;
  const ehIOS = /iphone|ipad|ipod/i.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const ehAndroid = /android/i.test(ua);
  const instalado = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); promptInstalar = e; avisar(); });
  window.addEventListener('appinstalled', () => { promptInstalar = null; avisar(); });
  // sw.js fica na raiz do app (mesma pasta do comum.js), também quando a página está em compras/
  const urlSW = RAIZ + 'sw.js';
  if ('serviceWorker' in navigator) navigator.serviceWorker.register(urlSW).catch(() => {});

  function ajudaInstalar(nomeApp) {
    const alvo = document.getElementById('modal'); if (!alvo) return;
    const passos = ehIOS ? `
      <ol class="passos">
        <li>Abra este link no <b>Safari</b> (no iPhone com iOS 16.4 ou mais novo, o Chrome também serve).</li>
        <li>Toque no botão <b>Compartilhar</b> — o quadrado com a seta para cima <b>⬆</b> (embaixo no Safari; no Chrome, no alto à direita).</li>
        <li>Role a lista e toque em <b>"Adicionar à Tela de Início"</b>.</li>
        <li>Confira o nome <b>${esc(nomeApp)}</b> e toque em <b>Adicionar</b>.</li>
      </ol>
      <p class="meta">O iPhone nunca mostra um aviso automático de instalação — é sempre por esse caminho.</p>` : `
      <ol class="passos">
        <li>Abra este link no <b>Chrome</b>.</li>
        <li>Toque no menu <b>⋮</b> (três pontinhos, no alto à direita).</li>
        <li>Toque em <b>"Instalar app"</b> ou <b>"Adicionar à tela inicial"</b>.</li>
        <li>Confirme em <b>Instalar</b>. O ícone <b>${esc(nomeApp)}</b> aparece junto dos outros aplicativos.</li>
      </ol>
      <p class="meta">Se a opção não aparecer, talvez o app já esteja instalado neste aparelho — procure o ícone na tela inicial.</p>`;
    alvo.innerHTML = `<div class="fundo-modal" id="fmInst"><div class="modal">
      <div class="titulo" style="margin-top:0">📲 Instalar ${esc(nomeApp)}</div>${passos}
      <button class="btn btn-claro" id="instX">Fechar</button></div></div>`;
    document.getElementById('instX').onclick = () => alvo.innerHTML = '';
    document.getElementById('fmInst').onclick = e => { if (e.target.id === 'fmInst') alvo.innerHTML = ''; };
  }

  async function instalar(nomeApp) {
    if (promptInstalar) {
      const p = promptInstalar; promptInstalar = null;
      p.prompt();
      try { await p.userChoice; } catch (e) {}
      avisar(); return;
    }
    ajudaInstalar(nomeApp);
  }

  const instalacao = { ehIOS, ehAndroid, instalado, podeMostrar: () => !instalado(), instalar, aoMudar: f => ouvintes.push(f) };

  window.App = { DEMO, instalacao, ls, get, post, foto: f => fotos(f)[0], fotos, esc, dataBR, qtdBR, uid };
})();
