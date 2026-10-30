// Funções compartilhadas entre o app da produção (index.html) e o painel de compras (compras.html).
(function () {
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
    const j = await r.json(); if (j.erro) throw new Error(j.erro); return j;
  }

  // ---------- modo demonstração (localStorage) ----------
  function demoGet(p) {
    const linhas = ls.get('demo_pedidos', []);
    if (p.action === 'catalogo') return Promise.resolve(window.DEMO);
    if (p.action === 'meus') return Promise.resolve({ linhas: linhas.filter(l => l.funcionario === p.funcionario) });
    if (p.action === 'pedidos') return Promise.resolve(p.pin === '1234' ? { linhas } : { erro: 'PIN inválido' })
      .then(j => { if (j.erro) throw new Error(j.erro); return j; });
    return Promise.resolve({ ok: true });
  }
  function demoPost(b) {
    const linhas = ls.get('demo_pedidos', []);
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
    if (!f) return 'img/outro.svg';
    const drive = f.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?id=)([\w-]+)/);
    if (drive) return 'https://drive.google.com/thumbnail?id=' + drive[1] + '&sz=w600';
    if (/^https?:\/\//.test(f)) return f;
    return 'img/' + f;
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

  window.App = { DEMO, ls, get, post, foto, esc, dataBR, qtdBR, uid };
})();
