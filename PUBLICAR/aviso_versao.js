/* ==========================================================================
   AVISO DE VERSAO NOVA dos paineis (23/10/2026 -- pedido: "quando eu atualizar
   algo informar igual no painel faturamento e mandar reiniciar falando que nao
   salva ou em caso de erro" + "a opcao de incluir no windows quando houver
   atualizacao de um painel especifico")
   --------------------------------------------------------------------------
   1. VERSAO NOVA: de 2 em 2 min (e antes de gravar, no maximo 1x por minuto)
      confere a impressao do painel publicado = ETag do proprio endereco e dos
      scripts do mesmo site. Mudou = faixa no topo "O PAINEL FOI ATUALIZADO"
      pedindo F5, e TODA gravacao desta aba volta recusada (503) sem chegar ao
      servidor: a aba da versao anterior nunca grava. Modelo: painel OPERACAO.
      A recusa sai como 503 de proposito: as filas dos paineis (checklist,
      ponto, pessoal...) guardam o que voltou 0/5xx para tentar depois e
      DESCARTAM o resto -- com 409/423 o que estava na fila se perderia.
   2. NAO SALVOU: gravacao que falhou por rede, erro 5xx do servidor ou
      resposta que nao e' a do servidor (pagina HTML no lugar) -> faixa
      vermelha dizendo que NAO salvou e pedindo F5 + fazer de novo.
   3. AVISO NO WINDOWS: botao do sino (canto de baixo, a esquerda). Ligado, a
      versao nova tambem sai como notificacao do Windows enquanto o painel
      estiver aberto (mesmo em outra aba ou janela). Desligado por padrao.

   Gravacao = metodo diferente de GET/HEAD/OPTIONS para
     - /api/... do proprio site (menos /api/entrar e /api/sair: entrar/sair);
     - Supabase /rest/v1, /storage/v1, /functions/v1 (menos /auth/v1, hub-sso
       e as leituras por POST do storage: sign e list).
   Configuracao pelos atributos da tag (CSP: nada de script inline):
     <script src="aviso_versao.js" data-painel="pontodp" data-nome="PONTO DP"></script>
     data-grava="0"    painel so de leitura: a faixa so pede F5, nada a bloquear
     data-erros="0"    painel que ja avisa sozinho quando nao grava (fila com
                       nova tentativa): sem a faixa vermelha (F5 perderia a fila)
     data-livres="a,b" trechos de endereco que nunca contam como gravacao
     data-sino="0"     sem o botao do sino
     data-fila="1"     painel que guarda no aparelho o que nao subiu (checklist, ponto):
                       a faixa diz que o guardado sobe depois de atualizar (sem refazer)
   Tem que carregar ANTES dos outros scripts: o supabase-js guarda o fetch ao
   criar o cliente. Copia mestre: PAINEIS\_COMUM\aviso_versao.js -- cada painel
   leva uma copia (sincronizar_aviso_versao.py iguala as copias).
   ========================================================================== */
(function () {
  'use strict';
  if (window.AvisoVersao || !window.fetch) return;

  var eu = document.currentScript;
  function attr(n, padrao) {
    var v = eu && eu.getAttribute(n);
    return v === null || v === undefined || v === '' ? padrao : v;
  }
  var CFG = {
    painel: attr('data-painel', location.hostname.split('.')[0] || 'painel'),
    nome: attr('data-nome', ''),
    grava: attr('data-grava', '1') !== '0',
    erros: attr('data-erros', '1') !== '0',
    sino: attr('data-sino', '1') !== '0',
    fila: attr('data-fila', '0') === '1',
    livres: attr('data-livres', '').split(',').map(function (s) { return s.trim(); }).filter(Boolean)
  };
  var FONTE = 'Calibri,Carlito,"Segoe UI",Arial,sans-serif';
  var Z = 2147483000;
  var FETCH0 = window.fetch.bind(window);
  var BASE = null;                 // impressao da versao aberta nesta aba (null = ainda nao lida)
  var NOVA = false;                // ja ha versao nova publicada
  var MODO = {};                   // endereco -> 'get' quando o HEAD nao traz ETag
  var confEm = 0, confTudoEm = 0, conferindo = null;
  var CH_SINO = 'aviso_windows::' + CFG.painel;
  var caixa = null, faixaV = null, faixaE = null, sino = null, recado = null, recadoT = null, falhas = 0;

  function nome() { return (CFG.nome || document.title || 'O PAINEL').trim(); }
  var CELULAR = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent || '');
  var F5 = CELULAR ? 'Toque em ATUALIZAR AGORA' : 'Aperte F5';
  function ler(ch) { try { return window.localStorage.getItem(ch); } catch (e) { return null; } }
  function guardar(ch, v) { try { window.localStorage.setItem(ch, v); } catch (e) { /* sem storage: vale so nesta aba */ } }

  // ---------- impressao da versao publicada ----------
  function limpo(u) { var x = new URL(u, location.href); return x.origin + x.pathname + x.search; }
  function comMarca(u) { return u + (u.indexOf('?') < 0 ? '?' : '&') + '__versao=' + Date.now(); }
  function enderecos(tudo) {
    var l = [location.origin + location.pathname];
    if (!tudo) return l;
    var nos = document.querySelectorAll('script[src]');
    for (var i = 0; i < nos.length; i++) {
      var x;
      try { x = new URL(nos[i].src, location.href); } catch (e) { continue; }
      if (x.origin !== location.origin) continue;
      var u = limpo(x.href);
      if (l.indexOf(u) < 0) l.push(u);
    }
    return l;
  }
  function resumo(t) {
    // impressao do conteudo quando nao ha ETag (pagina montada por funcao)
    var h1 = 0x811c9dc5, h2 = 0x01000193;
    for (var i = 0; i < t.length; i++) {
      var c = t.charCodeAt(i);
      h1 = Math.imul(h1 ^ c, 16777619) >>> 0;
      h2 = Math.imul(h2 + c, 2246822519) >>> 0;
    }
    return 'h' + h1.toString(16) + h2.toString(16) + '.' + t.length;
  }
  function etiquetaDe(r) {
    var e = r.headers.get('etag');
    return e ? e.replace(/^W\//, '') : '';
  }
  async function etiqueta(u, primeira) {
    var r;
    if (primeira) {
      // a 1a leitura pega o que o navegador JA tem (o que esta aberto nesta aba), nao o publicado agora
      r = await FETCH0(u, { cache: 'force-cache', credentials: 'same-origin' });
      if (r.status === 404) return '404';
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return etiquetaDe(r) || resumo(await r.text());
    }
    if (MODO[u] !== 'get') {
      r = await FETCH0(comMarca(u), { method: 'HEAD', cache: 'no-store', credentials: 'same-origin' });
      if (r.status === 404) return '404';
      if (r.ok && etiquetaDe(r)) return etiquetaDe(r);
      MODO[u] = 'get';
    }
    r = await FETCH0(comMarca(u), { cache: 'no-store', credentials: 'same-origin' });
    if (r.status === 404) return '404';
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return etiquetaDe(r) || resumo(await r.text());
  }
  function conferir(tudo) {
    if (NOVA) return Promise.resolve(true);
    if (conferindo) return conferindo;
    conferindo = (async function () {
      var primeira = !BASE;
      var lista = enderecos(tudo || primeira);
      var vals = await Promise.all(lista.map(function (u) { return etiqueta(u, primeira); }));
      confEm = Date.now();
      if (tudo || primeira) confTudoEm = confEm;
      if (primeira) {
        BASE = {};
        lista.forEach(function (u, i) { BASE[u] = vals[i]; });
        return false;
      }
      for (var i = 0; i < lista.length; i++) {
        if (BASE[lista[i]] === undefined) { BASE[lista[i]] = vals[i]; continue; }   // script que entrou depois
        if (BASE[lista[i]] !== vals[i]) { versaoNova(); return true; }
      }
      return false;
    })().catch(function () {
      return false;                                                 // sem rede agora: confere na proxima
    }).then(function (v) { conferindo = null; return v; });
    return conferindo;
  }

  // ---------- o que e' gravacao ----------
  function ehGravacao(url, metodo) {
    if (/^(GET|HEAD|OPTIONS)$/.test(metodo)) return false;
    var u;
    try { u = new URL(url, location.href); } catch (e) { return false; }
    for (var i = 0; i < CFG.livres.length; i++) if (u.href.indexOf(CFG.livres[i]) >= 0) return false;
    if (u.origin === location.origin) {
      return /^\/api\//.test(u.pathname) && !/^\/api\/(entrar|sair)(\.js)?\/?$/.test(u.pathname);
    }
    if (/\.supabase\.(co|in)$/.test(u.hostname)) {
      if (!/^\/(rest|storage|functions)\/v1\//.test(u.pathname)) return false;
      if (/^\/functions\/v1\/hub-sso/.test(u.pathname)) return false;
      if (/^\/storage\/v1\/object\/(sign|list)\//.test(u.pathname)) return false;
      return true;
    }
    return false;
  }
  function recusada() {
    var msg = 'O PAINEL FOI ATUALIZADO: aperte F5 (ATUALIZAR) antes de gravar. Esta tela é da versão anterior e NÃO salva mais nada.';
    faixaVersao();
    return new Response(JSON.stringify({ ok: false, erro: msg, error: msg, message: msg, code: 'PAINEL_ATUALIZADO' }),
                        { status: 503, statusText: 'Painel atualizado', headers: { 'Content-Type': 'application/json', 'Retry-After': '60' } });
  }
  async function gravar(u, o) {
    if (CFG.grava) {
      // keepalive = gravacao de quem esta fechando a aba: nao pode esperar a conferencia
      var podeEsperar = !(o && o.keepalive) && !(u && typeof u === 'object' && u.keepalive);
      if (!NOVA && podeEsperar && Date.now() - confTudoEm > 60000) await conferir(true);
      if (NOVA) return recusada();
    }
    var r;
    try {
      r = await FETCH0(u, o);
    } catch (e) {
      if (CFG.grava && CFG.erros && !(e && e.name === 'AbortError')) {
        naoSalvou('sem conexão com o servidor (a internet caiu ou o servidor não respondeu)');
      }
      throw e;
    }
    if (CFG.grava && CFG.erros) {
      if (r.status >= 500) naoSalvou('o servidor respondeu com erro ' + r.status);
      else if (r.ok && /text\/html/i.test(r.headers.get('content-type') || '')) naoSalvou('o servidor devolveu uma página no lugar da confirmação');
    }
    return r;
  }
  window.fetch = function (u, o) {
    var metodo = String((o && o.method) || (u && typeof u === 'object' && u.method) || 'GET').toUpperCase();
    var url = String((u && typeof u === 'object' && u.url) || u || '');
    if (!ehGravacao(url, metodo)) return FETCH0(u, o);
    return gravar(u, o);
  };
  if (navigator.sendBeacon) {
    var BEACON0 = navigator.sendBeacon.bind(navigator);
    navigator.sendBeacon = function (url, dados) {
      if (CFG.grava && NOVA && ehGravacao(url, 'POST')) { faixaVersao(); return false; }
      return BEACON0(url, dados);
    };
  }

  // ---------- tela ----------
  function el(tag, css, txt) {
    var n = document.createElement(tag);
    if (css) n.style.cssText = css;
    if (txt !== undefined) n.textContent = txt;
    return n;
  }
  function pronto(f) {
    if (document.body) f();
    else document.addEventListener('DOMContentLoaded', f);
  }
  function caixaTopo() {
    if (caixa && caixa.isConnected) return caixa;
    caixa = el('div', 'position:fixed;left:50%;top:10px;transform:translateX(-50%);z-index:' + Z + ';display:flex;' +
                      'flex-direction:column;gap:8px;align-items:stretch;width:max-content;max-width:calc(100% - 24px);' +
                      'font:14px/1.35 ' + FONTE + ';pointer-events:none');
    caixa.setAttribute('data-aviso-versao', 'topo');
    document.body.appendChild(caixa);
    return caixa;
  }
  function faixa(fundo, cor, borda) {
    return el('div', 'pointer-events:auto;display:flex;gap:12px;align-items:center;flex-wrap:wrap;background:' + fundo +
                     ';color:' + cor + ';border:1px solid ' + borda + ';border-radius:10px;padding:10px 14px;' +
                     'box-shadow:0 8px 24px rgba(0,0,0,.22);text-align:left');
  }
  function botao(txt, fundo, cor, f) {
    var b = el('button', 'background:' + fundo + ';color:' + cor + ';border:0;border-radius:8px;padding:7px 12px;' +
                         'font:700 13px ' + FONTE + ';cursor:pointer;white-space:nowrap', txt);
    b.type = 'button';
    b.addEventListener('click', f);
    return b;
  }
  function texto(partes) {
    var s = el('span', 'flex:1 1 280px;min-width:0');
    partes.forEach(function (p) {
      if (typeof p === 'string') s.appendChild(document.createTextNode(p));
      else s.appendChild(el('b', '', p.b));
    });
    return s;
  }
  function recarregar() { location.reload(); }

  function faixaVersao() {
    pronto(function () {
      if (faixaV && faixaV.isConnected) return;
      faixaV = faixa('#fff4d6', '#5c3d00', '#e0b75f');
      faixaV.setAttribute('role', 'alert');
      faixaV.setAttribute('data-aviso-versao', 'versao');
      faixaV.appendChild(texto(CFG.grava && CFG.fila
        ? [{ b: 'O PAINEL FOI ATUALIZADO.' }, ' Esta tela é da versão anterior e ', { b: 'NÃO SALVA MAIS NADA' },
           '. ' + F5 + ' para atualizar: o que já estava guardado neste aparelho sobe sozinho na versão nova (não precisa refazer).']
        : CFG.grava
        ? [{ b: 'O PAINEL FOI ATUALIZADO.' }, ' Esta tela é da versão anterior: o que for feito aqui ',
           { b: 'NÃO SERÁ SALVO' }, '. ' + F5 + ' para atualizar e continue na versão nova.']
        : [{ b: 'O PAINEL FOI ATUALIZADO.' }, ' ' + F5 + ' para ver a versão nova.']));
      faixaV.appendChild(botao(CELULAR ? 'ATUALIZAR AGORA' : 'ATUALIZAR AGORA (F5)', '#7a5200', '#ffffff', recarregar));
      caixaTopo().insertBefore(faixaV, caixaTopo().firstChild);
    });
  }
  function versaoNova() {
    if (NOVA) return;
    NOVA = true;
    faixaVersao();
    avisarWindows();
  }
  function naoSalvou(motivo) {
    falhas++;
    pronto(function () {
      if (!faixaE || !faixaE.isConnected) {
        faixaE = faixa('#fde8e8', '#7a1212', '#e6a3a3');
        faixaE.setAttribute('role', 'alert');
        faixaE.setAttribute('data-aviso-versao', 'erro');
        caixaTopo().appendChild(faixaE);
      }
      faixaE.textContent = '';
      faixaE.appendChild(texto([{ b: 'NÃO SALVOU.' }, ' A última gravação não foi confirmada: ' + motivo +
        (falhas > 1 ? ' (' + falhas + ' vezes)' : '') + '. ' + (CELULAR ? 'Toque em ATUALIZAR' : 'Aperte F5') +
        ' para recarregar e faça de novo o que não ficou salvo.']));
      faixaE.appendChild(botao(CELULAR ? 'ATUALIZAR' : 'ATUALIZAR (F5)', '#9b1c1c', '#ffffff', recarregar));
      faixaE.appendChild(botao('FECHAR', 'transparent', '#7a1212', function () { faixaE.remove(); falhas = 0; }));
    });
  }

  // ---------- aviso no Windows ----------
  function temWindows() {
    if (!('Notification' in window)) return false;
    // celular: o Chrome do Android nao deixa criar notificacao pela pagina
    return !CELULAR;
  }
  function sinoLigado() {
    return temWindows() && Notification.permission === 'granted' && ler(CH_SINO) === '1';
  }
  function icone() {
    var l = document.querySelector('link[rel~="icon"]');
    return l && l.href ? l.href : undefined;
  }
  function notificar(titulo, corpo) {
    try {
      var n = new Notification(titulo, { body: corpo, tag: 'versao-' + CFG.painel, icon: icone() });
      n.onclick = function () { try { window.focus(); } catch (e) { /* janela ja na frente */ } n.close(); };
    } catch (e) { /* navegador recusou: fica a faixa na tela */ }
  }
  function avisarWindows() {
    if (!sinoLigado()) return;
    notificar(nome() + ' FOI ATUALIZADO', CFG.grava && CFG.fila
      ? 'Aperte F5 no painel antes de continuar: a versão anterior não salva mais nada (o que ficou guardado sobe depois de atualizar).'
      : CFG.grava
      ? 'Aperte F5 no painel antes de continuar: o que for feito na versão anterior NÃO será salvo.'
      : 'Aperte F5 no painel para ver a versão nova.');
  }
  function mostrarRecado(t) {
    if (!recado || !recado.isConnected) {
      recado = el('div', 'position:fixed;left:54px;bottom:14px;z-index:' + Z + ';max-width:min(420px,calc(100% - 70px));' +
                         'background:#1f2d2c;color:#ffffff;border-radius:8px;padding:8px 12px;font:13px/1.35 ' + FONTE +
                         ';box-shadow:0 6px 18px rgba(0,0,0,.25)');
      document.body.appendChild(recado);
      if (sino) recado.style.bottom = (parseInt(sino.style.bottom, 10) + 2 || 14) + 'px';
    }
    recado.textContent = t;
    clearTimeout(recadoT);
    recadoT = setTimeout(function () { if (recado) recado.remove(); }, 7000);
  }
  function desenharSino() {
    if (!sino) return;
    var on = sinoLigado();
    sino.style.background = on ? '#D5DF66' : '#ffffff';
    sino.style.color = on ? '#1B302F' : '#3d4f4e';
    sino.style.borderColor = on ? '#9fae2c' : '#9aa9a8';
    sino.style.opacity = on ? '1' : '.8';
    sino.setAttribute('aria-pressed', on ? 'true' : 'false');
    sino.title = on
      ? 'Aviso no Windows LIGADO: quando ' + nome() + ' for atualizado, o Windows avisa. Clique para desligar.'
      : 'Avisar no Windows quando ' + nome() + ' for atualizado (clique para ligar)';
  }
  function clicarSino() {
    if (sinoLigado()) {
      guardar(CH_SINO, '0');
      desenharSino();
      mostrarRecado('Aviso no Windows DESLIGADO para ' + nome() + '.');
      return;
    }
    if (Notification.permission === 'denied') {
      mostrarRecado('O navegador está bloqueando os avisos deste painel. Libere no cadeado ao lado do endereço: Notificações > Permitir. Depois clique no sino de novo.');
      return;
    }
    var feito = false;
    function depois(p) {
      if (feito) return;          // requestPermission chama o callback E resolve a promise
      feito = true;
      if (p === 'granted') {
        guardar(CH_SINO, '1');
        mostrarRecado('Aviso no Windows LIGADO: quando ' + nome() + ' for atualizado, o Windows avisa (com o painel aberto).');
        notificar('Pronto: aviso de ' + nome() + ' ligado', 'Quando este painel for atualizado, o Windows avisa aqui.');
      } else {
        mostrarRecado('Sem permissão do navegador: o aviso no Windows continua desligado.');
      }
      desenharSino();
    }
    try {
      var p = Notification.requestPermission(depois);
      if (p && p.then) p.then(depois);
    } catch (e) {
      Notification.requestPermission().then(depois);
    }
  }
  function montarSino() {
    if (!CFG.sino || !temWindows() || sino) return;
    sino = el('button', 'position:fixed;left:12px;bottom:12px;z-index:' + (Z - 1) + ';width:34px;height:34px;border-radius:50%;' +
                        'border:1px solid #9aa9a8;display:flex;align-items:center;justify-content:center;padding:0;cursor:pointer;' +
                        'box-shadow:0 2px 8px rgba(0,0,0,.18)');
    sino.type = 'button';
    sino.setAttribute('data-aviso-versao', 'sino');
    sino.setAttribute('aria-label', 'Aviso no Windows quando o painel for atualizado');
    var NS = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('width', '18');
    svg.setAttribute('height', '18');
    svg.setAttribute('fill', 'none');
    svg.setAttribute('stroke', 'currentColor');
    svg.setAttribute('stroke-width', '2');
    svg.setAttribute('stroke-linecap', 'round');
    svg.setAttribute('stroke-linejoin', 'round');
    ['M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9', 'M13.73 21a2 2 0 0 1-3.46 0'].forEach(function (d) {
      var p = document.createElementNS(NS, 'path');
      p.setAttribute('d', d);
      svg.appendChild(p);
    });
    sino.appendChild(svg);
    sino.addEventListener('click', clicarSino);
    document.body.appendChild(sino);
    desenharSino();
    posicionarSino();
    window.addEventListener('resize', posicionarSino);
    setInterval(posicionarSino, 2000);
  }
  function fixoEmbaixo(x, y) {
    // rodape/barra fixa do proprio painel naquele canto: o sino sobe para cima dela
    var nos = document.elementsFromPoint ? document.elementsFromPoint(x, y) : [];
    for (var i = 0; i < nos.length; i++) {
      for (var n = nos[i]; n && n !== document.body && n !== document.documentElement; n = n.parentElement) {
        if (n === sino || n === recado || n === caixa) break;
        var p = getComputedStyle(n).position;
        if (p === 'fixed' || p === 'sticky') return n.getBoundingClientRect().top;
      }
    }
    return null;
  }
  function posicionarSino() {
    // pagina reescrita por document.write (FECHAMENTO, PLANO): a janela e o fetch envolvido continuam, o sino volta
    if (sino && !sino.isConnected && document.body) document.body.appendChild(sino);
    if (!sino || !sino.isConnected) return;
    var h = window.innerHeight, topo = null;
    [h - 6, h - 29, h - 46].forEach(function (y) {
      var t = fixoEmbaixo(29, y);
      if (t !== null && t > h / 2 && (topo === null || t < topo)) topo = t;
    });
    var baixo = topo === null ? 12 : Math.round(h - topo + 10);
    if (sino.style.bottom !== baixo + 'px') sino.style.bottom = baixo + 'px';
    if (recado && recado.isConnected) recado.style.bottom = (baixo + 2) + 'px';
  }

  // ---------- partida ----------
  window.AvisoVersao = {
    conferir: function () { return conferir(true); },
    nova: function () { return NOVA; },
    ehGravacao: ehGravacao,
    sinoLigado: sinoLigado,
    config: CFG
  };
  pronto(montarSino);
  function comecar() { conferir(true); }
  if (document.readyState === 'complete') setTimeout(comecar, 0);
  else window.addEventListener('load', comecar);
  setInterval(function () { conferir(Date.now() - confTudoEm > 600000); }, 120000);
  // a faixa de versao nova nao some se a pagina for reescrita (document.write) depois do aviso
  setInterval(function () { if (NOVA && faixaV && !faixaV.isConnected && document.body) faixaVersao(); }, 2000);
  document.addEventListener('visibilitychange', function () {
    if (!document.hidden && Date.now() - confEm > 60000) conferir(false);
  });
})();
