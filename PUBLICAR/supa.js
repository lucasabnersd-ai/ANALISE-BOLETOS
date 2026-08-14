/* Camada de login do painel ANALISE BOLETOS.
 *
 * O index.html publicado nasce SEM dados -- o Pages serve o site publicamente
 * mesmo com o repositorio privado. A carteira mora no Postgres, atras de RLS,
 * e so desce depois do login.
 *
 * Regras herdadas dos outros paineis:
 *  - persistSession:false. A sessao vive so na memoria da aba; em maquina
 *    compartilhada, a proxima pessoa nao entra ja logada.
 *  - link de recuperacao NAO vale como sessao: quem tem acesso ao e-mail
 *    entraria sem saber a senha.
 *  - nada de polling curto em getSession() -- o supabase-js serializa a
 *    chamada num lock e a aba congela.
 */
(function () {
  "use strict";

  var SUPABASE_URL = "https://pyrniqluywejmgzqkari.supabase.co";
  var SUPABASE_KEY = "sb_publishable_fXWQGDirOvs5xfxZDaSOtg_Jgd7vcbu";
  var LIB = "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.js";

  var sb = null;
  var resolver, rejeitar;
  var entregue = false;   // a carteira ja foi entregue ao painel uma vez?
  // O painel espera esta promessa; ela so resolve depois do login.
  window.__DADOS_BOLETOS__ = new Promise(function (res, rej) { resolver = res; rejeitar = rej; });

  /* ---------------- tela de login ---------------- */
  function el(tag, css, html) {
    var e = document.createElement(tag);
    if (css) e.style.cssText = css;
    if (html != null) e.innerHTML = html;
    return e;
  }

  function overlay() {
    var pronto = document.getElementById("ab-login");
    if (pronto) return pronto;

    var fundo = el("div", "position:fixed;inset:0;z-index:9999;display:flex;align-items:center;" +
      "justify-content:center;background:#0c1320;padding:20px;" +
      "font:14px/1.45 'Segoe UI',Calibri,system-ui,sans-serif");
    fundo.id = "ab-login";

    var caixa = el("div", "width:min(380px,100%);background:#111a2b;color:#e8eefb;" +
      "border:1px solid #22314b;border-radius:12px;padding:22px;box-shadow:0 18px 50px rgba(0,0,0,.45)");
    caixa.innerHTML =
      '<h1 style="margin:0 0 4px;font-size:17px;font-weight:700">Análise de Boletos</h1>' +
      '<p style="margin:0 0 18px;font-size:12.5px;color:#a3b3ca">Entre para ver a carteira.</p>' +
      '<label style="display:block;font-size:11px;color:#a3b3ca;margin-bottom:4px">E-MAIL</label>' +
      '<input id="ab-email" type="email" autocomplete="username" style="width:100%;box-sizing:border-box;' +
      'font:inherit;padding:9px 11px;margin-bottom:12px;background:#0e1626;color:#e8eefb;' +
      'border:1px solid #22314b;border-radius:7px;outline:none">' +
      '<label style="display:block;font-size:11px;color:#a3b3ca;margin-bottom:4px">SENHA</label>' +
      '<input id="ab-senha" type="password" autocomplete="current-password" style="width:100%;box-sizing:border-box;' +
      'font:inherit;padding:9px 11px;margin-bottom:16px;background:#0e1626;color:#e8eefb;' +
      'border:1px solid #22314b;border-radius:7px;outline:none">' +
      '<button id="ab-entrar" type="button" style="width:100%;font:inherit;font-weight:700;cursor:pointer;' +
      'padding:10px;background:#2f6ba8;color:#fff;border:0;border-radius:7px">Entrar</button>' +
      '<p id="ab-msg" style="margin:12px 0 0;font-size:12px;min-height:17px"></p>';

    fundo.appendChild(caixa);
    document.body.appendChild(fundo);

    var msg = caixa.querySelector("#ab-msg");
    fundo._setMsg = function (t, ok) {
      msg.textContent = t || "";
      msg.style.color = ok ? "#4ade80" : "#fca5a5";
    };
    caixa.querySelector("#ab-entrar").addEventListener("click", entrar);
    caixa.querySelector("#ab-senha").addEventListener("keydown", function (e) {
      if (e.key === "Enter") entrar();
    });
    return fundo;
  }

  function mostrarLogin(t, ok) {
    var o = overlay();
    o.style.display = "flex";
    if (t) o._setMsg(t, ok);
  }
  function esconderLogin() {
    var o = document.getElementById("ab-login");
    if (o) o.style.display = "none";
  }

  /* ---------------- carregar a lib ---------------- */
  function carregarLib() {
    return new Promise(function (res, rej) {
      if (window.supabase && window.supabase.createClient) return res();
      var s = document.createElement("script");
      s.src = LIB;
      s.onload = res;
      s.onerror = function () { rej(new Error("nao consegui carregar a biblioteca do Supabase")); };
      document.head.appendChild(s);
    });
  }

  /* Um link type=recovery gera sessao valida. Sem tela de troca de senha,
     aceitar essa sessao deixaria entrar so com acesso ao e-mail. */
  function ehRecuperacao(sessao) {
    try {
      if (!sessao || !sessao.access_token) return false;
      var corpo = JSON.parse(atob(sessao.access_token.split(".")[1]));
      return corpo && (corpo.aal === "aal0" || corpo.amr || []).some
        ? (corpo.amr || []).some(function (m) { return m.method === "recovery"; })
        : false;
    } catch (e) { return false; }
  }

  /* ---------------- dados ---------------- */

  /* Le a tabela inteira em blocos.
   *
   * ⚠ NAO troque isto por .limit(): o PostgREST do Supabase tem um TETO PROPRIO
   * de linhas por resposta (1000). O .limit(20000) que existia aqui nao vence
   * esse teto -- ele so pede menos, nunca mais. O resultado era a carteira
   * chegando cortada SEM ERRO NENHUM: em 12/08/2026 a carteira passou de 1000
   * ativos ao ganhar a aba da classificacao e 22 titulos sumiram da tela em
   * silencio (os uuid "SF1:" ordenam por ultimo, entao o corte caiu todo neles).
   * Paginar resolve para qualquer tamanho, hoje e depois.
   *
   * A ordem por uuid nao e enfeite: sem ordem estavel, dois blocos podem repetir
   * ou pular linha.
   */
  var BLOCO = 1000;
  var MAX_BLOCOS = 200;             // trava contra laco infinito (200k linhas)

  async function lerTudo(tabela, colunas, filtrar) {
    var tudo = [];
    for (var pagina = 0; pagina < MAX_BLOCOS; pagina++) {
      var q = sb.from(tabela).select(colunas);
      if (filtrar) q = filtrar(q);
      var r = await q.order("uuid", { ascending: true })
                     .range(pagina * BLOCO, pagina * BLOCO + BLOCO - 1);
      if (r.error) throw r.error;
      var lote = r.data || [];
      tudo = tudo.concat(lote);
      // bloco incompleto = acabou. (O teto do servidor nunca devolve mais que
      // BLOCO, entao um bloco cheio significa "pode haver mais".)
      if (lote.length < BLOCO) return tudo;
    }
    throw new Error("carteira grande demais para carregar (" + tudo.length + " linhas)");
  }

  /* ---------------- carregamento SOB DEMANDA (13/08/2026) ----------------
   * Antes, entrar no painel baixava a carteira INTEIRA -- 3.892 linhas, 5,2 MB
   * -- para mostrar uma aba de 97. Com a aba de pedidos a conta passou a doer.
   *
   * Agora sao duas etapas: o login traz so os CABECALHOS (uma linha por aba,
   * uuid "#meta:<aba>") e as MARCACOES; as linhas de cada aba chegam quando
   * alguem abre a aba, e ficam em memoria para nao rebaixar.
   *
   * ⚠ As marcacoes vem INTEIRAS, e nao por aba, de proposito: elas nao dizem a
   * que aba pertencem (a tabela e uma so, por uuid) e sao a fonte de quem esta
   * tratado -- que o painel precisa saber para CONTAR as abas que ainda nem
   * abriu. Sao poucas linhas perto da carteira.
   */
  var marcasPorUuid = {};

  function montarLinha(row) {
    var d = row.dados || {};
    var marca = marcasPorUuid[row.uuid] || {};
    return {
      uuid: row.uuid,
      c: d.c || {}, copia: d.copia || {}, difere: d.difere || {},
      delta: d.delta || {}, forte: !!d.forte,
      alerta: d.alerta || { tipo: "", nf: "" },
      match: d.match || null,
      texto: d.texto || null,
      selo: d.selo || "", ocr: !!d.ocr,
      feito: marca.feito != null ? marca.feito : !!d.feito,
      tratativa: marca.tratativa || "",
      // tipo de pagamento REAL, informado a mao na aba de associados. So
      // existe na marcacao -- a carteira nunca traz este campo.
      tipo_pgto_real: marca.tipo_pgto_real || "",
      tratado_em: marca.tratado_em || "",
      tratado_por: marca.tratado_por || "",
      quem: marca.quem || "",
      // sem check, a data da marcacao e o que diz QUANDO a NF foi tratada --
      // o painel usa isso para decidir visao futura x visao passado
      atualizado_em: marca.atualizado_em || "",
      // `o` so traz o que difere do exibido (numero cru, data ISO); o resto
      // cai no proprio `c`. Ordenar pelo texto formatado daria ordem errada.
      o: Object.assign({}, d.c || {}, d.o || {}),
      busca: Object.keys(d.c || {}).map(function (k) { return d.c[k]; }).join(" ").toLowerCase(),
    };
  }

  /* As linhas de UMA aba. Chamada pelo painel quando a aba e aberta.
   * ⚠ O filtro e por `dados->>aba`, o mesmo campo que o gerador grava em cada
   * linha -- e o `not like '#meta:%'` tira o cabecalho, que tambem carrega o
   * `aba` e viraria uma linha fantasma na tabela. */
  window.__CARREGAR_LINHAS_ABA__ = async function (id) {
    var linhas = await lerTudo("analise_boletos_titulos", "uuid,dados",
      function (q) {
        return q.eq("ativo", true).eq("dados->>aba", id)
                .not("uuid", "like", "#meta:%");
      });
    return linhas.map(montarLinha);
  };

  async function baixarCarteira() {
    // ETAPA 1: so os cabecalhos. Sao 8 linhas -- uma por aba.
    var metaRows = await lerTudo("analise_boletos_titulos", "uuid,dados",
      function (q) { return q.eq("ativo", true).like("uuid", "#meta:%"); });
    if (!metaRows.length) throw new Error("a carteira está vazia no servidor");

    var m = { data: await lerTudo("analise_boletos_marcacoes",
      "uuid,feito,tratativa,tratado_em,tratado_por,quem,atualizado_em,tipo_pgto_real") };
    marcasPorUuid = {};
    (m.data || []).forEach(function (r) { marcasPorUuid[r.uuid] = r; });

    var metas = {}, porAba = {};
    metaRows.forEach(function (row) {
      var d = row.dados || {};
      var aba = d.aba || "associacoes";
      if (d._meta) { metas[aba] = d._meta; porAba[aba] = null; }
    });
    // "Gerado em" e afins valem para a carteira inteira: vem do cabecalho mais
    // novo que existir, nunca de um que ficou para tras.
    var geral = metas[Object.keys(metas)[0]] || null;
    var abas = Object.keys(porAba).map(function (id) {
      var m = metas[id] || {};
      if (m.compacta && m.compacta.length) geral = geral || m;
      return {
        id: id, nome: m.nome || id, cor: m.cor || "azul",
        guia: m.guia || m.nome || id,
        ordem: m.ordem == null ? 99 : m.ordem,
        compacta: m.compacta || [], parcelas: m.parcelas || {},
        pills: m.pills || null, ocr: m.ocr || 0,
        particoes: m.particoes || null,
        // veredito da SE2 (titulos que ja tem boleto lancado ou ja foram
        // baixados). Vem no cabecalho da aba, nao na linha -- a carga nao
        // regrava titulo ja tratado. null = nao houve conferencia.
        se2: m.se2 || null,
        // rotulos dos contadores do resumo; ausente = os rotulos padrao
        resumo: m.resumo || null,
        com_alerta: m.com_alerta || 0, divergentes: m.divergentes || 0,
        // ⚠ `linhas` nasce como lista VAZIA, e nao `null`: assim todo o painel
        // (resumo, sub-abas, pills, drill) continua funcionando sem nenhuma
        // guarda espalhada. Quem diz se ja chegou e o `carregada` -- e e so ele
        // que o desenho consulta para mostrar "carregando" em vez de "vazio".
        linhas: [], carregada: false,
        total: m.total || 0,
        // aba desligada no gerador: o cabecalho continua aqui so para APAGAR do
        // painel a que ja estava no banco -- ver publicar_dados.cabecalho_da_aba
        oculta: !!m.oculta,
      };
    }).filter(function (a) { return a.compacta.length && !a.oculta; });
    if (!geral) throw new Error("carteira sem cabeçalho (_meta)");

    // os titulos vem ordenados por uuid; a ordem das abas e a do gerador
    abas.sort(function (a, b) { return a.ordem - b.ordem; });

    return {
      gerado_em: geral.gerado_em, atualizado_em: geral.atualizado_em,
      salva_em: geral.salva_em || "", sem_codigo: geral.sem_codigo,
      abas: abas,
      // as marcacoes inteiras: e delas que o painel monta quem esta tratado,
      // inclusive das abas que ainda nao foram abertas
      marcacoes: m.data || [],
    };
  }

  /* Check e tratativa vao para o banco -- uma linha por titulo, para duas
     pessoas mexerem sem uma apagar a marcacao da outra.
     Devolve {ok, erro, relogar} em vez de um booleano: quem chama e a fila do
     painel, que precisa saber POR QUE falhou para decidir entre tentar de novo
     e pedir login. O `.select()` no fim e de proposito -- so contamos como
     gravado o que o Postgres devolveu; sem ele, "sem erro" nao prova escrita. */
  window.__SALVAR_MARCACAO__ = async function (uuid, campos) {
    if (!sb) return { ok: false, erro: "ainda sem conexão com o servidor" };
    var linha = Object.assign({ uuid: uuid }, campos);
    var r;
    try {
      r = await sb.from("analise_boletos_marcacoes")
        .upsert(linha, { onConflict: "uuid" }).select("uuid");
    } catch (e) {
      return { ok: false, erro: (e && e.message) || "falha de rede" };
    }
    if (r.error) {
      var m = r.error.message || "";
      console.error("não consegui gravar a marcação:", m);
      // sessao expirada chega como JWT/401; nesse caso insistir nao resolve
      return {
        ok: false, erro: m,
        relogar: /jwt|token|expired|401|not authenticated/i.test(m),
      };
    }
    if (!r.data || !r.data.length) {
      return { ok: false, erro: "o banco não confirmou a gravação" };
    }
    return { ok: true };
  };

  /* O painel chama isto quando a fila detecta sessao expirada: a tela de login
     volta por cima, e o que estava pendente sobe assim que a pessoa entrar. */
  window.__PEDIR_LOGIN__ = function (msg) { mostrarLogin(msg || "", false); };

  async function entrar() {
    var email = (document.getElementById("ab-email").value || "").trim();
    var senha = document.getElementById("ab-senha").value || "";
    if (!email || !senha) return mostrarLogin("Preencha e-mail e senha.");
    mostrarLogin("Entrando…", true);
    var r = await sb.auth.signInWithPassword({ email: email, password: senha });
    if (r.error) return mostrarLogin(motivo(r.error));
    await depoisDoLogin();
  }

  function motivo(e) {
    var m = (e && e.message) || "";
    if (/Invalid login/i.test(m)) return "E-mail ou senha incorretos.";
    if (/Email not confirmed/i.test(m)) return "E-mail ainda não confirmado.";
    if (/rate limit|too many/i.test(m)) return "Muitas tentativas. Aguarde um pouco.";
    return m || "Não consegui entrar.";
  }

  async function depoisDoLogin() {
    var s = await sb.auth.getSession();
    var sessao = s.data && s.data.session;
    if (!sessao) return mostrarLogin("Sessão não iniciada.");
    if (ehRecuperacao(sessao)) {
      await sb.auth.signOut();
      return mostrarLogin("Link de recuperação não serve para entrar. Use sua senha.");
    }
    try {
      var dados = await baixarCarteira();
      esconderLogin();
      // A promessa so resolve uma vez; num segundo login (sessao que expirou no
      // meio do trabalho) e preciso reentregar a carteira na mao, senao a tela
      // fica com os dados velhos e a fila subiria por baixo do pano.
      if (entregue && window.__INICIAR_PAINEL__) window.__INICIAR_PAINEL__(dados);
      else { entregue = true; resolver(dados); }
    } catch (e) {
      // RLS negando aparece aqui: autenticou, mas nao esta autorizado.
      var m = (e && e.message) || "";
      await sb.auth.signOut();
      mostrarLogin(/permission|denied|policy/i.test(m)
        ? "Sua conta não tem acesso a este painel."
        : "Não consegui carregar a carteira: " + m);
    }
  }

  async function iniciar() {
    try {
      await carregarLib();
    } catch (e) {
      return mostrarLogin(e.message);
    }
    sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
      auth: { persistSession: false, autoRefreshToken: true, detectSessionInUrl: false },
    });
    window.__SB_PAINEL__ = sb;

    // Faxina: versoes antigas podiam ter gravado sessao aqui.
    try {
      Object.keys(localStorage).forEach(function (k) {
        if (/^sb-.*-auth-token$/.test(k)) localStorage.removeItem(k);
      });
    } catch (e) {}

    var s = await sb.auth.getSession();
    if (s.data && s.data.session) return depoisDoLogin();
    mostrarLogin();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar);
  } else {
    iniciar();
  }
})();
