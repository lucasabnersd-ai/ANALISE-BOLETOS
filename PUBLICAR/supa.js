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
  async function baixarCarteira() {
    // so os ativos: os que sairam da base ficam no banco, mas fora do painel
    var t = await sb.from("analise_boletos_titulos")
      .select("uuid,dados").eq("ativo", true)
      .order("uuid", { ascending: true }).limit(20000);
    if (t.error) throw t.error;
    if (!t.data.length) throw new Error("a carteira está vazia no servidor");

    var m = await sb.from("analise_boletos_marcacoes")
      .select("uuid,feito,tratativa,tratado_em,tratado_por,quem,atualizado_em").limit(20000);
    if (m.error) throw m.error;
    var porUuid = {};
    (m.data || []).forEach(function (r) { porUuid[r.uuid] = r; });

    // O cabecalho de cada aba (colunas, parcelas, contagens) vem numa linha
    // propria, de uuid "#meta:<aba>". Titulos antigos podem ainda trazer o
    // _meta embutido -- vale so enquanto nao houver a linha propria, senao uma
    // NF que ficou no banco de uma carga velha serviria um cabecalho vencido.
    var metas = {}, metasVelhas = {}, porAba = {};
    t.data.forEach(function (row) {
      var d = row.dados || {};
      var aba = d.aba || "associacoes";
      if (String(row.uuid).indexOf("#meta:") === 0) {
        if (d._meta) { metas[aba] = d._meta; porAba[aba] = porAba[aba] || []; }
        return;                                  // cabecalho nao e titulo
      }
      if (d._meta && !metasVelhas[aba]) metasVelhas[aba] = d._meta;
      var marca = porUuid[row.uuid] || {};
      (porAba[aba] = porAba[aba] || []).push({
        uuid: row.uuid,
        c: d.c || {}, copia: d.copia || {}, difere: d.difere || {},
        delta: d.delta || {}, forte: !!d.forte,
        alerta: d.alerta || { tipo: "", nf: "" },
        match: d.match || null,
        selo: d.selo || "", ocr: !!d.ocr,
        feito: marca.feito != null ? marca.feito : !!d.feito,
        tratativa: marca.tratativa || "",
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
      });
    });
    // "Gerado em" e afins valem para a carteira inteira: vem do cabecalho mais
    // novo que existir, nunca de um que ficou para tras.
    var geral = metas[Object.keys(metas)[0]] || null;
    var abas = Object.keys(porAba).map(function (id) {
      var m = metas[id] || metasVelhas[id] || {};
      if (m.compacta && m.compacta.length) geral = geral || m;
      return {
        id: id, nome: m.nome || id, cor: m.cor || "azul",
        guia: m.guia || m.nome || id,
        ordem: m.ordem == null ? 99 : m.ordem,
        compacta: m.compacta || [], parcelas: m.parcelas || {},
        pills: m.pills || null, ocr: m.ocr || 0,
        particoes: m.particoes || null,
        com_alerta: m.com_alerta || 0, divergentes: m.divergentes || 0,
        linhas: porAba[id],
      };
    }).filter(function (a) { return a.compacta.length; });
    if (!geral) throw new Error("carteira sem cabeçalho (_meta)");

    // os titulos vem ordenados por uuid; a ordem das abas e a do gerador
    abas.sort(function (a, b) { return a.ordem - b.ordem; });

    // carregado_em e a hora DESTA tela: diz se o que esta na frente da pessoa
    // acabou de vir do banco ou esta aberto ha horas.
    var agora = new Date();
    return {
      gerado_em: geral.gerado_em, atualizado_em: geral.atualizado_em,
      salva_em: geral.salva_em || "", sem_codigo: geral.sem_codigo,
      carregado_em: agora.toLocaleString("pt-BR", {day:"2-digit", month:"2-digit",
        year:"numeric", hour:"2-digit", minute:"2-digit"}).replace(", ", " às "),
      abas: abas,
    };
  }

  /* Check e tratativa vao para o banco -- uma linha por titulo, para duas
     pessoas mexerem sem uma apagar a marcacao da outra. */
  window.__SALVAR_MARCACAO__ = async function (uuid, campos) {
    if (!sb) return false;
    var linha = Object.assign({ uuid: uuid }, campos);
    var r = await sb.from("analise_boletos_marcacoes").upsert(linha, { onConflict: "uuid" });
    if (r.error) { console.error("não consegui gravar a marcação:", r.error.message); return false; }
    return true;
  };

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
      resolver(dados);
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
