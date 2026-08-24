// Carga da carteira do painel ANALISE BOLETOS (e o backup completo).
//
// verify_jwt fica DESLIGADO de proposito: quem chama e um script local, nao um
// usuario logado. A autenticacao e propria -- token no header x-carga-token,
// conferido pelo RPC ab_carga_token_ok, que so a service_role pode executar.
//
// Dois caminhos:
//   POST {"titulos":[...]}   -> grava a carteira
//   POST {"acao":"backup"}  -> devolve TUDO (titulos inclusive arquivados,
//                               marcacoes, historico e autorizados)
//
// QUATRO REGRAS DO USUARIO, todas sobre nao perder historico:
//
// 1. NADA E APAGADO. Titulo que nao veio na carga vira ativo=false com saiu_em
//    preenchido: some do painel mas continua no banco, com a marcacao e a
//    tratativa dele.
//
// 2. NF NAO ASSOCIADA (uuid "BOL:<linha digitavel>") nem isso: fica no painel
//    mesmo quando some da planilha. So sai quando passa a ter titulo associado,
//    e ai chega na carga por uma das outras abas -- com a mesma linha digitavel.
//
// 3. TITULO JA TRATADO NAO E ATUALIZADO. Se a pessoa ja marcou o check, o que
//    vale e o que esta no painel: a carga nao regrava os dados dele. Ele fica
//    onde esta, na aba Tratados, com os numeros que ela viu quando tratou.
//
// 4. DISPENSADO MANUALMENTE NAO VOLTA (dispensado_em preenchido). E o oposto da
//    regra 2: quem foi tirado do painel de proposito -- a limpeza da visao
//    passado -- fica fora, mesmo continuando na planilha de origem. Sem isto o
//    UPDATE ativo=false duraria ate a carga seguinte, porque o upsert regrava
//    ativo=true para tudo que vier na planilha. Para trazer de volta:
//    update analise_boletos_titulos set dispensado_em=null, ativo=true where ...
//
// ⚠ TETO DE 1.000 LINHAS DO POSTGREST (corrigido 24/08/2026). As leituras de
// existentes/tratados NAO podem usar `.limit(N)` grande: o PostgREST corta a
// resposta em max-rows (1.000 por padrao) e o `.limit()` NAO vence esse teto.
// Com a tabela acima de 1.000 linhas, a funcao so "via" as primeiras 1.000 e
// deixava de arquivar (regra 1) tudo que estivesse alem -- notas que sairam da
// carga continuavam ativas para sempre. A leitura agora e PAGINADA com `.range()`
// e `.order("uuid")` (ordem estavel para nao pular nem repetir pagina). Ver
// `lerTudo`.
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "jsr:@supabase/supabase-js@2";

const sb = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  { auth: { persistSession: false } },
);

// Abas de NF nao associada: o que esta nelas nunca e arquivado por ausencia.
const NAO_ASSOCIADAS = new Set(["futuros", "passado"]);
const PREFIXO_BOLETO = "BOL:";
// Linha de cabecalho da aba (colunas, contagens) -- nao e titulo de ninguem.
const PREFIXO_META = "#meta:";
const LOTE = 200;
// Tamanho da pagina ao LER a tabela inteira. Fica <= ao max-rows do PostgREST:
// uma pagina que volta com menos que isto e sinal de que acabou (ver lerTudo).
const PAGINA = 1000;

const json = (corpo: unknown, status = 200) =>
  new Response(JSON.stringify(corpo), {
    status,
    headers: { "Content-Type": "application/json" },
  });

const nega = () => json({ erro: "nao autorizado" }, 401);

// Le uma tabela/consulta INTEIRA, vencendo o teto de 1.000 linhas do PostgREST.
// `consulta(de, ate)` deve devolver a query ja com o filtro e a ORDEM estavel,
// faltando so o intervalo -- que este helper aplica com `.range()`. Para quando
// uma pagina volta com menos linhas que o tamanho pedido.
async function lerTudo<T>(
  consulta: (de: number, ate: number) => PromiseLike<{ data: T[] | null; error: unknown }>,
): Promise<T[]> {
  const tudo: T[] = [];
  for (let de = 0; ; de += PAGINA) {
    const { data, error } = await consulta(de, de + PAGINA - 1);
    if (error) throw error;
    const lote = data ?? [];
    tudo.push(...lote);
    if (lote.length < PAGINA) break;
  }
  return tudo;
}

Deno.serve(async (req: Request) => {
  if (req.method !== "POST") return new Response("metodo nao permitido", { status: 405 });

  const enviado = (req.headers.get("x-carga-token") ?? "").trim();
  if (!enviado || enviado.length > 200) return nega();

  const { data: ok, error: erroToken } = await sb.rpc("ab_carga_token_ok", { p_token: enviado });
  if (erroToken || ok !== true) return nega();

  let carga: { titulos?: Array<{ uuid: string; dados: unknown }>; acao?: string };
  try {
    carga = await req.json();
  } catch {
    return json({ erro: "json invalido" }, 400);
  }

  // ---------------------------------------------------------------- backup
  if (carga.acao === "backup") {
    const { data, error } = await sb.rpc("ab_backup_completo");
    if (error) return json({ erro: error.message }, 500);
    return json(data);
  }

  // ---------------------------------------------------------------- carga
  const recebidos = carga.titulos ?? [];
  if (!Array.isArray(recebidos) || recebidos.length === 0 || recebidos.length > 20000) {
    return json({ erro: "lista de titulos invalida" }, 400);
  }
  for (const t of recebidos) {
    if (!t || typeof t.uuid !== "string" || !t.uuid || t.uuid.length > 160 || t.dados == null) {
      return json({ erro: "titulo invalido" }, 400);
    }
  }

  // Titulo repetido dentro da propria carga: vale a PRIMEIRA aparicao. Deixar o
  // upsert resolver faria a ultima vencer, em silencio.
  const titulos: Array<{ uuid: string; dados: unknown }> = [];
  const vistos = new Set<string>();
  let repetidos = 0;
  for (const t of recebidos) {
    if (vistos.has(t.uuid)) { repetidos++; continue; }
    vistos.add(t.uuid);
    titulos.push(t);
  }

  // Quem ja foi tratado: a carga nao mexe nos dados dele (regra 3).
  // ⚠ PAGINADO (teto de 1.000): antes `.limit(20000)` so trazia as primeiras
  // 1.000 marcacoes -- um titulo tratado alem disso seria REGRAVADO pela carga.
  const tratados = new Set<string>();
  try {
    const linhas = await lerTudo<{ uuid: string }>((de, ate) =>
      sb.from("analise_boletos_marcacoes").select("uuid").eq("feito", true)
        .order("uuid").range(de, ate)
    );
    linhas.forEach((r) => tratados.add(r.uuid));
  } catch (e) {
    return json({ erro: String((e as { message?: string })?.message ?? e) }, 500);
  }

  // Quem ja existe no banco, se esta no painel e se foi dispensado a mao.
  // ⚠ PAGINADO (teto de 1.000): esta era a leitura que quebrava o arquivamento
  // -- com a tabela acima de 1.000 linhas, tudo alem da primeira pagina ficava
  // invisivel para a regra 1 e nunca saia do painel.
  const existentes = new Map<string, boolean>();
  const dispensados = new Set<string>();
  try {
    const linhas = await lerTudo<{ uuid: string; ativo: boolean; dispensado_em: string | null }>(
      (de, ate) =>
        sb.from("analise_boletos_titulos").select("uuid,ativo,dispensado_em")
          .order("uuid").range(de, ate)
    );
    linhas.forEach((r) => {
      existentes.set(r.uuid, r.ativo);
      if (r.dispensado_em) dispensados.add(r.uuid);
    });
  } catch (e) {
    return json({ erro: String((e as { message?: string })?.message ?? e) }, 500);
  }

  const agora = new Date().toISOString();
  const gravar: Array<Record<string, unknown>> = [];
  const reativar: string[] = [];
  let intocados = 0;
  let ignorados_dispensados = 0;

  for (const t of titulos) {
    const eMeta = t.uuid.startsWith(PREFIXO_META);
    // Dispensado a mao (regra 4): a carga passa por cima sem tocar. Nao entra no
    // upsert -- que gravaria ativo=true -- nem na lista de reativar.
    if (!eMeta && dispensados.has(t.uuid)) {
      ignorados_dispensados++;
      continue;
    }
    // O cabecalho da aba SEMPRE e regravado: ele descreve as colunas do painel,
    // nao o titulo de ninguem, e um cabecalho velho monta a tabela errada.
    const preservar = !eMeta && tratados.has(t.uuid) && existentes.has(t.uuid);
    if (preservar) {
      intocados++;
      if (existentes.get(t.uuid) === false) reativar.push(t.uuid);
      continue;
    }
    gravar.push({
      uuid: t.uuid,
      dados: t.dados,
      atualizado_em: agora,
      ativo: true,
      saiu_em: null,
    });
  }

  // upsert em lotes: 958+ linhas de uma vez estoura o limite do PostgREST
  for (let i = 0; i < gravar.length; i += LOTE) {
    const { error } = await sb
      .from("analise_boletos_titulos")
      .upsert(gravar.slice(i, i + LOTE), { onConflict: "uuid" });
    if (error) return json({ erro: error.message }, 500);
  }

  // Tratado que estava arquivado e voltou a aparecer: volta para o painel, mas
  // com os dados que ja tinha -- so o ativo muda.
  for (let i = 0; i < reativar.length; i += LOTE) {
    await sb.from("analise_boletos_titulos")
      .update({ ativo: true, saiu_em: null })
      .in("uuid", reativar.slice(i, i + LOTE));
  }

  // Boletos que agora TEM titulo associado: sao os que vieram nesta carga por
  // uma aba que nao e de nao associadas. A chave e a linha digitavel, que e o
  // que forma o uuid "BOL:..." do lado nao associado. Cabecalho de aba nao
  // conta -- ele nao representa titulo nenhum.
  const associadas = new Set<string>();
  for (const t of titulos) {
    if (t.uuid.startsWith(PREFIXO_META)) continue;
    // deno-lint-ignore no-explicit-any
    const d = t.dados as any;
    if (!d || NAO_ASSOCIADAS.has(d.aba)) continue;
    const digitos = String(d?.copia?.linha_digitavel ?? "").replace(/\D/g, "");
    if (digitos) associadas.add(PREFIXO_BOLETO + digitos);
  }

  // Quem nao veio nesta carga sai do painel, mas CONTINUA no banco.
  const sumidos: string[] = [];
  for (const [uuid, ativo] of existentes) {
    if (ativo && !vistos.has(uuid)) sumidos.push(uuid);
  }

  // A NF nao associada so sai do painel quando aparece associada.
  const arquivar = sumidos.filter((u) =>
    !u.startsWith(PREFIXO_BOLETO) || associadas.has(u)
  );
  const preservados = sumidos.length - arquivar.length;

  let arquivados = 0;
  for (let i = 0; i < arquivar.length; i += LOTE) {
    const { error, count } = await sb
      .from("analise_boletos_titulos")
      .update({ ativo: false, saiu_em: agora }, { count: "exact" })
      .in("uuid", arquivar.slice(i, i + LOTE));
    if (!error) arquivados += count ?? 0;
  }

  return json({
    ok: true,
    recebidos: recebidos.length,
    repetidos_na_carga: repetidos,
    gravados: gravar.length,
    intocados_por_ja_tratados: intocados,
    ignorados_por_dispensa_manual: ignorados_dispensados,
    reativados: reativar.length,
    arquivados,
    preservados_nao_associados: preservados,
    apagados: 0,
  });
});
