# -*- coding: utf-8 -*-
"""
NUCLEO JS DO GERADOR DE PLANILHA (.xlsx) - fonte unica.

Quem usa:
  - aplicar_exportar_excel.py  -> painel SE2 (PUBLICAR_GITHUB/index.html)
  - aplicar_excel_fluxo.py     -> painel Fluxo de Caixa (fluxo_template.html, que gera
                                  PUBLICAR_GITHUB/fluxo.html e PUBLICAR_FLUXO/index.html)

Por que um modulo Python e nao um .js compartilhado: cada painel e um HTML unico e
autocontido, publicado em repositorio diferente. Um arquivo .js a mais exigiria mexer
nos dois deploys e no publicar_seguro. Assim o codigo tem UMA fonte aqui e cada HTML
sai autocontido, sem nenhuma dependencia externa (nada de CDN).

O que ele gera: ZIP (deflate nativo do navegador) + XML do OpenXML. Cabecalho escuro
congelado, autofiltro, largura por coluna, numero como numero (#,##0.00), data como
data (dd/mm/yyyy) e texto em CAIXA ALTA - o mesmo visual nos dois paineis.

API exposta em window.PlanilhaSE2:
  gerar(aba, colunas, linhas)  -> Blob   (uma aba, cabecalho + autofiltro; usado pelo SE2)
  gerarLivro(abas)             -> Blob   (varias abas; usado pelo Fluxo)
  baixar(blob, nome)                     (dispara o download)
  lerCsv(texto) -> matriz

Formato de cada aba em gerarLivro:
  {
    aba: 'FLUXO',                 // nome da guia
    colunas: [{titulo, tipo, largura}],   // tipo: texto | moeda | inteiro | data
    linhas:  [[...], [...]],      // matriz de valores
    cabecalho: true,              // false = nao emite a linha de titulos (matriz ja traz)
    linhasCabecalho: [0, 7],      // indices em `linhas` que devem sair com estilo de cabecalho
    congelar: {linha:1, coluna:1},// quantas linhas/colunas ficam fixas
    filtro: true                  // autofiltro na linha de cabecalho
  }
"""

NUCLEO_JS = r"""
  /* ---------------------------------------------------------------------- *
   * ZIP (o .xlsx e um zip)                                                 *
   * ---------------------------------------------------------------------- */
  var TABELA_CRC = (function(){
    var t = new Uint32Array(256);
    for(var n = 0; n < 256; n++){
      var c = n;
      for(var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(u8){
    var c = 0xFFFFFFFF;
    for(var i = 0; i < u8.length; i++) c = TABELA_CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // deflate-raw nativo quando existir (Chrome/Edge atuais); senao guarda sem
  // compressao - o arquivo fica maior, mas abre igual.
  async function comprimir(u8){
    if(typeof CompressionStream !== 'function') return { dados: u8, metodo: 0 };
    try{
      var fluxo = new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate-raw'));
      var buf = new Uint8Array(await new Response(fluxo).arrayBuffer());
      return { dados: buf, metodo: 8 };
    }catch(e){ return { dados: u8, metodo: 0 }; }
  }

  async function zipar(arquivos){
    var enc = new TextEncoder(), partes = [], central = [], deslocamento = 0;
    for(var i = 0; i < arquivos.length; i++){
      var f = arquivos[i];
      var nome = enc.encode(f.nome);
      var cru = typeof f.dados === 'string' ? enc.encode(f.dados) : f.dados;
      var crc = crc32(cru);
      var c = await comprimir(cru);
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0, true);
      lh.setUint16(8, c.metodo, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x2821, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, c.dados.length, true);
      lh.setUint32(22, cru.length, true); lh.setUint16(26, nome.length, true); lh.setUint16(28, 0, true);
      partes.push(new Uint8Array(lh.buffer), nome, c.dados);
      var cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true);
      cd.setUint16(8, 0, true); cd.setUint16(10, c.metodo, true);
      cd.setUint16(12, 0, true); cd.setUint16(14, 0x2821, true);
      cd.setUint32(16, crc, true); cd.setUint32(20, c.dados.length, true);
      cd.setUint32(24, cru.length, true); cd.setUint16(28, nome.length, true);
      cd.setUint32(42, deslocamento, true);
      central.push(new Uint8Array(cd.buffer), nome);
      deslocamento += 30 + nome.length + c.dados.length;
    }
    var tamanhoCentral = central.reduce(function(s, p){ return s + p.length; }, 0);
    var fim = new DataView(new ArrayBuffer(22));
    fim.setUint32(0, 0x06054b50, true);
    fim.setUint16(8, arquivos.length, true); fim.setUint16(10, arquivos.length, true);
    fim.setUint32(12, tamanhoCentral, true); fim.setUint32(16, deslocamento, true);
    return new Blob(partes.concat(central, [new Uint8Array(fim.buffer)]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ---------------------------------------------------------------------- *
   * Valores                                                                *
   * ---------------------------------------------------------------------- */
  function esc(v){
    return String(v == null ? '' : v)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function letra(n){
    var s = ''; n++;
    while(n > 0){ var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }
  // Excel conta dias desde 1899-12-30; 1970-01-01 = 25569.
  function serial(a, m, d){ return Date.UTC(a, m - 1, d) / 86400000 + 25569; }

  function maiuscula(v){
    try{ return String(v).toLocaleUpperCase('pt-BR'); }
    catch(e){ return String(v).toUpperCase(); }
  }
  // Aceita "1234.56" e "1.234,56" / "1234,56".
  function numero(v){
    if(typeof v === 'number') return isFinite(v) ? String(v) : null;
    var s = String(v).trim();
    if(!s) return null;
    if(/^-?\d+(\.\d+)?$/.test(s)) return s;
    s = s.replace(/\s|R\$/g, '');
    if(/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if(/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.');
    else return null;
    var n = Number(s);
    return isFinite(n) ? String(n) : null;
  }
  // Aceita ISO (2026-08-09) e brasileiro (09/08/2026).
  function dataSerial(v){
    var s = String(v).trim(), m;
    if((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return serial(+m[1], +m[2], +m[3]);
    if((m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/))) return serial(+m[3], +m[2], +m[1]);
    return null;
  }

  /* ---------------------------------------------------------------------- *
   * Estilos                                                                *
   * ---------------------------------------------------------------------- */
  var ESTILOS = { geral: 0, cabecalho: 1, moeda: 2, data: 3, inteiro: 4, texto: 5 };

  var XML_ESTILOS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00"/>' +
    '<numFmt numFmtId="165" formatCode="dd/mm/yyyy"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1E2340"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FF334155"/></left><right style="thin"><color rgb="FF334155"/></right>' +
    '<top style="thin"><color rgb="FF334155"/></top><bottom style="thin"><color rgb="FF334155"/></bottom>' +
    '<diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
    '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1">' +
    '<alignment horizontal="center"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>' +
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  /* ---------------------------------------------------------------------- *
   * Planilha                                                               *
   * ---------------------------------------------------------------------- */
  function celulaTexto(ref, v){
    return '<c r="' + ref + '" s="' + ESTILOS.texto + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
  }
  function celulaCabecalho(ref, v){
    return '<c r="' + ref + '" s="' + ESTILOS.cabecalho + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
  }

  function painelCongelado(cong){
    var lin = (cong && cong.linha) || 0, col = (cong && cong.coluna) || 0;
    if(!lin && !col) return '';
    var canto = letra(col) + (lin + 1);
    var atr = (col ? ' xSplit="' + col + '"' : '') + (lin ? ' ySplit="' + lin + '"' : '');
    var ativo = col && lin ? 'bottomRight' : (col ? 'topRight' : 'bottomLeft');
    return '<pane' + atr + ' topLeftCell="' + canto + '" activePane="' + ativo + '" state="frozen"/>';
  }

  function planilhaXml(spec){
    var colunas = spec.colunas || [];
    var linhas = spec.linhas || [];
    var comCabecalho = spec.cabecalho !== false && colunas.length > 0;
    var extras = {};
    (spec.linhasCabecalho || []).forEach(function(i){ extras[i] = true; });

    var nCols = colunas.length;
    linhas.forEach(function(l){ if(l && l.length > nCols) nCols = l.length; });
    if(!nCols) nCols = 1;
    var ultima = letra(nCols - 1);
    var totalLinhas = linhas.length + (comCabecalho ? 1 : 0);

    var partes = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'];
    partes.push('<dimension ref="A1:' + ultima + Math.max(1, totalLinhas) + '"/>');
    var congelado = painelCongelado(spec.congelar);
    partes.push('<sheetViews><sheetView' + (spec.ativa ? ' tabSelected="1"' : '') +
      ' workbookViewId="0">' + congelado + '</sheetView></sheetViews>');
    partes.push('<sheetFormatPr defaultRowHeight="15"/>');
    if(nCols){
      var cols = [];
      for(var i = 0; i < nCols; i++){
        var larg = (colunas[i] && colunas[i].largura) || spec.larguraPadrao || 15;
        cols.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + larg + '" customWidth="1"/>');
      }
      partes.push('<cols>' + cols.join('') + '</cols>');
    }

    partes.push('<sheetData>');
    var r = 1;
    if(comCabecalho){
      partes.push('<row r="1" ht="30" customHeight="1">' + colunas.map(function(c, i){
        return celulaCabecalho(letra(i) + '1', c.titulo);
      }).join('') + '</row>');
      r = 2;
    }
    for(var l = 0; l < linhas.length; l++){
      var linha = linhas[l] || [], celulas = [], ehCab = !!extras[l];
      for(var i2 = 0; i2 < linha.length; i2++){
        var bruto = linha[i2], ref = letra(i2) + r;
        if(bruto == null || bruto === '') continue;
        if(ehCab){ celulas.push(celulaCabecalho(ref, maiuscula(bruto))); continue; }
        var col = colunas[i2] || {};
        var tipo = col.tipo || 'texto';
        if(tipo === 'moeda' || tipo === 'inteiro'){
          var n = numero(bruto);
          if(n === null) celulas.push(celulaTexto(ref, maiuscula(bruto)));
          else celulas.push('<c r="' + ref + '" s="' + (tipo === 'moeda' ? ESTILOS.moeda : ESTILOS.inteiro) + '"><v>' + n + '</v></c>');
        }else if(tipo === 'data'){
          var s = dataSerial(bruto);
          if(s === null) celulas.push(celulaTexto(ref, maiuscula(bruto)));
          else celulas.push('<c r="' + ref + '" s="' + ESTILOS.data + '"><v>' + s + '</v></c>');
        }else{
          celulas.push(celulaTexto(ref, maiuscula(bruto)));
        }
      }
      partes.push('<row r="' + r + (ehCab ? '" ht="26" customHeight="1' : '') + '">' + celulas.join('') + '</row>');
      r++;
    }
    partes.push('</sheetData>');
    if(spec.filtro !== false && comCabecalho && linhas.length){
      partes.push('<autoFilter ref="A1:' + ultima + totalLinhas + '"/>');
    }
    partes.push('</worksheet>');
    return partes.join('');
  }

  function nomeAba(s, usados){
    var n = String(s || 'DADOS').replace(/[\[\]\:\*\?\/\\]/g, ' ').slice(0, 31).trim() || 'DADOS';
    var base = n, i = 2;
    while(usados[n.toLowerCase()]){ n = (base.slice(0, 28) + '_' + i).slice(0, 31); i++; }
    usados[n.toLowerCase()] = true;
    return n;
  }

  async function gerarLivro(abas){
    abas = (abas || []).filter(Boolean);
    if(!abas.length) abas = [{ aba: 'DADOS', colunas: [], linhas: [] }];
    var usados = {};
    var nomes = abas.map(function(a){ return nomeAba(a.aba, usados); });

    var tipos = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'];
    var folhas = [], relsWb = [], arquivos = [];
    abas.forEach(function(a, i){
      var n = i + 1;
      tipos.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ' +
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
      folhas.push('<sheet name="' + esc(nomes[i]) + '" sheetId="' + n + '" r:id="rId' + n + '"/>');
      relsWb.push('<Relationship Id="rId' + n + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
        'Target="worksheets/sheet' + n + '.xml"/>');
      a.ativa = (i === 0);
      arquivos.push({ nome: 'xl/worksheets/sheet' + n + '.xml', dados: planilhaXml(a) });
    });
    tipos.push('</Types>');
    relsWb.push('<Relationship Id="rIdStyles" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>');

    return zipar([
      { nome: '[Content_Types].xml', dados: tipos.join('') },
      { nome: '_rels/.rels', dados:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>' },
      { nome: 'xl/workbook.xml', dados:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets>' + folhas.join('') + '</sheets></workbook>' },
      { nome: 'xl/_rels/workbook.xml.rels', dados:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relsWb.join('') + '</Relationships>' },
      { nome: 'xl/styles.xml', dados: XML_ESTILOS }
    ].concat(arquivos));
  }

  // Assinatura usada pelo painel SE2: uma aba, com cabecalho e autofiltro.
  function gerarXlsx(aba, colunas, linhas){
    return gerarLivro([{ aba: aba, colunas: colunas, linhas: linhas,
                         cabecalho: true, congelar: { linha: 1 }, filtro: true }]);
  }

  function baixarPlanilha(blob, nome){
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 60000);
  }

  function carimboHoje(){
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
"""


def bloco_autonomo(marca):
    """Bloco <script> completo que instala window.PlanilhaSE2 sozinho.

    Usado pelo painel de fluxo, onde nao ha a camada de interceptacao de CSV do SE2.
    """
    return (
        "<script>\n"
        "/* === " + marca + " =============================\n"
        "   Gerador de planilha .xlsx sem biblioteca externa (ZIP + OpenXML).\n"
        "   Mesmo motor e mesmo visual do painel SE2 - a fonte unica dos dois e o\n"
        "   modulo planilha_excel_js.py. Nao editar a mao.\n"
        "   ========================================================================== */\n"
        "(function(){\n"
        "  'use strict';\n"
        + NUCLEO_JS +
        "\n  window.PlanilhaSE2 = {\n"
        "    gerar: gerarXlsx,\n"
        "    gerarLivro: gerarLivro,\n"
        "    baixar: baixarPlanilha,\n"
        "    hoje: carimboHoje\n"
        "  };\n"
        "})();\n"
        "</script>\n"
    )
