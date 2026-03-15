#!/usr/bin/env node
/**
 * gerar_laudo_AMA_v2.js
 * =====================
 * Script genérico de geração do Laudo AMA em formato .docx
 * Plataforma AMA — Oxy Recovery Wellness & Performance
 * Dr. Mateus Antunes Nogueira — CRM-SP 97.070
 *
 * Versão: v2.0 | Março 2026
 *
 * MUDANÇAS v1 → v2
 * ----------------
 * 1. Dimensões dos gráficos em gerarSlotGraficos() atualizadas para gráficos v4:
 *    - G1–G4: figsize=(8.07, 7.78) → largura=16.2 cm, altura=15.6 cm (ratio 0.9641)
 *    - G5:    figsize=(7.94, 7.78) → largura=16.2 cm, altura=15.8 cm (ratio 0.9798)
 *    Anterior (v3): alturas de 10 cm e 8 cm — incompatíveis com a nova proporção.
 *
 * 2. Nomenclatura de arquivos PNG alinhada à saída do gerar_graficos_AMA_v4.py:
 *    G1_mapa_metabolico_v4.png, G2_resposta_cardiaca_v4.png,
 *    G3_equivalentes_ventilatorios_v4.png, G4_rer_v4.png, G5_fatmax_v4.png
 *
 * 3. Referência ao pós-processador atualizada para ama_docx_postprocess_v2.py
 *    no cabeçalho de documentação (sem impacto em código).
 *
 * USO:
 *   node gerar_laudo_AMA_v2.js <input.json> [output.docx] [graficos_dir] [logo_path]
 *   cat input.json | node gerar_laudo_AMA_v2.js - [output.docx] [graficos_dir] [logo_path]
 *
 * ESTRUTURA DO JSON DE INPUT:
 *   Ver T29_v1_UserMessage_Template.md + campo `laudo_gerado` (output da Claude API)
 *
 * CAMPOS ESPERADOS NO JSON:
 *   - paciente         → dados do paciente
 *   - perfil_laudo     → "saude" | "desempenho"
 *   - laudo_gerado     → objeto com todas as seções (S01–S19) geradas pela Claude API
 *   - graficos_dir     → (opcional) caminho para pasta com PNGs dos gráficos AMA
 *
 * OUTPUT:
 *   Arquivo .docx pronto para pós-processamento (ama_docx_postprocess_v2.py)
 *
 * REGRAS INEGOCIÁVEIS:
 *   - Script single-file (sem require chains externas)
 *   - Footer em texto puro (sem logo — gera rId0 inválido no docx-js)
 *   - Cabeçalho com logo Oxy Recovery via imageRun inline
 *   - ShadingType.CLEAR obrigatório (SOLID gera fundo preto)
 *   - Imagens: fs.readFileSync no momento da chamada + dimensões em EMU
 */

'use strict';

// ─── DEPENDÊNCIAS ─────────────────────────────────────────────────────────────

const fs   = require('fs');
const path = require('path');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageNumberElement, PageBreak,
  LevelFormat, UnderlineType,
} = require('docx');


// ─── CONSTANTES DE LAYOUT ─────────────────────────────────────────────────────

// A4 em DXA (1440 DXA = 1 polegada = 2,54 cm)
const PAGE_W     = 11906;   // 21 cm
const PAGE_H     = 16838;   // 29,7 cm
const MARGIN     = 1134;    // 2 cm
const CONTENT_W  = PAGE_W - MARGIN * 2;  // ~9638 DXA (~17 cm úteis)

// EMU para imagens (1 cm = 360000 EMU)
const CM         = 360000;

// Paleta Oxy Recovery
const COR = {
  VERDE_ESC : '1A4A3A',
  VERDE_MED : '2E7D5E',
  VERDE_CLA : '4CAF82',
  LARANJA   : 'E07B39',
  CINZA_BG  : 'F7F9F7',
  CINZA_TAB : 'E8EDE8',
  BRANCO    : 'FFFFFF',
  PRETO     : '1A1A1A',
  CINZA_TXT : '4A4A4A',
};

// Fontes — Calibri conforme identidade Oxy Recovery
const FONTE = 'Calibri';

// Bordas padrão de tabela
const borda = (cor = 'CCCCCC') => ({
  style: BorderStyle.SINGLE, size: 1, color: cor,
});
const BORDAS_TAB = {
  top    : borda(), bottom : borda(),
  left   : borda(), right  : borda(),
  insideHorizontal: borda(), insideVertical: borda(),
};
const BORDAS_HEADER_TAB = {
  top    : borda(COR.VERDE_ESC), bottom : borda(COR.VERDE_ESC),
  left   : borda(COR.VERDE_ESC), right  : borda(COR.VERDE_ESC),
  insideHorizontal: borda(COR.VERDE_ESC), insideVertical: borda(COR.VERDE_ESC),
};

// Padding de célula padrão
const PAD = { top: 80, bottom: 80, left: 120, right: 120 };


// ─── HELPERS DE PARÁGRAFO ─────────────────────────────────────────────────────

/** Parágrafo de texto corrido — usado em corpo de seções */
function paraTexto(texto, { bold = false, size = 22, cor = COR.PRETO, espacoAntes = 0, espacoDepois = 120, italico = false, alignment = AlignmentType.JUSTIFIED } = {}) {
  return new Paragraph({
    alignment,
    spacing: { before: espacoAntes, after: espacoDepois, line: 276, lineRule: 'auto' },
    children: [new TextRun({
      text   : texto || '',
      bold, italics: italico, size, color: cor,
      font   : { name: FONTE },
    })],
  });
}

/** Título de seção — numerado, verde escuro */
function paraTituloSecao(numero, titulo, nivel = HeadingLevel.HEADING_2) {
  return new Paragraph({
    heading  : nivel,
    spacing  : { before: 360, after: 120 },
    children : [
      new TextRun({ text: `${numero}. `, bold: true, size: 26, color: COR.VERDE_ESC, font: { name: FONTE } }),
      new TextRun({ text: titulo.toUpperCase(), bold: true, size: 26, color: COR.VERDE_ESC, font: { name: FONTE } }),
    ],
  });
}

/** Subtítulo dentro de seção */
function paraSubtitulo(texto) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [new TextRun({
      text: texto, bold: true, size: 23,
      color: COR.VERDE_MED, font: { name: FONTE },
    })],
  });
}

/** Parágrafo vazio — espaçamento */
function paraVazio(altura = 80) {
  return new Paragraph({ spacing: { before: 0, after: altura }, children: [] });
}

/** Linha divisória usando borda inferior */
function paraDivisoria() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border : { bottom: { style: BorderStyle.SINGLE, size: 2, color: COR.VERDE_CLA } },
    children: [],
  });
}

/** Quebra de página */
function quebraPagina() {
  return new Paragraph({ children: [new PageBreak()] });
}

/** Label + Valor em linha única */
function paraLabelValor(label, valor, { sizeLabel = 20, sizeValor = 22 } = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: `${label}: `, bold: true, size: sizeLabel, color: COR.VERDE_ESC, font: { name: FONTE } }),
      new TextRun({ text: valor || '—', size: sizeValor, color: COR.PRETO, font: { name: FONTE } }),
    ],
  });
}


// ─── HELPERS DE TABELA ────────────────────────────────────────────────────────

/** Célula de cabeçalho de tabela */
function celulaHeader(texto, largura, { colspan = 1 } = {}) {
  return new TableCell({
    columnSpan : colspan,
    width      : { size: largura, type: WidthType.DXA },
    shading    : { fill: COR.VERDE_ESC, type: ShadingType.CLEAR },
    margins    : PAD,
    borders    : BORDAS_HEADER_TAB,
    children   : [new Paragraph({
      alignment: AlignmentType.CENTER,
      children : [new TextRun({ text: texto || '', bold: true, size: 20, color: COR.BRANCO, font: { name: FONTE } })],
    })],
  });
}

/** Célula de dado de tabela */
function celulaValor(texto, largura, { bold = false, corFundo = COR.BRANCO, alignment = AlignmentType.CENTER } = {}) {
  return new TableCell({
    width  : { size: largura, type: WidthType.DXA },
    shading: { fill: corFundo, type: ShadingType.CLEAR },
    margins: PAD,
    borders: BORDAS_TAB,
    children: [new Paragraph({
      alignment,
      children: [new TextRun({ text: texto != null ? String(texto) : '—', bold, size: 21, color: COR.PRETO, font: { name: FONTE } })],
    })],
  });
}

/** Linha de tabela simples com colunas alternadas */
function linhaTabela(celulas, alternar = false) {
  return new TableRow({
    children: celulas,
    cantSplit: true,
    tableHeader: false,
  });
}


// ─── CARGA DE IMAGEM ──────────────────────────────────────────────────────────

/**
 * Carrega imagem do disco e retorna ImageRun configurado.
 * @param {string} caminhoArquivo - Path absoluto ou relativo ao arquivo de imagem
 * @param {number} larguraCm - Largura em centímetros
 * @param {number} alturaCm  - Altura em centímetros
 */
function carregarImagem(caminhoArquivo, larguraCm, alturaCm) {
  try {
    const dados = fs.readFileSync(caminhoArquivo);
    const ext   = path.extname(caminhoArquivo).toLowerCase();
    const tipo  = ext === '.png' ? 'png' : ext === '.jpg' || ext === '.jpeg' ? 'jpg' : 'png';
    return new ImageRun({
      data  : dados,
      type  : tipo,
      transformation: {
        width : Math.round(larguraCm * CM),
        height: Math.round(alturaCm  * CM),
      },
    });
  } catch (e) {
    console.warn(`[WARN] Imagem não encontrada: ${caminhoArquivo} — slot vazio inserido`);
    return null;
  }
}

/** Parágrafo com imagem centralizada + legenda */
function paraImagem(imagemRun, legenda = '') {
  const blocos = [];
  if (imagemRun) {
    blocos.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing  : { before: 120, after: 60 },
      children : [imagemRun],
    }));
  }
  if (legenda) {
    blocos.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing  : { before: 0, after: 120 },
      children : [new TextRun({ text: legenda, italics: true, size: 18, color: COR.CINZA_TXT, font: { name: FONTE } })],
    }));
  }
  return blocos;
}


// ─── CABEÇALHO E RODAPÉ ───────────────────────────────────────────────────────

function construirCabecalho(logoPath) {
  const children = [];

  // Logo Oxy Recovery — carregado do disco no momento da chamada
  const logo = carregarImagem(logoPath, 4.5, 1.5);
  if (logo) {
    children.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing  : { before: 0, after: 80 },
      children : [logo],
    }));
  }

  // Linha de separação visual abaixo do logo
  children.push(new Paragraph({
    border  : { bottom: { style: BorderStyle.SINGLE, size: 3, color: COR.VERDE_ESC } },
    spacing : { before: 0, after: 80 },
    children: [],
  }));

  return new Header({ children });
}

function construirRodape(nomeCompleto, crm) {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing  : { before: 80, after: 0 },
        children : [
          new TextRun({ text: `Dr. Mateus Antunes Nogueira | CRM-SP ${crm || '97.070'} | Oxy Recovery Wellness & Performance`, size: 16, color: COR.CINZA_TXT, font: { name: FONTE } }),
          new TextRun({ text: '   |   Página ', size: 16, color: COR.CINZA_TXT, font: { name: FONTE } }),
          new PageNumberElement({ page: PageNumber.CURRENT }),
          new TextRun({ text: '   |   Documento Médico Confidencial', size: 16, color: COR.CINZA_TXT, font: { name: FONTE } }),
        ],
      }),
    ],
  });
}


// ─── CAPA (S01) ───────────────────────────────────────────────────────────────

function gerarCapa(paciente, perfil, dataAvaliacao, logoPath) {
  const filhos = [];

  // Espaço inicial
  filhos.push(paraVazio(720));

  // Logo grande centralizado na capa
  const logo = carregarImagem(logoPath, 6, 2);
  if (logo) {
    filhos.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing  : { before: 0, after: 360 },
      children : [logo],
    }));
  }

  // Título do documento
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 0, after: 80 },
    children : [new TextRun({ text: 'LAUDO DE AVALIAÇÃO METABÓLICA AVANÇADA', bold: true, size: 36, color: COR.VERDE_ESC, font: { name: FONTE } })],
  }));

  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 0, after: 600 },
    children : [new TextRun({ text: 'Plataforma AMA — Oxy Recovery Wellness & Performance', size: 24, color: COR.VERDE_MED, font: { name: FONTE } })],
  }));

  // Dados do paciente em tabela centralizada
  // Identificador usado pelo ama_docx_postprocess_v2.py para centralização: tblW w:w="5000"
  const perfilLabel = perfil === 'desempenho' ? 'Perfil B — Atleta de Desempenho' : 'Perfil A — Atleta da Saúde';
  const colW = 5000 / 2;  // 2500 DXA por coluna → tabela total = 5000 DXA

  const linhasTabela = [
    ['Paciente', paciente.nome_completo || ''],
    ['Data de Nascimento', paciente.data_nascimento || ''],
    ['Data da Avaliação', dataAvaliacao || paciente.data_avaliacao || ''],
    ['Perfil', perfilLabel],
    ['Médico Responsável', 'Dr. Mateus Antunes Nogueira'],
    ['CRM', 'SP 97.070'],
  ];

  const rows = linhasTabela.map(([label, valor]) => new TableRow({
    children: [
      celulaHeader(label, colW),
      celulaValor(valor, colW, { alignment: AlignmentType.LEFT }),
    ],
  }));

  filhos.push(new Table({
    width: { size: 5000, type: WidthType.DXA },
    columnWidths: [colW, colW],
    rows,
  }));

  filhos.push(paraVazio(480));

  // Aviso de confidencialidade
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 0, after: 0 },
    children : [new TextRun({ text: 'DOCUMENTO MÉDICO CONFIDENCIAL — USO EXCLUSIVO DO PACIENTE', size: 18, italics: true, color: COR.CINZA_TXT, font: { name: FONTE } })],
  }));

  return filhos;
}


// ─── SEÇÃO GENÉRICA DE TEXTO (S02–S13, S15–S18) ──────────────────────────────

/**
 * Renderiza uma seção a partir de texto livre (string ou array de strings).
 * O conteúdo vem da Claude API como texto narrativo já redigido.
 */
function gerarSecaoTexto(numero, titulo, conteudo) {
  const filhos = [];

  filhos.push(paraTituloSecao(numero, titulo));
  filhos.push(paraDivisoria());

  if (!conteudo) {
    filhos.push(paraTexto('Esta seção não foi gerada para o perfil deste paciente.', { italico: true, cor: COR.CINZA_TXT }));
    return filhos;
  }

  // Aceita string ou array de strings (um parágrafo por item)
  const paragrafos = Array.isArray(conteudo) ? conteudo : conteudo.split('\n\n').filter(p => p.trim());

  paragrafos.forEach(par => {
    const texto = par.trim();
    if (texto) {
      filhos.push(paraTexto(texto));
      filhos.push(paraVazio(40));
    }
  });

  return filhos;
}


// ─── SEÇÃO DE COMPOSIÇÃO CORPORAL (S03) — TABELA ─────────────────────────────

function gerarSecaoComposicao(numero, dados, reavaliacao = null) {
  const filhos = [];

  filhos.push(paraTituloSecao(numero, 'Composição Corporal'));
  filhos.push(paraDivisoria());

  const temReavaliacao = reavaliacao && reavaliacao.peso_kg;
  const col1 = Math.floor(CONTENT_W * 0.35);
  const col2 = Math.floor(CONTENT_W * (temReavaliacao ? 0.22 : 0.32));
  const col3 = temReavaliacao ? Math.floor(CONTENT_W * 0.22) : 0;
  const col4 = temReavaliacao ? Math.floor(CONTENT_W * 0.21) : Math.floor(CONTENT_W * 0.33);

  const headers = temReavaliacao
    ? [celulaHeader('Parâmetro', col1), celulaHeader('Baseline', col2), celulaHeader('Atual', col3), celulaHeader('Δ Variação', col4)]
    : [celulaHeader('Parâmetro', col1), celulaHeader('Valor', col2), celulaHeader('Unidade', col4)];

  const d = dados || {};
  const r = reavaliacao || {};

  const delta = (v1, v2, casas = 1) => {
    if (v1 == null || v2 == null) return '—';
    const diff = (parseFloat(v2) - parseFloat(v1)).toFixed(casas);
    return diff > 0 ? `+${diff}` : String(diff);
  };

  const linhas = [
    ['Peso corporal', d.peso_kg, r.peso_kg, 'kg'],
    ['IMC', d.imc, r.imc, 'kg/m²'],
    ['% Gordura', d.percentual_gordura, r.percentual_gordura, '%'],
    ['Massa gorda', d.massa_gorda_kg, r.massa_gorda_kg, 'kg'],
    ['Massa magra (SMM)', d.smm_kg, r.smm_kg, 'kg'],
    ['Gordura de tronco', d.gordura_tronco_kg, r.gordura_tronco_kg, 'kg'],
    ['% Gordura tronco', d.gordura_tronco_percentual, r.gordura_tronco_percentual, '%'],
    ['Score visceral', d.gordura_visceral_score, r.gordura_visceral_score, ''],
    ['Água corporal total', d.agua_corporal_total_l, r.agua_corporal_total_l, 'L'],
    ['TMR InBody', d.tmr_inbody_kcal, r.tmr_inbody_kcal, 'kcal/dia'],
    ['Circunferência abdominal', d.circunferencia_abdominal_cm, r.circunferencia_abdominal_cm, 'cm'],
  ].filter(([, v]) => v != null);

  const rows = [
    new TableRow({ tableHeader: true, children: headers }),
    ...linhas.map(([label, v1, v2, unidade], i) => {
      const corFundo = i % 2 === 0 ? COR.BRANCO : COR.CINZA_BG;
      if (temReavaliacao) {
        return new TableRow({ children: [
          celulaValor(label, col1, { alignment: AlignmentType.LEFT }),
          celulaValor(v1 != null ? `${v1} ${unidade}`.trim() : '—', col2, { corFundo }),
          celulaValor(v2 != null ? `${v2} ${unidade}`.trim() : '—', col3, { corFundo }),
          celulaValor(delta(v1, v2), col4, { bold: true, corFundo }),
        ]});
      } else {
        return new TableRow({ children: [
          celulaValor(label, col1, { alignment: AlignmentType.LEFT }),
          celulaValor(v1 != null ? String(v1) : '—', col2, { corFundo }),
          celulaValor(unidade, col4, { corFundo }),
        ]});
      }
    }),
  ];

  filhos.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: temReavaliacao ? [col1, col2, col3, col4] : [col1, col2, col4],
    rows,
  }));

  return filhos;
}


// ─── SEÇÃO DE ERGOESPIROMETRIA (S05) — TABELA DE LIMIARES ────────────────────

function gerarSecaoMapa(numero, ergo, ergoBaseline = null) {
  const filhos = [];

  filhos.push(paraTituloSecao(numero, 'Mapa Metabólico — Ergoespirometria'));
  filhos.push(paraDivisoria());

  const temBaseline = ergoBaseline && ergoBaseline.l1_fc_bpm;
  const col1 = Math.floor(CONTENT_W * 0.30);
  const col2 = Math.floor(CONTENT_W * (temBaseline ? 0.23 : 0.35));
  const col3 = temBaseline ? Math.floor(CONTENT_W * 0.23) : 0;
  const col4 = temBaseline ? Math.floor(CONTENT_W * 0.24) : Math.floor(CONTENT_W * 0.35);

  const e = ergo || {};
  const b = ergoBaseline || {};

  const modalidade = e.protocolo_equipamento === 'cicloergometro' ? 'W' : 'km/h';

  const linhas = [
    ['L1 — Limiar Aeróbio', `${e.l1_fc_bpm || '—'} bpm`, b.l1_fc_bpm ? `${b.l1_fc_bpm} bpm` : null, `${e.l1_velocidade_ou_potencia || '—'} ${modalidade}`],
    ['Crossover (RER≈1,0)', `${e.crossover_fc_bpm || '—'} bpm`, b.crossover_fc_bpm ? `${b.crossover_fc_bpm} bpm` : null, `${e.crossover_velocidade_ou_potencia || '—'} ${modalidade}`],
    ['L2 — Limiar Anaeróbio', `${e.l2_fc_bpm || '—'} bpm`, b.l2_fc_bpm ? `${b.l2_fc_bpm} bpm` : null, `${e.l2_velocidade_ou_potencia || '—'} ${modalidade}`],
    ['FC Pico', `${e.fc_pico_bpm || '—'} bpm`, b.fc_pico_bpm ? `${b.fc_pico_bpm} bpm` : null, '—'],
    ['VO₂ pico', `${e.vo2max_ml_kg_min || '—'} ml/kg/min`, b.vo2max_ml_kg_min ? `${b.vo2max_ml_kg_min} ml/kg/min` : null, '—'],
    ['FATmax', `${e.fatmax_fc_bpm || '—'} bpm`, b.fatmax_fc_bpm ? `${b.fatmax_fc_bpm} bpm` : null, `${e.fatmax_g_min || '—'} g/min`],
  ];

  const headers = temBaseline
    ? [celulaHeader('Parâmetro', col1), celulaHeader('FC Baseline', col2), celulaHeader('FC Atual', col3), celulaHeader('Carga / VO₂', col4)]
    : [celulaHeader('Parâmetro', col1), celulaHeader('FC', col2), celulaHeader('Carga / VO₂', col4)];

  const rows = [
    new TableRow({ tableHeader: true, children: headers }),
    ...linhas.map(([label, fc, fcBase, carga], i) => {
      const corFundo = i % 2 === 0 ? COR.BRANCO : COR.CINZA_BG;
      if (temBaseline) {
        return new TableRow({ children: [
          celulaValor(label, col1, { alignment: AlignmentType.LEFT }),
          celulaValor(fcBase || '—', col2, { corFundo }),
          celulaValor(fc, col3, { bold: true, corFundo }),
          celulaValor(carga, col4, { corFundo }),
        ]});
      } else {
        return new TableRow({ children: [
          celulaValor(label, col1, { alignment: AlignmentType.LEFT }),
          celulaValor(fc, col2, { bold: true, corFundo }),
          celulaValor(carga, col4, { corFundo }),
        ]});
      }
    }),
  ];

  filhos.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: temBaseline ? [col1, col2, col3, col4] : [col1, col2, col4],
    rows,
  }));

  return filhos;
}


// ─── SEÇÃO DE SLOTS DE GRÁFICOS ───────────────────────────────────────────────

/**
 * Insere os 5 gráficos clínicos padrão AMA a partir de uma pasta.
 *
 * Nomenclatura esperada (saída do gerar_graficos_AMA_v4.py):
 *   G1_mapa_metabolico_v4.png
 *   G2_resposta_cardiaca_v4.png
 *   G3_equivalentes_ventilatorios_v4.png
 *   G4_rer_v4.png
 *   G5_fatmax_v4.png
 *
 * Dimensões calculadas para gráficos v4 (figsize a 180 DPI, 95% da largura útil):
 *   G1–G4: figsize=(8.07, 7.78) → largura=16.2 cm, altura=15.6 cm  (ratio 0.9641)
 *   G5:    figsize=(7.94, 7.78) → largura=16.2 cm, altura=15.8 cm  (ratio 0.9798)
 *
 * ATENÇÃO: se renomear os PNGs manualmente, manter o mapeamento correto aqui.
 */
function gerarSlotGraficos(graficosDir) {
  const filhos = [];

  if (!graficosDir || !fs.existsSync(graficosDir)) {
    filhos.push(paraTexto('[Gráficos não disponíveis — pasta não encontrada]', { italico: true, cor: COR.CINZA_TXT }));
    return filhos;
  }

  // ── v4: dimensões atualizadas para nova altura (~17,6 cm) dos gráficos ──────
  const graficos = [
    {
      arquivo : 'G1_mapa_metabolico_v4.png',
      legenda : 'Figura 1 — Mapa Metabólico Modificado: Oxidação de Substratos por Intensidade',
      largura : 16.2,
      altura  : 15.6,   // ratio 0.9641 — figsize=(8.07, 7.78) a 180 DPI
    },
    {
      arquivo : 'G2_resposta_cardiaca_v4.png',
      legenda : 'Figura 2 — Resposta Cardíaca e Cinética de Gases ao Exercício',
      largura : 16.2,
      altura  : 15.6,
    },
    {
      arquivo : 'G3_equivalentes_ventilatorios_v4.png',
      legenda : 'Figura 3 — Equivalentes Ventilatórios (VE/VO₂ e VE/VCO₂)',
      largura : 16.2,
      altura  : 15.6,
    },
    {
      arquivo : 'G4_rer_v4.png',
      legenda : 'Figura 4 — Relação de Troca Respiratória (RER) ao Longo do Protocolo',
      largura : 16.2,
      altura  : 15.6,
    },
    {
      arquivo : 'G5_fatmax_v4.png',
      legenda : 'Figura 5 — FATmax: Pico de Oxidação de Gordura e Deslocamento da Curva',
      largura : 16.2,
      altura  : 15.8,   // ratio 0.9798 — figsize=(7.94, 7.78) a 180 DPI
    },
  ];

  graficos.forEach(g => {
    const caminho = path.join(graficosDir, g.arquivo);
    const imagem  = carregarImagem(caminho, g.largura, g.altura);
    if (imagem) {
      const blocos = paraImagem(imagem, g.legenda);
      blocos.forEach(b => filhos.push(b));
      filhos.push(paraVazio(240));
    }
  });

  return filhos;
}


// ─── TABELA DO TREINADOR (S13 — Perfil B exclusivo) ──────────────────────────

function gerarTabelaTreinador(dadosTreinador, competicoes) {
  const filhos = [];

  filhos.push(paraSubtitulo('Informações para o Treinador'));

  if (!dadosTreinador) {
    filhos.push(paraTexto('Tabela do Treinador não disponível.', { italico: true }));
    return filhos;
  }

  const col1 = Math.floor(CONTENT_W * 0.35);
  const col2 = Math.floor(CONTENT_W * 0.65);

  const linhasConteudo = [
    ['VO₂ pico', dadosTreinador.vo2pico || '—'],
    ['FC Máxima no Teste', dadosTreinador.fc_max_teste || '—'],
    ['L1 — FC / Carga', dadosTreinador.l1_resumo || '—'],
    ['L2 — FC / Carga', dadosTreinador.l2_resumo || '—'],
    ['FATmax — FC / Carga', dadosTreinador.fatmax_resumo || '—'],
    ['Zona FATmax (Z2)', dadosTreinador.zona_fatmax || '—'],
    ['Distribuição recomendada', dadosTreinador.distribuicao || '80% Z1-Z2 / 20% Z4-Z5'],
    ['Observação clínica', dadosTreinador.observacao || '—'],
  ];

  const rows = [
    new TableRow({ tableHeader: true, children: [celulaHeader('Parâmetro', col1), celulaHeader('Referência Clínica', col2)] }),
    ...linhasConteudo.map(([label, valor], i) => new TableRow({ children: [
      celulaValor(label, col1, { alignment: AlignmentType.LEFT, corFundo: i % 2 === 0 ? COR.BRANCO : COR.CINZA_BG }),
      celulaValor(valor, col2, { alignment: AlignmentType.LEFT, corFundo: i % 2 === 0 ? COR.BRANCO : COR.CINZA_BG }),
    ]})),
  ];

  filhos.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [col1, col2],
    rows,
  }));

  // Competições próximas
  if (competicoes && competicoes.length > 0) {
    filhos.push(paraVazio(160));
    filhos.push(paraSubtitulo('Metas para o Próximo Ciclo'));

    const colC = [Math.floor(CONTENT_W * 0.40), Math.floor(CONTENT_W * 0.30), Math.floor(CONTENT_W * 0.30)];
    const rowsComp = [
      new TableRow({ tableHeader: true, children: [celulaHeader('Competição / Meta', colC[0]), celulaHeader('Data', colC[1]), celulaHeader('Distância / Formato', colC[2])] }),
      ...competicoes.map((c, i) => new TableRow({ children: [
        celulaValor(c.nome || '—', colC[0], { alignment: AlignmentType.LEFT, corFundo: i % 2 === 0 ? COR.BRANCO : COR.CINZA_BG }),
        celulaValor(c.data || '—', colC[1]),
        celulaValor(c.distancia_formato || '—', colC[2]),
      ]})),
    ];

    filhos.push(new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: colC,
      rows: rowsComp,
    }));
  }

  return filhos;
}


// ─── MONTAGEM DO DOCUMENTO ────────────────────────────────────────────────────

/**
 * Função principal — monta o documento .docx completo.
 *
 * @param {Object} input - JSON completo do pipeline (T29 UserMessage + laudo_gerado)
 * @param {string} graficosDir - Caminho absoluto para pasta com PNGs dos gráficos
 * @param {string} logoPath    - Caminho absoluto para o logo Oxy Recovery
 * @returns {Promise<Buffer>} Buffer do .docx pronto para gravar em disco
 */
async function montarDocumento(input, graficosDir, logoPath) {
  const paciente = input.paciente  || {};
  const perfil   = input.perfil_laudo || 'saude';
  const ergo     = input.ergoespirometria || {};
  const calorim  = input.calorimetria_indireta || {};
  const corpo    = input.composicao_corporal || {};
  const anamese  = input.anamnese || {};
  const lab      = input.laboratorio || {};
  const instruc  = input.instrucoes_especiais || {};
  const laudo    = input.laudo_gerado || {};  // output da Claude API

  // Suporte a reavaliação longitudinal
  const ergoBase = input.ergoespirometria_baseline || null;
  const corpoBase= input.composicao_corporal_baseline || null;

  const isPerfilB = perfil === 'desempenho';

  // ── Construir children do documento ──────────────────────────────────────

  let children = [];

  // ── CAPA ──────────────────────────────────────────────────────────────────
  children.push(...gerarCapa(paciente, perfil, paciente.data_avaliacao, logoPath));
  children.push(quebraPagina());

  // ── S02 — Anamnese ────────────────────────────────────────────────────────
  children.push(...gerarSecaoTexto('S02', 'Anamnese Clínica', laudo.S02));
  children.push(paraVazio(120));

  // ── S03 — Composição Corporal ─────────────────────────────────────────────
  children.push(...gerarSecaoComposicao('S03', corpo, corpoBase));
  children.push(paraVazio(120));
  if (laudo.S03_narrativa) {
    children.push(...gerarSecaoTexto('', '', laudo.S03_narrativa));
  }

  // ── S04 — Calorimetria Indireta ───────────────────────────────────────────
  children.push(...gerarSecaoTexto('S04', 'Calorimetria Indireta', laudo.S04));
  children.push(paraVazio(120));

  // ── S05 — Mapa Metabólico ─────────────────────────────────────────────────
  children.push(...gerarSecaoMapa('S05', ergo, ergoBase));
  children.push(paraVazio(120));
  if (laudo.S05_narrativa) {
    children.push(...gerarSecaoTexto('', '', laudo.S05_narrativa));
  }

  // ── GRÁFICOS CLÍNICOS ─────────────────────────────────────────────────────
  children.push(quebraPagina());
  children.push(paraTituloSecao('', 'Análise Gráfica — Dados Ergoespirométricos', HeadingLevel.HEADING_2));
  children.push(paraDivisoria());
  children.push(...gerarSlotGraficos(graficosDir));

  // ── S06 — Resposta Cardíaca ───────────────────────────────────────────────
  children.push(quebraPagina());
  children.push(...gerarSecaoTexto('S06', 'Resposta Cardíaca ao Exercício', laudo.S06));

  // ── S07 — Equivalentes Ventilatórios ──────────────────────────────────────
  children.push(...gerarSecaoTexto('S07', 'Equivalentes Ventilatórios', laudo.S07));

  // ── S08 — RER ─────────────────────────────────────────────────────────────
  children.push(...gerarSecaoTexto('S08', 'Razão de Troca Respiratória (RER)', laudo.S08));

  // ── S09 — FATmax ──────────────────────────────────────────────────────────
  children.push(...gerarSecaoTexto('S09', 'FATmax — Curva de Oxidação de Gordura', laudo.S09));

  // ── S10 — Laboratório ─────────────────────────────────────────────────────
  children.push(...gerarSecaoTexto('S10', 'Perfil Laboratorial', laudo.S10));

  // ── S11 + S12 — Exclusivo Perfil B ────────────────────────────────────────
  if (isPerfilB) {
    children.push(quebraPagina());
    children.push(...gerarSecaoTexto('S11', 'Análise do Condicionamento Aeróbio', laudo.S11));
    children.push(...gerarSecaoTexto('S12', 'Análise do Condicionamento Anaeróbio', laudo.S12));
  }

  // ── S13 — Diagnóstico Funcional Integrado ─────────────────────────────────
  children.push(quebraPagina());
  children.push(...gerarSecaoTexto('S13', 'Diagnóstico Funcional Integrado', laudo.S13));

  // Tabela do Treinador — Perfil B exclusivo
  if (isPerfilB && laudo.tabela_treinador) {
    children.push(paraVazio(120));
    children.push(...gerarTabelaTreinador(laudo.tabela_treinador, anamese.competicoes_proximas));
  }

  // ── S13b — Condutas Médicas Complementares ────────────────────────────────
  if (laudo.S13b) {
    children.push(...gerarSecaoTexto('S13b', 'Condutas Médicas Complementares', laudo.S13b));
  }

  // ── S14 — Prescrição de Atividade Física (Perfil A exclusivo) ────────────
  if (!isPerfilB && laudo.S14) {
    children.push(quebraPagina());
    children.push(...gerarSecaoTexto('S14', 'Prescrição de Atividade Física', laudo.S14));
  }

  // ── S15 — Prescrição Nutricional ──────────────────────────────────────────
  children.push(quebraPagina());
  children.push(...gerarSecaoTexto('S15', 'Prescrição Nutricional', laudo.S15));

  // ── S16 — Suplementação ───────────────────────────────────────────────────
  children.push(...gerarSecaoTexto('S16', 'Suplementação', laudo.S16));

  // ── S17 — Conclusão ───────────────────────────────────────────────────────
  children.push(quebraPagina());
  children.push(...gerarSecaoTexto('S17', 'Conclusão e Próximos Passos', laudo.S17));

  // ── S18 — Orientações ao Profissional de Exercício ────────────────────────
  const tituloS18 = isPerfilB ? 'Orientações ao Treinador' : 'Orientações ao Personal Trainer';
  children.push(...gerarSecaoTexto('S18', tituloS18, laudo.S18));

  // ── S19 — Documentos (receitas, suplementação, encaminhamentos) ───────────
  if (laudo.S19 && Object.keys(laudo.S19).length > 0) {
    children.push(quebraPagina());
    children.push(paraTituloSecao('S19', 'Documentos Médicos', HeadingLevel.HEADING_1));

    if (laudo.S19.receitas)           children.push(...gerarSecaoTexto('S19a', 'Receituário Médico', laudo.S19.receitas));
    if (laudo.S19.suplementacao)      children.push(...gerarSecaoTexto('S19b', 'Protocolo de Suplementação', laudo.S19.suplementacao));
    if (laudo.S19.encaminhamentos)    children.push(...gerarSecaoTexto('S19c', 'Encaminhamentos', laudo.S19.encaminhamentos));
    if (laudo.S19.ohb)                children.push(...gerarSecaoTexto('S19d', 'Solicitação de Oxigenoterapia Hiperbárica', laudo.S19.ohb));
    if (laudo.S19.investigacao_geno)  children.push(...gerarSecaoTexto('S19e', 'Solicitação de Investigação Genômica', laudo.S19.investigacao_geno));
    if (laudo.S19.investigacao_meta)  children.push(...gerarSecaoTexto('S19f', 'Solicitação de Investigação Metabolômica', laudo.S19.investigacao_meta));
  }

  // ── Assinatura final ──────────────────────────────────────────────────────
  children.push(paraVazio(480));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    children : [new TextRun({ text: '_______________________________________________', size: 22, font: { name: FONTE } })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 60, after: 0 },
    children : [new TextRun({ text: 'Dr. Mateus Antunes Nogueira', bold: true, size: 22, font: { name: FONTE } })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 0, after: 0 },
    children : [new TextRun({ text: 'CRM-SP 97.070 | Cirurgião do Aparelho Digestivo | Médico do Exercício e do Esporte | Nutrólogo', size: 19, color: COR.CINZA_TXT, font: { name: FONTE } })],
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing  : { before: 0, after: 0 },
    children : [new TextRun({ text: 'Oxy Recovery Wellness & Performance', size: 19, color: COR.VERDE_MED, font: { name: FONTE } })],
  }));

  // ── Numeração de listas (bullets) ─────────────────────────────────────────
  const numeracao = {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  };

  // ── Estilos globais ───────────────────────────────────────────────────────
  const estilos = {
    default: {
      document: { run: { font: FONTE, size: 22, color: COR.PRETO } },
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: FONTE, color: COR.VERDE_ESC },
        paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: FONTE, color: COR.VERDE_ESC },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 1 },
      },
    ],
  };

  // ── Documento final ───────────────────────────────────────────────────────
  const doc = new Document({
    styles  : estilos,
    numbering: numeracao,
    sections: [{
      properties: {
        page: {
          size  : { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: { default: construirCabecalho(logoPath) },
      footers: { default: construirRodape(paciente.nome_completo, '97.070') },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}


// ─── PONTO DE ENTRADA (CLI) ───────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);

  if (args.length === 0 || args[0] === '--help') {
    console.log(`
USO:
  node gerar_laudo_AMA_v2.js <input.json> [output.docx] [graficos_dir] [logo_path]

ARGUMENTOS:
  input.json    JSON do pipeline (T29 UserMessage + campo laudo_gerado)
                Use "-" para ler do stdin

  output.docx   Caminho de saída (default: laudo_AMA_output.docx)

  graficos_dir  Pasta com PNGs dos gráficos AMA v4:
                  G1_mapa_metabolico_v4.png
                  G2_resposta_cardiaca_v4.png
                  G3_equivalentes_ventilatorios_v4.png
                  G4_rer_v4.png
                  G5_fatmax_v4.png
                (default: ./graficos)

  logo_path     Caminho para Logo_Principal_Oxy_Recovery_Verde.jpg
                (default: ./Logo_Principal_Oxy_Recovery_Verde.jpg)

PIPELINE COMPLETO:
  1. python gerar_graficos_AMA_v4.py          → gera PNGs em ./graficos/
  2. node   gerar_laudo_AMA_v2.js dados.json laudo.docx ./graficos ./logo.jpg
  3. python ama_docx_postprocess_v2.py laudo.docx laudo_final.docx

EXEMPLO:
  node gerar_laudo_AMA_v2.js dados_paciente.json laudo_ana.docx ./graficos_ana ./assets/logo.jpg
`);
    process.exit(0);
  }

  // Ler JSON de input
  let jsonBruto;
  if (args[0] === '-') {
    jsonBruto = '';
    process.stdin.setEncoding('utf8');
    for await (const chunk of process.stdin) jsonBruto += chunk;
  } else {
    jsonBruto = fs.readFileSync(args[0], 'utf8');
  }

  let input;
  try {
    input = JSON.parse(jsonBruto);
  } catch (e) {
    console.error('[ERRO] JSON inválido:', e.message);
    process.exit(1);
  }

  const outputPath   = args[1] || 'laudo_AMA_output.docx';
  const graficosDir  = args[2] || input.graficos_dir || path.join(process.cwd(), 'graficos');
  const logoPath     = args[3] || input.logo_path    || path.join(process.cwd(), 'Logo_Principal_Oxy_Recovery_Verde.jpg');

  console.log(`[AMA v2] Gerando laudo para: ${input.paciente?.nome_completo || 'Paciente'}`);
  console.log(`[AMA v2] Perfil: ${input.perfil_laudo || 'saude'}`);
  console.log(`[AMA v2] Gráficos (v4): ${graficosDir}`);
  console.log(`[AMA v2] Logo: ${logoPath}`);

  try {
    const buffer = await montarDocumento(input, graficosDir, logoPath);
    fs.writeFileSync(outputPath, buffer);
    console.log(`[AMA v2] ✅ Laudo gerado: ${outputPath} (${(buffer.length / 1024).toFixed(0)} KB)`);
    console.log(`[AMA v2] → Próximo passo: python ama_docx_postprocess_v2.py ${outputPath}`);
  } catch (e) {
    console.error('[ERRO] Falha na geração do .docx:', e.message);
    console.error(e.stack);
    process.exit(1);
  }
}

main().catch(e => { console.error(e); process.exit(1); });
