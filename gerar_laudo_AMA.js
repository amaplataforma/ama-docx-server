'use strict';
// gerar_laudo_AMA.js — Plataforma AMA
// PLACEHOLDER: Substituir pelo arquivo completo do Project Knowledge
// Assinatura: node gerar_laudo_AMA.js <input.json> <output.docx> <graficos_dir> <logo_path>

const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, AlignmentType } = require('docx');

async function main() {
  const [,, inputJson, outputDocx, graficosDir, logoPath] = process.argv;

  let dados = {};
  if (inputJson && inputJson !== '-') {
    dados = JSON.parse(fs.readFileSync(inputJson, 'utf8'));
  } else {
    const stdin = fs.readFileSync('/dev/stdin', 'utf8');
    dados = JSON.parse(stdin);
  }

  const laudoGerado = dados.laudo_gerado || {};
  const paciente = dados.paciente || {};

  const paragraphs = [];

  // Titulo
  paragraphs.push(new Paragraph({
    children: [new TextRun({ text: 'LAUDO — AVALIACAO METABOLICA AVANCADA (AMA)', bold: true, size: 28 })],
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
  }));

  // Dados do paciente
  paragraphs.push(new Paragraph({
    children: [new TextRun({ text: 'Paciente: ' + (paciente.nome_completo || ''), bold: true })],
  }));
  paragraphs.push(new Paragraph({
    children: [new TextRun({ text: 'Data: ' + (paciente.data_avaliacao || '') })],
  }));
  paragraphs.push(new Paragraph({ children: [new TextRun('')] }));

  // Secoes do laudo gerado
  const secoes = Object.keys(laudoGerado).sort();
  for (const secao of secoes) {
    const texto = laudoGerado[secao];
    if (texto) {
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: secao, bold: true, size: 24 })],
        heading: HeadingLevel.HEADING_2,
      }));
      const linhas = texto.split('\n');
      for (const linha of linhas) {
        paragraphs.push(new Paragraph({
          children: [new TextRun({ text: linha })],
        }));
      }
      paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
    }
  }

  // Graficos
  const graficosNomes = [
    'G1_mapa_metabolico_v4.png',
    'G2_resposta_cardiaca_v4.png',
    'G3_equivalentes_ventilatorios_v4.png',
    'G4_rer_v4.png',
    'G5_fatmax_v4.png',
  ];

  for (const nomeGrafico of graficosNomes) {
    const grafPath = path.join(graficosDir, nomeGrafico);
    if (fs.existsSync(grafPath)) {
      try {
        const imgBuffer = fs.readFileSync(grafPath);
        paragraphs.push(new Paragraph({
          children: [new ImageRun({
            data: imgBuffer,
            transformation: { width: 500, height: 300 },
          })],
          alignment: AlignmentType.CENTER,
        }));
        paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
        console.log('[AMA v2] Grafico ' + nomeGrafico + ' inserido');
      } catch (e) {
        console.error('[AMA v2] Erro ao inserir grafico ' + nomeGrafico + ': ' + e.message);
      }
    }
  }

  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputDocx, buffer);
  console.log('[AMA v2] Laudo gerado: ' + outputDocx);
}

main().catch(err => {
  console.error('[AMA v2] Erro:', err.message);
  process.exit(1);
});
