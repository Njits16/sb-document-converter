const fetch = require("node-fetch");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const { Document, Paragraph, TextRun, HeadingLevel, Packer } = require("docx");
const PptxGenJS = require("pptxgenjs");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");
const JSZip = require("jszip");

// ── SB HUISSTIJL KLEUREN ──
const SB = {
  DARK:   "FF222222",
  ORANGE: "FFFE9933",
  WHITE:  "FFFFFFFF",
  LIGHT:  "FFF5F5F5",
  DARK2:  "FF444444",
  YELLOW: "FFFFDE7",
};

// ── FORMULE SHIFTER ──
function shiftFormula(formula, rowShift) {
  if (!formula || typeof formula !== "string" || !formula.startsWith("=")) return formula;
  const result = [];
  let i = 0;
  let inString = false;
  while (i < formula.length) {
    if (formula[i] === '"') { inString = !inString; result.push(formula[i++]); continue; }
    if (inString) { result.push(formula[i++]); continue; }
    if (formula[i] === "'") {
      let j = i + 1;
      while (j < formula.length && formula[j] !== "'") j++;
      result.push(formula.slice(i, j + 1)); i = j + 1; continue;
    }
    const m = formula.slice(i).match(/^(\$?)([A-Z]+)(\$?)(\d+)/);
    if (m) {
      const [full, dc, col, dr, rowStr] = m;
      const rowNum = dr ? parseInt(rowStr) : parseInt(rowStr) + rowShift;
      result.push(`${dc}${col}${dr}${rowNum}`);
      i += full.length;
    } else {
      result.push(formula[i++]);
    }
  }
  return result.join("");
}

// ── SB EXCEL STYLING ENGINE ──
async function applyExcelStyle(inputBuffer) {
  const ROW_SHIFT = 2;

  // Lees het originele bestand met ExcelJS (behoudt formules)
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(inputBuffer);

  for (const worksheet of workbook.worksheets) {
    const maxRow = worksheet.rowCount;
    const maxCol = worksheet.columnCount;

    // ── STAP 1: Verzamel en verschuif alle formules ──
    const formulas = {};
    worksheet.eachRow((row, rowNum) => {
      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        if (cell.formula) {
          formulas[`${rowNum}_${colNum}`] = shiftFormula(`=${cell.formula}`, ROW_SHIFT);
        }
      });
    });

    // ── STAP 2: Verschuif bestaande rijen 2 naar beneden ──
    // Kopieer alle rijen van onder naar boven om te voorkomen dat data overschreven wordt
    for (let r = maxRow; r >= 1; r--) {
      const srcRow = worksheet.getRow(r);
      const dstRow = worksheet.getRow(r + ROW_SHIFT);

      srcRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const dstCell = dstRow.getCell(colNum);
        if (cell.formula) {
          dstCell.value = { formula: formulas[`${r}_${colNum}`]?.slice(1) || cell.formula };
        } else {
          dstCell.value = cell.value;
        }
        dstCell.style = JSON.parse(JSON.stringify(cell.style));
      });
      dstRow.height = srcRow.height;
    }

    // Wis originele rijen 1 en 2
    for (let r = 1; r <= ROW_SHIFT; r++) {
      const row = worksheet.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
        cell.style = {};
      });
    }

    // ── STAP 3: Kolombreedtes instellen ──
    // Bewaar bestaande breedtes maar stel kolom A smaller in als er nummering is
    const colAWidth = worksheet.getColumn(1).width;
    if (!colAWidth || colAWidth > 10) {
      worksheet.getColumn(1).width = 6;
    }

    // ── STAP 4: Rij 1 — SB header ──
    const headerRow = worksheet.getRow(1);
    headerRow.height = 40;
    const headerCell = headerRow.getCell(1);
    headerCell.value = `SB PROCESMANAGEMENT  |  ${worksheet.name}`;
    headerCell.font = { name: "Impact", size: 14, color: { argb: "FFFFFFFF" }, bold: false };
    headerCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
    headerCell.alignment = { horizontal: "left", vertical: "middle" };

    // Kleur alle cellen in rij 1 donker
    for (let c = 2; c <= maxCol; c++) {
      const cell = headerRow.getCell(c);
      cell.value = null;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
    }

    // ── STAP 5: Rij 2 — oranje streep ──
    const orangeRow = worksheet.getRow(2);
    orangeRow.height = 4;
    for (let c = 1; c <= maxCol; c++) {
      const cell = orangeRow.getCell(c);
      cell.value = null;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFE9933" } };
    }

    // ── STAP 6: Detecteer rij 3 — metainfo of kolomkoppen ──
    const row3 = worksheet.getRow(3);
    const row3Vals = [];
    for (let c = 1; c <= maxCol; c++) {
      const v = row3.getCell(c).value;
      if (v && typeof v === "string" && !v.startsWith("=")) row3Vals.push(v.trim());
    }
    const hasMetaInfo = row3Vals.some(v => v.match(/^(Datum|Ingevuld|Scope|Date|Name)/i));
    const hasColHeaders = row3Vals.length > 0 && !hasMetaInfo;

    let dataStart;

    if (hasMetaInfo) {
      // Rijen 3-5: metainfo (lichtgrijs)
      for (let r = 3; r <= 5; r++) {
        const row = worksheet.getRow(r);
        row.height = 18;
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          const hasVal = cell.value && typeof cell.value === "string" && cell.value.trim();
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
          cell.font = { name: "Aptos", size: 10, bold: !!hasVal, color: { argb: "FF222222" } };
          if (hasVal) cell.alignment = { horizontal: "left", vertical: "middle" };
        }
      }
      // Rij 6: lege spacer
      const spacer = worksheet.getRow(6);
      spacer.height = 6;
      for (let c = 1; c <= maxCol; c++) {
        spacer.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
      }
      // Rij 7: kolomkoppen
      const colHeaderRow = worksheet.getRow(7);
      colHeaderRow.height = 55;
      for (let c = 1; c <= maxCol; c++) {
        const cell = colHeaderRow.getCell(c);
        if (cell.value) {
          cell.font = { name: "Aptos", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
          cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        }
      }
      // Rij 8: scorewaarden rij (oranje)
      const scoreRow = worksheet.getRow(8);
      scoreRow.height = 18;
      for (let c = 1; c <= maxCol; c++) {
        const cell = scoreRow.getCell(c);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFE9933" } };
        if (cell.value) {
          cell.font = { name: "Aptos", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
      dataStart = 9;
    } else if (hasColHeaders) {
      const colHeaderRow = worksheet.getRow(3);
      colHeaderRow.height = 22;
      for (let c = 1; c <= maxCol; c++) {
        const cell = colHeaderRow.getCell(c);
        cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
        cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      }
      dataStart = 4;
    } else {
      dataStart = 3;
    }

    // ── STAP 7: Data rijen ──
    const newMaxRow = maxRow + ROW_SHIFT;
    for (let r = dataStart; r <= newMaxRow; r++) {
      const row = worksheet.getRow(r);
      const isOdd = (r - dataStart) % 2 === 0;
      const rowBg = isOdd ? "FFF5F5F5" : "FFFFFFFF";

      const firstVal = row.getCell(1).value;
      const isNumbered = (
        firstVal !== null &&
        firstVal !== undefined &&
        !isNaN(firstVal) &&
        typeof firstVal !== "boolean" &&
        Number.isInteger(Number(firstVal))
      );

      // Check of dit de totaalrij is (laatste rij met tekst in kolom A)
      const isTotal = (r === newMaxRow && firstVal && typeof firstVal === "string" && firstVal.length > 3);

      if (isTotal) {
        row.height = 32;
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF222222" } };
          cell.font = {
            name: "Impact", size: c === 1 ? 12 : 14, bold: c > 1,
            color: { argb: c === 1 ? "FFFFFFFF" : "FFFE9933" }
          };
          cell.alignment = { horizontal: c === 1 ? "left" : "center", vertical: "middle" };
        }
        continue;
      }

      if (!row.height || row.height < 15) row.height = 30;

      for (let c = 1; c <= maxCol; c++) {
        const cell = row.getCell(c);
        const val = cell.value;
        const isFormula = val && typeof val === "object" && val.formula;

        if (c === 1 && isNumbered) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isOdd ? "FFFE9933" : "FF444444" } };
          cell.font = { name: "Aptos", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (isFormula) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFDE7" } };
          cell.font = { name: "Aptos", size: 9, color: { argb: "FF222222" } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else if (c === 2 && isNumbered) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowBg } };
          cell.font = { name: "Aptos", size: 9, bold: true, color: { argb: "FF222222" } };
          cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        } else {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowBg } };
          cell.font = { name: "Aptos", size: 9, color: { argb: "FF222222" } };
          cell.alignment = { horizontal: "left", vertical: isNumbered ? "top" : "middle", wrapText: true };
        }
      }
    }
  }

  // Sla op als buffer
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

// ── CLAUDE API ──
async function callClaude(apiKey, systemPrompt, text) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 4096,
      system: systemPrompt,
      messages: [{ role: "user", content: `Verwerk het volgende document:\n\n${text}` }],
    }),
  });
  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `API-fout (${response.status})`);
  }
  const data = await response.json();
  return data.content[0].text;
}

// ── HOOFDFUNCTIE ──
exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: JSON.stringify({ error: "API-key niet geconfigureerd op de server." }) };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Ongeldig verzoek." }) };
  }

  const { fileData, fileType, fileName, options } = body;
  if (!fileData) {
    return { statusCode: 400, body: JSON.stringify({ error: "Geen bestandsdata ontvangen." }) };
  }

  const inputBuffer = Buffer.from(fileData, "base64");
  const baseName = (fileName || "document").replace(/\.[^.]+$/, "");

  try {
    let fileBuffer, mimeType, outputFileName;

    if (fileType === "xlsx") {
      // ── EXCEL: Huisstijl toepassen ──
      fileBuffer = await applyExcelStyle(inputBuffer);
      mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      outputFileName = baseName + "_SB.xlsx";

    } else {
      // ── OVERIGE FORMATEN: Tekst verbeteren via Claude ──
      let text = "";

      if (fileType === "docx") {
        const mammoth = require("mammoth");
        const result = await mammoth.extractRawText({ buffer: inputBuffer });
        text = result.value;
      } else if (fileType === "txt") {
        text = inputBuffer.toString("utf-8");
      } else if (fileType === "pptx") {
        const zip = await JSZip.loadAsync(inputBuffer);
        const slideFiles = Object.keys(zip.files)
          .filter(n => n.match(/ppt\/slides\/slide\d+\.xml/))
          .sort();
        for (const slideName of slideFiles) {
          const xml = await zip.files[slideName].async("string");
          const matches = xml.match(/<a:t[^>]*>([^<]+)<\/a:t>/g) || [];
          const slideNum = slideName.match(/slide(\d+)/)?.[1];
          text += `=== Slide ${slideNum} ===\n${matches.map(m => m.replace(/<[^>]+>/g, "")).join(" ")}\n\n`;
        }
      }

      // System prompt
      let systemPrompt = "Je bent een document-assistent voor SB Procesmanagement.\n\n";
      if (options?.style) {
        systemPrompt += "Pas een professionele, zakelijke toon toe. Gebruik actieve zinnen, heldere kopjes, geen wollige taal.\n\n";
      }
      if (options?.improve) {
        systemPrompt += "Verbeter grammatica, spelling en leesbaarheid.\n\n";
      }
      if (fileType === "pptx") {
        systemPrompt += "Geef slides terug als:\n=== Slide N ===\nTitel: ...\nInhoud: ...\n\nAlleen slides, geen uitleg.";
      } else {
        systemPrompt += "Geef ALLEEN de verbeterde tekst terug, geen uitleg.";
      }

      const improved = await callClaude(apiKey, systemPrompt, text);

      if (fileType === "docx") {
        const lines = improved.split("\n").filter(l => l.trim());
        const paragraphs = lines.map(line => {
          if (line.startsWith("# ")) return new Paragraph({ text: line.replace(/^# /, ""), heading: HeadingLevel.HEADING_1 });
          if (line.startsWith("## ")) return new Paragraph({ text: line.replace(/^## /, ""), heading: HeadingLevel.HEADING_2 });
          return new Paragraph({ children: [new TextRun({ text: line, size: 24, font: "Aptos" })] });
        });
        const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
        fileBuffer = await Packer.toBuffer(doc);
        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        outputFileName = baseName + "_SB.docx";

      } else if (fileType === "pptx") {
        const pptx = new PptxGenJS();
        const blocks = improved.split(/=== Slide \d+ ===/);
        for (const block of blocks) {
          if (!block.trim()) continue;
          const title = block.match(/Titel:\s*(.+)/)?.[1]?.trim() || "";
          const content = block.match(/Inhoud:\s*([\s\S]+?)(?=Titel:|$)/)?.[1]?.trim() || block.trim();
          const slide = pptx.addSlide();
          slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: "100%", h: 1.2, fill: { color: "E8822A" } });
          slide.addText("SB PROCESMANAGEMENT", { x: 0.3, y: 0.15, w: 6, h: 0.5, fontSize: 14, bold: true, color: "FFFFFF", fontFace: "Calibri" });
          if (title) slide.addText(title, { x: 0.5, y: 1.4, w: 12.3, h: 0.8, fontSize: 24, bold: true, color: "222222" });
          if (content) slide.addText(content, { x: 0.5, y: 2.4, w: 12.3, h: 4.5, fontSize: 16, color: "444444", wrap: true });
        }
        fileBuffer = await pptx.write({ outputType: "nodebuffer" });
        mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
        outputFileName = baseName + "_SB.pptx";

      } else if (fileType === "pdf") {
        const pdfDoc = await PDFDocument.create();
        const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
        const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        let page = pdfDoc.addPage([595, 842]);
        const { width, height } = page.getSize();
        let y = height - 80;
        const margin = 50;
        page.drawRectangle({ x: 0, y: height - 50, width, height: 50, color: rgb(0.91, 0.51, 0.17) });
        page.drawText("SB PROCESMANAGEMENT", { x: margin, y: height - 34, size: 16, font: boldFont, color: rgb(1, 1, 1) });
        for (const line of improved.split("\n")) {
          if (y < 60) { page = pdfDoc.addPage([595, 842]); y = height - 80; }
          if (!line.trim()) { y -= 9; continue; }
          const isH = line.startsWith("# ") || line.startsWith("## ");
          const text = line.replace(/^#{1,3} /, "");
          const words = text.split(" ");
          let cur = "";
          for (const word of words) {
            const test = cur ? cur + " " + word : word;
            if ((isH ? boldFont : font).widthOfTextAtSize(test, isH ? 14 : 11) > width - margin * 2 && cur) {
              page.drawText(cur, { x: margin, y, size: isH ? 14 : 11, font: isH ? boldFont : font, color: isH ? rgb(0.91, 0.51, 0.17) : rgb(0.13, 0.13, 0.13) });
              y -= 18; cur = word;
            } else { cur = test; }
          }
          if (cur) { page.drawText(cur, { x: margin, y, size: isH ? 14 : 11, font: isH ? boldFont : font, color: isH ? rgb(0.91, 0.51, 0.17) : rgb(0.13, 0.13, 0.13) }); y -= 18; }
        }
        fileBuffer = await pdfDoc.save();
        mimeType = "application/pdf";
        outputFileName = baseName + "_SB.pdf";

      } else {
        fileBuffer = Buffer.from(improved, "utf-8");
        mimeType = "text/plain";
        outputFileName = baseName + "_SB.txt";
      }
    }

    return {
      statusCode: 200,
      headers: {
        "Content-Type": mimeType,
        "Content-Disposition": `attachment; filename="${outputFileName}"`,
        "X-File-Name": outputFileName,
      },
      body: fileBuffer.toString("base64"),
      isBase64Encoded: true,
    };

  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: "Fout: " + err.message }) };
  }
};
