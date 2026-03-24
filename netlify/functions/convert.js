const fetch = require("node-fetch");
const mammoth = require("mammoth");
const XLSX = require("xlsx");
const { Document, Paragraph, TextRun, HeadingLevel, Packer } = require("docx");
const PptxGenJS = require("pptxgenjs");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");

exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "API-key niet geconfigureerd op de server." }),
    };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Ongeldig verzoek." }) };
  }

  const { text, fileType, options, fileName } = body;
  if (!text) {
    return { statusCode: 400, body: JSON.stringify({ error: "Geen tekst ontvangen." }) };
  }

  // Bouw system prompt
  let systemPrompt =
    "Je bent een document-assistent voor SB Procesmanagement, een freelance interim procesmanager gespecialiseerd in logistiek en warehouse-optimalisatie.\n\n";

  if (options?.style) {
    systemPrompt += `Pas de volgende stijlregels toe:
- Professionele, zakelijke toon - direct en helder
- Gebruik actieve zinsconstructies (vermijd lijdende vorm)
- Structureer het document met duidelijke kopjes waar passend
- Verwijder wollige taal en overbodige herhalingen
- Houd de inhoud bondig maar volledig\n\n`;
  }

  if (options?.improve) {
    systemPrompt += `Verbeter ook:
- Grammatica en spelling
- Leesbaarheid en doorstroming van de tekst
- Logische volgorde van de inhoud\n\n`;
  }

  if (fileType === "xlsx") {
    systemPrompt += `Dit is een Excel-document. Geef de verbeterde inhoud terug als CSV met puntkomma als scheidingsteken.
Bewaar de tabelstructuur zo goed mogelijk. Gebruik === Tabblad: naam === om tabbladen te scheiden.
Geef ALLEEN de CSV-data terug, geen uitleg.`;
  } else if (fileType === "pptx") {
    systemPrompt += `Dit is een PowerPoint-presentatie. Geef de verbeterde inhoud terug in dit exacte formaat:
=== Slide 1 ===
Titel: [titel van de slide]
Inhoud: [inhoud van de slide]

=== Slide 2 ===
Titel: [titel]
Inhoud: [inhoud]

Geef ALLEEN de slides terug in dit formaat, geen uitleg.`;
  } else {
    systemPrompt += "Geef ALLEEN de verbeterde tekst terug. Geen uitleg, geen commentaar - alleen de inhoud.";
  }

  // Claude API aanroepen
  let improvedText;
  try {
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
      return {
        statusCode: response.status,
        body: JSON.stringify({ error: err.error?.message || `API-fout (${response.status})` }),
      };
    }

    const data = await response.json();
    improvedText = data.content[0].text;
  } catch (err) {
    return { statusCode: 500, body: JSON.stringify({ error: "API-aanroep mislukt: " + err.message }) };
  }

  // Bestand opbouwen in het juiste formaat
  try {
    let fileBuffer;
    let mimeType;
    let outputFileName;
    const baseName = (fileName || "document").replace(/\.[^.]+$/, "");

    if (fileType === "docx") {
      const lines = improvedText.split("\n").filter(l => l.trim());
      const paragraphs = lines.map(line => {
        if (line.startsWith("# ")) {
          return new Paragraph({ text: line.replace(/^# /, ""), heading: HeadingLevel.HEADING_1 });
        } else if (line.startsWith("## ")) {
          return new Paragraph({ text: line.replace(/^## /, ""), heading: HeadingLevel.HEADING_2 });
        } else if (line.startsWith("### ")) {
          return new Paragraph({ text: line.replace(/^### /, ""), heading: HeadingLevel.HEADING_3 });
        } else {
          return new Paragraph({ children: [new TextRun({ text: line, size: 24, font: "Aptos" })] });
        }
      });

      const doc = new Document({
        sections: [{ properties: {}, children: paragraphs }],
      });

      fileBuffer = await Packer.toBuffer(doc);
      mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      outputFileName = baseName + "_SB.docx";

    } else if (fileType === "xlsx") {
      const workbook = XLSX.utils.book_new();
      const sections = improvedText.split(/=== Tabblad: (.+?) ===/);

      if (sections.length > 1) {
        for (let i = 1; i < sections.length; i += 2) {
          const sheetName = sections[i].trim().substring(0, 31);
          const csvData = sections[i + 1]?.trim() || "";
          const rows = csvData.split("\n").map(row => row.split(";"));
          const worksheet = XLSX.utils.aoa_to_sheet(rows);
          XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }
      } else {
        const rows = improvedText.split("\n").map(row => row.split(";"));
        const worksheet = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Blad1");
      }

      fileBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
      mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      outputFileName = baseName + "_SB.xlsx";

    } else if (fileType === "pptx") {
      const pptx = new PptxGenJS();
      const slideBlocks = improvedText.split(/=== Slide \d+ ===/);

      for (const block of slideBlocks) {
        if (!block.trim()) continue;
        const titleMatch = block.match(/Titel:\s*(.+)/);
        const contentMatch = block.match(/Inhoud:\s*([\s\S]+?)(?=Titel:|$)/);
        const title = titleMatch ? titleMatch[1].trim() : "";
        const content = contentMatch ? contentMatch[1].trim() : block.trim();
        const slide = pptx.addSlide();

        slide.addShape(pptx.ShapeType.rect, {
          x: 0, y: 0, w: "100%", h: 1.2, fill: { color: "E8822A" },
        });
        slide.addText("SB PROCESMANAGEMENT", {
          x: 0.3, y: 0.15, w: 6, h: 0.5,
          fontSize: 14, bold: true, color: "FFFFFF", fontFace: "Calibri",
        });
        if (title) {
          slide.addText(title, {
            x: 0.5, y: 1.4, w: 12.3, h: 0.8,
            fontSize: 24, bold: true, color: "222222", fontFace: "Calibri",
          });
        }
        if (content) {
          slide.addText(content, {
            x: 0.5, y: 2.4, w: 12.3, h: 4.5,
            fontSize: 16, color: "444444", fontFace: "Calibri", valign: "top", wrap: true,
          });
        }
      }

      fileBuffer = await pptx.write({ outputType: "nodebuffer" });
      mimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
      outputFileName = baseName + "_SB.pptx";

    } else if (fileType === "pdf") {
      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const lines = improvedText.split("\n");
      let page = pdfDoc.addPage([595, 842]);
      const { width, height } = page.getSize();
      let y = height - 80;
      const margin = 50;
      const maxWidth = width - margin * 2;
      const lineHeight = 18;

      page.drawRectangle({ x: 0, y: height - 50, width, height: 50, color: rgb(0.91, 0.51, 0.17) });
      page.drawText("SB PROCESMANAGEMENT", {
        x: margin, y: height - 34, size: 16, font: boldFont, color: rgb(1, 1, 1),
      });

      for (const line of lines) {
        if (y < 60) {
          page = pdfDoc.addPage([595, 842]);
          y = height - 80;
          page.drawRectangle({ x: 0, y: height - 50, width, height: 50, color: rgb(0.91, 0.51, 0.17) });
          page.drawText("SB PROCESMANAGEMENT", {
            x: margin, y: height - 34, size: 16, font: boldFont, color: rgb(1, 1, 1),
          });
        }
        if (!line.trim()) { y -= lineHeight / 2; continue; }
        const isHeading = line.startsWith("# ") || line.startsWith("## ");
        const text = line.replace(/^#{1,3} /, "");
        const usedFont = isHeading ? boldFont : font;
        const fontSize = isHeading ? 14 : 11;
        const color = isHeading ? rgb(0.91, 0.51, 0.17) : rgb(0.13, 0.13, 0.13);
        const words = text.split(" ");
        let currentLine = "";
        for (const word of words) {
          const testLine = currentLine ? currentLine + " " + word : word;
          const testWidth = usedFont.widthOfTextAtSize(testLine, fontSize);
          if (testWidth > maxWidth && currentLine) {
            page.drawText(currentLine, { x: margin, y, size: fontSize, font: usedFont, color });
            y -= lineHeight;
            currentLine = word;
          } else {
            currentLine = testLine;
          }
        }
        if (currentLine) {
          page.drawText(currentLine, { x: margin, y, size: fontSize, font: usedFont, color });
          y -= lineHeight;
        }
        if (isHeading) y -= 4;
      }

      fileBuffer = await pdfDoc.save();
      mimeType = "application/pdf";
      outputFileName = baseName + "_SB.pdf";

    } else {
      fileBuffer = Buffer.from(improvedText, "utf-8");
      mimeType = "text/plain";
      outputFileName = baseName + "_SB.txt";
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
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "Fout bij het aanmaken van het bestand: " + err.message }),
    };
  }
};
