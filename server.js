const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();
app.use(express.json({ limit: "10mb" }));

// --- COULEURS CREATICAL ---
const C = {
  bgDark: "0D0221",
  bgMid: "1A0A2E",
  violet: "7C3AED",
  violetLight: "A855F7",
  violetPale: "C4B5FD",
  white: "FFFFFF",
  gray: "A0A0B0",
  grayLight: "E0E0E8",
  accent: "D946EF",
};

// --- HELPERS ---
function addFooter(slide) {
  slide.addText("CREATICAL", {
    x: 4, y: 5.1, w: 2, h: 0.35,
    fontSize: 9, color: C.gray, align: "center", fontFace: "Arial",
    italic: true,
  });
}

function addSlideBackground(pres, slide) {
  slide.background = { color: C.bgDark };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 0, w: 5, h: 5.625,
    fill: { color: C.violet, transparency: 85 },
  });
}

function makeShadow() {
  return { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 };
}

// --- HEALTH CHECK ---
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "CREATICAL PPTX Generator" });
});

// --- GENERATE PPTX ---
app.post("/generate", async (req, res) => {
  try {
    const { slidesData } = req.body;
    if (!slidesData || !slidesData.slides) {
      return res.status(400).json({ error: "slidesData.slides is required" });
    }

    const slides = slidesData.slides;
    const meta = slidesData.metadata || {};

    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "CREATICAL";
    pres.title = meta.title || "Stratégie de Contenu";

    for (const s of slides) {
      const sl = pres.addSlide();
      addSlideBackground(pres, sl);

      switch (s.type) {
        case "title": {
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.4, y: 0.4, w: 9.2, h: 4.8,
            line: { color: C.gray, width: 1 },
            fill: { color: C.bgMid, transparency: 40 },
            rectRadius: 0.15,
          });
          sl.addText(s.subtitle || "", {
            x: 1, y: 1, w: 5, h: 0.5,
            fontSize: 16, color: C.white, fontFace: "Georgia", italic: true,
          });
          sl.addText(s.title || "STRATEGIE", {
            x: 1, y: 1.6, w: 6, h: 1.2,
            fontSize: 44, color: C.white, fontFace: "Arial Black", bold: true,
          });
          sl.addText(s.client || "", {
            x: 1, y: 3.3, w: 5, h: 0.5,
            fontSize: 14, color: C.violetPale, fontFace: "Arial",
          });
          addFooter(sl);
          break;
        }

        case "objectifs": {
          sl.addText(s.title || "OBJECTIFS", {
            x: 0.8, y: 0.4, w: 5, h: 0.8,
            fontSize: 36, color: C.white, fontFace: "Arial Black", bold: true,
          });
          const items = s.items || [];
          items.forEach((item, i) => {
            const yPos = 1.6 + i * 0.9;
            sl.addShape(pres.shapes.OVAL, {
              x: 5.5, y: yPos + 0.05, w: 0.25, h: 0.25,
              fill: { color: C.violet },
            });
            sl.addText(item, {
              x: 6, y: yPos, w: 3.5, h: 0.7,
              fontSize: 13, color: C.white, fontFace: "Arial", valign: "top",
            });
          });
          addFooter(sl);
          break;
        }

        case "section": {
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.4, y: 0.4, w: 9.2, h: 4.8,
            line: { color: C.gray, width: 1 },
            fill: { color: C.bgMid, transparency: 40 },
            rectRadius: 0.15,
          });
          sl.addText(s.title || "", {
            x: 1, y: 1.5, w: 8, h: 1.5,
            fontSize: 40, color: C.white, fontFace: "Arial Black", bold: true, align: "center",
          });
          sl.addText(s.subtitle || "", {
            x: 1, y: 3, w: 8, h: 0.8,
            fontSize: 18, color: C.violetLight, fontFace: "Georgia", italic: true, align: "center",
          });
          addFooter(sl);
          break;
        }

        case "brand_archetype": {
          sl.addText(s.title || "BRAND ARCHETYPE", {
            x: 0.8, y: 0.3, w: 8, h: 0.7,
            fontSize: 28, color: C.white, fontFace: "Arial Black", bold: true,
          });
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.1, w: 4, h: 0.45,
            fill: { color: C.white }, rectRadius: 0.2,
          });
          sl.addText(s.subtitle || "", {
            x: 0.8, y: 1.1, w: 4, h: 0.45,
            fontSize: 13, color: C.bgDark, fontFace: "Arial", bold: true, align: "center", margin: 0,
          });
          sl.addText(s.description || "", {
            x: 0.8, y: 1.8, w: 8.4, h: 3.2,
            fontSize: 13, color: C.grayLight, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.3,
          });
          addFooter(sl);
          break;
        }

        case "positionnement": {
          sl.addText(s.title || "POSITIONNEMENT", {
            x: 0.8, y: 0.3, w: 8, h: 0.7,
            fontSize: 28, color: C.white, fontFace: "Arial Black", bold: true,
          });
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.2, w: 8.4, h: 0.6,
            fill: { color: C.violet }, rectRadius: 0.1,
          });
          sl.addText(s.tagline || "", {
            x: 0.8, y: 1.2, w: 8.4, h: 0.6,
            fontSize: 15, color: C.white, fontFace: "Georgia", italic: true, align: "center", margin: 0,
          });
          const elems = s.elements || [];
          elems.forEach((el, i) => {
            const yPos = 2.2 + i * 0.7;
            sl.addShape(pres.shapes.OVAL, {
              x: 1, y: yPos + 0.08, w: 0.2, h: 0.2,
              fill: { color: C.violetLight },
            });
            sl.addText(el, {
              x: 1.4, y: yPos, w: 7.5, h: 0.55,
              fontSize: 12, color: C.white, fontFace: "Arial",
            });
          });
          addFooter(sl);
          break;
        }

        case "piliers": {
          sl.addText(s.title || "PILIERS EDITORIAUX", {
            x: 0.8, y: 0.3, w: 8, h: 0.7,
            fontSize: 28, color: C.white, fontFace: "Arial Black", bold: true,
          });
          const piliers = s.piliers || [];
          piliers.forEach((p, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const xPos = 0.6 + col * 4.6;
            const yPos = 1.3 + row * 1.9;
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: yPos, w: 4.2, h: 1.6,
              fill: { color: C.bgMid, transparency: 20 },
              line: { color: C.violet, width: 1 },
            });
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: yPos, w: 0.06, h: 1.6,
              fill: { color: C.violetLight },
            });
            sl.addText(p.nom || "", {
              x: xPos + 0.2, y: yPos + 0.1, w: 3.8, h: 0.4,
              fontSize: 14, color: C.violetPale, fontFace: "Arial", bold: true, margin: 0,
            });
            sl.addText(p.description || "", {
              x: xPos + 0.2, y: yPos + 0.5, w: 3.8, h: 1,
              fontSize: 11, color: C.grayLight, fontFace: "Arial", valign: "top", margin: 0,
            });
          });
          addFooter(sl);
          break;
        }

        case "funnel": {
          sl.addText(s.title || "FUNNEL D'ACQUISITION", {
            x: 0.8, y: 0.3, w: 8, h: 0.7,
            fontSize: 28, color: C.white, fontFace: "Arial Black", bold: true,
          });
          const stages = [s.tofu, s.mofu, s.bofu];
          const stageColors = [C.violet, C.violetLight, C.accent];
          stages.forEach((stage, i) => {
            if (!stage) return;
            const xPos = 0.4 + i * 3.1;
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: 1.2, w: 2.9, h: 4,
              fill: { color: C.bgMid, transparency: 20 },
              line: { color: stageColors[i], width: 1 },
            });
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: 1.2, w: 2.9, h: 0.45,
              fill: { color: stageColors[i] },
            });
            sl.addText(stage.titre || "", {
              x: xPos, y: 1.2, w: 2.9, h: 0.45,
              fontSize: 14, color: C.white, fontFace: "Arial", bold: true, align: "center", margin: 0,
            });
            const videos = stage.videos || [];
            const videoText = videos.map((v, j) => ({
              text: `${j + 1}. ${v}`,
              options: { breakLine: true, fontSize: 10, color: C.grayLight, fontFace: "Arial" },
            }));
            if (videoText.length > 0) {
              sl.addText(videoText, {
                x: xPos + 0.15, y: 1.8, w: 2.6, h: 3.2, valign: "top",
              });
            }
          });
          break;
        }

        case "videos": {
          sl.addText(s.title || "12 VIDEOS", {
            x: 0.8, y: 0.3, w: 8, h: 0.6,
            fontSize: 24, color: C.white, fontFace: "Arial Black", bold: true,
          });
          const vids = (s.videos || []).slice(0, 12);
          vids.forEach((v, i) => {
            const col = i < 6 ? 0 : 1;
            const row = i < 6 ? i : i - 6;
            const xPos = 0.5 + col * 4.8;
            const yPos = 1.1 + row * 0.72;
            sl.addShape(pres.shapes.OVAL, {
              x: xPos, y: yPos + 0.05, w: 0.25, h: 0.25,
              fill: { color: C.violet },
            });
            sl.addText(String(v.numero || i + 1), {
              x: xPos, y: yPos + 0.05, w: 0.25, h: 0.25,
              fontSize: 9, color: C.white, fontFace: "Arial", bold: true, align: "center", margin: 0,
            });
            sl.addText(v.titre || "", {
              x: xPos + 0.35, y: yPos, w: 4.2, h: 0.35,
              fontSize: 10, color: C.white, fontFace: "Arial", bold: true, margin: 0,
            });
            sl.addText(`${v.funnel || ""} | ${v.pilier || ""}`, {
              x: xPos + 0.35, y: yPos + 0.3, w: 4.2, h: 0.3,
              fontSize: 8, color: C.gray, fontFace: "Arial", margin: 0,
            });
          });
          addFooter(sl);
          break;
        }

        case "plan_action": {
          sl.addText(s.title || "PLAN D'ACTION 90 JOURS", {
            x: 0.8, y: 0.3, w: 8, h: 0.6,
            fontSize: 24, color: C.white, fontFace: "Arial Black", bold: true,
          });
          sl.addShape(pres.shapes.LINE, {
            x: 0.5, y: 2.7, w: 9, h: 0,
            line: { color: C.violet, width: 2 },
          });
          const phases = s.phases || [];
          phases.forEach((phase, i) => {
            const xPos = 0.5 + i * 3.1;
            sl.addShape(pres.shapes.OVAL, {
              x: xPos + 1.1, y: 2.35, w: 0.7, h: 0.7,
              fill: { color: C.violet },
            });
            sl.addText(String(i + 1), {
              x: xPos + 1.1, y: 2.35, w: 0.7, h: 0.7,
              fontSize: 20, color: C.white, fontFace: "Arial", bold: true, align: "center", margin: 0,
            });
            sl.addText(phase.titre || "", {
              x: xPos, y: 1.2, w: 3, h: 0.4,
              fontSize: 13, color: C.violetPale, fontFace: "Arial", bold: true, align: "center",
            });
            sl.addText(phase.periode || "", {
              x: xPos, y: 1.55, w: 3, h: 0.3,
              fontSize: 10, color: C.gray, fontFace: "Arial", align: "center",
            });
            const actions = phase.actions || [];
            const actionText = actions.map((a) => ({
              text: a,
              options: { bullet: true, breakLine: true, fontSize: 10, color: C.grayLight, fontFace: "Arial" },
            }));
            if (actionText.length > 0) {
              sl.addText(actionText, {
                x: xPos + 0.2, y: 3.2, w: 2.6, h: 2, valign: "top",
              });
            }
          });
          break;
        }

        case "resume": {
          sl.addText(s.title || "RESUME", {
            x: 0.8, y: 0.3, w: 4, h: 1.5,
            fontSize: 40, color: C.white, fontFace: "Arial Black", bold: true,
          });
          const points = s.points || [];
          points.forEach((pt, i) => {
            const yPos = 1.2 + i * 1.3;
            sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
              x: 5, y: yPos, w: 4.3, h: 0.4,
              fill: { color: C.white }, rectRadius: 0.05,
            });
            sl.addText(pt.titre || "", {
              x: 5, y: yPos, w: 4.3, h: 0.4,
              fontSize: 12, color: C.bgDark, fontFace: "Arial", bold: true, align: "center", margin: 0,
            });
            sl.addText(pt.description || "", {
              x: 5, y: yPos + 0.45, w: 4.3, h: 0.7,
              fontSize: 10, color: C.grayLight, fontFace: "Arial", valign: "top",
            });
          });
          addFooter(sl);
          break;
        }

        case "contact": {
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.4, y: 0.4, w: 9.2, h: 4.8,
            line: { color: C.gray, width: 1 },
            fill: { color: C.bgMid, transparency: 40 },
            rectRadius: 0.15,
          });
          sl.addText(s.name || "CREATICAL", {
            x: 5, y: 1.8, w: 4, h: 0.6,
            fontSize: 22, color: C.white, fontFace: "Arial", bold: true,
          });
          sl.addText(s.role || "Social Media Manager", {
            x: 5, y: 2.4, w: 4, h: 0.5,
            fontSize: 16, color: C.violetPale, fontFace: "Georgia", italic: true,
          });
          sl.addText(s.email || "amine@creatical.be", {
            x: 5, y: 2.9, w: 4, h: 0.5,
            fontSize: 14, color: C.gray, fontFace: "Georgia", italic: true,
          });
          addFooter(sl);
          break;
        }

        default: {
          sl.addText(s.title || "", {
            x: 0.8, y: 0.3, w: 8, h: 0.7,
            fontSize: 28, color: C.white, fontFace: "Arial Black", bold: true,
          });
          if (s.content) {
            sl.addText(s.content, {
              x: 0.8, y: 1.3, w: 8.4, h: 3.8,
              fontSize: 12, color: C.grayLight, fontFace: "Arial", valign: "top",
            });
          }
          addFooter(sl);
          break;
        }
      }
    }

    // Generate as base64
    const base64 = await pres.write({ outputType: "base64" });

    res.json({
      success: true,
      pptxBase64: base64,
      filename: `${meta.title || "Presentation"}.pptx`,
    });
  } catch (error) {
    console.error("Error generating PPTX:", error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CREATICAL PPTX Server running on port ${PORT}`);
});
