const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ═══════════════════════════════════════════════
// COULEURS CREATICAL - Style Reporting Premium
// ═══════════════════════════════════════════════
const C = {
  bgDark: "0D0221",
  bgMid: "1A0A2E",
  bgCard: "150930",
  violet: "7C3AED",
  violetLight: "A855F7",
  violetPale: "C4B5FD",
  violetMuted: "6D28D9",
  white: "FFFFFF",
  offWhite: "F0E6FF",
  gray: "9CA3AF",
  grayLight: "D1D5DB",
  accent: "D946EF",
  accentPink: "EC4899",
  success: "10B981",
  warning: "F59E0B",
};

// Slide dimensions: 10" x 5.625"
const SW = 10;
const SH = 5.625;
const MARGIN = 0.5;

// ═══════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════

function addBg(pres, sl) {
  sl.background = { color: C.bgDark };
  // Violet glow on right side
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 4.5, y: 0, w: 5.5, h: SH,
    fill: { color: C.violet, transparency: 88 },
  });
  // Subtle second glow top-right
  sl.addShape(pres.shapes.OVAL, {
    x: 7, y: -1, w: 4, h: 4,
    fill: { color: C.accent, transparency: 93 },
  });
}

function addContentFrame(pres, sl) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: 0.3, w: 9.4, h: 5.025,
    line: { color: C.gray, width: 0.5 },
    fill: { color: "000000", transparency: 95 },
    rectRadius: 0.15,
  });
}

function addFooterLogo(sl) {
  // Simple "C" logo circle
  sl.addShape("ellipse", {
    x: 4.7, y: 5.05, w: 0.35, h: 0.35,
    fill: { color: C.violet, transparency: 40 },
    line: { color: C.violetLight, width: 0.5 },
  });
  sl.addText("C", {
    x: 4.7, y: 5.05, w: 0.35, h: 0.35,
    fontSize: 11, color: C.violetPale, fontFace: "Arial",
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

function pill(sl, x, y, w, h, text, opts = {}) {
  const bgColor = opts.bg || C.white;
  const textColor = opts.color || C.bgDark;
  const fontSize = opts.fontSize || 12;
  sl.addShape("roundedRectangle", {
    x, y, w, h,
    fill: { color: bgColor },
    rectRadius: h / 2,
    shadow: opts.shadow ? { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.2 } : undefined,
  });
  sl.addText(text, {
    x, y, w, h,
    fontSize, color: textColor, fontFace: opts.fontFace || "Arial",
    bold: opts.bold !== false, align: "center", valign: "middle", margin: 0,
    italic: opts.italic || false,
  });
}

function sectionNumber(pres, sl, x, y, num) {
  sl.addShape(pres.shapes.OVAL, {
    x, y, w: 0.5, h: 0.5,
    fill: { color: C.violet },
  });
  sl.addText(String(num), {
    x, y, w: 0.5, h: 0.5,
    fontSize: 16, color: C.white, fontFace: "Arial",
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

function card(pres, sl, x, y, w, h, opts = {}) {
  sl.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.bgCard, transparency: 10 },
    line: { color: opts.borderColor || C.violet, width: 0.75 },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 },
  });
  // Left accent bar
  if (opts.accent !== false) {
    sl.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.05, h,
      fill: { color: opts.accentColor || C.violetLight },
    });
  }
}

// ═══════════════════════════════════════════════
// HEALTH CHECK
// ═══════════════════════════════════════════════
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "CREATICAL PPTX Generator v2" });
});

// ═══════════════════════════════════════════════
// GENERATE PPTX
// ═══════════════════════════════════════════════
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
      addBg(pres, sl);

      switch (s.type) {

        // ─────────────────────────────────────
        // SLIDE TITRE
        // ─────────────────────────────────────
        case "title": {
          addContentFrame(pres, sl);
          // Subtitle (mois)
          sl.addText(s.subtitle || "", {
            x: 1, y: 0.8, w: 5, h: 0.5,
            fontSize: 16, color: C.offWhite, fontFace: "Georgia", italic: true,
          });
          // Big title
          sl.addText((s.title || "STRATEGIE").toUpperCase(), {
            x: 1, y: 1.4, w: 6, h: 1.0,
            fontSize: 42, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          // Cursive "de contenu"
          sl.addText("de contenu", {
            x: 1, y: 2.3, w: 4, h: 0.6,
            fontSize: 28, color: C.violetPale, fontFace: "Georgia", italic: true,
          });
          // Client pill
          pill(sl, 1, 3.2, 4.5, 0.45, s.client || "", { shadow: true });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // OBJECTIFS
        // ─────────────────────────────────────
        case "objectifs": {
          addContentFrame(pres, sl);
          // Big "OBJECTIFS 2026" on left
          sl.addText("OBJECTIFS", {
            x: 0.8, y: 0.8, w: 4, h: 0.8,
            fontSize: 38, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          sl.addText(s.year || "2026", {
            x: 0.8, y: 1.6, w: 4, h: 1.2,
            fontSize: 72, color: C.white, fontFace: "Arial", bold: true,
          });
          // Client name pill
          pill(sl, 0.8, 3.0, 3, 0.4, s.clientName || meta.client || "", { fontSize: 11, shadow: true });

          // Objectives on right with arrows
          const items = s.items || [];
          items.forEach((item, i) => {
            const yPos = 0.9 + i * 1.0;
            // Arrow line from center to text
            sl.addShape(pres.shapes.LINE, {
              x: 5.2, y: yPos + 0.25, w: 0.6, h: 0,
              line: { color: C.white, width: 1 },
            });
            // Triangle arrowhead
            sl.addText(item, {
              x: 6, y: yPos, w: 3.5, h: 0.8,
              fontSize: 12, color: C.white, fontFace: "Arial", valign: "middle",
            });
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // SECTION DIVIDER
        // ─────────────────────────────────────
        case "section": {
          addContentFrame(pres, sl);
          sl.addText((s.title || "").toUpperCase(), {
            x: 1, y: 1.2, w: 8, h: 1.2,
            fontSize: 44, color: C.white, fontFace: "Arial", bold: true,
            align: "center", charSpacing: 3,
          });
          sl.addText(s.subtitle || "", {
            x: 1, y: 2.6, w: 8, h: 0.8,
            fontSize: 20, color: C.violetPale, fontFace: "Georgia",
            italic: true, align: "center",
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // BRAND ARCHETYPE
        // ─────────────────────────────────────
        case "brand_archetype": {
          addContentFrame(pres, sl);
          sl.addText("BRAND ARCHETYPE", {
            x: 0.8, y: 0.5, w: 8, h: 0.6,
            fontSize: 28, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          // Subtitle pill
          pill(sl, 0.8, 1.2, 4.5, 0.4, s.subtitle || "Qui est le client ?", { shadow: true });
          // Description in card
          card(pres, sl, 0.8, 1.9, 8.4, 2.8);
          sl.addText(s.description || "", {
            x: 1.1, y: 2.0, w: 7.9, h: 2.6,
            fontSize: 12, color: C.grayLight, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.4, margin: [5, 10, 5, 10],
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // PERSONA / CIBLE
        // ─────────────────────────────────────
        case "persona": {
          addContentFrame(pres, sl);
          sl.addText("PERSONA CIBLE", {
            x: 0.8, y: 0.5, w: 5, h: 0.6,
            fontSize: 28, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          // Persona name pill
          pill(sl, 0.8, 1.15, 4, 0.4, s.name || "Persona", { shadow: true, bg: C.violet, color: C.white });

          // Left card - Demographics
          card(pres, sl, 0.8, 1.8, 4, 3.0);
          sl.addText("Profil", {
            x: 1.0, y: 1.85, w: 3.6, h: 0.35,
            fontSize: 14, color: C.violetPale, fontFace: "Arial", bold: true, margin: 0,
          });
          sl.addText(s.profil || "", {
            x: 1.0, y: 2.2, w: 3.6, h: 1.2,
            fontSize: 10, color: C.grayLight, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.3, margin: 0,
          });
          sl.addText("Core Problem", {
            x: 1.0, y: 3.4, w: 3.6, h: 0.3,
            fontSize: 12, color: C.accentPink, fontFace: "Arial", bold: true, margin: 0,
          });
          sl.addText(s.core_problem || "", {
            x: 1.0, y: 3.7, w: 3.6, h: 0.6,
            fontSize: 10, color: C.grayLight, fontFace: "Arial", valign: "top", margin: 0,
          });

          // Right card - Emotions & Peurs
          card(pres, sl, 5.2, 1.8, 4.2, 3.0, { accentColor: C.accentPink });
          sl.addText("Emotions & Peurs", {
            x: 5.5, y: 1.85, w: 3.7, h: 0.35,
            fontSize: 14, color: C.violetPale, fontFace: "Arial", bold: true, margin: 0,
          });
          const emotions = s.emotions || [];
          emotions.forEach((em, i) => {
            sl.addText(em, {
              x: 5.5, y: 2.25 + i * 0.35, w: 3.7, h: 0.35,
              fontSize: 9, color: C.grayLight, fontFace: "Arial", valign: "top",
              bullet: true, margin: 0,
            });
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // POSITIONNEMENT
        // ─────────────────────────────────────
        case "positionnement": {
          addContentFrame(pres, sl);
          sl.addText("POSITIONNEMENT", {
            x: 0.8, y: 0.5, w: 8, h: 0.6,
            fontSize: 28, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          // Tagline in violet pill
          pill(sl, 0.8, 1.3, 8.4, 0.55, s.tagline || "", {
            bg: C.violet, color: C.white, fontSize: 14, italic: true, fontFace: "Georgia",
          });
          // Elements in cards
          const elems = s.elements || [];
          elems.forEach((el, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const xPos = 0.8 + col * 4.4;
            const yPos = 2.2 + row * 1.3;
            card(pres, sl, xPos, yPos, 4.0, 1.1);
            sectionNumber(pres, sl, xPos + 0.15, yPos + 0.3, i + 1);
            sl.addText(el, {
              x: xPos + 0.8, y: yPos + 0.15, w: 3.1, h: 0.8,
              fontSize: 11, color: C.grayLight, fontFace: "Arial", valign: "middle",
              lineSpacingMultiple: 1.2,
            });
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // PILIERS EDITORIAUX
        // ─────────────────────────────────────
        case "piliers": {
          addContentFrame(pres, sl);
          sl.addText("PILIERS EDITORIAUX", {
            x: 0.8, y: 0.5, w: 8, h: 0.6,
            fontSize: 28, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          const piliers = s.piliers || [];
          const pilierColors = [C.violet, C.violetLight, C.accent, C.accentPink];
          piliers.forEach((p, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const xPos = 0.8 + col * 4.4;
            const yPos = 1.3 + row * 2.0;
            card(pres, sl, xPos, yPos, 4.0, 1.7, { accentColor: pilierColors[i] || C.violet });
            // Pilier number + name
            sectionNumber(pres, sl, xPos + 0.2, yPos + 0.15, i + 1);
            sl.addText(p.nom || "", {
              x: xPos + 0.85, y: yPos + 0.1, w: 2.9, h: 0.4,
              fontSize: 13, color: C.white, fontFace: "Arial", bold: true, margin: 0,
            });
            sl.addText(p.description || "", {
              x: xPos + 0.2, y: yPos + 0.6, w: 3.6, h: 1.0,
              fontSize: 10, color: C.grayLight, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.3, margin: 0,
            });
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // FUNNEL TOFU/MOFU/BOFU
        // ─────────────────────────────────────
        case "funnel": {
          addContentFrame(pres, sl);
          sl.addText("FUNNEL D'ACQUISITION", {
            x: 0.8, y: 0.4, w: 8, h: 0.5,
            fontSize: 24, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          const stages = [
            { data: s.tofu, color: C.violet, label: "TOFU" },
            { data: s.mofu, color: C.violetLight, label: "MOFU" },
            { data: s.bofu, color: C.accent, label: "BOFU" },
          ];
          stages.forEach((stage, i) => {
            if (!stage.data) return;
            const xPos = 0.5 + i * 3.15;
            // Column card
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: 1.05, w: 2.95, h: 4.2,
              fill: { color: C.bgCard, transparency: 10 },
              line: { color: stage.color, width: 0.75 },
            });
            // Header bar
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xPos, y: 1.05, w: 2.95, h: 0.5,
              fill: { color: stage.color },
            });
            sl.addText(stage.label, {
              x: xPos, y: 1.05, w: 2.95, h: 0.5,
              fontSize: 15, color: C.white, fontFace: "Arial",
              bold: true, align: "center", valign: "middle", margin: 0,
            });
            // Description
            sl.addText(stage.data.description || "", {
              x: xPos + 0.15, y: 1.6, w: 2.65, h: 0.5,
              fontSize: 9, color: C.violetPale, fontFace: "Arial",
              italic: true, valign: "top", margin: 0,
            });
            // Video ideas
            const videos = stage.data.videos || [];
            videos.forEach((v, j) => {
              sl.addText(`${j + 1}. ${v}`, {
                x: xPos + 0.15, y: 2.15 + j * 0.4, w: 2.65, h: 0.38,
                fontSize: 9, color: C.grayLight, fontFace: "Arial",
                valign: "top", margin: 0,
              });
            });
          });
          break;
        }

        // ─────────────────────────────────────
        // PLAN D'ACTION 90 JOURS
        // ─────────────────────────────────────
        case "plan_action": {
          addContentFrame(pres, sl);
          sl.addText("PLAN D'ACTION 90 JOURS", {
            x: 0.8, y: 0.4, w: 8, h: 0.6,
            fontSize: 24, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          // Timeline line
          sl.addShape(pres.shapes.LINE, {
            x: 0.8, y: 2.8, w: 8.4, h: 0,
            line: { color: C.violet, width: 2 },
          });
          const phases = s.phases || [];
          phases.forEach((phase, i) => {
            const xPos = 0.8 + i * 3.0;
            // Phase number circle on timeline
            sectionNumber(pres, sl, xPos + 1.2, 2.55, i + 1);
            // Phase title above
            pill(sl, xPos + 0.2, 1.4, 2.5, 0.35, phase.titre || "", {
              bg: C.violet, color: C.white, fontSize: 11,
            });
            sl.addText(phase.periode || "", {
              x: xPos + 0.2, y: 1.85, w: 2.5, h: 0.3,
              fontSize: 9, color: C.gray, fontFace: "Arial", align: "center",
            });
            // Actions below timeline
            const actions = phase.actions || [];
            actions.forEach((a, j) => {
              sl.addText(a, {
                x: xPos + 0.15, y: 3.2 + j * 0.4, w: 2.7, h: 0.38,
                fontSize: 9, color: C.grayLight, fontFace: "Arial",
                bullet: true, valign: "top", margin: 0,
              });
            });
          });
          break;
        }

        // ─────────────────────────────────────
        // RESUME
        // ─────────────────────────────────────
        case "resume": {
          addContentFrame(pres, sl);
          // Big title on left
          sl.addText("RESUME", {
            x: 0.8, y: 1.0, w: 4, h: 0.8,
            fontSize: 44, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          sl.addText("du mois", {
            x: 0.8, y: 1.8, w: 4, h: 0.6,
            fontSize: 26, color: C.violetPale, fontFace: "Georgia", italic: true,
          });
          // Points on right
          const points = s.points || [];
          points.forEach((pt, i) => {
            const yPos = 0.8 + i * 1.4;
            // White pill title
            pill(sl, 5.2, yPos, 4.0, 0.4, pt.titre || "", { shadow: true, fontSize: 11 });
            sl.addText(pt.description || "", {
              x: 5.2, y: yPos + 0.5, w: 4.0, h: 0.7,
              fontSize: 10, color: C.grayLight, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.3,
            });
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // CONTACT
        // ─────────────────────────────────────
        case "contact": {
          addContentFrame(pres, sl);
          // Logo area on right
          sl.addText("C", {
            x: 6.5, y: 1.2, w: 1.5, h: 1.5,
            fontSize: 60, color: C.violetPale, fontFace: "Arial", bold: true,
            align: "center", valign: "middle",
          });
          sl.addText(s.name || "CREATICAL", {
            x: 5.5, y: 2.8, w: 3.5, h: 0.5,
            fontSize: 20, color: C.white, fontFace: "Arial", bold: true,
          });
          sl.addText(s.role || "Social Media Manager", {
            x: 5.5, y: 3.3, w: 3.5, h: 0.5,
            fontSize: 16, color: C.violetPale, fontFace: "Georgia", italic: true,
          });
          sl.addText(s.email || "contact@creatical.be", {
            x: 5.5, y: 3.8, w: 3.5, h: 0.4,
            fontSize: 13, color: C.gray, fontFace: "Georgia", italic: true,
          });
          addFooterLogo(sl);
          break;
        }

        // ─────────────────────────────────────
        // GENERIC / FALLBACK
        // ─────────────────────────────────────
        default: {
          addContentFrame(pres, sl);
          sl.addText((s.title || "").toUpperCase(), {
            x: 0.8, y: 0.5, w: 8, h: 0.6,
            fontSize: 24, color: C.white, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          if (s.content) {
            card(pres, sl, 0.8, 1.3, 8.4, 3.8);
            sl.addText(s.content, {
              x: 1.1, y: 1.4, w: 7.9, h: 3.5,
              fontSize: 11, color: C.grayLight, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.3,
            });
          }
          addFooterLogo(sl);
          break;
        }
      }
    }

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
  console.log(`CREATICAL PPTX Server v2 running on port ${PORT}`);
});
