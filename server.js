const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ═══════════════════════════════════════════════
// PALETTE CREATICAL - Vincent Denis Style
// ═══════════════════════════════════════════════
const C = {
  bg1: "0D0221",    // fond très sombre
  bg2: "150930",    // fond cards
  bg3: "1E0F45",    // fond cards plus clair
  v1: "7C3AED",     // violet principal
  v2: "A855F7",     // violet clair
  v3: "C4B5FD",     // violet très clair
  v4: "6D28D9",     // violet foncé
  v5: "8B5CF6",     // violet moyen
  pk: "EC4899",     // rose/pink
  mg: "D946EF",     // magenta
  w: "FFFFFF",
  ow: "F0E6FF",     // off-white
  g1: "D1D5DB",     // gris clair
  g2: "9CA3AF",     // gris moyen
  g3: "6B7280",     // gris foncé
};

const SW = 10;
const SH = 5.625;

// ═══════════════════════════════════════════════
// DESIGN HELPERS
// ═══════════════════════════════════════════════

// Background avec zone violet sur la droite
function addBg(pres, sl) {
  sl.background = { color: C.bg1 };
  // Zone violette à droite (simule dégradé)
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 5, y: 0, w: 5, h: SH,
    fill: { color: C.v4, transparency: 82 },
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7, y: 0, w: 3, h: SH,
    fill: { color: C.v1, transparency: 88 },
  });
  // Glow subtil top-right
  sl.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.5, w: 4, h: 4,
    fill: { color: C.mg, transparency: 92 },
  });
}

// Cadre arrondi fin
function addFrame(pres, sl) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.25, y: 0.25, w: 9.5, h: 5.125,
    line: { color: C.w, width: 0.5, transparency: 85 },
    fill: { color: "000000", transparency: 98 },
    rectRadius: 0.15,
  });
}

// Logo "C" en footer
function addLogo(pres, sl) {
  sl.addShape(pres.shapes.OVAL, {
    x: 4.72, y: 5.1, w: 0.36, h: 0.36,
    fill: { color: C.v1, transparency: 30 },
    line: { color: C.v2, width: 0.5 },
  });
  sl.addText("C", {
    x: 4.72, y: 5.1, w: 0.36, h: 0.36,
    fontSize: 12, color: C.v3, fontFace: "Arial",
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// Pill / Badge blanc
function pill(sl, x, y, w, h, text, opts = {}) {
  const bg = opts.bg || C.w;
  const color = opts.color || C.bg1;
  const fs = opts.fontSize || 13;
  sl.addShape("roundedRectangle", {
    x, y, w, h,
    fill: { color: bg, transparency: opts.bgTransparency || 0 },
    rectRadius: h / 2,
    shadow: opts.shadow !== false ? { type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.25 } : undefined,
    line: opts.border ? { color: opts.border, width: 0.5 } : undefined,
  });
  sl.addText(text, {
    x, y, w, h,
    fontSize: fs, color, fontFace: opts.font || "Arial",
    bold: opts.bold !== false, align: "center", valign: "middle", margin: 0,
    italic: opts.italic || false,
  });
}

// Cercle numéroté
function numCircle(pres, sl, x, y, size, num, opts = {}) {
  sl.addShape(pres.shapes.OVAL, {
    x, y, w: size, h: size,
    fill: { color: opts.bg || C.v1 },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.2 },
  });
  sl.addText(String(num), {
    x, y, w: size, h: size,
    fontSize: opts.fontSize || Math.round(size * 14),
    color: C.w, fontFace: "Arial",
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// Card avec accent latéral
function cardShape(pres, sl, x, y, w, h, opts = {}) {
  sl.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: opts.fill || C.bg2, transparency: opts.fillTransparency || 15 },
    line: { color: opts.border || C.v1, width: 0.5, transparency: 50 },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 },
  });
  if (opts.accent !== false) {
    sl.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.04, h,
      fill: { color: opts.accentColor || C.v2 },
    });
  }
}

// Ligne de séparation
function hLine(pres, sl, x, y, w, opts = {}) {
  sl.addShape(pres.shapes.LINE, {
    x, y, w, h: 0,
    line: { color: opts.color || C.v1, width: opts.width || 2 },
  });
}

// ═══════════════════════════════════════════════
// HEALTH CHECK
// ═══════════════════════════════════════════════
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "CREATICAL PPTX v2.5 Max Design" });
});

// ═══════════════════════════════════════════════
// GENERATE PPTX
// ═══════════════════════════════════════════════
app.post("/generate", async (req, res) => {
  try {
    const { slidesData } = req.body;
    if (!slidesData || !slidesData.slides) {
      return res.status(400).json({ error: "slidesData.slides required" });
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
      addFrame(pres, sl);

      switch (s.type) {

        // ═══════════ TITRE ═══════════
        case "title": {
          // Subtitle date
          sl.addText(s.subtitle || "", {
            x: 0.8, y: 0.9, w: 5, h: 0.4,
            fontSize: 16, color: C.ow, fontFace: "Georgia", italic: true,
          });
          // Main title
          sl.addText((s.title || "STRATEGIE").toUpperCase(), {
            x: 0.8, y: 1.4, w: 5.5, h: 1.0,
            fontSize: 46, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          // "de contenu" en italique
          sl.addText("de contenu", {
            x: 0.8, y: 2.35, w: 4, h: 0.6,
            fontSize: 30, color: C.v3, fontFace: "Georgia", italic: true,
          });
          // Client pill
          pill(sl, 0.8, 3.2, 4.5, 0.45, s.client || "");
          // Grand "C" logo à droite
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 7, y: 1.2, w: 1.8, h: 1.8,
            fill: { color: C.v1 },
            rectRadius: 0.3,
            shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 },
          });
          sl.addText("C", {
            x: 7, y: 1.2, w: 1.8, h: 1.8,
            fontSize: 80, color: C.w, fontFace: "Arial", bold: true,
            align: "center", valign: "middle", margin: 0,
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ OBJECTIFS ═══════════
        case "objectifs": {
          // Titre + année à gauche
          sl.addText("OBJECTIFS", {
            x: 0.8, y: 0.6, w: 4, h: 0.5,
            fontSize: 18, color: C.v3, fontFace: "Arial", bold: true,
            charSpacing: 4, italic: true,
          });
          sl.addText(s.year || "2026", {
            x: 0.8, y: 1.0, w: 4, h: 1.8,
            fontSize: 96, color: C.w, fontFace: "Arial", bold: true,
          });
          // Client pill
          pill(sl, 0.8, 3.1, 3.5, 0.4, s.clientName || meta.client || "", { fontSize: 11 });
          // Objectifs à droite avec flèches
          const items = s.items || [];
          // Point central d'où partent les lignes
          const cx = 5.2, cy = 2.6;
          items.forEach((item, i) => {
            const yPos = 0.7 + i * 1.1;
            // Ligne de connexion
            sl.addShape(pres.shapes.LINE, {
              x: cx, y: cy, w: 0.8, h: yPos - cy,
              line: { color: C.w, width: 1 },
            });
            // Texte avec bold pour mots clés
            sl.addText(item, {
              x: 6.2, y: yPos - 0.15, w: 3.3, h: 0.8,
              fontSize: 12, color: C.w, fontFace: "Arial", valign: "middle",
              bold: true, lineSpacingMultiple: 1.2,
            });
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ SECTION DIVIDER ═══════════
        case "section": {
          sl.addText((s.title || "").toUpperCase(), {
            x: 0.8, y: 1.0, w: 8.4, h: 1.5,
            fontSize: 52, color: C.w, fontFace: "Arial", bold: true,
            align: "center", charSpacing: 4,
          });
          sl.addText(s.subtitle || "", {
            x: 1, y: 2.7, w: 8, h: 0.7,
            fontSize: 22, color: C.v3, fontFace: "Georgia", italic: true,
            align: "center",
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ BRAND ARCHETYPE ═══════════
        case "brand_archetype": {
          sl.addText("BRAND ARCHETYPE", {
            x: 0.8, y: 0.5, w: 8, h: 0.5,
            fontSize: 26, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          pill(sl, 0.8, 1.15, 5, 0.4, s.subtitle || "Qui est le client ?");
          // Description card
          cardShape(pres, sl, 0.8, 1.8, 8.4, 3.0);
          sl.addText(s.description || "", {
            x: 1.05, y: 1.9, w: 8.0, h: 2.8,
            fontSize: 12, color: C.g1, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.5,
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ PERSONA ═══════════
        case "persona": {
          sl.addText("PERSONA CIBLE", {
            x: 0.8, y: 0.45, w: 5, h: 0.5,
            fontSize: 26, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          pill(sl, 0.8, 1.05, 4.5, 0.38, s.name || "Persona", { bg: C.v1, color: C.w });
          // Left card - Profil
          cardShape(pres, sl, 0.8, 1.65, 4.1, 3.2);
          sl.addText("PROFIL", {
            x: 1.05, y: 1.72, w: 3.6, h: 0.3,
            fontSize: 11, color: C.v3, fontFace: "Arial", bold: true,
            charSpacing: 2, margin: 0,
          });
          sl.addText(s.profil || "", {
            x: 1.05, y: 2.05, w: 3.6, h: 1.2,
            fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.4, margin: 0,
          });
          sl.addText("CORE PROBLEM", {
            x: 1.05, y: 3.3, w: 3.6, h: 0.25,
            fontSize: 11, color: C.pk, fontFace: "Arial", bold: true,
            charSpacing: 1, margin: 0,
          });
          sl.addText(s.core_problem || "", {
            x: 1.05, y: 3.58, w: 3.6, h: 0.8,
            fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top",
            lineSpacingMultiple: 1.3, margin: 0,
          });
          // Right card - Emotions
          cardShape(pres, sl, 5.1, 1.65, 4.3, 3.2, { accentColor: C.pk });
          sl.addText("EMOTIONS & PEURS", {
            x: 5.35, y: 1.72, w: 3.8, h: 0.3,
            fontSize: 11, color: C.v3, fontFace: "Arial", bold: true,
            charSpacing: 2, margin: 0,
          });
          const emotions = s.emotions || [];
          const emoText = emotions.map((e, i) => ({
            text: e,
            options: { breakLine: true, fontSize: 10, color: C.g1, fontFace: "Arial", bullet: true },
          }));
          if (emoText.length) {
            sl.addText(emoText, {
              x: 5.35, y: 2.1, w: 3.8, h: 2.5, valign: "top",
              paraSpaceAfter: 4,
            });
          }
          addLogo(pres, sl);
          break;
        }

        // ═══════════ POSITIONNEMENT ═══════════
        case "positionnement": {
          sl.addText("POSITIONNEMENT", {
            x: 0.8, y: 0.45, w: 8, h: 0.5,
            fontSize: 26, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          // Tagline dans bande violette
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 0.8, y: 1.15, w: 8.4, h: 0.55,
            fill: { color: C.v1 }, rectRadius: 0.08,
          });
          sl.addText(s.tagline || "", {
            x: 0.8, y: 1.15, w: 8.4, h: 0.55,
            fontSize: 14, color: C.w, fontFace: "Georgia", italic: true,
            align: "center", valign: "middle", margin: 0,
          });
          // Elements en cards 2x2
          const elems = s.elements || [];
          elems.forEach((el, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const xP = 0.8 + col * 4.3;
            const yP = 2.0 + row * 1.4;
            cardShape(pres, sl, xP, yP, 4.0, 1.2, {
              accentColor: [C.v1, C.v2, C.mg, C.pk][i] || C.v1,
            });
            numCircle(pres, sl, xP + 0.15, yP + 0.3, 0.4, i + 1, { fontSize: 12 });
            sl.addText(el, {
              x: xP + 0.7, y: yP + 0.15, w: 3.1, h: 0.9,
              fontSize: 11, color: C.g1, fontFace: "Arial", valign: "middle",
              lineSpacingMultiple: 1.3,
            });
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ PILIERS ═══════════
        case "piliers": {
          sl.addText("PILIERS ÉDITORIAUX", {
            x: 0.8, y: 0.45, w: 8, h: 0.5,
            fontSize: 26, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          const piliers = s.piliers || [];
          const pColors = [C.v1, C.v2, C.mg, C.pk];
          piliers.forEach((p, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            const xP = 0.8 + col * 4.3;
            const yP = 1.2 + row * 2.05;
            cardShape(pres, sl, xP, yP, 4.0, 1.8, { accentColor: pColors[i] });
            numCircle(pres, sl, xP + 0.2, yP + 0.18, 0.38, i + 1, {
              bg: pColors[i], fontSize: 11,
            });
            sl.addText(p.nom || "", {
              x: xP + 0.7, y: yP + 0.12, w: 3.1, h: 0.35,
              fontSize: 13, color: C.w, fontFace: "Arial", bold: true, margin: 0,
            });
            sl.addText(p.description || "", {
              x: xP + 0.2, y: yP + 0.6, w: 3.6, h: 1.1,
              fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.4, margin: 0,
            });
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ FUNNEL ═══════════
        case "funnel": {
          sl.addText("FUNNEL D'ACQUISITION", {
            x: 0.8, y: 0.35, w: 8, h: 0.45,
            fontSize: 22, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          const stages = [
            { data: s.tofu, color: C.v1 },
            { data: s.mofu, color: C.v2 },
            { data: s.bofu, color: C.mg },
          ];
          stages.forEach((stage, i) => {
            if (!stage.data) return;
            const xP = 0.5 + i * 3.1;
            // Column border
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xP, y: 0.95, w: 2.95, h: 4.3,
              fill: { color: C.bg2, transparency: 15 },
              line: { color: stage.color, width: 0.5, transparency: 40 },
            });
            // Header bar
            sl.addShape(pres.shapes.RECTANGLE, {
              x: xP, y: 0.95, w: 2.95, h: 0.5,
              fill: { color: stage.color },
            });
            sl.addText(stage.data.titre || "", {
              x: xP, y: 0.95, w: 2.95, h: 0.5,
              fontSize: 14, color: C.w, fontFace: "Arial",
              bold: true, align: "center", valign: "middle", margin: 0,
              charSpacing: 1,
            });
            // Description
            sl.addText(stage.data.description || "", {
              x: xP + 0.12, y: 1.5, w: 2.7, h: 0.5,
              fontSize: 9, color: C.v3, fontFace: "Arial",
              italic: true, valign: "top", margin: 0,
            });
            // Video ideas
            const videos = stage.data.videos || [];
            videos.forEach((v, j) => {
              // Numéro
              sl.addText(`${j + 1}.`, {
                x: xP + 0.12, y: 2.05 + j * 0.42, w: 0.25, h: 0.38,
                fontSize: 9, color: C.v2, fontFace: "Arial", bold: true,
                valign: "top", margin: 0,
              });
              sl.addText(v, {
                x: xP + 0.35, y: 2.05 + j * 0.42, w: 2.4, h: 0.38,
                fontSize: 9, color: C.g1, fontFace: "Arial",
                valign: "top", margin: 0,
              });
              // Separator line
              if (j < videos.length - 1) {
                sl.addShape(pres.shapes.LINE, {
                  x: xP + 0.12, y: 2.39 + j * 0.42, w: 2.7, h: 0,
                  line: { color: C.w, width: 0.3, transparency: 90 },
                });
              }
            });
          });
          break;
        }

        // ═══════════ PLAN D'ACTION ═══════════
        case "plan_action": {
          sl.addText("PLAN D'ACTION 90 JOURS", {
            x: 0.8, y: 0.35, w: 8, h: 0.5,
            fontSize: 22, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 3,
          });
          // Timeline line
          hLine(pres, sl, 0.5, 2.65, 9);
          const phases = s.phases || [];
          phases.forEach((phase, i) => {
            const xP = 0.5 + i * 3.1;
            // Phase pill above
            pill(sl, xP + 0.3, 1.2, 2.4, 0.35, phase.titre || "", {
              bg: C.v1, color: C.w, fontSize: 10, shadow: false,
            });
            // Période
            sl.addText(phase.periode || "", {
              x: xP + 0.3, y: 1.6, w: 2.4, h: 0.25,
              fontSize: 9, color: C.g2, fontFace: "Arial", align: "center",
            });
            // Circle on timeline
            numCircle(pres, sl, xP + 1.2, 2.37, 0.55, i + 1, { bg: C.v1 });
            // Actions below
            const actions = phase.actions || [];
            actions.forEach((a, j) => {
              sl.addText(a, {
                x: xP + 0.15, y: 3.1 + j * 0.45, w: 2.7, h: 0.42,
                fontSize: 9, color: C.g1, fontFace: "Arial",
                bullet: true, valign: "top", margin: 0,
              });
            });
          });
          break;
        }

        // ═══════════ RESUME ═══════════
        case "resume": {
          // Big title left
          sl.addText("RÉSUMÉ", {
            x: 0.8, y: 1.0, w: 4, h: 0.8,
            fontSize: 44, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 4,
          });
          sl.addText("du mois", {
            x: 0.8, y: 1.85, w: 4, h: 0.6,
            fontSize: 28, color: C.v3, fontFace: "Georgia", italic: true,
          });
          // Client name pill
          pill(sl, 0.8, 2.7, 3, 0.35, meta.client || "", { fontSize: 10 });
          // Points on right
          const points = s.points || [];
          points.forEach((pt, i) => {
            const yP = 0.7 + i * 1.45;
            pill(sl, 5.2, yP, 4.2, 0.38, pt.titre || "", { fontSize: 11 });
            sl.addText(pt.description || "", {
              x: 5.2, y: yP + 0.48, w: 4.2, h: 0.8,
              fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.4,
            });
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ CONTACT ═══════════
        case "contact": {
          // Logo "C" centré à droite
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
            x: 6.2, y: 1.0, w: 1.4, h: 1.4,
            fill: { color: C.v1 }, rectRadius: 0.25,
            shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 },
          });
          sl.addText("C", {
            x: 6.2, y: 1.0, w: 1.4, h: 1.4,
            fontSize: 64, color: C.w, fontFace: "Arial", bold: true,
            align: "center", valign: "middle", margin: 0,
          });
          sl.addText(s.name || "CREATICAL", {
            x: 5.5, y: 2.7, w: 3.5, h: 0.5,
            fontSize: 20, color: C.w, fontFace: "Arial", bold: true,
          });
          sl.addText(s.role || "Social Media Manager", {
            x: 5.5, y: 3.2, w: 3.5, h: 0.4,
            fontSize: 15, color: C.v3, fontFace: "Georgia", italic: true,
          });
          sl.addText(s.email || "contact@creatical.be", {
            x: 5.5, y: 3.6, w: 3.5, h: 0.4,
            fontSize: 13, color: C.g2, fontFace: "Georgia", italic: true,
          });
          addLogo(pres, sl);
          break;
        }

        // ═══════════ GENERIC ═══════════
        default: {
          sl.addText((s.title || "").toUpperCase(), {
            x: 0.8, y: 0.45, w: 8, h: 0.5,
            fontSize: 24, color: C.w, fontFace: "Arial", bold: true,
            charSpacing: 2,
          });
          if (s.content) {
            cardShape(pres, sl, 0.8, 1.2, 8.4, 3.8);
            sl.addText(s.content, {
              x: 1.05, y: 1.3, w: 8.0, h: 3.6,
              fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top",
              lineSpacingMultiple: 1.4,
            });
          }
          addLogo(pres, sl);
          break;
        }
      }
    }

    const base64 = await pres.write({ outputType: "base64" });
    res.json({
      success: true,
      pptxBase64: base64,
      filename: `${meta.title || "Presentation"}.pptx`,
      slideCount: slides.length,
    });
  } catch (error) {
    console.error("Error:", error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CREATICAL PPTX v2.5 Max Design on port ${PORT}`);
});
