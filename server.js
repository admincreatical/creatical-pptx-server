const express = require("express");
const pptxgen = require("pptxgenjs");

const app = express();
app.use(express.json({ limit: "10mb" }));

const C = {
  bg1: "0D0221", bg2: "150930", bg3: "1E0F45",
  v1: "7C3AED", v2: "A855F7", v3: "C4B5FD", v4: "6D28D9", v5: "8B5CF6",
  pk: "EC4899", mg: "D946EF",
  w: "FFFFFF", ow: "F0E6FF",
  g1: "D1D5DB", g2: "9CA3AF", g3: "6B7280",
};
const SW = 10, SH = 5.625;

function addBg(pres, sl) {
  sl.background = { color: C.bg1 };
  sl.addShape(pres.shapes.RECTANGLE, { x: 5, y: 0, w: 5, h: SH, fill: { color: C.v4, transparency: 82 } });
  sl.addShape(pres.shapes.RECTANGLE, { x: 7, y: 0, w: 3, h: SH, fill: { color: C.v1, transparency: 88 } });
  sl.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.5, w: 4, h: 4, fill: { color: C.mg, transparency: 92 } });
}
function addFrame(pres, sl) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.25, y: 0.25, w: 9.5, h: 5.125,
    line: { color: C.w, width: 0.5, transparency: 85 },
    fill: { color: "000000", transparency: 98 }, rectRadius: 0.15,
  });
}
function addLogo(pres, sl) {
  sl.addShape(pres.shapes.OVAL, { x: 4.72, y: 5.1, w: 0.36, h: 0.36, fill: { color: C.v1, transparency: 30 }, line: { color: C.v2, width: 0.5 } });
  sl.addText("C", { x: 4.72, y: 5.1, w: 0.36, h: 0.36, fontSize: 12, color: C.v3, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
}
function pill(sl, x, y, w, h, text, opts = {}) {
  const bg = opts.bg || C.w, color = opts.color || C.bg1, fs = opts.fontSize || 13;
  sl.addShape("roundedRectangle", { x, y, w, h, fill: { color: bg, transparency: opts.bgTransparency || 0 }, rectRadius: h / 2, shadow: opts.shadow !== false ? { type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.25 } : undefined, line: opts.border ? { color: opts.border, width: 0.5 } : undefined });
  sl.addText(text, { x, y, w, h, fontSize: fs, color, fontFace: opts.font || "Arial", bold: opts.bold !== false, align: "center", valign: "middle", margin: 0, italic: opts.italic || false });
}
function numCircle(pres, sl, x, y, size, num, opts = {}) {
  sl.addShape(pres.shapes.OVAL, { x, y, w: size, h: size, fill: { color: opts.bg || C.v1 }, shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.2 } });
  sl.addText(String(num), { x, y, w: size, h: size, fontSize: opts.fontSize || Math.round(size * 14), color: C.w, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
}
function cardShape(pres, sl, x, y, w, h, opts = {}) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: opts.fill || C.bg2, transparency: opts.fillTransparency || 15 }, line: { color: opts.border || C.v1, width: 0.5, transparency: 50 }, shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 } });
  if (opts.accent !== false) { sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.04, h, fill: { color: opts.accentColor || C.v2 } }); }
}
function hLine(pres, sl, x, y, w, opts = {}) {
  sl.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: opts.color || C.v1, width: opts.width || 2 } });
}

app.get("/", (req, res) => { res.json({ status: "ok", service: "CREATICAL PPTX v3.0 Full Strategy" }); });

app.post("/generate", async (req, res) => {
  try {
    const { slidesData } = req.body;
    if (!slidesData || !slidesData.slides) return res.status(400).json({ error: "slidesData.slides required" });
    const slides = slidesData.slides, meta = slidesData.metadata || {};
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "CREATICAL";
    pres.title = meta.title || "Stratégie de Contenu";

    for (const s of slides) {
      const sl = pres.addSlide();
      addBg(pres, sl);
      addFrame(pres, sl);

      switch (s.type) {

        case "title": {
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.9, w: 5, h: 0.4, fontSize: 16, color: C.ow, fontFace: "Georgia", italic: true });
          sl.addText((s.title || "STRATEGIE").toUpperCase(), { x: 0.8, y: 1.4, w: 5.5, h: 1.0, fontSize: 46, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          sl.addText("de contenu", { x: 0.8, y: 2.35, w: 4, h: 0.6, fontSize: 30, color: C.v3, fontFace: "Georgia", italic: true });
          pill(sl, 0.8, 3.2, 4.5, 0.45, s.client || "");
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7, y: 1.2, w: 1.8, h: 1.8, fill: { color: C.v1 }, rectRadius: 0.3, shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 } });
          sl.addText("C", { x: 7, y: 1.2, w: 1.8, h: 1.8, fontSize: 80, color: C.w, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
          addLogo(pres, sl);
          break;
        }

        case "section": {
          sl.addText((s.title || "").toUpperCase(), { x: 0.8, y: 1.0, w: 8.4, h: 1.5, fontSize: 52, color: C.w, fontFace: "Arial", bold: true, align: "center", charSpacing: 4 });
          sl.addText(s.subtitle || "", { x: 1, y: 2.7, w: 8, h: 0.7, fontSize: 22, color: C.v3, fontFace: "Georgia", italic: true, align: "center" });
          addLogo(pres, sl);
          break;
        }

        case "objectifs": {
          sl.addText("OBJECTIFS", { x: 0.8, y: 0.6, w: 4, h: 0.5, fontSize: 18, color: C.v3, fontFace: "Arial", bold: true, charSpacing: 4, italic: true });
          sl.addText(s.year || "2026", { x: 0.8, y: 1.0, w: 4, h: 1.8, fontSize: 96, color: C.w, fontFace: "Arial", bold: true });
          pill(sl, 0.8, 3.1, 3.5, 0.4, s.clientName || meta.client || "", { fontSize: 11 });
          const items = s.items || [];
          const cx = 5.2, cy = 2.6;
          items.forEach((item, i) => {
            const yPos = 0.7 + i * 1.1;
            sl.addShape(pres.shapes.LINE, { x: cx, y: cy, w: 0.8, h: yPos - cy, line: { color: C.w, width: 1 } });
            sl.addText(item, { x: 6.2, y: yPos - 0.15, w: 3.3, h: 0.8, fontSize: 12, color: C.w, fontFace: "Arial", valign: "middle", bold: true, lineSpacingMultiple: 1.2 });
          });
          addLogo(pres, sl);
          break;
        }

        case "contexte": {
          sl.addText("CONTEXTE", { x: 0.8, y: 0.45, w: 8, h: 0.5, fontSize: 26, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          cardShape(pres, sl, 0.8, 1.2, 8.4, 3.8);
          sl.addText(s.description || "", { x: 1.05, y: 1.35, w: 8.0, h: 3.5, fontSize: 13, color: C.g1, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.6, align: "center" });
          addLogo(pres, sl);
          break;
        }

        case "marche": {
          sl.addText("À QUOI RESSEMBLE TON MARCHÉ ?", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 20, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.85, w: 8.4, h: 0.4, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true, align: "center" });
          const stats = s.stats || [];
          stats.forEach((stat, i) => {
            const col = i % 2, row = Math.floor(i / 2);
            const xP = 0.5 + col * 4.8, yP = 1.45 + row * 2.0;
            cardShape(pres, sl, xP, yP, 4.5, 1.75);
            numCircle(pres, sl, xP + 0.18, yP + 0.18, 0.42, stat.numero || i+1, { bg: C.v1, fontSize: 12 });
            sl.addText(stat.chiffre || "", { x: xP + 0.75, y: yP + 0.12, w: 3.5, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            sl.addText(stat.description || "", { x: xP + 0.2, y: yP + 0.7, w: 4.1, h: 0.9, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.35, margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "geographie": {
          sl.addText("OÙ SE TROUVENT TES CLIENTS ?", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.85, w: 8.4, h: 0.4, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true, align: "center" });
          cardShape(pres, sl, 0.5, 1.45, 4.0, 3.7, { accentColor: C.v1 });
          sl.addText("Principaux marchés", { x: 0.75, y: 1.55, w: 3.5, h: 0.35, fontSize: 12, color: C.v3, fontFace: "Georgia", italic: true, bold: true, margin: 0 });
          (s.marches_principaux || []).forEach((m, i) => {
            sl.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 2.05 + i * 0.7, w: 3.5, h: 0.55, fill: { color: C.bg3, transparency: 20 }, line: { color: C.v1, width: 0.3, transparency: 60 } });
            sl.addText(m, { x: 0.75, y: 2.05 + i * 0.7, w: 3.5, h: 0.55, fontSize: 14, color: C.w, fontFace: "Arial", align: "center", valign: "middle", margin: 0 });
          });
          cardShape(pres, sl, 5.0, 1.45, 4.2, 3.7, { accentColor: C.v2 });
          sl.addText("Villes prioritaires à cibler", { x: 5.25, y: 1.55, w: 3.7, h: 0.35, fontSize: 12, color: C.v3, fontFace: "Georgia", italic: true, bold: true, margin: 0 });
          (s.villes_prioritaires || []).forEach((v, i) => {
            sl.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 2.05 + i * 0.65, w: 3.7, h: 0.5, fill: { color: C.bg3, transparency: 20 }, line: { color: C.v2, width: 0.3, transparency: 60 } });
            sl.addText(v, { x: 5.25, y: 2.05 + i * 0.65, w: 3.7, h: 0.5, fontSize: 13, color: C.w, fontFace: "Arial", align: "center", valign: "middle", margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "concurrence": {
          sl.addText("ANALYSE CONCURRENTIELLE", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "Connais ton ennemi et connais-toi toi-même.", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true });
          (s.concurrents || []).forEach((c, i) => {
            const xP = 0.5 + i * 4.75;
            cardShape(pres, sl, xP, 1.3, 4.4, 3.5, { accentColor: [C.v1, C.v2][i] || C.v1 });
            sl.addText(c.nom || "", { x: xP + 0.2, y: 1.38, w: 4.0, h: 0.35, fontSize: 13, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            [{ label: "Positionnement", val: c.positionnement, color: C.v3 }, { label: "Forces", val: c.forces, color: C.v2 }, { label: "Faiblesses", val: c.faiblesses, color: C.pk }, { label: "Opportunité", val: c.opportunite, color: C.mg }].forEach((item, j) => {
              sl.addText(item.label, { x: xP + 0.2, y: 1.85 + j * 0.75, w: 4.0, h: 0.22, fontSize: 9, color: item.color, fontFace: "Arial", bold: true, charSpacing: 1, margin: 0 });
              sl.addText(item.val || "", { x: xP + 0.2, y: 2.08 + j * 0.75, w: 4.0, h: 0.42, fontSize: 9, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.3, margin: 0 });
            });
          });
          addLogo(pres, sl);
          break;
        }

        case "niveau_conscience": {
          sl.addText("NIVEAU DE CONSCIENCE DE L'AUDIENCE", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 18, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          const niveaux = [{ key: "inconscient", label: "Inconscient du problème" }, { key: "conscient_p", label: "Conscient du problème" }, { key: "conscient_s", label: "Conscient de la solution" }, { key: "conscient_pr", label: "Conscient du produit" }, { key: "pret", label: "Prêt à acheter" }];
          const currentN = (s.niveau || "").toLowerCase();
          niveaux.forEach((n, i) => {
            const isActive = currentN.includes(n.key.substring(0, 6));
            const xP = 0.5 + i * 1.82;
            sl.addShape(pres.shapes.RECTANGLE, { x: xP, y: 1.1, w: 1.65, h: 0.6, fill: { color: isActive ? C.v1 : C.bg3, transparency: isActive ? 0 : 30 }, line: { color: isActive ? C.v2 : C.g3, width: isActive ? 1 : 0.3 } });
            sl.addText(n.label, { x: xP, y: 1.1, w: 1.65, h: 0.6, fontSize: 8.5, color: isActive ? C.w : C.g2, fontFace: "Arial", bold: isActive, align: "center", valign: "middle", margin: 2 });
          });
          cardShape(pres, sl, 0.5, 2.1, 8.8, 1.4);
          sl.addText("Où se situe l'audience :", { x: 0.75, y: 2.18, w: 8.2, h: 0.3, fontSize: 11, color: C.v3, fontFace: "Arial", bold: true, margin: 0 });
          sl.addText(s.explication || "", { x: 0.75, y: 2.5, w: 8.2, h: 0.85, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          cardShape(pres, sl, 0.5, 3.65, 8.8, 1.3, { accentColor: C.mg });
          sl.addText("Impact sur le contenu :", { x: 0.75, y: 3.73, w: 8.2, h: 0.3, fontSize: 11, color: C.mg, fontFace: "Arial", bold: true, margin: 0 });
          sl.addText(s.impact_message || "", { x: 0.75, y: 4.05, w: 8.2, h: 0.75, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          addLogo(pres, sl);
          break;
        }

        case "sophistication_marche": {
          sl.addText("SOPHISTICATION DU MARCHÉ", { x: 0.8, y: 0.45, w: 8, h: 0.5, fontSize: 24, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          const niveauLabel = (s.niveau || "").toUpperCase();
          const niveauColor = niveauLabel.includes("ELEV") ? C.pk : niveauLabel.includes("MOY") ? C.v2 : C.v1;
          pill(sl, 0.8, 1.15, 3.5, 0.55, "Niveau : " + (s.niveau || ""), { bg: niveauColor, color: C.w, fontSize: 16, bold: true });
          cardShape(pres, sl, 0.8, 1.9, 8.4, 1.5);
          sl.addText("Analyse :", { x: 1.05, y: 1.98, w: 8.0, h: 0.3, fontSize: 11, color: C.v3, fontFace: "Arial", bold: true, margin: 0 });
          sl.addText(s.explication || "", { x: 1.05, y: 2.3, w: 8.0, h: 0.95, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.5, margin: 0 });
          cardShape(pres, sl, 0.8, 3.55, 8.4, 1.4, { accentColor: C.mg });
          sl.addText("Conséquence sur le contenu :", { x: 1.05, y: 3.63, w: 8.0, h: 0.3, fontSize: 11, color: C.mg, fontFace: "Arial", bold: true, margin: 0 });
          sl.addText(s.consequence_contenu || "", { x: 1.05, y: 3.95, w: 8.0, h: 0.85, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          addLogo(pres, sl);
          break;
        }

        case "brand_archetype": {
          sl.addText("BRAND ARCHETYPE", { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 26, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          pill(sl, 0.8, 1.15, 5, 0.4, s.subtitle || "Qui est le client ?");
          cardShape(pres, sl, 0.8, 1.8, 8.4, 3.0);
          sl.addText(s.description || "", { x: 1.05, y: 1.9, w: 8.0, h: 2.8, fontSize: 12, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.5 });
          addLogo(pres, sl);
          break;
        }

        case "persona": {
          sl.addText("PERSONA CIBLE", { x: 0.8, y: 0.45, w: 5, h: 0.5, fontSize: 26, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          pill(sl, 0.8, 1.05, 4.5, 0.38, s.name || "Persona", { bg: C.v1, color: C.w });
          cardShape(pres, sl, 0.8, 1.65, 4.1, 3.2);
          sl.addText("PROFIL", { x: 1.05, y: 1.72, w: 3.6, h: 0.3, fontSize: 11, color: C.v3, fontFace: "Arial", bold: true, charSpacing: 2, margin: 0 });
          sl.addText(s.profil || "", { x: 1.05, y: 2.05, w: 3.6, h: 1.2, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          sl.addText("CORE PROBLEM", { x: 1.05, y: 3.3, w: 3.6, h: 0.25, fontSize: 11, color: C.pk, fontFace: "Arial", bold: true, charSpacing: 1, margin: 0 });
          sl.addText(s.core_problem || "", { x: 1.05, y: 3.58, w: 3.6, h: 0.8, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.3, margin: 0 });
          cardShape(pres, sl, 5.1, 1.65, 4.3, 3.2, { accentColor: C.pk });
          sl.addText("EMOTIONS & PEURS", { x: 5.35, y: 1.72, w: 3.8, h: 0.3, fontSize: 11, color: C.v3, fontFace: "Arial", bold: true, charSpacing: 2, margin: 0 });
          const emotions = s.emotions || [];
          const emoText = emotions.map(e => ({ text: e, options: { breakLine: true, fontSize: 10, color: C.g1, fontFace: "Arial", bullet: true } }));
          if (emoText.length) sl.addText(emoText, { x: 5.35, y: 2.1, w: 3.8, h: 2.5, valign: "top", paraSpaceAfter: 4 });
          addLogo(pres, sl);
          break;
        }

        case "positionnement": {
          sl.addText("POSITIONNEMENT", { x: 0.8, y: 0.45, w: 8, h: 0.5, fontSize: 26, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.15, w: 8.4, h: 0.55, fill: { color: C.v1 }, rectRadius: 0.08 });
          sl.addText(s.tagline || "", { x: 0.8, y: 1.15, w: 8.4, h: 0.55, fontSize: 14, color: C.w, fontFace: "Georgia", italic: true, align: "center", valign: "middle", margin: 0 });
          const elems = s.elements || [];
          elems.forEach((el, i) => {
            const col = i % 2, row = Math.floor(i / 2);
            const xP = 0.8 + col * 4.3, yP = 2.0 + row * 1.4;
            cardShape(pres, sl, xP, yP, 4.0, 1.2, { accentColor: [C.v1, C.v2, C.mg, C.pk][i] || C.v1 });
            numCircle(pres, sl, xP + 0.15, yP + 0.3, 0.4, i + 1, { fontSize: 12 });
            sl.addText(el, { x: xP + 0.7, y: yP + 0.15, w: 3.1, h: 0.9, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.3 });
          });
          addLogo(pres, sl);
          break;
        }

        case "plateformes": {
          sl.addText("JUSTIFICATION DES PLATEFORMES", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 20, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "Pourquoi ces plateformes pour cette cible ?", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true });
          (s.plateformes || []).forEach((p, i) => {
            const xP = 0.4 + i * 3.1;
            cardShape(pres, sl, xP, 1.3, 2.9, 3.6, { accentColor: [C.v1, C.v2, C.mg][i] || C.v1 });
            sl.addText(p.nom || "", { x: xP + 0.2, y: 1.38, w: 2.5, h: 0.38, fontSize: 14, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            [{ label: "Pourquoi", val: p.pourquoi }, { label: "Audience", val: p.audience }, { label: "Contenu", val: p.contenu }, { label: "Fréquence", val: p.frequence }].forEach((item, j) => {
              sl.addText(item.label, { x: xP + 0.2, y: 1.88 + j * 0.72, w: 2.5, h: 0.22, fontSize: 9, color: C.v3, fontFace: "Arial", bold: true, charSpacing: 1, margin: 0 });
              sl.addText(item.val || "", { x: xP + 0.2, y: 2.1 + j * 0.72, w: 2.5, h: 0.42, fontSize: 9, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.3, margin: 0 });
            });
          });
          addLogo(pres, sl);
          break;
        }

        case "kpis": {
          sl.addText("KPIs & MÉTRIQUES DE SUCCÈS", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "Comment on mesure le succès mois par mois", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true });
          ["Objectif", "KPI Principal", "Mois 1", "Mois 2", "Mois 3"].forEach((h, i) => {
            const xCols = [0.4, 2.6, 5.5, 6.8, 8.1], wCols = [2.0, 2.7, 1.1, 1.1, 1.5];
            sl.addShape(pres.shapes.RECTANGLE, { x: xCols[i], y: 1.3, w: wCols[i], h: 0.38, fill: { color: C.v1 } });
            sl.addText(h, { x: xCols[i], y: 1.3, w: wCols[i], h: 0.38, fontSize: 10, color: C.w, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
          });
          (s.metriques || []).forEach((m, i) => {
            const yR = 1.78 + i * 1.0;
            sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: yR, w: 9.2, h: 0.88, fill: { color: i % 2 === 0 ? C.bg2 : C.bg3, transparency: 20 }, line: { color: C.v1, width: 0.3, transparency: 70 } });
            [m.objectif, m.kpi, m.mois1, m.mois2, m.mois3].forEach((v, j) => {
              const xCols = [0.4, 2.6, 5.5, 6.8, 8.1], wCols = [2.0, 2.7, 1.1, 1.1, 1.5];
              sl.addText(v || "", { x: xCols[j] + 0.08, y: yR + 0.1, w: wCols[j] - 0.16, h: 0.7, fontSize: 10, color: j >= 2 ? C.v3 : C.g1, fontFace: "Arial", bold: j >= 2, align: j >= 2 ? "center" : "left", valign: "middle", margin: 0, lineSpacingMultiple: 1.3 });
            });
          });
          addLogo(pres, sl);
          break;
        }

        case "piliers": {
          sl.addText("PILIERS ÉDITORIAUX", { x: 0.8, y: 0.45, w: 8, h: 0.5, fontSize: 26, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          const pColors = [C.v1, C.v2, C.mg, C.pk];
          (s.piliers || []).forEach((p, i) => {
            const col = i % 2, row = Math.floor(i / 2);
            const xP = 0.8 + col * 4.3, yP = 1.2 + row * 2.05;
            cardShape(pres, sl, xP, yP, 4.0, 1.8, { accentColor: pColors[i] });
            numCircle(pres, sl, xP + 0.2, yP + 0.18, 0.38, i + 1, { bg: pColors[i], fontSize: 11 });
            sl.addText(p.nom || "", { x: xP + 0.7, y: yP + 0.12, w: 3.1, h: 0.35, fontSize: 13, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            sl.addText(p.description || "", { x: xP + 0.2, y: yP + 0.6, w: 3.6, h: 1.1, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "contenu_strategique": {
          sl.addText("CONTENU ORDINAIRE VS CONTENU STRATÉGIQUE", { x: 0.8, y: 0.4, w: 8.4, h: 0.5, fontSize: 18, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.description || "", { x: 0.8, y: 1.0, w: 8.4, h: 0.5, fontSize: 12, color: C.g1, fontFace: "Arial", align: "center", lineSpacingMultiple: 1.4 });
          (s.differences || []).forEach((d, i) => {
            const yP = 1.7 + i * 0.8;
            sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yP, w: 3.8, h: 0.62, fill: { color: C.bg3, transparency: 20 }, line: { color: C.g3, width: 0.5 }, rectRadius: 0.3 });
            sl.addText("✕  " + (d.mauvais || ""), { x: 0.8, y: yP, w: 3.8, h: 0.62, fontSize: 13, color: C.g2, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
            sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.4, y: yP, w: 3.8, h: 0.62, fill: { color: C.v1, transparency: 10 }, line: { color: C.v2, width: 0.5 }, rectRadius: 0.3 });
            sl.addText("✓  " + (d.bon || ""), { x: 5.4, y: yP, w: 3.8, h: 0.62, fontSize: 13, color: C.w, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "piliers_contenu": {
          sl.addText("LES 3 PILIERS D'UN CONTENU QUI VEND", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 20, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "Un contenu efficace remplit 3 fonctions essentielles dans ton tunnel de vente.", { x: 0.8, y: 0.87, w: 8.4, h: 0.35, fontSize: 11, color: C.g1, fontFace: "Arial" });
          const fColors = [C.v1, C.v2, C.mg], fWidths = [8.0, 6.5, 5.0], fXl = [0.9, 1.65, 2.4];
          (s.piliers || []).forEach((p, i) => {
            const yP = 1.4 + i * 1.1;
            sl.addShape(pres.shapes.TRAPEZOID, { x: fXl[i], y: yP, w: fWidths[i], h: 0.9, fill: { color: fColors[i] }, shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.2 } });
            sl.addText(p.nom || "", { x: fXl[i], y: yP + 0.1, w: fWidths[i], h: 0.7, fontSize: 18, color: C.w, fontFace: "Georgia", italic: true, bold: true, align: "center", valign: "middle", margin: 0 });
            const descX = i % 2 === 0 ? 0.4 : 6.5;
            sl.addText(p.description || "", { x: descX, y: yP + 0.15, w: 2.8, h: 0.65, fontSize: 9, color: C.g2, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.3 });
          });
          addLogo(pres, sl);
          break;
        }

        case "methode": {
          sl.addText("LA MÉTHODE CREATICAL™", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.g1, fontFace: "Arial" });
          (s.etapes || []).forEach((e, i) => {
            const xP = 0.4 + i * 1.88;
            if (i < (s.etapes || []).length - 1) sl.addShape(pres.shapes.OVAL, { x: xP + 1.55, y: 1.95, w: 0.35, h: 0.35, fill: { color: C.v1, transparency: 60 } });
            cardShape(pres, sl, xP, 1.3, 1.6, 3.4, { accentColor: C.v1, accent: false });
            sl.addText(String(i + 1), { x: xP, y: 1.35, w: 1.6, h: 0.4, fontSize: 20, color: C.v3, fontFace: "Arial", bold: true, align: "center" });
            sl.addText(e, { x: xP + 0.1, y: 1.85, w: 1.4, h: 2.7, fontSize: 10, color: C.g1, fontFace: "Arial", align: "center", valign: "top", lineSpacingMultiple: 1.4, margin: 0, bold: true });
          });
          sl.addText("LE CONTENU AU SERVICE D'UNE APPROCHE SCIENTIFIQUE", { x: 0.4, y: 4.8, w: 9.2, h: 0.4, fontSize: 11, color: C.v3, fontFace: "Arial", bold: true, align: "center", charSpacing: 1 });
          addLogo(pres, sl);
          break;
        }

        case "hook_regles": {
          sl.addText("LA MÉTHODE CREATICAL™", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.g1, fontFace: "Arial" });
          (s.hooks || []).forEach((h, i) => {
            const yP = 1.4 + i * 1.3;
            cardShape(pres, sl, 0.8, yP, 8.4, 1.15, { accentColor: [C.v1, C.v2, C.mg][i] });
            sl.addText(h.type || "", { x: 1.05, y: yP + 0.1, w: 3.5, h: 0.3, fontSize: 12, color: C.v3, fontFace: "Arial", bold: true, charSpacing: 1, margin: 0 });
            sl.addText(h.description || "", { x: 1.05, y: yP + 0.45, w: 8.0, h: 0.6, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4, margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "emotion": {
          sl.addText("LA MÉTHODE CREATICAL™", { x: 0.8, y: 0.35, w: 8.4, h: 0.45, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "", { x: 0.8, y: 0.82, w: 8.4, h: 0.35, fontSize: 11, color: C.g1, fontFace: "Arial" });
          cardShape(pres, sl, 1.5, 1.3, 7.0, 2.5);
          sl.addText(s.description || "", { x: 1.7, y: 1.4, w: 6.6, h: 2.3, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.5, bullet: true });
          const eColors = [C.v1, C.v2, C.mg, C.pk, C.v3, C.v4, C.g2, C.ow];
          const positions = [{ x: 0.3, y: 0.9 }, { x: 8.5, y: 1.2 }, { x: 0.2, y: 2.5 }, { x: 8.3, y: 2.3 }, { x: 0.5, y: 3.8 }, { x: 8.2, y: 3.6 }, { x: 1.5, y: 4.5 }, { x: 7.0, y: 4.6 }];
          (s.emotions || []).forEach((e, i) => {
            const pos = positions[i] || { x: 1 + i * 1.2, y: 4.3 };
            const isBoxed = [0, 2, 5].includes(i);
            if (isBoxed) sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: pos.x - 0.1, y: pos.y - 0.05, w: 1.5, h: 0.4, fill: { color: eColors[i], transparency: 20 }, rectRadius: 0.2 });
            sl.addText(e, { x: pos.x, y: pos.y, w: 1.4, h: 0.38, fontSize: isBoxed ? 13 : 16, color: C.w, fontFace: "Georgia", italic: !isBoxed, bold: isBoxed, align: "center", valign: "middle" });
          });
          addLogo(pres, sl);
          break;
        }

        case "phase_funnel": {
          sl.addText((s.phase || "PHASE").toUpperCase(), { x: 0.8, y: 0.8, w: 8.4, h: 0.7, fontSize: 36, color: C.w, fontFace: "Arial", bold: true, align: "center", charSpacing: 6 });
          sl.addText(s.description || "", { x: 1, y: 1.65, w: 8, h: 0.5, fontSize: 14, color: C.g1, fontFace: "Arial", align: "center" });
          const fPhases = ["attirer", "convaincre", "convertir"], fCurrent = (s.funnel_focus || "").toLowerCase();
          const fColorsP = [C.v1, C.v2, C.mg], fWidthsP = [7.5, 6.0, 4.5], fXP = [1.15, 1.9, 2.65];
          fPhases.forEach((ph, i) => {
            const isActive = fCurrent === ph;
            sl.addShape(pres.shapes.TRAPEZOID, { x: fXP[i], y: 2.4 + i * 0.88, w: fWidthsP[i], h: 0.8, fill: { color: fColorsP[i], transparency: isActive ? 0 : 60 } });
            sl.addText(ph, { x: fXP[i], y: 2.4 + i * 0.88, w: fWidthsP[i], h: 0.8, fontSize: 18, color: isActive ? C.w : C.g3, fontFace: "Georgia", italic: true, bold: isActive, align: "center", valign: "middle", margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "funnel_idees": {
          const phaseColors = { TOFU: C.v1, MOFU: C.v2, BOFU: C.mg };
          const phaseColor = phaseColors[s.phase] || C.v1;
          sl.addText((s.title || "").toUpperCase(), { x: 0.8, y: 0.35, w: 7.5, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          pill(sl, 8.1, 0.3, 1.4, 0.5, s.phase || "", { bg: phaseColor, color: C.w, fontSize: 13 });
          sl.addText("Objectif : " + (s.objectif || ""), { x: 0.8, y: 0.9, w: 8.4, h: 0.4, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true });
          (s.idees || []).forEach((idee, i) => {
            const yP = 1.5 + i * 0.75;
            cardShape(pres, sl, 0.8, yP, 8.4, 0.65, { accent: false });
            numCircle(pres, sl, 0.85, yP + 0.1, 0.42, i + 1, { bg: phaseColor, fontSize: 12 });
            sl.addText(idee.titre || "", { x: 1.45, y: yP + 0.05, w: 4.8, h: 0.32, fontSize: 12, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            sl.addText(idee.objectif_video || "", { x: 1.45, y: yP + 0.36, w: 7.5, h: 0.25, fontSize: 9, color: C.g2, fontFace: "Arial", italic: true, margin: 0 });
            if (i < (s.idees || []).length - 1) sl.addShape(pres.shapes.LINE, { x: 0.8, y: yP + 0.65, w: 8.4, h: 0, line: { color: C.w, width: 0.3, transparency: 88 } });
          });
          addLogo(pres, sl);
          break;
        }

        case "plan_action": {
          sl.addText("PLAN D'ACTION 90 JOURS", { x: 0.8, y: 0.35, w: 8, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 3 });
          hLine(pres, sl, 0.5, 2.65, 9);
          (s.phases || []).forEach((phase, i) => {
            const xP = 0.5 + i * 3.1;
            pill(sl, xP + 0.3, 1.2, 2.4, 0.35, phase.titre || "", { bg: C.v1, color: C.w, fontSize: 10, shadow: false });
            sl.addText(phase.periode || "", { x: xP + 0.3, y: 1.6, w: 2.4, h: 0.25, fontSize: 9, color: C.g2, fontFace: "Arial", align: "center" });
            numCircle(pres, sl, xP + 1.2, 2.37, 0.55, i + 1, { bg: C.v1 });
            (phase.actions || []).forEach((a, j) => {
              sl.addText(a, { x: xP + 0.15, y: 3.1 + j * 0.45, w: 2.7, h: 0.42, fontSize: 9, color: C.g1, fontFace: "Arial", bullet: true, valign: "top", margin: 0 });
            });
          });
          addLogo(pres, sl);
          break;
        }

        case "fidelisation": {
          sl.addText("STRATÉGIE DE FIDÉLISATION", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText(s.subtitle || "Transformer les clients en ambassadeurs", { x: 0.8, y: 0.87, w: 8.4, h: 0.35, fontSize: 11, color: C.v3, fontFace: "Georgia", italic: true });
          (s.actions || []).forEach((a, i) => {
            const yP = 1.4 + i * 1.3;
            cardShape(pres, sl, 0.8, yP, 8.4, 1.15, { accentColor: [C.v1, C.v2, C.mg][i] });
            numCircle(pres, sl, 0.95, yP + 0.22, 0.45, i + 1, { bg: [C.v1, C.v2, C.mg][i], fontSize: 12 });
            sl.addText(a.titre || "", { x: 1.6, y: yP + 0.1, w: 7.4, h: 0.32, fontSize: 13, color: C.w, fontFace: "Arial", bold: true, margin: 0 });
            sl.addText(a.description || "", { x: 1.6, y: yP + 0.48, w: 7.4, h: 0.58, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.35, margin: 0 });
          });
          addLogo(pres, sl);
          break;
        }

        case "calendrier": {
          sl.addText("CALENDRIER DE PUBLICATION", { x: 0.8, y: 0.35, w: 8.4, h: 0.5, fontSize: 22, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          sl.addText("Heure de publication : " + (s.heure || "18h"), { x: 6.5, y: 0.87, w: 3.0, h: 0.3, fontSize: 10, color: C.g2, fontFace: "Arial", align: "right" });
          ["Semaine 1", "Semaine 2", "Semaine 3", "Semaine 4"].forEach((sem, i) => {
            pill(sl, 1.5 + i * 2.05, 1.25, 1.85, 0.42, sem, { bg: C.v1, color: C.w, fontSize: 11, shadow: false });
          });
          (s.plateformes || ["Instagram & Facebook", "TikTok", "YouTube Shorts"]).forEach((plat, i) => {
            const yRow = 1.85 + i * 1.0;
            sl.addText(plat, { x: 0.4, y: yRow + 0.15, w: 1.0, h: 0.7, fontSize: 9, color: C.g1, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.3 });
            [0, 1, 2, 3].forEach((week) => {
              const bars = [4, 3, 4, 3];
              for (let b = 0; b < bars[week]; b++) {
                sl.addShape(pres.shapes.RECTANGLE, { x: 1.55 + week * 2.05 + b * 0.42, y: yRow + 0.1, w: 0.32, h: 0.7, fill: { color: C.v1, transparency: 30 + b * 10 }, line: { color: C.v2, width: 0.3, transparency: 60 } });
              }
            });
          });
          addLogo(pres, sl);
          break;
        }

        case "tournage": {
          sl.addText("DÉROULEMENT DU TOURNAGE", { x: 0.8, y: 0.45, w: 8.4, h: 0.5, fontSize: 24, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2, align: "center" });
          cardShape(pres, sl, 0.8, 1.2, 8.4, 3.8);
          sl.addText(s.description || "", { x: 1.05, y: 1.35, w: 8.0, h: 3.5, fontSize: 13, color: C.g1, fontFace: "Arial", valign: "middle", lineSpacingMultiple: 1.7, align: "center" });
          addLogo(pres, sl);
          break;
        }

        case "resume": {
          sl.addText("RÉSUMÉ", { x: 0.8, y: 1.0, w: 4, h: 0.8, fontSize: 44, color: C.w, fontFace: "Arial", bold: true, charSpacing: 4 });
          sl.addText("du mois", { x: 0.8, y: 1.85, w: 4, h: 0.6, fontSize: 28, color: C.v3, fontFace: "Georgia", italic: true });
          pill(sl, 0.8, 2.7, 3, 0.35, meta.client || "", { fontSize: 10 });
          (s.points || []).forEach((pt, i) => {
            const yP = 0.7 + i * 1.45;
            pill(sl, 5.2, yP, 4.2, 0.38, pt.titre || "", { fontSize: 11 });
            sl.addText(pt.description || "", { x: 5.2, y: yP + 0.48, w: 4.2, h: 0.8, fontSize: 10, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4 });
          });
          addLogo(pres, sl);
          break;
        }

        case "contact": {
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.2, y: 1.0, w: 1.4, h: 1.4, fill: { color: C.v1 }, rectRadius: 0.25, shadow: { type: "outer", color: "000000", blur: 15, offset: 5, angle: 135, opacity: 0.3 } });
          sl.addText("C", { x: 6.2, y: 1.0, w: 1.4, h: 1.4, fontSize: 64, color: C.w, fontFace: "Arial", bold: true, align: "center", valign: "middle", margin: 0 });
          sl.addText(s.name || "CREATICAL", { x: 5.5, y: 2.7, w: 3.5, h: 0.5, fontSize: 20, color: C.w, fontFace: "Arial", bold: true });
          sl.addText(s.role || "Social Media Manager", { x: 5.5, y: 3.2, w: 3.5, h: 0.4, fontSize: 15, color: C.v3, fontFace: "Georgia", italic: true });
          sl.addText(s.email || "contact@creatical.be", { x: 5.5, y: 3.6, w: 3.5, h: 0.4, fontSize: 13, color: C.g2, fontFace: "Georgia", italic: true });
          addLogo(pres, sl);
          break;
        }

        default: {
          sl.addText((s.title || "").toUpperCase(), { x: 0.8, y: 0.45, w: 8, h: 0.5, fontSize: 24, color: C.w, fontFace: "Arial", bold: true, charSpacing: 2 });
          if (s.content) {
            cardShape(pres, sl, 0.8, 1.2, 8.4, 3.8);
            sl.addText(s.content, { x: 1.05, y: 1.3, w: 8.0, h: 3.6, fontSize: 11, color: C.g1, fontFace: "Arial", valign: "top", lineSpacingMultiple: 1.4 });
          }
          addLogo(pres, sl);
          break;
        }
      }
    }

    const base64 = await pres.write({ outputType: "base64" });
    res.json({ success: true, pptxBase64: base64, filename: `${meta.title || "Presentation"}.pptx`, slideCount: slides.length });
  } catch (error) {
    console.error("Error:", error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => { console.log(`CREATICAL PPTX v3.0 Full Strategy on port ${PORT}`); });
