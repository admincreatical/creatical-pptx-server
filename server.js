/**
 * CREATICAL PPTX SERVER — Design System v3.0
 * Node.js + pptxgenjs
 * POST /generate → { pptxBase64, filename, slideCount }
 *
 * DESIGN SYSTEM CREATICAL
 * ─────────────────────────────────────────────────────
 * Palette :
 *   BG_DARK    #0D0221  — fond principal sombre
 *   BG_MID     #160836  — fond secondaire / cards sombres
 *   BG_LIGHT   #F5F3FF  — fond clair (slides contenu)
 *   VIOLET     #7C3AED  — couleur signature
 *   VIOLET_L   #9D65F5  — violet clair (accents, highlights)
 *   VIOLET_XL  #C4A8FF  — violet très clair (texte sur dark)
 *   WHITE      #FFFFFF
 *   GRAY_1     #E9E4FF  — texte sur fond clair (titres)
 *   GRAY_2     #8B7BAB  — texte secondaire sur dark
 *   GRAY_3     #4B4666  — texte secondaire sur light
 *   ORANGE     #FF6B35  — accent chaud (CTA, highlights)
 *
 * Typographie (toujours disponible) :
 *   Titres     : Trebuchet MS Bold
 *   Corps      : Calibri
 *   Accent     : Georgia Italic
 *
 * Layout fixe :
 *   Slide      : 13.33" × 7.5" (widescreen standard)
 *   Margin     : 0.45" tous côtés
 *   Content W  : 12.43" (13.33 - 2×0.45)
 *   Content H  : 6.6" (7.5 - 2×0.45)
 *   Header H   : 1.1" (titre + barre accentuation)
 *   Body top   : 1.65" (après header)
 *   Body H     : 5.4" (7.5 - 1.65 - 0.45)
 *
 * Règles anti-débordement :
 *   - Tout texte a shrinkText:true + autoFit
 *   - Tailles max strictes par zone
 *   - Cartes : hauteur fixe, texte tronqué par CSS shrink
 */

const express = require('express');
const PptxGenJS = require('pptxgenjs');
const app = express();
app.use(express.json({ limit: '10mb' }));

// ══════════════════════════════════════════════════════
// DESIGN TOKENS
// ══════════════════════════════════════════════════════
const C = {
  BG_DARK:   '0D0221',
  BG_MID:    '160836',
  BG_CARD:   '1E0F40',
  BG_LIGHT:  'F5F3FF',
  BG_WHITE:  'FFFFFF',
  VIOLET:    '7C3AED',
  VIOLET_L:  '9D65F5',
  VIOLET_XL: 'C4A8FF',
  WHITE:     'FFFFFF',
  GRAY_1:    'E9E4FF',
  GRAY_2:    '8B7BAB',
  GRAY_3:    '4B4666',
  ORANGE:    'FF6B35',
  ORANGE_L:  'FF8C5A',
  SUCCESS:   '10B981',
  INFO:      '3B82F6',
};

const FONT = {
  TITLE:  'Trebuchet MS',
  BODY:   'Calibri',
  ACCENT: 'Georgia',
};

// Slide dimensions (inches)
const S = {
  W: 13.33, H: 7.5,
  MX: 0.45, MY: 0.45,   // margins
  get CW() { return this.W - 2 * this.MX; },  // 12.43
  HEADER_H: 1.1,
  BODY_TOP: 1.65,
  get BODY_H() { return this.H - this.BODY_TOP - this.MY; },  // 5.4
};

// ══════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════

/**
 * Truncate text to max characters + ellipsis
 */
function trunc(str, max = 120) {
  if (!str) return '';
  str = String(str);
  return str.length > max ? str.slice(0, max - 1) + '…' : str;
}

/**
 * Truncate array to max items
 */
function truncArr(arr, max = 6) {
  if (!Array.isArray(arr)) return [];
  return arr.slice(0, max);
}

/**
 * Add a standard slide header (title + accent bar + breadcrumb)
 * dark = true → dark theme, false → light theme
 */
function addHeader(slide, title, subtitle, dark = true) {
  const bg = dark ? C.BG_DARK : C.BG_LIGHT;
  const titleColor = dark ? C.WHITE : C.BG_DARK;
  const subtitleColor = dark ? C.VIOLET_XL : C.GRAY_3;

  // Violet left accent bar
  slide.addShape('rect', {
    x: S.MX, y: S.MY,
    w: 0.07, h: S.HEADER_H - 0.1,
    fill: { color: C.VIOLET },
    line: { type: 'none' },
  });

  // Title
  slide.addText(trunc(title, 60).toUpperCase(), {
    x: S.MX + 0.18, y: S.MY,
    w: S.CW - 0.18, h: 0.72,
    fontFace: FONT.TITLE,
    fontSize: 28,
    bold: true,
    color: titleColor,
    valign: 'middle',
    shrinkText: true,
  });

  // Subtitle / breadcrumb
  if (subtitle) {
    slide.addText(trunc(subtitle, 80), {
      x: S.MX + 0.18, y: S.MY + 0.72,
      w: S.CW - 0.18, h: 0.35,
      fontFace: FONT.BODY,
      fontSize: 11,
      color: subtitleColor,
      italic: true,
      valign: 'top',
      shrinkText: true,
    });
  }

  // Thin separator line
  slide.addShape('line', {
    x: S.MX + 0.18, y: S.MY + S.HEADER_H - 0.05,
    w: S.CW - 0.18, h: 0,
    line: { color: dark ? C.VIOLET : C.VIOLET_L, width: 1, dashType: 'solid' },
  });
}

/**
 * Add a card (rounded rect + content)
 * Returns nothing — use x,y,w,h to position
 */
function addCard(slide, { x, y, w, h, bg = C.BG_CARD, borderColor = null, radius = 0.12 }) {
  const opts = {
    x, y, w, h,
    fill: { color: bg },
    line: borderColor ? { color: borderColor, width: 1.5 } : { type: 'none' },
    rectRadius: radius,
  };
  slide.addShape('roundRect', opts);
}

/**
 * Add a pill/badge shape with label
 */
function addPill(slide, label, x, y, w = 2.2, h = 0.38, bg = C.VIOLET, textColor = C.WHITE) {
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: bg },
    line: { type: 'none' },
    rectRadius: 0.19,
  });
  slide.addText(trunc(label, 30), {
    x, y, w, h,
    fontFace: FONT.BODY,
    fontSize: 11,
    bold: true,
    color: textColor,
    align: 'center',
    valign: 'middle',
    shrinkText: true,
  });
}

/**
 * Add a numbered circle badge
 */
function addCircleBadge(slide, num, x, y, size = 0.45, bg = C.VIOLET) {
  slide.addShape('ellipse', {
    x, y, w: size, h: size,
    fill: { color: bg },
    line: { type: 'none' },
  });
  slide.addText(String(num), {
    x, y, w: size, h: size,
    fontFace: FONT.TITLE,
    fontSize: 14,
    bold: true,
    color: C.WHITE,
    align: 'center',
    valign: 'middle',
  });
}

/**
 * Add a Creatical "C" logo placeholder top-right
 */
function addLogo(slide, dark = true) {
  const color = dark ? C.VIOLET_L : C.VIOLET;
  slide.addText('C', {
    x: S.W - 0.75, y: 0.12,
    w: 0.5, h: 0.4,
    fontFace: FONT.TITLE,
    fontSize: 20,
    bold: true,
    color,
    align: 'center',
    valign: 'middle',
  });
  // small dot below C
  slide.addShape('ellipse', {
    x: S.W - 0.55, y: 0.52,
    w: 0.08, h: 0.08,
    fill: { color: C.ORANGE },
    line: { type: 'none' },
  });
}

// ══════════════════════════════════════════════════════
// SLIDE RENDERERS
// ══════════════════════════════════════════════════════

// ─── TITLE ───────────────────────────────────────────
function renderTitle(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);

  // Big decorative violet rectangle left
  slide.addShape('rect', {
    x: 0, y: 0,
    w: 0.22, h: S.H,
    fill: { color: C.VIOLET },
    line: { type: 'none' },
  });

  // Bottom accent strip
  slide.addShape('rect', {
    x: 0.22, y: S.H - 0.18,
    w: S.W - 0.22, h: 0.18,
    fill: { color: C.BG_MID },
    line: { type: 'none' },
  });

  // Orange dot accent
  slide.addShape('ellipse', {
    x: 1.1, y: 1.8,
    w: 0.18, h: 0.18,
    fill: { color: C.ORANGE },
    line: { type: 'none' },
  });

  // Client name (top)
  if (data.client) {
    slide.addText(trunc(data.client, 40).toUpperCase(), {
      x: 1.0, y: 0.6,
      w: 9, h: 0.45,
      fontFace: FONT.BODY,
      fontSize: 12,
      bold: true,
      color: C.VIOLET_XL,
      charSpacing: 4,
      valign: 'middle',
      shrinkText: true,
    });
  }

  // Main title — HUGE
  slide.addText(trunc(data.titre || 'STRATÉGIE DE CONTENU', 50), {
    x: 1.0, y: 1.2,
    w: 8.5, h: 2.8,
    fontFace: FONT.TITLE,
    fontSize: 58,
    bold: true,
    color: C.WHITE,
    valign: 'top',
    shrinkText: true,
    lineSpacingMultiple: 1.05,
  });

  // Subtitle
  if (data.sous_titre) {
    slide.addText(trunc(data.sous_titre, 80), {
      x: 1.0, y: 4.1,
      w: 8.5, h: 0.6,
      fontFace: FONT.ACCENT,
      fontSize: 18,
      italic: true,
      color: C.VIOLET_XL,
      valign: 'middle',
      shrinkText: true,
    });
  }

  // Date + prepared by
  const meta = [];
  if (data.date) meta.push(data.date);
  if (data.stratege) meta.push(`Stratège : ${data.stratege}`);
  if (meta.length) {
    slide.addText(meta.join('  ·  '), {
      x: 1.0, y: S.H - 0.55,
      w: 10, h: 0.35,
      fontFace: FONT.BODY,
      fontSize: 10,
      color: C.GRAY_2,
      valign: 'middle',
      shrinkText: true,
    });
  }
}

// ─── SECTION ─────────────────────────────────────────
function renderSection(pptx, slide, data) {
  slide.background = { color: C.VIOLET };
  addLogo(slide, false);

  // Dark overlay block right
  slide.addShape('rect', {
    x: S.W * 0.55, y: 0,
    w: S.W * 0.45, h: S.H,
    fill: { color: C.BG_DARK },
    line: { type: 'none' },
  });

  // Section number
  if (data.numero) {
    slide.addText(String(data.numero).padStart(2, '0'), {
      x: S.MX, y: S.MY,
      w: 2, h: 1.5,
      fontFace: FONT.TITLE,
      fontSize: 72,
      bold: true,
      color: 'FFFFFF',
      alpha: 30,
      valign: 'top',
    });
  }

  // Section title
  slide.addText(trunc(data.titre || '', 45).toUpperCase(), {
    x: S.MX, y: 1.8,
    w: S.W * 0.5,
    h: 2.5,
    fontFace: FONT.TITLE,
    fontSize: 40,
    bold: true,
    color: C.WHITE,
    valign: 'middle',
    shrinkText: true,
    lineSpacingMultiple: 1.1,
  });

  // Description
  if (data.description) {
    slide.addText(trunc(data.description, 140), {
      x: S.MX, y: 4.5,
      w: S.W * 0.5,
      h: 1.5,
      fontFace: FONT.BODY,
      fontSize: 14,
      color: 'FFFFFF',
      alpha: 80,
      valign: 'top',
      shrinkText: true,
    });
  }

  // Phases pills (right column)
  const phases = truncArr(data.phases, 5);
  if (phases.length) {
    phases.forEach((phase, i) => {
      const py = 1.2 + i * 0.85;
      addPill(slide, trunc(String(phase), 35), S.W * 0.58, py, S.W * 0.38, 0.55, C.BG_MID, C.VIOLET_XL);
    });
  }
}

// ─── CONTEXTE ────────────────────────────────────────
function renderContexte(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Contexte', data.sous_titre, true);

  const items = truncArr(data.points || [], 4);
  const cols = items.length <= 2 ? 1 : 2;
  const colW = cols === 1 ? S.CW : (S.CW - 0.25) / 2;
  const cardH = cols === 1 ? 1.1 : (S.BODY_H - 0.2) / Math.ceil(items.length / 2);

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (colW + 0.25);
    const cy = S.BODY_TOP + row * (cardH + 0.18);

    addCard(slide, { x: cx, y: cy, w: colW, h: cardH - 0.12, bg: C.BG_CARD });

    const label = typeof item === 'object' ? item.label : null;
    const text = typeof item === 'object' ? item.texte : item;

    let textY = cy + 0.12;
    let textH = cardH - 0.3;

    if (label) {
      slide.addText(trunc(label, 40), {
        x: cx + 0.18, y: cy + 0.1,
        w: colW - 0.36, h: 0.35,
        fontFace: FONT.TITLE,
        fontSize: 12,
        bold: true,
        color: C.VIOLET_L,
        shrinkText: true,
      });
      textY = cy + 0.45;
      textH = cardH - 0.55;
    }

    slide.addText(trunc(String(text || ''), 180), {
      x: cx + 0.18, y: textY,
      w: colW - 0.36, h: textH,
      fontFace: FONT.BODY,
      fontSize: 13,
      color: C.GRAY_1,
      valign: 'top',
      shrinkText: true,
    });
  });
}

// ─── MARCHE / GÉOGRAPHIE ──────────────────────────────
function renderMarche(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Marché', data.sous_titre, true);

  // Big stat left
  if (data.stat_principal) {
    addCard(slide, { x: S.MX, y: S.BODY_TOP, w: 3.8, h: S.BODY_H, bg: C.VIOLET });
    slide.addText(trunc(data.stat_principal.valeur || '', 12), {
      x: S.MX, y: S.BODY_TOP + 0.6,
      w: 3.8, h: 1.8,
      fontFace: FONT.TITLE,
      fontSize: 64,
      bold: true,
      color: C.WHITE,
      align: 'center',
      valign: 'middle',
      shrinkText: true,
    });
    slide.addText(trunc(data.stat_principal.label || '', 50), {
      x: S.MX, y: S.BODY_TOP + 2.4,
      w: 3.8, h: 0.6,
      fontFace: FONT.BODY,
      fontSize: 13,
      color: C.VIOLET_XL,
      align: 'center',
      valign: 'middle',
      shrinkText: true,
    });
  }

  // Points right
  const points = truncArr(data.points || [], 4);
  const rightX = S.MX + 4.1;
  const rightW = S.CW - 4.1;
  const cardH = (S.BODY_H - 0.15 * (points.length - 1)) / Math.max(points.length, 1);

  points.forEach((pt, i) => {
    const cy = S.BODY_TOP + i * (cardH + 0.15);
    addCard(slide, { x: rightX, y: cy, w: rightW, h: cardH, bg: C.BG_CARD });
    addCircleBadge(slide, i + 1, rightX + 0.15, cy + (cardH - 0.45) / 2);
    const label = typeof pt === 'object' ? pt.label : null;
    const text = typeof pt === 'object' ? pt.texte : pt;
    if (label) {
      slide.addText(trunc(label, 45), {
        x: rightX + 0.72, y: cy + 0.1,
        w: rightW - 0.88, h: 0.35,
        fontFace: FONT.TITLE, fontSize: 12, bold: true,
        color: C.VIOLET_L, shrinkText: true,
      });
    }
    slide.addText(trunc(String(text || ''), 130), {
      x: rightX + 0.72, y: label ? cy + 0.42 : cy + 0.12,
      w: rightW - 0.88, h: label ? cardH - 0.55 : cardH - 0.24,
      fontFace: FONT.BODY, fontSize: 12,
      color: C.GRAY_1, valign: 'top', shrinkText: true,
    });
  });
}

// ─── GÉOGRAPHIE ──────────────────────────────────────
function renderGeographie(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Géographie cible', data.sous_titre, true);

  const zones = truncArr(data.zones || [], 6);
  const perRow = zones.length <= 3 ? zones.length : 3;
  const rows = Math.ceil(zones.length / perRow);
  const cardW = (S.CW - 0.2 * (perRow - 1)) / perRow;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  zones.forEach((zone, i) => {
    const col = i % perRow;
    const row = Math.floor(i / perRow);
    const cx = S.MX + col * (cardW + 0.2);
    const cy = S.BODY_TOP + row * (cardH + 0.2);
    const isBg = zone.principal || i === 0;

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: isBg ? C.VIOLET : C.BG_CARD });

    slide.addText(zone.emoji || '📍', {
      x: cx + 0.15, y: cy + 0.15,
      w: 0.5, h: 0.4,
      fontSize: 18, align: 'center',
    });

    slide.addText(trunc(zone.nom || String(zone), 25), {
      x: cx + 0.7, y: cy + 0.18,
      w: cardW - 0.85, h: 0.4,
      fontFace: FONT.TITLE, fontSize: 14, bold: true,
      color: C.WHITE, valign: 'middle', shrinkText: true,
    });

    if (zone.description) {
      slide.addText(trunc(zone.description, 90), {
        x: cx + 0.15, y: cy + 0.65,
        w: cardW - 0.3, h: cardH - 0.8,
        fontFace: FONT.BODY, fontSize: 11,
        color: isBg ? C.VIOLET_XL : C.GRAY_1,
        valign: 'top', shrinkText: true,
      });
    }
  });
}

// ─── CONCURRENCE ─────────────────────────────────────
function renderConcurrence(pptx, slide, data) {
  slide.background = { color: C.BG_LIGHT };
  addLogo(slide, false);
  addHeader(slide, data.titre || 'Analyse Concurrentielle', data.sous_titre, false);

  const concurrents = truncArr(data.concurrents || [], 4);
  const colW = (S.CW - 0.25 * (concurrents.length - 1)) / Math.max(concurrents.length, 1);
  const cardH = S.BODY_H;

  concurrents.forEach((c, i) => {
    const cx = S.MX + i * (colW + 0.25);
    const isClient = c.client || i === 0;
    const cardBg = isClient ? C.VIOLET : C.BG_WHITE;
    const border = isClient ? null : C.VIOLET_L;

    addCard(slide, { x: cx, y: S.BODY_TOP, w: colW, h: cardH, bg: cardBg, borderColor: border });

    // Name
    slide.addText(trunc(c.nom || '', 25), {
      x: cx + 0.15, y: S.BODY_TOP + 0.18,
      w: colW - 0.3, h: 0.5,
      fontFace: FONT.TITLE, fontSize: 15, bold: true,
      color: isClient ? C.WHITE : C.BG_DARK,
      valign: 'middle', shrinkText: true,
    });

    if (isClient) {
      slide.addText('VOUS', {
        x: cx + 0.15, y: S.BODY_TOP + 0.65,
        w: colW - 0.3, h: 0.3,
        fontFace: FONT.BODY, fontSize: 9, bold: true,
        color: C.ORANGE, charSpacing: 2, shrinkText: true,
      });
    }

    // Stats or points
    const items = truncArr(c.points || [], 5);
    items.forEach((pt, j) => {
      const py = S.BODY_TOP + (isClient ? 1.1 : 0.85) + j * 0.78;
      addCard(slide, {
        x: cx + 0.1, y: py,
        w: colW - 0.2, h: 0.65,
        bg: isClient ? C.BG_MID : C.BG_LIGHT,
      });
      slide.addText(trunc(String(pt), 70), {
        x: cx + 0.2, y: py + 0.08,
        w: colW - 0.4, h: 0.5,
        fontFace: FONT.BODY, fontSize: 11,
        color: isClient ? C.GRAY_1 : C.GRAY_3,
        valign: 'top', shrinkText: true,
      });
    });
  });
}

// ─── NIVEAU CONSCIENCE ───────────────────────────────
function renderNiveauConscience(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Niveau de Conscience', data.sous_titre, true);

  const niveaux = truncArr(data.niveaux || [], 5);
  const totalH = S.BODY_H - 0.1;
  const heights = niveaux.map((n, i) => {
    const weight = niveaux.length - i; // bottom = largest
    return weight;
  });
  const totalW = heights.reduce((a, b) => a + b, 0);

  let currentX = S.MX;
  niveaux.forEach((niveau, i) => {
    const w = (heights[i] / totalW) * S.CW;
    const isActive = niveau.actuel || niveau.active;
    const bg = isActive ? C.VIOLET : (i % 2 === 0 ? C.BG_CARD : C.BG_MID);

    addCard(slide, { x: currentX, y: S.BODY_TOP, w: w - 0.1, h: totalH, bg, radius: 0.1 });

    // Number
    slide.addText(String(i + 1), {
      x: currentX + 0.1, y: S.BODY_TOP + 0.15,
      w: w - 0.2, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 22, bold: true,
      color: isActive ? C.WHITE : C.VIOLET_L,
      align: 'center', shrinkText: true,
    });

    // Label
    slide.addText(trunc(niveau.nom || niveau.label || '', 20), {
      x: currentX + 0.1, y: S.BODY_TOP + 0.65,
      w: w - 0.2, h: 0.55,
      fontFace: FONT.TITLE, fontSize: 11, bold: true,
      color: C.WHITE, align: 'center', shrinkText: true,
    });

    // Description
    if (niveau.description) {
      slide.addText(trunc(niveau.description, 80), {
        x: currentX + 0.1, y: S.BODY_TOP + 1.3,
        w: w - 0.2, h: totalH - 1.4,
        fontFace: FONT.BODY, fontSize: 10,
        color: isActive ? C.VIOLET_XL : C.GRAY_2,
        align: 'center', valign: 'top', shrinkText: true,
      });
    }

    currentX += w;
  });
}

// ─── SOPHISTICATION MARCHÉ ───────────────────────────
function renderSophisticationMarche(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Sophistication Marché', data.sous_titre, true);

  // Big gauge visual
  const niveau = Math.min(Math.max(parseInt(data.niveau) || 3, 1), 5);
  const gaugeFill = niveau / 5;

  // Gauge background
  slide.addShape('roundRect', {
    x: S.MX, y: S.BODY_TOP,
    w: S.CW, h: 0.55,
    fill: { color: C.BG_CARD },
    line: { type: 'none' },
    rectRadius: 0.27,
  });
  // Gauge fill
  slide.addShape('roundRect', {
    x: S.MX, y: S.BODY_TOP,
    w: S.CW * gaugeFill, h: 0.55,
    fill: { color: C.VIOLET },
    line: { type: 'none' },
    rectRadius: 0.27,
  });

  // Level labels
  ['Naïf', 'Basique', 'Informé', 'Averti', 'Expert'].forEach((label, i) => {
    slide.addText(label, {
      x: S.MX + (S.CW / 5) * i, y: S.BODY_TOP + 0.62,
      w: S.CW / 5, h: 0.3,
      fontFace: FONT.BODY, fontSize: 10,
      color: i < niveau ? C.VIOLET_XL : C.GRAY_2,
      align: 'center', shrinkText: true,
    });
  });

  // Current level callout
  slide.addText(`Niveau ${niveau}/5`, {
    x: S.MX, y: S.BODY_TOP + 1.05,
    w: 3, h: 0.6,
    fontFace: FONT.TITLE, fontSize: 24, bold: true,
    color: C.VIOLET_L, shrinkText: true,
  });

  // Implications
  const implications = truncArr(data.implications || [], 3);
  implications.forEach((impl, i) => {
    const cy = S.BODY_TOP + 1.85 + i * 1.1;
    addCard(slide, { x: S.MX, y: cy, w: S.CW, h: 0.95, bg: C.BG_CARD });
    addCircleBadge(slide, i + 1, S.MX + 0.18, cy + 0.25, 0.42, C.VIOLET);
    slide.addText(trunc(String(impl), 150), {
      x: S.MX + 0.75, y: cy + 0.1,
      w: S.CW - 0.9, h: 0.78,
      fontFace: FONT.BODY, fontSize: 13,
      color: C.GRAY_1, valign: 'middle', shrinkText: true,
    });
  });
}

// ─── OBJECTIFS ───────────────────────────────────────
function renderObjectifs(pptx, slide, data) {
  slide.background = { color: C.BG_LIGHT };
  addLogo(slide, false);
  addHeader(slide, data.titre || 'Objectifs', data.sous_titre, false);

  const objectifs = truncArr(data.objectifs || [], 4);
  const cols = objectifs.length <= 2 ? objectifs.length : 2;
  const colW = (S.CW - 0.3) / cols;
  const rows = Math.ceil(objectifs.length / cols);
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  objectifs.forEach((obj, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (colW + 0.3);
    const cy = S.BODY_TOP + row * (cardH + 0.2);

    addCard(slide, { x: cx, y: cy, w: colW, h: cardH, bg: C.BG_WHITE, borderColor: C.VIOLET_L });

    // Number badge
    addCircleBadge(slide, i + 1, cx + 0.2, cy + 0.2, 0.45, C.VIOLET);

    // Title
    slide.addText(trunc(obj.titre || obj.label || '', 40), {
      x: cx + 0.8, y: cy + 0.2,
      w: colW - 1.0, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 14, bold: true,
      color: C.BG_DARK, valign: 'middle', shrinkText: true,
    });

    // Description
    if (obj.description || obj.texte) {
      slide.addText(trunc(obj.description || obj.texte, 150), {
        x: cx + 0.2, y: cy + 0.82,
        w: colW - 0.4, h: cardH - 1.1,
        fontFace: FONT.BODY, fontSize: 12,
        color: C.GRAY_3, valign: 'top', shrinkText: true,
      });
    }

    // KPI badge if present
    if (obj.kpi) {
      addPill(slide, trunc(obj.kpi, 25), cx + 0.2, cy + cardH - 0.52, colW - 0.4, 0.38, C.VIOLET, C.WHITE);
    }
  });
}

// ─── BRAND ARCHETYPE ─────────────────────────────────
function renderBrandArchetype(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Archétype de Marque', data.sous_titre, true);

  // Big archetype display
  const archetype = data.archetype || {};

  addCard(slide, { x: S.MX, y: S.BODY_TOP, w: 4.5, h: S.BODY_H, bg: C.VIOLET });

  slide.addText(archetype.emoji || '🏛️', {
    x: S.MX, y: S.BODY_TOP + 0.4,
    w: 4.5, h: 1.2,
    fontSize: 56, align: 'center',
  });

  slide.addText(trunc(archetype.nom || '', 25).toUpperCase(), {
    x: S.MX + 0.2, y: S.BODY_TOP + 1.7,
    w: 4.1, h: 0.7,
    fontFace: FONT.TITLE, fontSize: 22, bold: true,
    color: C.WHITE, align: 'center', shrinkText: true,
  });

  slide.addText(trunc(archetype.slogan || archetype.description || '', 80), {
    x: S.MX + 0.2, y: S.BODY_TOP + 2.5,
    w: 4.1, h: 1.5,
    fontFace: FONT.ACCENT, fontSize: 14, italic: true,
    color: C.VIOLET_XL, align: 'center', valign: 'top', shrinkText: true,
  });

  // Traits right side
  const traits = truncArr(data.traits || archetype.traits || [], 5);
  const rightX = S.MX + 4.8;
  const rightW = S.CW - 4.8;
  const traitH = S.BODY_H / Math.max(traits.length, 1);

  traits.forEach((trait, i) => {
    const cy = S.BODY_TOP + i * (traitH + 0.08);
    addCard(slide, { x: rightX, y: cy, w: rightW, h: traitH - 0.08, bg: C.BG_CARD });
    slide.addText(trunc(String(trait), 100), {
      x: rightX + 0.25, y: cy + 0.1,
      w: rightW - 0.4, h: traitH - 0.25,
      fontFace: FONT.BODY, fontSize: 13,
      color: C.GRAY_1, valign: 'middle', shrinkText: true,
    });
  });
}

// ─── PERSONA ─────────────────────────────────────────
function renderPersona(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Persona Cible', data.sous_titre, true);

  const persona = data.persona || {};

  // Avatar area
  addCard(slide, { x: S.MX, y: S.BODY_TOP, w: 3.2, h: S.BODY_H, bg: C.BG_CARD });

  slide.addText(persona.emoji || '👤', {
    x: S.MX, y: S.BODY_TOP + 0.3,
    w: 3.2, h: 1.1,
    fontSize: 48, align: 'center',
  });

  slide.addText(trunc(persona.prenom || persona.nom || 'Le Persona', 20), {
    x: S.MX + 0.15, y: S.BODY_TOP + 1.5,
    w: 2.9, h: 0.5,
    fontFace: FONT.TITLE, fontSize: 18, bold: true,
    color: C.WHITE, align: 'center', shrinkText: true,
  });

  if (persona.age || persona.profession) {
    slide.addText([persona.age, persona.profession].filter(Boolean).join(' · '), {
      x: S.MX + 0.15, y: S.BODY_TOP + 2.05,
      w: 2.9, h: 0.35,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.VIOLET_XL, align: 'center', shrinkText: true,
    });
  }

  // Quick stats
  const stats = truncArr(persona.stats || [], 3);
  stats.forEach((stat, i) => {
    const sy = S.BODY_TOP + 2.6 + i * 0.85;
    addCard(slide, { x: S.MX + 0.2, y: sy, w: 2.8, h: 0.72, bg: C.VIOLET });
    slide.addText(trunc(String(stat), 50), {
      x: S.MX + 0.35, y: sy + 0.1,
      w: 2.5, h: 0.54,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.WHITE, valign: 'middle', align: 'center', shrinkText: true,
    });
  });

  // Right: traits, douleurs, désirs
  const rightX = S.MX + 3.45;
  const rightW = S.CW - 3.45;
  const sections = [
    { label: '🎯 Douleurs', items: persona.douleurs, color: C.ORANGE },
    { label: '✨ Désirs', items: persona.desirs, color: C.SUCCESS },
    { label: '💬 Objections', items: persona.objections, color: C.INFO },
  ];

  let currentY = S.BODY_TOP;
  sections.forEach((sec) => {
    const items = truncArr(sec.items || [], 3);
    if (!items.length) return;

    const sectionH = 0.38 + items.length * 0.55;
    if (currentY + sectionH > S.H - S.MY) return;

    slide.addText(sec.label, {
      x: rightX, y: currentY,
      w: rightW, h: 0.32,
      fontFace: FONT.TITLE, fontSize: 12, bold: true,
      color: sec.color, shrinkText: true,
    });

    items.forEach((item, j) => {
      const iy = currentY + 0.35 + j * 0.55;
      addCard(slide, { x: rightX, y: iy, w: rightW, h: 0.48, bg: C.BG_CARD });
      slide.addText(trunc(String(item), 90), {
        x: rightX + 0.15, y: iy + 0.06,
        w: rightW - 0.25, h: 0.38,
        fontFace: FONT.BODY, fontSize: 11,
        color: C.GRAY_1, valign: 'middle', shrinkText: true,
      });
    });

    currentY += sectionH + 0.2;
  });
}

// ─── POSITIONNEMENT ──────────────────────────────────
function renderPositionnement(pptx, slide, data) {
  slide.background = { color: C.BG_LIGHT };
  addLogo(slide, false);
  addHeader(slide, data.titre || 'Positionnement', data.sous_titre, false);

  const elements = truncArr(data.elements || [], 5);
  const cols = elements.length <= 3 ? elements.length : Math.ceil(elements.length / 2);
  const rows = Math.ceil(elements.length / cols);
  const cardW = (S.CW - 0.2 * (cols - 1)) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  elements.forEach((el, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.2);
    const cy = S.BODY_TOP + row * (cardH + 0.2);
    const colors = [C.VIOLET, C.ORANGE, C.SUCCESS, C.INFO, C.VIOLET_L];
    const accent = colors[i % colors.length];

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_WHITE, borderColor: accent });

    // Top accent bar
    slide.addShape('rect', {
      x: cx, y: cy,
      w: cardW, h: 0.08,
      fill: { color: accent },
      line: { type: 'none' },
    });

    slide.addText(el.emoji || '', {
      x: cx + 0.15, y: cy + 0.18,
      w: 0.55, h: 0.55,
      fontSize: 22,
    });

    slide.addText(trunc(el.label || el.titre || '', 30), {
      x: cx + 0.15, y: cy + 0.82,
      w: cardW - 0.3, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.BG_DARK, shrinkText: true,
    });

    slide.addText(trunc(el.valeur || el.description || '', 100), {
      x: cx + 0.15, y: cy + 1.28,
      w: cardW - 0.3, h: cardH - 1.45,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_3, valign: 'top', shrinkText: true,
    });
  });
}

// ─── PLATEFORMES ─────────────────────────────────────
function renderPlateformes(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Plateformes', data.sous_titre, true);

  const plateformes = truncArr(data.plateformes || [], 5);
  const colW = S.CW / Math.max(plateformes.length, 1) - 0.15;

  plateformes.forEach((p, i) => {
    const cx = S.MX + i * (colW + 0.15);
    const isPrimary = p.principale || p.priorite === 1 || i === 0;
    const bg = isPrimary ? C.VIOLET : C.BG_CARD;

    addCard(slide, { x: cx, y: S.BODY_TOP, w: colW, h: S.BODY_H, bg });

    slide.addText(p.emoji || p.icone || '📱', {
      x: cx, y: S.BODY_TOP + 0.25,
      w: colW, h: 0.9,
      fontSize: 36, align: 'center',
    });

    slide.addText(trunc(p.nom || '', 20), {
      x: cx + 0.1, y: S.BODY_TOP + 1.25,
      w: colW - 0.2, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 14, bold: true,
      color: C.WHITE, align: 'center', shrinkText: true,
    });

    if (p.frequence) {
      addPill(slide, trunc(p.frequence, 20), cx + 0.15, S.BODY_TOP + 1.82, colW - 0.3, 0.35,
        isPrimary ? C.BG_MID : C.BG_MID, C.VIOLET_XL);
    }

    const details = truncArr(p.details || p.formats || [], 4);
    details.forEach((d, j) => {
      slide.addText('• ' + trunc(String(d), 45), {
        x: cx + 0.15, y: S.BODY_TOP + 2.35 + j * 0.55,
        w: colW - 0.3, h: 0.48,
        fontFace: FONT.BODY, fontSize: 10,
        color: isPrimary ? C.VIOLET_XL : C.GRAY_1,
        valign: 'top', shrinkText: true,
      });
    });
  });
}

// ─── KPIs ─────────────────────────────────────────────
function renderKpis(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'KPIs & Objectifs', data.sous_titre, true);

  const kpis = truncArr(data.kpis || [], 6);
  const cols = kpis.length <= 3 ? kpis.length : 3;
  const rows = Math.ceil(kpis.length / cols);
  const cardW = (S.CW - 0.2 * (cols - 1)) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  kpis.forEach((kpi, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.2);
    const cy = S.BODY_TOP + row * (cardH + 0.2);

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_CARD });

    // Accent top
    slide.addShape('rect', {
      x: cx, y: cy,
      w: 0.06, h: cardH,
      fill: { color: C.VIOLET },
      line: { type: 'none' },
    });

    // Metric value (big)
    if (kpi.valeur || kpi.objectif) {
      slide.addText(trunc(String(kpi.valeur || kpi.objectif || ''), 12), {
        x: cx + 0.22, y: cy + 0.15,
        w: cardW - 0.35, h: 0.8,
        fontFace: FONT.TITLE, fontSize: 32, bold: true,
        color: C.VIOLET_L, shrinkText: true,
      });
    }

    // KPI name
    slide.addText(trunc(kpi.nom || kpi.label || '', 40), {
      x: cx + 0.22, y: cy + 0.95,
      w: cardW - 0.35, h: 0.4,
      fontFace: FONT.TITLE, fontSize: 12, bold: true,
      color: C.WHITE, shrinkText: true,
    });

    // Description
    if (kpi.description) {
      slide.addText(trunc(kpi.description, 80), {
        x: cx + 0.22, y: cy + 1.38,
        w: cardW - 0.35, h: cardH - 1.52,
        fontFace: FONT.BODY, fontSize: 10,
        color: C.GRAY_2, valign: 'top', shrinkText: true,
      });
    }
  });
}

// ─── PILIERS ─────────────────────────────────────────
function renderPiliers(pptx, slide, data) {
  slide.background = { color: C.BG_LIGHT };
  addLogo(slide, false);
  addHeader(slide, data.titre || 'Piliers de Marque', data.sous_titre, false);

  const piliers = truncArr(data.piliers || [], 4);
  const cols = piliers.length <= 2 ? piliers.length : 2;
  const rows = Math.ceil(piliers.length / cols);
  const cardW = (S.CW - 0.3) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;
  const accentColors = [C.VIOLET, C.ORANGE, C.SUCCESS, C.INFO];

  piliers.forEach((pilier, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.3);
    const cy = S.BODY_TOP + row * (cardH + 0.2);
    const accent = accentColors[i % accentColors.length];

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_WHITE, borderColor: accent });

    // Top full-width accent
    slide.addShape('rect', {
      x: cx, y: cy,
      w: cardW, h: 0.12,
      fill: { color: accent },
      line: { type: 'none' },
    });

    slide.addText(pilier.emoji || `0${i + 1}`, {
      x: cx + 0.2, y: cy + 0.22,
      w: 0.7, h: 0.6,
      fontSize: 26, shrinkText: true,
    });

    slide.addText(trunc(pilier.nom || pilier.titre || '', 35), {
      x: cx + 0.2, y: cy + 0.92,
      w: cardW - 0.4, h: 0.5,
      fontFace: FONT.TITLE, fontSize: 15, bold: true,
      color: C.BG_DARK, shrinkText: true,
    });

    slide.addText(trunc(pilier.description || pilier.texte || '', 160), {
      x: cx + 0.2, y: cy + 1.5,
      w: cardW - 0.4, h: cardH - 1.68,
      fontFace: FONT.BODY, fontSize: 12,
      color: C.GRAY_3, valign: 'top', shrinkText: true,
    });
  });
}

// ─── PILIERS CONTENU ─────────────────────────────────
function renderPiliersContenu(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Piliers de Contenu', data.sous_titre, true);

  const piliers = truncArr(data.piliers || [], 5);
  const isHorizontal = piliers.length <= 3;

  if (isHorizontal) {
    const cardW = (S.CW - 0.2 * (piliers.length - 1)) / piliers.length;
    piliers.forEach((pilier, i) => {
      const cx = S.MX + i * (cardW + 0.2);
      const accent = [C.VIOLET, C.ORANGE, C.SUCCESS, C.INFO, C.VIOLET_L][i];
      addCard(slide, { x: cx, y: S.BODY_TOP, w: cardW, h: S.BODY_H, bg: C.BG_CARD });
      slide.addShape('rect', { x: cx, y: S.BODY_TOP, w: cardW, h: 0.1, fill: { color: accent }, line: { type: 'none' } });
      slide.addText(String(i + 1), {
        x: cx + 0.15, y: S.BODY_TOP + 0.2,
        w: 0.6, h: 0.6,
        fontFace: FONT.TITLE, fontSize: 28, bold: true,
        color: accent, shrinkText: true,
      });
      slide.addText(trunc(pilier.nom || pilier.titre || '', 25), {
        x: cx + 0.15, y: S.BODY_TOP + 0.9,
        w: cardW - 0.3, h: 0.55,
        fontFace: FONT.TITLE, fontSize: 14, bold: true,
        color: C.WHITE, shrinkText: true,
      });
      if (pilier.pourcentage) {
        addPill(slide, pilier.pourcentage, cx + 0.15, S.BODY_TOP + 1.55, cardW - 0.3, 0.38, accent, C.WHITE);
      }
      slide.addText(trunc(pilier.description || '', 160), {
        x: cx + 0.15, y: S.BODY_TOP + 2.05,
        w: cardW - 0.3, h: S.BODY_H - 2.2,
        fontFace: FONT.BODY, fontSize: 11,
        color: C.GRAY_1, valign: 'top', shrinkText: true,
      });
    });
  } else {
    // 4-5 piliers: horizontal rows
    const cardH = (S.BODY_H - 0.15 * (piliers.length - 1)) / piliers.length;
    piliers.forEach((pilier, i) => {
      const cy = S.BODY_TOP + i * (cardH + 0.15);
      const accent = [C.VIOLET, C.ORANGE, C.SUCCESS, C.INFO, C.VIOLET_L][i];
      addCard(slide, { x: S.MX, y: cy, w: S.CW, h: cardH, bg: C.BG_CARD });
      slide.addShape('rect', { x: S.MX, y: cy, w: 0.08, h: cardH, fill: { color: accent }, line: { type: 'none' } });
      addCircleBadge(slide, i + 1, S.MX + 0.22, cy + (cardH - 0.42) / 2, 0.42, accent);
      slide.addText(trunc(pilier.nom || pilier.titre || '', 30), {
        x: S.MX + 0.82, y: cy + 0.08,
        w: 3.5, h: cardH - 0.16,
        fontFace: FONT.TITLE, fontSize: 14, bold: true,
        color: C.WHITE, valign: 'middle', shrinkText: true,
      });
      slide.addText(trunc(pilier.description || '', 150), {
        x: S.MX + 4.5, y: cy + 0.08,
        w: S.CW - 4.65, h: cardH - 0.16,
        fontFace: FONT.BODY, fontSize: 11,
        color: C.GRAY_1, valign: 'middle', shrinkText: true,
      });
    });
  }
}

// ─── CONTENU STRATÉGIQUE ─────────────────────────────
function renderContenuStrategique(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Contenu Stratégique', data.sous_titre, true);

  const formats = truncArr(data.formats || [], 6);
  const cols = Math.min(formats.length, 3);
  const rows = Math.ceil(formats.length / cols);
  const cardW = (S.CW - 0.2 * (cols - 1)) / cols;
  const cardH = (S.BODY_H - 0.15 * (rows - 1)) / rows;

  formats.forEach((fmt, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.2);
    const cy = S.BODY_TOP + row * (cardH + 0.15);

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_CARD });

    slide.addText(fmt.emoji || '🎬', {
      x: cx + 0.15, y: cy + 0.12,
      w: 0.6, h: 0.55,
      fontSize: 22,
    });

    slide.addText(trunc(fmt.type || fmt.nom || '', 30), {
      x: cx + 0.8, y: cy + 0.15,
      w: cardW - 0.95, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 12, bold: true,
      color: C.VIOLET_L, shrinkText: true,
    });

    slide.addText(trunc(fmt.description || '', 100), {
      x: cx + 0.15, y: cy + 0.72,
      w: cardW - 0.3, h: cardH - 0.88,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_1, valign: 'top', shrinkText: true,
    });
  });
}

// ─── MÉTHODE ─────────────────────────────────────────
function renderMethode(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Notre Méthode', data.sous_titre, true);

  const etapes = truncArr(data.etapes || [], 5);
  const cols = etapes.length <= 3 ? etapes.length : Math.ceil(etapes.length / 2);
  const rows = Math.ceil(etapes.length / cols);
  const cardW = (S.CW - 0.25 * (cols - 1)) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  etapes.forEach((etape, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.25);
    const cy = S.BODY_TOP + row * (cardH + 0.2);

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_CARD });
    addCircleBadge(slide, i + 1, cx + 0.2, cy + 0.18, 0.45, C.VIOLET);

    slide.addText(trunc(etape.titre || etape.nom || '', 35), {
      x: cx + 0.82, y: cy + 0.18,
      w: cardW - 0.98, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.WHITE, valign: 'middle', shrinkText: true,
    });

    slide.addText(trunc(etape.description || etape.texte || '', 140), {
      x: cx + 0.2, y: cy + 0.78,
      w: cardW - 0.4, h: cardH - 0.95,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_1, valign: 'top', shrinkText: true,
    });

    // Arrow between steps (horizontal)
    if (rows === 1 && i < etapes.length - 1) {
      slide.addText('→', {
        x: cx + cardW + 0.02, y: cy + cardH / 2 - 0.2,
        w: 0.2, h: 0.4,
        fontFace: FONT.TITLE, fontSize: 16,
        color: C.VIOLET_L, align: 'center',
      });
    }
  });
}

// ─── HOOK RÈGLES ─────────────────────────────────────
function renderHookRegles(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Les Règles du Hook', data.sous_titre, true);

  const regles = truncArr(data.regles || [], 5);
  const cardH = (S.BODY_H - 0.15 * (regles.length - 1)) / Math.max(regles.length, 1);

  regles.forEach((regle, i) => {
    const cy = S.BODY_TOP + i * (cardH + 0.15);
    addCard(slide, { x: S.MX, y: cy, w: S.CW, h: cardH, bg: C.BG_CARD });

    // Number
    addCircleBadge(slide, i + 1, S.MX + 0.18, cy + (cardH - 0.44) / 2, 0.44, C.ORANGE);

    // Rule title
    slide.addText(trunc(regle.titre || regle.nom || '', 45), {
      x: S.MX + 0.78, y: cy + 0.08,
      w: 4.5, h: cardH - 0.16,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.WHITE, valign: 'middle', shrinkText: true,
    });

    // Rule description
    if (regle.description || regle.exemple) {
      slide.addText(trunc(regle.description || regle.exemple, 150), {
        x: S.MX + 5.5, y: cy + 0.08,
        w: S.CW - 5.65, h: cardH - 0.16,
        fontFace: FONT.BODY, fontSize: 11,
        color: C.GRAY_1, valign: 'middle', shrinkText: true,
      });
    }
  });
}

// ─── EMOTION ─────────────────────────────────────────
function renderEmotion(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Marketing Émotionnel', data.sous_titre, true);

  const emotions = truncArr(data.emotions || [], 6);
  const cols = Math.min(emotions.length, 3);
  const rows = Math.ceil(emotions.length / cols);
  const cardW = (S.CW - 0.2 * (cols - 1)) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  emotions.forEach((emotion, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.2);
    const cy = S.BODY_TOP + row * (cardH + 0.2);
    const accents = [C.ORANGE, C.VIOLET, C.SUCCESS, C.INFO, C.ORANGE_L, C.VIOLET_L];
    const accent = accents[i % accents.length];

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_CARD });

    slide.addText(emotion.emoji || '❤️', {
      x: cx, y: cy + 0.15,
      w: cardW, h: 0.7,
      fontSize: 28, align: 'center',
    });

    slide.addText(trunc(emotion.nom || '', 20), {
      x: cx + 0.1, y: cy + 0.92,
      w: cardW - 0.2, h: 0.4,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: accent, align: 'center', shrinkText: true,
    });

    slide.addText(trunc(emotion.description || '', 90), {
      x: cx + 0.1, y: cy + 1.38,
      w: cardW - 0.2, h: cardH - 1.52,
      fontFace: FONT.BODY, fontSize: 10,
      color: C.GRAY_1, align: 'center', valign: 'top', shrinkText: true,
    });
  });
}

// ─── FUNNEL IDÉES ────────────────────────────────────
function renderFunnelIdees(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Idées de Contenu', data.sous_titre, true);

  const phases = truncArr(data.phases || [], 3);
  const colW = (S.CW - 0.2 * (phases.length - 1)) / Math.max(phases.length, 1);
  const funnelColors = [C.VIOLET_L, C.VIOLET, C.BG_CARD];

  phases.forEach((phase, i) => {
    const cx = S.MX + i * (colW + 0.2);
    const bg = funnelColors[i % funnelColors.length] || C.BG_CARD;

    addCard(slide, { x: cx, y: S.BODY_TOP, w: colW, h: 0.55, bg });
    slide.addText(trunc(phase.nom || phase.label || '', 25).toUpperCase(), {
      x: cx, y: S.BODY_TOP,
      w: colW, h: 0.55,
      fontFace: FONT.TITLE, fontSize: 12, bold: true,
      color: C.WHITE, align: 'center', valign: 'middle', shrinkText: true,
    });

    const idees = truncArr(phase.idees || [], 5);
    idees.forEach((idee, j) => {
      const iy = S.BODY_TOP + 0.7 + j * 0.92;
      addCard(slide, { x: cx, y: iy, w: colW, h: 0.8, bg: C.BG_CARD });
      slide.addText(idee.emoji || '💡', {
        x: cx + 0.1, y: iy + 0.1,
        w: 0.45, h: 0.6, fontSize: 16,
      });
      slide.addText(trunc(idee.titre || String(idee), 55), {
        x: cx + 0.6, y: iy + 0.08,
        w: colW - 0.72, h: 0.65,
        fontFace: FONT.BODY, fontSize: 11,
        color: C.GRAY_1, valign: 'middle', shrinkText: true,
      });
    });
  });
}

// ─── PHASE FUNNEL ────────────────────────────────────
function renderPhaseFunnel(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Phase Funnel', data.sous_titre, true);

  const etapes = truncArr(data.etapes || [], 4);
  const cardW = S.CW / Math.max(etapes.length, 1) - 0.15;

  etapes.forEach((etape, i) => {
    const cx = S.MX + i * (cardW + 0.15);
    const funnelH = S.BODY_H * (1 - i * 0.08);
    const cy = S.BODY_TOP + (S.BODY_H - funnelH);
    const alpha = 100 - i * 15;

    addCard(slide, { x: cx, y: cy, w: cardW, h: funnelH, bg: C.VIOLET });

    slide.addText(trunc(etape.nom || etape.titre || '', 18).toUpperCase(), {
      x: cx + 0.1, y: cy + 0.2,
      w: cardW - 0.2, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 12, bold: true,
      color: C.WHITE, align: 'center', shrinkText: true,
    });

    if (etape.pourcentage || etape.volume) {
      slide.addText(String(etape.pourcentage || etape.volume), {
        x: cx + 0.1, y: cy + 0.72,
        w: cardW - 0.2, h: 0.6,
        fontFace: FONT.TITLE, fontSize: 22, bold: true,
        color: C.VIOLET_XL, align: 'center', shrinkText: true,
      });
    }

    slide.addText(trunc(etape.description || '', 80), {
      x: cx + 0.1, y: cy + 1.4,
      w: cardW - 0.2, h: funnelH - 1.6,
      fontFace: FONT.BODY, fontSize: 10,
      color: C.VIOLET_XL, align: 'center', valign: 'top', shrinkText: true,
    });
  });
}

// ─── PLAN ACTION ─────────────────────────────────────
function renderPlanAction(pptx, slide, data) {
  slide.background = { color: C.BG_LIGHT };
  addLogo(slide, false);
  addHeader(slide, data.titre || "Plan d'Action", data.sous_titre, false);

  const actions = truncArr(data.actions || [], 6);
  const cols = 2;
  const rows = Math.ceil(actions.length / cols);
  const cardW = (S.CW - 0.25) / 2;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  actions.forEach((action, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.25);
    const cy = S.BODY_TOP + row * (cardH + 0.2);

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_WHITE, borderColor: C.VIOLET_L });

    addCircleBadge(slide, i + 1, cx + 0.18, cy + 0.18, 0.42, C.VIOLET);

    slide.addText(trunc(action.titre || action.nom || '', 40), {
      x: cx + 0.76, y: cy + 0.18,
      w: cardW - 0.95, h: 0.42,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.BG_DARK, valign: 'middle', shrinkText: true,
    });

    slide.addText(trunc(action.description || '', 120), {
      x: cx + 0.18, y: cy + 0.73,
      w: cardW - 0.36, h: cardH - 0.92,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_3, valign: 'top', shrinkText: true,
    });

    if (action.delai || action.deadline) {
      addPill(slide, trunc(action.delai || action.deadline, 20), cx + 0.18, cy + cardH - 0.5, 2.5, 0.35, C.VIOLET_XL + '33', C.VIOLET);
    }
  });
}

// ─── FIDÉLISATION ────────────────────────────────────
function renderFidelisation(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Fidélisation', data.sous_titre, true);

  const strategies = truncArr(data.strategies || data.elements || [], 4);
  const cols = strategies.length <= 2 ? strategies.length : 2;
  const rows = Math.ceil(strategies.length / cols);
  const cardW = (S.CW - 0.25) / cols;
  const cardH = (S.BODY_H - 0.2 * (rows - 1)) / rows;

  strategies.forEach((strat, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const cx = S.MX + col * (cardW + 0.25);
    const cy = S.BODY_TOP + row * (cardH + 0.2);

    addCard(slide, { x: cx, y: cy, w: cardW, h: cardH, bg: C.BG_CARD });

    slide.addText(strat.emoji || '🔄', {
      x: cx + 0.2, y: cy + 0.18,
      w: 0.65, h: 0.65, fontSize: 24,
    });

    slide.addText(trunc(strat.titre || strat.nom || '', 35), {
      x: cx + 0.2, y: cy + 0.92,
      w: cardW - 0.4, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.WHITE, shrinkText: true,
    });

    slide.addText(trunc(strat.description || '', 140), {
      x: cx + 0.2, y: cy + 1.45,
      w: cardW - 0.4, h: cardH - 1.62,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_1, valign: 'top', shrinkText: true,
    });
  });
}

// ─── CALENDRIER ──────────────────────────────────────
function renderCalendrier(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Calendrier Éditorial', data.sous_titre, true);

  const semaines = truncArr(data.semaines || [], 8);

  // Column headers
  const headers = ['Plateforme', 'S1', 'S2', 'S3', 'S4'];
  const colWidths = [3.2, 2.3, 2.3, 2.3, 2.3]; // total ≈ 12.4
  const rowH = Math.min((S.BODY_H - 0.5) / (semaines.length + 1), 0.88);

  // Header row
  let x = S.MX;
  headers.forEach((h, i) => {
    addCard(slide, { x, y: S.BODY_TOP, w: colWidths[i] - 0.08, h: 0.45, bg: C.VIOLET });
    slide.addText(h, {
      x, y: S.BODY_TOP,
      w: colWidths[i] - 0.08, h: 0.45,
      fontFace: FONT.TITLE, fontSize: 11, bold: true,
      color: C.WHITE, align: 'center', valign: 'middle', shrinkText: true,
    });
    x += colWidths[i];
  });

  // Data rows
  semaines.forEach((sem, row) => {
    const cy = S.BODY_TOP + 0.52 + row * (rowH + 0.06);
    let rx = S.MX;
    const values = [
      sem.plateforme || sem.platform || '',
      sem.s1 || sem.semaine_1 || '',
      sem.s2 || sem.semaine_2 || '',
      sem.s3 || sem.semaine_3 || '',
      sem.s4 || sem.semaine_4 || '',
    ];
    values.forEach((val, ci) => {
      const bg = ci === 0 ? C.BG_CARD : (row % 2 === 0 ? C.BG_MID : C.BG_CARD);
      addCard(slide, { x: rx, y: cy, w: colWidths[ci] - 0.08, h: rowH, bg });
      slide.addText(trunc(String(val), ci === 0 ? 25 : 50), {
        x: rx + 0.1, y: cy + 0.05,
        w: colWidths[ci] - 0.25, h: rowH - 0.1,
        fontFace: FONT.BODY,
        fontSize: ci === 0 ? 11 : 10,
        bold: ci === 0,
        color: ci === 0 ? C.VIOLET_L : C.GRAY_1,
        valign: 'middle', shrinkText: true,
      });
      rx += colWidths[ci];
    });
  });
}

// ─── TOURNAGE ────────────────────────────────────────
function renderTournage(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Guide de Tournage', data.sous_titre, true);

  const etapes = truncArr(data.etapes || [], 5);
  const cardH = (S.BODY_H - 0.15 * (etapes.length - 1)) / Math.max(etapes.length, 1);

  etapes.forEach((etape, i) => {
    const cy = S.BODY_TOP + i * (cardH + 0.15);
    addCard(slide, { x: S.MX, y: cy, w: S.CW, h: cardH, bg: C.BG_CARD });

    const icons = ['🎬', '🎙️', '💡', '🎨', '✂️'];
    slide.addText(etape.emoji || icons[i % icons.length], {
      x: S.MX + 0.15, y: cy + (cardH - 0.55) / 2,
      w: 0.55, h: 0.55, fontSize: 20,
    });

    slide.addText(trunc(etape.titre || etape.nom || '', 35), {
      x: S.MX + 0.88, y: cy + 0.08,
      w: 3.5, h: cardH - 0.16,
      fontFace: FONT.TITLE, fontSize: 13, bold: true,
      color: C.VIOLET_L, valign: 'middle', shrinkText: true,
    });

    slide.addText(trunc(etape.description || '', 180), {
      x: S.MX + 4.65, y: cy + 0.08,
      w: S.CW - 4.8, h: cardH - 0.16,
      fontFace: FONT.BODY, fontSize: 11,
      color: C.GRAY_1, valign: 'middle', shrinkText: true,
    });
  });
}

// ─── RÉSUMÉ ──────────────────────────────────────────
function renderResume(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };
  addLogo(slide);
  addHeader(slide, data.titre || 'Récapitulatif', data.sous_titre, true);

  const points = truncArr(data.points || [], 5);
  const cardH = (S.BODY_H - 0.15 * (points.length - 1)) / Math.max(points.length, 1);

  points.forEach((point, i) => {
    const cy = S.BODY_TOP + i * (cardH + 0.15);
    addCard(slide, { x: S.MX, y: cy, w: S.CW, h: cardH, bg: C.BG_CARD });

    // Accent dot
    slide.addShape('ellipse', {
      x: S.MX + 0.2, y: cy + (cardH - 0.2) / 2,
      w: 0.2, h: 0.2,
      fill: { color: C.VIOLET_L },
      line: { type: 'none' },
    });

    slide.addText(trunc(String(typeof point === 'object' ? (point.texte || point.label) : point), 200), {
      x: S.MX + 0.6, y: cy + 0.08,
      w: S.CW - 0.8, h: cardH - 0.16,
      fontFace: FONT.BODY, fontSize: 13,
      color: C.GRAY_1, valign: 'middle', shrinkText: true,
    });
  });
}

// ─── CONTACT ─────────────────────────────────────────
function renderContact(pptx, slide, data) {
  slide.background = { color: C.BG_DARK };

  // Big violet left panel
  slide.addShape('rect', {
    x: 0, y: 0,
    w: S.W * 0.45, h: S.H,
    fill: { color: C.VIOLET },
    line: { type: 'none' },
  });

  // C logo large
  slide.addText('C', {
    x: 0, y: S.H / 2 - 1.5,
    w: S.W * 0.45, h: 3,
    fontFace: FONT.TITLE,
    fontSize: 180,
    bold: true,
    color: 'FFFFFF',
    alpha: 15,
    align: 'center',
    valign: 'middle',
  });

  slide.addText('CREATICAL', {
    x: 0, y: S.H * 0.4,
    w: S.W * 0.45, h: 0.55,
    fontFace: FONT.TITLE,
    fontSize: 18,
    bold: true,
    color: C.WHITE,
    charSpacing: 5,
    align: 'center',
    shrinkText: true,
  });

  slide.addText('Stratégie · Contenu · Croissance', {
    x: 0, y: S.H * 0.4 + 0.62,
    w: S.W * 0.45, h: 0.35,
    fontFace: FONT.BODY,
    fontSize: 11,
    color: C.VIOLET_XL,
    align: 'center',
    shrinkText: true,
  });

  // Right side
  const rightX = S.W * 0.45 + S.MX;
  const rightW = S.W - S.W * 0.45 - S.MX * 2;

  slide.addText(trunc(data.titre || 'Passons à l\'action', 35), {
    x: rightX, y: 1.2,
    w: rightW, h: 1.2,
    fontFace: FONT.TITLE,
    fontSize: 28,
    bold: true,
    color: C.WHITE,
    valign: 'top',
    shrinkText: true,
  });

  const contacts = [
    { emoji: '📧', value: data.email },
    { emoji: '📱', value: data.telephone },
    { emoji: '🌐', value: data.site },
    { emoji: '📍', value: data.adresse },
  ].filter(c => c.value);

  contacts.forEach((c, i) => {
    const cy = 2.8 + i * 0.75;
    slide.addText(`${c.emoji}  ${trunc(c.value, 45)}`, {
      x: rightX, y: cy,
      w: rightW, h: 0.55,
      fontFace: FONT.BODY,
      fontSize: 13,
      color: C.GRAY_1,
      valign: 'middle',
      shrinkText: true,
    });
  });

  if (data.cta) {
    addCard(slide, { x: rightX, y: S.H - 1.2, w: rightW, h: 0.62, bg: C.ORANGE });
    slide.addText(trunc(data.cta, 50), {
      x: rightX, y: S.H - 1.2,
      w: rightW, h: 0.62,
      fontFace: FONT.TITLE,
      fontSize: 14,
      bold: true,
      color: C.WHITE,
      align: 'center',
      valign: 'middle',
      shrinkText: true,
    });
  }
}

// ─── GENERIC FALLBACK ────────────────────────────────
function renderGeneric(pptx, slide, data) {
  const isDark = !data.theme || data.theme !== 'light';
  slide.background = { color: isDark ? C.BG_DARK : C.BG_LIGHT };
  addLogo(slide, isDark);
  addHeader(slide, data.titre || 'Section', data.sous_titre, isDark);

  // Try to render content intelligently
  const content = data.contenu || data.texte || data.description || data.content;
  if (content) {
    slide.addText(trunc(String(content), 500), {
      x: S.MX, y: S.BODY_TOP,
      w: S.CW, h: S.BODY_H,
      fontFace: FONT.BODY,
      fontSize: 14,
      color: isDark ? C.GRAY_1 : C.GRAY_3,
      valign: 'top',
      shrinkText: true,
    });
  }

  // Handle any array-like data
  const items = data.items || data.elements || data.liste;
  if (Array.isArray(items) && items.length) {
    const truncItems = truncArr(items, 6);
    const cardH = (S.BODY_H - 0.15 * (truncItems.length - 1)) / truncItems.length;
    truncItems.forEach((item, i) => {
      const cy = S.BODY_TOP + i * (cardH + 0.15);
      addCard(slide, { x: S.MX, y: cy, w: S.CW, h: cardH, bg: isDark ? C.BG_CARD : C.BG_WHITE });
      slide.addText(trunc(String(typeof item === 'object' ? JSON.stringify(item) : item), 200), {
        x: S.MX + 0.25, y: cy + 0.1,
        w: S.CW - 0.5, h: cardH - 0.2,
        fontFace: FONT.BODY, fontSize: 12,
        color: isDark ? C.GRAY_1 : C.GRAY_3,
        valign: 'middle', shrinkText: true,
      });
    });
  }
}

// ══════════════════════════════════════════════════════
// SLIDE DISPATCH TABLE
// ══════════════════════════════════════════════════════
const RENDERERS = {
  title:                  renderTitle,
  section:                renderSection,
  contexte:               renderContexte,
  marche:                 renderMarche,
  geographie:             renderGeographie,
  concurrence:            renderConcurrence,
  niveau_conscience:      renderNiveauConscience,
  sophistication_marche:  renderSophisticationMarche,
  objectifs:              renderObjectifs,
  brand_archetype:        renderBrandArchetype,
  persona:                renderPersona,
  positionnement:         renderPositionnement,
  plateformes:            renderPlateformes,
  kpis:                   renderKpis,
  piliers:                renderPiliers,
  piliers_contenu:        renderPiliersContenu,
  contenu_strategique:    renderContenuStrategique,
  methode:                renderMethode,
  hook_regles:            renderHookRegles,
  emotion:                renderEmotion,
  funnel_idees:           renderFunnelIdees,
  phase_funnel:           renderPhaseFunnel,
  plan_action:            renderPlanAction,
  fidelisation:           renderFidelisation,
  calendrier:             renderCalendrier,
  tournage:               renderTournage,
  resume:                 renderResume,
  contact:                renderContact,
};

// ══════════════════════════════════════════════════════
// API ENDPOINT
// ══════════════════════════════════════════════════════
app.post('/generate', async (req, res) => {
  try {
    const { slidesData } = req.body;
    if (!slidesData) return res.status(400).json({ error: 'slidesData manquant' });

    const { metadata = {}, slides = [] } = slidesData;

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33" x 7.5"
    pptx.author = 'Creatical';
    pptx.subject = metadata.client || 'Stratégie de Contenu';
    pptx.title = metadata.titre || 'Présentation Creatical';

    for (const slideData of slides) {
      const slideObj = pptx.addSlide();
      const type = (slideData.type || 'generic').toLowerCase().trim();
      const renderer = RENDERERS[type] || renderGeneric;

      try {
        renderer(pptx, slideObj, slideData);
      } catch (err) {
        console.error(`[ERROR] Slide type="${type}":`, err.message);
        // Fallback: render error info on slide
        slideObj.background = { color: C.BG_DARK };
        slideObj.addText(`⚠️ Erreur slide "${type}": ${err.message}`, {
          x: 0.5, y: 3, w: 12, h: 1,
          fontFace: FONT.BODY, fontSize: 14,
          color: C.ORANGE, align: 'center',
        });
      }
    }

    const pptxBase64 = await pptx.write({ outputType: 'base64' });
    const client = metadata.client ? metadata.client.replace(/[^a-zA-Z0-9]/g, '_') : 'client';
    const date = new Date().toISOString().slice(0, 10);
    const filename = `Strategie_${client}_${date}.pptx`;

    res.json({ pptxBase64, filename, slideCount: slides.length });
  } catch (err) {
    console.error('[FATAL]', err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok', version: '3.0.0' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Creatical PPTX Server v3.0 — port ${PORT}`);
});

module.exports = app;
