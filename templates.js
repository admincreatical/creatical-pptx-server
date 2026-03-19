// ═══════════════════════════════════════════════════════
// CREATICAL SLIDE HTML TEMPLATE ENGINE
// Generates HTML for each slide type matching Vincent Denis style
// ═══════════════════════════════════════════════════════

const FONTS_LINK = `<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&family=Caveat:wght@400;700&display=swap" rel="stylesheet">`;

const BASE_CSS = `
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  width: 1280px; height: 720px; overflow: hidden;
  font-family: 'Inter', sans-serif; color: #FFFFFF;
}
.slide {
  width: 1280px; height: 720px; position: relative; overflow: hidden;
  background: linear-gradient(135deg, #0D0221 0%, #1A0A2E 35%, #2D1054 65%, #4C1D95 100%);
}
.slide::before {
  content: ''; position: absolute; top: -100px; right: -150px;
  width: 600px; height: 600px;
  background: radial-gradient(circle, rgba(124,58,237,0.3) 0%, transparent 70%);
}
.slide::after {
  content: ''; position: absolute; bottom: -200px; right: -100px;
  width: 500px; height: 500px;
  background: radial-gradient(circle, rgba(168,85,247,0.15) 0%, transparent 70%);
}
.frame {
  position: absolute; top: 24px; left: 24px; right: 24px; bottom: 24px;
  border: 1px solid rgba(255,255,255,0.12); border-radius: 16px;
  pointer-events: none; z-index: 1;
}
.content { position: relative; z-index: 2; padding: 48px 56px; height: 100%; display: flex; flex-direction: column; }
.footer-logo {
  position: absolute; bottom: 16px; left: 50%; transform: translateX(-50%);
  width: 36px; height: 36px; border-radius: 50%;
  background: linear-gradient(135deg, #7C3AED, #D946EF, #EC4899);
  display: flex; align-items: center; justify-content: center;
  font-weight: 800; font-size: 16px; z-index: 10;
  box-shadow: 0 0 20px rgba(124,58,237,0.4);
}
.pill {
  display: inline-flex; align-items: center; padding: 8px 24px;
  border-radius: 50px; background: #fff; color: #0D0221;
  font-weight: 700; font-size: 14px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);
}
.pill--violet { background: #7C3AED; color: #fff; }
.pill--sm { padding: 6px 18px; font-size: 12px; }
.num-circle {
  width: 48px; height: 48px; border-radius: 50%; background: #7C3AED;
  display: flex; align-items: center; justify-content: center;
  font-weight: 800; font-size: 20px; flex-shrink: 0;
}
.num-circle--sm { width: 36px; height: 36px; font-size: 14px; }
.card {
  background: rgba(26,10,46,0.5); border: 1px solid rgba(124,58,237,0.3);
  border-radius: 8px; padding: 20px; position: relative;
  backdrop-filter: blur(10px);
}
.card--accent::before {
  content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
  background: #A855F7; border-radius: 8px 0 0 8px;
}
.card--pink::before { background: #EC4899; }
.card--magenta::before { background: #D946EF; }
.label { font-size: 12px; font-weight: 600; color: #C4B5FD; text-transform: uppercase; letter-spacing: 1px; }
.body { font-size: 14px; line-height: 1.6; color: #D1D5DB; }
.body--sm { font-size: 12px; }
.title-xl { font-size: 52px; font-weight: 900; letter-spacing: 3px; text-transform: uppercase; line-height: 1.05; }
.title-lg { font-size: 38px; font-weight: 800; letter-spacing: 2px; text-transform: uppercase; }
.title-md { font-size: 28px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; }
.serif { font-family: 'Playfair Display', serif; font-style: italic; }
.handwritten { font-family: 'Caveat', cursive; }
.text-violet { color: #C4B5FD; }
.text-pink { color: #EC4899; }
.text-gray { color: #9CA3AF; }
.emoji { font-size: 28px; }
.emoji--lg { font-size: 42px; }
</style>`;

function esc(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function wrapSlide(inner) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">${FONTS_LINK}${BASE_CSS}</head><body>${inner}</body></html>`;
}

function renderTitle(s) {
  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content" style="justify-content: center; gap: 10px;">
        <div class="serif text-violet" style="font-size: 18px;">${esc(s.subtitle)}</div>
        <div style="display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap;">
          <span class="title-xl" style="font-size: 56px;">${esc(s.title || 'STRATEGIE')}</span>
        </div>
        <div class="serif text-violet" style="font-size: 36px; margin-top: -6px;">de contenu</div>
        <div style="margin-top: 16px;">
          <span class="pill">${esc(s.client)}</span>
        </div>
        <div style="position: absolute; right: 80px; top: 50%; transform: translateY(-50%);">
          <div style="width: 160px; height: 160px; border-radius: 24px; background: linear-gradient(135deg, #7C3AED, #D946EF, #06B6D4, #EC4899); display: flex; align-items: center; justify-content: center; font-size: 100px; font-weight: 900; box-shadow: 0 12px 40px rgba(124,58,237,0.5);">C</div>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderObjectifs(s, meta) {
  const items = (s.items || []).map(item => `
    <div style="display: flex; align-items: flex-start; gap: 10px; font-size: 14px; line-height: 1.5;">
      <span style="color: #A855F7; font-size: 18px; flex-shrink: 0;">→</span>
      <span>${esc(item)}</span>
    </div>`).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div style="display: flex; height: 100%;">
          <div style="flex: 1; display: flex; flex-direction: column; justify-content: center;">
            <div style="display: flex; gap: 16px; align-items: flex-start;">
              <span class="emoji--lg">🎯</span>
              <div>
                <div class="serif text-violet" style="font-size: 16px;">OBJECTIFS</div>
                <div style="font-size: 120px; font-weight: 900; line-height: 0.85; margin-top: -8px;">2026</div>
              </div>
              <span class="emoji--lg" style="margin-top: -8px;">🔥</span>
            </div>
            <div style="margin-top: 16px;"><span class="pill" style="font-size: 12px;">${esc(s.clientName || meta?.client || '')}</span></div>
          </div>
          <div style="width: 420px; display: flex; flex-direction: column; gap: 18px; justify-content: center;">
            ${items}
          </div>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderSection(s) {
  const emojis = ['💡','📊','🔥','✅','🎯','📈'];
  const randomEmojis = () => emojis.sort(() => Math.random() - 0.5).slice(0, 3);
  const [e1, e2, e3] = randomEmojis();

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content" style="justify-content: center; align-items: center; text-align: center;">
        <div style="position: absolute; top: 80px; left: 160px;" class="emoji--lg">${e1}</div>
        <div style="position: absolute; top: 60px; right: 120px;" class="emoji--lg">${e2}</div>
        <div style="position: absolute; bottom: 100px; right: 200px;" class="emoji--lg">${e3}</div>
        <div class="title-xl" style="font-size: 60px; max-width: 800px;">${esc(s.title)}</div>
        <div class="serif text-violet" style="font-size: 24px; margin-top: 12px;">${esc(s.subtitle)}</div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderBrand(s) {
  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">BRAND ARCHETYPE</div>
        <div style="margin-top: 10px;"><span class="pill">${esc(s.subtitle)}</span></div>
        <div class="card card--accent" style="margin-top: 20px; flex: 1;">
          <p class="body" style="line-height: 1.7;">${esc(s.description)}</p>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderPersona(s) {
  const emotions = (s.emotions || []).map(e => `<div class="body--sm" style="padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,0.06);">${esc(e)}</div>`).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">PERSONA CIBLE</div>
        <div style="margin-top: 8px;"><span class="pill pill--violet">${esc(s.name)}</span></div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 16px; flex: 1;">
          <div class="card card--accent">
            <div class="label" style="margin-bottom: 10px;">Profil</div>
            <p class="body--sm">${esc(s.profil)}</p>
            <div class="label text-pink" style="margin-top: 16px; margin-bottom: 6px;">Core Problem</div>
            <p class="body--sm">${esc(s.core_problem)}</p>
          </div>
          <div class="card card--pink">
            <div class="label" style="margin-bottom: 10px;">Émotions & Peurs</div>
            ${emotions}
          </div>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderPositionnement(s) {
  const elems = (s.elements || []).map((el, i) => `
    <div class="card card--accent" style="display: flex; gap: 12px; align-items: flex-start; padding: 14px 16px;">
      <div class="num-circle--sm" style="width: 32px; height: 32px; border-radius: 50%; background: #7C3AED; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 13px; flex-shrink: 0;">${i + 1}</div>
      <p class="body--sm" style="padding-top: 4px;">${esc(el)}</p>
    </div>`).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">POSITIONNEMENT</div>
        <div style="margin-top: 14px; padding: 14px 28px; background: #7C3AED; border-radius: 8px; text-align: center;">
          <span class="serif" style="color: #fff; font-size: 16px;">${esc(s.tagline)}</span>
        </div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 14px; flex: 1;">
          ${elems}
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderPiliers(s) {
  const colors = ['#7C3AED', '#A855F7', '#D946EF', '#EC4899'];
  const piliers = (s.piliers || []).map((p, i) => `
    <div class="card" style="border-left: 4px solid ${colors[i] || colors[0]}; padding: 16px 18px;">
      <div style="display: flex; gap: 12px; align-items: flex-start;">
        <div style="width: 36px; height: 36px; border-radius: 50%; background: ${colors[i] || colors[0]}; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 15px; flex-shrink: 0;">${i + 1}</div>
        <div>
          <div style="font-weight: 700; font-size: 14px; margin-bottom: 6px;">${esc(p.nom)}</div>
          <p class="body--sm" style="line-height: 1.5;">${esc(p.description)}</p>
        </div>
      </div>
    </div>`).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">PILIERS ÉDITORIAUX</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 16px; flex: 1;">
          ${piliers}
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderFunnel(s) {
  function col(data, cssClass, color) {
    if (!data) return '';
    const vids = (data.videos || []).map((v, i) => `<div style="font-size: 11px; color: #D1D5DB; padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,0.05);">${i + 1}. ${esc(v)}</div>`).join('');
    return `
      <div style="background: rgba(26,10,46,0.5); border: 1px solid ${color}40; border-radius: 8px; overflow: hidden; display: flex; flex-direction: column;">
        <div style="padding: 12px 16px; background: ${color}; font-weight: 700; font-size: 15px; text-align: center; letter-spacing: 1px;">${esc(data.titre)}</div>
        <div style="padding: 14px 16px; flex: 1;">
          <div style="font-size: 11px; color: #C4B5FD; font-style: italic; margin-bottom: 10px;">${esc(data.description)}</div>
          ${vids}
        </div>
      </div>`;
  }

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">FUNNEL D'ACQUISITION</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; margin-top: 14px; flex: 1;">
          ${col(s.tofu, 'tofu', '#7C3AED')}
          ${col(s.mofu, 'mofu', '#A855F7')}
          ${col(s.bofu, 'bofu', '#D946EF')}
        </div>
      </div>
    </div>`);
}

function renderPlan(s) {
  const phases = (s.phases || []).map((p, i) => {
    const actions = (p.actions || []).map(a => `<div style="font-size: 11px; color: #D1D5DB; padding: 4px 0 4px 14px; position: relative;"><span style="position: absolute; left: 0; color: #A855F7; font-weight: 700;">›</span>${esc(a)}</div>`).join('');
    return `
      <div style="display: flex; flex-direction: column; align-items: center; text-align: center;">
        <span class="pill pill--violet" style="font-size: 11px; padding: 6px 16px;">${esc(p.titre)}</span>
        <div class="text-gray" style="font-size: 10px; margin-top: 4px;">${esc(p.periode)}</div>
        <div style="width: 56px; height: 56px; border-radius: 50%; background: #7C3AED; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 22px; margin: 14px 0; box-shadow: 0 0 20px rgba(124,58,237,0.5); z-index: 3;">${i + 1}</div>
        <div style="text-align: left; width: 100%; padding: 0 8px;">${actions}</div>
      </div>`;
  }).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">PLAN D'ACTION 90 JOURS</div>
        <div style="position: relative; margin-top: 20px; flex: 1;">
          <div style="position: absolute; top: 80px; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #7C3AED, #A855F7, #D946EF); border-radius: 2px;"></div>
          <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; height: 100%;">
            ${phases}
          </div>
        </div>
      </div>
    </div>`);
}

function renderResume(s) {
  const points = (s.points || []).map(pt => `
    <div>
      <span class="pill pill--sm">${esc(pt.titre)}</span>
      <p class="body--sm" style="margin-top: 8px; line-height: 1.5;">${esc(pt.description)}</p>
    </div>`).join('');

  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 40px; height: 100%; align-items: center;">
          <div>
            <div class="title-xl" style="font-size: 48px;">RÉSUMÉ</div>
            <div class="serif text-violet" style="font-size: 36px; margin-top: 4px;">du mois</div>
          </div>
          <div style="display: flex; flex-direction: column; gap: 22px;">
            ${points}
          </div>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderContact(s) {
  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content" style="justify-content: center; align-items: flex-end; padding-right: 120px; gap: 8px;">
        <div style="width: 90px; height: 90px; border-radius: 20px; background: linear-gradient(135deg, #7C3AED, #D946EF, #06B6D4, #EC4899); display: flex; align-items: center; justify-content: center; font-size: 48px; font-weight: 900; margin-bottom: 16px; box-shadow: 0 8px 30px rgba(124,58,237,0.4);">C</div>
        <div style="font-size: 22px; font-weight: 700;">${esc(s.name)}</div>
        <div class="serif text-violet" style="font-size: 18px;">${esc(s.role)}</div>
        <div class="serif text-gray" style="font-size: 14px;">${esc(s.email)}</div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderGeneric(s) {
  return wrapSlide(`
    <div class="slide">
      <div class="frame"></div>
      <div class="content">
        <div class="title-md">${esc(s.title)}</div>
        <div class="card card--accent" style="margin-top: 16px; flex: 1;">
          <p class="body">${esc(s.content || '')}</p>
        </div>
      </div>
      <div class="footer-logo">C</div>
    </div>`);
}

function renderSlideHTML(slide, meta) {
  switch (slide.type) {
    case 'title': return renderTitle(slide);
    case 'objectifs': return renderObjectifs(slide, meta);
    case 'section': return renderSection(slide);
    case 'brand_archetype': return renderBrand(slide);
    case 'persona': return renderPersona(slide);
    case 'positionnement': return renderPositionnement(slide);
    case 'piliers': return renderPiliers(slide);
    case 'funnel': return renderFunnel(slide);
    case 'plan_action': return renderPlan(slide);
    case 'resume': return renderResume(slide);
    case 'contact': return renderContact(slide);
    default: return renderGeneric(slide);
  }
}

module.exports = { renderSlideHTML };
