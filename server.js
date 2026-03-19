const express = require("express");
const pptxgen = require("pptxgenjs");
const puppeteer = require("puppeteer");
const { renderSlideHTML } = require("./templates");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ═══════════════════════════════════════════════
// PUPPETEER BROWSER POOL
// ═══════════════════════════════════════════════
let browser = null;

async function getBrowser() {
  if (!browser) {
    browser = await puppeteer.launch({
      headless: "new",
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--no-first-run",
        "--no-zygote",
        "--single-process",
      ],
    });
  }
  return browser;
}

// ═══════════════════════════════════════════════
// RENDER HTML TO PNG (1280x720)
// ═══════════════════════════════════════════════
async function htmlToPng(html) {
  const b = await getBrowser();
  const page = await b.newPage();
  await page.setViewport({ width: 1280, height: 720 });
  await page.setContent(html, { waitUntil: "networkidle0", timeout: 15000 });
  // Wait a bit for fonts to load
  await new Promise(r => setTimeout(r, 500));
  const screenshot = await page.screenshot({
    type: "png",
    clip: { x: 0, y: 0, width: 1280, height: 720 },
  });
  await page.close();
  return screenshot;
}

// ═══════════════════════════════════════════════
// HEALTH CHECK
// ═══════════════════════════════════════════════
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "CREATICAL PPTX Generator v3 (Puppeteer)" });
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

    console.log(`Generating ${slides.length} slides...`);

    // 1. Render each slide HTML → PNG
    const pngBuffers = [];
    for (let i = 0; i < slides.length; i++) {
      console.log(`  Rendering slide ${i + 1}/${slides.length} (${slides[i].type})`);
      const html = renderSlideHTML(slides[i], meta);
      const png = await htmlToPng(html);
      pngBuffers.push(png);
    }

    // 2. Assemble PNGs into PPTX
    console.log("Assembling PPTX...");
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    pres.author = "CREATICAL";
    pres.title = meta.title || "Stratégie de Contenu";

    for (const pngBuf of pngBuffers) {
      const sl = pres.addSlide();
      const base64Img = "image/png;base64," + pngBuf.toString("base64");
      sl.addImage({
        data: base64Img,
        x: 0, y: 0, w: 10, h: 5.625,
      });
    }

    const pptxBase64 = await pres.write({ outputType: "base64" });

    console.log("Done! PPTX generated successfully.");

    res.json({
      success: true,
      pptxBase64,
      filename: `${meta.title || "Presentation"}.pptx`,
      slideCount: slides.length,
    });
  } catch (error) {
    console.error("Error generating PPTX:", error);
    res.status(500).json({ error: error.message, stack: error.stack });
  }
});

// ═══════════════════════════════════════════════
// PREVIEW (returns HTML of a single slide for debugging)
// ═══════════════════════════════════════════════
app.post("/preview", (req, res) => {
  try {
    const { slide, metadata } = req.body;
    const html = renderSlideHTML(slide, metadata);
    res.setHeader("Content-Type", "text/html");
    res.send(html);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ═══════════════════════════════════════════════
// GRACEFUL SHUTDOWN
// ═══════════════════════════════════════════════
process.on("SIGTERM", async () => {
  if (browser) await browser.close();
  process.exit(0);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`CREATICAL PPTX Server v3 (Puppeteer) running on port ${PORT}`);
});
