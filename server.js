const express = require('express');
const bodyParser = require('body-parser');
const { chromium } = require('playwright');
const PptxGenJS = require('pptxgenjs');
const { Readable } = require('stream');

const app = express();

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

const PORT = 9000;

// -----------------------------
// GLOBAL BROWSER
// -----------------------------
let browser = null;

async function getBrowser() {
  if (!browser) {
    browser = await chromium.launch({
      headless: true,
      args: ['--disable-dev-shm-usage']
    });
    console.log('Chromium launched.');
  }
  return browser;
}

// -----------------------------
// CONCURRENCY LIMIT
// -----------------------------
const MAX_CONCURRENCY = 4;
let active = 0;

async function withConcurrency(fn) {
  while (active >= MAX_CONCURRENCY) {
    await new Promise((r) => setTimeout(r, 40));
  }
  active++;
  try {
    return await fn();
  } finally {
    active--;
  }
}

// -----------------------------
// RENDER SLIDE
// -----------------------------
async function renderSlide(html, css, width, height) {
  const browser = await getBrowser();
  const context = await browser.newContext({
    viewport: { width, height },
    deviceScaleFactor: 2
  });
  const page = await context.newPage();

  const content = `
    <html>
      <head><style>${css}</style></head>
      <body style="margin:0;padding:0;">${html}</body>
    </html>
  `;

  await page.setContent(content, { waitUntil: 'networkidle' });

  const buffer = await page.screenshot({ type: 'png' });

  await page.close();
  await context.close();

  return buffer;
}

// -----------------------------
// BUILD PPT
// -----------------------------
function normalizeTextConfig(text) {
  if (!text) return null;

  if (typeof text === 'string') {
    return { value: text, options: {} };
  }

  if (typeof text === 'object') {
    const { value, content, text: innerText, options = {}, ...rest } = text;
    const finalValue = value ?? content ?? innerText;
    if (!finalValue) return null;

    return { value: finalValue, options: { ...rest, ...options } };
  }

  return null;
}

async function buildPpt(slideData) {
  const ppt = new PptxGenJS();

  slideData.forEach(({ image, text }) => {
    const base64 = image.toString('base64');

    const slide = ppt.addSlide();

    slide.addImage({
      // CHANGE THIS LINE:
      // Use the standard Data URL format instead of the internal "base64:IMAGE_PNG"
      data: `data:image/png;base64,${base64}`,
      x: 0,
      y: 0,
      w: '100%',
      h: '100%'
    });

    const normalizedText = normalizeTextConfig(text);

    if (normalizedText) {
      const defaultTextOptions = {
        x: '5%',
        y: '5%',
        w: '90%',
        color: '363636',
        fontSize: 18,
        fontFace: 'Arial',
        align: 'left'
      };

      slide.addText(normalizedText.value, {
        ...defaultTextOptions,
        ...normalizedText.options
      });
    }
  });

  const base64Out = await ppt.write('base64');
  return Buffer.from(base64Out, 'base64');
}



// -----------------------------------------------
// EXAMPLE PAYLOAD FOR TESTING
// -----------------------------------------------
//
// {
//   "css": "body { font-family: Arial; } td { border:1px solid black; padding:8px; }",
//   "slides": [
//     {
//       "html": "<div><h2>Hello!</h2><table><tr><td>A</td><td>B</td></tr></table></div>",
//       "width": 1400,
//       "height": 900
//     }
//   ]
// }
//
// -----------------------------------------------

// -----------------------------
// ENDPOINT: /generate-ppt
// -----------------------------
app.post('/generate-ppt', async (req, res) => {
  try {
    const { css, slides } = req.body;

    if (!slides || slides.length === 0) {
      return res.status(400).json({ error: 'No slides provided' });
    }
    console.log(req.body, 'PAYLOAD')
    console.log(`Rendering ${slides.length} slides...`);

    const slidesWithImages = await Promise.all(
      slides.map((s) =>
        withConcurrency(() =>
          renderSlide(s.html, css, s.width, s.height).then((image) => ({
            image,
            text: s.text
          }))
        )
      )
    );

    const pptBuffer = await buildPpt(slidesWithImages);

    const stream = Readable.from(pptBuffer);

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="report.pptx"'
    );

    stream.pipe(res);

  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'Failed to generate PPT' });
  }
});

// -----------------------------
// TEST ROUTE
// -----------------------------
app.get('/hello', (req, res) => {
  res.send('Hello World!');
});

// -----------------------------
// START SERVER
// -----------------------------
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
