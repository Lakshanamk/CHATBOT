import puppeteer from 'puppeteer';
import fs from 'fs';
import path from 'path';

const SCREENSHOT_DIR = path.join(process.cwd(), 'screenshots');
if (!fs.existsSync(SCREENSHOT_DIR)) {
  fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });
}

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function run() {
  const browser = await puppeteer.launch({
    headless: 'new',
    defaultViewport: { width: 1280, height: 800 }
  });
  const page = await browser.newPage();

  console.log('Navigating to http://localhost:5173 ...');
  await page.goto('http://localhost:5173', { waitUntil: 'networkidle2' });
  await delay(1000);

  // Screenshot 1: Main Dashboard (Welcome Screen)
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '01_main_dashboard.png') });
  console.log('Saved 01_main_dashboard.png');

  // Helper to type query and wait
  async function typeAndSend(text) {
    const inputSelector = 'textarea, input[type="text"]';
    await page.waitForSelector(inputSelector);
    await page.type(inputSelector, text);
    await page.keyboard.press('Enter');
    await delay(1200);
  }

  // Screenshot 2: Fee Calculator
  await typeAndSend('How much is the tuition fee?');
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '02_fee_calculator.png') });
  console.log('Saved 02_fee_calculator.png');

  // Screenshot 3: Certificate Request Widget
  await typeAndSend('How can I apply for a bonafide certificate?');
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '03_certificate_request.png') });
  console.log('Saved 03_certificate_request.png');

  // Screenshot 4: Exam Schedule Widget
  await typeAndSend('When are the semester exams?');
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '04_exam_schedule.png') });
  console.log('Saved 04_exam_schedule.png');

  // Screenshot 5: Hostel Mess Menu Widget
  await typeAndSend('What is today mess menu?');
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '05_hostel_mess.png') });
  console.log('Saved 05_hostel_mess.png');

  // Screenshot 6: Placement Stats Widget
  await typeAndSend('Show me placement statistics');
  await page.screenshot({ path: path.join(SCREENSHOT_DIR, '06_placement_stats.png') });
  console.log('Saved 06_placement_stats.png');

  // Screenshot 7: Knowledge Base Modal
  try {
    const buttons = await page.$$('button');
    for (const btn of buttons) {
      const text = await page.evaluate(el => el.textContent || el.getAttribute('title') || '', btn);
      if (text.includes('Knowledge') || text.includes('Search') || text.includes('FAQ')) {
        await btn.click();
        await delay(800);
        await page.screenshot({ path: path.join(SCREENSHOT_DIR, '07_knowledge_base.png') });
        console.log('Saved 07_knowledge_base.png');
        await page.keyboard.press('Escape');
        await delay(500);
        break;
      }
    }
  } catch (e) {
    console.log('Error opening KB modal:', e.message);
  }

  // Screenshot 8: Light Theme Mode
  try {
    const buttons = await page.$$('button');
    for (const btn of buttons) {
      const title = await page.evaluate(el => el.getAttribute('title') || el.ariaLabel || '', btn);
      if (title.toLowerCase().includes('theme') || title.toLowerCase().includes('dark') || title.toLowerCase().includes('light')) {
        await btn.click();
        await delay(600);
        await page.screenshot({ path: path.join(SCREENSHOT_DIR, '08_light_mode.png') });
        console.log('Saved 08_light_mode.png');
        break;
      }
    }
  } catch (e) {
    console.log('Error toggling theme:', e.message);
  }

  await browser.close();
  console.log('All screenshots captured successfully!');
}

run().catch(err => {
  console.error('Error taking screenshots:', err);
  process.exit(1);
});
