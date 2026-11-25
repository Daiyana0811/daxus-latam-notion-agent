const path = require('path');

require('dotenv').config({ path: process.env.DOTENV_CONFIG_PATH });

const { chromium } = require('playwright');

const siteUrl = (process.env.SHAREPOINT_SITE_BASE_URL || 'https://zakidatasas.sharepoint.com/sites/general').replace(/\/$/, '');
const outputPath = process.env.SHAREPOINT_STORAGE_STATE_OUTPUT ||
  path.join(process.cwd(), '.sharepoint-storage-state.json');
const email = process.env.SHAREPOINT_EMAIL || process.env.MICROSOFT_EMAIL || '';
const password = process.env.SHAREPOINT_PASSWORD || process.env.MICROSOFT_PASSWORD || '';

function isLoginUrl(url) {
  return /login\.microsoftonline\.com|login\.live\.com/i.test(String(url || ''));
}

async function clickVisible(locator) {
  if (!(await locator.isVisible().catch(() => false))) {
    return false;
  }

  await locator.click();
  return true;
}

async function submit(page) {
  const button = page.locator('#idSIButton9, input[type="submit"], button[type="submit"]').first();
  if (await button.isVisible().catch(() => false)) {
    await button.click();
  } else {
    await page.keyboard.press('Enter');
  }
  await page.waitForTimeout(1200);
}

async function main() {
  if (!email || !password) {
    throw new Error('Configura SHAREPOINT_EMAIL y SHAREPOINT_PASSWORD antes de renovar la sesion.');
  }

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    console.log('Abriendo SharePoint. Completa solamente el MFA si Microsoft lo solicita.');
    await page.goto(siteUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    const deadline = Date.now() + 10 * 60 * 1000;

    while (Date.now() < deadline) {
      if (!isLoginUrl(page.url()) && /sharepoint\.com/i.test(page.url())) {
        await page.waitForTimeout(3000);
        await context.storageState({ path: outputPath });
        console.log(`Sesion guardada en ${outputPath}`);
        return;
      }

      const emailInput = page.locator('input[type="email"], input[name="loginfmt"], #i0116').first();
      if (await emailInput.isVisible().catch(() => false)) {
        await emailInput.fill(email);
        await submit(page);
        continue;
      }

      const passwordInput = page.locator('input[type="password"], input[name="passwd"], #i0118').first();
      if (await passwordInput.isVisible().catch(() => false)) {
        await passwordInput.fill(password);
        await submit(page);
        continue;
      }

      const account = page.getByText(email, { exact: false }).first();
      if (await clickVisible(account)) {
        await page.waitForTimeout(1200);
        continue;
      }

      const staySignedIn = page.locator(
        'input[type="submit"][value="Sí"], input[type="submit"][value="Yes"], button:has-text("Sí"), button:has-text("Yes")'
      ).first();
      if (await clickVisible(staySignedIn)) {
        await page.waitForTimeout(1200);
        continue;
      }

      const title = await page.title().catch(() => page.url());
      console.log(`Esperando MFA o confirmacion manual... ${title}`);
      await page.waitForTimeout(2500);
    }

    throw new Error('No se completo el login de SharePoint dentro de 10 minutos.');
  } finally {
    await browser.close();
  }
}

main().catch(error => {
  console.error('Error renovando la sesion de SharePoint:', error);
  process.exit(1);
});
