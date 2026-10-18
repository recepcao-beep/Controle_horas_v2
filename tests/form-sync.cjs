// Run after npm run build. PLAYWRIGHT_MODULE may point to an existing Playwright installation.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
(async () => {
  const browser = await chromium.launch({ headless: true, channel: process.env.BROWSER_CHANNEL || 'msedge' });
  try {
    for (const mobile of [false, true]) {
      const context = await browser.newContext(mobile ? { viewport: { width: 390, height: 844 }, isMobile: true } : {});
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      await page.clock.install();
      let loads = 0;
      await context.route('**/*', async route => {
        const url = new URL(route.request().url());
        if (url.pathname === '/api/sheets/load') {
          loads++;
          return route.fulfill({ json: {
            sectors: [{ id: 's', name: 'Recepção', fixedRate: 100 }],
            employees: [{ id: 'e', name: 'Pessoa Teste', sectorId: 's', type: 'REGISTRADO', salary: 2000 + loads, monthlyHours: 220, fixedDayOff: 6 }],
            requests: [], config: { overtimeMultiplier: 1.25, valeTransporteValue: 12 }
          } });
        }
        if (url.pathname.startsWith('/api/')) throw new Error('Unexpected API write: ' + url.pathname);
        if (url.hostname === 'cdn.tailwindcss.com') return route.fulfill({ body: 'window.tailwind = {};', contentType: 'application/javascript' });
        if (url.hostname !== 'form.test') return route.abort();
        const file = path.join(process.cwd(), 'dist', url.pathname === '/' ? 'index.html' : url.pathname);
        return fs.existsSync(file) ? route.fulfill({ body: fs.readFileSync(file), contentType: file.endsWith('.js') ? 'application/javascript' : 'text/html' }) : route.fulfill({ status: 404, body: '' });
      });
      await page.goto('http://form.test/');
      await page.waitForFunction(() => JSON.parse(localStorage.getItem('controle_horas_db_v3') || '{}').employees?.length === 1);
      await page.getByRole('button', { name: 'Começar' }).click();
      await page.getByRole('button', { name: 'Registrado', exact: true }).click();
      await page.locator('select').nth(0).selectOption('s');
      await page.locator('select').nth(1).selectOption('e');
      await page.getByRole('button', { name: 'Lançar Horários' }).click();
      await page.locator('input[type=date]').fill('2026-10-12');
      const times = page.locator('input[type=time]');
      for (const [i, value] of ['08:00','08:10','17:00','18:00'].entries()) await times.nth(i).fill(value);
      await page.getByRole('button', { name: 'Folga Vendida' }).nth(1).click();
      const snapshot = () => times.evaluateAll(inputs => inputs.map(input => input.value));
      const expected = await snapshot();
      const folgaClass = await page.getByRole('button', { name: 'Folga Vendida' }).nth(1).getAttribute('class');
      for (const trigger of ['focus','online','timer']) {
        const previous = loads;
        if (trigger === 'timer') await page.clock.fastForward(31000);
        else await page.evaluate(event => window.dispatchEvent(new Event(event)), trigger);
        await page.waitForFunction(salary => JSON.parse(localStorage.getItem('controle_horas_db_v3')).employees[0].salary > salary, 2000 + previous);
        assert.deepEqual(await snapshot(), expected, `${mobile ? 'mobile' : 'desktop'}: ${trigger} must preserve all time fields`);
        assert.equal(await page.getByRole('button', { name: 'Folga Vendida' }).nth(1).getAttribute('class'), folgaClass);
      }
      // A deliberate context change still initializes the appropriate blank week.
      await page.locator('input[type=date]').fill('2026-10-19');
      assert.ok((await snapshot()).every(value => value === ''), 'new week must initialize blank fields');
      await times.nth(0).fill('09:00');
      await page.getByRole('button', { name: 'Cancelar', exact: true }).click();
      await page.getByRole('button', { name: 'Lançar Horários' }).click();
      await page.waitForFunction(() => [...document.querySelectorAll('input[type=time]')].every(input => input.value === ''));
      assert.ok((await snapshot()).every(value => value === ''), 'reopening after cancel starts a new form');
      assert.deepEqual(errors, []);
      console.log(`PASS ${mobile ? 'mobile' : 'desktop'}: focus, reconnection and 30-second polling preserve the draft; week change and cancel/reopen still work.`);
      await context.close();
    }
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });

