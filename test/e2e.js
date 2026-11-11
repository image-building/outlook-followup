// Headless smoke test: real task pane + mocked Office.js + mocked Graph.
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const root = path.join(__dirname, '..', 'src');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });

  // State of the fake mailbox
  await page.addInitScript(() => {
    window.__mail = [
      { id: 'A', conversationId: 'CONV1', subject: 'Offerte Q4', from: { emailAddress: { address: 'philip@example.com' } }, sentDateTime: '2026-11-01T10:00:00Z', parentFolderId: 'SENT', isDraft: false, webLink: 'https://x/A', categories: [] },
      { id: 'T', conversationId: 'CONV2', subject: 'Tagged in compose', from: { emailAddress: { address: 'philip@example.com' } }, sentDateTime: '2026-11-05T10:00:00Z', parentFolderId: 'SENT', isDraft: false, webLink: 'https://x/T', categories: ['Follow-up'] },
    ];
    window.__flags = {}; window.__events = {}; let evn = 0;
    window.__roam = {};
    const json = (o, s = 200) => new Response(o === null ? null : JSON.stringify(o), { status: s, headers: { 'Content-Type': 'application/json' } });
    const realFetch = window.fetch;
    window.fetch = async (url, opt = {}) => {
      if (!String(url).startsWith('https://graph.microsoft.com')) return realFetch(url, opt);
      const u = new URL(url); const p = u.pathname.replace('/v1.0', ''); const m = opt.method || 'GET';
      const body = opt.body ? JSON.parse(opt.body) : null;
      if (p === '/me') return json({ mail: 'philip@example.com', userPrincipalName: 'philip@example.com' });
      if (p === '/me/mailFolders/sentitems') return json({ id: 'SENT' });
      if (p === '/me/mailFolders/sentitems/messages') return json({ value: __mail.filter(x => x.parentFolderId === 'SENT' && x.categories.includes('Follow-up')) });
      if (p === '/me/outlook/masterCategories') return json(m === 'GET' ? { value: [] } : {}, m === 'GET' ? 200 : 201);
      if (p === '/me/messages' && m === 'GET') {
        const conv = decodeURIComponent(u.searchParams.get('$filter')).match(/'(.+)'/)[1];
        return json({ value: __mail.filter(x => x.conversationId === conv) });
      }
      let mm;
      if ((mm = p.match(/^\/me\/messages\/(.+)$/))) {
        if (m === 'GET') return json({ internetMessageHeaders: __mail.find(x => x.id === mm[1]).headers || [] });
        if (body.flag) __flags[mm[1]] = body.flag.flagStatus;
        if (body.categories) __mail.find(x => x.id === mm[1]).categories = body.categories;
        return json({});
      }
      if (p === '/me/events' && m === 'POST') { const id = 'EV' + (++evn); __events[id] = body; return json({ id }, 201); }
      if ((mm = p.match(/^\/me\/events\/(.+)$/))) {
        if (m === 'DELETE') { delete __events[mm[1]]; return new Response(null, { status: 204 }); }
        __events[mm[1]] = body; return json({ id: mm[1] });
      }
      return json({ error: { message: 'unmocked ' + m + ' ' + p } }, 400);
    };
    const ok = (v) => ({ status: 'succeeded', value: v });
    window.Office = {
      AsyncResultStatus: { Succeeded: 'succeeded' }, EventType: { ItemChanged: 'ic' }, MailboxEnums: { RestVersion: { v2_0: 'v2' } },
      onReady: (cb) => setTimeout(cb, 0),
      context: {
        requirements: { isSetSupported: () => true },
        roamingSettings: { get: k => __roam[k], set: (k, v) => { __roam[k] = v; }, saveAsync: cb => cb(ok()) },
        ui: {},
        mailbox: {
          userProfile: { emailAddress: 'philip@example.com' },
          addHandlerAsync: () => {},
          item: { subject: 'Offerte Q4', conversationId: 'CONV1' },
        },
      },
    };
    window.msal = { createNestablePublicClientApplication: async () => ({ acquireTokenSilent: async () => ({ accessToken: 't' }) }) };
  });

  await page.route('https://appsforoffice.microsoft.com/**', r => r.fulfill({ body: '/* stubbed */', contentType: 'text/javascript' }));
  await page.route('file://**/lib/msal-browser.min.js', r => r.fulfill({ body: '/* stubbed */', contentType: 'text/javascript' }));
  await page.goto('file://' + root + '/taskpane.html');
  await page.waitForTimeout(800);
  const text = async (id) => (await page.textContent('#' + id)).trim();
  const st = () => page.evaluate(() => ({ flags: { ...__flags }, events: Object.keys(__events).length, items: JSON.parse(__roam['fu.v1'] || '{"items":[]}').items }));

  console.log('auto-sync adopted compose-tagged email:', (await st()).items.map(i => i.s + '=' + i.st));
  console.log('current email status:', await text('currentStatus'));

  await page.click('#trackBtn'); await page.waitForTimeout(400);
  let s = await st();
  console.log('after Track: flag A =', s.flags.A, '| reminder events =', s.events, '| status:', await text('currentStatus'));

  // Reply arrives
  await page.evaluate(() => __mail.push({ id: 'B', conversationId: 'CONV1', subject: 'RE: Offerte Q4', from: { emailAddress: { address: 'client@acme.nl' } }, sentDateTime: '2026-11-03T09:00:00Z', parentFolderId: 'INBOX', isDraft: false, categories: [] }));
  await page.click('#syncBtn'); await page.waitForTimeout(400);
  s = await st();
  console.log('after reply: flag A =', s.flags.A, '| events =', s.events, '| answered list:', await text('answeredList'));

  // I send again
  await page.evaluate(() => __mail.push({ id: 'C', conversationId: 'CONV1', subject: 'RE: Offerte Q4', from: { emailAddress: { address: 'philip@example.com' } }, sentDateTime: '2026-11-04T09:00:00Z', parentFolderId: 'SENT', isDraft: false, categories: [] }));
  await page.click('#syncBtn'); await page.waitForTimeout(400);
  s = await st();
  console.log('after my new email: flag C =', s.flags.C, '| flag A =', s.flags.A, '| events =', s.events);

  // Out-of-office should not tick it off
  await page.evaluate(() => __mail.push({ id: 'D', conversationId: 'CONV1', subject: 'Automatisch antwoord: RE: Offerte Q4', from: { emailAddress: { address: 'client@acme.nl' } }, sentDateTime: '2026-11-04T09:01:00Z', parentFolderId: 'INBOX', isDraft: false, categories: [] }));
  await page.click('#syncBtn'); await page.waitForTimeout(400);
  s = await st();
  console.log('after out-of-office: CONV1 state =', s.items.find(i => i.c === 'CONV1').st);

  // Untrack via list X
  await page.click('#waitingList li:has-text("Tagged in compose") .x'); await page.waitForTimeout(400);
  s = await st();
  console.log('after untrack: items =', s.items.map(i => i.s), '| T categories =', await page.evaluate(() => __mail.find(x => x.id === 'T').categories));
  console.log('footer:', await text('footer'));
  await page.screenshot({ path: path.join(__dirname, 'pane.png'), fullPage: true });
  console.log('page errors:', errors.length ? errors : 'none');
  await browser.close();
})();
