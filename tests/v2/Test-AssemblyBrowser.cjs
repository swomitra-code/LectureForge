// Local Studio browser acceptance. Uses an isolated visible Edge profile.
const {spawn} = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(process.argv[2] || path.join(process.env.LOCALAPPDATA, 'LectureForgeE'));
const fixture = JSON.parse(fs.readFileSync(path.join(root, 'results.json'), 'utf8').replace(/^\uFEFF/, ''));
const id = fixture.project_id, url = fixture.url, port = 9239;
const checks = [], pending = new Map();
const profile = path.join(root, 'browser-' + Date.now());
fs.mkdirSync(profile);
const edge = spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', [
  '--no-first-run', '--disable-background-networking', '--disable-sync',
  '--remote-debugging-port=' + port, '--user-data-dir=' + profile, 'about:blank'
], {stdio: 'ignore'});
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
let ws, sequence = 0;
function command(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++sequence;
    const timer = setTimeout(() => { pending.delete(id); reject(Error('Browser timeout: ' + method)); }, 30000);
    pending.set(id, {resolve: value => {clearTimeout(timer); resolve(value);}, reject});
    ws.send(JSON.stringify({id, method, params}));
  });
}
async function evaluate(expression) {
  const r = await command('Runtime.evaluate', {expression, awaitPromise: true, returnByValue: true});
  if (r.exceptionDetails) throw Error(r.exceptionDetails.text);
  return r.result.value;
}
async function until(expression) {
  for (let i = 0; i < 120; i++) { if (await evaluate(expression)) return; await sleep(500); }
  throw Error('UI condition timed out: ' + expression);
}
function check(name, ok) {
  if (!ok) throw Error('FAIL: ' + name);
  checks.push({name, passed: true}); console.log('PASS: ' + name);
}
(async () => {
  try {
    let target;
    for (let i = 0; i < 60; i++) {
      try { target = (await (await fetch('http://127.0.0.1:' + port + '/json')).json()).find(t => t.type === 'page'); if (target) break; } catch {}
      await sleep(250);
    }
    if (!target) throw Error('Isolated Edge did not start');
    ws = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {ws.onopen = resolve; ws.onerror = reject;});
    ws.onmessage = event => {
      const r = JSON.parse(event.data), p = pending.get(r.id);
      if (p) { pending.delete(r.id); r.error ? p.reject(Error(r.error.message)) : p.resolve(r.result); }
    };
    await command('Page.enable');
    await command('Page.navigate', {url});
    await until("document.getElementById('open')!==null");
    await evaluate(`localStorage.setItem('lectureforge-project','${id}')`);
    await command('Page.reload');
    await until("document.getElementById('assembly-progress').textContent==='RECORDING POWERPOINT READY'");
    check('Browser restores the completed Recording PowerPoint', await evaluate("!document.getElementById('open-recording').hidden"));
    await until("document.getElementById('assembly-gate').textContent.includes('14 slides ready')");
    check('Readiness enables Create for all 14 slides', await evaluate("!document.getElementById('create-recording').disabled"));
    check('Dashboard retains 14 Avatar Ready and Slide 8 pause', await evaluate("document.getElementById('avatar-summary').textContent.includes('14 Avatar Ready')&&document.getElementById('avatar-jobs').textContent.includes('6.000 s pause verified')"));
    await evaluate("document.getElementById('placement-slide').value='6';document.getElementById('placement-slide').dispatchEvent(new Event('change'))");
    check('Placement form shows the Slide 6 override', await evaluate("Number(document.getElementById('placement-left').value)===84"));
    const before = await evaluate(`fetch('/api/assembly-status?id=${id}').then(r=>r.json())`);
    await evaluate(`document.getElementById('output-folder').value=${JSON.stringify(before.current.folder)};document.getElementById('create-recording').click()`);
    await sleep(4000);
    const after = await evaluate(`fetch('/api/assembly-status?id=${id}').then(r=>r.json())`);
    check('Create click reuses the identical completed assembly', before.current.id === after.current.id && after.current.stage === 'Complete');
    await evaluate("document.getElementById('open-recording').click();document.getElementById('open-output').click()");
    await sleep(1500);
    check('Open action controls execute without a Studio error', await evaluate("!document.getElementById('message').classList.contains('error')"));
    await command('Page.reload');
    await until("document.getElementById('assembly-progress').textContent==='RECORDING POWERPOINT READY'");
    check('Browser refresh preserves the final output path', await evaluate(`document.getElementById('recording-path').textContent===${JSON.stringify(before.current.output)}`));
    const avatars = await evaluate(`fetch('/api/avatar-status?id=${id}').then(r=>r.json())`);
    check('Browser actions make zero provider calls', avatars.provider_calls === 0 && avatars.ready === 14);
    await evaluate("document.getElementById('assembly-panel').scrollIntoView()");
    const shot = await command('Page.captureScreenshot', {format: 'png'});
    fs.writeFileSync(path.join(root, 'assembly-browser.png'), Buffer.from(shot.data, 'base64'));
    fs.writeFileSync(path.join(root, 'browser-results.json'), JSON.stringify({checks, elevenlabs_calls: 0, heygen_calls: 0}, null, 2));
  } finally { if (ws) ws.close(); edge.kill(); }
})().catch(error => {console.error(error.message); process.exitCode = 1;});
