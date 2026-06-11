// SAFA Dashboard — static server with password login
// Set DASHBOARD_PASSWORD in Railway (Variables tab) to enable the login page.
const express = require('express');
const crypto = require('crypto');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const PASSWORD = process.env.DASHBOARD_PASSWORD || '';
const SECRET = process.env.SESSION_SECRET || crypto.createHash('sha256').update('safa::' + PASSWORD).digest('hex');
const COOKIE = 'safa_auth';
const MAX_AGE_DAYS = 30;

const sign = v => crypto.createHmac('sha256', SECRET).update(v).digest('hex');
const expectedToken = () => sign('ok::' + PASSWORD);

function getCookie(req, name) {
  const c = req.headers.cookie || '';
  const m = c.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]*)'));
  return m ? decodeURIComponent(m[1]) : null;
}
function isAuthed(req) {
  if (!PASSWORD) return true; // no password configured -> open (local dev)
  const t = getCookie(req, COOKIE);
  return !!t && crypto.timingSafeEqual(Buffer.from(t.padEnd(64).slice(0, 64)), Buffer.from(expectedToken()));
}

const loginPage = (error) => `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SAFA — Sign in</title><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:#f7f7f8;color:#0e0e10;display:flex;align-items:center;justify-content:center;min-height:100vh;-webkit-font-smoothing:antialiased}
.card{background:#fff;border:1px solid #e6e6ea;border-radius:16px;padding:40px 36px;width:94vw;max-width:380px;box-shadow:0 1px 2px rgba(14,14,16,.04),0 8px 32px rgba(14,14,16,.07)}
h1{font-size:20px;font-weight:800;letter-spacing:-.02em}
h1 span{color:#8a8a93;font-weight:600}
p{color:#8a8a93;font-size:13px;margin:6px 0 24px}
label{display:block;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#8a8a93;margin-bottom:6px}
input{width:100%;border:1px solid #e6e6ea;border-radius:10px;padding:12px 14px;font-size:15px;outline:none;font-family:inherit}
input:focus{border-color:#0e0e10}
button{width:100%;margin-top:14px;background:#0e0e10;color:#fff;border:none;border-radius:10px;padding:12px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit}
button:hover{background:#26262b}
.err{background:#fee4e2;color:#b42318;border-radius:8px;padding:9px 12px;font-size:12.5px;font-weight:600;margin-bottom:14px}
</style></head><body>
<form class="card" method="POST" action="/login">
  <h1>SAFA <span>/ Incoming Shipments</span></h1>
  <p>Enter the team password to continue.</p>
  ${error ? '<div class="err">Wrong password — try again.</div>' : ''}
  <label for="pw">Password</label>
  <input id="pw" name="password" type="password" autofocus autocomplete="current-password">
  <button type="submit">Sign in</button>
</form></body></html>`;

app.use(express.urlencoded({ extended: false }));

app.get('/login', (req, res) => {
  if (isAuthed(req)) return res.redirect('/');
  res.send(loginPage(false));
});
app.post('/login', (req, res) => {
  const given = String(req.body.password || '');
  const a = crypto.createHash('sha256').update(given).digest();
  const b = crypto.createHash('sha256').update(PASSWORD).digest();
  if (PASSWORD && crypto.timingSafeEqual(a, b)) {
    res.setHeader('Set-Cookie',
      `${COOKIE}=${expectedToken()}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${MAX_AGE_DAYS * 86400}` +
      (req.secure || req.headers['x-forwarded-proto'] === 'https' ? '; Secure' : ''));
    return res.redirect('/');
  }
  res.status(401).send(loginPage(true));
});
app.get('/logout', (req, res) => {
  res.setHeader('Set-Cookie', `${COOKIE}=; Path=/; HttpOnly; Max-Age=0`);
  res.redirect('/login');
});

app.use((req, res, next) => {
  if (isAuthed(req)) return next();
  res.redirect('/login');
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, () => console.log(`SAFA dashboard on port ${PORT}${PASSWORD ? ' (password protected)' : ' (NO PASSWORD SET — open access)'}`));
