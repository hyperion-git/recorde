// Minimal HTTPS static-file server for the add-in. Uses the cert pair
// installed by `office-addin-dev-certs install`, which lives under the
// current OS user's home — `os.homedir()` handles spaces in the path
// (e.g. a user name containing a space) correctly because no shell expansion
// is involved.
//
// Serves the same layout build.js publishes: /src/<f> is addin/src/<f> or,
// failing that, core/<f> (the pane imports the shared modules as ./x.js);
// /assets/… is addin/assets/…; / and /index.html are the pane and the support page.

const fs = require('fs');
const path = require('path');
const https = require('https');
const os = require('os');

const PORT = 3000;
const CERT_DIR = path.join(os.homedir(), '.office-addin-dev-certs');
const ADDIN = __dirname;
const CORE = path.join(ADDIN, '..', 'core');

const mime = {
  '.html': 'text/html',
  '.js':   'application/javascript',
  '.mjs':  'application/javascript',
  '.css':  'text/css',
  '.json': 'application/json',
  '.svg':  'image/svg+xml',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
};

const opts = {
  cert: fs.readFileSync(path.join(CERT_DIR, 'localhost.crt')),
  key:  fs.readFileSync(path.join(CERT_DIR, 'localhost.key')),
};

https.createServer(opts, (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-cache');

  let pathname = decodeURIComponent(new URL(req.url, 'https://localhost').pathname);
  if (pathname === '/') pathname = '/src/taskpane.html';

  const candidates = [path.normalize(path.join(ADDIN, pathname))];
  if (pathname.startsWith('/src/')) candidates.push(path.normalize(path.join(CORE, pathname.slice('/src/'.length))));
  const filePath = candidates.find((p) => (p.startsWith(ADDIN) || p.startsWith(CORE)) && fs.existsSync(p));
  if (!filePath) {
    res.writeHead(404); res.end(`Not found: ${pathname}`); return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) { res.writeHead(404); res.end(`Not found: ${pathname}`); return; }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': mime[ext] || 'application/octet-stream' });
    res.end(data);
  });
}).listen(PORT, () => {
  console.log(`Serving ${ADDIN} (+ core/ under /src/) on https://localhost:${PORT}`);
});
