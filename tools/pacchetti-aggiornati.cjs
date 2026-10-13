// Usato da avvia.bat / avvia.sh: esce con 0 se node_modules corrisponde a
// package-lock.json, altrimenti con 1 (allora il lanciatore riesegue npm install).
// Serve a chi aggiorna il progetto senza cancellare la vecchia cartella node_modules.
const fs = require('fs');
const path = require('path');

const radice = path.join(__dirname, '..');
try {
  const voluti = JSON.parse(fs.readFileSync(path.join(radice, 'package-lock.json'), 'utf8')).packages;
  const installati = JSON.parse(
    fs.readFileSync(path.join(radice, 'node_modules', '.package-lock.json'), 'utf8'),
  ).packages;
  const allineati = Object.keys(voluti).every((nome) => {
    const voluto = voluti[nome];
    // La radice e i pacchetti facoltativi (es. binari per un altro sistema) non contano.
    if (nome === '' || voluto.optional || voluto.peer) return true;
    return installati[nome] !== undefined && installati[nome].version === voluto.version;
  });
  process.exit(allineati ? 0 : 1);
} catch {
  process.exit(1);
}
