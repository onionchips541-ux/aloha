const fs = require('fs'), p = require('path'), nm = (...a) => p.join('node_modules', ...a);

function exigir(caminho) {
  if (!fs.existsSync(caminho)) {
    console.error('ERRO: arquivo não encontrado: ' + caminho);
    console.error('Arquivos na raiz do projeto: ' + fs.readdirSync('.').join(', '));
    console.error('Confira se as pastas src/ e assets/ foram enviadas para a raiz do repositório.');
    process.exit(1);
  }
}

exigir('src/index.html');
fs.rmSync('www', { recursive: true, force: true });
fs.mkdirSync('www/ocr', { recursive: true });
fs.copyFileSync('src/index.html', 'www/index.html');
if (fs.existsSync('assets')) fs.cpSync('assets', 'www/assets', { recursive: true });
else console.warn('AVISO: pasta assets/ não encontrada; logos da tela de carregamento não aparecerão.');

fs.copyFileSync(nm('tesseract.js', 'dist', 'tesseract.min.js'), 'www/ocr/tesseract.min.js');
fs.copyFileSync(nm('tesseract.js', 'dist', 'worker.min.js'), 'www/ocr/worker.min.js');
for (const f of fs.readdirSync(nm('tesseract.js-core')))
  if (/^tesseract-core.*\.js$/.test(f)) fs.copyFileSync(nm('tesseract.js-core', f), 'www/ocr/' + f);
fs.copyFileSync(nm('@tesseract.js-data', 'por', '4.0.0_best_int', 'por.traineddata.gz'), 'www/ocr/por.traineddata.gz');
console.log('www pronto');

fs.mkdirSync('www/lib', { recursive: true });
fs.copyFileSync(nm('jspdf', 'dist', 'jspdf.umd.min.js'), 'www/lib/jspdf.umd.min.js');
fs.copyFileSync(nm('jspdf-autotable', 'dist', 'jspdf.plugin.autotable.min.js'), 'www/lib/jspdf.plugin.autotable.min.js');
