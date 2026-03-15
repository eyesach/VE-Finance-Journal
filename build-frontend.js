const fs = require('fs');
const path = require('path');

const dist = path.join(__dirname, 'dist');

// Clean dist
if (fs.existsSync(dist)) {
  fs.rmSync(dist, { recursive: true });
}
fs.mkdirSync(dist);

// Copy index.html
fs.copyFileSync('index.html', path.join(dist, 'index.html'));

// Copy directories
function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

for (const dir of ['css', 'js', 'lib']) {
  if (fs.existsSync(dir)) {
    copyDir(dir, path.join(dist, dir));
  }
}

console.log('Frontend assets copied to dist/');
