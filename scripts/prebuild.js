const { execSync } = require('child_process');
const path = require('path');

const betterSqlite3Dir = path.join(__dirname, '..', 'node_modules', 'better-sqlite3');
const prebuildInstall = path.join(__dirname, '..', 'node_modules', 'prebuild-install', 'bin.js');

try {
  execSync(`node "${prebuildInstall}"`, {
    cwd: betterSqlite3Dir,
    stdio: 'inherit'
  });
  console.log('better-sqlite3 prebuilt binary ready');
} catch (e) {
  console.error('prebuild-install failed:', e.message);
  process.exit(1);
}
