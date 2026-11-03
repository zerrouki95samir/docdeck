/* Regenerate app icon assets: build/icon.png + build/icon.iconset -> icon.icns via iconutil. */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { Resvg } from '@resvg/resvg-js';

const svg = readFileSync('public/icon.svg', 'utf8');
const render = (px) => new Resvg(svg, { fitTo: { mode: 'width', value: px } }).render().asPng();
writeFileSync('build/icon.png', render(512));
writeFileSync('build/icon.svg', svg);
mkdirSync('build/icon.iconset', { recursive: true });
const names = {
  'icon_16x16': 16, 'icon_16x16@2x': 32, 'icon_32x32': 32, 'icon_32x32@2x': 64,
  'icon_128x128': 128, 'icon_128x128@2x': 256, 'icon_256x256': 256, 'icon_256x256@2x': 512,
  'icon_512x512': 512, 'icon_512x512@2x': 1024,
};
for (const [name, px] of Object.entries(names)) writeFileSync(`build/icon.iconset/${name}.png`, render(px));
execSync('iconutil -c icns build/icon.iconset -o build/icon.icns');
console.log('icons: build/icon.png + build/icon.icns');
