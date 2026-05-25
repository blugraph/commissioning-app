// Generates icon-192.png and icon-512.png using the Canvas API via a headless approach.
// Run: node gen_icons.js
const { createCanvas } = require('canvas');
const fs = require('fs');

function drawIcon(size) {
  const c = createCanvas(size, size);
  const ctx = c.getContext('2d');
  const s = size;

  // Background
  ctx.fillStyle = '#1e3a5f';
  ctx.beginPath();
  ctx.roundRect(0, 0, s, s, s * 0.18);
  ctx.fill();

  // Water wave
  ctx.fillStyle = '#2563eb';
  const wy = s * 0.58;
  ctx.beginPath();
  ctx.moveTo(0, wy);
  const amp = s * 0.04, ww = s / 2;
  ctx.quadraticCurveTo(ww * 0.5, wy - amp, ww, wy);
  ctx.quadraticCurveTo(ww * 1.5, wy + amp, s, wy);
  ctx.lineTo(s, s);
  ctx.lineTo(0, s);
  ctx.closePath();
  ctx.fill();

  // Staff gauge (vertical bar)
  const bx = s * 0.44, bw = s * 0.12, by = s * 0.18, bh = s * 0.52;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(bx, by, bw, bh);

  // Tick marks
  ctx.fillStyle = '#1e3a5f';
  const ticks = 5;
  for (let i = 0; i <= ticks; i++) {
    const ty = by + (bh / ticks) * i;
    const tw = i % 2 === 0 ? bw * 0.7 : bw * 0.4;
    ctx.fillRect(bx + bw * 0.15, ty - 1, tw, 2);
  }

  // Water line indicator arrow
  ctx.fillStyle = '#60a5fa';
  const ax = bx + bw + s * 0.04, ay = wy - 1;
  ctx.beginPath();
  ctx.moveTo(ax, ay);
  ctx.lineTo(ax + s * 0.08, ay - s * 0.04);
  ctx.lineTo(ax + s * 0.08, ay + s * 0.04);
  ctx.closePath();
  ctx.fill();

  return c.toBuffer('image/png');
}

try {
  fs.writeFileSync('icon-192.png', drawIcon(192));
  fs.writeFileSync('icon-512.png', drawIcon(512));
  console.log('Icons generated.');
} catch (e) {
  console.error('canvas package not available:', e.message);
  process.exit(1);
}
