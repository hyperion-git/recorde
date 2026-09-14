// PNG fallback for the equation picture (Word keeps a raster blip next to the
// SVG; older readers show it). Sized like the add-in's canvas path
// (computeRasterSize: 300 DPI, 4 MP budget). Backends in order of preference:
// @resvg/resvg-js (optional npm dep, no external tools) → cairosvg (python) →
// ImageMagick `convert`. The first that works wins; the choice is reported.
import { spawnSync } from 'node:child_process';
import { computeRasterSize } from '../../src/raster.js';

let resvgMod;   // undefined = not tried, null = unavailable

export async function svgToPng(svgString, widthPt, heightPt) {
  const { pxW, pxH } = computeRasterSize(widthPt, heightPt, 1);
  if (resvgMod === undefined) {
    try { resvgMod = await import('@resvg/resvg-js'); } catch { resvgMod = null; }
  }
  if (resvgMod) {
    const r = new resvgMod.Resvg(svgString, { fitTo: { mode: 'width', value: pxW }, background: 'rgba(0,0,0,0)' });
    return { png: Buffer.from(r.render().asPng()), pxW, pxH, tool: 'resvg' };
  }
  const py = `import sys,cairosvg; sys.stdout.buffer.write(cairosvg.svg2png(bytestring=sys.stdin.buffer.read(), output_width=${pxW}, output_height=${pxH}))`;
  for (const [cmd, args] of [['python3', ['-c', py]], ['micromamba', ['run', '-n', 'sci-base', 'python', '-c', py]]]) {
    const r = spawnSync(cmd, args, { input: svgString, maxBuffer: 64 << 20 });
    if (r.status === 0 && r.stdout.length) return { png: r.stdout, pxW, pxH, tool: 'cairosvg' };
  }
  const r = spawnSync('convert', ['-background', 'none', '-density', '300', 'svg:-', '-resize', `${pxW}x${pxH}!`, 'png:-'],
    { input: svgString, maxBuffer: 64 << 20 });
  if (r.status === 0 && r.stdout.length) return { png: r.stdout, pxW, pxH, tool: 'imagemagick' };
  throw new Error('no SVG rasterizer available: install @resvg/resvg-js (npm), cairosvg (pip) or ImageMagick');
}
