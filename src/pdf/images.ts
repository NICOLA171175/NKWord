/**
 * Immagini per il PDF. JPEG e PNG si incorporano senza ricodificarli
 * (DCTDecode; FlateDecode con i predittori PNG); i PNG con trasparenza si
 * decomprimono per separare il canale alfa in una maschera (SMask). Gli
 * altri formati si possono convertire in PNG con un `ImageRasterizer`, che
 * nel browser usa un canvas.
 */

import { parseDataUri } from '../io/util/encoding';

/** Immagine pronta da scrivere come XObject. */
export interface PdfImage {
  readonly width: number;
  readonly height: number;
  /** Voci del dizionario oltre a Type, Subtype, Width, Height e Length. */
  readonly dict: string;
  readonly data: Uint8Array;
  /** Maschera di trasparenza (scala di grigi, 8 bit). */
  readonly alpha?: Uint8Array;
  readonly alphaFilter?: string;
}

/** Converte un'immagine (URL o data URI) in PNG; `null` se non ci riesce. */
export type ImageRasterizer = (src: string) => Promise<Uint8Array | null>;

const u32 = (b: Uint8Array, i: number) => ((b[i]! << 24) | (b[i + 1]! << 16) | (b[i + 2]! << 8) | b[i + 3]!) >>> 0;

export function isJpeg(b: Uint8Array): boolean {
  return b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;
}

export function isPng(b: Uint8Array): boolean {
  return b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
}

/** JPEG: dimensioni e componenti dal marcatore SOF. */
export function jpegImage(b: Uint8Array): PdfImage | null {
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) return null;
    const marker = b[i + 1]!;
    if (marker === 0xff) {
      i++;
      continue;
    }
    const len = (b[i + 2]! << 8) | b[i + 3]!;
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      const height = (b[i + 5]! << 8) | b[i + 6]!;
      const width = (b[i + 7]! << 8) | b[i + 8]!;
      const comps = b[i + 9]!;
      const cs = comps === 1 ? '/DeviceGray' : comps === 4 ? '/DeviceCMYK /Decode [1 0 1 0 1 0 1 0]' : '/DeviceRGB';
      if (!width || !height) return null;
      return { width, height, dict: `/ColorSpace ${cs} /BitsPerComponent 8 /Filter /DCTDecode`, data: b };
    }
    i += 2 + len;
  }
  return null;
}

interface PngInfo {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlace: number;
  palette?: Uint8Array;
  trns: boolean;
  idat: Uint8Array;
}

function readPng(b: Uint8Array): PngInfo | null {
  let i = 8;
  let info: Omit<PngInfo, 'idat' | 'trns'> | undefined;
  let trns = false;
  const idat: Uint8Array[] = [];
  while (i + 8 <= b.length) {
    const len = u32(b, i);
    const type = String.fromCharCode(b[i + 4]!, b[i + 5]!, b[i + 6]!, b[i + 7]!);
    const data = b.subarray(i + 8, i + 8 + len);
    if (type === 'IHDR') {
      info = { width: u32(data, 0), height: u32(data, 4), bitDepth: data[8]!, colorType: data[9]!, interlace: data[12]! };
    } else if (type === 'PLTE' && info) info.palette = data;
    else if (type === 'tRNS') trns = true;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    i += 12 + len;
  }
  if (!info || idat.length === 0) return null;
  const all = new Uint8Array(idat.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const c of idat) {
    all.set(c, at);
    at += c.length;
  }
  return { ...info, trns, idat: all };
}

async function pipe(data: Uint8Array, stream: CompressionStream | DecompressionStream): Promise<Uint8Array> {
  const out = new Response(new Blob([data as BlobPart]).stream().pipeThrough(stream));
  return new Uint8Array(await out.arrayBuffer());
}

/** Toglie i filtri di riga PNG (8 bit per campione). */
export function unfilterPng(raw: Uint8Array, width: number, height: number, bpp: number): Uint8Array {
  const stride = width * bpp;
  const out = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)]!;
    const src = y * (stride + 1) + 1;
    const dst = y * stride;
    for (let x = 0; x < stride; x++) {
      const v = raw[src + x]!;
      const a = x >= bpp ? out[dst + x - bpp]! : 0;
      const up = y > 0 ? out[dst - stride + x]! : 0;
      const c = x >= bpp && y > 0 ? out[dst - stride + x - bpp]! : 0;
      let p: number;
      switch (filter) {
        case 1:
          p = a;
          break;
        case 2:
          p = up;
          break;
        case 3:
          p = (a + up) >> 1;
          break;
        case 4: {
          const e = a + up - c;
          const pa = Math.abs(e - a);
          const pb = Math.abs(e - up);
          const pc = Math.abs(e - c);
          p = pa <= pb && pa <= pc ? a : pb <= pc ? up : c;
          break;
        }
        default:
          p = 0;
      }
      out[dst + x] = (v + p) & 0xff;
    }
  }
  return out;
}

const canStream = () => typeof DecompressionStream !== 'undefined' && typeof CompressionStream !== 'undefined';

export async function pngImage(b: Uint8Array): Promise<PdfImage | null> {
  const png = readPng(b);
  if (!png || png.interlace !== 0) return null;
  const { width, height, bitDepth, colorType } = png;
  // Senza trasparenza i dati compressi del PNG vanno bene così come sono.
  if ((colorType === 0 || colorType === 2 || colorType === 3) && !png.trns) {
    const colors = colorType === 2 ? 3 : 1;
    let cs = colorType === 2 ? '/DeviceRGB' : '/DeviceGray';
    if (colorType === 3) {
      if (!png.palette) return null;
      let hex = '';
      for (const v of png.palette) hex += v.toString(16).padStart(2, '0');
      cs = `[/Indexed /DeviceRGB ${png.palette.length / 3 - 1} <${hex}>]`;
    }
    return {
      width,
      height,
      dict: `/ColorSpace ${cs} /BitsPerComponent ${bitDepth} /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors ${colors} /BitsPerComponent ${bitDepth} /Columns ${width} >>`,
      data: png.idat,
    };
  }
  // Canale alfa: si decomprime e si separa (solo 8 bit, il caso comune).
  if ((colorType === 4 || colorType === 6) && bitDepth === 8 && canStream()) {
    const colors = colorType === 6 ? 3 : 1;
    const bpp = colors + 1;
    const pixels = unfilterPng(await pipe(png.idat, new DecompressionStream('deflate')), width, height, bpp);
    const color = new Uint8Array(width * height * colors);
    const alpha = new Uint8Array(width * height);
    for (let i = 0, c = 0; i < width * height; i++) {
      for (let k = 0; k < colors; k++) color[c++] = pixels[i * bpp + k]!;
      alpha[i] = pixels[i * bpp + colors]!;
    }
    return {
      width,
      height,
      dict: `/ColorSpace ${colors === 3 ? '/DeviceRGB' : '/DeviceGray'} /BitsPerComponent 8 /Filter /FlateDecode`,
      data: await pipe(color, new CompressionStream('deflate')),
      alpha: await pipe(alpha, new CompressionStream('deflate')),
      alphaFilter: '/FlateDecode',
    };
  }
  return null;
}

/** Nel browser: disegna l'immagine su un canvas e la riprende come PNG. */
export const browserRasterizer: ImageRasterizer = async (src) => {
  if (typeof document === 'undefined' || typeof Image === 'undefined') return null;
  try {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.src = src;
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext('2d');
    if (!ctx || !canvas.width || !canvas.height) return null;
    ctx.drawImage(img, 0, 0);
    const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch {
    return null;
  }
};

/** Carica un'immagine del documento; `null` se il formato non è gestito. */
export async function loadImage(src: string, rasterize?: ImageRasterizer): Promise<PdfImage | null> {
  const parsed = src.startsWith('data:') ? parseDataUri(src) : null;
  const bytes = parsed?.bytes;
  try {
    if (bytes && isJpeg(bytes)) return jpegImage(bytes);
    if (bytes && isPng(bytes)) {
      const img = await pngImage(bytes);
      if (img) return img;
    }
    if (!rasterize) return null;
    const png = await rasterize(src);
    return png && isPng(png) ? await pngImage(png) : null;
  } catch {
    return null;
  }
}
