/** Conversioni fra testo, byte, base64 e data URI. */

import type { FileData } from '../../core';

export function bytesToBase64(bytes: Uint8Array): string {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64.replace(/\s+/g, ''));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function toDataUri(bytes: Uint8Array, mimeType: string): string {
  return `data:${mimeType};base64,${bytesToBase64(bytes)}`;
}

/** Decodifica un data URI; `null` se `uri` non lo è. */
export function parseDataUri(uri: string): { mimeType: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]*)((?:;[^;,]*)*?)(;base64)?,(.*)$/s.exec(uri);
  if (!m) return null;
  const mimeType = m[1] || 'application/octet-stream';
  const payload = m[4] ?? '';
  const bytes = m[3] ? base64ToBytes(payload) : new TextEncoder().encode(decodeURIComponent(payload));
  return { mimeType, bytes };
}

/** Testo da `FileData`: i byte sono letti come UTF-8 (o UTF-16 con BOM); il BOM viene tolto. */
export function dataToText(data: FileData): string {
  if (typeof data === 'string') return data.replace(/^﻿/, '');
  if (data[0] === 0xff && data[1] === 0xfe) return new TextDecoder('utf-16le').decode(data.subarray(2));
  if (data[0] === 0xfe && data[1] === 0xff) return new TextDecoder('utf-16be').decode(data.subarray(2));
  return new TextDecoder('utf-8').decode(data).replace(/^﻿/, '');
}

export function dataToBytes(data: FileData): Uint8Array {
  return typeof data === 'string' ? new TextEncoder().encode(data) : data;
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export function mimeFromExtension(name: string): string {
  const ext = name.split('.').pop()?.toLowerCase() ?? '';
  return MIME_BY_EXT[ext] ?? 'application/octet-stream';
}

export function extensionFromMime(mime: string): string {
  for (const [ext, m] of Object.entries(MIME_BY_EXT)) if (m === mime) return ext;
  return 'bin';
}
