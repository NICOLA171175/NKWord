import { describe, expect, it } from 'vitest';
import { normalizeColor } from '../../src/io/util/color';
import { base64ToBytes, bytesToBase64, dataToText, parseDataUri, toDataUri } from '../../src/io/util/encoding';
import { child, elements, parseHtml, parseXml, textContent, type MarkupElement } from '../../src/io/util/markup';
import { crc32, readZip, writeZip } from '../../src/io/util/zip';

const names = (el: MarkupElement) => elements(el).map((e) => e.name);

describe('markup', () => {
  it('analizza XML con prefissi, entità, CDATA e tag autochiusi', () => {
    const root = parseXml('<?xml version="1.0"?><w:p a="x &amp; y"><w:t xml:space="preserve"> a &lt;b&gt; &#233;</w:t><w:br/><![CDATA[<raw>]]></w:p>');
    const p = elements(root)[0]!;
    expect(p.name).toBe('w:p');
    expect(p.attrs.a).toBe('x & y');
    expect(textContent(child(p, 'w:t')!)).toBe(' a <b> é');
    expect(names(p)).toEqual(['w:t', 'w:br']);
    expect(p.children[2]).toBe('<raw>');
  });

  it('analizza HTML malformato come un browser', () => {
    const root = parseHtml('<P class=x>uno<p>due<ul><li>a<li>b</ul><BR><img src="i.png">tre</div>&nbsp;&euro;');
    const ps = elements(root, 'p');
    expect(ps).toHaveLength(2);
    expect(ps[0]!.attrs.class).toBe('x');
    expect(textContent(ps[0]!)).toBe('uno');
    const ul = elements(root, 'ul')[0]!;
    expect(elements(ul, 'li').map(textContent)).toEqual(['a', 'b']);
    expect(names(root)).toEqual(['p', 'p', 'ul', 'br', 'img']);
    expect(root.children.at(-1)).toBe('tre €');
  });

  it('tratta script e style come testo grezzo e ignora commenti', () => {
    const root = parseHtml('<style>p > b { color: red }</style><!-- <p>no</p> --><p>sì</p>');
    expect(textContent(elements(root, 'style')[0]!)).toBe('p > b { color: red }');
    expect(elements(root, 'p').map(textContent)).toEqual(['sì']);
  });
});

describe('zip', () => {
  it('scrive e rilegge file compressi e non compressi', async () => {
    const enc = new TextEncoder();
    const big = enc.encode('ciao '.repeat(1000));
    const small = enc.encode('x');
    const bin = new Uint8Array([0, 1, 2, 255]);
    const zip = await writeZip([
      ['a.txt', big],
      ['cartella/b.txt', small],
      ['àccento.bin', bin],
    ]);
    expect(zip.length).toBeLessThan(big.length);
    const out = await readZip(zip);
    expect([...out.keys()]).toEqual(['a.txt', 'cartella/b.txt', 'àccento.bin']);
    expect(out.get('a.txt')).toEqual(big);
    expect(out.get('cartella/b.txt')).toEqual(small);
    expect(out.get('àccento.bin')).toEqual(bin);
  });

  it('calcola il CRC-32 standard', () => {
    expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926);
  });

  it('rifiuta dati che non sono ZIP', async () => {
    await expect(readZip(new Uint8Array(30))).rejects.toThrow(/ZIP/);
  });
});

describe('codifiche e colori', () => {
  it('converte base64 e data URI', () => {
    const bytes = new Uint8Array([137, 80, 78, 71, 0, 255]);
    expect(base64ToBytes(bytesToBase64(bytes))).toEqual(bytes);
    const uri = toDataUri(bytes, 'image/png');
    expect(parseDataUri(uri)).toEqual({ mimeType: 'image/png', bytes });
    expect(parseDataUri('https://x/y.png')).toBeNull();
  });

  it('decodifica testo con BOM UTF-8 e UTF-16', () => {
    expect(dataToText(new Uint8Array([0xef, 0xbb, 0xbf, 0x63, 0x69, 0x61, 0x6f]))).toBe('ciao');
    expect(dataToText(new Uint8Array([0xff, 0xfe, 0x61, 0x00, 0xe8, 0x00]))).toBe('aè');
  });

  it('normalizza i colori CSS e DOCX', () => {
    expect(normalizeColor('#ABC')).toBe('#aabbcc');
    expect(normalizeColor('FF0000')).toBe('#ff0000');
    expect(normalizeColor('rgb(0, 128, 255)')).toBe('#0080ff');
    expect(normalizeColor('rgba(0,0,0,0)')).toBeUndefined();
    expect(normalizeColor('auto')).toBeUndefined();
    expect(normalizeColor('Navy')).toBe('#000080');
  });
});
