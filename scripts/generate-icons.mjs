import { writeFileSync, mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { deflateSync } from 'node:zlib';

function crc32(buf) {
    let crc = 0xffffffff;
    for (const byte of buf) {
        crc ^= byte;
        for (let i = 0; i < 8; i++) {
            crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
        }
    }
    return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
    const typeBuf = Buffer.from(type, 'ascii');
    const out = Buffer.alloc(12 + data.length);
    out.writeUInt32BE(data.length, 0);
    typeBuf.copy(out, 4);
    data.copy(out, 8);
    out.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 8 + data.length);
    return out;
}

function makePng(size) {
    const raw = Buffer.alloc((size * 4 + 1) * size);
    for (let y = 0; y < size; y++) {
        const row = y * (size * 4 + 1);
        raw[row] = 0;
        for (let x = 0; x < size; x++) {
            const i = row + 1 + x * 4;
            const edge = x === 0 || y === 0 || x === size - 1 || y === size - 1;
            const diagonal = Math.abs(x - y) <= Math.max(1, Math.floor(size / 16));
            const band = x > size * 0.22 && x < size * 0.78 && y > size * 0.32 && y < size * 0.68;
            raw[i] = edge ? 15 : diagonal ? 0 : band ? 0 : 26;
            raw[i + 1] = edge ? 23 : diagonal ? 210 : band ? 210 : 26;
            raw[i + 2] = edge ? 42 : diagonal ? 255 : band ? 255 : 46;
            raw[i + 3] = 255;
        }
    }

    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(size, 0);
    ihdr.writeUInt32BE(size, 4);
    ihdr[8] = 8;
    ihdr[9] = 6;
    ihdr[10] = 0;
    ihdr[11] = 0;
    ihdr[12] = 0;

    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        chunk('IHDR', ihdr),
        chunk('IDAT', deflateSync(raw)),
        chunk('IEND', Buffer.alloc(0)),
    ]);
}

mkdirSync(resolve('assets'), { recursive: true });
for (const size of [16, 32, 64, 80]) {
    writeFileSync(resolve(`assets/icon-${size}.png`), makePng(size));
    console.log(`Generated assets/icon-${size}.png`);
}
