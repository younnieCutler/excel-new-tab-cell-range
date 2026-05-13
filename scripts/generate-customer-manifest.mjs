import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const DEFAULT_BASE_URL = 'https://cellfocus.example.com';
const DEFAULT_SUPPORT_URL = `${DEFAULT_BASE_URL}/support.html`;

function getArg(name) {
    const prefix = `--${name}=`;
    const inline = process.argv.find((arg) => arg.startsWith(prefix));
    if (inline) return inline.slice(prefix.length);

    const index = process.argv.indexOf(`--${name}`);
    return index >= 0 ? process.argv[index + 1] : undefined;
}

function normalizeBaseUrl(value) {
    if (!value) {
        throw new Error('Missing required --base-url argument.');
    }
    const url = new URL(value);
    if (url.protocol !== 'https:') {
        throw new Error('Base URL must use HTTPS.');
    }
    return url.href.replace(/\/$/, '');
}

function normalizeSupportUrl(value, baseUrl) {
    if (!value) return `${baseUrl}/support.html`;
    const url = new URL(value);
    if (url.protocol !== 'https:') {
        throw new Error('Support URL must use HTTPS.');
    }
    return url.href.replace(/\/$/, '');
}

const baseUrl = normalizeBaseUrl(getArg('base-url'));
const supportUrl = normalizeSupportUrl(getArg('support-url'), baseUrl);
const output = resolve(getArg('out') ?? 'dist/manifest.xml');
const label = getArg('label') ?? 'deployment';
const appDomain = new URL(baseUrl).origin;

let manifest = readFileSync(resolve('manifest.xml'), 'utf8');
manifest = manifest
    .replaceAll(DEFAULT_SUPPORT_URL, supportUrl)
    .replaceAll(DEFAULT_BASE_URL, baseUrl)
    .replace(`<AppDomain>${baseUrl}</AppDomain>`, `<AppDomain>${appDomain}</AppDomain>`);

mkdirSync(dirname(output), { recursive: true });
writeFileSync(output, manifest.endsWith('\n') ? manifest : `${manifest}\n`);

console.log(`Generated ${output}`);
console.log(`Profile: ${label}`);
console.log(`Base URL: ${baseUrl}`);
console.log(`AppDomain: ${appDomain}`);
console.log(`Support URL: ${supportUrl}`);
