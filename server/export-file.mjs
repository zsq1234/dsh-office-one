import { createServer } from 'node:http';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, extname } from 'node:path';
import { ExchangeFormat, exportToFile, importBuffer } from '@univerjs-pro/exchange-node';
import { UniverInstanceType } from '@univerjs/core';
import { normalizeSlides } from './normalize-slides.mjs';

const port = Number(process.env.PORT ?? 8787);
const maxBodyBytes = 300 * 1024 * 1024;
const units = {
  sheet: UniverInstanceType.UNIVER_SHEET,
  doc: UniverInstanceType.UNIVER_DOC,
  slide: UniverInstanceType.UNIVER_SLIDE,
};
const formats = new Set(Object.values(ExchangeFormat));
const importFormats = {
  sheet: new Set([ExchangeFormat.XLS, ExchangeFormat.XLSX, ExchangeFormat.CSV, ExchangeFormat.TSV]),
  doc: new Set([ExchangeFormat.DOC, ExchangeFormat.DOCX]),
  slide: new Set([ExchangeFormat.PPT, ExchangeFormat.PPTX]),
};

function corsOrigin(request) {
  return process.env.CORS_ORIGIN ?? request.headers.origin ?? 'http://127.0.0.1:3080';
}

function sendJson(request, response, status, payload) {
  response.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'access-control-allow-origin': corsOrigin(request),
  });
  response.end(JSON.stringify(payload));
}

async function readBody(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBodyBytes) throw new Error('Request body is too large.');
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function readJson(request) {
  return JSON.parse((await readBody(request)).toString('utf8'));
}

function importRequest(request, payload) {
  const contentType = request.headers['content-type'] ?? '';
  if (contentType.startsWith('application/json')) {
    if (typeof payload?.data !== 'string') throw new Error('JSON import requests require base64 data.');
    return Buffer.from(payload.data, 'base64');
  }
  return payload.data;
}

function importFormat(unitType, format, fileName) {
  const inferred = format || extname(fileName).slice(1).toLowerCase();
  if (!importFormats[unitType]?.has(inferred)) return null;
  return inferred;
}

const server = createServer(async (request, response) => {
  if (request.method === 'OPTIONS') {
    response.writeHead(204, {
      'access-control-allow-origin': corsOrigin(request),
      'access-control-allow-methods': 'POST, OPTIONS',
      'access-control-allow-headers': 'content-type',
    });
    response.end();
    return;
  }

  const isExport = request.method === 'POST' && request.url === '/api/univer/export-file';
  const isImport = request.method === 'POST' && request.url.startsWith('/api/univer/import-file');
  if (!isExport && !isImport) {
    sendJson(request, response, 404, { error: 'Not found.' });
    return;
  }

  let directory;
  try {
    if (isImport) {
      const contentType = request.headers['content-type'] ?? '';
      const url = new URL(request.url, 'http://127.0.0.1');
      const payload = contentType.startsWith('application/json')
        ? await readJson(request)
        : {
            unitType: url.searchParams.get('unitType'),
            format: url.searchParams.get('format'),
            fileName: url.searchParams.get('fileName'),
            data: await readBody(request),
          };
      const unitType = typeof payload?.unitType === 'string' ? payload.unitType : '';
      const fileName = typeof payload?.fileName === 'string' ? payload.fileName : '';
      const format = importFormat(unitType, typeof payload?.format === 'string' ? payload.format.toLowerCase() : '', fileName);
      const input = importRequest(request, payload);
      const type = units[unitType];
      if (!type || !format || !fileName || input.byteLength === 0) {
        sendJson(request, response, 400, {
          error: 'Expected a file plus unitType (sheet, doc, or slide), fileName, and a supported format.',
        });
        return;
      }
      const data = await importBuffer(input, { type, format, fileName });
      if (unitType === 'slide') normalizeSlides(data);
      sendJson(request, response, 200, { unitType, format, fileName, data });
      return;
    }

    const payload = await readJson(request);
    const unitType = typeof payload?.unitType === 'string' ? payload.unitType : '';
    const format = typeof payload?.format === 'string' ? payload.format.toLowerCase() : '';
    const type = units[unitType];
    if (!type || !formats.has(format) || !payload?.data || typeof payload.data !== 'object') {
      sendJson(request, response, 400, {
        error: 'Expected { unitType, format, data } with a supported Univer Unit.',
      });
      return;
    }

    directory = await mkdtemp(join(tmpdir(), 'dsh-univer-export-'));
    const outputPath = join(directory, `document.${format}`);
    if (unitType === 'slide') normalizeSlides(payload.data);
    const options = { type, format };
    if (unitType === 'sheet' && (format === ExchangeFormat.CSV || format === ExchangeFormat.TSV)) {
      const worksheetId = payload.data?.sheetOrder?.[0] ?? Object.keys(payload.data?.sheets ?? {})[0];
      if (!worksheetId) throw new Error('CSV/TSV export requires a workbook with at least one worksheet.');
      options.csv = { worksheetId };
    }
    await exportToFile(payload.data, outputPath, options);
    const file = await readFile(outputPath);
    const filename = `univer-export${extname(outputPath)}`;

    response.writeHead(200, {
      'content-type': 'application/octet-stream',
      'content-disposition': `attachment; filename="${filename}"`,
      'content-length': file.byteLength,
      'access-control-allow-origin': corsOrigin(request),
    });
    response.end(file);
  } catch (error) {
    console.error(error);
    if (!response.headersSent) sendJson(request, response, 500, { error: error instanceof Error ? error.message : String(error) });
  } finally {
    if (directory) await rm(directory, { recursive: true, force: true });
  }
});

server.listen(port, () => {
  console.log(`Univer file export service listening on http://127.0.0.1:${port}`);
});
