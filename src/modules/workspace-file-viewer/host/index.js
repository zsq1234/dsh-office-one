import { chmod, link, lstat, readdir, readFile, realpath, rename, stat, unlink, writeFile } from "node:fs/promises";
import { basename, dirname, extname, isAbsolute, relative, resolve, sep } from "node:path";
import { randomUUID } from "node:crypto";

const name = "workspace-file-viewer";
const inject = ["webServer", "workspaceRegistry"];
const IGNORED = new Set([".git", "node_modules", ".DS_Store"]);
const DEFAULT_PAGE_SIZE = 100;
const MAX_PAGE_SIZE = 500;
const MAX_FILE_BYTES = 300 * 1024 * 1024;
const UNIVER_IMPORT_ENDPOINT = process.env.UNIVER_FILE_IMPORT_ENDPOINT ?? "http://127.0.0.1:8787/api/univer/import-file";
const UNIVER_EXPORT_ENDPOINT = process.env.UNIVER_FILE_EXPORT_ENDPOINT ?? "http://127.0.0.1:8787/api/univer/export-file";
import { readFile as readRuntimeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
const RUNTIME_DIR = fileURLToPath(new URL("./", import.meta.url));

function isInside(root, candidate) {
  const rel = relative(root, candidate);
  return rel === "" || (!rel.startsWith(`..${sep}`) && rel !== ".." && !isAbsolute(rel));
}

async function safePath(root, input = "") {
  const candidate = resolve(root, input);
  if (!isInside(resolve(root), candidate)) throw new Error("path must stay inside the workspace");
  const [canonicalRoot, canonicalCandidate] = await Promise.all([realpath(root), realpath(candidate)]);
  if (isInside(canonicalRoot, canonicalCandidate)) return canonicalCandidate;
  throw new Error("path must stay inside the workspace (symlinks included)");
}

function directoryEntryRank(entry) { return entry.isDirectory() ? 0 : 1; }
function compareDirectoryEntry(entry, cursorKey) {
  return directoryEntryRank(entry) - cursorKey[0] || entry.name.localeCompare(cursorKey[1]);
}
function decodeDirectoryCursor(cursor) {
  if (cursor === "") return null;
  try {
    const value = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
    if (!Array.isArray(value) || (value[0] !== 0 && value[0] !== 1) || typeof value[1] !== "string") throw new Error();
    return value;
  } catch {
    throw new Error("invalid directory cursor");
  }
}
function encodeDirectoryCursor(entry) {
  return Buffer.from(JSON.stringify([directoryEntryRank(entry), entry.name])).toString("base64url");
}

async function listDirectory(root, inputPath = "", cursor = "", requestedLimit = "") {
  const directory = await safePath(root, inputPath);
  const directoryInfo = await stat(directory);
  if (!directoryInfo.isDirectory()) throw new Error("not a directory");

  const cursorKey = decodeDirectoryCursor(cursor);
  const parsedLimit = requestedLimit === "" ? DEFAULT_PAGE_SIZE : Number(requestedLimit);
  if (!Number.isSafeInteger(parsedLimit) || parsedLimit < 1) throw new Error("invalid page size");
  const limit = Math.min(parsedLimit, MAX_PAGE_SIZE);

  const entries = (await readdir(directory, { withFileTypes: true }))
    .filter((entry) => !IGNORED.has(entry.name) && (entry.isDirectory() || entry.isFile()));
  entries.sort((a, b) => directoryEntryRank(a) - directoryEntryRank(b) || a.name.localeCompare(b.name));

  const startIndex = cursorKey === null ? 0 : (() => {
    const index = entries.findIndex((entry) => compareDirectoryEntry(entry, cursorKey) > 0);
    return index < 0 ? entries.length : index;
  })();
  const directoryPath = relative(root, directory).split(sep).join("/");
  const page = entries.slice(startIndex, startIndex + limit);
  const result = (await Promise.all(page.map(async (entry) => {
    const path = directoryPath ? `${directoryPath}/${entry.name}` : entry.name;
    if (entry.isDirectory()) return { path, name: entry.name, type: "directory" };
    const info = await lstat(resolve(directory, entry.name));
    if (!info.isFile()) return null;
    return { path, name: entry.name, type: "file", size: info.size, modified: info.mtime.toISOString() };
  }))).filter(Boolean);
  return {
    path: directoryPath,
    entries: result,
    nextCursor: startIndex + page.length < entries.length && page.length > 0 ? encodeDirectoryCursor(page[page.length - 1]) : null,
    total: entries.length,
  };
}

async function readOfficeFile(file, unitType) {
  const input = await readFile(file);
  const format = extname(file).slice(1).toLowerCase();
  const params = new URLSearchParams({ unitType, format, fileName: basename(file) });
  let response;
  try {
    response = await fetch(`${UNIVER_IMPORT_ENDPOINT}?${params}`, {
      method: "POST",
      headers: { "content-type": "application/octet-stream" },
      body: input,
      signal: AbortSignal.timeout(300_000),
    });
  } catch (error) {
    throw new Error(`cannot reach dsh-univer-file-export: ${error instanceof Error ? error.message : String(error)}`);
  }
  let payload;
  try { payload = await response.json(); } catch { throw new Error(`dsh-univer-file-export returned HTTP ${response.status} without JSON`); }
  if (!response.ok) throw new Error(payload?.error || `dsh-univer-file-export returned HTTP ${response.status}`);
  if (payload?.data === null || typeof payload?.data !== "object") throw new Error("dsh-univer-file-export did not return valid Univer UnitData");
  return payload.data;
}

const readXlsx = (file) => readOfficeFile(file, "sheet");
const readDoc = (file) => readOfficeFile(file, "doc");
const readSlides = (file) => readOfficeFile(file, "slide");

async function readJsonBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_FILE_BYTES) throw new Error("save request is larger than 300 MiB");
    chunks.push(chunk);
  }
  if (size === 0) throw new Error("save request body is required");
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

const SAVE_FORMATS = {
  sheet: new Set(["xlsx", "csv"]),
  doc: new Set(["docx"]),
  slide: new Set(["pptx"]),
};

async function exportOfficeFile(data, unitType, format) {
  let response;
  try {
    response = await fetch(UNIVER_EXPORT_ENDPOINT, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ unitType, format, data }),
      signal: AbortSignal.timeout(300_000),
    });
  } catch (error) {
    throw new Error(`cannot reach dsh-univer-file-export: ${error instanceof Error ? error.message : String(error)}`);
  }
  if (!response.ok) {
    let message = `dsh-univer-file-export returned HTTP ${response.status}`;
    try { const payload = await response.json(); if (payload?.error) message = payload.error; } catch {}
    throw new Error(message);
  }
  const output = Buffer.from(await response.arrayBuffer());
  if (output.byteLength === 0) throw new Error("dsh-univer-file-export returned an empty file");
  if (output.byteLength > MAX_FILE_BYTES) throw new Error("exported file is larger than 300 MiB");
  if (["xlsx", "docx", "pptx"].includes(format) && (output[0] !== 0x50 || output[1] !== 0x4b)) {
    throw new Error("dsh-univer-file-export returned an invalid OOXML file");
  }
  return output;
}

async function saveOfficeFile(root, payload) {
  const path = typeof payload?.path === "string" ? payload.path : "";
  const unitType = typeof payload?.unitType === "string" ? payload.unitType : "";
  const format = typeof payload?.format === "string" ? payload.format.toLowerCase() : "";
  if (!path || !SAVE_FORMATS[unitType]?.has(format) || payload?.data === null || typeof payload?.data !== "object") {
    throw new Error("expected path, supported unitType/format, and Univer UnitData");
  }
  const sourceFormat = extname(path).slice(1).toLowerCase();
  const saveLegacyXlsAsXlsx = unitType === "sheet" && sourceFormat === "xls" && format === "xlsx";
  if (sourceFormat !== format && !saveLegacyXlsAsXlsx) throw new Error("file extension does not match the export format");
  const absolute = await safePath(root, path);
  const info = await stat(absolute);
  if (!info.isFile()) throw new Error("not a file");
  const expectedModified = typeof payload?.expectedModified === "string" ? payload.expectedModified : "";
  if (!expectedModified || info.mtime.toISOString() !== expectedModified) {
    throw new Error("file changed after it was opened; reload it before saving");
  }

  const target = saveLegacyXlsAsXlsx
    ? resolve(dirname(absolute), `${basename(absolute).replace(/\.xls$/i, "")}.xlsx`)
    : absolute;
  if (saveLegacyXlsAsXlsx) {
    try {
      await lstat(target);
      throw new Error(`${basename(target)} already exists; rename or remove it before saving the XLS file as XLSX`);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
    }
  }

  const output = await exportOfficeFile(payload.data, unitType, format);
  const temporary = resolve(dirname(target), `.${basename(target)}.${randomUUID()}.tmp`);
  try {
    await writeFile(temporary, output, { flag: "wx" });
    await chmod(temporary, info.mode);
    const current = await stat(absolute);
    if (current.mtime.toISOString() !== expectedModified) throw new Error("file changed while it was being exported; reload it before saving");
    if (saveLegacyXlsAsXlsx) {
      await link(temporary, target);
      await unlink(temporary);
    } else {
      await rename(temporary, target);
    }
  } catch (error) {
    await unlink(temporary).catch(() => {});
    throw error;
  }
  const saved = await stat(target);
  return { path: relative(root, target).split(sep).join("/"), size: output.byteLength, modified: saved.mtime.toISOString() };
}

function json(res, status, payload) { res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }); res.end(JSON.stringify(payload)); }
function sessionRoot(registry, sessionId) { if (!sessionId) throw new Error("sessionId is required"); const root = registry.host.sessionPath(sessionId); if (!root) throw new Error("cannot resolve the workspace directory for this session"); return root; }
async function handleFiles(req, res, registry) {
  try {
    const url = new URL(req.url ?? "/", "http://localhost");
    const root = sessionRoot(registry, url.searchParams.get("sessionId"));
    if (url.pathname === "/api/workspace-office-save") {
      if (req.method !== "POST") return json(res, 405, { error: "method not allowed" });
      const result = await saveOfficeFile(root, await readJsonBody(req));
      return json(res, 200, result);
    }
    if (req.method !== "GET") return json(res, 405, { error: "method not allowed" });
    if (url.pathname === "/api/workspace-files") {
      const page = await listDirectory(
        root,
        url.searchParams.get("path") ?? "",
        url.searchParams.get("cursor") ?? "",
        url.searchParams.get("limit") ?? "",
      );
      return json(res, 200, page);
    }
    if (url.pathname === "/api/workspace-file") {
      const absolute = await safePath(root, url.searchParams.get("path") ?? ""); const info = await stat(absolute);
      if (!info.isFile()) return json(res, 400, { error: "not a file" });
      if (info.size > MAX_FILE_BYTES) return json(res, 413, { error: "file is larger than 300 MiB" });
      const path = relative(root, absolute).split(sep).join("/");
      if (/\.(ppt|pptx)$/i.test(path)) return json(res, 200, { path, presentation: await readSlides(absolute) });
      return json(res, 200, { path, content: await readFile(absolute, "utf8"), size: info.size, modified: info.mtime.toISOString() });
    }
    if (url.pathname === "/api/workspace-xlsx") {
      const path = url.searchParams.get("path") ?? ""; if (!/\.(xls|xlsx|csv)$/i.test(path)) throw new Error("only .xls, .xlsx, and .csv files are supported");
      const absolute = await safePath(root, path); const info = await stat(absolute); if (!info.isFile()) throw new Error("not a file");
      if (info.size > MAX_FILE_BYTES) return json(res, 413, { error: "workbook is larger than 300 MiB" });
      return json(res, 200, { path, modified: info.mtime.toISOString(), workbook: await readXlsx(absolute) });
    }
    if (url.pathname === "/api/workspace-doc") {
      const path = url.searchParams.get("path") ?? ""; if (!/\.(doc|docx)$/i.test(path)) throw new Error("only .doc and .docx files are supported");
      const absolute = await safePath(root, path); const info = await stat(absolute); if (!info.isFile()) throw new Error("not a file");
      if (info.size > MAX_FILE_BYTES) return json(res, 413, { error: "document is larger than 300 MiB" });
      return json(res, 200, { path, modified: info.mtime.toISOString(), document: await readDoc(absolute) });
    }
    if (url.pathname === "/api/workspace-slides") {
      const path = url.searchParams.get("path") ?? ""; if (!/\.(ppt|pptx)$/i.test(path)) throw new Error("only .ppt and .pptx files are supported");
      const absolute = await safePath(root, path); const info = await stat(absolute); if (!info.isFile()) throw new Error("not a file");
      if (info.size > MAX_FILE_BYTES) return json(res, 413, { error: "presentation is larger than 300 MiB" });
      return json(res, 200, { path, modified: info.mtime.toISOString(), presentation: await readSlides(absolute) });
    }
    return json(res, 404, { error: "not found" });
  } catch (error) { return json(res, 400, { error: error instanceof Error ? error.message : String(error) }); }
}
async function handleRuntime(req, res) { const kind = new URL(req.url ?? '/', 'http://localhost').pathname.split('/').pop(); if (kind !== 'sheet.js' && kind !== 'docs.js' && kind !== 'slides.js') return json(res, 404, { error: 'not found' }); try { const body = await readRuntimeFile(resolve(RUNTIME_DIR, 'runtimes', kind)); res.writeHead(200, { 'content-type': 'application/javascript; charset=utf-8', 'cache-control': 'no-store' }); res.end(body); } catch { json(res, 404, { error: 'runtime not built' }); } }
function apply(ctx) { const handler = (req, res) => handleFiles(req, res, ctx.workspaceRegistry); const runtimeHandler = (req, res) => handleRuntime(req, res); ctx.effect(() => { const a = ctx.webServer.register({ kind: "exact", path: "/api/workspace-files", handler }); const b = ctx.webServer.register({ kind: "exact", path: "/api/workspace-file", handler }); const c = ctx.webServer.register({ kind: "exact", path: "/api/workspace-xlsx", handler }); const d = ctx.webServer.register({ kind: "exact", path: "/api/workspace-doc", handler }); const e = ctx.webServer.register({ kind: "exact", path: "/api/workspace-slides", handler }); const s = ctx.webServer.register({ kind: "exact", path: "/api/workspace-office-save", handler }); const rs = ctx.webServer.register({ kind: "exact", path: "/api/workspace-file-viewer/runtime/docs.js", handler: runtimeHandler }); const rt = ctx.webServer.register({ kind: "exact", path: "/api/workspace-file-viewer/runtime/sheet.js", handler: runtimeHandler }); const rv = ctx.webServer.register({ kind: "exact", path: "/api/workspace-file-viewer/runtime/slides.js", handler: runtimeHandler }); return () => { a(); b(); c(); d(); e(); s(); rs(); rt(); rv(); }; }, "workspace-file-viewer: api"); }
export { apply, inject, name, readDoc, readXlsx };
