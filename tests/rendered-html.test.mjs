import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the BOM comparison workspace", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<html lang="zh-Hant">/i);
  assert.match(html, /<title>BOMLens｜BOM 與線路圖版本比對<\/title>/i);
  assert.match(html, /版本比對/);
  assert.match(html, /BOM 差異/);
  assert.match(html, /線路圖比對/);
  assert.match(html, /PCB_Main_v1\.3\.xlsx/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("ships real BOM parsing, comparison, and export behavior", async () => {
  const [page, layout, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(page, /XLSX\.read/);
  assert.match(page, /function compareBom/);
  assert.match(page, /XLSX\.writeFile/);
  assert.match(page, /getImageData/);
  assert.match(page, /application\/pdf/);
  assert.match(layout, /lang="zh-Hant"/);
  assert.match(packageJson, /"xlsx"/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
});
