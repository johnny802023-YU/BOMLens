import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { createServer, request as httpRequest } from "node:http";
import path from "node:path";
import { pathToFileURL } from "node:url";

const root = process.cwd();
const clientRoot = path.resolve(root, "dist", "client");
const serverModule = path.join(root, "node_modules", "vinext", "dist", "server", "prod-server.js");
const { startProdServer } = await import(pathToFileURL(serverModule).href);

const contentTypes = new Map([
  [".css", "text/css; charset=utf-8"],
  [".js", "application/javascript; charset=utf-8"],
  [".mjs", "application/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".svg", "image/svg+xml"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".woff2", "font/woff2"],
  [".wasm", "application/wasm"],
]);

async function resolveStaticFile(rawUrl) {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(rawUrl ?? "/", "http://127.0.0.1").pathname);
  } catch {
    return null;
  }
  if (pathname === "/") return null;
  const relative = pathname.replace(/^\/+/, "").split("/").join(path.sep);
  const candidate = path.resolve(clientRoot, relative);
  if (candidate !== clientRoot && !candidate.startsWith(`${clientRoot}${path.sep}`)) return null;
  try {
    const info = await stat(candidate);
    return info.isFile() ? { path: candidate, size: info.size, pathname } : null;
  } catch {
    return null;
  }
}

function proxyToBackend(req, res, backendPort) {
  const proxy = httpRequest({
    hostname: "127.0.0.1",
    port: backendPort,
    method: req.method,
    path: req.url,
    headers: req.headers,
  }, (backendResponse) => {
    res.writeHead(backendResponse.statusCode ?? 502, backendResponse.headers);
    backendResponse.pipe(res);
  });
  proxy.on("error", () => {
    if (!res.headersSent) res.writeHead(502, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("BOMLens local service unavailable");
  });
  req.pipe(proxy);
}

const backend = await startProdServer({
  port: 0,
  host: "127.0.0.1",
  outDir: path.join(root, "dist"),
  noCompression: true,
  purpose: "BOMLens backend",
});

const port = Number.parseInt(process.env.BOMLENS_PORT ?? "3784", 10);
const server = createServer(async (req, res) => {
  const file = await resolveStaticFile(req.url);
  if (!file) {
    proxyToBackend(req, res, backend.port);
    return;
  }
  const extension = path.extname(file.path).toLowerCase();
  res.writeHead(200, {
    "Content-Type": contentTypes.get(extension) ?? "application/octet-stream",
    "Content-Length": String(file.size),
    "Cache-Control": file.pathname.startsWith("/assets/") ? "public, max-age=31536000, immutable" : "no-cache",
    "X-Content-Type-Options": "nosniff",
  });
  if (req.method === "HEAD") {
    res.end();
    return;
  }
  createReadStream(file.path).pipe(res);
});

server.listen(port, "127.0.0.1", () => {
  console.log(`[BOMLens] Offline server running at http://127.0.0.1:${port}`);
});
