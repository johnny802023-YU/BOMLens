import path from "node:path";
import { pathToFileURL } from "node:url";

const root = process.cwd();
const serverModule = path.join(root, "node_modules", "vinext", "dist", "server", "prod-server.js");
const { startProdServer } = await import(pathToFileURL(serverModule).href);

await startProdServer({
  port: Number.parseInt(process.env.BOMLENS_PORT ?? "3784", 10),
  host: "127.0.0.1",
  outDir: path.join(root, "dist"),
  noCompression: true,
});
