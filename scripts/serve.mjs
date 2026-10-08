import { createReadStream, existsSync } from "node:fs";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.join(projectRoot, "out");
const types = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
  ".webmanifest": "application/manifest+json",
  ".woff2": "font/woff2",
};

if (!existsSync(outputRoot)) {
  console.error("Static export not found. Run npm run build first.");
  process.exit(1);
}

createServer((request, response) => {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(request.url ?? "/", "http://localhost").pathname);
  } catch {
    response.writeHead(400).end("Bad request");
    return;
  }
  const relativePath = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  const filePath = path.resolve(outputRoot, relativePath);
  if (!filePath.startsWith(`${outputRoot}${path.sep}`) && filePath !== outputRoot) {
    response.writeHead(403).end("Forbidden");
    return;
  }
  if (!existsSync(filePath)) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" }).end("Not found");
    return;
  }
  response.writeHead(200, { "Content-Type": types[path.extname(filePath)] ?? "application/octet-stream", "Cache-Control": "no-cache" });
  if (request.method === "HEAD") response.end();
  else createReadStream(filePath).pipe(response);
}).listen(4173, "0.0.0.0", () => console.log("BatchScan preview: http://localhost:4173"));
