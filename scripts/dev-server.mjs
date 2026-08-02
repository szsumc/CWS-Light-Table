import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const port = Number(process.env.PORT || 4173);

const mimeTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
]);

createServer(async (request, response) => {
  try {
    const requestPath = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`).pathname;
    const relativePath = requestPath === "/" ? "index.html" : requestPath.slice(1);
    const absolutePath = path.resolve(rootDir, relativePath);
    if (!absolutePath.startsWith(rootDir)) {
      response.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Forbidden");
      return;
    }
    const body = await readFile(absolutePath);
    response.writeHead(200, {
      "Content-Type": mimeTypes.get(path.extname(absolutePath).toLowerCase()) || "application/octet-stream",
    });
    response.end(body);
  } catch (error) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(`Not found: ${error.message}`);
  }
}).listen(port, () => {
  console.log(`CWS Light Table dev server running at http://127.0.0.1:${port}`);
});
