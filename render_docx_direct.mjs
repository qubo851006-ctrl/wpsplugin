import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const inputPath = String.raw`D:\claude\wps-ai-qa\codex修改建议.docx`;
const outputDir = String.raw`D:\claude\wps-ai-qa\codex_render_png`;
const artifactPkg = String.raw`C:\Users\shinh\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules\@oai\artifact-tool`;

const artifactTool = await import(
  pathToFileURL(path.join(artifactPkg, "dist", "artifact_tool.mjs")).href
);

await fs.rm(outputDir, { recursive: true, force: true });
await fs.mkdir(outputDir, { recursive: true });

const model = await artifactTool.DocumentFile.importDocx(
  await artifactTool.FileBlob.load(inputPath)
);

const bootstrapCanvas = new OffscreenCanvas(1, 1);
const bootstrapCtx = bootstrapCanvas.getContext("2d");
const initialDraw = artifactTool.drawDocumentToCtx(model, bootstrapCtx, { pageIndex: 0 });
const pages = initialDraw.pages;
const imageBitmaps = await artifactTool.preloadDocumentImageBitmaps(model, pages);

for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
  const page = pages[pageIndex];
  const width = Math.max(1, Math.ceil(page.widthPx * 2));
  const height = Math.max(1, Math.ceil(page.heightPx * 2));
  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.save();
  ctx.scale(2, 2);
  artifactTool.drawDocumentToCtx(model, ctx, {
    pageIndex,
    clear: true,
    imageBitmaps,
  });
  ctx.restore();
  const blob = await canvas.convertToBlob({ type: "image/png" });
  const buf = Buffer.from(await blob.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `page-${pageIndex + 1}.png`), buf);
}

console.log(outputDir);
