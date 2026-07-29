// build.js - bundles taskpane.css + taskpane.js into a single self-contained
// docs/taskpane.html for zero-dependency static hosting (e.g. GitHub Pages).
const fs = require("fs");
const path = require("path");

const root = __dirname;
const outDir = path.join(root, "docs");
fs.mkdirSync(outDir, { recursive: true });

let html = fs.readFileSync(path.join(root, "taskpane.html"), "utf8");
const css = fs.readFileSync(path.join(root, "taskpane.css"), "utf8");
const js = fs.readFileSync(path.join(root, "taskpane.js"), "utf8");

// Inline the stylesheet.
html = html.replace(
  /<link rel="stylesheet" href="taskpane\.css"\s*\/>/,
  `<style>\n${css}\n</style>`
);

// Inline the script.
html = html.replace(
  /<script src="taskpane\.js"><\/script>/,
  `<script>\n${js}\n</script>`
);

fs.writeFileSync(path.join(outDir, "taskpane.html"), html, "utf8");

// Copy icons so the manifest icon URLs resolve on the host too.
const assetsIn = path.join(root, "assets");
const assetsOut = path.join(outDir, "assets");
if (fs.existsSync(assetsIn)) {
  fs.mkdirSync(assetsOut, { recursive: true });
  for (const f of fs.readdirSync(assetsIn)) {
    fs.copyFileSync(path.join(assetsIn, f), path.join(assetsOut, f));
  }
}

console.log("Built docs/taskpane.html (self-contained).");
