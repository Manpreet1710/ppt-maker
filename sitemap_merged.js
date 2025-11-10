/**
 * merge-simple-sitemaps.js
 *
 * Description:
 * Takes an array of sitemap URLs and generates a sitemap index file (merged-sitemap.xml)
 * without fetching or reading the actual sitemaps.
 *
 * Usage:
 *   node merge-simple-sitemaps.js
 */

const fs = require("fs");
const path = require("path");

// 👉 Step 1: Add your sitemap URLs here
const sitemapUrls = [
  "https://wordswithletters.org/words-with-letters/sitemap.xml",
  "https://wordswithletters.org/x-letter-words-with-letters/sitemap.xml",
  "https://wordswithletters.org/words-collections/sitemap.xml",
  "https://wordswithletters.org/words-with-letters/sitemap.xml",
  "https://wordswithletters.org/words-with-letters/sitemap.xml"
];

// 👉 Step 2: Generate sitemap index XML
function generateSitemapIndex(urls) {
  const header =
    '<?xml version="1.0" encoding="UTF-8"?>\n<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n';
  const footer = "</sitemapindex>";

  const body = urls
    .map(
      (url) => `  <sitemap>\n    <loc>${escapeXml(url)}</loc>\n  </sitemap>`
    )
    .join("\n");

  return header + body + "\n" + footer + "\n";
}

// Simple XML escape
function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// 👉 Step 3: Write to file
function writeMergedSitemap(outputFile, content) {
  fs.writeFileSync(outputFile, content, "utf8");
  console.log(`✅ Merged sitemap created: ${outputFile}`);
}

// Run
(function main() {
  if (!sitemapUrls.length) {
    console.error("❌ No sitemap URLs provided.");
    process.exit(1);
  }

  const xmlContent = generateSitemapIndex(sitemapUrls);
  const outputPath = path.resolve(process.cwd(), "merged-sitemap.xml");
  writeMergedSitemap(outputPath, xmlContent);
})();
