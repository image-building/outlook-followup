#!/usr/bin/env node
/* Usage: node scripts/configure.js https://yourname.github.io/outlook-followup <entra-client-id> <tenant-id>
 * Writes your hosting URL and Entra app ID into manifest.xml and src/config.js. */
const fs = require("fs");
const path = require("path");

const [baseUrlArg, clientId, tenantId] = process.argv.slice(2);
if (!baseUrlArg || !clientId || !tenantId) {
  console.error("Usage: node scripts/configure.js <https base url> <entra client id> <tenant id>");
  process.exit(1);
}
const baseUrl = baseUrlArg.replace(/\/+$/, "");
if (!/^https:\/\//.test(baseUrl)) { console.error("The base URL must start with https://"); process.exit(1); }
if (!/^[0-9a-f-]{36}$/i.test(clientId)) { console.error("The client ID should look like 00000000-0000-0000-0000-000000000000"); process.exit(1); }

const root = path.join(__dirname, "..");
const manifestPath = path.join(root, "manifest.xml");
let manifest = fs.readFileSync(manifestPath, "utf8");
manifest = manifest.replace(/https:\/\/[^"]*?\/outlook-followup(?=\/)/g, baseUrl)
  .replace(/https:\/\/YOUR-HOST\/outlook-followup/g, baseUrl);
fs.writeFileSync(manifestPath, manifest);

const configPath = path.join(root, "src", "config.js");
let config = fs.readFileSync(configPath, "utf8");
config = config.replace(/clientId: "[^"]*"/, `clientId: "${clientId}"`)
  .replace(/baseUrl: "[^"]*"/, `baseUrl: "${baseUrl}"`)
  .replace(/authority: "[^"]*"/, `authority: "https://login.microsoftonline.com/${tenantId}"`);
fs.writeFileSync(configPath, config);

const host = new URL(baseUrl).host;
console.log("Configured.");
console.log("Add these redirect URIs to your Entra app (platform: Single-page application):");
console.log("  brk-multihub://" + host);
console.log("  " + baseUrl + "/src/auth-dialog.html");
