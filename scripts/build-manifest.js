/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const { applyPort, resolveDevServerPort } = require("./dev-server-port");

const root = path.join(__dirname, "..");
const source = path.join(root, "manifest.json");
const target = path.join(root, "manifest.local.json");

const port = resolveDevServerPort();
fs.writeFileSync(target, applyPort(fs.readFileSync(source, "utf8"), port));

console.log(`Wrote manifest.local.json pointing at https://localhost:${port}`);
