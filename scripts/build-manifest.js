/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const { applyAddinId, resolveAddinId } = require("./addin-id");
const { applyPort, resolveDevServerPort } = require("./dev-server-port");

const root = path.join(__dirname, "..");
const source = path.join(root, "manifest.json");
const target = path.join(root, "manifest.local.json");

const port = resolveDevServerPort();
const addinId = resolveAddinId();

const manifest = applyAddinId(applyPort(fs.readFileSync(source, "utf8"), port), addinId);
fs.writeFileSync(target, manifest);

console.log(`Wrote manifest.local.json for https://localhost:${port}, add-in id ${addinId}`);
