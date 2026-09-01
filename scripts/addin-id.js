/* eslint-disable no-undef */

const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const PLACEHOLDER_ADDIN_ID = "fe737f47-102d-4d29-8a47-50844e10ac76";
const ID_FILE = path.join(__dirname, "..", ".addin-id");
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function assertUuid(value) {
  if (typeof value !== "string" || !UUID.test(value.trim())) {
    throw new Error(`Invalid add-in id: ${value}. Expected a uuid such as ${PLACEHOLDER_ADDIN_ID}.`);
  }
  return value.trim();
}

function applyAddinId(content, id) {
  return content.replace(new RegExp(PLACEHOLDER_ADDIN_ID, "g"), id);
}

/**
 * Each person who installs the add-in gets their own id, so several people can sideload it from
 * the same tenant without sharing one identity. The id is generated once and kept in .addin-id,
 * which is untracked, so repeat installs replace the previous one instead of piling up ghosts.
 */
function resolveAddinId() {
  if (process.env.ADDIN_ID) return assertUuid(process.env.ADDIN_ID);

  if (fs.existsSync(ID_FILE)) {
    return assertUuid(fs.readFileSync(ID_FILE, "utf8"));
  }

  const generated = crypto.randomUUID();
  fs.writeFileSync(ID_FILE, `${generated}\n`);
  return generated;
}

module.exports = { PLACEHOLDER_ADDIN_ID, ID_FILE, assertUuid, applyAddinId, resolveAddinId };
