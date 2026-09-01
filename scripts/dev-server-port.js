/* eslint-disable no-undef */

const DEFAULT_DEV_SERVER_PORT = 3000;

/**
 * The manifest ships with localhost:3000 as its canonical placeholder. Every URL in it is
 * rewritten to the resolved port before the add-in is sideloaded, so a second person can run
 * the dev server on a free port without editing tracked files.
 */
function resolveDevServerPort() {
  const raw = process.env.DEV_SERVER_PORT || process.env.npm_package_config_dev_server_port;
  if (!raw) return DEFAULT_DEV_SERVER_PORT;

  const port = Number(raw);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`Invalid dev server port: ${raw}. Use an integer between 1 and 65535.`);
  }
  return port;
}

function applyPort(content, port) {
  return content.replace(new RegExp(`localhost:${DEFAULT_DEV_SERVER_PORT}`, "g"), `localhost:${port}`);
}

module.exports = { DEFAULT_DEV_SERVER_PORT, resolveDevServerPort, applyPort };
