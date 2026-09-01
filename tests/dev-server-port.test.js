const test = require("node:test");
const assert = require("node:assert/strict");
const {
  DEFAULT_DEV_SERVER_PORT,
  resolveDevServerPort,
} = require("../scripts/dev-server-port.js");

function withEnv(values, run) {
  const keys = ["DEV_SERVER_PORT", "npm_package_config_dev_server_port"];
  const saved = {};
  keys.forEach((key) => {
    saved[key] = process.env[key];
    delete process.env[key];
  });
  Object.entries(values).forEach(([key, value]) => {
    process.env[key] = value;
  });
  try {
    return run();
  } finally {
    keys.forEach((key) => {
      if (saved[key] === undefined) delete process.env[key];
      else process.env[key] = saved[key];
    });
  }
}

test("falls back to 3000 when nothing is configured", () => {
  withEnv({}, () => {
    assert.equal(resolveDevServerPort(), DEFAULT_DEV_SERVER_PORT);
    assert.equal(DEFAULT_DEV_SERVER_PORT, 3000);
  });
});

test("uses the npm package config port", () => {
  withEnv({ npm_package_config_dev_server_port: "3100" }, () => {
    assert.equal(resolveDevServerPort(), 3100);
  });
});

test("lets DEV_SERVER_PORT override the package config", () => {
  withEnv({ npm_package_config_dev_server_port: "3000", DEV_SERVER_PORT: "4200" }, () => {
    assert.equal(resolveDevServerPort(), 4200);
  });
});

test("rejects ports that are not usable", () => {
  ["0", "-1", "70000", "abc", "3000.5", ""].forEach((value) => {
    withEnv(value === "" ? {} : { DEV_SERVER_PORT: value }, () => {
      if (value === "") {
        assert.equal(resolveDevServerPort(), DEFAULT_DEV_SERVER_PORT);
        return;
      }
      assert.throws(() => resolveDevServerPort(), /dev server port/i, `expected ${value} to be rejected`);
    });
  });
});

test("rewrites every localhost:3000 reference in a manifest", () => {
  const { applyPort } = require("../scripts/dev-server-port.js");
  const source = JSON.stringify({
    a: "https://localhost:3000",
    b: "https://localhost:3000/taskpane.html",
    c: "https://localhost:3000/assets/icon-16.png",
  });

  const rewritten = applyPort(source, 4200);

  assert.ok(!rewritten.includes("localhost:3000"), "no 3000 references should survive");
  assert.equal((rewritten.match(/localhost:4200/g) || []).length, 3);
});
