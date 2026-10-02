import assert from "node:assert/strict";
import { readFileSync, statSync } from "node:fs";
import test from "node:test";
import { JSDOM } from "jsdom";

// Exercise the actual compiled visual with DOM events and a fake Power BI host.
const bundlePath = new URL("../.tmp/drop/visual.js", import.meta.url);
for (const source of ["../src/visual.ts", "../src/settings.ts", "../pbiviz.json", "../capabilities.json"]) {
  assert.ok(statSync(bundlePath).mtimeMs >= statSync(new URL(source, import.meta.url)).mtimeMs,
    "Rebuild the visual with npm run package before running npm test.");
}
const bundle = readFileSync(bundlePath, "utf8");
const config = JSON.parse(readFileSync(new URL("../pbiviz.json", import.meta.url)));
const merge = 0;
const remove = 1;

function setup(t, withField = true) {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const dom = new JSDOM('<div id="visual"></div>', { runScripts: "outside-only" });
  const { window } = dom;
  window.setTimeout = globalThis.setTimeout;
  window.clearTimeout = globalThis.clearTimeout;
  window.powerbi = {};
  window.eval(bundle);
  const calls = [];
  const host = {
    eventService: { renderingStarted() {}, renderingFinished() {} },
    createLocalizationManager: () => ({ getDisplayName: key => key }),
    createSelectionManager: () => ({ showContextMenu() {} }),
    applyJsonFilter: (...args) => calls.push(JSON.parse(JSON.stringify(args)))
  };
  const visual = window.powerbi.visuals.plugins[config.visual.guid].create({
    element: window.document.querySelector("#visual"), host
  });
  const input = window.document.querySelector("input");
  const clear = window.document.querySelector('[name="clear-button"]');
  const update = (jsonFilters = [], queryName = "Names.Name", extra = {}) => visual.update({
    type: 2, viewport: { width: 400, height: 60 },
    dataViews: [{ metadata: { columns: queryName ? [{ queryName, displayName: "Name", type: { text: true } }] : [] } }],
    jsonFilters, ...extra
  });
  const enter = (text, inputType = "insertText") => {
    input.value = text;
    input.dispatchEvent(new window.InputEvent("input", { bubbles: true, inputType }));
  };
  update([], withField ? "Names.Name" : null);
  t.after(() => { visual.destroy(); window.close(); });
  return { visual, input, clear, calls, update, enter, window, tick: ms => t.mock.timers.tick(ms) };
}

function contains(value, table = "Names", column = "Name") {
  return {
    $schema: "http://powerbi.com/product/schema#advanced", filterType: 0,
    target: { table, column }, logicalOperator: "And",
    conditions: [{ operator: "Contains", value }]
  };
}

test("rapid typing submits only the final Contains search after 300 ms", t => {
  const s = setup(t);
  s.enter("w"); s.tick(200);
  s.enter("wi"); s.tick(200);
  s.enter("will"); s.tick(299);
  assert.equal(s.calls.length, 0);
  s.tick(1);
  assert.deepEqual(s.calls, [[contains("will"), "general", "filter", merge]]);
  assert.equal(s.window.document.querySelector('[name="search-button"]'), null);
});

test("pasting filters automatically, including literal punctuation and accents", t => {
  const s = setup(t);
  s.enter("Jo\u00e3o O'Neil [A]", "insertFromPaste"); s.tick(300);
  assert.deepEqual(s.calls[0], [contains("Jo\u00e3o O'Neil [A]"), "general", "filter", merge]);
});

test("deletion and whitespace remove the filter immediately without a stale search", t => {
  const s = setup(t);
  s.enter("pending"); s.tick(100);
  s.enter("", "deleteContentBackward");
  assert.deepEqual(s.calls, [[null, "general", "filter", remove]]);
  s.tick(500);
  assert.equal(s.calls.length, 1);
  s.enter("   ");
  assert.equal(s.calls[1][3], remove);
  assert.equal(s.input.value, "");
});

test("clear cancels the timer and returns focus to the input", t => {
  const s = setup(t);
  s.enter("pending"); s.clear.click(); s.tick(500);
  assert.deepEqual(s.calls, [[null, "general", "filter", remove]]);
  assert.equal(s.input.value, "");
  assert.equal(s.window.document.activeElement, s.input);
});

test("Enter applies immediately and cancels the automatic duplicate", t => {
  const s = setup(t);
  s.enter("immediate");
  s.input.dispatchEvent(new s.window.KeyboardEvent("keydown", { key: "Enter" }));
  s.tick(500);
  assert.deepEqual(s.calls, [[contains("immediate"), "general", "filter", merge]]);
});

test("text composition does not submit an unfinished word", t => {
  const s = setup(t);
  s.enter("old");
  s.input.dispatchEvent(new s.window.CompositionEvent("compositionstart"));
  s.enter("\u306b"); s.tick(500);
  s.input.dispatchEvent(new s.window.KeyboardEvent("keydown", { key: "Enter", isComposing: true }));
  assert.equal(s.calls.length, 0);
  s.input.value = "\u65e5\u672c";
  s.input.dispatchEvent(new s.window.CompositionEvent("compositionend"));
  s.tick(300);
  assert.deepEqual(s.calls[0], [contains("\u65e5\u672c"), "general", "filter", merge]);
});

test("host updates preserve the draft and ignore earlier query replies", t => {
  const s = setup(t);
  s.update([contains("saved")]);
  s.enter("latest");
  s.update([contains("saved")]);
  assert.equal(s.input.value, "latest");
  s.tick(300);
  s.update([contains("saved")]);
  assert.equal(s.input.value, "latest");
  s.update([contains("latest")]);
  s.update([contains("bookmark")]);
  assert.equal(s.input.value, "bookmark");
  s.update([]);
  assert.equal(s.input.value, "");
  assert.equal(s.calls.length, 1);
});

test("saved searches restore safely alongside unrelated Basic filters", t => {
  const s = setup(t);
  s.update([{ target: { table: "Other", column: "Category" }, operator: "In", values: ["A"] }, contains("will")]);
  assert.equal(s.input.value, "will");
  s.update([contains("unrelated", "Other")]);
  assert.equal(s.input.value, "");
});

test("changing or removing the field cancels the pending search and clears the old filter", t => {
  const s = setup(t);
  s.enter("pending");
  s.update([], "Names.Surname"); s.tick(500);
  assert.deepEqual(s.calls, [[null, "general", "filter", remove]]);
  assert.equal(s.input.value, "");
  s.enter("new"); s.tick(300);
  assert.deepEqual(s.calls[1][0].target, { table: "Names", column: "Surname" });
  s.update([], null);
  assert.equal(s.input.disabled, true);
  assert.equal(s.clear.disabled, true);
});

test("a resize without data views preserves the selected field and draft", t => {
  const s = setup(t);
  s.enter("resize");
  s.update([], "Names.Name", { dataViews: undefined, type: 4 });
  s.tick(300);
  assert.deepEqual(s.calls, [[contains("resize"), "general", "filter", merge]]);
});

test("without a field the input is disabled and no filter is submitted", t => {
  const s = setup(t, false);
  s.enter("ignored"); s.tick(500);
  assert.equal(s.input.disabled, true);
  s.visual.performSearch("ignored");
  assert.equal(s.calls.length, 0);
});

test("destroy cancels a search before it can reach the Power BI host", t => {
  const s = setup(t);
  s.enter("pending"); s.visual.destroy(); s.tick(500);
  assert.equal(s.calls.length, 0);
});
