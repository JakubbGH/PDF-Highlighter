const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const zlib = require("node:zlib");

const repoRoot = path.resolve(__dirname, "..");
const appPath = path.join(repoRoot, "app.js");

const TEST_EXPORTS = [
  "contentTypesXml",
  "workbookRelsXml",
  "planSheetXml",
  "progressSheetXml",
  "drawingXml",
  "createZip",
  "extractVbaProjectFromWorkbook",
  "findZipEntry",
  "inflateRawDeflate",
  "progressColor",
  "parseCsv",
  "csvRowsToUpdates",
  "parseCsvPoints",
  "csvCell",
  "extractZlCode",
  "zlLabelsFromTextRuns",
  "zlLabelsFromDetectedText",
  "mergeDetectedZlLabels",
  "findBoxBoundaryForLabel",
  "roomLabelBox",
  "planToPages",
  "projectNameFromFiles",
  "pageTabLabel",
  "excelPlanSheetNames",
  "normalizeRotation",
  "rotatedPlanSize",
  "rotatePoint",
  "rotatePoints",
  "rotateDetectedLabel",
  "copyVbaSourceCode",
  "fetchVbaSourceCode",
  "copyTextToClipboard",
  "renderMacroStatus",
  "saveInstalledVbaProject",
  "downloadInstalledMacroProject",
  "loadVbaProject"
];

function loadAppHarness() {
  let source = fs.readFileSync(appPath, "utf8");
  source = source.replace("\n  init();\n", "\n  // init skipped by tools/smoke_test.js\n");
  source = source.replace(
    /\}\)\(\);\s*$/,
    `  globalThis.__testApi = { ${TEST_EXPORTS.join(", ")} };\n})();`
  );

  const elements = new Map();
  const storage = new Map();

  function fakeElement(id) {
    if (elements.has(id)) return elements.get(id);
    const element = {
      id,
      addEventListener() {},
      append() {},
      appendChild() {},
      click() {},
      remove() {},
      replaceChildren() {},
      setAttribute() {},
      removeAttribute() {},
      querySelector() {
        return null;
      },
      getBoundingClientRect() {
        return { left: 0, top: 0, width: 1200, height: 800 };
      },
      classList: {
        add() {},
        remove() {},
        toggle() {}
      },
      style: {},
      dataset: {},
      files: [],
      value: "",
      checked: false,
      disabled: false,
      hidden: false,
      textContent: "",
      innerHTML: "",
      title: ""
    };
    elements.set(id, element);
    return element;
  }

  const context = {
    console,
    TextDecoder,
    TextEncoder,
    Uint8Array,
    Array,
    Math,
    JSON,
    Date,
    Error,
    Blob,
    Response,
    setTimeout,
    clearTimeout,
    document: {
      documentElement: { dataset: {} },
      body: fakeElement("body"),
      getElementById: fakeElement,
      createElement(tag) {
        return fakeElement(`created-${tag}-${elements.size}`);
      },
      createElementNS(_namespace, tag) {
        return fakeElement(`created-${tag}-${elements.size}`);
      }
    },
    localStorage: {
      getItem(key) {
        return storage.has(key) ? storage.get(key) : null;
      },
      setItem(key, value) {
        storage.set(key, value);
      },
      removeItem(key) {
        storage.delete(key);
      }
    },
    window: {
      crypto: null,
      confirm() {
        return true;
      }
    },
    navigator: {},
    pdfjsLib: null,
    URL: {
      createObjectURL() {
        return "blob:test";
      },
      revokeObjectURL() {}
    },
    alert() {},
    confirm() {
      return true;
    },
    fetch() {
      throw new Error("Unexpected fetch in smoke test");
    },
    btoa(value) {
      return Buffer.from(value, "binary").toString("base64");
    },
    atob(value) {
      return Buffer.from(value, "base64").toString("binary");
    }
  };
  context.globalThis = context;

  vm.runInNewContext(source, context, { filename: appPath });
  return { api: context.__testApi, elements, storage, context };
}

function makeDeflatedZip(name, data) {
  const nameBytes = Buffer.from(name, "utf8");
  const compressed = zlib.deflateRawSync(Buffer.from(data));
  const localHeader = Buffer.alloc(30);
  let offset = 0;

  localHeader.writeUInt32LE(0x04034b50, 0);
  localHeader.writeUInt16LE(20, 4);
  localHeader.writeUInt16LE(0, 6);
  localHeader.writeUInt16LE(8, 8);
  localHeader.writeUInt16LE(0, 10);
  localHeader.writeUInt16LE(0, 12);
  localHeader.writeUInt32LE(0, 14);
  localHeader.writeUInt32LE(compressed.length, 18);
  localHeader.writeUInt32LE(data.length, 22);
  localHeader.writeUInt16LE(nameBytes.length, 26);
  localHeader.writeUInt16LE(0, 28);
  offset += localHeader.length + nameBytes.length + compressed.length;

  const centralDirectory = Buffer.alloc(46);
  centralDirectory.writeUInt32LE(0x02014b50, 0);
  centralDirectory.writeUInt16LE(20, 4);
  centralDirectory.writeUInt16LE(20, 6);
  centralDirectory.writeUInt16LE(0, 8);
  centralDirectory.writeUInt16LE(8, 10);
  centralDirectory.writeUInt16LE(0, 12);
  centralDirectory.writeUInt16LE(0, 14);
  centralDirectory.writeUInt32LE(0, 16);
  centralDirectory.writeUInt32LE(compressed.length, 20);
  centralDirectory.writeUInt32LE(data.length, 24);
  centralDirectory.writeUInt16LE(nameBytes.length, 28);
  centralDirectory.writeUInt16LE(0, 30);
  centralDirectory.writeUInt16LE(0, 32);
  centralDirectory.writeUInt16LE(0, 34);
  centralDirectory.writeUInt16LE(0, 36);
  centralDirectory.writeUInt32LE(0, 38);
  centralDirectory.writeUInt32LE(0, 42);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(1, 8);
  eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(centralDirectory.length + nameBytes.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([
    localHeader,
    nameBytes,
    compressed,
    centralDirectory,
    nameBytes,
    eocd
  ]);
}

function findEndOfCentralDirectory(bytes) {
  for (let offset = bytes.length - 22; offset >= 0; offset -= 1) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) return offset;
  }
  throw new Error("EOCD not found");
}

function listZipEntries(bytes) {
  const eocdOffset = findEndOfCentralDirectory(bytes);
  const entryCount = bytes.readUInt16LE(eocdOffset + 10);
  let offset = bytes.readUInt32LE(eocdOffset + 16);
  const entries = [];

  for (let index = 0; index < entryCount; index += 1) {
    assert.equal(bytes.readUInt32LE(offset), 0x02014b50, "central directory signature");
    const method = bytes.readUInt16LE(offset + 10);
    const compressedSize = bytes.readUInt32LE(offset + 20);
    const nameLength = bytes.readUInt16LE(offset + 28);
    const extraLength = bytes.readUInt16LE(offset + 30);
    const commentLength = bytes.readUInt16LE(offset + 32);
    const localHeaderOffset = bytes.readUInt32LE(offset + 42);
    const name = bytes.slice(offset + 46, offset + 46 + nameLength).toString("utf8");
    const localNameLength = bytes.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = bytes.readUInt16LE(localHeaderOffset + 28);
    const dataOffset = localHeaderOffset + 30 + localNameLength + localExtraLength;

    entries.push({
      name,
      method,
      bytes: bytes.slice(dataOffset, dataOffset + compressedSize)
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

function syntheticRoomImage() {
  const width = 220;
  const height = 150;
  const data = new Uint8ClampedArray(width * height * 4);
  for (let index = 0; index < data.length; index += 4) {
    data[index] = 255;
    data[index + 1] = 255;
    data[index + 2] = 255;
    data[index + 3] = 255;
  }

  function darkPixel(x, y) {
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    const index = ((y * width) + x) * 4;
    data[index] = 0;
    data[index + 1] = 0;
    data[index + 2] = 0;
    data[index + 3] = 255;
  }

  function line(x1, y1, x2, y2, gap) {
    const minX = Math.min(x1, x2);
    const maxX = Math.max(x1, x2);
    const minY = Math.min(y1, y2);
    const maxY = Math.max(y1, y2);
    for (let y = minY; y <= maxY; y += 1) {
      for (let x = minX; x <= maxX; x += 1) {
        if (gap && x >= gap.x1 && x <= gap.x2 && y >= gap.y1 && y <= gap.y2) continue;
        darkPixel(x, y);
      }
    }
  }

  line(40, 30, 42, 120);
  line(158, 30, 160, 120, { x1: 158, x2: 160, y1: 70, y2: 88 });
  line(40, 30, 160, 32);
  line(40, 118, 160, 120);

  return { data, width, height };
}

async function run() {
  const { api, elements, context } = loadAppHarness();

  assert.equal(api.progressColor(0), "rgb(216, 66, 47)", "0% should be red");
  assert.equal(api.progressColor(50), "rgb(217, 163, 33)", "50% should be amber");
  assert.equal(api.progressColor(100), "rgb(8, 88, 43)", "100% should be dark green");

  const csvText = [
    "page,page_name,room_id,percent,points,plan_width,plan_height",
    "2,Plan 2,A101,25,\"[[72,72],[357,72],[357,258],[72,258]]\",1200,800"
  ].join("\n");
  const csvRows = api.parseCsv(csvText);
  const updates = api.csvRowsToUpdates(csvRows);
  assert.equal(updates.length, 1, "CSV should produce one room update");
  assert.equal(updates[0].id, "A101");
  assert.equal(updates[0].percent, 25);
  assert.equal(updates[0].pageIndex, 1);
  assert.equal(updates[0].pageName, "Plan 2");
  assert.equal(JSON.stringify(updates[0].points[2]), JSON.stringify([357, 258]));
  assert.equal(api.csvCell("[[1,2],[3,4]]"), '"[[1,2],[3,4]]"', "CSV cell quoting");
  assert.equal(JSON.stringify(api.parseCsvPoints("1:2;3:4")), JSON.stringify([[1, 2], [3, 4]]), "semicolon point parsing");
  assert.equal(api.extractZlCode("Room ZL-101"), "ZL-101", "ZL code extraction");
  assert.equal(api.extractZlCode("ZONE ZL 204A"), "ZL204A", "spaced ZL code extraction");
  assert.equal(
    api.projectNameFromFiles([{ name: "Building Level 01.pdf" }, { name: "Building Level 02.pdf" }]),
    "Building Level",
    "multi-file project name should use the shared filename prefix"
  );
  assert.deepEqual(
    api.planToPages({
      name: "Level Pack",
      pages: [
        { id: "page-1", name: "Level 01", plan: { name: "Level 01", src: "data:one", width: 10, height: 10 }, rooms: [] },
        { id: "page-2", name: "Level 02", plan: { name: "Level 02", src: "data:two", width: 20, height: 20 }, rooms: [] }
      ]
    }, 1).map((page) => page.id),
    ["file-2-page-1", "file-2-page-2"],
    "multi-file PDFs should flatten into unique page ids"
  );
  const twoSinglePagePdfs = [
    {
      name: "Ground Floor",
      plan: { name: "Ground Floor", sourceType: "pdf", originalFileName: "Ground Floor.pdf", pageNumber: 1, pageCount: 1 }
    },
    {
      name: "First Floor",
      plan: { name: "First Floor", sourceType: "pdf", originalFileName: "First Floor.pdf", pageNumber: 1, pageCount: 1 }
    }
  ];
  assert.deepEqual(
    twoSinglePagePdfs.map((page, index) => api.pageTabLabel(page, index, twoSinglePagePdfs)),
    ["Ground Floor", "First Floor"],
    "separate single-page PDFs should use filename tab labels, not duplicate Page 1 labels"
  );
  assert.equal(
    api.pageTabLabel({ customName: true, name: "Renamed Basement", plan: { sourceType: "pdf", pageNumber: 1, pageCount: 1 } }, 0, []),
    "Renamed Basement",
    "renamed pages should use the custom tab label"
  );
  assert.deepEqual(
    api.excelPlanSheetNames([
      { name: "Ground Floor", plan: {} },
      { name: "First/Floor:East", plan: {} },
      { name: "Progress", plan: {} },
      { name: "Ground Floor", plan: {} }
    ]),
    ["Ground Floor", "First Floor East", "Progress 2", "Ground Floor 2"],
    "Excel plan sheets should use unique safe page names"
  );
  assert.equal(api.normalizeRotation(-90), 270, "negative rotation normalization");
  assert.equal(JSON.stringify(api.rotatedPlanSize(1200, 800, 90)), JSON.stringify({ width: 800, height: 1200 }), "90 degree rotation swaps page dimensions");
  assert.equal(JSON.stringify(api.rotatePoint([100, 200], 1200, 800, 90)), JSON.stringify([600, 100]), "rotate point clockwise");
  assert.equal(JSON.stringify(api.rotatePoint([100, 200], 1200, 800, 270)), JSON.stringify([200, 1100]), "rotate point counter-clockwise");
  assert.equal(JSON.stringify(api.rotatePoints([[0, 0], [1200, 800]], 1200, 800, 180)), JSON.stringify([[1200, 800], [0, 0]]), "rotate room points 180 degrees");
  assert.equal(
    JSON.stringify(api.rotateDetectedLabel({ id: "ZL-1", x: 100, y: 200, width: 40, height: 12 }, 1200, 800, 90)),
    JSON.stringify({ id: "ZL-1", x: 600, y: 100, width: 12, height: 40 }),
    "rotated ZL labels should move and swap label bounds"
  );

  const zlRuns = [
    { text: "ZL", x: 80, y: 50, width: 14, height: 10, centerX: 87, centerY: 55 },
    { text: "204A", x: 98, y: 50, width: 28, height: 10, centerX: 112, centerY: 55 }
  ];
  const zlLabels = api.zlLabelsFromTextRuns(zlRuns, 220, 150);
  assert.equal(zlLabels.length, 1, "ZL text runs should produce one label");
  assert.equal(zlLabels[0].id, "ZL204A");

  const detectedTextLabels = api.zlLabelsFromDetectedText([
    { rawValue: "Room ZL-301", boundingBox: { x: 81, y: 52, width: 58, height: 14 } },
    { rawValue: "A101", boundingBox: { x: 10, y: 10, width: 30, height: 12 } }
  ], 220, 150);
  assert.equal(detectedTextLabels.length, 1, "drawing text detection should filter to ZL labels");
  assert.equal(detectedTextLabels[0].id, "ZL-301");
  assert.equal(JSON.stringify(api.mergeDetectedZlLabels(zlLabels, detectedTextLabels).map((label) => label.id).sort()), JSON.stringify(["ZL-301", "ZL204A"]));

  const roomImage = syntheticRoomImage();
  const detectedBox = api.findBoxBoundaryForLabel(roomImage, roomImage.width, roomImage.height, {
    id: "ZL204A",
    x: 100,
    y: 75,
    width: 44,
    height: 12
  });
  assert(detectedBox, "box detector should tolerate a door gap");
  assert(Math.abs(detectedBox.points[0][0] - 40) <= 5, "left wall detection");
  assert(Math.abs(detectedBox.points[1][0] - 160) <= 5, "right wall detection");
  assert(Math.abs(detectedBox.points[0][1] - 30) <= 5, "top wall detection");
  assert(Math.abs(detectedBox.points[2][1] - 120) <= 5, "bottom wall detection");

  const vbaSource = "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\nIf Intersect(Target, Sh.Range(\"A:B,E:F\")) Is Nothing Then Exit Sub\nEnd Sub\nPublic Sub RefreshZoneColours()\nEnd Sub";
  let fetchedVbaUrl = "";
  let copiedText = "";
  context.fetch = async (url, options) => {
    fetchedVbaUrl = url;
    assert.equal(options.cache, "no-cache", "VBA source should bypass stale browser cache");
    return {
      ok: true,
      async text() {
        return vbaSource;
      }
    };
  };
  context.navigator.clipboard = {
    async writeText(text) {
      copiedText = text;
    }
  };
  assert.equal(await api.fetchVbaSourceCode(), vbaSource, "VBA source fetch");
  assert.equal(fetchedVbaUrl, "vendor/excel/ThisWorkbookCode.bas", "VBA source URL");
  await api.copyVbaSourceCode();
  assert.equal(copiedText, vbaSource, "VBA source should copy to clipboard");
  assert.equal(elements.get("saveStatus").textContent, "VBA copied");
  assert.equal(elements.get("copyVbaButton").disabled, false);

  copiedText = "";
  context.navigator.clipboard = null;
  await api.copyVbaSourceCode();
  assert.equal(copiedText, "", "clipboard should not be used when unavailable");
  assert.equal(elements.get("saveStatus").textContent, "VBA downloaded");

  context.fetch = async () => ({ ok: false });
  const originalWarn = context.console.warn;
  context.console.warn = () => {};
  await api.copyVbaSourceCode();
  context.console.warn = originalWarn;
  assert.equal(elements.get("saveStatus").textContent, "VBA file opened", "blocked fetch should fall back to linked VBA file");

  const liveMacro = { live: true, sourceName: "local-template.xlsm", bytes: new Uint8Array([1, 2, 3]) };
  const snapshotMacro = { live: false, sourceName: "placeholder sample macro", bytes: new Uint8Array([4, 5, 6]) };
  const room = { id: "A101", percent: 25, points: [[72, 72], [357, 72], [357, 258], [72, 258]] };
  const livePlan = api.planSheetXml(liveMacro);
  const liveProgress = api.progressSheetXml([room], liveMacro);
  const renamedProgress = api.progressSheetXml([{ room, sheetName: "Ground Floor", plan: {} }], liveMacro);
  const snapshotProgress = api.progressSheetXml([], snapshotMacro);
  const drawing = api.drawingXml([room], 1200, 800);

  assert.match(livePlan, /Live macro included/);
  assert.match(livePlan, /columns A, B, E, or F/);
  assert.match(liveProgress, /<c r="B2" t="inlineStr" s="4"><is><t>Plan<\/t><\/is><\/c>/, "progress row exports target plan sheet");
  assert.match(renamedProgress, /<c r="B2" t="inlineStr" s="4"><is><t>Ground Floor<\/t><\/is><\/c>/, "renamed page sheet reference exports to Progress column B");
  assert.match(liveProgress, /<c r="E2" s="4"><v>25<\/v><\/c>/, "percent exports as numeric E cell");
  assert.match(liveProgress, /Live macro included/);
  assert.match(liveProgress, /edit A, B, E, or F/);
  assert.match(liveProgress, /local-template\.xlsm/);
  assert.match(snapshotProgress, /Snapshot only/);
  assert.match(snapshotProgress, /Install macro template/);
  assert.match(drawing, /name="Zone_A101"/);
  assert.match(drawing, /name="Label_A101"/);
  assert.match(drawing, /FFFFFF/, "label text should be white");

  const lShapeRoom = {
    id: "L101",
    percent: 55,
    points: [[0, 0], [120, 0], [120, 40], [40, 40], [40, 120], [0, 120]]
  };
  const lLabelBox = api.roomLabelBox(lShapeRoom.points, lShapeRoom.id);
  assert(
    lLabelBox.x + lLabelBox.width <= 40.1 || lLabelBox.y + lLabelBox.height <= 40.1,
    `L-shaped room label should stay inside a coloured arm, got ${JSON.stringify(lLabelBox)}`
  );
  assert.match(api.drawingXml([lShapeRoom], 140, 140), /name="Label_L101"/, "L-shaped room label exports");

  const packageBytes = Buffer.from(api.createZip([
    ["[Content_Types].xml", api.contentTypesXml()],
    ["xl/_rels/workbook.xml.rels", api.workbookRelsXml()],
    ["xl/worksheets/sheet1.xml", livePlan],
    ["xl/worksheets/sheet2.xml", liveProgress],
    ["xl/drawings/drawing1.xml", drawing],
    ["xl/vbaProject.bin", liveMacro.bytes]
  ]));
  const entries = listZipEntries(packageBytes);
  assert(entries.some((entry) => entry.name === "xl/vbaProject.bin"), "XLSM package should contain vbaProject.bin");
  assert(entries.some((entry) => entry.name === "xl/drawings/drawing1.xml"), "XLSM package should contain drawing XML");
  assert.equal(entries.find((entry) => entry.name === "xl/vbaProject.bin").bytes.length, 3);
  assert.match(entries.find((entry) => entry.name === "[Content_Types].xml").bytes.toString("utf8"), /macroEnabled\.main\+xml/);
  assert.match(entries.find((entry) => entry.name === "xl/_rels/workbook.xml.rels").bytes.toString("utf8"), /vbaProject/);
  assert.match(api.contentTypesXml(3, 2), /\/xl\/worksheets\/sheet3\.xml/, "multi-page content types include extra worksheet");
  assert.match(api.contentTypesXml(3, 2), /\/xl\/drawings\/drawing2\.xml/, "multi-page content types include extra drawing");
  assert.match(api.workbookRelsXml(3), /worksheets\/sheet3\.xml/, "multi-page workbook rels include extra worksheet");

  const macroPayload = new Uint8Array(4096).map((_, index) => index % 251);
  const deflatedWorkbook = new Uint8Array(makeDeflatedZip("xl/vbaProject.bin", macroPayload));
  const extractedMacro = await api.extractVbaProjectFromWorkbook(deflatedWorkbook);
  assert.equal(extractedMacro.length, macroPayload.length, "extracted macro length");
  assert(extractedMacro.every((value, index) => value === macroPayload[index]), "deflated vbaProject.bin extraction");

  const installedBytes = new Uint8Array(2048).map((_, index) => index % 255);
  api.saveInstalledVbaProject(installedBytes, "local-template.xlsm");
  api.renderMacroStatus();
  const project = await api.loadVbaProject();
  assert.equal(project.live, true, "installed macro should be treated as live");
  assert.equal(project.sourceName, "local-template.xlsm");
  assert.equal(elements.get("macroStatus").textContent, "Live XLSM");
  assert.equal(elements.get("macroStatus").dataset.state, "live");
  assert.equal(elements.get("downloadMacroButton").hidden, false);
  assert.equal(elements.get("clearMacroButton").hidden, false);
  assert.match(elements.get("exportExcelButton").title, /live macro-enabled/);

  console.log("Smoke test passed: CSV, colours, Auto ZL detection, VBA setup, XLSM XML, ZIP packaging, macro extraction, and macro status.");
}

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
