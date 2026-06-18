(() => {
  const failures = [];
  const checks = {};

  function check(name, condition, detail = "") {
    checks[name] = Boolean(condition);
    if (!condition) failures.push(detail ? `${name}: ${detail}` : name);
  }

  function text(selector) {
    return document.querySelector(selector)?.textContent?.trim() || "";
  }

  const appScript = Array.from(document.scripts)
    .map((script) => script.getAttribute("src"))
    .filter(Boolean)
    .find((src) => src.startsWith("app.js"));
  const stylesheet = Array.from(document.querySelectorAll('link[rel="stylesheet"]'))
    .map((link) => link.getAttribute("href"))
    .filter(Boolean)
    .find((href) => href.startsWith("styles.css"));
  const shapes = Array.from(document.querySelectorAll("polygon.room-shape[data-room-id]"));
  const labels = Array.from(document.querySelectorAll("g.room-label[data-room-id]"));
  const roomIds = shapes.map((shape) => shape.dataset.roomId).sort();
  const macroStatus = document.getElementById("macroStatus");
  const planFileInput = document.getElementById("planFileInput");
  const overlay = document.getElementById("overlay");
  const floorImage = document.getElementById("floorImage");
  const b203 = document.querySelector('polygon.room-shape[data-room-id="B203"]');
  const a101Label = document.querySelector('g.room-label[data-room-id="A101"]');
  const opacityValues = shapes
    .map((shape) => Number(shape.getAttribute("fill-opacity")))
    .filter(Number.isFinite);
  const desktopLayout = window.innerWidth > 920;
  const bodyStyle = getComputedStyle(document.body);
  const sidePanelStyle = getComputedStyle(document.querySelector(".side-panel"));
  const toolsStyle = getComputedStyle(document.querySelector(".side-panel .panel-section"));

  check("title", document.title === "Floor Plan Progress Tracker", document.title);
  check("current app script", appScript === "app.js?v=21", appScript || "missing");
  check("current stylesheet", stylesheet === "styles.css?v=3", stylesheet || "missing");
  check("desktop body scroll lock", !desktopLayout || bodyStyle.overflow === "hidden", bodyStyle.overflow);
  check("sidebar scroll pane", !desktopLayout || sidePanelStyle.overflowY === "auto", sidePanelStyle.overflowY);
  check("sticky tools panel", !desktopLayout || toolsStyle.position === "sticky", toolsStyle.position);
  check("pdf support", document.documentElement.dataset.pdfSupport === "ready", document.documentElement.dataset.pdfSupport || "missing");
  check("floor image loaded", Boolean(floorImage?.getAttribute("src")), "missing floor image src");
  check("overlay viewBox", overlay?.getAttribute("viewBox") === "0 0 1200 800", overlay?.getAttribute("viewBox") || "missing");
  check("sample room count", shapes.length === 6, `found ${shapes.length}`);
  check("sample label count", labels.length === 6, `found ${labels.length}`);
  check("page tabs ready", Boolean(document.getElementById("pageTabs")), "missing page tabs");
  check("sample room ids", JSON.stringify(roomIds) === JSON.stringify(["A101", "A102", "A103", "B201", "B202", "B203"]), roomIds.join(","));
  check("100 percent colour", b203?.getAttribute("fill") === "rgb(8, 88, 43)", b203?.getAttribute("fill") || "missing");
  check("transparent overlays", opacityValues.length === 6 && opacityValues.every((value) => value > 0 && value < 1), opacityValues.join(","));
  check("label text", a101Label?.textContent?.includes("15%") && a101Label?.textContent?.includes("A101"), a101Label?.textContent || "missing");
  check("room list count", document.querySelectorAll(".room-item").length === 6, `found ${document.querySelectorAll(".room-item").length}`);
  check("header average", /6 rooms mapped, 51% average complete/.test(text("#planMeta")), text("#planMeta"));
  check("macro status", ["snapshot", "live"].includes(macroStatus?.dataset.state), macroStatus?.outerHTML || "missing");
  check("export button", Boolean(document.getElementById("exportExcelButton")), "missing export button");
  check("multi-plan input", Boolean(planFileInput?.multiple), "plan file picker should allow multiple files");
  check("auto ZL button", Boolean(document.getElementById("autoZlButton")), "missing auto ZL button");
  check("copy VBA button", Boolean(document.getElementById("copyVbaButton")), "missing copy VBA button");
  check("macro installer", Boolean(document.getElementById("installMacroButton")), "missing install macro button");
  check("macro bin button", Boolean(document.getElementById("downloadMacroButton")), "missing macro bin button");
  check("csv controls", Boolean(document.getElementById("importCsvButton") && document.getElementById("exportCsvButton")), "missing CSV controls");
  check("legend", document.querySelectorAll(".legend .swatch").length === 3, `found ${document.querySelectorAll(".legend .swatch").length}`);

  return {
    ok: failures.length === 0,
    failures,
    checks,
    details: {
      appScript,
      stylesheet,
      roomIds,
      macroStatus: macroStatus?.textContent?.trim() || "",
      macroState: macroStatus?.dataset.state || "",
      pdfSupport: document.documentElement.dataset.pdfSupport || "",
      opacityValues
    }
  };
})()
