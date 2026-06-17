(function () {
  "use strict";

  const STORAGE_KEY = "floor-plan-progress-tracker";
  const PDF_WORKER_SRC = "vendor/pdf.worker.min.js";
  const PDF_RENDER_LONG_EDGE = 2400;
  const EXCEL_VBA_PROJECT_SRC = "vendor/excel/vbaProject.bin";
  const EMUS_PER_PIXEL = 9525;
  const PROGRESS_COLORS = {
    low: [216, 66, 47],
    mid: [217, 163, 33],
    high: [8, 88, 43]
  };
  const SAMPLE_PROJECT = {
    name: "Sample Level 01",
    plan: {
      name: "Sample Level 01",
      src: "assets/sample-floor.svg",
      sourceType: "image",
      width: 1200,
      height: 800
    },
    settings: {
      opacity: 48,
      showLabels: true,
      zoom: 1
    },
    rooms: [
      { id: "A101", percent: 15, points: [[72, 72], [357, 72], [357, 258], [72, 258]] },
      { id: "A102", percent: 45, points: [[379, 72], [655, 72], [655, 258], [379, 258]] },
      { id: "A103", percent: 80, points: [[678, 72], [1118, 72], [1118, 258], [678, 258]] },
      { id: "B201", percent: 5, points: [[72, 456], [357, 456], [357, 720], [72, 720]] },
      { id: "B202", percent: 60, points: [[379, 456], [655, 456], [655, 720], [379, 720]] },
      { id: "B203", percent: 100, points: [[678, 456], [1118, 456], [1118, 720], [678, 720]] }
    ]
  };

  const el = {
    projectName: document.getElementById("projectName"),
    planTitle: document.getElementById("planTitle"),
    planMeta: document.getElementById("planMeta"),
    floorImage: document.getElementById("floorImage"),
    overlay: document.getElementById("overlay"),
    stage: document.getElementById("stage"),
    stageScroller: document.getElementById("stageScroller"),
    saveStatus: document.getElementById("saveStatus"),
    roomCount: document.getElementById("roomCount"),
    roomList: document.getElementById("roomList"),
    roomSearch: document.getElementById("roomSearch"),
    selectModeButton: document.getElementById("selectModeButton"),
    drawModeButton: document.getElementById("drawModeButton"),
    drawCard: document.getElementById("drawCard"),
    toolHint: document.getElementById("toolHint"),
    draftRoomId: document.getElementById("draftRoomId"),
    draftPercent: document.getElementById("draftPercent"),
    finishRoomButton: document.getElementById("finishRoomButton"),
    cancelDraftButton: document.getElementById("cancelDraftButton"),
    noSelection: document.getElementById("noSelection"),
    roomEditor: document.getElementById("roomEditor"),
    selectedRoomId: document.getElementById("selectedRoomId"),
    selectedPercent: document.getElementById("selectedPercent"),
    selectedPercentNumber: document.getElementById("selectedPercentNumber"),
    selectedPercentOutput: document.getElementById("selectedPercentOutput"),
    deleteRoomButton: document.getElementById("deleteRoomButton"),
    zoomControl: document.getElementById("zoomControl"),
    opacityControl: document.getElementById("opacityControl"),
    labelToggle: document.getElementById("labelToggle"),
    loadPlanButton: document.getElementById("loadPlanButton"),
    loadProjectButton: document.getElementById("loadProjectButton"),
    saveProjectButton: document.getElementById("saveProjectButton"),
    importCsvButton: document.getElementById("importCsvButton"),
    exportCsvButton: document.getElementById("exportCsvButton"),
    exportExcelButton: document.getElementById("exportExcelButton"),
    resetSampleButton: document.getElementById("resetSampleButton"),
    planFileInput: document.getElementById("planFileInput"),
    projectFileInput: document.getElementById("projectFileInput"),
    csvFileInput: document.getElementById("csvFileInput")
  };

  let state = loadInitialProject();
  let selectedRoomId = null;
  let mode = "select";
  let draftPoints = [];
  let draggingVertex = null;
  let saveTimer = null;

  init();

  function init() {
    normalizeProject(state);
    configurePdfRenderer();
    bindEvents();
    loadPlanImage(state.plan.src);
    render();
  }

  function bindEvents() {
    el.selectModeButton.addEventListener("click", () => setMode("select"));
    el.drawModeButton.addEventListener("click", () => setMode("draw"));
    el.finishRoomButton.addEventListener("click", finishDraftRoom);
    el.cancelDraftButton.addEventListener("click", cancelDraft);
    el.draftRoomId.addEventListener("input", updateFinishButton);
    el.draftPercent.addEventListener("input", updateFinishButton);

    el.selectedRoomId.addEventListener("change", commitSelectedRoomId);
    el.selectedRoomId.addEventListener("blur", commitSelectedRoomId);
    el.selectedRoomId.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        el.selectedRoomId.blur();
      }
    });

    function commitSelectedRoomId() {
      const room = getSelectedRoom();
      if (!room) return;
      const nextId = el.selectedRoomId.value.trim();
      if (!nextId) {
        el.selectedRoomId.value = room.id;
        return;
      }
      room.id = uniqueRoomId(nextId, room.id);
      selectedRoomId = room.id;
      queueSave();
      render();
    }

    el.selectedPercent.addEventListener("input", () => updateSelectedPercent(el.selectedPercent.value));
    el.selectedPercentNumber.addEventListener("input", () => updateSelectedPercent(el.selectedPercentNumber.value));
    el.deleteRoomButton.addEventListener("click", deleteSelectedRoom);

    el.roomSearch.addEventListener("input", renderRoomList);
    el.zoomControl.addEventListener("change", () => {
      state.settings.zoom = Number(el.zoomControl.value);
      queueSave();
      renderStageSize();
    });
    el.opacityControl.addEventListener("input", () => {
      state.settings.opacity = Number(el.opacityControl.value);
      queueSave();
      render();
    });
    el.labelToggle.addEventListener("change", () => {
      state.settings.showLabels = el.labelToggle.checked;
      queueSave();
      render();
    });

    el.overlay.addEventListener("pointerdown", handleOverlayPointerDown);
    el.overlay.addEventListener("pointermove", handleOverlayPointerMove);
    el.overlay.addEventListener("pointerup", stopDraggingVertex);
    el.overlay.addEventListener("pointercancel", stopDraggingVertex);
    el.overlay.addEventListener("dblclick", (event) => {
      if (mode === "draw" && draftPoints.length >= 3) {
        event.preventDefault();
        finishDraftRoom();
      }
    });

    window.addEventListener("keydown", (event) => {
      if ((event.key === "Delete" || event.key === "Backspace") && selectedRoomId && !isTyping()) {
        event.preventDefault();
        deleteSelectedRoom();
      }
      if (event.key === "Escape" && mode === "draw") {
        cancelDraft();
      }
    });

    el.loadPlanButton.addEventListener("click", () => el.planFileInput.click());
    el.planFileInput.addEventListener("change", handlePlanFile);
    el.loadProjectButton.addEventListener("click", () => el.projectFileInput.click());
    el.projectFileInput.addEventListener("change", handleProjectFile);
    el.saveProjectButton.addEventListener("click", downloadProject);
    el.importCsvButton.addEventListener("click", () => el.csvFileInput.click());
    el.csvFileInput.addEventListener("change", handleCsvFile);
    el.exportCsvButton.addEventListener("click", downloadCsv);
    el.exportExcelButton.addEventListener("click", downloadExcelWorkbook);
    el.resetSampleButton.addEventListener("click", () => {
      if (!confirm("Reload the sample project? Unsaved changes in this browser will be replaced.")) return;
      state = clone(SAMPLE_PROJECT);
      selectedRoomId = null;
      setMode("select");
      loadPlanImage(state.plan.src);
      queueSave();
      render();
    });
  }

  function loadInitialProject() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (error) {
        console.warn("Could not read saved project", error);
      }
    }
    return clone(SAMPLE_PROJECT);
  }

  function normalizeProject(project) {
    project.name = project.name || "Untitled Floor Plan";
    project.plan = project.plan || clone(SAMPLE_PROJECT.plan);
    project.plan.width = Number(project.plan.width) || 1200;
    project.plan.height = Number(project.plan.height) || 800;
    project.plan.sourceType = project.plan.sourceType || "image";
    project.rooms = Array.isArray(project.rooms) ? project.rooms : [];
    project.settings = Object.assign({ opacity: 48, showLabels: true, zoom: 1 }, project.settings || {});
    project.rooms.forEach((room, index) => {
      room.id = room.id || `ROOM-${index + 1}`;
      room.percent = clamp(Number(room.percent) || 0, 0, 100);
      room.points = Array.isArray(room.points) ? room.points : [];
    });
  }

  function loadPlanImage(src) {
    el.floorImage.onload = () => {
      if (!state.plan.width || !state.plan.height) {
        state.plan.width = el.floorImage.naturalWidth || 1200;
        state.plan.height = el.floorImage.naturalHeight || 800;
      }
      renderStageSize();
    };
    el.floorImage.onerror = () => {
      el.saveStatus.textContent = "Plan missing";
    };
    el.floorImage.src = src;
  }

  function render() {
    normalizeProject(state);
    renderHeader();
    renderStageSize();
    renderOverlay();
    renderRoomList();
    renderEditor();
    renderControls();
  }

  function renderHeader() {
    const roomTotal = state.rooms.length;
    const avg = roomTotal ? Math.round(state.rooms.reduce((sum, room) => sum + room.percent, 0) / roomTotal) : 0;
    const pdfMeta = state.plan.sourceType === "pdf"
      ? `, PDF page ${state.plan.pageNumber || 1}${state.plan.pageCount ? ` of ${state.plan.pageCount}` : ""}`
      : "";
    el.projectName.textContent = state.name;
    el.planTitle.textContent = state.plan.name || state.name;
    el.planMeta.textContent = `${roomTotal} room${roomTotal === 1 ? "" : "s"} mapped, ${avg}% average complete${pdfMeta}`;
    el.roomCount.textContent = String(roomTotal);
  }

  function renderStageSize() {
    const zoom = Number(state.settings.zoom) || 1;
    const width = Math.max(320, state.plan.width * zoom);
    const height = Math.max(240, state.plan.height * zoom);
    el.stage.style.width = `${width}px`;
    el.stage.style.height = `${height}px`;
    el.overlay.setAttribute("viewBox", `0 0 ${state.plan.width} ${state.plan.height}`);
    el.overlay.setAttribute("width", String(width));
    el.overlay.setAttribute("height", String(height));
    el.zoomControl.value = String(zoom);
  }

  function renderOverlay() {
    el.overlay.innerHTML = "";

    for (const room of state.rooms) {
      if (room.points.length < 3) continue;
      const polygon = makeSvg("polygon", {
        points: pointsToString(room.points),
        fill: progressColor(room.percent),
        "fill-opacity": String(state.settings.opacity / 100),
        stroke: selectedRoomId === room.id ? "#0b5368" : "#243140",
        "stroke-width": selectedRoomId === room.id ? "3" : "1.5",
        class: `room-shape${selectedRoomId === room.id ? " is-selected" : ""}`,
        "data-room-id": room.id
      });
      polygon.addEventListener("pointerdown", (event) => {
        if (mode !== "select") return;
        event.stopPropagation();
        selectedRoomId = room.id;
        render();
      });
      el.overlay.appendChild(polygon);
    }

    if (state.settings.showLabels) {
      for (const room of state.rooms) {
        if (room.points.length < 3) continue;
        el.overlay.appendChild(makeRoomLabel(room));
      }
    }

    if (selectedRoomId && mode === "select") {
      const selected = getSelectedRoom();
      if (selected) {
        selected.points.forEach((point, index) => {
          const handle = makeSvg("circle", {
            cx: point[0],
            cy: point[1],
            r: vertexRadius(),
            class: "vertex",
            "data-index": String(index)
          });
          handle.addEventListener("pointerdown", (event) => {
            event.stopPropagation();
            draggingVertex = { roomId: selected.id, index };
            handle.setPointerCapture(event.pointerId);
          });
          el.overlay.appendChild(handle);
        });
      }
    }

    if (draftPoints.length) {
      el.overlay.appendChild(makeSvg("polyline", {
        points: pointsToString(draftPoints),
        class: "draft-line"
      }));
      draftPoints.forEach((point) => {
        el.overlay.appendChild(makeSvg("circle", {
          cx: point[0],
          cy: point[1],
          r: vertexRadius(),
          class: "draft-point"
        }));
      });
    }
  }

  function makeRoomLabel(room) {
    const center = polygonCentroid(room.points);
    const group = makeSvg("g", { class: "room-label", "data-room-id": room.id });
    const labelWidth = Math.max(74, Math.min(128, Math.max(room.id.length * 9 + 24, 78)));
    const labelHeight = 44;
    const bg = makeSvg("rect", {
      x: center[0] - labelWidth / 2,
      y: center[1] - labelHeight / 2,
      width: labelWidth,
      height: labelHeight,
      rx: 5,
      class: "room-label-bg"
    });
    const text = makeSvg("text", { x: center[0], y: center[1] - 4 });
    const percent = makeSvg("tspan", {
      x: center[0],
      dy: 0,
      class: "room-label-percent"
    });
    percent.textContent = `${room.percent}%`;
    const id = makeSvg("tspan", {
      x: center[0],
      dy: 18,
      class: "room-label-id"
    });
    id.textContent = room.id;
    text.append(percent, id);
    group.append(bg, text);
    group.addEventListener("pointerdown", (event) => {
      if (mode !== "select") return;
      event.stopPropagation();
      selectedRoomId = room.id;
      render();
    });
    return group;
  }

  function renderRoomList() {
    const query = el.roomSearch.value.trim().toLowerCase();
    const rooms = state.rooms
      .slice()
      .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }))
      .filter((room) => room.id.toLowerCase().includes(query));

    el.roomList.innerHTML = "";
    if (!rooms.length) {
      const empty = document.createElement("div");
      empty.className = "empty-room-list";
      empty.textContent = query ? "No matching rooms." : "Draw a room to start.";
      el.roomList.appendChild(empty);
      return;
    }

    for (const room of rooms) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `room-item${room.id === selectedRoomId ? " is-selected" : ""}`;
      button.addEventListener("click", () => {
        selectedRoomId = room.id;
        setMode("select");
        render();
      });

      const dot = document.createElement("span");
      dot.className = "room-dot";
      dot.style.backgroundColor = progressColor(room.percent);

      const label = document.createElement("span");
      const code = document.createElement("span");
      code.className = "room-code";
      code.textContent = room.id;
      const points = document.createElement("span");
      points.className = "room-points";
      points.textContent = `${room.points.length} points`;
      label.append(code, points);

      const pct = document.createElement("span");
      pct.className = "room-percent";
      pct.textContent = `${room.percent}%`;

      button.append(dot, label, pct);
      el.roomList.appendChild(button);
    }
  }

  function renderEditor() {
    const room = getSelectedRoom();
    el.noSelection.hidden = Boolean(room);
    el.roomEditor.hidden = !room;
    el.deleteRoomButton.disabled = !room;
    if (!room) return;

    if (document.activeElement !== el.selectedRoomId) {
      el.selectedRoomId.value = room.id;
    }
    el.selectedPercent.value = String(room.percent);
    el.selectedPercentNumber.value = String(room.percent);
    el.selectedPercentOutput.textContent = `${room.percent}%`;
  }

  function renderControls() {
    el.selectModeButton.classList.toggle("is-active", mode === "select");
    el.drawModeButton.classList.toggle("is-active", mode === "draw");
    el.drawCard.hidden = mode !== "draw";
    el.opacityControl.value = String(state.settings.opacity);
    el.labelToggle.checked = Boolean(state.settings.showLabels);
    el.toolHint.textContent = mode === "draw"
      ? `${draftPoints.length} point${draftPoints.length === 1 ? "" : "s"} placed. Click room corners, then Finish.`
      : "Select a room to edit it, or switch to Draw Room and click around a room boundary.";
    updateFinishButton();
  }

  function setMode(nextMode) {
    mode = nextMode;
    if (mode !== "draw") draftPoints = [];
    render();
  }

  function handleOverlayPointerDown(event) {
    const point = svgPointFromEvent(event);
    if (mode === "draw") {
      draftPoints.push(point);
      updateFinishButton();
      render();
      return;
    }

    selectedRoomId = null;
    render();
  }

  function handleOverlayPointerMove(event) {
    if (!draggingVertex) return;
    const room = state.rooms.find((item) => item.id === draggingVertex.roomId);
    if (!room) return;
    room.points[draggingVertex.index] = svgPointFromEvent(event);
    queueSave();
    render();
  }

  function stopDraggingVertex() {
    if (!draggingVertex) return;
    draggingVertex = null;
    queueSave();
  }

  function finishDraftRoom() {
    const id = uniqueRoomId(el.draftRoomId.value.trim() || `ROOM-${state.rooms.length + 1}`);
    if (draftPoints.length < 3) return;
    state.rooms.push({
      id,
      percent: clamp(Number(el.draftPercent.value) || 0, 0, 100),
      points: draftPoints.map((point) => [round(point[0]), round(point[1])])
    });
    selectedRoomId = id;
    draftPoints = [];
    el.draftRoomId.value = "";
    el.draftPercent.value = "0";
    setMode("select");
    queueSave();
  }

  function cancelDraft() {
    draftPoints = [];
    render();
  }

  function updateFinishButton() {
    const hasId = Boolean(el.draftRoomId.value.trim());
    el.finishRoomButton.disabled = draftPoints.length < 3 || !hasId;
  }

  function updateSelectedPercent(value) {
    const room = getSelectedRoom();
    if (!room) return;
    room.percent = clamp(Number(value) || 0, 0, 100);
    queueSave();
    render();
  }

  function deleteSelectedRoom() {
    const room = getSelectedRoom();
    if (!room) return;
    if (!confirm(`Delete room ${room.id}?`)) return;
    state.rooms = state.rooms.filter((item) => item.id !== room.id);
    selectedRoomId = null;
    queueSave();
    render();
  }

  async function handlePlanFile(event) {
    const file = event.target.files[0];
    event.target.value = "";
    if (!file) return;

    el.loadPlanButton.disabled = true;
    el.saveStatus.textContent = "Loading plan...";

    try {
      const plan = isPdfFile(file) ? await readPdfPlan(file) : await readImagePlan(file);
      applyLoadedPlan(plan);
    } catch (error) {
      if (error.name !== "AbortError") {
        alert(error.message || "That floor plan could not be opened.");
        console.error(error);
      }
      el.saveStatus.textContent = "Autosaved";
    } finally {
      el.loadPlanButton.disabled = false;
    }
  }

  function handleProjectFile(event) {
    const file = event.target.files[0];
    event.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const project = JSON.parse(String(reader.result));
        state = project;
        selectedRoomId = null;
        normalizeProject(state);
        loadPlanImage(state.plan.src);
        queueSave();
        render();
      } catch (error) {
        alert("That JSON file could not be opened.");
        console.error(error);
      }
    };
    reader.readAsText(file);
  }

  function handleCsvFile(event) {
    const file = event.target.files[0];
    event.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const rows = parseCsv(String(reader.result));
      const updates = csvRowsToUpdates(rows);
      let applied = 0;
      let created = 0;
      let geometryUpdated = 0;
      const missing = [];

      for (const update of updates) {
        const room = state.rooms.find((item) => item.id.toLowerCase() === update.id.toLowerCase());
        const points = scaleImportedPoints(update.points, update.planWidth, update.planHeight);
        if (!room) {
          if (points.length >= 3) {
            state.rooms.push({
              id: uniqueRoomId(update.id),
              percent: update.percent,
              points
            });
            created += 1;
            applied += 1;
            geometryUpdated += 1;
          } else {
            missing.push(update.id);
          }
          continue;
        }
        room.percent = update.percent;
        if (points.length >= 3) {
          room.points = points;
          geometryUpdated += 1;
        }
        applied += 1;
      }

      queueSave();
      render();
      const detail = [
        `${applied} room${applied === 1 ? "" : "s"} updated`,
        created ? `${created} room${created === 1 ? "" : "s"} created` : "",
        geometryUpdated ? `${geometryUpdated} overlay${geometryUpdated === 1 ? "" : "s"} mapped` : "",
        missing.length ? `${missing.length} room code${missing.length === 1 ? "" : "s"} had no saved points` : ""
      ].filter(Boolean).join(". ");
      alert(`${detail}.`);
    };
    reader.readAsText(file);
  }

  function downloadProject() {
    downloadFile(`${safeFileName(state.name)}.floor-plan.json`, JSON.stringify(state, null, 2), "application/json");
  }

  function downloadCsv() {
    const rows = [["room_id", "percent", "points", "plan_width", "plan_height"]];
    state.rooms
      .slice()
      .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }))
      .forEach((room) => rows.push([
        room.id,
        String(room.percent),
        JSON.stringify(room.points),
        String(state.plan.width),
        String(state.plan.height)
      ]));
    const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
    downloadFile(`${safeFileName(state.name)}-progress.csv`, csv, "text/csv");
  }

  async function downloadExcelWorkbook() {
    el.exportExcelButton.disabled = true;
    el.saveStatus.textContent = "Building XLSM...";

    try {
      const blob = await buildExcelWorkbook();
      downloadBlob(`${safeFileName(state.name)}-floor-plan-progress.xlsm`, blob);
      el.saveStatus.textContent = "Autosaved";
    } catch (error) {
      alert(error.message || "The Excel workbook could not be exported.");
      console.error(error);
      el.saveStatus.textContent = "Use Save";
    } finally {
      el.exportExcelButton.disabled = false;
    }
  }

  async function buildExcelWorkbook() {
    const planImage = await renderPlanToPng();
    const vbaProject = await loadVbaProject();
    const rooms = state.rooms
      .slice()
      .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
    const files = [
      ["[Content_Types].xml", contentTypesXml()],
      ["_rels/.rels", rootRelsXml()],
      ["xl/workbook.xml", workbookXml()],
      ["xl/_rels/workbook.xml.rels", workbookRelsXml()],
      ["xl/styles.xml", stylesXml()],
      ["xl/worksheets/sheet1.xml", planSheetXml()],
      ["xl/worksheets/_rels/sheet1.xml.rels", planSheetRelsXml()],
      ["xl/worksheets/sheet2.xml", progressSheetXml(rooms)],
      ["xl/drawings/drawing1.xml", drawingXml(rooms, planImage.width, planImage.height)],
      ["xl/drawings/_rels/drawing1.xml.rels", drawingRelsXml()],
      ["xl/media/floor-plan.png", planImage.bytes],
      ["xl/vbaProject.bin", vbaProject],
      ["docProps/core.xml", corePropsXml()],
      ["docProps/app.xml", appPropsXml()]
    ];

    return new Blob([createZip(files)], {
      type: "application/vnd.ms-excel.sheet.macroEnabled.12"
    });
  }

  async function renderPlanToPng() {
    const image = await loadImage(state.plan.src);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(state.plan.width || image.naturalWidth || 1200));
    canvas.height = Math.max(1, Math.round(state.plan.height || image.naturalHeight || 800));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return {
      bytes: dataUrlToBytes(canvas.toDataURL("image/png")).bytes,
      width: canvas.width,
      height: canvas.height
    };
  }

  function loadImage(src) {
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("The floor plan image could not be prepared for Excel export."));
      image.src = src;
    });
  }

  async function loadVbaProject() {
    const response = await fetch(EXCEL_VBA_PROJECT_SRC);
    if (!response.ok) {
      throw new Error("The macro template is missing. Upload vendor/excel/vbaProject.bin with the site.");
    }
    return new Uint8Array(await response.arrayBuffer());
  }

  function contentTypesXml() {
    return xmlDecl(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    </Types>`);
  }

  function rootRelsXml() {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>`);
  }

  function workbookXml() {
    return xmlDecl(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="23426"/>
      <workbookPr codeName="ThisWorkbook"/>
      <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="18000" windowHeight="10000"/></bookViews>
      <sheets>
        <sheet name="Plan" sheetId="1" r:id="rId1"/>
        <sheet name="Progress" sheetId="2" r:id="rId2"/>
      </sheets>
    </workbook>`);
  }

  function workbookRelsXml() {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
    </Relationships>`);
  }

  function stylesXml() {
    return xmlDecl(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="3">
        <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
        <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
        <font><b/><sz val="14"/><color rgb="FF18202B"/><name val="Calibri"/><family val="2"/></font>
      </fonts>
      <fills count="4">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF152330"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFEFF4F8"/><bgColor indexed="64"/></patternFill></fill>
      </fills>
      <borders count="2">
        <border><left/><right/><top/><bottom/><diagonal/></border>
        <border><left style="thin"><color rgb="FFD7DEE8"/></left><right style="thin"><color rgb="FFD7DEE8"/></right><top style="thin"><color rgb="FFD7DEE8"/></top><bottom style="thin"><color rgb="FFD7DEE8"/></bottom><diagonal/></border>
      </borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
      <cellXfs count="5">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
      </cellXfs>
      <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
    </styleSheet>`);
  }

  function planSheetXml() {
    return xmlDecl(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr codeName="Sheet1"/>
      <sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
      <sheetFormatPr defaultRowHeight="15"/>
      <cols><col min="1" max="12" width="14" customWidth="1"/></cols>
      <sheetData>
        <row r="1" ht="24" customHeight="1">
          ${cell("A1", "Floor Plan Progress Export", "inlineStr", 2)}
        </row>
        <row r="2">
          ${cell("A2", "Edit room percentages on the Progress sheet. Zone drawing shapes are preserved on this plan sheet.", "inlineStr", 0)}
        </row>
      </sheetData>
      <pageMargins left="0.25" right="0.25" top="0.25" bottom="0.25" header="0.1" footer="0.1"/>
      <drawing r:id="rId1"/>
    </worksheet>`);
  }

  function planSheetRelsXml() {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
    </Relationships>`);
  }

  function progressSheetXml(rooms) {
    const rows = [
      `<row r="1">${cell("A1", "Room ID", "inlineStr", 1)}${cell("B1", "Percent", "inlineStr", 1)}${cell("C1", "Current Colour", "inlineStr", 1)}${cell("D1", "Overlay Opacity", "inlineStr", 1)}${cell("E1", "Zone Shape", "inlineStr", 1)}${cell("F1", "Points", "inlineStr", 1)}</row>`
    ];

    rooms.forEach((room, index) => {
      const row = index + 2;
      rows.push(`<row r="${row}">
        ${cell(`A${row}`, room.id, "inlineStr", 4)}
        ${cell(`B${row}`, room.percent, "n", 4)}
        ${cell(`C${row}`, rgbToHex(progressColor(room.percent)), "inlineStr", 4)}
        ${cell(`D${row}`, `${state.settings.opacity}%`, "inlineStr", 4)}
        ${cell(`E${row}`, excelShapeName(room), "inlineStr", 4)}
        ${cell(`F${row}`, JSON.stringify(room.points), "inlineStr", 4)}
      </row>`);
    });

    return xmlDecl(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr codeName="Sheet2"/>
      <sheetViews><sheetView workbookViewId="0"/></sheetViews>
      <sheetFormatPr defaultRowHeight="18"/>
      <cols>
        <col min="1" max="1" width="18" customWidth="1"/>
        <col min="2" max="2" width="12" customWidth="1"/>
        <col min="3" max="3" width="16" customWidth="1"/>
        <col min="4" max="4" width="16" customWidth="1"/>
        <col min="5" max="5" width="26" customWidth="1"/>
        <col min="6" max="6" width="80" customWidth="1"/>
      </cols>
      <sheetData>${rows.join("")}</sheetData>
      <autoFilter ref="A1:F${Math.max(1, rooms.length + 1)}"/>
      <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    </worksheet>`);
  }

  function drawingXml(rooms, imageWidth, imageHeight) {
    const imageExtent = { width: imageWidth, height: imageHeight };
    const image = oneCellAnchor(0, 0, 0, 4, imageExtent.width, imageExtent.height, pictureXml(2, "Floor plan image", "rId1", imageExtent.width, imageExtent.height));
    const shapes = rooms
      .filter((room) => room.points.length >= 3)
      .map((room, index) => polygonShapeAnchor(room, index + 3))
      .join("");

    return xmlDecl(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      ${image}
      ${shapes}
    </xdr:wsDr>`);
  }

  function drawingRelsXml() {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/floor-plan.png"/>
    </Relationships>`);
  }

  function pictureXml(id, name, relationshipId, width, height) {
    return `<xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="${id}" name="${xmlEscape(name)}"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="${relationshipId}"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>`;
  }

  function polygonShapeAnchor(room, id) {
    const bounds = pointBounds(room.points);
    const width = Math.max(1, bounds.maxX - bounds.minX);
    const height = Math.max(1, bounds.maxY - bounds.minY);
    const fill = rgbToHex(progressColor(room.percent)).replace("#", "");
    const alpha = Math.round(clamp(state.settings.opacity, 0, 100) * 1000);
    const center = polygonCentroid(room.points);
    const labelY = clamp(center[1] - bounds.minY - 16, 8, height - 8);
    const idY = clamp(center[1] - bounds.minY + 8, 8, height - 8);
    const pathPoints = room.points.map((point) => [
      Math.round(((point[0] - bounds.minX) / width) * 100000),
      Math.round(((point[1] - bounds.minY) / height) * 100000)
    ]);
    const first = pathPoints[0];
    const rest = pathPoints.slice(1).map((point) => `<a:lnTo><a:pt x="${point[0]}" y="${point[1]}"/></a:lnTo>`).join("");

    const shape = `<xdr:sp macro="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="${id}" name="${xmlEscape(excelShapeName(room))}" descr="${xmlEscape(room.id)}"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/></a:xfrm>
        <a:custGeom>
          <a:avLst/>
          <a:gdLst/>
          <a:ahLst/>
          <a:cxnLst/>
          <a:rect l="0" t="0" r="100000" b="100000"/>
          <a:pathLst>
            <a:path w="100000" h="100000">
              <a:moveTo><a:pt x="${first[0]}" y="${first[1]}"/></a:moveTo>
              ${rest}
              <a:close/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
        <a:solidFill><a:srgbClr val="${fill}"><a:alpha val="${alpha}"/></a:srgbClr></a:solidFill>
        <a:ln w="14288"><a:solidFill><a:srgbClr val="243140"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr wrap="square" anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>${xmlEscape(`${room.percent}%`)}</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:rPr lang="en-US" sz="1100" b="1"/><a:t>${xmlEscape(room.id)}</a:t></a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>`;

    return oneCellAnchor(bounds.minX, bounds.minY, 0, 4, width, height, shape);
  }

  function oneCellAnchor(x, y, col, row, width, height, body) {
    return `<xdr:oneCellAnchor>
      <xdr:from>
        <xdr:col>${col}</xdr:col>
        <xdr:colOff>${pxToEmu(x)}</xdr:colOff>
        <xdr:row>${row}</xdr:row>
        <xdr:rowOff>${pxToEmu(y)}</xdr:rowOff>
      </xdr:from>
      <xdr:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/>
      ${body}
      <xdr:clientData/>
    </xdr:oneCellAnchor>`;
  }

  function corePropsXml() {
    const now = new Date().toISOString();
    return xmlDecl(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dc:title>${xmlEscape(state.name)} Floor Plan Progress</dc:title>
      <dc:creator>Floor Plan Progress Tracker</dc:creator>
      <cp:lastModifiedBy>Floor Plan Progress Tracker</cp:lastModifiedBy>
      <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
      <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
    </cp:coreProperties>`);
  }

  function appPropsXml() {
    return xmlDecl(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <Application>Microsoft Excel</Application>
      <DocSecurity>0</DocSecurity>
      <ScaleCrop>false</ScaleCrop>
      <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>
      <TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Plan</vt:lpstr><vt:lpstr>Progress</vt:lpstr></vt:vector></TitlesOfParts>
    </Properties>`);
  }

  function createZip(entries) {
    const encoder = new TextEncoder();
    const prepared = entries.map(([name, data]) => {
      const nameBytes = encoder.encode(name);
      const bytes = typeof data === "string" ? encoder.encode(data) : data;
      return { name, nameBytes, bytes, crc: crc32(bytes) };
    });
    const chunks = [];
    const central = [];
    let offset = 0;

    for (const file of prepared) {
      const localHeader = new Uint8Array(30 + file.nameBytes.length);
      const view = new DataView(localHeader.buffer);
      view.setUint32(0, 0x04034b50, true);
      view.setUint16(4, 20, true);
      view.setUint16(6, 0x0800, true);
      view.setUint16(8, 0, true);
      view.setUint16(10, 0, true);
      view.setUint16(12, 0, true);
      view.setUint32(14, file.crc, true);
      view.setUint32(18, file.bytes.length, true);
      view.setUint32(22, file.bytes.length, true);
      view.setUint16(26, file.nameBytes.length, true);
      view.setUint16(28, 0, true);
      localHeader.set(file.nameBytes, 30);
      chunks.push(localHeader, file.bytes);

      const centralHeader = new Uint8Array(46 + file.nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0x0800, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, 0, true);
      centralView.setUint16(14, 0, true);
      centralView.setUint32(16, file.crc, true);
      centralView.setUint32(20, file.bytes.length, true);
      centralView.setUint32(24, file.bytes.length, true);
      centralView.setUint16(28, file.nameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(file.nameBytes, 46);
      central.push(centralHeader);

      offset += localHeader.length + file.bytes.length;
    }

    const centralOffset = offset;
    const centralSize = central.reduce((sum, item) => sum + item.length, 0);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, prepared.length, true);
    endView.setUint16(10, prepared.length, true);
    endView.setUint32(12, centralSize, true);
    endView.setUint32(16, centralOffset, true);

    return concatUint8(chunks.concat(central, [end]));
  }

  function cell(ref, value, type, style) {
    const styleAttr = style ? ` s="${style}"` : "";
    if (type === "n") {
      return `<c r="${ref}"${styleAttr}><v>${Number(value) || 0}</v></c>`;
    }
    return `<c r="${ref}" t="inlineStr"${styleAttr}><is><t>${xmlEscape(value)}</t></is></c>`;
  }

  function xmlDecl(xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml.replace(/>\s+</g, "><").trim()}`;
  }

  function xmlEscape(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function pointBounds(points) {
    return points.reduce((bounds, point) => ({
      minX: Math.min(bounds.minX, point[0]),
      minY: Math.min(bounds.minY, point[1]),
      maxX: Math.max(bounds.maxX, point[0]),
      maxY: Math.max(bounds.maxY, point[1])
    }), { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity });
  }

  function excelShapeName(room) {
    return `Zone_${room.id.replace(/[^A-Za-z0-9_]/g, "_")}`;
  }

  function rgbToHex(rgb) {
    const match = String(rgb).match(/\d+/g);
    if (!match || match.length < 3) return "#000000";
    return `#${match.slice(0, 3).map((item) => Number(item).toString(16).padStart(2, "0")).join("")}`;
  }

  function pxToEmu(value) {
    return Math.round(Number(value) * EMUS_PER_PIXEL);
  }

  function dataUrlToBytes(dataUrl) {
    const match = String(dataUrl).match(/^data:([^;,]+)?(;base64)?,(.*)$/);
    if (!match) throw new Error("Invalid data URL.");
    const mimeType = match[1] || "application/octet-stream";
    const isBase64 = Boolean(match[2]);
    const payload = match[3];
    const binary = isBase64 ? atob(payload) : decodeURIComponent(payload);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return { mimeType, bytes };
  }

  function concatUint8(parts) {
    const length = parts.reduce((sum, item) => sum + item.length, 0);
    const out = new Uint8Array(length);
    let offset = 0;
    for (const part of parts) {
      out.set(part, offset);
      offset += part.length;
    }
    return out;
  }

  function crc32(bytes) {
    let crc = -1;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xff];
    }
    return (crc ^ -1) >>> 0;
  }

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let crc = i;
      for (let j = 0; j < 8; j += 1) {
        crc = (crc & 1) ? (0xedb88320 ^ (crc >>> 1)) : (crc >>> 1);
      }
      table[i] = crc >>> 0;
    }
    return table;
  })();

  function configurePdfRenderer() {
    const pdfLib = window.pdfjsLib || window["pdfjs-dist/build/pdf"];
    document.documentElement.dataset.pdfSupport = pdfLib ? "ready" : "missing";
    if (!pdfLib || !pdfLib.GlobalWorkerOptions) return;
    window.pdfjsLib = pdfLib;
    pdfLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_SRC;
  }

  function isPdfFile(file) {
    return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
  }

  function readImagePlan(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("That image file could not be read."));
      reader.onload = () => {
        const src = String(reader.result);
        const image = new Image();
        image.onload = () => {
          const name = file.name.replace(/\.[^.]+$/, "");
          resolve({
            name,
            src,
            sourceType: "image",
            originalFileName: file.name,
            width: image.naturalWidth || 1200,
            height: image.naturalHeight || 800
          });
        };
        image.onerror = () => reject(new Error("That image file could not be opened as a floor plan."));
        image.src = src;
      };
      reader.readAsDataURL(file);
    });
  }

  async function readPdfPlan(file) {
    if (!window.pdfjsLib) {
      throw new Error("PDF support is missing. Make sure vendor/pdf.min.js and vendor/pdf.worker.min.js are uploaded with the site.");
    }

    const data = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data }).promise;

    try {
      const pageNumber = choosePdfPageNumber(pdf.numPages);
      const page = await pdf.getPage(pageNumber);
      const baseViewport = page.getViewport({ scale: 1 });
      const scale = pdfRenderScale(baseViewport.width, baseViewport.height);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);

      const context = canvas.getContext("2d", { alpha: false });
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, canvas.width, canvas.height);

      await page.render({
        canvasContext: context,
        viewport,
        background: "white"
      }).promise;

      const baseName = file.name.replace(/\.[^.]+$/, "");
      return {
        name: pdf.numPages > 1 ? `${baseName} - Page ${pageNumber}` : baseName,
        src: canvas.toDataURL("image/png"),
        sourceType: "pdf",
        originalFileName: file.name,
        pageNumber,
        pageCount: pdf.numPages,
        width: canvas.width,
        height: canvas.height
      };
    } finally {
      if (typeof pdf.destroy === "function") {
        await pdf.destroy();
      }
    }
  }

  function choosePdfPageNumber(pageCount) {
    if (pageCount <= 1) return 1;
    const answer = prompt(`This PDF has ${pageCount} pages. Which page number should be used?`, "1");
    if (answer === null) {
      const error = new Error("PDF loading cancelled.");
      error.name = "AbortError";
      throw error;
    }
    return clamp(Math.floor(Number(answer) || 1), 1, pageCount);
  }

  function pdfRenderScale(width, height) {
    const longestEdge = Math.max(width, height) || PDF_RENDER_LONG_EDGE;
    return PDF_RENDER_LONG_EDGE / longestEdge;
  }

  function applyLoadedPlan(plan) {
    const keepRooms = state.rooms.length && confirm("Keep existing room areas on this new floor plan?");
    state.name = plan.name;
    state.plan = plan;

    if (!keepRooms) {
      state.rooms = [];
      selectedRoomId = null;
    }

    loadPlanImage(state.plan.src);
    queueSave();
    render();
  }

  function queueSave() {
    el.saveStatus.textContent = "Saving...";
    window.clearTimeout(saveTimer);
    saveTimer = window.setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        el.saveStatus.textContent = "Autosaved";
      } catch (error) {
        console.warn("Autosave failed. Download a project JSON backup instead.", error);
        el.saveStatus.textContent = "Use Save";
      }
    }, 180);
  }

  function getSelectedRoom() {
    return state.rooms.find((room) => room.id === selectedRoomId) || null;
  }

  function uniqueRoomId(rawId, currentId) {
    const base = rawId || "ROOM";
    if (base === currentId) return base;
    const existing = new Set(state.rooms.filter((room) => room.id !== currentId).map((room) => room.id.toLowerCase()));
    if (!existing.has(base.toLowerCase())) return base;
    let index = 2;
    while (existing.has(`${base}-${index}`.toLowerCase())) {
      index += 1;
    }
    return `${base}-${index}`;
  }

  function svgPointFromEvent(event) {
    const rect = el.overlay.getBoundingClientRect();
    const x = ((event.clientX - rect.left) / rect.width) * state.plan.width;
    const y = ((event.clientY - rect.top) / rect.height) * state.plan.height;
    return [round(clamp(x, 0, state.plan.width)), round(clamp(y, 0, state.plan.height))];
  }

  function progressColor(percent) {
    const pct = clamp(Number(percent) || 0, 0, 100);
    if (pct <= 50) {
      return mixColor(PROGRESS_COLORS.low, PROGRESS_COLORS.mid, pct / 50);
    }
    return mixColor(PROGRESS_COLORS.mid, PROGRESS_COLORS.high, (pct - 50) / 50);
  }

  function mixColor(a, b, amount) {
    const mixed = a.map((start, index) => Math.round(start + (b[index] - start) * amount));
    return `rgb(${mixed[0]}, ${mixed[1]}, ${mixed[2]})`;
  }

  function polygonCentroid(points) {
    let twiceArea = 0;
    let cx = 0;
    let cy = 0;

    for (let i = 0; i < points.length; i += 1) {
      const current = points[i];
      const next = points[(i + 1) % points.length];
      const cross = current[0] * next[1] - next[0] * current[1];
      twiceArea += cross;
      cx += (current[0] + next[0]) * cross;
      cy += (current[1] + next[1]) * cross;
    }

    if (Math.abs(twiceArea) < 0.001) {
      const average = points.reduce((sum, point) => [sum[0] + point[0], sum[1] + point[1]], [0, 0]);
      return [average[0] / points.length, average[1] / points.length];
    }

    return [cx / (3 * twiceArea), cy / (3 * twiceArea)];
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let quoted = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      const next = text[i + 1];

      if (char === "\"" && quoted && next === "\"") {
        cell += "\"";
        i += 1;
      } else if (char === "\"") {
        quoted = !quoted;
      } else if (char === "," && !quoted) {
        row.push(cell);
        cell = "";
      } else if ((char === "\n" || char === "\r") && !quoted) {
        if (char === "\r" && next === "\n") i += 1;
        row.push(cell);
        if (row.some((value) => value.trim() !== "")) rows.push(row);
        row = [];
        cell = "";
      } else {
        cell += char;
      }
    }

    row.push(cell);
    if (row.some((value) => value.trim() !== "")) rows.push(row);
    return rows;
  }

  function csvRowsToUpdates(rows) {
    if (!rows.length) return [];
    const header = rows[0].map((value) => value.trim().toLowerCase());
    const hasHeader = header.includes("room_id") || header.includes("room") || header.includes("code");
    const roomIndex = hasHeader
      ? firstExistingIndex(header, ["room_id", "room", "code", "room code", "roomid"])
      : 0;
    const percentIndex = hasHeader
      ? firstExistingIndex(header, ["percent", "percentage", "%", "complete", "progress"])
      : 1;
    const pointsIndex = hasHeader
      ? optionalHeaderIndex(header, ["points", "polygon_points", "room_points", "points_json", "overlay_points"])
      : -1;
    const planWidthIndex = hasHeader
      ? optionalHeaderIndex(header, ["plan_width", "width", "source_width", "drawing_width"])
      : -1;
    const planHeightIndex = hasHeader
      ? optionalHeaderIndex(header, ["plan_height", "height", "source_height", "drawing_height"])
      : -1;
    const dataRows = hasHeader ? rows.slice(1) : rows;

    return dataRows
      .map((row) => ({
        id: String(row[roomIndex] || "").trim(),
        percent: clamp(Number(String(row[percentIndex] || "").replace("%", "").trim()) || 0, 0, 100),
        points: pointsIndex >= 0 ? parseCsvPoints(row[pointsIndex]) : [],
        planWidth: planWidthIndex >= 0 ? Number(row[planWidthIndex]) : null,
        planHeight: planHeightIndex >= 0 ? Number(row[planHeightIndex]) : null
      }))
      .filter((row) => row.id);
  }

  function optionalHeaderIndex(values, options) {
    for (const option of options) {
      const index = values.indexOf(option);
      if (index >= 0) return index;
    }
    return -1;
  }

  function parseCsvPoints(value) {
    const text = String(value || "").trim();
    if (!text) return [];

    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) return parsed;
    } catch (error) {
      console.warn("CSV points were not JSON.", error);
    }

    return text
      .split(";")
      .map((pair) => pair.trim().split(/[|\s:]+/).map(Number))
      .filter((pair) => pair.length >= 2 && pair.every(Number.isFinite))
      .map((pair) => [pair[0], pair[1]]);
  }

  function scaleImportedPoints(points, sourceWidth, sourceHeight) {
    if (!Array.isArray(points)) return [];
    const scaleX = sourceWidth && state.plan.width ? state.plan.width / sourceWidth : 1;
    const scaleY = sourceHeight && state.plan.height ? state.plan.height / sourceHeight : 1;

    return points
      .map((point) => {
        if (!Array.isArray(point) || point.length < 2) return null;
        const x = Number(point[0]);
        const y = Number(point[1]);
        if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
        return [
          round(clamp(x * scaleX, 0, state.plan.width)),
          round(clamp(y * scaleY, 0, state.plan.height))
        ];
      })
      .filter(Boolean);
  }

  function firstExistingIndex(values, options) {
    for (const option of options) {
      const index = values.indexOf(option);
      if (index >= 0) return index;
    }
    return 0;
  }

  function downloadFile(fileName, content, type) {
    const blob = new Blob([content], { type });
    downloadBlob(fileName, blob);
  }

  function downloadBlob(fileName, blob) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function csvCell(value) {
    const text = String(value);
    if (!/[",\n\r]/.test(text)) return text;
    return `"${text.replace(/"/g, "\"\"")}"`;
  }

  function makeSvg(tag, attrs) {
    const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
    for (const [key, value] of Object.entries(attrs)) {
      node.setAttribute(key, value);
    }
    return node;
  }

  function pointsToString(points) {
    return points.map((point) => `${point[0]},${point[1]}`).join(" ");
  }

  function vertexRadius() {
    return Math.max(5, 7 / (Number(state.settings.zoom) || 1));
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function round(value) {
    return Math.round(value * 10) / 10;
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function safeFileName(value) {
    return String(value || "floor-plan").replace(/[^a-z0-9-_]+/gi, "-").replace(/^-+|-+$/g, "") || "floor-plan";
  }

  function isTyping() {
    const tag = document.activeElement && document.activeElement.tagName;
    return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";
  }
})();
