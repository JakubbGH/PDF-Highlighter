(function () {
  "use strict";

  const STORAGE_KEY = "floor-plan-progress-tracker";
  const VBA_PROJECT_STORAGE_KEY = "floor-plan-progress-tracker-vba-project";
  const PDF_WORKER_SRC = "vendor/pdf.worker.min.js";
  const PDF_RENDER_LONG_EDGE = 2400;
  const EXCEL_VBA_PROJECT_SRC = "vendor/excel/vbaProject.bin";
  const EXCEL_VBA_SOURCE_SRC = "vendor/excel/ThisWorkbookCode.bas";
  const PLACEHOLDER_VBA_PROJECT_SHA256 = "0ced1464b3677e98f5e3a8c5d80135e18dc98dca39299f1a8cfd2a00999fbf9f";
  const PLACEHOLDER_VBA_PROJECT_BYTES = 15872;
  const EXCEL_PLAN_TOP_OFFSET = 84;
  const EXCEL_LABEL_BOX_PADDING = 4;
  const EMUS_PER_PIXEL = 9525;
  const AUTO_ZL_ROOM_PATTERN = /ZL/i;
  const AUTO_DETECT_WALL = {
    darkThreshold: 150,
    lineThickness: 5,
    minCoverage: 0.38,
    minDarkRatio: 0.055,
    minRoomSize: 36,
    minLabelDistance: 24,
    scanStep: 2,
    bandSize: 180,
    maxSpanRatio: 0.46
  };
  const DEFLATE_LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  const DEFLATE_LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  const DEFLATE_DISTANCE_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  const DEFLATE_DISTANCE_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
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
    pageTabs: document.getElementById("pageTabs"),
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
    autoZlButton: document.getElementById("autoZlButton"),
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
    copyVbaButton: document.getElementById("copyVbaButton"),
    installMacroButton: document.getElementById("installMacroButton"),
    downloadMacroButton: document.getElementById("downloadMacroButton"),
    clearMacroButton: document.getElementById("clearMacroButton"),
    macroStatus: document.getElementById("macroStatus"),
    resetSampleButton: document.getElementById("resetSampleButton"),
    planFileInput: document.getElementById("planFileInput"),
    projectFileInput: document.getElementById("projectFileInput"),
    csvFileInput: document.getElementById("csvFileInput"),
    macroTemplateInput: document.getElementById("macroTemplateInput")
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
    renderMacroStatus();
    render();
  }

  function bindEvents() {
    el.selectModeButton.addEventListener("click", () => setMode("select"));
    el.drawModeButton.addEventListener("click", () => setMode("draw"));
    el.autoZlButton.addEventListener("click", autoDetectZlRooms);
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
    el.copyVbaButton.addEventListener("click", copyVbaSourceCode);
    el.installMacroButton.addEventListener("click", () => el.macroTemplateInput.click());
    el.downloadMacroButton.addEventListener("click", downloadInstalledMacroProject);
    el.clearMacroButton.addEventListener("click", clearInstalledMacroTemplate);
    el.macroTemplateInput.addEventListener("change", handleMacroTemplateFile);
    el.resetSampleButton.addEventListener("click", () => {
      if (!confirm("Reload the sample project? Unsaved changes in this browser will be replaced.")) return;
      state = clone(SAMPLE_PROJECT);
      normalizeProject(state);
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
    project.settings = Object.assign({ opacity: 48, showLabels: true, zoom: 1 }, project.settings || {});
    if (!Array.isArray(project.pages) || !project.pages.length) {
      project.pages = [pageFromLegacyProject(project)];
    }

    project.pages = project.pages
      .map((page, index) => normalizePage(page, index))
      .filter(Boolean);

    if (!project.pages.length) {
      project.pages = [normalizePage(pageFromLegacyProject(clone(SAMPLE_PROJECT)), 0)];
    }

    project.activePageIndex = clamp(Math.floor(Number(project.activePageIndex) || 0), 0, project.pages.length - 1);
    syncActivePageReferences(project);
  }

  function pageFromLegacyProject(project) {
    const plan = project.plan || clone(SAMPLE_PROJECT.plan);
    const name = plan.name || project.name || `Page 1`;
    return {
      id: "page-1",
      name,
      plan,
      rooms: Array.isArray(project.rooms) ? project.rooms : []
    };
  }

  function normalizePage(page, index) {
    if (!page) return null;
    const normalized = typeof page === "object" ? page : {};
    normalized.id = normalized.id || `page-${index + 1}`;
    normalized.plan = normalized.plan || clone(SAMPLE_PROJECT.plan);
    normalized.plan.name = normalized.plan.name || normalized.name || `Page ${index + 1}`;
    normalized.plan.width = Number(normalized.plan.width) || 1200;
    normalized.plan.height = Number(normalized.plan.height) || 800;
    normalized.plan.sourceType = normalized.plan.sourceType || "image";
    normalized.plan.zlLabels = Array.isArray(normalized.plan.zlLabels)
      ? normalized.plan.zlLabels
        .map((label) => normalizeDetectedZlLabel(label, normalized.plan.width, normalized.plan.height))
        .filter(Boolean)
      : [];
    normalized.name = normalized.name || normalized.plan.name || `Page ${index + 1}`;
    normalized.rooms = Array.isArray(normalized.rooms) ? normalized.rooms : [];
    normalized.rooms.forEach((room, roomIndex) => {
      room.id = room.id || `ROOM-${roomIndex + 1}`;
      room.percent = clamp(Number(room.percent) || 0, 0, 100);
      room.points = Array.isArray(room.points) ? room.points : [];
      room.points = room.points
        .map((point) => {
          if (!Array.isArray(point) || point.length < 2) return null;
          const x = Number(point[0]);
          const y = Number(point[1]);
          if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
          return [
            round(clamp(x, 0, normalized.plan.width)),
            round(clamp(y, 0, normalized.plan.height))
          ];
        })
        .filter(Boolean);
    });
    return normalized;
  }

  function activePage(project = state) {
    return project.pages[project.activePageIndex] || project.pages[0];
  }

  function syncActivePageReferences(project = state) {
    const page = activePage(project);
    project.plan = page.plan;
    project.rooms = page.rooms;
  }

  function setCurrentRooms(rooms) {
    const page = activePage();
    page.rooms = rooms;
    state.rooms = page.rooms;
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
    renderPageTabs();
    renderStageSize();
    renderOverlay();
    renderRoomList();
    renderEditor();
    renderControls();
  }

  function renderHeader() {
    const roomTotal = state.rooms.length;
    const avg = roomTotal ? Math.round(state.rooms.reduce((sum, room) => sum + room.percent, 0) / roomTotal) : 0;
    const pageCount = state.pages.length;
    const pageNumber = state.activePageIndex + 1;
    const pdfMeta = state.plan.sourceType === "pdf"
      ? `, PDF page ${state.plan.pageNumber || pageNumber}${state.plan.pageCount ? ` of ${state.plan.pageCount}` : ""}`
      : "";
    const multiPageMeta = pageCount > 1 ? `, tab ${pageNumber} of ${pageCount}` : "";
    el.projectName.textContent = state.name;
    el.planTitle.textContent = state.plan.name || state.name;
    el.planMeta.textContent = `${roomTotal} room${roomTotal === 1 ? "" : "s"} mapped, ${avg}% average complete${pdfMeta}${multiPageMeta}`;
    el.roomCount.textContent = String(roomTotal);
  }

  function renderPageTabs() {
    if (!el.pageTabs) return;
    const pages = state.pages || [];
    el.pageTabs.hidden = pages.length <= 1;
    el.pageTabs.innerHTML = "";
    if (pages.length <= 1) return;

    pages.forEach((page, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `page-tab${index === state.activePageIndex ? " is-active" : ""}`;
      button.textContent = pageTabLabel(page, index);
      button.title = page.name || page.plan.name || button.textContent;
      button.addEventListener("click", () => setActivePage(index));
      el.pageTabs.appendChild(button);
    });
  }

  function pageTabLabel(page, index) {
    if (page.plan.sourceType === "pdf") {
      return `Page ${page.plan.pageNumber || index + 1}`;
    }
    return page.name || page.plan.name || `Page ${index + 1}`;
  }

  function setActivePage(index) {
    const nextIndex = clamp(Math.floor(Number(index) || 0), 0, state.pages.length - 1);
    if (nextIndex === state.activePageIndex) return;

    state.activePageIndex = nextIndex;
    syncActivePageReferences();
    selectedRoomId = null;
    draftPoints = [];
    draggingVertex = null;
    mode = "select";
    loadPlanImage(state.plan.src);
    queueSave();
    render();
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

    renderDetectedZlPins();

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

  function renderDetectedZlPins() {
    const labels = state.plan.zlLabels || [];
    if (!labels.length) return;

    const mappedRoomIds = new Set(state.rooms.map((room) => room.id.toLowerCase()));
    for (const label of labels) {
      if (mappedRoomIds.has(label.id.toLowerCase())) continue;
      const group = makeSvg("g", { class: "zl-pin", "data-room-id": label.id });
      const marker = makeSvg("circle", {
        cx: label.x,
        cy: label.y,
        r: 7
      });
      const text = makeSvg("text", {
        x: label.x,
        y: label.y - 12
      });
      text.textContent = label.id;
      group.append(marker, text);
      el.overlay.appendChild(group);
    }
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
    const zlLabelCount = (state.plan.zlLabels || []).length;
    el.autoZlButton.disabled = !state.plan.src;
    el.autoZlButton.title = zlLabelCount
      ? `Detect box-like boundaries for ${zlLabelCount} ZL label${zlLabelCount === 1 ? "" : "s"}`
      : "Scan the visible drawing for ZL text, then detect box-like boundaries";
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
    setCurrentRooms(state.rooms.filter((item) => item.id !== room.id));
    selectedRoomId = null;
    queueSave();
    render();
  }

  async function handlePlanFile(event) {
    const files = Array.from(event.target.files || []);
    event.target.value = "";
    if (!files.length) return;

    el.loadPlanButton.disabled = true;
    el.saveStatus.textContent = files.length === 1 ? "Loading plan..." : `Loading ${files.length} plans...`;

    try {
      const plan = await readPlanFiles(files);
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

  async function readPlanFiles(files) {
    const loadedPlans = [];
    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      el.saveStatus.textContent = files.length === 1
        ? "Loading plan..."
        : `Loading ${index + 1} of ${files.length}: ${file.name}`;
      loadedPlans.push(isPdfFile(file) ? await readPdfPlan(file) : await readImagePlan(file));
    }

    if (loadedPlans.length === 1) return loadedPlans[0];

    return {
      name: projectNameFromFiles(files),
      pages: loadedPlans.flatMap((plan, fileIndex) => planToPages(plan, fileIndex))
    };
  }

  function planToPages(plan, fileIndex = 0) {
    const pages = Array.isArray(plan.pages) && plan.pages.length
      ? plan.pages
      : [{
        id: "page-1",
        name: plan.name,
        plan,
        rooms: []
      }];

    return pages.map((page, pageIndex) => {
      const planData = page.plan || page;
      const name = page.name || planData.name || `Plan ${fileIndex + 1}`;
      return {
        id: `file-${fileIndex + 1}-page-${pageIndex + 1}`,
        name,
        plan: Object.assign({}, planData, { name }),
        rooms: Array.isArray(page.rooms) ? page.rooms : []
      };
    });
  }

  function projectNameFromFiles(files) {
    const names = Array.from(files)
      .map((file) => String(file.name || "").replace(/\.[^.]+$/, "").trim())
      .filter(Boolean);
    if (!names.length) return "Floor Plans";
    if (names.length === 1) return names[0];

    const prefix = commonFileNamePrefix(names);
    return prefix || `${names[0]} + ${names.length - 1} more`;
  }

  function commonFileNamePrefix(names) {
    let prefix = names[0] || "";
    for (const name of names.slice(1)) {
      while (prefix && !name.toLowerCase().startsWith(prefix.toLowerCase())) {
        prefix = prefix.slice(0, -1);
      }
      if (!prefix) break;
    }

    return prefix
      .replace(/[\s._-]*\d+$/g, "")
      .replace(/[\s._-]+$/g, "")
      .trim();
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
        const page = resolveCsvUpdatePage(update);
        const rooms = page.rooms;
        const room = rooms.find((item) => item.id.toLowerCase() === update.id.toLowerCase());
        const points = scaleImportedPoints(update.points, update.planWidth, update.planHeight, page.plan);
        if (!room) {
          if (points.length >= 3) {
            rooms.push({
              id: uniqueRoomId(update.id, null, rooms),
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
    const rows = [["page", "page_name", "room_id", "percent", "points", "plan_width", "plan_height"]];
    state.pages.forEach((page, pageIndex) => {
      page.rooms
        .slice()
        .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }))
        .forEach((room) => rows.push([
          String(pageIndex + 1),
          page.name || page.plan.name || `Page ${pageIndex + 1}`,
          room.id,
          String(room.percent),
          JSON.stringify(room.points),
          String(page.plan.width),
          String(page.plan.height)
        ]));
    });
    const csv = rows.map((row) => row.map(csvCell).join(",")).join("\n");
    downloadFile(`${safeFileName(state.name)}-progress.csv`, csv, "text/csv");
  }

  async function autoDetectZlRooms() {
    let labels = (state.plan.zlLabels || []).filter((label) => AUTO_ZL_ROOM_PATTERN.test(label.id));
    if (!labels.length) {
      el.autoZlButton.disabled = true;
      el.saveStatus.textContent = "Reading drawing text...";
      const detectionResult = await detectZlLabelsFromDrawing();
      labels = detectionResult.labels;
      if (!labels.length) {
        alert(detectionResult.message);
        el.saveStatus.textContent = "No ZL text found";
        renderControls();
        return;
      }
      state.plan.zlLabels = mergeDetectedZlLabels(state.plan.zlLabels || [], labels);
      queueSave();
      render();
    }

    const existingZlRooms = state.rooms.filter((room) => AUTO_ZL_ROOM_PATTERN.test(room.id));
    const replaceExisting = existingZlRooms.length
      ? confirm(`There are already ${existingZlRooms.length} ZL room${existingZlRooms.length === 1 ? "" : "s"} mapped. Update their boundaries where detected?`)
      : false;

    el.autoZlButton.disabled = true;
    el.saveStatus.textContent = "Detecting ZL rooms...";

    try {
      const imageData = await getPlanImageData();
      let created = 0;
      let updated = 0;
      let skipped = 0;
      let failed = 0;
      const seen = new Set();

      for (const label of labels) {
        const key = label.id.toLowerCase();
        if (seen.has(key)) {
          skipped += 1;
          continue;
        }
        seen.add(key);

        const detection = findBoxBoundaryForLabel(imageData, imageData.width, imageData.height, label);
        if (!detection) {
          failed += 1;
          continue;
        }

        const existing = state.rooms.find((room) => room.id.toLowerCase() === key);
        if (existing) {
          if (!replaceExisting) {
            skipped += 1;
            continue;
          }
          existing.points = detection.points;
          updated += 1;
        } else {
          state.rooms.push({
            id: uniqueRoomId(label.id),
            percent: 0,
            points: detection.points
          });
          created += 1;
        }
      }

      if (created || updated) {
        selectedRoomId = state.rooms.find((room) => AUTO_ZL_ROOM_PATTERN.test(room.id))?.id || selectedRoomId;
        queueSave();
        render();
      }

      const detail = [
        created ? `${created} created` : "",
        updated ? `${updated} updated` : "",
        skipped ? `${skipped} skipped` : "",
        failed ? `${failed} not boxed` : ""
      ].filter(Boolean).join(", ") || "no changes";
      alert(`Auto ZL detection finished: ${detail}. Review the boxes and adjust any doorway or broken-wall misses.`);
    } catch (error) {
      alert(error.message || "Auto ZL detection could not run.");
      console.error(error);
      el.saveStatus.textContent = "Use Save";
    } finally {
      renderControls();
    }
  }

  async function getPlanImageData() {
    const { canvas, context } = await getPlanCanvas();
    return context.getImageData(0, 0, canvas.width, canvas.height);
  }

  async function getPlanCanvas() {
    const image = await loadImage(state.plan.src);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(state.plan.width || image.naturalWidth || 1200));
    canvas.height = Math.max(1, Math.round(state.plan.height || image.naturalHeight || 800));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return { canvas, context };
  }

  async function detectZlLabelsFromDrawing() {
    const TextDetectorCtor = window.TextDetector || window.ShapeTextDetector || null;
    if (typeof TextDetectorCtor !== "function") {
      return {
        labels: [],
        message: "No ZL room labels were found, and this browser does not provide local drawing-text detection. The plan may have flattened text; use a browser with TextDetector support, a text-based PDF, or draw the ZL rooms manually."
      };
    }

    try {
      const { canvas } = await getPlanCanvas();
      const detector = new TextDetectorCtor();
      const detections = await detector.detect(canvas);
      const labels = zlLabelsFromDetectedText(detections, canvas.width, canvas.height);
      return {
        labels,
        message: labels.length
          ? `${labels.length} ZL label${labels.length === 1 ? "" : "s"} found in the drawing.`
          : "The drawing-text detector ran, but it did not find any text strings containing ZL."
      };
    } catch (error) {
      console.warn("Drawing text detection failed.", error);
      return {
        labels: [],
        message: "The browser could not read text from the drawing image. The plan may need a bundled OCR engine or manual ZL seed labels."
      };
    }
  }

  async function downloadExcelWorkbook() {
    el.exportExcelButton.disabled = true;
    el.saveStatus.textContent = "Building XLSM...";

    try {
      const workbook = await buildExcelWorkbook();
      downloadBlob(`${safeFileName(state.name)}-floor-plan-progress.xlsm`, workbook.blob);
      el.saveStatus.textContent = workbook.macroProject.live ? "Live XLSM exported" : "Snapshot exported";
    } catch (error) {
      alert(error.message || "The Excel workbook could not be exported.");
      console.error(error);
      el.saveStatus.textContent = "Use Save";
    } finally {
      el.exportExcelButton.disabled = false;
    }
  }

  async function handleMacroTemplateFile() {
    const [file] = el.macroTemplateInput.files;
    if (!file) return;

    el.installMacroButton.disabled = true;
    el.saveStatus.textContent = "Installing macro...";

    try {
      const workbookBytes = new Uint8Array(await file.arrayBuffer());
      const vbaProject = await extractVbaProjectFromWorkbook(workbookBytes);
      if (await isPlaceholderVbaProject(vbaProject)) {
        throw new Error("That workbook contains the placeholder sample macro, not the floor plan refresh macro.");
      }

      saveInstalledVbaProject(vbaProject, file.name);
      el.saveStatus.textContent = "Macro installed";
      renderMacroStatus();
      alert("Excel macro template installed locally. Future XLSM exports from this browser will update box colours and labels when macros are enabled in Excel. Use Macro Bin if you need the compiled file for the repo.");
    } catch (error) {
      alert(error.message || "The macro template could not be installed.");
      console.error(error);
      el.saveStatus.textContent = "Use Save";
    } finally {
      el.installMacroButton.disabled = false;
      el.macroTemplateInput.value = "";
    }
  }

  async function copyVbaSourceCode() {
    el.copyVbaButton.disabled = true;
    el.saveStatus.textContent = "Loading VBA...";

    try {
      const sourceCode = await fetchVbaSourceCode();
      const copied = await copyTextToClipboard(sourceCode);
      if (copied) {
        el.saveStatus.textContent = "VBA copied";
        alert("VBA code copied. In Excel, press Alt+F11, open ThisWorkbook, paste it there, then save as .xlsm.");
      } else {
        downloadFile("ThisWorkbookCode.bas", sourceCode, "text/plain");
        el.saveStatus.textContent = "VBA downloaded";
        alert("Clipboard access was not available, so the VBA source file was downloaded instead. Open it, copy all text, paste it into ThisWorkbook in Excel, then save as .xlsm.");
      }
    } catch (error) {
      downloadLinkedFile("ThisWorkbookCode.bas", EXCEL_VBA_SOURCE_SRC);
      alert("The browser could not read the VBA source directly, so the source file was opened or downloaded instead. Copy all text from ThisWorkbookCode.bas, paste it into ThisWorkbook in Excel, then save as .xlsm.");
      console.warn(error.message || "The VBA source could not be copied directly.", error);
      el.saveStatus.textContent = "VBA file opened";
    } finally {
      el.copyVbaButton.disabled = false;
    }
  }

  async function fetchVbaSourceCode() {
    const response = await fetch(EXCEL_VBA_SOURCE_SRC, { cache: "no-cache" });
    if (!response.ok) {
      throw new Error("The VBA source file is missing. Upload vendor/excel/ThisWorkbookCode.bas with the site.");
    }

    const sourceCode = await response.text();
    if (!/Workbook_SheetChange/.test(sourceCode) || !/RefreshZoneColours/.test(sourceCode) || !/A:B,E:F/.test(sourceCode)) {
      throw new Error("The VBA source file does not look like the floor plan refresh macro.");
    }
    return sourceCode;
  }

  async function copyTextToClipboard(text) {
    if (!navigator.clipboard || typeof navigator.clipboard.writeText !== "function") {
      return false;
    }

    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (error) {
      console.warn("Clipboard copy failed; falling back to download.", error);
      return false;
    }
  }

  function clearInstalledMacroTemplate() {
    if (!confirm("Remove the locally installed Excel macro template from this browser?")) return;

    localStorage.removeItem(VBA_PROJECT_STORAGE_KEY);
    el.saveStatus.textContent = "Macro removed";
    renderMacroStatus();
  }

  function downloadInstalledMacroProject() {
    const installedProject = loadInstalledVbaProject();
    if (!installedProject) {
      alert("Install a macro template first.");
      renderMacroStatus();
      return;
    }

    downloadBlob("vbaProject.bin", new Blob([installedProject.bytes], {
      type: "application/vnd.ms-office.vbaProject"
    }));
    el.saveStatus.textContent = "Macro bin downloaded";
  }

  async function buildExcelWorkbook() {
    const macroProject = await loadVbaProject();
    const pageExports = [];
    for (let index = 0; index < state.pages.length; index += 1) {
      const page = state.pages[index];
      const planImage = await renderPlanToPng(page.plan);
      pageExports.push({
        page,
        sheetName: excelPlanSheetName(page, index, state.pages.length),
        sheetIndex: index + 1,
        drawingIndex: index + 1,
        mediaName: `floor-plan-${index + 1}.png`,
        image: planImage,
        rooms: page.rooms
          .slice()
          .sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }))
      });
    }

    const progressSheetIndex = pageExports.length + 1;
    const worksheetCount = progressSheetIndex;
    const progressRows = pageExports.flatMap((pageExport) => (
      pageExport.rooms.map((room) => ({
        room,
        sheetName: pageExport.sheetName,
        plan: pageExport.page.plan
      }))
    ));
    const files = [
      ["[Content_Types].xml", contentTypesXml(worksheetCount, pageExports.length)],
      ["_rels/.rels", rootRelsXml()],
      ["xl/workbook.xml", workbookXml(pageExports.map((pageExport) => pageExport.sheetName))],
      ["xl/_rels/workbook.xml.rels", workbookRelsXml(worksheetCount)],
      ["xl/styles.xml", stylesXml()]
    ];

    pageExports.forEach((pageExport) => {
      files.push(
        [`xl/worksheets/sheet${pageExport.sheetIndex}.xml`, planSheetXml(macroProject, pageExport.page, pageExport.sheetIndex)],
        [`xl/worksheets/_rels/sheet${pageExport.sheetIndex}.xml.rels`, planSheetRelsXml(pageExport.drawingIndex)],
        [`xl/drawings/drawing${pageExport.drawingIndex}.xml`, drawingXml(pageExport.rooms, pageExport.image.width, pageExport.image.height)],
        [`xl/drawings/_rels/drawing${pageExport.drawingIndex}.xml.rels`, drawingRelsXml(pageExport.mediaName)],
        [`xl/media/${pageExport.mediaName}`, pageExport.image.bytes]
      );
    });

    files.push(
      [`xl/worksheets/sheet${progressSheetIndex}.xml`, progressSheetXml(progressRows, macroProject, progressSheetIndex)],
      ["xl/vbaProject.bin", macroProject.bytes],
      ["docProps/core.xml", corePropsXml()],
      ["docProps/app.xml", appPropsXml(pageExports.map((pageExport) => pageExport.sheetName).concat("Progress"))]
    );

    return {
      blob: new Blob([createZip(files)], {
        type: "application/vnd.ms-excel.sheet.macroEnabled.12"
      }),
      macroProject
    };
  }

  async function renderPlanToPng(plan = state.plan) {
    const image = await loadImage(plan.src);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(plan.width || image.naturalWidth || 1200));
    canvas.height = Math.max(1, Math.round(plan.height || image.naturalHeight || 800));
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
    const installedProject = loadInstalledVbaProject();
    if (installedProject) {
      if (!(await isPlaceholderVbaProject(installedProject.bytes))) {
        return {
          bytes: installedProject.bytes,
          live: true,
          sourceName: installedProject.sourceName,
          statusText: "Live macro installed"
        };
      }
      localStorage.removeItem(VBA_PROJECT_STORAGE_KEY);
      renderMacroStatus();
    }

    const response = await fetch(EXCEL_VBA_PROJECT_SRC);
    if (!response.ok) {
      throw new Error("The macro template is missing. Upload vendor/excel/vbaProject.bin with the site.");
    }
    const bytes = new Uint8Array(await response.arrayBuffer());
    const isPlaceholder = await isPlaceholderVbaProject(bytes);
    if (isPlaceholder) {
      const exportSnapshot = window.confirm(
        "The Excel live-update macro template has not been installed yet. The workbook will export as a snapshot, but changing Progress columns A, B, E, or F in Excel will not refresh labels, colours, or opacity. Export the snapshot anyway?"
      );
      if (!exportSnapshot) {
        throw new Error("Excel live-update macro template is not installed yet. See vendor/excel/README.md for the one-time setup.");
      }
    }

    return {
      bytes,
      live: !isPlaceholder,
      sourceName: isPlaceholder ? "placeholder sample macro" : "site macro template",
      statusText: isPlaceholder ? "Snapshot only" : "Live macro bundled"
    };
  }

  function saveInstalledVbaProject(bytes, sourceName) {
    const payload = {
      sourceName,
      installedAt: new Date().toISOString(),
      bytes: bytesToBase64(bytes)
    };
    localStorage.setItem(VBA_PROJECT_STORAGE_KEY, JSON.stringify(payload));
  }

  function loadInstalledVbaProject() {
    const raw = localStorage.getItem(VBA_PROJECT_STORAGE_KEY);
    if (!raw) return null;

    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed.bytes !== "string") return null;
      return {
        sourceName: typeof parsed.sourceName === "string" ? parsed.sourceName : "local template",
        bytes: base64ToBytes(parsed.bytes)
      };
    } catch (error) {
      console.warn("Could not read the installed Excel macro template", error);
      localStorage.removeItem(VBA_PROJECT_STORAGE_KEY);
      return null;
    }
  }

  function loadInstalledVbaMetadata() {
    const raw = localStorage.getItem(VBA_PROJECT_STORAGE_KEY);
    if (!raw) return null;

    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed.bytes !== "string") return null;
      return {
        sourceName: typeof parsed.sourceName === "string" ? parsed.sourceName : "local template",
        installedAt: typeof parsed.installedAt === "string" ? parsed.installedAt : ""
      };
    } catch (error) {
      console.warn("Could not read the installed Excel macro template metadata", error);
      localStorage.removeItem(VBA_PROJECT_STORAGE_KEY);
      return null;
    }
  }

  function renderMacroStatus() {
    const metadata = loadInstalledVbaMetadata();
    const isInstalled = Boolean(metadata);
    const sourceName = metadata ? metadata.sourceName : "no local macro template";

    el.macroStatus.textContent = isInstalled ? "Live XLSM" : "Snapshot XLSM";
    el.macroStatus.title = isInstalled
      ? `Exports use the locally installed macro template: ${sourceName}`
      : "Exports are snapshots until a macro template is installed.";
    el.macroStatus.dataset.state = isInstalled ? "live" : "snapshot";
    el.downloadMacroButton.hidden = !isInstalled;
    el.clearMacroButton.hidden = !isInstalled;
    el.exportExcelButton.title = isInstalled
      ? "Download a live macro-enabled Excel workbook"
      : "Download a snapshot XLSM, or install a macro template for live Excel updates";
  }

  async function isPlaceholderVbaProject(bytes) {
    if (bytes.length !== PLACEHOLDER_VBA_PROJECT_BYTES) return false;
    if (!window.crypto || !window.crypto.subtle) return true;

    const digest = await window.crypto.subtle.digest("SHA-256", bytes);
    const hash = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    return hash === PLACEHOLDER_VBA_PROJECT_SHA256;
  }

  async function extractVbaProjectFromWorkbook(workbookBytes) {
    const entry = findZipEntry(workbookBytes, "xl/vbaProject.bin");
    if (!entry) {
      throw new Error("That workbook does not contain xl/vbaProject.bin. Save the template as a macro-enabled .xlsm file first.");
    }

    const compressed = workbookBytes.slice(entry.dataOffset, entry.dataOffset + entry.compressedSize);
    let data;
    if (entry.compressionMethod === 0) {
      data = compressed;
    } else if (entry.compressionMethod === 8) {
      data = await inflateZipEntry(compressed);
    } else {
      throw new Error(`That workbook uses ZIP compression method ${entry.compressionMethod}, which this local page cannot read.`);
    }

    if (entry.uncompressedSize && data.length !== entry.uncompressedSize) {
      throw new Error("The extracted VBA project size did not match the workbook directory.");
    }
    if (data.length < 1024) {
      throw new Error("The extracted VBA project is unexpectedly small. The template may not contain a compiled macro project.");
    }

    return data;
  }

  function findZipEntry(bytes, requestedName) {
    const eocdOffset = findEndOfCentralDirectory(bytes);
    if (eocdOffset < 0) {
      throw new Error("That file is not a readable Excel .xlsm package.");
    }

    const entryCount = readUint16(bytes, eocdOffset + 10);
    const centralDirectoryOffset = readUint32(bytes, eocdOffset + 16);
    let offset = centralDirectoryOffset;
    const target = requestedName.toLowerCase();
    const decoder = new TextDecoder("utf-8");

    for (let index = 0; index < entryCount; index += 1) {
      if (readUint32(bytes, offset) !== 0x02014b50) break;

      const compressionMethod = readUint16(bytes, offset + 10);
      const compressedSize = readUint32(bytes, offset + 20);
      const uncompressedSize = readUint32(bytes, offset + 24);
      const nameLength = readUint16(bytes, offset + 28);
      const extraLength = readUint16(bytes, offset + 30);
      const commentLength = readUint16(bytes, offset + 32);
      const localHeaderOffset = readUint32(bytes, offset + 42);
      const name = decoder.decode(bytes.slice(offset + 46, offset + 46 + nameLength));

      if (name.toLowerCase() === target) {
        if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff || localHeaderOffset === 0xffffffff) {
          throw new Error("ZIP64 macro templates are not supported by the local installer.");
        }
        if (readUint32(bytes, localHeaderOffset) !== 0x04034b50) {
          throw new Error("The macro template has an invalid ZIP local header.");
        }

        const localNameLength = readUint16(bytes, localHeaderOffset + 26);
        const localExtraLength = readUint16(bytes, localHeaderOffset + 28);
        return {
          compressionMethod,
          compressedSize,
          uncompressedSize,
          dataOffset: localHeaderOffset + 30 + localNameLength + localExtraLength
        };
      }

      offset += 46 + nameLength + extraLength + commentLength;
    }

    return null;
  }

  function findEndOfCentralDirectory(bytes) {
    const minOffset = Math.max(0, bytes.length - 22 - 65535);
    for (let offset = bytes.length - 22; offset >= minOffset; offset -= 1) {
      if (readUint32(bytes, offset) === 0x06054b50) return offset;
    }
    return -1;
  }

  async function inflateZipEntry(bytes) {
    if (typeof DecompressionStream !== "undefined") {
      const formats = ["deflate-raw", "deflate"];
      for (const format of formats) {
        try {
          const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(format));
          return new Uint8Array(await new Response(stream).arrayBuffer());
        } catch (error) {
          console.warn(`Browser ${format} decompression failed; using the local fallback.`, error);
        }
      }
    }

    return inflateRawDeflate(bytes);
  }

  function inflateRawDeflate(bytes) {
    const reader = createBitReader(bytes);
    const output = [];
    let isFinal = false;

    while (!isFinal) {
      isFinal = reader.readBits(1) === 1;
      const blockType = reader.readBits(2);

      if (blockType === 0) {
        inflateStoredBlock(reader, output);
      } else if (blockType === 1) {
        inflateCompressedBlock(reader, output, fixedLiteralTree(), fixedDistanceTree());
      } else if (blockType === 2) {
        const trees = readDynamicTrees(reader);
        inflateCompressedBlock(reader, output, trees.literalTree, trees.distanceTree);
      } else {
        throw new Error("The macro template contains an invalid deflate block.");
      }
    }

    return new Uint8Array(output);
  }

  function createBitReader(bytes) {
    let position = 0;
    let bitBuffer = 0;
    let bitCount = 0;

    return {
      readBits(count) {
        while (bitCount < count) {
          if (position >= bytes.length) throw new Error("Unexpected end of deflate data.");
          bitBuffer |= bytes[position] << bitCount;
          position += 1;
          bitCount += 8;
        }

        const value = bitBuffer & ((1 << count) - 1);
        bitBuffer >>>= count;
        bitCount -= count;
        return value;
      },
      alignByte() {
        bitBuffer = 0;
        bitCount = 0;
      },
      readByte() {
        if (position >= bytes.length) throw new Error("Unexpected end of stored deflate block.");
        return bytes[position++];
      }
    };
  }

  function inflateStoredBlock(reader, output) {
    reader.alignByte();
    const length = reader.readByte() | (reader.readByte() << 8);
    const inverseLength = reader.readByte() | (reader.readByte() << 8);
    if (((length ^ inverseLength) & 0xffff) !== 0xffff) {
      throw new Error("The macro template contains an invalid stored deflate block.");
    }

    for (let index = 0; index < length; index += 1) {
      output.push(reader.readByte());
    }
  }

  function inflateCompressedBlock(reader, output, literalTree, distanceTree) {
    while (true) {
      const symbol = decodeHuffmanSymbol(reader, literalTree);
      if (symbol < 256) {
        output.push(symbol);
      } else if (symbol === 256) {
        return;
      } else if (symbol <= 285) {
        const lengthIndex = symbol - 257;
        let length = DEFLATE_LENGTH_BASE[lengthIndex];
        length += reader.readBits(DEFLATE_LENGTH_EXTRA[lengthIndex]);

        const distanceSymbol = decodeHuffmanSymbol(reader, distanceTree);
        if (distanceSymbol >= DEFLATE_DISTANCE_BASE.length) {
          throw new Error("The macro template contains an invalid deflate distance.");
        }

        let distance = DEFLATE_DISTANCE_BASE[distanceSymbol];
        distance += reader.readBits(DEFLATE_DISTANCE_EXTRA[distanceSymbol]);
        if (distance > output.length) {
          throw new Error("The macro template references data before the start of the deflate output.");
        }

        for (let index = 0; index < length; index += 1) {
          output.push(output[output.length - distance]);
        }
      } else {
        throw new Error("The macro template contains an invalid deflate symbol.");
      }
    }
  }

  function readDynamicTrees(reader) {
    const literalCount = reader.readBits(5) + 257;
    const distanceCount = reader.readBits(5) + 1;
    const codeLengthCount = reader.readBits(4) + 4;
    const codeLengthOrder = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
    const codeLengthLengths = new Array(19).fill(0);

    for (let index = 0; index < codeLengthCount; index += 1) {
      codeLengthLengths[codeLengthOrder[index]] = reader.readBits(3);
    }

    const codeLengthTree = buildHuffmanTree(codeLengthLengths);
    const lengths = [];
    while (lengths.length < literalCount + distanceCount) {
      const symbol = decodeHuffmanSymbol(reader, codeLengthTree);
      if (symbol <= 15) {
        lengths.push(symbol);
      } else if (symbol === 16) {
        if (!lengths.length) throw new Error("The macro template has an invalid repeated deflate length.");
        const repeat = reader.readBits(2) + 3;
        const value = lengths[lengths.length - 1];
        for (let index = 0; index < repeat; index += 1) lengths.push(value);
      } else if (symbol === 17) {
        const repeat = reader.readBits(3) + 3;
        for (let index = 0; index < repeat; index += 1) lengths.push(0);
      } else if (symbol === 18) {
        const repeat = reader.readBits(7) + 11;
        for (let index = 0; index < repeat; index += 1) lengths.push(0);
      } else {
        throw new Error("The macro template has an invalid deflate code-length symbol.");
      }
    }

    return {
      literalTree: buildHuffmanTree(lengths.slice(0, literalCount)),
      distanceTree: buildHuffmanTree(lengths.slice(literalCount, literalCount + distanceCount))
    };
  }

  let cachedFixedLiteralTree = null;
  let cachedFixedDistanceTree = null;

  function fixedLiteralTree() {
    if (!cachedFixedLiteralTree) {
      const lengths = new Array(288).fill(0);
      for (let symbol = 0; symbol <= 143; symbol += 1) lengths[symbol] = 8;
      for (let symbol = 144; symbol <= 255; symbol += 1) lengths[symbol] = 9;
      for (let symbol = 256; symbol <= 279; symbol += 1) lengths[symbol] = 7;
      for (let symbol = 280; symbol <= 287; symbol += 1) lengths[symbol] = 8;
      cachedFixedLiteralTree = buildHuffmanTree(lengths);
    }
    return cachedFixedLiteralTree;
  }

  function fixedDistanceTree() {
    if (!cachedFixedDistanceTree) {
      cachedFixedDistanceTree = buildHuffmanTree(new Array(32).fill(5));
    }
    return cachedFixedDistanceTree;
  }

  function buildHuffmanTree(lengths) {
    const maxBits = Math.max(...lengths);
    const counts = new Array(maxBits + 1).fill(0);
    const nextCodes = new Array(maxBits + 1).fill(0);
    const tables = Array.from({ length: maxBits + 1 }, () => new Map());
    let code = 0;

    for (const length of lengths) {
      if (length > 0) counts[length] += 1;
    }

    for (let bits = 1; bits <= maxBits; bits += 1) {
      code = (code + (counts[bits - 1] || 0)) << 1;
      nextCodes[bits] = code;
    }

    lengths.forEach((length, symbol) => {
      if (!length) return;
      const symbolCode = nextCodes[length];
      nextCodes[length] += 1;
      tables[length].set(reverseBits(symbolCode, length), symbol);
    });

    return { maxBits, tables };
  }

  function reverseBits(value, length) {
    let reversed = 0;
    for (let index = 0; index < length; index += 1) {
      reversed = (reversed << 1) | (value & 1);
      value >>>= 1;
    }
    return reversed;
  }

  function decodeHuffmanSymbol(reader, tree) {
    let code = 0;
    for (let length = 1; length <= tree.maxBits; length += 1) {
      code |= reader.readBits(1) << (length - 1);
      if (tree.tables[length].has(code)) return tree.tables[length].get(code);
    }
    throw new Error("The macro template contains an invalid deflate code.");
  }

  function readUint16(bytes, offset) {
    return bytes[offset] | (bytes[offset + 1] << 8);
  }

  function readUint32(bytes, offset) {
    return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
  }

  function bytesToBase64(bytes) {
    let binary = "";
    const chunkSize = 0x8000;
    for (let index = 0; index < bytes.length; index += chunkSize) {
      binary += String.fromCharCode(...bytes.slice(index, index + chunkSize));
    }
    return btoa(binary);
  }

  function base64ToBytes(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  function contentTypesXml(worksheetCount = 2, drawingCount = 1) {
    const worksheetOverrides = Array.from({ length: worksheetCount }, (_item, index) => (
      `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    )).join("");
    const drawingOverrides = Array.from({ length: drawingCount }, (_item, index) => (
      `<Override PartName="/xl/drawings/drawing${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    )).join("");

    return xmlDecl(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
      ${worksheetOverrides}
      ${drawingOverrides}
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

  function workbookXml(planSheetNames = ["Plan"]) {
    const sheetNames = planSheetNames.concat("Progress");
    const sheets = sheetNames.map((name, index) => (
      `<sheet name="${xmlEscape(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
    )).join("");

    return xmlDecl(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="23426"/>
      <workbookPr codeName="ThisWorkbook"/>
      <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="18000" windowHeight="10000"/></bookViews>
      <sheets>
        ${sheets}
      </sheets>
      <calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>
    </workbook>`);
  }

  function workbookRelsXml(worksheetCount = 2) {
    const worksheetRels = Array.from({ length: worksheetCount }, (_item, index) => (
      `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    )).join("");
    const stylesId = worksheetCount + 1;
    const vbaId = worksheetCount + 2;

    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${worksheetRels}
      <Relationship Id="rId${stylesId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId${vbaId}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
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

  function planSheetXml(macroProject, page = defaultExcelPage(), sheetIndex = 1) {
    const statusMessage = macroProject.live
      ? "Live macro included. Enable macros in Excel, then edits to Progress columns A, B, E, or F update zone labels, colours, and opacity."
      : "Snapshot export. Install the macro template before exporting if this workbook must update zone colours and labels in Excel.";

    return xmlDecl(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr codeName="Sheet${sheetIndex}"/>
      <sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>
      <sheetFormatPr defaultRowHeight="15"/>
      <cols><col min="1" max="12" width="14" customWidth="1"/></cols>
      <sheetData>
        <row r="1" ht="24" customHeight="1">
          ${cell("A1", page.plan.name || page.name || "Floor Plan Progress Export", "inlineStr", 2)}
        </row>
        <row r="2">
          ${cell("A2", statusMessage, "inlineStr", 0)}
        </row>
      </sheetData>
      <pageMargins left="0.25" right="0.25" top="0.25" bottom="0.25" header="0.1" footer="0.1"/>
      <drawing r:id="rId1"/>
    </worksheet>`);
  }

  function defaultExcelPage() {
    return Array.isArray(state.pages) && state.pages.length
      ? activePage()
      : pageFromLegacyProject(state);
  }

  function planSheetRelsXml(drawingIndex = 1) {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/>
    </Relationships>`);
  }

  function progressSheetXml(roomEntries = [], macroProject, sheetIndex = 2) {
    const entries = roomEntries.map((entry) => (
      entry && entry.room
        ? entry
        : { room: entry, sheetName: "Plan", plan: state.plan }
    ));
    const refreshStatus = macroProject.live ? "Live macro included" : "Snapshot only";
    const refreshNote = macroProject.live
      ? "Enable macros, then edit A, B, E, or F."
      : "Install macro template and export again for live updates.";
    const rows = [
      `<row r="1">${cell("A1", "Room ID", "inlineStr", 1)}${cell("B1", "Plan Sheet", "inlineStr", 1)}${cell("C1", "Zone Shape", "inlineStr", 1)}${cell("D1", "Current Colour", "inlineStr", 1)}${cell("E1", "Percent Complete", "inlineStr", 1)}${cell("F1", "Overlay Opacity", "inlineStr", 1)}${cell("G1", "Points", "inlineStr", 1)}${cell("H1", "Label Shape", "inlineStr", 1)}${cell("I1", "Excel Refresh", "inlineStr", 1)}${cell("J1", "Macro Source", "inlineStr", 1)}${cell("K1", "Note", "inlineStr", 1)}</row>`
    ];

    entries.forEach((entry, index) => {
      const room = entry.room;
      const row = index + 2;
      rows.push(`<row r="${row}">
        ${cell(`A${row}`, room.id, "inlineStr", 4)}
        ${cell(`B${row}`, entry.sheetName || "Plan", "inlineStr", 4)}
        ${cell(`C${row}`, excelShapeName(room), "inlineStr", 4)}
        ${formulaTextCell(`D${row}`, excelColourFormula(row), rgbToHex(progressColor(room.percent)), 4)}
        ${cell(`E${row}`, room.percent, "n", 4)}
        ${cell(`F${row}`, state.settings.opacity, "n", 4)}
        ${cell(`G${row}`, JSON.stringify(room.points), "inlineStr", 4)}
        ${cell(`H${row}`, excelLabelShapeName(room), "inlineStr", 4)}
        ${index === 0 ? cell(`I${row}`, refreshStatus, "inlineStr", 4) : ""}
        ${index === 0 ? cell(`J${row}`, macroProject.sourceName, "inlineStr", 4) : ""}
        ${index === 0 ? cell(`K${row}`, refreshNote, "inlineStr", 4) : ""}
      </row>`);
    });

    if (!entries.length) {
      rows.push(`<row r="2">
        ${cell("I2", refreshStatus, "inlineStr", 4)}
        ${cell("J2", macroProject.sourceName, "inlineStr", 4)}
        ${cell("K2", refreshNote, "inlineStr", 4)}
      </row>`);
    }

    return xmlDecl(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr codeName="Sheet${sheetIndex}"/>
      <sheetViews><sheetView workbookViewId="0"/></sheetViews>
      <sheetFormatPr defaultRowHeight="18"/>
      <cols>
        <col min="1" max="1" width="18" customWidth="1"/>
        <col min="2" max="2" width="16" customWidth="1"/>
        <col min="3" max="3" width="26" customWidth="1"/>
        <col min="4" max="4" width="16" customWidth="1"/>
        <col min="5" max="5" width="18" customWidth="1"/>
        <col min="6" max="6" width="16" customWidth="1"/>
        <col min="7" max="7" width="80" customWidth="1"/>
        <col min="8" max="8" width="26" customWidth="1"/>
        <col min="9" max="9" width="22" customWidth="1"/>
        <col min="10" max="10" width="26" customWidth="1"/>
        <col min="11" max="11" width="44" customWidth="1"/>
      </cols>
      <sheetData>${rows.join("")}</sheetData>
      <autoFilter ref="A1:K${Math.max(1, entries.length + 1)}"/>
      <dataValidations count="2">
        <dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="E2:E${Math.max(2, entries.length + 1)}"><formula1>0</formula1><formula2>100</formula2></dataValidation>
        <dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="F2:F${Math.max(2, entries.length + 1)}"><formula1>0</formula1><formula2>100</formula2></dataValidation>
      </dataValidations>
      <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
    </worksheet>`);
  }

  function drawingXml(rooms, imageWidth, imageHeight) {
    const imageExtent = { width: imageWidth, height: imageHeight };
    const mappedRooms = rooms.filter((room) => room.points.length >= 3);
    const image = absoluteAnchor(0, EXCEL_PLAN_TOP_OFFSET, imageExtent.width, imageExtent.height, pictureXml(2, "Floor plan image", "rId1", imageExtent.width, imageExtent.height));
    const shapes = mappedRooms
      .map((room, index) => polygonShapeAnchor(room, index + 3))
      .join("");
    const labels = mappedRooms
      .map((room, index) => roomLabelAnchor(room, index + 3 + mappedRooms.length))
      .join("");

    return xmlDecl(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      ${image}
      ${shapes}
      ${labels}
    </xdr:wsDr>`);
  }

  function drawingRelsXml(mediaFileName = "floor-plan.png") {
    return xmlDecl(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${xmlEscape(mediaFileName)}"/>
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
    </xdr:sp>`;

    return absoluteAnchor(bounds.minX, EXCEL_PLAN_TOP_OFFSET + bounds.minY, width, height, shape);
  }

  function roomLabelAnchor(room, id) {
    const labelBox = roomLabelBox(room.points, room.id);
    const width = Math.max(1, labelBox.width);
    const height = Math.max(1, labelBox.height);
    const percentText = `${room.percent}%`;
    const percentFont = Math.round(excelLabelFontSize(percentText, width, height, 19) * 100);
    const idFont = Math.round(Math.min(
      excelLabelFontSize(room.id, width, height, 14),
      (percentFont / 100) * 0.78
    ) * 100);
    const shape = `<xdr:sp macro="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="${id}" name="${xmlEscape(excelLabelShapeName(room))}" descr="${xmlEscape(`${room.id} label`)}"/>
        <xdr:cNvSpPr txBox="1"/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr wrap="square" anchor="ctr" anchorCtr="1" lIns="0" tIns="0" rIns="0" bIns="0" vertOverflow="clip" horzOverflow="clip"><a:normAutofit/></a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:rPr lang="en-US" sz="${percentFont}" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>${xmlEscape(percentText)}</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:rPr lang="en-US" sz="${idFont}" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>${xmlEscape(room.id)}</a:t></a:r>
        </a:p>
      </xdr:txBody>
    </xdr:sp>`;

    return absoluteAnchor(labelBox.x, EXCEL_PLAN_TOP_OFFSET + labelBox.y, width, height, shape);
  }

  function roomLabelBox(points, roomId = "") {
    const bounds = pointBounds(points);
    const fallback = {
      x: bounds.minX,
      y: bounds.minY,
      width: Math.max(1, bounds.maxX - bounds.minX),
      height: Math.max(1, bounds.maxY - bounds.minY)
    };

    if (!Array.isArray(points) || points.length < 3 || !Number.isFinite(fallback.x) || !Number.isFinite(fallback.y)) {
      return fallback;
    }

    const textRatio = clamp((Math.max(4, String(roomId || "").length) * 8 + 26) / 34, 0.85, 3.8);
    const ratios = uniqueNumbers([textRatio, 2.2, 1.55, 1.05, 0.78]);
    const candidates = interiorLabelCandidates(points, bounds);
    let best = null;

    for (const center of candidates) {
      for (const ratio of ratios) {
        const box = largestCenteredLabelBox(points, bounds, center, ratio);
        if (box && (!best || box.width * box.height > best.width * best.height)) {
          best = box;
        }
      }
    }

    if (!best) return fallback;

    const padding = Math.min(EXCEL_LABEL_BOX_PADDING, Math.max(0, Math.min(best.width, best.height) / 8));
    return {
      x: round(best.x + padding),
      y: round(best.y + padding),
      width: round(Math.max(1, best.width - padding * 2)),
      height: round(Math.max(1, best.height - padding * 2))
    };
  }

  function interiorLabelCandidates(points, bounds) {
    const candidates = [];
    const addCandidate = (point) => {
      if (!point || !pointInPolygon(point, points)) return;
      const key = `${Math.round(point[0] * 10)}|${Math.round(point[1] * 10)}`;
      if (candidates.some((candidate) => `${Math.round(candidate[0] * 10)}|${Math.round(candidate[1] * 10)}` === key)) return;
      candidates.push(point);
    };

    addCandidate(polygonCentroid(points));
    addCandidate([(bounds.minX + bounds.maxX) / 2, (bounds.minY + bounds.maxY) / 2]);

    const steps = 8;
    for (let yIndex = 1; yIndex < steps; yIndex += 1) {
      for (let xIndex = 1; xIndex < steps; xIndex += 1) {
        addCandidate([
          bounds.minX + ((bounds.maxX - bounds.minX) * xIndex) / steps,
          bounds.minY + ((bounds.maxY - bounds.minY) * yIndex) / steps
        ]);
      }
    }

    return candidates;
  }

  function largestCenteredLabelBox(points, bounds, center, ratio) {
    const maxHalfWidth = Math.min(center[0] - bounds.minX, bounds.maxX - center[0]);
    const maxHalfHeight = Math.min(center[1] - bounds.minY, bounds.maxY - center[1]);
    let low = 0;
    let high = Math.min(maxHalfHeight, maxHalfWidth / ratio);
    if (high <= 0) return null;

    for (let index = 0; index < 18; index += 1) {
      const halfHeight = (low + high) / 2;
      const halfWidth = halfHeight * ratio;
      const box = {
        x: center[0] - halfWidth,
        y: center[1] - halfHeight,
        width: halfWidth * 2,
        height: halfHeight * 2
      };

      if (rectangleInsidePolygon(box, points)) {
        low = halfHeight;
      } else {
        high = halfHeight;
      }
    }

    const height = low * 2;
    const width = height * ratio;
    if (width < 4 || height < 4) return null;

    return {
      x: center[0] - width / 2,
      y: center[1] - height / 2,
      width,
      height
    };
  }

  function rectangleInsidePolygon(box, points) {
    const x1 = box.x;
    const y1 = box.y;
    const x2 = box.x + box.width;
    const y2 = box.y + box.height;
    const samples = [
      [x1, y1], [x2, y1], [x2, y2], [x1, y2],
      [(x1 + x2) / 2, y1], [x2, (y1 + y2) / 2],
      [(x1 + x2) / 2, y2], [x1, (y1 + y2) / 2],
      [(x1 + x2) / 2, (y1 + y2) / 2]
    ];

    if (samples.some((point) => !pointInPolygon(point, points))) return false;

    const rectEdges = [
      [[x1, y1], [x2, y1]],
      [[x2, y1], [x2, y2]],
      [[x2, y2], [x1, y2]],
      [[x1, y2], [x1, y1]]
    ];

    for (const rectEdge of rectEdges) {
      for (let index = 0; index < points.length; index += 1) {
        const a = points[index];
        const b = points[(index + 1) % points.length];
        if (segmentsIntersect(rectEdge[0], rectEdge[1], a, b)) return false;
      }
    }

    return true;
  }

  function pointInPolygon(point, polygon) {
    const x = point[0];
    const y = point[1];
    let inside = false;

    for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
      const xi = polygon[i][0];
      const yi = polygon[i][1];
      const xj = polygon[j][0];
      const yj = polygon[j][1];

      if (pointOnSegment(point, polygon[j], polygon[i])) return true;
      const intersects = ((yi > y) !== (yj > y)) && (x < ((xj - xi) * (y - yi)) / (yj - yi) + xi);
      if (intersects) inside = !inside;
    }

    return inside;
  }

  function segmentsIntersect(a, b, c, d) {
    const o1 = orientation(a, b, c);
    const o2 = orientation(a, b, d);
    const o3 = orientation(c, d, a);
    const o4 = orientation(c, d, b);

    if (o1 !== o2 && o3 !== o4) return true;
    if (o1 === 0 && pointOnSegment(c, a, b)) return true;
    if (o2 === 0 && pointOnSegment(d, a, b)) return true;
    if (o3 === 0 && pointOnSegment(a, c, d)) return true;
    if (o4 === 0 && pointOnSegment(b, c, d)) return true;
    return false;
  }

  function orientation(a, b, c) {
    const value = ((b[1] - a[1]) * (c[0] - b[0])) - ((b[0] - a[0]) * (c[1] - b[1]));
    if (Math.abs(value) < 0.0001) return 0;
    return value > 0 ? 1 : 2;
  }

  function pointOnSegment(point, start, end) {
    const cross = ((point[1] - start[1]) * (end[0] - start[0])) - ((point[0] - start[0]) * (end[1] - start[1]));
    if (Math.abs(cross) > 0.0001) return false;
    return point[0] >= Math.min(start[0], end[0]) - 0.0001
      && point[0] <= Math.max(start[0], end[0]) + 0.0001
      && point[1] >= Math.min(start[1], end[1]) - 0.0001
      && point[1] <= Math.max(start[1], end[1]) + 0.0001;
  }

  function excelLabelFontSize(text, width, height, maxSize) {
    const length = Math.max(1, String(text || "").length);
    const widthFit = width / (length * 0.68);
    const heightFit = height / 3.1;
    return clamp(Math.min(widthFit, heightFit), 5, maxSize);
  }

  function uniqueNumbers(values) {
    const seen = new Set();
    return values.filter((value) => {
      const key = Math.round(value * 100);
      if (seen.has(key)) return false;
      seen.add(key);
      return Number.isFinite(value) && value > 0;
    });
  }

  function absoluteAnchor(x, y, width, height, body) {
    return `<xdr:absoluteAnchor>
      <xdr:pos x="${pxToEmu(x)}" y="${pxToEmu(y)}"/>
      <xdr:ext cx="${pxToEmu(width)}" cy="${pxToEmu(height)}"/>
      ${body}
      <xdr:clientData/>
    </xdr:absoluteAnchor>`;
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

  function appPropsXml(sheetNames = ["Plan", "Progress"]) {
    const sheetList = sheetNames.map((name) => `<vt:lpstr>${xmlEscape(name)}</vt:lpstr>`).join("");
    return xmlDecl(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
      <Application>Microsoft Excel</Application>
      <DocSecurity>0</DocSecurity>
      <ScaleCrop>false</ScaleCrop>
      <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>
      <TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">${sheetList}</vt:vector></TitlesOfParts>
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

  function formulaTextCell(ref, formula, cachedValue, style) {
    const styleAttr = style ? ` s="${style}"` : "";
    return `<c r="${ref}" t="str"${styleAttr}><f>${xmlEscape(formula)}</f><v>${xmlEscape(cachedValue)}</v></c>`;
  }

  function excelColourFormula(row, percentColumn = "E") {
    const p = `MAX(0,MIN(100,${percentColumn}${row}))`;
    const lowerAmount = `(${p}/50)`;
    const upperAmount = `((${p}-50)/50)`;
    const lower = `"#"&DEC2HEX(ROUND(216+(217-216)*${lowerAmount},0),2)&DEC2HEX(ROUND(66+(163-66)*${lowerAmount},0),2)&DEC2HEX(ROUND(47+(33-47)*${lowerAmount},0),2)`;
    const upper = `"#"&DEC2HEX(ROUND(217+(8-217)*${upperAmount},0),2)&DEC2HEX(ROUND(163+(88-163)*${upperAmount},0),2)&DEC2HEX(ROUND(33+(43-33)*${upperAmount},0),2)`;
    return `IF(${p}<=50,${lower},${upper})`;
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

  function excelPlanSheetName(_page, index, pageCount) {
    return pageCount > 1 ? `Plan ${index + 1}` : "Plan";
  }

  function excelShapeName(room) {
    return `Zone_${room.id.replace(/[^A-Za-z0-9_]/g, "_")}`;
  }

  function excelLabelShapeName(room) {
    return `Label_${room.id.replace(/[^A-Za-z0-9_]/g, "_")}`;
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
      const baseName = file.name.replace(/\.[^.]+$/, "");
      const pages = [];
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
        el.saveStatus.textContent = `Loading PDF page ${pageNumber} of ${pdf.numPages}...`;
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

        let zlLabels = [];
        try {
          const textContent = await page.getTextContent();
          zlLabels = extractZlLabelsFromTextContent(textContent, viewport, scale, canvas.width, canvas.height);
        } catch (error) {
          console.warn(`Could not read PDF text labels on page ${pageNumber}.`, error);
        }

        const pageName = pdf.numPages > 1 ? `${baseName} - Page ${pageNumber}` : baseName;
        pages.push({
          id: `page-${pageNumber}`,
          name: pageName,
          plan: {
            name: pageName,
            src: canvas.toDataURL("image/png"),
            sourceType: "pdf",
            originalFileName: file.name,
            pageNumber,
            pageCount: pdf.numPages,
            width: canvas.width,
            height: canvas.height,
            zlLabels
          },
          rooms: []
        });
      }

      return {
        name: baseName,
        pages
      };
    } finally {
      if (typeof pdf.destroy === "function") {
        await pdf.destroy();
      }
    }
  }

  function extractZlLabelsFromTextContent(textContent, viewport, scale, planWidth, planHeight) {
    const util = window.pdfjsLib && window.pdfjsLib.Util;
    if (!textContent || !Array.isArray(textContent.items) || !util || typeof util.transform !== "function") {
      return [];
    }

    const runs = textContent.items
      .map((item) => {
        const rawText = String(item.str || "").trim();
        if (!rawText) return null;
        const transform = util.transform(viewport.transform, item.transform);
        const fontHeight = Math.max(7, Math.abs(item.height || transform[3] || 0) * scale);
        const width = Math.max(rawText.length * fontHeight * 0.42, Math.abs(item.width || 0) * scale);
        const x = transform[4];
        const y = transform[5] - fontHeight;
        return {
          text: rawText,
          x,
          y,
          width,
          height: fontHeight,
          centerX: x + width / 2,
          centerY: y + fontHeight / 2
        };
      })
      .filter(Boolean);

    return zlLabelsFromTextRuns(runs, planWidth, planHeight);
  }

  function zlLabelsFromTextRuns(runs, planWidth, planHeight) {
    const lines = [];
    const sortedRuns = runs
      .filter((run) => run && String(run.text || "").trim())
      .sort((a, b) => (a.centerY - b.centerY) || (a.x - b.x));

    for (const run of sortedRuns) {
      const tolerance = Math.max(8, run.height * 0.85);
      let line = lines.find((item) => Math.abs(item.centerY - run.centerY) <= tolerance);
      if (!line) {
        line = { centerY: run.centerY, runs: [] };
        lines.push(line);
      }
      line.runs.push(run);
      line.centerY = line.runs.reduce((sum, item) => sum + item.centerY, 0) / line.runs.length;
    }

    const labels = [];
    const seen = new Set();
    for (const line of lines) {
      const lineRuns = line.runs.slice().sort((a, b) => a.x - b.x);
      let cluster = [];
      for (const run of lineRuns) {
        const previous = cluster[cluster.length - 1];
        const gap = previous ? run.x - (previous.x + previous.width) : 0;
        const allowedGap = Math.max(28, run.height * 4);
        if (previous && gap > allowedGap) {
          pushZlCluster(labels, seen, cluster, planWidth, planHeight);
          cluster = [];
        }
        cluster.push(run);
      }
      pushZlCluster(labels, seen, cluster, planWidth, planHeight);
    }

    return labels;
  }

  function zlLabelsFromDetectedText(detections, planWidth, planHeight) {
    if (!Array.isArray(detections)) return [];
    const runs = detections
      .map((detection) => detectedTextToRun(detection))
      .filter(Boolean);
    return zlLabelsFromTextRuns(runs, planWidth, planHeight);
  }

  function detectedTextToRun(detection) {
    const text = String(detection && (detection.rawValue || detection.text || detection.value || "")).trim();
    if (!text) return null;

    const box = detection.boundingBox || detection.bounds || {};
    const cornerPoints = Array.isArray(detection.cornerPoints) ? detection.cornerPoints : [];
    const pointBoundsValue = cornerPoints.length
      ? cornerPoints.reduce((bounds, point) => ({
          minX: Math.min(bounds.minX, Number(point.x)),
          minY: Math.min(bounds.minY, Number(point.y)),
          maxX: Math.max(bounds.maxX, Number(point.x)),
          maxY: Math.max(bounds.maxY, Number(point.y))
        }), { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity })
      : null;

    const x = Number.isFinite(Number(box.x)) ? Number(box.x) : pointBoundsValue?.minX;
    const y = Number.isFinite(Number(box.y)) ? Number(box.y) : pointBoundsValue?.minY;
    const width = Number.isFinite(Number(box.width)) ? Number(box.width) : (pointBoundsValue ? pointBoundsValue.maxX - pointBoundsValue.minX : NaN);
    const height = Number.isFinite(Number(box.height)) ? Number(box.height) : (pointBoundsValue ? pointBoundsValue.maxY - pointBoundsValue.minY : NaN);
    if (![x, y, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;

    return {
      text,
      x,
      y,
      width,
      height,
      centerX: x + width / 2,
      centerY: y + height / 2
    };
  }

  function mergeDetectedZlLabels(existingLabels, newLabels) {
    const merged = (existingLabels || [])
      .map((label) => normalizeDetectedZlLabel(label))
      .filter(Boolean);

    for (const newLabel of newLabels || []) {
      const normalized = normalizeDetectedZlLabel(newLabel);
      if (!normalized) continue;

      const match = merged.find((label) => (
        label.id.toLowerCase() === normalized.id.toLowerCase()
        && Math.abs(label.x - normalized.x) <= 12
        && Math.abs(label.y - normalized.y) <= 12
      ));
      if (match) {
        Object.assign(match, normalized);
      } else {
        merged.push(normalized);
      }
    }

    return merged;
  }

  function pushZlCluster(labels, seen, cluster, planWidth, planHeight) {
    if (!cluster.length) return;
    const text = cluster.map((run) => run.text).join(" ");
    const id = extractZlCode(text);
    if (!id) return;

    const bounds = cluster.reduce((box, run) => ({
      minX: Math.min(box.minX, run.x),
      minY: Math.min(box.minY, run.y),
      maxX: Math.max(box.maxX, run.x + run.width),
      maxY: Math.max(box.maxY, run.y + run.height)
    }), { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity });
    const label = normalizeDetectedZlLabel({
      id,
      x: (bounds.minX + bounds.maxX) / 2,
      y: (bounds.minY + bounds.maxY) / 2,
      width: bounds.maxX - bounds.minX,
      height: bounds.maxY - bounds.minY
    }, planWidth, planHeight);
    if (!label) return;

    const key = `${label.id.toLowerCase()}|${Math.round(label.x / 8)}|${Math.round(label.y / 8)}`;
    if (seen.has(key)) return;
    seen.add(key);
    labels.push(label);
  }

  function normalizeDetectedZlLabel(label, planWidth = state.plan.width, planHeight = state.plan.height) {
    const id = extractZlCode(label.id);
    const x = Number(label.x);
    const y = Number(label.y);
    if (!id || !Number.isFinite(x) || !Number.isFinite(y)) return null;

    return {
      id,
      x: round(clamp(x, 0, Number(planWidth) || 1)),
      y: round(clamp(y, 0, Number(planHeight) || 1)),
      width: round(Math.max(1, Number(label.width) || 1)),
      height: round(Math.max(1, Number(label.height) || 1))
    };
  }

  function extractZlCode(text) {
    const value = String(text || "").toUpperCase().replace(/\s+/g, " ").trim();
    const match = value.match(/\b([A-Z0-9]*ZL[\s_-]*[A-Z0-9_-]*)\b/);
    if (!match) return "";
    const code = match[1].replace(/\s+/g, "");
    return code === "ZL" ? "" : code;
  }

  function findBoxBoundaryForLabel(imageData, width, height, label) {
    const map = buildDarkPixelMap(imageData, width, height);
    const cx = clamp(Math.round(label.x), 0, width - 1);
    const cy = clamp(Math.round(label.y), 0, height - 1);
    const minDistance = Math.max(AUTO_DETECT_WALL.minLabelDistance, Math.round(Math.max(label.width || 0, label.height || 0) * 0.85));
    const maxXDistance = Math.max(minDistance + 8, Math.round(width * AUTO_DETECT_WALL.maxSpanRatio));
    const maxYDistance = Math.max(minDistance + 8, Math.round(height * AUTO_DETECT_WALL.maxSpanRatio));

    let left = findVerticalWall(map, width, height, cx, cy, -1, minDistance, maxXDistance);
    let right = findVerticalWall(map, width, height, cx, cy, 1, minDistance, maxXDistance);
    let top = null;
    let bottom = null;

    if (left && right) {
      top = findHorizontalWall(map, width, height, cx, cy, -1, minDistance, maxYDistance, left.position, right.position);
      bottom = findHorizontalWall(map, width, height, cx, cy, 1, minDistance, maxYDistance, left.position, right.position);
    }

    if ((!top || !bottom) && (!left || !right)) {
      top = top || findHorizontalWall(map, width, height, cx, cy, -1, minDistance, maxYDistance);
      bottom = bottom || findHorizontalWall(map, width, height, cx, cy, 1, minDistance, maxYDistance);
      if (top && bottom) {
        left = left || findVerticalWall(map, width, height, cx, cy, -1, minDistance, maxXDistance, top.position, bottom.position);
        right = right || findVerticalWall(map, width, height, cx, cy, 1, minDistance, maxXDistance, top.position, bottom.position);
      }
    }

    if (!left || !right || !top || !bottom) return null;
    if (right.position - left.position < AUTO_DETECT_WALL.minRoomSize) return null;
    if (bottom.position - top.position < AUTO_DETECT_WALL.minRoomSize) return null;
    if (cx <= left.position || cx >= right.position || cy <= top.position || cy >= bottom.position) return null;

    return {
      confidence: round((left.score + right.score + top.score + bottom.score) / 4),
      points: [
        [round(left.position), round(top.position)],
        [round(right.position), round(top.position)],
        [round(right.position), round(bottom.position)],
        [round(left.position), round(bottom.position)]
      ]
    };
  }

  function buildDarkPixelMap(imageData, width, height) {
    const source = imageData.data || imageData;
    const map = new Uint8Array(width * height);
    for (let index = 0, pixel = 0; index < source.length; index += 4, pixel += 1) {
      const alpha = source[index + 3];
      if (alpha < 30) continue;
      const luminance = (source[index] * 0.2126) + (source[index + 1] * 0.7152) + (source[index + 2] * 0.0722);
      map[pixel] = luminance <= AUTO_DETECT_WALL.darkThreshold ? 1 : 0;
    }
    return map;
  }

  function findVerticalWall(map, width, height, cx, cy, direction, minDistance, maxDistance, topLimit, bottomLimit) {
    const bandHalf = Math.min(AUTO_DETECT_WALL.bandSize / 2, Math.max(50, height * 0.08));
    const y1 = topLimit == null ? cy - bandHalf : topLimit + 6;
    const y2 = bottomLimit == null ? cy + bandHalf : bottomLimit - 6;

    for (let distance = minDistance; distance <= maxDistance; distance += AUTO_DETECT_WALL.scanStep) {
      const x = cx + direction * distance;
      if (x <= 1 || x >= width - 2) break;
      const score = verticalWallScore(map, width, height, x, y1, y2);
      if (score >= 1) return { position: x, score };
    }
    return null;
  }

  function findHorizontalWall(map, width, height, cx, cy, direction, minDistance, maxDistance, leftLimit, rightLimit) {
    const bandHalf = Math.min(AUTO_DETECT_WALL.bandSize / 2, Math.max(50, width * 0.08));
    const x1 = leftLimit == null ? cx - bandHalf : leftLimit + 6;
    const x2 = rightLimit == null ? cx + bandHalf : rightLimit - 6;

    for (let distance = minDistance; distance <= maxDistance; distance += AUTO_DETECT_WALL.scanStep) {
      const y = cy + direction * distance;
      if (y <= 1 || y >= height - 2) break;
      const score = horizontalWallScore(map, width, height, y, x1, x2);
      if (score >= 1) return { position: y, score };
    }
    return null;
  }

  function verticalWallScore(map, width, height, x, y1, y2) {
    const half = Math.floor(AUTO_DETECT_WALL.lineThickness / 2);
    const startY = Math.max(0, Math.round(Math.min(y1, y2)));
    const endY = Math.min(height - 1, Math.round(Math.max(y1, y2)));
    let rowsWithDark = 0;
    let darkPixels = 0;
    let totalPixels = 0;

    for (let y = startY; y <= endY; y += 1) {
      let rowHasDark = false;
      for (let offset = -half; offset <= half; offset += 1) {
        const px = Math.round(x + offset);
        if (px < 0 || px >= width) continue;
        totalPixels += 1;
        if (map[(y * width) + px]) {
          darkPixels += 1;
          rowHasDark = true;
        }
      }
      if (rowHasDark) rowsWithDark += 1;
    }

    return wallScore(rowsWithDark, endY - startY + 1, darkPixels, totalPixels);
  }

  function horizontalWallScore(map, width, height, y, x1, x2) {
    const half = Math.floor(AUTO_DETECT_WALL.lineThickness / 2);
    const startX = Math.max(0, Math.round(Math.min(x1, x2)));
    const endX = Math.min(width - 1, Math.round(Math.max(x1, x2)));
    let columnsWithDark = 0;
    let darkPixels = 0;
    let totalPixels = 0;

    for (let x = startX; x <= endX; x += 1) {
      let columnHasDark = false;
      for (let offset = -half; offset <= half; offset += 1) {
        const py = Math.round(y + offset);
        if (py < 0 || py >= height) continue;
        totalPixels += 1;
        if (map[(py * width) + x]) {
          darkPixels += 1;
          columnHasDark = true;
        }
      }
      if (columnHasDark) columnsWithDark += 1;
    }

    return wallScore(columnsWithDark, endX - startX + 1, darkPixels, totalPixels);
  }

  function wallScore(linesWithDark, lineCount, darkPixels, totalPixels) {
    if (!lineCount || !totalPixels) return 0;
    const coverage = linesWithDark / lineCount;
    const ratio = darkPixels / totalPixels;
    if (coverage < AUTO_DETECT_WALL.minCoverage || ratio < AUTO_DETECT_WALL.minDarkRatio) return 0;
    return Math.min(1.5, (coverage / AUTO_DETECT_WALL.minCoverage + ratio / AUTO_DETECT_WALL.minDarkRatio) / 2);
  }

  function pdfRenderScale(width, height) {
    const longestEdge = Math.max(width, height) || PDF_RENDER_LONG_EDGE;
    return PDF_RENDER_LONG_EDGE / longestEdge;
  }

  function applyLoadedPlan(plan) {
    const previousPages = (state.pages || []).map((page) => ({
      rooms: clone(page.rooms || [])
    }));
    const keepRooms = countProjectRooms(state) && confirm("Keep existing room areas on this new floor plan?");
    const loadedPages = Array.isArray(plan.pages) && plan.pages.length
      ? plan.pages
      : [{
        id: "page-1",
        name: plan.name,
        plan,
        rooms: []
      }];

    state.name = plan.name;
    state.pages = loadedPages.map((page, index) => ({
      id: page.id || `page-${index + 1}`,
      name: page.name || page.plan?.name || `Page ${index + 1}`,
      plan: page.plan || page,
      rooms: keepRooms ? clone(previousPages[index]?.rooms || []) : []
    }));
    state.activePageIndex = 0;
    normalizeProject(state);

    selectedRoomId = null;
    draftPoints = [];
    draggingVertex = null;
    mode = "select";

    loadPlanImage(state.plan.src);
    queueSave();
    render();
  }

  function countProjectRooms(project = state) {
    if (Array.isArray(project.pages) && project.pages.length) {
      return project.pages.reduce((sum, page) => sum + (Array.isArray(page.rooms) ? page.rooms.length : 0), 0);
    }
    return Array.isArray(project.rooms) ? project.rooms.length : 0;
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

  function uniqueRoomId(rawId, currentId, rooms = state.rooms) {
    const base = rawId || "ROOM";
    if (base === currentId) return base;
    const existing = new Set(rooms.filter((room) => room.id !== currentId).map((room) => room.id.toLowerCase()));
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
    const pageIndex = hasHeader
      ? optionalHeaderIndex(header, ["page", "page_index", "page number", "tab", "plan_page"])
      : -1;
    const pageNameIndex = hasHeader
      ? optionalHeaderIndex(header, ["page_name", "page name", "plan", "plan_sheet", "sheet", "sheet_name"])
      : -1;
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
        pageIndex: pageIndex >= 0 ? parsePageIndex(row[pageIndex]) : null,
        pageName: pageNameIndex >= 0 ? String(row[pageNameIndex] || "").trim() : "",
        points: pointsIndex >= 0 ? parseCsvPoints(row[pointsIndex]) : [],
        planWidth: planWidthIndex >= 0 ? Number(row[planWidthIndex]) : null,
        planHeight: planHeightIndex >= 0 ? Number(row[planHeightIndex]) : null
      }))
      .filter((row) => row.id);
  }

  function parsePageIndex(value) {
    const text = String(value || "").trim();
    const match = text.match(/\d+/);
    if (!match) return null;
    const pageNumber = Number(match[0]);
    if (!Number.isFinite(pageNumber) || pageNumber < 1) return null;
    return pageNumber - 1;
  }

  function resolveCsvUpdatePage(update) {
    if (Number.isInteger(update.pageIndex) && state.pages[update.pageIndex]) {
      return state.pages[update.pageIndex];
    }

    const normalizedName = String(update.pageName || "").trim().toLowerCase();
    if (normalizedName) {
      const match = state.pages.find((page, index) => {
        const names = [
          page.name,
          page.plan.name,
          excelPlanSheetName(page, index, state.pages.length),
          `page ${index + 1}`
        ].filter(Boolean).map((name) => String(name).toLowerCase());
        return names.includes(normalizedName);
      });
      if (match) return match;
    }

    return activePage();
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

    if (/^[\[{]/.test(text)) {
      try {
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) return parsed;
      } catch (error) {
        console.warn("CSV points looked like JSON but could not be parsed.", error);
      }
    }

    return text
      .split(";")
      .map((pair) => pair.trim().split(/[|\s:]+/).map(Number))
      .filter((pair) => pair.length >= 2 && pair.every(Number.isFinite))
      .map((pair) => [pair[0], pair[1]]);
  }

  function scaleImportedPoints(points, sourceWidth, sourceHeight, plan = state.plan) {
    if (!Array.isArray(points)) return [];
    const scaleX = sourceWidth && plan.width ? plan.width / sourceWidth : 1;
    const scaleY = sourceHeight && plan.height ? plan.height / sourceHeight : 1;

    return points
      .map((point) => {
        if (!Array.isArray(point) || point.length < 2) return null;
        const x = Number(point[0]);
        const y = Number(point[1]);
        if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
        return [
          round(clamp(x * scaleX, 0, plan.width)),
          round(clamp(y * scaleY, 0, plan.height))
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

  function downloadLinkedFile(fileName, href) {
    const link = document.createElement("a");
    link.href = href;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
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
