/* global Office, PowerPoint, document */

// Percent 0-100 maps to normalized Office.js adjustment value 0-0.5,
// which covers the visually-useful range for a rounded rectangle
// (0 = square corners, 0.5 = fully round / capped at half the shorter side).
// The user can flip "raw mode" to send a value through unchanged.
const PERCENT_TO_RAW = 0.005; // 100% -> 0.5
const POLL_INTERVAL_MS = 500; // Check selection every 500ms

const els = {};
let pollTimer = null;
let lastSelectedShapeId = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("This add-in only runs in PowerPoint.", "err");
    return;
  }
  els.input = document.getElementById("radiusInput");
  els.autoRead = document.getElementById("autoRead");
  els.modePercent = document.getElementById("modePercent");
  els.modeRaw = document.getElementById("modeRaw");
  els.apply = document.getElementById("applyBtn");
  els.read = document.getElementById("readBtn");
  els.status = document.getElementById("status");

  els.apply.addEventListener("click", applyToSelected);
  els.read.addEventListener("click", readFromSelected);
  els.modePercent.addEventListener("change", onModeToggle);
  els.modeRaw.addEventListener("change", onModeToggle);
  els.autoRead.addEventListener("change", onAutoReadToggle);
  onModeToggle();
  onAutoReadToggle();
});

function onModeToggle() {
  if (els.modeRaw.checked) {
    els.input.min = "0";
    els.input.max = "1";
    els.input.step = "0.01";
    document.querySelector('label[for="radiusInput"]').textContent = "Adjustment (0-1)";
  } else {
    els.input.min = "0";
    els.input.max = "100";
    els.input.step = "1";
    document.querySelector('label[for="radiusInput"]').textContent = "Radius (%)";
  }
}

function onAutoReadToggle() {
  if (els.autoRead.checked) {
    startPolling();
  } else {
    stopPolling();
  }
}

function startPolling() {
  if (pollTimer) return; // Already running
  lastSelectedShapeId = null;
  pollTimer = setInterval(pollSelection, POLL_INTERVAL_MS);
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
}

async function pollSelection() {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id,items/name");
      await context.sync();

      // No selection or multiple shapes — clear input if changed
      if (shapes.items.length !== 1) {
        if (lastSelectedShapeId !== null) {
          lastSelectedShapeId = null;
          setStatus("Select a single shape to auto-read its radius.", "");
        }
        return;
      }

      const shape = shapes.items[0];

      // Same shape still selected — skip re-read
      if (shape.id === lastSelectedShapeId) {
        return;
      }

      lastSelectedShapeId = shape.id;

      // Check if shape has adjustments
      shape.adjustments.load("count");
      await context.sync();

      if (!shape.adjustments || shape.adjustments.count < 1) {
        setStatus(`"${shape.name}" has no adjustment handles.`, "");
        return;
      }

      // Read the adjustment value
      const result = shape.adjustments.get(0);
      await context.sync();

      const raw = result.value;
      if (typeof raw !== "number") {
        setStatus(`Could not read adjustment from "${shape.name}".`, "");
        return;
      }

      const pct = raw / PERCENT_TO_RAW;
      els.input.value = els.modeRaw.checked ? raw.toFixed(4) : pct.toFixed(1);
      setStatus(`Auto-read from "${shape.name}": ${raw.toFixed(4)} (${pct.toFixed(1)}%).`, "ok");
    });
  } catch (e) {
    // Silently fail during polling — don't spam errors if PowerPoint is busy
    if (e && e.code !== "GeneralException") {
      console.error("Auto-read poll error:", e);
    }
  }
}

function getTargetValue() {
  const n = parseFloat(els.input.value);
  if (Number.isNaN(n)) throw new Error("Enter a number first.");
  if (els.modeRaw.checked) {
    if (n < 0 || n > 1) throw new Error("Raw value must be between 0 and 1.");
    return n;
  }
  if (n < 0 || n > 100) throw new Error("Percent must be between 0 and 100.");
  return n * PERCENT_TO_RAW;
}

async function applyToSelected() {
  setStatus("", null);
  let value;
  try {
    value = getTargetValue();
  } catch (e) {
    setStatus(e.message, "err");
    return;
  }

  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id");
      await context.sync();

      if (shapes.items.length === 0) {
        setStatus("No shapes selected.", "err");
        return;
      }

      // Load adjustment counts so we can skip shapes with no handles
      // (lines, pictures, text boxes without geometry adjustments).
      const counts = shapes.items.map((s) => {
        s.adjustments.load("count");
        return s.adjustments;
      });
      await context.sync();

      let applied = 0;
      let skipped = 0;
      shapes.items.forEach((shape, i) => {
        const adj = counts[i];
        if (!adj || typeof adj.count !== "number" || adj.count < 1) {
          skipped++;
          return;
        }
        try {
          shape.adjustments.set(0, value);
          applied++;
        } catch {
          skipped++;
        }
      });

      await context.sync();

      const msg = `Set adjustment[0] = ${value.toFixed(4)} on ${applied} shape${applied === 1 ? "" : "s"}.` +
        (skipped > 0 ? ` (${skipped} skipped — no adjustment handle.)` : "");
      setStatus(msg, "ok");
    });
  } catch (e) {
    setStatus(formatError(e), "err");
  }
}

async function readFromSelected() {
  setStatus("", null);
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id,items/name");
      await context.sync();

      if (shapes.items.length === 0) {
        setStatus("No shapes selected.", "err");
        return;
      }

      const first = shapes.items[0];
      first.adjustments.load("count");
      await context.sync();

      if (!first.adjustments || first.adjustments.count < 1) {
        setStatus(`"${first.name}" has no adjustment handles.`, "err");
        return;
      }

      // Use get(index) which returns a ClientResult<number>.
      // Must sync before accessing .value.
      const result = first.adjustments.get(0);
      await context.sync();

      const raw = result.value;
      if (typeof raw !== "number") {
        setStatus(`"${first.name}" — could not read adjustment[0] value (preview API).`, "err");
        return;
      }

      const pct = raw / PERCENT_TO_RAW;
      els.input.value = els.modeRaw.checked ? raw.toFixed(4) : pct.toFixed(1);
      setStatus(`Read from "${first.name}": raw ${raw.toFixed(4)} (${pct.toFixed(1)}%).`, "ok");
    });
  } catch (e) {
    setStatus(formatError(e), "err");
  }
}

function setStatus(msg, kind) {
  els.status.textContent = msg || "";
  els.status.className = "status" + (kind ? " " + kind : "");
}

function formatError(e) {
  if (e && e.code === "ApiNotFound") {
    return "Shape adjustments require PowerPointApi 1.10 (preview). Use Office Insider/Beta build.";
  }
  if (e && e.debugInfo) {
    return `${e.message || "Error"}\n${e.debugInfo.message || ""}`;
  }
  return e && e.message ? e.message : String(e);
}
