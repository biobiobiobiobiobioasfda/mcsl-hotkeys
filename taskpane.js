/* LSCM Hotkeys (Office.js)
 *   Ctrl+W = yellow fill (#FFFF00)
 *   Ctrl+Q = blue font (#0000FF)
 *   Ctrl+J = center across selection, copying each row's source cell
 *            formatting (fill, font, number format, top/bottom borders)
 * Each action is one normal Excel undo step (ExcelApi 1.20 undo support),
 * so Ctrl+Z works like any built-in command and earlier edits stay undoable.
 */

const MAX_COPY_CELLS = 50000; // above this, Ctrl+J only sets alignment

function runOptions() {
  try {
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.20")) {
      return { mergeUndoGroup: true };
    }
  } catch (e) { /* ignore */ }
  return null;
}

function excelRun(batch) {
  const opts = runOptions();
  return opts ? Excel.run(opts, batch) : Excel.run(batch);
}

async function fillYellow() {
  try {
    await excelRun(async (ctx) => {
      ctx.workbook.getSelectedRanges().format.fill.color = "#FFFF00";
      await ctx.sync();
    });
  } catch (e) { console.error("FillYellow", e); }
}

async function fontBlue() {
  try {
    await excelRun(async (ctx) => {
      ctx.workbook.getSelectedRanges().format.font.color = "#0000FF";
      await ctx.sync();
    });
  } catch (e) { console.error("FontBlue", e); }
}

function hasFill(cell) {
  const f = cell.format && cell.format.fill;
  if (!f) return false;
  if (f.pattern !== undefined && f.pattern !== null) return f.pattern !== "None";
  return f.color && f.color.toUpperCase() !== "#FFFFFF";
}

function border(b) {
  if (!b || !b.style || b.style === "None") return { style: "None" };
  const out = { style: b.style };
  if (b.color) out.color = b.color;
  if (b.weight) out.weight = b.weight;
  return out;
}

function fillFrom(src) {
  const f = src.format.fill || {};
  if (!hasFill(src)) return { pattern: "None" };
  const out = { color: f.color };
  if (f.pattern) out.pattern = f.pattern;
  return out;
}

async function centerAcross() {
  try {
    await excelRun(async (ctx) => {
      const areas = ctx.workbook.getSelectedRanges().areas;
      areas.load("items/cellCount");
      await ctx.sync();

      const jobs = [];
      for (const r of areas.items) {
        if (r.cellCount > MAX_COPY_CELLS) { jobs.push({ r, skip: true }); continue; }
        const props = r.getCellProperties({
          format: {
            fill: { color: true, pattern: true },
            font: { bold: true, italic: true, color: true, name: true, size: true,
                    underline: true, strikethrough: true },
            borders: { color: true, style: true, weight: true }
          }
        });
        r.load("numberFormat, formulas");
        jobs.push({ r, props, skip: false });
      }
      await ctx.sync();

      for (const job of jobs) {
        const r = job.r;
        if (!job.skip) {
          const cp = job.props.value;
          const nf = r.numberFormat;
          const fx = r.formulas;
          const outProps = [];
          const outNf = [];
          for (let i = 0; i < cp.length; i++) {
            const row = cp[i];
            // Source: first filled cell, else first non-empty cell, else leftmost
            let s = row.findIndex(hasFill);
            if (s < 0) s = fx[i].findIndex((v) => v !== "" && v !== null);
            if (s < 0) s = 0;
            const src = row[s];
            const b = src.format.borders || {};
            const fmt = {
              fill: fillFrom(src),
              font: src.format.font,
              borders: { top: border(b.top), bottom: border(b.bottom) }
            };
            outProps.push(row.map((_, j) => (j === s ? {} : { format: fmt })));
            outNf.push(row.map(() => nf[i][s]));
          }
          r.setCellProperties(outProps);
          r.numberFormat = outNf;
        }
        r.format.horizontalAlignment = Excel.HorizontalAlignment.centerAcrossSelection;
      }
      await ctx.sync(); // all writes in one sync = one undo step
    });
  } catch (e) { console.error("CenterAcross", e); }
}

// Ribbon button: starts the add-in in this workbook and makes it auto-start
// the next time this workbook opens.
function startHotkeys(event) {
  setAutoStart().finally(() => { if (event && event.completed) event.completed(); });
}

function setAutoStart() {
  try {
    return Office.addin.setStartupBehavior(Office.StartupBehavior.load).catch(() => {});
  } catch (e) { return Promise.resolve(); }
}

Office.onReady(() => {
  Office.actions.associate("FillYellow", fillYellow);
  Office.actions.associate("FontBlue", fontBlue);
  Office.actions.associate("CenterAcross", centerAcross);
  Office.actions.associate("StartHotkeys", startHotkeys);
  setAutoStart();
  const s = document.getElementById("status");
  if (s) s.textContent = "Hotkeys are on in this workbook.";
});
