(() => {
  "use strict";

  const state = { batchId: null, tracers: [], template: null, jobId: null, logLen: 0, polling: null };
  const $ = (id) => document.getElementById(id);
  const el = {
    dzTracers: $("dz-tracers"), inTracers: $("input-tracers"), btnTracers: $("btn-tracers"),
    btnMore: $("btn-tracers-more"), btnClear: $("btn-tracers-clear"),
    block: $("tracer-block"), body: $("tracer-body"), count: $("tracer-count"),
    dzTemplate: $("dz-template"), inTemplate: $("input-template"), btnTemplate: $("btn-template"),
    tplCard: $("template-card"), tplDot: $("template-dot"), tplName: $("template-name"), tplSub: $("template-sub"),
    btnTplReplace: $("btn-template-replace"),
    month: $("month"), outName: $("output-name"),
    genSummary: $("generate-summary"), btnGenerate: $("btn-generate"),
    pProgress: $("panel-progress"), progTitle: $("progress-title"), progFill: $("progress-fill"),
    progNow: $("progress-now"), log: $("log"),
    pResult: $("panel-result"), resIcon: $("result-icon"), resTitle: $("result-title"), resSub: $("result-sub"),
    stats: $("stats"), btnDownload: $("btn-download"), btnNew: $("btn-new"),
    failedBlock: $("failed-block"), failedList: $("failed-list"), unitsBody: $("units-body"),
    checksSummary: $("checks-summary"), checkList: $("check-list"), logFinal: $("log-final"),
    errBanner: $("error-banner"), errMsg: $("error-message"), btnErrClose: $("btn-error-close"),
  };

  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function showError(msg) {
    el.errMsg.textContent = msg;
    el.errBanner.hidden = false;
    el.errBanner.scrollIntoView({ behavior: "smooth", block: "center" });
  }
  function clearError() { el.errBanner.hidden = true; }
  el.btnErrClose.addEventListener("click", clearError);

  async function api(url, opts) {
    let resp;
    try { resp = await fetch(url, opts); }
    catch (e) { throw new Error("Could not reach the tool. Is the Tracer Summary window (run.bat) still open?"); }
    let data = {};
    try { data = await resp.json(); } catch (e) { /* ignore */ }
    if (!resp.ok) throw new Error(data.error || ("Request failed (" + resp.status + ")."));
    return data;
  }

  function applyBatch(data) {
    state.batchId = data.batch_id;
    state.tracers = data.tracers || [];
    state.template = data.template;
    render();
  }

  // ---------------- uploads ----------------
  async function uploadTracers(files) {
    const list = Array.from(files || []);
    if (!list.length) return;
    clearError();
    const fd = new FormData();
    if (state.batchId) fd.append("batch_id", state.batchId);
    list.forEach((f) => fd.append("files", f));
    el.genSummary.textContent = "Checking " + list.length + " file(s)…";
    try {
      const data = await api("/api/tracers", { method: "POST", body: fd });
      applyBatch(data);
      if (data.rejected && data.rejected.length) showError("Skipped: " + data.rejected.join(", "));
    } catch (e) { showError(e.message); render(); }
  }

  async function uploadTemplate(file) {
    if (!file) return;
    clearError();
    const fd = new FormData();
    if (state.batchId) fd.append("batch_id", state.batchId);
    fd.append("file", file);
    try { applyBatch(await api("/api/template", { method: "POST", body: fd })); }
    catch (e) { showError(e.message); }
  }

  async function removeTracer(id) {
    try { applyBatch(await api("/api/tracers/" + state.batchId + "/" + id, { method: "DELETE" })); }
    catch (e) { showError(e.message); }
  }

  async function clearTracers() {
    if (!state.batchId) return;
    try { applyBatch(await api("/api/tracers/" + state.batchId + "/clear", { method: "POST" })); }
    catch (e) { showError(e.message); }
  }

  async function move(idx, delta) {
    const j = idx + delta;
    if (j < 0 || j >= state.tracers.length) return;
    const order = state.tracers.map((t) => t.id);
    [order[idx], order[j]] = [order[j], order[idx]];
    try {
      applyBatch(await api("/api/tracers/" + state.batchId + "/order", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ order }),
      }));
    } catch (e) { showError(e.message); }
  }

  function wireDropzone(dz, input, onFiles, multiple) {
    ["dragenter", "dragover"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => {
      const files = Array.from(e.dataTransfer.files || []);
      onFiles(multiple ? files : files.slice(0, 1));
    });
    dz.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
    input.addEventListener("change", () => { onFiles(Array.from(input.files)); input.value = ""; });
  }

  wireDropzone(el.dzTracers, el.inTracers, uploadTracers, true);
  wireDropzone(el.dzTemplate, el.inTemplate, (f) => uploadTemplate(f[0]), false);
  el.btnTracers.addEventListener("click", () => el.inTracers.click());
  el.btnMore.addEventListener("click", () => el.inTracers.click());
  el.btnClear.addEventListener("click", clearTracers);
  el.btnTemplate.addEventListener("click", () => el.inTemplate.click());
  el.btnTplReplace.addEventListener("click", () => el.inTemplate.click());

  el.body.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-act]");
    if (!b) return;
    const idx = Number(b.dataset.idx);
    if (b.dataset.act === "up") move(idx, -1);
    if (b.dataset.act === "down") move(idx, 1);
    if (b.dataset.act === "remove") removeTracer(b.dataset.id);
  });

  // ---------------- render ----------------
  function statusOf(t) {
    if (!t.valid) return ["badge-err", "Error"];
    if (t.duplicate_of) return ["badge-warn", "Duplicate"];
    if (t.same_unit_as) return ["badge-warn", "Same unit"];
    if (t.warnings && t.warnings.length) return ["badge-warn", "Ready · check"];
    return ["badge-ok", "Ready"];
  }

  function render() {
    const n = state.tracers.length;
    el.block.hidden = n === 0;
    el.dzTracers.hidden = n > 0;
    el.count.textContent = n + (n === 1 ? " file" : " files");
    el.body.innerHTML = state.tracers.map((t, i) => {
      const [cls, label] = statusOf(t);
      let issues = "";
      (t.errors || []).forEach((m) => { issues += '<span class="issue err">' + esc(m) + "</span>"; });
      if (t.duplicate_of) issues += '<span class="issue warn">Duplicate file detected: same content as ' + esc(t.duplicate_of) + ". Remove it, or keep it to include it twice.</span>";
      if (t.same_unit_as) issues += '<span class="issue warn">Same unit as ' + esc(t.same_unit_as) + " (e.g. two months) — both will be included.</span>";
      (t.warnings || []).forEach((m) => { issues += '<span class="issue warn">' + esc(m) + "</span>"; });
      return "<tr>" +
        '<td class="c-num">' + (i + 1) + "</td>" +
        '<td><div class="fname"><span class="' + (t.valid ? "tick" : "cross") + '">' + (t.valid ? "✓" : "✗") + "</span><div>" + esc(t.filename) + issues + "</div></div></td>" +
        '<td class="unit-cell">' + esc(t.unit || "—") + "</td>" +
        '<td class="c-right">' + esc(t.score_text || "—") + "</td>" +
        '<td class="c-right">' + (t.valid ? t.findings : "—") + "</td>" +
        '<td><span class="badge ' + cls + '">' + label + "</span></td>" +
        '<td class="c-actions">' +
          '<button class="icon-btn" title="Move up" data-act="up" data-idx="' + i + '"' + (i === 0 ? " disabled" : "") + ">↑</button> " +
          '<button class="icon-btn" title="Move down" data-act="down" data-idx="' + i + '"' + (i === n - 1 ? " disabled" : "") + ">↓</button> " +
          '<button class="icon-btn remove" title="Remove" data-act="remove" data-id="' + t.id + '">×</button>' +
        "</td></tr>";
    }).join("");

    const tpl = state.template;
    el.tplCard.hidden = !tpl;
    el.dzTemplate.hidden = !!tpl;
    if (tpl) {
      el.tplName.textContent = (tpl.valid ? "✓ " : "✗ ") + tpl.filename;
      el.tplDot.classList.toggle("bad", !tpl.valid);
      el.tplSub.classList.toggle("err", !tpl.valid);
      el.tplSub.textContent = tpl.valid ? "Summary sheet detected: " + tpl.sheet : tpl.error;
    }

    const valid = state.tracers.filter((t) => t.valid);
    const bad = n - valid.length;
    const findings = valid.reduce((a, t) => a + (t.findings || 0), 0);
    const ready = valid.length > 0 && tpl && tpl.valid;
    el.btnGenerate.disabled = !ready || !!state.polling;
    if (!n && !tpl) el.genSummary.textContent = "Upload tracer files and the template to begin.";
    else if (!n) el.genSummary.textContent = "Upload the tracer files (step 1).";
    else if (!tpl) el.genSummary.textContent = valid.length + " tracer file(s) ready — now upload the Summary template (step 2).";
    else if (!tpl.valid) el.genSummary.textContent = "The template is not valid — please upload the Tracer Summary template.";
    else if (!valid.length) el.genSummary.textContent = "None of the tracer files could be read.";
    else el.genSummary.innerHTML = "<b>" + valid.length + "</b> tracer file(s) ready · <b>" + findings + "</b> findings" +
      (bad ? ' · <span style="color:var(--error)">' + bad + " file(s) with errors will be skipped</span>" : "");
  }

  // ---------------- generate ----------------
  el.btnGenerate.addEventListener("click", async () => {
    clearError();
    let job;
    try {
      job = await api("/api/generate", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ batch_id: state.batchId, month: el.month.value, output_name: el.outName.value }),
      });
    } catch (e) { showError(e.message); return; }
    state.jobId = job.job_id;
    state.logLen = 0;
    el.log.innerHTML = "";
    el.pResult.hidden = true;
    el.pProgress.hidden = false;
    el.progTitle.textContent = "Processing…";
    el.progFill.style.width = "0%";
    el.progNow.textContent = "Starting…";
    el.pProgress.scrollIntoView({ behavior: "smooth", block: "start" });
    state.polling = setInterval(poll, 300);
    render();
  });

  function logLine(entry) {
    const d = document.createElement("div");
    d.className = "l-" + entry.t;
    d.textContent = entry.m;
    return d;
  }

  async function poll() {
    let data;
    try { data = await api("/api/job/" + state.jobId + "?since=" + state.logLen); }
    catch (e) { stopPolling(); showError(e.message); return; }
    data.log.forEach((entry) => el.log.appendChild(logLine(entry)));
    state.logLen = data.log_len;
    el.log.scrollTop = el.log.scrollHeight;
    const total = Math.max(1, data.total);
    let pct = (data.current / total) * 85;
    if (data.stage === "writing") pct = 90;
    if (data.stage === "validating") pct = 96;
    if (data.stage === "done") pct = 100;
    el.progFill.style.width = pct + "%";
    if (data.stage === "processing") el.progNow.textContent = "Processing " + data.current + " of " + data.total + " — " + data.current_file;
    else if (data.stage === "writing") el.progNow.textContent = "Writing the Tracer Summary into the template…";
    else if (data.stage === "validating") el.progNow.textContent = "Checking the generated workbook…";
    if (data.status !== "running") { stopPolling(); showResult(data); }
  }

  function stopPolling() { clearInterval(state.polling); state.polling = null; render(); }

  function stat(num, label, err) {
    return '<div class="stat"><div class="stat-num' + (err ? " err" : "") + '">' + num + '</div><div class="stat-label">' + label + "</div></div>";
  }

  function showResult(data) {
    const r = data.result || {};
    el.progTitle.textContent = "Processing complete";
    el.progNow.textContent = "";
    el.pResult.classList.remove("bad", "partial");
    el.logFinal.innerHTML = "";
    Array.from(el.log.children).forEach((c) => el.logFinal.appendChild(c.cloneNode(true)));

    if (data.status !== "done") {
      el.pResult.classList.add("bad");
      el.resIcon.textContent = "!";
      el.resTitle.textContent = "TRACER SUMMARY COULD NOT BE GENERATED";
      el.resSub.textContent = r.message || "See the processing log below.";
      el.stats.innerHTML = stat(0, "files processed") + stat(r.errors || 0, "errors", true);
      el.btnDownload.hidden = true;
    } else {
      const partial = !r.success || r.errors > 0;
      if (partial) el.pResult.classList.add("partial");
      el.resIcon.textContent = partial ? "!" : "✓";
      el.resTitle.textContent = r.success && !r.errors ? "TRACER SUMMARY GENERATED SUCCESSFULLY"
        : (r.success ? "TRACER SUMMARY GENERATED — SOME FILES WERE SKIPPED" : "TRACER SUMMARY GENERATED — PLEASE REVIEW THE CHECKS");
      el.resSub.textContent = "All output checks: " + r.checks_passed + " of " + r.checks_total + " passed" +
        (r.warnings ? " · " + r.warnings + " warning(s) in the log" : "");
      el.stats.innerHTML = stat(r.files_processed, "files processed") + stat(r.findings, "findings extracted") +
        stat(r.errors, r.errors === 1 ? "error" : "errors", r.errors > 0) + stat(r.rows, "unit rows written");
      el.btnDownload.hidden = false;
      el.btnDownload.href = r.download_url;
      el.btnDownload.textContent = "DOWNLOAD " + r.download_name;
      el.btnDownload.setAttribute("download", r.download_name);
      el.unitsBody.innerHTML = (r.units || []).map((u) =>
        "<tr><td><b>" + esc(u.unit) + '</b></td><td class="c-right">' + esc(u.score) + '</td><td class="c-right">' + u.findings + "</td><td>" + esc(u.file) + "</td></tr>").join("");
      el.checksSummary.textContent = "Output checks (" + r.checks_passed + "/" + r.checks_total + " passed)";
      el.checkList.innerHTML = (r.checks || []).map((c) => '<li class="' + (c.ok ? "ok" : "bad") + '">' + esc(c.m) + "</li>").join("");
    }
    const failed = r.failed || [];
    el.failedBlock.hidden = !failed.length;
    el.failedList.innerHTML = failed.map((f) => "<li><b>" + esc(f.filename) + "</b> — " + esc((f.errors || []).join("; ")) + "</li>").join("");
    el.pResult.hidden = false;
    el.pResult.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  el.btnNew.addEventListener("click", () => {
    el.pResult.hidden = true;
    el.pProgress.hidden = true;
    clearTracers();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  el.outName.addEventListener("input", () => {
    const v = el.outName.value.replace(/[\\/:*?"<>|]/g, "");
    if (v !== el.outName.value) el.outName.value = v;
  });

  render();
})();
