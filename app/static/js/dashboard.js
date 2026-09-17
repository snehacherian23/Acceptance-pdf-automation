// dashboard.js
// Handles: sidebar navigation (mobile toggle + workflow switching),
// multi-file chip management for both workflows, and a SHARED
// generic job runner (submit -> poll -> render results) used by both
// Workflow 1 and Workflow 2 -- they differ only in which fields they
// send and which DOM elements they render into, described by the
// WORKFLOWS config below. This avoids duplicating the submit/poll/
// render logic for each workflow.

// =====================================================================
// Sidebar: mobile toggle + workflow switching
// =====================================================================
const HEADER_TEXT = {
    workflow1: {
        title: "Generate Acceptance Documents",
        subtitle: "Excel rows in, signed-ready Word documents out.",
    },
    workflow2: {
        title: "Prepare Final Signed Documents",
        subtitle: "Combine signed documents with Consolidated Sheets into the final deliverable.",
    },
};

function switchWorkflow(workflow) {
    document.querySelectorAll(".sidebar-nav-item").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.workflow === workflow);
    });
    document.querySelectorAll(".workflow-panel").forEach(panel => {
        panel.classList.toggle("d-none", panel.id !== workflow + "Panel");
    });

    const text = HEADER_TEXT[workflow];
    if (text) {
        document.getElementById("headerTitle").textContent = text.title;
        document.getElementById("headerSubtitle").textContent = text.subtitle;
    }

    closeSidebar();
}

document.querySelectorAll(".sidebar-nav-item").forEach(btn => {
    btn.addEventListener("click", () => switchWorkflow(btn.dataset.workflow));
});

function openSidebar() {
    document.getElementById("appSidebar").classList.add("open");
    document.getElementById("sidebarOverlay").classList.add("open");
}
function closeSidebar() {
    document.getElementById("appSidebar").classList.remove("open");
    document.getElementById("sidebarOverlay").classList.remove("open");
}

document.getElementById("sidebarToggle").addEventListener("click", openSidebar);
document.getElementById("sidebarOverlay").addEventListener("click", closeSidebar);

// =====================================================================
// Shared file-chip component (used by both workflows)
// =====================================================================
const fileState = {
    excel: [],
    signed: [],
    finalConsolidated: [],
};

const CHIP_CONFIG = {
    excel: {
        input: "excelInput", list: "excelChipList", count: "excelCount", clear: "clearExcel",
        icon: () => "bi-file-earmark-spreadsheet",
    },
    signed: {
        input: "signedInput", list: "signedChipList", count: "signedCount", clear: "clearSigned",
        icon: (file) => file.name.toLowerCase().endsWith(".pdf") ? "bi-file-earmark-pdf" : "bi-file-earmark-word",
    },
    finalConsolidated: {
        input: "finalConsolidatedInput", list: "finalConsolidatedChipList", count: "finalConsolidatedCount", clear: "clearFinalConsolidated",
        icon: (file) => file.name.toLowerCase().endsWith(".docx") ? "bi-file-earmark-word" : "bi-file-earmark-pdf",
    },
};

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function renderChips(category) {
    const cfg = CHIP_CONFIG[category];
    const files = fileState[category];
    const list = document.getElementById(cfg.list);
    const countEl = document.getElementById(cfg.count);
    const clearBtn = document.getElementById(cfg.clear);

    list.innerHTML = "";
    files.forEach((file, index) => {
        const chip = document.createElement("div");
        chip.className = "file-chip";
        chip.innerHTML = `
            <span class="chip-icon"><i class="bi ${cfg.icon(file)}"></i></span>
            <span class="chip-info">
                <div class="chip-name">${file.name}</div>
                <div class="chip-size">${formatFileSize(file.size)}</div>
            </span>
            <span class="chip-ready" title="Ready"><i class="bi bi-check-circle-fill"></i></span>
            <button type="button" class="chip-remove" data-category="${category}" data-index="${index}" title="Remove">
                <i class="bi bi-x-lg"></i>
            </button>
        `;
        list.appendChild(chip);
    });

    countEl.textContent = `${files.length} selected`;
    clearBtn.classList.toggle("d-none", files.length === 0);
}

function addFiles(category, newFiles) {
    const existingKeys = new Set(fileState[category].map(f => f.name + "|" + f.size));
    Array.from(newFiles).forEach(file => {
        const key = file.name + "|" + file.size;
        if (!existingKeys.has(key)) {
            fileState[category].push(file);
            existingKeys.add(key);
        }
    });
    renderChips(category);
}

function removeFile(category, index) {
    fileState[category].splice(index, 1);
    renderChips(category);
}

Object.keys(CHIP_CONFIG).forEach(category => {
    const cfg = CHIP_CONFIG[category];
    document.getElementById(cfg.input).addEventListener("change", function (e) {
        addFiles(category, e.target.files);
        e.target.value = "";  // reset so re-selecting the same file later still fires 'change'
    });
    document.getElementById(cfg.clear).addEventListener("click", function () {
        fileState[category] = [];
        renderChips(category);
    });
});

// Event delegation for remove (×) buttons across BOTH workflow forms
document.getElementById("workflow1Panel").addEventListener("click", handleChipRemoveClick);
document.getElementById("workflow2Panel").addEventListener("click", handleChipRemoveClick);

function handleChipRemoveClick(e) {
    const btn = e.target.closest(".chip-remove");
    if (btn) {
        removeFile(btn.dataset.category, parseInt(btn.dataset.index, 10));
    }
}

// =====================================================================
// Generic job runner -- shared by both workflows
// =====================================================================
// Each workflow is described once here: which form to listen on, how
// to validate + build its FormData, which endpoint to call, and which
// DOM elements to update. The actual submit/poll/render functions
// below are written once and driven entirely by this config.
const WORKFLOWS = {
    workflow1: {
        formId: "processingForm",
        endpoint: "/processing/run",
        submitBtnId: "processBtn",
        submitBtnIdleHtml: '<i class="bi bi-file-earmark-plus-fill"></i> Generate Documents',
        submitBtnBusyHtml: '<span class="spinner-border spinner-border-sm"></span> Generating...',
        validate: () => fileState.excel.length > 0 || "Please add at least one Excel sheet.",
        buildFormData: (fd) => { fileState.excel.forEach(f => fd.append("excel_files", f)); },
        el: {
            progressTrack: "progressTrack", processingStatus: "processingStatus",
            progressCount: "progressCount", currentFile: "currentFileLabel",
            completionPanel: "completionPanel", resultsPanel: "resultsPanel",
            cTotal: "cTotal", cSuccess: "cSuccess", cFailed: "cFailed", cReview: "cReview",
            downloadZip: "downloadZipBtn", downloadReport: "downloadReportBtn",
            statTotal: "statTotal", statSuccess: "statSuccess", statFailed: "statFailed",
            statReview: "statReview", statTime: "statTime", logPanel: "logPanel",
        },
    },
    workflow2: {
        formId: "workflow2Form",
        endpoint: "/processing/run-final",
        submitBtnId: "generateFinalBtn",
        submitBtnIdleHtml: '<i class="bi bi-file-earmark-pdf-fill"></i> Generate Final PDFs',
        submitBtnBusyHtml: '<span class="spinner-border spinner-border-sm"></span> Finalizing...',
        validate: () => {
            if (fileState.signed.length === 0) return "Please add at least one signed Acceptance Test document.";
            if (fileState.finalConsolidated.length === 0) return "Please add at least one Consolidated Sheet.";
            return true;
        },
        buildFormData: (fd) => {
            fileState.signed.forEach(f => fd.append("signed_files", f));
            fileState.finalConsolidated.forEach(f => fd.append("consolidated_files", f));
        },
        el: {
            progressTrack: "finalProgressTrack", processingStatus: "finalProcessingStatus",
            progressCount: "finalProgressCount", currentFile: "finalCurrentFileLabel",
            completionPanel: "finalCompletionPanel", resultsPanel: "workflow2ResultsPanel",
            cTotal: "fcTotal", cSuccess: "fcSuccess", cFailed: "fcFailed", cReview: "fcReview",
            downloadZip: "finalDownloadZipBtn", downloadReport: "finalDownloadReportBtn",
            statTotal: "finalStatTotal", statSuccess: "finalStatSuccess", statFailed: "finalStatFailed",
            statReview: "finalStatReview", statTime: "finalStatTime", logPanel: "finalLogPanel",
        },
    },
};

const pollTimers = {};

function setupWorkflow(key) {
    const wf = WORKFLOWS[key];
    const form = document.getElementById(wf.formId);

    form.addEventListener("submit", async function (e) {
        e.preventDefault();

        const validation = wf.validate();
        if (validation !== true) { alert(validation); return; }

        const btn = document.getElementById(wf.submitBtnId);
        const progressTrack = document.getElementById(wf.el.progressTrack);
        const processingStatus = document.getElementById(wf.el.processingStatus);

        const formData = new FormData();
        wf.buildFormData(formData);

        btn.disabled = true;
        btn.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Starting...';
        progressTrack.classList.add("active");
        progressTrack.classList.remove("determinate");
        document.getElementById(wf.el.completionPanel).classList.add("d-none");
        document.getElementById(wf.el.resultsPanel).classList.add("d-none");

        try {
            const response = await fetch(wf.endpoint, { method: "POST", body: formData });
            const data = await response.json();

            if (!response.ok) {
                alert("Error: " + (data.error || "Could not start processing."));
                resetSubmitButton(wf, btn);
                progressTrack.classList.remove("active");
                return;
            }

            btn.innerHTML = wf.submitBtnBusyHtml;
            processingStatus.classList.add("active");
            pollStatus(key, data.job_id);

        } catch (err) {
            alert("Request failed: " + err.message);
            resetSubmitButton(wf, btn);
            progressTrack.classList.remove("active");
        }
    });
}

function resetSubmitButton(wf, btn) {
    btn.disabled = false;
    btn.innerHTML = wf.submitBtnIdleHtml;
}

function pollStatus(workflowKey, jobId) {
    const wf = WORKFLOWS[workflowKey];
    pollTimers[workflowKey] = setInterval(async function () {
        try {
            const response = await fetch("/processing/status/" + jobId);
            const job = await response.json();
            if (!response.ok) { return; }

            updateLiveProgress(wf, job);

            if (job.status === "completed") {
                clearInterval(pollTimers[workflowKey]);
                finishProcessing(wf, job);
            }
        } catch (err) {
            // transient network hiccup while polling -- ignore and retry next tick
        }
    }, 700);
}

function updateLiveProgress(wf, job) {
    const total = job.total_files || 0;
    const processed = job.processed_count || 0;
    const pct = total > 0 ? Math.round((processed / total) * 100) : 0;

    const track = document.getElementById(wf.el.progressTrack);
    track.classList.add("determinate");
    track.style.setProperty("--progress-pct", pct + "%");

    document.getElementById(wf.el.progressCount).textContent = `${processed} / ${total}`;
    document.getElementById(wf.el.currentFile).textContent = job.current_file
        ? job.current_file
        : (processed >= total && total > 0 ? "Finishing up..." : "Starting...");
}

function finishProcessing(wf, job) {
    resetSubmitButton(wf, document.getElementById(wf.submitBtnId));
    document.getElementById(wf.el.progressTrack).classList.remove("active");
    document.getElementById(wf.el.processingStatus).classList.remove("active");

    renderCompletion(wf, job);
    renderStatCards(wf, job);
    renderLogs(wf, job);

    document.getElementById(wf.el.completionPanel).classList.remove("d-none");
    document.getElementById(wf.el.resultsPanel).classList.remove("d-none");
}

function renderCompletion(wf, job) {
    const stats = job.stats || { total: 0, success: 0, failed: 0 };
    document.getElementById(wf.el.cTotal).textContent = stats.total;
    document.getElementById(wf.el.cSuccess).textContent = stats.success;
    document.getElementById(wf.el.cFailed).textContent = stats.failed;
    document.getElementById(wf.el.cReview).textContent = stats.failed; // failed items need manual review

    const zipBtn = document.getElementById(wf.el.downloadZip);
    if (job.zip_path) {
        zipBtn.href = "/processing/download/" + job.id;
        zipBtn.classList.remove("d-none");
    } else {
        zipBtn.classList.add("d-none");
    }

    const reportBtn = document.getElementById(wf.el.downloadReport);
    if (job.report_path) {
        reportBtn.href = "/processing/download-report/" + job.id;
        reportBtn.classList.remove("d-none");
    } else {
        reportBtn.classList.add("d-none");
    }
}

function renderStatCards(wf, job) {
    const stats = job.stats || { total: 0, success: 0, failed: 0 };
    document.getElementById(wf.el.statTotal).textContent = stats.total;
    document.getElementById(wf.el.statSuccess).textContent = stats.success;
    document.getElementById(wf.el.statFailed).textContent = stats.failed;
    document.getElementById(wf.el.statReview).textContent = stats.failed;

    const duration = job.duration_seconds;
    document.getElementById(wf.el.statTime).textContent = duration != null ? `${duration}s` : "--";
}

function renderLogs(wf, job) {
    const logPanel = document.getElementById(wf.el.logPanel);
    logPanel.innerHTML = "";

    if (!job.logs || job.logs.length === 0) {
        logPanel.innerHTML = '<div class="empty-state"><i class="bi bi-inbox"></i>No log entries.</div>';
        return;
    }

    job.logs.forEach(function (entry) {
        const line = document.createElement("div");
        line.className = "log-line level-" + (entry.level || "info");
        line.textContent = "[" + entry.level.toUpperCase() + "]  " + entry.message;
        logPanel.appendChild(line);
    });

    logPanel.scrollTop = logPanel.scrollHeight;
}

// Wire up both workflows using the exact same code path
setupWorkflow("workflow1");
setupWorkflow("workflow2");
