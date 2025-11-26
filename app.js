// ============================================================================
// SmartAI Proposal Builder - Consultant Workspace
// ============================================================================

// State
let state = {
  apiBase: "https://sgdev-smartai-api-01.azurewebsites.net",
  uploadBrokerBase: "https://sgdev-smartai-func-sas-ehbpeahwg2b6bzcy.southeastasia-01.azurewebsites.net",
  env: {
    feature_psg_enabled: false,
    model_worker: null,
    model_manager: null,
    packs_latest: { EDG: null, PSG: null },
    appconfig_label: "default"
  },
  sessionId: null,
  workflowType: "grant", // "grant" or "other"
  grant: "EDG",
  checklist: { uploads: [], drafts: [] },
  evidenceDetected: [],
  evidenceSelected: [],
  evidenceUploadStatus: {}, // label -> "uploaded" | "detected"
  facts: {},
  validationChecks: [],
  solutionAnchor: "",
  sharedDraft: {
    style: "Formal, outcome-oriented",
    length_limit: 300,
    evidence_char_cap: 0
  },
  outputs: {}, // section_id -> DraftResponse
  activeStep: 0,
  stepStatus: ["todo", "todo", "todo", "todo", "todo"] // todo | in_progress | done
};

const STEP_NAMES = [
  "Start Proposal Run",
  "Evidence Collection",
  "SME Snapshot",
  "Solution Anchor",
  "Draft Document Set",
  "Review / Export"
];

// ============================================================================
// API Helpers
// ============================================================================

function getApiBase() {
  const input = document.getElementById("apiBase");
  if (input) {
    const base = input.value.trim().replace(/\/+$/, "");
    state.apiBase = base;
    localStorage.setItem("apiBase", base);
    return base;
  }
  return state.apiBase;
}

async function apiGet(path) {
  const base = getApiBase();
  const res = await fetch(`${base}${path}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
  return res.json();
}

async function apiPost(path, body) {
  const base = getApiBase();
  const res = await fetch(`${base}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const text = await res.text();
    let error;
    try {
      error = JSON.parse(text);
    } catch {
      error = { error: text };
    }
    throw new Error(error.error || `HTTP ${res.status}: ${res.statusText}`);
  }
  return res.json();
}

async function brokerPost(path, body) {
  const base = state.uploadBrokerBase.replace(/\/+$/, "");
  const url = base + path;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const text = await res.text();
    let error;
    try {
      error = JSON.parse(text);
    } catch {
      error = { error: text };
    }
    throw new Error(error.error || `HTTP ${res.status}: ${res.statusText}`);
  }
  return res.json();
}

// ============================================================================
// Persistence
// ============================================================================

function persistState() {
  if (state.sessionId) localStorage.setItem("sid", state.sessionId);
  if (state.grant) localStorage.setItem("grant", state.grant);
  if (state.workflowType) localStorage.setItem("workflowType", state.workflowType);
  if (state.solutionAnchor) localStorage.setItem("solutionAnchor", state.solutionAnchor);
  if (state.sharedDraft.style) localStorage.setItem("sharedDraftStyle", state.sharedDraft.style);
  if (state.sharedDraft.length_limit) localStorage.setItem("sharedDraftLengthLimit", state.sharedDraft.length_limit.toString());
  if (state.sharedDraft.evidence_char_cap) localStorage.setItem("sharedDraftEvidenceCap", state.sharedDraft.evidence_char_cap.toString());
}

function hydrateState() {
  const apiBase = localStorage.getItem("apiBase");
  if (apiBase) {
    state.apiBase = apiBase;
    const input = document.getElementById("apiBase");
    if (input) input.value = apiBase;
  }
  const sid = localStorage.getItem("sid");
  if (sid) state.sessionId = sid;
  const grant = localStorage.getItem("grant");
  if (grant) state.grant = grant;
  const workflowType = localStorage.getItem("workflowType");
  if (workflowType) state.workflowType = workflowType;
  const solutionAnchor = localStorage.getItem("solutionAnchor");
  if (solutionAnchor) state.solutionAnchor = solutionAnchor;
  const style = localStorage.getItem("sharedDraftStyle");
  if (style) state.sharedDraft.style = style;
  const lengthLimit = localStorage.getItem("sharedDraftLengthLimit");
  if (lengthLimit) state.sharedDraft.length_limit = parseInt(lengthLimit, 10);
  const evidenceCap = localStorage.getItem("sharedDraftEvidenceCap");
  if (evidenceCap) state.sharedDraft.evidence_char_cap = parseInt(evidenceCap, 10);
}

// ============================================================================
// Renderers: Top Bar
// ============================================================================

function renderTopBar() {
  const container = document.getElementById("envPills");
  if (!container) return;

  const { env } = state;
  const pills = [];

  if (env.appconfig_label) {
    pills.push(`<span class="px-2 py-1 bg-gray-100 border border-gray-300 rounded-full text-xs text-gray-700"><strong>Env:</strong> ${env.appconfig_label}</span>`);
  }

  if (env.packs_latest?.EDG) {
    pills.push(`<span class="px-2 py-1 bg-gray-100 border border-gray-300 rounded-full text-xs text-gray-700"><strong>EDG:</strong> ${env.packs_latest.EDG}</span>`);
  }

  if (env.packs_latest?.PSG) {
    pills.push(`<span class="px-2 py-1 bg-gray-100 border border-gray-300 rounded-full text-xs text-gray-700"><strong>PSG:</strong> ${env.packs_latest.PSG}</span>`);
  }

  if (env.model_worker) {
    pills.push(`<span class="px-2 py-1 bg-gray-100 border border-gray-300 rounded-full text-xs text-gray-700"><strong>Model:</strong> ${env.model_worker}</span>`);
  }

  if (env.model_manager) {
    pills.push(`<span class="px-2 py-1 bg-gray-100 border border-gray-300 rounded-full text-xs text-gray-700"><strong>Manager:</strong> ${env.model_manager}</span>`);
  }

  pills.push(`<span class="px-2 py-1 ${env.feature_psg_enabled ? "bg-green-50 border-green-300 text-green-700" : "bg-gray-100 border-gray-300 text-gray-700"} rounded-full text-xs"><strong>PSG:</strong> ${env.feature_psg_enabled ? "enabled" : "disabled"}</span>`);

  container.innerHTML = pills.join("");
}

async function loadEnvConfig() {
  try {
    const features = await apiGet("/v1/config/features");
    state.env.feature_psg_enabled = features.feature_psg_enabled || false;
    state.env.model_worker = features.model_worker || null;
    state.env.packs_latest = features.packs_latest || { EDG: null, PSG: null };

    try {
      const active = await apiGet("/v1/prompts/active");
      state.env.appconfig_label = active.appconfig_label || "default";
      if (active.model_worker) state.env.model_worker = active.model_worker;
      if (active.model_manager) state.env.model_manager = active.model_manager;
      if (active.packs_latest) state.env.packs_latest = active.packs_latest;
    } catch (e) {
      console.warn("Could not load /v1/prompts/active:", e);
    }

    renderTopBar();
  } catch (e) {
    console.error("Failed to load env config:", e);
    renderTopBar();
  }
}

// ============================================================================
// Renderers: Left Rail
// ============================================================================

function getStepStatusIcon(status) {
  if (status === "done") return "✓";
  if (status === "in_progress") return "→";
  return "○";
}

function getStepStatusClass(status, stepIndex) {
  const isActive = state.activeStep === stepIndex;
  if (status === "done") return isActive ? "bg-green-50 border-green-300 text-green-900" : "text-green-700";
  if (status === "in_progress" || isActive) return "bg-blue-50 border-blue-300 text-blue-900 font-medium";
  return "text-gray-600";
}

function renderLeftRail() {
  const container = document.getElementById("leftRail");
  if (!container) return;

  const steps = STEP_NAMES.map((name, idx) => {
    const status = state.stepStatus[idx] || "todo";
    const isActive = state.activeStep === idx;
    const icon = getStepStatusIcon(status);
    const statusClass = getStepStatusClass(status, idx);

    return `
      <button
        onclick="navigateToStep(${idx})"
        class="w-full text-left px-4 py-3 border-l-2 ${isActive ? "border-blue-500" : "border-transparent"} ${statusClass} hover:bg-gray-50 transition-colors"
      >
        <div class="flex items-center gap-2">
          <span class="text-sm font-medium">${icon}</span>
          <span class="text-sm">${name}</span>
        </div>
      </button>
    `;
  }).join("");

  container.innerHTML = `
    <div class="p-4 border-b border-gray-200">
      <h2 class="text-sm font-semibold text-gray-700 uppercase tracking-wide">Workflow</h2>
    </div>
    <div class="py-2">
      ${steps}
    </div>
  `;
}

// ============================================================================
// Renderers: Step Panels
// ============================================================================

function renderStepPanel() {
  const container = document.getElementById("mainPanel");
  if (!container) return;

  const step = state.activeStep;
  let html = "";

  switch (step) {
    case 0: html = renderStep0(); break;
    case 1: html = renderStep1(); break;
    case 2: html = renderStep2(); break;
    case 3: html = renderStep3(); break;
    case 4: html = renderStep4(); break;
    case 5: html = renderStep5(); break;
    default: html = renderStep0();
  }

  container.innerHTML = html;
  
  // Re-attach event listeners after render
  attachStepEventListeners();
}

function attachStepEventListeners() {
  // Workflow type change
  const workflowTypeSelect = document.getElementById("workflowType");
  if (workflowTypeSelect) {
    workflowTypeSelect.onchange = (e) => {
      state.workflowType = e.target.value;
      const grantSelector = document.getElementById("grantSelector");
      if (grantSelector) {
        grantSelector.classList.toggle("hidden", e.target.value !== "grant");
      }
    };
  }

  // Evidence select change
  const evidenceSelect = document.getElementById("evidenceSelect");
  if (evidenceSelect) {
    evidenceSelect.onchange = (e) => {
      state.evidenceSelected = Array.from(e.target.selectedOptions).map(o => o.value);
    };
  }
}

function renderStep0() {
  const psgDisabled = !state.env.feature_psg_enabled;
  return `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900 mb-2">Start Proposal Run</h2>
        <p class="text-sm text-gray-600">Create a new proposal session and load the grant-aware checklist.</p>
      </div>

      <div class="space-y-4">
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-2">Workflow Type</label>
          <select id="workflowType" class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500">
            <option value="grant" ${state.workflowType === "grant" ? "selected" : ""}>Grant Proposal (EDG/PSG)</option>
            <option value="other" ${state.workflowType === "other" ? "selected" : ""}>Other document set (future-ready)</option>
          </select>
        </div>

        <div id="grantSelector" class="${state.workflowType === "grant" ? "" : "hidden"}">
          <label class="block text-sm font-medium text-gray-700 mb-2">Grant Type</label>
          <select id="grantSelect" class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500">
            <option value="EDG" ${state.grant === "EDG" ? "selected" : ""}>EDG (Enterprise Development Grant)</option>
            <option value="PSG" ${state.grant === "PSG" ? "selected" : ""} ${psgDisabled ? "disabled" : ""}>PSG (Productivity Solutions Grant)${psgDisabled ? " (disabled)" : ""}</option>
          </select>
        </div>

        <div class="flex items-center gap-3 pt-4">
          <button
            onclick="handleCreateSession()"
            class="px-4 py-2 bg-gradient-to-r from-green-500 to-green-600 text-white font-medium rounded-lg hover:from-green-600 hover:to-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all shadow-sm"
          >
            Create Run
          </button>
          ${state.sessionId ? `<span class="text-sm text-gray-600">Session: <code class="px-2 py-1 bg-gray-100 rounded text-xs">${state.sessionId}</code></span>` : ""}
        </div>
      </div>
    </div>
  `;
}

function renderStep1() {
  const uploadTasks = state.checklist.uploads || [];
  const detectedLabels = new Set(state.evidenceDetected.map(e => e.label));

  return `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900 mb-2">Evidence Collection</h2>
        <p class="text-sm text-gray-600">Upload required documents. Evidence will be processed and available for drafting.</p>
      </div>

      <div class="space-y-4 mb-6">
        <div class="flex items-center justify-between">
          <h3 class="text-lg font-semibold text-gray-900">Upload Tasks</h3>
          <button
            onclick="handleRefreshEvidence()"
            class="px-3 py-1.5 text-sm bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg transition-colors"
          >
            Refresh Evidence Detected
          </button>
        </div>

        ${uploadTasks.length === 0 ? `
          <p class="text-sm text-gray-500 italic">No upload tasks in checklist.</p>
        ` : `
          <div class="overflow-x-auto">
            <table class="min-w-full divide-y divide-gray-200">
              <thead class="bg-gray-50">
                <tr>
                  <th class="px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase">Label</th>
                  <th class="px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase">Upload</th>
                  <th class="px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase">Status</th>
                </tr>
              </thead>
              <tbody class="bg-white divide-y divide-gray-200">
                ${uploadTasks.map(task => {
                  const label = task.id || task.label;
                  const displayLabel = label.replace(/_/g, " "); // Human-readable display only
                  const uploadStatus = state.evidenceUploadStatus[label] || "not_uploaded";
                  const isDetected = detectedLabels.has(label);
                  let statusText = "Not uploaded";
                  let statusClass = "bg-gray-100 text-gray-700";
                  if (isDetected) {
                    statusText = "Detected";
                    statusClass = "bg-green-100 text-green-700";
                  } else if (uploadStatus === "uploaded") {
                    statusText = "Uploaded (pending parse)";
                    statusClass = "bg-yellow-100 text-yellow-700";
                  } else if (uploadStatus === "uploading") {
                    statusText = "Uploading…";
                    statusClass = "bg-blue-100 text-blue-700";
                  } else if (uploadStatus === "error") {
                    statusText = "Upload failed";
                    statusClass = "bg-red-100 text-red-700";
                  }
                  return `
                    <tr>
                      <td class="px-4 py-3 text-sm font-medium text-gray-900">${displayLabel}</td>
                      <td class="px-4 py-3">
                        <input
                          type="file"
                          data-label="${label}"
                          onchange="handleFileSelect('${label}', this)"
                          class="text-xs"
                        />
                      </td>
                      <td class="px-4 py-3">
                        <span class="px-2 py-1 text-xs rounded-full ${statusClass}">${statusText}</span>
                      </td>
                    </tr>
                  `;
                }).join("")}
              </tbody>
            </table>
          </div>
        `}
      </div>

      <div class="border-t border-gray-200 pt-4">
        <details class="cursor-pointer">
          <summary class="text-sm font-medium text-gray-700 hover:text-gray-900">Evidence Inspector</summary>
          <div id="evidenceInspector" class="mt-4 space-y-2 text-sm">
            ${state.evidenceDetected.length === 0 ? `
              <p class="text-gray-500 italic text-sm">No evidence detected yet.</p>
            ` : state.evidenceDetected.map(item => {
              const displayLabel = item.label.replace(/_/g, " "); // Human-readable display only
              return `
                <div class="border border-gray-200 rounded p-2 bg-gray-50">
                  <div class="flex items-center gap-2 mb-1">
                    <span class="px-2 py-0.5 bg-green-100 text-green-700 rounded text-xs font-medium">${displayLabel}</span>
                    <span class="text-xs text-gray-500">${item.chars || 0} chars</span>
                  </div>
                  <div class="text-xs text-gray-600 font-mono whitespace-pre-wrap">${(item.preview || "").replace(/</g, "&lt;")}</div>
                </div>
              `;
            }).join("")}
          </div>
        </details>
      </div>
    </div>
  `;
}

function renderStep2() {
  const checks = state.validationChecks || [];
  const warnings = checks.filter(c => (c.level || "").toLowerCase() === "warning");
  const errors = checks.filter(c => (c.level || "").toLowerCase() === "error");

  return `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900 mb-2">SME Snapshot + Eligibility</h2>
        <p class="text-sm text-gray-600">Capture company facts and run non-blocking eligibility validation.</p>
      </div>

      ${(warnings.length > 0 || errors.length > 0) ? `
        <div class="mb-4 p-3 rounded-lg ${errors.length > 0 ? "bg-red-50 border border-red-200" : "bg-yellow-50 border border-yellow-200"}">
          <div class="flex items-center gap-2 mb-1">
            <span class="text-sm font-medium ${errors.length > 0 ? "text-red-800" : "text-yellow-800"}">
              ${warnings.length} warning${warnings.length !== 1 ? "s" : ""}${errors.length > 0 ? `, ${errors.length} error${errors.length !== 1 ? "s" : ""}` : ""}
            </span>
          </div>
          <div class="text-xs ${errors.length > 0 ? "text-red-700" : "text-yellow-700"}">
            ${checks.map(c => `${c.code}: ${c.message}`).join(" • ")}
          </div>
        </div>
      ` : checks.length === 0 ? "" : `
        <div class="mb-4 p-3 rounded-lg bg-green-50 border border-green-200">
          <span class="text-sm text-green-800">No validation issues detected.</span>
        </div>
      `}

      <div class="space-y-4">
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Local Equity (%)</label>
            <input
              id="factEquity"
              type="number"
              min="0"
              max="100"
              value="${state.facts.local_equity_pct || ""}"
              placeholder="e.g. 45"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Turnover (SGD)</label>
            <input
              id="factTurnover"
              type="number"
              min="0"
              value="${state.facts.turnover || ""}"
              placeholder="optional"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Headcount</label>
            <input
              id="factHeadcount"
              type="number"
              min="0"
              value="${state.facts.headcount || ""}"
              placeholder="optional"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
        </div>

        <div>
          <label class="block text-sm font-medium text-gray-700 mb-1">Extra Facts (JSON or key:value lines)</label>
          <textarea
            id="factExtra"
            rows="3"
            placeholder='{"industry":"F&B","budget_range":"<50k"}'
            class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500 font-mono text-sm"
          >${state.facts.extra ? JSON.stringify(state.facts.extra, null, 2) : ""}</textarea>
        </div>

        <div class="flex items-center gap-3 pt-2">
          <button
            onclick="handleSaveFacts()"
            class="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 font-medium rounded-lg transition-colors"
          >
            Save SME Snapshot
          </button>
          <button
            onclick="handleValidate()"
            class="px-4 py-2 bg-gradient-to-r from-green-500 to-green-600 text-white font-medium rounded-lg hover:from-green-600 hover:to-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all shadow-sm"
          >
            Run Eligibility Checks
          </button>
        </div>
      </div>
    </div>
  `;
}

function renderStep3() {
  return `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900 mb-2">Solution Anchor</h2>
        <p class="text-sm text-gray-600">Define the core problem statement and solution context. This will be used as the prompt for all draft sections.</p>
      </div>

      <div class="space-y-4">
        <div>
          <label class="block text-sm font-medium text-gray-700 mb-2">Solution Anchor / Tailored Problem Statement</label>
          <textarea
            id="solutionAnchor"
            rows="6"
            placeholder="Our revenue fell 12% due to low education and confidence in AI; propose a plan."
            class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
          >${state.solutionAnchor}</textarea>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 pt-4 border-t border-gray-200">
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Style</label>
            <input
              id="sharedStyle"
              type="text"
              value="${state.sharedDraft.style}"
              placeholder="Formal, outcome-oriented"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Length Limit (words)</label>
            <input
              id="sharedLengthLimit"
              type="number"
              min="50"
              step="50"
              value="${state.sharedDraft.length_limit}"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
          <div>
            <label class="block text-sm font-medium text-gray-700 mb-1">Evidence Char Cap</label>
            <input
              id="sharedEvidenceCap"
              type="number"
              min="500"
              step="500"
              value="${state.sharedDraft.evidence_char_cap || ""}"
              placeholder="6000"
              class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500"
            />
          </div>
        </div>

        <div class="pt-4">
          <button
            onclick="handleSaveAnchor()"
            class="px-4 py-2 bg-gradient-to-r from-green-500 to-green-600 text-white font-medium rounded-lg hover:from-green-600 hover:to-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all shadow-sm"
          >
            Save Anchor
          </button>
        </div>
      </div>
    </div>
  `;
}

function renderStep4() {
  const draftTasks = state.checklist.drafts || [];
  const detectedLabels = state.evidenceDetected.map(e => e.label);

  return `
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div class="lg:col-span-2 bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <div class="mb-6">
          <h2 class="text-2xl font-bold text-gray-900 mb-2">Draft Document Set</h2>
          <p class="text-sm text-gray-600">Draft individual sections or generate all sections sequentially.</p>
        </div>

        ${draftTasks.length === 0 ? `
          <p class="text-sm text-gray-500 italic">No draft tasks in checklist.</p>
        ` : `
          <div class="space-y-3 mb-4">
            ${draftTasks.map(task => {
              const sectionId = task.id;
              const sectionVariant = task.section_variant || null;
              const output = state.outputs[sectionId];
              const hasOutput = !!output;
              return `
                <div class="border border-gray-200 rounded-lg p-4 ${hasOutput ? "bg-green-50 border-green-300" : ""}">
                  <div class="flex items-center justify-between mb-2">
                    <div>
                      <span class="font-medium text-gray-900">${sectionId}</span>
                      ${sectionVariant ? `<span class="text-xs text-gray-500 ml-2">(${sectionVariant})</span>` : ""}
                    </div>
                    <button
                      onclick="handleDraftSection('${sectionId}', ${sectionVariant ? `'${sectionVariant}'` : "null"})"
                      class="px-3 py-1.5 text-sm bg-green-500 hover:bg-green-600 text-white rounded-lg transition-colors"
                      ${!state.solutionAnchor ? "disabled title='Save Solution Anchor first'" : ""}
                    >
                      Draft
                    </button>
                  </div>
                  ${hasOutput ? `
                    <div class="mt-2 text-xs text-gray-600">
                      <span>Framework: <strong>${output.framework || "-"}</strong></span>
                      <span class="mx-2">|</span>
                      <span>Score: <strong>${output.evaluation?.score ?? "-"}</strong></span>
                      ${output.evidence_used?.length ? `
                        <span class="mx-2">|</span>
                        <span>Evidence: ${output.evidence_used.map(l => `<span class="px-1.5 py-0.5 bg-gray-200 rounded text-xs">${l}</span>`).join(" ")}</span>
                      ` : ""}
                    </div>
                  ` : ""}
                </div>
              `;
            }).join("")}
          </div>

          <div class="pt-4 border-t border-gray-200">
            <div class="mb-3">
              <label class="block text-sm font-medium text-gray-700 mb-2">Evidence Labels (optional, leave empty for defaults)</label>
              <select
                id="evidenceSelect"
                multiple
                class="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-green-500 text-sm"
                size="4"
              >
                ${detectedLabels.map(label => `
                  <option value="${label}" ${state.evidenceSelected.includes(label) ? "selected" : ""}>${label}</option>
                `).join("")}
              </select>
            </div>
            <button
              onclick="handleDraftAll()"
              class="w-full px-4 py-2 bg-gradient-to-r from-green-500 to-green-600 text-white font-medium rounded-lg hover:from-green-600 hover:to-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all shadow-sm"
              ${!state.solutionAnchor ? "disabled title='Save Solution Anchor first'" : ""}
            >
              Draft All Sections
            </button>
          </div>
        `}
      </div>

      <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <h3 class="text-lg font-semibold text-gray-900 mb-4">Latest Output</h3>
        <div id="latestOutput" class="text-sm text-gray-600 font-mono whitespace-pre-wrap bg-gray-50 p-3 rounded border border-gray-200 max-h-96 overflow-y-auto">
          ${Object.keys(state.outputs).length === 0 ? "No sections drafted yet." : "Select a section to view its output."}
        </div>
      </div>
    </div>
  `;
}

function renderStep5() {
  const draftTasks = state.checklist.drafts || [];
  const draftedSections = draftTasks.filter(t => state.outputs[t.id]);

  return `
    <div class="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div class="mb-6">
        <h2 class="text-2xl font-bold text-gray-900 mb-2">Review / Export</h2>
        <p class="text-sm text-gray-600">Review drafted sections and export the complete proposal.</p>
      </div>

      ${draftedSections.length === 0 ? `
        <p class="text-sm text-gray-500 italic">No sections have been drafted yet. Go to Step 4 to draft sections.</p>
      ` : `
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div>
            <h3 class="text-lg font-semibold text-gray-900 mb-3">Solution Anchor</h3>
            <div class="bg-gray-50 border border-gray-200 rounded-lg p-4 text-sm text-gray-700 whitespace-pre-wrap font-mono max-h-64 overflow-y-auto">
              ${state.solutionAnchor || "(not set)"}
            </div>
          </div>
          <div>
            <h3 class="text-lg font-semibold text-gray-900 mb-3">Drafted Sections</h3>
            <div class="space-y-2 mb-4">
              ${draftedSections.map(task => {
                const sectionId = task.id;
                const output = state.outputs[sectionId];
                return `
                  <div class="border border-gray-200 rounded-lg p-3">
                    <div class="flex items-center justify-between mb-2">
                      <span class="font-medium text-sm text-gray-900">${sectionId}${task.section_variant ? ` (${task.section_variant})` : ""}</span>
                      <button
                        onclick="handleCopySection('${sectionId}')"
                        class="px-2 py-1 text-xs bg-gray-100 hover:bg-gray-200 text-gray-700 rounded transition-colors"
                      >
                        Copy
                      </button>
                    </div>
                    <div class="text-xs text-gray-600 font-mono whitespace-pre-wrap max-h-32 overflow-y-auto bg-gray-50 p-2 rounded border border-gray-200">
                      ${(output?.output || "").substring(0, 200)}${(output?.output || "").length > 200 ? "..." : ""}
                    </div>
                  </div>
                `;
              }).join("")}
            </div>
            <div class="flex flex-col gap-2">
              <button
                onclick="handleCopyAll()"
                class="w-full px-4 py-2 bg-gradient-to-r from-green-500 to-green-600 text-white font-medium rounded-lg hover:from-green-600 hover:to-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-all shadow-sm"
              >
                Copy All Sections
              </button>
              <button
                disabled
                title="Export functionality not implemented"
                class="w-full px-4 py-2 bg-gray-200 text-gray-500 font-medium rounded-lg cursor-not-allowed"
              >
                Download DOCX/PDF (stub)
              </button>
            </div>
          </div>
        </div>
      `}
    </div>
  `;
}

// ============================================================================
// Event Handlers
// ============================================================================

function navigateToStep(step) {
  if (step < 0 || step >= STEP_NAMES.length) return;
  state.activeStep = step;
  
  // Update step status: mark current step as in_progress if it's still todo
  if (state.stepStatus[step] === "todo") {
    state.stepStatus[step] = "in_progress";
  }
  
  renderLeftRail();
  renderStepPanel();
}

async function handleCreateSession() {
  const workflowType = document.getElementById("workflowType")?.value || "grant";
  state.workflowType = workflowType;

  let grant = "EDG";
  if (workflowType === "grant") {
    grant = document.getElementById("grantSelect")?.value || "EDG";
    state.grant = grant;
  }

  try {
    const body = { grant, company_name: "SmartGrant Pte Ltd" };
    const result = await apiPost("/v1/session", body);
    state.sessionId = result.session_id;
    persistState();

    // Auto-fetch checklist
    await handleLoadChecklist();

    // Transition to Step 1
    state.stepStatus[0] = "done";
    state.stepStatus[1] = "in_progress";
    navigateToStep(1);
  } catch (e) {
    alert(`Failed to create session: ${e.message}`);
  }
}

async function handleLoadChecklist() {
  if (!state.sessionId) return;
  try {
    const result = await apiGet(`/v1/session/${encodeURIComponent(state.sessionId)}/checklisttest`);
    const tasks = result.tasks || [];
    state.checklist.uploads = tasks.filter(t => t.type === "upload");
    state.checklist.drafts = tasks.filter(t => t.type === "draft");
    renderStepPanel();
  } catch (e) {
    alert(`Failed to load checklist: ${e.message}`);
  }
}

async function handleFileSelect(label, input) {
  // Guard: require session
  if (!state.sessionId) {
    alert("Create Run first before uploading evidence.");
    input.value = "";
    return;
  }

  // Extract file
  const file = input.files?.[0];
  if (!file) {
    return;
  }

  // Prevent double upload if already uploading
  if (state.evidenceUploadStatus[label] === "uploading") {
    alert("Upload already in progress. Please wait.");
    input.value = "";
    return;
  }

  try {
    // Set status to uploading
    state.evidenceUploadStatus[label] = "uploading";
    renderStepPanel();

    // Get SAS URL from broker
    const sasResp = await brokerPost("/api/upload/sas", {
      sid: state.sessionId,
      label: label,
      filename: file.name
    });

    // PUT file to blob storage
    const putRes = await fetch(sasResp.uploadUrl, {
      method: "PUT",
      headers: {
        "x-ms-blob-type": "BlockBlob",
        "Content-Type": file.type || "application/octet-stream"
      },
      body: file
    });

    if (!putRes.ok) {
      throw new Error(`Blob upload failed: HTTP ${putRes.status}`);
    }

    // Success: mark as uploaded
    state.evidenceUploadStatus[label] = "uploaded";
    renderStepPanel();

    // Optional non-blocking alert
    setTimeout(() => {
      alert("Uploaded. Parsing will take a few seconds; click Refresh Evidence Detected.");
    }, 100);

  } catch (error) {
    // Failure: mark as error
    state.evidenceUploadStatus[label] = "error";
    renderStepPanel();

    // Alert with error details
    alert(`Upload failed: ${error.message}\n\nPlease try again or check your connection.`);
    
    // Reset input to allow retry
    input.value = "";
  }
}

async function handleRefreshEvidence() {
  if (!state.sessionId) return;
  try {
    const result = await apiGet(`/v1/debug/evidence/${encodeURIComponent(state.sessionId)}?preview=120`);
    state.evidenceDetected = result.items || [];

    // Update upload status for detected items
    state.evidenceDetected.forEach(item => {
      if (state.evidenceUploadStatus[item.label] !== "detected") {
        state.evidenceUploadStatus[item.label] = "detected";
      }
    });

    // Re-render to update Evidence Inspector and status pills
    renderStepPanel();
  } catch (e) {
    alert(`Failed to refresh evidence: ${e.message}`);
  }
}

async function handleSaveFacts() {
  if (!state.sessionId) {
    alert("Create a session first");
    return;
  }

  const payload = {};
  const equity = document.getElementById("factEquity")?.value;
  const turnover = document.getElementById("factTurnover")?.value;
  const headcount = document.getElementById("factHeadcount")?.value;
  const extra = document.getElementById("factExtra")?.value;

  if (equity !== "") payload.local_equity_pct = Number(equity);
  if (turnover !== "") payload.turnover = Number(turnover);
  if (headcount !== "") payload.headcount = Number(headcount);

  if (extra) {
    try {
      payload.extra = JSON.parse(extra);
    } catch {
      // Try key:value parsing
      const lines = extra.split("\n").filter(l => l.trim());
      const obj = {};
      lines.forEach(line => {
        const [key, ...vals] = line.split(":");
        if (key && vals.length) {
          obj[key.trim()] = vals.join(":").trim();
        }
      });
      if (Object.keys(obj).length > 0) payload.extra = obj;
    }
  }

  try {
    if (Object.keys(payload).length > 0) {
      await apiPost(`/v1/session/${encodeURIComponent(state.sessionId)}/facts`, payload);
      state.facts = { ...state.facts, ...payload };
      persistState();
    }
  } catch (e) {
    alert(`Failed to save facts: ${e.message}`);
  }
}

async function handleValidate() {
  if (!state.sessionId) {
    alert("Create a session first");
    return;
  }

  // Save facts first
  await handleSaveFacts();

  try {
    const result = await apiPost(`/v1/session/${encodeURIComponent(state.sessionId)}/validate`, {});
    state.validationChecks = result.checks || [];
    state.stepStatus[2] = state.validationChecks.length === 0 ? "done" : "in_progress";
    renderLeftRail();
    renderStepPanel();
  } catch (e) {
    alert(`Failed to validate: ${e.message}`);
  }
}

function handleSaveAnchor() {
  const anchor = document.getElementById("solutionAnchor")?.value || "";
  const style = document.getElementById("sharedStyle")?.value || "Formal, outcome-oriented";
  const lengthLimit = parseInt(document.getElementById("sharedLengthLimit")?.value || "300", 10);
  const evidenceCap = parseInt(document.getElementById("sharedEvidenceCap")?.value || "0", 10);

  state.solutionAnchor = anchor;
  state.sharedDraft.style = style;
  state.sharedDraft.length_limit = lengthLimit;
  state.sharedDraft.evidence_char_cap = evidenceCap > 0 ? evidenceCap : 0;

  state.stepStatus[3] = "done";
  persistState();
  renderLeftRail();
  renderStepPanel();
}

async function handleDraftSection(sectionId, sectionVariant) {
  if (!state.sessionId) {
    alert("Create a session first");
    return;
  }
  if (!state.solutionAnchor) {
    alert("Save Solution Anchor first (Step 3)");
    return;
  }

  const evidenceSelect = document.getElementById("evidenceSelect");
  const selectedLabels = evidenceSelect ? Array.from(evidenceSelect.selectedOptions).map(o => o.value) : [];

  const inputs = {
    prompt: state.solutionAnchor,
    solution_anchor: state.solutionAnchor,
    style: state.sharedDraft.style,
    length_limit: state.sharedDraft.length_limit
  };

  if (state.sharedDraft.evidence_char_cap > 0) {
    inputs.evidence_char_cap = state.sharedDraft.evidence_char_cap;
  }

  if (selectedLabels.length > 0) {
    inputs.evidence_labels = selectedLabels;
  }

  const payload = {
    session_id: state.sessionId,
    section_id: sectionId,
    inputs
  };

  if (sectionVariant) {
    payload.section_variant = sectionVariant;
  }

  try {
    const result = await apiPost("/v1/draft", payload);
    state.outputs[sectionId] = result;

    // Update latest output display
    const latestOutput = document.getElementById("latestOutput");
    if (latestOutput) {
      latestOutput.textContent = result.output || "";
    }

    state.stepStatus[4] = "done";
    renderLeftRail();
    renderStepPanel();
  } catch (e) {
    alert(`Failed to draft section: ${e.message}`);
  }
}

async function handleDraftAll() {
  if (!state.sessionId) {
    alert("Create a session first");
    return;
  }
  if (!state.solutionAnchor) {
    alert("Save Solution Anchor first (Step 3)");
    return;
  }

  const draftTasks = state.checklist.drafts || [];
  if (draftTasks.length === 0) {
    alert("No draft tasks in checklist");
    return;
  }

  const evidenceSelect = document.getElementById("evidenceSelect");
  const selectedLabels = evidenceSelect ? Array.from(evidenceSelect.selectedOptions).map(o => o.value) : [];

  for (const task of draftTasks) {
    const sectionId = task.id;
    const sectionVariant = task.section_variant || null;

    try {
      await handleDraftSection(sectionId, sectionVariant);
      // Small delay to avoid overwhelming the API
      await new Promise(resolve => setTimeout(resolve, 500));
    } catch (e) {
      console.error(`Failed to draft ${sectionId}:`, e);
      // Continue with next section
    }
  }

  state.stepStatus[4] = "done";
  state.stepStatus[5] = "in_progress";
  renderLeftRail();
  renderStepPanel();
}

function handleCopySection(sectionId) {
  const output = state.outputs[sectionId];
  if (!output || !output.output) {
    alert("No output for this section");
    return;
  }
  navigator.clipboard.writeText(output.output).then(() => {
    alert("Copied to clipboard!");
    state.stepStatus[5] = "done";
    renderLeftRail();
  }).catch(e => {
    alert(`Failed to copy: ${e.message}`);
  });
}

function handleCopyAll() {
  const draftTasks = state.checklist.drafts || [];
  const sections = draftTasks
    .filter(t => state.outputs[t.id])
    .map(t => {
      const output = state.outputs[t.id];
      return `=== ${t.id}${t.section_variant ? ` (${t.section_variant})` : ""} ===\n\n${output.output || ""}`;
    });

  if (sections.length === 0) {
    alert("No sections have been drafted");
    return;
  }

  const fullText = sections.join("\n\n");
  navigator.clipboard.writeText(fullText).then(() => {
    alert("All sections copied to clipboard!");
    state.stepStatus[5] = "done";
    renderLeftRail();
  }).catch(e => {
    alert(`Failed to copy: ${e.message}`);
  });
}

// ============================================================================
// Initialization
// ============================================================================

async function init() {
  hydrateState();
  await loadEnvConfig();
  
  // If we have a session, try to load checklist
  if (state.sessionId) {
    try {
      await handleLoadChecklist();
      // Also refresh evidence if we're on step 1
      if (state.activeStep === 1) {
        await handleRefreshEvidence();
      }
    } catch (e) {
      console.warn("Could not load checklist for existing session:", e);
    }
  }
  
  renderLeftRail();
  renderStepPanel();
}

// Make functions globally available for onclick handlers
window.navigateToStep = navigateToStep;
window.handleCreateSession = handleCreateSession;
window.handleFileSelect = handleFileSelect;
window.handleRefreshEvidence = handleRefreshEvidence;
window.handleSaveFacts = handleSaveFacts;
window.handleValidate = handleValidate;
window.handleSaveAnchor = handleSaveAnchor;
window.handleDraftSection = handleDraftSection;
window.handleDraftAll = handleDraftAll;
window.handleCopySection = handleCopySection;
window.handleCopyAll = handleCopyAll;

// Initialize on load
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}

