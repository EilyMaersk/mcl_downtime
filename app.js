/* global XLSX, d3, topojson */
const STORAGE_KEY = "mcl-operational-downtime-records-v1";
const EDIT_KEY = "mcl-operational-downtime-edit-id";
const EXCEL_IDB_NAME = "mcl-operational-downtime-fs";
const EXCEL_IDB_STORE = "handles";
const EXCEL_HANDLE_KEY = "excel-file";
const EXCEL_SHEET_NAME = "Records";
const EXCEL_COLUMNS = [
  "id",
  "incidentReference",
  "incidentDate",
  "outageStart",
  "outageEnd",
  "outageDurationHours",
  "operationalLostHours",
  "siteId",
  "country",
  "region",
  "customer",
  "application",
  "incidentCategory",
  "rootCause",
  "impactDescription",
  "submitterName",
  "submitterEmail",
  "submittedAt",
  "status"
];
const supportsFileSystemAccess = typeof window !== "undefined" && "showSaveFilePicker" in window;

let excelFileHandle = null;
let worldTopo = null;
let mapResizeTimer = null;

const COUNTRY_NAME_ALIASES = {
  "United States": "United States of America"
};

const masterData = {
  sites: [
    { siteId: "CN0203", country: "China", region: "Greater China" },
    { siteId: "CN0204", country: "China", region: "Greater China" },
    { siteId: "CN0205", country: "China", region: "Greater China" },
    { siteId: "CN0206", country: "China", region: "Greater China" },
    { siteId: "CN0207", country: "China", region: "Greater China" },
    { siteId: "CN0208", country: "China", region: "Greater China" },
    { siteId: "CN0213", country: "China", region: "Greater China" },
    { siteId: "CN00131", country: "China", region: "Greater China" },
    { siteId: "PL1001", country: "Poland", region: "Europe" },
    { siteId: "DE2002", country: "Germany", region: "Europe" },
    { siteId: "MX3001", country: "Mexico", region: "North America" },
    { siteId: "US4105", country: "United States", region: "North America" }
  ],
  customers: [
    "Multiple Projects",
    "Retail Customer A",
    "Industrial Customer B",
    "Consumer Customer C"
  ],
  applications: [
    "LogiReport",
    "SAP EWM",
    "Warehouse RF",
    "Integration Layer",
    "Reporting Platform"
  ],
  incidentCategories: [
    "Outage",
    "Performance Degradation",
    "Technology Incident"
  ],
  rootCauses: [
    "Application defect",
    "Infrastructure",
    "Network",
    "Third-party dependency",
    "User access / device issue",
    "Unknown"
  ]
};

const demoRecords = [
  {
    id: createId(),
    incidentReference: "INC-2026-0610-A",
    incidentDate: "2026-06-10",
    outageStart: "2026-06-10T10:15",
    outageEnd: "2026-06-10T11:05",
    outageDurationHours: 0.83,
    operationalLostHours: 1.15,
    siteId: "CN0203",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Outage",
    rootCause: "Application defect",
    impactDescription: "Inbound and outbound operations were interrupted while LogiReport was unavailable.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-06-10T11:10",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0610-A",
    incidentDate: "2026-06-10",
    outageStart: "2026-06-10T10:15",
    outageEnd: "2026-06-10T11:05",
    outageDurationHours: 0.83,
    operationalLostHours: 1.15,
    siteId: "CN0206",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Outage",
    rootCause: "Application defect",
    impactDescription: "Team members were stood down during the outage and needed time to ramp operations back up.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-06-10T11:13",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0612-A",
    incidentDate: "2026-06-12",
    outageStart: "2026-06-12T14:15",
    outageEnd: "2026-06-12T15:00",
    outageDurationHours: 0.75,
    operationalLostHours: 1.5,
    siteId: "CN0204",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Performance Degradation",
    rootCause: "Infrastructure",
    impactDescription: "Warehouse teams experienced intermittent delays that doubled the operational impact versus the outage duration.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-06-12T15:05",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0616-A",
    incidentDate: "2026-06-16",
    outageStart: "2026-06-16T17:14",
    outageEnd: "2026-06-16T17:34",
    outageDurationHours: 0.33,
    operationalLostHours: 0.36,
    siteId: "CN0213",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Outage",
    rootCause: "Application defect",
    impactDescription: "Rapid issue recovery limited lost hours but affected same-day throughput.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-06-16T17:41",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0713-A",
    incidentDate: "2026-07-13",
    outageStart: "2026-07-13T07:35",
    outageEnd: "2026-07-13T09:00",
    outageDurationHours: 1.42,
    operationalLostHours: 4.42,
    siteId: "CN0204",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Outage",
    rootCause: "Application defect",
    impactDescription: "One hour of outage translated into extended operational recovery across the site.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-07-13T09:11",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0713-A",
    incidentDate: "2026-07-13",
    outageStart: "2026-07-13T09:00",
    outageEnd: "2026-07-13T09:45",
    outageDurationHours: 0.75,
    operationalLostHours: 3,
    siteId: "CN0208",
    country: "China",
    region: "Greater China",
    customer: "Multiple Projects",
    application: "LogiReport",
    incidentCategory: "Outage",
    rootCause: "Application defect",
    impactDescription: "Wave picking and outbound staging resumed slowly after the outage cleared.",
    submitterName: "China Ops Lead",
    submitterEmail: "china.ops@maersk.com",
    submittedAt: "2026-07-13T09:18",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0721-EU",
    incidentDate: "2026-07-21",
    outageStart: "2026-07-21T08:10",
    outageEnd: "2026-07-21T09:05",
    outageDurationHours: 0.92,
    operationalLostHours: 2.5,
    siteId: "PL1001",
    country: "Poland",
    region: "Europe",
    customer: "Retail Customer A",
    application: "SAP EWM",
    incidentCategory: "Technology Incident",
    rootCause: "Network",
    impactDescription: "Radio-frequency devices failed to reconnect quickly, creating prolonged picking delays.",
    submitterName: "Europe Ops Lead",
    submitterEmail: "europe.ops@maersk.com",
    submittedAt: "2026-07-21T09:22",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0801-NA",
    incidentDate: "2026-08-01",
    outageStart: "2026-08-01T06:40",
    outageEnd: "2026-08-01T07:25",
    outageDurationHours: 0.75,
    operationalLostHours: 1.9,
    siteId: "US4105",
    country: "United States",
    region: "North America",
    customer: "Industrial Customer B",
    application: "Warehouse RF",
    incidentCategory: "Performance Degradation",
    rootCause: "User access / device issue",
    impactDescription: "Users lost productive time logging back into devices after service restoration.",
    submitterName: "North America Ops Lead",
    submitterEmail: "na.ops@maersk.com",
    submittedAt: "2026-08-01T07:31",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0803-NA",
    incidentDate: "2026-08-03",
    outageStart: "2026-08-03T13:05",
    outageEnd: "2026-08-03T14:15",
    outageDurationHours: 1.17,
    operationalLostHours: 2.8,
    siteId: "MX3001",
    country: "Mexico",
    region: "North America",
    customer: "Consumer Customer C",
    application: "Integration Layer",
    incidentCategory: "Outage",
    rootCause: "Third-party dependency",
    impactDescription: "Inbound transaction failures disrupted receiving and delayed recovery after the incident closed.",
    submitterName: "North America Ops Lead",
    submitterEmail: "na.ops@maersk.com",
    submittedAt: "2026-08-03T14:23",
    status: "Submitted"
  },
  {
    id: createId(),
    incidentReference: "INC-2026-0804-EU",
    incidentDate: "2026-08-04",
    outageStart: "2026-08-04T09:30",
    outageEnd: "2026-08-04T10:20",
    outageDurationHours: 0.83,
    operationalLostHours: 1.6,
    siteId: "DE2002",
    country: "Germany",
    region: "Europe",
    customer: "Retail Customer A",
    application: "Reporting Platform",
    incidentCategory: "Technology Incident",
    rootCause: "Infrastructure",
    impactDescription: "Operational reporting delays slowed release prioritization and recovery sequencing.",
    submitterName: "Europe Ops Lead",
    submitterEmail: "europe.ops@maersk.com",
    submittedAt: "2026-08-04T10:29",
    status: "Submitted"
  }
];

const elements = {
  form: document.getElementById("record-form"),
  recordId: document.getElementById("record-id"),
  incidentReference: document.getElementById("incident-reference"),
  incidentDate: document.getElementById("incident-date"),
  siteId: document.getElementById("site-id"),
  country: document.getElementById("country"),
  region: document.getElementById("region"),
  customer: document.getElementById("customer"),
  application: document.getElementById("application"),
  incidentCategory: document.getElementById("incident-category"),
  rootCause: document.getElementById("root-cause"),
  outageStart: document.getElementById("outage-start"),
  outageEnd: document.getElementById("outage-end"),
  outageDuration: document.getElementById("outage-duration"),
  lostHours: document.getElementById("lost-hours"),
  submitterName: document.getElementById("submitter-name"),
  submitterEmail: document.getElementById("submitter-email"),
  impactDescription: document.getElementById("impact-description"),
  cancelEdit: document.getElementById("cancel-edit"),
  clearFilters: document.getElementById("clear-filters"),
  seedDemoData: document.getElementById("seed-demo-data"),
  exportCsv: document.getElementById("export-csv"),
  clearAllData: document.getElementById("clear-all-data"),
  filterMonth: document.getElementById("filter-month"),
  filterRegion: document.getElementById("filter-region"),
  filterCountry: document.getElementById("filter-country"),
  filterSite: document.getElementById("filter-site"),
  filterCustomer: document.getElementById("filter-customer"),
  filterApplication: document.getElementById("filter-application"),
  kpiGrid: document.getElementById("kpi-grid"),
  monthlyChart: document.getElementById("chart-monthly"),
  regionChart: document.getElementById("chart-region"),
  countryChart: document.getElementById("chart-country"),
  topSites: document.getElementById("top-sites"),
  topApplications: document.getElementById("top-applications"),
  recordsBody: document.getElementById("records-body"),
  emptyStateTemplate: document.getElementById("empty-state-template"),
  worldMap: document.getElementById("world-map"),
  linkExcel: document.getElementById("link-excel"),
  syncExcel: document.getElementById("sync-excel"),
  excelStatus: document.getElementById("excel-status")
};

const state = {
  records: loadRecords(),
  filters: {
    month: "",
    region: "",
    country: "",
    siteId: "",
    customer: "",
    application: ""
  }
};

initialize();

async function initialize() {
  await restoreExcelHandle();
  ensureSeedData();

  if (elements.form) {
    initFormPage();
  }
  if (elements.kpiGrid) {
    initMetricsPage();
  }
  if (elements.recordsBody) {
    initRecordsPage();
  }

  updateExcelStatus();
}

function ensureSeedData() {
  if (state.records.length === 0) {
    state.records = cloneDemoRecords();
    persistRecords();
  }
}

function initFormPage() {
  populateStaticOptions();
  bindFormEvents();
  setDefaultFormValues();
  loadPendingEdit();
}

function initMetricsPage() {
  bindFilterEvents();
  populateFilterOptions();
  renderDashboard();
  loadWorldMap();
  window.addEventListener("resize", () => {
    window.clearTimeout(mapResizeTimer);
    mapResizeTimer = window.setTimeout(() => renderWorldMap(getFilteredRecords()), 200);
  });
}

function initRecordsPage() {
  bindFilterEvents();
  bindRecordsEvents();
  populateFilterOptions();
  renderTable();
}

function populateStaticOptions() {
  populateSelect(elements.siteId, masterData.sites.map((site) => site.siteId), "Select site");
  populateSelect(elements.customer, masterData.customers, "Select customer");
  populateSelect(elements.application, masterData.applications, "Select application / system");
  populateSelect(elements.incidentCategory, masterData.incidentCategories, "Select category");
  populateSelect(elements.rootCause, masterData.rootCauses, "Select root cause");
}

function bindFormEvents() {
  elements.siteId.addEventListener("change", syncLocationFromSite);
  elements.outageStart.addEventListener("change", updateDurationFromTimes);
  elements.outageEnd.addEventListener("change", updateDurationFromTimes);
  elements.form.addEventListener("submit", handleFormSubmit);
  elements.form.addEventListener("reset", () => {
    window.setTimeout(() => {
      clearEditState();
      setDefaultFormValues();
    }, 0);
  });
  elements.cancelEdit.addEventListener("click", () => {
    elements.form.reset();
    clearEditState();
    setDefaultFormValues();
  });
}

function bindFilterEvents() {
  [
    elements.filterMonth,
    elements.filterRegion,
    elements.filterCountry,
    elements.filterSite,
    elements.filterCustomer,
    elements.filterApplication
  ].forEach((input) => {
    if (input) {
      input.addEventListener("change", handleFilterChange);
    }
  });

  if (elements.clearFilters) {
    elements.clearFilters.addEventListener("click", clearFilters);
  }
}

function bindRecordsEvents() {
  if (elements.seedDemoData) {
    elements.seedDemoData.addEventListener("click", resetDemoData);
  }
  if (elements.exportCsv) {
    elements.exportCsv.addEventListener("click", exportFilteredCsv);
  }
  if (elements.clearAllData) {
    elements.clearAllData.addEventListener("click", clearAllData);
  }
  if (elements.linkExcel) {
    elements.linkExcel.addEventListener("click", linkExcelFile);
  }
  if (elements.syncExcel) {
    elements.syncExcel.addEventListener("click", syncFromExcelFile);
  }
  elements.recordsBody.addEventListener("click", handleTableAction);
}

function loadPendingEdit() {
  const editId = window.localStorage.getItem(EDIT_KEY);
  if (!editId) {
    return;
  }

  window.localStorage.removeItem(EDIT_KEY);
  const record = state.records.find((entry) => entry.id === editId);
  if (record) {
    loadRecordIntoForm(record);
  }
}

function refreshData() {
  populateFilterOptions();
  if (elements.kpiGrid) {
    renderDashboard();
  }
  if (elements.recordsBody) {
    renderTable();
  }
}

function setDefaultFormValues() {
  const today = new Date();
  elements.incidentDate.value = today.toISOString().slice(0, 10);
  if (!elements.siteId.value) {
    elements.siteId.selectedIndex = 0;
  }
  syncLocationFromSite();
  elements.outageDuration.value = "";
}

function handleFormSubmit(event) {
  event.preventDefault();

  const outageDurationHours = calculateDurationHours(elements.outageStart.value, elements.outageEnd.value);
  if (outageDurationHours === null) {
    window.alert("Outage end must be later than outage start.");
    return;
  }
  const site = masterData.sites.find((entry) => entry.siteId === elements.siteId.value);
  if (!site) {
    window.alert("Please select a valid site.");
    return;
  }

  const record = {
    id: elements.recordId.value || createId(),
    incidentReference: elements.incidentReference.value.trim(),
    incidentDate: elements.incidentDate.value,
    outageStart: elements.outageStart.value,
    outageEnd: elements.outageEnd.value,
    outageDurationHours,
    operationalLostHours: roundNumber(Number(elements.lostHours.value)),
    siteId: site.siteId,
    country: site.country,
    region: site.region,
    customer: elements.customer.value,
    application: elements.application.value,
    incidentCategory: elements.incidentCategory.value,
    rootCause: elements.rootCause.value,
    impactDescription: elements.impactDescription.value.trim(),
    submitterName: elements.submitterName.value.trim(),
    submitterEmail: elements.submitterEmail.value.trim(),
    submittedAt: new Date().toISOString(),
    status: "Submitted"
  };

  if (elements.recordId.value) {
    state.records = state.records.map((existing) => existing.id === record.id ? record : existing);
  } else {
    state.records.unshift(record);
  }

  saveRecords().finally(() => {
    window.location.href = "records.html";
  });
}

function handleFilterChange() {
  state.filters = {
    month: elements.filterMonth.value,
    region: elements.filterRegion.value,
    country: elements.filterCountry.value,
    siteId: elements.filterSite.value,
    customer: elements.filterCustomer.value,
    application: elements.filterApplication.value
  };
  if (elements.kpiGrid) {
    renderDashboard();
  }
  if (elements.recordsBody) {
    renderTable();
  }
}

function clearFilters() {
  state.filters = {
    month: "",
    region: "",
    country: "",
    siteId: "",
    customer: "",
    application: ""
  };

  refreshData();
}

function resetDemoData() {
  if (!window.confirm("Replace the current local records with the demo dataset?")) {
    return;
  }

  state.records = cloneDemoRecords();
  saveRecords();
  state.filters = { month: "", region: "", country: "", siteId: "", customer: "", application: "" };
  refreshData();
}

function clearAllData() {
  if (!window.confirm("Delete all locally stored prototype records from this browser?")) {
    return;
  }

  state.records = [];
  saveRecords();
  state.filters = { month: "", region: "", country: "", siteId: "", customer: "", application: "" };
  refreshData();
}

function handleTableAction(event) {
  const actionButton = event.target.closest("button[data-action]");
  if (!actionButton) {
    return;
  }

  const { action, recordId } = actionButton.dataset;
  const record = state.records.find((entry) => entry.id === recordId);
  if (!record) {
    return;
  }

  if (action === "edit") {
    window.localStorage.setItem(EDIT_KEY, recordId);
    window.location.href = "index.html";
    return;
  }

  if (action === "delete" && window.confirm("Delete this record from the browser prototype?")) {
    state.records = state.records.filter((entry) => entry.id !== recordId);
    saveRecords();
    refreshData();
  }
}

function loadRecordIntoForm(record) {
  elements.recordId.value = record.id;
  elements.incidentReference.value = record.incidentReference;
  elements.incidentDate.value = record.incidentDate;
  elements.siteId.value = record.siteId;
  syncLocationFromSite();
  elements.customer.value = record.customer;
  elements.application.value = record.application;
  elements.incidentCategory.value = record.incidentCategory;
  elements.rootCause.value = record.rootCause;
  elements.outageStart.value = record.outageStart;
  elements.outageEnd.value = record.outageEnd;
  elements.outageDuration.value = record.outageDurationHours.toFixed(2);
  elements.lostHours.value = record.operationalLostHours.toFixed(2);
  elements.submitterName.value = record.submitterName;
  elements.submitterEmail.value = record.submitterEmail;
  elements.impactDescription.value = record.impactDescription;
  elements.cancelEdit.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function clearEditState() {
  elements.recordId.value = "";
  elements.cancelEdit.classList.add("hidden");
}

function syncLocationFromSite() {
  const site = masterData.sites.find((entry) => entry.siteId === elements.siteId.value);
  elements.country.value = site ? site.country : "";
  elements.region.value = site ? site.region : "";
}

function updateDurationFromTimes() {
  const duration = calculateDurationHours(elements.outageStart.value, elements.outageEnd.value);
  elements.outageDuration.value = duration === null ? "" : duration.toFixed(2);
}

function calculateDurationHours(startValue, endValue) {
  if (!startValue || !endValue) {
    return null;
  }

  const start = new Date(startValue);
  const end = new Date(endValue);
  const milliseconds = end.getTime() - start.getTime();
  if (Number.isNaN(milliseconds) || milliseconds <= 0) {
    return null;
  }

  return roundNumber(milliseconds / 3600000);
}

function loadRecords() {
  const raw = window.localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      throw new Error("Stored records are not an array.");
    }
    return parsed;
  } catch (error) {
    console.error("Failed to load stored prototype data.", error);
    window.alert("Stored prototype data could not be read and will be reset for this browser.");
    window.localStorage.removeItem(STORAGE_KEY);
    return [];
  }
}

function persistRecords() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state.records));
}

function saveRecords() {
  persistRecords();
  return writeRecordsToExcel();
}

function openHandleDb() {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(EXCEL_IDB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(EXCEL_IDB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function storeHandle(handle) {
  const db = await openHandleDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(EXCEL_IDB_STORE, "readwrite");
    tx.objectStore(EXCEL_IDB_STORE).put(handle, EXCEL_HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function readStoredHandle() {
  const db = await openHandleDb();
  const handle = await new Promise((resolve, reject) => {
    const tx = db.transaction(EXCEL_IDB_STORE, "readonly");
    const request = tx.objectStore(EXCEL_IDB_STORE).get(EXCEL_HANDLE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return handle;
}

async function ensureHandlePermission(handle, mode, { prompt = true } = {}) {
  const options = { mode };
  if ((await handle.queryPermission(options)) === "granted") {
    return true;
  }
  if (prompt && (await handle.requestPermission(options)) === "granted") {
    return true;
  }
  return false;
}

async function restoreExcelHandle() {
  if (!supportsFileSystemAccess) {
    return;
  }

  try {
    const handle = await readStoredHandle();
    if (!handle) {
      return;
    }

    excelFileHandle = handle;

    if (typeof XLSX === "undefined") {
      return;
    }

    if (await ensureHandlePermission(handle, "read", { prompt: false })) {
      const imported = await readRecordsFromExcel(handle);
      if (imported) {
        state.records = imported;
        persistRecords();
      }
    }
  } catch (error) {
    console.error("Could not restore the linked Excel file.", error);
  }
}

async function linkExcelFile() {
  if (!supportsFileSystemAccess) {
    window.alert("This browser cannot link a local Excel file. Use Chrome or Edge, or export a CSV instead.");
    return;
  }
  if (typeof XLSX === "undefined") {
    window.alert("The Excel library failed to load. Check your internet connection and reload the page.");
    return;
  }

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: "operational-downtime-records.xlsx",
      types: [
        {
          description: "Excel Workbook",
          accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] }
        }
      ]
    });

    if (!(await ensureHandlePermission(handle, "readwrite"))) {
      window.alert("Permission to write the Excel file was not granted.");
      return;
    }

    excelFileHandle = handle;
    await storeHandle(handle);

    const imported = await readRecordsFromExcel(handle);
    if (imported && imported.length) {
      const replace = window.confirm(
        `The selected file already has ${imported.length} record(s). ` +
        "Load them into the app? Cancel keeps the current records and overwrites the file."
      );
      if (replace) {
        state.records = imported;
        persistRecords();
      }
    }

    await writeRecordsToExcel();
    updateExcelStatus();
    refreshData();
    window.alert("Excel file linked. Records will now be saved to it.");
  } catch (error) {
    if (error && error.name === "AbortError") {
      return;
    }
    console.error("Could not link the Excel file.", error);
    window.alert("Could not link the Excel file.");
  }
}

async function syncFromExcelFile() {
  if (!excelFileHandle) {
    window.alert("No Excel file is linked yet. Use \"Link Excel file\" first.");
    return;
  }
  if (typeof XLSX === "undefined") {
    window.alert("The Excel library failed to load. Check your internet connection and reload the page.");
    return;
  }

  try {
    if (!(await ensureHandlePermission(excelFileHandle, "read"))) {
      window.alert("Permission to read the Excel file was not granted.");
      return;
    }

    const imported = await readRecordsFromExcel(excelFileHandle);
    state.records = imported || [];
    persistRecords();
    updateExcelStatus();
    refreshData();
    window.alert(`Loaded ${state.records.length} record(s) from the linked Excel file.`);
  } catch (error) {
    console.error("Could not read the linked Excel file.", error);
    window.alert("Could not read the linked Excel file.");
  }
}

async function readRecordsFromExcel(handle) {
  const file = await handle.getFile();
  if (!file || file.size === 0) {
    return [];
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[EXCEL_SHEET_NAME] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) {
    return [];
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map(rowToRecord).filter(Boolean);
}

async function writeRecordsToExcel() {
  if (!excelFileHandle || typeof XLSX === "undefined") {
    return;
  }

  try {
    if (!(await ensureHandlePermission(excelFileHandle, "readwrite"))) {
      return;
    }

    const rows = state.records.map(recordToRow);
    const sheet = XLSX.utils.json_to_sheet(rows, { header: EXCEL_COLUMNS });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, EXCEL_SHEET_NAME);
    const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });

    const writable = await excelFileHandle.createWritable();
    await writable.write(new Blob([output]));
    await writable.close();
    updateExcelStatus();
  } catch (error) {
    console.error("Could not write to the linked Excel file.", error);
  }
}

function recordToRow(record) {
  const row = {};
  EXCEL_COLUMNS.forEach((column) => {
    row[column] = record[column] ?? "";
  });
  return row;
}

function rowToRecord(row) {
  if (!row || !row.incidentDate) {
    return null;
  }

  return {
    id: String(row.id || createId()),
    incidentReference: String(row.incidentReference || ""),
    incidentDate: String(row.incidentDate || ""),
    outageStart: String(row.outageStart || ""),
    outageEnd: String(row.outageEnd || ""),
    outageDurationHours: roundNumber(Number(row.outageDurationHours) || 0),
    operationalLostHours: roundNumber(Number(row.operationalLostHours) || 0),
    siteId: String(row.siteId || ""),
    country: String(row.country || ""),
    region: String(row.region || ""),
    customer: String(row.customer || ""),
    application: String(row.application || ""),
    incidentCategory: String(row.incidentCategory || ""),
    rootCause: String(row.rootCause || ""),
    impactDescription: String(row.impactDescription || ""),
    submitterName: String(row.submitterName || ""),
    submitterEmail: String(row.submitterEmail || ""),
    submittedAt: String(row.submittedAt || ""),
    status: String(row.status || "Submitted")
  };
}

function updateExcelStatus() {
  if (!elements.excelStatus) {
    return;
  }

  if (!supportsFileSystemAccess) {
    elements.excelStatus.textContent = "Linked Excel saving needs Chrome or Edge. Records are stored in this browser.";
    if (elements.linkExcel) {
      elements.linkExcel.disabled = true;
    }
    if (elements.syncExcel) {
      elements.syncExcel.disabled = true;
    }
    return;
  }

  if (excelFileHandle) {
    elements.excelStatus.textContent = `Linked to ${excelFileHandle.name}. Saves also write to this file.`;
  } else {
    elements.excelStatus.textContent = "Not linked to an Excel file. Records are stored in this browser only.";
  }
}

function getFilteredRecords() {
  return state.records
    .filter((record) => {
      const recordMonth = record.incidentDate.slice(0, 7);
      return (
        (!state.filters.month || recordMonth === state.filters.month) &&
        (!state.filters.region || record.region === state.filters.region) &&
        (!state.filters.country || record.country === state.filters.country) &&
        (!state.filters.siteId || record.siteId === state.filters.siteId) &&
        (!state.filters.customer || record.customer === state.filters.customer) &&
        (!state.filters.application || record.application === state.filters.application)
      );
    })
    .sort((a, b) => b.incidentDate.localeCompare(a.incidentDate));
}

function renderDashboard() {
  const records = getFilteredRecords();
  renderKpis(records);
  renderBarChart(elements.monthlyChart, aggregate(records, "month"), formatMonthLabel);
  renderBarChart(elements.regionChart, aggregate(records, "region"), (label) => label);
  renderBarChart(elements.countryChart, aggregate(records, "country"), (label) => label);
  renderMiniRanking(elements.topSites, aggregate(records, "siteId"));
  renderMiniRanking(elements.topApplications, aggregate(records, "application"));
  renderWorldMap(records);
}

const KPI_ICONS = {
  warehouse: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21V8.5L12 4l9 4.5V21"/><path d="M2 21h20"/><rect x="7" y="13" width="10" height="8"/><path d="M7 17h10"/></svg>',
  box: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 4.5v9L12 21l-8-4.5v-9L12 3z"/><path d="M4 7.5l8 4.5 8-4.5"/><path d="M12 12v9"/><path d="M8 5.25l8 4.5"/></svg>',
  truck: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6h12v10H2z"/><path d="M14 9h4l3 3v4h-7z"/><circle cx="6.5" cy="18" r="2"/><circle cx="17.5" cy="18" r="2"/></svg>',
  gauge: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M3.5 18a9 9 0 1 1 17 0"/><path d="M12 14l4.5-3"/><circle cx="12" cy="14" r="1.3"/></svg>'
};

function renderKpis(records) {
  const totalLostHours = records.reduce((sum, record) => sum + record.operationalLostHours, 0);
  const totalOutageHours = records.reduce((sum, record) => sum + record.outageDurationHours, 0);
  const impactedSites = new Set(records.map((record) => record.siteId)).size;
  const averageLostHours = records.length ? totalLostHours / records.length : 0;
  const impactMultiplier = totalOutageHours ? totalLostHours / totalOutageHours : 0;
  const periodText = state.filters.month ? formatMonthLabel(state.filters.month) : "All months";

  const kpis = [
    {
      icon: KPI_ICONS.warehouse,
      label: "Total lost hours",
      value: formatHours(totalLostHours),
      subtext: `${records.length} site submissions in scope`
    },
    {
      icon: KPI_ICONS.box,
      label: "Total incidents",
      value: String(records.length),
      subtext: `${impactedSites} impacted sites`
    },
    {
      icon: KPI_ICONS.truck,
      label: "Average impact",
      value: formatHours(averageLostHours),
      subtext: "Average lost hours per submitted record"
    },
    {
      icon: KPI_ICONS.gauge,
      label: "Impact multiplier",
      value: `${roundNumber(impactMultiplier).toFixed(2)}x`,
      subtext: `${impactedSites} impacted sites | ${periodText}`
    }
  ];

  elements.kpiGrid.innerHTML = kpis.map((kpi) => `
    <article class="kpi-card">
      <span class="kpi-icon">${kpi.icon}</span>
      <div class="kpi-body">
        <span class="kpi-label">${escapeHtml(kpi.label)}</span>
        <span class="kpi-value">${escapeHtml(kpi.value)}</span>
        <span class="kpi-subtext">${escapeHtml(kpi.subtext)}</span>
      </div>
    </article>
  `).join("");
}

function renderBarChart(container, data, labelFormatter) {
  if (!data.length) {
    renderEmptyState(container);
    return;
  }

  const maxValue = Math.max(...data.map((entry) => entry.value));
  container.innerHTML = data.map((entry) => {
    const width = maxValue === 0 ? 0 : (entry.value / maxValue) * 100;
    return `
      <div class="chart-row">
        <span class="chart-label">${escapeHtml(labelFormatter(entry.label))}</span>
        <div class="chart-bar-track">
          <div class="chart-bar-fill" style="width: ${width}%;"></div>
        </div>
        <span class="chart-value">${escapeHtml(formatHours(entry.value))}</span>
      </div>
    `;
  }).join("");
}

function renderMiniRanking(container, data) {
  if (!data.length) {
    renderEmptyState(container);
    return;
  }

  container.innerHTML = data.slice(0, 5).map((entry, index) => `
    <div class="mini-row">
      <div>
        <strong>${index + 1}. ${escapeHtml(entry.label)}</strong>
        <span>${escapeHtml(formatHours(entry.value))} lost hours</span>
      </div>
      <span>${Math.round(entry.value * 100) / 100}</span>
    </div>
  `).join("");
}

async function loadWorldMap() {
  if (!elements.worldMap) {
    return;
  }
  if (typeof d3 === "undefined" || typeof topojson === "undefined") {
    elements.worldMap.innerHTML = '<p class="map-fallback">Map library could not be loaded. Check your internet connection and reload.</p>';
    return;
  }
  if (worldTopo) {
    renderWorldMap(getFilteredRecords());
    return;
  }

  try {
    worldTopo = await d3.json("https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json");
    renderWorldMap(getFilteredRecords());
  } catch (error) {
    console.error("Could not load the world map data.", error);
    elements.worldMap.innerHTML = '<p class="map-fallback">Map data could not be loaded.</p>';
  }
}

function renderWorldMap(records) {
  if (!elements.worldMap || !worldTopo || typeof d3 === "undefined" || typeof topojson === "undefined") {
    return;
  }

  const container = elements.worldMap;
  const width = container.clientWidth || 720;
  const height = Math.round(width * 0.52);

  const totals = new Map();
  records.forEach((record) => {
    totals.set(record.country, (totals.get(record.country) || 0) + record.operationalLostHours);
  });
  const maxValue = totals.size ? Math.max(...totals.values()) : 0;
  const color = d3.scaleSequential(d3.interpolateBlues).domain([0, maxValue || 1]);

  const valueForFeature = (name) => {
    if (totals.has(name)) {
      return totals.get(name);
    }
    for (const [appName, mapName] of Object.entries(COUNTRY_NAME_ALIASES)) {
      if (mapName === name) {
        return totals.get(appName) || 0;
      }
    }
    return 0;
  };

  const features = topojson.feature(worldTopo, worldTopo.objects.countries).features;
  const projection = d3.geoNaturalEarth1().fitSize([width, height], { type: "Sphere" });
  const path = d3.geoPath(projection);

  container.innerHTML = "";
  const svg = d3.select(container)
    .append("svg")
    .attr("viewBox", `0 0 ${width} ${height}`)
    .attr("role", "img");

  svg.append("path")
    .attr("class", "sphere")
    .attr("d", path({ type: "Sphere" }));

  const tooltip = d3.select(container)
    .append("div")
    .attr("class", "map-tooltip");

  svg.append("g")
    .selectAll("path")
    .data(features)
    .join("path")
    .attr("d", path)
    .attr("class", (feature) => valueForFeature(feature.properties.name) > 0 ? "country has-value" : "country")
    .attr("fill", (feature) => {
      const value = valueForFeature(feature.properties.name);
      return value > 0 ? color(value) : null;
    })
    .on("mousemove", (event, feature) => {
      const value = valueForFeature(feature.properties.name);
      const bounds = container.getBoundingClientRect();
      tooltip
        .style("left", `${event.clientX - bounds.left}px`)
        .style("top", `${event.clientY - bounds.top}px`)
        .style("opacity", 1)
        .html(`<strong>${escapeHtml(feature.properties.name)}</strong><br>${escapeHtml(formatHours(value))} lost hours`);
    })
    .on("mouseleave", () => tooltip.style("opacity", 0));

  renderMapLegend(container, color, maxValue);
}

function renderMapLegend(container, color, maxValue) {
  if (maxValue <= 0) {
    return;
  }

  const steps = 5;
  const legend = d3.select(container).append("div").attr("class", "map-legend");
  legend.append("span").text("0h");
  const swatches = legend.append("div").attr("class", "swatches");
  for (let i = 1; i <= steps; i += 1) {
    swatches.append("span")
      .attr("class", "swatch")
      .style("background", color((maxValue / steps) * i));
  }
  legend.append("span").text(formatHours(maxValue));
}

function renderTable() {
  const records = getFilteredRecords();
  if (!records.length) {
    elements.recordsBody.innerHTML = `
      <tr>
        <td colspan="11">${elements.emptyStateTemplate.innerHTML}</td>
      </tr>
    `;
    return;
  }

  elements.recordsBody.innerHTML = records.map((record) => `
    <tr>
      <td>${escapeHtml(record.incidentDate)}</td>
      <td>${escapeHtml(record.incidentReference)}</td>
      <td>${escapeHtml(record.siteId)}</td>
      <td>${escapeHtml(record.country)}</td>
      <td>${escapeHtml(record.application)}</td>
      <td>${escapeHtml(formatHours(record.outageDurationHours))}</td>
      <td>${escapeHtml(formatHours(record.operationalLostHours))}</td>
      <td>${escapeHtml(record.incidentCategory)}</td>
      <td>${escapeHtml(record.rootCause)}</td>
      <td>${escapeHtml(record.submitterName)}</td>
      <td class="actions-cell">
        <button type="button" class="table-button edit" data-action="edit" data-record-id="${record.id}">Edit</button>
        <button type="button" class="table-button delete" data-action="delete" data-record-id="${record.id}">Delete</button>
      </td>
    </tr>
  `).join("");
}

function populateFilterOptions() {
  if (!elements.filterMonth) {
    return;
  }

  const months = [...new Set(state.records.map((record) => record.incidentDate.slice(0, 7)))].sort().reverse();
  const regions = [...new Set(state.records.map((record) => record.region))].sort();
  const countries = [...new Set(state.records.map((record) => record.country))].sort();
  const sites = [...new Set(state.records.map((record) => record.siteId))].sort();
  const customers = [...new Set(state.records.map((record) => record.customer))].sort();
  const applications = [...new Set(state.records.map((record) => record.application))].sort();

  populateSelect(elements.filterMonth, months, "All months", state.filters.month, formatMonthLabel);
  populateSelect(elements.filterRegion, regions, "All regions", state.filters.region);
  populateSelect(elements.filterCountry, countries, "All countries", state.filters.country);
  populateSelect(elements.filterSite, sites, "All sites", state.filters.siteId);
  populateSelect(elements.filterCustomer, customers, "All customers", state.filters.customer);
  populateSelect(elements.filterApplication, applications, "All applications", state.filters.application);
}

function populateSelect(select, values, emptyLabel, currentValue = "", formatter = (value) => value) {
  const options = [`<option value="">${escapeHtml(emptyLabel)}</option>`];
  values.forEach((value) => {
    options.push(`<option value="${escapeAttribute(value)}">${escapeHtml(formatter(value))}</option>`);
  });
  select.innerHTML = options.join("");
  if (currentValue) {
    select.value = currentValue;
  }
}

function aggregate(records, dimension) {
  const buckets = new Map();
  records.forEach((record) => {
    const label = dimension === "month" ? record.incidentDate.slice(0, 7) : record[dimension];
    buckets.set(label, (buckets.get(label) || 0) + record.operationalLostHours);
  });

  return [...buckets.entries()]
    .map(([label, value]) => ({ label, value: roundNumber(value) }))
    .sort((a, b) => dimension === "month"
      ? a.label.localeCompare(b.label)
      : b.value - a.value || a.label.localeCompare(b.label));
}

function renderEmptyState(container) {
  container.innerHTML = elements.emptyStateTemplate.innerHTML;
}

function exportFilteredCsv() {
  const records = getFilteredRecords();
  if (!records.length) {
    window.alert("There are no filtered records to export.");
    return;
  }

  const columns = [
    "incidentDate",
    "incidentReference",
    "siteId",
    "country",
    "region",
    "customer",
    "application",
    "incidentCategory",
    "rootCause",
    "outageStart",
    "outageEnd",
    "outageDurationHours",
    "operationalLostHours",
    "submitterName",
    "submitterEmail",
    "impactDescription",
    "submittedAt"
  ];

  const rows = [
    columns.join(","),
    ...records.map((record) => columns.map((column) => csvEscape(record[column])).join(","))
  ];

  const blob = new Blob([rows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "operational-downtime-export.csv";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function createId() {
  return `rec-${Math.random().toString(36).slice(2, 10)}-${Date.now().toString(36)}`;
}

function cloneDemoRecords() {
  return demoRecords.map((record) => ({ ...record }));
}

function roundNumber(value) {
  return Math.round(value * 100) / 100;
}

function formatHours(value) {
  return `${roundNumber(value).toFixed(2)}h`;
}

function formatMonthLabel(value) {
  const [year, month] = value.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return date.toLocaleString("en", { month: "short", year: "numeric" });
}

function csvEscape(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll("\"", "\"\"")}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}
