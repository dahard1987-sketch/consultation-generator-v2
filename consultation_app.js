const firebaseConfig = {
  apiKey: "AIzaSyCzwqK-ylC2Rbi57z9p9owWSbzeYOBm8Yk",
  authDomain: "cols-doc-maker.firebaseapp.com",
  projectId: "cols-doc-maker",
  storageBucket: "cols-doc-maker.firebasestorage.app",
  messagingSenderId: "824934533685",
  appId: "1:824934533685:web:8b14f84290274912f2ff7c",
  measurementId: "G-C4EQTBKDP8"
};

const $ = (id) => document.getElementById(id);
const FIELD_LABELS = {
  roster: "학생/반 목록",
  scores: "성적 입력",
  attitude: "수업태도",
  bookHomework: "교재숙제",
  readi: "READi",
  alex: "Alex",
  vocabTest: "지면 단어시험 결과",
  score: "성적 분석",
  lowScoreCare: "성적 저조자 관리",
  specialNote: "특이사항",
  summary: "전체 진행표"
};
const INCLUDE_FIELD_DEFS = [
  ["attitude", "수업태도"],
  ["bookHomework", "교재숙제"],
  ["readi", "READi"],
  ["alex", "Alex"],
  ["vocabTest", "지면 단어시험 결과"],
  ["score", "성적 분석"],
  ["lowScoreCare", "성적 저조자 관리"],
  ["specialNote", "특이사항"]
];
const RATING_OPTIONS = ["", "우수", "보통", "미흡"];
const SCORE_COLUMNS = [["lc", "LC"], ["rc", "RC"], ["vo", "VO"], ["gr", "GR"], ["total", "총점"]];
const CONSULTATION_LABELS = {
  "1차정기": "1차 정기",
  "2차정기": "2차 정기",
  "3차정기": "3차 정기",
  "4차정기": "4차 정기",
  "학기상담": "학기 상담",
  "보충상담": "보충 상담"
};
const SEMESTER_LABELS = { SP: "봄학기", SU: "여름학기", FA: "가을학기", WI: "겨울학기" };
const SEMESTER_ORDER = ["SP", "SU", "FA", "WI"];

let firebaseApp = null;
let db = null;
let auth = null;
let firebaseReady = false;
let currentUser = null;
let fb = {};
let googleProvider = null;

const state = {
  activeCriterion: "roster",
  currentPage: 0,
  search: "",
  classFilter: "",
  completionFilter: "all",
  selectedStudentId: ""
};

function createDefaultSchedule() {
  return {
    "2D": { nextTestName: "", nextTestDate: "", eopAlexAward: "", marketDay: "", summerStart: "", termEnd: "" },
    "3D": { nextTestName: "", nextTestDate: "", eopAlexAward: "", marketDay: "", summerStart: "", termEnd: "" }
  };
}

function createDefaultProject() {
  const now = new Date().toISOString();
  return {
    version: 3,
    projectName: "2026 봄학기 3차 정기 MT2 상담",
    createdAt: now,
    updatedAt: now,
    settings: {
      teacherName: "",
      year: "26",
      semester: "SP",
      nextSemester: "SU",
      consultationType: "3차정기",
      testType: "MT2",
      pageSize: 20,
      includeFields: {
        attitude: true,
        bookHomework: true,
        readi: true,
        alex: true,
        vocabTest: true,
        score: true,
        lowScoreCare: true,
        specialNote: true
      },
      extraFields: [],
      schedule: createDefaultSchedule()
    },
    students: [],
    sync: {
      projectId: "",
      settingsDirty: false,
      dirtyStudentIds: {},
      deletedStudentIds: {},
      hasUnsavedChanges: false,
      lastServerSavedAt: "",
      lastServerLoadedAt: "",
      firebaseEnabled: false
    }
  };
}

let project = createDefaultProject();

function safeParseJson(text, fallback = null) {
  try {
    return JSON.parse(text);
  } catch {
    return fallback;
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function showToast(message) {
  const el = $("toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => el.classList.remove("show"), 2200);
}

function debounce(fn, wait = 120) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function makeSafeId(value) {
  return String(value || "").trim().replace(/\s+/g, "_").replace(/[^\w가-힣-]/g, "");
}

function displayConsultationType(value) {
  return CONSULTATION_LABELS[value] || value || "";
}

function semesterText(year, semester) {
  return `${year || ""}${SEMESTER_LABELS[semester] || semester || ""}`.trim();
}

function getAutoNextSemester(semester) {
  const idx = Math.max(0, SEMESTER_ORDER.indexOf(semester));
  return SEMESTER_ORDER[(idx + 1) % SEMESTER_ORDER.length];
}

function makeAutoProjectNameFromSettings(settings) {
  return `${semesterText(settings.year, settings.semester)} ${displayConsultationType(settings.consultationType)} ${settings.testType || ""} 상담`
    .replace(/\s+/g, " ")
    .trim();
}

function makeAutoProjectName() {
  return makeAutoProjectNameFromSettings(project.settings);
}

function makeProjectId(targetProject = project) {
  const s = targetProject.settings || {};
  return [
    makeSafeId(s.teacherName || "unknown"),
    makeSafeId(s.year || "year"),
    makeSafeId(s.semester || "semester"),
    makeSafeId(s.nextSemester || "next"),
    makeSafeId(s.testType || "test"),
    makeSafeId(s.consultationType || "consultation")
  ].join("__");
}

function getLocalStorageKey() {
  return `consultationProject__${makeProjectId(project)}`;
}

function structuredCloneSafe(value) {
  if (typeof structuredClone === "function") return structuredClone(value);
  return JSON.parse(JSON.stringify(value));
}

function normalizeRatingMemo(raw) {
  if (!raw) return { rating: "", memo: "" };
  if (typeof raw === "string") return { rating: raw, memo: "" };
  return { rating: String(raw.rating || "").trim(), memo: String(raw.memo || "").trim() };
}

function normalizeVocabTest(raw) {
  if (!raw) return { rating: "", percent: "" };
  if (typeof raw === "string") {
    const match = raw.match(/(우수|보통|미흡)\s*\(?\s*(\d+(?:\.\d+)?)?%?\s*\)?/);
    return { rating: match?.[1] || "", percent: match?.[2] || "" };
  }
  return { rating: String(raw.rating || "").trim(), percent: String(raw.percent || "").replace(/%/g, "").trim() };
}

function normalizeStudent(raw, index = 0) {
  const name = String(raw?.name || "").trim();
  const className = String(raw?.className || raw?.class || "").trim();
  const classType = raw?.classType || detectClassType(className);
  const id = raw?.id || makeStudentId(className, name, index);
  return {
    id,
    className,
    name,
    classType,
    attitude: String(raw?.attitude || "").trim(),
    bookHomework: normalizeRatingMemo(raw?.bookHomework),
    readi: normalizeRatingMemo(raw?.readi),
    alex: normalizeRatingMemo(raw?.alex),
    vocabTest: normalizeVocabTest(raw?.vocabTest),
    scores: raw?.scores || {},
    scoreAnalysis: {
      good: String(raw?.scoreAnalysis?.good || raw?.scores?.[project.settings.testType || "MT2"]?.good || "").trim(),
      problem: String(raw?.scoreAnalysis?.problem || raw?.scores?.[project.settings.testType || "MT2"]?.problem || "").trim(),
      improvement: String(raw?.scoreAnalysis?.improvement || raw?.scores?.[project.settings.testType || "MT2"]?.improvement || "").trim()
    },
    lowScoreCare: String(raw?.lowScoreCare || "").trim(),
    extraFieldValues: raw?.extraFieldValues || {},
    specialNote: String(raw?.specialNote || "").trim(),
    manualComplete: Boolean(raw?.manualComplete),
    _dirty: Boolean(raw?._dirty),
    _lastSavedAt: raw?._lastSavedAt || ""
  };
}

function normalizeProjectShape(raw) {
  const base = createDefaultProject();
  const merged = {
    ...base,
    ...raw,
    settings: {
      ...base.settings,
      ...(raw?.settings || {}),
      includeFields: { ...base.settings.includeFields, ...(raw?.settings?.includeFields || {}) },
      schedule: {
        ...base.settings.schedule,
        ...(raw?.settings?.schedule || {}),
        "2D": { ...base.settings.schedule["2D"], ...(raw?.settings?.schedule?.["2D"] || {}) },
        "3D": { ...base.settings.schedule["3D"], ...(raw?.settings?.schedule?.["3D"] || {}) }
      },
      extraFields: raw?.settings?.extraFields || []
    },
    sync: {
      ...base.sync,
      ...(raw?.sync || {}),
      dirtyStudentIds: raw?.sync?.dirtyStudentIds || {},
      deletedStudentIds: raw?.sync?.deletedStudentIds || {}
    },
    students: Array.isArray(raw?.students) ? raw.students : []
  };
  if (!merged.settings.nextSemester) merged.settings.nextSemester = getAutoNextSemester(merged.settings.semester);
  merged.version = 3;
  merged.students = merged.students.map((student, index) => normalizeStudent(student, index));
  merged.projectName = merged.projectName || makeAutoProjectNameFromSettings(merged.settings);
  return merged;
}

function detectClassType(className) {
  const text = String(className || "").toUpperCase();
  if (/\b3D\b|MWF|MON|WED|FRI/.test(text)) return "3D";
  if (/\b2D\b|TTH|TUE|THU/.test(text)) return "2D";
  return "2D";
}

function makeStudentId(className, name, index) {
  return `${makeSafeId(className || "class")}__${makeSafeId(name || "student")}__${index}`;
}

function getScore(student) {
  const testType = project.settings.testType || "MT2";
  if (!student.scores) student.scores = {};
  if (!student.scores[testType]) student.scores[testType] = { lc: "", rc: "", vo: "", gr: "", total: "" };
  return student.scores[testType];
}

function getStudentByNameAndClass(name, className = "") {
  const candidates = project.students.filter((student) => student.name === name);
  if (!className) return candidates[0];
  return candidates.find((student) => student.className === className);
}

function studentHasAnyInput(student) {
  const score = getScore(student);
  if (student.attitude) return true;
  if ([student.bookHomework, student.readi, student.alex].some((item) => item?.rating || item?.memo)) return true;
  if (student.vocabTest?.rating || student.vocabTest?.percent) return true;
  if ([student.lowScoreCare, student.specialNote, student.scoreAnalysis?.good, student.scoreAnalysis?.problem, student.scoreAnalysis?.improvement].some(Boolean)) return true;
  if (Object.values(score).some(Boolean)) return true;
  return Object.values(student.extraFieldValues || {}).some(Boolean);
}

function getStudentProgressState(student) {
  if (student.manualComplete) return "complete";
  if (studentHasAnyInput(student)) return "in-progress";
  return "empty";
}

function setLocalStatus(message) {
  $("localSaveStatus").textContent = message;
}

function setServerStatus(message) {
  $("serverSaveStatus").textContent = message;
}

function renderDirtyCount() {
  const count = Object.keys(project.sync.dirtyStudentIds || {}).length;
  $("dirtyCount").textContent = `변경된 학생 ${count}명`;
  $("unsavedBadge").textContent = project.sync.hasUnsavedChanges ? "서버 저장 필요" : "서버 저장 완료";
  $("unsavedBadge").className = `badge ${project.sync.hasUnsavedChanges ? "warn" : "ok"}`;
}

function renderFirebaseBadge() {
  const badge = $("firebaseBadge");
  if (!firebaseReady) {
    badge.textContent = "서버 저장 필요 / 미로그인";
    badge.className = "badge warn";
  } else if (currentUser) {
    badge.textContent = `서버 저장 가능 / ${currentUser.displayName || currentUser.email || "로그인됨"}`;
    badge.className = "badge ok";
  } else {
    badge.textContent = "서버 저장 가능 / 미로그인";
    badge.className = "badge warn";
  }
}

function renderSaveStatus() {
  setLocalStatus("로컬 저장됨");
  if (project.sync.hasUnsavedChanges) {
    setServerStatus("서버 저장 필요");
  } else if (project.sync.lastServerSavedAt) {
    setServerStatus(`서버 저장 완료 ${project.sync.lastServerSavedAt.slice(0, 19).replace("T", " ")}`);
  } else if (!firebaseReady) {
    setServerStatus("서버 미연결");
  } else {
    setServerStatus("서버 저장 대기");
  }
  $("lastSavedAtView").textContent = project.sync.lastServerSavedAt ? project.sync.lastServerSavedAt.slice(0, 19).replace("T", " ") : "-";
  renderDirtyCount();
  renderFirebaseBadge();
}

const saveProjectToLocalStorage = debounce(() => {
  project.updatedAt = new Date().toISOString();
  project.sync.projectId = makeProjectId(project);
  localStorage.setItem(getLocalStorageKey(), JSON.stringify(project));
  setLocalStatus("로컬 저장됨");
  renderProjectId();
}, 60);

function saveProjectToLocalStorageNow() {
  project.updatedAt = new Date().toISOString();
  project.sync.projectId = makeProjectId(project);
  localStorage.setItem(getLocalStorageKey(), JSON.stringify(project));
  setLocalStatus("로컬 저장됨");
  renderProjectId();
}

function markStudentDirty(student, { rerender = false } = {}) {
  student._dirty = true;
  project.sync.dirtyStudentIds[student.id] = true;
  project.sync.hasUnsavedChanges = true;
  saveProjectToLocalStorage();
  renderSaveStatus();
  renderPreview();
  renderProgressSidebar();
  if (rerender) renderAll({ keepPage: true });
}

function markSettingsDirty({ rerender = true } = {}) {
  project.sync.settingsDirty = true;
  project.sync.hasUnsavedChanges = true;
  saveProjectToLocalStorage();
  renderSaveStatus();
  if (rerender) renderAll({ keepPage: true });
}

function renderProjectId() {
  $("projectIdView").textContent = makeProjectId(project);
}

function syncSettingsFromInputs({ autoNext = false } = {}) {
  const prevAuto = getAutoNextSemester(project.settings.semester);
  const currentNext = project.settings.nextSemester;
  project.settings.teacherName = $("teacherName").value.trim();
  project.settings.year = $("year").value.trim();
  project.settings.semester = $("semester").value;
  project.settings.testType = $("testType").value;
  project.settings.consultationType = $("consultationType").value;
  project.settings.pageSize = Number($("pageSize").value) || 20;
  if (autoNext || !currentNext || currentNext === prevAuto) {
    project.settings.nextSemester = getAutoNextSemester(project.settings.semester);
  } else {
    project.settings.nextSemester = $("nextSemester").value;
  }
  project.projectName = $("projectName").value.trim() || makeAutoProjectName();
}

function bindEvents() {
  ["teacherName", "year", "semester", "testType", "consultationType", "projectName", "pageSize", "nextSemester"].forEach((id) => {
    $(id).addEventListener("input", () => {
      syncSettingsFromInputs({ autoNext: id === "semester" });
      state.currentPage = 0;
      markSettingsDirty();
    });
    $(id).addEventListener("change", () => {
      syncSettingsFromInputs({ autoNext: id === "semester" });
      state.currentPage = 0;
      markSettingsDirty();
    });
  });

  $("searchInput").addEventListener("input", () => {
    state.search = $("searchInput").value.trim();
    state.currentPage = 0;
    renderStudents();
    renderPreviewStudentSelect();
  });
  $("classFilter").addEventListener("change", () => {
    state.classFilter = $("classFilter").value;
    state.currentPage = 0;
    renderStudents();
    renderPreviewStudentSelect();
  });
  $("completionFilter").addEventListener("change", () => {
    state.completionFilter = $("completionFilter").value;
    state.currentPage = 0;
    renderStudents();
  });
  $("prevPageBtn").addEventListener("click", () => {
    state.currentPage = Math.max(0, state.currentPage - 1);
    renderStudents();
  });
  $("nextPageBtn").addEventListener("click", () => {
    state.currentPage = Math.min(getTotalPages() - 1, state.currentPage + 1);
    renderStudents();
  });
  $("previewStudentSelect").addEventListener("change", () => {
    state.selectedStudentId = $("previewStudentSelect").value;
    renderPreview();
  });
  $("copyCurrentBtn").addEventListener("click", copyCurrentConsultationText);
  $("copyAllBtn").addEventListener("click", copyAllConsultationTexts);
  $("saveJsonBtn").addEventListener("click", saveProjectAsJSON);
  $("loadJsonBtn").addEventListener("click", () => $("jsonFileInput").click());
  $("jsonFileInput").addEventListener("change", loadProjectFromJSONFile);
  $("scoreFileInput").addEventListener("change", importScoresFromFile);
  $("saveServerBtn").addEventListener("click", saveProjectToFirebase);
  $("loadServerBtn").addEventListener("click", loadProjectFromFirebase);
  $("newProjectBtn").addEventListener("click", createNewProject);
  $("clearProjectDataBtn").addEventListener("click", clearCurrentProjectData);
  $("addExtraFieldBtn").addEventListener("click", addExtraField);
  $("showLocalProjectsBtn").addEventListener("click", showLocalProjectsModal);
  $("closeLocalProjectsBtn").addEventListener("click", () => $("localProjectsModal").classList.remove("show"));
  $("localProjectsModal").addEventListener("click", (event) => {
    if (event.target.id === "localProjectsModal") $("localProjectsModal").classList.remove("show");
  });
  $("loginBtn").addEventListener("click", loginWithGoogle);
  $("logoutBtn").addEventListener("click", logout);
  $("exportTxtVerticalBtn").addEventListener("click", () => exportConsultationTxt("vertical"));
  $("exportTxtTsvBtn").addEventListener("click", () => exportConsultationTxt("tsv"));

  const moreBtn = $("moreBtn");
  const moreMenu = $("moreMenu");
  moreBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    const open = moreMenu.classList.toggle("show");
    moreBtn.setAttribute("aria-expanded", String(open));
  });
  document.addEventListener("click", (event) => {
    if (!moreMenu.contains(event.target) && event.target !== moreBtn) {
      moreMenu.classList.remove("show");
      moreBtn.setAttribute("aria-expanded", "false");
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      moreMenu.classList.remove("show");
      moreBtn.setAttribute("aria-expanded", "false");
      $("localProjectsModal").classList.remove("show");
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      saveProjectToFirebase();
    }
  });
  window.addEventListener("beforeunload", (event) => {
    if (project.sync?.hasUnsavedChanges) {
      event.preventDefault();
      event.returnValue = "";
    }
  });
}

function renderAll({ keepPage = false } = {}) {
  if (!keepPage) state.currentPage = 0;
  renderSettingsInputs();
  renderProjectId();
  renderIncludeFields();
  renderScheduleInputs();
  renderExtraFieldsList();
  renderTabs();
  renderClassFilter();
  renderBulkTools();
  renderStudents();
  renderPreviewStudentSelect();
  renderPreview();
  renderSaveStatus();
  renderProgressSidebar();
  $("studentCountBadge").textContent = `학생 ${project.students.length}명`;
}

function renderSettingsInputs() {
  $("teacherName").value = project.settings.teacherName || "";
  $("year").value = project.settings.year || "26";
  $("semester").value = project.settings.semester || "SP";
  $("nextSemester").value = project.settings.nextSemester || getAutoNextSemester(project.settings.semester || "SP");
  $("consultationType").value = project.settings.consultationType || "3차정기";
  $("testType").value = project.settings.testType || "MT2";
  $("projectName").value = project.projectName || makeAutoProjectName();
  $("pageSize").value = String(project.settings.pageSize || 20);
  $("searchInput").value = state.search;
  $("completionFilter").value = state.completionFilter;
}

function renderIncludeFields() {
  const container = $("includeFields");
  container.innerHTML = "";
  INCLUDE_FIELD_DEFS.forEach(([key, label]) => {
    const item = document.createElement("label");
    item.className = "check-item";
    item.innerHTML = `<input type="checkbox" data-field="${key}"> <span>${escapeHtml(label)}</span>`;
    const input = item.querySelector("input");
    input.checked = Boolean(project.settings.includeFields[key]);
    input.addEventListener("change", () => {
      project.settings.includeFields[key] = input.checked;
      markSettingsDirty();
    });
    container.appendChild(item);
  });
}

function renderScheduleInputs() {
  const map = [
    ["schedule3DNextTestName", "3D", "nextTestName"],
    ["schedule3DNextTestDate", "3D", "nextTestDate"],
    ["schedule3DEopAlexAward", "3D", "eopAlexAward"],
    ["schedule3DMarketDay", "3D", "marketDay"],
    ["schedule3DSummerStart", "3D", "summerStart"],
    ["schedule3DTermEnd", "3D", "termEnd"],
    ["schedule2DNextTestName", "2D", "nextTestName"],
    ["schedule2DNextTestDate", "2D", "nextTestDate"],
    ["schedule2DEopAlexAward", "2D", "eopAlexAward"],
    ["schedule2DMarketDay", "2D", "marketDay"],
    ["schedule2DSummerStart", "2D", "summerStart"],
    ["schedule2DTermEnd", "2D", "termEnd"]
  ];
  map.forEach(([id, type, key]) => {
    const el = $(id);
    el.value = project.settings.schedule?.[type]?.[key] || "";
    const handler = () => {
      project.settings.schedule[type][key] = el.value.trim();
      markSettingsDirty({ rerender: false });
      renderPreview();
    };
    el.oninput = handler;
    el.onchange = handler;
  });
}

function renderTabs() {
  const baseTabs = ["roster", "scores", "attitude", "bookHomework", "readi", "alex", "vocabTest", "score", "lowScoreCare", "specialNote"];
  const tabs = baseTabs.filter((key) => key === "roster" || key === "scores" || project.settings.includeFields[key]);
  project.settings.extraFields.forEach((field) => tabs.push(`extra:${field.id}`));
  tabs.push("summary");
  if (!tabs.includes(state.activeCriterion)) state.activeCriterion = tabs[0] || "roster";
  const container = $("criterionTabs");
  container.innerHTML = "";
  tabs.forEach((key) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `tab ${state.activeCriterion === key ? "active" : ""}`;
    btn.textContent = getTabLabel(key);
    btn.addEventListener("click", () => {
      state.activeCriterion = key;
      state.currentPage = 0;
      renderTabs();
      renderBulkTools();
      renderStudents();
    });
    container.appendChild(btn);
  });
}

function getTabLabel(key) {
  if (key.startsWith("extra:")) {
    const fieldId = key.slice(6);
    return project.settings.extraFields.find((field) => field.id === fieldId)?.name || "추가 항목";
  }
  return FIELD_LABELS[key] || key;
}

function renderClassFilter() {
  const classes = Array.from(new Set(project.students.map((student) => student.className).filter(Boolean))).sort((a, b) => a.localeCompare(b, "ko"));
  $("classFilter").innerHTML = `<option value="">전체</option>` + classes.map((cls) => `<option value="${escapeHtml(cls)}">${escapeHtml(cls)}</option>`).join("");
  $("classFilter").value = state.classFilter;
}

function renderExtraFieldsList() {
  const container = $("extraFieldsList");
  if (project.settings.extraFields.length === 0) {
    container.innerHTML = `<div class="help">추가한 자유 항목이 없습니다.</div>`;
    return;
  }
  container.innerHTML = `<table class="mini-table"><thead><tr><th>항목명</th><th style="width:80px;">관리</th></tr></thead><tbody></tbody></table>`;
  const tbody = container.querySelector("tbody");
  project.settings.extraFields.forEach((field) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(field.name)}</td><td><button class="small danger" type="button">삭제</button></td>`;
    tr.querySelector("button").addEventListener("click", () => removeExtraField(field.id));
    tbody.appendChild(tr);
  });
}

function addExtraField() {
  const name = $("extraFieldName").value.trim();
  if (!name) return alert("추가할 항목명을 입력해주세요.");
  const field = { id: `extra_${Date.now()}`, name };
  project.settings.extraFields.push(field);
  project.students.forEach((student) => {
    if (!student.extraFieldValues) student.extraFieldValues = {};
    student.extraFieldValues[field.id] = "";
  });
  $("extraFieldName").value = "";
  markSettingsDirty();
}

function removeExtraField(fieldId) {
  const field = project.settings.extraFields.find((item) => item.id === fieldId);
  if (!field) return;
  if (!confirm(`'${field.name}' 항목을 삭제하시겠습니까? 학생 입력값도 함께 제거됩니다.`)) return;
  project.settings.extraFields = project.settings.extraFields.filter((item) => item.id !== fieldId);
  project.students.forEach((student) => {
    delete student.extraFieldValues?.[fieldId];
    markStudentDirty(student);
  });
  if (state.activeCriterion === `extra:${fieldId}`) state.activeCriterion = "attitude";
  markSettingsDirty();
}

function renderBulkTools() {
  const container = $("bulkTools");
  if (state.activeCriterion === "roster") {
    container.innerHTML = `
      <div class="student-card">
        <div class="section-title">학생/반 목록 입력</div>
        <div class="help">권장 형식: <span class="mono">반\t이름</span>. 첫 행이 헤더여도 자동으로 무시합니다.</div>
        <textarea id="rosterTsv" placeholder="Octa 1 2D 2부-3&#9;김도형&#10;Octa 1 2D 2부-3&#9;박서연"></textarea>
        <div class="actions" style="margin-top:10px;">
          <button id="applyRosterBtn" class="primary" type="button">학생 목록 반영</button>
          <button id="appendRosterBtn" type="button">기존 목록에 추가</button>
        </div>
      </div>`;
    $("applyRosterBtn").addEventListener("click", () => applyRosterFromTextarea(false));
    $("appendRosterBtn").addEventListener("click", () => applyRosterFromTextarea(true));
  } else if (state.activeCriterion === "scores") {
    const pageStudents = getPagedStudents();
    container.innerHTML = `
      <div class="student-card score-entry-card">
        <div>
          <div class="section-title">성적 입력</div>
          <div class="help">표에서 점수만 바로 입력하거나, 예시 엑셀 파일을 내려받아 업로드할 수 있습니다.</div>
        </div>
        <div class="bulk-score-actions">
          <button id="downloadScoreTemplateBtn" type="button">예시 엑셀 다운로드</button>
          <button id="uploadScoreFileBtn" type="button">엑셀 업로드</button>
          <button id="applyScoresBtn" class="primary" type="button">붙여넣기 반영</button>
        </div>
        <textarea id="scoreTsv" placeholder="이름&#9;반&#9;LC&#9;RC&#9;VO&#9;GR&#9;Total&#10;김도형&#9;Octa 1 2D 2부-3&#9;18&#9;17&#9;20&#9;19&#9;74"></textarea>
        <div class="bulk-score-table-wrap">
          <table class="bulk-score-table">
            <thead>
              <tr><th>이름</th><th>반</th><th>LC</th><th>RC</th><th>VO</th><th>GR</th><th>총점</th></tr>
            </thead>
            <tbody id="bulkScoreTableBody"></tbody>
          </table>
        </div>
      </div>`;
    $("applyScoresBtn").addEventListener("click", applyScoresFromTextarea);
    $("downloadScoreTemplateBtn").addEventListener("click", downloadScoreTemplate);
    $("uploadScoreFileBtn").addEventListener("click", () => $("scoreFileInput").click());
    renderBulkScoreTable(pageStudents);
  } else {
    container.innerHTML = "";
  }
}

function renderBulkScoreTable(students) {
  const tbody = $("bulkScoreTableBody");
  if (!tbody) return;
  tbody.innerHTML = "";
  if (students.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7">표시할 학생이 없습니다.</td></tr>`;
    return;
  }
  students.forEach((student) => {
    const score = getScore(student);
    const tr = document.createElement("tr");
    tr.innerHTML =
      `<td>${escapeHtml(student.name)}</td><td>${escapeHtml(student.className)}</td>` +
      SCORE_COLUMNS.map(([key, label]) => `<td><input data-student-id="${escapeHtml(student.id)}" data-score-key="${key}" placeholder="${label}" value="${escapeHtml(score[key] || "")}" /></td>`).join("");
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("input[data-score-key]").forEach((input) => {
    input.addEventListener("input", () => {
      const student = project.students.find((item) => item.id === input.dataset.studentId);
      if (!student) return;
      const score = getScore(student);
      score[input.dataset.scoreKey] = input.value.trim();
      markStudentDirty(student);
    });
  });
}

function parseTsv(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.split("\t").map((cell) => cell.trim()));
}

function applyRosterFromTextarea(append) {
  const rows = parseTsv($("rosterTsv")?.value || "").filter((row) => row.some((cell) => cell.trim()));
  if (rows.length === 0) return alert("붙여넣은 학생 목록이 없습니다.");
  const parsed = [];
  rows.forEach((row) => {
    const compact = row.filter(Boolean);
    if (compact.length < 2) return;
    const joined = compact.join(" ").toLowerCase().replace(/\s+/g, "");
    if (/^(반|class|classname|학생|이름|name)/.test(joined)) return;
    parsed.push({ className: compact[0], name: compact[1] });
  });
  if (parsed.length === 0) return alert("반과 이름 형식의 데이터를 찾지 못했습니다.");
  const existing = new Map(project.students.map((student) => [`${student.className}__${student.name}`, student]));
  const nextStudents = parsed.map((row, index) => {
    const matched = existing.get(`${row.className}__${row.name}`);
    if (matched) return normalizeStudent(matched, index);
    const student = normalizeStudent(row, index);
    student._dirty = true;
    return student;
  });
  if (append) {
    const keys = new Set(project.students.map((student) => `${student.className}__${student.name}`));
    nextStudents.forEach((student) => {
      const key = `${student.className}__${student.name}`;
      if (!keys.has(key)) project.students.push(student);
    });
  } else {
    project.students = nextStudents;
  }
  project.students = project.students.map((student, index) => normalizeStudent(student, index));
  project.students.forEach((student) => {
    if (student._dirty) project.sync.dirtyStudentIds[student.id] = true;
  });
  project.sync.settingsDirty = true;
  project.sync.hasUnsavedChanges = true;
  if (!state.selectedStudentId && project.students[0]) state.selectedStudentId = project.students[0].id;
  saveProjectToLocalStorageNow();
  renderAll();
  showToast(`학생 ${parsed.length}명 반영 완료`);
}

function normalizeHeader(value) {
  const h = String(value || "").trim().toLowerCase().replace(/\s+/g, "");
  const map = {
    "이름": "name",
    "학생": "name",
    "student": "name",
    "name": "name",
    "반": "className",
    "수강반": "className",
    "class": "className",
    "classname": "className",
    "class_name": "className",
    "lc": "lc",
    "listening": "lc",
    "rc": "rc",
    "reading": "rc",
    "vo": "vo",
    "voca": "vo",
    "vocab": "vo",
    "vocabulary": "vo",
    "단어": "vo",
    "gr": "gr",
    "grammar": "gr",
    "total": "total",
    "overall": "total",
    "총점": "total"
  };
  return map[h] || h;
}

function detectLevelFromClass(className) {
  const match = String(className || "").match(/^(Octa|Nona|Deca|Hepta|Penta|Demi|Alpha)/i);
  return match ? match[1] : "";
}

function getTestQuestionCounts(testType, className) {
  const test = String(testType || "").trim().toUpperCase();
  if (test.startsWith("MT")) {
    return { lc: 15, rc: 15, vo: 10, gr: 10 };
  }
  if (test === "TT") {
    if (/Nona/i.test(className)) return { lc: 20, rc: 20, vo: 10, gr: 10 };
    return { lc: 15, rc: 15, vo: 5, gr: 5 };
  }
  return null;
}

function convertRawScoresToWrong(scores, testType, className) {
  const counts = getTestQuestionCounts(testType, className);
  if (!counts) return scores;
  const result = { ...scores };
  if (scores.lc !== "") {
    const v = parseFloat(scores.lc);
    if (!isNaN(v)) result.lc = String(Math.round(v * counts.lc / 100) - counts.lc);
  }
  if (scores.rc !== "") {
    const v = parseFloat(scores.rc);
    if (!isNaN(v)) result.rc = String(Math.round(v * counts.rc / 100) - counts.rc);
  }
  if (scores.vo !== "") {
    const v = parseFloat(scores.vo);
    if (!isNaN(v)) result.vo = String(Math.round(v * counts.vo / 30) - counts.vo);
  }
  if (scores.gr !== "") {
    const v = parseFloat(scores.gr);
    if (!isNaN(v)) result.gr = String(Math.round(v * counts.gr / 30) - counts.gr);
  }
  return result;
}

function applyScoreRows(rows) {
  if (rows.length === 0) return { matched: 0, missed: 0 };
  const rawFirstRow = rows[0].map((cell) => String(cell || "").trim().toLowerCase());
  const isRawFormat = rawFirstRow.some((h) => h === "listening" || h === "overall");
  const header = rows[0].map(normalizeHeader);
  const hasHeader = header.some((key) => ["name", "className", "lc", "rc", "vo", "gr", "total"].includes(key));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const columns = hasHeader ? header : ["name", "className", "lc", "rc", "vo", "gr", "total"];
  let matched = 0;
  let missed = 0;
  dataRows.forEach((row) => {
    const obj = {};
    columns.forEach((key, index) => {
      if (key) obj[key] = row[index] ?? "";
    });
    const name = String(obj.name || row[0] || "").trim();
    const className = String(obj.className || row[1] || "").trim();
    if (!name) return;
    const student = getStudentByNameAndClass(name, className);
    if (!student) {
      missed += 1;
      return;
    }
    const score = getScore(student);
    if (isRawFormat) {
      const rawScores = {
        lc: String(obj.lc ?? "").trim(),
        rc: String(obj.rc ?? "").trim(),
        vo: String(obj.vo ?? "").trim(),
        gr: String(obj.gr ?? "").trim(),
        total: String(obj.total ?? "").trim()
      };
      const converted = convertRawScoresToWrong(rawScores, project.settings.testType, student.className);
      SCORE_COLUMNS.forEach(([key]) => {
        if (converted[key] !== "") score[key] = converted[key];
      });
    } else {
      SCORE_COLUMNS.forEach(([key]) => {
        if (obj[key] !== undefined) score[key] = String(obj[key] || "").trim();
      });
    }
    markStudentDirty(student);
    matched += 1;
  });
  saveProjectToLocalStorageNow();
  renderAll({ keepPage: true });
  return { matched, missed };
}

function applyScoresFromTextarea() {
  const rows = parseTsv($("scoreTsv")?.value || "").filter((row) => row.some((cell) => cell.trim()));
  if (rows.length === 0) return alert("붙여넣은 성적 데이터가 없습니다.");
  const result = applyScoreRows(rows);
  showToast(`성적 반영 완료: 매칭 ${result.matched}명, 미매칭 ${result.missed}명`);
}

function downloadScoreTemplate() {
  if (!window.XLSX) return alert("엑셀 기능을 불러오지 못했습니다.");
  const rawYear = String(project.settings.year || "").trim();
  const yearFull = rawYear.length === 2 ? `20${rawYear}` : rawYear;
  const semesterLabel = SEMESTER_LABELS[project.settings.semester] || project.settings.semester || "";
  const testType = project.settings.testType || "";
  const rows = [["캠퍼스", "이름", "학생코드", "레벨", "수강반", "시험명", "담임", "응시일", "Overall", "Listening", "Reading", "Vocabulary", "Grammar", "소요시간(분)", "응시여부"]];
  project.students.forEach((student) => {
    const level = detectLevelFromClass(student.className);
    const testName = [`${yearFull} ${semesterLabel}`.trim(), level, testType].filter(Boolean).join(" ");
    rows.push([
      "",
      student.name,
      "",
      level,
      student.className,
      testName,
      project.settings.teacherName || "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      ""
    ]);
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "성적입력예시");
  XLSX.writeFile(wb, `${makeProjectId(project) || "consultation"}_score_template.xlsx`);
}

async function importScoresFromFile(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;
  try {
    const rows = await readRowsFromFile(file);
    const result = applyScoreRows(rows.filter((row) => row.some((cell) => String(cell || "").trim())));
    showToast(`엑셀 반영 완료: 매칭 ${result.matched}명, 미매칭 ${result.missed}명`);
  } catch (error) {
    console.error(error);
    alert(`파일을 읽는 중 문제가 생겼습니다.\n\n${error.message || error}`);
  }
}

async function readRowsFromFile(file) {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (["xlsx", "xls", "csv"].includes(ext)) {
    if (!window.XLSX) throw new Error("엑셀 라이브러리를 불러오지 못했습니다.");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }).map((row) => row.map((cell) => String(cell ?? "").trim()));
  }
  const text = await file.text();
  return parseTsv(text);
}

function getFilteredStudents() {
  const search = state.search.toLowerCase();
  return project.students.filter((student) => {
    if (state.classFilter && student.className !== state.classFilter) return false;
    if (search && !`${student.name} ${student.className}`.toLowerCase().includes(search)) return false;
    if (state.completionFilter === "dirty" && !student._dirty) return false;
    if (state.completionFilter === "complete" && !student.manualComplete) return false;
    if (state.completionFilter === "incomplete" && student.manualComplete) return false;
    return true;
  });
}

function getTotalPages() {
  const count = getFilteredStudents().length;
  return Math.max(1, Math.ceil(count / (project.settings.pageSize || 20)));
}

function getPagedStudents() {
  const filtered = getFilteredStudents();
  const pageSize = project.settings.pageSize || 20;
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  if (state.currentPage >= totalPages) state.currentPage = totalPages - 1;
  const start = state.currentPage * pageSize;
  return filtered.slice(start, start + pageSize);
}

function renderStudents() {
  const container = $("studentsContainer");
  const filtered = getFilteredStudents();
  const pageStudents = getPagedStudents();
  const totalPages = getTotalPages();
  $("pageInfo").textContent = `${state.currentPage + 1} / ${totalPages} · ${filtered.length}명`;
  $("prevPageBtn").disabled = state.currentPage <= 0;
  $("nextPageBtn").disabled = state.currentPage >= totalPages - 1;
  if (state.activeCriterion === "summary") {
    renderSummary(container, filtered);
    return;
  }
  if (state.activeCriterion === "scores") {
    container.innerHTML = "";
    renderBulkTools();
    return;
  }
  if (state.activeCriterion === "roster" && project.students.length === 0) {
    container.innerHTML = `<div class="empty-state">학생/반 목록을 먼저 붙여넣고 [학생 목록 반영]을 눌러주세요.</div>`;
    return;
  }
  if (pageStudents.length === 0) {
    container.innerHTML = `<div class="empty-state">조건에 맞는 학생이 없습니다.</div>`;
    return;
  }
  container.innerHTML = "";
  pageStudents.forEach((student) => container.appendChild(renderStudentCard(student)));
}

function renderSummary(container, students) {
  const rows = students.map((student) => {
    const score = getScore(student);
    const progress = getStudentProgressState(student);
    const progressText = progress === "complete" ? "완료" : progress === "in-progress" ? "입력 중" : "미입력";
    return `<tr>
      <td><strong>${escapeHtml(student.name)}</strong><div class="student-meta">${escapeHtml(student.className)} · ${escapeHtml(student.classType)}</div></td>
      <td>${student._dirty ? "저장 필요" : "저장됨"}</td>
      <td>${progressText}</td>
      <td>${escapeHtml(score.total || "")}</td>
      <td>${escapeHtml(formatVocabTest(student.vocabTest))}</td>
      <td>${escapeHtml(student.specialNote || "")}</td>
    </tr>`;
  }).join("");
  container.innerHTML = `
    <div class="student-card">
      <div class="section-title">전체 진행표</div>
      <table class="mini-table">
        <thead><tr><th>학생</th><th>저장</th><th>입력</th><th>총점</th><th>지면 단어시험</th><th>특이사항</th></tr></thead>
        <tbody>${rows || `<tr><td colspan="6">학생이 없습니다.</td></tr>`}</tbody>
      </table>
    </div>`;
}

function renderStudentCard(student) {
  const card = document.createElement("div");
  card.className = "student-card";
  card.dataset.studentId = student.id;
  const status = student._dirty ? `<span class="dirty-dot">서버 저장 필요</span>` : `<span class="saved-dot">저장됨</span>`;
  card.innerHTML = `
    <div class="student-head">
      <div>
        <div class="student-name">${escapeHtml(student.name)}</div>
        <div class="student-meta">${escapeHtml(student.className)} · ${escapeHtml(student.classType)}</div>
      </div>
      <div class="actions">${status}</div>
    </div>
    <div class="student-inputs"></div>`;
  card.querySelector(".student-inputs").appendChild(renderCriterionInputs(student, state.activeCriterion));
  return card;
}

function renderCriterionInputs(student, criterion) {
  const wrap = document.createElement("div");
  if (criterion === "roster") {
    wrap.appendChild(makeInputRow("반", makeBoundInput(student.className, (value) => {
      student.className = value;
      student.classType = detectClassType(value);
      markStudentDirty(student, { rerender: true });
    })));
    wrap.appendChild(makeInputRow("이름", makeBoundInput(student.name, (value) => {
      student.name = value;
      markStudentDirty(student);
      renderPreviewStudentSelect();
    })));
    wrap.appendChild(makeInputRow("2D/3D", makeBoundSelect(student.classType, ["2D", "3D"], (value) => {
      student.classType = value;
      markStudentDirty(student);
    })));
  } else if (["bookHomework", "readi", "alex"].includes(criterion)) {
    const obj = student[criterion];
    wrap.appendChild(makeInputRow("평가", makeBoundSelect(obj.rating, RATING_OPTIONS, (value) => {
      obj.rating = value;
      markStudentDirty(student);
    })));
    wrap.appendChild(makeInputRow("메모", makeBoundTextarea(obj.memo, (value) => {
      obj.memo = value;
      markStudentDirty(student);
    })));
  } else if (criterion === "attitude") {
    wrap.appendChild(makeInputRow("수업태도", makeBoundTextarea(student.attitude, (value) => {
      student.attitude = value;
      markStudentDirty(student);
    })));
  } else if (criterion === "vocabTest") {
    wrap.appendChild(makeInputRow("평가", makeBoundSelect(student.vocabTest.rating, RATING_OPTIONS, (value) => {
      student.vocabTest.rating = value;
      markStudentDirty(student);
    })));
    wrap.appendChild(makeInputRow("백분율(선택)", makeBoundInput(student.vocabTest.percent, (value) => {
      student.vocabTest.percent = value.replace(/[^\d.]/g, "");
      markStudentDirty(student, { rerender: true });
    }, "예: 90")));
  } else if (criterion === "score") {
    wrap.appendChild(makeInputRow("잘한 점", makeBoundTextarea(student.scoreAnalysis.good, (value) => {
      student.scoreAnalysis.good = value;
      markStudentDirty(student);
    })));
    wrap.appendChild(makeInputRow("못한 점", makeBoundTextarea(student.scoreAnalysis.problem, (value) => {
      student.scoreAnalysis.problem = value;
      markStudentDirty(student);
    })));
    wrap.appendChild(makeInputRow("개선 방향", makeBoundTextarea(student.scoreAnalysis.improvement, (value) => {
      student.scoreAnalysis.improvement = value;
      markStudentDirty(student);
    })));
  } else if (criterion === "lowScoreCare") {
    wrap.appendChild(makeInputRow("성적 저조자 관리", makeBoundTextarea(student.lowScoreCare, (value) => {
      student.lowScoreCare = value;
      markStudentDirty(student);
    })));
  } else if (criterion === "specialNote") {
    wrap.appendChild(makeInputRow("특이사항", makeBoundTextarea(student.specialNote, (value) => {
      student.specialNote = value;
      markStudentDirty(student);
    })));
  } else if (criterion.startsWith("extra:")) {
    const fieldId = criterion.slice(6);
    if (!student.extraFieldValues) student.extraFieldValues = {};
    wrap.appendChild(makeInputRow("내용", makeBoundTextarea(student.extraFieldValues[fieldId] || "", (value) => {
      student.extraFieldValues[fieldId] = value;
      markStudentDirty(student);
    })));
  }
  return wrap;
}

function makeInputRow(labelText, control) {
  const row = document.createElement("div");
  row.className = "input-row";
  const label = document.createElement("label");
  label.textContent = labelText;
  row.appendChild(label);
  row.appendChild(control);
  return row;
}

function makeBoundInput(value, onChange, placeholder = "") {
  const input = document.createElement("input");
  input.value = value || "";
  input.placeholder = placeholder;
  input.addEventListener("input", () => onChange(input.value.trimStart()));
  return input;
}

function makeBoundTextarea(value, onChange) {
  const textarea = document.createElement("textarea");
  textarea.value = value || "";
  textarea.addEventListener("input", () => onChange(textarea.value));
  return textarea;
}

function makeBoundSelect(value, options, onChange) {
  const select = document.createElement("select");
  select.innerHTML = options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option || "선택")}</option>`).join("");
  select.value = value || "";
  select.addEventListener("change", () => onChange(select.value));
  return select;
}

function renderPreviewStudentSelect() {
  const students = getFilteredStudents();
  if (!state.selectedStudentId || !project.students.some((student) => student.id === state.selectedStudentId)) {
    state.selectedStudentId = students[0]?.id || project.students[0]?.id || "";
  }
  $("previewStudentSelect").innerHTML = project.students.map((student) => `<option value="${escapeHtml(student.id)}">${escapeHtml(student.className)} · ${escapeHtml(student.name)}</option>`).join("");
  $("previewStudentSelect").value = state.selectedStudentId;
}

function renderPreview() {
  const student = project.students.find((item) => item.id === state.selectedStudentId) || project.students[0];
  $("previewBox").textContent = student ? buildConsultationText(student) : "학생을 먼저 입력해주세요.";
}

function formatRatingMemo(obj) {
  return [obj?.rating, obj?.memo].filter(Boolean).join(" - ");
}

function formatVocabTest(vocabTest) {
  if (!vocabTest) return "";
  const rating = String(vocabTest.rating || "").trim();
  const percent = String(vocabTest.percent || "").replace(/%/g, "").trim();
  if (rating && percent) return `${rating}(${percent}%)`;
  return rating || (percent ? `${percent}%` : "");
}

function splitMultiline(text) {
  return String(text ?? "").replace(/\r\n/g, "\n").split("\n");
}

function getScheduleLines(student) {
  const schedule = project.settings.schedule?.[student.classType] || {};
  const consultationType = project.settings.consultationType;
  const lines = [];
  if (schedule.nextTestName || schedule.nextTestDate) {
    const parts = [schedule.nextTestName, schedule.nextTestDate].filter(Boolean);
    lines.push(`다음 시험: ${parts.join(" / ")}`);
  }
  if (["3차정기", "4차정기", "학기상담"].includes(consultationType)) {
    if (project.settings.nextSemester) lines.push(`다음 학기: ${semesterText(project.settings.year, project.settings.nextSemester)}`);
    if (schedule.summerStart) lines.push(`다음 학기 시작: ${schedule.summerStart}`);
    if (schedule.termEnd) lines.push(`학기 종료일: ${schedule.termEnd}`);
    if (schedule.eopAlexAward) lines.push(`EOP/ALEX: ${schedule.eopAlexAward}`);
    if (schedule.marketDay) lines.push(`Market Day: ${schedule.marketDay}`);
  }
  return lines;
}

function buildConsultationTitle() {
  const settings = project.settings || {};
  const rawYear = String(settings.year || "").trim();
  const year = rawYear.length === 2 ? `20${rawYear}` : rawYear;
  const semester = SEMESTER_LABELS[settings.semester] || settings.semester || "";
  const consultationType = displayConsultationType(settings.consultationType)
    .replace(/\s*정기/g, "")
    .replace(/\s*상담$/g, "")
    .trim();
  return [`${year} ${semester}`.trim(), consultationType, "상담"].filter(Boolean).join(" ");
}

function buildConsultationText(student) {
  const title = buildConsultationTitle();
  const score = getScore(student);
  const sections = [];
  if (project.settings.includeFields.attitude && student.attitude) sections.push({ title: "수업태도", lines: splitMultiline(student.attitude) });
  if (project.settings.includeFields.bookHomework && formatRatingMemo(student.bookHomework)) sections.push({ title: "교재숙제", lines: splitMultiline(formatRatingMemo(student.bookHomework)) });
  if (project.settings.includeFields.readi && formatRatingMemo(student.readi)) sections.push({ title: "READi", lines: splitMultiline(formatRatingMemo(student.readi)) });
  if (project.settings.includeFields.alex && formatRatingMemo(student.alex)) sections.push({ title: "Alex", lines: splitMultiline(formatRatingMemo(student.alex)) });
  if (project.settings.includeFields.vocabTest && formatVocabTest(student.vocabTest)) sections.push({ title: "지면 단어시험 결과", lines: [formatVocabTest(student.vocabTest)] });
  const scoreBits = SCORE_COLUMNS.filter(([key]) => score[key]).map(([key, label]) => `${label} ${score[key]}`);
  if (scoreBits.length) sections.push({ title: `${project.settings.testType} 성적`, lines: [scoreBits.join(" / ")] });
  if (project.settings.includeFields.score) {
    const analysisLines = [];
    if (student.scoreAnalysis.good) analysisLines.push(`잘한 점: ${student.scoreAnalysis.good}`);
    if (student.scoreAnalysis.problem) analysisLines.push(`못한 점: ${student.scoreAnalysis.problem}`);
    if (student.scoreAnalysis.improvement) analysisLines.push(`개선 방향: ${student.scoreAnalysis.improvement}`);
    if (analysisLines.length) sections.push({ title: "성적 분석", lines: analysisLines.flatMap(splitMultiline) });
  }
  if (project.settings.includeFields.lowScoreCare && student.lowScoreCare) sections.push({ title: "성적 저조자 관리", lines: splitMultiline(student.lowScoreCare) });
  project.settings.extraFields.forEach((field) => {
    const value = student.extraFieldValues?.[field.id];
    if (value) sections.push({ title: field.name, lines: splitMultiline(value) });
  });
  if (project.settings.includeFields.specialNote && student.specialNote) sections.push({ title: "특이사항", lines: splitMultiline(student.specialNote) });
  const scheduleLines = getScheduleLines(student);
  if (scheduleLines.length) sections.push({ title: "일정 안내", lines: scheduleLines });

  const out = [`[${title}]`, ""];
  sections.forEach((section, index) => {
    out.push(`${index + 1}. ${section.title}`);
    section.lines.forEach((line) => out.push(line.trim() ? `   ${line}` : ""));
    out.push("");
  });
  return out.join("\n").replace(/\n+$/, "");
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  }
  showToast("복사했습니다.");
}

function copyCurrentConsultationText() {
  const student = project.students.find((item) => item.id === state.selectedStudentId) || project.students[0];
  if (!student) return alert("복사할 학생이 없습니다.");
  copyText(buildConsultationText(student));
}

function copyAllConsultationTexts() {
  if (project.students.length === 0) return alert("복사할 학생이 없습니다.");
  copyText(project.students.map((student) => buildConsultationText(student)).join("\n\n------------------------------\n\n"));
}

function openCopyPopup(title, content) {
  const popup = window.open("", "_blank", "width=900,height=720");
  if (!popup) {
    copyText(content);
    alert("팝업이 차단되어 텍스트를 클립보드에 복사했습니다.");
    return;
  }
  popup.document.write(`<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(title)}</title>
  <style>
    body { margin: 0; padding: 18px; font-family: "Noto Sans KR", system-ui, sans-serif; background: #f4f6fb; color: #111827; }
    header { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 12px; }
    h1 { margin: 0; font-size: 18px; }
    button { border: 0; border-radius: 8px; padding: 9px 14px; background: #2563eb; color: #fff; font-weight: 700; cursor: pointer; }
    textarea { width: 100%; height: calc(100vh - 92px); box-sizing: border-box; border: 1px solid #d1d5db; border-radius: 10px; padding: 14px; font: 14px/1.6 "Noto Sans KR", system-ui, sans-serif; resize: none; background: #fff; color: #111827; }
  </style>
</head>
<body>
  <header>
    <h1>${escapeHtml(title)}</h1>
    <button id="copyBtn" type="button">전체 복사</button>
  </header>
  <textarea id="exportText">${escapeHtml(content)}</textarea>
  <script>
    const textarea = document.getElementById("exportText");
    document.getElementById("copyBtn").addEventListener("click", async () => {
      textarea.focus();
      textarea.select();
      try {
        await navigator.clipboard.writeText(textarea.value);
      } catch {
        document.execCommand("copy");
      }
    });
    textarea.focus();
    textarea.select();
  <\/script>
</body>
</html>`);
  popup.document.close();
}

function exportConsultationTxt(mode) {
  if (project.students.length === 0) return alert("내보낼 학생이 없습니다.");
  if (mode === "vertical") {
    const content = project.students.map((student) => buildConsultationText(student)).join("\n\n------------------------------\n\n");
    openCopyPopup("TXT 내보내기: 세로형", content);
    return;
  }
  const content = project.students.map((student) => {
    const consultation = buildConsultationText(student).replace(/\t/g, " ").replace(/\r?\n/g, "\n");
    const escaped = `"${consultation.replace(/"/g, "\"\"")}"`;
    return `${student.name}\t${escaped}`;
  }).join("\n");
  openCopyPopup("TXT 내보내기: 한 줄형", content);
}

function saveProjectAsJSON() {
  saveProjectToLocalStorageNow();
  const blob = new Blob([JSON.stringify(project, null, 2)], { type: "application/json;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${makeProjectId(project)}_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  a.remove();
}

function loadProjectFromJSONFile(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const data = safeParseJson(reader.result);
    if (!data) return alert("JSON 파일을 읽을 수 없습니다.");
    if (project.sync.hasUnsavedChanges && !confirm("현재 저장되지 않은 변경사항이 있습니다. JSON 파일로 덮어쓸까요?")) return;
    project = normalizeProjectShape(data);
    saveProjectToLocalStorageNow();
    renderAll();
    showToast("JSON 불러오기 완료");
  };
  reader.readAsText(file, "utf-8");
}

function loadProjectFromLocalStorage() {
  const data = localStorage.getItem(getLocalStorageKey());
  if (!data) return false;
  const parsed = safeParseJson(data);
  if (!parsed) return false;
  project = normalizeProjectShape(parsed);
  return true;
}

function showLocalProjectsModal() {
  const keys = Object.keys(localStorage).filter((key) => key.startsWith("consultationProject__")).sort();
  const list = $("localProjectsList");
  if (keys.length === 0) {
    list.innerHTML = `<div class="empty-state">로컬 저장 프로젝트가 없습니다.</div>`;
  } else {
    list.innerHTML = keys.map((key) => {
      const data = safeParseJson(localStorage.getItem(key), {});
      const name = data?.projectName || key.replace("consultationProject__", "");
      const updated = data?.updatedAt ? data.updatedAt.slice(0, 19).replace("T", " ") : "-";
      return `<div class="local-item">
        <div><strong>${escapeHtml(name)}</strong><div class="student-meta mono">${escapeHtml(key.replace("consultationProject__", ""))}</div><div class="student-meta">업데이트: ${escapeHtml(updated)} · 학생 ${data?.students?.length || 0}명</div></div>
        <button class="small" data-key="${escapeHtml(key)}" type="button">불러오기</button>
      </div>`;
    }).join("");
    list.querySelectorAll("button[data-key]").forEach((button) => {
      button.addEventListener("click", () => {
        if (project.sync.hasUnsavedChanges && !confirm("현재 저장되지 않은 변경사항이 있습니다. 로컬 프로젝트로 바꿀까요?")) return;
        const parsed = safeParseJson(localStorage.getItem(button.dataset.key));
        if (!parsed) return alert("로컬 프로젝트를 읽을 수 없습니다.");
        project = normalizeProjectShape(parsed);
        state.selectedStudentId = project.students[0]?.id || "";
        $("localProjectsModal").classList.remove("show");
        renderAll();
        showToast("로컬 프로젝트 불러오기 완료");
      });
    });
  }
  $("localProjectsModal").classList.add("show");
}

function createNewProject() {
  if (project.sync.hasUnsavedChanges && !confirm("현재 저장되지 않은 변경사항이 있습니다. 새 프로젝트를 만들까요?")) return;
  project = createDefaultProject();
  state.activeCriterion = "roster";
  state.currentPage = 0;
  state.selectedStudentId = "";
  saveProjectToLocalStorageNow();
  renderAll();
}

function clearCurrentProjectData() {
  if (project.students.length === 0) {
    showToast("비울 학생 명단이 없습니다.");
    return;
  }
  const message = "현재 프로젝트의 학생 명단과 입력한 상담 내용을 모두 비울까요?\n\n서버에 반영하려면 이후 '서버 저장'을 눌러주세요.";
  if (!confirm(message)) return;
  project.sync.deletedStudentIds = {
    ...(project.sync.deletedStudentIds || {}),
    ...Object.fromEntries(project.students.map((student) => [student.id, true]))
  };
  project.students = [];
  project.sync.dirtyStudentIds = {};
  project.sync.settingsDirty = true;
  project.sync.hasUnsavedChanges = true;
  state.selectedStudentId = "";
  state.currentPage = 0;
  state.activeCriterion = "roster";
  saveProjectToLocalStorageNow();
  renderAll();
  showToast("현재 프로젝트 내용을 비웠습니다.");
}

function isFirebaseConfigured() {
  return Boolean(firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.appId);
}

async function initFirebase() {
  if (!isFirebaseConfigured()) {
    firebaseReady = false;
    project.sync.firebaseEnabled = false;
    renderAuthState();
    renderFirebaseBadge();
    return;
  }
  try {
    const [appMod, firestoreMod, authMod] = await Promise.all([
      import("https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js"),
      import("https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js"),
      import("https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js")
    ]);
    fb = { ...appMod, ...firestoreMod, ...authMod };
    firebaseApp = fb.initializeApp(firebaseConfig);
    db = fb.getFirestore(firebaseApp);
    try {
      auth = fb.initializeAuth(firebaseApp, {
        persistence: fb.browserLocalPersistence,
        popupRedirectResolver: fb.browserPopupRedirectResolver
      });
    } catch (error) {
      auth = fb.getAuth(firebaseApp);
    }
    googleProvider = new fb.GoogleAuthProvider();
    firebaseReady = true;
    project.sync.firebaseEnabled = true;
    fb.onAuthStateChanged(auth, (user) => {
      currentUser = user || null;
      renderAuthState();
      renderFirebaseBadge();
    });
  } catch (error) {
    console.error(error);
    firebaseReady = false;
    project.sync.firebaseEnabled = false;
  }
  renderAuthState();
  renderFirebaseBadge();
}

function renderAuthState() {
  const info = $("authInfo");
  const infoTop = $("authInfoTop");
  const loginBtn = $("loginBtn");
  const logoutBtn = $("logoutBtn");
  if (!firebaseReady) {
    info.textContent = "서버 연결이 설정되지 않았습니다. 로컬 저장과 JSON 백업은 그대로 사용할 수 있습니다.";
    infoTop.textContent = "미로그인";
    infoTop.className = "auth-info warn";
    infoTop.hidden = false;
    loginBtn.hidden = false;
    logoutBtn.hidden = true;
  } else if (currentUser) {
    info.textContent = `로그인됨: ${currentUser.displayName || currentUser.email || "구글 계정"}`;
    infoTop.textContent = currentUser.displayName || currentUser.email || "로그인됨";
    infoTop.className = "auth-info ok";
    infoTop.hidden = false;
    loginBtn.hidden = true;
    logoutBtn.hidden = false;
  } else {
    info.textContent = location.protocol === "file:"
      ? "구글 로그인은 로컬 서버 주소에서 더 안정적으로 동작합니다. 로그인하지 않아도 로컬 작업은 가능합니다."
      : "구글 로그인 후 서버 저장을 사용할 수 있습니다. 로그인하지 않아도 로컬 작업은 가능합니다.";
    infoTop.textContent = "미로그인";
    infoTop.className = "auth-info warn";
    infoTop.hidden = false;
    loginBtn.hidden = false;
    logoutBtn.hidden = true;
  }
}

async function loginWithGoogle() {
  if (!firebaseReady || !auth) return alert("서버 설정이 필요합니다.");
  try {
    await fb.signInWithPopup(auth, googleProvider, fb.browserPopupRedirectResolver);
    showToast("구글 로그인 완료");
  } catch (error) {
    console.error(error);
    const fileHint = location.protocol === "file:"
      ? "\n\n현재 파일을 직접 열고 있다면 로그인 제한이 생길 수 있습니다. 이 폴더에서 로컬 서버를 띄운 뒤 http://localhost 주소로 접속해 주세요."
      : "";
    alert(`로그인 실패: ${error.message || error}${fileHint}`);
  }
}

async function logout() {
  if (!firebaseReady || !auth) return;
  await fb.signOut(auth);
  showToast("로그아웃했습니다.");
}

function removeInternalFields(student) {
  const clean = structuredCloneSafe(student);
  delete clean._dirty;
  delete clean._lastSavedAt;
  return clean;
}

async function saveProjectToFirebase() {
  saveProjectToLocalStorageNow();
  if (!firebaseReady || !db) {
    if (confirm("서버에 연결되어 있지 않습니다. 현재 프로젝트를 JSON 파일로 저장할까요?")) saveProjectAsJSON();
    return;
  }
  if (!currentUser) {
    alert("서버 저장은 구글 로그인 후 사용할 수 있습니다.");
    return;
  }
  const projectId = makeProjectId(project);
  project.sync.projectId = projectId;
  setServerStatus("서버 저장 중...");
  try {
    const dirtyIds = Object.keys(project.sync.dirtyStudentIds || {});
    const deletedIds = Object.keys(project.sync.deletedStudentIds || {});
    const dirtyStudents = project.students.filter((student) => dirtyIds.includes(student.id));
    if (project.sync.settingsDirty || !project.sync.lastServerSavedAt) {
      const metaRef = fb.doc(db, "consultationProjects", projectId);
      await fb.setDoc(metaRef, {
        version: project.version,
        projectName: project.projectName,
        teacherName: project.settings.teacherName,
        year: project.settings.year,
        semester: project.settings.semester,
        nextSemester: project.settings.nextSemester,
        consultationType: project.settings.consultationType,
        testType: project.settings.testType,
        settings: {
          includeFields: project.settings.includeFields,
          extraFields: project.settings.extraFields,
          schedule: project.settings.schedule,
          pageSize: project.settings.pageSize
        },
        studentCount: project.students.length,
        updatedAt: fb.serverTimestamp(),
        updatedBy: currentUser?.email || "anonymous"
      }, { merge: true });
    }
    if (dirtyStudents.length > 0) {
      for (let i = 0; i < dirtyStudents.length; i += 400) {
        const batch = fb.writeBatch(db);
        dirtyStudents.slice(i, i + 400).forEach((student) => {
          const ref = fb.doc(db, "consultationProjects", projectId, "students", student.id);
          batch.set(ref, { ...removeInternalFields(student), updatedAt: fb.serverTimestamp(), updatedBy: currentUser?.email || "anonymous" }, { merge: true });
        });
        await batch.commit();
      }
    }
    if (deletedIds.length > 0) {
      for (let i = 0; i < deletedIds.length; i += 400) {
        const batch = fb.writeBatch(db);
        deletedIds.slice(i, i + 400).forEach((studentId) => {
          const ref = fb.doc(db, "consultationProjects", projectId, "students", studentId);
          batch.delete(ref);
        });
        await batch.commit();
      }
    }
    const savedAt = new Date().toISOString();
    project.students.forEach((student) => {
      if (dirtyIds.includes(student.id)) {
        student._dirty = false;
        student._lastSavedAt = savedAt;
      }
    });
    project.sync.settingsDirty = false;
    project.sync.dirtyStudentIds = {};
    project.sync.deletedStudentIds = {};
    project.sync.hasUnsavedChanges = false;
    project.sync.lastServerSavedAt = savedAt;
    saveProjectToLocalStorageNow();
    renderAll({ keepPage: true });
    showToast(`서버 저장 완료: 저장 ${dirtyStudents.length}명, 삭제 ${deletedIds.length}명`);
  } catch (error) {
    console.error(error);
    setServerStatus("서버 저장 실패");
    alert(`서버 저장에 실패했습니다.\n\n${error.message || error}`);
  }
}

async function loadProjectFromFirebase() {
  if (!firebaseReady || !db) return alert("서버 설정이 필요합니다.");
  if (!currentUser) return alert("서버 불러오기는 구글 로그인 후 사용할 수 있습니다.");
  if (project.sync.hasUnsavedChanges && !confirm("현재 저장되지 않은 변경사항이 있습니다. 서버 데이터로 덮어쓸까요?")) return;
  const projectId = makeProjectId(project);
  setServerStatus("서버에서 불러오는 중...");
  try {
    const metaRef = fb.doc(db, "consultationProjects", projectId);
    const metaSnap = await fb.getDoc(metaRef);
    if (!metaSnap.exists()) {
      setServerStatus("서버 데이터 없음");
      return alert("이 프로젝트의 서버 저장 데이터가 없습니다.");
    }
    const meta = metaSnap.data();
    const studentsRef = fb.collection(db, "consultationProjects", projectId, "students");
    const studentSnap = await fb.getDocs(studentsRef);
    const loaded = normalizeProjectShape({
      ...meta,
      settings: {
        ...(meta.settings || {}),
        teacherName: meta.teacherName,
        year: meta.year,
        semester: meta.semester,
        nextSemester: meta.nextSemester,
        consultationType: meta.consultationType,
        testType: meta.testType
      },
      students: studentSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
    });
    loaded.sync.lastServerLoadedAt = new Date().toISOString();
    loaded.sync.lastServerSavedAt = new Date().toISOString();
    loaded.sync.hasUnsavedChanges = false;
    loaded.sync.dirtyStudentIds = {};
    loaded.sync.deletedStudentIds = {};
    loaded.sync.settingsDirty = false;
    loaded.students.forEach((student) => {
      student._dirty = false;
    });
    project = loaded;
    state.selectedStudentId = project.students[0]?.id || "";
    saveProjectToLocalStorageNow();
    renderAll();
    setServerStatus("서버 불러오기 완료");
    showToast("서버 불러오기 완료");
  } catch (error) {
    console.error(error);
    setServerStatus("서버 불러오기 실패");
    alert(`서버 불러오기에 실패했습니다.\n\n${error.message || error}`);
  }
}

function renderProgressSidebar() {
  const students = project.students.slice().sort((a, b) => `${a.className} ${a.name}`.localeCompare(`${b.className} ${b.name}`, "ko"));
  $("progressSummary").textContent = students.length === 0
    ? "학생이 없습니다."
    : `완료 ${students.filter((student) => student.manualComplete).length}명 / 전체 ${students.length}명`;
  const container = $("progressList");
  container.innerHTML = students.map((student) => {
    const progressState = getStudentProgressState(student);
    const dotClass = progressState === "complete" ? "complete" : progressState === "in-progress" ? "in-progress" : "";
    return `<label class="progress-item">
      <input type="checkbox" data-progress-id="${escapeHtml(student.id)}" ${student.manualComplete ? "checked" : ""} />
      <span class="progress-name">${escapeHtml(student.className)} ${escapeHtml(student.name)}</span>
      <span class="progress-dot ${dotClass}"></span>
    </label>`;
  }).join("");
  container.querySelectorAll("input[data-progress-id]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const student = project.students.find((item) => item.id === checkbox.dataset.progressId);
      if (!student) return;
      student.manualComplete = checkbox.checked;
      markStudentDirty(student);
      renderStudents();
    });
  });
}

function loadFallbackProject() {
  const keys = Object.keys(localStorage).filter((key) => key.startsWith("consultationProject__")).sort();
  const latestKey = keys[keys.length - 1];
  if (!latestKey) return false;
  const parsed = safeParseJson(localStorage.getItem(latestKey));
  if (!parsed) return false;
  project = normalizeProjectShape(parsed);
  return true;
}

(async function init() {
  bindEvents();
  if (!loadProjectFromLocalStorage()) loadFallbackProject();
  renderAll();
  await initFirebase();
  renderAll({ keepPage: true });
})();
