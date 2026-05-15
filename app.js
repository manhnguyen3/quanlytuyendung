/* global XLSX */

const STORAGE_KEY = "tuyenDungWorkflow_v1";

/** Google Sheet — link chia sẻ /edit hoặc xuất bản pubhtml */
const DEFAULT_GOOGLE_SHEET_PUB_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQs96Y38iag14gR2tVg72YahtcmbfOFyuZbwUNFlzp7qA2juTWHxon6SDe6hRdLDBpx_sDkcMGvV8MR/pubhtml";

/** Google AI Studio (Gemini) — mặc định trong app, không hiển thị trên giao diện */
const DEFAULT_GEMINI_API_KEY = "AIzaSyB6qO0GLO_GMRy5h0lukgU3nXk3OI3L5Sc";
const DEFAULT_GEMINI_MODEL = "gemini-2.5-flash";

const STATUS_LABELS = {
  new: "Chưa tạo câu hỏi",
  pending_hr: "Chờ HR duyệt",
  approved: "Đã duyệt — chưa gửi HRM",
  sent_hrm: "Đã gửi HRM",
  done: "Hoàn tất",
};

const HEADER_ALIASES = {
  name: ["họ và tên", "ho ten", "họ tên", "hoten", "ten", "tên", "fullname", "full name", "name", "ung vien", "ứng viên"],
  email: ["email", "mail", "e-mail"],
  phone: ["sdt", "điện thoại", "dien thoai", "phone", "mobile", "số điện thoại"],
  industry: ["ngành", "nganh", "chuyên ngành", "chuyen nganh", "industry", "lĩnh vực", "linh vuc", "sector", "nhóm ngành"],
  position: ["vị trí", "vi tri", "position", "job", "title", "chức danh", "chuc danh", "vị trí ứng tuyển"],
  experience: ["kinh nghiệm", "kinh nghiem", "experience", "số năm", "so nam"],
  education: ["học vấn", "hoc van", "education", "bằng cấp", "bang cap", "trình độ", "trinh do"],
  skills: ["kỹ năng", "ky nang", "skills", "skill", "competency"],
  note: ["ghi chú", "ghi chu", "note", "mô tả", "mo ta", "description"],
};

function norm(s) {
  return String(s ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function pickMappedRow(row, headerMap) {
  const out = {};
  for (const [key, colIdx] of Object.entries(headerMap)) {
    if (colIdx >= 0) out[key] = row[colIdx];
  }
  return out;
}

function detectHeaderMap(headers) {
  const normalized = headers.map((h, i) => ({ raw: h, i, n: norm(h) }));
  const map = {
    name: -1,
    email: -1,
    phone: -1,
    industry: -1,
    position: -1,
    experience: -1,
    education: -1,
    skills: -1,
    note: -1,
  };

  for (const key of Object.keys(map)) {
    const aliases = HEADER_ALIASES[key];
    for (const { i, n } of normalized) {
      if (!n) continue;
      if (aliases.some((a) => n === a || n.includes(a))) {
        map[key] = i;
        break;
      }
    }
  }

  if (map.name < 0 && headers.length) {
    map.name = 0;
  }

  return map;
}

function rowToCandidate(rowArr, headers, headerMap, sheetName, rowIndex) {
  const cells = rowArr.map((c) => (c === undefined || c === null ? "" : c));
  const mapped = pickMappedRow(cells, headerMap);
  const displayName =
    String(mapped.name || "").trim() ||
    `Ứng viên dòng ${rowIndex + 2}`;

  const raw = {};
  headers.forEach((h, idx) => {
    if (h != null && String(h).trim() !== "") raw[String(h).trim()] = cells[idx];
  });

  return {
    id: `${sheetName}::${rowIndex}`,
    sheetName,
    rowIndex,
    displayName,
    industry: String(mapped.industry || "").trim() || "—",
    position: String(mapped.position || "").trim() || "—",
    email: String(mapped.email || "").trim(),
    phone: String(mapped.phone || "").trim(),
    experience: String(mapped.experience || "").trim(),
    education: String(mapped.education || "").trim(),
    skills: String(mapped.skills || "").trim(),
    note: String(mapped.note || "").trim(),
    raw,
  };
}

/**
 * Chuyển link Google Sheet thành URL CSV:
 * - /edit, /view, /d/{id} → export?format=csv (cần quyền «Bất kỳ ai có link» xem được)
 * - /pubhtml, /pub → pub?output=csv (xuất bản lên web)
 */
function toGoogleSheetCsvUrl(pubUrl) {
  const raw = String(pubUrl ?? "").trim();
  if (!raw) return "";
  let u;
  try {
    u = new URL(raw);
  } catch {
    return "";
  }
  if (!u.hostname.includes("google.com") || !u.pathname.includes("/spreadsheets/")) return "";

  const path = u.pathname.replace(/\/+$/u, "");

  const docMatch = path.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (docMatch && !path.includes("/e/")) {
    const id = docMatch[1];
    const gid = u.searchParams.get("gid") || "0";
    return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${encodeURIComponent(gid)}`;
  }

  if (path.includes("/e/")) {
    const pubPath = path.replace(/\/pubhtml\/?$/i, "/pub");
    if (pubPath.endsWith("/pub")) {
      const base = `${u.protocol}//${u.host}${pubPath}`;
      return `${base}?output=csv`;
    }
  }

  return "";
}

/** Parse CSV (hỗ trợ dấu ngoặc kép, xuống dòng trong ô). */
function parseCsv(text) {
  const t = String(text ?? "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    const next = t[i + 1];
    if (inQuotes) {
      if (c === '"') {
        if (next === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n" || (c === "\r" && next === "\n")) {
      if (c === "\r") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (c !== "\r") {
      field += c;
    }
  }
  row.push(field);
  if (row.some((cell) => String(cell).trim() !== "") || rows.length === 0) {
    rows.push(row);
  }
  return rows.map((r) => r.map((cell) => (cell === undefined || cell === null ? "" : cell)));
}

function candidatesFromMatrix(rows, sheetName) {
  if (!rows.length) return [];

  let headerRowIdx = 0;
  let headers = (rows[0] || []).map((h) => String(h).trim());

  const nonEmpty = headers.filter(Boolean).length;
  if (nonEmpty < 2 && rows.length > 1) {
    headerRowIdx = 1;
    headers = (rows[1] || []).map((h) => String(h).trim());
  }

  const headerMap = detectHeaderMap(headers);
  const dataStart = headerRowIdx + 1;
  const out = [];

  for (let r = dataStart; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.some((c) => String(c).trim() !== "")) continue;
    out.push(rowToCandidate(row, headers, headerMap, sheetName, r));
  }
  return out;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const candidates = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows.length) continue;
    candidates.push(...candidatesFromMatrix(rows, sheetName));
  }

  return candidates;
}

/**
 * Tải nội dung CSV từ URL Google (pub).
 * Trình duyệt thường báo "Failed to fetch" khi mở file qua file:// hoặc khi Google không cho CORS —
 * khi đó thử lại qua api.allorigins.win (chỉ dùng cho sheet đã công khai).
 */
async function fetchPublishedCsvText(csvUrl) {
  const direct = async () => {
    const res = await fetch(csvUrl, { credentials: "omit", cache: "no-store" });
    if (!res.ok) throw new Error(`Tải sheet lỗi HTTP ${res.status}`);
    return res.text();
  };

  const viaAllOrigins = async () => {
    const proxy = `https://api.allorigins.win/get?url=${encodeURIComponent(csvUrl)}`;
    const res = await fetch(proxy, { credentials: "omit", cache: "no-store" });
    if (!res.ok) throw new Error(`Proxy AllOrigins HTTP ${res.status}`);
    const data = await res.json();
    const raw = data?.contents;
    if (typeof raw !== "string") throw new Error("Proxy không trả về nội dung.");
    return raw;
  };

  let text;
  let usedProxy = false;
  try {
    text = await direct();
  } catch (e1) {
    const m = String(e1?.message ?? e1).toLowerCase();
    const looksLikeBlockedFetch =
      e1 instanceof TypeError || m.includes("failed to fetch") || m.includes("networkerror");
    if (!looksLikeBlockedFetch) throw e1;
    try {
      text = await viaAllOrigins();
      usedProxy = true;
    } catch (e2) {
      throw new Error(
        `Không tải được sheet (thường do mở app bằng file:// — trình duyệt chặn CORS). ` +
          `Hãy mở qua địa chỉ http://localhost (VS Code «Live Server», hoặc chạy \`npx serve .\` trong thư mục app). ` +
          `Lỗi gốc: ${e1?.message || e1}; proxy: ${e2?.message || e2}`
      );
    }
  }

  const t = text.trim();
  if (t.startsWith("<!DOCTYPE") || (t.includes("<html") && t.length < 8000)) {
    throw new Error(
      "Google trả về trang HTML thay vì CSV — kiểm tra Sheet vẫn bật «Xuất bản lên web» và link pubhtml đúng."
    );
  }

  return { text, usedProxy };
}

async function loadCandidatesFromGoogleSheet(pubUrl) {
  const csvUrl = toGoogleSheetCsvUrl(pubUrl);
  if (!csvUrl) {
    throw new Error("Không chuyển được link Google Sheet sang CSV. Dùng link dạng .../pubhtml hoặc .../pub.");
  }
  const { text, usedProxy } = await fetchPublishedCsvText(csvUrl);
  const rows = parseCsv(text);
  const list = candidatesFromMatrix(rows, "Danh sách (Google Sheet)");
  return { candidates: list, usedProxy };
}

function defaultWorkflowEntry() {
  return {
    status: "new",
    questions: [],
    questionsText: "",
    hrComment: "",
    approvedAt: null,
    rejectedAt: null,
    hrmSentAt: null,
    hrmResponse: null,
    reportAt: null,
    log: [],
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { byId: {}, settings: {} };
    return JSON.parse(raw);
  } catch {
    return { byId: {}, settings: {} };
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function pushLog(entry, message) {
  entry.log = entry.log || [];
  entry.log.push({ t: new Date().toISOString(), message });
}

function templateQuestions(candidate) {
  const pos = candidate.position || "vị trí ứng tuyển";
  const ind = candidate.industry || "ngành";
  const skills = candidate.skills || "kỹ năng được nêu trong hồ sơ";
  const edu = candidate.education || "trình độ của bạn";
  const exp = candidate.experience || "kinh nghiệm làm việc";

  return [
    `Với vị trí ${pos} trong lĩnh vực ${ind}, anh/chị mô tả ngắn gọn một dự án hoặc công việc tiêu biểu nhất liên quan trực tiếp đến vai trò này?`,
    `Anh/chị đánh giá thế mạnh chính của mình so với các ứng viên khác cho vị trí ${pos} là gì?`,
    `Kỹ năng/kiến thức nào trong nhóm "${skills}" anh/chị đã áp dụng thực tế và kết quả đạt được?`,
    `Liên quan ${edu} và ${exp}, anh/chị đã học được điều gì quan trọng nhất cho công việc hiện tại?`,
    `Trong môi trường ${ind}, anh/chị xử lý thế nào khi deadline gấp nhưng chất lượng vẫn phải đảm bảo?`,
    `Một tình huống khó với đồng nghiệp/khách hàng nội bộ: anh/chị chọn hướng xử lý nào và vì sao?`,
    `Trong 90 ngày đầu nếu được nhận, anh/chị ưu tiên 3 mục tiêu cụ thể nào và cách đo lường?`,
    `Anh/chị có câu hỏi nào cho chúng tôi về vị trí, đội ngũ hoặc văn hóa làm việc không?`,
  ];
}

function parseQuestionsJsonFromModelText(text) {
  const raw = String(text ?? "").trim();
  if (!raw) throw new Error("Gemini không trả về nội dung.");
  try {
    return JSON.parse(raw);
  } catch {
    const m = raw.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("Không parse được JSON từ Gemini.");
    return JSON.parse(m[0]);
  }
}

async function generateQuestionsWithGemini(candidate, apiKey, modelId) {
  const model = String(modelId || DEFAULT_GEMINI_MODEL).replace(/^models\//, "");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
    model
  )}:generateContent?key=${encodeURIComponent(apiKey)}`;

  const systemText =
    'Bạn là chuyên gia tuyển dụng. Chỉ trả về một JSON hợp lệ dạng {"questions": ["...", ...]} với đúng 8 chuỗi: mỗi chuỗi là một câu hỏi phỏng vấn tiếng Việt, súc tích, không trùng lặp, phù hợp ứng viên và vị trí.';

  const userText = JSON.stringify({
    candidate: {
      name: candidate.displayName,
      industry: candidate.industry,
      position: candidate.position,
      education: candidate.education,
      experience: candidate.experience,
      skills: candidate.skills,
      note: candidate.note,
    },
  });

  const body = {
    systemInstruction: { parts: [{ text: systemText }] },
    contents: [{ role: "user", parts: [{ text: userText }] }],
    generationConfig: {
      temperature: 0.5,
      responseMimeType: "application/json",
    },
  };

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const textRaw = await res.text();
  let data = {};
  try {
    data = textRaw ? JSON.parse(textRaw) : {};
  } catch {
    if (!res.ok) throw new Error(textRaw || res.statusText);
    throw new Error("Phản hồi Gemini không phải JSON.");
  }

  if (!res.ok) {
    const msg = data?.error?.message || textRaw || res.statusText;
    throw new Error(typeof msg === "string" ? msg : JSON.stringify(msg));
  }

  if (data?.promptFeedback?.blockReason) {
    throw new Error(`Yêu cầu bị chặn: ${data.promptFeedback.blockReason}`);
  }

  const cand = data?.candidates?.[0];
  const text = cand?.content?.parts?.map((p) => p.text).join("") || "";
  if (!text && cand?.finishReason) {
    throw new Error(`Gemini kết thúc sớm: ${cand.finishReason}`);
  }

  const parsed = parseQuestionsJsonFromModelText(text);
  const qs = Array.isArray(parsed.questions) ? parsed.questions.map(String) : [];
  if (qs.length < 3) throw new Error("Gemini trả về quá ít câu hỏi.");
  return qs.slice(0, 12);
}

function downloadBlob(filename, blob) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** --- App --- */

let candidates = [];
let state = loadState();
let selectedId = null;
let sheetFetchLoading = false;
let lastSheetError = "";

const els = {
  fileInput: document.getElementById("fileInput"),
  btnReloadSheet: document.getElementById("btnReloadSheet"),
  sheetHint: document.getElementById("sheetHint"),
  filterIndustry: document.getElementById("filterIndustry"),
  filterStatus: document.getElementById("filterStatus"),
  search: document.getElementById("search"),
  stats: document.getElementById("stats"),
  candidateList: document.getElementById("candidateList"),
  emptyDetail: document.getElementById("emptyDetail"),
  detail: document.getElementById("detail"),
  hrmUrl: document.getElementById("hrmUrl"),
  btnExportState: document.getElementById("btnExportState"),
  btnImportState: document.getElementById("btnImportState"),
  stateImport: document.getElementById("stateImport"),
};

function persistSettings() {
  state.settings = {
    ...(state.settings || {}),
    hrmUrl: els.hrmUrl.value,
  };
  saveState(state);
}

function hydrateSettings() {
  const s = state.settings || {};
  if (s.hrmUrl) els.hrmUrl.value = s.hrmUrl;
}

function ensureEntry(id) {
  if (!state.byId[id]) state.byId[id] = defaultWorkflowEntry();
  return state.byId[id];
}

function statusCounts() {
  const c = { new: 0, pending_hr: 0, approved: 0, sent_hrm: 0, done: 0 };
  for (const cand of candidates) {
    const st = ensureEntry(cand.id).status;
    if (c[st] !== undefined) c[st]++;
  }
  return c;
}

function filteredCandidates() {
  const ind = els.filterIndustry.value;
  const st = els.filterStatus.value;
  const q = norm(els.search.value);

  return candidates.filter((c) => {
    if (ind && c.industry !== ind) return false;
    const entry = ensureEntry(c.id);
    if (st && entry.status !== st) return false;
    if (q) {
      const hay = norm(
        [c.displayName, c.position, c.industry, c.email, c.phone, c.skills].join(" ")
      );
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function renderIndustryFilter() {
  const set = new Set(candidates.map((c) => c.industry).filter(Boolean));
  const prev = els.filterIndustry.value;
  els.filterIndustry.innerHTML = '<option value="">Tất cả ngành</option>';
  [...set].sort().forEach((v) => {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    els.filterIndustry.appendChild(o);
  });
  if ([...set].includes(prev)) els.filterIndustry.value = prev;
}

function renderStats() {
  const c = statusCounts();
  els.stats.textContent = `Tổng ${candidates.length} ứng viên — Mới: ${c.new} | Chờ HR: ${c.pending_hr} | Đã duyệt: ${c.approved} | Đã gửi HRM: ${c.sent_hrm} | Hoàn tất: ${c.done}`;
}

function renderList() {
  const list = filteredCandidates();
  els.candidateList.innerHTML = "";

  if (!list.length) {
    const p = document.createElement("p");
    p.className = "hint";
    if (sheetFetchLoading && !candidates.length) {
      p.textContent = "Đang tải danh sách từ Google Sheet...";
    } else if (!candidates.length && lastSheetError) {
      p.textContent = `Không tải được sheet: ${lastSheetError} Bạn có thể thử «Tải lại» hoặc nhập Excel.`;
    } else if (!candidates.length) {
      p.textContent = "Chưa có dữ liệu. Nhấn «Tải lại từ Google Sheet» hoặc chọn file Excel.";
    } else {
      p.textContent = "Không có ứng viên khớp bộ lọc.";
    }
    els.candidateList.appendChild(p);
    return;
  }

  for (const c of list) {
    const entry = ensureEntry(c.id);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "candidate-item" + (c.id === selectedId ? " active" : "");
    btn.innerHTML = `
      <div class="name"></div>
      <div class="meta"></div>
      <span class="badge ${entry.status}"></span>
    `;
    btn.querySelector(".name").textContent = c.displayName;
    btn.querySelector(".meta").textContent = `${c.position} · ${c.industry}`;
    btn.querySelector(".badge").textContent = STATUS_LABELS[entry.status] || entry.status;
    btn.addEventListener("click", () => {
      selectedId = c.id;
      renderList();
      renderDetail();
    });
    els.candidateList.appendChild(btn);
  }
}

function renderDetail() {
  const c = candidates.find((x) => x.id === selectedId);
  if (!c) {
    els.emptyDetail.classList.remove("hidden");
    els.detail.classList.add("hidden");
    els.detail.innerHTML = "";
    return;
  }

  els.emptyDetail.classList.add("hidden");
  els.detail.classList.remove("hidden");
  const entry = ensureEntry(c.id);

  const questions =
    entry.questions?.length > 0
      ? entry.questions
      : (entry.questionsText || "")
          .split(/\n+/)
          .map((s) => s.replace(/^\d+[\).\s]+/, "").trim())
          .filter(Boolean);

  const canApprove = entry.status === "pending_hr" && questions.length > 0;
  const canSend =
    entry.status === "approved" && questions.length > 0;
  const canFinalize = entry.status === "sent_hrm";

  els.detail.innerHTML = `
    <div class="detail-header">
      <div>
        <h3></h3>
        <p class="hint" style="margin:0.35rem 0 0"></p>
      </div>
      <div class="actions">
        <button type="button" class="btn btn-primary" id="btnGen"></button>
      </div>
    </div>

    <div class="panel" style="margin:0;padding:0.85rem;border:1px solid rgba(255,255,255,0.08)">
      <h4 class="section-title">Câu hỏi phỏng vấn (Gemini)</h4>
      <div class="questions-editor" id="questionsMount"></div>
    </div>

    <div class="panel hr-section" style="margin:0;padding:0.85rem;border:1px solid rgba(255,255,255,0.08)">
      <h4 class="section-title">HR duyệt</h4>
      <label>Ghi chú HR (tùy chọn)
        <textarea id="hrComment"></textarea>
      </label>
      <div class="actions" style="margin-top:0.5rem">
        <button type="button" class="btn btn-success" id="btnApprove" ${canApprove ? "" : "disabled"}>Duyệt câu hỏi</button>
        <button type="button" class="btn btn-danger" id="btnReject" ${entry.status === "pending_hr" ? "" : "disabled"}>Từ chối / gửi lại</button>
      </div>
    </div>

    <div class="panel hrm-section" style="margin:0;padding:0.85rem;border:1px solid rgba(255,255,255,0.08)">
      <h4 class="section-title">Gửi lên hệ thống HR (ATS/HRM)</h4>
      <p class="hint" id="hrmHint"></p>
      <div class="actions">
        <button type="button" class="btn btn-primary" id="btnSendHrm" ${canSend ? "" : "disabled"}>Gửi dữ liệu</button>
      </div>
      <pre id="hrmOut" class="hint hidden" style="white-space:pre-wrap;max-height:140px;overflow:auto;margin-top:0.5rem"></pre>
    </div>

    <div class="panel report-section" style="margin:0;padding:0.85rem;border:1px solid rgba(255,255,255,0.08)">
      <h4 class="section-title">Cập nhật &amp; báo cáo</h4>
      <div class="actions">
        <button type="button" class="btn btn-secondary" id="btnReport" ${canFinalize ? "" : "disabled"}>Hoàn tất &amp; tải báo cáo JSON</button>
        <button type="button" class="btn btn-ghost" id="btnCsv">Xuất CSV tóm tất cả ứng viên</button>
      </div>
      <ul class="timeline" id="timeline"></ul>
    </div>

    <details>
      <summary class="hint" style="cursor:pointer">Dữ liệu gốc (Google Sheet / Excel)</summary>
      <div class="raw-fields" style="margin-top:0.5rem"></div>
    </details>
  `;

  els.detail.querySelector("h3").textContent = c.displayName;
  els.detail.querySelector(".detail-header .hint").textContent = `${c.position} · ${c.industry}`;

  const btnGen = els.detail.querySelector("#btnGen");
  btnGen.textContent =
    entry.status === "new"
      ? "Tạo câu hỏi (Gemini)"
      : entry.status === "pending_hr"
        ? "Tạo lại câu hỏi"
        : "Đã khóa (đã duyệt hoặc đã gửi)";
  btnGen.disabled = !(entry.status === "new" || entry.status === "pending_hr");

  const qMount = els.detail.querySelector("#questionsMount");
  const listQs = questions.length ? questions : templateQuestions(c);
  listQs.forEach((text, idx) => {
    const wrap = document.createElement("div");
    wrap.className = "question-block";
    wrap.innerHTML = `<label>Câu ${idx + 1}<textarea data-q="${idx}"></textarea></label>`;
    wrap.querySelector("textarea").value = text;
    qMount.appendChild(wrap);
  });

  const hrTa = els.detail.querySelector("#hrComment");
  hrTa.value = entry.hrComment || "";

  const hrmHint = els.detail.querySelector("#hrmHint");
  hrmHint.textContent = els.hrmUrl.value.trim()
    ? "Sẽ POST JSON tới URL bạn cấu hình (cần máy chủ cho phép CORS từ origin này)."
    : "Chưa có URL — bước gửi sẽ mô phỏng thành công (ghi log).";

  const hrmOut = els.detail.querySelector("#hrmOut");
  if (entry.hrmResponse) {
    hrmOut.classList.remove("hidden");
    hrmOut.textContent =
      typeof entry.hrmResponse === "string"
        ? entry.hrmResponse
        : JSON.stringify(entry.hrmResponse, null, 2);
  }

  const rawBox = els.detail.querySelector(".raw-fields");
  for (const [k, v] of Object.entries(c.raw)) {
    const d = document.createElement("div");
    d.innerHTML = `<strong></strong><span></span>`;
    d.querySelector("strong").textContent = k;
    d.querySelector("span").textContent = String(v ?? "");
    rawBox.appendChild(d);
  }

  const tl = els.detail.querySelector("#timeline");
  (entry.log || []).slice(-12).reverse().forEach((l) => {
    const li = document.createElement("li");
    li.textContent = `${l.t} — ${l.message}`;
    tl.appendChild(li);
  });

  btnGen.addEventListener("click", async () => {
    btnGen.disabled = true;
    btnGen.textContent = "Đang tạo...";
    try {
      const qs = await generateQuestionsWithGemini(
        c,
        DEFAULT_GEMINI_API_KEY,
        DEFAULT_GEMINI_MODEL
      );
      entry.questions = qs;
      entry.questionsText = qs.join("\n");
      entry.status = "pending_hr";
      pushLog(entry, "Đã tạo câu hỏi phỏng vấn (Gemini).");
      saveState(state);
      persistSettings();
    } catch (e) {
      alert(String(e.message || e));
    }
    renderStats();
    renderList();
    renderDetail();
  });

  hrTa.addEventListener("change", () => {
    entry.hrComment = hrTa.value;
    saveState(state);
  });

  qMount.querySelectorAll("textarea[data-q]").forEach((ta) => {
    ta.addEventListener("change", () => {
      const idx = Number(ta.getAttribute("data-q"));
      const arr = [...qMount.querySelectorAll("textarea[data-q]")].map((x) =>
        x.value.trim()
      );
      entry.questions = arr;
      entry.questionsText = arr.join("\n");
      saveState(state);
    });
  });

  els.detail.querySelector("#btnApprove").addEventListener("click", () => {
    const arr = [...qMount.querySelectorAll("textarea[data-q]")].map((x) =>
      x.value.trim()
    ).filter(Boolean);
    if (!arr.length) {
      alert("Cần ít nhất một câu hỏi.");
      return;
    }
    entry.questions = arr;
    entry.status = "approved";
    entry.approvedAt = new Date().toISOString();
    entry.hrComment = hrTa.value;
    pushLog(entry, "HR đã duyệt câu hỏi.");
    saveState(state);
    renderStats();
    renderList();
    renderDetail();
  });

  els.detail.querySelector("#btnReject").addEventListener("click", () => {
    entry.status = "new";
    entry.rejectedAt = new Date().toISOString();
    pushLog(entry, "HR từ chối / yêu cầu tạo lại câu hỏi.");
    saveState(state);
    renderStats();
    renderList();
    renderDetail();
  });

  els.detail.querySelector("#btnSendHrm").addEventListener("click", async () => {
    const payload = {
      candidate: c,
      questions: entry.questions,
      hrComment: entry.hrComment,
      approvedAt: entry.approvedAt,
      clientTs: new Date().toISOString(),
    };
    const url = els.hrmUrl.value.trim();
    try {
      if (url) {
        const res = await fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        const txt = await res.text();
        entry.hrmResponse = { ok: res.ok, status: res.status, body: txt.slice(0, 4000) };
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
      } else {
        entry.hrmResponse = { mock: true, message: "Không cấu hình URL — mô phỏng thành công." };
      }
      entry.status = "sent_hrm";
      entry.hrmSentAt = new Date().toISOString();
      pushLog(entry, "Đã gửi sang hệ thống HR.");
      saveState(state);
      persistSettings();
    } catch (e) {
      alert(`Gửi HRM lỗi: ${e.message || e}`);
      entry.hrmResponse = { error: String(e.message || e) };
      saveState(state);
    }
    renderStats();
    renderList();
    renderDetail();
  });

  els.detail.querySelector("#btnReport").addEventListener("click", () => {
    const report = {
      generatedAt: new Date().toISOString(),
      candidate: c,
      workflow: entry,
    };
    entry.status = "done";
    entry.reportAt = new Date().toISOString();
    pushLog(entry, "Đã hoàn tất và xuất báo cáo.");
    saveState(state);
    downloadBlob(
      `bao_cao_${norm(c.displayName).replace(/\s+/g, "_")}.json`,
      new Blob([JSON.stringify(report, null, 2)], { type: "application/json" })
    );
    renderStats();
    renderList();
    renderDetail();
  });

  els.detail.querySelector("#btnCsv").addEventListener("click", () => {
    const headers = [
      "id",
      "name",
      "industry",
      "position",
      "status",
      "questions",
      "approvedAt",
      "hrmSentAt",
      "reportAt",
    ];
    const lines = [headers.join(",")];
    for (const cand of candidates) {
      const e = ensureEntry(cand.id);
      const row = [
        cand.id,
        cand.displayName,
        cand.industry,
        cand.position,
        e.status,
        JSON.stringify((e.questions || []).join(" | ")),
        e.approvedAt || "",
        e.hrmSentAt || "",
        e.reportAt || "",
      ].map((cell) => {
        const s = String(cell ?? "").replace(/"/g, '""');
        return `"${s}"`;
      });
      lines.push(row.join(","));
    }
    downloadBlob(
      `tom_tat_ung_vien_${new Date().toISOString().slice(0, 10)}.csv`,
      new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" })
    );
  });
}

els.fileInput.addEventListener("change", async (ev) => {
  const file = ev.target.files?.[0];
  if (!file) return;
  const buf = await file.arrayBuffer();
  try {
    candidates = parseWorkbook(buf);
    lastSheetError = "";
    if (els.sheetHint) {
      els.sheetHint.textContent = candidates.length
        ? `Đang dùng file Excel (${candidates.length} dòng).`
        : "";
    }
    if (!candidates.length) {
      alert("Không đọc được dòng dữ liệu nào. Kiểm tra dòng tiêu đề cột trong Excel.");
    }
    selectedId = candidates[0]?.id || null;
    renderIndustryFilter();
    renderStats();
    renderList();
    renderDetail();
  } catch (e) {
    alert("Lỗi đọc Excel: " + (e.message || e));
  }
  ev.target.value = "";
});

async function refreshCandidatesFromGoogleSheet() {
  sheetFetchLoading = true;
  lastSheetError = "";
  if (els.sheetHint) els.sheetHint.textContent = "Đang tải...";
  renderList();
  renderDetail();
  const btn = els.btnReloadSheet;
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Đang tải...";
  }
  try {
    const { candidates: list, usedProxy } = await loadCandidatesFromGoogleSheet(
      DEFAULT_GOOGLE_SHEET_PUB_URL
    );
    candidates = list;
    selectedId = candidates[0]?.id || null;
    if (!candidates.length) {
      lastSheetError = "Không có dòng dữ liệu (kiểm tra dòng tiêu đề cột).";
    }
    if (els.sheetHint) {
      const proxyNote = usedProxy ? " (qua proxy vì trình duyệt chặn tải trực tiếp)." : "";
      els.sheetHint.textContent = lastSheetError
        ? lastSheetError
        : `Đã tải ${candidates.length} ứng viên từ Google Sheet.${proxyNote}`;
    }
  } catch (e) {
    lastSheetError = String(e.message || e);
    candidates = [];
    selectedId = null;
    if (els.sheetHint) {
      els.sheetHint.textContent = `Lỗi: ${lastSheetError}`;
    }
  }
  sheetFetchLoading = false;
  if (btn) {
    btn.disabled = false;
    btn.textContent = "Tải lại từ Google Sheet";
  }
  renderIndustryFilter();
  renderStats();
  renderList();
  renderDetail();
}

if (els.btnReloadSheet) {
  els.btnReloadSheet.addEventListener("click", () => {
    refreshCandidatesFromGoogleSheet();
  });
}

["change", "input"].forEach((evt) => {
  els.filterIndustry.addEventListener(evt, () => {
    renderList();
  });
  els.filterStatus.addEventListener(evt, () => {
    renderList();
  });
  els.search.addEventListener(evt, () => {
    renderList();
  });
});

["change", "blur"].forEach((evt) => {
  els.hrmUrl.addEventListener(evt, persistSettings);
});

els.btnExportState.addEventListener("click", () => {
  const blob = new Blob([JSON.stringify({ ...state, exportedAt: new Date().toISOString() }, null, 2)], {
    type: "application/json",
  });
  downloadBlob("tuyen_dung_trang_thai.json", blob);
});

els.btnImportState.addEventListener("click", () => els.stateImport.click());

els.stateImport.addEventListener("change", async (ev) => {
  const f = ev.target.files?.[0];
  if (!f) return;
  try {
    const text = await f.text();
    const parsed = JSON.parse(text);
    if (!parsed.byId) throw new Error("File không hợp lệ.");
    state = parsed;
    saveState(state);
    hydrateSettings();
    renderStats();
    renderList();
    renderDetail();
  } catch (e) {
    alert("Nhập trạng thái lỗi: " + (e.message || e));
  }
  ev.target.value = "";
});

hydrateSettings();
renderIndustryFilter();
renderStats();
renderList();
renderDetail();
refreshCandidatesFromGoogleSheet();
