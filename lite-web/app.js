const PERIOD_OPTIONS = ['1', '2', '3', '4', '5', '6', '7', '8', '9'];
const HOURS_OPTIONS = ['1', '2', '3', '4', '5', '6', '7', '8', '9'];
const LAST_SUBJECT_KEY = 'liteTitleMaker:lastSubject';

const state = {
  mode: 'lecture',
  noteAuthor: '',
  entries: [],
  files: [],
  lastSelectedSubject: '',
  manualRowSeq: 0,
};

const els = {
  formatText: document.getElementById('formatText'),
  entryCount: document.getElementById('entryCount'),
  modeSwitch: document.getElementById('modeSwitch'),
  lectureModeBtn: document.getElementById('lectureModeBtn'),
  noteModeBtn: document.getElementById('noteModeBtn'),
  noteAuthorSection: document.getElementById('noteAuthorSection'),
  noteAuthorInput: document.getElementById('noteAuthorInput'),
  rowsBody: document.getElementById('rowsBody'),
  addRowBtn: document.getElementById('addRowBtn'),
  toggleAllBtn: document.getElementById('toggleAllBtn'),
  clearBtn: document.getElementById('clearBtn'),
  copyBtn: document.getElementById('copyBtn'),
  csvBtn: document.getElementById('csvBtn'),
  exportArea: document.getElementById('exportArea'),
};

init().catch((error) => {
  window.alert(`초기화 실패: ${error.message}`);
});

async function init() {
  loadLastSelectedSubject();
  bindEvents();
  await loadDefaultTimetable();
  renderMode();
  ensureDefaultRow();
  renderRows();
  refreshExportText();
}

function bindEvents() {
  els.lectureModeBtn.addEventListener('click', () => setMode('lecture'));
  els.noteModeBtn.addEventListener('click', () => setMode('note'));
  els.noteAuthorInput.addEventListener('input', (event) => {
    state.noteAuthor = String(event.target.value ?? '').trim();
    for (const file of state.files) {
      updateAutoTargetName(file);
    }
    renderRows();
    refreshExportText();
  });

  els.rowsBody.addEventListener('change', onTableChange);
  els.rowsBody.addEventListener('input', onTableInput);

  els.addRowBtn.addEventListener('click', addManualRow);
  els.toggleAllBtn.addEventListener('click', toggleAllRows);
  els.clearBtn.addEventListener('click', () => {
    state.files = [];
    ensureDefaultRow();
    renderRows();
    refreshExportText();
  });
  els.copyBtn.addEventListener('click', copySelectedNames);
  els.csvBtn.addEventListener('click', exportCsv);
}

async function loadDefaultTimetable() {
  const bundled = window.DEFAULT_TIMETABLE;
  if (bundled && Array.isArray(bundled.entries)) {
    state.entries = bundled.entries;
    els.entryCount.textContent = `${state.entries.length}개 강의`;
    return;
  }

  try {
    const response = await fetch('./default-timetable.json');
    if (!response.ok) {
      throw new Error('default-timetable.json을 불러올 수 없습니다.');
    }
    const json = await response.json();
    state.entries = Array.isArray(json?.entries) ? json.entries : [];
    els.entryCount.textContent = `${state.entries.length}개 강의`;
  } catch (error) {
    throw new Error(`시간표 로드 실패: ${error.message}`);
  }
}

function setMode(next) {
  if (next !== 'lecture' && next !== 'note') {
    return;
  }
  if (state.mode === next) {
    return;
  }
  state.mode = next;
  renderMode();
  for (const file of state.files) {
    if (state.mode === 'note') {
      file.hours = '1';
    }
    updateAutoTargetName(file);
  }
  renderRows();
  refreshExportText();
}

function renderMode() {
  const isNote = state.mode === 'note';
  els.modeSwitch?.classList.toggle('lecture', !isNote);
  els.modeSwitch?.classList.toggle('note', isNote);
  els.lectureModeBtn.classList.toggle('active', !isNote);
  els.noteModeBtn.classList.toggle('active', isNote);
  els.noteAuthorSection.classList.toggle('hidden', !isNote);

  if (isNote) {
    els.formatText.textContent =
      '필족 형식: [수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목]-[작성자명].docx';
  } else {
    els.formatText.textContent =
      '강의자료 형식: [수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목].pdf';
  }
}

function ensureDefaultRow() {
  if (state.files.length > 0) {
    return;
  }
  state.manualRowSeq += 1;
  const next = createFileModel({
    sourceName: `manual-${state.manualRowSeq}.pdf`,
    sourceLabel: `직접입력-${state.manualRowSeq}`,
  });
  updateAutoTargetName(next);
  state.files.push(next);
}

function addManualRow() {
  state.manualRowSeq += 1;
  const next = createFileModel({
    sourceName: `manual-${state.manualRowSeq}.pdf`,
    sourceLabel: `직접입력-${state.manualRowSeq}`,
  });
  updateAutoTargetName(next);
  state.files.push(next);
  renderRows();
  refreshExportText();
}

function createFileModel({ sourceName, sourceLabel }) {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    sourceName,
    sourceLabel,
    enabled: true,
    subject: resolveDefaultSubject(),
    order: '',
    classDate: '',
    periodStart: '',
    hours: '1',
    professor: '',
    title: '',
    targetName: sourceName,
    targetEdited: false,
  };
}

function onTableChange(event) {
  const action = event.target.dataset.action;
  const id = event.target.dataset.id;
  const file = state.files.find((item) => item.id === id);
  if (!file || !action) {
    return;
  }

  if (action === 'enabled') {
    file.enabled = event.target.checked;
    refreshExportText();
    return;
  }

  if (action === 'subject') {
    file.subject = event.target.value;
    rememberLastSelectedSubject(file.subject);
    const orderOptions = getOrderOptions(file.subject);
    if (!orderOptions.includes(file.order)) {
      file.order = '';
      clearDerivedFields(file);
    }
    applyPresetFromSelection(file);
    updateAutoTargetName(file);
    renderRows();
    refreshExportText();
    return;
  }

  if (action === 'order') {
    file.order = event.target.value;
    applyPresetFromSelection(file);
    updateAutoTargetName(file);
    renderRows();
    refreshExportText();
    return;
  }

  if (action === 'date') file.classDate = event.target.value;
  if (action === 'period') file.periodStart = event.target.value;
  if (action === 'hours') {
    file.hours = state.mode === 'note' ? '1' : event.target.value;
  }
  if (action === 'target') {
    file.targetName = event.target.value;
    file.targetEdited = true;
    refreshExportText();
    return;
  }

  updateAutoTargetName(file);
  refreshExportText();
}

function onTableInput(event) {
  const action = event.target.dataset.action;
  const id = event.target.dataset.id;
  const file = state.files.find((item) => item.id === id);
  if (!file || !action) {
    return;
  }

  if (action === 'professor') file.professor = event.target.value;
  if (action === 'title') file.title = event.target.value;
  if (action === 'target') {
    file.targetName = event.target.value;
    file.targetEdited = true;
    refreshExportText();
    return;
  }

  updateAutoTargetName(file);
  refreshExportText();
}

function renderRows() {
  els.rowsBody.innerHTML = '';
  const subjects = getSubjectOptions();

  if (state.files.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 9;
    td.textContent = '"행 추가" 버튼으로 제목 행을 생성하세요.';
    td.style.color = '#5f6e88';
    tr.appendChild(td);
    els.rowsBody.appendChild(tr);
    return;
  }

  for (const file of state.files) {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td><input type="checkbox" data-action="enabled" data-id="${escapeHtml(file.id)}" ${file.enabled ? 'checked' : ''}></td>
      <td>${escapeHtml(file.sourceLabel || file.sourceName)}</td>
      <td>${renderSubjectSelect(file, subjects)}</td>
      <td>${renderOrderSelect(file)}</td>
      <td><input type="date" data-action="date" data-id="${escapeHtml(file.id)}" value="${escapeHtml(file.classDate)}"></td>
      <td>${renderOptionSelect('period', file.id, file.periodStart, PERIOD_OPTIONS, '시작교시')}</td>
      <td>${renderOptionSelect('hours', file.id, state.mode === 'note' ? '1' : file.hours, state.mode === 'note' ? ['1'] : HOURS_OPTIONS, '시수', state.mode === 'note')}</td>
      <td><input data-action="professor" data-id="${escapeHtml(file.id)}" value="${escapeHtml(file.professor)}" placeholder="예: 차동훈"></td>
      <td><input data-action="title" data-id="${escapeHtml(file.id)}" value="${escapeHtml(file.title)}" placeholder="강의제목"></td>
    `;
    els.rowsBody.appendChild(row);

    const nameRow = document.createElement('tr');
    nameRow.className = 'name-row';
    nameRow.innerHTML = `
      <td colspan="9">
        <span class="name-label">최종 파일명 검토</span>
        <input data-action="target" data-id="${escapeHtml(file.id)}" value="${escapeHtml(file.targetName)}">
      </td>
    `;
    els.rowsBody.appendChild(nameRow);
  }
}

function renderSubjectSelect(file, subjects) {
  const id = escapeHtml(file.id);
  let options = '<option value="">과목 선택</option>';
  for (const subject of subjects) {
    const selected = subject === file.subject ? 'selected' : '';
    options += `<option value="${escapeHtml(subject)}" ${selected}>${escapeHtml(subject)}</option>`;
  }
  return `<select data-action="subject" data-id="${id}">${options}</select>`;
}

function renderOrderSelect(file) {
  const id = escapeHtml(file.id);
  const options = getOrderOptions(file.subject);
  let html = '<option value="">순서 선택</option>';
  for (const value of options) {
    const selected = value === String(file.order ?? '') ? 'selected' : '';
    html += `<option value="${escapeHtml(value)}" ${selected}>${escapeHtml(value)}</option>`;
  }
  if (file.order && !options.includes(String(file.order))) {
    html += `<option value="${escapeHtml(file.order)}" selected>${escapeHtml(file.order)}</option>`;
  }
  return `<select data-action="order" data-id="${id}">${html}</select>`;
}

function renderOptionSelect(action, idRaw, currentValue, options, placeholder, disabled = false) {
  const id = escapeHtml(idRaw);
  let html = `<option value="">${escapeHtml(placeholder)}</option>`;
  for (const option of options) {
    const selected = option === String(currentValue ?? '') ? 'selected' : '';
    html += `<option value="${option}" ${selected}>${option}</option>`;
  }
  return `<select data-action="${action}" data-id="${id}" ${disabled ? 'disabled' : ''}>${html}</select>`;
}

function toggleAllRows() {
  if (state.files.length === 0) {
    return;
  }
  const hasUnchecked = state.files.some((file) => !file.enabled);
  for (const file of state.files) {
    file.enabled = hasUnchecked;
  }
  renderRows();
  refreshExportText();
}

function getSelectedNames() {
  return state.files
    .filter((file) => file.enabled)
    .map((file) => String(file.targetName ?? '').trim())
    .filter(Boolean);
}

function refreshExportText() {
  els.exportArea.value = getSelectedNames().join('\n');
}

async function copySelectedNames() {
  const text = getSelectedNames().join('\n');
  if (!text) {
    window.alert('복사할 파일명이 없습니다.');
    return;
  }

  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    window.alert('복사 완료');
    return;
  }

  els.exportArea.focus();
  els.exportArea.select();
  document.execCommand('copy');
  window.alert('복사 완료');
}

function exportCsv() {
  const rows = state.files
    .filter((file) => file.enabled)
    .map((file) => [file.sourceLabel || file.sourceName, file.targetName]);

  if (rows.length === 0) {
    window.alert('CSV로 내보낼 항목이 없습니다.');
    return;
  }

  const lines = [['입력 행', '생성 파일명'], ...rows].map((cols) =>
    cols.map((value) => `"${String(value ?? '').replace(/"/g, '""')}"`).join(','),
  );
  const csv = `\uFEFF${lines.join('\n')}`;
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `title-maker-${state.mode}-${todayToken()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function todayToken() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}${m}${d}`;
}

function updateAutoTargetName(file) {
  if (file.targetEdited) {
    return;
  }
  file.targetName = buildTargetName(file);
}

function buildTargetName(file) {
  const subject = String(file.subject ?? '').trim();
  const order = normalizeOrderToken(file.order);
  const dateToken = formatDateToken(file.classDate);
  const period = buildPeriodToken(file.periodStart, file.hours);
  const professorRaw = String(file.professor ?? '').trim();
  const title = String(file.title ?? '').trim();
  const author = String(state.noteAuthor ?? '').trim();

  if (!subject || !order || !dateToken || !period || !professorRaw || !title) {
    return file.targetName || file.sourceName;
  }

  if (state.mode === 'note' && !author) {
    return file.targetName || file.sourceName;
  }

  const professor = normalizeProfessor(professorRaw);
  if (state.mode === 'note') {
    const rawName = `${subject}${order}-${dateToken}-${period}교시-${professor}-${title}-${author}.docx`;
    return sanitizeOutputName(rawName);
  }

  const rawName = `${subject}${order}-${dateToken}-${period}교시-${professor}-${title}.pdf`;
  return sanitizeOutputName(rawName);
}

function clearDerivedFields(file) {
  file.classDate = '';
  file.periodStart = '';
  file.hours = '1';
  file.professor = '';
  file.title = '';
}

function applyPresetFromSelection(file) {
  const entry = findEntry(file.subject, file.order);
  if (!entry) {
    return false;
  }
  file.classDate = toDateInputValue(entry.dateToken);
  file.periodStart = String(entry.period ?? '');
  file.hours = state.mode === 'note' ? '1' : String(entry.hours ?? '1');
  file.professor = String(entry.professor ?? '');
  file.title = String(entry.lectureTitle ?? '');
  return true;
}

function findEntry(subject, orderValue) {
  const subjectText = String(subject ?? '').trim();
  const orderText = String(orderValue ?? '').trim();
  if (!subjectText || !orderText) {
    return null;
  }
  const first = orderText.split(',')[0];
  const normalized = normalizeSingleOrder(first);
  if (!normalized) {
    return null;
  }
  return state.entries.find(
    (entry) => String(entry.subject) === subjectText && String(entry.orderNumber) === normalized,
  ) ?? null;
}

function normalizeSingleOrder(value) {
  const num = Number(String(value ?? '').trim());
  if (!Number.isFinite(num) || num <= 0) {
    return '';
  }
  return String(Math.round(num)).padStart(2, '0');
}

function normalizeOrderToken(value) {
  const text = String(value ?? '').trim();
  if (!text) {
    return '';
  }
  const tokens = text
    .split(',')
    .map((token) => token.trim())
    .filter(Boolean)
    .map((token) => Number(token))
    .filter((num) => Number.isFinite(num) && num > 0)
    .map((num) => String(Math.round(num)).padStart(2, '0'));
  return tokens.length > 0 ? tokens.join(',') : '';
}

function formatDateToken(yyyyMmDd) {
  const text = String(yyyyMmDd ?? '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return '';
  }
  const [yyyy, mm, dd] = text.split('-');
  return `${yyyy.slice(-2)}${mm}${dd}`;
}

function buildPeriodToken(startPeriod, hours) {
  const start = Number(startPeriod);
  const duration = Number(hours);
  if (!Number.isFinite(start) || start < 1 || start > 9) {
    return '';
  }
  const count = Number.isFinite(duration) && duration > 0 ? Math.round(duration) : 1;
  const values = [];
  for (let i = 0; i < count; i += 1) {
    const period = start + i;
    if (period > 9) {
      break;
    }
    values.push(String(period));
  }
  return values.join(',');
}

function normalizeProfessor(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (text.endsWith('교수님')) return text;
  if (text.endsWith('교수')) return `${text}님`;
  return `${text}교수님`;
}

function getSubjectOptions() {
  return unique(state.entries.map((entry) => entry.subject)).sort((a, b) => a.localeCompare(b, 'ko-KR'));
}

function getOrderOptions(subject) {
  const subjectText = String(subject ?? '').trim();
  if (!subjectText) {
    return [];
  }
  return unique(
    state.entries
      .filter((entry) => String(entry.subject) === subjectText)
      .map((entry) => String(entry.orderNumber)),
  ).sort((a, b) => Number(a) - Number(b));
}

function resolveDefaultSubject() {
  const value = String(state.lastSelectedSubject ?? '').trim();
  if (!value) {
    return '';
  }
  const subjects = getSubjectOptions();
  return subjects.includes(value) ? value : '';
}

function rememberLastSelectedSubject(subject) {
  const value = String(subject ?? '').trim();
  if (!value) {
    return;
  }
  state.lastSelectedSubject = value;
  try {
    localStorage.setItem(LAST_SUBJECT_KEY, value);
  } catch {
    // ignore
  }
}

function loadLastSelectedSubject() {
  try {
    state.lastSelectedSubject = String(localStorage.getItem(LAST_SUBJECT_KEY) ?? '').trim();
  } catch {
    state.lastSelectedSubject = '';
  }
}

function toDateInputValue(dateToken) {
  const token = String(dateToken ?? '').trim();
  if (!/^\d{6}$/.test(token)) {
    return '';
  }
  const yy = Number(token.slice(0, 2));
  const mm = token.slice(2, 4);
  const dd = token.slice(4, 6);
  return `${2000 + yy}-${mm}-${dd}`;
}

function sanitizeOutputName(filename) {
  return String(filename ?? '')
    .replace(/[\\/]/g, '-')
    .replace(/[<>:"|?*\x00-\x1F]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/-+/g, '-')
    .trim();
}

function unique(values) {
  return [...new Set(values)];
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
