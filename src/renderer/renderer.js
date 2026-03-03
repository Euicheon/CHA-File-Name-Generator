const PERIOD_OPTIONS = ['1', '2', '3', '4', '5', '6', '7', '8', '9'];
const HOURS_OPTIONS = ['1', '2', '3', '4', '5', '6', '7', '8', '9'];
const LAST_SUBJECT_STORAGE_KEY = 'titleMaker:lastSelectedSubject';

const state = {
  mode: 'lecture',
  noteAuthor: '',
  lastSelectedSubject: '',
  timetable: null,
  outputDir: null,
  files: [],
  customPlanPath: '',
  adminYear: new Date().getFullYear(),
  adminRows: [],
};

const els = {
  subtitleText: document.getElementById('subtitleText'),
  modeSwitch: document.getElementById('modeSwitch'),
  lectureModeBtn: document.getElementById('lectureModeBtn'),
  noteModeBtn: document.getElementById('noteModeBtn'),
  selectTimetableBtn: document.getElementById('selectTimetableBtn'),
  openAdminBtn: document.getElementById('openAdminBtn'),
  selectOutputBtn: document.getElementById('selectOutputBtn'),
  noteModePanel: document.getElementById('noteModePanel'),
  noteAuthorInput: document.getElementById('noteAuthorInput'),
  pickFilesBtn: document.getElementById('pickFilesBtn'),
  filePicker: document.getElementById('filePicker'),
  dropZone: document.getElementById('dropZone'),
  dropTitle: document.getElementById('dropTitle'),
  dropSubtitle: document.getElementById('dropSubtitle'),
  timetablePath: document.getElementById('timetablePath'),
  lectureCount: document.getElementById('lectureCount'),
  outputPath: document.getElementById('outputPath'),
  fileTableBody: document.getElementById('fileTableBody'),
  toggleAllBtn: document.getElementById('toggleAllBtn'),
  clearBtn: document.getElementById('clearBtn'),
  runBtn: document.getElementById('runBtn'),
  logOutput: document.getElementById('logOutput'),
  adminModal: document.getElementById('adminModal'),
  closeAdminBtn: document.getElementById('closeAdminBtn'),
  adminYearInput: document.getElementById('adminYearInput'),
  customPlanPath: document.getElementById('customPlanPath'),
  addAdminRowBtn: document.getElementById('addAdminRowBtn'),
  deleteAdminRowBtn: document.getElementById('deleteAdminRowBtn'),
  saveAdminBtn: document.getElementById('saveAdminBtn'),
  resetAdminBtn: document.getElementById('resetAdminBtn'),
  adminTableBody: document.getElementById('adminTableBody'),
  subtleVersion: document.getElementById('subtleVersion'),
};

init().catch((error) => {
  writeLog(`초기화 오류: ${error.message}`);
});

async function init() {
  loadLastSelectedSubject();
  bindEvents();
  renderModeUI();
  renderMeta();
  renderTable();
  renderAppVersion();

  const initial = await window.titleMaker.getInitialState();
  if (initial?.error) {
    writeLog(initial.error);
  }

  if (initial?.customPlan?.path) {
    state.customPlanPath = initial.customPlan.path;
  }

  if (initial?.timetable) {
    applyTimetable(initial.timetable);
    let modeText = '기본 시간표 적용';
    if (initial?.customPlan?.active) {
      modeText = '사용자 시간표 적용';
    } else if (initial?.customPlan?.baseSource === 'bundled') {
      modeText = '내장 시간표 적용';
    }
    writeLog(`${modeText} (${initial.timetable.entryCount}개 강의)`);
  } else {
    writeLog('시간표 파일을 먼저 선택해 주세요.');
  }
}

async function renderAppVersion() {
  if (!els.subtleVersion || !window.titleMaker?.getAppVersion) {
    return;
  }

  try {
    const version = await window.titleMaker.getAppVersion();
    els.subtleVersion.textContent = version ? `v${version}` : 'v-';
  } catch {
    els.subtleVersion.textContent = 'v-';
  }
}

function bindEvents() {
  els.lectureModeBtn.addEventListener('click', () => setMode('lecture'));
  els.noteModeBtn.addEventListener('click', () => setMode('note'));
  els.noteAuthorInput.addEventListener('input', onNoteAuthorInput);
  els.selectTimetableBtn.addEventListener('click', onSelectTimetable);
  els.openAdminBtn.addEventListener('click', openAdminModal);
  els.closeAdminBtn.addEventListener('click', closeAdminModal);
  els.adminYearInput.addEventListener('change', onAdminYearChange);
  els.selectOutputBtn.addEventListener('click', onSelectOutputDir);
  els.pickFilesBtn.addEventListener('click', () => els.filePicker.click());
  els.filePicker.addEventListener('change', onFilePickerChange);

  els.addAdminRowBtn.addEventListener('click', onAddAdminRow);
  els.deleteAdminRowBtn.addEventListener('click', onDeleteAdminRows);
  els.saveAdminBtn.addEventListener('click', onSaveAdminPlan);
  els.resetAdminBtn.addEventListener('click', onClearAdminPlan);

  bindDropZoneEvents();

  els.fileTableBody.addEventListener('change', onTableChange);
  els.fileTableBody.addEventListener('input', onTableInput);
  els.adminTableBody.addEventListener('change', onAdminTableChange);
  els.adminTableBody.addEventListener('input', onAdminTableInput);

  els.toggleAllBtn.addEventListener('click', toggleAllRows);
  els.clearBtn.addEventListener('click', clearRows);
  els.runBtn.addEventListener('click', runRename);
}

function setMode(nextMode) {
  if (nextMode !== 'lecture' && nextMode !== 'note') {
    return;
  }

  if (state.mode === nextMode) {
    return;
  }

  state.mode = nextMode;
  state.files = [];
  renderModeUI();
  renderTable();

  if (nextMode === 'note') {
    writeLog('필족 모드로 전환했습니다. (.docx 입력 후 제목 규칙으로 저장)');
  } else {
    writeLog('강의자료 모드로 전환했습니다.');
  }
}

function onNoteAuthorInput(event) {
  state.noteAuthor = String(event.target.value ?? '').trim();
  for (const file of state.files) {
    updateAutoTargetName(file);
  }
  renderTable();
}

function renderModeUI() {
  const isNoteMode = state.mode === 'note';

  els.modeSwitch?.classList.toggle('lecture', !isNoteMode);
  els.modeSwitch?.classList.toggle('note', isNoteMode);
  els.lectureModeBtn.classList.toggle('active', !isNoteMode);
  els.noteModeBtn.classList.toggle('active', isNoteMode);
  els.noteModePanel.classList.toggle('hidden', !isNoteMode);
  els.openAdminBtn.disabled = isNoteMode;
  els.noteAuthorInput.value = state.noteAuthor;

  if (isNoteMode) {
    els.filePicker.setAttribute('accept', '.docx');
    els.subtitleText.textContent =
      '필족 형식: [수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목]-[작성자명].docx';
    els.dropTitle.textContent = '여기에 필족 원본 DOCX 파일을 드래그앤드롭';
    els.dropSubtitle.textContent = '저장 시 제목 규칙에 맞춘 DOCX 파일을 만듭니다.';
    els.runBtn.textContent = '필족 저장 (DOCX)';
    return;
  }

  els.filePicker.removeAttribute('accept');
  els.subtitleText.textContent =
    '강의자료 형식: [수업명][순서]-[수업일]-[교시]-[교수님성함]-[강의제목].pdf';
  els.dropTitle.textContent = '여기에 교수님 파일을 드래그앤드롭';
  els.dropSubtitle.textContent = '또는 아래 버튼으로 선택';
  els.runBtn.textContent = '파일명 변환 + 복사';
}

function bindDropZoneEvents() {
  const preventWindowDefaults = (event) => {
    event.preventDefault();
  };
  window.addEventListener('dragover', preventWindowDefaults);
  window.addEventListener('drop', preventWindowDefaults);

  const preventDefaults = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, preventDefaults);
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, () => {
      els.dropZone.classList.add('dragging');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    els.dropZone.addEventListener(eventName, () => {
      els.dropZone.classList.remove('dragging');
    });
  });

  els.dropZone.addEventListener('drop', async (event) => {
    const paths = extractFilePaths(event.dataTransfer?.files);
    await importFiles(paths);
  });
}

async function onSelectTimetable() {
  const selected = await window.titleMaker.selectTimetable();
  if (!selected) {
    return;
  }

  try {
    const timetable = await window.titleMaker.loadTimetable(selected);
    applyTimetable(timetable);
    writeLog(`시간표 로드 완료: ${selected}`);
  } catch (error) {
    writeLog(`시간표 로드 실패: ${error.message}`);
  }
}

async function onSelectOutputDir() {
  const selected = await window.titleMaker.selectOutputDir();
  if (!selected) {
    return;
  }

  state.outputDir = selected;
  renderMeta();
  writeLog(`출력 폴더 선택: ${selected}`);
}

async function onFilePickerChange(event) {
  const paths = extractFilePaths(event.target.files);
  await importFiles(paths);
  event.target.value = '';
}

function openAdminModal() {
  if (!state.timetable) {
    writeLog('시간표가 없어 관리자 창을 열 수 없습니다.');
    return;
  }

  state.adminYear = Number(state.timetable.year) || new Date().getFullYear();
  state.adminRows = (state.timetable.entries ?? []).map((entry) => ({
    remove: false,
    subject: entry.subject,
    prefix: entry.prefix,
    orderNumber: entry.orderNumber,
    classDate: toDateInputValue(entry, state.timetable.year),
    period: String(entry.period ?? ''),
    hours: String(entry.hours ?? 1),
    professor: entry.professor,
    lectureTitle: entry.lectureTitle,
  }));

  if (state.adminRows.length === 0) {
    state.adminRows.push(makeEmptyAdminRow());
  }

  els.adminYearInput.value = String(state.adminYear);
  els.customPlanPath.textContent = state.customPlanPath || '';
  els.adminModal.classList.remove('hidden');
  renderAdminTable();
}

function closeAdminModal() {
  els.adminModal.classList.add('hidden');
}

function onAdminYearChange(event) {
  const year = Number(event.target.value);
  if (Number.isFinite(year) && year >= 2000 && year <= 2099) {
    state.adminYear = year;
  }
}

function onAddAdminRow() {
  state.adminRows.push(makeEmptyAdminRow());
  renderAdminTable();
}

function onDeleteAdminRows() {
  const remaining = state.adminRows.filter((row) => !row.remove);
  state.adminRows = remaining.length > 0 ? remaining : [makeEmptyAdminRow()];
  renderAdminTable();
}

function onAdminTableChange(event) {
  const rowIndex = Number(event.target.dataset.row);
  const action = event.target.dataset.action;

  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= state.adminRows.length) {
    return;
  }

  const row = state.adminRows[rowIndex];
  if (!row || !action) {
    return;
  }

  if (action === 'remove') {
    row.remove = event.target.checked;
    return;
  }

  if (action === 'subject') row.subject = event.target.value;
  if (action === 'prefix') row.prefix = event.target.value;
  if (action === 'order') row.orderNumber = normalizeOrder(event.target.value);
  if (action === 'date') row.classDate = event.target.value;
  if (action === 'period') row.period = event.target.value;
  if (action === 'hours') row.hours = event.target.value;
  if (action === 'title') row.lectureTitle = event.target.value;

  if (action === 'order') {
    renderAdminTable();
  }
}

function onAdminTableInput(event) {
  const rowIndex = Number(event.target.dataset.row);
  const action = event.target.dataset.action;

  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= state.adminRows.length) {
    return;
  }

  const row = state.adminRows[rowIndex];
  if (!row || !action) {
    return;
  }

  if (action === 'professor') row.professor = event.target.value;
  if (action === 'subject') row.subject = event.target.value;
  if (action === 'prefix') row.prefix = event.target.value;
  if (action === 'title') row.lectureTitle = event.target.value;
}

async function onSaveAdminPlan() {
  if (state.adminRows.length === 0) {
    writeLog('저장할 항목이 없습니다.');
    return;
  }

  const rows = state.adminRows.filter((row) => !row.remove);
  const serializedEntries = [];

  for (const row of rows) {
    if (isAdminRowEmpty(row)) {
      continue;
    }

    const normalized = normalizeAdminRow(row, state.adminYear);
    if (!normalized.ok) {
      writeLog(`관리자 저장 오류: ${normalized.reason}`);
      return;
    }

    serializedEntries.push(normalized.entry);
  }

  if (serializedEntries.length === 0) {
    writeLog('유효한 항목이 없습니다.');
    return;
  }

  try {
    const result = await window.titleMaker.saveCustomPlan({
      year: state.adminYear,
      entries: serializedEntries,
    });

    state.customPlanPath = result?.path || state.customPlanPath;
    if (result?.timetable) {
      applyTimetable(result.timetable);
    }

    closeAdminModal();
    writeLog(`사용자 시간표 저장 완료 (${serializedEntries.length}개)`);
  } catch (error) {
    writeLog(`사용자 시간표 저장 실패: ${error.message}`);
  }
}

async function onClearAdminPlan() {
  try {
    await window.titleMaker.clearCustomPlan();
    const initial = await window.titleMaker.getInitialState();
    if (initial?.timetable) {
      applyTimetable(initial.timetable);
    }
    closeAdminModal();
    writeLog('사용자 시간표를 삭제하고 기본 시간표로 복원했습니다.');
  } catch (error) {
    writeLog(`사용자 시간표 삭제 실패: ${error.message}`);
  }
}

async function importFiles(paths) {
  let targetPaths = unique(paths).filter(Boolean);
  if (targetPaths.length === 0) {
    writeLog('드롭된 파일의 경로를 읽지 못했습니다. 파일을 다시 드롭해 주세요.');
    return;
  }

  if (state.mode === 'note') {
    const beforeCount = targetPaths.length;
    targetPaths = targetPaths.filter((filePath) => getFileExt(filePath).toLowerCase() === '.docx');
    const filteredCount = beforeCount - targetPaths.length;
    if (filteredCount > 0) {
      writeLog(`필족 모드에서는 .docx만 처리합니다. ${filteredCount}개 파일은 제외했습니다.`);
    }
    if (targetPaths.length === 0) {
      return;
    }
  }

  if (!state.timetable?.path) {
    writeLog('먼저 시간표 파일을 선택해 주세요.');
    return;
  }

  try {
    const suggestions = await window.titleMaker.buildSuggestions({
      timetablePath: state.timetable.path,
      filePaths: targetPaths,
    });

    if (suggestions.length === 0) {
      writeLog('처리 가능한 파일이 없었습니다. 폴더가 아닌 파일을 드롭해 주세요.');
      return;
    }

    mergeImportedFiles(suggestions);
    renderTable();
    writeLog(`파일 ${suggestions.length}개 추가`);
  } catch (error) {
    writeLog(`파일 분석 실패: ${error.message}`);
  }
}

function mergeImportedFiles(suggestions) {
  const byPath = new Map(state.files.map((item) => [item.sourcePath, item]));
  const entryById = new Map((state.timetable?.entries ?? []).map((entry) => [entry.id, entry]));

  for (const suggestion of suggestions) {
    const matchedEntry = suggestion.matchedEntryIds
      ?.map((id) => entryById.get(id))
      .find(Boolean);

    const existing = byPath.get(suggestion.sourcePath);
    if (existing) {
      if (matchedEntry) {
        applyEntryPreset(existing, matchedEntry);
      }
      updateAutoTargetName(existing);
      continue;
    }

    const next = createFileModel(suggestion);
    if (matchedEntry) {
      applyEntryPreset(next, matchedEntry);
    }
    next.targetName = buildTargetName(next);
    byPath.set(suggestion.sourcePath, next);
  }

  state.files = Array.from(byPath.values()).sort((a, b) => a.sourceName.localeCompare(b.sourceName, 'ko-KR'));
}

function createFileModel(suggestion) {
  const ext = getFileExt(suggestion.sourceName);
  const defaultSubject = resolveDefaultSubject();

  return {
    sourcePath: suggestion.sourcePath,
    sourceName: suggestion.sourceName,
    ext,
    enabled: true,
    subject: defaultSubject,
    order: '',
    classDate: '',
    periodStart: '',
    hours: '1',
    professor: '',
    title: '',
    targetName: suggestion.suggestedName,
    targetEdited: false,
  };
}

function applyEntryPreset(file, entry) {
  file.subject = entry.subject;
  file.order = entry.orderNumber;
  file.classDate = toDateInputValue(entry, state.timetable?.year);
  file.periodStart = String(entry.period ?? '');
  file.hours = state.mode === 'note' ? '1' : String(entry.hours ?? '1');
  file.professor = entry.professor;
  file.title = entry.lectureTitle;
}

function onTableChange(event) {
  const action = event.target.dataset.action;
  const rowId = event.target.dataset.id;
  if (!action || !rowId) {
    return;
  }

  const file = state.files.find((item) => item.sourcePath === rowId);
  if (!file) {
    return;
  }

  if (action === 'toggle') {
    file.enabled = event.target.checked;
    return;
  }

  if (action === 'subject') {
    file.subject = event.target.value;
    rememberLastSelectedSubject(file.subject);

    const orders = getOrderOptions(file.subject);
    if (!orders.includes(file.order)) {
      file.order = '';
      file.title = '';
    }

    updateAutoTargetName(file);
    renderTable();
    return;
  }

  if (action === 'order') {
    file.order = event.target.value;
    const entry = findEntry(file.subject, file.order);
    if (entry) {
      applyEntryPreset(file, entry);
    }
    updateAutoTargetName(file);
    renderTable();
    return;
  }

  if (action === 'date') {
    file.classDate = event.target.value;
    updateAutoTargetName(file);
    return;
  }

  if (action === 'period') {
    file.periodStart = event.target.value;
    updateAutoTargetName(file);
    return;
  }

  if (action === 'hours') {
    if (state.mode === 'note') {
      file.hours = '1';
      updateAutoTargetName(file);
      renderTable();
      return;
    }
    file.hours = event.target.value;
    updateAutoTargetName(file);
    return;
  }

  if (action === 'title') {
    file.title = event.target.value;
    updateAutoTargetName(file);
    return;
  }

  if (action === 'target') {
    file.targetName = event.target.value;
    file.targetEdited = true;
  }
}

function onTableInput(event) {
  const action = event.target.dataset.action;
  const rowId = event.target.dataset.id;
  if (!action || !rowId) {
    return;
  }

  const file = state.files.find((item) => item.sourcePath === rowId);
  if (!file) {
    return;
  }

  if (action === 'professor') {
    file.professor = event.target.value;
    updateAutoTargetName(file);
    return;
  }

  if (action === 'target') {
    file.targetName = event.target.value;
    file.targetEdited = true;
  }
}

function toggleAllRows() {
  if (state.files.length === 0) {
    return;
  }

  const hasUnchecked = state.files.some((item) => !item.enabled);
  for (const file of state.files) {
    file.enabled = hasUnchecked;
  }

  renderTable();
}

function clearRows() {
  state.files = [];
  renderTable();
  writeLog('목록을 비웠습니다.');
}

async function runRename() {
  const selectedFiles = state.files.filter((item) => item.enabled);
  const isNoteMode = state.mode === 'note';

  if (selectedFiles.length === 0) {
    writeLog('선택된 파일이 없습니다.');
    return;
  }

  if (isNoteMode && !state.noteAuthor) {
    writeLog('필족 모드에서는 작성자 이름을 먼저 입력해 주세요.');
    return;
  }

  for (const file of selectedFiles) {
    const missing = [];
    if (!file.subject) missing.push('과목명');
    if (!file.order) missing.push('순서');
    if (!file.classDate) missing.push('수업일');
    if (!file.periodStart) missing.push('시작 교시');
    if (!file.hours) missing.push('시수');
    if (!String(file.professor ?? '').trim()) missing.push('교수님 성함');
    if (!String(file.title ?? '').trim()) missing.push('강의 제목');

    if (missing.length > 0) {
      writeLog(`입력 누락: ${file.sourceName} (${missing.join(', ')})`);
      return;
    }

    if (!String(file.targetName ?? '').trim()) {
      writeLog(`파일명이 비어 있습니다: ${file.sourceName}`);
      return;
    }

    if (isNoteMode && String(file.ext ?? '').toLowerCase() !== '.docx') {
      writeLog(`필족 모드는 docx만 지원합니다: ${file.sourceName}`);
      return;
    }

    // 사용자가 최종 파일명을 직접 수정하지 않은 경우에는 저장 직전에 규칙명으로 다시 계산합니다.
    if (!file.targetEdited) {
      file.targetName = buildTargetName(file);
    }
  }

  if (!state.outputDir) {
    const selected = await window.titleMaker.selectOutputDir();
    if (!selected) {
      writeLog('출력 폴더 선택이 취소되었습니다.');
      return;
    }

    state.outputDir = selected;
    renderMeta();
  }

  try {
    const results = await window.titleMaker.copyRenamedFiles({
      outputDir: state.outputDir,
      tasks: selectedFiles.map((item) => ({
        sourcePath: item.sourcePath,
        targetName: item.targetName,
      })),
    });

    const copied = results.filter((item) => item.status === 'copied');
    const failed = results.filter((item) => item.status === 'failed');

    writeLog(`완료: 성공 ${copied.length}개 / 실패 ${failed.length}개`);

    for (const item of copied) {
      writeLog(`  + ${item.targetPath}`);
    }

    for (const item of failed) {
      writeLog(`  - 실패 (${item.sourcePath ?? 'unknown'}): ${item.error}`);
    }
  } catch (error) {
    writeLog(`복사 실행 실패: ${error.message}`);
  }
}

function applyTimetable(timetable) {
  state.timetable = timetable;
  state.adminYear = Number(timetable.year) || state.adminYear;
  renderMeta();

  for (const file of state.files) {
    updateAutoTargetName(file);
  }

  renderTable();
}

function renderMeta() {
  els.timetablePath.textContent = state.timetable?.path ?? '선택되지 않음';
  els.lectureCount.textContent = String(state.timetable?.entryCount ?? 0);
  els.outputPath.textContent = state.outputDir ?? '선택되지 않음';
}

function renderTable() {
  els.fileTableBody.innerHTML = '';
  const mainColumnCount = 9;

  if (state.files.length === 0) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = mainColumnCount;
    td.textContent =
      state.mode === 'note'
        ? 'docx 파일을 드래그앤드롭하면 행별로 과목/순서/일자/교시/시수와 작성자명 기준 파일명을 생성합니다.'
        : '파일을 드래그앤드롭하면 행별로 과목/순서/일자/교시/시수를 선택할 수 있습니다.';
    td.style.color = '#5f6f8a';
    tr.appendChild(td);
    els.fileTableBody.appendChild(tr);
    return;
  }

  const subjectOptions = getSubjectOptions();

  for (const file of state.files) {
    const tr = document.createElement('tr');

    const checkTd = document.createElement('td');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = file.enabled;
    checkbox.dataset.action = 'toggle';
    checkbox.dataset.id = file.sourcePath;
    checkTd.appendChild(checkbox);

    const sourceTd = document.createElement('td');
    sourceTd.className = 'source';
    sourceTd.textContent = file.sourceName;

    const subjectTd = document.createElement('td');
    subjectTd.appendChild(
      makeSelect({
        action: 'subject',
        id: file.sourcePath,
        value: file.subject,
        options: subjectOptions,
        placeholder: '과목 선택',
      }),
    );

    const orderTd = document.createElement('td');
    orderTd.appendChild(
      makeSelect({
        action: 'order',
        id: file.sourcePath,
        value: file.order,
        options: getOrderOptions(file.subject),
        placeholder: '순서 선택',
      }),
    );

    const dateTd = document.createElement('td');
    const dateInput = document.createElement('input');
    dateInput.className = 'name-input';
    dateInput.type = 'date';
    dateInput.value = file.classDate ?? '';
    dateInput.dataset.action = 'date';
    dateInput.dataset.id = file.sourcePath;
    dateTd.appendChild(dateInput);

    const periodTd = document.createElement('td');
    periodTd.appendChild(
      makeSelect({
        action: 'period',
        id: file.sourcePath,
        value: file.periodStart,
        options: PERIOD_OPTIONS,
        placeholder: '시작 교시',
      }),
    );

    const hoursTd = document.createElement('td');
    const hoursSelect = makeSelect({
      action: 'hours',
      id: file.sourcePath,
      value: state.mode === 'note' ? '1' : file.hours,
      options: state.mode === 'note' ? ['1'] : HOURS_OPTIONS,
      placeholder: '시수',
    });
    if (state.mode === 'note') {
      hoursSelect.disabled = true;
    }
    hoursTd.appendChild(hoursSelect);

    const professorTd = document.createElement('td');
    const professorInput = document.createElement('input');
    professorInput.className = 'name-input';
    professorInput.type = 'text';
    professorInput.value = file.professor ?? '';
    professorInput.placeholder = '예: 차동훈';
    professorInput.dataset.action = 'professor';
    professorInput.dataset.id = file.sourcePath;
    professorTd.appendChild(professorInput);

    const titleTd = document.createElement('td');
    titleTd.appendChild(
      makeSelect({
        action: 'title',
        id: file.sourcePath,
        value: file.title,
        options: getTitleOptions(file.subject, file.order, file.title),
        placeholder: '강의 제목 선택',
      }),
    );

    tr.appendChild(checkTd);
    tr.appendChild(sourceTd);
    tr.appendChild(subjectTd);
    tr.appendChild(orderTd);
    tr.appendChild(dateTd);
    tr.appendChild(periodTd);
    tr.appendChild(hoursTd);
    tr.appendChild(professorTd);
    tr.appendChild(titleTd);
    els.fileTableBody.appendChild(tr);

    const targetRow = document.createElement('tr');
    targetRow.className = 'target-review-row';
    const targetTd = document.createElement('td');
    targetTd.colSpan = mainColumnCount;
    const targetWrap = document.createElement('div');
    targetWrap.className = 'target-review-wrap';
    const targetLabel = document.createElement('span');
    targetLabel.className = 'target-review-label';
    targetLabel.textContent = '최종 파일명 검토';
    const targetInput = document.createElement('input');
    targetInput.className = 'name-input target-review-input';
    targetInput.type = 'text';
    targetInput.value = file.targetName ?? '';
    targetInput.dataset.action = 'target';
    targetInput.dataset.id = file.sourcePath;
    targetWrap.appendChild(targetLabel);
    targetWrap.appendChild(targetInput);
    targetTd.appendChild(targetWrap);
    targetRow.appendChild(targetTd);
    els.fileTableBody.appendChild(targetRow);
  }
}

function renderAdminTable() {
  els.adminTableBody.innerHTML = '';

  if (state.adminRows.length === 0) {
    state.adminRows.push(makeEmptyAdminRow());
  }

  for (let index = 0; index < state.adminRows.length; index += 1) {
    const row = state.adminRows[index];
    const tr = document.createElement('tr');

    const removeTd = document.createElement('td');
    const removeInput = document.createElement('input');
    removeInput.type = 'checkbox';
    removeInput.checked = Boolean(row.remove);
    removeInput.dataset.row = String(index);
    removeInput.dataset.action = 'remove';
    removeTd.appendChild(removeInput);

    const subjectTd = document.createElement('td');
    subjectTd.appendChild(makeAdminInput(index, 'subject', row.subject, 'text', '순환기학'));

    const prefixTd = document.createElement('td');
    prefixTd.appendChild(makeAdminInput(index, 'prefix', row.prefix, 'text', '순환'));

    const orderTd = document.createElement('td');
    orderTd.appendChild(makeAdminInput(index, 'order', row.orderNumber, 'text', '01'));

    const dateTd = document.createElement('td');
    dateTd.appendChild(makeAdminInput(index, 'date', row.classDate, 'date'));

    const periodTd = document.createElement('td');
    periodTd.appendChild(
      makeSelect({
        action: 'period',
        row: index,
        value: row.period,
        options: PERIOD_OPTIONS,
        placeholder: '시작교시',
        isAdmin: true,
      }),
    );

    const hoursTd = document.createElement('td');
    hoursTd.appendChild(
      makeSelect({
        action: 'hours',
        row: index,
        value: row.hours,
        options: HOURS_OPTIONS,
        placeholder: '시수',
        isAdmin: true,
      }),
    );

    const professorTd = document.createElement('td');
    professorTd.appendChild(makeAdminInput(index, 'professor', row.professor, 'text', '차동훈'));

    const titleTd = document.createElement('td');
    titleTd.appendChild(makeAdminInput(index, 'title', row.lectureTitle, 'text', '서론/신체 검사'));

    tr.appendChild(removeTd);
    tr.appendChild(subjectTd);
    tr.appendChild(prefixTd);
    tr.appendChild(orderTd);
    tr.appendChild(dateTd);
    tr.appendChild(periodTd);
    tr.appendChild(hoursTd);
    tr.appendChild(professorTd);
    tr.appendChild(titleTd);

    els.adminTableBody.appendChild(tr);
  }
}

function makeAdminInput(row, action, value, type, placeholder = '') {
  const input = document.createElement('input');
  input.className = 'name-input';
  input.type = type;
  input.value = value ?? '';
  input.placeholder = placeholder;
  input.dataset.row = String(row);
  input.dataset.action = action;
  return input;
}

function makeSelect({ action, id, row, value, options, placeholder, isAdmin = false }) {
  const select = document.createElement('select');
  select.className = 'name-input';
  select.dataset.action = action;

  if (isAdmin) {
    select.dataset.row = String(row);
  } else {
    select.dataset.id = id;
  }

  const empty = document.createElement('option');
  empty.value = '';
  empty.textContent = placeholder;
  select.appendChild(empty);

  for (const optionValue of options) {
    const option = document.createElement('option');
    option.value = optionValue;
    option.textContent = optionValue;
    select.appendChild(option);
  }

  if (value && !options.includes(value)) {
    const custom = document.createElement('option');
    custom.value = value;
    custom.textContent = value;
    select.appendChild(custom);
  }

  select.value = value ?? '';
  return select;
}

function buildTargetName(file) {
  if (!file.ext) {
    return file.targetName ?? file.sourceName;
  }

  const isNoteMode = state.mode === 'note';
  const subject = String(file.subject ?? '').trim();
  const order = normalizeOrder(file.order);
  const dateToken = formatDateToken(file.classDate);
  const effectiveHours = isNoteMode ? '1' : file.hours;
  const period = buildPeriodToken(file.periodStart, effectiveHours);
  const professorRaw = String(file.professor ?? '').trim();
  const title = String(file.title ?? '').trim();
  const author = String(state.noteAuthor ?? '').trim();

  if (!subject || !order || !dateToken || !period || !professorRaw || !title) {
    return file.targetName || file.sourceName;
  }

  if (isNoteMode && !author) {
    return file.targetName || file.sourceName;
  }

  const professor = normalizeProfessor(professorRaw);
  if (isNoteMode) {
    const rawName = `${subject}${order}-${dateToken}-${period}교시-${professor}-${title}-${author}${file.ext}`;
    return sanitizeOutputName(rawName);
  }

  const rawName = `${subject}${order}-${dateToken}-${period}교시-${professor}-${title}${file.ext}`;
  return sanitizeOutputName(rawName);
}

function updateAutoTargetName(file) {
  if (file.targetEdited) {
    return;
  }
  file.targetName = buildTargetName(file);
}

function normalizeProfessor(value) {
  const text = String(value ?? '').trim();
  if (!text) {
    return '';
  }
  if (text.endsWith('교수님')) {
    return text;
  }
  if (text.endsWith('교수')) {
    return `${text}님`;
  }
  return `${text}교수님`;
}

function formatDateToken(yyyyMmDd) {
  const text = String(yyyyMmDd ?? '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return '';
  }

  const [yyyy, mm, dd] = text.split('-');
  return `${yyyy.slice(-2)}${mm}${dd}`;
}

function toDateInputValue(entry, fallbackYear) {
  if (!entry) {
    return '';
  }

  if (entry.dateToken && /^\d{6}$/.test(entry.dateToken)) {
    const yy = Number(entry.dateToken.slice(0, 2));
    const mm = entry.dateToken.slice(2, 4);
    const dd = entry.dateToken.slice(4, 6);

    const inferredCentury = (fallbackYear ?? new Date().getFullYear()) >= 2000 ? 2000 : 1900;
    return `${inferredCentury + yy}-${mm}-${dd}`;
  }

  const year = Number(fallbackYear) || new Date().getFullYear();
  const month = String(entry.month ?? '').padStart(2, '0');
  const day = String(entry.day ?? '').padStart(2, '0');

  if (!month || !day) {
    return '';
  }

  return `${year}-${month}-${day}`;
}

function buildPeriodToken(startPeriod, hours) {
  const start = Number(startPeriod);
  const duration = Number(hours);

  if (!Number.isFinite(start) || start < 1 || start > 9) {
    return '';
  }

  const count = Number.isFinite(duration) && duration > 0 ? Math.round(duration) : 1;
  const periods = [];

  for (let i = 0; i < count; i += 1) {
    const value = start + i;
    if (value > 9) {
      break;
    }
    periods.push(String(value));
  }

  return periods.join(',');
}

function normalizeOrder(value) {
  const text = String(value ?? '').trim();
  const num = Number(text);
  if (!Number.isFinite(num) || num <= 0) {
    return '';
  }
  return String(Math.round(num)).padStart(2, '0');
}

function findEntry(subject, order) {
  const orderNumber = normalizeOrder(order);
  if (!subject || !orderNumber) {
    return null;
  }

  return (state.timetable?.entries ?? []).find(
    (entry) => entry.subject === subject && entry.orderNumber === orderNumber,
  );
}

function getSubjectOptions() {
  return unique((state.timetable?.entries ?? []).map((entry) => entry.subject)).sort((a, b) =>
    a.localeCompare(b, 'ko-KR'),
  );
}

function getOrderOptions(subject) {
  if (!subject) {
    return [];
  }

  return unique(
    (state.timetable?.entries ?? [])
      .filter((entry) => entry.subject === subject)
      .map((entry) => entry.orderNumber),
  ).sort((a, b) => Number(a) - Number(b));
}

function getTitleOptions(subject, order, currentValue) {
  if (!subject) {
    return currentValue ? [currentValue] : [];
  }

  const base = (state.timetable?.entries ?? []).filter((entry) => entry.subject === subject);
  const normalizedOrder = normalizeOrder(order);
  const filtered = normalizedOrder ? base.filter((entry) => entry.orderNumber === normalizedOrder) : base;
  const options = unique(filtered.map((entry) => entry.lectureTitle));

  if (currentValue && !options.includes(currentValue)) {
    options.unshift(currentValue);
  }

  return options;
}

function makeEmptyAdminRow() {
  return {
    remove: false,
    subject: '',
    prefix: '',
    orderNumber: '',
    classDate: '',
    period: '',
    hours: '1',
    professor: '',
    lectureTitle: '',
  };
}

function isAdminRowEmpty(row) {
  return !String(row.subject ?? '').trim() && !String(row.lectureTitle ?? '').trim() && !String(row.professor ?? '').trim();
}

function normalizeAdminRow(row, year) {
  const subject = String(row.subject ?? '').trim().replace(/\s+/g, '');
  const prefix = String(row.prefix ?? '').trim().replace(/\s+/g, '') || subject;
  const orderNumber = normalizeOrder(row.orderNumber);
  const classDate = String(row.classDate ?? '').trim();
  const period = String(row.period ?? '').trim();
  const hours = String(row.hours ?? '1').trim();
  const professor = String(row.professor ?? '').trim();
  const lectureTitle = String(row.lectureTitle ?? '').trim();

  if (!subject) return { ok: false, reason: '관리자 항목의 과목명이 비어 있습니다.' };
  if (!prefix) return { ok: false, reason: `관리자 항목(${subject})의 코드 접두가 비어 있습니다.` };
  if (!orderNumber) return { ok: false, reason: `관리자 항목(${subject})의 순서가 잘못되었습니다.` };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(classDate)) return { ok: false, reason: `관리자 항목(${subject}${orderNumber})의 수업일 형식이 잘못되었습니다.` };
  if (!period) return { ok: false, reason: `관리자 항목(${subject}${orderNumber})의 시작 교시가 비어 있습니다.` };
  if (!hours) return { ok: false, reason: `관리자 항목(${subject}${orderNumber})의 시수가 비어 있습니다.` };
  if (!professor) return { ok: false, reason: `관리자 항목(${subject}${orderNumber})의 교수님 성함이 비어 있습니다.` };
  if (!lectureTitle) return { ok: false, reason: `관리자 항목(${subject}${orderNumber})의 강의 제목이 비어 있습니다.` };

  const [yyyy, mm, dd] = classDate.split('-');
  const normalizedYear = Number(year) || Number(yyyy);
  const dateToken = `${String(normalizedYear).slice(-2)}${mm}${dd}`;

  return {
    ok: true,
    entry: {
      subject,
      prefix,
      orderNumber,
      lectureTitle,
      professor,
      month: Number(mm),
      day: Number(dd),
      dateToken,
      period,
      hours: Number(hours),
    },
  };
}

function writeLog(message) {
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  const stamp = `[${hh}:${mm}:${ss}]`;

  if (els.logOutput.textContent.length > 0) {
    els.logOutput.textContent += '\n';
  }

  els.logOutput.textContent += `${stamp} ${message}`;
  els.logOutput.scrollTop = els.logOutput.scrollHeight;
}

function unique(values) {
  return [...new Set(values)];
}

function loadLastSelectedSubject() {
  try {
    const value = String(localStorage.getItem(LAST_SUBJECT_STORAGE_KEY) ?? '').trim();
    state.lastSelectedSubject = value;
  } catch {
    state.lastSelectedSubject = '';
  }
}

function rememberLastSelectedSubject(subject) {
  const value = String(subject ?? '').trim();
  if (!value) {
    return;
  }

  state.lastSelectedSubject = value;
  try {
    localStorage.setItem(LAST_SUBJECT_STORAGE_KEY, value);
  } catch {
    // ignore persistence failures and continue in-memory
  }
}

function resolveDefaultSubject() {
  const value = String(state.lastSelectedSubject ?? '').trim();
  if (!value) {
    return '';
  }

  const subjects = getSubjectOptions();
  return subjects.includes(value) ? value : '';
}

function extractFilePaths(fileList) {
  return unique(
    Array.from(fileList ?? [])
      .map((file) => window.titleMaker.getPathForFile(file))
      .filter(Boolean),
  );
}

function getFileExt(filename) {
  const match = String(filename ?? '').match(/(\.[^.]+)$/);
  return match ? match[1] : '';
}

function sanitizeOutputName(filename) {
  return String(filename ?? '')
    .replace(/[\\/]/g, '-')
    .replace(/[<>:"|?*\x00-\x1F]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/-+/g, '-')
    .trim();
}
