const fsSync = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function findDefaultTimetablePath(rootDir) {
  if (!rootDir || !fsSync.existsSync(rootDir)) {
    return null;
  }

  const files = fsSync
    .readdirSync(rootDir, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith('.xlsx'))
    .map((entry) => entry.name)
    .sort((a, b) => a.localeCompare(b, 'ko-KR'));

  if (files.length === 0) {
    return null;
  }

  const priority = files.find((name) => /강의|시간표|계획/i.test(name));
  return path.join(rootDir, priority ?? files[0]);
}

function loadTimetableFromFile(filePath) {
  const fullPath = path.resolve(filePath);
  if (!isRegularFile(fullPath)) {
    throw new Error(`시간표 파일을 찾을 수 없습니다: ${fullPath}`);
  }

  const workbook = XLSX.readFile(fullPath, {
    cellDates: true,
    raw: true,
  });

  const targetSheetName = workbook.SheetNames.find((name) => name.trim() === '강의계획표') ?? workbook.SheetNames[0];
  if (!targetSheetName) {
    throw new Error('강의계획표 시트를 찾을 수 없습니다.');
  }

  const sheet = workbook.Sheets[targetSheetName];
  const decodedRange = XLSX.utils.decode_range(sheet['!ref']);
  const colOffset = decodedRange.s.c;
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });

  const year = inferYear(fullPath, workbook);
  const entries = parseLectureEntries(rows, year, colOffset);

  if (entries.length === 0) {
    throw new Error('강의계획표에서 강의 행을 찾지 못했습니다. 파일 형식을 확인해 주세요.');
  }

  return buildTimetableFromEntries(entries, year, fullPath);
}

function inferYear(filePath, workbook) {
  const filename = path.basename(filePath);
  const fileMatch = filename.match(/(20\d{2})/);
  if (fileMatch) {
    return Number(fileMatch[1]);
  }

  const timeTableSheet = workbook.Sheets['시간표'];
  if (timeTableSheet) {
    const rows = XLSX.utils.sheet_to_json(timeTableSheet, {
      header: 1,
      defval: null,
      raw: true,
    });

    for (const row of rows) {
      for (const cell of row) {
        if (cell instanceof Date) {
          return cell.getFullYear();
        }
      }
    }
  }

  return new Date().getFullYear();
}

function parseLectureEntries(rows, year, colOffset) {
  const entries = [];
  let currentSubject = '';

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] ?? [];
    const colB = normalizeText(getCellByColNumber(row, 2, colOffset));
    const colC = normalizeText(getCellByColNumber(row, 3, colOffset));

    if (
      colB &&
      !colC &&
      !/순서/.test(colB) &&
      !/[-–]\s*\d+$/.test(colB) &&
      !/^과목책임교수/.test(colB)
    ) {
      currentSubject = compactSubject(colB);
    }

    const orderMatch = colB.match(/^(.+?)\s*-\s*(\d{1,2})$/);
    if (!orderMatch) {
      continue;
    }

    const prefix = normalizePrefix(orderMatch[1]);
    const orderNumber = String(Number(orderMatch[2])).padStart(2, '0');
    const lectureTitle = normalizeText(getCellByColNumber(row, 3, colOffset));
    const hours = normalizeHours(getCellByColNumber(row, 4, colOffset));
    const professor = normalizeText(getCellByColNumber(row, 8, colOffset));
    const month = Number(getCellByColNumber(row, 9, colOffset));
    const day = Number(getCellByColNumber(row, 10, colOffset));
    const period = normalizePeriod(getCellByColNumber(row, 12, colOffset));

    if (
      !prefix ||
      !currentSubject ||
      !professor ||
      !lectureTitle ||
      !Number.isFinite(month) ||
      !Number.isFinite(day) ||
      !period
    ) {
      continue;
    }

    const mm = String(month).padStart(2, '0');
    const dd = String(day).padStart(2, '0');
    const yy = String(year).slice(-2);
    const dateToken = `${yy}${mm}${dd}`;

    entries.push({
      id: `${prefix}-${orderNumber}`,
      subject: currentSubject,
      prefix,
      orderNumber,
      lectureTitle,
      professor,
      month,
      day,
      dateToken,
      period,
      hours,
      summary: `${currentSubject}${orderNumber} | ${dateToken} | ${buildPeriodToken(period, hours)}교시 | ${professor} | ${lectureTitle}`,
    });
  }

  return entries;
}

function buildTimetableFromEntries(entries, year, filePath) {
  const normalized = entries
    .map((entry) => normalizeEntry(entry, year))
    .filter(Boolean)
    .sort(sortLectureEntries);

  return {
    path: path.resolve(filePath),
    year: Number(year) || new Date().getFullYear(),
    entryCount: normalized.length,
    entries: normalized,
  };
}

function getCellByColNumber(row, colNumber, colOffset) {
  const index = colNumber - 1 - colOffset;
  if (index < 0) {
    return null;
  }
  return row[index] ?? null;
}

function buildSuggestionForFile(filePath, entries) {
  const ext = path.extname(filePath);
  const basename = path.basename(filePath, ext);
  const matchedEntries = findMatchedEntries(basename, entries);

  if (matchedEntries.length === 0) {
    return {
      sourcePath: filePath,
      sourceName: path.basename(filePath),
      suggestedName: sanitizeFilename(`${basename}${ext}`),
      matchStatus: 'unmatched',
      matchSummary: '자동 매칭 실패 (필요 시 새 파일명 직접 수정)',
      matchedEntryIds: [],
    };
  }

  const suggestedName = buildFilenameFromEntries(matchedEntries, ext);
  const matchLabels = matchedEntries.map((entry) => `${entry.subject}${entry.orderNumber}`);

  return {
    sourcePath: filePath,
    sourceName: path.basename(filePath),
    suggestedName,
    matchStatus: 'matched',
    matchSummary: `자동 매칭: ${matchLabels.join(', ')}`,
    matchedEntryIds: matchedEntries.map((entry) => entry.id),
  };
}

function findMatchedEntries(baseName, entries) {
  const byAlias = matchByAliasAndOrder(baseName, entries);
  if (byAlias.length > 0) {
    return byAlias;
  }

  const byToken = matchByToken(baseName, entries);
  if (byToken.length > 0) {
    return byToken;
  }

  return [];
}

function matchByAliasAndOrder(baseName, entries) {
  const normalizedBase = normalizeText(baseName);
  const subjects = [...new Set(entries.map((entry) => entry.subject))].sort((a, b) => b.length - a.length);
  const prefixes = [...new Set(entries.map((entry) => entry.prefix))].sort((a, b) => b.length - a.length);

  const subjectMatches = findAliasMatches(normalizedBase, subjects, entries, 'subject');
  if (subjectMatches.length > 0) {
    return subjectMatches;
  }

  return findAliasMatches(normalizedBase, prefixes, entries, 'prefix');
}

function findAliasMatches(baseName, aliases, entries, mode) {
  for (const alias of aliases) {
    const regex = new RegExp(`${escapeRegExp(alias)}\\s*[-_]?\\s*(\\d{1,2}(?:\\s*,\\s*\\d{1,2})*)`, 'g');
    const wanted = new Set();

    for (const match of baseName.matchAll(regex)) {
      const rawNumbers = match[1].split(',').map((token) => token.trim()).filter(Boolean);
      for (const numberToken of rawNumbers) {
        wanted.add(String(Number(numberToken)).padStart(2, '0'));
      }
    }

    if (wanted.size === 0) {
      continue;
    }

    const matched = entries
      .filter((entry) => (mode === 'subject' ? entry.subject === alias : entry.prefix === alias))
      .filter((entry) => wanted.has(entry.orderNumber))
      .sort(sortLectureEntries);

    if (matched.length > 0) {
      return matched;
    }
  }

  return [];
}

function matchByToken(baseName, entries) {
  const normalizedBase = normalizeLoose(baseName);

  const matched = entries.filter((entry) => {
    const orderNoLeadingZero = String(Number(entry.orderNumber));
    const tokens = [
      `${entry.subject}${entry.orderNumber}`,
      `${entry.subject}${orderNoLeadingZero}`,
      `${entry.prefix}${entry.orderNumber}`,
      `${entry.prefix}${orderNoLeadingZero}`,
      `${entry.prefix}-${entry.orderNumber}`,
      `${entry.subject}-${entry.orderNumber}`,
    ];

    return tokens
      .map((token) => normalizeLoose(token))
      .some((token) => token.length > 0 && normalizedBase.includes(token));
  });

  return dedupeEntries(matched).sort(sortLectureEntries);
}

function buildFilenameFromEntries(entries, ext) {
  const sorted = [...entries].sort(sortLectureEntries);
  const subject = sorted[0]?.subject ?? '강의';
  const orderPart = sorted.map((entry) => entry.orderNumber).join(',');
  const datePart = uniqueValues(sorted.map((entry) => entry.dateToken)).join(',');
  const periodPart = uniqueValues(sorted.map((entry) => buildPeriodToken(entry.period, entry.hours))).join(',');
  const professorPart = uniqueValues(sorted.map((entry) => `${entry.professor}교수님`)).join(',');
  const titlePart = uniqueValues(sorted.map((entry) => entry.lectureTitle)).join('-');

  const rawName = `${subject}${orderPart}-${datePart}-${periodPart}교시-${professorPart}-${titlePart}${ext}`;
  return sanitizeFilename(rawName);
}

function ensureExt(name, fallbackExt) {
  const trimmed = String(name ?? '').trim();
  if (!trimmed) {
    return `untitled${fallbackExt || ''}`;
  }

  if (path.extname(trimmed)) {
    return trimmed;
  }

  return `${trimmed}${fallbackExt || ''}`;
}

function sanitizeFilename(filename) {
  const raw = String(filename ?? '').trim().replace(/[\\/]/g, '-');
  const extMatch = raw.match(/(\.[^.]+)$/);
  const ext = extMatch ? extMatch[1] : '';
  const base = ext ? raw.slice(0, -ext.length) : raw;

  const cleanBase = base
    .replace(/[<>:"|?*\x00-\x1F]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/-+/g, '-')
    .trim()
    .replace(/[\.\s]+$/g, '');

  const safeBase = cleanBase || 'untitled';
  return `${safeBase}${ext}`;
}

function isRegularFile(filePath) {
  try {
    return fsSync.statSync(filePath).isFile();
  } catch {
    return false;
  }
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function compactSubject(value) {
  return normalizeText(value).replace(/\s+/g, '');
}

function normalizePrefix(value) {
  return normalizeText(value).replace(/\s+/g, '');
}

function normalizePeriod(value) {
  const text = normalizeText(value).replace(/교시/g, '');
  if (!text) {
    return '';
  }

  if (!Number.isNaN(Number(text))) {
    return String(Number(text));
  }

  return text.replace(/\s+/g, '');
}

function normalizeHours(value) {
  const num = Number(value);
  if (!Number.isFinite(num) || num <= 0) {
    return 1;
  }
  return Math.max(1, Math.min(9, Math.round(num)));
}

function buildPeriodToken(startPeriod, hours) {
  const start = Number(startPeriod);
  const duration = normalizeHours(hours);
  if (!Number.isFinite(start) || start <= 0) {
    return normalizePeriod(startPeriod);
  }

  const periods = [];
  for (let i = 0; i < duration; i += 1) {
    const period = start + i;
    if (period > 9) {
      break;
    }
    periods.push(String(period));
  }

  if (periods.length === 0) {
    return String(start);
  }

  return periods.join(',');
}

function normalizeEntry(entry, year) {
  const subject = compactSubject(entry?.subject);
  const prefix = normalizePrefix(entry?.prefix || subject);
  const orderNumber = String(Number(entry?.orderNumber)).padStart(2, '0');
  const lectureTitle = normalizeText(entry?.lectureTitle);
  const professor = normalizeText(entry?.professor);
  const period = normalizePeriod(entry?.period);
  const hours = normalizeHours(entry?.hours);

  const month = Number(entry?.month);
  const day = Number(entry?.day);
  const dateTokenFromEntry = normalizeText(entry?.dateToken);

  let dateToken = '';
  let normalizedMonth = month;
  let normalizedDay = day;

  if (/^\d{6}$/.test(dateTokenFromEntry)) {
    dateToken = dateTokenFromEntry;
    normalizedMonth = Number(dateToken.slice(2, 4));
    normalizedDay = Number(dateToken.slice(4, 6));
  } else if (Number.isFinite(month) && Number.isFinite(day)) {
    const yy = String(Number(year) || new Date().getFullYear()).slice(-2);
    dateToken = `${yy}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
  }

  if (!subject || !prefix || !/^\d{2}$/.test(orderNumber) || !lectureTitle || !professor || !period || !/^\d{6}$/.test(dateToken)) {
    return null;
  }

  return {
    id: `${prefix}-${orderNumber}`,
    subject,
    prefix,
    orderNumber,
    lectureTitle,
    professor,
    month: normalizedMonth,
    day: normalizedDay,
    dateToken,
    period,
    hours,
    summary: `${subject}${orderNumber} | ${dateToken} | ${buildPeriodToken(period, hours)}교시 | ${professor} | ${lectureTitle}`,
  };
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeLoose(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[\s_\-]/g, '');
}

function sortLectureEntries(a, b) {
  const keyA = `${a.dateToken}-${a.orderNumber}`;
  const keyB = `${b.dateToken}-${b.orderNumber}`;
  return keyA.localeCompare(keyB, 'ko-KR', { numeric: true });
}

function dedupeEntries(entries) {
  const map = new Map();
  for (const entry of entries) {
    map.set(entry.id, entry);
  }
  return [...map.values()];
}

function uniqueValues(values) {
  return [...new Set(values.filter(Boolean))];
}

module.exports = {
  buildPeriodToken,
  buildTimetableFromEntries,
  buildFilenameFromEntries,
  buildSuggestionForFile,
  ensureExt,
  findDefaultTimetablePath,
  isRegularFile,
  loadTimetableFromFile,
  sanitizeFilename,
};
