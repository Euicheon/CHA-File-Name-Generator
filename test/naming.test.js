const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const {
  buildPeriodToken,
  buildSuggestionForFile,
  loadTimetableFromFile,
  sanitizeFilename,
} = require('../src/core/naming');

function resolveTimetablePath() {
  const cwd = path.resolve(__dirname, '..');
  const candidates = fs
    .readdirSync(cwd)
    .filter((name) => name.toLowerCase().endsWith('.xlsx'))
    .sort((a, b) => a.localeCompare(b, 'ko-KR'));

  if (candidates.length === 0) {
    throw new Error('테스트용 .xlsx 파일을 찾을 수 없습니다.');
  }

  const prioritized = candidates.find((name) => /강의|시간표|계획/i.test(name));
  return path.join(cwd, prioritized ?? candidates[0]);
}

test('시간표 로드가 강의 엔트리를 생성한다', () => {
  const timetablePath = resolveTimetablePath();
  const timetable = loadTimetableFromFile(timetablePath);

  assert.ok(timetable.entryCount > 0);
  assert.equal(timetable.entries.length, timetable.entryCount);
  assert.ok(timetable.entries.some((entry) => entry.subject === '순환기학'));
});

test('순환-01 파일명이 자동 매칭된다', () => {
  const timetable = loadTimetableFromFile(resolveTimetablePath());
  const suggestion = buildSuggestionForFile('/tmp/순환-01-intro.pdf', timetable.entries);

  assert.equal(suggestion.matchStatus, 'matched');
  assert.ok(suggestion.suggestedName.startsWith('순환기학01-'));
  assert.ok(suggestion.suggestedName.endsWith('.pdf'));
  assert.ok(suggestion.matchSummary.includes('순환기학01'));
});

test('복수 순서(03,04)도 한 파일로 묶어 제안한다', () => {
  const timetable = loadTimetableFromFile(resolveTimetablePath());
  const suggestion = buildSuggestionForFile('/tmp/호흡03,04_강의자료.pdf', timetable.entries);

  assert.equal(suggestion.matchStatus, 'matched');
  assert.ok(suggestion.matchedEntryIds.length >= 2);
  assert.ok(suggestion.suggestedName.startsWith('호흡기학03,04-'));
});

test('파일명 sanitize가 윈도우 금지문자를 제거한다', () => {
  const value = sanitizeFilename('test:<>"|?*.pdf');
  assert.equal(value, 'test-.pdf');
});

test('시수 기준으로 교시 문자열을 확장한다', () => {
  assert.equal(buildPeriodToken('5', 3), '5,6,7');
  assert.equal(buildPeriodToken('9', 3), '9');
});
