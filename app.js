const baseFileInput = document.getElementById('baseFile');
const weeklyFilesInput = document.getElementById('weeklyFiles');
const mergeBtn = document.getElementById('mergeBtn');
const downloadBtn = document.getElementById('downloadBtn');
const logEl = document.getElementById('log');

let resultBlob = null;
let resultName = '';

function log(message) {
  logEl.textContent += `${message}\n`;
}

function clearLog() {
  logEl.textContent = '';
}

function normalizeSpaces(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

function weekSortKey(label) {
  const m = label.match(/(\d{2})\.(\d{2})-(\d{2})\.(\d{2})/);
  if (!m) return [999, 999, 999, 999];
  return [Number(m[2]), Number(m[1]), Number(m[4]), Number(m[3])];
}

function weekLabelFromFileName(fileName) {
  const m = fileName.match(/(\d{2}\.\d{2})\s*-\s*(\d{2}\.\d{2})/);
  if (m) return `${m[1]}-${m[2]}`;
  return fileName.replace(/\.xlsx$/i, '');
}

function extractDuration(groupName, minutes) {
  const m = normalizeSpaces(groupName).match(/\((\d+)\s*минут\)\s*$/i);
  if (m) return Number(m[1]);
  if (minutes === null || minutes === undefined || minutes === '') return null;
  const numeric = Number(minutes);
  return Number.isFinite(numeric) ? numeric : null;
}

function groupDisplayAndKey(groupName) {
  const base = normalizeSpaces(groupName).replace(/\s*\(\d+\s*минут\)\s*$/i, '');
  return { display: normalizeSpaces(base), key: normalizeSpaces(base).toLowerCase() };
}

function cleanTheme(text) {
  const original = normalizeSpaces(text);
  if (!original) return '';

  let cleaned = original;
  if (cleaned.includes('КОММЕНТ')) {
    cleaned = cleaned.split('КОММЕНТ')[0].trim();
  }
  cleaned = normalizeSpaces(cleaned.replaceAll('$', '').replaceAll('&', ' ').replaceAll("'", ''));

  if (cleaned.startsWith('✨ Регулярный репорт')) {
    const lesson = original.match(/(\d+)[-–]?(?:ый|ой|ий|ая)?\s+урок[^.]*успешно пройден/i);
    return lesson ? `Подг. курс — урок ${lesson[1]}` : 'Регулярный репорт';
  }

  return cleaned;
}

function parseWeeklyRecords(fileName, workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  let teacher = '';
  const week = weekLabelFromFileName(fileName);
  const records = [];

  for (const row of rows) {
    if (!row || row.length === 0) continue;
    const c0 = normalizeSpaces(row[0]);

    if (row.length === 1 && c0 && !c0.startsWith('За период') && c0 !== '№') {
      teacher = c0;
      continue;
    }

    if (!/^\d+$/.test(c0) || row.length < 8) continue;

    const dateStr = normalizeSpaces(row[1]);
    const groupRaw = normalizeSpaces(row[2]);
    const minutes = row[5];
    const themeRaw = normalizeSpaces(row[7]);
    if (!groupRaw) continue;

    const { display, key } = groupDisplayAndKey(groupRaw);
    records.push({
      week,
      dateStr,
      group: display,
      groupKey: key,
      duration: extractDuration(groupRaw, minutes),
      teacher,
      themeShort: cleanTheme(themeRaw),
      themeFull: themeRaw,
    });
  }

  return records;
}

function mapFromBaseWorkbook(workbook) {
  const map = new Map();
  if (!workbook) return map;

  const sheet = workbook.Sheets['Сырые данные'];
  if (!sheet) {
    log('⚠ В общей таблице нет листа "Сырые данные". Будет использована только новая неделя.');
    return map;
  }

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  for (let i = 1; i < rows.length; i += 1) {
    const r = rows[i];
    if (!r || r.length < 7) continue;

    const group = normalizeSpaces(r[2]);
    const groupKey = group.toLowerCase();
    if (!groupKey) continue;

    if (!map.has(groupKey)) {
      map.set(groupKey, {
        group,
        durations: [],
        teacherCount: new Map(),
        byWeek: new Map(),
      });
    }

    const item = map.get(groupKey);
    const durationNum = Number(r[3]);
    if (Number.isFinite(durationNum)) item.durations.push(durationNum);

    const teacher = normalizeSpaces(r[4]);
    if (teacher) item.teacherCount.set(teacher, (item.teacherCount.get(teacher) || 0) + 1);

    const week = normalizeSpaces(r[0]);
    if (week) {
      const line = `${normalizeSpaces(r[5])} (${normalizeSpaces(r[1])})`;
      const list = item.byWeek.get(week) || [];
      if (!list.includes(line)) list.push(line);
      item.byWeek.set(week, list);
    }
  }

  log(`Загружены существующие данные: ${map.size} групп.`);
  return map;
}

function addRecordsToMap(map, records) {
  for (const r of records) {
    if (!map.has(r.groupKey)) {
      map.set(r.groupKey, {
        group: r.group,
        durations: [],
        teacherCount: new Map(),
        byWeek: new Map(),
      });
    }

    const item = map.get(r.groupKey);
    if (r.duration !== null && r.duration !== undefined) item.durations.push(r.duration);
    if (r.teacher) item.teacherCount.set(r.teacher, (item.teacherCount.get(r.teacher) || 0) + 1);

    const line = `${r.themeShort} (${r.dateStr})`;
    const list = item.byWeek.get(r.week) || [];
    if (!list.includes(line)) list.push(line);
    item.byWeek.set(r.week, list);
  }
}

function mostCommon(numbers) {
  const count = new Map();
  for (const n of numbers) {
    count.set(n, (count.get(n) || 0) + 1);
  }
  let best = '';
  let bestCount = -1;
  for (const [value, c] of count.entries()) {
    if (c > bestCount) {
      bestCount = c;
      best = value;
    }
  }
  return best;
}

function aggregateToRows(groupMap) {
  const weekSet = new Set();
  for (const item of groupMap.values()) {
    for (const week of item.byWeek.keys()) weekSet.add(week);
  }

  const weeks = [...weekSet].sort((a, b) => {
    const ka = weekSortKey(a);
    const kb = weekSortKey(b);
    return ka.join(',').localeCompare(kb.join(','), 'ru');
  });

  const trajectoryRows = [['Группа', 'Длительность, мин', 'Преподаватель', ...weeks]];
  const rawRows = [['Неделя', 'Дата', 'Группа', 'Длительность, мин', 'Преподаватель', 'Тема коротко', 'Тема полностью']];

  const groups = [...groupMap.values()].sort((a, b) => a.group.localeCompare(b.group, 'ru'));
  for (const item of groups) {
    const duration = mostCommon(item.durations);
    const teachers = [...item.teacherCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([name]) => name)
      .join(' / ');

    const row = [item.group, duration, teachers];
    for (const week of weeks) {
      row.push((item.byWeek.get(week) || []).join('\n'));
    }
    trajectoryRows.push(row);

    for (const week of weeks) {
      for (const line of item.byWeek.get(week) || []) {
        const m = line.match(/^(.*)\s\((\d{2}\.\d{2}\.\d{4})\)$/);
        rawRows.push([
          week,
          m ? m[2] : '',
          item.group,
          duration,
          teachers,
          m ? m[1] : line,
          m ? m[1] : line,
        ]);
      }
    }
  }

  return { weeks, trajectoryRows, rawRows };
}

function buildWorkbook(trajectoryRows, rawRows) {
  const wb = XLSX.utils.book_new();
  const trajectory = XLSX.utils.aoa_to_sheet(trajectoryRows);
  const raw = XLSX.utils.aoa_to_sheet(rawRows);
  const instructions = XLSX.utils.aoa_to_sheet([
    ['Как обновлять файл'],
    ['1. Загрузите существующую общую таблицу (если есть).'],
    ['2. Добавьте новые недельные выгрузки.'],
    ['3. Нажмите «Собрать обновлённую таблицу».'],
    ['4. Скачайте новый файл и используйте его как общую таблицу в следующий раз.'],
  ]);

  XLSX.utils.book_append_sheet(wb, trajectory, 'Траектория');
  XLSX.utils.book_append_sheet(wb, raw, 'Сырые данные');
  XLSX.utils.book_append_sheet(wb, instructions, 'Инструкция');
  return wb;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

mergeBtn.addEventListener('click', async () => {
  try {
    clearLog();
    downloadBtn.disabled = true;
    resultBlob = null;

    const weeklyFiles = [...weeklyFilesInput.files];
    if (weeklyFiles.length === 0) {
      log('Ошибка: выберите хотя бы одну недельную выгрузку.');
      return;
    }

    const baseFile = baseFileInput.files[0] || null;
    const groupMap = new Map();

    if (baseFile) {
      const baseBuffer = await readFileAsArrayBuffer(baseFile);
      const baseWb = XLSX.read(baseBuffer, { type: 'array' });
      const baseMap = mapFromBaseWorkbook(baseWb);
      for (const [k, v] of baseMap.entries()) groupMap.set(k, v);
    }

    let totalRecords = 0;
    for (const file of weeklyFiles) {
      const buffer = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(buffer, { type: 'array' });
      const records = parseWeeklyRecords(file.name, wb);
      addRecordsToMap(groupMap, records);
      totalRecords += records.length;
      log(`✓ ${file.name}: найдено занятий ${records.length}`);
    }

    const { weeks, trajectoryRows, rawRows } = aggregateToRows(groupMap);
    const resultWb = buildWorkbook(trajectoryRows, rawRows);
    const array = XLSX.write(resultWb, { type: 'array', bookType: 'xlsx' });

    resultBlob = new Blob([array], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const now = new Date();
    const dateTag = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
      now.getDate(),
    ).padStart(2, '0')}`;
    resultName = `Траектория_групп_${dateTag}.xlsx`;

    log('---');
    log(`Всего недель: ${weeks.length}`);
    log(`Всего групп: ${trajectoryRows.length - 1}`);
    log(`Всего новых занятий: ${totalRecords}`);
    log(`Файл готов: ${resultName}`);

    downloadBtn.disabled = false;
  } catch (error) {
    log(`Ошибка: ${error.message || error}`);
  }
});

downloadBtn.addEventListener('click', () => {
  if (!resultBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(resultBlob);
  a.download = resultName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
});
