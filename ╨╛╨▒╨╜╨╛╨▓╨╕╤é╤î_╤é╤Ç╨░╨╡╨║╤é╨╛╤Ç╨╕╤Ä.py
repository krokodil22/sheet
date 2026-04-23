import os, re, zipfile, xml.etree.ElementTree as ET
from collections import defaultdict, Counter
from datetime import datetime, timezone
from xml.sax.saxutils import escape

OUTPUT_FILE = 'Траектория_групп_обновляемая.xlsx'
THIS_SCRIPT = os.path.basename(__file__)


def normalize_spaces(s):
    return re.sub(r'\s+', ' ', (s or '')).strip()


def week_sort_key(label):
    m = re.match(r'(\d{2})\.(\d{2})-(\d{2})\.(\d{2})', label)
    if m:
        return (int(m.group(2)), int(m.group(1)), int(m.group(4)), int(m.group(3)))
    return (999, 999, 999, 999)


def col_letter(n):
    s = ''
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def xml_text(value):
    value = '' if value is None else str(value)
    return escape(value).replace('\n', '&#10;')


def col_letter_to_index(col_letters):
    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch.upper()) - 64)
    return idx - 1


def parse_xlsx_rows(path):
    ns = {
        'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('a:si', ns):
                texts = []
                for t in si.iterfind('.//a:t', ns):
                    texts.append(t.text or '')
                shared.append(''.join(texts))
        workbook_root = ET.fromstring(z.read('xl/workbook.xml'))
        first_sheet = workbook_root.find('a:sheets', ns)[0]
        rel_id = first_sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target = None
        for rel in rels:
            if rel.attrib.get('Id') == rel_id:
                target = rel.attrib['Target']
                break
        if not target:
            raise RuntimeError(f'Не удалось найти лист в {path}')
        if not target.startswith('xl/'):
            target = 'xl/' + target
        root = ET.fromstring(z.read(target))
        sheet_data = root.find('a:sheetData', ns)
        rows = []
        for row in sheet_data.findall('a:row', ns):
            vals = {}
            maxc = -1
            for c in row.findall('a:c', ns):
                ref = c.attrib.get('r', '')
                m = re.match(r'([A-Z]+)(\d+)', ref)
                if not m:
                    continue
                ci = col_letter_to_index(m.group(1))
                maxc = max(maxc, ci)
                t = c.attrib.get('t')
                v = c.find('a:v', ns)
                is_el = c.find('a:is', ns)
                value = None
                if t == 's':
                    value = shared[int(v.text)] if v is not None and v.text is not None else ''
                elif t == 'inlineStr':
                    texts = []
                    if is_el is not None:
                        for te in is_el.iterfind('.//a:t', ns):
                            texts.append(te.text or '')
                    value = ''.join(texts)
                else:
                    value = v.text if v is not None else None
                vals[ci] = value
            if maxc >= 0:
                rows.append([vals.get(i) for i in range(maxc + 1)])
        return rows


def week_label_from_filename(path):
    name = os.path.basename(path)
    m = re.search(r'(\d{2}\.\d{2})\s*-\s*(\d{2}\.\d{2})', name)
    if m:
        return f'{m.group(1)}-{m.group(2)}'
    return os.path.splitext(name)[0]


def extract_duration(group_name, minutes):
    group_name = normalize_spaces(group_name)
    m = re.search(r'\((\d+)\s*минут\)\s*$', group_name, re.I)
    if m:
        return int(m.group(1))
    try:
        if minutes not in (None, ''):
            return int(float(minutes))
    except Exception:
        pass
    return None


def group_display_and_key(group_name):
    group_name = normalize_spaces(group_name)
    base = re.sub(r'\s*\(\d+\s*минут\)\s*$', '', group_name, flags=re.I)
    base = normalize_spaces(base)
    key = base.lower()
    return base, key


def clean_theme(text):
    original = normalize_spaces(text)
    if not original:
        return ''
    cleaned = original
    if 'КОММЕНТ' in cleaned:
        cleaned = cleaned.split('КОММЕНТ')[0].strip(" '\"")
    cleaned = cleaned.replace('$', '').replace('&', ' ').replace("'", '').strip()
    cleaned = normalize_spaces(cleaned)
    if cleaned.startswith('✨ Регулярный репорт'):
        lesson = re.search(r'(\d+)[-–]?(?:ый|ой|ий|ая)?\s+урок[^.]*успешно пройден', original, re.I)
        if lesson:
            return f'Подг. курс — урок {lesson.group(1)}'
        return 'Регулярный репорт'
    return cleaned


def parse_records(path):
    rows = parse_xlsx_rows(path)
    week = week_label_from_filename(path)
    teacher = None
    records = []
    for r in rows:
        if not r:
            continue
        c0 = normalize_spaces(r[0] if len(r) > 0 and r[0] is not None else '')
        if len(r) == 1 and c0 and not c0.startswith('За период') and c0 != '№':
            teacher = c0
            continue
        if c0.isdigit() and len(r) >= 8:
            date_str = normalize_spaces(r[1] or '')
            group_raw = normalize_spaces(r[2] or '')
            minutes = r[5] if len(r) > 5 else None
            theme_raw = r[7] if len(r) > 7 else ''
            if not group_raw:
                continue
            try:
                dt = datetime.strptime(date_str, '%d.%m.%Y').date() if date_str else None
            except Exception:
                dt = None
            group_display, group_key = group_display_and_key(group_raw)
            records.append({
                'week': week,
                'date': dt,
                'date_str': date_str,
                'group': group_display,
                'group_key': group_key,
                'duration': extract_duration(group_raw, minutes),
                'teacher': teacher or '',
                'theme_short': clean_theme(theme_raw),
                'theme_full': normalize_spaces(theme_raw),
            })
    return records


def collect_source_files(folder):
    files = []
    for name in os.listdir(folder):
        if not name.lower().endswith('.xlsx'):
            continue
        if name.startswith('~$'):
            continue
        if name == OUTPUT_FILE:
            continue
        if name.startswith('Траектория_'):
            continue
        files.append(os.path.join(folder, name))
    return sorted(files)


def aggregate(records):
    weeks = sorted({r['week'] for r in records}, key=week_sort_key)
    grouped = defaultdict(list)
    for r in records:
        grouped[r['group_key']].append(r)
    trajectory = []
    raw = []
    for r in sorted(records, key=lambda x: (x['group_key'], x['date'] or datetime.min.date(), x['teacher'])):
        raw.append([
            r['week'], r['date_str'], r['group'], r['duration'], r['teacher'], r['theme_short'], r['theme_full']
        ])
    for group_key, items in grouped.items():
        items_sorted = sorted(items, key=lambda x: (x['date'] or datetime.min.date(), x['theme_short']))
        group_name = items_sorted[0]['group']
        duration_counts = Counter(i['duration'] for i in items if i['duration'] is not None)
        duration = duration_counts.most_common(1)[0][0] if duration_counts else ''
        teacher_counts = Counter(i['teacher'] for i in items if i['teacher'])
        teachers = ' / '.join([t for t, _ in teacher_counts.most_common(3)])
        by_week = defaultdict(list)
        for item in items_sorted:
            by_week[item['week']].append(f"{item['theme_short']} ({item['date_str']})")
        row = [group_name, duration, teachers]
        for week in weeks:
            row.append('\n'.join(by_week.get(week, [])))
        trajectory.append(row)
    trajectory.sort(key=lambda x: x[0].lower())
    return weeks, trajectory, raw


def make_cell(ref, value, style_id=0, numeric=False):
    if value is None or value == '':
        return f'<c r="{ref}" s="{style_id}"/>'
    if numeric:
        return f'<c r="{ref}" s="{style_id}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{style_id}" t="inlineStr"><is><t xml:space="preserve">{xml_text(value)}</t></is></c>'


def build_sheet_xml(rows, widths, freeze=None, autofilter_ref=None):
    lines = []
    for idx, width in enumerate(widths, start=1):
        lines.append(f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>')
    cols_xml = '<cols>' + ''.join(lines) + '</cols>' if lines else ''

    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f'{col_letter(c_idx)}{r_idx}'
            if r_idx == 1:
                style_id = 1
            else:
                style_id = 2
            numeric = isinstance(value, (int, float)) and c_idx == 2 and r_idx > 1
            cells.append(make_cell(ref, value, style_id=style_id, numeric=numeric))
        row_attrs = f' r="{r_idx}"'
        if r_idx == 1:
            row_attrs += ' ht="24" customHeight="1"'
        sheet_rows.append(f'<row{row_attrs}>' + ''.join(cells) + '</row>')
    sheet_data_xml = '<sheetData>' + ''.join(sheet_rows) + '</sheetData>'

    pane_xml = ''
    if freeze == 'trajectory':
        pane_xml = (
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/>'
            '<selection pane="topRight" activeCell="D1" sqref="D1"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '<selection pane="bottomRight" activeCell="D2" sqref="D2"/>'
            '</sheetView></sheetViews>'
        )
    elif freeze == 'raw':
        pane_xml = (
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '</sheetView></sheetViews>'
        )
    else:
        pane_xml = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'

    autofilter_xml = f'<autoFilter ref="{autofilter_ref}"/>' if autofilter_ref else ''
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'{pane_xml}'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'{cols_xml}'
        f'{sheet_data_xml}'
        f'{autofilter_xml}'
        '</worksheet>'
    )


def build_instruction_xml(lines):
    rows = [[line] for line in lines]
    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        style_id = 3 if r_idx == 1 else 4
        cells = [make_cell(f'A{r_idx}', row[0], style_id=style_id, numeric=False)]
        row_attrs = f' r="{r_idx}"'
        if r_idx == 1:
            row_attrs += ' ht="24" customHeight="1"'
        sheet_rows.append(f'<row{row_attrs}>' + ''.join(cells) + '</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="90" customWidth="1"/></cols>'
        '<sheetData>' + ''.join(sheet_rows) + '</sheetData>'
        '</worksheet>'
    )


def build_styles_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="14"/><name val="Calibri"/><family val="2"/></font>
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD6DEE8"/></left><right style="thin"><color rgb="FFD6DEE8"/></right><top style="thin"><color rgb="FFD6DEE8"/></top><bottom style="thin"><color rgb="FFD6DEE8"/></bottom><diagonal/></border>
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="2" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="2" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def build_workbook_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>
  <sheets>
    <sheet name="Траектория" sheetId="1" r:id="rId1"/>
    <sheet name="Сырые данные" sheetId="2" r:id="rId2"/>
    <sheet name="Инструкция" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>'''


def build_workbook_rels_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def build_root_rels_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def build_content_types_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''


def build_core_xml():
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>OpenAI</dc:creator>
  <cp:lastModifiedBy>OpenAI</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''


def build_app_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>'''


def write_xlsx(output_path, weeks, trajectory, raw):
    trajectory_rows = [['Группа', 'Длительность, мин', 'Преподаватель'] + weeks] + trajectory
    raw_rows = [['Неделя', 'Дата', 'Группа', 'Длительность, мин', 'Преподаватель', 'Тема коротко', 'Тема полностью']] + raw
    instruction_lines = [
        'Как обновлять файл',
        '1. Положите новые недельные выгрузки .xlsx в эту же папку.',
        "2. Запустите файл 'Запустить обновление.bat'.",
        "3. Откройте 'Траектория_групп_обновляемая.xlsx'.",
        '',
        'Что делает обновление',
        '— каждая группа остается в одной строке;',
        '— каждая новая неделя добавляется вправо новым столбцом;',
        '— если группа новая, для нее создается новая строка;',
        '— если у группы несколько занятий в одну неделю, они пишутся в одной ячейке с новой строки.',
        '',
        'Важно',
        'Скрипт использует только стандартный Python. Если Python на компьютере не установлен, батник не запустится.',
    ]
    trajectory_widths = [28, 14, 24] + [34] * len(weeks)
    raw_widths = [14, 12, 28, 14, 24, 28, 60]
    with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', build_content_types_xml())
        z.writestr('_rels/.rels', build_root_rels_xml())
        z.writestr('docProps/core.xml', build_core_xml())
        z.writestr('docProps/app.xml', build_app_xml())
        z.writestr('xl/workbook.xml', build_workbook_xml())
        z.writestr('xl/_rels/workbook.xml.rels', build_workbook_rels_xml())
        z.writestr('xl/styles.xml', build_styles_xml())
        z.writestr(
            'xl/worksheets/sheet1.xml',
            build_sheet_xml(
                trajectory_rows,
                trajectory_widths,
                freeze='trajectory',
                autofilter_ref=f'A1:{col_letter(len(trajectory_rows[0]))}{len(trajectory_rows)}',
            ),
        )
        z.writestr(
            'xl/worksheets/sheet2.xml',
            build_sheet_xml(
                raw_rows,
                raw_widths,
                freeze='raw',
                autofilter_ref=f'A1:{col_letter(len(raw_rows[0]))}{len(raw_rows)}',
            ),
        )
        z.writestr('xl/worksheets/sheet3.xml', build_instruction_xml(instruction_lines))


def main():
    folder = os.path.dirname(os.path.abspath(__file__))
    source_files = collect_source_files(folder)
    if not source_files:
        print('Не найдено ни одной недельной выгрузки .xlsx в папке.')
        return
    records = []
    for path in source_files:
        try:
            records.extend(parse_records(path))
        except Exception as e:
            print(f'Ошибка при чтении файла: {os.path.basename(path)} -> {e}')
    if not records:
        print('Занятия в выгрузках не найдены.')
        return
    weeks, trajectory, raw = aggregate(records)
    output_path = os.path.join(folder, OUTPUT_FILE)
    write_xlsx(output_path, weeks, trajectory, raw)
    print('Готово.')
    print(f'Файл создан: {OUTPUT_FILE}')
    print(f'Обработано недель: {len(weeks)}')
    print(f'Групп в итоге: {len(trajectory)}')


if __name__ == '__main__':
    main()
