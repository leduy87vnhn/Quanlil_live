import csv, os, datetime

CSV_PATH = r"d:\CHUNG - LINH TINH\Phần mềm NTC\Quản lí livestream\QUỐC THUẬN.csv"
LOG_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'vmix_debug.log')

def append_log(line):
    try:
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(f"[{datetime.datetime.now().isoformat()}] {line}\n")
    except Exception as e:
        print('ERROR writing log:', e)


def parse_csv(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception as e:
        append_log(f"TOOLBAR_LOAD error reading CSV: {e}")
        return None
    # find header row 'Số trận'
    table_start = None
    for i, r in enumerate(rows):
        if not r:
            continue
        first = r[0].strip().replace('\ufeff','').lower()
        if first == 'số trận' and len(r) >= 6:
            table_start = i
            break
    if table_start is None:
        append_log('TOOLBAR_LOAD CSV parse: table header not found')
        return None
    data = rows[table_start+1:]
    # count non-empty rows
    valid = [row for row in data if any(cell.strip() for cell in row)]
    return {'table_start': table_start, 'rows_total': len(rows), 'data_rows': valid}


if __name__ == '__main__':
    append_log('TOOLBAR_LOAD pressed (simulated)')
    res = parse_csv(CSV_PATH)
    if res is None:
        append_log(f'TOOLBAR_LOAD result: failed to parse {CSV_PATH}')
        print('FAILED')
    else:
        append_log(f'TOOLBAR_LOAD loaded {CSV_PATH} rows_total={res["rows_total"]} data_rows={len(res["data_rows"])}')
        if res['data_rows']:
            preview = res['data_rows'][0][:6]
            append_log(f'TOOLBAR_LOAD preview first_row={preview}')
        print('OK')
