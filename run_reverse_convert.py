# -*- coding: utf-8 -*-
"""
역변환 시뮬레이션: 검증.xlsx DB -> 수정된 결과 생성
read_only로 읽어서 새 파일로 저장 (openpyxl MergeCells 호환 이슈 회피)
"""
import openpyxl
from openpyxl import Workbook

def normalize_cat(s):
    if not s: return ''
    s = str(s).strip()
    s = s.replace('KT-①','KT-1').replace('KT-②','KT-2').replace('KT-③','KT-3')
    s = s.replace('KT-④','KT-4').replace('KT-⑤','KT-5').replace('KT-⑥','FT')
    s = s.replace('BT-⑤','KT-5').replace('지관','PRE')
    if len(s) >= 3 and s[:2].upper() == 'KT' and s[2] != '-':
        if s[2].isdigit(): s = 'KT-' + s[2]
    if len(s) >= 4 and s[:3].upper() == 'JRT' and s[3] != '-':
        if s[3].isdigit(): s = 'JRT-' + s[3]
    if len(s) >= 3 and s[:2].upper() == 'BT' and s[2] != '-':
        if s[2].isdigit(): s = 'BT-' + s[2]
    return s

def get_mapping(shift_raw, cat):
    shift_raw = str(shift_raw or '')
    is_kt = cat.startswith('KT') or cat == 'FT'
    base_col = {'KT-1':3,'JRT-1':3,'KT-2':15,'JRT-2':15,'KT-3':27,'BT-1':27,
                'KT-4':39,'BT-2':39,'KT-5':51,'BT-3':51,'FT':63,'PRE':63}.get(cat, -1)
    if base_col < 0: return None
    roll_col, kg_col = base_col + 2, base_col + 9
    if '주간' in shift_raw: start, end = (7, 16) if is_kt else (20, 29)
    elif '야간' in shift_raw: start, end = (35, 44) if is_kt else (48, 57)
    elif '생산일보' in shift_raw: start, end = (7, 16) if is_kt else (20, 29)
    else: return None
    return (start, end, roll_col, kg_col)

def run():
    # read_only로 읽기 (MergeCells 오류 회피)
    wb_in = openpyxl.load_workbook('검증.xlsx', read_only=True, data_only=True)
    ws_db = wb_in['DB']
    ws_res = wb_in['결과']
    
    # 결과 시트 데이터 복사 (최대 70행 80열)
    res_data = [[ws_res.cell(r, c).value for c in range(1, 81)] for r in range(1, 71)]
    res_data_orig = [row[:] for row in res_data]
    
    # 원지 영역 초기화
    for r1, r2 in [(7,17), (20,30), (35,45), (48,58)]:
        for r in range(r1, min(r2, len(res_data))):
            for c in [5,17,29,41,53,65, 12,24,36,48,60,72]:
                if c <= len(res_data[0]): res_data[r-1][c-1] = None
    
    last_row = ws_db.max_row
    
    def write_origin(cat_filter, is_jrt):
        for r in range(2, last_row + 1):
            if not ws_db.cell(r, 4).value or '원지' not in str(ws_db.cell(r, 4).value): continue
            cat = normalize_cat(ws_db.cell(r, 6).value)
            if cat not in cat_filter: continue
            try: seq = int(ws_db.cell(r, 5).value or 0)
            except: seq = 0
            if seq < 1 or seq > 20: continue
            m = get_mapping(ws_db.cell(r, 3).value, cat)
            if not m: continue
            start, end, roll_col, kg_col = m
            target_row = start + seq - 1
            if target_row > end: continue
            roll = ws_db.cell(r, 7).value
            kg = ws_db.cell(r, 8).value
            if roll: res_data[target_row-1][roll_col-1] = roll
            if kg is not None: res_data[target_row-1][kg_col-1] = kg
    
    write_origin(('JRT-1','JRT-2','BT-1','BT-2','BT-3','PRE'), True)
    write_origin(('KT-1','KT-2','KT-3','KT-4','KT-5','FT'), False)
    
    # DB 데이터 수집 (close 전에)
    db_data = [[ws_db.cell(r, c).value for c in range(1, 15)] for r in range(1, last_row + 1)]
    wb_in.close()
    
    # 새 워크북 생성 (기존 파일 로드 시 MergeCells 오류 회피)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)
    
    # DB 시트
    ws_db_dst = wb_out.create_sheet('DB')
    for r, row in enumerate(db_data, 1):
        for c, val in enumerate(row, 1):
            if val is not None: ws_db_dst.cell(r, c).value = val
    
    # 결과 시트 (기존)
    ws_orig = wb_out.create_sheet('결과')
    for r, row in enumerate(res_data_orig, 1):
        for c, val in enumerate(row, 1):
            if c <= 80 and val is not None: ws_orig.cell(r, c).value = val
    
    # 결과_수정후 (역변환 적용)
    ws_new = wb_out.create_sheet('결과_수정후')
    for r, row in enumerate(res_data, 1):
        for c, val in enumerate(row, 1):
            if c <= 80: ws_new.cell(r, c).value = val
    
    wb_out.save('검증_수정.xlsx')
    wb_out.close()
    print('역변환 시뮬레이션 완료. 검증_수정.xlsx 저장 (DB, 결과, 결과_수정후 시트).')

if __name__ == '__main__':
    run()
