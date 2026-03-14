"""
데이터업데이트 매크로 실행 및 오류 감지
- sync_to_excel 후 데이터업데이트 매크로 실행
- COM 오류 시 상세 출력
"""
import os
import sys

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 필요. pip install pywin32")
    sys.exit(1)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

def find_workbook():
    for ext in ("*.xlsb", "*.xlsm", "*.xlsx"):
        import glob
        found = glob.glob(os.path.join(SCRIPT_DIR, ext))
        if found:
            return os.path.abspath(found[0])
    return None

def main():
    wb_path = find_workbook()
    if not wb_path:
        print("ERROR: Excel 워크북(.xlsb/.xlsm) 없음")
        sys.exit(1)
    print(f"워크북: {wb_path}")

    excel = None
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True

        wb_full = os.path.abspath(wb_path)
        workbook = None
        for wb in excel.Workbooks:
            if os.path.abspath(wb.FullName) == wb_full:
                workbook = wb
                break
        if workbook is None:
            workbook = excel.Workbooks.Open(wb_full, ReadOnly=False)

        print("1. VBA 동기화...")
        from sync_to_excel import sync_to_excel
        sync_to_excel(wb_path, close_after=False)
        workbook.Save()

        print("2. 데이터업데이트 매크로 실행...")
        excel.Application.Run("데이터업데이트")
        print("완료: 데이터업데이트 성공")
        return 0

    except Exception as e:
        err = sys.exc_info()
        print(f"\n오류 발생: {e}")
        if hasattr(e, 'args') and e.args:
            for i, a in enumerate(e.args):
                print(f"  args[{i}]: {a}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())
