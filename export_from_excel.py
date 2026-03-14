"""
VBA 모듈 Export: Excel 워크북의 VBA 모듈을 vba_modules 폴더로 자동 Export
사용법: python export_from_excel.py [workbook_path]
"""
import os
import sys
import glob
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32가 설치되지 않았습니다.")
    print("다음 명령으로 설치하세요: pip install pywin32")
    sys.exit(1)


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_MODULES_DIR = os.path.join(SCRIPT_DIR, "vba_modules")
DEFAULT_WORKBOOK = os.path.join(SCRIPT_DIR, "workbook.xlsm")

# Export 제외 컴포넌트 (Excel 기본 객체)
SKIP_COMPONENTS = {"ThisWorkbook", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"}

# 허용된 모듈만 Export (이 목록 외 모듈은 절대 Export하지 않음 - 기존 모듈만 수정)
ALLOWED_MODULES = {
    "Module1", "Module2", "Module3", "Module4", "Module5", "Module6", "Module7",
    "Module8", "Module9", "Module10", "Module11", "Module12", "Module13", "Module14",
    "Module15", "Module16", "Module17", "Module18", "Module19", "Module20", "Module21",
    "Module22", "Module23", "Module24", "Module25", "Module26",
    "UserForm1", "UserForm2",
}

# vbext_ComponentType
VBEXT_CT_STDMODULE = 1  # 표준 모듈 (.bas)
VBEXT_CT_CLASSMODULE = 2  # 클래스 모듈 (.cls)
VBEXT_CT_MSFORM = 3  # UserForm (.frm)


def deduplicate_frm(path):
    """ .frm 파일에서 중복된 Option Explicit, Sub 블록 제거 """
    for enc in ("cp949", "utf-8"):
        try:
            with open(path, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except Exception:
            continue
    else:
        return
    # 두 번째 "Option Explicit" 또는 동일 이름 Sub 발견 시 그 이전까지만 유지
    seen_option = False
    seen_subs = set()
    cutoff = len(lines)
    i = 0
    while i < len(lines):
        line = lines[i]
        s = line.strip()
        if s == "Option Explicit":
            if seen_option:
                cutoff = i
                break
            seen_option = True
        if s.startswith("Private Sub ") or s.startswith("Public Sub "):
            # "Private Sub CommandButton1_Click()" -> "CommandButton1_Click"
            part = s.split("(", 1)[0].replace("Private Sub ", "").replace("Public Sub ", "").strip()
            if part in seen_subs:
                cutoff = i
                break
            seen_subs.add(part)
        i += 1
    if cutoff < len(lines):
        lines = lines[:cutoff]
        while lines and lines[-1].strip() == "":
            lines.pop()
        try:
            with open(path, "w", encoding="cp949") as f:
                f.writelines(lines)
            print(f"  [DEDUP] {os.path.basename(path)} 중복 제거")
        except Exception:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.writelines(lines)
            except Exception:
                pass


def find_workbook(workbook_path=None):
    """워크북 경로 결정"""
    if workbook_path and os.path.isfile(workbook_path):
        return os.path.abspath(workbook_path)
    candidate = os.environ.get("VBA_WORKBOOK", DEFAULT_WORKBOOK)
    if os.path.isfile(candidate):
        return os.path.abspath(candidate)
    return None


def export_from_excel(workbook_path=None, close_after=False, skip_git=False):
    """Excel의 VBA 모듈을 vba_modules 폴더로 Export. skip_git=True면 자동 커밋 건너뜀"""
    wb_path = find_workbook(workbook_path)
    if not wb_path:
        print(f"ERROR: 워크북을 찾을 수 없습니다.")
        print(f"  기본 경로: {DEFAULT_WORKBOOK}")
        print(f"  사용법: python export_from_excel.py path/to/workbook.xlsm")
        return False

    os.makedirs(VBA_MODULES_DIR, exist_ok=True)

    excel = None
    workbook = None

    try:
        # 새 Excel 인스턴스 사용 (기존 인스턴스와 충돌 방지)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_full = os.path.abspath(wb_path)
        for wb in excel.Workbooks:
            try:
                fn = os.path.normpath(os.path.abspath(wb.FullName))
                if fn == wb_full:
                    workbook = wb
                    break
            except Exception:
                pass

        if workbook is None:
            try:
                workbook = excel.Workbooks.Open(wb_full, 0, True)  # UpdateLinks=0, ReadOnly=True
            except Exception:
                pass
            if workbook is None:
                try:
                    workbook = excel.Workbooks.Open(wb_full, 0, False)  # ReadOnly=False
                except Exception as open_err:
                    print(f"ERROR: 워크북 열기 실패: {open_err}")
                    print("  - 파일이 다른 Excel에서 열려 있으면 먼저 닫아보세요.")
                    print("  - Excel: 파일→옵션→신뢰 센터→'VBA 프로젝트 개체 모델 신뢰' 체크.")
                    excel.Quit()
                    return False

        if workbook is None:
            print("ERROR: 워크북을 열 수 없습니다.")
            print("  - 파일 사용 중이거나 손상 여부 확인")
            excel.Quit()
            return False

        try:
            vb_components = workbook.VBProject.VBComponents
        except Exception as vba_err:
            print(f"ERROR: VBA 프로젝트 접근 실패: {vba_err}")
            print("  Excel: 파일 → 옵션 → 신뢰 센터 → 신뢰 센터 설정 → 매크로 설정")
            print("  → 'VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스' 체크 필요")
            workbook.Close(SaveChanges=False)
            return False

        exported = 0

        for i in range(1, vb_components.Count + 1):
            comp = vb_components.Item(i)
            name = comp.Name
            if name in SKIP_COMPONENTS:
                continue
            if name not in ALLOWED_MODULES:
                continue  # 허용 목록 외 모듈은 절대 Export하지 않음 (Module111 등 중복 생성 방지)
            if comp.Type == VBEXT_CT_STDMODULE:
                ext = ".bas"
            elif comp.Type == VBEXT_CT_CLASSMODULE:
                ext = ".cls"
            elif comp.Type == VBEXT_CT_MSFORM:
                ext = ".frm"
            else:
                continue

            out_path = os.path.join(VBA_MODULES_DIR, f"{name}{ext}")
            try:
                comp.Export(out_path)
                # .bas/.cls: CP949→UTF-8 변환 (편집기 한글). .frm은 CP949 유지 (Excel Import 시 한글 깨짐 방지)
                if ext == ".bas" or ext == ".cls":
                    try:
                        with open(out_path, "r", encoding="cp949") as f:
                            content = f.read()
                        with open(out_path, "w", encoding="utf-8") as f:
                            f.write(content)
                    except Exception:
                        pass
                elif ext == ".frm":
                    deduplicate_frm(out_path)
                exported += 1
                print(f"  [OK] Export: {name}{ext}")
            except Exception as e:
                print(f"  [FAIL] Export {name}: {e}")

        if close_after:
            workbook.Close(SaveChanges=False)
            if excel.Workbooks.Count == 0:
                excel.Quit()

        # 허용 목록 외 .bas/.frm/.cls 파일 삭제 (Module111 등 중복 생성물 제거)
        for fname in os.listdir(VBA_MODULES_DIR):
            if fname.lower().endswith((".bas", ".frm", ".cls")):
                stem = os.path.splitext(fname)[0]
                if stem not in ALLOWED_MODULES:
                    try:
                        p = os.path.join(VBA_MODULES_DIR, fname)
                        if os.path.isfile(p):
                            os.remove(p)
                            print(f"  [DEL] 제거: {fname}")
                    except Exception as e:
                        print(f"  [WARN] {fname} 제거 실패: {e}")

        print(f"\nExport 완료: {exported}개 모듈을 {VBA_MODULES_DIR}에 저장했습니다. (저장시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
        if not skip_git:
            try:
                from git_helpers import git_commit_after_change
                git_commit_after_change("export", f"{exported}개 Export")
            except Exception:
                pass
        return True

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook_path", nargs="?", help="워크북 경로")
    parser.add_argument("--close", action="store_true", help="Export 후 워크북 닫기")
    parser.add_argument("--no-git", action="store_true", help="자동 Git 커밋 건너뜀")
    args = parser.parse_args()
    workbook_path = args.workbook_path
    print("=" * 50)
    print("VBA Export from Excel")
    print("=" * 50)
    success = export_from_excel(workbook_path, close_after=args.close, skip_git=args.no_git)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
