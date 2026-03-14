"""
VBA 모듈 동기화: vba_modules 폴더의 .bas 파일을 Excel 워크북에 자동 Import
사용법: python sync_to_excel.py [workbook_path]
"""
import os
import sys
import glob
import time
from datetime import datetime
import tempfile
import shutil

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("ERROR: pywin32가 설치되지 않았습니다.")
    print("다음 명령으로 설치하세요: pip install pywin32")
    sys.exit(1)


# 기본값
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_MODULES_DIR = os.path.join(SCRIPT_DIR, "vba_modules")
DEFAULT_WORKBOOK = os.path.join(SCRIPT_DIR, "이루모스_생산관리.xlsb")
SYNC_LOCK_FILE = os.path.join(SCRIPT_DIR, ".sync_lock")
SYNC_FAIL_COOLDOWN = 300  # 저장 실패 시 5분간 sync 비활성화


# Import 시 건너뛸 컴포넌트 (Excel 기본 객체)
SKIP_COMPONENTS = {"ThisWorkbook", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"}

# vbext_ComponentType: Document(Sheet/Workbook)=100은 제거 불가
VBEXT_CT_STDMODULE = 1
VBEXT_CT_CLASSMODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOCUMENT = 100

# 허용된 모듈만 유지 (이 목록 외 모듈은 절대 생성/유지하지 않음)
ALLOWED_MODULES = {
    "Module1", "Module2", "Module3", "Module4", "Module5", "Module6", "Module7",
    "Module8", "Module9", "Module10", "Module11", "Module12", "Module13", "Module14",
    "Module15", "Module16", "Module17", "Module18", "Module19", "Module20", "Module21",
    "Module22", "Module23", "Module24", "Module25", "Module26",
    "UserForm1", "UserForm2",
}


def find_workbook(workbook_path=None):
    """워크북 경로 결정 (인자 > 환경변수 > 기본값 > 폴더 내 .xlsb/.xlsm 자동탐지)"""
    if workbook_path and os.path.isfile(workbook_path):
        return os.path.abspath(workbook_path)
    candidate = os.environ.get("VBA_WORKBOOK", DEFAULT_WORKBOOK)
    if os.path.isfile(candidate):
        return os.path.abspath(candidate)
    # 폴더 내 Excel 매크로 파일 자동 탐지
    for ext in ("*.xlsb", "*.xlsm"):
        found = glob.glob(os.path.join(SCRIPT_DIR, ext))
        if found:
            return os.path.abspath(found[0])
    return None


def is_excel_running():
    """Excel 프로세스 실행 여부 확인"""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        return excel is not None
    except Exception:
        return False


def _remove_components_not_allowed(vb_components):
    """ALLOWED_MODULES에 없는 모듈/UserForm만 제거 (Module101, Module102 등 엑셀 자동생성 잔여물)"""
    removed = 0
    while True:
        found = False
        for i in range(vb_components.Count, 0, -1):
            try:
                comp = vb_components.Item(i)
                ctype = getattr(comp, "Type", VBEXT_CT_DOCUMENT)
                if ctype == VBEXT_CT_DOCUMENT:
                    continue
                name = comp.Name
                if name in SKIP_COMPONENTS or name in ALLOWED_MODULES:
                    continue
                vb_components.Remove(comp)
                removed += 1
                print(f"  [제거] {name}")
                found = True
                break
            except Exception:
                pass
        if not found:
            break
    return removed


def _remove_all_vba_components(vb_components):
    """표준모듈/클래스/UserForm 전부 제거 (Document=Sheet/Workbook 제외)"""
    removed = 0
    while True:
        found = False
        for i in range(vb_components.Count, 0, -1):
            try:
                comp = vb_components.Item(i)
                ctype = getattr(comp, "Type", VBEXT_CT_DOCUMENT)
                if ctype == VBEXT_CT_DOCUMENT:
                    continue
                name = comp.Name
                if name in SKIP_COMPONENTS:
                    continue
                if ctype in (VBEXT_CT_STDMODULE, VBEXT_CT_CLASSMODULE, VBEXT_CT_MSFORM):
                    vb_components.Remove(comp)
                    removed += 1
                    print(f"  [제거] {name}")
                    found = True
                    break
            except Exception as e:
                print(f"  WARNING: 제거 실패 (인덱스 {i}): {e}")
        if not found:
            break
    return removed


def _sync_lock_acquire():
    """동시 sync 방지 - lock 파일 사용"""
    if os.path.isfile(SYNC_LOCK_FILE):
        try:
            age = time.time() - os.path.getmtime(SYNC_LOCK_FILE)
            if age < 300:  # 5분 이내면 다른 sync 진행 중
                return False
        except Exception:
            pass
        try:
            os.remove(SYNC_LOCK_FILE)
        except Exception:
            pass
    try:
        with open(SYNC_LOCK_FILE, "w") as f:
            f.write(str(os.getpid()))
        return True
    except Exception:
        return False


def _sync_lock_release():
    try:
        if os.path.isfile(SYNC_LOCK_FILE):
            os.remove(SYNC_LOCK_FILE)
    except Exception:
        pass


def _write_sync_fail_time():
    try:
        with open(os.path.join(SCRIPT_DIR, ".sync_last_fail"), "w") as f:
            f.write(str(time.time()))
    except Exception:
        pass


def _sync_fail_cooldown_active():
    try:
        path = os.path.join(SCRIPT_DIR, ".sync_last_fail")
        if not os.path.isfile(path):
            return False
        with open(path, "r") as f:
            t = float(f.read().strip())
        return (time.time() - t) < SYNC_FAIL_COOLDOWN
    except Exception:
        return False


def sync_to_excel(workbook_path=None, close_after=True, clean_only=False, skip_git=False):
    """
    vba_modules 폴더의 .bas/.frm → Excel Import
    clean_only=True: 무작위 생성된 모듈/UserForm 전부 삭제만 수행
    skip_git=True: 성공 후 Git 자동 커밋 건너뜀
    미허용 모듈(Module101 등) 절대 생성 방지
    """
    if _sync_fail_cooldown_active():
        print("WARNING: 이전 sync 저장 실패로 5분간 sync 비활성화. 워크북을 닫고 직접 실행하세요.")
        return False
    if not _sync_lock_acquire():
        print("WARNING: 다른 sync가 진행 중입니다. 잠시 후 다시 시도하세요.")
        return False
    try:
        return _sync_to_excel_impl(workbook_path, close_after, clean_only, skip_git)
    finally:
        _sync_lock_release()


def _sync_to_excel_impl(workbook_path=None, close_after=True, clean_only=False, skip_git=False):
    wb_path = find_workbook(workbook_path)
    if not wb_path:
        print(f"ERROR: 워크북을 찾을 수 없습니다. 경로를 확인하세요.")
        print(f"  기본 경로: {DEFAULT_WORKBOOK}")
        print(f"  사용법: python sync_to_excel.py path/to/your/workbook.xlsm")
        return False

    if not clean_only and not os.path.isdir(VBA_MODULES_DIR):
        print(f"ERROR: vba_modules 폴더가 없습니다: {VBA_MODULES_DIR}")
        return False

    # 허용된 모듈만 Import (Module18 대용량 모듈 먼저 → 실패 시 원본파일 재시도)
    _all_bas = [p for p in glob.glob(os.path.join(VBA_MODULES_DIR, "*.bas"))
                 if os.path.splitext(os.path.basename(p))[0] in ALLOWED_MODULES]
    def _import_order(p):
        n = os.path.splitext(os.path.basename(p))[0]
        if n == "Module18":
            return (0, p)
        if n == "Module8":
            return (1, p)
        return (2, p)
    bas_files = [] if clean_only else sorted(_all_bas, key=_import_order)
    frm_files = [] if clean_only else sorted([
        p for p in glob.glob(os.path.join(VBA_MODULES_DIR, "*.frm"))
        if os.path.splitext(os.path.basename(p))[0] in ALLOWED_MODULES
    ])
    if not clean_only and not bas_files and not frm_files:
        print(f"WARNING: vba_modules 폴더에 허용된 .bas/.frm 파일이 없습니다.")
        return True

    excel = None
    workbook = None
    use_existing = False  # 열린 상태에서 동기화 시 Excel 닫지 않음

    wb_abs = os.path.abspath(wb_path)
    wb_abs_norm = os.path.normpath(os.path.normcase(wb_abs))
    wb_basename = os.path.basename(wb_abs)

    def _paths_match(fn):
        if not fn or not isinstance(fn, str):
            return False
        fn_clean = fn.strip().replace("/", "\\")
        fn_norm = os.path.normpath(os.path.normcase(fn_clean))
        if fn_norm == wb_abs_norm:
            return True
        fn_base = os.path.basename(fn_norm)
        if os.path.normcase(fn_base) == os.path.normcase(wb_basename):
            try:
                return os.path.samefile(fn_norm, wb_abs)
            except (OSError, TypeError):
                pass
            if fn_clean.endswith(wb_basename) or fn_clean.endswith(wb_basename.replace("\\", "/")):
                return True
        return False

    try:
        # 1) 이미 열린 워크북 직접 조회 시도 (GetObject로 경로 지정)
        try:
            wb_direct = win32com.client.GetObject(wb_abs)
            if wb_direct and hasattr(wb_direct, "VBProject"):
                workbook = wb_direct
                excel = wb_direct.Application
                use_existing = True
        except Exception:
            pass

        # 2) 못 찾으면 열린 Excel에서 워크북 검색
        if workbook is None:
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                for i in range(1, excel.Workbooks.Count + 1):
                    try:
                        wb = excel.Workbooks(i)
                        fn = str(getattr(wb, "FullName", "") or "")
                        if _paths_match(fn):
                            workbook = wb
                            use_existing = True
                            break
                        wb_name = str(getattr(wb, "Name", "") or "")
                        wb_dir = str(getattr(wb, "Path", "") or "")
                        if os.path.normcase(wb_name) == os.path.normcase(wb_basename):
                            combined = os.path.join(wb_dir, wb_name)
                            if _paths_match(combined):
                                workbook = wb
                                use_existing = True
                                break
                        if wb_name and wb_basename and wb_name.lower() == wb_basename.lower():
                            workbook = wb
                            use_existing = True
                            break
                    except Exception:
                        pass
            except Exception:
                pass

        # 3) 못 찾으면 새 Excel 인스턴스로 열기
        if workbook is None:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                workbook = excel.Workbooks.Open(
                    wb_abs,
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                )
            except Exception as e:
                excel.Quit()
                print(f"ERROR: 워크북 열기 실패: {e}")
                print(f"  경로: {wb_abs}")
                return False

        if workbook is None:
            if not use_existing and excel:
                excel.Quit()
            print("ERROR: 워크북을 열 수 없습니다.")
            return False

        if workbook.ReadOnly and not use_existing:
            workbook.Close(SaveChanges=False)
            excel.Quit()
            print("ERROR: 워크북이 읽기 전용입니다. Excel에서 해당 파일을 활성 창으로 두고 다시 시도하세요.")
            return False

        vb_project = workbook.VBProject
        vb_components = vb_project.VBComponents

        # .bas 파일에서 VB_Name 추출 (Import 전 해당 모듈 제거용)
        def get_bas_vbname(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        if 'Attribute VB_Name' in line and '=' in line:
                            parts = line.split("=", 1)
                            if len(parts) >= 2:
                                return parts[1].strip().strip('"').strip()
                return None
            except Exception:
                try:
                    with open(path, "r", encoding="cp949") as f:
                        for line in f:
                            if 'Attribute VB_Name' in line and '=' in line:
                                parts = line.split("=", 1)
                                if len(parts) >= 2:
                                    return parts[1].strip().strip('"').strip()
                    return None
                except Exception:
                    return None

        # 표준모듈/클래스/UserForm 전부 제거
        n_removed = _remove_all_vba_components(vb_components)
        if clean_only:
            try:
                workbook.Save()
            except Exception:
                workbook.Close(SaveChanges=False)
                excel.Quit()
                _write_sync_fail_time()
                print(f"\nERROR: 저장 실패 - 워크북이 다른 Excel 창에서 열려 있습니다.")
                print("  해당 파일을 모두 닫은 뒤 다시 실행하세요. 5분간 sync 비활성화.")
                return False
            workbook.Close(SaveChanges=True)
            excel.Quit()
            print(f"\n정리 완료: {n_removed}개 모듈/UserForm 삭제됨. 워크북 저장됨.")
            print("  열린 Excel 창이 있다면 '다시 로드' 하세요.")
            if not skip_git:
                try:
                    from git_helpers import git_commit_after_change
                    git_commit_after_change("sync", "정리 모드")
                except Exception:
                    pass
            return True

        def get_frm_vbname(path):
            """ .frm 파일에서 VB_Name 추출 """
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        if 'Attribute VB_Name' in line and '=' in line:
                            parts = line.split("=", 1)
                            if len(parts) >= 2:
                                return parts[1].strip().strip('"').strip()
                return os.path.splitext(os.path.basename(path))[0]
            except Exception:
                return os.path.splitext(os.path.basename(path))[0]

        # .bas, .frm 파일 Import (허용 목록만)
        imported = 0
        for bas_path in bas_files:
            imp_path = bas_path
            temp_path = None  # temp 파일 경로 (삭제용, 원본 bas_path 제외)
            try:
                # Import 전 동일 VB_Name 모듈 완전 제거 (중복 생성 방지)
                vb_name = get_bas_vbname(bas_path)
                if vb_name and vb_name not in SKIP_COMPONENTS:
                    while True:
                        try:
                            comp = vb_components.Item(vb_name)
                            if comp.Type == VBEXT_CT_STDMODULE:
                                vb_components.Remove(comp)
                                continue
                            break
                        except Exception:
                            break
                try:
                    with open(bas_path, "r", encoding="utf-8") as f:
                        content = f.read()
                    fd, imp_path = tempfile.mkstemp(suffix=".bas")
                    os.close(fd)
                    temp_path = imp_path
                    with open(imp_path, "w", encoding="cp949", errors="replace") as f:
                        f.write(content)
                except UnicodeDecodeError:
                    try:
                        with open(bas_path, "r", encoding="cp949") as f:
                            content = f.read()
                        fd, imp_path = tempfile.mkstemp(suffix=".bas")
                        os.close(fd)
                        temp_path = imp_path
                        with open(imp_path, "w", encoding="cp949", errors="replace") as f:
                            f.write(content)
                    except Exception:
                        imp_path = bas_path
                except Exception:
                    imp_path = bas_path
                for attempt in range(3):
                    try:
                        vb_components.Import(imp_path)
                        break
                    except Exception as imp_err:
                        if attempt == 0:
                            time.sleep(1)
                            continue
                        if attempt == 1 and imp_path != bas_path and os.path.isfile(bas_path):
                            imp_path = bas_path
                            time.sleep(0.5)
                            continue
                        raise imp_err
                imported += 1
                print(f"  [OK] Import: {os.path.basename(bas_path)}")
                _remove_components_not_allowed(vb_components)  # Import 직후 즉시 미허용 제거
            except Exception as e:
                print(f"  [FAIL] Import {os.path.basename(bas_path)}: {e}")
                if "Module18" in bas_path:
                    import traceback
                    traceback.print_exc()
            finally:
                if temp_path and os.path.isfile(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

        for frm_path in frm_files:
            imp_frm = frm_path
            temp_frm_dir = None
            try:
                vb_name = get_frm_vbname(frm_path)
                if vb_name and vb_name in ALLOWED_MODULES:
                    while True:
                        try:
                            comp = vb_components.Item(vb_name)
                            if comp.Type in (VBEXT_CT_CLASSMODULE, VBEXT_CT_MSFORM):
                                vb_components.Remove(comp)
                                continue
                            break
                        except Exception:
                            break
                # Excel은 .frm 한글을 CP949로 기대. UTF-8/CP949 둘 다 처리
                try:
                    try:
                        with open(frm_path, "r", encoding="utf-8") as f:
                            content = f.read()
                    except UnicodeDecodeError:
                        with open(frm_path, "r", encoding="cp949") as f:
                            content = f.read()
                    temp_frm_dir = tempfile.mkdtemp()
                    imp_frm = os.path.join(temp_frm_dir, os.path.basename(frm_path))
                    with open(imp_frm, "w", encoding="cp949") as f:
                        f.write(content)
                    frx_path = os.path.join(os.path.dirname(frm_path), os.path.splitext(os.path.basename(frm_path))[0] + ".frx")
                    if os.path.isfile(frx_path):
                        shutil.copy2(frx_path, os.path.join(temp_frm_dir, os.path.basename(frx_path)))
                except Exception:
                    imp_frm = frm_path
                vb_components.Import(imp_frm)
                imported += 1
                print(f"  [OK] Import: {os.path.basename(frm_path)}")
                _remove_components_not_allowed(vb_components)  # Import 직후 즉시 미허용 제거
            except Exception as e:
                print(f"  [FAIL] Import {os.path.basename(frm_path)}: {e}")
            finally:
                if temp_frm_dir and os.path.isdir(temp_frm_dir):
                    try:
                        shutil.rmtree(temp_frm_dir)
                    except Exception:
                        pass

        # Import 후 엑셀이 자동생성한 미허용 모듈(Module101 등) 제거
        extra = _remove_components_not_allowed(vb_components)
        if extra:
            print(f"  [정리] 미허용 모듈 {extra}개 추가 제거")

        try:
            workbook.Save()
        except Exception as save_err:
            if not use_existing:
                workbook.Close(SaveChanges=False)
                excel.Quit()
            _write_sync_fail_time()
            print("\nERROR: 저장 실패 - 워크북이 읽기 전용이거나 다른 Excel에서 열려 있습니다.")
            print("  Excel에서 워크북을 모두 닫은 뒤 다시 실행하세요.")
            print("  이후 5분간 자동 sync 비활성화됨.")
            return False
        if not use_existing:
            workbook.Close(SaveChanges=True)
            excel.Quit()
        done_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n동기화 완료: {imported}개 모듈(.bas/.frm) Import, 워크북 저장됨. 완료: {done_at}")
        if not skip_git:
            try:
                from git_helpers import git_commit_after_change
                git_commit_after_change("sync", f"{imported}개 Import")
            except Exception:
                pass
        return True

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        if excel and not use_existing:
            try:
                for wb in excel.Workbooks:
                    wb.Close(SaveChanges=False)
                excel.Quit()
            except Exception:
                pass
        _write_sync_fail_time()
        return False


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook_path", nargs="?", help="워크북 경로")
    parser.add_argument("--close", action="store_true", help="Sync 후 워크북 닫기")
    parser.add_argument("--clean", action="store_true",
                        help="엑셀 내 모듈/UserForm 전부 삭제 (Import 없음)")
    args = parser.parse_args()
    print("=" * 50)
    print("VBA Sync to Excel" + (" [정리 모드]" if args.clean else ""))
    print("=" * 50)
    success = sync_to_excel(args.workbook_path, close_after=args.close, clean_only=args.clean, skip_git=args.no_git)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
