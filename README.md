# VBA 자동 동기화 개발 환경

Excel VBA 모듈을 외부 `.bas` 파일로 관리하고, Python 스크립트로 자동 동기화합니다.

## 프로젝트 구조

```
project/
├── workbook.xlsm          # Excel 워크북 (경로 변경 가능)
├── vba_modules/           # .bas 파일 저장 폴더
│   ├── modMain.bas
│   ├── modUtil.bas
│   └── ...
├── sync_to_excel.py       # .bas → Excel Import
├── export_from_excel.py   # Excel → .bas Export
├── watch_and_sync.py      # 파일 변경 감지 → 자동 Sync
└── requirements.txt
```

## 1. Excel 보안 설정 (필수)

VBA 프로젝트에 프로그래밍 방식으로 접근하려면 다음을 활성화해야 합니다.

1. **Excel** → **파일** → **옵션**
2. **신뢰 센터** → **신뢰 센터 설정**
3. **매크로 설정** → **"VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스"** 체크
4. 확인 후 Excel 재시작

## 2. 설치

```bash
pip install -r requirements.txt
```

## 3. 사용법

### 기존 Excel에서 모듈 Export

```bash
python export_from_excel.py path/to/your/workbook.xlsm
```

→ `vba_modules/` 폴더에 모든 `.bas` 모듈이 저장됩니다.

### .bas → Excel 동기화

```bash
python sync_to_excel.py path/to/your/workbook.xlsm
```

워크북 경로를 생략하면 `workbook.xlsm` 또는 환경변수 `VBA_WORKBOOK` 사용.

### 워크플로우

```
AI가 VBA 코드 생성
    ↓
.bas 파일 저장 (vba_modules/)
    ↓
python sync_to_excel.py 실행
    ↓
Excel 모듈 자동 업데이트
```

### 자동 Sync (파일 변경 감지)

```bash
python watch_and_sync.py path/to/your/workbook.xlsm
```

`.bas` 파일을 수정하면 1.5초 후 자동으로 Excel에 동기화됩니다.

```
.bas 수정
    ↓
자동 감지
    ↓
Excel 자동 업데이트
```

## 4. 환경변수

| 변수 | 설명 |
|------|------|
| `VBA_WORKBOOK` | 기본 워크북 경로 (sync/export 시 사용) |

## 5. 주의사항

- **Sheet**, **ThisWorkbook** 등은 Export/Import 대상에서 제외됩니다.
- 동기화 시 기존 표준 모듈(.bas)은 모두 제거 후 다시 Import됩니다.
- 30개 이상 모듈도 폴더 전체를 자동 처리합니다.
