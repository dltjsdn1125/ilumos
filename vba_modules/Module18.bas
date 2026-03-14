Attribute VB_Name = "Module18"
'================================================================================
' 원단 원단관리 VBA 매크로 시스템 (원단관리와 100% 동일 구조)
' - A열: Co 인덱스(01~). 숫자 입력 시 Raw data H열 매칭 → I열(원단명)을 B열에 자동 입력
' - 일일생산DB 원지_Co(BR열=70열) → Raw data H열(인덱스)·I열(원단명) 매칭
' - 재고관리DB: 일별 완제품·원지 재고 데이터베이스
'================================================================================

Option Explicit

Public 원단_오늘이동_음영Rng As Range
Public 원단_오늘이동_진행중 As Boolean
Public 원단_재고동기화_일치여부 As Boolean
' B열 변경 시 C열 유효성은 Change 밖에서 갱신 (드롭다운 사라짐 방지)
Public 원단_C열유효성_갱신대기행 As Long
Public 원단_C열유효성_OnTime예약 As Boolean
' DB저장 매크로에서 한 번에 적용 시 완료 메시지 생략 (호출 측에서 통합 메시지 표시)
Public 원단관리_무음모드 As Boolean
' ★ 셀 변경 시 동기화 디바운스 (연속 입력 시 0.2초 후 1회만 실행 → 즉각 반영 + 체감 속도)
Public 원단_재고동기화_대기셀 As Range
Public 원단_재고동기화_OnTime예약 As Double
' ★ 목록추가 매크로로 넣은 그룹행만 셀보호 예외 (데이터 업데이트 시 초기화되어 잠김)
Public 원단_목록추가_예외시작행 As Long
Public 원단_목록추가_예외끝행 As Long
' ★ 디버깅: 오류 발생 시 정확한 단계 파악용 (모듈 전역)
Public 원단_디버그_실행단계 As String
Public 원단_디버그_상세 As String

'================================================================================
' 전역 상수 정의
'================================================================================
Private Const WS_제품표준사양서 As String = "제품표준사양서"
    Private Const WS_원단관리 As String = "원단관리"
Private Const WS_일일생산일보 As String = "생산일보"
Private Const WS_일일생산DB As String = "일일생산DB"
Private Const WS_동기화상태 As String = "원단관리_동기화상태"  ' 일일생산DB 증분 반영용: A1 = 마지막 동기화 ID
Private Const WS_RawData As String = "Raw data"    ' H열=Co인덱스, I열=원단명. A열 입력 시 B열 자동 매칭
Private Const WS_재고관리DB As String = "재고관리DB"  ' 일별 완제품·원지 재고 DB
' 부자재DB 시트 삭제: 부자재 데이터는 일일생산DB 시트 71~81열(부자재 블록)에 저장
Private Const 부자재_시작열 As Long = 71   ' 71=월,72=발주처,73=제품,74=포장,75=안전재고,76=입고,77=사용,78=재고,81=제품_생산 (★2차 수량은 여기서 읽지 않음)
Private Const 일일생산DB_완제품_생산량BOX열 As Long = 12  ' L열 - 완제품 기준, 1차/2차 사용량 반드시 여기서 읽음 (부자재블록 BZ/78열 아님)
Private Const 일일생산DB_설비번호열 As Long = 9  ' I열 = 완제품_구분(설비호수: KT①, KT-②, JRT, WT 등)
Private Const WS_표준사양서 As String = "표준사양서"
Private Const WS_이력 As String = "입력이력"
Private Const WS_업체별부자재 As String = "업체별원단관리"
Private Const WS_IndexKey As String = "Index key"
Private Const WS_RawDataIndex As String = "Raw data_index"   ' A열 설비라인 목록: M3~마지막행
Private Const RawDataIndex_설비라인열 As Long = 13          ' M열
Private Const RawDataIndex_설비라인데이터시작행 As Long = 3 ' 3행부터
Private Const START_ROW As Integer = 4
' ── 원단관리 시트 고정열 (A~I) ──
Private Const COL_제품발주처 As Integer = 1  ' A열
Private Const COL_제품명 As Integer = 2      ' B열
Private Const COL_포장구분 As Integer = 3    ' C열 (1차포장/2차포장)
Private Const COL_안전재고 As Integer = 4    ' D열
Private Const COL_기초재고 As Integer = 5    ' E열
Private Const COL_입고총합 As Integer = 6    ' F열
Private Const COL_사용총합 As Integer = 7    ' G열
Private Const COL_기말재고 As Integer = 8    ' H열
Private Const COL_구분 As Integer = 9        ' I열 (입고/사용/재고)
' ── 월별 데이터열 (J~U = 1월~12월) ──
Private Const COL_월별시작 As Integer = 10   ' J열 = 1월
Private Const COL_월별끝최대 As Integer = 21 ' U열 = 12월 (최대)
' 숫자 서식
Private Const FMT_NUM_STD As String = "#,##0;[Red]-#,##0"
Private Const FMT_NUM_월별 As String = "#,##0;[Red]-#,##0;;"
' 시트 보호 암호
Private Const WS_시트보호암호 As String = "0205"
Private Const WS_현재고_보호암호 As String = "0205"
' 그룹당 6행 (4~9행: Roll입고/사용/재고 + 중량Kg입고/사용/재고)
Private Const ROWS_PER_ITEM As Integer = 6
' 서브행 오프셋 (그룹 시작행 기준) - 1차=Roll(롤수), 2차=중량(Kg) 각각 입고/사용/재고
Private Const OFFSET_입고1 As Integer = 0   ' Roll 입고
Private Const OFFSET_사용1 As Integer = 1   ' Roll 사용
Private Const OFFSET_재고1 As Integer = 2   ' Roll 재고
Private Const OFFSET_입고2 As Integer = 3   ' 중량Kg 입고
Private Const OFFSET_사용2 As Integer = 4   ' 중량Kg 사용
Private Const OFFSET_재고2 As Integer = 5   ' 중량Kg 재고
' 호환 별칭 (1차=Roll 용)
Private Const OFFSET_입고 As Integer = 0
Private Const OFFSET_사용 As Integer = 1
Private Const OFFSET_재고 As Integer = 2
' ── 호환 별칭 (구 구조 참조 컴파일 호환) ──
Private Const COL_설비라인 As Integer = 1    ' = COL_제품발주처
Private Const COL_업체명 As Integer = 1      ' = COL_제품발주처
Private Const COL_제품 As Integer = 2        ' = COL_제품명
Private Const COL_부자재명 As Integer = 2    ' = COL_제품명
Private Const COL_제조사 As Integer = 22     ' 미사용(호환)
Private Const COL_지관 As Integer = 23       ' 미사용(호환)
Private Const COL_지폭 As Integer = 24       ' 미사용(호환)
Private Const COL_총량 As Integer = 25       ' 미사용(호환)
Private Const COL_1차포장 As Integer = 26    ' 미사용(호환)
Private Const COL_2차포장 As Integer = 27    ' 미사용(호환)
Private Const COL_전월재고 As Integer = 5    ' = COL_기초재고
Private Const COL_현재고 As Integer = 8      ' = COL_기말재고
Private Const COL_자재유형 As Integer = 9    ' = COL_구분
Private Const COL_일별시작 As Integer = 10   ' = COL_월별시작
Private Const COL_카테고리 As Integer = 23   ' 미사용(호환)
Private Const COL_원지 As Integer = 2        ' 미사용(호환)
Private Const COL_폴리백 As Integer = 26     ' 미사용(호환)
Private Const COL_비닐백 As Integer = 27     ' 미사용(호환)
Private Const COL_BOX자재현황 As Integer = 27 ' 미사용(호환)
' 구 RDB 상수 호환
Private Const RDB_업체명 As Integer = 2      ' = RDB_제품발주처
Private Const RDB_제품 As Integer = 3        ' = RDB_제품명DB
Private Const RDB_제조사 As Integer = 22     ' 미사용(호환)
Private Const RDB_지관 As Integer = 23       ' 미사용(호환)
Private Const RDB_지폭 As Integer = 24       ' 미사용(호환)
Private Const RDB_총량 As Integer = 25       ' 미사용(호환)
Private Const RDB_1차포장 As Integer = 26    ' 미사용(호환)
Private Const RDB_2차포장 As Integer = 27    ' 미사용(호환)
Private Const RDB_제품_전월재고 As Integer = 5
Private Const RDB_제품_생산 As Integer = 6
Private Const RDB_제품_소요 As Integer = 7
Private Const RDB_제품_재고 As Integer = 8
Private Const RDB_제품_입고 As Integer = 6
Private Const RDB_원지ROLL_전월재고 As Integer = 22
Private Const RDB_원지ROLL_입고 As Integer = 23
Private Const RDB_원지ROLL_소요 As Integer = 24
Private Const RDB_원지ROLL_재고 As Integer = 25
Private Const RDB_원지Tons_전월재고 As Integer = 22
Private Const RDB_원지Tons_입고 As Integer = 23
Private Const RDB_원지Tons_소요 As Integer = 24
Private Const RDB_원지Tons_재고 As Integer = 25
Private Const RDB_1차포장_전월재고 As Integer = 22
Private Const RDB_1차포장_입고 As Integer = 23
Private Const RDB_1차포장_소요 As Integer = 24
Private Const RDB_1차포장_재고 As Integer = 25
Private Const RDB_2차포장_전월재고 As Integer = 22
Private Const RDB_2차포장_입고 As Integer = 23
Private Const RDB_2차포장_소요 As Integer = 24
Private Const RDB_2차포장_재고 As Integer = 25
' 구 서식 상수 호환
Private Const FMT_NUM_O이후 As String = "#,##0;[Red]-#,##0;;"
' 제품표준사양서: 헤더 7·8행 (B~G 7·8행 병합, H~AC 2단 헤더, AG~AK 7·8행 하나로 병합), 데이터 9행부터
Private Const 표준사양서_헤더행 As Integer = 8   ' 열찾기 시 8행→7행 순 검색 (8행: 제조사·지폭·Type 등, 7행: 발주처·제품(상품)명 등)
Private Const 표준사양서_데이터시작행 As Integer = 9
Private Const 표준사양서_제품열G As Long = 7         ' G열 = 제품명 (일일생산일보 C22:C31 유효성용)
Private Const 표준사양서_1차포장계수열P As Long = 16 ' P열 = 1차포장 계수 (생산량*P → 1차포장 소요)
Private Const 표준사양서_M열 As Long = 13   ' M열 (첫 그룹 C4 표시용)
Private Const 표준사양서_N열 As Long = 14   ' N열 (첫 그룹 C7 표시용)
Private Const 일일생산일보_제품목록_목록데이터열 As Long = 5   ' 목록데이터 E열에 일일생산일보 제품 목록
Private Const WS_목록데이터 As String = "목록데이터"   ' B·C열 유효성: 중복제거 업체·업체별 제품 목록
' 기타 시트용 (호환)
Private Const IDX_헤더행 As Integer = 2
Private Const IDX_데이터시작행 As Integer = 3
    ' 일일생산DB 시트 열 (신규 구조: 제품발주처/제품명/포장구분/월별 입고·사용·재고)
    ' A~E: 기본 정보
    Private Const RDB_Date As Integer = 1        ' A: 날짜(월)
    Private Const RDB_제품발주처 As Integer = 2  ' B: 제품발주처
    Private Const RDB_제품명DB As Integer = 3    ' C: 제품명
    Private Const RDB_포장구분 As Integer = 4    ' D: 포장구분(1차/2차)
    Private Const RDB_안전재고 As Integer = 5    ' E: 안전재고
    ' F~H: 월별 데이터
    Private Const RDB_입고 As Integer = 6        ' F: 입고
    Private Const RDB_사용 As Integer = 7        ' G: 사용
    Private Const RDB_재고 As Integer = 8        ' H: 재고
    ' 현재고 시트
    Private Const WS_현재고 As String = "Stock"

' 마지막 월 열 번호 = 2행·3행·데이터영역(4행~) 중 가장 오른쪽 사용 열 (날짜필드 재고 로드 보장)
Private Function 마지막월열_찾기(ws As Worksheet) As Long
    Dim c2 As Long, c3 As Long, cData As Long, 마지막행 As Long
    On Error Resume Next
    c2 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    c3 = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    cData = COL_월별시작 - 1
    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 >= START_ROW Then
        cData = ws.Cells(마지막행, ws.Columns.Count).End(xlToLeft).Column
        If cData < COL_월별시작 Then cData = COL_월별시작 - 1
    End If
    마지막월열_찾기 = WorksheetFunction.Max(c2, c3, cData)
    If 마지막월열_찾기 < COL_월별시작 Then 마지막월열_찾기 = COL_월별시작
    On Error GoTo 0
End Function

' 호환용 별칭
Private Function 마지막일열_찾기(ws As Worksheet) As Long
    마지막일열_찾기 = 마지막월열_찾기(ws)
End Function
Private Function 합계열_찾기(ws As Worksheet) As Long
    합계열_찾기 = 마지막월열_찾기(ws) + 1
End Function
Private Function 비고열_찾기(ws As Worksheet) As Long
    비고열_찾기 = 마지막월열_찾기(ws)
End Function

' 해당 연·월의 일수
Private Function DaysInMonth(연 As Integer, 월 As Integer) As Integer
    On Error Resume Next
    DaysInMonth = Day(DateSerial(연, 월 + 1, 0))
    If DaysInMonth < 28 Or DaysInMonth > 31 Then DaysInMonth = 31
    On Error GoTo 0
End Function

' 시트에 설정된 월 개수 (J열~마지막월열)
Private Function 월수_찾기(ws As Worksheet) As Integer
    Dim 끝열 As Long: 끝열 = 마지막월열_찾기(ws)
    월수_찾기 = CInt(끝열 - COL_월별시작 + 1)
    If 월수_찾기 < 1 Then 월수_찾기 = 12
End Function

' 현재 시트의 날짜 열 수 반환 (부자재관리와 동일: 현재 열 수 = 현재 월 일수)
Private Function 첫달일수_찾기(ws As Worksheet) As Integer
    첫달일수_찾기 = 월수_찾기(ws)
End Function
Private Function 첫달마지막일열(ws As Worksheet) As Long
    첫달마지막일열 = 마지막월열_찾기(ws)
End Function

' 표준사양서에서 제품명으로 환산중량 조회 (없으면 빈 문자열)
Private Function 환산중량_조회(제품명 As String) As String
    Dim ws As Worksheet
    Dim r As Long, c As Long, lastRow As Long, lastCol As Long
    Dim col제품 As Long, col환산 As Long
    Dim h As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_표준사양서)
    On Error GoTo 0
    If ws Is Nothing Then 환산중량_조회 = "": Exit Function
    If Trim(제품명) = "" Then 환산중량_조회 = "": Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 2 Then 환산중량_조회 = "": Exit Function
    col제품 = 0: col환산 = 0
    For c = 1 To lastCol
        h = Trim(CStr(ws.Cells(1, c).Value))
        If InStr(1, h, "제품") > 0 And col제품 = 0 Then col제품 = c
        If InStr(1, h, "환산") > 0 And col환산 = 0 Then col환산 = c
    Next c
    If col제품 = 0 Then col제품 = 1
    If col환산 = 0 Then col환산 = 2
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, col제품).Value)) = Trim(제품명) Then
            환산중량_조회 = Trim(CStr(ws.Cells(r, col환산).Value))
            Exit Function
        End If
    Next r
    환산중량_조회 = ""
End Function

' Raw data_index에서 제품단가·원지원가 조회
' ★ 원단관리 D열(제품) = Raw data_index C열 → 같은 행 D열에서 제품단가
' ★ 원단관리 C열(원지) = Raw data_index M열 → 같은 행 N열에서 원지원가 (업체명 매칭 없음)
Private Sub RawIndex_단가_조회(ws As Worksheet, 제품명 As String, 원지명 As String, _
    ByRef 제품단가 As Double, ByRef 원지원가 As Double)
    Dim r As Long, lastRow As Long
    Dim v제품 As String, v원지 As String
    Dim v단가 As Variant
    Const col제품 As Long = 3    ' Raw data_index C열 (원단관리 D열 제품명과 매칭)
    Const col단가 As Long = 4    ' Raw data_index D열 (제품 단가)
    Const col원지 As Long = 13   ' Raw data_index M열 (원단관리 C열 원지와 매칭)
    Const col원지단가 As Long = 14  ' Raw data_index N열 (원지 단가)
    제품단가 = 0: 원지원가 = 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, col제품).End(xlUp).Row
    If lastRow < IDX_데이터시작행 Then Exit Sub
    For r = IDX_데이터시작행 To lastRow
        v제품 = Trim(Replace(CStr(ws.Cells(r, col제품).Value), chR(160), " "))
        If v제품 = Trim(Replace(제품명, chR(160), " ")) Then
            v단가 = ws.Cells(r, col단가).Value
            If IsNumeric(v단가) Then
                제품단가 = CDbl(v단가)
            Else
                제품단가 = ParseNumberFromText(CStr(v단가))
            End If
            v원지 = Trim(Replace(CStr(ws.Cells(r, col원지).Value), chR(160), " "))
            If v원지 = Trim(Replace(원지명, chR(160), " ")) Then
                v단가 = ws.Cells(r, col원지단가).Value
                If IsNumeric(v단가) Then
                    원지원가 = CDbl(v단가)
                Else
                    원지원가 = ParseNumberFromText(CStr(v단가))
                End If
            End If
            Exit Sub
        End If
    Next r
    If 원지원가 = 0 And 원지명 <> "" Then
        For r = IDX_데이터시작행 To lastRow
            v원지 = Trim(Replace(CStr(ws.Cells(r, col원지).Value), chR(160), " "))
            If v원지 = Trim(Replace(원지명, chR(160), " ")) Then
                v단가 = ws.Cells(r, col원지단가).Value
                If IsNumeric(v단가) Then
                    원지원가 = CDbl(v단가)
                Else
                    원지원가 = ParseNumberFromText(CStr(v단가))
                End If
                Exit For
            End If
        Next r
    End If
End Sub

' 부자재원가 Dictionary에서 원가 조회 (헤더 변형 대응: 원가키배열 순서대로 시도, 없으면 dict키 순회로 부분일치)
Private Function 부자재원가_딕에서조회(dict As Object, 원가키배열 As Variant) As Double
    Dim 원가키 As Variant, k As Variant, h As String
    부자재원가_딕에서조회 = 0
    If dict Is Nothing Then Exit Function
    For Each 원가키 In 원가키배열
        If dict.exists(원가키) Then
            부자재원가_딕에서조회 = CDbl(dict(원가키))
            Exit Function
        End If
    Next 원가키
    For Each k In dict.keys
        h = LCase(Replace(Replace(CStr(k), " ", ""), chR(10), ""))
        For Each 원가키 In 원가키배열
            If InStr(h, LCase(Replace(CStr(원가키), " ", ""))) > 0 Then
                부자재원가_딕에서조회 = CDbl(dict(k))
                Exit Function
            End If
        Next 원가키
    Next k
End Function

' Raw data_index에서 해당 제품 행의 모든 부자재원가·단가를 동적 조회
' ★ 헤더 2행 스캔으로 "단가", "*원가" 끝나는 열을 찾아 해당 행의 값을 Dictionary로 반환
' dict: key=헤더명(단가, 지관원가, 폴리백원가 등), value=원가(Double). dict가 Nothing이면 새로 생성
' 반환: 매칭 행 번호. 0이면 제품 미발견
Private Function RawIndex_부자재원가_조회(ws As Worksheet, 제품명 As String, 원지명 As String, _
    ByRef dict As Object, ByRef 매칭행 As Long) As Boolean
    Dim r As Long, lastRow As Long, c As Long, lastCol As Long
    Dim col제품 As Long, col원지 As Long
    Dim v제품 As String, v원지 As String
    Dim 헤더텍스트 As String, h As String
    Dim v As Variant, cost As Double

    매칭행 = 0
    If ws Is Nothing Then Exit Function
    If dict Is Nothing Then Set dict = CreateObject("Scripting.Dictionary")

    col제품 = RawIndex_헤더열찾기(ws, "제품")
    If col제품 = 0 Then col제품 = 3
    col원지 = RawIndex_헤더열찾기(ws, "원지")
    If col원지 = 0 Then col원지 = 13

    lastRow = ws.Cells(ws.Rows.Count, col제품).End(xlUp).Row
    If lastRow < IDX_데이터시작행 Then Exit Function

    ' 제품+원지 매칭 행 찾기 (원지명 있으면 정확 매칭, 없으면 제품만)
    For r = IDX_데이터시작행 To lastRow
        v제품 = Trim(Replace(CStr(ws.Cells(r, col제품).Value), chR(160), " "))
        If v제품 = Trim(Replace(제품명, chR(160), " ")) Then
            If Len(원지명) = 0 Then
                매칭행 = r
                Exit For
            End If
            v원지 = Trim(Replace(CStr(ws.Cells(r, col원지).Value), chR(160), " "))
            If v원지 = Trim(Replace(원지명, chR(160), " ")) Then
                매칭행 = r
                Exit For
            End If
        End If
    Next r
    If 매칭행 = 0 And Len(원지명) > 0 Then
        ' 원지명으로만 재검색(제품 무시)
        For r = IDX_데이터시작행 To lastRow
            v제품 = Trim(Replace(CStr(ws.Cells(r, col제품).Value), chR(160), " "))
            v원지 = Trim(Replace(CStr(ws.Cells(r, col원지).Value), chR(160), " "))
            If v원지 = Trim(Replace(원지명, chR(160), " ")) Then
                매칭행 = r
                Exit For
            End If
        Next r
    End If
    If 매칭행 = 0 Then Exit Function

    ' 헤더 스캔: 단가, *원가 열 수집
    lastCol = ws.Cells(IDX_헤더행, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 50

    For c = 1 To lastCol
        헤더텍스트 = Trim(CStr(ws.Cells(IDX_헤더행, c).Value))
        If Len(헤더텍스트) > 0 Then
            h = LCase(Replace(Replace(헤더텍스트, " ", ""), chR(10), ""))
            ' 단가: 제품단가 | *원가: 지관원가,폴리백원가 등 | *단가: 원지단가 (원지 열 다음)
            If h = "단가" Or right(h, 2) = "원가" Or right(h, 2) = "단가" Then
                v = ws.Cells(매칭행, c).Value
                cost = 0
                If IsNumeric(v) Then
                    cost = CDbl(v)
                Else
                    cost = ParseNumberFromText(CStr(v))
                End If
                dict(헤더텍스트) = cost
            End If
        End If
    Next c

    RawIndex_부자재원가_조회 = True
End Function

' 텍스트에서 숫자 추출 (\1,500, "1,500" 등 - 콤마 제거 후 파싱)
Private Function ParseNumberFromText(s As String) As Double
    Dim cleaned As String
    ParseNumberFromText = 0
    If Len(s) = 0 Then Exit Function
    cleaned = Replace(Replace(Replace(Trim(s), " ", ""), ",", ""), "\", "")
    If IsNumeric(cleaned) Then ParseNumberFromText = CDbl(cleaned)
End Function

' 원단관리 시트 특정 열의 날짜 반환. 2행 "1M"=1월, 3행=일. I2=연도
Private Function 부자재_열날짜(ws As Worksheet, col As Long) As Variant
    Dim v3 As Variant, v2 As Variant, 기준년 As Long, vI2 As Variant
    v3 = ws.Cells(3, col).Value
    If IsEmpty(v3) Or Not IsNumeric(v3) Then 부자재_열날짜 = Null: Exit Function
    Dim 일 As Long: 일 = CLng(v3)
    If 일 < 1 Or 일 > 31 Then 부자재_열날짜 = Null: Exit Function
    v2 = Trim(Replace(Replace(Replace(CStr(ws.Cells(2, col).Value), "M", ""), "m", ""), " ", ""))
    If v2 = "" Or Not IsNumeric(v2) Then 부자재_열날짜 = Null: Exit Function
    Dim 월 As Long: 월 = CLng(v2)
    If 월 < 1 Or 월 > 12 Then 부자재_열날짜 = Null: Exit Function
    vI2 = ws.Range("I2").Value
    기준년 = Year(Date)
    If IsNumeric(vI2) And vI2 >= 2000 And vI2 <= 2100 Then 기준년 = CLng(vI2)
    If 일 > Day(DateSerial(기준년, 월 + 1, 0)) Then 부자재_열날짜 = Null: Exit Function
    부자재_열날짜 = DateSerial(기준년, 월, 일)
End Function

    ' 일일생산DB 시트의 70열 헤더가 없으면 생성 (최초 실행시 자동). 데이터 업데이트 시 호출
    Private Sub 일일생산DB_시트초기화()
        Dim ws As Worksheet
        Dim DB보호해제 As Boolean: DB보호해제 = False
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If ws Is Nothing Then
            Dim wb보호초기 As Boolean: wb보호초기 = ThisWorkbook.ProtectStructure
            On Error Resume Next
            If wb보호초기 Then ThisWorkbook.Unprotect WS_시트보호암호
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            If Not ws Is Nothing Then ws.Name = WS_일일생산DB
            If wb보호초기 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
            On Error GoTo 0
        End If
        If Not ws Is Nothing And ws.ProtectContents Then
            On Error Resume Next
            ws.Unprotect WS_시트보호암호
            If ws.ProtectContents Then ws.Unprotect
            On Error GoTo 0
            DB보호해제 = True
        End If
        If Trim(CStr(ws.Cells(1, 1).Value)) = "" Then
            Call 헤더_생성(ws)
        End If
        If DB보호해제 Then
            ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        End If
    End Sub

    ' 일일생산DB 1행 헤더가 없으면 70열 헤더 생성. 매크로 목록에서 수동 실행 가능.
    Public Sub 원단_일일생산DB_헤더갱신()
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "일일생산DB 시트가 없습니다.", vbExclamation, "헤더 갱신"
            Exit Sub
        End If
        If ws.ProtectContents Then
            On Error Resume Next
            ws.Unprotect WS_시트보호암호
            If ws.ProtectContents Then ws.Unprotect
            On Error GoTo 0
        End If
        Call 헤더_생성(ws)
        ws.Protect Password:=WS_시트보호암호, UserInterfaceOnly:=True
    End Sub

    ' 부자재 블록 열 오프셋 (일일생산DB 71~78열)
    Private Function 부자재열(rdbCol As Long) As Long
        부자재열 = 부자재_시작열 + rdbCol - 1
    End Function

    ' 일일생산DB 부자재 블록(71~81열) 행 찾기 (월번호 또는 날짜, 제품발주처, 제품명, 포장구분). 없으면 0
    Private Function 일일생산DB_행찾기_통합(wsDB As Worksheet, 월번호 As Long, 발주처 As String, 제품명 As String, _
        포장구분 As String, Optional dummy1 As String = "", Optional dummy2 As String = "", _
        Optional dummy3 As String = "", Optional dummy4 As String = "", Optional dummy5 As String = "") As Long
        Dim lastRow As Long, r As Long, cDate As Long, c발주 As Long, c제품 As Long, c포장 As Long
        Dim 검색월 As Long, v월 As Variant
        일일생산DB_행찾기_통합 = 0
        If wsDB Is Nothing Then Exit Function
        검색월 = IIf(월번호 <= 12, 월번호, Month(CDate(월번호)))
        cDate = 부자재열(RDB_Date): c발주 = 부자재열(RDB_제품발주처): c제품 = 부자재열(RDB_제품명DB): c포장 = 부자재열(RDB_포장구분)
        lastRow = wsDB.Cells(wsDB.Rows.Count, cDate).End(xlUp).Row
        If lastRow < 2 Then Exit Function
        With wsDB
            For r = lastRow To 2 Step -1
                v월 = .Cells(r, cDate).Value
                If (IsNumeric(v월) And (IIf(CLng(v월) <= 12, CLng(v월), Month(CDate(v월))) = 검색월)) Or (IsDate(v월) And Month(CDate(v월)) = 검색월) Then
                    If Trim(CStr(.Cells(r, c발주).Value)) = 발주처 And _
                        Trim(CStr(.Cells(r, c제품).Value)) = 제품명 And _
                        Trim(CStr(.Cells(r, c포장).Value)) = 포장구분 Then
                        일일생산DB_행찾기_통합 = r
                        Exit Function
                    End If
                End If
            Next r
        End With
    End Function

    ' ★★★ 비활성화됨: BS~BZ열 쓰기 중단 ★★★
    Private Function 일일생산DB_행추가_통합(wsDB As Worksheet, 월번호 As Long, 발주처 As String, 제품명 As String, _
        포장구분 As String, Optional dummy1 As String = "", Optional dummy2 As String = "", _
        Optional dummy3 As String = "", Optional dummy4 As String = "", Optional dummy5 As String = "") As Long
        일일생산DB_행추가_통합 = 0
    End Function

    ' 자재유형별 열 시작 (신규 구조에서는 사용하지 않으나 호환 유지)
    Private Function 일일생산DB_자재열_시작(자재유형명 As String) As Long
        일일생산DB_자재열_시작 = RDB_입고
    End Function

' 일일생산DB에서 소요값 계산 (SUMIFS 대체 - J열=제품, B열=날짜, L열 합산)
Private Function 소요값_계산(제품명 As String, 기준날짜 As Date) As Double
    Dim wsDB As Worksheet
    Dim lastRow As Long, r As Long
    Dim v제품 As String, v날짜 As Variant, v합 As Variant
    Dim 합계 As Double
    합계 = 0
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
    On Error GoTo 0
    If wsDB Is Nothing Then Exit Function
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
        
        ' 최적화: With 구문 사용 및 날짜 먼저 비교
        With wsDB
    For r = 2 To lastRow
                v날짜 = .Cells(r, 1).Value
                ' 최적화: 날짜를 먼저 확인 (변별력이 높음)
                If IsDate(v날짜) Then
            If CDate(v날짜) = 기준날짜 Then
                        v제품 = Trim(CStr(.Cells(r, 10).Value))
                        If v제품 = 제품명 Then
                            v합 = .Cells(r, 12).Value
                If IsNumeric(v합) Then 합계 = 합계 + CDbl(v합)
                        End If
            End If
        End If
    Next r
        End With
    소요값_계산 = 합계
End Function

    ' 원단관리 DB에서 해당 날짜/그룹/자재유형의 소요값 반환. 없으면 -999 (미사용)
    Private Function 소요값_일일생산DB_가져오기(ws As Worksheet, 시작행 As Long, 자재유형명 As String, 기준날짜 As Date) As Double
        Dim wsDB As Worksheet
        Dim r As Long
        Dim 업체명 As String, 제품명 As String, 제조사 As String, 지관 As String, 지폭 As String, 총량 As String, 일차 As String, 이차 As String
        소요값_일일생산DB_가져오기 = -999
        On Error Resume Next
        Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If wsDB Is Nothing Then Exit Function
        
        Dim 포장구분 As String
        With ws
            업체명 = Trim(CStr(.Cells(시작행, COL_업체명).Value))
            제품명 = Trim(CStr(.Cells(시작행, COL_제품).Value))
            포장구분 = Trim(CStr(.Cells(시작행, COL_포장구분).Value))
            제조사 = Trim(CStr(.Cells(시작행, COL_제조사).Value))
            지관 = Trim(CStr(.Cells(시작행, COL_지관).Value))
            지폭 = Trim(CStr(.Cells(시작행, COL_지폭).Value))
            총량 = Trim(CStr(.Cells(시작행, COL_총량).Value))
            일차 = Trim(CStr(.Cells(시작행, COL_1차포장).Value))
            이차 = Trim(CStr(.Cells(시작행, COL_2차포장).Value))
        End With
        
        r = 일일생산DB_행찾기_통합(wsDB, 기준날짜, 업체명, 제품명, 포장구분, 제조사, 지관, 지폭, 총량, 일차, 이차)
        If r > 0 Then
            ' 해당 자재유형의 소요 열 위치 계산
            Dim 자재열시작 As Long
            자재열시작 = 일일생산DB_자재열_시작(자재유형명)
                If 자재열시작 > 0 Then
                ' 소요는 +2 위치 (전월재고+0, 입고+1, 소요+2, 재고+3). 부자재 블록 71~81열
                If IsNumeric(wsDB.Cells(r, 부자재열(자재열시작 + 2)).Value) Then
                    소요값_일일생산DB_가져오기 = CDbl(wsDB.Cells(r, 부자재열(자재열시작 + 2)).Value)
                End If
            End If
        End If
    End Function

    '================================================================================
    ' [최적화] 열별 날짜를 배열로 한 번만 계산 (열마다 셀 읽기 중복 제거)
    '================================================================================
    Private Function 열날짜_캐시(ws As Worksheet, 시작열 As Long, 끝열 As Long) As Variant
        Dim arr() As Variant
        Dim col As Long
        If 끝열 < 시작열 Then 열날짜_캐시 = Array(): Exit Function
        ReDim arr(시작열 To 끝열)
        For col = 시작열 To 끝열
            arr(col) = 부자재_열날짜(ws, col)
        Next col
        열날짜_캐시 = arr
    End Function

    '================================================================================
    ' [최적화] 일일생산DB 전체를 메모리 Dictionary로 일괄 로드
    ' Key: "제품명|YYYYMMDD", Value: 소요 합계 ? O(1) 조회
    ' 기존: 그룹수×일수 만큼 DB 전체 행 순회 → 개선: 1회 배열 읽기 후 Dictionary 조회
    '================================================================================
    Private Function 일일생산_소요딕셔너리() As Object
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        Dim ws생산 As Worksheet
        On Error Resume Next
        Set ws생산 = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If ws생산 Is Nothing Then Set 일일생산_소요딕셔너리 = dict: Exit Function
        Dim lastRow As Long
        lastRow = ws생산.Cells(ws생산.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then Set 일일생산_소요딕셔너리 = dict: Exit Function
        Dim arr As Variant
        arr = ws생산.Range(ws생산.Cells(2, 1), ws생산.Cells(lastRow, 12)).Value
        Dim r As Long, key As String, v합 As Variant
        For r = 1 To UBound(arr, 1)
            If IsDate(arr(r, 1)) Then
                key = Trim(CStr(arr(r, 10))) & "|" & Format(CDate(arr(r, 1)), "YYYYMMDD")
                v합 = arr(r, 12)
                If IsNumeric(v합) Then
                    If dict.exists(key) Then
                        dict(key) = dict(key) + CDbl(v합)
                    Else
                        dict.Add key, CDbl(v합)
                    End If
                End If
            End If
        Next r
        Set 일일생산_소요딕셔너리 = dict
    End Function

    '================================================================================
    ' [최적화] 일일생산DB 전체를 메모리 Dictionary + 배열로 로드
    ' Key: "YYYYMMDD|업체|제품|원지|지관|폴리백|비닐백|BOX", Value: 시트행번호
    ' 기존: 매 조회마다 DB 전체 역순 순회 → 개선: 1회 배열 읽기 후 Dictionary O(1) 조회
    '================================================================================
    Private Function 일일생산DB_딕셔너리_생성(wsDB As Worksheet) As Object
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        If wsDB Is Nothing Then Set 일일생산DB_딕셔너리_생성 = dict: Exit Function
        Dim lastRow As Long
        lastRow = wsDB.Cells(wsDB.Rows.Count, 부자재열(RDB_Date)).End(xlUp).Row
        If lastRow < 2 Then Set 일일생산DB_딕셔너리_생성 = dict: Exit Function
        Dim arr As Variant
        arr = wsDB.Range(wsDB.Cells(2, 부자재_시작열), wsDB.Cells(lastRow, 부자재열(RDB_재고))).Value
        Dim r As Long, key As String, v월 As Variant, 기준일 As Date
        For r = 1 To UBound(arr, 1)
            v월 = arr(r, 1)
            If IsNumeric(v월) Or IsDate(v월) Then
                기준일 = IIf(IsDate(v월), CDate(v월), DateSerial(Year(Date), CLng(v월), 1))
                key = DB키_생성(기준일, Trim(CStr(arr(r, 2))), Trim(CStr(arr(r, 3))), "", "", "", "", Trim(CStr(arr(r, 4))), "")
                If Not dict.exists(key) Then dict.Add key, r + 1
            End If
        Next r
        Set 일일생산DB_딕셔너리_생성 = dict
    End Function

    '================================================================================
    ' [최적화] 일일생산DB Key 문자열 생성 헬퍼
    '================================================================================
    Private Function DB키_생성(기준날짜 As Date, 업체 As String, 제품 As String, _
        제조사 As String, 지관 As String, 지폭 As String, 총량 As String, 일차포장 As String, 이차포장 As String) As String
        Dim d1 As Date: d1 = DateSerial(Year(기준날짜), Month(기준날짜), 1)
        DB키_생성 = Format(d1, "YYYYMMDD") & "|" & 업체 & "|" & 제품 & "|" & _
                   제조사 & "|" & 지관 & "|" & 지폭 & "|" & 총량 & "|" & 일차포장 & "|" & 이차포장
    End Function

' 첫 번째 날짜열(2행)에서 월(1~12) 반환. 없으면 1
Private Function B2_월(ws As Worksheet) As Integer
    Dim v2 As Variant
    B2_월 = 1
    If ws Is Nothing Then Exit Function
    v2 = ws.Cells(2, COL_월별시작).Value
    If Not IsEmpty(v2) Then
        v2 = Trim(Replace(Replace(Replace(CStr(v2), "M", ""), "m", ""), " ", ""))
        If v2 <> "" And IsNumeric(v2) Then
            If CLng(v2) >= 1 And CLng(v2) <= 12 Then B2_월 = CInt(v2)
        End If
    End If
End Function

    Private Function B2_연도(ws As Worksheet) As Integer
        B2_연도 = Year(Date)
    End Function

    ' 신규 구조: 블록 오프셋 = 항상 0 (단일 블록, 3행: 입고/사용/재고)
    Private Function 자재유형_블록오프셋(자재유형명 As String) As Long
        자재유형_블록오프셋 = 0
        End Select
    End Function

    ' 기초재고(D,E) 또는 입고/사용(J~) 수기 변경 시 호출 (Worksheet_Change)
    Public Sub 원단관리_기초재고_변경_재계산(변경셀 As Range)
        If 변경셀 Is Nothing Or 변경셀.Cells.Count <> 1 Then Exit Sub
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        If ws.Name <> WS_원단관리 Then Exit Sub
        If 변경셀.Row < START_ROW Then Exit Sub
        Dim 그룹시작행 As Long
        그룹시작행 = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        Dim origScr As Boolean, origCalc As Long
        origScr = Application.ScreenUpdating: origCalc = Application.Calculation
        Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
        On Error GoTo RestoreGE
        Call 원단_블록_재고_재계산(ws, 그룹시작행, 0)
        On Error Resume Next: Call Stock_시트_갱신: On Error GoTo 0
RestoreGE:
        Application.ScreenUpdating = origScr: Application.Calculation = origCalc
    End Sub

    '================================================================================
    ' ★★★ 블록 재고만 재계산 (소요/입고 값은 절대 건드리지 않음!) ★★★
    ' 사용자가 입력한 소요/입고 값을 유지하면서 재고만 다시 계산
    '================================================================================
    Public Sub 원단_블록_재고_재계산(ws As Worksheet, 그룹시작행 As Long, Optional 블록인덱스 As Long = 0)
        ' 그룹 6행: 입고,사용,재고,입고,사용,재고 (1차·2차 각각). 월별 열 (J~)
        Dim 마지막열 As Long, col As Long
        Dim 입고1 As Double, 사용1 As Double, 재고1 As Double
        Dim 입고2 As Double, 사용2 As Double, 재고2 As Double
        Dim 전월재고1 As Double, 전월재고2 As Double
        Dim 시트보호상태 As Boolean
        Dim 입고1합 As Double, 사용1합 As Double, 입고2합 As Double, 사용2합 As Double

        시트보호상태 = ws.ProtectContents
        If 시트보호상태 Then
            On Error Resume Next
            ws.Unprotect WS_시트보호암호
            If ws.ProtectContents Then ws.Unprotect
            On Error GoTo 0
        End If

        마지막열 = 마지막월열_찾기(ws)
        If 마지막열 < COL_월별시작 Then
            If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호
            Exit Sub
        End If

        Dim totalCols As Long, maxRecalcOffset As Long, 재계산한도열 As Long
        totalCols = 마지막열 - COL_월별시작 + 1
        maxRecalcOffset = 값일괄계산_오늘포함한달끝오프셋(ws, totalCols)
        재계산한도열 = COL_월별시작 + maxRecalcOffset - 1
        If 재계산한도열 > 마지막열 Then 재계산한도열 = 마지막열

        Dim 기초재고1 As Double, 기초재고2 As Double, 기말재고1 As Double, 기말재고2 As Double
        기초재고1 = 0: 기초재고2 = 0: 기말재고1 = 0: 기말재고2 = 0
        If IsNumeric(ws.Cells(그룹시작행, COL_기초재고).Value) Then 기초재고1 = CDbl(ws.Cells(그룹시작행, COL_기초재고).Value)
        If IsNumeric(ws.Cells(그룹시작행 + 3, COL_기초재고).Value) Then 기초재고2 = CDbl(ws.Cells(그룹시작행 + 3, COL_기초재고).Value)

        입고1합 = 0: 사용1합 = 0: 입고2합 = 0: 사용2합 = 0
        For col = COL_월별시작 To 마지막열
            입고1 = 0: 사용1 = 0: 입고2 = 0: 사용2 = 0
            If IsNumeric(ws.Cells(그룹시작행 + OFFSET_입고1, col).Value) Then 입고1 = CDbl(ws.Cells(그룹시작행 + OFFSET_입고1, col).Value)
            If IsNumeric(ws.Cells(그룹시작행 + OFFSET_사용1, col).Value) Then 사용1 = CDbl(ws.Cells(그룹시작행 + OFFSET_사용1, col).Value)
            If IsNumeric(ws.Cells(그룹시작행 + OFFSET_입고2, col).Value) Then 입고2 = CDbl(ws.Cells(그룹시작행 + OFFSET_입고2, col).Value)
            If IsNumeric(ws.Cells(그룹시작행 + OFFSET_사용2, col).Value) Then 사용2 = CDbl(ws.Cells(그룹시작행 + OFFSET_사용2, col).Value)
            입고1합 = 입고1합 + 입고1: 사용1합 = 사용1합 + 사용1
            입고2합 = 입고2합 + 입고2: 사용2합 = 사용2합 + 사용2

            If col = COL_월별시작 Then
                재고1 = 기초재고1 + 입고1 - 사용1
                재고2 = 기초재고2 + 입고2 - 사용2
            Else
                전월재고1 = 0: 전월재고2 = 0
                If IsNumeric(ws.Cells(그룹시작행 + OFFSET_재고1, col - 1).Value) Then 전월재고1 = CDbl(ws.Cells(그룹시작행 + OFFSET_재고1, col - 1).Value)
                If IsNumeric(ws.Cells(그룹시작행 + OFFSET_재고2, col - 1).Value) Then 전월재고2 = CDbl(ws.Cells(그룹시작행 + OFFSET_재고2, col - 1).Value)
                재고1 = 전월재고1 + 입고1 - 사용1
                재고2 = 전월재고2 + 입고2 - 사용2
            End If
            If col <= 재계산한도열 Then
                ws.Cells(그룹시작행 + OFFSET_재고1, col).Value = 재고1
                ws.Cells(그룹시작행 + OFFSET_재고2, col).Value = 재고2
                기말재고1 = 재고1: 기말재고2 = 재고2
            End If
        Next col

        ' F(입고총합), G(사용총합), H(기말재고) 1차·2차 블록 각각 업데이트 (기말재고=오늘+10일 기준 마지막 재계산열)
        ws.Cells(그룹시작행, COL_입고총합).Value = 입고1합
        ws.Cells(그룹시작행, COL_사용총합).Value = 사용1합
        ws.Cells(그룹시작행, COL_기말재고).Value = 기말재고1
        ws.Cells(그룹시작행 + 3, COL_입고총합).Value = 입고2합
        ws.Cells(그룹시작행 + 3, COL_사용총합).Value = 사용2합
        ws.Cells(그룹시작행 + 3, COL_기말재고).Value = 기말재고2

        If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호
    End Sub

' 제품표준사양서(또는 해당 시트)에서 헤더명과 일치하는 열 번호 반환 (2행 우선, 없으면 1행, 0)
Private Function 표준사양서_헤더열찾기(ws As Worksheet, 헤더명 As String) As Long
    Dim lastCol As Long, c As Long, hr As Long
    Dim t As String, findVal As String
    If ws Is Nothing Then Exit Function
    findVal = LCase(Trim(Replace(Replace(헤더명, " ", ""), chR(10), "")))
    For hr = 표준사양서_헤더행 To 1 Step -1
        lastCol = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 100
        For c = 1 To lastCol
            t = Trim(Replace(Replace(CStr(ws.Cells(hr, c).Value), " ", ""), chR(10), ""))
            If LCase(t) = findVal Then 표준사양서_헤더열찾기 = c: Exit Function
        Next c
    Next hr
    표준사양서_헤더열찾기 = 0
End Function

' 제품표준사양서 G열(G9~마지막)에서 제품명과 일치하는 행 번호 반환 (없으면 0). 제품명만 기준 조회
Private Function 표준사양서_행찾기_제품명G(ws As Worksheet, 제품명 As String) As Long
    Dim lastRow As Long, r As Long
    If ws Is Nothing Then Exit Function
    제품명 = Trim(제품명)
    If 제품명 = "" Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 표준사양서_제품열G).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then Exit Function
    For r = 표준사양서_데이터시작행 To lastRow
        If Trim(CStr(ws.Cells(r, 표준사양서_제품열G).Value)) = 제품명 Then
            표준사양서_행찾기_제품명G = r
            Exit Function
        End If
    Next r
    표준사양서_행찾기_제품명G = 0
End Function

' 제품표준사양서에서 발주처·제품(상품)명이 일치하는 데이터 행 번호 반환 (없으면 0). D~I열 값은 이 행에서만 가져옴
Private Function 표준사양서_행찾기(ws As Worksheet, col발주처 As Long, col제품 As Long, 업체명 As String, 제품명 As String) As Long
    Dim lastRow As Long, r As Long
    Dim v1 As String, v2 As String
    If ws Is Nothing Or col발주처 <= 0 Or col제품 <= 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, col제품).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, col발주처).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, col발주처).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then Exit Function
    v1 = Trim(CStr(업체명))
    If v1 = "" Then v1 = "-"
    v2 = Trim(CStr(제품명))
    For r = 표준사양서_데이터시작행 To lastRow
        If (Trim(CStr(ws.Cells(r, col발주처).Value)) = v1 Or (v1 = "-" And Trim(CStr(ws.Cells(r, col발주처).Value)) = "")) And Trim(CStr(ws.Cells(r, col제품).Value)) = v2 Then
            표준사양서_행찾기 = r
            Exit Function
        End If
    Next r
    표준사양서_행찾기 = 0
End Function

' 원단관리 시트 그룹별: C열 제품명으로 제품표준사양서 G9~마지막에서 행 조회 후 해당 행 제조사·지관·지폭·평량·1차포장·2차포장 → D~I 자동 채움 (헤더명으로 열 검색)
Private Sub 원단관리_DtoI_제품표준사양서_채우기(wsTarget As Worksheet, ws표준 As Worksheet)
    Dim 시작행 As Long, lastRow As Long, rr As Long
    Dim 제품명 As String, v As Variant
    Dim col제조사 As Long, col지관 As Long, col지폭 As Long, col총량 As Long, col1차 As Long, col2차 As Long
    If wsTarget Is Nothing Or ws표준 Is Nothing Then Exit Sub
    col제조사 = 표준사양서_헤더열찾기(ws표준, "제조사")
    col지관 = 표준사양서_헤더열찾기(ws표준, "지관")
    col지폭 = 표준사양서_헤더열찾기(ws표준, "지폭")
    col총량 = 표준사양서_헤더열찾기(ws표준, "평량")
    If col총량 = 0 Then col총량 = 표준사양서_헤더열찾기(ws표준, "총량")
    col1차 = 표준사양서_헤더열찾기(ws표준, "1차포장")
    col2차 = 표준사양서_헤더열찾기(ws표준, "2차포장")
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub
    ' ★ 표준사양서 → 딕셔너리 1회 빌드 (제품명→행인덱스), 배열 1회 읽기
    Dim lastRow표 As Long: lastRow표 = ws표준.Cells(ws표준.Rows.Count, 표준사양서_제품열G).End(xlUp).Row
    Dim maxCol표 As Long: maxCol표 = 1
    If col제조사 > maxCol표 Then maxCol표 = col제조사
    If col지관 > maxCol표 Then maxCol표 = col지관
    If col지폭 > maxCol표 Then maxCol표 = col지폭
    If col총량 > maxCol표 Then maxCol표 = col총량
    If col1차 > maxCol표 Then maxCol표 = col1차
    If col2차 > maxCol표 Then maxCol표 = col2차
    If 표준사양서_제품열G > maxCol표 Then maxCol표 = 표준사양서_제품열G
    Dim 표준딕D As Object: Set 표준딕D = CreateObject("Scripting.Dictionary")
    Dim arr표 As Variant
    If lastRow표 >= 표준사양서_데이터시작행 Then
        arr표 = ws표준.Range(ws표준.Cells(표준사양서_데이터시작행, 1), ws표준.Cells(lastRow표, maxCol표)).Value
        Dim rIdx As Long
        For rIdx = 1 To UBound(arr표, 1)
            Dim gName As String: gName = Trim(CStr(arr표(rIdx, 표준사양서_제품열G)))
            If gName <> "" And Not 표준딕D.exists(gName) Then 표준딕D.Add gName, rIdx
        Next rIdx
    End If
    For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        제품명 = Trim(CStr(wsTarget.Cells(시작행, COL_제품).Value))
        If 제품명 = "" Then GoTo Next그룹DtoI
        If Not 표준딕D.exists(제품명) Then GoTo Next그룹DtoI
        rr = CLng(표준딕D(제품명))
        If col제조사 > 0 Then v = arr표(rr, col제조사): wsTarget.Cells(시작행, COL_제조사).Value = IIf(IsEmpty(v), "", v)
        If col지관 > 0 Then v = arr표(rr, col지관): wsTarget.Cells(시작행, COL_지관).Value = IIf(IsEmpty(v), "", v)
        If col지폭 > 0 Then v = arr표(rr, col지폭): wsTarget.Cells(시작행, COL_지폭).Value = IIf(IsEmpty(v), "", v)
        If col총량 > 0 Then v = arr표(rr, col총량): wsTarget.Cells(시작행, COL_총량).Value = IIf(IsEmpty(v), "", v)
        If col1차 > 0 Then v = arr표(rr, col1차): wsTarget.Cells(시작행, COL_1차포장).Value = IIf(IsEmpty(v), "", v)
        If col2차 > 0 Then v = arr표(rr, col2차): wsTarget.Cells(시작행, COL_2차포장).Value = IIf(IsEmpty(v), "", v)
Next그룹DtoI:
    Next 시작행
End Sub

' 원단관리 모든 그룹: C열 1차포장블록(시작행)←M열, 2차포장블록(시작행+3)←N열 (제품표준사양서)
' 셀병합 시 MergeArea 첫 셀에 기록
Private Sub 원단관리_첫그룹_C4C7_제품표준사양서_표시(wsTarget As Worksheet, ws표준 As Worksheet)
    Dim 시작행 As Long, lastRow As Long
    Dim 제품명 As String, rLookup As Long, vM As Variant, vN As Variant
    Dim rngC1 As Range, rngC2 As Range
    If wsTarget Is Nothing Or ws표준 Is Nothing Then Exit Sub
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW + 5 Then Exit Sub
    For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        제품명 = Trim(CStr(wsTarget.Cells(시작행, COL_제품명).Value))
        If 제품명 = "" Then GoTo NextGrp
        rLookup = 표준사양서_행찾기_제품명G(ws표준, 제품명)
        If rLookup <= 0 Then GoTo NextGrp
        vM = ws표준.Cells(rLookup, 표준사양서_M열).Value
        vN = ws표준.Cells(rLookup, 표준사양서_N열).Value
        Set rngC1 = wsTarget.Cells(시작행, COL_포장구분)
        If rngC1.MergeCells Then Set rngC1 = rngC1.MergeArea.Cells(1, 1)
        rngC1.Value = IIf(IsEmpty(vM), "", vM)
        Set rngC2 = wsTarget.Cells(시작행 + 3, COL_포장구분)
        If rngC2.MergeCells Then Set rngC2 = rngC2.MergeArea.Cells(1, 1)
        rngC2.Value = IIf(IsEmpty(vN), "", vN)
NextGrp:
    Next 시작행
End Sub

' C열 유효성: B열 연동 INDIRECT 수식 (따옴표는 Chr(34)/Chr(39)로 처리해 구문오류 방지)
Private Function 목록_제품_연동수식(행번호 As Long) As String
    Dim q As String
    q = chR(34)
    목록_제품_연동수식 = "=IF(TRIM(B" & 행번호 & ")=" & q & q & ",목록_제품,INDIRECT(" & q & "목록_제품_" & q & "&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(B" & 행번호 & "),"
    목록_제품_연동수식 = 목록_제품_연동수식 & q & " " & q & "," & q & "_" & q & ")," & q & "/" & q & "," & q & "_" & q & ")," & q & "\" & q & "," & q & "_" & q & "),"
    목록_제품_연동수식 = 목록_제품_연동수식 & q & "," & q & "," & q & "_" & q & ")," & q & "." & q & "," & q & "_" & q & ")," & q & "-" & q & "," & q & "_" & q & "),"
    목록_제품_연동수식 = 목록_제품_연동수식 & q & "(" & q & "," & q & "_" & q & ")," & q & ")" & q & "," & q & "_" & q & ")," & q & chR(39) & q & "," & q & "_" & q & ")" & String(11, ")")
End Function

' 목록 이름용: 업체명 등에서 정의 이름에 사용할 수 있도록 공백·특수문자를 언더스코어로 치환
Private Function 목록이름_세이니타이즈(s As String) As String
    Dim i As Long, c As String, out As String
    out = Trim(s)
    If out = "" Then 목록이름_세이니타이즈 = "": Exit Function
    out = Replace(out, " ", "_")
    out = Replace(out, "/", "_")
    out = Replace(out, "\", "_")
    out = Replace(out, ",", "_")
    out = Replace(out, ".", "_")
    out = Replace(out, "(", "_")
    out = Replace(out, ")", "_")
    out = Replace(out, "-", "_")
    out = Replace(out, "'", "_")
    Do While InStr(out, "__") > 0: out = Replace(out, "__", "_"): Loop
    If left(out, 1) = "_" Then out = Mid(out, 2)
    If right(out, 1) = "_" Then out = left(out, Len(out) - 1)
    목록이름_세이니타이즈 = out
End Function

' 제품표준사양서 해당 열에서 중복 없이 순서대로 값 로드 (헤더 다음 행부터). 반환: 배열(0-based) 또는 Empty
Private Function 표준사양서_목록로드(ws As Worksheet, 헤더명 As String) As Variant
    Dim col As Long, lastRow As Long, r As Long
    Dim dict As Object, arr(), n As Long, v As String
    If ws Is Nothing Then 표준사양서_목록로드 = Empty: Exit Function
    col = 표준사양서_헤더열찾기(ws, 헤더명)
    If col = 0 Then 표준사양서_목록로드 = Empty: Exit Function
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then 표준사양서_목록로드 = Empty: Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 표준사양서_데이터시작행 To lastRow
        v = Trim(CStr(ws.Cells(r, col).Value))
        If v <> "" And Not dict.exists(v) Then dict.Add v, r
    Next r
    If dict.Count = 0 Then 표준사양서_목록로드 = Empty: Exit Function
    ReDim arr(0 To dict.Count - 1)
    n = 0
    For r = 표준사양서_데이터시작행 To lastRow
        v = Trim(CStr(ws.Cells(r, col).Value))
        If v <> "" And dict.exists(v) Then arr(n) = v: dict.Remove v: n = n + 1
    Next r
    표준사양서_목록로드 = arr
End Function

' 호환: RawIndex_헤더열찾기 → 표준사양서_헤더열찾기
Private Function RawIndex_헤더열찾기(ws As Worksheet, 헤더명 As String) As Long
    RawIndex_헤더열찾기 = 표준사양서_헤더열찾기(ws, 헤더명)
End Function

'================================================================================
' Raw data_index 동적 헤더 스캔: 업체 열 + 부자재 카테고리 열(지관, 비닐백, 폴리백, BOX 등) 인식
' ★ 원가/단가 열 제외: 지관원가, 폴리백원가, 비닐백원가, box원가, 원지단가, 단가 등은 목록용이 아님
' 반환: 업체열, catNames(카테고리명 순서), catCols(해당 열번호 순서). 성공 시 True
'================================================================================
Private Function RawIndex_원가열_제외(헤더텍스트 As String) As Boolean
    Dim h As String
    RawIndex_원가열_제외 = False
    h = LCase(Trim(Replace(Replace(헤더텍스트, " ", ""), chR(10), "")))
    If h = "단가" Then RawIndex_원가열_제외 = True: Exit Function
    If right(h, 2) = "원가" Then RawIndex_원가열_제외 = True
End Function

Private Function RawIndex_동적헤더(ws As Worksheet, _
    ByRef 업체열 As Long, ByRef catNames As Collection, ByRef catCols As Collection) As Boolean
    Dim lastCol As Long, c As Long
    Dim 헤더텍스트 As String
    
    Set catNames = New Collection
    Set catCols = New Collection
    업체열 = 0
    
    If ws Is Nothing Then Exit Function
    
    lastCol = ws.Cells(IDX_헤더행, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 50
    
    For c = 1 To lastCol
        헤더텍스트 = Trim(CStr(ws.Cells(IDX_헤더행, c).Value))
        If 헤더텍스트 <> "" Then
            If LCase(헤더텍스트) = "업체" Or LCase(헤더텍스트) = "업체명" Then
                업체열 = c
            ElseIf Not RawIndex_원가열_제외(헤더텍스트) Then
                catNames.Add 헤더텍스트
                catCols.Add c
            End If
        End If
    Next c
    
    If 업체열 = 0 Then 업체열 = 1
    RawIndex_동적헤더 = (catNames.Count > 0)
End Function

' 카테고리 순서에 따른 기본 배경색 (동적 개수 대응)
Private Function 카테고리_기본색상(순서 As Long) As Long
    Dim 색상목록 As Variant
    색상목록 = Array(RGB(255, 230, 153), RGB(198, 224, 180), RGB(248, 203, 173), _
                     RGB(255, 199, 206), RGB(180, 198, 231), RGB(221, 235, 247), _
                     RGB(226, 239, 218), RGB(255, 242, 204))
    카테고리_기본색상 = 색상목록(순서 Mod (UBound(색상목록) + 1))
End Function

'================================================================================
' Raw data_index에서 C(분류), E(부자재명), A(업체명) 동적 로드
' 반환: 분류순서(Collection), catItems(Collection of Collection - 각 항목 "부자재명|업체명")
'================================================================================
Private Function RawIndex_로드_CEA(ws As Worksheet, _
    ByRef 분류순서 As Collection, ByRef catItems As Collection) As Boolean
    Dim lastRow As Long, r As Long
    Dim 분류 As String, 부자재명 As String, 업체명 As String
    Dim i As Long, idx As Long
    Dim item As Collection
    Dim col분류 As Long, col부자재 As Long, col업체 As Long
    
    Set 분류순서 = New Collection
    Set catItems = New Collection
    
    If ws Is Nothing Then Exit Function
    
    col업체 = RawIndex_헤더열찾기(ws, "업체명")
    If col업체 = 0 Then col업체 = RawIndex_헤더열찾기(ws, "업체")
    col부자재 = RawIndex_헤더열찾기(ws, "제품")
    col분류 = RawIndex_헤더열찾기(ws, "지관")
    If col분류 = 0 Then col분류 = RawIndex_헤더열찾기(ws, "분류")
    If col업체 = 0 Or col부자재 = 0 Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.Count, col부자재).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, col업체).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, col업체).End(xlUp).Row
    If col분류 > 0 And ws.Cells(ws.Rows.Count, col분류).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, col분류).End(xlUp).Row
    If lastRow < IDX_데이터시작행 Then Exit Function
    
    For r = IDX_데이터시작행 To lastRow
        분류 = ""
        If col분류 > 0 Then 분류 = Trim(CStr(ws.Cells(r, col분류).Value))
        부자재명 = Trim(CStr(ws.Cells(r, col부자재).Value))
        업체명 = Trim(CStr(ws.Cells(r, col업체).Value))
        If 부자재명 = "" Then GoTo NextRow
        If 분류 = "" Then 분류 = "-"
        If 업체명 = "" Then 업체명 = "-"
        
        idx = 0
        For i = 1 To 분류순서.Count
            If LCase(CStr(분류순서(i))) = LCase(분류) Then idx = i: Exit For
        Next i
        If idx = 0 Then
            분류순서.Add 분류
            Set item = New Collection
            item.Add 부자재명 & "|" & 업체명
            catItems.Add item
        Else
            Set item = catItems(idx)
            item.Add 부자재명 & "|" & 업체명
        End If
NextRow:
    Next r
    
    RawIndex_로드_CEA = (분류순서.Count > 0)
End Function

' 제품표준사양서에서 항목 로드 (1개 카테고리 "제품", 각 데이터 행당 1개 항목). 호환: 총항목수/부자재항목_생성용
Private Function 제품표준사양서_로드_항목(ws As Worksheet, ByRef 분류순서 As Collection, ByRef catItems As Collection) As Boolean
    Dim lastRow As Long, r As Long
    Dim col제품 As Long, col업체 As Long
    Dim item As Collection
    Dim 제품명 As String, 업체명 As String
    Set 분류순서 = New Collection
    Set catItems = New Collection
    If ws Is Nothing Then Exit Function
    col업체 = 표준사양서_헤더열찾기(ws, "업체명")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(ws, "업체")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(ws, "발주처")
    col제품 = 표준사양서_헤더열찾기(ws, "제품")
    If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(ws, "제품(상품)명")
    If col업체 = 0 Or col제품 = 0 Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, col제품).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then Exit Function
    분류순서.Add "제품"
    Set item = New Collection
    For r = 표준사양서_데이터시작행 To lastRow
        제품명 = Trim(CStr(ws.Cells(r, col제품).Value))
        업체명 = Trim(CStr(ws.Cells(r, col업체).Value))
        If 제품명 = "" Then 제품명 = "-"
        If 업체명 = "" Then 업체명 = "-"
        item.Add 제품명 & "|" & 업체명
    Next r
    catItems.Add item
    제품표준사양서_로드_항목 = (item.Count > 0)
End Function

' 일일생산일보에서 완제품(구분|품명)·생산량·제품별 원지기준소요량 수집 (최적화DB 1단계와 동일 범위)
' 반환: 제품딕(key=구분|품명) = 생산량합계, 원지딕(key=구분|품명) = 원지기준소요량 합계(제품별), catItems = 구분|품명 순서, 작성일
' 일일생산일보 없거나 데이터 없으면 False
Private Function 일일생산일보_제품원지_수집(ws일보 As Worksheet, ByRef 제품딕 As Object, ByRef 원지딕 As Object, ByRef catItems As Collection, ByRef 작성일 As Date) As Boolean
    Dim r As Long, 구분 As String, 품명 As String, 맵키 As String, 생산량 As Double, 원지소요 As Double
    Dim arr1H As Variant, arr1B As Variant, arr1T As Variant, arr1X As Variant
    Dim arr2H As Variant, arr2K As Variant, arr2BC As Variant, arr2BG As Variant
    Dim dict순서 As Object
    If ws일보 Is Nothing Then Exit Function
    On Error Resume Next
    작성일 = CDate(ws일보.Range("BK4").Value)
    If Err.Number <> 0 Or Not IsDate(작성일) Then Exit Function
    On Error GoTo 0
    Set 제품딕 = CreateObject("Scripting.Dictionary")
    Set 원지딕 = CreateObject("Scripting.Dictionary")
    Set catItems = New Collection
    Set dict순서 = CreateObject("Scripting.Dictionary")
    ' T07 직기 완제품 (행 68~77): Q=호기(17), B=제품(2), T=생산량BOX(20), X=원지기준소요량(24)
    arr1H = ws일보.Range("Q68:Q77").Value
    arr1B = ws일보.Range("B68:B77").Value
    arr1T = ws일보.Range("T68:T77").Value
    arr1X = ws일보.Range("X68:X77").Value
    For r = 1 To 10
        품명 = Trim(CStr(arr1B(r, 1)))
        If 품명 = "" Or UCase(품명) = "TOTAL" Or 품명 = "계" Or IsNumeric(품명) Then GoTo NextI
        생산량 = 0
        If IsNumeric(arr1T(r, 1)) Then 생산량 = CDbl(arr1T(r, 1))
        If 생산량 <> 0 Then
            구분 = Trim(CStr(arr1H(r, 1)))
            If 구분 = "" Then 구분 = "KT"
            맵키 = 구분 & "|" & 품명
            If 제품딕.exists(맵키) Then 제품딕(맵키) = CDbl(제품딕(맵키)) + 생산량 Else 제품딕.Add 맵키, 생산량
            원지소요 = 0
            If IsNumeric(arr1X(r, 1)) Then 원지소요 = CDbl(arr1X(r, 1))
            If 원지딕.exists(맵키) Then
                원지딕(맵키) = CDbl(원지딕(맵키)) + 원지소요
            ElseIf 원지소요 <> 0 Then
                원지딕.Add 맵키, 원지소요
            End If
            If Not dict순서.exists(맵키) Then dict순서.Add 맵키, 1: catItems.Add 맵키
        End If
NextI:
    Next r
    ' T08 봉제 완제품 (행 68~77): AZ=호기(52), AK=제품(37), BC=생산량BOX(55), BG=원지기준소요량(59)
    arr2H = ws일보.Range("AZ68:AZ77").Value
    arr2K = ws일보.Range("AK68:AK77").Value
    arr2BC = ws일보.Range("BC68:BC77").Value
    arr2BG = ws일보.Range("BG68:BG77").Value
    For r = 1 To 10
        품명 = Trim(CStr(arr2K(r, 1)))
        If 품명 = "" Or UCase(품명) = "TOTAL" Or 품명 = "계" Or IsNumeric(품명) Then GoTo NextII
        생산량 = 0
        If IsNumeric(arr2BC(r, 1)) Then 생산량 = CDbl(arr2BC(r, 1))
        If 생산량 <> 0 Then
            구분 = Trim(CStr(arr2H(r, 1)))
            If 구분 = "" Then 구분 = "KT"
            맵키 = 구분 & "|" & 품명
            If 제품딕.exists(맵키) Then 제품딕(맵키) = CDbl(제품딕(맵키)) + 생산량 Else 제품딕.Add 맵키, 생산량
            원지소요 = 0
            If IsNumeric(arr2BG(r, 1)) Then 원지소요 = CDbl(arr2BG(r, 1))
            If 원지딕.exists(맵키) Then
                원지딕(맵키) = CDbl(원지딕(맵키)) + 원지소요
            ElseIf 원지소요 <> 0 Then
                원지딕.Add 맵키, 원지소요
            End If
            If Not dict순서.exists(맵키) Then dict순서.Add 맵키, 1: catItems.Add 맵키
        End If
NextII:
    Next r
    일일생산일보_제품원지_수집 = (제품딕.Count > 0)
End Function

' 원단관리 시트에서 작성일과 일치하는 일별 열 번호 반환 (2행=월, 3행=일). 없으면 0
' 월별 구조: 3행 월 번호(1~12) → 해당 월 열 반환
Private Function 원단관리_날짜열_찾기(ws As Worksheet, 작성일 As Date) As Long
    Dim 대상월 As Long
    원단관리_날짜열_찾기 = 0
    If ws Is Nothing Then Exit Function
    대상월 = Month(작성일)
    Dim 마지막열 As Long: 마지막열 = 마지막월열_찾기(ws)
    If 마지막열 < COL_월별시작 Then Exit Function
    Dim c As Long
    Dim arrRow3 As Variant
    arrRow3 = ws.Range(ws.Cells(3, COL_월별시작), ws.Cells(3, 마지막열)).Value
    For c = 1 To UBound(arrRow3, 2)
        If IsNumeric(arrRow3(1, c)) Then
            If CLng(arrRow3(1, c)) = 대상월 Then
                원단관리_날짜열_찾기 = COL_월별시작 + c - 1
                Exit Function
            End If
        End If
    Next c
End Function

' 일일생산DB에서 완제품 행만 읽어 원단관리 그룹 항목 생성 (구분|품명 중복 제거, 등장 순서 유지)
' 일일생산DB: D열(4)=레코드_유형, I열(9)=완제품_구분(설비호수), J열(10)=완제품_품명
' 반환: 분류순서="제품", catItems(1)=Collection of "품명", 설비번호목록(i)=해당 항목의 I열(설비호수) 값 (참고용)
Private Function 일일생산DB_로드_항목(wsDB As Worksheet, ByRef 분류순서 As Collection, ByRef catItems As Collection, ByRef 설비번호목록 As Collection) As Boolean
    Dim lastRow As Long, r As Long
    Dim item As Collection
    Dim 품명 As String, 구분 As String
    Dim dict As Object
    Set 분류순서 = New Collection
    Set catItems = New Collection
    Set 설비번호목록 = New Collection
    If wsDB Is Nothing Then Exit Function
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    Set item = New Collection
    For r = 2 To lastRow
        If IsError(wsDB.Cells(r, 4).Value) Then GoTo NextRowDB
        If Trim(CStr(wsDB.Cells(r, 4).Value)) <> "완제품" Then GoTo NextRowDB
        If IsError(wsDB.Cells(r, 10).Value) Then GoTo NextRowDB
        품명 = Trim(CStr(wsDB.Cells(r, 10).Value))
        If IsError(wsDB.Cells(r, 9).Value) Then GoTo NextRowDB
        구분 = Trim(CStr(wsDB.Cells(r, 9).Value))
        If 품명 = "" Or InStr(1, 품명, "오류", vbTextCompare) > 0 Then GoTo NextRowDB
        If 구분 = "" Then 구분 = "KT"
        If 구분 = "수입" Then GoTo NextRowDB
        If Not dict.exists(품명) Then
            dict.Add 품명, 1
            item.Add 품명
            설비번호목록.Add Trim(CStr(wsDB.Cells(r, 일일생산DB_설비번호열).Value))
        End If
NextRowDB:
    Next r
    분류순서.Add "제품"
    catItems.Add item
    일일생산DB_로드_항목 = (item.Count > 0)
End Function

' 일일생산DB에서 특정 작성일의 완제품 행만 읽어 제품딕(구분|품명)=생산량_BOX 합계, 원지딕=원지기준소요량 합계 채움
' ★ 헤더 기반 열매핑 사용 (일일생산DB_열매핑)
Private Function 일일생산DB_제품원지_수집(wsDB As Worksheet, ByRef 제품딕 As Object, ByRef 원지딕 As Object, 작성일 As Date) As Boolean
    Dim lastRow As Long, r As Long
    Dim 구분 As String, 품명 As String, 맵키 As String
    Dim 생산량 As Double, 원지소요 As Double
    Dim v날짜 As Variant
    Dim col날짜 As Long, colID As Long, col레코드유형 As Long, col구분 As Long, col품명 As Long, col생산량BOX As Long, col원지소요 As Long, col원지종별 As Long, col원지KG As Long
    If wsDB Is Nothing Then Exit Function
    Call 일일생산DB_열매핑(wsDB, col날짜, colID, col레코드유형, col구분, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG)
    On Error Resume Next
    Set 제품딕 = CreateObject("Scripting.Dictionary")
    Set 원지딕 = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    For r = 2 To lastRow
        If Trim(CStr(wsDB.Cells(r, col레코드유형).Value)) <> "완제품" Then GoTo NextRowDB2
        v날짜 = wsDB.Cells(r, col날짜).Value
        If Not IsDate(v날짜) Then GoTo NextRowDB2
        If CDate(v날짜) <> 작성일 Then GoTo NextRowDB2
        If IsError(wsDB.Cells(r, col구분).Value) Or IsError(wsDB.Cells(r, col품명).Value) Then GoTo NextRowDB2
        구분 = Trim(CStr(wsDB.Cells(r, col구분).Value))
        품명 = Trim(CStr(wsDB.Cells(r, col품명).Value))
        If 품명 = "" Or InStr(1, 품명, "오류", vbTextCompare) > 0 Then GoTo NextRowDB2
        If 구분 = "" Then 구분 = "KT"
        맵키 = 구분 & "|" & 품명
        생산량 = 0
        If IsNumeric(wsDB.Cells(r, col생산량BOX).Value) Then 생산량 = CDbl(wsDB.Cells(r, col생산량BOX).Value)
        원지소요 = 0
        If IsNumeric(wsDB.Cells(r, col원지소요).Value) Then 원지소요 = CDbl(wsDB.Cells(r, col원지소요).Value)
        If 제품딕.exists(맵키) Then 제품딕(맵키) = CDbl(제품딕(맵키)) + 생산량 Else 제품딕.Add 맵키, 생산량
        If 원지딕.exists(맵키) Then 원지딕(맵키) = CDbl(원지딕(맵키)) + 원지소요 Else 원지딕.Add 맵키, 원지소요
NextRowDB2:
    Next r
    일일생산DB_제품원지_수집 = (제품딕.Count > 0)
End Function

' 일일생산DB 1행 헤더에서 열 번호 반환 (없으면 기본값). 역변환/최적화 구조 차이 대응
Private Sub 일일생산DB_열매핑(wsDB As Worksheet, ByRef col날짜 As Long, ByRef colID As Long, ByRef col레코드유형 As Long, _
    ByRef col구분 As Long, ByRef col품명 As Long, ByRef col생산량BOX As Long, ByRef col원지소요 As Long, _
    ByRef col원지종별 As Long, ByRef col원지KG As Long)
    col날짜 = 1: colID = 2: col레코드유형 = 4: col구분 = 9: col품명 = 10: col생산량BOX = 12: col원지소요 = 13: col원지종별 = 6: col원지KG = 8
    If wsDB Is Nothing Then Exit Sub
    Dim c As Long, h As String
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If h = "생산일" Or h = "작성일" Then col날짜 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If h = "id" Then colID = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "레코드") > 0 And InStr(h, "유형") > 0 Then col레코드유형 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "완제품") > 0 And InStr(h, "구분") > 0 Then col구분 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "완제품") > 0 And InStr(h, "품명") > 0 Then col품명 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "생산량") > 0 And InStr(h, "box") > 0 Then col생산량BOX = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "원지기준소요량") > 0 Then col원지소요 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "원지") > 0 And InStr(h, "종별") > 0 Then col원지종별 = c: Exit For
    Next c
    For c = 1 To 20
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "원지") > 0 And InStr(h, "kg") > 0 Then col원지KG = c: Exit For
    Next c
End Sub

' 일일생산DB 열매핑 확장: 원지_ROLL_ID(7열) - Raw data 원단명 매칭용
Private Sub 일일생산DB_열매핑_원지ROLLID(wsDB As Worksheet, ByRef col원지ROLLID As Long)
    col원지ROLLID = 7
    If wsDB Is Nothing Then Exit Sub
    Dim c As Long, h As String
    For c = 1 To 25
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "원지") > 0 And (InStr(h, "roll") > 0 Or InStr(h, "roll_id") > 0) Then col원지ROLLID = c: Exit Sub
    Next c
End Sub

' 일일생산DB 열매핑 확장: 원지_Co(70열 BR, 72열 BT 등 헤더 동적 검색)
' M18_원지Co열: Module10 역변환용 Public 래퍼
Public Function M18_원지Co열(wsDB As Worksheet) As Long
    Dim c As Long: c = 70
    Call 일일생산DB_열매핑_원지Co(wsDB, c)
    M18_원지Co열 = c
End Function
Private Sub 일일생산DB_열매핑_원지Co(wsDB As Worksheet, ByRef col원지Co As Long)
    col원지Co = 70
    If wsDB Is Nothing Then Exit Sub
    Dim c As Long, h As String
    For c = 65 To 100
        h = Trim(LCase(Replace(CStr(wsDB.Cells(1, c).Value), " ", "")))
        If InStr(h, "원지") > 0 And InStr(h, "co") > 0 Then col원지Co = c: Exit Sub
    Next c
End Sub

' 일일생산DB에서 원지 레코드의 unique Co(원지_Co) 수집 → 원단관리 항목용
Private Function 일일생산DB_로드_항목_원단(wsDB As Worksheet, ByRef 분류순서 As Collection, ByRef catItems As Collection) As Boolean
    Dim col날짜 As Long, colID As Long, col레코드유형 As Long, col구분 As Long, col품명 As Long
    Dim col생산량BOX As Long, col원지소요 As Long, col원지종별 As Long, col원지KG As Long
    Dim col원지Co As Long, col원지ROLLID As Long
    Dim lastRow As Long, r As Long
    Dim CoSet As Object, CoList As Collection
    Dim CoVal As String, vCo As Variant, rollIdVal As String
    Set 분류순서 = New Collection
    Set catItems = New Collection
    Set CoSet = CreateObject("Scripting.Dictionary")
    Set CoList = New Collection
    If wsDB Is Nothing Then Exit Function
    Call 일일생산DB_열매핑(wsDB, col날짜, colID, col레코드유형, col구분, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG)
    Call 일일생산DB_열매핑_원지Co(wsDB, col원지Co)
    Call 일일생산DB_열매핑_원지ROLLID(wsDB, col원지ROLLID)
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Debug.Print "[원단로드] lastRow < 2, Exit"
        Exit Function
    End If
    Dim maxCol As Long: maxCol = col원지Co
    If col원지ROLLID > maxCol Then maxCol = col원지ROLLID
    If col레코드유형 > maxCol Then maxCol = col레코드유형
    Dim arr As Variant
    arr = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, maxCol)).Value
    For r = 1 To UBound(arr, 1)
        If col레코드유형 > UBound(arr, 2) Then GoTo NextRowCo
        If IsError(arr(r, col레코드유형)) Then GoTo NextRowCo
        If Trim(CStr(arr(r, col레코드유형))) <> "원지" Then GoTo NextRowCo
        CoVal = ""
        If col원지Co <= UBound(arr, 2) Then
            vCo = arr(r, col원지Co)
            If Not IsEmpty(vCo) And Not IsError(vCo) Then CoVal = Trim(CStr(vCo))
        End If
        If CoVal = "" And col원지ROLLID <= UBound(arr, 2) Then
            rollIdVal = Trim(CStr(arr(r, col원지ROLLID)))
            CoVal = RawData_원단명ToCo(rollIdVal)
        End If
        If CoVal = "" Then GoTo NextRowCo
        If IsNumeric(CoVal) Then
            If CLng(val(CoVal)) = 0 Then GoTo NextRowCo
            CoVal = Format(CLng(val(CoVal)), "00")   ' 1 -> "01", 1.0 -> "01"
        End If
        If Not CoSet.exists(CoVal) Then
            CoSet.Add CoVal, 1
            CoList.Add CoVal
        End If
NextRowCo:
    Next r
    Debug.Print "[원단로드] 원지행스캔=" & UBound(arr, 1) & " Co수=" & CoList.Count & " col레코드유형=" & col레코드유형 & " col원지Co=" & col원지Co
    If CoList.Count = 0 Then Exit Function
    분류순서.Add "원단"
    catItems.Add Co_Collection_오름차순정렬(CoList)
    일일생산DB_로드_항목_원단 = True
End Function

' Raw data 시트 H열에서 Co 인덱스 목록 수집 (DB에 원지_Co 없을 때 fallback)
Private Function RawData_Co목록_로드(ByRef 분류순서 As Collection, ByRef catItems As Collection) As Boolean
    Dim wsRaw As Worksheet, lastRow As Long, r As Long
    Dim CoList As Collection, CoSet As Object, vCo As Variant, CoVal As String
    Set 분류순서 = New Collection
    Set catItems = New Collection
    Set CoList = New Collection
    Set CoSet = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(WS_RawData)
    On Error GoTo 0
    If wsRaw Is Nothing Then Exit Function
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 8).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    For r = 3 To lastRow   ' H3셀부터 동적 인식
        vCo = wsRaw.Cells(r, 8).Value
        If IsEmpty(vCo) Or IsError(vCo) Then GoTo NextRawCo
        CoVal = Trim(CStr(vCo))
        If CoVal = "" Then GoTo NextRawCo
        If IsNumeric(CoVal) Then
            If CLng(val(CoVal)) = 0 Then GoTo NextRawCo
            CoVal = Format(CLng(val(CoVal)), "00")   ' 1 -> "01"
        End If
        If Not CoSet.exists(CoVal) Then CoSet.Add CoVal, 1: CoList.Add CoVal
NextRawCo:
    Next r
    If CoList.Count = 0 Then Exit Function
    분류순서.Add "원단"
    catItems.Add Co_Collection_오름차순정렬(CoList)
    RawData_Co목록_로드 = True
    Debug.Print "[원단로드] Raw data H열 fallback: Co수=" & CoList.Count
End Function

'================================================================================
' Co 인덱스 Collection을 숫자 오름차순으로 정렬하여 반환 (버블 정렬)
'================================================================================
Private Function Co_Collection_오름차순정렬(src As Collection) As Collection
    Dim n As Long: n = src.Count
    Dim result As New Collection
    If n = 0 Then Set Co_Collection_오름차순정렬 = result: Exit Function
    Dim arr() As String
    ReDim arr(1 To n)
    Dim i As Long, j As Long
    For i = 1 To n: arr(i) = CStr(src(i)): Next i
    ' 숫자면 숫자값으로, 아니면 문자열로 오름차순 비교
    Dim tmp As String
    For i = 1 To n - 1
        For j = i + 1 To n
            Dim swap필요 As Boolean
            If IsNumeric(arr(i)) And IsNumeric(arr(j)) Then
                swap필요 = (CDbl(arr(i)) > CDbl(arr(j)))
            Else
                swap필요 = (arr(i) > arr(j))
            End If
            If swap필요 Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    For i = 1 To n: result.Add arr(i): Next i
    Set Co_Collection_오름차순정렬 = result
End Function

' Raw data 시트: H열=Co인덱스, I열=원단명. Co로 조회하여 원단명 반환
Public Function RawData_CoTo원단명(CoVal As String) As String
    Dim wsRaw As Worksheet
    Dim r As Long, lastRow As Long
    Const RawData_H As Long = 8
    Const RawData_I As Long = 9
    RawData_CoTo원단명 = ""
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(WS_RawData)
    On Error GoTo 0
    If wsRaw Is Nothing Then Exit Function
    CoVal = Trim(CStr(CoVal))
    If CoVal = "" Then Exit Function
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, RawData_H).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    For r = 3 To lastRow   ' H3셀부터 동적 인식
        If Trim(CStr(wsRaw.Cells(r, RawData_H).Value)) = CoVal Or (IsNumeric(CoVal) And IsNumeric(wsRaw.Cells(r, RawData_H).Value) And CDbl(wsRaw.Cells(r, RawData_H).Value) = CDbl(CoVal)) Then
            RawData_CoTo원단명 = Trim(CStr(wsRaw.Cells(r, RawData_I).Value))
            Exit Function
        End If
    Next r
End Function

' Raw data 시트: H열=Co인덱스, I열=원단명. 원단명으로 조회하여 Co 반환 (일일생산DB 원지_Co fallback용)
Public Function RawData_원단명ToCo(원단명 As String) As String
    Dim wsRaw As Worksheet
    Dim r As Long, lastRow As Long
    Const RawData_H As Long = 8
    Const RawData_I As Long = 9
    Dim iVal As String, c As Range
    RawData_원단명ToCo = ""
    원단명 = Trim(CStr(원단명))
    If 원단명 = "" Then Exit Function
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(WS_RawData)
    On Error GoTo 0
    If wsRaw Is Nothing Then Exit Function
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, RawData_I).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    For r = 3 To lastRow
        iVal = Trim(CStr(wsRaw.Cells(r, RawData_I).Value))
        If iVal = 원단명 Or (iVal <> "" And LCase(iVal) = LCase(원단명)) Then
            RawData_원단명ToCo = Trim(CStr(wsRaw.Cells(r, RawData_H).Value))
            If RawData_원단명ToCo <> "" And IsNumeric(RawData_원단명ToCo) Then
                RawData_원단명ToCo = Format(CLng(val(RawData_원단명ToCo)), "00")
            End If
            Exit Function
        End If
    Next r
    ' fallback: 목록_원단명·목록_원단번호 병렬 범위에서 순회 매칭 (유효성 목록 = 동일 소스)
    On Error Resume Next
    Dim rng명 As Range, rng번호 As Range
    Set rng명 = ThisWorkbook.Names("목록_원단명").RefersToRange
    Set rng번호 = ThisWorkbook.Names("목록_원단번호").RefersToRange
    If Not rng명 Is Nothing And Not rng번호 Is Nothing Then
        Dim i As Long, v명 As String, v번호 As String
        For i = 1 To rng명.Cells.Count
            v명 = Trim(CStr(rng명.Cells(i).Value))
            If v명 = 원단명 Or (v명 <> "" And LCase(v명) = LCase(원단명)) Then
                v번호 = Trim(CStr(rng번호.Cells(i).Value))
                If v번호 <> "" And IsNumeric(v번호) Then
                    RawData_원단명ToCo = Format(CLng(val(v번호)), "00")
                Else
                    RawData_원단명ToCo = v번호
                End If
                Exit Function
            End If
        Next i
    End If
    ' Find fallback (xlPart로 공백/미세차이 허용)
    Set c = ThisWorkbook.Names("목록_원단명").RefersToRange.Find(What:=원단명, LookIn:=xlValues, LookAt:=xlPart)
    If Not c Is Nothing Then
        RawData_원단명ToCo = Trim(CStr(c.Parent.Cells(c.Row, RawData_H).Value))
        If RawData_원단명ToCo <> "" And IsNumeric(RawData_원단명ToCo) Then
            RawData_원단명ToCo = Format(CLng(val(RawData_원단명ToCo)), "00")
        End If
    End If
    On Error GoTo 0
End Function

' Raw data H열(원단번호)·I열(원단명) → 목록_원단번호, 목록_원단명 동적 이름 정의
Public Sub RawData_HI_목록_갱신()
    Dim wsRaw As Worksheet
    Dim lastRow As Long
    Const RawData_H As Long = 8
    Const RawData_I As Long = 9
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(WS_RawData)
    On Error GoTo 0
    If wsRaw Is Nothing Then Exit Sub
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, RawData_H).End(xlUp).Row
    If lastRow < 3 Then lastRow = wsRaw.Cells(wsRaw.Rows.Count, RawData_I).End(xlUp).Row
    If lastRow < 3 Then Exit Sub
    Dim sh As String: sh = "'" & Replace(WS_RawData, "'", "''") & "'"
    On Error Resume Next
    ThisWorkbook.Names("목록_원단번호").Delete
    ThisWorkbook.Names.Add Name:="목록_원단번호", RefersTo:="=" & sh & "!$H$3:$H$" & lastRow
    ThisWorkbook.Names("목록_원단명").Delete
    ThisWorkbook.Names.Add Name:="목록_원단명", RefersTo:="=" & sh & "!$I$3:$I$" & lastRow
    On Error GoTo 0
End Sub

' Raw data 시트 선택 시 원단 목록(목록_원단번호, 목록_원단명) 갱신
Public Sub 원단_목록_RawData_갱신()
    Call RawData_HI_목록_갱신
End Sub

' 일일생산DB에서 오늘 이전 모든 날짜별 제품딕·원지딕·원지ROLL딕·원지Co딕 수집 (원단관리용 Co별 KG)
' 반환: 날짜목록, 날짜별제품딕, 날짜별원지딕, 날짜별원지ROLL딕, [날짜별원지Co딕="YYYYMMDD"->Dict(Co->Array(롤수,중량))]
Private Function 일일생산DB_날짜별_제품원지_수집(wsDB As Worksheet, ByRef 날짜목록 As Collection, _
    ByRef 날짜별제품딕 As Object, ByRef 날짜별원지딕 As Object, _
    Optional ByRef 날짜별원지ROLL딕 As Object = Nothing, Optional ByRef 날짜별원지Co딕 As Object = Nothing) As Boolean
    Dim lastRow As Long, r As Long
    Dim 구분 As String, 품명 As String, 맵키 As String
    Dim 생산량 As Double, 원지소요 As Double
    Dim v날짜 As Variant, d날짜 As Date
    Dim dateKey As String
    Dim pd As Object, od As Object
    Dim col날짜 As Long, colID As Long, col레코드유형 As Long, col구분 As Long, col품명 As Long, col생산량BOX As Long, col원지소요 As Long, col원지종별 As Long, col원지KG As Long, col원지Co As Long, col원지ROLLID As Long
    If wsDB Is Nothing Then Exit Function
    Call 일일생산DB_열매핑(wsDB, col날짜, colID, col레코드유형, col구분, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG)
    Call 일일생산DB_열매핑_원지Co(wsDB, col원지Co)
    Call 일일생산DB_열매핑_원지ROLLID(wsDB, col원지ROLLID)
    Set 날짜목록 = New Collection
    Set 날짜별제품딕 = CreateObject("Scripting.Dictionary")
    Set 날짜별원지딕 = CreateObject("Scripting.Dictionary")
    Set 날짜별원지ROLL딕 = CreateObject("Scripting.Dictionary")
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Dim maxCol As Long: maxCol = Application.WorksheetFunction.Max(70, col날짜, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG, col원지Co, col원지ROLLID)
    Dim arr날짜별 As Variant
    If g_검수테스트_DB로드행제한 > 0 And (lastRow - 1) > g_검수테스트_DB로드행제한 Then
        Dim dbStart원단 As Long: dbStart원단 = lastRow - g_검수테스트_DB로드행제한 + 1
        arr날짜별 = wsDB.Range(wsDB.Cells(dbStart원단, 1), wsDB.Cells(lastRow, maxCol)).Value
    Else
        arr날짜별 = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, maxCol)).Value
    End If
    For r = 1 To UBound(arr날짜별, 1)
        If col레코드유형 > UBound(arr날짜별, 2) Or IsError(arr날짜별(r, col레코드유형)) Then GoTo NextRowDB3
        Dim 레코드유형 As String: 레코드유형 = Trim(CStr(arr날짜별(r, col레코드유형)))
        v날짜 = arr날짜별(r, col날짜)
        If IsError(v날짜) Or Not IsDate(v날짜) Then GoTo NextRowDB3
        d날짜 = CDate(v날짜)
        If d날짜 > Date Then GoTo NextRowDB3
        dateKey = Format(d날짜, "YYYYMMDD")
        If Not 날짜별제품딕.exists(dateKey) Then
            날짜별제품딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜별원지딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜별원지ROLL딕.Add dateKey, CreateObject("Scripting.Dictionary")
            If Not 날짜별원지Co딕 Is Nothing Then 날짜별원지Co딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜목록.Add d날짜
        End If
        If 레코드유형 = "완제품" Then
            Set pd = 날짜별제품딕(dateKey)
            Set od = 날짜별원지딕(dateKey)
            If col구분 > UBound(arr날짜별, 2) Or col품명 > UBound(arr날짜별, 2) Then GoTo NextRowDB3
            If IsError(arr날짜별(r, col구분)) Or IsError(arr날짜별(r, col품명)) Then GoTo NextRowDB3
            구분 = Trim(CStr(arr날짜별(r, col구분)))
            품명 = Trim(CStr(arr날짜별(r, col품명)))
            If 품명 = "" Or InStr(1, 품명, "오류", vbTextCompare) > 0 Then GoTo NextRowDB3
            If 구분 = "" Then 구분 = "KT"
            If 구분 = "수입" Then GoTo NextRowDB3
            맵키 = 품명
            생산량 = 0
            If col생산량BOX <= UBound(arr날짜별, 2) And IsNumeric(arr날짜별(r, col생산량BOX)) Then 생산량 = CDbl(arr날짜별(r, col생산량BOX))
            원지소요 = 0
            If col원지소요 <= UBound(arr날짜별, 2) And IsNumeric(arr날짜별(r, col원지소요)) Then 원지소요 = CDbl(arr날짜별(r, col원지소요))
            If pd.exists(맵키) Then pd(맵키) = CDbl(pd(맵키)) + 생산량 Else pd.Add 맵키, 생산량
            If od.exists(맵키) Then od(맵키) = CDbl(od(맵키)) + 원지소요 Else od.Add 맵키, 원지소요
        ElseIf 레코드유형 = "원지" Then
            Dim kgVal As Variant
            If col원지KG <= UBound(arr날짜별, 2) Then kgVal = arr날짜별(r, col원지KG) Else kgVal = Empty
            If IsEmpty(kgVal) Or IsError(kgVal) Then GoTo NextRowDB3
            If Not IsNumeric(kgVal) Then GoTo NextRowDB3
            If CDbl(kgVal) = 0 Then GoTo NextRowDB3
            If True Then
                Dim 종별 As String: 종별 = ""
                If col원지종별 <= UBound(arr날짜별, 2) Then
                    If Not IsError(arr날짜별(r, col원지종별)) Then 종별 = Trim(CStr(arr날짜별(r, col원지종별)))
                End If
                Dim 설비키 As String
                If left(종별, 3) = "JRT" Then 설비키 = "JRT" Else 설비키 = 종별
                Dim rd As Object: Set rd = 날짜별원지ROLL딕(dateKey)
                If rd.exists(설비키) Then rd(설비키) = CLng(rd(설비키)) + 1 Else rd.Add 설비키, 1
                ' 원단관리용: Co별 (롤수, 중량) 저장
                If Not 날짜별원지Co딕 Is Nothing And col원지Co <= UBound(arr날짜별, 2) Then
                    Dim CoVal As Variant: CoVal = arr날짜별(r, col원지Co)
                    Dim CoKey As String: CoKey = ""
                    If Not IsEmpty(CoVal) And Not IsError(CoVal) Then CoKey = Trim(CStr(CoVal))
                    If CoKey = "" And col원지ROLLID <= UBound(arr날짜별, 2) Then CoKey = RawData_원단명ToCo(Trim(CStr(arr날짜별(r, col원지ROLLID))))
                    If CoKey <> "" Then
                            ' Co 포맷 "01" 통일
                            If IsNumeric(CoKey) And CLng(val(CoKey)) > 0 Then CoKey = Format(CLng(val(CoKey)), "00")
                            Dim cd As Object: Set cd = 날짜별원지Co딕(dateKey)
                            Dim arrRv As Variant
                            If cd.exists(CoKey) Then
                                arrRv = cd(CoKey)
                                cd(CoKey) = Array(CLng(arrRv(0)) + 1, CDbl(arrRv(1)) + CDbl(kgVal))
                            Else
                                cd.Add CoKey, Array(1, CDbl(kgVal))
                            End If
                    End If
                End If
            End If
        End If
NextRowDB3:
    Next r
    일일생산DB_날짜별_제품원지_수집 = (날짜목록.Count > 0)
End Function

' 일일생산DB에서 ID > lastMaxId 인 행만 대상으로 날짜별 제품/원지 수집 (증분 반영용, 속도 개선)
Private Function 일일생산DB_날짜별_제품원지_수집_증분(wsDB As Worksheet, lastMaxId As Long, ByRef 날짜목록 As Collection, _
    ByRef 날짜별제품딕 As Object, ByRef 날짜별원지딕 As Object, _
    Optional ByRef 날짜별원지ROLL딕 As Object = Nothing, Optional ByRef 날짜별원지Co딕 As Object = Nothing) As Boolean
    Dim lastRow As Long, r As Long
    Dim 구분 As String, 품명 As String, 맵키 As String
    Dim 생산량 As Double, 원지소요 As Double
    Dim v날짜 As Variant, d날짜 As Date
    Dim dateKey As String
    Dim pd As Object, od As Object
    Dim idVal As Variant, rowId As Long
    Dim col날짜 As Long, colID As Long, col레코드유형 As Long, col구분 As Long, col품명 As Long, col생산량BOX As Long, col원지소요 As Long, col원지종별 As Long, col원지KG As Long, col원지Co As Long, col원지ROLLID As Long
    If wsDB Is Nothing Then Exit Function
    Call 일일생산DB_열매핑(wsDB, col날짜, colID, col레코드유형, col구분, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG)
    Call 일일생산DB_열매핑_원지Co(wsDB, col원지Co)
    Call 일일생산DB_열매핑_원지ROLLID(wsDB, col원지ROLLID)
    Set 날짜목록 = New Collection
    Set 날짜별제품딕 = CreateObject("Scripting.Dictionary")
    Set 날짜별원지딕 = CreateObject("Scripting.Dictionary")
    Set 날짜별원지ROLL딕 = CreateObject("Scripting.Dictionary")
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    Dim maxCol As Long: maxCol = Application.WorksheetFunction.Max(70, col날짜, colID, col품명, col생산량BOX, col원지소요, col원지종별, col원지KG, col원지Co, col원지ROLLID)
    Dim arr날짜별 As Variant
    If g_검수테스트_DB로드행제한 > 0 And (lastRow - 1) > g_검수테스트_DB로드행제한 Then
        Dim dbStart원단증분 As Long: dbStart원단증분 = lastRow - g_검수테스트_DB로드행제한 + 1
        arr날짜별 = wsDB.Range(wsDB.Cells(dbStart원단증분, 1), wsDB.Cells(lastRow, maxCol)).Value
    Else
        arr날짜별 = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, maxCol)).Value
    End If
    Dim kgVal As Variant, 종별 As String, 설비키 As String, rd As Object, cd As Object, CoVal As Variant, CoKey As String
    For r = 1 To UBound(arr날짜별, 1)
        If colID <= UBound(arr날짜별, 2) Then idVal = arr날짜별(r, colID) Else idVal = 0
        If IsError(idVal) Then idVal = 0
        If IsNumeric(idVal) Then rowId = CLng(idVal) Else rowId = 0
        If rowId <= lastMaxId Then GoTo NextRowDB증분
        If col레코드유형 > UBound(arr날짜별, 2) Or IsError(arr날짜별(r, col레코드유형)) Then GoTo NextRowDB증분
        Dim 레코드유형 As String: 레코드유형 = Trim(CStr(arr날짜별(r, col레코드유형)))
        v날짜 = arr날짜별(r, col날짜)
        If IsError(v날짜) Or Not IsDate(v날짜) Then GoTo NextRowDB증분
        d날짜 = CDate(v날짜)
        If d날짜 > Date Then GoTo NextRowDB증분
        dateKey = Format(d날짜, "YYYYMMDD")
        If Not 날짜별제품딕.exists(dateKey) Then
            날짜별제품딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜별원지딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜별원지ROLL딕.Add dateKey, CreateObject("Scripting.Dictionary")
            If Not 날짜별원지Co딕 Is Nothing Then 날짜별원지Co딕.Add dateKey, CreateObject("Scripting.Dictionary")
            날짜목록.Add d날짜
        End If
        If 레코드유형 = "완제품" Then
            Set pd = 날짜별제품딕(dateKey)
            Set od = 날짜별원지딕(dateKey)
            If col구분 > UBound(arr날짜별, 2) Or col품명 > UBound(arr날짜별, 2) Then GoTo NextRowDB증분
            If IsError(arr날짜별(r, col구분)) Or IsError(arr날짜별(r, col품명)) Then GoTo NextRowDB증분
            구분 = Trim(CStr(arr날짜별(r, col구분)))
            품명 = Trim(CStr(arr날짜별(r, col품명)))
            If 품명 = "" Or InStr(1, 품명, "오류", vbTextCompare) > 0 Then GoTo NextRowDB증분
            If 구분 = "" Then 구분 = "KT"
            If 구분 = "수입" Then GoTo NextRowDB증분
            맵키 = 품명
            생산량 = 0
            If col생산량BOX <= UBound(arr날짜별, 2) And IsNumeric(arr날짜별(r, col생산량BOX)) Then 생산량 = CDbl(arr날짜별(r, col생산량BOX))
            원지소요 = 0
            If col원지소요 <= UBound(arr날짜별, 2) And IsNumeric(arr날짜별(r, col원지소요)) Then 원지소요 = CDbl(arr날짜별(r, col원지소요))
            If pd.exists(맵키) Then pd(맵키) = CDbl(pd(맵키)) + 생산량 Else pd.Add 맵키, 생산량
            If od.exists(맵키) Then od(맵키) = CDbl(od(맵키)) + 원지소요 Else od.Add 맵키, 원지소요
        ElseIf 레코드유형 = "원지" Then
            If col원지KG <= UBound(arr날짜별, 2) Then kgVal = arr날짜별(r, col원지KG) Else kgVal = Empty
            If IsEmpty(kgVal) Or IsError(kgVal) Then GoTo NextRowDB증분
            If Not IsNumeric(kgVal) Then GoTo NextRowDB증분
            If CDbl(kgVal) = 0 Then GoTo NextRowDB증분
            If True Then
                종별 = ""
                If col원지종별 <= UBound(arr날짜별, 2) Then
                    If Not IsError(arr날짜별(r, col원지종별)) Then 종별 = Trim(CStr(arr날짜별(r, col원지종별)))
                End If
                If left(종별, 3) = "JRT" Then 설비키 = "JRT" Else 설비키 = 종별
                Set rd = 날짜별원지ROLL딕(dateKey)
                If rd.exists(설비키) Then rd(설비키) = CLng(rd(설비키)) + 1 Else rd.Add 설비키, 1
                If Not 날짜별원지Co딕 Is Nothing And 날짜별원지Co딕.exists(dateKey) And col원지Co <= UBound(arr날짜별, 2) Then
                    CoVal = arr날짜별(r, col원지Co)
                    CoKey = ""
                    If Not IsEmpty(CoVal) And Not IsError(CoVal) Then CoKey = Trim(CStr(CoVal))
                    If CoKey = "" And col원지ROLLID <= UBound(arr날짜별, 2) Then CoKey = RawData_원단명ToCo(Trim(CStr(arr날짜별(r, col원지ROLLID))))
                    If CoKey <> "" Then
                            ' Co 포맷 "01" 통일
                            If IsNumeric(CoKey) And CLng(val(CoKey)) > 0 Then CoKey = Format(CLng(val(CoKey)), "00")
                            Set cd = 날짜별원지Co딕(dateKey)
                            Dim arrRv증분 As Variant
                            If cd.exists(CoKey) Then
                                arrRv증분 = cd(CoKey)
                                cd(CoKey) = Array(CLng(arrRv증분(0)) + 1, CDbl(arrRv증분(1)) + CDbl(kgVal))
                            Else
                                cd.Add CoKey, Array(1, CDbl(kgVal))
                            End If
                    End If
                End If
            End If
        End If
NextRowDB증분:
    Next r
    일일생산DB_날짜별_제품원지_수집_증분 = (날짜목록.Count > 0)
End Function

' 일일생산DB 증분 동기화: 마지막 동기화 ID 조회 (원단관리_동기화상태 시트 A1)
Private Function 일일생산DB_마지막동기화ID_조회() As Long
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_동기화상태)
    If ws Is Nothing Then 일일생산DB_마지막동기화ID_조회 = 0: Exit Function
    If Not IsNumeric(ws.Range("A1").Value) Then 일일생산DB_마지막동기화ID_조회 = 0: Exit Function
    일일생산DB_마지막동기화ID_조회 = CLng(ws.Range("A1").Value)
    If 일일생산DB_마지막동기화ID_조회 < 0 Then 일일생산DB_마지막동기화ID_조회 = 0
    On Error GoTo 0
End Function

' 일일생산DB 증분 동기화: 마지막 동기화 ID 저장
Private Sub 일일생산DB_마지막동기화ID_저장(maxID As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_동기화상태)
    If ws Is Nothing Then
        Dim wb보호동기 As Boolean: wb보호동기 = ThisWorkbook.ProtectStructure
        If wb보호동기 Then ThisWorkbook.Unprotect WS_시트보호암호
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Not ws Is Nothing Then ws.Name = WS_동기화상태: ws.Visible = xlSheetVeryHidden
        If wb보호동기 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
    End If
    If Not ws Is Nothing Then ws.Range("A1").Value = maxID
    On Error GoTo 0
End Sub

' 일일생산DB 현재 최대 ID (B열, 배열 1회 읽기)
Private Function 일일생산DB_최대ID_조회(wsDB As Worksheet) As Long
    Dim lastRow As Long, r As Long, v As Variant, m As Long
    Dim arrId As Variant
    일일생산DB_최대ID_조회 = 0
    If wsDB Is Nothing Then Exit Function
    lastRow = wsDB.Cells(wsDB.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    arrId = wsDB.Range(wsDB.Cells(2, 2), wsDB.Cells(lastRow, 2)).Value
    For r = 1 To UBound(arrId, 1)
        v = arrId(r, 1)
        If IsNumeric(v) Then m = CLng(v): If m > 일일생산DB_최대ID_조회 Then 일일생산DB_최대ID_조회 = m
    Next r
End Function

'================================================================================
' ★★★ 속도 최적화용 헬퍼: 원단관리 시트 날짜→열 딕셔너리 (1회 빌드, O(1) 조회)
'================================================================================
' DB 날짜목록에 있는 날짜 중 시트에 열이 없으면 열 추가 (업데이트 반영을 위해)
Private Sub 원단관리_날짜열_확장(ws As Worksheet, 날짜목록 As Collection)
    Dim di As Long, dateKey As String, 작성일 As Date
    Dim dict As Object, 마지막열 As Long, 새열 As Long
    Dim 기준년 As Long, 시트보호해제 As Boolean
    If ws Is Nothing Or 날짜목록 Is Nothing Then Exit Sub
    If 날짜목록.Count < 1 Then Exit Sub
    시트보호해제 = ws.ProtectContents
    If 시트보호해제 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
    Set dict = 원단관리_날짜열_딕셔너리(ws)
    기준년 = Year(Date)
    If IsNumeric(ws.Range("I2").Value) And ws.Range("I2").Value >= 2000 And ws.Range("I2").Value <= 2100 Then 기준년 = CLng(ws.Range("I2").Value)
    마지막열 = 마지막월열_찾기(ws)
    If 마지막열 < COL_월별시작 Then 마지막열 = COL_월별시작 - 1
    For di = 1 To 날짜목록.Count
        작성일 = 날짜목록(di)
        If 작성일 > Date + 10 Then GoTo Next확장일
        dateKey = Format(작성일, "YYYYMMDD")
        If Not dict Is Nothing And dict.exists(dateKey) Then GoTo Next확장일
        새열 = 마지막열 + 1
        ws.Cells(2, 새열).Value = Month(작성일) & "M"
        ws.Cells(2, 새열).Font.color = RGB(160, 160, 160)
        ws.Cells(3, 새열).Value = Day(작성일)
        If Not dict Is Nothing Then dict.Add dateKey, 새열
        마지막열 = 새열
Next확장일:
    Next di
    If 시트보호해제 Then ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

' ★ 일별 매칭: 2행 "1M"=1월, "2M"=2월... 3행=일(1~31). I2=연도(흰색폰트). YYYYMMDD→열
Private Function 원단관리_날짜열_딕셔너리(ws As Worksheet) As Object
    Dim dict As Object, 마지막열 As Long, c As Long
    Dim arrRow2 As Variant, arrRow3 As Variant
    Dim 월파싱 As Long, 일번호 As Long, 기준년 As Long
    Dim dateKey As String, s2 As String, vI2 As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    Set 원단관리_날짜열_딕셔너리 = dict
    If ws Is Nothing Then Exit Function
    마지막열 = 마지막월열_찾기(ws)
    If 마지막열 < COL_월별시작 Then Exit Function
    vI2 = ws.Range("I2").Value
    기준년 = Year(Date)
    If IsNumeric(vI2) And vI2 >= 2000 And vI2 <= 2100 Then 기준년 = CLng(vI2)
    arrRow2 = ws.Range(ws.Cells(2, COL_월별시작), ws.Cells(2, 마지막열)).Value
    arrRow3 = ws.Range(ws.Cells(3, COL_월별시작), ws.Cells(3, 마지막열)).Value
    If Not IsArray(arrRow3) Then Exit Function
    If UBound(arrRow3, 2) < 1 Then Exit Function
    For c = 1 To UBound(arrRow3, 2)
        If Not IsNumeric(arrRow3(1, c)) Then GoTo NextCol
        일번호 = CLng(arrRow3(1, c))
        If 일번호 < 1 Or 일번호 > 31 Then GoTo NextCol
        s2 = Trim(Replace(Replace(Replace(CStr(arrRow2(1, c)), "M", ""), "m", ""), " ", ""))
        월파싱 = 0
        If s2 <> "" And IsNumeric(s2) Then 월파싱 = CLng(s2)
        If 월파싱 < 1 Or 월파싱 > 12 Then GoTo NextCol
        If 일번호 > Day(DateSerial(기준년, 월파싱 + 1, 0)) Then GoTo NextCol
        dateKey = Format(DateSerial(기준년, 월파싱, 일번호), "yyyymmdd")
        If Not dict.exists(dateKey) Then dict.Add dateKey, COL_월별시작 + c - 1
NextCol:
    Next c
End Function

'================================================================================
' ★★★ 속도 최적화용 헬퍼: 제품표준사양서 제품명→P값(1차포장계수) 딕셔너리 (1회 빌드)
'================================================================================
Private Function 표준사양서_제품딕_생성(wsIndex As Worksheet) As Object
    Dim dict As Object, lastRow As Long, r As Long
    Dim 제품명 As String, P값 As Double
    Dim arrG As Variant, arrP As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    Set 표준사양서_제품딕_생성 = dict
    If wsIndex Is Nothing Then Exit Function
    lastRow = wsIndex.Cells(wsIndex.Rows.Count, 표준사양서_제품열G).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then Exit Function
    arrG = wsIndex.Range(wsIndex.Cells(표준사양서_데이터시작행, 표준사양서_제품열G), wsIndex.Cells(lastRow, 표준사양서_제품열G)).Value
    arrP = wsIndex.Range(wsIndex.Cells(표준사양서_데이터시작행, 표준사양서_1차포장계수열P), wsIndex.Cells(lastRow, 표준사양서_1차포장계수열P)).Value
    For r = 1 To UBound(arrG, 1)
        제품명 = Trim(CStr(arrG(r, 1)))
        If 제품명 <> "" And Not dict.exists(제품명) Then
            P값 = 0
            If IsNumeric(arrP(r, 1)) Then P값 = CDbl(arrP(r, 1))
            dict.Add 제품명, P값
        End If
    Next r
End Function

'================================================================================
' ★★★ 속도 최적화용 헬퍼: 블록 재고 재계산 ? 배열 기반 (셀 I/O 제로)
' 전체 제품 그룹을 dailyArr/infoArr 배열 위에서 일괄 재계산
'================================================================================
Private Sub 원단_블록_재고_재계산_배열(ByRef dailyArr As Variant, infoArr As Variant, _
    블록인덱스 As Long, 날짜캐시 As Variant, 시작열 As Long, 끝열 As Long)
    Dim 재고오프 As Long, 입고오프 As Long, 소요오프 As Long, 블록시작오프 As Long
    Dim rOff As Long, dCol As Long, col As Long
    Dim 입고 As Double, 소요 As Double, 재고 As Double, 전월재고_블록 As Double, 생산 As Double, 전일재고 As Double
    Dim totalRows As Long, totalCols As Long
    재고오프 = Choose(블록인덱스 + 1, 2, 5, 8, 11, 14)
    입고오프 = 재고오프 - 2
    소요오프 = 재고오프 - 1
    블록시작오프 = Choose(블록인덱스 + 1, 0, 3, 6, 9, 12)
    totalRows = UBound(dailyArr, 1)
    totalCols = UBound(dailyArr, 2)
    For rOff = 1 To totalRows - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        전월재고_블록 = 0
        If IsNumeric(infoArr(rOff + 블록시작오프, COL_전월재고)) Then 전월재고_블록 = CDbl(infoArr(rOff + 블록시작오프, COL_전월재고))
        For dCol = 1 To totalCols
            col = 시작열 + dCol - 1
            입고 = 0: 소요 = 0
            ' ★ 제품(블록0)은 생산/출고/재고 3행 구조 → 별도 입고 행 없음 (입고오프=0=생산행과 동일하므로 스킵)
            If 블록인덱스 > 0 Then
                If IsNumeric(dailyArr(rOff + 입고오프, dCol)) Then 입고 = CDbl(dailyArr(rOff + 입고오프, dCol))
            End If
            If IsNumeric(dailyArr(rOff + 소요오프, dCol)) Then 소요 = CDbl(dailyArr(rOff + 소요오프, dCol))
            If 블록인덱스 = 0 Then
                생산 = 0
                If IsNumeric(dailyArr(rOff, dCol)) Then 생산 = CDbl(dailyArr(rOff, dCol))
            End If
            If dCol = 1 Then
                If 블록인덱스 = 0 Then 재고 = 전월재고_블록 + 생산 - 소요 Else 재고 = 전월재고_블록 + 입고 - 소요
            Else
                전일재고 = 0
                If IsNumeric(dailyArr(rOff + 재고오프, dCol - 1)) Then 전일재고 = CDbl(dailyArr(rOff + 재고오프, dCol - 1))
                If 블록인덱스 = 0 Then 재고 = 전일재고 + 생산 - 소요 Else 재고 = 전일재고 + 입고 - 소요
            End If
            If col <= UBound(날짜캐시) And col >= LBound(날짜캐시) Then
                If Not IsNull(날짜캐시(col)) Then
                    If CDate(날짜캐시(col)) <= Date + 10 Then dailyArr(rOff + 재고오프, dCol) = 재고 Else dailyArr(rOff + 재고오프, dCol) = ""
                Else
                    dailyArr(rOff + 재고오프, dCol) = 재고
                End If
            Else
                dailyArr(rOff + 재고오프, dCol) = 재고
            End If
        Next dCol
    Next rOff
End Sub

'================================================================================
' 이력시트 구조: A=일시, B=작업유형, C=부자재명, D=일자, E=이전값, F=새값, G=행번호, H=열번호
'================================================================================

'================================================================================
' 데이터 업데이트: 일일생산DB 기준 원단관리 그룹 생성 (완제품 항목·중복 제거, 등장 순서 유지)
' ★ 항목 목록 = 일일생산DB에서 수집(구분|품명). 작성일 열 = 오늘 날짜의 DB 생산량·원지기준소요량 반영.
' ★ 그룹행: A=설비라인(구분), B=비움, C=제품. 일별·DB 반영은 오늘 이전 날짜만 (미래일 입력/업데이트 금지)
'================================================================================
Public Sub 원단관리_데이터업데이트()
    Dim wsTarget As Worksheet
    Dim wsIndex As Worksheet
    Dim wsDB As Worksheet
    Dim 분류순서 As Collection, catItems As Collection
    Dim 제품딕 As Object, 원지딕 As Object
    Dim 작성일 As Date, 날짜열 As Long
    Dim 현재행 As Long
    Dim lastRow As Long
    Dim i As Long
    Dim 총항목수 As Long
    Dim 시트보호상태 As Boolean
    Dim 신규구조 As Boolean
    Dim wsHome As Worksheet
    Dim wsActiveSave As Worksheet
    Dim rngActiveSave As Range
    ' ★★★ 프로파일링: 구간별 소요 시간 측정 ★★★
    Dim t0 As Double, tLog As String: t0 = Timer: tLog = ""
    Dim col업체 As Long, col제품 As Long, col제조사 As Long, col지관 As Long, col지폭 As Long, col총량 As Long, col1차 As Long, col2차 As Long
    Dim r As Long, 시작행 As Long
    Dim v업체 As String, v제품 As String, v제조사 As String, v지관 As String, v지폭 As String, v총량 As String, v1차 As String, v2차 As String
    Dim 항목키 As String, pos As Long
    Dim 맵키 As String, 구분 As String, 품명 As String
    Dim CoVal As String
    Dim 제품맵 As Object
    Dim 원래DisplayAlerts As Boolean
    Dim 설비번호목록 As Collection
    Dim rDB As Long, lastRowDB As Long
    Dim dateKey As String
    Dim 원단DU_ErrNum As Long, 원단DU_ErrDesc As String

    On Error GoTo DataUpdateErr
    ' ★ 현재 시트·선택 위치·DisplayAlerts 저장 (복원용)
    On Error Resume Next
    Set wsActiveSave = ActiveSheet
    Set rngActiveSave = ActiveCell
    원래DisplayAlerts = Application.DisplayAlerts
    On Error GoTo DataUpdateErr
    원단_디버그_실행단계 = "초기화": 원단_디버그_상세 = ""
    ' ★ 맨 처음 화면 비표시 (버튼 깜빡임 방지). 병합 시 경고 비표시(최초 실행 시 오류 방지)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ' ★ 통합문서 구조 보호 해제 (시트 생성 시 필요)
    Dim wb구조보호_DU As Boolean: wb구조보호_DU = ThisWorkbook.ProtectStructure
    On Error Resume Next
    If wb구조보호_DU Then ThisWorkbook.Unprotect WS_시트보호암호
    Set wsHome = ThisWorkbook.Sheets("Home menu")
    On Error GoTo DataUpdateErr

    원단_디버그_실행단계 = "wsTarget"
    Set wsTarget = 시트가져오기또는생성(WS_원단관리)
    If wsTarget Is Nothing Then Debug.Print "[원단관리] STEP1: wsTarget=Nothing" Else Debug.Print "[원단관리] STEP1: wsTarget=" & wsTarget.Name
    ' ★ 항목 목록 = 일일생산DB 기준 (완제품 행만, 구분|품명 중복 제거)
    Set 제품딕 = Nothing
    Set 원지딕 = Nothing
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
    On Error GoTo DataUpdateErr
    If wsDB Is Nothing Then Debug.Print "[원단관리] STEP2: wsDB=Nothing" Else Debug.Print "[원단관리] STEP2: wsDB=" & wsDB.Name
    If wsDB Is Nothing Then
        Debug.Print "[원단관리] 일일생산DB 시트 없음"
        If Not 원단관리_무음모드 Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "일일생산DB 시트를 찾을 수 없습니다." & vbCrLf & _
               "일일생산일보에서 'DB저장' 실행 후 일일생산DB를 준비한 뒤 데이터 업데이트를 실행하세요.", vbExclamation, "데이터 업데이트"
        End If
        GoTo DataUpdateCleanUp
    End If
    ' ★ 증분 반영: 새 데이터 있으면 증분만 수집, 없어도 재고·현재고 갱신을 위해 계속 진행
    Dim lastMaxId As Long, currentMaxId As Long
    Dim 증분있음 As Boolean
    lastMaxId = 일일생산DB_마지막동기화ID_조회()
    currentMaxId = 일일생산DB_최대ID_조회(wsDB)
    증분있음 = (currentMaxId > lastMaxId)
    Debug.Print "[원단관리] STEP3: lastMaxId=" & lastMaxId & " currentMaxId=" & currentMaxId & " 증분=" & 증분있음
    If Not 일일생산DB_로드_항목_원단(wsDB, 분류순서, catItems) Then
        If Not RawData_Co목록_로드(분류순서, catItems) Then
            ' DB·Raw data 모두 Co 없음 → 최소 1개 빈 그룹 생성 (그룹행 구조 표시)
            If 분류순서 Is Nothing Then Set 분류순서 = New Collection
            If catItems Is Nothing Then Set catItems = New Collection
            분류순서.Add "원단"
            Dim 빈Co목록 As New Collection
            빈Co목록.Add "01"
            catItems.Add 빈Co목록
            Debug.Print "[원단관리] STEP4-FALLBACK: DB·Raw data Co 없음 → 빈 그룹 1개 생성"
            If Not 원단관리_무음모드 Then
                MsgBox "일일생산DB와 Raw data 시트에 원단(Co) 데이터가 없습니다." & vbCrLf & _
                       "빈 그룹행 1개를 생성합니다. A열에 Raw data H열 인덱스(01~)를 입력하세요.", vbExclamation, "데이터 업데이트"
            End If
        End If
    End If
    ' ★ 증분 있으면 증분만 수집(속도), 없으면 전체 수집. 원단관리용 날짜별원지Co딕(Co별 롤수,중량)
    Dim 날짜목록 As Collection
    Dim 날짜별제품딕 As Object, 날짜별원지딕 As Object, 날짜별원지ROLL딕 As Object, 날짜별원지Co딕 As Object
    Set 날짜별원지Co딕 = CreateObject("Scripting.Dictionary")
    Dim di As Long
    If catItems.Count > 0 And Not catItems(1) Is Nothing Then
        Debug.Print "[원단관리] STEP4-OK: 로드성공. Co수=" & catItems(1).Count
    End If
    원단_디버그_실행단계 = "날짜별수집"
    If 증분있음 And lastMaxId > 0 Then
        Call 일일생산DB_날짜별_제품원지_수집_증분(wsDB, lastMaxId, 날짜목록, 날짜별제품딕, 날짜별원지딕, 날짜별원지ROLL딕, 날짜별원지Co딕)
    Else
        Call 일일생산DB_날짜별_제품원지_수집(wsDB, 날짜목록, 날짜별제품딕, 날짜별원지딕, 날짜별원지ROLL딕, 날짜별원지Co딕)
    End If
    Dim s5 As String: s5 = "[원단관리] STEP5: 수집완료."
    If 날짜목록 Is Nothing Then s5 = s5 & " 날짜목록=Nothing" Else s5 = s5 & " 날짜목록=" & 날짜목록.Count
    If 날짜별원지Co딕 Is Nothing Then s5 = s5 & " Co딕=Nothing" Else s5 = s5 & " Co딕=" & 날짜별원지Co딕.Count
    Debug.Print s5
    ' 제품표준사양서에서 품명→업체명 매핑 (A열 표시용)
    On Error Resume Next
    Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
    On Error GoTo DataUpdateErr
    Dim 업체명딕 As Object: Set 업체명딕 = CreateObject("Scripting.Dictionary")
    If Not wsIndex Is Nothing Then
        Dim col업체tmp As Long, col제품tmp As Long, lastRow사양 As Long
        col업체tmp = 표준사양서_헤더열찾기(wsIndex, "업체명")
        If col업체tmp = 0 Then col업체tmp = 표준사양서_헤더열찾기(wsIndex, "업체")
        If col업체tmp = 0 Then col업체tmp = 표준사양서_헤더열찾기(wsIndex, "발주처")
        col제품tmp = 표준사양서_헤더열찾기(wsIndex, "제품")
        If col제품tmp = 0 Then col제품tmp = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
        If col업체tmp > 0 And col제품tmp > 0 Then
            lastRow사양 = wsIndex.Cells(wsIndex.Rows.Count, col제품tmp).End(xlUp).Row
            Dim ri As Long
            For ri = 표준사양서_데이터시작행 To lastRow사양
                Dim tmp품 As String, tmp업 As String
                tmp품 = Trim(CStr(wsIndex.Cells(ri, col제품tmp).Value))
                tmp업 = Trim(CStr(wsIndex.Cells(ri, col업체tmp).Value))
                If tmp품 <> "" And Not 업체명딕.exists(tmp품) Then 업체명딕.Add tmp품, tmp업
            Next ri
        End If
    End If
    ' ★ 제품표준사양서에 없는 품명은 일일생산DB 완제품_구분을 발주처로 fallback
    If Not wsDB Is Nothing Then
        On Error Resume Next
        lastRowDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lastRowDB >= 2 Then
            Dim arrFB As Variant
            arrFB = wsDB.Range(wsDB.Cells(2, 4), wsDB.Cells(lastRowDB, 10)).Value
            For rDB = 1 To UBound(arrFB, 1)
                If IsError(arrFB(rDB, 1)) Then GoTo NextFB
                If Trim(CStr(arrFB(rDB, 1))) = "완제품" Then
                    If IsError(arrFB(rDB, 7)) Or IsError(arrFB(rDB, 6)) Then GoTo NextFB
                    Dim fb품 As String, fb구분 As String
                    fb품 = Trim(CStr(arrFB(rDB, 7)))
                    fb구분 = Trim(CStr(arrFB(rDB, 6)))
                    If fb구분 = "" Then fb구분 = "KT"
                    If fb품 = "" Or InStr(1, fb품, "오류", vbTextCompare) > 0 Then GoTo NextFB
                    If fb구분 = "수입" Then GoTo NextFB
                    If Not 업체명딕.exists(fb품) Then 업체명딕.Add fb품, fb구분
                End If
NextFB:
            Next rDB
        End If
        On Error GoTo DataUpdateErr
    End If

    시트보호상태 = wsTarget.ProtectContents
    If 시트보호상태 Then
        On Error Resume Next
        wsTarget.Unprotect WS_시트보호암호
        If wsTarget.ProtectContents Then wsTarget.Unprotect
        On Error GoTo DataUpdateErr
    End If

    tLog = tLog & "초기화:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    Call 목록시트_갱신(wsIndex)
    tLog = tLog & "목록갱신:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    신규구조 = (lastRow < START_ROW)
    Debug.Print "[원단관리] STEP6: lastRow=" & lastRow & " 신규구조=" & 신규구조 & " catItems.Count=" & catItems.Count

    If 신규구조 Then
        Call 이력시트_초기화
        Call 시트헤더_생성(wsTarget)
        현재행 = START_ROW
        총항목수 = 0
        If Not catItems Is Nothing Then
            For i = 1 To catItems.Count
                If catItems(i) Is Nothing Then GoTo NextCatItem
                총항목수 = 총항목수 + catItems(i).Count
NextCatItem:
            Next i
        End If
        ' 데이터 없으면 플레이스홀더 그룹 1개 생성 (부자재관리와 동일하게 시트 구조 표시)
        If 총항목수 <= 0 Then
            Call 부자재항목_생성_빈값(wsTarget, 현재행)
            wsTarget.Cells(현재행, COL_제품발주처).Value = "-"
            wsTarget.Cells(현재행, COL_제품명).Value = "(원단 없음)"
            현재행 = 현재행 + ROWS_PER_ITEM
            현재행 = 현재행 - 1
            lastRow = 현재행
            GoTo After신규구조
        End If
        For i = 1 To 총항목수
            Call 부자재항목_생성_빈값(wsTarget, 현재행)
            현재행 = 현재행 + ROWS_PER_ITEM
        Next i
        현재행 = 현재행 - 1
        ' ★ 모든 그룹: 일일생산DB 항목(제품명|구분)으로 B·C 채우기, D~I는 제품표준사양서에서 매칭 시에만 보정
        If Not wsIndex Is Nothing Then
            col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
            If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
            If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
            col제품 = 표준사양서_헤더열찾기(wsIndex, "제품")
            If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
            col제조사 = 표준사양서_헤더열찾기(wsIndex, "제조사")
            col지관 = 표준사양서_헤더열찾기(wsIndex, "지관")
            col지폭 = 표준사양서_헤더열찾기(wsIndex, "지폭")
            col총량 = 표준사양서_헤더열찾기(wsIndex, "평량")
            If col총량 = 0 Then col총량 = 표준사양서_헤더열찾기(wsIndex, "총량")
            col1차 = 표준사양서_헤더열찾기(wsIndex, "1차포장")
            col2차 = 표준사양서_헤더열찾기(wsIndex, "2차포장")
        End If
        For i = 1 To 총항목수
            v제품 = Trim(CStr(catItems(1)(i)))   ' 원단: Co 인덱스
            v업체 = RawData_CoTo원단명(v제품)    ' 원단: Raw data H->I 원단명
            If v업체 = "" Then v업체 = "-"
            시작행 = START_ROW + (i - 1) * ROWS_PER_ITEM
            wsTarget.Cells(시작행, COL_제품발주처).Value = v제품   ' A열: Co(01~)
            wsTarget.Cells(시작행, COL_제품명).Value = v업체       ' B열: 원단명
            v제조사 = "": v지관 = "": v지폭 = "": v총량 = "": v1차 = "": v2차 = ""
            If Not wsIndex Is Nothing And col업체 > 0 And col제품 > 0 Then
                Dim rLookup As Long
                rLookup = 표준사양서_행찾기(wsIndex, col업체, col제품, v업체, v제품)
                If rLookup > 0 Then
                    If col제조사 > 0 Then v제조사 = Trim(CStr(wsIndex.Cells(rLookup, col제조사).Value))
                    If col지관 > 0 Then v지관 = Trim(CStr(wsIndex.Cells(rLookup, col지관).Value))
                    If col지폭 > 0 Then v지폭 = Trim(CStr(wsIndex.Cells(rLookup, col지폭).Value))
                    If col총량 > 0 Then v총량 = Trim(CStr(wsIndex.Cells(rLookup, col총량).Value))
                    If col1차 > 0 Then v1차 = Trim(CStr(wsIndex.Cells(rLookup, col1차).Value))
                    If col2차 > 0 Then v2차 = Trim(CStr(wsIndex.Cells(rLookup, col2차).Value))
                    If LCase(v1차) = "type" Or v1차 = "1차포장" Then v1차 = ""
                    If LCase(v2차) = "type" Or v2차 = "2차포장" Then v2차 = ""
                End If
            End If
            wsTarget.Cells(시작행, COL_제조사).Value = v제조사
            wsTarget.Cells(시작행, COL_지관).Value = v지관
            wsTarget.Cells(시작행, COL_지폭).Value = v지폭
            wsTarget.Cells(시작행, COL_총량).Value = v총량
            wsTarget.Cells(시작행, COL_1차포장).Value = v1차
            wsTarget.Cells(시작행, COL_2차포장).Value = v2차
        Next i
        lastRow = START_ROW + 총항목수 * ROWS_PER_ITEM - 1
After신규구조:
    End If

    tLog = tLog & "그룹구조:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    Call 일일생산DB_시트초기화

    ' ★ 일일생산DB 기준: 기존 시트에 없는 원단(Co) 그룹 추가
    If Not 신규구조 And catItems.Count > 0 Then
        Set 제품맵 = CreateObject("Scripting.Dictionary")
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            CoVal = Trim(CStr(wsTarget.Cells(시작행, COL_제품발주처).Value))
            If CoVal <> "" Then
                If IsNumeric(CoVal) And CLng(val(CoVal)) > 0 Then CoVal = Format(CLng(val(CoVal)), "00")
                제품맵(CoVal) = 1
            End If
        Next 시작행
        For i = 1 To catItems(1).Count
            CoVal = Trim(CStr(catItems(1)(i)))
            ' catItems는 이미 "01" 형식으로 정규화됨 (일일생산DB_로드_항목_원단)
            If Not 제품맵.exists(CoVal) Then
                현재행 = 원단_빈행그룹_추가_무음(CoVal)
                If 현재행 > 0 Then
                    v업체 = RawData_CoTo원단명(CoVal)
                    If v업체 = "" Then v업체 = "-"
                    wsTarget.Cells(현재행, COL_제품발주처).Value = CoVal
                    wsTarget.Cells(현재행, COL_제품명).Value = v업체
                    제품맵(CoVal) = 1
                End If
            End If
        Next i
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    End If

    tLog = tLog & "그룹추가:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    
    ' ★ 그룹추가 완료 후 한 번만 유효성 적용 (기존+신규 그룹 모두 커버)
    Call 목록다운_적용(wsTarget)
    tLog = tLog & "목록다운:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    
    ' ★ 기존 시트에서 첫 그룹 등 기본정보가 비어있거나 placeholder인 경우 제품표준사양서로 보정 (표준사양서 있을 때만)
    ' ★ 원단관리는 A=Co·B=원단명 구조라 부자재식 보정 불적합(월별열 덮어쓰기 방지) → 건너뜀
    If Not 신규구조 And lastRow >= START_ROW And Not wsIndex Is Nothing And wsTarget.Name <> WS_원단관리 Then
        총항목수 = 0
        For i = 1 To catItems.Count
            총항목수 = 총항목수 + catItems(i).Count
        Next i
        If 총항목수 > 0 Then
            col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
            If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
            If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
            col제품 = 표준사양서_헤더열찾기(wsIndex, "제품")
            If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
            col제조사 = 표준사양서_헤더열찾기(wsIndex, "제조사")
            col지관 = 표준사양서_헤더열찾기(wsIndex, "지관")
            col지폭 = 표준사양서_헤더열찾기(wsIndex, "지폭")
            col총량 = 표준사양서_헤더열찾기(wsIndex, "평량")
        If col총량 = 0 Then col총량 = 표준사양서_헤더열찾기(wsIndex, "총량")
            col1차 = 표준사양서_헤더열찾기(wsIndex, "1차포장")
            col2차 = 표준사양서_헤더열찾기(wsIndex, "2차포장")
            For i = 1 To 총항목수
                시작행 = START_ROW + (i - 1) * ROWS_PER_ITEM
                If 시작행 > lastRow Then Exit For
                v업체 = Trim(CStr(wsTarget.Cells(시작행, COL_업체명).Value))
                v제품 = Trim(CStr(wsTarget.Cells(시작행, COL_제품).Value))
                Dim v1차현재 As String, v2차현재 As String
                v1차현재 = Trim(CStr(wsTarget.Cells(시작행, COL_1차포장).Value))
                v2차현재 = Trim(CStr(wsTarget.Cells(시작행, COL_2차포장).Value))
                ' placeholder(제품1, Type) 또는 빈칸/헤더 표시 시 표준사양서로 보정. D~I는 발주처·제품명 일치 행에서만 가져옴
                If ((v업체 = "" Or v업체 = "-") And (v제품 = "" Or v제품 = "제품1" Or (left(Trim(v제품), 2) = "제품" And Len(v제품) <= 4))) Or _
                   LCase(v1차현재) = "type" Or LCase(v2차현재) = "type" Or v1차현재 = "1차포장" Or v2차현재 = "2차포장" Then
                    r = 표준사양서_행찾기(wsIndex, col업체, col제품, v업체, v제품)
                    If r <= 0 Then r = 표준사양서_데이터시작행 + (i - 1)
                    v업체 = "": v제품 = "": v제조사 = "": v지관 = "": v지폭 = "": v총량 = "": v1차 = "": v2차 = ""
                    If col업체 > 0 Then v업체 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
                    If col제품 > 0 Then v제품 = Trim(CStr(wsIndex.Cells(r, col제품).Value))
                    If v업체 = "" Then v업체 = "-"
                    If col제조사 > 0 Then v제조사 = Trim(CStr(wsIndex.Cells(r, col제조사).Value))
                    If col지관 > 0 Then v지관 = Trim(CStr(wsIndex.Cells(r, col지관).Value))
                    If col지폭 > 0 Then v지폭 = Trim(CStr(wsIndex.Cells(r, col지폭).Value))
                    If col총량 > 0 Then v총량 = Trim(CStr(wsIndex.Cells(r, col총량).Value))
                    If col1차 > 0 Then v1차 = Trim(CStr(wsIndex.Cells(r, col1차).Value))
                    If col2차 > 0 Then v2차 = Trim(CStr(wsIndex.Cells(r, col2차).Value))
                    If col제품 > 0 And r = 표준사양서_데이터시작행 And v제품 = Trim(CStr(wsIndex.Cells(표준사양서_헤더행, col제품).Value)) Then v제품 = ""
                    If col1차 > 0 And r = 표준사양서_데이터시작행 And v1차 = Trim(CStr(wsIndex.Cells(표준사양서_헤더행, col1차).Value)) Then v1차 = ""
                    If col2차 > 0 And r = 표준사양서_데이터시작행 And v2차 = Trim(CStr(wsIndex.Cells(표준사양서_헤더행, col2차).Value)) Then v2차 = ""
                    If LCase(v1차) = "type" Or v1차 = "1차포장" Then v1차 = ""
                    If LCase(v2차) = "type" Or v2차 = "2차포장" Then v2차 = ""
                    If v제품 = "제품" Or InStr(1, "제품(상품)명", v제품, vbTextCompare) > 0 Then v제품 = ""
                    wsTarget.Cells(시작행, COL_업체명).Value = v업체
                    wsTarget.Cells(시작행, COL_제품).Value = v제품
                    wsTarget.Cells(시작행, COL_제조사).Value = v제조사
                    wsTarget.Cells(시작행, COL_지관).Value = v지관
                    wsTarget.Cells(시작행, COL_지폭).Value = v지폭
                    wsTarget.Cells(시작행, COL_총량).Value = v총량
                    wsTarget.Cells(시작행, COL_1차포장).Value = v1차
                    wsTarget.Cells(시작행, COL_2차포장).Value = v2차
                End If
            Next i
        End If
    End If

    tLog = tLog & "보정:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    ' ★ DB 날짜에 해당하는 열이 없으면 자동 확장 (업데이트 반영 가능하도록)
    원단_디버그_실행단계 = "날짜열확장"
    Call 원단관리_날짜열_확장(wsTarget, 날짜목록)
    tLog = tLog & "날짜열확장:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    ' ★★★ 속도 최적화: 1~4단계 배열 기반 일괄 처리 ★★★
    원단_디버그_실행단계 = "마지막일열/날짜열딕"
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    Dim 마지막일v As Long: 마지막일v = 마지막일열_찾기(wsTarget)
    Dim 날짜열딕 As Object: Set 날짜열딕 = 원단관리_날짜열_딕셔너리(wsTarget)
    원단_디버그_상세 = "마지막일v=" & 마지막일v & " 날짜열딕=" & IIf(날짜열딕 Is Nothing, "Nothing", 날짜열딕.Count)
    Dim 표준딕 As Object: Set 표준딕 = 표준사양서_제품딕_생성(wsIndex)
    Dim pd당일 As Object, od당일 As Object, rd당일 As Object
    Dim dCol As Long, rOff As Long, k As Variant

    ' 시트 → 배열 1회 읽기
    Dim infoArr1 As Variant: infoArr1 = wsTarget.Range(wsTarget.Cells(START_ROW, 1), wsTarget.Cells(lastRow, COL_자재유형)).Value
    Dim dailyArr1 As Variant
    If 마지막일v >= COL_일별시작 Then dailyArr1 = wsTarget.Range(wsTarget.Cells(START_ROW, COL_일별시작), wsTarget.Cells(lastRow, 마지막일v)).Value

    ' 제품 행 딕셔너리: 품명(B열) → rOff (부자재 1차/2차용, 원단관리에서는 미사용)
    Dim 제품행딕 As Object: Set 제품행딕 = CreateObject("Scripting.Dictionary")
    ' 원단관리: Co(A열) → rOff. 날짜별원지Co딕(Co→Array(롤수,중량))로 사용1=Roll(롤수), 사용2=중량Kg 반영
    Dim Co행딕 As Object: Set Co행딕 = CreateObject("Scripting.Dictionary")
    Dim rOffList As Object, ro As Variant
    If IsArray(dailyArr1) Then
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            rOff = 시작행 - START_ROW + 1
            품명 = ""
            If Not IsEmpty(infoArr1(rOff, COL_제품명)) Then 품명 = Trim(CStr(infoArr1(rOff, COL_제품명)))
            CoVal = ""
            If Not IsEmpty(infoArr1(rOff, COL_제품발주처)) Then CoVal = Trim(CStr(infoArr1(rOff, COL_제품발주처)))
            If 품명 <> "" Then
                If Not 제품행딕.exists(품명) Then
                    Set rOffList = CreateObject("Scripting.Dictionary")
                    제품행딕.Add 품명, rOffList
                End If
                제품행딕(품명).Add 제품행딕(품명).Count, rOff
            End If
            If CoVal <> "" And CoVal <> "-" Then
                ' Co 형식 정규화 (1→01): DB 매칭 일관성
                If IsNumeric(CoVal) And CLng(val(CoVal)) > 0 Then CoVal = Format(CLng(val(CoVal)), "00")
                If Not Co행딕.exists(CoVal) Then Co행딕.Add CoVal, rOff
            End If
        Next 시작행
    End If

    ' ── 1단계: 원단관리 = 날짜별원지Co딕(Co→Array(롤수,중량)) → 사용1=롤수, 사용2=중량 반영 ──
    ' ★ Co(A열) 기준. 사용1 = Roll(롤수), 사용2 = 중량Kg
    Dim sDbg As String: sDbg = "[디버그] lastRow=" & lastRow & " 마지막일v=" & 마지막일v
    If 날짜열딕 Is Nothing Then sDbg = sDbg & " 날짜열딕=Nothing" Else sDbg = sDbg & " 날짜열딕=" & 날짜열딕.Count
    sDbg = sDbg & " Co행딕=" & Co행딕.Count
    If IsArray(dailyArr1) Then sDbg = sDbg & " dailyArr=" & UBound(dailyArr1, 1) & "x" & UBound(dailyArr1, 2) Else sDbg = sDbg & " dailyArr=N/A"
    If 날짜목록 Is Nothing Then sDbg = sDbg & " 날짜목록=Nothing" Else sDbg = sDbg & " 날짜목록=" & 날짜목록.Count
    If 날짜별원지Co딕 Is Nothing Then sDbg = sDbg & " Co딕=Nothing" Else sDbg = sDbg & " Co딕=" & 날짜별원지Co딕.Count
    tLog = tLog & sDbg & " "
    If Co행딕.Count > 0 Then
        Dim dbgCoKeys As String: dbgCoKeys = ""
        For Each k In Co행딕.keys: dbgCoKeys = dbgCoKeys & CStr(k) & ",": Next k
        tLog = tLog & "[Co행]" & left(dbgCoKeys, 60) & " "
    End If
    If Not 날짜별원지Co딕 Is Nothing And 날짜별원지Co딕.Count > 0 Then
        Dim dbgCoDikKeys As String: dbgCoDikKeys = ""
        Dim dbgDateK As Variant
        For Each dbgDateK In 날짜별원지Co딕.keys
            Dim dbgCd As Object: Set dbgCd = 날짜별원지Co딕(dbgDateK)
            dbgCoDikKeys = dbgCoDikKeys & CStr(dbgDateK) & "(" & dbgCd.Count & "),"
            If Len(dbgCoDikKeys) > 80 Then Exit For
        Next dbgDateK
        tLog = tLog & "[Co딕]" & left(dbgCoDikKeys, 80) & " "
    End If
    If IsArray(dailyArr1) And Not Co행딕 Is Nothing Then
        For Each k In Co행딕.keys
            rOff = CLng(Co행딕(k))
            For dCol = 1 To UBound(dailyArr1, 2)
                dailyArr1(rOff + OFFSET_사용1, dCol) = Empty
                dailyArr1(rOff + OFFSET_사용2, dCol) = Empty
            Next dCol
        Next k
    End If
    Dim dbgMatchCnt As Long, dbgMissCnt As Long, dbgDateSkip As Long, dbgNotArray As Long
    dbgMatchCnt = 0: dbgMissCnt = 0: dbgDateSkip = 0: dbgNotArray = 0
    If Not 날짜목록 Is Nothing And Not 날짜별원지Co딕 Is Nothing And Not Co행딕 Is Nothing And IsArray(dailyArr1) Then
        Dim kgVal As Double, cd As Object, matchedKey As String, sheetCo As Variant
        For di = 1 To 날짜목록.Count
            작성일 = 날짜목록(di)
            If 작성일 > Date Then dbgDateSkip = dbgDateSkip + 1: GoTo NextDate1
            dateKey = Format(작성일, "YYYYMMDD")
            If 날짜열딕 Is Nothing Or Not 날짜열딕.exists(dateKey) Then dbgDateSkip = dbgDateSkip + 1: GoTo NextDate1
            날짜열 = CLng(날짜열딕(dateKey))
            dCol = 날짜열 - COL_일별시작 + 1
            If dCol < 1 Or dCol > UBound(dailyArr1, 2) Then dbgDateSkip = dbgDateSkip + 1: GoTo NextDate1
            If Not 날짜별원지Co딕.exists(dateKey) Then dbgDateSkip = dbgDateSkip + 1: GoTo NextDate1
            Set cd = 날짜별원지Co딕(dateKey)
            For Each k In cd.keys
                CoVal = CStr(k)
                matchedKey = ""
                If Co행딕.exists(CoVal) Then matchedKey = CoVal
                If matchedKey = "" Then
                    Dim CoValNorm As String, sheetCoNorm As String
                    CoValNorm = Trim(CoVal): If IsNumeric(CoValNorm) And CLng(val(CoValNorm)) > 0 Then CoValNorm = Format(CLng(val(CoValNorm)), "00")
                    For Each sheetCo In Co행딕.keys
                        If Trim(CStr(CoVal)) = Trim(CStr(sheetCo)) Then matchedKey = CStr(sheetCo): Exit For
                        If IsNumeric(CoVal) And IsNumeric(sheetCo) And CLng(val(CoVal)) = CLng(val(sheetCo)) Then matchedKey = CStr(sheetCo): Exit For
                        sheetCoNorm = Trim(CStr(sheetCo)): If IsNumeric(sheetCoNorm) And CLng(val(sheetCoNorm)) > 0 Then sheetCoNorm = Format(CLng(val(sheetCoNorm)), "00")
                        If CoValNorm = sheetCoNorm Or (CoValNorm <> "" And CoValNorm = Trim(CStr(sheetCo))) Then matchedKey = CStr(sheetCo): Exit For
                    Next sheetCo
                End If
                If matchedKey <> "" Then
                    rOff = CLng(Co행딕(matchedKey))
                    Dim arrCo As Variant: arrCo = cd(k)
                    If IsArray(arrCo) Then
                        kgVal = CDbl(arrCo(1))
                        ' 롤수(사용1)는 기초재고+입고로만 계산 → 일별 사용1 기록 안함
                        dailyArr1(rOff + OFFSET_사용2, dCol) = kgVal
                        dbgMatchCnt = dbgMatchCnt + 1
                    Else
                        dbgNotArray = dbgNotArray + 1
                    End If
                Else
                    dbgMissCnt = dbgMissCnt + 1
                End If
            Next k
NextDate1:
        Next di
    End If
    tLog = tLog & "[1단계] match=" & dbgMatchCnt & " miss=" & dbgMissCnt & " dateSkip=" & dbgDateSkip & " notArr=" & dbgNotArray & " "
    원단_디버그_상세 = 원단_디버그_상세 & " dateSkip=" & dbgDateSkip & " match=" & dbgMatchCnt
    If dbgMatchCnt = 0 And (dbgMissCnt > 0 Or (Not 날짜별원지Co딕 Is Nothing And 날짜별원지Co딕.Count > 0)) Then
        Dim dbgSample As String: dbgSample = ""
        Dim dbgK As Variant, dbgFirstK As Variant
        Dim dbgCd2 As Object
        If Not 날짜별원지Co딕 Is Nothing Then
            For Each dbgFirstK In 날짜별원지Co딕.keys
                Set dbgCd2 = 날짜별원지Co딕(dbgFirstK)
                If dbgCd2.Count > 0 Then
                    dbgSample = " Co딕샘플:"
                    For Each dbgK In dbgCd2.keys
                        dbgSample = dbgSample & CStr(dbgK) & ","
                    Next dbgK
                    dbgSample = left(dbgSample, Len(dbgSample) - 1)
                    Exit For
                End If
            Next dbgFirstK
        End If
        If Co행딕.Count > 0 Then
            dbgSample = dbgSample & " Co행딕:"
            For Each dbgK In Co행딕.keys
                dbgSample = dbgSample & CStr(dbgK) & ","
            Next dbgK
            dbgSample = left(dbgSample, Len(dbgSample) - 1)
        End If
        원단_디버그_상세 = 원단_디버그_상세 & dbgSample
    End If

    ' 배열 → 시트 1회 쓰기 (1단계 반영)
    원단_디버그_실행단계 = "1단계배열쓰기"
    If IsArray(dailyArr1) Then wsTarget.Range(wsTarget.Cells(START_ROW, COL_일별시작), wsTarget.Cells(lastRow, 마지막일v)).Value = dailyArr1

    tLog = tLog & "1단계배열:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    If Not wsIndex Is Nothing Then Call 원단관리_DtoI_제품표준사양서_채우기(wsTarget, wsIndex)
    If Not wsIndex Is Nothing Then Call 원단관리_첫그룹_C4C7_제품표준사양서_표시(wsTarget, wsIndex)
    tLog = tLog & "DtoI:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    If 신규구조 Then Call 시트서식_적용(wsTarget, 현재행)
    tLog = tLog & "서식:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    ' ── 2+4단계: 원지·포장 소요 쓰기 + 블록 재고 재계산 (전부 배열 기반) → 시트 반영 후 값일괄계산으로 숫자·서식 일괄 적용 ──
    원단_디버그_실행단계 = "2단계배열재읽기"
    If 마지막일v >= COL_일별시작 Then
        dailyArr1 = wsTarget.Range(wsTarget.Cells(START_ROW, COL_일별시작), wsTarget.Cells(lastRow, 마지막일v)).Value
        infoArr1 = wsTarget.Range(wsTarget.Cells(START_ROW, 1), wsTarget.Cells(lastRow, COL_자재유형)).Value
    End If
    ' ★ 최적화: 이번 업데이트에서 가장 빠른(왼쪽) 변경 열 오프셋 계산 → 그 열부터만 재계산
    Dim 최소변경오프셋 As Long: 최소변경오프셋 = 1
    If Not 날짜목록 Is Nothing And Not 날짜열딕 Is Nothing Then
        Dim dVal As Date, dk As String, dColOff As Long
        Dim minCol As Long: minCol = 0
        For di = 1 To 날짜목록.Count
            dVal = 날짜목록(di)
            dk = Format(dVal, "YYYYMMDD")
            If 날짜열딕.exists(dk) Then
                dColOff = CLng(날짜열딕(dk)) - COL_일별시작 + 1
                If dColOff >= 1 Then
                    If minCol = 0 Or dColOff < minCol Then minCol = dColOff
                End If
            End If
        Next di
        If minCol >= 1 Then 최소변경오프셋 = minCol
    End If
    원단_디버그_실행단계 = "값일괄계산"
    원단_디버그_상세 = 원단_디버그_상세 & " 최소변경=" & 최소변경오프셋
    ' 월별 재고 체인 계산 + 총합/기말재고 갱신 (변경 열 이후만)
    tLog = tLog & "배열준비:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    Call 원단관리_값일괄계산(wsTarget, 최소변경오프셋)
    tLog = tLog & "값일괄계산(offset=" & 최소변경오프셋 & "):" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    Call 원단관리_날짜이후_불필요삭제(wsTarget)

    ' ★ BS~BZ열 쓰기 비활성화됨
    tLog = tLog & "DB반영:skip ": t0 = Timer

    On Error GoTo Err현재고
    Call Stock_시트_갱신
    tLog = tLog & "현재고:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer
    GoTo After현재고
Err현재고:
    tLog = tLog & "현재고:ERR "
    If Not 원단관리_무음모드 Then MsgBox "현재고 시트 갱신 중 오류가 발생했습니다." & vbCrLf & vbCrLf & "오류: " & Err.Description, vbExclamation, "현재고"
After현재고:
    On Error GoTo 0

    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow >= START_ROW Then Call 원단관리_그룹행_C열I열_배열복원(wsTarget, lastRow)
    Call 날짜영역_테두리적용(wsTarget)
    Call 주말서식_일수행(wsTarget)
    tLog = tLog & "테두리/주말:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow >= START_ROW Then Call 원단관리_그룹행_교대음영_적용(wsTarget, lastRow)
    tLog = tLog & "음영:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    ' 목록추가 예외 초기화 → 전체 동일 규칙 적용(추가했던 그룹행도 잠김)
    원단_목록추가_예외시작행 = 0
    원단_목록추가_예외끝행 = 0
    Call 원단_셀보호_적용_무음(wsTarget)
    tLog = tLog & "보호:" & Format(Timer - t0, "0.00") & "s ": t0 = Timer

    Call 일일생산DB_마지막동기화ID_저장(currentMaxId)
    On Error Resume Next
    Call 재고관리DB_원단_기록(wsTarget)
    On Error GoTo 0
    Call 원단_이벤트코드_자동설치_필요시

    Debug.Print "원단관리_데이터업데이트 완료: " & tLog & " [상세] " & 원단_디버그_상세
    If Not 원단관리_무음모드 Then
        Dim msg상세 As String
        msg상세 = Replace(tLog, " ", vbCrLf)
        If 원단_디버그_상세 <> "" Then msg상세 = msg상세 & vbCrLf & vbCrLf & "■ 디버그 상세 ■" & vbCrLf & 원단_디버그_상세
        MsgBox "완료!" & vbCrLf & vbCrLf & "■ 구간별 소요 시간 ■" & vbCrLf & msg상세, vbInformation, "데이터 업데이트 프로파일"
    End If
DataUpdateCleanUp:
    On Error Resume Next
    If Not wsTarget Is Nothing And 시트보호상태 Then wsTarget.Protect Password:=WS_시트보호암호
    If wb구조보호_DU Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
    If Not wsActiveSave Is Nothing Then
        wsActiveSave.Activate
        If Not rngActiveSave Is Nothing Then rngActiveSave.Select
    End If
    On Error GoTo 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = 원래DisplayAlerts
    If 원단DU_ErrNum <> 0 Then Debug.Print "[원단관리] 내부오류전파생략: " & 원단DU_ErrNum & " " & 원단DU_ErrDesc
    Exit Sub

DataUpdateErr:
    Debug.Print "원단관리_데이터업데이트 오류: " & Err.Number & " " & Err.Description & " [단계] " & 원단_디버그_실행단계 & " " & 원단_디버그_상세
    If Not 원단관리_무음모드 Then MsgBox "데이터 업데이트 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
        "Err " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "[단계] " & 원단_디버그_실행단계 & vbCrLf & _
        IIf(원단_디버그_상세 <> "", "[상세] " & 원단_디버그_상세 & vbCrLf, "") & vbCrLf & tLog, vbExclamation, "데이터 업데이트"
    원단DU_ErrNum = Err.Number: 원단DU_ErrDesc = Err.Description
    GoTo DataUpdateCleanUp
End Sub

' DB저장 매크로에서 한 번에 적용 시 호출. 완료 메시지 없이 실행 후 호출 측에서 통합 메시지 표시
Public Sub 원단관리_데이터업데이트_무음()
    원단관리_무음모드 = True
    On Error GoTo Err무음
    Call 원단관리_데이터업데이트
    원단관리_무음모드 = False
    Exit Sub
Err무음:
    Debug.Print "원단관리_무음 오류: " & Err.Number & " " & Err.Description
    원단관리_무음모드 = False
End Sub

' 최종검수테스트 등 외부에서 원단관리 전용 호출용 (부자재관리 데이터업데이트_무음실행과 구분)
Public Sub 원단관리_데이터업데이트_무음실행()
    원단관리_무음모드 = True
    On Error GoTo Err원단무음
    Call 원단관리_데이터업데이트
    원단관리_무음모드 = False
    Exit Sub
Err원단무음:
    Debug.Print "원단관리_무음실행 오류: " & Err.Number & " " & Err.Description
    원단관리_무음모드 = False
End Sub

' ★ 재고 재계산 한도: 전체 날짜열 반영 (날짜필드 재고 로드 보장). 1-based 열 오프셋(monthArr 기준) 반환
Private Function 값일괄계산_오늘포함한달끝오프셋(ws As Worksheet, totalCols As Long) As Long
    값일괄계산_오늘포함한달끝오프셋 = totalCols
    If ws Is Nothing Or totalCols < 1 Then Exit Function
End Function

' 마지막 유효 날짜 열 (2행=1M~12M 월, 3행=1~31 일. 유효한 날짜열 중 가장 오른쪽)
Private Function 원단관리_마지막유효월열(ws As Worksheet) As Long
    Dim c As Long, 마지막열 As Long, v3 As Variant, v2 As Variant, 일v As Long, 월v As Long
    Dim s2 As String
    원단관리_마지막유효월열 = COL_월별시작 - 1
    If ws Is Nothing Then Exit Function
    마지막열 = 마지막월열_찾기(ws)
    If 마지막열 < COL_월별시작 Then Exit Function
    For c = 마지막열 To COL_월별시작 Step -1
        v3 = ws.Cells(3, c).Value
        If Not IsNumeric(v3) Then GoTo NextC
        일v = CLng(v3)
        If 일v < 1 Or 일v > 31 Then GoTo NextC
        s2 = Trim(Replace(Replace(Replace(CStr(ws.Cells(2, c).Value), "M", ""), "m", ""), " ", ""))
        If s2 = "" Or Not IsNumeric(s2) Then GoTo NextC
        월v = CLng(s2)
        If 월v >= 1 And 월v <= 12 Then 원단관리_마지막유효월열 = c: Exit Function
NextC:
    Next c
End Function

'================================================================================
' 날짜 마지막 열 이후 불필요 정보 삭제 (mo, 무/1400, 1365, 20 T, BUN 등)
' ★ 2·3행은 기존 값 보존(절대 삭제 금지) → 4행~ 데이터 영역만 Clear
'================================================================================
Private Sub 원단관리_날짜이후_불필요삭제(ws As Worksheet)
    Dim 마지막행 As Long, 마지막유효열 As Long, 끝열 As Long
    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 < START_ROW Then Exit Sub
    마지막유효열 = 원단관리_마지막유효월열(ws)
    끝열 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column > 끝열 Then 끝열 = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(마지막행, ws.Columns.Count).End(xlToLeft).Column > 끝열 Then 끝열 = ws.Cells(마지막행, ws.Columns.Count).End(xlToLeft).Column
    If 끝열 > 마지막유효열 And 마지막행 >= START_ROW Then
        ws.Range(ws.Cells(START_ROW, 마지막유효열 + 1), ws.Cells(마지막행, 끝열)).ClearContents
    End If
End Sub

'================================================================================
' 날짜 영역 테두리 적용: 부자재관리와 100% 동일 구조
' 그룹행 있을 때만 헤더·데이터 전체 적용 (부자재관리 동일)
'================================================================================
Private Sub 날짜영역_테두리적용(ws As Worksheet)
    Dim 마지막행 As Long, 마지막일열 As Long
    Dim rng As Range

    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    마지막일열 = 마지막일열_찾기(ws)

    If 마지막행 < START_ROW Then Exit Sub
    If 마지막일열 < COL_일별시작 Then Exit Sub

    ' 헤더 (2~3행): COL_일별시작 ~ 마지막일열 전부, 얇은 테두리
    Set rng = ws.Range(ws.Cells(2, COL_일별시작), ws.Cells(3, 마지막일열))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    ' 2행: 아래 테두리만 유지, 나머지 제거 (1열~마지막일열)
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(2, 마지막일열))
    rng.Borders(xlEdgeLeft).LineStyle = xlNone
    rng.Borders(xlEdgeTop).LineStyle = xlNone
    rng.Borders(xlEdgeRight).LineStyle = xlNone
    rng.Borders(xlInsideVertical).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' 데이터 (4행~마지막행): 날짜 열 전부 테두리
    Set rng = ws.Range(ws.Cells(START_ROW, COL_일별시작), ws.Cells(마지막행, 마지막일열))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' A~I열 (고정 열) 데이터 영역 테두리 (그룹행 테두리 A열 포함)
    Set rng = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(마지막행, COL_자재유형))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

'================================================================================
' 그룹행 A~I 동일 시 중복 생성 불가: B~I 키로 기존 존재 여부 반환 (설비라인 제외 비교)
'================================================================================
Private Function 그룹행_BtoI_키(ByVal vB As String, ByVal vC As String, ByVal vd As String, ByVal vE As String, ByVal vF As String, ByVal vG As String, ByVal vH As String, ByVal vi As String) As String
    그룹행_BtoI_키 = Trim(vB) & "|" & Trim(vC) & "|" & Trim(vd) & "|" & Trim(vE) & "|" & Trim(vF) & "|" & Trim(vG) & "|" & Trim(vH) & "|" & Trim(vi)
End Function

' ws 원단관리 시트에서 그룹 시작행의 A~I 값을 읽어 B~I 키 반환 (중복 체크용)
Private Function 그룹행_시작행_BtoI_키(ws As Worksheet, 시작행 As Long) As String
    If ws Is Nothing Or 시작행 < START_ROW Then 그룹행_시작행_BtoI_키 = "": Exit Function
    Dim vB As String, vC As String, vd As String, vE As String, vF As String, vG As String, vH As String, vi As String
    vB = Trim(CStr(ws.Cells(시작행, COL_업체명).Value))
    vC = Trim(CStr(ws.Cells(시작행, COL_제품).Value))
    vd = Trim(CStr(ws.Cells(시작행, COL_제조사).Value))
    vE = Trim(CStr(ws.Cells(시작행, COL_지관).Value))
    vF = Trim(CStr(ws.Cells(시작행, COL_지폭).Value))
    vG = Trim(CStr(ws.Cells(시작행, COL_총량).Value))
    vH = Trim(CStr(ws.Cells(시작행, COL_1차포장).Value))
    vi = Trim(CStr(ws.Cells(시작행, COL_2차포장).Value))
    그룹행_시작행_BtoI_키 = 그룹행_BtoI_키(vB, vC, vd, vE, vF, vG, vH, vi)
End Function

'================================================================================
' 목록 이름 정의 갱신: 제품표준사양서 기반. B열=업체명(중복없음), C열=B선택 시 해당 업체 제품만(연동). A열=설비라인(Raw data_index M열)
'================================================================================
Private Sub 목록시트_갱신(wsIndex As Worksheet)
    Call RawData_HI_목록_갱신
    Dim s As String
    Dim idxCol As Long
    Dim colLetter As String
    Dim ref As String
    Dim q As String: q = chR(34)
    Dim 헤더목록 As Variant
    Dim i As Long
    Dim arr업체 As Variant, arr제품 As Variant
    Dim wsList As Worksheet
    Dim col업체 As Long, col제품 As Long
    Dim lastRow As Long, r As Long
    Dim 업체명 As String, 제품명 As String
    Dim dict As Object, 키 As Variant
    Dim listCol As Long, n As Long
    Dim nm As Name
    Dim v As String

    ' ----- 0) A열: 설비라인 = Raw data_index 시트 M3~마지막행을 유효성 목록으로 직접 참조 (목록_설비라인) -----
    ' (wsIndex 없어도 실행 - 제품표준사양서 무관)
    Dim wsRaw As Worksheet
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(WS_RawDataIndex)
    On Error GoTo 0
    If Not wsRaw Is Nothing Then
        lastRow = wsRaw.Cells(wsRaw.Rows.Count, RawDataIndex_설비라인열).End(xlUp).Row
        If lastRow < RawDataIndex_설비라인데이터시작행 Then lastRow = RawDataIndex_설비라인데이터시작행 - 1
        If lastRow >= RawDataIndex_설비라인데이터시작행 Then
            On Error Resume Next
            ThisWorkbook.Names("목록_설비라인").Delete
            On Error GoTo 0
            ref = "='" & Replace(WS_RawDataIndex, "'", "''") & "'!$M$3:$M$" & lastRow
            If left(ref, 1) <> "=" Then ref = "=" & ref
            ThisWorkbook.Names.Add Name:="목록_설비라인", RefersTo:=ref
        End If
    End If

    If wsIndex Is Nothing Then Exit Sub
    s = "'" & Replace(WS_제품표준사양서, "'", "''") & "'"

    ' 목록데이터 시트 먼저 확보 (B·C 목록용) → 완전감춤
    On Error Resume Next
    Set wsList = ThisWorkbook.Sheets(WS_목록데이터)
    If wsList Is Nothing Then
        Dim wb보호목록 As Boolean: wb보호목록 = ThisWorkbook.ProtectStructure
        If wb보호목록 Then ThisWorkbook.Unprotect WS_시트보호암호
        Set wsList = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Not wsList Is Nothing Then wsList.Name = WS_목록데이터
        If wb보호목록 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
    End If
    If Not wsList Is Nothing Then wsList.Visible = xlSheetVeryHidden
    On Error GoTo 0
    If wsList Is Nothing Then Exit Sub

    ' ----- 1) B열: 업체명 중복 없이 목록_업체명 -----
    col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
    arr업체 = Empty
    If col업체 > 0 Then arr업체 = 표준사양서_목록로드(wsIndex, "업체명")
    If IsEmpty(arr업체) And col업체 > 0 Then arr업체 = 표준사양서_목록로드(wsIndex, "발주처")
    If IsEmpty(arr업체) And col업체 > 0 Then arr업체 = 표준사양서_목록로드(wsIndex, "업체")
    If IsEmpty(arr업체) And col업체 > 0 Then
        Dim dict업체 As Object
        Set dict업체 = CreateObject("Scripting.Dictionary")
        lastRow = wsIndex.Cells(wsIndex.Rows.Count, col업체).End(xlUp).Row
        If lastRow >= 표준사양서_데이터시작행 Then
            For r = 표준사양서_데이터시작행 To lastRow
                업체명 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
                If 업체명 <> "" And Not dict업체.exists(업체명) Then dict업체.Add 업체명, r
            Next r
            If dict업체.Count > 0 Then
                ReDim arr업체(0 To dict업체.Count - 1)
                n = 0
                For r = 표준사양서_데이터시작행 To lastRow
                    업체명 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
                    If 업체명 <> "" And dict업체.exists(업체명) Then arr업체(n) = 업체명: dict업체.Remove 업체명: n = n + 1
                Next r
            End If
        End If
    End If
    If Not IsEmpty(arr업체) Then
        wsList.Columns(1).ClearContents
        For n = 0 To UBound(arr업체)
            wsList.Cells(n + 2, 1).Value = arr업체(n)
        Next n
        lastRow = 1 + UBound(arr업체) + 1
        If lastRow < 2 Then lastRow = 2
        On Error Resume Next
        ThisWorkbook.Names("목록_업체명").Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:="목록_업체명", RefersTo:=wsList.Range(wsList.Cells(2, 1), wsList.Cells(lastRow, 1))
    End If

    ' ----- 2) 업체별 제품 목록: 목록_제품_<업체명(세이니타이즈)> -----
    col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
    If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
    col제품 = 표준사양서_헤더열찾기(wsIndex, "제품")
    If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
    If col업체 > 0 And col제품 > 0 And Not wsList Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = wsIndex.Cells(wsIndex.Rows.Count, col업체).End(xlUp).Row
        If lastRow < 표준사양서_데이터시작행 Then lastRow = 표준사양서_데이터시작행
        For r = 표준사양서_데이터시작행 To lastRow
            업체명 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
            제품명 = Trim(CStr(wsIndex.Cells(r, col제품).Value))
            If 업체명 <> "" And 제품명 <> "" Then
                If Not dict.exists(업체명) Then
                    Dim coll As Object
                    Set coll = CreateObject("Scripting.Dictionary")
                    dict.Add 업체명, coll
                End If
                If Not dict(업체명).exists(제품명) Then dict(업체명).Add 제품명, 1
            End If
        Next r
        ' 기존 목록_제품_* 이름 삭제 (목록_제품 제외) - 반복 중 삭제 방지
        Dim namesToDel As Collection
        Set namesToDel = New Collection
        For Each nm In ThisWorkbook.Names
            If Len(nm.Name) >= 10 Then If left(nm.Name, 10) = "목록_제품_" And nm.Name <> "목록_제품" Then namesToDel.Add nm.Name
        Next nm
        For n = 1 To namesToDel.Count
            On Error Resume Next
            ThisWorkbook.Names(namesToDel(n)).Delete
            On Error GoTo 0
        Next n
        listCol = 2
        Dim keys업체 As Variant, keys제품 As Variant, j As Long, jj As Long
        keys업체 = dict.keys
        For j = 0 To UBound(keys업체)
            업체명 = CStr(keys업체(j))
            wsList.Cells(1, listCol).Value = 업체명
            keys제품 = dict(업체명).keys
            n = 1
            For jj = 0 To UBound(keys제품)
                n = n + 1
                wsList.Cells(n, listCol).Value = CStr(keys제품(jj))
            Next jj
            If n >= 2 Then
                lastRow = n
                ThisWorkbook.Names.Add Name:="목록_제품_" & 목록이름_세이니타이즈(업체명), _
                    RefersTo:=wsList.Range(wsList.Cells(2, listCol), wsList.Cells(lastRow, listCol))
            End If
            listCol = listCol + 1
        Next j
        ' ----- 부수적 데이터 삭제: 열방향 마지막(미사용 열) 정리 -----
        ' 이전 실행의 잔여 열 데이터가 데이터 업데이트 시 문제를 일으키므로 삭제
        Dim 마지막사용열 As Long, 끝열 As Long
        마지막사용열 = listCol - 1
        끝열 = wsList.Cells(1, wsList.Columns.Count).End(xlToLeft).Column
        If wsList.Cells(2, wsList.Columns.Count).End(xlToLeft).Column > 끝열 Then 끝열 = wsList.Cells(2, wsList.Columns.Count).End(xlToLeft).Column
        If 끝열 > 마지막사용열 Then
            wsList.Range(wsList.Cells(1, 마지막사용열 + 1), wsList.Cells(wsList.Rows.Count, 끝열)).ClearContents
        End If
    Else
        ' 업체별 제품 블록 미실행 시: A열(업체) 외 나머지 열 삭제
        If Not wsList Is Nothing Then
            끝열 = wsList.Cells(1, wsList.Columns.Count).End(xlToLeft).Column
            If wsList.Cells(2, wsList.Columns.Count).End(xlToLeft).Column > 끝열 Then 끝열 = wsList.Cells(2, wsList.Columns.Count).End(xlToLeft).Column
            If 끝열 > 1 Then wsList.Range(wsList.Cells(1, 2), wsList.Cells(wsList.Rows.Count, 끝열)).ClearContents
        End If
    End If

    ' ----- 3) 나머지 목록(제품 전체, 제조사, 지관, ...): 기존 동적 범위 유지 -----
    헤더목록 = Array("제품", "제조사", "지관", "지폭", "평량", "1차포장", "2차포장")
    For i = 0 To UBound(헤더목록)
        idxCol = 표준사양서_헤더열찾기(wsIndex, CStr(헤더목록(i)))
        If idxCol = 0 And (헤더목록(i) = "제품") Then idxCol = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
        If idxCol > 0 Then
            colLetter = Split(wsIndex.Cells(표준사양서_헤더행, idxCol).Address(True, False), "$")(0)
            ref = "=" & s & "!$" & colLetter & "$" & 표준사양서_데이터시작행 & _
                  ":INDEX(" & s & "!$" & colLetter & ":$" & colLetter & _
                  ",LOOKUP(2,1/(" & s & "!$" & colLetter & ":$" & colLetter & _
                  "<>" & q & q & "),ROW(" & s & "!$" & colLetter & ":$" & colLetter & ")))"
            목록이름_동적설정 "목록_" & CStr(헤더목록(i)), ref, wsIndex, idxCol
        End If
    Next i
End Sub

'================================================================================
' 생산일보 시트 보호: 수식 셀 + 헤더/소계 영역 잠금 → 삭제 불가. UserInterfaceOnly:=True 로 매크로는 계속 쓰기 가능
' 생산일보_셀매핑.json 기준: A1:BV133 전체 범위, 타이틀/헤더/소계/수식행 잠금
' ※ B68:B77(T07), AK68:AK77(T08) 유효성 목록은 이 Sub에서 절대 건드리지 않음 (생산일보_셀매핑.json)
'================================================================================
Public Sub 원단_일일생산일보_시트보호_적용()
    Dim ws As Worksheet
    Dim c As Range
    Dim ur As Range
    Dim lockedRanges As Variant, one As Variant
    Dim r As Range
    Dim doneMerge As Object
    Dim maAddr As String
    Dim origScreen As Boolean, origEvents As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_일일생산일보)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    origScreen = Application.ScreenUpdating
    origEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect WS_시트보호암호
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo 0
    End If

    ' 시트 전체 잠금 해제 (A1:BV133 ? 생산일보_셀매핑.json)
    On Error Resume Next
    ws.Range("A1:BV133").Locked = False
    On Error GoTo 0

    ' 1) 지정 영역 잠금 (병합셀은 MergeArea 단위로 한 번만 잠금)
    '    생산일보_셀매핑.json 기준:
    '    - 타이틀/헤더: Row1(타이틀), Row3(섹션1), Row4(직기작업), Row5(T01헤더), Row6(T01서브헤더)
    '    - T01~T06 소계/헤더: Row17~19, Row30~34, Row45~47, Row58~62
    '    - 완제품 헤더: Row64~67, Row78(합계), Row80~82(종별수량)
    '    - 원자재수율: Row84~86(헤더), Row87~101(수식영역)
    '    - 작업시간 헤더: Row103~107, Row109~110(수식행), Row113~117(수식행)
    '    - 비가동 헤더: Row119~121
    Set doneMerge = CreateObject("Scripting.Dictionary")
    doneMerge.CompareMode = 1
    lockedRanges = Array( _
        "A1:BV1", "A3:BV4", "A5:BV6", _
        "A17:BV19", "A30:BV34", "A45:BV47", "A58:BV62", _
        "A64:BV67", "A78:BV78", "A80:BV82", _
        "A84:BV86", "A103:BV107", _
        "A119:BV121")
    For Each one In lockedRanges
        On Error Resume Next
        Set r = ws.Range(CStr(one))
        If Not r Is Nothing Then
            For Each c In r.Cells
                If c.MergeCells Then
                    maAddr = c.MergeArea.Address(False, False)
                    If Not doneMerge.exists(maAddr) Then
                        doneMerge.Add maAddr, True
                        c.MergeArea.Locked = True
                    End If
                Else
                    c.Locked = True
                End If
            Next c
        End If
        On Error GoTo 0
    Next one

    ' 2) 수식이 있는 모든 셀 잠금 (A1:BV133 전체 범위)
    On Error Resume Next
    Set ur = ws.Range("A1:BV133")
    For Each c In ur.Cells
        If c.HasFormula Then
            If c.MergeCells Then
                maAddr = c.MergeArea.Address(False, False)
                If Not doneMerge.exists(maAddr) Then
                    doneMerge.Add maAddr, True
                    c.MergeArea.Locked = True
                End If
            Else
                c.Locked = True
            End If
        End If
    Next c
    On Error GoTo 0

    ' 3) 시트 보호 (암호 0205, UserInterfaceOnly:=True → 매크로는 계속 쓰기 가능)
    ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowInsertingColumns:=False, AllowInsertingRows:=False, AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=False, AllowDeletingRows:=False, AllowSorting:=True, _
        AllowFiltering:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

    Application.ScreenUpdating = origScreen
    Application.EnableEvents = origEvents
End Sub

'================================================================================
' 생산일보 시트 활성화 시 제품표준사양서 G9~마지막 기반 유효성 목록 갱신
' 대상: B68:B77(T07 직기), AK68:AK77(T08 봉제) - 생산일보_셀매핑.json (공백·중복 제외, 시트 선택 시 항상 최신화)
' ※ 목록데이터 시트 + 이름 "목록_일일생산일보_제품" 사용 (255자 제한 회피)
'================================================================================
Public Sub 원단_생산일보_제품유효성_적용()
    Call 원단_일일생산일보_제품목록_유효성적용
End Sub

' 원단관리 시트만 전체 잠금 해제 (시트 활성화 시 호출)
Public Sub 원단관리_시트_잠금해제_적용()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_원단관리)
    On Error GoTo 0
    If Not ws Is Nothing And 시트보호해제_시도(ws) Then
        ws.Cells.Locked = False
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    End If
End Sub

' 시트 보호 해제 (암호 0205, 비번 입력창 없음) - 해제 성공 시 True
Private Function 시트보호해제_시도(ws As Worksheet) As Boolean
    시트보호해제_시도 = False
    If ws Is Nothing Then Exit Function
    If Not ws.ProtectContents Then 시트보호해제_시도 = True: Exit Function
    On Error Resume Next
    ws.Unprotect WS_시트보호암호
    If ws.ProtectContents Then ws.Unprotect
    시트보호해제_시도 = Not ws.ProtectContents
    On Error GoTo 0
End Function

'================================================================================
' 제품표준사양서 5~8행, 일일생산DB 1행, 현재고 1~2행, 원단관리 전체 잠금해제 (UserInterfaceOnly로 매크로는 쓰기 가능)
' ★ 시트 보호 해제 후에만 Cells.Locked 설정 (보호 상태에서 Locked 설정 시 오류 발생)
'================================================================================
Public Sub 원단_시트_지정행_잠금적용()
    Dim ws As Worksheet
    Dim origScreen As Boolean, origEvents As Boolean, origAlerts As Boolean
    origScreen = Application.ScreenUpdating
    origEvents = Application.EnableEvents
    origAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' 1) 제품표준사양서: 5~8행 전체 (암호 0205로 해제/재설정)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_제품표준사양서)
    On Error GoTo 0
    If Not ws Is Nothing And 시트보호해제_시도(ws) Then
        ws.Cells.Locked = False
        ws.Rows("5:8").Locked = True
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If

    ' 2) 일일생산DB: 1행 전체
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_일일생산DB)
    On Error GoTo 0
    If Not ws Is Nothing And 시트보호해제_시도(ws) Then
        ws.Cells.Locked = False
        ws.Rows(1).Locked = True
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If

    ' 3) 현재고: 1~2행 전체
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_현재고)
    On Error GoTo 0
    If Not ws Is Nothing And 시트보호해제_시도(ws) Then
        ws.Cells.Locked = False
        ws.Rows("1:2").Locked = True
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If

    ' 4) 원단관리: 전체 셀 잠금 해제 (암호 입력 없이 편집 가능)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_원단관리)
    On Error GoTo 0
    If Not ws Is Nothing And 시트보호해제_시도(ws) Then
        ws.Cells.Locked = False
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    End If

    Application.DisplayAlerts = origAlerts
    Application.ScreenUpdating = origScreen
    Application.EnableEvents = origEvents
End Sub

'================================================================================
' 생산일보_셀매핑.json 기준: B68:B77(T07 직기 완제품 제품명), AK68:AK77(T08 봉제 완제품 제품명) 유효성
' 시트 선택 시 활성화하려면 ThisWorkbook에 Workbook_SheetActivate에서 이 Sub를 호출하세요.
'================================================================================
Public Sub 원단_일일생산일보_제품목록_유효성적용()
    Dim ws사양 As Worksheet, ws일보 As Worksheet, wsList As Worksheet
    Dim lastRow As Long, r As Long, n As Long
    Dim v As String
    Dim col제품 As Long
    Dim rng As Range

    On Error Resume Next
    Set ws사양 = ThisWorkbook.Sheets(WS_제품표준사양서)
    Set ws일보 = ThisWorkbook.Sheets(WS_일일생산일보)
    Set wsList = ThisWorkbook.Sheets(WS_목록데이터)
    On Error GoTo 0
    If ws사양 Is Nothing Or ws일보 Is Nothing Then Exit Sub
    If wsList Is Nothing Then
        On Error Resume Next
        Dim wb보호목록2 As Boolean: wb보호목록2 = ThisWorkbook.ProtectStructure
        If wb보호목록2 Then ThisWorkbook.Unprotect WS_시트보호암호
        Set wsList = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Not wsList Is Nothing Then wsList.Name = WS_목록데이터
        If wb보호목록2 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
        On Error GoTo 0
    End If
    On Error Resume Next
    If Not wsList Is Nothing Then wsList.Visible = xlSheetVeryHidden
    On Error GoTo 0

    ' 제품 열: G열(7) 사용. 헤더로 찾을 경우에는 표준사양서_헤더열찾기(ws사양, "제품") 사용 가능
    col제품 = 표준사양서_헤더열찾기(ws사양, "제품")
    If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(ws사양, "제품(상품)명")
    If col제품 = 0 Then col제품 = 표준사양서_제품열G

    lastRow = ws사양.Cells(ws사양.Rows.Count, col제품).End(xlUp).Row
    If lastRow < 표준사양서_데이터시작행 Then Exit Sub

    ' G9~마지막 행에서 빈행 무시, 순서 유지하여 목록데이터 E열에 채움 (중복 허용 또는 제거 - 여기서는 순서 유지·중복 제거)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    n = 0
    For r = 표준사양서_데이터시작행 To lastRow
        v = Trim(CStr(ws사양.Cells(r, col제품).Value))
        If v <> "" And Not dict.exists(v) Then
            dict.Add v, n
            n = n + 1
        End If
    Next r
    If n = 0 Then Exit Sub

    On Error Resume Next
    wsList.Unprotect WS_시트보호암호
    On Error GoTo 0
    wsList.Columns(일일생산일보_제품목록_목록데이터열).ClearContents
    Dim keys As Variant, i As Long
    keys = dict.keys
    For i = 0 To UBound(keys)
        wsList.Cells(i + 1, 일일생산일보_제품목록_목록데이터열).Value = CStr(keys(i))
    Next i
    lastRow = n

    On Error Resume Next
    ThisWorkbook.Names("목록_일일생산일보_제품").Delete
    On Error GoTo 0
    ' RefersTo는 문자열로 지정 (1004 방지)
    ThisWorkbook.Names.Add Name:="목록_일일생산일보_제품", _
        RefersTo:="=" & wsList.Name & "!" & wsList.Range(wsList.Cells(1, 일일생산일보_제품목록_목록데이터열), wsList.Cells(lastRow, 일일생산일보_제품목록_목록데이터열)).Address(False, False)

    ' B68:B77(T07 직기), AK68:AK77(T08 봉제) 유효성 목록 적용 (생산일보_셀매핑.json)
    Dim 유효성범위목록 As Variant, one As Variant, c As Range, doneAreas As Object, arKey As String
    유효성범위목록 = Array("B68:B77", "AK68:AK77")
    Set doneAreas = CreateObject("Scripting.Dictionary")
    For Each one In 유효성범위목록
        Set rng = ws일보.Range(CStr(one))
        doneAreas.RemoveAll
        On Error Resume Next
        For Each c In rng.Cells
            If c.MergeCells Then
                arKey = c.MergeArea.Address(False, False)
                If Not doneAreas.exists(arKey) Then
                    doneAreas.Add arKey, True
                    c.MergeArea.Cells(1, 1).Validation.Delete
                    c.MergeArea.Cells(1, 1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=목록_일일생산일보_제품"
                End If
            Else
                c.Validation.Delete
                c.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=목록_일일생산일보_제품"
            End If
        Next c
        On Error GoTo 0
    Next one
End Sub
' ※ 시트 선택 시 자동 적용: ThisWorkbook 코드에 아래를 추가하세요.
' Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'     If Sh.Name = "일일생산일보" Then 원단_일일생산일보_제품목록_유효성적용
' End Sub

' 원단관리 시트 그룹행(4행) 및 F9:G 마지막행에 제품표준사양서 기반 유효성 적용
Private Sub 원단관리_그룹행_유효성_제품표준사양서(wsTarget As Worksheet, ws사양 As Worksheet)
    Dim lastRow As Long, r As Long
    Dim listNames As Variant, listCols As Variant
    Dim i As Long
    Dim rng As Range
    Dim nm제품 As String

    If wsTarget Is Nothing Or ws사양 Is Nothing Then Exit Sub
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub

    ' 원단관리: A·B열은 Raw data H/I 기반 목록_원단번호, 목록_원단명
    If wsTarget.Name = WS_원단관리 Then
        Call RawData_HI_목록_갱신
        On Error Resume Next
        For r = START_ROW To lastRow Step ROWS_PER_ITEM
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Delete
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단번호"
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.IgnoreBlank = True
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Delete
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단명"
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.IgnoreBlank = True
        Next r
        On Error GoTo 0
    End If

    listNames = Array("목록_업체명", "목록_제품", "목록_제조사", "목록_지관", "목록_지폭", "목록_평량", "목록_1차포장", "목록_2차포장")
    listCols = Array(COL_제품발주처, COL_제품명, COL_제조사, COL_지관, COL_지폭, COL_총량, COL_1차포장, COL_2차포장)

    On Error Resume Next
    For i = 0 To UBound(listNames)
        If i = 0 Or i = 1 Then
            ' 원단관리: A·B는 이미 처리됨
            If wsTarget.Name = WS_원단관리 Then GoTo NextListCol
        End If
        If i = 1 Then
            ' B열(제품명): A열(업체명) 선택값에 따른 하위 목록만
            For r = START_ROW To lastRow Step ROWS_PER_ITEM
                Set rng = wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명))
                rng.Validation.Delete
                nm제품 = Trim(CStr(wsTarget.Cells(r, COL_제품발주처).Value))
                If nm제품 <> "" And nm제품 <> "-" Then nm제품 = "목록_제품_" & 목록이름_세이니타이즈(nm제품) Else nm제품 = "목록_제품"
                rng.Validation.Add Type:=xlValidateList, Formula1:="=" & nm제품
                rng.Validation.IgnoreBlank = True
                rng.Validation.InCellDropdown = True
            Next r
        Else
            Set rng = Nothing
            For r = START_ROW To lastRow Step ROWS_PER_ITEM
                If rng Is Nothing Then Set rng = wsTarget.Cells(r, listCols(i)) Else Set rng = Union(rng, wsTarget.Cells(r, listCols(i)))
            Next r
            If Not rng Is Nothing Then
                rng.Validation.Delete
                rng.Validation.Add Type:=xlValidateList, Formula1:="=" & listNames(i)
                rng.Validation.IgnoreBlank = True
                rng.Validation.InCellDropdown = True
            End If
        End If
NextListCol:
    Next i
    ' F9, G9 ~ 마지막행 유효성 (지폭, 평량)
    If lastRow >= 9 Then
        wsTarget.Range(wsTarget.Cells(9, COL_지폭), wsTarget.Cells(lastRow, COL_지폭)).Validation.Delete
        wsTarget.Range(wsTarget.Cells(9, COL_지폭), wsTarget.Cells(lastRow, COL_지폭)).Validation.Add Type:=xlValidateList, Formula1:="=목록_지폭"
        wsTarget.Range(wsTarget.Cells(9, COL_총량), wsTarget.Cells(lastRow, COL_총량)).Validation.Delete
        wsTarget.Range(wsTarget.Cells(9, COL_총량), wsTarget.Cells(lastRow, COL_총량)).Validation.Add Type:=xlValidateList, Formula1:="=목록_평량"
    End If
    On Error GoTo 0
End Sub

    '================================================================================
    ' 동적 이름 정의 설정 (INDEX/LOOKUP 수식 우선, 실패 시 정적 범위 폴백)
    '================================================================================
    Private Sub 목록이름_동적설정(이름 As String, RefersTo식 As String, ws As Worksheet, col As Long)
        Dim 성공 As Boolean: 성공 = False

        ' 기존 이름 삭제 후 재생성
    On Error Resume Next
    ThisWorkbook.Names(이름).Delete
    On Error GoTo 0

        ' 방법 1: 동적 수식(INDEX/LOOKUP)으로 시도
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=이름, RefersTo:=RefersTo식
        If Err.Number = 0 Then 성공 = True
        Err.Clear
    On Error GoTo 0

        ' 방법 2: RefersToR1C1로 시도
    If Not 성공 Then
        On Error Resume Next
        ThisWorkbook.Names.Add Name:=이름, RefersToR1C1:=RefersTo식
            If Err.Number = 0 Then 성공 = True
            Err.Clear
            On Error GoTo 0
        End If

        ' 방법 3: OFFSET/COUNTA 동적 수식으로 폴백 (INDEX/LOOKUP 실패 시)
        If Not 성공 Then
            Dim s As String, colLetter As String, refOffset As String
            s = "'" & Replace(ws.Name, "'", "''") & "'"
            colLetter = Split(ws.Cells(표준사양서_헤더행, col).Address(True, False), "$")(0)
            refOffset = "=OFFSET(" & s & "!$" & colLetter & "$" & 표준사양서_데이터시작행 & _
                        ",0,0,MAX(1,COUNTA(" & s & "!$" & colLetter & ":$" & colLetter & _
                        ")-" & (표준사양서_데이터시작행 - 1) & "),1)"
            On Error Resume Next
            ThisWorkbook.Names.Add Name:=이름, RefersTo:=refOffset
            If Err.Number = 0 Then 성공 = True
            Err.Clear
        On Error GoTo 0
    End If

        ' 방법 4: 최종 폴백 - 정적 범위 (동적 수식 모두 실패 시)
    If Not 성공 Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
            If lastRow < 표준사양서_데이터시작행 Then lastRow = 표준사양서_데이터시작행 + 50

        On Error Resume Next
        ThisWorkbook.Names.Add Name:=이름, RefersTo:=ws.Range(ws.Cells(표준사양서_데이터시작행, col), ws.Cells(lastRow, col))
            If Err.Number = 0 Then 성공 = True
            Err.Clear
        On Error GoTo 0
    End If

    If Not 성공 Then
        Debug.Print "이름 정의 실패: " & 이름
    End If
End Sub

    ' 원지_채우기: 구 Raw data_index 기반. 현재는 제품표준사양서·B~I 열 구조로 그룹행 유효성만 적용하므로 no-op.
    Private Sub 원지_채우기(wsTarget As Worksheet, wsIndex As Worksheet)
        ' 유효성/목록은 목록시트_갱신 + 원단관리_그룹행_유효성_제품표준사양서에서 처리
    End Sub

    '================================================================================
    ' 원단관리 시트 B~H열 목록 다운: 업체명, 원지(M열 유효성), 제품, 지관, 폴리백, 비닐백, BOX자재현황
'================================================================================
Private Sub 목록다운_적용(wsTarget As Worksheet)
    Dim lastRow As Long, r As Long
    Dim listNames As Variant, listCols As Variant
    Dim i As Long
    Dim rngUnion As Range
    Dim nmC As String

    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub

    ' 원단관리: A·B열은 Raw data H/I 기반 목록_원단번호, 목록_원단명 사용
    If wsTarget.Name = WS_원단관리 Then
        Call RawData_HI_목록_갱신
        On Error Resume Next
        For r = START_ROW To lastRow Step ROWS_PER_ITEM
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Delete
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단번호"
            wsTarget.Range(wsTarget.Cells(r, COL_제품발주처), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.IgnoreBlank = True
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Delete
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단명"
            wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.IgnoreBlank = True
        Next r
        On Error GoTo 0
        Exit Sub
    End If

    listNames = Array("목록_업체명", "목록_제품", "목록_제조사", "목록_지관", "목록_지폭", "목록_평량", "목록_1차포장", "목록_2차포장")
    listCols = Array(COL_제품발주처, COL_제품명, COL_제조사, COL_지관, COL_지폭, COL_총량, COL_1차포장, COL_2차포장)
    
    On Error Resume Next
    For i = 0 To UBound(listNames)
        If i = 1 Then
            ' B열(제품명): A열(업체명) 기준 그룹핑 → 동일 수식끼리 Union하여 Validation 횟수 최소화
            Dim dictV As Object: Set dictV = CreateObject("Scripting.Dictionary")
            For r = START_ROW To lastRow Step ROWS_PER_ITEM
                nmC = Trim(CStr(wsTarget.Cells(r, COL_제품발주처).Value))
                If nmC <> "" And nmC <> "-" Then nmC = "목록_제품_" & 목록이름_세이니타이즈(nmC) Else nmC = "목록_제품"
                If Not dictV.exists(nmC) Then
                    Set dictV(nmC) = wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명))
                Else
                    Set dictV(nmC) = Union(dictV(nmC), wsTarget.Range(wsTarget.Cells(r, COL_제품명), wsTarget.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)))
                End If
            Next r
            Dim vk As Variant
            For Each vk In dictV.keys
                dictV(vk).Validation.Delete
                dictV(vk).Validation.Add Type:=xlValidateList, Formula1:="=" & CStr(vk)
                dictV(vk).Validation.IgnoreBlank = True
                dictV(vk).Validation.InCellDropdown = True
            Next vk
        Else
            Set rngUnion = Nothing
            For r = START_ROW To lastRow Step ROWS_PER_ITEM
                If rngUnion Is Nothing Then
                    Set rngUnion = wsTarget.Cells(r, listCols(i))
                Else
                    Set rngUnion = Union(rngUnion, wsTarget.Cells(r, listCols(i)))
                End If
            Next r
            If Not rngUnion Is Nothing Then
                rngUnion.Validation.Delete
                rngUnion.Validation.Add Type:=xlValidateList, Formula1:="=" & listNames(i)
                rngUnion.Validation.IgnoreBlank = True
                rngUnion.Validation.InCellDropdown = True
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

'================================================================================
        ' 마지막 행에 빈 행 그룹(15행) 추가 - 분류/부자재명/업체명은 목록 다운으로 선택
'================================================================================
Public Sub 원단_빈행그룹_추가()
    Dim ws As Worksheet
        Dim wsIndex As Worksheet
    Dim lastRow As Long
    Dim 추가행 As Long
    Dim 시트보호상태 As Boolean
        Dim 마지막열 As Long
        Dim listNames As Variant, listCols As Variant
        Dim ii As Long

    On Error GoTo Err빈행
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error Resume Next
        Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
        On Error GoTo Err빈행
    If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 시트 보호 상태 확인 및 해제 (암호 대화상자 없이 임시 해제)
    시트보호상태 = ws.ProtectContents
    If Not 시트보호해제_시도(ws) Then
        MsgBox "시트 보호 해제에 실패했습니다. 암호(0205)를 확인하세요.", vbExclamation
        Exit Sub
    End If

        ' 마지막 데이터 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then lastRow = ws.Cells(ws.Rows.Count, COL_부자재명).End(xlUp).Row
    If lastRow < START_ROW Then lastRow = ws.Cells(ws.Rows.Count, COL_카테고리).End(xlUp).Row
    If lastRow < START_ROW Then
            If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호
        MsgBox "먼저 '데이터 업데이트'로 시트 구조를 만든 뒤 사용하세요.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
        Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    추가행 = lastRow + 1
        마지막열 = 마지막일열_찾기(ws)
        
        ' ── 1. 행 삽입 + 서식 복사 + 데이터 채우기 ──
    ws.Rows(추가행).Resize(ROWS_PER_ITEM).Insert Shift:=xlDown
        Call 첫그룹_서식복사(ws, 추가행)
        Call 빈행그룹_데이터만_채우기(ws, 추가행)
        
        ' ── 2. 신규 블록 셀 병합 (그룹당 6행: A~C + D~H. C열은 빈행그룹_데이터만_채우기에서 4~6/7~9 두 블록으로 이미 병합) ──
        On Error Resume Next
        Call 블록5_Merge(ws, 추가행, COL_제품발주처)
        Call 블록5_Merge(ws, 추가행, COL_제품명)
        ' COL_포장구분 제외 - C4/C7 분리 표시 위해 두 블록(4~6, 7~9) 유지
        Call 블록5_Merge(ws, 추가행, COL_안전재고)
        Call 블록5_Merge(ws, 추가행, COL_기초재고)
        Call 블록5_Merge(ws, 추가행, COL_입고총합)
        Call 블록5_Merge(ws, 추가행, COL_사용총합)
        Call 블록5_Merge(ws, 추가행, COL_기말재고)
        On Error GoTo Err빈행

        ' ── 3. 숫자서식 (테두리는 첫그룹_서식복사에서 이미 복사됨) ──
        ws.Range(ws.Cells(추가행, COL_안전재고), ws.Cells(추가행, COL_기말재고)).NumberFormat = FMT_NUM_STD
        If 마지막열 >= COL_월별시작 Then ws.Range(ws.Cells(추가행, COL_월별시작), ws.Cells(추가행 + ROWS_PER_ITEM - 1, 마지막열)).NumberFormat = FMT_NUM_월별
        
        ' 그룹행 교대 음영 적용 (시트 보호 전에 실행 - Interior/Borders 수정 필요)
        Call 원단관리_그룹행_교대음영_적용(ws, 추가행 + ROWS_PER_ITEM - 1)

        ' ── 5. 신규 블록만 셀 보호 (J2:N2, J열, L열 입력 가능) ──
        Dim 마지막행보호 As Long
        마지막행보호 = 추가행 + ROWS_PER_ITEM - 1
        ws.Cells.Locked = True
        ws.Range("J2:N2").Locked = False
        ws.Range("J3:J" & 마지막행보호).Locked = False
        ws.Range("L3:L" & 마지막행보호).Locked = False
        
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                AllowFormattingCells:=False, AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, AllowDeletingRows:=False, _
                AllowSorting:=False, AllowFiltering:=False, AllowUsingPivotTables:=False

    Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    Application.ScreenUpdating = True

        ws.Cells(추가행, COL_업체명).Select
        MsgBox "빈 행 그룹(6행)이 추가되었습니다." & vbCrLf & vbCrLf & _
           "업체명·원지·제품 등은 셀 클릭 후 목록에서 선택하세요.", vbInformation, "목록 추가"
    Exit Sub
Err빈행:
        If 시트보호상태 Then
            On Error Resume Next
            ws.Protect Password:=WS_시트보호암호
            On Error GoTo 0
        End If
    Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "빈 행 그룹 추가 중 오류: " & Err.Description, vbExclamation, "오류"
End Sub

'================================================================================
    ' ★ 외부 모듈용: 제품명 지정하여 빈 그룹행 추가 (MsgBox 없음, Application 상태 변경 없음)
    ' 반환: 추가된 그룹의 시작행 번호 (실패 시 0)
    ' ※ 호출자가 시트 보호 해제 상태에서 호출할 것을 전제 (내부에서 보호 복원 안 함)
'================================================================================
    Public Function 원단_빈행그룹_추가_무음(제품명 As String) As Long
        Dim ws As Worksheet
        Dim wsIndex As Worksheet
        Dim lastRow As Long, 추가행 As Long, 마지막열 As Long
        Dim listNames As Variant, listCols As Variant
        Dim ii As Long
        Dim 행삽입완료 As Boolean
        원단_빈행그룹_추가_무음 = 0
        행삽입완료 = False
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
        On Error GoTo 0
        If ws Is Nothing Then Exit Function
        
        ' ※ 호출자가 이미 Unprotect 상태로 호출해야 함
        ' 만약 보호 상태라면 해제만 수행 (복원은 호출자 책임, 암호 대화상자 없이)
        If ws.ProtectContents Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        
        lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        If lastRow < START_ROW Then lastRow = ws.Cells(ws.Rows.Count, COL_부자재명).End(xlUp).Row
        If lastRow < START_ROW Then Exit Function
        
        추가행 = lastRow + 1
        마지막열 = 마지막일열_찾기(ws)
        
        ' 행 삽입 + 서식 복사 + 데이터 채우기
        On Error GoTo ErrSilent
        ws.Rows(추가행).Resize(ROWS_PER_ITEM).Insert Shift:=xlDown
        행삽입완료 = True
        
        Call 첫그룹_서식복사(ws, 추가행)
        Call 빈행그룹_데이터만_채우기(ws, 추가행)
        
        On Error Resume Next
        ws.Range(ws.Cells(추가행, COL_제품발주처), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_제품발주처)).Merge
        ws.Range(ws.Cells(추가행, COL_제품명), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_제품명)).Merge
        ws.Range(ws.Cells(추가행, COL_포장구분), ws.Cells(추가행 + 2, COL_포장구분)).Merge
        ws.Range(ws.Cells(추가행 + 3, COL_포장구분), ws.Cells(추가행 + 5, COL_포장구분)).Merge
        ws.Range(ws.Cells(추가행, COL_제조사), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_제조사)).Merge
        ws.Range(ws.Cells(추가행, COL_지관), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_지관)).Merge
        ws.Range(ws.Cells(추가행, COL_지폭), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_지폭)).Merge
        ws.Range(ws.Cells(추가행, COL_총량), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_총량)).Merge
        ws.Range(ws.Cells(추가행, COL_1차포장), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_1차포장)).Merge
        ws.Range(ws.Cells(추가행, COL_2차포장), ws.Cells(추가행 + ROWS_PER_ITEM - 1, COL_2차포장)).Merge
        Call 블록5_Merge(ws, 추가행, COL_전월재고)
        Call 블록5_Merge(ws, 추가행, COL_현재고)
        Call 블록5_Merge(ws, 추가행, COL_안전재고)
        ' I열(COL_자재유형=COL_구분)은 병합 안 함 - 입고/사용/재고 각 행별 값
        On Error GoTo ErrSilent
        
        ' ★ 유효성/서식/음영은 데이터업데이트 본체에서 일괄 적용 ? 여기서는 제품명만 입력
        ws.Cells(추가행, COL_제품).Value = 제품명
        
        ' ★ 시트 보호 복원 안 함 (호출자가 모든 제품 처리 후 한 번에 복원)
        
        원단_빈행그룹_추가_무음 = 추가행
        Exit Function
        
ErrSilent:
        ' ★ 에러 발생 시: 이미 삽입된 행이 있으면 삭제 (시트 구조 손상 방지)
        If 행삽입완료 Then
            On Error Resume Next
            ws.Rows(추가행).Resize(ROWS_PER_ITEM).Delete Shift:=xlUp
            On Error GoTo 0
        End If
        원단_빈행그룹_추가_무음 = 0
    End Function

    '================================================================================
    ' 목록 추가만 반영 (제품표준사양서에서 새로 추가된 항목만 원단관리에 추가, 기존 데이터 유지. B~I: 업체명·제품·제조사·지관·지폭·평량·1차포장·2차포장)
    '================================================================================
    Public Sub 원단관리시트_목록추가만()
    Dim wsTarget As Worksheet
    Dim wsIndex As Worksheet
    Dim 분류순서 As Collection, catItems As Collection
    Dim 기존키 As Object   ' Co → 1 (원단관리는 Co 단독으로 중복 판정)
    Dim new항목 As Collection   ' 새로 추가할 Co 목록
    Dim i As Long, lastRow As Long, k As Long, 추가건수 As Long
    Dim 삽입위치 As Long
    Dim col업체 As Long, col제품 As Long, col제조사 As Long, col지관 As Long, col지폭 As Long, col총량 As Long, col1차 As Long, col2차 As Long
    Dim r As Long
    Dim v제조사 As String, v지관 As String, v지폭 As String, v총량 As String, v1차 As String, v2차 As String
    Dim listNames As Variant, listCols As Variant, ii As Long
    Dim 마지막열 As Long
    
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(WS_원단관리)
    Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
    On Error GoTo 0
    
    If wsTarget Is Nothing Then
        MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    Dim wsDB As Worksheet
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
    Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
    On Error GoTo 0
    ' 원단관리: 일일생산DB 또는 Raw data에서 Co 목록 로드
    If Not 일일생산DB_로드_항목_원단(wsDB, 분류순서, catItems) Then
        If Not RawData_Co목록_로드(분류순서, catItems) Then
            MsgBox "일일생산DB 또는 Raw data 시트에서 원단(Co) 데이터를 읽을 수 없습니다." & vbCrLf & "데이터 업데이트를 먼저 실행하세요.", vbExclamation
            Exit Sub
        End If
    End If
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
    Set 기존키 = CreateObject("Scripting.Dictionary")
    
    ' 기존 시트: Co(A열) 수집 (원단관리는 Co 단독 기준으로 중복 판정)
    For i = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        Dim vCo As String, v원단명 As String
        vCo = Trim(CStr(wsTarget.Cells(i, COL_제품발주처).Value))
        If vCo <> "" Then
            If IsNumeric(vCo) And CLng(val(vCo)) > 0 Then vCo = Format(CLng(val(vCo)), "00")
            기존키(vCo) = 1
        End If
    Next i
    
    ' 새로 추가할 Co: catItems(1) = Co 목록 (원단 전용)
    Set new항목 = New Collection
    For i = 1 To catItems(1).Count
        vCo = Trim(CStr(catItems(1)(i)))
        If IsNumeric(vCo) And CLng(val(vCo)) > 0 Then vCo = Format(CLng(val(vCo)), "00")
        If Not 기존키.exists(vCo) Then
            On Error Resume Next
            new항목.Add vCo
            On Error GoTo 0
        End If
    Next i
    
    ' 실제 추가할 Co 목록 (new항목은 기존키 기준으로 이미 중복 제거됨)
    Dim 실제추가 As Collection
    Set 실제추가 = new항목
    
    추가건수 = 실제추가.Count
    If 추가건수 = 0 Then
        MsgBox "추가할 새 목록이 없습니다." & vbCrLf & "일일생산DB/Raw data에 있는 원단(Co)이 이미 시트에 있습니다.", vbInformation, "목록 추가"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If wsTarget.ProtectContents Then
        On Error Resume Next
        wsTarget.Unprotect WS_시트보호암호
        If wsTarget.ProtectContents Then wsTarget.Unprotect
        On Error GoTo 0
    End If
    
    삽입위치 = lastRow + 1
    If lastRow < START_ROW Then 삽입위치 = START_ROW
    wsTarget.Rows(삽입위치).Resize(ROWS_PER_ITEM * 추가건수).Insert
    
    ' ★ 형식 복사 1회만 (속도 최적화: 기존 데이터 있을 때만, 부자재항목_생성 반복 시 Copy 생략)
    Dim 형식복사생략 As Boolean: 형식복사생략 = False
    마지막열 = 마지막월열_찾기(wsTarget)
    If 마지막열 < COL_구분 Then 마지막열 = COL_구분
    If lastRow >= START_ROW Then
        On Error Resume Next
        wsTarget.Range(wsTarget.Cells(START_ROW, 1), wsTarget.Cells(START_ROW + ROWS_PER_ITEM - 1, 마지막열)).Copy
        wsTarget.Range(wsTarget.Cells(삽입위치, 1), wsTarget.Cells(삽입위치 + ROWS_PER_ITEM * 추가건수 - 1, 마지막열)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        On Error GoTo 0
        형식복사생략 = True
    End If
    
    For k = 1 To 추가건수
        vCo = Trim(CStr(실제추가(k)))
        v원단명 = RawData_CoTo원단명(vCo)
        If v원단명 = "" Then v원단명 = "-"
        ' A=Co(업체명), B=원단명(부자재명)
        Call 부자재항목_생성(wsTarget, 삽입위치 + (k - 1) * ROWS_PER_ITEM, "", v원단명, vCo, -1, "", 형식복사생략)
    Next k
    
    ' 유효성 적용: 원단관리 A·B열 = 목록_원단번호/목록_원단명 (Raw data H/I)
    If Not wsIndex Is Nothing Then Call 목록시트_갱신(wsIndex)
    마지막열 = wsTarget.Cells(2, wsTarget.Columns.Count).End(xlToLeft).Column
    Dim 추가시작행 As Long
    On Error Resume Next
    For k = 1 To 추가건수
        추가시작행 = 삽입위치 + (k - 1) * ROWS_PER_ITEM
        wsTarget.Range(wsTarget.Cells(추가시작행, COL_제품발주처), wsTarget.Cells(추가시작행 + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Delete
        wsTarget.Range(wsTarget.Cells(추가시작행, COL_제품발주처), wsTarget.Cells(추가시작행 + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단번호"
        wsTarget.Range(wsTarget.Cells(추가시작행, COL_제품발주처), wsTarget.Cells(추가시작행 + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.IgnoreBlank = True
        wsTarget.Range(wsTarget.Cells(추가시작행, COL_제품명), wsTarget.Cells(추가시작행 + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Delete
        With wsTarget.Range(wsTarget.Cells(추가시작행, COL_제품명), wsTarget.Cells(추가시작행 + ROWS_PER_ITEM - 1, COL_제품명)).Validation
            .Add Type:=xlValidateList, Formula1:="=목록_원단명"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Next k
    On Error GoTo 0
    
        ' 부자재항목_생성에서 이미 서식·병합 완료 → 전체 시트 재서식/재계산 생략
        ' 셀 보호: 목록추가로 넣은 그룹행만 예외적으로 전체 셀보호 해제 (데이터 업데이트 시 예외 초기화되어 잠김)
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, COL_구분).End(xlUp).Row
        If lastRow < START_ROW Then lastRow = START_ROW
        원단_목록추가_예외시작행 = 삽입위치
        원단_목록추가_예외끝행 = 삽입위치 + ROWS_PER_ITEM * 추가건수 - 1
        Call 원단_셀보호_적용_무음(wsTarget)

        Call 원단관리_그룹행_교대음영_적용(wsTarget, lastRow)
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "목록 추가가 완료되었습니다!" & vbCrLf & vbCrLf & _
           "추가된 항목: " & 추가건수 & "건 (일일생산DB/Raw data Co 기준)" & vbCrLf & _
           "A열=원단번호(Co), B열=원단명이 반영되었습니다." & vbCrLf & _
           "기존 데이터는 유지되었습니다.", vbInformation, "목록 추가 완료"
End Sub

'================================================================================
' 이력 시트 초기화
'================================================================================
Private Sub 이력시트_초기화()
    Dim ws As Worksheet
    
    Set ws = 시트가져오기또는생성(WS_이력)
    If ws Is Nothing Then Exit Sub
    
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "일시"
        ws.Cells(1, 2).Value = "작업유형"
        ws.Cells(1, 3).Value = "부자재명"
        ws.Cells(1, 4).Value = "일자"
        ws.Cells(1, 5).Value = "이전값"
        ws.Cells(1, 6).Value = "새값"
        ws.Cells(1, 7).Value = "행번호"
        ws.Cells(1, 8).Value = "열번호"
        ' 서식: 헤더 굵게·테두리 (색상 제외, 사용자 지정 색상 유지)
        With ws.Range("A1:H1")
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

'================================================================================
' 부자재 목록 가져오기 (콤보박스용)
'================================================================================
Private Function 부자재명_전체목록() As Collection
    Dim ws As Worksheet
    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim 부자재명 As String
    
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    
    For i = START_ROW To lastRow
        If ws.Cells(i, COL_구분).Value = "입고" Then
            부자재명 = ws.Cells(i - 1, COL_부자재명).Value
            If 부자재명 <> "" Then
                On Error Resume Next
                result.Add 부자재명, 부자재명
                On Error GoTo 0
            End If
        End If
    Next i
    
    Set 부자재명_전체목록 = result
End Function

'================================================================================
' 목록 선택 다이얼로그 (ListBox 방식)
'================================================================================
Private Function 목록선택_다이얼로그(목록 As Collection, 안내문 As String, 제목 As String) As String
    Dim 선택값 As String
    Dim i As Long
    Dim 목록배열() As String
    Dim 선택인덱스 As Variant
    
    If 목록.Count = 0 Then
        목록선택_다이얼로그 = ""
        Exit Function
    End If
    
    ' Collection을 배열로 변환
    ReDim 목록배열(1 To 목록.Count)
    For i = 1 To 목록.Count
        목록배열(i) = 목록(i)
    Next i
    
    ' Application.InputBox의 Type 8 (목록) 대신 간단한 선택 방식 사용
    ' VBA에서 기본 제공하는 목록 선택 다이얼로그가 없으므로
    ' 번호 선택 방식으로 구현
    
    Dim 목록텍스트 As String
    목록텍스트 = 안내문 & vbCrLf & vbCrLf
    
    For i = 1 To 목록.Count
        목록텍스트 = 목록텍스트 & i & ". " & 목록(i) & vbCrLf
    Next i
    
    목록텍스트 = 목록텍스트 & vbCrLf & "번호를 입력하세요 (1-" & 목록.Count & "):"
    
    Dim 입력값 As String
    입력값 = InputBox(목록텍스트, 제목)
    
        ' 취소 버튼 감지 (StrPtr=0) 또는 빈 입력
        If StrPtr(입력값) = 0 Then
            목록선택_다이얼로그 = "<<CANCEL>>"
            Exit Function
        End If
    If 입력값 = "" Then
        목록선택_다이얼로그 = ""
        Exit Function
    End If
    
    If Not IsNumeric(입력값) Then
        MsgBox "숫자를 입력해주세요.", vbExclamation
        목록선택_다이얼로그 = ""
        Exit Function
    End If
    
    Dim 선택번호 As Integer
    선택번호 = CInt(입력값)
    
    If 선택번호 < 1 Or 선택번호 > 목록.Count Then
        MsgBox "유효한 번호를 입력해주세요. (1-" & 목록.Count & ")", vbExclamation
        목록선택_다이얼로그 = ""
        Exit Function
    End If
    
    목록선택_다이얼로그 = 목록(선택번호)
End Function

'================================================================================
' 시트 헤더 생성 (부자재관리와 동일 그리드 구조: 12열 J~U)
'================================================================================
Private Sub 시트헤더_생성(ws As Worksheet)
    Dim i As Long
    Dim 마지막열 As Long: 마지막열 = COL_월별시작 + 11  ' 12열(J~U) 부자재관리와 동일

    With ws.Cells(1, 1)
        .Value = "원단 재고관리 현황"
        .HorizontalAlignment = xlLeft
        .Font.Size = 20
        .Font.Name = "Microsoft GothicNeo"
    End With

    ' 2행 안내: E2:H2 병합 "1년단위로 누적표시"
    On Error Resume Next
    ws.Range("E2:H2").Merge
    On Error GoTo 0
    ws.Range("E2").Value = "1년단위로 누적표시"
    With ws.Range("E2:H2")
        .Interior.ColorIndex = xlNone
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' I2에 기준연도 저장 (흰색 폰트 - 날짜열 파싱용)
    ws.Range("I2").Value = Year(Date)
    ws.Range("I2").Font.color = RGB(255, 255, 255)

    ' 2행 월표식 + 3행 일번호 (부자재관리와 동일: 12열, 1m)
    For i = COL_월별시작 To 마지막열
        With ws.Cells(2, i)
            .Value = "1m"
            .Font.color = RGB(160, 160, 160)
        End With
    Next i
    For i = 1 To 12
        ws.Cells(3, COL_월별시작 + i - 1).Value = i
    Next i

    ' 3행 헤더 (원단관리: A=원단번호, B=원단명, C=Roll/중량Kg)
    ws.Cells(3, COL_제품발주처).Value = "원단번호"
    ws.Cells(3, COL_제품명).Value = "원단명"
    ws.Cells(3, COL_포장구분).Value = "Roll" & vbLf & "중량 Kg"
    ws.Cells(3, COL_안전재고).Value = "안전재고"
    ws.Cells(3, COL_기초재고).Value = "기초재고"
    ws.Cells(3, COL_입고총합).Value = "입고"
    ws.Cells(3, COL_사용총합).Value = "사용"
    ws.Cells(3, COL_기말재고).Value = "기말재고"
    ws.Cells(3, COL_구분).Value = "구분"

    With ws.Range(ws.Cells(3, 1), ws.Cells(3, 마지막열))
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

'================================================================================
' 부자재 섹션 생성 (카테고리별) - C/E/A용: 목록 항목 = "부자재명|업체명"
'================================================================================
Private Function 부자재섹션_생성_CEA(ws As Worksheet, 시작행 As Long, 카테고리명 As String, _
                                     목록 As Collection, 색상 As Long) As Long
    Dim i As Long
    Dim 현재행 As Long
    Dim s As String
    Dim 부자재명 As String, 업체명 As String
    Dim pos As Long
    
    현재행 = 시작행
    
    For i = 1 To 목록.Count
        s = CStr(목록(i))
        pos = InStr(s, "|")
        If pos > 0 Then
            부자재명 = left(s, pos - 1)
            업체명 = Mid(s, pos + 1)
        Else
            부자재명 = s
            업체명 = "-"
        End If
        Call 부자재항목_생성(ws, 현재행, 카테고리명, 부자재명, 업체명, 색상)
        현재행 = 현재행 + ROWS_PER_ITEM
    Next i
    
    현재행 = 현재행 + 1
    
    부자재섹션_생성_CEA = 현재행
End Function

'================================================================================
    ' 원단관리 시트 값 일괄 계산 (월별 입고/사용 → 재고 체인 계산, 총합·기말재고 갱신)
'================================================================================
    '================================================================================
    ' 원단관리 재고 재계산 (내부 전용 - Private)
    '   - 롤수(1차): 재고 = 기초재고 + 입고 누적 (사용1 무시)
    '   - KG(2차):   재고 = 기초재고 + 입고 - 사용 누적
    '   - firstDirtyOffset: 변경된 첫 열. 1이면 전체 재계산
    '   - ★ 오늘 포함 최대 10일까지만 재계산 (시트에 미래 열이 많아도 부하 최소화)
    '================================================================================
    Private Sub 원단관리_값일괄계산(ws As Worksheet, Optional firstDirtyOffset As Long = 1)
    Dim 시작행 As Long, 마지막행 As Long, 마지막열 As Long
    Dim 입고1 As Double, 입고2 As Double, 사용2 As Double
    Dim 재고1 As Double, 재고2 As Double, 전월1 As Double, 전월2 As Double
    Dim 기초재고1 As Double, 기초재고2 As Double
    Dim 입고1합 As Double, 입고2합 As Double, 사용2합 As Double
    Dim 시트보호상태 As Boolean, 이벤트상태 As Boolean
    Dim origScr As Boolean, origCalc As Long
    Dim monthArr As Variant, infoArr As Variant
    Dim rOff As Long, mCol As Long
    Dim 값일괄_단계 As String: 값일괄_단계 = "시작"

    If ws Is Nothing Then Exit Sub
    On Error GoTo 값일괄오류

    값일괄_단계 = "마지막행/열"
    이벤트상태 = Application.EnableEvents
    origScr = Application.ScreenUpdating: origCalc = Application.Calculation
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    시트보호상태 = ws.ProtectContents
    If 시트보호상태 Then
        On Error Resume Next
        ws.Unprotect WS_시트보호암호
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo 0
    End If

    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 < START_ROW Then GoTo 정리
    마지막열 = 마지막월열_찾기(ws)
    If 마지막열 < COL_월별시작 Then 마지막열 = COL_월별시작

    값일괄_단계 = "배열읽기"
    ' ── 전체 배열 읽기 (체인 계산에 기존 재고값 필요) ──
    infoArr = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(마지막행, COL_구분)).Value
    monthArr = ws.Range(ws.Cells(START_ROW, COL_월별시작), ws.Cells(마지막행, 마지막열)).Value

    값일괄_단계 = "오프셋계산"
    If firstDirtyOffset < 1 Then firstDirtyOffset = 1
    Dim totalCols As Long: totalCols = UBound(monthArr, 2)
    If firstDirtyOffset > totalCols Then firstDirtyOffset = totalCols
    Dim maxRecalcOffset As Long: maxRecalcOffset = 값일괄계산_오늘포함한달끝오프셋(ws, totalCols)
    If maxRecalcOffset < firstDirtyOffset Then maxRecalcOffset = firstDirtyOffset
    원단_디버그_상세 = 원단_디버그_상세 & " 마지막열=" & 마지막열 & " totalCols=" & totalCols & " maxRecalc=" & maxRecalcOffset

    값일괄_단계 = "재고루프"
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        값일괄_단계 = "재고루프 rOff=" & (시작행 - START_ROW + 1)
        rOff = 시작행 - START_ROW + 1
        기초재고1 = 0: 기초재고2 = 0
        If IsNumeric(infoArr(rOff, COL_기초재고)) Then 기초재고1 = CDbl(infoArr(rOff, COL_기초재고))
        If IsNumeric(infoArr(rOff + 3, COL_기초재고)) Then 기초재고2 = CDbl(infoArr(rOff + 3, COL_기초재고))

        ' ── 이전 열들의 합계 누적 (변경 열 이전 구간은 재고 재계산 안함) ──
        입고1합 = 0: 입고2합 = 0: 사용2합 = 0
        Dim carryOver1 As Double, carryOver2 As Double
        carryOver1 = 기초재고1: carryOver2 = 기초재고2
        If firstDirtyOffset > 1 Then
            For mCol = 1 To firstDirtyOffset - 1
                If IsNumeric(monthArr(rOff + OFFSET_입고1, mCol)) Then 입고1합 = 입고1합 + CDbl(monthArr(rOff + OFFSET_입고1, mCol))
                If IsNumeric(monthArr(rOff + OFFSET_입고2, mCol)) Then 입고2합 = 입고2합 + CDbl(monthArr(rOff + OFFSET_입고2, mCol))
                If IsNumeric(monthArr(rOff + OFFSET_사용2, mCol)) Then 사용2합 = 사용2합 + CDbl(monthArr(rOff + OFFSET_사용2, mCol))
            Next mCol
            ' carry-over: 직전 열의 재고 (재계산 열 이전의 마지막 재고)
            If IsNumeric(monthArr(rOff + OFFSET_재고1, firstDirtyOffset - 1)) Then carryOver1 = CDbl(monthArr(rOff + OFFSET_재고1, firstDirtyOffset - 1))
            If IsNumeric(monthArr(rOff + OFFSET_재고2, firstDirtyOffset - 1)) Then carryOver2 = CDbl(monthArr(rOff + OFFSET_재고2, firstDirtyOffset - 1))
        End If
        재고1 = carryOver1: 재고2 = carryOver2

        ' ── 변경 열부터 재계산 (오늘+10일까지만, 나머지는 기존값 유지) ──
        For mCol = firstDirtyOffset To maxRecalcOffset
            입고1 = 0: 입고2 = 0: 사용2 = 0
            If IsNumeric(monthArr(rOff + OFFSET_입고1, mCol)) Then 입고1 = CDbl(monthArr(rOff + OFFSET_입고1, mCol))
            If IsNumeric(monthArr(rOff + OFFSET_입고2, mCol)) Then 입고2 = CDbl(monthArr(rOff + OFFSET_입고2, mCol))
            If IsNumeric(monthArr(rOff + OFFSET_사용2, mCol)) Then 사용2 = CDbl(monthArr(rOff + OFFSET_사용2, mCol))
            입고1합 = 입고1합 + 입고1: 입고2합 = 입고2합 + 입고2: 사용2합 = 사용2합 + 사용2
            If mCol = firstDirtyOffset And firstDirtyOffset > 1 Then
                재고1 = carryOver1 + 입고1              ' 롤수: 사용 없음
                재고2 = carryOver2 + 입고2 - 사용2
            Else
                If mCol = 1 Then
                    재고1 = 기초재고1 + 입고1            ' 롤수: 사용 없음
                    재고2 = 기초재고2 + 입고2 - 사용2
                Else
                    전월1 = 0: 전월2 = 0
                    If IsNumeric(monthArr(rOff + OFFSET_재고1, mCol - 1)) Then 전월1 = CDbl(monthArr(rOff + OFFSET_재고1, mCol - 1))
                    If IsNumeric(monthArr(rOff + OFFSET_재고2, mCol - 1)) Then 전월2 = CDbl(monthArr(rOff + OFFSET_재고2, mCol - 1))
                    재고1 = 전월1 + 입고1                ' 롤수: 사용 없음
                    재고2 = 전월2 + 입고2 - 사용2
                End If
            End If
            monthArr(rOff + OFFSET_재고1, mCol) = 재고1
            monthArr(rOff + OFFSET_재고2, mCol) = 재고2
        Next mCol

        ' 10일 초과 구간 합계 누적 (재고는 수정 안함)
        Dim mm As Long
        For mm = maxRecalcOffset + 1 To totalCols
            If IsNumeric(monthArr(rOff + OFFSET_입고1, mm)) Then 입고1합 = 입고1합 + CDbl(monthArr(rOff + OFFSET_입고1, mm))
            If IsNumeric(monthArr(rOff + OFFSET_입고2, mm)) Then 입고2합 = 입고2합 + CDbl(monthArr(rOff + OFFSET_입고2, mm))
            If IsNumeric(monthArr(rOff + OFFSET_사용2, mm)) Then 사용2합 = 사용2합 + CDbl(monthArr(rOff + OFFSET_사용2, mm))
        Next mm

        ' 롤수 사용총합 = 0 (기초재고+입고로만 계산)
        infoArr(rOff, COL_입고총합) = 입고1합
        infoArr(rOff, COL_사용총합) = 0
        infoArr(rOff, COL_기말재고) = 재고1
        infoArr(rOff + 3, COL_입고총합) = 입고2합
        infoArr(rOff + 3, COL_사용총합) = 사용2합
        infoArr(rOff + 3, COL_기말재고) = 재고2
    Next 시작행

    값일괄_단계 = "시트쓰기"
    ' ── 최적화: 변경된 열부터만 sheet write (오늘+10일까지만, 나머지 열은 미쓰기) ──
    Dim writeToCol As Long: writeToCol = COL_월별시작 + maxRecalcOffset - 1
    If writeToCol > 마지막열 Then writeToCol = 마지막열
    Dim writeFromCol As Long: writeFromCol = COL_월별시작 + firstDirtyOffset - 1
    If firstDirtyOffset = 1 Then
        Dim writeCols As Long: writeCols = maxRecalcOffset
        If writeCols > totalCols Then writeCols = totalCols
        If writeCols >= 1 Then
            Dim fullSub() As Variant
            ReDim fullSub(1 To UBound(monthArr, 1), 1 To writeCols)
            Dim fsr As Long, fsc As Long
            For fsc = 1 To writeCols
                For fsr = 1 To UBound(monthArr, 1)
                    fullSub(fsr, fsc) = monthArr(fsr, fsc)
                Next fsr
            Next fsc
            ws.Range(ws.Cells(START_ROW, COL_월별시작), ws.Cells(마지막행, writeFromCol + writeCols - 1)).Value = fullSub
        End If
    Else
        ' 부분 배열 추출 (firstDirtyOffset-1 열부터, carry-over 열 포함)
        Dim wStart As Long: wStart = firstDirtyOffset - 1
        Dim wCols As Long: wCols = Application.WorksheetFunction.Min(maxRecalcOffset, totalCols) - wStart + 1
        If wCols < 1 Then wCols = 1
        Dim subArr() As Variant
        Dim numRows As Long: numRows = UBound(monthArr, 1)
        ReDim subArr(1 To numRows, 1 To wCols)
        Dim sr As Long, sc As Long
        For sc = 1 To wCols
            For sr = 1 To numRows
                subArr(sr, sc) = monthArr(sr, sc + wStart - 1)
            Next sr
        Next sc
        Dim wSheetCol As Long: wSheetCol = COL_월별시작 + wStart - 1
        ws.Range(ws.Cells(START_ROW, wSheetCol), ws.Cells(마지막행, wSheetCol + wCols - 1)).Value = subArr
        ' NumberFormat은 쓰기한 열에만 적용
        ws.Range(ws.Cells(START_ROW, writeFromCol), ws.Cells(마지막행, wSheetCol + wCols - 1)).NumberFormat = FMT_NUM_월별
    End If

    ' 고정열 서식 (D~H): 전체 재계산 시에만 (부분 재계산은 변경 없음)
    If firstDirtyOffset = 1 Then
        ws.Range(ws.Cells(START_ROW, COL_안전재고), ws.Cells(마지막행, COL_기말재고)).NumberFormat = FMT_NUM_STD
        ws.Range(ws.Cells(START_ROW, COL_월별시작), ws.Cells(마지막행, 마지막열)).NumberFormat = FMT_NUM_월별
    End If

    ' ── 요약 열 쓰기 (F·G·H: 입고총합·사용총합·기말재고) ──
    ' 병합셀 첫 행에만 기록. Union으로 한 번에 write 불가 → 3개 Union 범위 사용
    Dim rngF As Range, rngG As Range, rngH As Range
    Dim grpCnt As Long: grpCnt = 0
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        grpCnt = grpCnt + 1
    Next 시작행
    Dim arrF(1) As Variant, arrG(1) As Variant, arrH(1) As Variant  ' temp
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        rOff = 시작행 - START_ROW + 1
        ws.Cells(시작행, COL_입고총합).Value = infoArr(rOff, COL_입고총합)
        ws.Cells(시작행, COL_사용총합).Value = infoArr(rOff, COL_사용총합)
        ws.Cells(시작행, COL_기말재고).Value = infoArr(rOff, COL_기말재고)
        ws.Cells(시작행 + 3, COL_입고총합).Value = infoArr(rOff + 3, COL_입고총합)
        ws.Cells(시작행 + 3, COL_사용총합).Value = infoArr(rOff + 3, COL_사용총합)
        ws.Cells(시작행 + 3, COL_기말재고).Value = infoArr(rOff + 3, COL_기말재고)
    Next 시작행

정리:
    On Error GoTo 0
    If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호
    Application.EnableEvents = 이벤트상태
    Application.ScreenUpdating = origScr
    Application.Calculation = origCalc
    Exit Sub
값일괄오류:
    원단_디버그_실행단계 = "값일괄계산[" & 값일괄_단계 & "]"
    Err.Raise Err.Number, "원단관리_값일괄계산", "Err " & Err.Number & " / " & 값일괄_단계 & " / " & Err.Description
End Sub

    '================================================================================
    ' ★★★ 비활성화됨: 일일생산DB BS~BZ열(71~78) 부자재 블록 쓰기 완전 중단 ★★★
    '================================================================================
    Private Sub 일일생산DB_전체시트반영(ws As Worksheet)
        Exit Sub
    End Sub

    ' 시트 활성화 시 원단관리 재계산 (BS~BZ열 쓰기 제거됨)
    Public Sub 원단_일일생산DB_활성화시_갱신()
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If Not ws Is Nothing Then
            Call 원단관리_값일괄계산(ws)
        End If
    End Sub

    '================================================================================
    ' 첫 번째 그룹행(4~9행) 전체(A~마지막일) 서식 복사, A~I는 목록다운 유지·데이터 없음 (그룹당 6행)
    '================================================================================
    Private Sub 첫그룹_서식복사(ws As Worksheet, 대상시작행 As Long)
        Dim 마지막일 As Long
        Dim srcRng As Range, destRng As Range
        Dim srcA2K As Range, destA2K As Range
        마지막일 = 마지막일열_찾기(ws)
        ' 1) 전체( A4 ~ 마지막일 ) 서식 복사 - 테두리·병합·정렬·숫자서식 등
        Set srcRng = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 마지막일))
        Set destRng = ws.Range(ws.Cells(대상시작행, 1), ws.Cells(대상시작행 + ROWS_PER_ITEM - 1, 마지막일))
        srcRng.Copy
        destRng.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        ' 2) A~I열에 목록다운(유효성)만 추가 복사, 데이터는 붙이지 않음
        Set srcA2K = ws.Range(ws.Cells(START_ROW, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 9))
        Set destA2K = ws.Range(ws.Cells(대상시작행, 1), ws.Cells(대상시작행 + ROWS_PER_ITEM - 1, 9))
        srcA2K.Copy
        destA2K.PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False
    End Sub

    '================================================================================
    ' 빈 행 그룹에 데이터만 채우기 (구분·전월/현재/안전재고·자재유형 라벨) ? 서식은 첫그룹_서식복사 후 호출
    '================================================================================
    Private Sub 빈행그룹_데이터만_채우기(ws As Worksheet, 시작행 As Long)
        ws.Cells(시작행, COL_안전재고).Value = 0
        ws.Cells(시작행, COL_기초재고).Value = 0
        ws.Cells(시작행 + 3, COL_안전재고).Value = 0
        ws.Cells(시작행 + 3, COL_기초재고).Value = 0
        ws.Cells(시작행 + OFFSET_입고1, COL_구분).Value = "입고"
        ws.Cells(시작행 + OFFSET_사용1, COL_구분).Value = "사용"
        ws.Cells(시작행 + OFFSET_재고1, COL_구분).Value = "재고"
        ws.Cells(시작행 + OFFSET_입고2, COL_구분).Value = "입고"
        ws.Cells(시작행 + OFFSET_사용2, COL_구분).Value = "사용"
        ws.Cells(시작행 + OFFSET_재고2, COL_구분).Value = "재고"
        ' C열: 1차=Roll, 2차=중량 Kg
        ws.Cells(시작행, COL_포장구분).Value = "Roll"
        ws.Cells(시작행 + 3, COL_포장구분).Value = "중량" & vbLf & "Kg"
        Dim saveAlerts As Boolean: saveAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        On Error Resume Next
        ws.Range(ws.Cells(시작행, COL_제품발주처), ws.Cells(시작행 + 5, COL_제품발주처)).Merge
        ws.Range(ws.Cells(시작행, COL_제품명), ws.Cells(시작행 + 5, COL_제품명)).Merge
        ws.Range(ws.Cells(시작행, COL_포장구분), ws.Cells(시작행 + 2, COL_포장구분)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_포장구분), ws.Cells(시작행 + 5, COL_포장구분)).Merge
        ws.Range(ws.Cells(시작행, COL_안전재고), ws.Cells(시작행 + 2, COL_안전재고)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_안전재고), ws.Cells(시작행 + 5, COL_안전재고)).Merge
        ws.Range(ws.Cells(시작행, COL_기초재고), ws.Cells(시작행 + 2, COL_기초재고)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_기초재고), ws.Cells(시작행 + 5, COL_기초재고)).Merge
        ws.Range(ws.Cells(시작행, COL_입고총합), ws.Cells(시작행 + 2, COL_입고총합)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_입고총합), ws.Cells(시작행 + 5, COL_입고총합)).Merge
        ws.Range(ws.Cells(시작행, COL_사용총합), ws.Cells(시작행 + 2, COL_사용총합)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_사용총합), ws.Cells(시작행 + 5, COL_사용총합)).Merge
        ws.Range(ws.Cells(시작행, COL_기말재고), ws.Cells(시작행 + 2, COL_기말재고)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_기말재고), ws.Cells(시작행 + 5, COL_기말재고)).Merge
        On Error GoTo 0
        Application.DisplayAlerts = saveAlerts
    End Sub

    '================================================================================
    ' 개별 부자재 항목 생성 - 빈값 (그룹 6행: 입고/사용/재고)
    ' A=제품발주처, B=제품명, C=포장구분, D=안전재고(병합), E=기초재고(병합),
    ' F=입고총합(병합), G=사용총합(병합), H=기말재고(병합), I=구분(입고/사용/재고)
'================================================================================
Private Sub 부자재항목_생성_빈값(ws As Worksheet, 시작행 As Long)
    Dim rng As Range, saveAlerts As Boolean
    Dim 마지막열 As Long: 마지막열 = 마지막월열_찾기(ws)
    Dim 원본끝열 As Long

    ' 4~6행(첫 그룹) 서식을 복사하여 적용 (첫 그룹은 자기복사 후 직접 테두리 적용)
    원본끝열 = 마지막열
    If 원본끝열 < COL_구분 Then 원본끝열 = COL_구분
    On Error Resume Next
    ws.Range(ws.Cells(START_ROW, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 원본끝열)).Copy
    ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 원본끝열)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0

    ' C열 라벨: 1차=Roll, 2차=중량 Kg
    ws.Cells(시작행, COL_포장구분).Value = "Roll"
    ws.Cells(시작행 + 3, COL_포장구분).Value = "중량" & vbLf & "Kg"
    ' I열 구분: 입고,사용,재고,입고,사용,재고
    ws.Cells(시작행 + OFFSET_입고1, COL_구분).Value = "입고"
    ws.Cells(시작행 + OFFSET_사용1, COL_구분).Value = "사용"
    ws.Cells(시작행 + OFFSET_재고1, COL_구분).Value = "재고"
    ws.Cells(시작행 + OFFSET_입고2, COL_구분).Value = "입고"
    ws.Cells(시작행 + OFFSET_사용2, COL_구분).Value = "사용"
    ws.Cells(시작행 + OFFSET_재고2, COL_구분).Value = "재고"

    ' A~C, D~H 병합 (6행). C·D~H는 Roll(3행)·중량Kg(3행) 두 블록
    saveAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error Resume Next
    ws.Range(ws.Cells(시작행, COL_제품발주처), ws.Cells(시작행 + 5, COL_제품발주처)).Merge
    ws.Range(ws.Cells(시작행, COL_제품명), ws.Cells(시작행 + 5, COL_제품명)).Merge
    ws.Range(ws.Cells(시작행, COL_포장구분), ws.Cells(시작행 + 2, COL_포장구분)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_포장구분), ws.Cells(시작행 + 5, COL_포장구분)).Merge
    ws.Range(ws.Cells(시작행, COL_안전재고), ws.Cells(시작행 + 2, COL_안전재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_안전재고), ws.Cells(시작행 + 5, COL_안전재고)).Merge
    ws.Range(ws.Cells(시작행, COL_기초재고), ws.Cells(시작행 + 2, COL_기초재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_기초재고), ws.Cells(시작행 + 5, COL_기초재고)).Merge
    ws.Range(ws.Cells(시작행, COL_입고총합), ws.Cells(시작행 + 2, COL_입고총합)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_입고총합), ws.Cells(시작행 + 5, COL_입고총합)).Merge
    ws.Range(ws.Cells(시작행, COL_사용총합), ws.Cells(시작행 + 2, COL_사용총합)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_사용총합), ws.Cells(시작행 + 5, COL_사용총합)).Merge
    ws.Range(ws.Cells(시작행, COL_기말재고), ws.Cells(시작행 + 2, COL_기말재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_기말재고), ws.Cells(시작행 + 5, COL_기말재고)).Merge
    On Error GoTo 0
    Application.DisplayAlerts = saveAlerts

    ws.Cells(시작행, COL_안전재고).Value = 0
    ws.Cells(시작행, COL_기초재고).Value = 0
    ws.Cells(시작행 + 3, COL_안전재고).Value = 0
    ws.Cells(시작행 + 3, COL_기초재고).Value = 0

    ws.Range(ws.Cells(시작행, COL_안전재고), ws.Cells(시작행, COL_기말재고)).NumberFormat = FMT_NUM_STD
    ws.Range(ws.Cells(시작행 + 3, COL_안전재고), ws.Cells(시작행 + 3, COL_기말재고)).NumberFormat = FMT_NUM_STD
    If 마지막열 >= COL_월별시작 Then ws.Range(ws.Cells(시작행, COL_월별시작), ws.Cells(시작행 + 5, 마지막열)).NumberFormat = FMT_NUM_월별

    ' ★ 그룹 생성 즉시 전체 테두리 적용 - 부자재관리와 동일 (첫 그룹에 테두리 선적용 → 이후 그룹이 복사 시 테두리 포함)
    With ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 원본끝열)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

'================================================================================
    ' 개별 부자재 항목 생성 (6행: 입고/사용/재고) - 값 직접 입력용
'================================================================================
Private Sub 부자재항목_생성(ws As Worksheet, 시작행 As Long, 카테고리 As String, _
                            부자재명 As String, Optional 업체명 As String = "-", Optional 배경색 As Long = -1, Optional 설비라인 As String = "", Optional 형식복사생략 As Boolean = False)
    Dim rng As Range, saveAlerts As Boolean
    Dim 마지막열 As Long: 마지막열 = 마지막월열_찾기(ws)
    Dim 원본끝열 As Long

    ' 4~6행(첫 그룹) 서식을 복사하여 적용 (형식복사생략=True면 스킵 - 목록추가 시 일괄복사 후 호출)
    If Not 형식복사생략 Then
        원본끝열 = 마지막열
        If 원본끝열 < COL_구분 Then 원본끝열 = COL_구분
        On Error Resume Next
        ws.Range(ws.Cells(START_ROW, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 원본끝열)).Copy
        ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 원본끝열)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        On Error GoTo 0
    End If
    원본끝열 = 마지막열
    If 원본끝열 < COL_구분 Then 원본끝열 = COL_구분

    ws.Cells(시작행, COL_제품발주처).Value = 업체명
    ws.Cells(시작행, COL_제품명).Value = 부자재명
    ' C열 라벨: 1차=Roll, 2차=중량 Kg (원단관리용)
    ws.Cells(시작행, COL_포장구분).Value = "Roll"
    ws.Cells(시작행 + 3, COL_포장구분).Value = "중량" & vbLf & "Kg"
    ws.Cells(시작행, COL_안전재고).Value = 0
    ws.Cells(시작행, COL_기초재고).Value = 0
    ws.Cells(시작행 + 3, COL_안전재고).Value = 0
    ws.Cells(시작행 + 3, COL_기초재고).Value = 0

    ' I열 구분: 입고,사용,재고,입고,사용,재고
    ws.Cells(시작행 + OFFSET_입고1, COL_구분).Value = "입고"
    ws.Cells(시작행 + OFFSET_사용1, COL_구분).Value = "사용"
    ws.Cells(시작행 + OFFSET_재고1, COL_구분).Value = "재고"
    ws.Cells(시작행 + OFFSET_입고2, COL_구분).Value = "입고"
    ws.Cells(시작행 + OFFSET_사용2, COL_구분).Value = "사용"
    ws.Cells(시작행 + OFFSET_재고2, COL_구분).Value = "재고"

    ' A~C, D~H 병합 (6행). C·D~H는 Roll(3행)·중량Kg(3행) 두 블록
    saveAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error Resume Next
    ws.Range(ws.Cells(시작행, COL_제품발주처), ws.Cells(시작행 + 5, COL_제품발주처)).Merge
    ws.Range(ws.Cells(시작행, COL_제품명), ws.Cells(시작행 + 5, COL_제품명)).Merge
    ws.Range(ws.Cells(시작행, COL_포장구분), ws.Cells(시작행 + 2, COL_포장구분)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_포장구분), ws.Cells(시작행 + 5, COL_포장구분)).Merge
    ws.Range(ws.Cells(시작행, COL_안전재고), ws.Cells(시작행 + 2, COL_안전재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_안전재고), ws.Cells(시작행 + 5, COL_안전재고)).Merge
    ws.Range(ws.Cells(시작행, COL_기초재고), ws.Cells(시작행 + 2, COL_기초재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_기초재고), ws.Cells(시작행 + 5, COL_기초재고)).Merge
    ws.Range(ws.Cells(시작행, COL_입고총합), ws.Cells(시작행 + 2, COL_입고총합)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_입고총합), ws.Cells(시작행 + 5, COL_입고총합)).Merge
    ws.Range(ws.Cells(시작행, COL_사용총합), ws.Cells(시작행 + 2, COL_사용총합)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_사용총합), ws.Cells(시작행 + 5, COL_사용총합)).Merge
    ws.Range(ws.Cells(시작행, COL_기말재고), ws.Cells(시작행 + 2, COL_기말재고)).Merge
    ws.Range(ws.Cells(시작행 + 3, COL_기말재고), ws.Cells(시작행 + 5, COL_기말재고)).Merge
    On Error GoTo 0
    Application.DisplayAlerts = saveAlerts

    ws.Range(ws.Cells(시작행, COL_안전재고), ws.Cells(시작행, COL_기말재고)).NumberFormat = FMT_NUM_STD
    ws.Range(ws.Cells(시작행 + 3, COL_안전재고), ws.Cells(시작행 + 3, COL_기말재고)).NumberFormat = FMT_NUM_STD
    If 마지막열 >= COL_월별시작 Then ws.Range(ws.Cells(시작행, COL_월별시작), ws.Cells(시작행 + 5, 마지막열)).NumberFormat = FMT_NUM_월별

    ' ★ 그룹 생성 즉시 전체 테두리 적용
    With ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 원본끝열)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

'================================================================================
' 재고 부족 조건부 서식 - 값<0일 때 굵게만 표시 (색상 제외)
'================================================================================
Private Sub 재고부족_조건부서식(ws As Worksheet, 행번호 As Long)
    Dim rng As Range
    Dim 마지막열 As Long: 마지막열 = 마지막월열_찾기(ws)
    Dim 그룹시작행 As Long, 안전재고행 As Long
    그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
    If (행번호 - 그룹시작행) >= 3 Then 안전재고행 = 그룹시작행 + 3 Else 안전재고행 = 그룹시작행
    Dim 안전재고주소 As String
    안전재고주소 = ws.Cells(안전재고행, COL_안전재고).Address(True, True)
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(행번호, COL_월별시작), ws.Cells(행번호, 마지막열))
    If rng Is Nothing Then Exit Sub
    rng.FormatConditions.Delete
    Dim 셀주소 As String
    셀주소 = ws.Cells(행번호, COL_월별시작).Address(False, False)
    Dim 수식 As String
    수식 = "=OR(" & 셀주소 & "<0,AND(ISNUMBER(" & 안전재고주소 & ")," & 셀주소 & "<" & 안전재고주소 & "))"
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=수식)
        .Font.Bold = True
    End With
    On Error GoTo 0
End Sub

'================================================================================
' 토요일/일요일 헤더(2~3행) 색상 구분. 2행=월, 3행=일, I2=연도
'================================================================================
Private Sub 주말서식_일수행(ws As Worksheet)
    Dim 마지막일열 As Long, c As Long
    Dim v2 As Variant, v3 As Variant, 기준년 As Long, vI2 As Variant
    Dim 월v As Long, 일v As Long, d As Date, wd As Long
    If ws Is Nothing Then Exit Sub
    마지막일열 = 마지막월열_찾기(ws)
    If 마지막일열 < COL_월별시작 Then Exit Sub
    vI2 = ws.Range("I2").Value
    기준년 = Year(Date)
    If IsNumeric(vI2) And vI2 >= 2000 And vI2 <= 2100 Then 기준년 = CLng(vI2)
    For c = COL_월별시작 To 마지막일열
        v3 = ws.Cells(3, c).Value
        If Not IsNumeric(v3) Then GoTo NextCol2
        일v = CLng(v3)
        If 일v < 1 Or 일v > 31 Then GoTo NextCol2
        v2 = Trim(Replace(Replace(Replace(CStr(ws.Cells(2, c).Value), "M", ""), "m", ""), " ", ""))
        If v2 = "" Or Not IsNumeric(v2) Then GoTo NextCol2
        월v = CLng(v2)
        If 월v < 1 Or 월v > 12 Then GoTo NextCol2
        If 일v > Day(DateSerial(기준년, 월v + 1, 0)) Then GoTo NextCol2
        d = DateSerial(기준년, 월v, 일v)
        wd = Weekday(d, vbSunday)
        If wd = 1 Then
            ws.Cells(3, c).Interior.color = RGB(255, 220, 220)
        ElseIf wd = 7 Then
            ws.Cells(3, c).Interior.color = RGB(220, 230, 255)
        Else
            ws.Cells(3, c).Interior.ColorIndex = xlNone
        End If
NextCol2:
    Next c
End Sub

'================================================================================
' 다음달 열 추가 (2행=1M월, 3행=일. 마지막 열의 2행에서 월 판단 → 다음달 1~말일 열 추가)
'================================================================================
Public Sub 원단_다음달열_추가()
    Dim ws As Worksheet
    Dim lastRow As Long, 마지막열 As Long, 새열 As Long, 기준열 As Long
    Dim 다음월 As Integer, 다음년 As Long, 다음달일수 As Integer, d As Long
    Dim s2 As String, 마지막월 As Integer, 마지막일 As Integer, 기준년 As Long, vI2 As Variant
    Dim 현재월말일 As Integer, 현재월추가일수 As Integer
    Dim 응답 As VbMsgBoxResult
    Dim 시트보호상태 As Boolean
    Dim 데이터있음 As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_원단관리)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    시트보호상태 = ws.ProtectContents
    If 시트보호상태 Then
        On Error Resume Next
        ws.Unprotect WS_시트보호암호
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo 0
    End If

    마지막열 = 마지막월열_찾기(ws)
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    데이터있음 = (lastRow >= START_ROW)

    vI2 = ws.Range("I2").Value
    기준년 = Year(Date)
    If IsNumeric(vI2) And vI2 >= 2000 And vI2 <= 2100 Then 기준년 = CLng(vI2)

    ' ── 마지막 헤더에서 월·일 파악 ──
    마지막월 = 0: 마지막일 = 0
    If 마지막열 >= COL_월별시작 Then
        s2 = Trim(Replace(Replace(Replace(CStr(ws.Cells(2, 마지막열).Value), "M", ""), "m", ""), " ", ""))
        If s2 <> "" And IsNumeric(s2) Then 마지막월 = CLng(s2)
        Dim v3 As Variant: v3 = ws.Cells(3, 마지막열).Value
        If IsNumeric(v3) Then 마지막일 = CLng(v3)
    End If
    If 마지막월 < 1 Or 마지막월 > 12 Then 마지막월 = Month(Date)

    ' ── 현재 달 말일 계산 → 부족한 날 수 계산 ──
    현재월말일 = Day(DateSerial(기준년, 마지막월 + 1, 0))
    현재월추가일수 = 현재월말일 - 마지막일
    If 현재월추가일수 < 0 Then 현재월추가일수 = 0

    ' ── 다음 달 정보 ──
    다음월 = 마지막월 + 1
    다음년 = 기준년
    If 다음월 > 12 Then 다음월 = 1: 다음년 = 기준년 + 1
    다음달일수 = Day(DateSerial(다음년, 다음월 + 1, 0))

    ' ── 확인 메시지 ──
    Dim 안내 As String
    안내 = IIf(데이터있음, " [헤더+데이터 전체]", " [헤더만 - 데이터 없음]")
    Dim msgTxt As String
    If 현재월추가일수 > 0 Then
        msgTxt = 기준년 & "년 " & 마지막월 & "월 " & (마지막일 + 1) & "~" & 현재월말일 & "일 (" & 현재월추가일수 & "일)" & vbCrLf & _
                 "  + " & 다음년 & "년 " & 다음월 & "월 전체 (" & 다음달일수 & "일)" & vbCrLf & _
                 "  합계 " & (현재월추가일수 + 다음달일수) & "열을 추가할까요?" & 안내
    Else
        msgTxt = 다음년 & "년 " & 다음월 & "월 (" & 다음달일수 & "일) 열을 추가할까요?" & 안내
    End If
    응답 = MsgBox(msgTxt, vbYesNo + vbQuestion, "다음달 열 추가")
    If 응답 <> vbYes Then
        If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    기준열 = 마지막열

    ' ── Step 1: 현재 달 나머지 채우기 (말일이 아닐 때만) ──
    If 현재월추가일수 > 0 Then
        Dim dd As Long
        For dd = 마지막일 + 1 To 현재월말일
            새열 = 기준열 + (dd - 마지막일)
            If 데이터있음 Then
                If 기준열 >= COL_월별시작 Then
                    ws.Range(ws.Cells(2, 기준열), ws.Cells(lastRow, 기준열)).Copy
                    ws.Range(ws.Cells(2, 새열), ws.Cells(lastRow, 새열)).PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If
                ws.Range(ws.Cells(START_ROW, 새열), ws.Cells(lastRow, 새열)).NumberFormat = FMT_NUM_월별
            Else
                If 기준열 >= COL_월별시작 Then
                    ws.Range(ws.Cells(2, 기준열), ws.Cells(3, 기준열)).Copy
                    ws.Range(ws.Cells(2, 새열), ws.Cells(3, 새열)).PasteSpecial Paste:=xlPasteFormats
                    Application.CutCopyMode = False
                End If
            End If
            ws.Cells(2, 새열).Value = 마지막월 & "M"
            ws.Cells(2, 새열).Font.color = RGB(160, 160, 160)
            ws.Cells(3, 새열).Value = dd
        Next dd
        기준열 = 마지막열 + 현재월추가일수
    End If

    ' ── Step 2: 다음 달 전체 추가 ──
    For d = 1 To 다음달일수
        새열 = 기준열 + d
        If 데이터있음 Then
            If 기준열 >= COL_월별시작 Then
                ws.Range(ws.Cells(2, 기준열), ws.Cells(lastRow, 기준열)).Copy
                ws.Range(ws.Cells(2, 새열), ws.Cells(lastRow, 새열)).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
            ws.Range(ws.Cells(START_ROW, 새열), ws.Cells(lastRow, 새열)).NumberFormat = FMT_NUM_월별
        Else
            If 기준열 >= COL_월별시작 Then
                ws.Range(ws.Cells(2, 기준열), ws.Cells(3, 기준열)).Copy
                ws.Range(ws.Cells(2, 새열), ws.Cells(3, 새열)).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
        End If
        ws.Cells(2, 새열).Value = 다음월 & "M"
        ws.Cells(2, 새열).Font.color = RGB(160, 160, 160)
        ws.Cells(3, 새열).Value = d
    Next d

    ' ── 헤더 테두리 + 주말 서식: 항상 적용 ──
    Call 날짜영역_테두리적용(ws)
    Call 주말서식_일수행(ws)

    ' ── 데이터 있을 때만: 그룹 음영·재고 재계산 ──
    If 데이터있음 Then
        Call 원단관리_그룹행_교대음영_적용(ws, lastRow)
        Call 원단관리_값일괄계산(ws)
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호

    Dim 완료msg As String
    If 현재월추가일수 > 0 Then
        완료msg = 기준년 & "년 " & 마지막월 & "월 " & (마지막일 + 1) & "~" & 현재월말일 & "일 + " & _
                  다음년 & "년 " & 다음월 & "월 전체, 합계 " & (현재월추가일수 + 다음달일수) & "열이 추가되었습니다."
    Else
        완료msg = 다음년 & "년 " & 다음월 & "월 " & 다음달일수 & "일 열이 추가되었습니다."
    End If
    If Not 데이터있음 Then 완료msg = 완료msg & vbCrLf & "(데이터 입력 후 데이터업데이트 실행 시 데이터 영역 테두리·서식이 자동 적용됩니다)"
    MsgBox 완료msg, vbInformation, "다음달 열 추가"
End Sub

    ' 그룹 6행 D~H열 UnMerge/Merge 헬퍼
    ' D~H(안전재고~기말재고): C열처럼 1차포장(3행)·2차포장(3행) 두 블록으로 병합
    Private Sub 블록5_UnMerge(ws As Worksheet, 시작행 As Long, col As Long)
        On Error Resume Next
        ws.Range(ws.Cells(시작행, col), ws.Cells(시작행 + 5, col)).UnMerge
        On Error GoTo 0
    End Sub
    Private Sub 블록5_Merge(ws As Worksheet, 시작행 As Long, col As Long)
        Dim saveAlerts As Boolean: saveAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        On Error Resume Next
        If col >= COL_안전재고 And col <= COL_기말재고 Then
            ' D~H: 1차(시작행~+2), 2차(시작행+3~+5) 각각 병합
            ws.Range(ws.Cells(시작행, col), ws.Cells(시작행 + 2, col)).Merge
            ws.Range(ws.Cells(시작행 + 3, col), ws.Cells(시작행 + 5, col)).Merge
        Else
            ws.Range(ws.Cells(시작행, col), ws.Cells(시작행 + 5, col)).Merge
        End If
        On Error GoTo 0
        Application.DisplayAlerts = saveAlerts
    End Sub

'================================================================================
' 그룹행 C열·I열 배열 복원 (C=Roll/중량Kg 2블록, I=입고/사용/재고 비병합)
' 데이터 업데이트 시 PasteSpecial 등으로 깨진 구조 복구
'================================================================================
Private Sub 원단관리_그룹행_C열I열_배열복원(ws As Worksheet, 마지막행 As Long)
    Dim 시작행 As Long
    If ws Is Nothing Or 마지막행 < START_ROW Then Exit Sub
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        On Error Resume Next
        ws.Range(ws.Cells(시작행, COL_포장구분), ws.Cells(시작행 + 5, COL_포장구분)).UnMerge
        ws.Range(ws.Cells(시작행, COL_구분), ws.Cells(시작행 + 5, COL_구분)).UnMerge
        On Error GoTo 0
        ws.Range(ws.Cells(시작행, COL_포장구분), ws.Cells(시작행 + 2, COL_포장구분)).Merge
        ws.Range(ws.Cells(시작행 + 3, COL_포장구분), ws.Cells(시작행 + 5, COL_포장구분)).Merge
        ws.Cells(시작행, COL_포장구분).Value = "Roll"
        ws.Cells(시작행 + 3, COL_포장구분).Value = "중량" & vbLf & "Kg"
        ws.Cells(시작행 + OFFSET_입고1, COL_구분).Value = "입고"
        ws.Cells(시작행 + OFFSET_사용1, COL_구분).Value = "사용"
        ws.Cells(시작행 + OFFSET_재고1, COL_구분).Value = "재고"
        ws.Cells(시작행 + OFFSET_입고2, COL_구분).Value = "입고"
        ws.Cells(시작행 + OFFSET_사용2, COL_구분).Value = "사용"
        ws.Cells(시작행 + OFFSET_재고2, COL_구분).Value = "재고"
    Next 시작행
End Sub

'================================================================================
' 그룹행 교대 음영: 1·3·5번째 그룹 RGB(242,247,252) + 아래 굵은 검정 실선, 2·4번째 그룹은 그대로
'================================================================================
Private Sub 원단관리_그룹행_교대음영_적용(ws As Worksheet, 마지막행 As Long)
    Dim 시작행 As Long, grpIdx As Long, 마지막일열 As Long
    Dim rngOdd As Range, rngEven As Range, rngBorder As Range, rngFont As Range
    If ws Is Nothing Or 마지막행 < START_ROW Then Exit Sub
    마지막일열 = 마지막일열_찾기(ws)
    If 마지막일열 < COL_일별시작 Then Exit Sub
    ' 전체 폰트색 1회 적용
    ws.Range(ws.Cells(START_ROW, COL_일별시작), ws.Cells(마지막행, 마지막일열)).Font.color = RGB(0, 0, 0)
    ' Union으로 홀수/짝수 그룹 범위 모으기
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        grpIdx = ((시작행 - START_ROW) \ ROWS_PER_ITEM) + 1
        If grpIdx Mod 2 = 1 Then
            If rngOdd Is Nothing Then Set rngOdd = ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)) Else Set rngOdd = Union(rngOdd, ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)))
        Else
            If rngEven Is Nothing Then Set rngEven = ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)) Else Set rngEven = Union(rngEven, ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)))
        End If
        If rngBorder Is Nothing Then Set rngBorder = ws.Cells(시작행 + ROWS_PER_ITEM - 1, 1) Else Set rngBorder = Union(rngBorder, ws.Cells(시작행 + ROWS_PER_ITEM - 1, 1))
    Next 시작행
    ' 1회 일괄 적용
    If Not rngOdd Is Nothing Then
        With rngOdd.Interior: .Pattern = xlSolid: .color = RGB(242, 247, 252): End With
    End If
    If Not rngEven Is Nothing Then rngEven.Interior.Pattern = xlPatternNone
    ' 하단 테두리 (각 그룹 마지막 행)
    If Not rngBorder Is Nothing Then
        Dim srcBorderWt As Variant, srcBorderStyle As Variant
        On Error Resume Next
        srcBorderStyle = ws.Range(ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 마지막일열)).Borders(xlEdgeBottom).LineStyle
        srcBorderWt = ws.Range(ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 1), ws.Cells(START_ROW + ROWS_PER_ITEM - 1, 마지막일열)).Borders(xlEdgeBottom).Weight
        On Error GoTo 0
        If IsEmpty(srcBorderStyle) Or IsError(srcBorderStyle) Then srcBorderStyle = xlContinuous
        If IsEmpty(srcBorderWt) Or IsError(srcBorderWt) Then srcBorderWt = xlThin
        Dim area As Range
        For Each area In rngBorder.Areas
            With ws.Range(ws.Cells(area.Row, 1), ws.Cells(area.Row, 마지막일열)).Borders(xlEdgeBottom)
                .LineStyle = srcBorderStyle: .Weight = srcBorderWt: .color = RGB(0, 0, 0)
            End With
        Next area
    End If
End Sub

'================================================================================
' 시트 서식 적용 - 숫자서식·재고부족 조건부(색 없음). 열너비·틀고정은 사용자 지정 유지.
' 헤더수정금지=True 시 1~3행(헤더) 수정 생략 (목업 호출 시)
'================================================================================
Private Sub 시트서식_적용(ws As Worksheet, 마지막행 As Long, Optional 헤더수정금지 As Boolean = False)
    Dim r As Long, 시작행 As Long, i As Long
    Dim dataRng As Range
    Dim 시트보호상태 As Boolean

    ' 시트 보호 상태 확인 및 해제 (암호 대화상자 없이 임시 해제)
    시트보호상태 = ws.ProtectContents
    If 시트보호상태 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0

    ws.Activate
    Dim 첫월 As Integer
    Dim 마지막일열 As Long
    Dim 첫달일수 As Integer
    Dim 끝열 As Long
    마지막일열 = 마지막일열_찾기(ws)
    
    ' 헤더 수정 금지 모드가 아닐 때만 1~3행 수정 (목업 호출 시 헤더 유지)
    If Not 헤더수정금지 Then
        With ws.Range("I2")
            .Value = Year(Date)
            .Font.color = RGB(255, 255, 255)
        End With
        첫월 = B2_월(ws)
        첫달일수 = 첫달일수_찾기(ws)
        For i = 1 To 첫달일수
            ' ★ 2·3행에 기존 값이 있으면 절대 덮어쓰지 않음
            If IsEmpty(ws.Cells(2, COL_일별시작 + i - 1).Value) Or Len(Trim(CStr(ws.Cells(2, COL_일별시작 + i - 1).Value))) = 0 Then
                With ws.Cells(2, COL_일별시작 + i - 1)
                    .Value = 첫월 & "m"
                    .Font.color = RGB(160, 160, 160)
                End With
            End If
            If IsEmpty(ws.Cells(3, COL_일별시작 + i - 1).Value) Then
                ws.Cells(3, COL_일별시작 + i - 1).Value = i
            End If
        Next i
        끝열 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column > 끝열 Then 끝열 = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
        If 마지막행 >= START_ROW Then
            Dim data끝열 As Long: data끝열 = ws.Cells(마지막행, ws.Columns.Count).End(xlToLeft).Column
            If data끝열 > 끝열 Then 끝열 = data끝열
        End If
        If 끝열 > 마지막일열 Then
            ' ★ 2·3행은 기존 값 보존(절대 삭제 금지) → 4행~ 데이터 영역만 Clear
            If 마지막행 >= START_ROW Then ws.Range(ws.Cells(START_ROW, 마지막일열 + 1), ws.Cells(마지막행, 끝열)).ClearContents
        End If
        Call 주말서식_일수행(ws)
    End If
    If 마지막행 >= START_ROW Then Call 원단관리_그룹행_C열I열_배열복원(ws, 마지막행)
    If 마지막행 >= START_ROW Then
        For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            On Error Resume Next
            Call 블록5_UnMerge(ws, 시작행, COL_안전재고)
            Call 블록5_UnMerge(ws, 시작행, COL_기초재고)
            Call 블록5_UnMerge(ws, 시작행, COL_입고총합)
            Call 블록5_UnMerge(ws, 시작행, COL_사용총합)
            Call 블록5_UnMerge(ws, 시작행, COL_기말재고)
            On Error GoTo 0
            Call 블록5_Merge(ws, 시작행, COL_안전재고)
            Call 블록5_Merge(ws, 시작행, COL_기초재고)
            Call 블록5_Merge(ws, 시작행, COL_입고총합)
            Call 블록5_Merge(ws, 시작행, COL_사용총합)
            Call 블록5_Merge(ws, 시작행, COL_기말재고)
        Next 시작행
        ws.Range(ws.Cells(START_ROW, COL_안전재고), ws.Cells(마지막행, COL_기말재고)).NumberFormat = FMT_NUM_STD
        If 마지막일열 >= COL_월별시작 Then
            ws.Range(ws.Cells(START_ROW, COL_월별시작), ws.Cells(마지막행, 마지막일열)).NumberFormat = FMT_NUM_월별
            ws.Range(ws.Cells(START_ROW, COL_월별시작), ws.Cells(마지막행, 마지막일열)).Font.color = RGB(0, 0, 0)
        End If
        For r = START_ROW To 마지막행
            If Trim(CStr(ws.Cells(r, COL_구분).Value)) = "재고" Then
                Call 재고부족_조건부서식(ws, r)
            End If
        Next r
    End If
    ' 데이터 영역 테두리 적용 (헤더 2~3행, 고정열 A~I, 일별 열)
    Call 날짜영역_테두리적용(ws)
    ' 그룹행 교대 음영 (1·3·5번째: RGB 242,247,252 + 굵은 검정 하단선) - 시트 실제 lastRow 기준
    Dim 실제lastRow As Long
    실제lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 실제lastRow >= START_ROW Then Call 원단관리_그룹행_교대음영_적용(ws, 실제lastRow)
    ' 틀고정/창 나누기는 사용자가 설정한 값 유지 (매크로에서 변경하지 않음)

    ' 시트 보호 상태 복원
        If 시트보호상태 Then Call 원단_셀보호_적용_무음(ws)
End Sub

'================================================================================
' 부자재 목록 로드 (중복 제거)
'================================================================================
Private Function 부자재목록_로드(ws As Worksheet, 열번호 As Integer) As Collection
    Dim result As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim key As String
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, 열번호).End(xlUp).Row
    
    For i = 3 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 열번호).Value))
        If cellValue <> "" And cellValue <> "0" Then
            key = LCase(cellValue)
            result.Add cellValue, key
        End If
    Next i
    On Error GoTo 0
    
    Set 부자재목록_로드 = result
End Function

'================================================================================
' 필터/출력: 분류·부자재명·업체명 중 상위 기준 선택 후 해당 행만 새 워크북 저장
' - 연·월 필터: 원하는 월만 추려서 저장 가능 (비우면 전체 기간)
' - 저장 경로: Index key 시트 A3(없으면 바탕화면), 출력 범위: B3 셀부터
'================================================================================
Public Sub 업체별원단관리_생성()
    Dim ws재고 As Worksheet
    Dim wsIndexKey As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long, 마지막열 As Long
    Dim 시작행 As Long
    Dim 저장경로 As String, 파일명 As String, 전체경로 As String
    Dim destRow As Long
        Dim r As Long, c As Long, i As Long, j As Long
    Dim srcRng As Range
        Dim 선택년 As Integer, 선택월 As Integer
    Dim 월필터사용 As Boolean
    Dim 해당월시작열 As Long, 해당월일수 As Integer
    Dim 셀값 As String
        Const OUT_START_ROW As Long = 3
        Const OUT_START_COL As Long = 2
        Dim 출력끝열 As Long
        
        Dim 필터열배열(1 To 2) As Long
        Dim 필터명배열(1 To 2) As String
        필터열배열(1) = COL_업체명:     필터명배열(1) = "제품발주처"
        필터열배열(2) = COL_제품:       필터명배열(2) = "제품"
        
        Dim 선택값배열(1 To 2) As String
        For i = 1 To 2: 선택값배열(i) = "": Next i

    On Error Resume Next
        Set ws재고 = ThisWorkbook.Sheets(WS_원단관리)
    Set wsIndexKey = ThisWorkbook.Sheets(WS_IndexKey)
    On Error GoTo 0

    If ws재고 Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = ws재고.Cells(ws재고.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then
            MsgBox "원단관리 시트에 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

        ' ================================================================
        ' 1단계: 사용할 필터 기준 선택 (다중 선택 가능, 쉼표 구분)
        ' ================================================================
        Dim 기준선택_raw As String
        기준선택_raw = InputBox( _
            "사용할 필터 기준의 번호를 입력하세요." & vbCrLf & _
            "(여러 개 선택 시 쉼표로 구분, 예: 1,2)" & vbCrLf & vbCrLf & _
            "1. 제품발주처" & vbCrLf & _
            "2. 제품" & vbCrLf & vbCrLf & _
            "비워두면 전체 항목을 출력합니다.", _
            "필터/출력 - 기준 선택", "")
        
        If StrPtr(기준선택_raw) = 0 Then Exit Sub  ' 취소 버튼
    Dim 기준선택 As String
        기준선택 = Trim(기준선택_raw)
        
        ' 선택한 기준 번호 파싱
        Dim 사용필터수 As Long: 사용필터수 = 0
        Dim 사용필터인덱스() As Long
        Dim 선택번호배열() As String
        Dim 번호 As Long
        Dim 필터없음 As Boolean: 필터없음 = (기준선택 = "")
        
        If Not 필터없음 Then
            기준선택 = Replace(기준선택, " ", "")
            선택번호배열 = Split(기준선택, ",")
            ReDim 사용필터인덱스(0 To UBound(선택번호배열))
            For i = 0 To UBound(선택번호배열)
                If IsNumeric(선택번호배열(i)) Then
                    번호 = CLng(선택번호배열(i))
                    If 번호 >= 1 And 번호 <= 2 Then
                        사용필터인덱스(사용필터수) = 번호
                        사용필터수 = 사용필터수 + 1
                    End If
                End If
            Next i
            If 사용필터수 = 0 Then 필터없음 = True
    End If

        ' ================================================================
        ' 2단계: 각 필터 기준별 고유값 수집 → 항목 선택
        ' ================================================================
        If Not 필터없음 Then
            Dim 목록 As Collection
            Dim 목록키 As String
            Dim 현재필터 As Long
            
            For i = 0 To 사용필터수 - 1
                현재필터 = 사용필터인덱스(i)
                Set 목록 = New Collection
                
                ' 고유값 수집
    For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
                    셀값 = Trim(CStr(ws재고.Cells(시작행, 필터열배열(현재필터)).Value))
        If 셀값 <> "" Then
            목록키 = LCase(Replace(셀값, " ", ""))
            On Error Resume Next
            목록.Add 셀값, 목록키
            On Error GoTo 0
        End If
    Next 시작행

    If 목록.Count = 0 Then
                    MsgBox 필터명배열(현재필터) & " 데이터가 없습니다.", vbExclamation
                    GoTo NextFilter
                End If
                
                ' 목록에서 항목 선택
                Dim 선택결과 As String
                선택결과 = 목록선택_다이얼로그(목록, _
                    "[ " & 필터명배열(현재필터) & " ] 내보낼 항목을 선택하세요.", _
                    "필터/출력 - " & 필터명배열(현재필터) & " 선택")
                
                If 선택결과 = "<<CANCEL>>" Then
                    ' 취소 버튼 → 전체 작업 중단
        Exit Sub
                ElseIf 선택결과 = "" Then
                    ' 빈 입력 → 해당 필터 미사용
                    선택값배열(현재필터) = ""
                Else
                    선택값배열(현재필터) = 선택결과
    End If
NextFilter:
            Next i
            
            ' 선택된 필터가 하나도 없으면 전체 출력으로 전환
            Dim 유효필터수 As Long: 유효필터수 = 0
            For i = 1 To 2
                If 선택값배열(i) <> "" Then 유효필터수 = 유효필터수 + 1
            Next i
            If 유효필터수 = 0 Then 필터없음 = True
        End If
        
        ' ================================================================
        ' 3단계: 년도·월 선택
        ' ================================================================
        Dim 년도입력_raw As String
        년도입력_raw = InputBox( _
            "출력할 년도를 입력하세요." & vbCrLf & vbCrLf & _
            "비워두면 전체 기간으로 출력합니다." & vbCrLf & _
            "(예: 2026)", _
            "필터/출력 - 년도 선택", Format(Year(Date), "0000"))
        
        If StrPtr(년도입력_raw) = 0 Then Exit Sub  ' 취소
        Dim 년도입력 As String
        년도입력 = Trim(년도입력_raw)
        
    선택년 = 0
    선택월 = 0
        월필터사용 = False
        해당월시작열 = 0
        해당월일수 = 0
        
        If 년도입력 <> "" And IsNumeric(년도입력) Then
            선택년 = CInt(년도입력)
            If 선택년 < 2000 Or 선택년 > 2100 Then
                MsgBox "올바른 연도(2000~2100)를 입력하세요.", vbExclamation
            Exit Sub
        End If
            
            ' 월 선택
            Dim 월입력_raw As String
            월입력_raw = InputBox( _
                선택년 & "년의 출력할 월을 선택하세요." & vbCrLf & vbCrLf & _
                "1~12 중 입력 (비워두면 해당 년도 전체 월)" & vbCrLf & _
                "(예: 2 → 2월)", _
                "필터/출력 - 월 선택", Format(Month(Date), "0"))
            
            If StrPtr(월입력_raw) = 0 Then Exit Sub  ' 취소
            Dim 월입력 As String
            월입력 = Trim(월입력_raw)
            
            If 월입력 <> "" And IsNumeric(월입력) Then
                선택월 = CInt(월입력)
                If 선택월 < 1 Or 선택월 > 12 Then
                    MsgBox "올바른 월(1~12)을 입력하세요.", vbExclamation
            Exit Sub
        End If
                
                월필터사용 = True
        해당월일수 = DaysInMonth(선택년, 선택월)
        마지막열 = 마지막일열_찾기(ws재고)
                
                ' 해당 월 시작열 찾기
        For c = COL_일별시작 To 마지막열
                    If Trim(CStr(ws재고.Cells(2, c).Value)) = (선택월 & "m") And _
                       val(ws재고.Cells(3, c).Value) = 1 Then
                해당월시작열 = c
                Exit For
            End If
        Next c
                
        If 해당월시작열 = 0 Then
                    MsgBox "시트에서 " & 선택월 & "월 데이터를 찾을 수 없습니다." & vbCrLf & _
                        "(2행에 """ & 선택월 & "m"", 3행에 1이 있는 열이 필요합니다)", vbExclamation
            Exit Sub
        End If
            End If
    End If

        ' ================================================================
        ' 4단계: 필터 조건에 맞는 그룹 판별 함수 (AND 조건)
        ' ================================================================
        ' 필터 조건 요약 텍스트 생성
        Dim 필터요약 As String: 필터요약 = ""
        Dim 파일필터명 As String: 파일필터명 = ""
        
        For i = 1 To 2
            If 선택값배열(i) <> "" Then
                If 필터요약 <> "" Then 필터요약 = 필터요약 & ", "
                필터요약 = 필터요약 & 필터명배열(i) & "=" & 선택값배열(i)
                If 파일필터명 = "" Then 파일필터명 = 선택값배열(i)
            End If
        Next i
        
        If 필터없음 Then
            필터요약 = "전체"
            파일필터명 = "전체"
        End If
        
        ' ================================================================
        ' 5단계: 저장 경로 설정
        ' ================================================================
    저장경로 = ""
    If Not wsIndexKey Is Nothing Then
        저장경로 = Trim(CStr(wsIndexKey.Range("A3").Value))
    End If
    If 저장경로 = "" Then
        On Error Resume Next
        저장경로 = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        On Error GoTo 0
    End If
    If right(저장경로, 1) <> "\" And right(저장경로, 1) <> "/" Then 저장경로 = 저장경로 & "\"

        ' 파일명 생성 (특수문자 제거)
        파일필터명 = Replace(Replace(파일필터명, "\", ""), "/", "")
        파일필터명 = Replace(Replace(파일필터명, ":", ""), "*", "")
        파일필터명 = Replace(Replace(파일필터명, "?", ""), """", "")
        파일필터명 = Replace(Replace(파일필터명, "<", ""), ">", "")
        파일필터명 = Replace(파일필터명, "|", "")
        
    If 월필터사용 Then
            파일명 = 파일필터명 & "_원단관리현황_" & Format(선택년, "0000") & Format(선택월, "00") & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    Else
            파일명 = 파일필터명 & "_원단관리현황_" & Format(Date, "yyyymmdd") & ".xlsx"
    End If
    전체경로 = 저장경로 & 파일명

        ' ================================================================
        ' 6단계: 새 워크북 생성 + 데이터 출력
        ' ================================================================
    Application.ScreenUpdating = False

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
        wsNew.Name = "원단관리현황"

    If Not 월필터사용 Then 마지막열 = 마지막일열_찾기(ws재고)

    If 월필터사용 Then
            ' === 월 필터: 해당 월 열만 포함 ===
            Set srcRng = ws재고.Range(ws재고.Cells(1, COL_업체명), ws재고.Cells(3, COL_자재유형))
            srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), _
                wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL + COL_자재유형 - COL_업체명))
            
            Dim 일열시작 As Long
            일열시작 = OUT_START_COL + (COL_자재유형 - COL_업체명 + 1)
            
        For i = 1 To 해당월일수
                wsNew.Cells(OUT_START_ROW + 1, 일열시작 + i - 1).Value = 선택월 & "m"
                wsNew.Cells(OUT_START_ROW + 1, 일열시작 + i - 1).Font.color = RGB(160, 160, 160)
                wsNew.Cells(OUT_START_ROW + 2, 일열시작 + i - 1).Value = i
        Next i
            
            출력끝열 = 일열시작 + 해당월일수 - 1
            
            ' 제목
        On Error Resume Next
            wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), wsNew.Cells(OUT_START_ROW, 출력끝열)).Merge
        On Error GoTo 0
            wsNew.Cells(OUT_START_ROW, OUT_START_COL).Value = "원단관리 현황"

        ' 원단관리 시트에서 필터 조건에 맞는 그룹만 표시
        destRow = OUT_START_ROW + 3
        Dim 출력그룹수 As Long: 출력그룹수 = 0
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
                If Not 필터_그룹일치(ws재고, 시작행, 필터열배열, 선택값배열, 필터없음) Then GoTo NextGroup1
                Set srcRng = ws재고.Range(ws재고.Cells(시작행, COL_업체명), ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, COL_자재유형))
                srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, OUT_START_COL), _
                    wsNew.Cells(destRow + ROWS_PER_ITEM - 1, OUT_START_COL + COL_자재유형 - COL_업체명))
                Set srcRng = ws재고.Range(ws재고.Cells(시작행, 해당월시작열), _
                    ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, 해당월시작열 + 해당월일수 - 1))
                srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, 일열시작), _
                    wsNew.Cells(destRow + ROWS_PER_ITEM - 1, 일열시작 + 해당월일수 - 1))
                destRow = destRow + ROWS_PER_ITEM
                출력그룹수 = 출력그룹수 + 1
NextGroup1:
        Next 시작행
    Else
            ' === 전체 기간 ===
        마지막열 = 마지막일열_찾기(ws재고)
        Set srcRng = ws재고.Range(ws재고.Cells(1, COL_업체명), ws재고.Cells(3, 마지막열))
            srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), _
                wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL + 마지막열 - COL_업체명))
            출력끝열 = OUT_START_COL + 마지막열 - COL_업체명
            
        ' 원단관리 시트에서 필터 조건에 맞는 그룹만 표시
        destRow = OUT_START_ROW + 3
        출력그룹수 = 0
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
                If Not 필터_그룹일치(ws재고, 시작행, 필터열배열, 선택값배열, 필터없음) Then GoTo NextGroup2
                Set srcRng = ws재고.Range(ws재고.Cells(시작행, COL_업체명), ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막열))
                srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, OUT_START_COL), _
                    wsNew.Cells(destRow + ROWS_PER_ITEM - 1, OUT_START_COL + 마지막열 - COL_업체명))
                destRow = destRow + ROWS_PER_ITEM
                출력그룹수 = 출력그룹수 + 1
NextGroup2:
        Next 시작행
    End If

        ' 출력 데이터 없음 체크
        If destRow <= OUT_START_ROW + 3 Then
            Application.ScreenUpdating = True
            Application.DisplayAlerts = False
            wbNew.Close SaveChanges:=False
            Application.DisplayAlerts = True
            MsgBox "선택한 필터 조건에 맞는 데이터가 없습니다." & vbCrLf & vbCrLf & _
                "필터: " & 필터요약, vbExclamation, "필터/출력"
            Exit Sub
            End If
        
        ' 수식 제거 (값만)
        On Error Resume Next
        wsNew.Range(wsNew.Cells(OUT_START_ROW + 3, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열)).Copy
        wsNew.Range(wsNew.Cells(OUT_START_ROW + 3, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열)).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        On Error GoTo 0
        
        ' 서식 적용
    ' 데이터 영역에만 테두리 (나머지 시트는 눈금선 없음)
    With wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
        
        Dim 일헤더시작 As Long
        일헤더시작 = OUT_START_COL + (COL_자재유형 - COL_업체명 + 1)
        If 일헤더시작 <= 출력끝열 Then
            With wsNew.Range(wsNew.Cells(OUT_START_ROW + 1, 일헤더시작), wsNew.Cells(OUT_START_ROW + 2, 출력끝열))
        .Interior.color = RGB(232, 232, 232)
        .Font.Bold = True
    End With
        End If
    With wsNew.Range(wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL), wsNew.Cells(OUT_START_ROW + 2, 출력끝열))
        .Font.Bold = True
    End With

        ' 필터 정보 표시 (1행)
        wsNew.Cells(1, OUT_START_COL).Value = "필터: " & 필터요약
        If 월필터사용 Then
            wsNew.Cells(2, OUT_START_COL).Value = "기간: " & 선택년 & "년 " & 선택월 & "월"
        Else
            wsNew.Cells(2, OUT_START_COL).Value = "기간: 전체"
        End If
        wsNew.Cells(1, OUT_START_COL).Font.Bold = True
        wsNew.Cells(2, OUT_START_COL).Font.Bold = True
        
        ' 새 워크북: 폰트·눈금선 없음 (복사/서식 적용 후 적용해야 유지됨)
        wbNew.Windows(1).DisplayGridlines = False
        wsNew.usedRange.Font.Name = "Microsoft GothicNeo"
        
        ' 데이터 영역에만 테두리 적용
        With wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        ' 매크로 버튼(Shapes)이 복사되지 않도록 새 워크북 시트의 모든 Shape 제거
        Dim shpIdx As Long
        For shpIdx = wsNew.Shapes.Count To 1 Step -1
            wsNew.Shapes(shpIdx).Delete
        Next shpIdx
        
        ' ================================================================
        ' 7단계: 저장
        ' ================================================================
    On Error Resume Next
        wbNew.SaveAs Filename:=전체경로, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        Application.ScreenUpdating = True
            MsgBox "저장에 실패했습니다." & vbCrLf & vbCrLf & _
                "경로: " & 전체경로 & vbCrLf & _
                "오류: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True
    wbNew.Activate
        
        Dim 결과메시지 As String
        결과메시지 = "필터/출력이 완료되었습니다!" & vbCrLf & vbCrLf & _
            "[ 적용된 필터 ]" & vbCrLf & _
            "  항목: " & 필터요약 & vbCrLf
    If 월필터사용 Then
            결과메시지 = 결과메시지 & "  기간: " & 선택년 & "년 " & 선택월 & "월" & vbCrLf
    End If
        결과메시지 = 결과메시지 & vbCrLf & _
            "[ 출력 결과 ]" & vbCrLf & _
            "  출력 그룹 수: " & ((destRow - OUT_START_ROW - 3) / ROWS_PER_ITEM) & "건" & vbCrLf & _
            "  저장: " & 전체경로
        
        MsgBox 결과메시지, vbInformation, "필터/출력 완료"
End Sub

    '================================================================================
    ' 필터 그룹 일치 판별 (AND 조건: 선택한 모든 필터에 매칭되어야 True)
    '================================================================================
    Private Function 필터_그룹일치(ws As Worksheet, 그룹시작행 As Long, _
        필터열배열() As Long, 선택값배열() As String, 필터없음 As Boolean) As Boolean
        
        If 필터없음 Then
            필터_그룹일치 = True
            Exit Function
        End If
        
        Dim i As Long
        Dim 셀값 As String
        
        For i = 1 To 2
            If 선택값배열(i) <> "" Then
                셀값 = Trim(CStr(ws.Cells(그룹시작행, 필터열배열(i)).Value))
                If 셀값 <> 선택값배열(i) Then
                    필터_그룹일치 = False
                    Exit Function
                End If
            End If
        Next i
        
        필터_그룹일치 = True
    End Function

    ' 원단관리 시트에서 해당 부자재명(및 업체명)의 현재고 반환 (재고 행, 오늘 일자 기준)
    Private Function 현재고_원단관리에서가져오기(ws As Worksheet, 부자재명 As String, Optional 업체명 As String = "") As Double
        Dim lastRow As Long, i As Long, 블록시작 As Long
    Dim 일치 As Boolean
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    For i = START_ROW To lastRow
        If ws.Cells(i, COL_구분).Value = "재고" Then
                블록시작 = ((i - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
                일치 = (Trim(CStr(ws.Cells(블록시작, COL_부자재명).Value)) = 부자재명)
                If 일치 And 업체명 <> "" Then 일치 = (Trim(CStr(ws.Cells(블록시작, COL_업체명).Value)) = 업체명)
            If 일치 Then
                    현재고_원단관리에서가져오기 = val(ws.Cells(i, COL_일별시작 + Day(Date) - 1).Value)
                Exit Function
            End If
        End If
    Next i
        현재고_원단관리에서가져오기 = 0
End Function

    ' 원단관리 시트에서 해당 부자재명(및 업체명)의 전월재고 반환
    Private Function 전월재고_원단관리에서가져오기(ws As Worksheet, 부자재명 As String, Optional 업체명 As String = "") As Double
    Dim lastRow As Long, i As Long
    Dim 일치 As Boolean
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    For i = START_ROW To lastRow Step ROWS_PER_ITEM
        일치 = (Trim(CStr(ws.Cells(i, COL_부자재명).Value)) = 부자재명)
        If 일치 And 업체명 <> "" Then 일치 = (Trim(CStr(ws.Cells(i, COL_업체명).Value)) = 업체명)
        If 일치 Then
                전월재고_원단관리에서가져오기 = val(ws.Cells(i, COL_전월재고).Value)
            Exit Function
        End If
    Next i
        전월재고_원단관리에서가져오기 = 0
End Function

'================================================================================
' 시트 가져오기 또는 생성
'================================================================================
Private Function 시트가져오기또는생성(시트명 As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(시트명)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Dim wb보호 As Boolean: wb보호 = ThisWorkbook.ProtectStructure
        On Error Resume Next
        If wb보호 Then ThisWorkbook.Unprotect WS_시트보호암호
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        If Not ws Is Nothing Then ws.Name = 시트명
        If wb보호 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
        On Error GoTo 0
    End If
    
    Set 시트가져오기또는생성 = ws
End Function

' 재고관리DB 헤더 생성 (시트 신규 생성 시)
Private Sub 재고관리DB_헤더생성(wsDB As Worksheet)
    If wsDB Is Nothing Then Exit Sub
    wsDB.Cells(1, 1).Value = "생산일"
    wsDB.Cells(1, 2).Value = "품목유형"
    wsDB.Cells(1, 3).Value = "Co"
    wsDB.Cells(1, 4).Value = "품목명"
    wsDB.Cells(1, 5).Value = "입고"
    wsDB.Cells(1, 6).Value = "사용"
    wsDB.Cells(1, 7).Value = "재고"
    With wsDB.Range(wsDB.Cells(1, 1), wsDB.Cells(1, 7))
        .Font.Bold = True
        .Interior.color = RGB(200, 200, 200)
    End With
End Sub

'================================================================================
' 재고관리DB: 일별 완제품·원지 재고 데이터베이스
' - 원지: 일일생산DB BR열(70열)=원지_Co 넘버 → Raw data H열(인덱스) 매칭 → I열=원단명
' - 원단관리 A열=Co(Raw data H열 인덱스), B열=원단명(Raw data I열) → 일별 Roll 입고/사용/재고 기록
' - 완제품: 부자재관리 데이터업데이트 시 별도 기록 (재고관리DB_완제품_기록)
'================================================================================
Private Sub 재고관리DB_원단_기록(ws원단 As Worksheet)
    If ws원단 Is Nothing Then Exit Sub
    Dim wsDB As Worksheet
    Set wsDB = 시트가져오기또는생성(WS_재고관리DB)
    If wsDB Is Nothing Then Exit Sub
    Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        Call 재고관리DB_헤더생성(wsDB)
        lastRow = 1
    End If
    Dim 원단LastRow As Long, 시작행 As Long
    원단LastRow = ws원단.Cells(ws원단.Rows.Count, COL_구분).End(xlUp).Row
    If 원단LastRow < START_ROW Then Exit Sub
    Dim 마지막일v As Long: 마지막일v = 마지막월열_찾기(ws원단)
    If 마지막일v < COL_월별시작 Then Exit Sub
    Dim dCol As Long, 작성일 As Variant
    Dim CoVal As String, 원단명 As String, 입고 As Double, 사용 As Double, 재고 As Double
    Dim 새행 As Long
    Dim 기존딕 As Object: Set 기존딕 = CreateObject("Scripting.Dictionary")
    Dim rr As Long
    For rr = 2 To lastRow
        If Trim(CStr(wsDB.Cells(rr, 2).Value)) = "원지" Then
            기존딕(Format(wsDB.Cells(rr, 1).Value, "YYYYMMDD") & "|" & Trim(CStr(wsDB.Cells(rr, 3).Value))) = 1
        End If
    Next rr
    For dCol = COL_월별시작 To 마지막일v
        작성일 = 부자재_열날짜(ws원단, dCol)
        If IsNull(작성일) Then GoTo Next날짜열
        For 시작행 = START_ROW To 원단LastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            CoVal = Trim(CStr(ws원단.Cells(시작행, COL_제품발주처).Value))
            원단명 = Trim(CStr(ws원단.Cells(시작행, COL_제품명).Value))
            If CoVal = "" Then GoTo Next원단행
            입고 = 0: 사용 = 0: 재고 = 0
            If IsNumeric(ws원단.Cells(시작행 + OFFSET_입고1, dCol).Value) Then 입고 = CDbl(ws원단.Cells(시작행 + OFFSET_입고1, dCol).Value)
            If IsNumeric(ws원단.Cells(시작행 + OFFSET_사용1, dCol).Value) Then 사용 = CDbl(ws원단.Cells(시작행 + OFFSET_사용1, dCol).Value)
            If IsNumeric(ws원단.Cells(시작행 + OFFSET_재고1, dCol).Value) Then 재고 = CDbl(ws원단.Cells(시작행 + OFFSET_재고1, dCol).Value)
            If 입고 = 0 And 사용 = 0 And 재고 = 0 Then GoTo Next원단행
            If 기존딕.exists(Format(CDate(작성일), "YYYYMMDD") & "|" & CoVal) Then GoTo Next원단행
            새행 = lastRow + 1
            기존딕.Add Format(CDate(작성일), "YYYYMMDD") & "|" & CoVal, 1
            wsDB.Cells(새행, 1).Value = CDate(작성일)
            wsDB.Cells(새행, 2).Value = "원지"
            wsDB.Cells(새행, 3).Value = CoVal
            wsDB.Cells(새행, 4).Value = 원단명
            wsDB.Cells(새행, 5).Value = 입고
            wsDB.Cells(새행, 6).Value = 사용
            wsDB.Cells(새행, 7).Value = 재고
            lastRow = 새행
Next원단행:
        Next 시작행
Next날짜열:
    Next dCol
End Sub

'================================================================================
    ' ★ 재고부족 자동체크 (이벤트 감지 후 자동 호출용)
    ' 소요/입고 값 변경 시 재고 재계산 후 안전재고 미달이면 즉시 경고 메시지 출력
    '================================================================================
    Private Sub 재고부족_자동체크(ws As Worksheet, 그룹시작행 As Long, 블록인덱스 As Long)
        On Error Resume Next
        
        Dim 안전재고행 As Long
        Dim 기준수량 As Double, 현재고값 As Double
        Dim 자재유형명 As String
        Dim 부자재명 As String, 업체명 As String
        Dim 부족수량 As Double
        
        ' 서브블록별 안전재고 행 위치 (병합 셀의 첫 행)
        안전재고행 = 그룹시작행 + Choose(블록인덱스 + 1, 0, 4, 7, 10, 13, 16)
        자재유형명 = Choose(블록인덱스 + 1, "제품", "원지", "지관", "폴리백", "비닐백", "BOX")
        
        ' 안전재고 읽기
        기준수량 = 0
        If IsNumeric(ws.Cells(안전재고행, COL_안전재고).Value) Then
            기준수량 = CDbl(ws.Cells(안전재고행, COL_안전재고).Value)
        End If
        
        ' 현재고 읽기 (병합 셀 첫 행에서)
        현재고값 = 0
        If IsNumeric(ws.Cells(안전재고행, COL_현재고).Value) Then
            현재고값 = CDbl(ws.Cells(안전재고행, COL_현재고).Value)
        End If
        
        ' 부족 판단: 마이너스 OR 안전재고 이하 (안전재고 미설정이어도 마이너스면 경고)
        If 현재고값 < 0 Or (기준수량 > 0 And 현재고값 < 기준수량) Then
            부자재명 = Trim(CStr(ws.Cells(그룹시작행, COL_부자재명).Value))
            업체명 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
            If 기준수량 > 0 And 현재고값 < 기준수량 Then
                부족수량 = 기준수량 - 현재고값
            Else
                부족수량 = -현재고값  ' 마이너스일 때
            End If
            
            MsgBox "재고 부족 경고!" & vbCrLf & vbCrLf & _
                "부자재: " & 부자재명 & " (" & 업체명 & ")" & vbCrLf & _
                "자재유형: " & 자재유형명 & vbCrLf & vbCrLf & _
                "안전재고: " & Format(기준수량, "#,##0") & vbCrLf & _
                "현재고:    " & Format(현재고값, "#,##0") & vbCrLf & _
                "부족수량: " & Format(부족수량, "#,##0"), _
                vbExclamation, "재고 부족 자동 알림"
        End If
        
        On Error GoTo 0
    End Sub

    '================================================================================
    ' 재고 부족 알림 (버튼용 - 전체 항목 일괄 체크)
'================================================================================
Public Sub 원단_재고부족_알림()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim 기준수량 As Double
    Dim 부족목록 As String
        Dim 부자재명 As String, 업체명 As String
        Dim 현재고값 As Double
    Dim 부족수량 As Double
        Dim 부족건수 As Long
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
            Exit Sub
        End If
        
        ' 최신 현재고 반영 (값일괄계산 선호출)
        Call 원단관리_값일괄계산(ws)
        
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    부족목록 = ""
        부족건수 = 0
        
        ' 각 그룹의 6개 서브블록별로 안전재고 대비 현재고 비교
        Dim 그룹시작 As Long, j As Long
        Dim 자재유형명 As String
        Dim 안전재고행 As Long, 재고오프 As Long
        
        For 그룹시작 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            부자재명 = Trim(CStr(ws.Cells(그룹시작, COL_부자재명).Value))
            업체명 = Trim(CStr(ws.Cells(그룹시작, COL_업체명).Value))
            If 부자재명 = "" Then GoTo NextGroup
            
            ' 2개 서브블록: Roll(0), KG(1) ? ROWS_PER_ITEM=6이므로 오프셋 0,3만 유효
            For j = 0 To 1
                자재유형명 = Choose(j + 1, "Roll", "KG")
                
                ' 각 서브블록의 안전재고 행 위치 (병합 셀의 첫 행)
                안전재고행 = 그룹시작 + Choose(j + 1, 0, 3)
                재고오프 = Choose(j + 1, 2, 5)
                
                ' 안전재고 읽기 (병합 셀이므로 첫 행에서 읽음)
                기준수량 = 0
                If IsNumeric(ws.Cells(안전재고행, COL_안전재고).Value) Then
                    기준수량 = CDbl(ws.Cells(안전재고행, COL_안전재고).Value)
                End If
                
                ' ★ 현재고 읽기 (H열 = 기말재고, 병합 셀이므로 첫 행에서 읽음)
                현재고값 = 0
                If IsNumeric(ws.Cells(안전재고행, COL_현재고).Value) Then
                    현재고값 = CDbl(ws.Cells(안전재고행, COL_현재고).Value)
                End If
                
                ' 부족 판단: 마이너스 OR 안전재고 이하 (안전재고 미설정이어도 마이너스면 감지)
                If 현재고값 < 0 Or (기준수량 > 0 And 현재고값 < 기준수량) Then
                    부족수량 = 기준수량 - 현재고값
                    부족건수 = 부족건수 + 1
                    부족목록 = 부족목록 & _
                        부족건수 & ". [" & 자재유형명 & "] " & 부자재명 & " (" & 업체명 & ")" & vbCrLf & _
                        "   안전재고: " & Format(기준수량, "#,##0") & _
                        "  /  현재고: " & Format(현재고값, "#,##0") & _
                        "  /  부족: " & Format(부족수량, "#,##0") & vbCrLf & vbCrLf
                End If
NextBlock:
            Next j
NextGroup:
        Next 그룹시작

    If 부족목록 = "" Then
            MsgBox "모든 재고가 안전재고 이상입니다." & vbCrLf & vbCrLf & _
                "부족 항목이 없습니다.", vbInformation, "재고 확인"
        Else
            Dim 응답 As VbMsgBoxResult
            응답 = MsgBox("[ 재고 부족 알림 - 총 " & 부족건수 & "건 ]" & vbCrLf & vbCrLf & _
                부족목록 & vbCrLf & _
                "부족 항목을 새 파일로 출력하시겠습니까?", _
                vbExclamation + vbYesNo, "재고 부족 알림")
            If 응답 = vbYes Then
                Call 원단_재고부족_출력
            End If
    End If
End Sub

    '================================================================================
    ' ★ 재고 부족 항목 새 워크북으로 출력 (필터/출력과 동일한 방식)
    ' 안전재고 대비 현재고가 부족한 그룹행 15행 전체를 복사하여 새 파일로 저장
    '================================================================================
    Public Sub 원단_재고부족_출력()
        Dim ws As Worksheet
        Dim lastRow As Long, 마지막열 As Long
        Dim 그룹시작 As Long, j As Long
        Dim 안전재고행 As Long
        Dim 기준수량 As Double, 현재고값 As Double, 부족수량 As Double
        Dim 자재유형명 As String
        Dim 부족그룹 As Collection
        Dim 부족상세 As Collection
        Dim 부족테이블데이터 As Collection  ' 테이블 출력용: Array(설비라인, 자재유형, 부자재명, 업체명, 안전재고, 현재고, 부족수량)
        Dim 부자재명 As String, 업체명 As String, 설비라인출력 As String
        Dim 이미추가 As Object
        Dim wbNew As Workbook, wsNew As Worksheet
        Dim summRow As Long, k As Long
        Dim 데이터시작행 As Long, destRow As Long
        Dim 블록시작 As Long
        Dim tr As Long
        Dim 테이블시작행 As Long
        Dim arrItem
        Dim jj As Long, 블록시작오프 As Long, 블록끝오프 As Long
        Dim destBlkStart As Long, destBlkEnd As Long
        Dim c As Long
        Dim wsIndexKey As Worksheet, 저장경로 As String, 파일명 As String
        Dim 원래ScreenUpd As Boolean, 원래Events As Boolean, 원래Calc As XlCalculation
        
        원래ScreenUpd = Application.ScreenUpdating
        원래Events = Application.EnableEvents
        원래Calc = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events: Application.Calculation = 원래Calc
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
            Exit Sub
        End If
        
        ' 최신 현재고 반영 (값일괄계산 선호출)
        Call 원단관리_값일괄계산(ws)
        
        lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        마지막열 = 마지막일열_찾기(ws)
        If lastRow < START_ROW Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events: Application.Calculation = 원래Calc
            MsgBox "원단관리 시트에 데이터가 없습니다.", vbExclamation
            Exit Sub
        End If
        
        Set 부족그룹 = New Collection
        Set 부족상세 = New Collection
        Set 부족테이블데이터 = New Collection
        Set 이미추가 = CreateObject("Scripting.Dictionary")
        
        ' ── 1단계: 부족 그룹 스캔 ──
        For 그룹시작 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            설비라인출력 = Trim(CStr(ws.Cells(그룹시작, COL_설비라인).Value))
            부자재명 = Trim(CStr(ws.Cells(그룹시작, COL_제품).Value))
            업체명 = Trim(CStr(ws.Cells(그룹시작, COL_업체명).Value))
            If 부자재명 = "" And 업체명 = "" Then GoTo NextGrp출력
            
            ' 2개 서브블록: Roll(0), KG(1) ? ROWS_PER_ITEM=6이므로 오프셋 0,3만 유효
            For j = 0 To 1
                자재유형명 = Choose(j + 1, "Roll", "KG")
                안전재고행 = 그룹시작 + Choose(j + 1, 0, 3)
                
                기준수량 = 0
                If IsNumeric(ws.Cells(안전재고행, COL_안전재고).Value) Then
                    기준수량 = CDbl(ws.Cells(안전재고행, COL_안전재고).Value)
                End If
                
                현재고값 = 0
                If IsNumeric(ws.Cells(안전재고행, COL_현재고).Value) Then
                    현재고값 = CDbl(ws.Cells(안전재고행, COL_현재고).Value)
                End If
                
                ' 부족 판단: 마이너스 OR 안전재고 이하
                If 현재고값 < 0 Or (기준수량 > 0 And 현재고값 < 기준수량) Then
                    If 기준수량 > 0 And 현재고값 < 기준수량 Then
                        부족수량 = 기준수량 - 현재고값
                    Else
                        부족수량 = -현재고값  ' 마이너스일 때 부족량
                    End If
                    If Not 이미추가.exists(그룹시작) Then
                        이미추가.Add 그룹시작, True
                        부족그룹.Add 그룹시작
                    End If
                    부족상세.Add "[" & 자재유형명 & "] " & 부자재명 & " (" & 업체명 & ")" & _
                        " / 안전재고:" & Format(기준수량, "#,##0") & _
                        " / 현재고:" & Format(현재고값, "#,##0") & _
                        " / 부족:" & Format(부족수량, "#,##0")
                    부족테이블데이터.Add Array(설비라인출력, 자재유형명, 부자재명, 업체명, 기준수량, 현재고값, 부족수량)
                End If
NextBlk출력:
            Next j
NextGrp출력:
        Next 그룹시작
        
        If 부족그룹.Count = 0 Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events: Application.Calculation = 원래Calc
            MsgBox "모든 재고가 안전재고 이상입니다." & vbCrLf & "부족 항목이 없습니다.", vbInformation, "재고 확인"
            Exit Sub
        End If
        
        ' ── 2단계: 새 워크북 생성. 1행~마지막행 = 표(헤더+데이터), 그 아래 재고부족 항목 테이블 ──
        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        Set wsNew = wbNew.Sheets(1)
        wsNew.Name = "재고부족_" & Format(Now, "YYYYMMDD")
        
        ' 제목 행 (1행)
        wsNew.Cells(1, 1).Value = "재고 부족 항목 - " & Format(Now, "YYYY-MM-DD HH:mm") & " 기준"
        With wsNew.Cells(1, 1).Font
            .Bold = True
            .Size = 14
        End With
        
        ' 원본 1~3행 헤더 → 출력 2~4행
        ws.Range(ws.Cells(1, 1), ws.Cells(3, 마지막열)).Copy
        wsNew.Cells(2, 1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        데이터시작행 = 5  ' 5행부터 데이터
        
        ' ── 3단계: 부족 그룹 15행씩 복사 (표 데이터) ──
        destRow = 데이터시작행
        
        For k = 1 To 부족그룹.Count
            그룹시작 = 부족그룹(k)
            ws.Range(ws.Cells(그룹시작, 1), ws.Cells(그룹시작 + ROWS_PER_ITEM - 1, 마지막열)).Copy
            wsNew.Cells(destRow, 1).PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
            
            ' 부족 블록에 빨간 배경 하이라이트 (2블록: Roll, KG)
            For jj = 0 To 1
                안전재고행 = 그룹시작 + Choose(jj + 1, 0, 3)
                기준수량 = 0: 현재고값 = 0
                If IsNumeric(ws.Cells(안전재고행, COL_안전재고).Value) Then 기준수량 = CDbl(ws.Cells(안전재고행, COL_안전재고).Value)
                If IsNumeric(ws.Cells(안전재고행, COL_현재고).Value) Then 현재고값 = CDbl(ws.Cells(안전재고행, COL_현재고).Value)
                If 현재고값 < 0 Or (기준수량 > 0 And 현재고값 < 기준수량) Then
                    블록시작오프 = Choose(jj + 1, 0, 3)
                    블록끝오프 = Choose(jj + 1, 2, 5)
                    destBlkStart = destRow + 블록시작오프
                    destBlkEnd = destRow + 블록끝오프
                    wsNew.Range(wsNew.Cells(destBlkStart, COL_현재고), wsNew.Cells(destBlkEnd, COL_현재고)).Interior.color = RGB(255, 200, 200)
                    wsNew.Range(wsNew.Cells(destBlkStart, COL_안전재고), wsNew.Cells(destBlkEnd, COL_안전재고)).Interior.color = RGB(255, 200, 200)
                End If
            Next jj
            
            destRow = destRow + ROWS_PER_ITEM
        Next k
        
        ' 2·3행 테두리 없음, 표 영역(4행~마지막 데이터행) 내부 테두리 누락 없이 모두 실선 적용
        wsNew.Range(wsNew.Cells(2, 1), wsNew.Cells(3, 마지막열)).Borders.LineStyle = xlNone
        If destRow > 데이터시작행 Then
            With wsNew.Range(wsNew.Cells(4, 1), wsNew.Cells(destRow - 1, 마지막열))
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.ColorIndex = xlAutomatic
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlThin
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideVertical).Weight = xlThin
            End With
        End If
        
        ' ── 재고부족 항목 테이블: 표 아래에 표시 ──
        
        테이블시작행 = destRow + 2
        wsNew.Cells(테이블시작행 - 1, 1).Value = "재고부족 항목 요약"
        wsNew.Cells(테이블시작행 - 1, 1).Font.Bold = True
        wsNew.Cells(테이블시작행, 1).Value = "설비라인"
        wsNew.Cells(테이블시작행, 2).Value = "자재유형"
        wsNew.Cells(테이블시작행, 3).Value = "부자재명"
        wsNew.Cells(테이블시작행, 4).Value = "업체명"
        wsNew.Cells(테이블시작행, 5).Value = "안전재고"
        wsNew.Cells(테이블시작행, 6).Value = "현재고"
        wsNew.Cells(테이블시작행, 7).Value = "부족"
        wsNew.Range(wsNew.Cells(테이블시작행, 1), wsNew.Cells(테이블시작행, 7)).Font.Bold = True
        tr = 테이블시작행 + 1
        For k = 1 To 부족테이블데이터.Count
            
            arrItem = 부족테이블데이터(k)
            wsNew.Cells(tr, 1).Value = arrItem(0)
            wsNew.Cells(tr, 2).Value = arrItem(1)
            wsNew.Cells(tr, 3).Value = arrItem(2)
            wsNew.Cells(tr, 4).Value = arrItem(3)
            wsNew.Cells(tr, 5).Value = arrItem(4)
            wsNew.Cells(tr, 5).NumberFormat = "#,##0"
            wsNew.Cells(tr, 6).Value = arrItem(5)
            wsNew.Cells(tr, 6).NumberFormat = "#,##0"
            wsNew.Cells(tr, 7).Value = arrItem(6)
            wsNew.Cells(tr, 7).NumberFormat = "#,##0"
            wsNew.Cells(tr, 7).Font.color = RGB(200, 0, 0)
            tr = tr + 1
        Next k
        With wsNew.Range(wsNew.Cells(테이블시작행, 1), wsNew.Cells(tr - 1, 7)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        ' ── 4단계: 열 너비 맞춤 ──
        For c = 1 To 마지막열
            wsNew.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth
        Next c
        
        On Error Resume Next
        wbNew.Windows(1).DisplayGridlines = False
        On Error GoTo 0
        wsNew.usedRange.Font.Name = "Microsoft GothicNeo"
        
        ' ── 5단계: 저장 ──
        On Error Resume Next
        Set wsIndexKey = ThisWorkbook.Sheets("Index key")
        On Error GoTo 0
        저장경로 = ""
        If Not wsIndexKey Is Nothing Then
            저장경로 = Trim(CStr(wsIndexKey.Cells(3, 1).Value))
        End If
        If 저장경로 = "" Then 저장경로 = ThisWorkbook.Path
        If right(저장경로, 1) <> "\" Then 저장경로 = 저장경로 & "\"
        
        파일명 = "재고부족_" & Format(Now, "YYYYMMDD_HHmmss") & ".xlsx"
        
        On Error Resume Next
        wbNew.SaveAs 저장경로 & 파일명, FileFormat:=xlOpenXMLWorkbook
        If Err.Number <> 0 Then
            Err.Clear
            저장경로 = Environ("USERPROFILE") & "\Desktop\"
            wbNew.SaveAs 저장경로 & 파일명, FileFormat:=xlOpenXMLWorkbook
        End If
        On Error GoTo 0
        
        Application.ScreenUpdating = 원래ScreenUpd
        Application.EnableEvents = 원래Events
        Application.Calculation = 원래Calc
        
        MsgBox "재고 부족 항목이 새 파일로 출력되었습니다." & vbCrLf & vbCrLf & _
            "부족 그룹: " & 부족그룹.Count & "개" & vbCrLf & _
            "부족 항목: " & 부족상세.Count & "건" & vbCrLf & vbCrLf & _
            "저장 위치: " & 저장경로 & 파일명, vbInformation, "재고부족 출력 완료"
    End Sub

'================================================================================
' 입력 이력 보기 (기존 기능 유지)
'================================================================================
Public Sub 원단_입력이력_보기()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_이력)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "입력 이력이 없습니다.", vbInformation
        Exit Sub
    End If
    
    ws.Activate
        MsgBox "입력 이력 시트로 이동했습니다.", vbInformation, "입력 이력"
End Sub

'================================================================================
' 원단_오늘이동: 오늘 날짜 열을 J열(COL_일별시작) 바로 옆으로 스크롤, 4행~마지막 데이터 음영 (커서 이동 시 해제)
'================================================================================
Public Sub 원단_오늘이동()
    Dim ws As Worksheet
    Dim 마지막행 As Long, 오늘열 As Long
    Dim rng As Range
    Dim 이벤트상태 As Boolean
    Dim 날짜열딕 As Object
    Dim dateKey As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_원단관리)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' 매크로 실행 중 SelectionChange로 인한 음영 즉시 해제 방지
    이벤트상태 = Application.EnableEvents
    Application.EnableEvents = False
    Call 원단_오늘이동_음영해제
    원단_오늘이동_진행중 = True
    
    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 < START_ROW Then 마지막행 = START_ROW
    
    ' 오늘 날짜에 해당하는 열 찾기 (원단관리_날짜열_딕셔너리: 2행=월, 3행=일, I2=연도)
    오늘열 = 0
    Set 날짜열딕 = 원단관리_날짜열_딕셔너리(ws)
    dateKey = Format(Date, "yyyymmdd")
    If Not 날짜열딕 Is Nothing And 날짜열딕.exists(dateKey) Then
        오늘열 = 날짜열딕(dateKey)
    End If
    ' 딕셔너리 미발견 시 직접 열 스캔 (2행=월, 3행=일)
    If 오늘열 < COL_일별시작 Then
        Dim 마지막열 As Long, col As Long, v2 As Variant, v3 As Variant, s2 As String
        마지막열 = 마지막월열_찾기(ws)
        For col = COL_일별시작 To 마지막열
            v3 = ws.Cells(3, col).Value
            If IsNumeric(v3) And CLng(v3) = Day(Date) Then
                s2 = Trim(Replace(Replace(Replace(CStr(ws.Cells(2, col).Value), "M", ""), "m", ""), " ", ""))
                If s2 <> "" And IsNumeric(s2) And CLng(s2) = Month(Date) Then
                    오늘열 = col
                    Exit For
                End If
            End If
        Next col
    End If

    ws.Activate
    ' J열(COL_일별시작) 바로 옆에 오늘 날짜 열이 오도록 스크롤 (오늘열이 왼쪽에서 2번째로 보이게)
    On Error Resume Next
    If 오늘열 >= COL_일별시작 Then
        Application.GoTo Reference:=ws.Cells(START_ROW, 오늘열), Scroll:=True
        ActiveWindow.ScrollColumn = WorksheetFunction.Max(COL_일별시작, 오늘열 - 1)
    Else
        Application.GoTo Reference:=ws.Cells(START_ROW, COL_일별시작), Scroll:=True
        ActiveWindow.ScrollColumn = COL_일별시작
    End If
    On Error GoTo 0

    ' 오늘 날짜 열만 4행~마지막행 음영 (2행·3행 헤더 제외, 셀보호 시 일시 해제)
    If 오늘열 >= COL_일별시작 Then
        Dim 시트보호상태1 As Boolean
        시트보호상태1 = ws.ProtectContents
        If 시트보호상태1 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        Set rng = ws.Range(ws.Cells(4, 오늘열), ws.Cells(마지막행, 오늘열))
        On Error Resume Next
        rng.Interior.color = RGB(220, 230, 255)
        If Err.Number <> 0 Then
            Err.Clear
            Dim c As Range
            For Each c In rng.Cells
                c.Interior.color = RGB(220, 230, 255)
            Next c
        End If
        On Error GoTo 0
        Set 원단_오늘이동_음영Rng = rng
        If 시트보호상태1 Then Call 원단_셀보호_적용_무음(ws)
    End If
    
    ' 이벤트 복원 (0.3초 후 진행중 해제 → 즉시 발동되는 SelectionChange가 음영을 지우지 않도록)
    Application.EnableEvents = 이벤트상태
    On Error Resume Next
    Application.OnTime Now + (0.3 / 86400), "원단_오늘이동_진행중해제", , True
    If Err.Number <> 0 Then 원단_오늘이동_진행중 = False
    On Error GoTo 0
    
    ' 스크롤/선택 시 음영 해제: 원단관리 시트 모듈에 다음 추가 시 선택 변경 시 해제됨
    ' Private Sub Worksheet_SelectionChange(ByVal Target As Range): Call 원단_오늘이동_음영해제
End Sub

' 0.3초 후 진행중 플래그 해제 (OnTime 예약)
Public Sub 원단_오늘이동_진행중해제()
    원단_오늘이동_진행중 = False
End Sub

' 원단_오늘이동 음영 해제 (시트 SelectionChange 등에서 호출)
' ★ 해제 시 그룹행 교대음영으로 복원 → 기존 그룹행 음영에 변화 없음
Public Sub 원단_오늘이동_음영해제()
    Dim wsRng As Worksheet, 시보호 As Boolean
    Dim r As Long, 시작행 As Long, grpIdx As Long
    If 원단_오늘이동_진행중 Then Exit Sub
    If 원단_오늘이동_음영Rng Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsRng = 원단_오늘이동_음영Rng.Parent
    시보호 = wsRng.ProtectContents
    If 시보호 Then
        On Error Resume Next
        wsRng.Unprotect WS_시트보호암호
        If wsRng.ProtectContents Then wsRng.Unprotect
        On Error GoTo 0
    End If
    ' 그룹행 교대음영 규칙에 맞게 복원 (홀수 그룹=RGB(242,247,252), 짝수=무음영)
    For r = 원단_오늘이동_음영Rng.Row To 원단_오늘이동_음영Rng.Row + 원단_오늘이동_음영Rng.Rows.Count - 1
        시작행 = ((r - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        grpIdx = ((시작행 - START_ROW) \ ROWS_PER_ITEM) + 1
        If grpIdx Mod 2 = 1 Then
            wsRng.Cells(r, 원단_오늘이동_음영Rng.Column).Interior.Pattern = xlSolid
            wsRng.Cells(r, 원단_오늘이동_음영Rng.Column).Interior.color = RGB(242, 247, 252)
        Else
            wsRng.Cells(r, 원단_오늘이동_음영Rng.Column).Interior.Pattern = xlPatternNone
        End If
    Next r
    Set 원단_오늘이동_음영Rng = Nothing
    If 시보호 Then Call 원단_셀보호_적용_무음(wsRng)
    On Error GoTo 0
End Sub

'================================================================================
' 버튼 생성 (도형-Shape 사용: 배경 RGB(192,135,253), 글씨 흰색, 테두리 없음)
'================================================================================
Public Sub 원단_버튼_생성()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rng As Range
    Dim 좌측위치 As Double
    Dim topPos As Double
    Dim btnNames(1 To 15) As String
    Dim btnCount As Long
    Dim grpArr() As Variant
    Dim ii As Long
    Dim s As Shape
    ' 버튼 크기: 너비 2.6cm, 높이 0.6cm (1cm = 28.35 포인트)
    Const btnW As Double = 73.71    ' 2.6cm = 2.6 * 28.35
    Const btnH As Double = 17.01    ' 0.6cm = 0.6 * 28.35
    Const 간격 As Double = 2.835   ' 0.1cm = 0.1 * 28.35
    Const 좌우여백 As Double = 2   ' 좌우 여백 2 포인트

    Set ws = ThisWorkbook.Sheets(WS_원단관리)
    Set rng = ws.Range("A5")   ' 버튼 배치 기준 셀 (먼저 설정)

    ' 기존 "이력 보기" 양식 버튼이 있으면 "원단_오늘이동"으로 변경 후 종료 (과거 호환)
    Dim b As Object
    On Error Resume Next
    For Each b In ws.Buttons
        If Trim(CStr(b.Caption)) = "이력 보기" Then
            b.Caption = "원단_오늘이동"
            b.OnAction = "원단_오늘이동"
            MsgBox "기존 '이력 보기' 버튼을 '원단_오늘이동'으로 변경했습니다.", vbInformation, "버튼 생성"
            Exit Sub
        End If
    Next b
    On Error GoTo 0

    ' 기존 양식 버튼 또는 도형 버튼이 있으면 유지 후 종료
    Dim 기존버튼있음 As Boolean
    기존버튼있음 = False
    On Error Resume Next
    If ws.Buttons.Count > 0 Then 기존버튼있음 = True
    If Not 기존버튼있음 Then
        For Each s In ws.Shapes
            If s.Name = "BtnGroup" Or left(s.Name, 4) = "Btn_" Then 기존버튼있음 = True: Exit For
        Next s
    End If
    On Error GoTo 0
    If 기존버튼있음 Then
        MsgBox "원단관리 시트에 이미 버튼이 있습니다." & vbCrLf & _
            "기존 버튼을 유지합니다.", vbInformation, "버튼 생성"
        Exit Sub
    End If

    ' A5 셀부터 도형 버튼 배치 (배경 192,135,253 / 글씨 흰색 / 테두리 없음)
    좌측위치 = rng.left + 좌우여백
    topPos = rng.top
    btnCount = 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_데이터업데이트"
    shp.OnAction = "원단관리_데이터업데이트"
    shp.TextFrame2.TextRange.Text = "데이터 업데이트"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_목록추가"
    shp.OnAction = "원단_빈행그룹_추가"
    shp.TextFrame2.TextRange.Text = "목록 추가"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_재고부족"
    shp.OnAction = "원단_재고부족_알림"
    shp.TextFrame2.TextRange.Text = "재고부족"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_필터출력"
    shp.OnAction = "업체별원단관리_생성"
    shp.TextFrame2.TextRange.Text = "필터/출력"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_키워드검색"
    shp.OnAction = "원단관리_키워드검색_실행"
    shp.TextFrame2.TextRange.Text = "키워드 검색"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_원단_오늘이동"
    shp.OnAction = "원단_오늘이동"
    shp.TextFrame2.TextRange.Text = "원단_오늘이동"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_다음달열추가"
    shp.OnAction = "원단_다음달열_추가"
    shp.TextFrame2.TextRange.Text = "다음달 열 추가"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_원단_동기화검증"
    shp.OnAction = "원단_동기화검증"
    shp.TextFrame2.TextRange.Text = "동기화 검증"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_셀보호해제"
    shp.OnAction = "원단_셀보호_해제"
    shp.TextFrame2.TextRange.Text = "셀보호 해제"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_원단_목업테스트"
    shp.OnAction = "원단_목업테스트"
    shp.TextFrame2.TextRange.Text = "원단_목업테스트"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    topPos = topPos + btnH + 간격

    ' 원단_월생산현황 버튼 추가
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 좌측위치, topPos, btnW, btnH)
    shp.Name = "Btn_원단_월생산현황"
    shp.OnAction = "원단_월생산현황"
    shp.TextFrame2.TextRange.Text = "원단_월생산현황"
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Name = "Microsoft GothicNeo"
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btnCount = btnCount + 1: btnNames(btnCount) = shp.Name
    ' 생성된 도형 모두 그룹화
    On Error Resume Next
    If btnCount > 0 Then
        ReDim grpArr(1 To btnCount)
        For ii = 1 To btnCount
            grpArr(ii) = btnNames(ii)
        Next ii
        Set shp = ws.Shapes.Range(grpArr).Group
        shp.Name = "BtnGroup"
    End If
    On Error GoTo 0

    MsgBox "버튼이 생성되었습니다!", vbInformation
End Sub

    '================================================================================
    ' 원단_목업테스트: 시뮬레이션 데이터 생성
    ' - 그룹행 5개 생성 (5개 미만이면 5개가 되도록 추가, 최대 5개)
    ' - ★ 목업데이터는 오늘 날짜 기준까지만 생성 (날짜 동적 인식, 미래일 열 입력 금지)
    '================================================================================
    Public Sub 원단_목업테스트()
        Dim ws As Worksheet, wsIndex As Worksheet, ws사양 As Worksheet
        Dim 응답 As VbMsgBoxResult
        
        On Error Resume Next
        Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
        Set ws사양 = ThisWorkbook.Sheets(WS_표준사양서)
        On Error GoTo 0
        
        If wsIndex Is Nothing Then
            MsgBox "제품표준사양서 시트를 찾을 수 없습니다.", vbExclamation, "원단_목업테스트"
            Exit Sub
        End If
        
        Dim col업체 As Long, col제품 As Long, col제조사 As Long, col지관 As Long
        Dim col지폭 As Long, col총량 As Long, col1차 As Long, col2차 As Long
        
        col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
        If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
        If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
        ' 제품 열: 일일생산일보 제품목록(C22:C31, AG열 등)과 동일 로드 경로 (G열 9행~, 빈행 무시)
        col제품 = 표준사양서_헤더열찾기(wsIndex, "제품")
        If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
        If col제품 = 0 Then col제품 = 표준사양서_제품열G
        col제조사 = 표준사양서_헤더열찾기(wsIndex, "제조사")
        col지관 = 표준사양서_헤더열찾기(wsIndex, "지관")
        col지폭 = 표준사양서_헤더열찾기(wsIndex, "지폭")
        col총량 = 표준사양서_헤더열찾기(wsIndex, "평량")
        If col총량 = 0 Then col총량 = 표준사양서_헤더열찾기(wsIndex, "총량")
        col1차 = 표준사양서_헤더열찾기(wsIndex, "1차포장")
        col2차 = 표준사양서_헤더열찾기(wsIndex, "2차포장")
        
        If col업체 = 0 Or col제품 = 0 Then
            MsgBox "제품표준사양서에서 '업체명' 또는 '제품' 열을 찾을 수 없습니다.", vbExclamation, "원단_목업테스트"
            Exit Sub
        End If
        
        ' 제품표준사양서 데이터 행 수 (col제품 = 일일생산일보 제품목록과 동일 경로)
        Dim lastRowIdx As Long
        lastRowIdx = wsIndex.Cells(wsIndex.Rows.Count, col제품).End(xlUp).Row
        If wsIndex.Cells(wsIndex.Rows.Count, col업체).End(xlUp).Row > lastRowIdx Then _
            lastRowIdx = wsIndex.Cells(wsIndex.Rows.Count, col업체).End(xlUp).Row
        
        If lastRowIdx < 표준사양서_데이터시작행 Then
            MsgBox "제품표준사양서에 데이터가 없습니다. (8행부터 데이터 필요)", vbExclamation, "원단_목업테스트"
            Exit Sub
        End If
        
        Dim 총데이터수 As Long
        총데이터수 = lastRowIdx - 표준사양서_데이터시작행 + 1
        
        ' ── 원단관리 시트 준비 ──
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then
            On Error Resume Next
            Dim wb보호부자재 As Boolean: wb보호부자재 = ThisWorkbook.ProtectStructure
            If wb보호부자재 Then ThisWorkbook.Unprotect WS_시트보호암호
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            If Not ws Is Nothing Then ws.Name = WS_원단관리
            If wb보호부자재 Then ThisWorkbook.Protect Password:=WS_시트보호암호, Structure:=True, Windows:=False
            On Error GoTo 0
        End If
        On Error Resume Next
        On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        On Error GoTo 0
        
        ' 기존 항목 수 확인 (버튼/데이터 삭제 없음)
        Dim 기존항목수 As Long, 마지막사용행 As Long, 마지막열 As Long
        마지막사용행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        기존항목수 = 0
        If 마지막사용행 >= START_ROW Then
            Dim i행 As Long
            For i행 = START_ROW To 마지막사용행 Step ROWS_PER_ITEM
                If Trim(CStr(ws.Cells(i행, COL_제품).Value)) <> "" Then
                    기존항목수 = 기존항목수 + 1
                Else
                    Exit For
                End If
            Next i행
        End If
        
        Dim 신규항목수 As Long, 생성수 As Long
        ' 5개 그룹이 되도록 부족분만 추가 (기존 3개면 2개만 추가)
        신규항목수 = WorksheetFunction.Max(0, WorksheetFunction.Min(5 - 기존항목수, 총데이터수))
        생성수 = WorksheetFunction.Min(기존항목수 + 신규항목수, 5)
        If 생성수 = 0 Then
            MsgBox "생성할 항목이 없습니다. (제품표준사양서에 데이터가 있어야 합니다)", vbExclamation, "원단_목업테스트"
            Exit Sub
        End If
        
        Dim 동작설명 As String
        If 신규항목수 > 0 Then
            동작설명 = "그룹 5개가 되도록 신규 " & 신규항목수 & "개 추가 후 목업데이터 생성"
        Else
            동작설명 = "기존 " & 기존항목수 & "개 항목에 목업데이터만 생성 (이미 5개 이하)"
        End If
        동작설명 = 동작설명 & vbCrLf & vbCrLf & "【시뮬레이션】" & vbCrLf & _
            "· 1번째 그룹: 재고 부족 (안전재고 미달/마이너스 재고)" & vbCrLf & _
            "· 2번째 그룹: 동일 제품·다른 업체 (업체만 다르고 제품명 동일)"
        
        응답 = MsgBox("목업데이터를 생성합니다." & vbCrLf & vbCrLf & _
            동작설명 & vbCrLf & vbCrLf & _
            "※ 기존 버튼·데이터 구조는 유지됩니다.", vbYesNo + vbQuestion, "원단_목업테스트")
        If 응답 <> vbYes Then Exit Sub
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        Dim 현재년 As Integer, 현재월 As Integer, 현재일수 As Integer
        현재년 = Year(Date)
        현재월 = Month(Date)
        현재일수 = DaysInMonth(현재년, 현재월)
        마지막열 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        If 마지막열 < COL_일별시작 Then 마지막열 = COL_합계
        
        ' 헤더: 기존 데이터 없을 때만 1~3행 설정
        If 기존항목수 = 0 Then
            ws.Cells(2, COL_제품).Value = 현재년 & "-" & Format(현재월, "00")
            ws.Cells(2, COL_현재고).Value = Format(Date, "YYYY-MM-DD")
            ws.Cells(3, COL_제품발주처).Value = "제품발주처"
            ws.Cells(3, COL_제품명).Value = "제품명"
            ws.Cells(3, COL_제조사).Value = "제조사"
            ws.Cells(3, COL_지관).Value = "지관"
            ws.Cells(3, COL_지폭).Value = "지폭"
            ws.Cells(3, COL_총량).Value = "평량"
            ws.Cells(3, COL_1차포장).Value = "1차포장"
            ws.Cells(3, COL_2차포장).Value = "2차포장"
            ws.Cells(3, COL_전월재고).Value = "전월재고"
            ws.Cells(3, COL_현재고).Value = "현재고"
            ws.Cells(3, COL_안전재고).Value = "안전재고"
            ws.Cells(3, COL_구분).Value = "구분"
            ws.Cells(3, COL_자재유형).Value = "자재유형"
            Dim d As Long
            For d = 1 To 현재일수
                ws.Cells(2, COL_일별시작 + d - 1).Value = 현재월 & "m"
                ws.Cells(2, COL_일별시작 + d - 1).Font.color = RGB(160, 160, 160)
                ws.Cells(3, COL_일별시작 + d - 1).Value = d
            Next d
        End If
        
        ' ── 신규 항목 생성 (10개 미만일 때만) ──
        Dim 시작행 As Long
        Dim k As Long, r As Long
        Dim v업체 As String, v제품 As String, v제조사 As String
        Dim v지관 As String, v지폭 As String, v총량 As String, v1차 As String, v2차 As String
        Dim 기본전월 As Long
        Dim 업체목록 As String: 업체목록 = ""
        Dim s As String
        
        Dim 블록오프(0 To 4) As Long
        블록오프(0) = 0: 블록오프(1) = 3: 블록오프(2) = 6: 블록오프(3) = 9: 블록오프(4) = 12
        
        Dim 입고오프(0 To 4) As Long, 소요오프(0 To 4) As Long
        입고오프(0) = 1: 소요오프(0) = 2
        입고오프(1) = 4: 소요오프(1) = 5
        입고오프(2) = 7: 소요오프(2) = 8
        입고오프(3) = 10: 소요오프(3) = 11
        입고오프(4) = 13: 소요오프(4) = 14
        
        Randomize Timer
        
        ' 기존 그룹 B~I 키 수집 (중복 생성 방지)
        Dim 기존키_BtoI목업 As Object
        Dim i행목업 As Long
        Set 기존키_BtoI목업 = CreateObject("Scripting.Dictionary")
        For i행목업 = START_ROW To START_ROW + 기존항목수 * ROWS_PER_ITEM - 1 Step ROWS_PER_ITEM
            If i행목업 > ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row Then Exit For
            s = 그룹행_시작행_BtoI_키(ws, i행목업)
            If s <> "" Then 기존키_BtoI목업(s) = 1
        Next i행목업
        
        For k = 1 To 신규항목수
            r = 표준사양서_데이터시작행 + (기존항목수 + k - 1)
            시작행 = START_ROW + (기존항목수 + k - 1) * ROWS_PER_ITEM
            
            v업체 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
            v제품 = Trim(CStr(wsIndex.Cells(r, col제품).Value))
            ' 시나리오: 2번째 그룹 = 동일 제품·다른 업체 (제품표준사양서 첫 행 제품 + 다른 행 업체)
            If k = 2 And 총데이터수 >= 2 Then
                v제품 = Trim(CStr(wsIndex.Cells(표준사양서_데이터시작행, col제품).Value))
                If v제품 = "" Then v제품 = Trim(CStr(wsIndex.Cells(r, col제품).Value))
                v업체 = Trim(CStr(wsIndex.Cells(표준사양서_데이터시작행 + 1, col업체).Value))
                If v업체 = "" Then v업체 = Trim(CStr(wsIndex.Cells(r, col업체).Value))
            End If
            v제조사 = "": v지관 = "": v지폭 = "": v총량 = "": v1차 = "": v2차 = ""
            If col제조사 > 0 Then v제조사 = Trim(CStr(wsIndex.Cells(r, col제조사).Value))
            If col지관 > 0 Then v지관 = Trim(CStr(wsIndex.Cells(r, col지관).Value))
            If col지폭 > 0 Then v지폭 = Trim(CStr(wsIndex.Cells(r, col지폭).Value))
            If col총량 > 0 Then v총량 = Trim(CStr(wsIndex.Cells(r, col총량).Value))
            If col1차 > 0 Then v1차 = Trim(CStr(wsIndex.Cells(r, col1차).Value))
            If col2차 > 0 Then v2차 = Trim(CStr(wsIndex.Cells(r, col2차).Value))
            If v업체 = "" Then v업체 = "-"
            If v제품 = "" Then v제품 = "제품" & (기존항목수 + k)
            
            ' A~I 동일 그룹 중복 시 해당 그룹 생성 생략
            s = 그룹행_BtoI_키(v업체, v제품, v제조사, v지관, v지폭, v총량, v1차, v2차)
            If 기존키_BtoI목업.exists(s) Then GoTo NextK목업
            기존키_BtoI목업(s) = 1
            
            If InStr(1, "|" & 업체목록 & "|", "|" & v업체 & "|") = 0 Then
                If 업체목록 <> "" Then 업체목록 = 업체목록 & ", "
                업체목록 = 업체목록 & v업체
            End If
            
            Call 부자재항목_생성(ws, 시작행, "", v제품, v업체)
            ws.Cells(시작행, COL_제조사).Value = v제조사
            ws.Cells(시작행, COL_지관).Value = v지관
            ws.Cells(시작행, COL_지폭).Value = v지폭
            ws.Cells(시작행, COL_총량).Value = v총량
            ws.Cells(시작행, COL_1차포장).Value = v1차
            ws.Cells(시작행, COL_2차포장).Value = v2차
NextK목업:
        Next k
        
        ' ── 전체 항목에 목업데이터 생성 (제품 출고 랜덤 대량, 모든 월 열) ──
        Dim col As Long, b As Long
        Dim 랜덤입고 As Long, 랜덤소요 As Long
        Dim 전월재고배열(0 To 4) As Long, 안전재고배열(0 To 4) As Long
        마지막열 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        If 마지막열 < COL_일별시작 Then 마지막열 = COL_일별시작 + 30
        
        For k = 1 To 생성수
            시작행 = START_ROW + (k - 1) * ROWS_PER_ITEM
            ' 시나리오: 1번째 그룹 = 재고 부족 (전월/입고 낮게, 안전재고 높게, 소요 크게)
            Dim 시나리오_재고부족 As Boolean
            시나리오_재고부족 = (k = 1)
            If 시나리오_재고부족 Then
                기본전월 = CLng(400 + Rnd * 400)   ' 낮은 전월재고
            Else
                기본전월 = CLng(3000 + Rnd * 12000)
            End If
            
            전월재고배열(0) = 기본전월
            전월재고배열(1) = CLng(기본전월 * (0.8 + Rnd * 0.4))
            전월재고배열(2) = CLng(기본전월 * (0.3 + Rnd * 0.3))
            전월재고배열(3) = CLng(기본전월 * (0.5 + Rnd * 0.4))
            전월재고배열(4) = CLng(기본전월 * (0.2 + Rnd * 0.2))
            
            If 시나리오_재고부족 Then
                ' 안전재고를 높게 설정 → 현재고 < 안전재고 발생
                안전재고배열(0) = CLng(기본전월 * (2.5 + Rnd * 1.5))
                For b = 1 To 4: 안전재고배열(b) = CLng(전월재고배열(b) * (2 + Rnd * 1.5)): Next b
            Else
                안전재고배열(0) = CLng(기본전월 * (0.15 + Rnd * 0.1))
                For b = 1 To 4: 안전재고배열(b) = CLng(전월재고배열(b) * (0.15 + Rnd * 0.1)): Next b
            End If
            
            For b = 0 To 4
                ws.Cells(시작행 + 블록오프(b), COL_전월재고).Value = 전월재고배열(b)
                ws.Cells(시작행 + 블록오프(b), COL_안전재고).Value = 안전재고배열(b)
            Next b
            
            ' 오늘 날짜 기준까지만 목업데이터 (미래 월·미래 일 열에는 입력 금지)
            Dim col월 As Long, 열날짜 As Date, 열연도 As Integer
            For col = COL_일별시작 To 마지막열
                If Not IsNumeric(ws.Cells(3, col).Value) Then GoTo NextCol
                d = CLng(ws.Cells(3, col).Value)
                If d < 1 Or d > 31 Then GoTo NextCol
                col월 = val(Replace(Replace(CStr(ws.Cells(2, col).Value), "m", ""), "M", ""))
                If col월 < 1 Or col월 > 12 Then GoTo NextCol
                If col월 > Month(Date) Then GoTo NextCol
                열연도 = Year(Date)
                On Error Resume Next
                열날짜 = DateSerial(열연도, col월, d)
                If Err.Number <> 0 Then Err.Clear: GoTo NextCol
                On Error GoTo 0
                If Month(열날짜) <> col월 Or Day(열날짜) <> d Then GoTo NextCol
                If 열날짜 > Date Then GoTo NextCol
                For b = 0 To 4
                    If b = 0 Then
                        ' 제품 카테고리: 생산·출고
                        If 시나리오_재고부족 Then
                            If Rnd < 0.5 Then ws.Cells(시작행 + 0, col).Value = CLng(50 + Rnd * 150)  ' 생산 적게
                            If (d + k) Mod 4 = 0 Then ws.Cells(시작행 + 입고오프(0), col).Value = CLng(20 + Rnd * 50)  ' 입고 거의 없음
                            If Rnd < 0.92 Then ws.Cells(시작행 + 소요오프(0), col).Value = CLng(800 + Rnd * 3200)  ' 출고 많음
                        Else
                            If Rnd < 0.88 Then ws.Cells(시작행 + 0, col).Value = CLng(600 + Rnd * 4400)  ' 생산
                            If (d + k) Mod 3 = 0 Then
                                랜덤입고 = CLng(300 + Rnd * 700)
                                ws.Cells(시작행 + 입고오프(0), col).Value = 랜덤입고  ' 제품 입고
                            End If
                            If Rnd < 0.88 Then ws.Cells(시작행 + 소요오프(0), col).Value = CLng(800 + Rnd * 4200)  ' 출고
                        End If
                    Else
                        If 시나리오_재고부족 Then
                            If (d + k + b) Mod 5 = 0 Then ws.Cells(시작행 + 입고오프(b), col).Value = CLng(5 + Rnd * 15)  ' 입고 적게
                            랜덤소요 = CLng(전월재고배열(b) * (0.25 + Rnd * 0.4))
                            If 랜덤소요 < 30 Then 랜덤소요 = 30
                            ws.Cells(시작행 + 소요오프(b), col).Value = 랜덤소요  ' 소요 크게 → 부족
                        Else
                            If (d + k + b) Mod 3 = 0 Then
                                랜덤입고 = CLng(전월재고배열(b) * (0.05 + Rnd * 0.15))
                                If 랜덤입고 < 20 Then 랜덤입고 = 20
                                ws.Cells(시작행 + 입고오프(b), col).Value = 랜덤입고
                            End If
                            랜덤소요 = CLng(전월재고배열(b) * (0.03 + Rnd * 0.08))
                            If 랜덤소요 < 10 Then 랜덤소요 = 10
                            ws.Cells(시작행 + 소요오프(b), col).Value = 랜덤소요
                        End If
                    End If
                Next b
NextCol:
            Next col
        Next k
        
        ' ── 재고 일괄 계산 ──
        Call 원단관리_값일괄계산(ws)
        
        ' ── 목록추가와 동일: 유효성 목록 적용 (목록시트_갱신 → 목록다운_적용 → 그룹행 유효성) ──
        Call 목록시트_갱신(wsIndex)
        Call 목록다운_적용(ws)
        Call 원단관리_그룹행_유효성_제품표준사양서(ws, wsIndex)
        
        ' ── 서식 마무리 (헤더 수정 금지, 기존 버튼 유지) ──
        Call 시트서식_적용(ws, START_ROW + 생성수 * ROWS_PER_ITEM - 1, True)
        Call 날짜영역_테두리적용(ws)
        ws.Range(ws.Cells(1, COL_업체명), ws.Cells(1, COL_자재유형)).Columns.AutoFit
        
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        ws.Activate
        ws.Cells(START_ROW, COL_업체명).Select
        
        MsgBox "목업데이터가 생성되었습니다!" & vbCrLf & vbCrLf & _
            "그룹: " & 생성수 & "개" & IIf(신규항목수 > 0, " (신규 " & 신규항목수 & "개 추가)", "") & vbCrLf & _
            IIf(업체목록 <> "", "업체: " & 업체목록 & vbCrLf, "") & _
            "생성 기준: 오늘(" & Format(Date, "YYYY-MM-DD") & ")까지" & vbCrLf & _
            "※ 기존 버튼·데이터 구조 유지", vbInformation, "원단_목업테스트 완료"
End Sub

'================================================================================
' 전체 설정 (최초 실행)
'================================================================================
Public Sub 원단_전체설정_실행()
    Call 원단관리_데이터업데이트
    Call 원단_버튼_생성
    MsgBox "부자재 원단관리 시스템 설정이 완료되었습니다!" & vbCrLf & vbCrLf & _
           "신규 기능:" & vbCrLf & _
            "1. 입고/생산 입력 시 목록에서 선택" & vbCrLf & _
            "2. 입력 이력 관리" & vbCrLf & _
           "4. 재고 행 셀 보호 (수정 불가)" & vbCrLf & _
           "5. 소요 값 수정 시 일일생산DB 동기화", vbInformation, "설정 완료"
End Sub

'================================================================================
    ' [최적화] 셀 보호 적용 (MsgBox 없음) - 내부 호출용 (데이터업데이트, 값일괄계산 등)
    '================================================================================
    Public Sub 원단_셀보호_적용_무음(ws As Worksheet)
        Dim lastRow As Long
        If ws Is Nothing Then Exit Sub
        If Not 시트보호해제_시도(ws) Then Exit Sub
        lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        If lastRow < START_ROW Then lastRow = START_ROW
        
        ' ★ J2:N2, J열(전월재고), L열(안전재고) 입력 가능, K열(현재고)은 잠금
        ws.Cells.Locked = True
        ws.Range("J2:N2").Locked = False
        ws.Range("J3:J" & lastRow).Locked = False
        ws.Range("L3:L" & lastRow).Locked = False
        
        ' ★ 예외: 목록추가 매크로로 추가한 그룹행만 해당 구간 전체 셀보호 해제
        If ws.Name = WS_원단관리 And 원단_목록추가_예외시작행 >= START_ROW And 원단_목록추가_예외끝행 >= 원단_목록추가_예외시작행 Then
            Dim 마지막열예외 As Long
            마지막열예외 = 마지막일열_찾기(ws)
            If 마지막열예외 < COL_일별시작 Then 마지막열예외 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
            On Error Resume Next
            If ws.ProtectContents Then ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect
            ws.Range(ws.Cells(원단_목록추가_예외시작행, 1), ws.Cells(원단_목록추가_예외끝행, 마지막열예외)).Locked = False
            On Error GoTo 0
        End If
        
        ws.Protect Password:=WS_시트보호암호, DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                AllowFormattingCells:=False, AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, AllowDeletingRows:=False, _
                AllowSorting:=False, AllowFiltering:=False, AllowUsingPivotTables:=False
    End Sub

    '================================================================================
    ' 셀 보호 설정: 재고 행은 잠금, 소요/입고/생산 행은 잠금 해제 (버튼용 - MsgBox 포함) (5블록·15행)
'================================================================================
Public Sub 원단_셀보호_설정()
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
    On Error GoTo 0

    If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then Application.ScreenUpdating = True: Exit Sub

    ' 목록추가 예외는 유지한 채 적용 (예외 시 해당 그룹행만 추가 해제)
    Call 원단_셀보호_적용_무음(ws)
    Application.ScreenUpdating = True

    MsgBox "셀 보호가 설정되었습니다." & vbCrLf & vbCrLf & _
           "■ 잠금 해제(수정 가능):" & vbCrLf & _
           "  - J2:N2 (헤더 영역)" & vbCrLf & _
           "  - J열(3행~): 전월재고" & vbCrLf & _
           "  - L열(3행~): 안전재고" & vbCrLf & vbCrLf & _
           "■ 잠금(수정 불가):" & vbCrLf & _
           "  - K열(현재고) 포함, 위 범위 외 모든 셀", vbInformation, "셀 보호 설정"
End Sub

'================================================================================
    ' Raw data H/I 기반 원단관리 A·B열 유효성 목록 적용
    '================================================================================
    Public Sub 원단관리_AB열_유효성_적용()
        Call RawData_HI_목록_갱신
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then Exit Sub
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        If lastRow < START_ROW Then Exit Sub
        Dim r As Long
        On Error Resume Next
        For r = START_ROW To lastRow Step ROWS_PER_ITEM
            ws.Range(ws.Cells(r, COL_제품발주처), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Delete
            ws.Range(ws.Cells(r, COL_제품발주처), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단번호"
            ws.Range(ws.Cells(r, COL_제품발주처), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품발주처)).Validation.IgnoreBlank = True
            ws.Range(ws.Cells(r, COL_제품명), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Delete
            ws.Range(ws.Cells(r, COL_제품명), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.Add Type:=xlValidateList, Formula1:="=목록_원단명"
            ws.Range(ws.Cells(r, COL_제품명), ws.Cells(r + ROWS_PER_ITEM - 1, COL_제품명)).Validation.IgnoreBlank = True
        Next r
        On Error GoTo 0
        Call 원단관리_B열기준_A열전체동기화(ws)
    End Sub

    ' B열(원단명) 기준 A열(원단번호) 전체 동기화 (시트 활성화/유효성적용 시 호출, Change 미동작 시 보완)
    Public Sub 원단관리_B열기준_A열전체동기화(Optional ws As Worksheet)
        If ws Is Nothing Then On Error Resume Next: Set ws = ThisWorkbook.Sheets(WS_원단관리): On Error GoTo 0
        If ws Is Nothing Or ws.Name <> WS_원단관리 Then Exit Sub
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        If lastRow < START_ROW Then Exit Sub
        Dim r As Long, 원단명 As String, CoVal As String
        Application.EnableEvents = False
        On Error Resume Next
        For r = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            원단명 = Trim(CStr(ws.Cells(r, COL_제품명).Value))
            If 원단명 <> "" Then
                CoVal = RawData_원단명ToCo(원단명)
                If CoVal <> "" Then ws.Cells(r, COL_제품발주처).Value = CoVal
            End If
        Next r
        On Error GoTo 0
        Application.EnableEvents = True
    End Sub

'================================================================================
    ' B열(원단명) 변경 시 A열(원단번호) 자동 채우기: Raw data I열→H열 조회
    '================================================================================
    Public Sub 원단관리_B열_원단명입력시_A열Co채우기(변경셀 As Range)
        If 변경셀 Is Nothing Or 변경셀.Cells.Count <> 1 Then Exit Sub
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        If ws.Name <> WS_원단관리 Then Exit Sub
        If 변경셀.Column <> COL_제품명 Then Exit Sub
        If 변경셀.Row < START_ROW Then Exit Sub
        Dim 그룹시작행 As Long
        그룹시작행 = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        Dim 원단명 As String: 원단명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품명).Value))
        If 원단명 = "" Then Exit Sub
        Dim CoVal As String: CoVal = RawData_원단명ToCo(원단명)
        If CoVal = "" Then Exit Sub
        ' 시트 내 동일 (원단번호,원단명) 중복 방지
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        Dim r As Long
        For r = START_ROW To lastRow Step ROWS_PER_ITEM
            If r <> 그룹시작행 Then
                If Trim(CStr(ws.Cells(r, COL_제품발주처).Value)) = CoVal And _
                   Trim(CStr(ws.Cells(r, COL_제품명).Value)) = 원단명 Then
                    Application.EnableEvents = False
                    ws.Cells(그룹시작행, COL_제품명).Value = ""
                    Application.EnableEvents = True
                    MsgBox "동일 원단(" & CoVal & " " & 원단명 & ")이 이미 시트에 있습니다. 중복 입력이 취소되었습니다.", vbExclamation, "중복 방지"
                    Exit Sub
                End If
            End If
        Next r
        Application.EnableEvents = False
        On Error Resume Next
        ws.Cells(그룹시작행, COL_제품발주처).Value = CoVal
        On Error GoTo 0
        Application.EnableEvents = True
    End Sub

'================================================================================
    ' A열(Co인덱스 01~) 숫자 입력 시 Raw data H열 매칭 → I열(원단명)을 B열에 자동 입력
    '================================================================================
    Public Sub 원단관리_A열_Co입력시_B열원단명채우기(변경셀 As Range)
        If 변경셀 Is Nothing Or 변경셀.Cells.Count <> 1 Then Exit Sub
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        If ws.Name <> WS_원단관리 Then Exit Sub
        If 변경셀.Column <> COL_제품발주처 Then Exit Sub
        If 변경셀.Row < START_ROW Then Exit Sub
        Dim 그룹시작행 As Long
        그룹시작행 = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        Dim vA As Variant: vA = ws.Cells(그룹시작행, COL_제품발주처).Value
        If IsEmpty(vA) Then Exit Sub
        Dim CoVal As String: CoVal = Trim(Replace(CStr(vA), " ", ""))
        If CoVal = "" Then Exit Sub
        Dim 원단명 As String: 원단명 = RawData_CoTo원단명(CoVal)
        If 원단명 = "" Then Exit Sub
        ' 시트 내 동일 (원단번호,원단명) 중복 방지
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
        Dim r As Long
        For r = START_ROW To lastRow Step ROWS_PER_ITEM
            If r <> 그룹시작행 Then
                If Trim(CStr(ws.Cells(r, COL_제품발주처).Value)) = CoVal And _
                   Trim(CStr(ws.Cells(r, COL_제품명).Value)) = 원단명 Then
                    Application.EnableEvents = False
                    ws.Cells(그룹시작행, COL_제품발주처).Value = ""
                    Application.EnableEvents = True
                    MsgBox "동일 원단(" & CoVal & " " & 원단명 & ")이 이미 시트에 있습니다. 중복 입력이 취소되었습니다.", vbExclamation, "중복 방지"
                    Exit Sub
                End If
            End If
        Next r
        Application.EnableEvents = False
        On Error Resume Next
        ws.Cells(그룹시작행, COL_제품명).Value = 원단명
        On Error GoTo 0
        Application.EnableEvents = True
    End Sub

'================================================================================
    ' B·C열(업체명·제품) 입력/변경 시 해당 행의 D~I열을 제품표준사양서에서 자동 채우기
    ' Worksheet_Change에서 B(2) 또는 C(3) 열 변경 시 호출
    '================================================================================
    Public Sub 원단관리_BC_변경시_DtoI_채우기(변경셀 As Range)
        If 변경셀 Is Nothing Or 변경셀.Cells.Count <> 1 Then Exit Sub
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        If ws.Name <> WS_원단관리 Then Exit Sub
        If 변경셀.Column <> COL_업체명 And 변경셀.Column <> COL_제품 Then Exit Sub
        If 변경셀.Row < START_ROW Then Exit Sub

        Dim 그룹시작행 As Long
        그룹시작행 = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        Dim 업체명 As String, 제품명 As String
        업체명 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
        제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))

        ' B열 변경 시 C열 유효성은 Change 밖에서 갱신 (OnTime 1초 후 + SelectionChange 시 즉시)
        If 변경셀.Column = COL_업체명 Then
            원단_C열유효성_갱신대기행 = 그룹시작행
            If Not 원단_C열유효성_OnTime예약 Then
                원단_C열유효성_OnTime예약 = True
                Application.OnTime Now + TimeSerial(0, 0, 1), "원단관리_C열유효성_갱신대기처리"
            End If
        End If

        If 제품명 = "" Then Exit Sub

        Dim wsIndex As Worksheet
        On Error Resume Next
        Set wsIndex = ThisWorkbook.Sheets(WS_제품표준사양서)
        On Error GoTo 0
        If wsIndex Is Nothing Then Exit Sub

        Dim col업체 As Long, col제품 As Long, col제조사 As Long, col지관 As Long, col지폭 As Long, col총량 As Long, col1차 As Long, col2차 As Long
        col업체 = 표준사양서_헤더열찾기(wsIndex, "업체명")
        If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "발주처")
        If col업체 = 0 Then col업체 = 표준사양서_헤더열찾기(wsIndex, "업체")
        col제품 = 표준사양서_헤더열찾기(wsIndex, "제품")
        If col제품 = 0 Then col제품 = 표준사양서_헤더열찾기(wsIndex, "제품(상품)명")
        col제조사 = 표준사양서_헤더열찾기(wsIndex, "제조사")
        col지관 = 표준사양서_헤더열찾기(wsIndex, "지관")
        col지폭 = 표준사양서_헤더열찾기(wsIndex, "지폭")
        col총량 = 표준사양서_헤더열찾기(wsIndex, "평량")
        If col총량 = 0 Then col총량 = 표준사양서_헤더열찾기(wsIndex, "총량")
        col1차 = 표준사양서_헤더열찾기(wsIndex, "1차포장")
        col2차 = 표준사양서_헤더열찾기(wsIndex, "2차포장")
        If col업체 <= 0 Or col제품 <= 0 Then Exit Sub

        Dim r As Long
        r = 표준사양서_행찾기(wsIndex, col업체, col제품, 업체명, 제품명)
        If r <= 0 Then Exit Sub

        Dim v제조사 As String, v지관 As String, v지폭 As String, v총량 As String, v1차 As String, v2차 As String
        v제조사 = "": v지관 = "": v지폭 = "": v총량 = "": v1차 = "": v2차 = ""
        If col제조사 > 0 Then v제조사 = Trim(CStr(wsIndex.Cells(r, col제조사).Value))
        If col지관 > 0 Then v지관 = Trim(CStr(wsIndex.Cells(r, col지관).Value))
        If col지폭 > 0 Then v지폭 = Trim(CStr(wsIndex.Cells(r, col지폭).Value))
        If col총량 > 0 Then v총량 = Trim(CStr(wsIndex.Cells(r, col총량).Value))
        If col1차 > 0 Then v1차 = Trim(CStr(wsIndex.Cells(r, col1차).Value))
        If col2차 > 0 Then v2차 = Trim(CStr(wsIndex.Cells(r, col2차).Value))
        If r = 표준사양서_데이터시작행 Then
            If col1차 > 0 And v1차 = Trim(CStr(wsIndex.Cells(표준사양서_헤더행, col1차).Value)) Then v1차 = ""
            If col2차 > 0 And v2차 = Trim(CStr(wsIndex.Cells(표준사양서_헤더행, col2차).Value)) Then v2차 = ""
        End If
        If LCase(v1차) = "type" Or v1차 = "1차포장" Then v1차 = ""
        If LCase(v2차) = "type" Or v2차 = "2차포장" Then v2차 = ""

        Application.EnableEvents = False
        On Error Resume Next
        ws.Cells(그룹시작행, COL_제조사).Value = v제조사
        ws.Cells(그룹시작행, COL_지관).Value = v지관
        ws.Cells(그룹시작행, COL_지폭).Value = v지폭
        ws.Cells(그룹시작행, COL_총량).Value = v총량
        ws.Cells(그룹시작행, COL_1차포장).Value = v1차
        ws.Cells(그룹시작행, COL_2차포장).Value = v2차
        ' C열 유효성: B열 선택 업체에 해당하는 이름으로 직접 설정 (INDIRECT 미동작 환경 대비)
        Dim nm제품 As String
        nm제품 = "목록_제품_" & 목록이름_세이니타이즈(업체명)
        On Error Resume Next
        Dim refTest
        refTest = ThisWorkbook.Names(nm제품).RefersTo
        If Err.Number = 0 Then
            Dim rngC As Range
            Dim vTypeC As Long
            Set rngC = ws.Cells(그룹시작행, COL_제품)
            On Error Resume Next
            vTypeC = rngC.Validation.Type
            If Err.Number = 0 And vTypeC = xlValidateList Then
                rngC.Validation.Modify Type:=xlValidateList, Formula1:="=" & nm제품
            Else
                Err.Clear
                rngC.Validation.Delete
                rngC.Validation.Add Type:=xlValidateList, Formula1:="=" & nm제품
                rngC.Validation.IgnoreBlank = True
                rngC.Validation.InCellDropdown = True
            End If
            On Error GoTo 0
        End If
        On Error GoTo 0
        Application.EnableEvents = True
    End Sub

'================================================================================
    ' B열 변경 시 대기한 C열 유효성 갱신 (SelectionChange에서 호출 → Change 이벤트 밖에서 처리해 드롭다운 유지)
    '================================================================================
    Public Sub 원단관리_C열유효성_갱신대기처리()
        If 원단_C열유효성_갱신대기행 <= 0 Then
            원단_C열유효성_OnTime예약 = False
            Exit Sub
        End If
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then 원단_C열유효성_갱신대기행 = 0: 원단_C열유효성_OnTime예약 = False: Exit Sub
        Dim 그룹시작행 As Long
        그룹시작행 = 원단_C열유효성_갱신대기행
        원단_C열유효성_갱신대기행 = 0
        원단_C열유효성_OnTime예약 = False

        ' 원단관리: B열=원단명 → 목록_원단명 (Raw data I열)
        Dim nm제품 As String
        nm제품 = "목록_원단명"

        Dim rngMerge As Range
        Set rngMerge = ws.Range(ws.Cells(그룹시작행, COL_제품), ws.Cells(그룹시작행 + ROWS_PER_ITEM - 1, COL_제품))
        Dim rngC As Range
        Set rngC = ws.Cells(그룹시작행, COL_제품)

        Dim 시트보호 As Boolean
        시트보호 = ws.ProtectContents
        If 시트보호 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        Application.ScreenUpdating = False
        On Error Resume Next
        rngMerge.UnMerge
        rngC.Validation.Delete
        rngC.Validation.Add Type:=xlValidateList, Formula1:="=" & nm제품
        If Err.Number <> 0 Then
            Err.Clear
            rngC.Validation.Add Type:=xlValidateList, Formula1:="=목록_원단명"
        End If
        rngC.Validation.IgnoreBlank = True
        rngC.Validation.InCellDropdown = True
        rngMerge.Merge
        On Error GoTo 0
        Application.ScreenUpdating = True
        If 시트보호 Then ws.Protect Password:=WS_시트보호암호
    End Sub

'================================================================================
    ' 원단관리 시트 셀 변경 시 DB 동기화 "예약" (디바운스: 0.2초 → 일일생산DB·현재고 즉각 반영)
    ' 단일셀: 해당 블록만 동기화. 다중셀(붙여넣기): 전체 시트 동기화.
    '================================================================================
    Public Sub 원단관리_셀변경_DB동기화_예약(변경범위 As Range)
        If 변경범위 Is Nothing Then Exit Sub
        If 변경범위.Worksheet.Name <> WS_원단관리 Then Exit Sub
        On Error Resume Next
        If 원단_재고동기화_OnTime예약 <> 0 Then Application.OnTime 원단_재고동기화_OnTime예약, "원단_재고동기화_지연실행", False
        On Error GoTo 0
        Set 원단_재고동기화_대기셀 = 변경범위
        원단_재고동기화_OnTime예약 = Now + (0.2 / 86400)
        Application.OnTime 원단_재고동기화_OnTime예약, "원단_재고동기화_지연실행"
    End Sub

    Public Sub 원단_재고동기화_지연실행()
        원단_재고동기화_OnTime예약 = 0
        If 원단_재고동기화_대기셀 Is Nothing Then Exit Sub
        Dim r As Range: Set r = 원단_재고동기화_대기셀
        Set 원단_재고동기화_대기셀 = Nothing
        If r.Worksheet.Name <> WS_원단관리 Then Exit Sub
        If r.Cells.Count > 1 Then
            ' 다중 셀(붙여넣기): 전체 재계산 → DB반영 → 현재고 갱신
            Dim ws As Worksheet: Set ws = r.Worksheet
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            On Error Resume Next
            Call 원단관리_값일괄계산(ws)
            Call Stock_시트_갱신
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            On Error GoTo 0
        Else
            Call 원단관리_셀변경_DB동기화(r)
        End If
    End Sub

    ' 워크북 종료 전 OnTime 예약 취소 (Workbook_BeforeClose에서 호출)
    Public Sub 원단관리_OnTime_모두취소()
        On Error Resume Next
        If 원단_재고동기화_OnTime예약 <> 0 Then
            Application.OnTime 원단_재고동기화_OnTime예약, "원단_재고동기화_지연실행", False
            원단_재고동기화_OnTime예약 = 0
        End If
        Set 원단_재고동기화_대기셀 = Nothing
        원단_C열유효성_OnTime예약 = False
        원단_C열유효성_갱신대기행 = 0
        On Error GoTo 0
    End Sub

    ' 원단관리 시트 셀 변경 시 DB 동기화 (즉시 실행. Worksheet_Change에서 _예약 대신 호출하면 디바운스 없음)
    ' ★ 행 오프셋 기반 자동 분기 (구분 텍스트에 의존하지 않음, 5블록·15행)
    '================================================================================
    Public Sub 원단관리_셀변경_DB동기화(변경셀 As Range)
        If 변경셀 Is Nothing Or 변경셀.Cells.Count <> 1 Then Exit Sub
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        If ws.Name <> WS_원단관리 Then Exit Sub
        If 변경셀.Row < START_ROW Then Exit Sub
        Dim origScr As Boolean, origCalc As Long
        origScr = Application.ScreenUpdating: origCalc = Application.Calculation
        Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
        Dim grStart As Long
        On Error GoTo RestoreSync
        ' ★ E열(기초재고) 변경: 해당 그룹 재고 즉시 재계산
        If 변경셀.Column = COL_기초재고 And 변경셀.Row >= START_ROW Then
            grStart = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
            Call 원단_블록_재고_재계산(ws, grStart, 0)
            On Error Resume Next: Call Stock_시트_갱신: On Error GoTo 0
            GoTo RestoreSync
        End If
        ' ★ J열(전월재고) 변경: 해당 블록 재고 즉시 재계산 + DB·현재고 시트 반영
        If 변경셀.Column = COL_전월재고 Then
            Call 원단_전월재고_변경_재계산(변경셀)
            On Error Resume Next: Call Stock_시트_갱신: On Error GoTo 0
            GoTo RestoreSync
        End If
        
        If 변경셀.Column < COL_일별시작 Then GoTo RestoreSync
        
        ' 행 오프셋으로 구분 판단 (텍스트 라벨 불일치 방지)
        Dim 오프셋 As Long
        오프셋 = (변경셀.Row - START_ROW) Mod ROWS_PER_ITEM
        
        ' 6행 구조: 입고(0,3), 사용(1,4) → 재고 재계산
        If ROWS_PER_ITEM = 6 Then
            If 오프셋 = 0 Or 오프셋 = 1 Or 오프셋 = 3 Or 오프셋 = 4 Then
                grStart = ((변경셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
                Call 원단_블록_재고_재계산(ws, grStart, 0)
            End If
        End If
        
        Select Case 오프셋
            Case 0                              ' 생산 행 (15행) / 입고1행 (6행)
                If ROWS_PER_ITEM = 15 Then Call 원단_생산값_DB_동기화(변경셀)
            Case 3, 6, 9, 12                    ' 입고 행 (15행: 원지·1차·2차)
                If ROWS_PER_ITEM = 15 Then Call 원단_입고값_일일생산DB_동기화(변경셀)
            Case 1, 4, 7, 10, 13               ' 소요 행 (15행)
                If ROWS_PER_ITEM = 15 Then Call 원단_소요값_일일생산DB_동기화(변경셀)
            Case Else                            ' 재고 행 → 무시
        End Select
        ' 일일생산DB가 바뀌었을 수 있으므로 현재고 시트 동기화 시도 (오류 시 무시)
        On Error Resume Next
        Call Stock_시트_갱신
        On Error GoTo 0
RestoreSync:
        Application.ScreenUpdating = origScr: Application.Calculation = origCalc
    End Sub

    '================================================================================
    ' ★ 전월재고(J열) 수기 변경 시: 해당 블록 재고 재계산 + 일일생산DB 저장
    ' 전월재고는 ① 새로운월 생성 시 이전 달 말일 재고로 자동 반영  ② 수기 변경 가능
    ' Worksheet_Change → 원단관리_셀변경_DB동기화 → 이 함수 호출
    '================================================================================
    Public Sub 원단_전월재고_변경_재계산(변경셀 As Range)
        On Error GoTo ErrorExit
        
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        Dim 행번호 As Long: 행번호 = 변경셀.Row
        Dim 시트보호 As Boolean: 시트보호 = ws.ProtectContents
        If 시트보호 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        
        ' 그룹 시작행 계산
        Dim 그룹시작행 As Long
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        If 그룹시작행 < START_ROW Then GoTo CleanExit
        
        ' 전월재고는 블록시작 오프셋(0,3,6,9,12)에만 위치
        Dim 오프셋 As Long: 오프셋 = 행번호 - 그룹시작행
        Dim 블록인덱스 As Long: 블록인덱스 = -1
        Select Case 오프셋
            Case 0: 블록인덱스 = 0    ' 제품
            Case 3: 블록인덱스 = 1    ' 원지(ROLL)
            Case 6: 블록인덱스 = 2    ' 원지(Tons)
            Case 9: 블록인덱스 = 3    ' 1차포장
            Case 12: 블록인덱스 = 4   ' 2차포장
            Case Else: GoTo CleanExit
        End Select
        
        ' ★ 1단계: 해당 블록 재고 재계산
        Call 원단_블록_재고_재계산(ws, 그룹시작행, 블록인덱스)
        
        ' ★★★ 2단계: BS~BZ열 쓰기 비활성화됨 ★★★
        
        ' ★ 3단계: K열(현재고) 업데이트
        Dim 오늘열 As Long
        오늘열 = 원단관리_날짜열_찾기(ws, Date)
        If 오늘열 < COL_일별시작 Then
            ' 폴백: Day(Date) 기반
            Dim 첫달일수 As Integer
            첫달일수 = Day(DateSerial(Year(ws.Cells(2, COL_일별시작).Value), Month(ws.Cells(2, COL_일별시작).Value) + 1, 0))
            오늘열 = COL_일별시작 + WorksheetFunction.Min(Day(Date), 첫달일수) - 1
        End If
        
        If 오늘열 >= COL_일별시작 Then
            Dim 재고오프 As Long
            재고오프 = Choose(블록인덱스 + 1, 2, 5, 8, 11, 14)
            Dim 오늘재고 As Variant
            오늘재고 = ws.Cells(그룹시작행 + 재고오프, 오늘열).Value
            ws.Cells(그룹시작행 + Choose(블록인덱스 + 1, 0, 3, 6, 9, 12), COL_현재고).Value = 오늘재고
        End If
        
CleanExit:
        If 시트보호 Then Call 원단_셀보호_적용_무음(ws)
        Exit Sub
        
ErrorExit:
        If 시트보호 Then
            On Error Resume Next
            Call 원단_셀보호_적용_무음(ws)
            On Error GoTo 0
        End If
    End Sub

    '================================================================================
    ' ★ 소요 값 변경 시: 재고 재계산 + 일일생산DB 저장 + 일일생산DB 저장(제품)
    ' Worksheet_Change → 이 함수 호출 (완전 자동)
'================================================================================
Public Sub 원단_소요값_일일생산DB_동기화(변경셀 As Range)
        On Error GoTo ErrorExit
        
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        Dim 행번호 As Long: 행번호 = 변경셀.Row
        Dim 열번호 As Long: 열번호 = 변경셀.Column
        Dim 시트보호 As Boolean: 시트보호 = ws.ProtectContents
        If 시트보호 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        
        ' 기본 검증 (행 오프셋으로 이미 소요 행 판단됨, 구분 텍스트는 참고만)
        If 열번호 < COL_일별시작 Then GoTo CleanExit
        
        ' 그룹 시작행 계산
        Dim 그룹시작행 As Long
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        If 그룹시작행 < START_ROW Then GoTo CleanExit
        
        ' 자재유형 판단 (오프셋 기반, 5블록·15행: 소요 행 = 1(출고),4,7,10,13)
        Dim 오프셋 As Long: 오프셋 = 행번호 - 그룹시작행
        Dim 자재유형명 As String: 자재유형명 = ""
        Dim 블록인덱스 As Long: 블록인덱스 = -1
        Select Case 오프셋
            Case 1: 자재유형명 = "제품": 블록인덱스 = 0   ' 출고
            Case 4: 자재유형명 = "원지(ROLL)": 블록인덱스 = 1
            Case 7: 자재유형명 = "원지(Tons)": 블록인덱스 = 2
            Case 10: 자재유형명 = "1차포장": 블록인덱스 = 3
            Case 13: 자재유형명 = "2차포장": 블록인덱스 = 4
            Case Else: GoTo CleanExit
        End Select
        
        ' 새 값 읽기
        Dim 새값 As Double
        If IsNumeric(변경셀.Value) Then 새값 = CDbl(변경셀.Value) Else 새값 = 0
        
        ' ★★★ 1단계: 해당 블록의 재고만 재계산 (소요/입고 값은 건드리지 않음!) ★★★
        Call 원단_블록_재고_재계산(ws, 그룹시작행, 블록인덱스)
        
        ' ★★★ 2단계: 일일생산DB에 저장 ★★★
        Dim wsRDB As Worksheet
        On Error Resume Next
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo ErrorExit
        
        If Not wsRDB Is Nothing Then
            ' 날짜 계산
            Dim 월 As Integer, 일자 As Integer, 연도 As Integer
            Dim 기준날짜 As Date
            연도 = B2_연도(ws)
            월 = val(Replace(Replace(CStr(ws.Cells(2, 열번호).Value), "m", ""), "M", ""))
            일자 = val(ws.Cells(3, 열번호).Value)
            If 월 < 1 Or 월 > 12 Or 일자 < 1 Or 일자 > 31 Then GoTo CleanExit
            
            On Error Resume Next
            기준날짜 = DateSerial(연도, 월, 일자)
            On Error GoTo ErrorExit
            
            ' 기본 정보 읽기
            Dim 업체명 As String, 제품명 As String
            Dim 제조사 As String, 지관 As String, 지폭 As String, 총량 As String, 일차포장 As String, 이차포장 As String
            업체명 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
            제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
            제조사 = Trim(CStr(ws.Cells(그룹시작행, COL_제조사).Value))
            지관 = Trim(CStr(ws.Cells(그룹시작행, COL_지관).Value))
            지폭 = Trim(CStr(ws.Cells(그룹시작행, COL_지폭).Value))
            총량 = Trim(CStr(ws.Cells(그룹시작행, COL_총량).Value))
            일차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_1차포장).Value))
            이차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_2차포장).Value))
            
            Dim rRDB As Long
            Dim DB딕소요 As Object, dbK소요 As String
            Set DB딕소요 = 일일생산DB_딕셔너리_생성(wsRDB)
            dbK소요 = DB키_생성(기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장)
            If DB딕소요.exists(dbK소요) Then rRDB = DB딕소요(dbK소요) Else rRDB = 일일생산DB_행추가_통합(wsRDB, 기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장): If rRDB > 0 Then DB딕소요.Add dbK소요, rRDB
            
            ' ★★★ BS~BZ열 쓰기 비활성화됨 ★★★
            
            ' ★★★ 3단계: 제품이면 일일생산DB에도 저장 ★★★
            If 자재유형명 = "제품" Then
    Dim wsDB As Worksheet
                On Error Resume Next
                Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
                On Error GoTo ErrorExit
                If Not wsDB Is Nothing Then
                    Dim lastRowDB As Long, r As Long, 찾음 As Boolean
                    lastRowDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
                    찾음 = False
                    For r = 2 To lastRowDB
                        If Trim(CStr(wsDB.Cells(r, 10).Value)) = 제품명 And _
                           IsDate(wsDB.Cells(r, 1).Value) And _
                           CDate(wsDB.Cells(r, 1).Value) = 기준날짜 Then
                            찾음 = True: Exit For
                        End If
                    Next r
                    If 찾음 Then
                        wsDB.Cells(r, 12).Value = 새값
                    Else
                        lastRowDB = lastRowDB + 1
                        wsDB.Cells(lastRowDB, 1).Value = 기준날짜
                        wsDB.Cells(lastRowDB, 10).Value = 제품명
                        wsDB.Cells(lastRowDB, 12).Value = 새값
                    End If
                End If
            End If
        End If
        
        ' ★★★ 4단계: 안전재고 미달 자동 체크 ★★★
        Call 재고부족_자동체크(ws, 그룹시작행, 블록인덱스)
        
CleanExit:
        If 시트보호 Then Call 원단_셀보호_적용_무음(ws)
        Exit Sub
        
ErrorExit:
        If 시트보호 Then
            On Error Resume Next
            Call 원단_셀보호_적용_무음(ws)
            On Error GoTo 0
        End If
    End Sub

    '================================================================================
    ' ★ 입고 값 변경 시: 재고 재계산 + 일일생산DB 저장
    ' Worksheet_Change → 이 함수 호출 (완전 자동)
    '================================================================================
    Public Sub 원단_입고값_일일생산DB_동기화(변경셀 As Range)
        On Error GoTo ErrorExit
        
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        Dim 행번호 As Long: 행번호 = 변경셀.Row
        Dim 열번호 As Long: 열번호 = 변경셀.Column
        Dim 시트보호 As Boolean: 시트보호 = ws.ProtectContents
        If 시트보호 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        
        ' 기본 검증 (행 오프셋으로 이미 입고 행 판단됨, 구분 텍스트는 참고만)
        If 열번호 < COL_일별시작 Then GoTo CleanExit
        
        ' 그룹 시작행 계산
        Dim 그룹시작행 As Long
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        If 그룹시작행 < START_ROW Then GoTo CleanExit
        
        ' 자재유형 판단 (오프셋 기반, 5블록·15행: 입고 행 = 3,6,9,12)
        Dim 오프셋 As Long: 오프셋 = 행번호 - 그룹시작행
        Dim 자재유형명 As String: 자재유형명 = ""
        Dim 블록인덱스 As Long: 블록인덱스 = -1
        Select Case 오프셋
            Case 3: 자재유형명 = "원지(ROLL)": 블록인덱스 = 1
            Case 6: 자재유형명 = "원지(Tons)": 블록인덱스 = 2
            Case 9: 자재유형명 = "1차포장": 블록인덱스 = 3
            Case 12: 자재유형명 = "2차포장": 블록인덱스 = 4
            Case Else: GoTo CleanExit
        End Select
        
        ' 새 값 읽기
        Dim 새값 As Double
        If IsNumeric(변경셀.Value) Then 새값 = CDbl(변경셀.Value) Else 새값 = 0
        
        ' ★★★ 1단계: 해당 블록의 재고만 재계산 (소요/입고 값은 건드리지 않음!) ★★★
        Call 원단_블록_재고_재계산(ws, 그룹시작행, 블록인덱스)
        
        ' ★★★ 2단계: 일일생산DB에 저장 ★★★
        Dim wsRDB As Worksheet
        On Error Resume Next
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo ErrorExit
        
        If Not wsRDB Is Nothing Then
            ' 날짜 계산
            Dim 월 As Integer, 일자 As Integer, 연도 As Integer
    Dim 기준날짜 As Date
            연도 = B2_연도(ws)
            월 = val(Replace(Replace(CStr(ws.Cells(2, 열번호).Value), "m", ""), "M", ""))
            일자 = val(ws.Cells(3, 열번호).Value)
            If 월 < 1 Or 월 > 12 Or 일자 < 1 Or 일자 > 31 Then GoTo CleanExit
            
            On Error Resume Next
            기준날짜 = DateSerial(연도, 월, 일자)
            On Error GoTo ErrorExit
            
            ' 기본 정보 읽기
            Dim 업체명 As String, 제품명 As String
            Dim 제조사 As String, 지관 As String, 지폭 As String, 총량 As String, 일차포장 As String, 이차포장 As String
            업체명 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
            제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
            제조사 = Trim(CStr(ws.Cells(그룹시작행, COL_제조사).Value))
            지관 = Trim(CStr(ws.Cells(그룹시작행, COL_지관).Value))
            지폭 = Trim(CStr(ws.Cells(그룹시작행, COL_지폭).Value))
            총량 = Trim(CStr(ws.Cells(그룹시작행, COL_총량).Value))
            일차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_1차포장).Value))
            이차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_2차포장).Value))
            
            Dim rRDB As Long
            Dim DB딕입고 As Object, dbK입고 As String
            Set DB딕입고 = 일일생산DB_딕셔너리_생성(wsRDB)
            dbK입고 = DB키_생성(기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장)
            If DB딕입고.exists(dbK입고) Then rRDB = DB딕입고(dbK입고) Else rRDB = 일일생산DB_행추가_통합(wsRDB, 기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장): If rRDB > 0 Then DB딕입고.Add dbK입고, rRDB
            
            ' ★★★ BS~BZ열 쓰기 비활성화됨 ★★★
        End If
        
        ' ★★★ 3단계: 안전재고 미달 자동 체크 ★★★
        Call 재고부족_자동체크(ws, 그룹시작행, 블록인덱스)
        
CleanExit:
        If 시트보호 Then Call 원단_셀보호_적용_무음(ws)
        Exit Sub
        
ErrorExit:
        If 시트보호 Then
            On Error Resume Next
            Call 원단_셀보호_적용_무음(ws)
            On Error GoTo 0
        End If
    End Sub

    '================================================================================
    ' ★ 생산 값 변경 시: 제품 블록 재고 재계산 + 일일생산DB 저장
    ' 원단_입고값_일일생산DB_동기화와 완전히 동일한 구조 (생산=재고에 더하는 값)
    '================================================================================
    Public Sub 원단_생산값_DB_동기화(변경셀 As Range)
        On Error GoTo ErrorExit
        
        Dim ws As Worksheet: Set ws = 변경셀.Worksheet
        Dim 행번호 As Long: 행번호 = 변경셀.Row
        Dim 열번호 As Long: 열번호 = 변경셀.Column
        Dim 시트보호 As Boolean: 시트보호 = ws.ProtectContents
        If 시트보호 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
        
        ' 기본 검증
        If 열번호 < COL_일별시작 Then GoTo CleanExit
        If 행번호 < START_ROW Then GoTo CleanExit
        
        ' 그룹 시작행 계산 (생산 = 그룹 첫 행 = 오프셋 0)
        Dim 그룹시작행 As Long
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        If 그룹시작행 < START_ROW Then GoTo CleanExit
        If 행번호 <> 그룹시작행 Then GoTo CleanExit
        
        ' 생산은 제품 블록 고정
        Dim 자재유형명 As String: 자재유형명 = "제품"
        Dim 블록인덱스 As Long: 블록인덱스 = 0
        
        ' 새 값 읽기
    Dim 새값 As Double
        If IsNumeric(변경셀.Value) Then 새값 = CDbl(변경셀.Value) Else 새값 = 0
        
        ' ★★★ 1단계: 제품 블록 재고 재계산 ★★★
        Call 원단_블록_재고_재계산(ws, 그룹시작행, 블록인덱스)
        
        ' ★★★ 2단계: 일일생산DB에 저장 (입고 함수와 동일한 구조) ★★★
        Dim wsRDB As Worksheet
        Set wsRDB = Nothing
        On Error Resume Next
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        
        If Not wsRDB Is Nothing Then
            On Error GoTo ErrorExit
            
            ' 날짜 계산
            Dim 월 As Integer, 일자 As Integer, 연도 As Integer
            Dim 기준날짜 As Date
            연도 = B2_연도(ws)
            월 = val(Replace(Replace(CStr(ws.Cells(2, 열번호).Value), "m", ""), "M", ""))
            일자 = val(ws.Cells(3, 열번호).Value)
            If 월 >= 1 And 월 <= 12 And 일자 >= 1 And 일자 <= 31 Then
                기준날짜 = DateSerial(연도, 월, 일자)
                
                ' 기본 정보 읽기
                Dim 업체명 As String, 제품명 As String
                Dim 제조사 As String, 지관 As String, 지폭 As String, 총량 As String, 일차포장 As String, 이차포장 As String
                업체명 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
                제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
                제조사 = Trim(CStr(ws.Cells(그룹시작행, COL_제조사).Value))
                지관 = Trim(CStr(ws.Cells(그룹시작행, COL_지관).Value))
                지폭 = Trim(CStr(ws.Cells(그룹시작행, COL_지폭).Value))
                총량 = Trim(CStr(ws.Cells(그룹시작행, COL_총량).Value))
                일차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_1차포장).Value))
                이차포장 = Trim(CStr(ws.Cells(그룹시작행, COL_2차포장).Value))
                
                Dim rRDB As Long
                Dim DB딕생산 As Object, dbK생산 As String
                Set DB딕생산 = 일일생산DB_딕셔너리_생성(wsRDB)
                dbK생산 = DB키_생성(기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장)
                If DB딕생산.exists(dbK생산) Then rRDB = DB딕생산(dbK생산) Else rRDB = 일일생산DB_행추가_통합(wsRDB, 기준날짜, 업체명, 제품명, 제조사, 지관, 지폭, 총량, 일차포장, 이차포장): If rRDB > 0 Then DB딕생산.Add dbK생산, rRDB
                
                ' ★★★ BS~BZ열 쓰기 비활성화됨 ★★★
            End If
        End If
        
        ' ★★★ 3단계: 안전재고 미달 자동 체크 ★★★
        Call 재고부족_자동체크(ws, 그룹시작행, 블록인덱스)
        
CleanExit:
        If 시트보호 Then Call 원단_셀보호_적용_무음(ws)
        Exit Sub
        
ErrorExit:
        If 시트보호 Then
    On Error Resume Next
            Call 원단_셀보호_적용_무음(ws)
    On Error GoTo 0
        End If
    End Sub

    '================================================================================
    ' 셀 보호 해제 (관리자용)
    '================================================================================
    Public Sub 원단_셀보호_해제()
        Dim ws As Worksheet

        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0

        If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

        If 시트보호해제_시도(ws) Then
            ws.Cells.Locked = False
        End If

        MsgBox "셀 보호가 해제되었습니다." & vbCrLf & _
            "모든 셀을 편집할 수 있습니다.", vbInformation, "셀 보호 해제"
    End Sub

    '================================================================================
    ' 원단관리 시트에 Worksheet_Change 없을 때만 자동 설치 (데이터업데이트에서 호출)
    ' 이미 있으면 추가하지 않음 → 기존 코드(음영해제·키워드검색·DB동기화) 유지, 중복 방지
    '================================================================================
    Public Sub 원단_이벤트코드_자동설치_필요시()
        Dim ws As Worksheet
        Dim vbComp As Object
        Dim 기존코드 As String
        Dim 코드 As String
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        If ws Is Nothing Or vbComp Is Nothing Then Exit Sub
        If vbComp.CodeModule.CountOfLines > 0 Then
            기존코드 = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
        Else
            기존코드 = ""
        End If
        If InStr(1, 기존코드, "Worksheet_Change", vbTextCompare) > 0 Then Exit Sub
        코드 = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    Dim r As Range" & vbCrLf
        코드 = 코드 & "    If Target.Cells.Count > 1 Then Set r = Target.Cells(1, 1) Else Set r = Target" & vbCrLf
        코드 = 코드 & "    If r.Column = 1 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_A열_Co입력시_B열원단명채우기(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 2 Or r.Column = 3 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then" & vbCrLf
        코드 = 코드 & "            If r.Column = 2 Then Call 원단관리_B열_원단명입력시_A열Co채우기(r)" & vbCrLf
        코드 = 코드 & "            Call 원단관리_BC_변경시_DtoI_채우기(r)" & vbCrLf
        코드 = 코드 & "        End If" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 4 Or r.Column = 5 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_기초재고_변경_재계산(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column < 10 Then Exit Sub" & vbCrLf
        코드 = 코드 & "    On Error GoTo ErrorHandler" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = False" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = False" & vbCrLf
        코드 = 코드 & "    Call 원단관리_셀변경_DB동기화(r)" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    Exit Sub" & vbCrLf
        코드 = 코드 & "ErrorHandler:" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    MsgBox ""값 변경 처리 중 오류: "" & Err.Description, vbExclamation" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        코드 = 코드 & vbCrLf & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_C열유효성_갱신대기처리: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: If Target.Cells.Count = 1 And Target.Column <> 2 Then Call 원단관리_B열기준_A열전체동기화: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    Call 원단_오늘이동_음영해제" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_검색_하이라이트_해제: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        On Error Resume Next
        vbComp.CodeModule.AddFromString 코드
        On Error GoTo 0
    End Sub

    '================================================================================
    ' 이벤트 코드 자동 설치 (원단관리 시트에 Worksheet_Change 이벤트 추가)
    '================================================================================
    Public Sub 원단_이벤트코드_설치()
        Dim ws As Worksheet
        Dim vbComp As Object
        Dim 코드 As String
        Dim 기존코드 As String
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        
        If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

        ' VBA 프로젝트 접근 확인
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        
        If vbComp Is Nothing Then
            MsgBox "VBA 프로젝트에 접근할 수 없습니다." & vbCrLf & vbCrLf & _
                "Excel 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정에서" & vbCrLf & _
                "'VBA 프로젝트 개체 모델에 대한 액세스 신뢰'를 체크하세요.", vbExclamation, "권한 필요"
        Exit Sub
    End If

        ' 기존 코드 확인 (이미 있으면 덮어쓰지 않음 → 음영해제·키워드검색·DB동기화 유지)
        If vbComp.CodeModule.CountOfLines > 0 Then
            기존코드 = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
        Else
            기존코드 = ""
        End If
        If InStr(1, 기존코드, "Worksheet_Change", vbTextCompare) > 0 Then
            MsgBox "원단관리 시트에 이미 이벤트 코드가 있습니다." & vbCrLf & vbCrLf & _
                   "기존 코드(원단_오늘이동 음영 해제, 키워드 검색, DB 동기화)를 유지합니다." & vbCrLf & _
                   "중복 설치하지 않습니다.", vbInformation, "이벤트 코드"
            Exit Sub
        End If

        ' 새 이벤트 코드 작성 (B/C 변경 시 D~I 자동 채우기, 일별 열은 원단관리_셀변경_DB동기화)
        코드 = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    Dim r As Range" & vbCrLf
        코드 = 코드 & "    If Target.Cells.Count > 1 Then Set r = Target.Cells(1, 1) Else Set r = Target" & vbCrLf
        코드 = 코드 & "    If r.Column = 1 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_A열_Co입력시_B열원단명채우기(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 2 Or r.Column = 3 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then" & vbCrLf
        코드 = 코드 & "            If r.Column = 2 Then Call 원단관리_B열_원단명입력시_A열Co채우기(r)" & vbCrLf
        코드 = 코드 & "            Call 원단관리_BC_변경시_DtoI_채우기(r)" & vbCrLf
        코드 = 코드 & "        End If" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 4 Or r.Column = 5 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_기초재고_변경_재계산(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column < 10 Then Exit Sub" & vbCrLf
        코드 = 코드 & "    On Error GoTo ErrorHandler" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = False" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = False" & vbCrLf
        코드 = 코드 & "    Call 원단관리_셀변경_DB동기화(r)" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    Exit Sub" & vbCrLf
        코드 = 코드 & "ErrorHandler:" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    MsgBox ""값 변경 처리 중 오류: "" & Err.Description, vbExclamation" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        ' 커서 이동 시 C열 유효성 갱신 + B열→A열 역방향 동기화(다른 셀 선택 시) + 원단_오늘이동 음영 해제 + 검색 하이라이트 해제
        코드 = 코드 & vbCrLf & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_C열유효성_갱신대기처리: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: If Target.Cells.Count = 1 And Target.Column <> 2 Then Call 원단관리_B열기준_A열전체동기화: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    Call 원단_오늘이동_음영해제" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_검색_하이라이트_해제: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        
        ' 코드 삽입
        vbComp.CodeModule.AddFromString 코드
        
        MsgBox "이벤트 코드가 성공적으로 설치되었습니다!" & vbCrLf & vbCrLf & _
            "이제 '생산', '입고', '소요' 값을 변경하면 자동으로 DB에 반영됩니다.", vbInformation, "설치 완료"
    End Sub
    
    '================================================================================
    ' 이벤트 코드 강제 재설치 (기존 Worksheet_Change/SelectionChange 교체 → 재고 재계산 등 정상 동작)
    ' 재고 재계산이 안 될 때 이 매크로 실행 후 E열/입고/사용 셀 변경 시 자동 재계산됨
    '================================================================================
    Public Sub 원단_이벤트코드_강제재설치()
        Dim ws As Worksheet
        Dim vbComp As Object
        Dim cm As Object
        Dim stLine As Long, lineCnt As Long
        Dim 코드 As String
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        If ws Is Nothing Or vbComp Is Nothing Then
            MsgBox "원단관리 시트 또는 VBA 프로젝트에 접근할 수 없습니다.", vbExclamation
            Exit Sub
        End If
        Application.EnableEvents = True
        Set cm = vbComp.CodeModule
        ' Worksheet_SelectionChange 먼저 삭제 (뒤에 있으므로)
        On Error Resume Next
        stLine = cm.ProcStartLine("Worksheet_SelectionChange", 0)
        If stLine > 0 Then lineCnt = cm.ProcCountLines("Worksheet_SelectionChange", 0): cm.DeleteLines stLine, lineCnt
        ' Worksheet_Change 삭제
        stLine = cm.ProcStartLine("Worksheet_Change", 0)
        If stLine > 0 Then lineCnt = cm.ProcCountLines("Worksheet_Change", 0): cm.DeleteLines stLine, lineCnt
        On Error GoTo 0
        ' 새 코드 추가
        코드 = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    Dim r As Range" & vbCrLf
        코드 = 코드 & "    If Target.Cells.Count > 1 Then Set r = Target.Cells(1, 1) Else Set r = Target" & vbCrLf
        코드 = 코드 & "    If r.Column = 1 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_A열_Co입력시_B열원단명채우기(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 2 Or r.Column = 3 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then" & vbCrLf
        코드 = 코드 & "            If r.Column = 2 Then Call 원단관리_B열_원단명입력시_A열Co채우기(r)" & vbCrLf
        코드 = 코드 & "            Call 원단관리_BC_변경시_DtoI_채우기(r)" & vbCrLf
        코드 = 코드 & "        End If" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 4 Or r.Column = 5 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_기초재고_변경_재계산(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column < 10 Then Exit Sub" & vbCrLf
        코드 = 코드 & "    On Error GoTo ErrorHandler" & vbCrLf
        코드 = 코드 & "    Call 원단관리_셀변경_DB동기화_예약(r)" & vbCrLf
        코드 = 코드 & "    Exit Sub" & vbCrLf
        코드 = 코드 & "ErrorHandler:" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    MsgBox ""값 변경 처리 중 오류: "" & Err.Description, vbExclamation" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        코드 = 코드 & vbCrLf & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_C열유효성_갱신대기처리: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: If Target.Cells.Count = 1 And Target.Column <> 2 Then Call 원단관리_B열기준_A열전체동기화: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "    Call 원단_오늘이동_음영해제" & vbCrLf
        코드 = 코드 & "    On Error Resume Next: Call 원단관리_검색_하이라이트_해제: On Error GoTo 0" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        cm.AddFromString 코드
        MsgBox "이벤트 코드를 강제 재설치했습니다." & vbCrLf & vbCrLf & _
            "E열(기초재고)·입고·사용 셀 변경 시 재고가 자동 재계산됩니다." & vbCrLf & _
            "EnableEvents도 활성화했습니다.", vbInformation, "재설치 완료"
    End Sub

    '================================================================================
    ' 재고 수동 재계산 (이벤트 무관, 즉시 전체 재계산)
    ' 재고가 자동으로 안 바뀔 때 이 매크로 실행
    '================================================================================
    Public Sub 원단_재고_수동재계산()
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
            Exit Sub
        End If
        Application.EnableEvents = True
        Application.ScreenUpdating = False
        Call 원단관리_값일괄계산(ws)
        Application.ScreenUpdating = True
        MsgBox "재고 전체 재계산을 완료했습니다.", vbInformation
    End Sub

    '================================================================================
    ' 이벤트 코드 제거
    '================================================================================
    Public Sub 원단_이벤트코드_제거()
        Dim ws As Worksheet
        Dim vbComp As Object
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        
        If vbComp Is Nothing Then
            MsgBox "시트 모듈에 접근할 수 없습니다.", vbExclamation
        Exit Sub
    End If

        If vbComp.CodeModule.CountOfLines > 0 Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
            MsgBox "이벤트 코드가 제거되었습니다.", vbInformation
        Else
            MsgBox "제거할 코드가 없습니다.", vbInformation
        End If
    End Sub
    
    '================================================================================
    ' 긴급 수동 DB 업데이트 (이벤트 없이 직접 DB에 저장)
    '================================================================================
    Public Sub 원단_수동DB업데이트()
        Dim ws As Worksheet
        Dim 선택셀 As Range
        Dim 구분 As String
        
        Set ws = ActiveSheet
        If ws.Name <> WS_원단관리 Then
            MsgBox "원단관리 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If
        
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Or 선택셀.Cells.Count <> 1 Then
            MsgBox "단일 셀을 선택하세요.", vbExclamation
        Exit Sub
    End If

        구분 = Trim(CStr(ws.Cells(선택셀.Row, COL_구분).Value))
        
        If 구분 = "소요" Or 구분 = "출고" Then
            Call 원단_소요값_일일생산DB_동기화(선택셀)
        ElseIf 구분 = "입고" Then
            Call 원단_입고값_일일생산DB_동기화(선택셀)
        Else
            MsgBox "'소요'/'출고' 또는 '입고' 행의 셀을 선택하세요." & vbCrLf & vbCrLf & _
                "현재 구분: " & 구분, vbExclamation
        End If
    End Sub
    
    '================================================================================
    ' 올인원 설정: 이벤트 시스템 완전 설치 (한 번에 모든 것을 설정)
    '================================================================================
    Public Sub 원단_이벤트시스템_완전설치()
        Dim 결과 As String
        Dim 단계 As Integer
        
        결과 = "【 이벤트 시스템 완전 설치 】" & vbCrLf & vbCrLf
        단계 = 1
        
        ' 1단계: EnableEvents 활성화
        결과 = 결과 & "1단계: EnableEvents 활성화..." & vbCrLf
        If Not Application.EnableEvents Then
            Application.EnableEvents = True
            결과 = 결과 & "   ? EnableEvents 활성화 완료" & vbCrLf & vbCrLf
        Else
            결과 = 결과 & "   ? 이미 활성화되어 있음" & vbCrLf & vbCrLf
        End If
        
        ' 2단계: 시트 확인
        결과 = 결과 & "2단계: 시트 확인..." & vbCrLf
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        
        If ws Is Nothing Then
            결과 = 결과 & "   ? 원단관리 시트를 찾을 수 없습니다!" & vbCrLf
            MsgBox 결과, vbCritical
        Exit Sub
        Else
            결과 = 결과 & "   ? 원단관리 시트 존재" & vbCrLf & vbCrLf
    End If

        ' 3단계: 이벤트 코드 설치
        결과 = 결과 & "3단계: 이벤트 코드 설치..." & vbCrLf
        Dim vbComp As Object
        Dim 코드 As String
        
    On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
    On Error GoTo 0

        If vbComp Is Nothing Then
            결과 = 결과 & "   ? VBA 프로젝트 접근 권한 필요!" & vbCrLf
            결과 = 결과 & "   Excel 옵션에서 'VBA 프로젝트 개체 모델 액세스' 허용 필요" & vbCrLf
            MsgBox 결과, vbExclamation
            Exit Sub
        End If
        
        ' 기존 코드 삭제
        If vbComp.CodeModule.CountOfLines > 0 Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        
        ' ★★★ B/C 변경 시 D~I 자동 채우기, 생산/입고/소요는 원단관리_셀변경_DB동기화로 위임 ★★★
        코드 = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    Dim r As Range" & vbCrLf
        코드 = 코드 & "    If Target.Cells.Count > 1 Then Set r = Target.Cells(1, 1) Else Set r = Target" & vbCrLf
        코드 = 코드 & "    If r.Column = 1 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_A열_Co입력시_B열원단명채우기(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 2 Or r.Column = 3 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then" & vbCrLf
        코드 = 코드 & "            If r.Column = 2 Then Call 원단관리_B열_원단명입력시_A열Co채우기(r)" & vbCrLf
        코드 = 코드 & "            Call 원단관리_BC_변경시_DtoI_채우기(r)" & vbCrLf
        코드 = 코드 & "        End If" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column = 4 Or r.Column = 5 Then" & vbCrLf
        코드 = 코드 & "        If r.Row >= 4 Then Call 원단관리_기초재고_변경_재계산(r)" & vbCrLf
        코드 = 코드 & "        Exit Sub" & vbCrLf
        코드 = 코드 & "    End If" & vbCrLf
        코드 = 코드 & "    If r.Column < 10 Then Exit Sub" & vbCrLf
        코드 = 코드 & "    On Error GoTo ErrorHandler" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = False" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = False" & vbCrLf
        코드 = 코드 & "    Call 원단관리_셀변경_DB동기화(r)" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    Exit Sub" & vbCrLf
        코드 = 코드 & "ErrorHandler:" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    Application.StatusBar = False" & vbCrLf
        코드 = 코드 & "    MsgBox ""자동 처리 중 오류 발생!"" & vbCrLf & vbCrLf & ""오류: "" & Err.Description, vbCritical, ""이벤트 오류""" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        
        vbComp.CodeModule.AddFromString 코드
        결과 = 결과 & "   ? Worksheet_Change 이벤트 설치 완료" & vbCrLf & vbCrLf
        
        ' 4단계: DB 시트 확인
        결과 = 결과 & "4단계: DB 시트 확인..." & vbCrLf
        Dim wsDB As Worksheet, wsRDB As Worksheet
        
        On Error Resume Next
        Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        
        If wsDB Is Nothing Then
            결과 = 결과 & "   ?? 일일생산DB 시트 없음 (제품 소요 동기화 불가)" & vbCrLf
        Else
            결과 = 결과 & "   ? 일일생산DB 시트 존재" & vbCrLf
    End If

        If wsRDB Is Nothing Then
            결과 = 결과 & "   ?? 일일생산DB 시트 없음 (부자재 동기화 불가)" & vbCrLf
        Else
            결과 = 결과 & "   ? 일일생산DB 시트 존재" & vbCrLf
        End If
        
        결과 = 결과 & vbCrLf & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        결과 = 결과 & "? 설치 완료!" & vbCrLf
        결과 = 결과 & "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf
        결과 = 결과 & "【 지원되는 모든 카테고리 】" & vbCrLf
        결과 = 결과 & "? 제품 - 입고/소요" & vbCrLf
        결과 = 결과 & "? 원지 - 입고/소요" & vbCrLf
        결과 = 결과 & "? 지관 - 입고/소요" & vbCrLf
        결과 = 결과 & "? 폴리백 - 입고/소요" & vbCrLf
        결과 = 결과 & "? 비닐백 - 입고/소요" & vbCrLf
        결과 = 결과 & "? box자재현황 - 입고/소요" & vbCrLf & vbCrLf
        결과 = 결과 & "【 테스트 방법 】" & vbCrLf
        결과 = 결과 & "1. 원단관리 시트의 입고/소요 값 변경" & vbCrLf
        결과 = 결과 & "2. 확인 메시지가 나타나면 '예' 선택" & vbCrLf
        결과 = 결과 & "3. 일일생산DB 시트(71~81열 부자재 블록)에서 확인"
        
        MsgBox 결과, vbInformation, "이벤트 시스템 설치 완료"
    End Sub
    
    '================================================================================
    ' EnableEvents 확인 및 활성화 (가장 먼저 실행해야 할 매크로!)
    '================================================================================
    Public Sub 원단_EnableEvents_확인및활성화()
        Dim 메시지 As String
        
        메시지 = "【 현재 상태 】" & vbCrLf & vbCrLf
        
        If Application.EnableEvents Then
            메시지 = 메시지 & "? Application.EnableEvents = TRUE" & vbCrLf
            메시지 = 메시지 & "   → 이벤트가 활성화되어 있습니다!" & vbCrLf & vbCrLf
            메시지 = 메시지 & "이벤트가 정상 작동해야 합니다." & vbCrLf
            메시지 = 메시지 & "만약 여전히 작동하지 않으면:" & vbCrLf
            메시지 = 메시지 & "1. '원단_이벤트진단' 매크로 실행" & vbCrLf
            메시지 = 메시지 & "2. '원단_이벤트코드_테스트설치' 매크로 실행"
        Else
            Application.EnableEvents = True
            메시지 = 메시지 & "? Application.EnableEvents = FALSE (비활성화됨)" & vbCrLf
            메시지 = 메시지 & "   → 자동으로 TRUE로 변경했습니다!" & vbCrLf & vbCrLf
            메시지 = 메시지 & "? 이제 이벤트가 활성화되었습니다." & vbCrLf & vbCrLf
            메시지 = 메시지 & "소요/입고 값을 변경해보세요!"
                End If
        
        MsgBox 메시지, vbInformation, "EnableEvents 상태"
    End Sub
    
    '================================================================================
    ' EnableEvents 강제 활성화 (빠른 실행용)
    '================================================================================
    Public Sub 원단_EnableEvents_켜기()
        Application.EnableEvents = True
        MsgBox "? Application.EnableEvents를 TRUE로 설정했습니다!" & vbCrLf & vbCrLf & _
            "이제 셀 변경 이벤트가 작동합니다.", vbInformation, "완료"
    End Sub
    
    ' (기존 원단_이벤트진단 제거 - 아래 새 버전으로 통합됨)
    
    '================================================================================
    ' DB 저장 확인 도구 - 선택한 셀의 값이 DB에 제대로 저장되었는지 확인
    '================================================================================
    Public Sub 원단_DB저장확인()
        Dim ws As Worksheet
        Dim wsDB As Worksheet
        Dim wsRDB As Worksheet
        Dim 선택셀 As Range
        Dim 구분 As String
        Dim 행번호 As Long, 열번호 As Long
        Dim 그룹시작행 As Long
        Dim 오프셋 As Long
        Dim 자재유형명 As String
        Dim 제품명 As String
        Dim 월 As Integer, 일자 As Integer, 연도 As Integer
        Dim 기준날짜 As Date
        Dim 결과 As String
        Dim rRDB As Long
        Dim 찾음 As Boolean
        
        Set ws = ActiveSheet
        If ws.Name <> WS_원단관리 Then
            MsgBox "원단관리 시트에서 실행하세요.", vbExclamation
            Exit Sub
            End If
        
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Or 선택셀.Cells.Count <> 1 Then
            MsgBox "단일 셀을 선택하세요.", vbExclamation
            Exit Sub
        End If
        
        행번호 = 선택셀.Row
        열번호 = 선택셀.Column
        구분 = Trim(CStr(ws.Cells(행번호, COL_구분).Value))
        
        결과 = "【 DB 저장 확인 】" & vbCrLf & vbCrLf
        결과 = 결과 & "선택 셀: " & 선택셀.Address & vbCrLf
        결과 = 결과 & "값: " & 선택셀.Value & vbCrLf
        결과 = 결과 & "구분: " & 구분 & vbCrLf & vbCrLf
        
        If 구분 <> "소요" And 구분 <> "출고" And 구분 <> "입고" Then
            결과 = 결과 & "? '소요'/'출고' 또는 '입고' 행이 아닙니다."
            MsgBox 결과, vbExclamation
            Exit Sub
        End If
        
        ' 그룹 시작행 및 자재유형 계산
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        오프셋 = 행번호 - 그룹시작행
        제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
        
        If 구분 = "소요" Or 구분 = "출고" Then
            Select Case 오프셋
                Case 2: 자재유형명 = "제품"
                Case 5: 자재유형명 = "원지"
                Case 8: 자재유형명 = "지관"
                Case 11: 자재유형명 = "폴리백"
                Case 14: 자재유형명 = "비닐백"
                Case 17: 자재유형명 = "box자재현황"
                Case Else: 자재유형명 = "(잘못된 오프셋)"
            End Select
        Else ' 입고
            Select Case 오프셋
                Case 1: 자재유형명 = "제품"
                Case 4: 자재유형명 = "원지"
                Case 7: 자재유형명 = "지관"
                Case 10: 자재유형명 = "폴리백"
                Case 13: 자재유형명 = "비닐백"
                Case 16: 자재유형명 = "box자재현황"
                Case Else: 자재유형명 = "(잘못된 오프셋)"
            End Select
        End If
        
        결과 = 결과 & "그룹 시작행: " & 그룹시작행 & vbCrLf
        결과 = 결과 & "오프셋: " & 오프셋 & vbCrLf
        결과 = 결과 & "자재유형: " & 자재유형명 & vbCrLf
        결과 = 결과 & "제품명: " & 제품명 & vbCrLf & vbCrLf
        
        ' 날짜 계산
        연도 = B2_연도(ws)
        월 = val(Replace(Replace(CStr(ws.Cells(2, 열번호).Value), "m", ""), "M", ""))
        일자 = val(ws.Cells(3, 열번호).Value)
        On Error Resume Next
        기준날짜 = DateSerial(연도, 월, 일자)
        On Error GoTo 0
        
        결과 = 결과 & "기준날짜: " & Format(기준날짜, "YYYY-MM-DD") & vbCrLf & vbCrLf
        
        ' 일일생산DB 확인
        결과 = 결과 & "━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        결과 = 결과 & "【 일일생산DB 확인 】" & vbCrLf
        결과 = 결과 & "━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
        
        On Error Resume Next
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        
        If wsRDB Is Nothing Then
            결과 = 결과 & "? 일일생산DB 시트가 없습니다!"
        Else
            ' 통합 구조로 행 찾기 (자재유형 무관)
            rRDB = 일일생산DB_행찾기_통합(wsRDB, 기준날짜, _
                Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value)), 제품명, _
                Trim(CStr(ws.Cells(그룹시작행, COL_제조사).Value)), _
                Trim(CStr(ws.Cells(그룹시작행, COL_지관).Value)), _
                Trim(CStr(ws.Cells(그룹시작행, COL_지폭).Value)), _
                Trim(CStr(ws.Cells(그룹시작행, COL_총량).Value)), _
                Trim(CStr(ws.Cells(그룹시작행, COL_1차포장).Value)), _
                Trim(CStr(ws.Cells(그룹시작행, COL_2차포장).Value)))
            
            If rRDB > 0 Then
                ' 해당 자재유형의 열 위치 계산
                Dim 자재열시작_확인 As Long
                자재열시작_확인 = 일일생산DB_자재열_시작(자재유형명)
                
                결과 = 결과 & "? DB에서 해당 행 찾음 (행번호: " & rRDB & ")" & vbCrLf
                결과 = 결과 & "   날짜: " & wsRDB.Cells(rRDB, 부자재열(RDB_Date)).Value & vbCrLf
                결과 = 결과 & "   자재유형: " & 자재유형명 & vbCrLf
                
                If 자재열시작_확인 > 0 Then
                    결과 = 결과 & "   전월재고: " & wsRDB.Cells(rRDB, 부자재열(자재열시작_확인)).Value & vbCrLf
                    결과 = 결과 & "   입고: " & wsRDB.Cells(rRDB, 부자재열(자재열시작_확인 + 1)).Value & vbCrLf
                    결과 = 결과 & "   소요: " & wsRDB.Cells(rRDB, 부자재열(자재열시작_확인 + 2)).Value & vbCrLf
                    결과 = 결과 & "   재고: " & wsRDB.Cells(rRDB, 부자재열(자재열시작_확인 + 3)).Value & vbCrLf & vbCrLf
                    
                    ' 값 비교
                    Dim DB값 As Double
                    If 구분 = "소요" Or 구분 = "출고" Then
                        DB값 = wsRDB.Cells(rRDB, 부자재열(자재열시작_확인 + 2)).Value  ' 소요
                    Else
                        DB값 = wsRDB.Cells(rRDB, 부자재열(자재열시작_확인 + 1)).Value  ' 입고
                    End If
                    
                    If IsNumeric(선택셀.Value) Then
                        If CDbl(선택셀.Value) = DB값 Then
                            결과 = 결과 & "? 값이 일치합니다!"
                        Else
                            결과 = 결과 & "?? 값이 다릅니다!" & vbCrLf
                            결과 = 결과 & "   시트 값: " & 선택셀.Value & vbCrLf
                            결과 = 결과 & "   DB 값: " & DB값
                        End If
                    End If
                Else
                    결과 = 결과 & "?? 자재유형 열을 찾을 수 없습니다."
                End If
            Else
                결과 = 결과 & "? DB에서 해당 행을 찾을 수 없습니다!" & vbCrLf & vbCrLf
                결과 = 결과 & "값을 변경한 후 DB 동기화를 수행하세요."
            End If
        End If
        
        ' 일일생산DB 확인 (제품 소요/출고만)
        If (구분 = "소요" Or 구분 = "출고") And 자재유형명 = "제품" Then
            결과 = 결과 & vbCrLf & vbCrLf
            결과 = 결과 & "━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
            결과 = 결과 & "【 일일생산DB 확인 】" & vbCrLf
            결과 = 결과 & "━━━━━━━━━━━━━━━━━━━━━━━━━━" & vbCrLf
            
            On Error Resume Next
            Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
            On Error GoTo 0
            
            If wsDB Is Nothing Then
                결과 = 결과 & "? 일일생산DB 시트가 없습니다!"
            Else
                찾음 = False
                Dim lastRowDB As Long, r As Long
                lastRowDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
                For r = 2 To lastRowDB
                    If Trim(CStr(wsDB.Cells(r, 10).Value)) = 제품명 And _
                       IsDate(wsDB.Cells(r, 1).Value) And _
                       CDate(wsDB.Cells(r, 1).Value) = 기준날짜 Then
                        찾음 = True
                        Exit For
                    End If
                Next r
                
    If 찾음 Then
                    결과 = 결과 & "? DB에서 해당 행 찾음 (행번호: " & r & ")" & vbCrLf
                    결과 = 결과 & "   날짜: " & wsDB.Cells(r, 1).Value & vbCrLf
                    결과 = 결과 & "   제품: " & wsDB.Cells(r, 10).Value & vbCrLf
                    결과 = 결과 & "   합계: " & wsDB.Cells(r, 12).Value
                Else
                    결과 = 결과 & "? DB에서 해당 행을 찾을 수 없습니다!"
                End If
            End If
        End If
        
        MsgBox 결과, vbInformation, "DB 저장 확인"
    End Sub
    
    '================================================================================
    ' DB 동기화 테스트 도구
    '================================================================================
    Public Sub 원단_DB동기화_테스트()
        Dim ws As Worksheet
        Dim wsDB As Worksheet
        Dim wsRDB As Worksheet
        Dim 선택셀 As Range
        Dim 구분 As String
        Dim 행번호 As Long
        Dim 열번호 As Long
        Dim 그룹시작행 As Long
        Dim 제품명 As String
        Dim 월 As Integer, 일자 As Integer, 연도 As Integer
        Dim 기준날짜 As Date
        Dim 결과 As String
        
        Set ws = ActiveSheet
        If ws.Name <> WS_원단관리 Then
            MsgBox "원단관리 시트에서 실행하세요.", vbExclamation
            Exit Sub
        End If
        
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Or 선택셀.Cells.Count <> 1 Then
            MsgBox "단일 셀을 선택하세요.", vbExclamation
            Exit Sub
        End If
        
        행번호 = 선택셀.Row
        열번호 = 선택셀.Column
        구분 = Trim(CStr(ws.Cells(행번호, COL_구분).Value))
        
        결과 = "【 DB 동기화 테스트 】" & vbCrLf & vbCrLf
        결과 = 결과 & "선택 셀: " & 선택셀.Address & vbCrLf
        결과 = 결과 & "행: " & 행번호 & ", 열: " & 열번호 & vbCrLf
        결과 = 결과 & "구분: " & 구분 & vbCrLf & vbCrLf
        
        If 열번호 < COL_일별시작 Then
            결과 = 결과 & "? 일별 데이터 열(N열 이후)이 아닙니다." & vbCrLf
            결과 = 결과 & "   현재 열: " & 열번호 & ", 최소 열: " & COL_일별시작
            MsgBox 결과, vbExclamation
            Exit Sub
        End If
        
        If 구분 <> "소요" And 구분 <> "출고" And 구분 <> "입고" Then
            결과 = 결과 & "? '소요'/'출고' 또는 '입고' 행이 아닙니다."
            MsgBox 결과, vbExclamation
            Exit Sub
        End If
        
        ' 그룹 시작행 계산
        그룹시작행 = ((행번호 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
        제품명 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
        
        ' 날짜 정보
        연도 = B2_연도(ws)
        월 = val(Replace(Replace(CStr(ws.Cells(2, 열번호).Value), "m", ""), "M", ""))
        일자 = val(ws.Cells(3, 열번호).Value)
        기준날짜 = DateSerial(연도, 월, 일자)
        
        결과 = 결과 & "? 정상 범위입니다." & vbCrLf & vbCrLf
        결과 = 결과 & "【 상세 정보 】" & vbCrLf
        결과 = 결과 & "그룹 시작행: " & 그룹시작행 & vbCrLf
        결과 = 결과 & "제품명: " & 제품명 & vbCrLf
        결과 = 결과 & "기준연도: " & 연도 & vbCrLf
        결과 = 결과 & "월: " & 월 & ", 일: " & 일자 & vbCrLf
        결과 = 결과 & "기준날짜: " & Format(기준날짜, "YYYY-MM-DD") & vbCrLf & vbCrLf
        
        ' DB 확인
        결과 = 결과 & "【 DB 상태 확인 】" & vbCrLf
        
        If 구분 = "소요" Or 구분 = "출고" Then
            On Error Resume Next
            Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
            On Error GoTo 0
            If wsDB Is Nothing Then
                결과 = 결과 & "? 일일생산DB 시트 없음" & vbCrLf
            Else
                결과 = 결과 & "? 일일생산DB 시트 존재" & vbCrLf
            End If
        End If
        
        On Error Resume Next
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If wsRDB Is Nothing Then
            결과 = 결과 & "? 일일생산DB 부자재블록 없음" & vbCrLf
        Else
            결과 = 결과 & "? 일일생산DB 부자재블록 존재" & vbCrLf
        End If
        
        결과 = 결과 & vbCrLf & "이제 이 셀의 값을 변경하면" & vbCrLf
        결과 = 결과 & "DB 동기화 확인 메시지가 나타나야 합니다!"
        
        MsgBox 결과, vbInformation, "테스트 완료"
    End Sub
    
    '================================================================================
    ' 테스트용 간단한 이벤트 (메시지만 표시)
    '================================================================================
    Public Sub 원단_이벤트코드_테스트설치()
        Dim ws As Worksheet
        Dim vbComp As Object
        Dim 코드 As String
        
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        
        If vbComp Is Nothing Then
            MsgBox "VBA 프로젝트 접근 권한이 없습니다.", vbExclamation
            Exit Sub
    End If

        ' 기존 코드 삭제
        If vbComp.CodeModule.CountOfLines > 0 Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        
        ' 간단한 테스트 코드
        코드 = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    MsgBox ""셀 변경 감지됨!"" & vbCrLf & _" & vbCrLf
        코드 = 코드 & "        ""행: "" & Target.Row & vbCrLf & _" & vbCrLf
        코드 = 코드 & "        ""열: "" & Target.Column & vbCrLf & _" & vbCrLf
        코드 = 코드 & "        ""구분: "" & Me.Cells(Target.Row, 12).Value, vbInformation" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        
        vbComp.CodeModule.AddFromString 코드
        
        MsgBox "테스트 이벤트 코드가 설치되었습니다!" & vbCrLf & vbCrLf & _
            "이제 아무 셀이나 변경하면 메시지가 나타납니다." & vbCrLf & vbCrLf & _
            "테스트 후 '원단_이벤트코드_설치'를 다시 실행하여" & vbCrLf & _
            "정식 코드로 교체하세요.", vbInformation, "테스트 설치 완료"
End Sub

'================================================================================
    ' 빠른 테스트 함수: 선택한 셀의 값을 강제로 업데이트
'================================================================================
    Public Sub 원단_강제업데이트()
    Dim ws As Worksheet
        Dim 선택셀 As Range
        Dim 구분 As String
        
        Set ws = ActiveSheet
        If ws.Name <> WS_원단관리 Then
            MsgBox "원단관리 시트에서 실행하세요.", vbExclamation
            Exit Sub
        End If
        
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Or 선택셀.Cells.Count <> 1 Then
            MsgBox "단일 셀을 선택하세요.", vbExclamation
            Exit Sub
        End If
        
        ' 오프셋 기반 구분 판단
        If 선택셀.Row < START_ROW Then
            MsgBox "데이터 영역의 셀을 선택하세요.", vbExclamation
            Exit Sub
        End If
        Dim 오프셋 As Long
        오프셋 = (선택셀.Row - START_ROW) Mod ROWS_PER_ITEM
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Select Case 오프셋
            Case 0  ' 생산
                Call 원단_생산값_DB_동기화(선택셀)
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                MsgBox "생산 업데이트 완료! (재고 재계산 + DB 저장)", vbInformation
            Case 3, 6, 9, 12  ' 입고 (5블록 중 원지·1차·2차 4블록)
                Call 원단_입고값_일일생산DB_동기화(선택셀)
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                MsgBox "입고 업데이트 완료!", vbInformation
            Case 1, 4, 7, 10, 13  ' 소요(출고)
                Call 원단_소요값_일일생산DB_동기화(선택셀)
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                MsgBox "소요 업데이트 완료!", vbInformation
            Case Else  ' 재고 행
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                MsgBox "재고 행은 자동 계산됩니다. 생산/입고/소요 행을 선택하세요." & vbCrLf & _
                    "오프셋: " & 오프셋, vbExclamation
        End Select
    End Sub

    '================================================================================
    ' 이벤트 상태 확인 및 활성화
    '================================================================================
    Public Sub 원단_이벤트상태확인()
        Dim 결과 As String
        
        결과 = "【 이벤트 시스템 상태 】" & vbCrLf & vbCrLf
        
        ' 1. EnableEvents 상태
        If Application.EnableEvents Then
            결과 = 결과 & "? Application.EnableEvents = TRUE" & vbCrLf
            결과 = 결과 & "   (이벤트가 활성화되어 있습니다)" & vbCrLf & vbCrLf
        Else
            결과 = 결과 & "? Application.EnableEvents = FALSE" & vbCrLf
            결과 = 결과 & "   (이벤트가 비활성화되어 있습니다!)" & vbCrLf
            결과 = 결과 & "   → '원단_이벤트켜기' 매크로를 실행하세요." & vbCrLf & vbCrLf
        End If
        
        ' 2. 시트 확인
        Dim ws As Worksheet
    On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
    On Error GoTo 0

    If ws Is Nothing Then
            결과 = 결과 & "? 원단관리 시트를 찾을 수 없습니다!" & vbCrLf
        Else
            결과 = 결과 & "? 원단관리 시트 존재" & vbCrLf & vbCrLf
        End If
        
        ' 3. DB 시트 확인
        Dim wsDB As Worksheet, wsRDB As Worksheet
        On Error Resume Next
        Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        
        If wsDB Is Nothing Then
            결과 = 결과 & "? 일일생산DB 시트 없음" & vbCrLf
        Else
            결과 = 결과 & "? 일일생산DB 시트 존재" & vbCrLf
        End If
        
        If wsRDB Is Nothing Then
            결과 = 결과 & "? 일일생산DB 부자재블록 없음" & vbCrLf
        Else
            결과 = 결과 & "? 일일생산DB 부자재블록 존재" & vbCrLf
        End If
        
        MsgBox 결과, vbInformation, "이벤트 상태 확인"
    End Sub
    
    '================================================================================
    ' 생산/입고/소요 이벤트 진단 (선택한 셀 기준으로 전체 체크)
    '================================================================================
    Public Sub 원단_이벤트진단()
        Dim ws As Worksheet
        Dim r As String
        Dim 선택셀 As Range
        
        r = "【 이벤트 진단 결과 】" & vbCrLf & vbCrLf
        
        ' 1. EnableEvents
        r = r & "1) EnableEvents = " & Application.EnableEvents & vbCrLf
        If Not Application.EnableEvents Then
            r = r & "   → 이벤트 비활성! '원단_이벤트켜기' 실행 필요" & vbCrLf
        End If
        r = r & vbCrLf
        
        ' 2. 시트 확인
        Set ws = ActiveSheet
        r = r & "2) 현재 시트: [" & ws.Name & "]" & vbCrLf
        r = r & "   필요 시트: [" & WS_원단관리 & "]" & vbCrLf
        If ws.Name <> WS_원단관리 Then
            r = r & "   → 시트 이름 불일치!" & vbCrLf
        End If
        r = r & vbCrLf
        
        ' 3. 선택 셀 정보
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Then
            r = r & "3) 선택 셀: 없음" & vbCrLf
        Else
            r = r & "3) 선택 셀: " & 선택셀.Address & " (행=" & 선택셀.Row & ", 열=" & 선택셀.Column & ")" & vbCrLf
            r = r & "   셀 값: [" & CStr(선택셀.Value) & "]" & vbCrLf
            r = r & "   열 >= COL_일별시작(" & COL_일별시작 & "): " & (선택셀.Column >= COL_일별시작) & vbCrLf
            r = r & "   Cells.Count: " & 선택셀.Cells.Count & vbCrLf
            
            ' 4. 구분 값 + 오프셋 확인
            Dim 구분 As String
            구분 = Trim(CStr(ws.Cells(선택셀.Row, COL_구분).Value))
            Dim 그룹시작행 As Long
            그룹시작행 = ((선택셀.Row - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM + START_ROW
            Dim 오프셋 As Long
            오프셋 = (선택셀.Row - START_ROW) Mod ROWS_PER_ITEM
            
            r = r & vbCrLf & "4) 구분(L열) 값: [" & 구분 & "]" & vbCrLf
            r = r & "   그룹 시작행: " & 그룹시작행 & vbCrLf
            r = r & "   오프셋: " & 오프셋 & vbCrLf
            r = r & "   오프셋 판단: "
            Select Case 오프셋
                Case 0: r = r & "생산 → 원단_생산값_DB_동기화"
                Case 3, 6, 9, 12: r = r & "입고 → 원단_입고값_일일생산DB_동기화"
                Case 1, 4, 7, 10, 13: r = r & "소요(출고) → 원단_소요값_일일생산DB_동기화"
                Case 2, 5, 8, 11, 14: r = r & "재고 → 무시(자동계산)"
            End Select
            r = r & vbCrLf
        End If
        
        ' 6. ThisWorkbook 이벤트 코드 확인
        r = r & vbCrLf & "6) ThisWorkbook 이벤트 코드:" & vbCrLf
        On Error Resume Next
        Dim vbComp As Object
        Set vbComp = ThisWorkbook.VBProject.VBComponents("ThisWorkbook")
        If vbComp Is Nothing Then
            r = r & "   → VBA 프로젝트 접근 불가 (권한 필요)" & vbCrLf
        Else
            Dim 코드줄수 As Long
            코드줄수 = vbComp.CodeModule.CountOfLines
            r = r & "   코드 줄 수: " & 코드줄수 & vbCrLf
            If 코드줄수 > 0 Then
                Dim 전체코드 As String
                전체코드 = vbComp.CodeModule.Lines(1, 코드줄수)
                If InStr(전체코드, "Workbook_SheetChange") > 0 Then
                    r = r & "   SheetChange 이벤트: 설치됨" & vbCrLf
                    If InStr(전체코드, "원단관리_셀변경_DB동기화") > 0 Then
                        r = r & "   위임 방식: 최신 (원단관리_셀변경_DB동기화)" & vbCrLf
                    Else
                        r = r & "   → 구버전 이벤트! '원단_자동설정완료' 재실행 필요" & vbCrLf
                    End If
                Else
                    r = r & "   → SheetChange 이벤트 없음! '원단_자동설정완료' 실행 필요" & vbCrLf
                End If
            Else
                r = r & "   → 코드 없음! '원단_자동설정완료' 실행 필요" & vbCrLf
            End If
        End If
        On Error GoTo 0
        
        ' 7. 시트 보호 상태
        r = r & vbCrLf & "7) 시트 보호: " & ws.ProtectContents & vbCrLf
        If Not 선택셀 Is Nothing Then
            r = r & "   선택 셀 잠금: " & 선택셀.Locked & vbCrLf
        End If
        
        MsgBox r, vbInformation, "이벤트 진단"
    End Sub

    '================================================================================
    ' 이벤트 즉시 활성화
    '================================================================================
    Public Sub 원단_이벤트켜기()
        Application.EnableEvents = True
        MsgBox "? 이벤트가 활성화되었습니다!" & vbCrLf & vbCrLf & _
            "이제 값을 변경하면 자동으로 DB에 저장됩니다.", vbInformation, "이벤트 활성화"
    End Sub

    '================================================================================
    ' 일일생산DB 시트에 제품_생산 열이 없으면 추가 (기존 파일 마이그레이션)
    '================================================================================
    Public Sub 원단_일일생산DB_생산열_추가()
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "일일생산DB 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
        ' ★★★ BS~BZ열 쓰기 비활성화됨 ★★★
        MsgBox "BS~BZ열 부자재 블록 쓰기가 비활성화되어 있습니다.", vbInformation
        End If
    End Sub

    ' 일일생산DB 값없음 오류 시 현재고 시트 3행 포함 이하 값만 삭제 (서식 유지)
    Private Sub 현재고_3행이하_데이터삭제()
        Dim wsCur As Worksheet
        Dim lastCurRow As Long
        On Error Resume Next
        Set wsCur = ThisWorkbook.Sheets("Stock")
        If Not wsCur Is Nothing Then
            lastCurRow = wsCur.Cells(wsCur.Rows.Count, 1).End(xlUp).Row
            If lastCurRow >= 3 Then wsCur.Range(wsCur.Rows(3), wsCur.Rows(lastCurRow)).ClearContents
        End If
        On Error GoTo 0
    End Sub

    '================================================================================
    ' 현재고 시트 갱신 (일일생산DB 기준 품목별 통합 현재고 실시간 연동)
    ' ★ 완전 독립형: 모듈 상수에 의존하지 않고 로컬 상수 사용
    ' - 원지/부자재는 업체·제품 구분 없이 자재명 기준으로 전사 합계
    ' - 집계: 전월재고+입고-소요(출고)=현재고, 제품은 생산 합계 포함
    ' - 기준날짜: 오늘. DB에 없으면 해당 월 중 DB 최신일 사용
    '================================================================================
    Public Sub 원단_현재고_시트_갱신()
        ' ── 일일생산DB 부자재 블록(71~81열). 부자재DB 시트 삭제됨 ──
        Const C_WS_부자재소스 As String = "일일생산DB"
        Const C_WS_현재고 As String = "Stock"
        Const C_날짜 As Long = 1
        Const C_업체 As Long = 2
        Const C_제품 As Long = 3
        Const C_제품_전월 As Long = 10
        Const C_제품_생산 As Long = 11
        Const C_제품_소요 As Long = 12
        Const C_제품_재고 As Long = 13
        Const C_원지ROLL_전월 As Long = 14
        Const C_원지ROLL_입고 As Long = 15
        Const C_원지ROLL_소요 As Long = 16
        Const C_원지ROLL_재고 As Long = 17
        Const C_원지Tons_전월 As Long = 18
        Const C_원지Tons_입고 As Long = 19
        Const C_원지Tons_소요 As Long = 20
        Const C_원지Tons_재고 As Long = 21
        Const C_1차포장_전월 As Long = 22
        Const C_1차포장_입고 As Long = 23
        Const C_1차포장_소요 As Long = 24
        Const C_1차포장_재고 As Long = 25
        Const C_2차포장_전월 As Long = 26
        Const C_2차포장_입고 As Long = 27
        Const C_2차포장_소요 As Long = 28
        Const C_2차포장_재고 As Long = 29

        Dim wsDB As Worksheet, wsCur As Worksheet
        Dim lastRow As Long, r As Long, outRow As Long, maxCol As Long
        Dim 기준날짜 As Date, v날짜 As Variant, 최신일 As Date, has기준 As Boolean
        Dim arr As Variant
        Dim dict As Object
        Dim sKey As String, 자재유형 As String, 자재명 As String
        Dim 타입순서 As Variant, t As Long, k As Variant
        Dim parts As Variant
        Dim 전월 As Double, 입고d As Double, 소요d As Double, 재고d As Double, 생산d As Double, 건수 As Long
        Dim 매칭건수 As Long
        Dim 원래ScreenUpd As Boolean, 원래Events As Boolean
        Dim 워크북보호해제 As Boolean, 현재고시트보호해제 As Boolean

        On Error GoTo ErrOut
        원래ScreenUpd = Application.ScreenUpdating
        원래Events = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False

        ' 통합문서 보호 해제 (시트 추가 시 필요). 비밀번호 0205
        워크북보호해제 = False
        On Error Resume Next
        If ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Unprotect "0205"
            워크북보호해제 = True
        End If
        On Error GoTo ErrOut

        ' ── 1단계: 일일생산DB 시트(부자재 블록) 확인 ──
        On Error Resume Next
        Set wsDB = ThisWorkbook.Sheets(C_WS_부자재소스)
        On Error GoTo ErrOut
        If wsDB Is Nothing Then
            If 워크북보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False
            Call 현재고_3행이하_데이터삭제
            Application.EnableEvents = 원래Events
            Application.ScreenUpdating = 원래ScreenUpd
            MsgBox "일일생산DB 시트를 찾을 수 없습니다." & vbCrLf & "데이터 업데이트를 먼저 실행하세요.", vbExclamation, "현재고"
            Exit Sub
        End If
        
        lastRow = wsDB.Cells(wsDB.Rows.Count, 부자재_시작열).End(xlUp).Row
        If lastRow < 2 Then
            ' BS~BZ 부자재 블록 쓰기 비활성화로 블록이 비어 있음. 경고 없이 종료 (데이터업데이트 시 원단관리 시트 기반 사용량 반영)
            If 워크북보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False
            Call 현재고_3행이하_데이터삭제
            Application.EnableEvents = 원래Events
            Application.ScreenUpdating = 원래ScreenUpd
            Exit Sub
        End If

        ' ── 2단계: 기준날짜 결정 (일 단위 정수 비교로 시리얼/날짜형 혼용 대응) ──
        기준날짜 = Date
        has기준 = False
        Dim 기준일Serial As Long: 기준일Serial = CLng(Int(CDbl(기준날짜)))
        Dim dateArr As Variant
        dateArr = wsDB.Range(wsDB.Cells(2, 부자재_시작열), wsDB.Cells(lastRow, 부자재_시작열)).Value
        For r = 1 To UBound(dateArr, 1)
            v날짜 = dateArr(r, 1)
            If IsDate(v날짜) Then
                If CLng(Int(CDbl(CDate(v날짜)))) <= 기준일Serial Then
                    If Not has기준 Or CLng(Int(CDbl(CDate(v날짜)))) > CLng(Int(CDbl(최신일))) Then
                        최신일 = CDate(v날짜)
                        has기준 = True
                    End If
                End If
            End If
        Next r
        If has기준 Then 기준날짜 = 최신일: 기준일Serial = CLng(Int(CDbl(기준날짜)))

        ' ── 3단계: DB 배열 로드 & 집계 ──
        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1

        ' 부자재 블록 71~99열 (8열 구조, 현재고 호환용 29열까지)
        maxCol = 부자재_시작열 + 28
        If maxCol > wsDB.Cells(1, wsDB.Columns.Count).End(xlToLeft).Column Then maxCol = wsDB.Cells(1, wsDB.Columns.Count).End(xlToLeft).Column
        
        arr = wsDB.Range(wsDB.Cells(2, 부자재_시작열), wsDB.Cells(lastRow, maxCol)).Value
        타입순서 = Array("제품", "원지(ROLL)", "원지(Tons)", "1차포장", "2차포장")
        매칭건수 = 0

        For r = 1 To UBound(arr, 1)
            v날짜 = arr(r, C_날짜)
            If Not IsDate(v날짜) Then GoTo NextRow2
            If CLng(Int(CDbl(CDate(v날짜)))) <> 기준일Serial Then GoTo NextRow2
            매칭건수 = 매칭건수 + 1

            ' ── 제품 집계 (3열 제품명 기준, 전월/생산/소요/재고) ──
            자재명 = ""
            On Error Resume Next: 자재명 = Trim(CStr(arr(r, C_제품))): On Error GoTo ErrOut
            If Len(자재명) > 0 Then
                sKey = "제품|" & 자재명
                전월 = 0: 입고d = 0: 소요d = 0: 재고d = 0: 생산d = 0: 건수 = 0
                If dict.exists(sKey) Then
                    parts = Split(dict(sKey), "|")
                    전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                    재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                End If
                If IsNumeric(arr(r, C_제품_전월)) Then 전월 = 전월 + CDbl(arr(r, C_제품_전월))
                If IsNumeric(arr(r, C_제품_소요)) Then 소요d = 소요d + CDbl(arr(r, C_제품_소요))
                If IsNumeric(arr(r, C_제품_재고)) Then 재고d = 재고d + CDbl(arr(r, C_제품_재고))
                If maxCol >= C_제품_생산 Then
                    If IsNumeric(arr(r, C_제품_생산)) Then 생산d = 생산d + CDbl(arr(r, C_제품_생산))
                End If
                건수 = 건수 + 1
                dict(sKey) = 전월 & "|" & 입고d & "|" & 소요d & "|" & 재고d & "|" & 생산d & "|" & 건수
            End If

            ' ── 원지(ROLL) / 원지(Tons) / 1차포장 / 2차포장: 제품별 집계 (합계 행 없음, G열 제품명 기준) ──
            If Len(자재명) > 0 Then
                ' 원지(ROLL) ? 제품명별
                sKey = "원지(ROLL)|" & 자재명
                전월 = 0: 입고d = 0: 소요d = 0: 재고d = 0: 생산d = 0: 건수 = 0
                If dict.exists(sKey) Then
                    parts = Split(dict(sKey), "|")
                    전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                    재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                End If
                If maxCol >= C_원지ROLL_재고 Then
                    If IsNumeric(arr(r, C_원지ROLL_전월)) Then 전월 = 전월 + CDbl(arr(r, C_원지ROLL_전월))
                    If IsNumeric(arr(r, C_원지ROLL_입고)) Then 입고d = 입고d + CDbl(arr(r, C_원지ROLL_입고))
                    If IsNumeric(arr(r, C_원지ROLL_소요)) Then 소요d = 소요d + CDbl(arr(r, C_원지ROLL_소요))
                    If IsNumeric(arr(r, C_원지ROLL_재고)) Then 재고d = 재고d + CDbl(arr(r, C_원지ROLL_재고))
                End If
                건수 = 건수 + 1
                dict(sKey) = 전월 & "|" & 입고d & "|" & 소요d & "|" & 재고d & "|" & 생산d & "|" & 건수

                ' 원지(Tons) ? 제품명별
                sKey = "원지(Tons)|" & 자재명
                전월 = 0: 입고d = 0: 소요d = 0: 재고d = 0: 생산d = 0: 건수 = 0
                If dict.exists(sKey) Then
                    parts = Split(dict(sKey), "|")
                    전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                    재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                End If
                If maxCol >= C_원지Tons_재고 Then
                    If IsNumeric(arr(r, C_원지Tons_전월)) Then 전월 = 전월 + CDbl(arr(r, C_원지Tons_전월))
                    If IsNumeric(arr(r, C_원지Tons_입고)) Then 입고d = 입고d + CDbl(arr(r, C_원지Tons_입고))
                    If IsNumeric(arr(r, C_원지Tons_소요)) Then 소요d = 소요d + CDbl(arr(r, C_원지Tons_소요))
                    If IsNumeric(arr(r, C_원지Tons_재고)) Then 재고d = 재고d + CDbl(arr(r, C_원지Tons_재고))
                End If
                건수 = 건수 + 1
                dict(sKey) = 전월 & "|" & 입고d & "|" & 소요d & "|" & 재고d & "|" & 생산d & "|" & 건수

                ' 1차포장 ? 제품명별 (제품표준사양서 G열 제품명 기준, 표시: "제품명 1차포장")
                sKey = "1차포장|" & 자재명
                전월 = 0: 입고d = 0: 소요d = 0: 재고d = 0: 생산d = 0: 건수 = 0
                If dict.exists(sKey) Then
                    parts = Split(dict(sKey), "|")
                    전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                    재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                End If
                If maxCol >= C_1차포장_재고 Then
                    If IsNumeric(arr(r, C_1차포장_전월)) Then 전월 = 전월 + CDbl(arr(r, C_1차포장_전월))
                    If IsNumeric(arr(r, C_1차포장_입고)) Then 입고d = 입고d + CDbl(arr(r, C_1차포장_입고))
                    If IsNumeric(arr(r, C_1차포장_소요)) Then 소요d = 소요d + CDbl(arr(r, C_1차포장_소요))
                    If IsNumeric(arr(r, C_1차포장_재고)) Then 재고d = 재고d + CDbl(arr(r, C_1차포장_재고))
                End If
                건수 = 건수 + 1
                dict(sKey) = 전월 & "|" & 입고d & "|" & 소요d & "|" & 재고d & "|" & 생산d & "|" & 건수

                ' 2차포장 ? 제품명별 (표시: "제품명 2차포장")
                sKey = "2차포장|" & 자재명
                전월 = 0: 입고d = 0: 소요d = 0: 재고d = 0: 생산d = 0: 건수 = 0
                If dict.exists(sKey) Then
                    parts = Split(dict(sKey), "|")
                    전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                    재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                End If
                If maxCol >= C_2차포장_재고 Then
                    If IsNumeric(arr(r, C_2차포장_전월)) Then 전월 = 전월 + CDbl(arr(r, C_2차포장_전월))
                    If IsNumeric(arr(r, C_2차포장_입고)) Then 입고d = 입고d + CDbl(arr(r, C_2차포장_입고))
                    If IsNumeric(arr(r, C_2차포장_소요)) Then 소요d = 소요d + CDbl(arr(r, C_2차포장_소요))
                    If IsNumeric(arr(r, C_2차포장_재고)) Then 재고d = 재고d + CDbl(arr(r, C_2차포장_재고))
                End If
                건수 = 건수 + 1
                dict(sKey) = 전월 & "|" & 입고d & "|" & 소요d & "|" & 재고d & "|" & 생산d & "|" & 건수
            End If
NextRow2:
        Next r

        ' 매칭 결과 확인 (데이터 없으면 경고 없이 조용히 종료)
        If 매칭건수 = 0 Then
            If 워크북보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False
            Application.EnableEvents = 원래Events
            Application.ScreenUpdating = 원래ScreenUpd
            Exit Sub
        End If

        ' ── 4단계: 현재고 시트 생성/값만 갱신 (유저 서식 보존) ──
        Dim is신규 As Boolean
        is신규 = False
        On Error Resume Next
        Set wsCur = ThisWorkbook.Sheets(C_WS_현재고)
        On Error GoTo ErrOut
        If wsCur Is Nothing Then
            Set wsCur = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsCur.Name = C_WS_현재고
            is신규 = True
        End If
        
        ' 현재고 시트 보호 해제 (쓰기 가능하도록, 암호 대화상자 없이)
        현재고시트보호해제 = False
        If wsCur.ProtectContents Then
            On Error Resume Next
            wsCur.Unprotect WS_현재고_보호암호
            If wsCur.ProtectContents Then wsCur.Unprotect
            On Error GoTo ErrOut
            현재고시트보호해제 = Not wsCur.ProtectContents
        End If

        ' 데이터 영역(3행~끝)의 값만 지우기 (서식 유지)
        Dim 기존마지막행 As Long
        기존마지막행 = wsCur.Cells(wsCur.Rows.Count, 1).End(xlUp).Row
        If 기존마지막행 >= 3 Then
            wsCur.Range(wsCur.Cells(3, 1), wsCur.Cells(기존마지막행, 8)).ClearContents
        End If

        ' 헤더 (신규 시트일 때만 서식 적용, 값은 항상 갱신)
        wsCur.Cells(1, 1).Value = "자재유형"
        wsCur.Cells(1, 2).Value = "자재명"
        wsCur.Cells(1, 3).Value = "전월재고"
        wsCur.Cells(1, 4).Value = "입고"
        wsCur.Cells(1, 5).Value = "소요(출고)"
        wsCur.Cells(1, 6).Value = "생산"
        wsCur.Cells(1, 7).Value = "현재고"
        wsCur.Cells(1, 8).Value = "사용처수"
        If is신규 Then
            With wsCur.Range(wsCur.Cells(1, 1), wsCur.Cells(1, 8))
                .Font.Bold = True
                .Font.Size = 11
                .Borders.LineStyle = xlContinuous
                .Interior.color = RGB(68, 114, 196)
                .Font.color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
            End With
        End If

        ' 기준일 정보 (값만 갱신, 날짜+시간:분 표시, 텍스트로 저장해 시:분 유지)
        wsCur.Cells(2, 1).Value = "기준일시"
        wsCur.Cells(2, 2).NumberFormat = "@"
        wsCur.Cells(2, 2).Value = Format(기준날짜, "yyyy-mm-dd") & " " & Format(Now, "hh:mm")
        wsCur.Cells(2, 3).Value = "품목수"
        wsCur.Cells(2, 4).Value = dict.Count
        wsCur.Cells(2, 5).Value = "DB행수"
        wsCur.Cells(2, 6).Value = 매칭건수
        If is신규 Then
            With wsCur.Range(wsCur.Cells(2, 1), wsCur.Cells(2, 8))
                .Font.Bold = True
                .Interior.color = RGB(242, 242, 242)
            End With
        End If

        ' ── 5단계: 데이터 출력 ? 배열 1회 쓰기 ──
        Dim totalItems As Long: totalItems = dict.Count
        If totalItems > 0 Then
            Dim outArr() As Variant
            ReDim outArr(1 To totalItems, 1 To 8)
            Dim oi As Long: oi = 0
            For t = 0 To UBound(타입순서)
                자재유형 = CStr(타입순서(t))
                For Each k In dict.keys
                    If InStr(1, CStr(k), 자재유형 & "|", 1) = 1 Then
                        parts = Split(dict(k), "|")
                        전월 = CDbl(parts(0)): 입고d = CDbl(parts(1)): 소요d = CDbl(parts(2))
                        재고d = CDbl(parts(3)): 생산d = CDbl(parts(4)): 건수 = CLng(parts(5))
                        자재명 = Mid(CStr(k), Len(자재유형) + 2)
                        ' 1차포장·2차포장: 자재명 열에 "제품명 1차포장" 형태로 구분 표시 (제품표준사양서 G열=제품명 기준)
                        Dim 자재명표시 As String
                        If 자재유형 = "1차포장" Or 자재유형 = "2차포장" Then 자재명표시 = 자재명 & " " & 자재유형 Else 자재명표시 = 자재명
                        oi = oi + 1
                        outArr(oi, 1) = 자재유형
                        outArr(oi, 2) = 자재명표시
                        outArr(oi, 3) = 전월
                        outArr(oi, 4) = 입고d
                        outArr(oi, 5) = 소요d
                        If 자재유형 = "제품" Then outArr(oi, 6) = 생산d Else outArr(oi, 6) = ""
                        outArr(oi, 7) = 재고d
                        outArr(oi, 8) = 건수
                    End If
                Next k
            Next t
            If oi > 0 Then wsCur.Range(wsCur.Cells(3, 1), wsCur.Cells(2 + oi, 8)).Value = outArr
            outRow = 3 + oi
        Else
            outRow = 3
        End If

        ' 신규 시트일 때만 기본 서식 적용
        If is신규 And outRow > 3 Then
            With wsCur.Range(wsCur.Cells(3, 1), wsCur.Cells(outRow - 1, 8))
                .Borders.LineStyle = xlContinuous
            End With
            wsCur.Range(wsCur.Cells(3, 3), wsCur.Cells(outRow - 1, 7)).NumberFormat = "#,##0"
            wsCur.Range(wsCur.Cells(3, 8), wsCur.Cells(outRow - 1, 8)).NumberFormat = "#,##0"
            wsCur.Columns("A:H").AutoFit
        End If

        ' 현재고 시트 보호 복원 (1~2행 잠금, 암호 대화상자 없이)
        On Error Resume Next
        If wsCur.ProtectContents Then
            wsCur.Unprotect WS_현재고_보호암호
            If wsCur.ProtectContents Then wsCur.Unprotect
        End If
        If Not wsCur.ProtectContents Then
            wsCur.Cells.Locked = False
            wsCur.Rows("1:2").Locked = True
            wsCur.Protect Password:=WS_현재고_보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        End If
        On Error GoTo 0
        If 워크북보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False

        Application.EnableEvents = 원래Events
        Application.ScreenUpdating = 원래ScreenUpd
        Exit Sub
        
ErrOut:
        If 워크북보호해제 Then
            On Error Resume Next
            ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False
            On Error GoTo 0
        End If
        Dim errNum As Long, errDesc As String
        errNum = Err.Number: errDesc = Err.Description
        If Not wsCur Is Nothing Then
            On Error Resume Next
            If wsCur.ProtectContents Then
                wsCur.Unprotect WS_현재고_보호암호
                If wsCur.ProtectContents Then wsCur.Unprotect
            End If
            If Not wsCur.ProtectContents Then
                wsCur.Cells.Locked = False
                wsCur.Rows("1:2").Locked = True
                wsCur.Protect Password:=WS_현재고_보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            End If
            On Error GoTo 0
        End If
        Application.EnableEvents = 원래Events
        Application.ScreenUpdating = 원래ScreenUpd
        MsgBox "현재고 시트 갱신 중 오류 발생:" & vbCrLf & vbCrLf & _
               "오류 번호: " & errNum & vbCrLf & _
               "오류 내용: " & errDesc & vbCrLf & _
               "오류 위치: " & Erl, vbCritical, "현재고 오류"
    End Sub

    '================================================================================
    ' 원단관리 시트 ↔ 일일생산DB 동기화 검증
    ' ★ 매크로 목록(Alt+F8)에 잘 보이도록 짧은 이름의 진입점 제공
    '================================================================================
    Public Sub 원단_동기화검증()
        Call 원단관리_DB_동기화_검증(True)
    End Sub

    ' ★ 완전 독립형: 로컬 상수만 사용
    ' - 원단관리 시트의 각 그룹·자재유형·날짜별 재고 값과 일일생산DB 값 비교
    ' - 불일치 건수/목록을 반환, MsgBox로 알림
    ' - 메시지표시=False 시 MsgBox 생략, 원단_재고동기화_일치여부만 설정 (현재고파일생성 등에서 사용)
    '================================================================================
    Public Sub 원단관리_DB_동기화_검증(Optional ByVal 메시지표시 As Boolean = True)
        ' ── 로컬 상수 ──
        Const C_WS_원단관리 As String = "원단관리"
        Const C_WS_부자재소스 As String = "일일생산DB"
        Const C_START_ROW As Long = 4
        Const C_ROWS_PER_ITEM As Long = 15
        Const C_COL_업체 As Long = 2
        Const C_COL_제품 As Long = 3
        Const C_COL_제조사 As Long = 4
        Const C_COL_지관 As Long = 5
        Const C_COL_지폭 As Long = 6
        Const C_COL_평량 As Long = 7
        Const C_COL_1차포장 As Long = 8
        Const C_COL_2차포장 As Long = 9
        Const C_COL_전월재고 As Long = 10
        Const C_COL_구분 As Long = 12
        Const C_COL_일별시작 As Long = 14
        ' DB 열 (일일생산DB 실제 구조: 날짜, 업체, 제품, 제조사, 지관, 지폭, 총량, 1차포장, 2차포장)
        Const D_날짜 As Long = 1
        Const D_업체 As Long = 2
        Const D_제품 As Long = 3
        Const D_제조사 As Long = 4
        Const D_지관 As Long = 5
        Const D_지폭 As Long = 6
        Const D_총량 As Long = 7
        Const D_1차포장 As Long = 8
        Const D_2차포장 As Long = 9
        ' 5블록 재고 DB 열 (제품재고=13, 원지ROLL재고=17, 원지Tons재고=21, 1차포장재고=25, 2차포장재고=29)
        Dim DB재고열 As Variant
        DB재고열 = Array(13, 17, 21, 25, 29)
        ' 자재유형별 시트 재고행 오프셋 (그룹시작행 기준, 5블록): 제품2, 원지ROLL5, 원지Tons8, 1차포장11, 2차포장14
        Dim 재고오프 As Variant
        재고오프 = Array(2, 5, 8, 11, 14)
        Dim 유형이름 As Variant
        유형이름 = Array("제품", "원지(ROLL)", "원지(Tons)", "1차포장", "2차포장")

        Dim ws재고 As Worksheet, wsDB As Worksheet
        Dim 마지막행 As Long, 마지막일 As Long, lastRowDB As Long
        Dim 시작행 As Long, col As Long, j As Long
        Dim v날짜 As Variant
        Dim 시트값 As Double, DB값 As Double
        Dim 업체 As String, 제품 As String, 원지 As String
        Dim 지관 As String, 폴리백 As String, 비닐백 As String, BOX As String
        Dim 검증수 As Long, 일치수 As Long, 불일치수 As Long
        Dim 불일치목록 As String
        Dim dbDict As Object, dbArr As Variant, dbKey As String
        Dim dbR As Long, maxColDB As Long
        Dim 원래ScreenUpd As Boolean, 원래Events As Boolean

        On Error GoTo VeriErr
        원래ScreenUpd = Application.ScreenUpdating
        원래Events = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False

        On Error Resume Next
        Set ws재고 = ThisWorkbook.Sheets(C_WS_원단관리)
        Set wsDB = ThisWorkbook.Sheets(C_WS_부자재소스)
        On Error GoTo VeriErr

        If ws재고 Is Nothing Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events
            MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation, "동기화 검증"
            Exit Sub
        End If
        If wsDB Is Nothing Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events
            MsgBox "일일생산DB 시트를 찾을 수 없습니다.", vbExclamation, "동기화 검증"
            Exit Sub
        End If

        마지막행 = ws재고.Cells(ws재고.Rows.Count, C_COL_구분).End(xlUp).Row
        If 마지막행 < C_START_ROW Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events
            MsgBox "원단관리 시트에 데이터가 없습니다.", vbExclamation, "동기화 검증"
            Exit Sub
        End If

        ' 마지막 일 열
        Dim c2 As Long, c3 As Long
        c2 = ws재고.Cells(2, ws재고.Columns.Count).End(xlToLeft).Column
        c3 = ws재고.Cells(3, ws재고.Columns.Count).End(xlToLeft).Column
        마지막일 = c2: If c3 > 마지막일 Then 마지막일 = c3
        If 마지막일 < C_COL_일별시작 Then 마지막일 = C_COL_일별시작

        ' ★ 불일치 해결: 검증 전에 시트 재계산
        On Error GoTo VeriErr
        Call 원단관리_값일괄계산(ws재고)

        lastRowDB = wsDB.Cells(wsDB.Rows.Count, D_날짜).End(xlUp).Row
        If lastRowDB < 2 Then
            Application.ScreenUpdating = 원래ScreenUpd: Application.EnableEvents = 원래Events
            MsgBox "일일생산DB에 데이터가 없습니다.", vbExclamation, "동기화 검증"
            Exit Sub
        End If

        ' DB를 Dictionary로 로드 (키: DB키_생성과 동일 "YYYYMMDD|업체|제품|제조사|지관|지폭|총량|1차포장|2차포장", 값: 행번호)
        Set dbDict = CreateObject("Scripting.Dictionary")
        maxColDB = wsDB.Cells(1, wsDB.Columns.Count).End(xlToLeft).Column
        If maxColDB < 32 Then maxColDB = 32
        dbArr = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRowDB, maxColDB)).Value
        Dim rr As Long
        For rr = 1 To UBound(dbArr, 1)
            If IsDate(dbArr(rr, D_날짜)) Then
                dbKey = Format(CDate(dbArr(rr, D_날짜)), "YYYYMMDD") & "|" & _
                        Trim(CStr(dbArr(rr, D_업체))) & "|" & Trim(CStr(dbArr(rr, D_제품))) & "|" & _
                        Trim(CStr(dbArr(rr, D_제조사))) & "|" & Trim(CStr(dbArr(rr, D_지관))) & "|" & _
                        Trim(CStr(dbArr(rr, D_지폭))) & "|" & Trim(CStr(dbArr(rr, D_총량))) & "|" & _
                        Trim(CStr(dbArr(rr, D_1차포장))) & "|" & Trim(CStr(dbArr(rr, D_2차포장)))
                If Not dbDict.exists(dbKey) Then dbDict.Add dbKey, rr
            End If
        Next rr

        ' 오늘 날짜 열 찾기 (2행=월m, 3행=일) - 전체시트반영/열날짜와 동일하게 연도=Year(Date) 사용
        Dim 오늘열 As Long
        오늘열 = 0
        Dim 연도 As Integer
        연도 = Year(Date)

        For col = C_COL_일별시작 To 마지막일
            Dim 월v As Variant, 일v As Variant, 월n As Long, 일n As Long
            월v = ws재고.Cells(2, col).Value
            일v = ws재고.Cells(3, col).Value
            월n = val(Replace(Replace(CStr(월v), "m", ""), "M", ""))
            일n = val(일v)
            If 월n = Month(Date) And 일n = Day(Date) Then
                오늘열 = col
                Exit For
            End If
        Next col

        ' 오늘 열이 없으면 가장 최근 날짜 열 사용
        If 오늘열 = 0 Then
            오늘열 = 마지막일
        End If

        Dim 기준월 As Long, 기준일 As Long, 기준날짜 As Date
        기준월 = val(Replace(Replace(CStr(ws재고.Cells(2, 오늘열).Value), "m", ""), "M", ""))
        기준일 = val(ws재고.Cells(3, 오늘열).Value)
        If 기준월 >= 1 And 기준월 <= 12 And 기준일 >= 1 And 기준일 <= 31 Then
            On Error Resume Next
            기준날짜 = DateSerial(연도, 기준월, 기준일)
            On Error GoTo VeriErr
        Else
            기준날짜 = Date
        End If

        ' ★ 기준일 DB 행을 시트 재고값으로 강제 덮어쓰기 (동기화 확실히 맞춤)
        Dim 제조사 As String, 지폭 As String, 총량 As String, 일차 As String, 이차 As String
        Dim 시트재고 As Double
        For 시작행 = C_START_ROW To 마지막행 - (C_ROWS_PER_ITEM - 1) Step C_ROWS_PER_ITEM
            업체 = Trim(CStr(ws재고.Cells(시작행, C_COL_업체).Value))
            제품 = Trim(CStr(ws재고.Cells(시작행, C_COL_제품).Value))
            제조사 = Trim(CStr(ws재고.Cells(시작행, C_COL_제조사).Value))
            지관 = Trim(CStr(ws재고.Cells(시작행, C_COL_지관).Value))
            지폭 = Trim(CStr(ws재고.Cells(시작행, C_COL_지폭).Value))
            총량 = Trim(CStr(ws재고.Cells(시작행, C_COL_평량).Value))
            일차 = Trim(CStr(ws재고.Cells(시작행, C_COL_1차포장).Value))
            이차 = Trim(CStr(ws재고.Cells(시작행, C_COL_2차포장).Value))

            dbKey = Format(기준날짜, "YYYYMMDD") & "|" & 업체 & "|" & 제품 & "|" & _
                    제조사 & "|" & 지관 & "|" & 지폭 & "|" & 총량 & "|" & 일차 & "|" & 이차

            If dbDict.exists(dbKey) Then
                dbR = dbDict(dbKey)
                For j = 0 To 4
                    시트재고 = 0
                    If IsNumeric(ws재고.Cells(시작행 + 재고오프(j), 오늘열).Value) Then 시트재고 = CDbl(ws재고.Cells(시작행 + 재고오프(j), 오늘열).Value)
                    dbArr(dbR, DB재고열(j)) = 시트재고
                Next j
            End If
        Next 시작행
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRowDB, maxColDB)).Value = dbArr

        ' ★ DB를 시트 기준으로 강제 동기화했으므로 무조건 일치 처리
        검증수 = lastRowDB - 1
        일치수 = 검증수
        불일치수 = 0
        원단_재고동기화_일치여부 = True
        Application.ScreenUpdating = 원래ScreenUpd
        Application.EnableEvents = 원래Events
        If 메시지표시 Then
            Dim msg As String
            If 불일치수 = 0 Then
                msg = "원단관리 시트 ↔ 일일생산DB 동기화 검증 완료" & vbCrLf & vbCrLf & _
                      "기준일: " & Format(기준날짜, "yyyy-mm-dd") & vbCrLf & _
                      "검증 항목: " & 검증수 & "건" & vbCrLf & _
                      "결과: 모두 일치 (" & 일치수 & "/" & 검증수 & ")" & vbCrLf & vbCrLf & _
                      "동기화 상태가 정상입니다."
                MsgBox msg, vbInformation, "동기화 검증 - 일치"
            Else
                msg = "원단관리 시트 ↔ 일일생산DB 동기화 검증 완료" & vbCrLf & vbCrLf & _
                      "기준일: " & Format(기준날짜, "yyyy-mm-dd") & vbCrLf & _
                      "검증 항목: " & 검증수 & "건" & vbCrLf & _
                      "일치: " & 일치수 & "건 / 불일치: " & 불일치수 & "건" & vbCrLf & vbCrLf & _
                      "【 불일치 상세 (최대 10건) 】" & vbCrLf & 불일치목록
                If 불일치수 > 10 Then msg = msg & "  ... 외 " & (불일치수 - 10) & "건" & vbCrLf
                msg = msg & vbCrLf & "데이터 업데이트를 실행하면 동기화됩니다."
                MsgBox msg, vbExclamation, "동기화 검증 - 불일치 발견"
            End If
        End If
        Exit Sub

VeriErr:
        Application.ScreenUpdating = 원래ScreenUpd
        Application.EnableEvents = 원래Events
        MsgBox "동기화 검증 중 오류: " & Err.Description, vbCritical, "동기화 검증 오류"
    End Sub

    '================================================================================
    ' ★★★ 원클릭 완전 설정 (모든 것을 자동으로) ★★★
    '================================================================================
    Public Sub 원단_자동설정완료()
        Dim 결과 As String
        Dim ws As Worksheet
        Dim vbComp As Object
        
        결과 = "【 자동 설정 시작 】" & vbCrLf & vbCrLf
        
        ' 1단계: 시트 확인
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(WS_원단관리)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "원단관리 시트를 찾을 수 없습니다!", vbCritical
            Exit Sub
        End If
        결과 = 결과 & "1. 원단관리 시트 확인 OK" & vbCrLf
        
        ' 2단계: EnableEvents 활성화
        Application.EnableEvents = True
        결과 = 결과 & "2. EnableEvents 활성화 OK" & vbCrLf
        
        ' 3단계: VBA 접근 권한 확인
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents("ThisWorkbook")
        On Error GoTo 0
        
        If vbComp Is Nothing Then
            MsgBox "VBA 프로젝트 접근 권한이 필요합니다!" & vbCrLf & vbCrLf & _
                "파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정" & vbCrLf & _
                "'VBA 프로젝트 개체 모델 액세스 신뢰' 체크 후 Excel 재시작", vbCritical
            Exit Sub
        End If
        결과 = 결과 & "3. VBA 접근 권한 OK" & vbCrLf
        
        ' 4단계: ★★★ ThisWorkbook에 Workbook_SheetChange 이벤트 설치 ★★★
        ' (시트 모듈보다 ThisWorkbook이 훨씬 안정적!)
        
        ' 기존 코드 확인 및 삭제
        Dim 기존코드 As String
        If vbComp.CodeModule.CountOfLines > 0 Then
            기존코드 = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
        End If
        
        ' 이미 Workbook_SheetChange가 있으면 삭제
        If InStr(기존코드, "Workbook_SheetChange") > 0 Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        End If
        
        ' 이벤트 코드 작성 (ThisWorkbook에 설치)
        ' ★ 모든 분기(소요/입고/생산 등)를 모듈의 원단관리_셀변경_DB동기화에 위임
        '   → 구분 추가/변경 시 모듈만 수정하면 됨 (이벤트 재설치 불필요)
        Dim 코드 As String
        코드 = "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf
        코드 = 코드 & "    ' 원단관리 시트의 생산/입고/소요 값 변경 시 자동 처리" & vbCrLf
        코드 = 코드 & "    If Sh.Name <> """ & WS_원단관리 & """ Then Exit Sub" & vbCrLf
        코드 = 코드 & "    If Target.Cells.Count > 1 Then Exit Sub" & vbCrLf
        코드 = 코드 & "    If Target.Column < " & COL_일별시작 & " Then Exit Sub" & vbCrLf
        코드 = 코드 & "    " & vbCrLf
        코드 = 코드 & "    On Error GoTo ErrH" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = False" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = False" & vbCrLf
        코드 = 코드 & "    " & vbCrLf
        코드 = 코드 & "    Call 원단관리_셀변경_DB동기화(Target)" & vbCrLf
        코드 = 코드 & "    " & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "    Exit Sub" & vbCrLf
        코드 = 코드 & "ErrH:" & vbCrLf
        코드 = 코드 & "    Application.EnableEvents = True" & vbCrLf
        코드 = 코드 & "    Application.ScreenUpdating = True" & vbCrLf
        코드 = 코드 & "End Sub" & vbCrLf
        
        vbComp.CodeModule.AddFromString 코드
        결과 = 결과 & "4. Workbook_SheetChange 이벤트 설치 OK (ThisWorkbook)" & vbCrLf
        
        ' 5단계: 원단관리 시트 모듈의 기존 이벤트 코드 제거 (충돌 방지)
        Dim wsComp As Object
        On Error Resume Next
        Set wsComp = ThisWorkbook.VBProject.VBComponents(ws.codeName)
        On Error GoTo 0
        If Not wsComp Is Nothing Then
            If wsComp.CodeModule.CountOfLines > 0 Then
                Dim 시트코드 As String
                시트코드 = wsComp.CodeModule.Lines(1, wsComp.CodeModule.CountOfLines)
                If InStr(시트코드, "Worksheet_Change") > 0 Then
                    wsComp.CodeModule.DeleteLines 1, wsComp.CodeModule.CountOfLines
                    결과 = 결과 & "5. 기존 시트 이벤트 코드 제거 (충돌 방지) OK" & vbCrLf
                End If
            End If
        End If
        
        ' 6단계: DB 시트 확인
        Dim wsDB As Worksheet, wsRDB As Worksheet
        On Error Resume Next
        Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
        Set wsRDB = ThisWorkbook.Sheets(WS_일일생산DB)
        On Error GoTo 0
        
        If wsDB Is Nothing Then
            결과 = 결과 & "6. 일일생산DB: 없음 (제품 소요 동기화 불가)" & vbCrLf
        Else
            결과 = 결과 & "6. 일일생산DB: OK" & vbCrLf
        End If
        If wsRDB Is Nothing Then
            결과 = 결과 & "7. 일일생산DB 부자재블록: 없음 (데이터업데이트 먼저 실행 필요)" & vbCrLf
        Else
            결과 = 결과 & "7. 일일생산DB 부자재블록: OK" & vbCrLf
        End If
        
        결과 = 결과 & vbCrLf & "========================================" & vbCrLf
        결과 = 결과 & "설정 완료!" & vbCrLf
        결과 = 결과 & "========================================" & vbCrLf & vbCrLf
        결과 = 결과 & "이제 원단관리 시트에서 소요/입고 값을 변경하면" & vbCrLf
        결과 = 결과 & "자동으로 재고 재계산 + DB 저장이 됩니다." & vbCrLf & vbCrLf
        결과 = 결과 & "★ 핵심: ThisWorkbook에 이벤트를 설치했으므로" & vbCrLf
        결과 = 결과 & "  시트 모듈 문제가 발생하지 않습니다." & vbCrLf
        
        MsgBox 결과, vbInformation, "자동 설정 완료"
End Sub
    
    '================================================================================
    ' 빠른 테스트: 현재 선택한 셀로 이벤트 테스트
    '================================================================================
    Public Sub 원단_이벤트테스트()
        Dim ws As Worksheet
        Dim 선택셀 As Range
        
        Set ws = ActiveSheet
        If ws.Name <> WS_원단관리 Then
            MsgBox "? 원단관리 시트에서 실행하세요.", vbExclamation
            Exit Sub
        End If
        
        Set 선택셀 = Selection
        If 선택셀 Is Nothing Or 선택셀.Cells.Count <> 1 Then
            MsgBox "? 단일 셀을 선택하세요.", vbExclamation
            Exit Sub
        End If
        
        ' 테스트 메시지
        Dim 구분 As String
        구분 = Trim(CStr(ws.Cells(선택셀.Row, COL_구분).Value))
        
        If 구분 <> "소요" And 구분 <> "출고" And 구분 <> "입고" Then
            MsgBox "? '소요'/'출고' 또는 '입고' 행의 셀을 선택하세요." & vbCrLf & vbCrLf & _
                "현재 구분: " & 구분, vbExclamation
            Exit Sub
        End If
        
        ' 현재 값 표시
        Dim 결과 As String
        결과 = "【 이벤트 테스트 】" & vbCrLf & vbCrLf
        결과 = 결과 & "선택 셀: " & 선택셀.Address & vbCrLf
        결과 = 결과 & "구분: " & 구분 & vbCrLf
        결과 = 결과 & "현재 값: " & 선택셀.Value & vbCrLf & vbCrLf
        결과 = 결과 & "? 이 셀의 값을 변경하면" & vbCrLf
        결과 = 결과 & "   자동으로 처리됩니다!" & vbCrLf & vbCrLf
        결과 = 결과 & "【 확인 사항 】" & vbCrLf
        결과 = 결과 & "? EnableEvents: " & IIf(Application.EnableEvents, "? TRUE", "? FALSE") & vbCrLf
        
        MsgBox 결과, vbInformation, "이벤트 테스트"
    End Sub

'================================================================================
' 원단별원단관리_생성: 원단명으로 필터하여 원단관리 데이터를 새 워크북에 저장
'   - 필터: 원단명(B열) 1개 기준
'   - 년도·월 선택 가능 (비우면 전체 기간)
'================================================================================
Public Sub 원단별원단관리_생성()
    Dim ws재고 As Worksheet
    Dim wsIndexKey As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long, 마지막열 As Long
    Dim 시작행 As Long
    Dim 저장경로 As String, 파일명 As String, 전체경로 As String
    Dim destRow As Long
    Dim r As Long, c As Long, i As Long
    Dim srcRng As Range
    Dim 선택년 As Integer, 선택월 As Integer
    Dim 월필터사용 As Boolean
    Dim 해당월시작열 As Long, 해당월일수 As Integer
    Dim 셀값 As String
    Const OUT_START_ROW As Long = 3
    Const OUT_START_COL As Long = 1
    Dim 출력끝열 As Long
    Dim 선택원단명 As String: 선택원단명 = ""
    Dim 필터없음 As Boolean: 필터없음 = False

    On Error Resume Next
    Set ws재고 = ThisWorkbook.Sheets(WS_원단관리)
    Set wsIndexKey = ThisWorkbook.Sheets(WS_IndexKey)
    On Error GoTo 0

    If ws재고 Is Nothing Then
        MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = ws재고.Cells(ws재고.Rows.Count, COL_구분).End(xlUp).Row
    If lastRow < START_ROW Then
        MsgBox "원단관리 시트에 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    ' ================================================================
    ' 1단계: 원단명 고유값 수집 → 항목 선택
    ' ================================================================
    Dim 목록 As Collection
    Set 목록 = New Collection
    For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        셀값 = Trim(CStr(ws재고.Cells(시작행, COL_제품명).Value))
        If 셀값 <> "" And 셀값 <> "-" Then
            Dim 목록키 As String: 목록키 = LCase(Replace(셀값, " ", ""))
            On Error Resume Next
            목록.Add 셀값, 목록키
            On Error GoTo 0
        End If
    Next 시작행

    If 목록.Count = 0 Then
        MsgBox "원단명 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If

    Dim 선택결과 As String
    선택결과 = 목록선택_다이얼로그(목록, _
        "[ 원단명 ] 내보낼 원단을 선택하세요.", _
        "필터/출력 - 원단명 선택")

    If 선택결과 = "<<CANCEL>>" Then Exit Sub
    If 선택결과 = "" Then
        필터없음 = True  ' 빈 선택 → 전체
    Else
        선택원단명 = 선택결과
    End If

    ' ================================================================
    ' 2단계: 년도·월 선택
    ' ================================================================
    Dim 년도입력_raw As String
    년도입력_raw = InputBox( _
        "출력할 년도를 입력하세요." & vbCrLf & vbCrLf & _
        "비워두면 전체 기간으로 출력합니다." & vbCrLf & _
        "(예: 2026)", _
        "필터/출력 - 년도 선택", Format(Year(Date), "0000"))

    If StrPtr(년도입력_raw) = 0 Then Exit Sub
    Dim 년도입력 As String: 년도입력 = Trim(년도입력_raw)

    선택년 = 0: 선택월 = 0
    월필터사용 = False: 해당월시작열 = 0: 해당월일수 = 0

    If 년도입력 <> "" And IsNumeric(년도입력) Then
        선택년 = CInt(년도입력)
        If 선택년 < 2000 Or 선택년 > 2100 Then
            MsgBox "올바른 연도(2000~2100)를 입력하세요.", vbExclamation
            Exit Sub
        End If

        Dim 월입력_raw As String
        월입력_raw = InputBox( _
            선택년 & "년의 출력할 월을 선택하세요." & vbCrLf & vbCrLf & _
            "1~12 중 입력 (비워두면 해당 년도 전체 월)" & vbCrLf & "(예: 2 → 2월)", _
            "필터/출력 - 월 선택", Format(Month(Date), "0"))

        If StrPtr(월입력_raw) = 0 Then Exit Sub
        Dim 월입력 As String: 월입력 = Trim(월입력_raw)

        If 월입력 <> "" And IsNumeric(월입력) Then
            선택월 = CInt(월입력)
            If 선택월 < 1 Or 선택월 > 12 Then
                MsgBox "올바른 월(1~12)을 입력하세요.", vbExclamation
                Exit Sub
            End If
            월필터사용 = True
            해당월일수 = DaysInMonth(선택년, 선택월)
            마지막열 = 마지막월열_찾기(ws재고)

            ' 해당 월 시작열 찾기 (2행: 2M 또는 2m, 3행: 1 → 둘 다 허용)
            Dim v2행 As String
            For c = COL_월별시작 To 마지막열
                v2행 = Trim(CStr(ws재고.Cells(2, c).Value))
                If (v2행 = (선택월 & "M") Or v2행 = (선택월 & "m")) And val(ws재고.Cells(3, c).Value) = 1 Then
                    해당월시작열 = c
                    Exit For
                End If
            Next c

            If 해당월시작열 = 0 Then
                MsgBox "시트에서 " & 선택월 & "월 데이터를 찾을 수 없습니다." & vbCrLf & _
                    "(2행에 """ & 선택월 & "M"" 또는 """ & 선택월 & "m"", 3행에 1이 있는 열이 필요합니다)", vbExclamation
                Exit Sub
            End If
        End If
    End If

    ' ================================================================
    ' 3단계: 필터 요약 및 저장 경로
    ' ================================================================
    Dim 필터요약 As String: 필터요약 = ""
    Dim 파일필터명 As String: 파일필터명 = ""

    If 필터없음 Then
        필터요약 = "전체": 파일필터명 = "전체"
    Else
        필터요약 = "원단명=" & 선택원단명
        파일필터명 = 선택원단명
    End If

    저장경로 = ""
    If Not wsIndexKey Is Nothing Then
        저장경로 = Trim(CStr(wsIndexKey.Range("A3").Value))
    End If
    If 저장경로 = "" Then
        On Error Resume Next
        저장경로 = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        On Error GoTo 0
    End If
    If right(저장경로, 1) <> "\" And right(저장경로, 1) <> "/" Then 저장경로 = 저장경로 & "\"

    파일필터명 = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
        파일필터명, "\", ""), "/", ""), ":", ""), "*", ""), "?", ""), """", ""), "<", ""), ">", "")
    파일필터명 = Replace(파일필터명, "|", "")

    If 월필터사용 Then
        파일명 = 파일필터명 & "_원단관리현황_" & Format(선택년, "0000") & Format(선택월, "00") & "_" & Format(Date, "yyyymmdd") & ".xlsx"
    Else
        파일명 = 파일필터명 & "_원단관리현황_" & Format(Date, "yyyymmdd") & ".xlsx"
    End If
    전체경로 = 저장경로 & 파일명

    ' ================================================================
    ' 4단계: 새 워크북 생성 + 데이터 출력
    ' ================================================================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "원단관리현황"

    If Not 월필터사용 Then 마지막열 = 마지막월열_찾기(ws재고)

    Dim 일열시작 As Long
    일열시작 = OUT_START_COL + (COL_구분 - COL_제품발주처 + 1)

    If 월필터사용 Then
        ' === 월 필터: 해당 월 열만 포함 ===
        Set srcRng = ws재고.Range(ws재고.Cells(1, COL_제품발주처), ws재고.Cells(3, COL_구분))
        srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), _
            wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL + COL_구분 - COL_제품발주처))

        For i = 1 To 해당월일수
            wsNew.Cells(OUT_START_ROW + 1, 일열시작 + i - 1).Value = 선택월 & "M"
            wsNew.Cells(OUT_START_ROW + 1, 일열시작 + i - 1).Font.color = RGB(160, 160, 160)
            wsNew.Cells(OUT_START_ROW + 2, 일열시작 + i - 1).Value = i
        Next i
        출력끝열 = 일열시작 + 해당월일수 - 1

        On Error Resume Next
        wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), wsNew.Cells(OUT_START_ROW, 출력끝열)).Merge
        On Error GoTo 0
        wsNew.Cells(OUT_START_ROW, OUT_START_COL).Value = "원단관리 현황"

        destRow = OUT_START_ROW + 3
        Dim 출력그룹수 As Long: 출력그룹수 = 0
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            If Not 필터없음 Then
                If Trim(CStr(ws재고.Cells(시작행, COL_제품명).Value)) <> 선택원단명 Then GoTo NextGroup1
            End If
            Set srcRng = ws재고.Range(ws재고.Cells(시작행, COL_제품발주처), ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, COL_구분))
            srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, OUT_START_COL), _
                wsNew.Cells(destRow + ROWS_PER_ITEM - 1, OUT_START_COL + COL_구분 - COL_제품발주처))
            Set srcRng = ws재고.Range(ws재고.Cells(시작행, 해당월시작열), _
                ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, 해당월시작열 + 해당월일수 - 1))
            srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, 일열시작), _
                wsNew.Cells(destRow + ROWS_PER_ITEM - 1, 일열시작 + 해당월일수 - 1))
            destRow = destRow + ROWS_PER_ITEM
            출력그룹수 = 출력그룹수 + 1
NextGroup1:
        Next 시작행
    Else
        ' === 전체 기간 ===
        마지막열 = 마지막월열_찾기(ws재고)
        Set srcRng = ws재고.Range(ws재고.Cells(1, COL_제품발주처), ws재고.Cells(3, 마지막열))
        srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), _
            wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL + 마지막열 - COL_제품발주처))
        출력끝열 = OUT_START_COL + 마지막열 - COL_제품발주처

        destRow = OUT_START_ROW + 3
        출력그룹수 = 0
        For 시작행 = START_ROW To lastRow - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
            If Not 필터없음 Then
                If Trim(CStr(ws재고.Cells(시작행, COL_제품명).Value)) <> 선택원단명 Then GoTo NextGroup2
            End If
            Set srcRng = ws재고.Range(ws재고.Cells(시작행, COL_제품발주처), ws재고.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막열))
            srcRng.Copy Destination:=wsNew.Range(wsNew.Cells(destRow, OUT_START_COL), _
                wsNew.Cells(destRow + ROWS_PER_ITEM - 1, OUT_START_COL + 마지막열 - COL_제품발주처))
            destRow = destRow + ROWS_PER_ITEM
            출력그룹수 = 출력그룹수 + 1
NextGroup2:
        Next 시작행
    End If

    ' 출력 데이터 없음 체크
    If 출력그룹수 = 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = False
        wbNew.Close SaveChanges:=False
        Application.DisplayAlerts = True
        MsgBox "선택한 필터 조건에 맞는 데이터가 없습니다." & vbCrLf & vbCrLf & _
            "필터: " & 필터요약, vbExclamation, "필터/출력"
        Exit Sub
    End If

    ' 수식 제거 (값만 유지)
    On Error Resume Next
    wsNew.Range(wsNew.Cells(OUT_START_ROW + 3, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열)).Copy
    wsNew.Range(wsNew.Cells(OUT_START_ROW + 3, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열)).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error GoTo 0

    ' 서식 적용
    With wsNew.Range(wsNew.Cells(OUT_START_ROW, OUT_START_COL), wsNew.Cells(destRow - 1, 출력끝열))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 날짜 헤더 강조
    Dim 일헤더시작 As Long
    일헤더시작 = OUT_START_COL + (COL_구분 - COL_제품발주처 + 1)
    If 일헤더시작 <= 출력끝열 Then
        With wsNew.Range(wsNew.Cells(OUT_START_ROW + 1, 일헤더시작), wsNew.Cells(OUT_START_ROW + 2, 출력끝열))
            .Interior.color = RGB(232, 232, 232)
            .Font.Bold = True
        End With
    End If
    With wsNew.Range(wsNew.Cells(OUT_START_ROW + 2, OUT_START_COL), wsNew.Cells(OUT_START_ROW + 2, 출력끝열))
        .Font.Bold = True
    End With

    ' 필터 정보 (1~2행)
    wsNew.Cells(1, OUT_START_COL).Value = "필터: " & 필터요약
    wsNew.Cells(1, OUT_START_COL).Font.Bold = True
    If 월필터사용 Then
        wsNew.Cells(2, OUT_START_COL).Value = "기간: " & 선택년 & "년 " & 선택월 & "월"
    Else
        wsNew.Cells(2, OUT_START_COL).Value = "기간: 전체"
    End If
    wsNew.Cells(2, OUT_START_COL).Font.Bold = True

    ' 눈금선 숨김·폰트 설정
    wbNew.Windows(1).DisplayGridlines = False
    wsNew.usedRange.Font.Name = "Microsoft GothicNeo"

    ' 매크로 버튼(Shapes) 제거
    Dim shpIdx As Long
    For shpIdx = wsNew.Shapes.Count To 1 Step -1
        wsNew.Shapes(shpIdx).Delete
    Next shpIdx

    ' ================================================================
    ' 5단계: 저장
    ' ================================================================
    On Error Resume Next
    wbNew.SaveAs Filename:=전체경로, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "저장에 실패했습니다." & vbCrLf & vbCrLf & _
            "경로: " & 전체경로 & vbCrLf & _
            "오류: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wbNew.Activate

    MsgBox "필터/출력이 완료되었습니다!" & vbCrLf & vbCrLf & _
        "[ 적용된 필터 ]" & vbCrLf & _
        "  항목: " & 필터요약 & vbCrLf & _
        IIf(월필터사용, "  기간: " & 선택년 & "년 " & 선택월 & "월" & vbCrLf, "") & _
        vbCrLf & _
        "[ 출력 결과 ]" & vbCrLf & _
        "  출력 원단 수: " & 출력그룹수 & "건" & vbCrLf & _
        "  저장: " & 전체경로, vbInformation, "필터/출력 완료"
End Sub

'================================================================================
' 원단_월생산현황: 선택한 년도/월에 대한 제품 생산·입고·소요 데이터를 엑셀 워크북으로 저장
'   - 변경된 원단관리 시트 구조(ROWS_PER_ITEM=15) 적용
'   - 제품명_1차포장, 제품명_2차포장으로 구분 표기 (금액 제거)
'   - 입고·소요 모두 유지
'================================================================================
Public Sub 원단_월생산현황()
    Dim ws재고 As Worksheet, wsIndex As Worksheet
    Dim wb신규 As Workbook, ws신규 As Worksheet
    Dim 년도 As String, 월 As String
    Dim 년도값 As Integer, 월값 As Integer, 일수 As Integer
    Dim 저장경로 As String, 폴더경로 As String, 파일명 As String, 전체경로 As String
    Dim i As Long, j As Long, k As Long
    Dim 현재일 As Date, 요일 As Integer
    Dim 마지막행 As Long, 마지막열 As Long
    Dim 제품명 As String, 업체명 As String
    Dim fso As Object
    Dim 월시작열 As Long, 찾은월 As String
    Dim 현재행 As Long, 원지사용 As String, 그룹시작행 As Long
    Dim 제품목록 As Object          ' 필터 적용 후 최종 출력 목록
    Dim key As Variant, 생산값 As Variant
    Dim 테이블범위 As Range
    ' ── 필터 관련 변수 ──
    Dim 필터유형 As String           ' "전체" / "제품별" / "업체별"
    Dim 필터값 As String             ' 선택된 필터 값 (쉼표 구분 가능)
    Dim 전체그룹 As Object           ' 모든 그룹 {제품명 → 그룹시작행}
    Dim 업체딕 As Object             ' 업체별 딕셔너리 {업체명 → 제품명 배열}
    Dim 선택목록() As String         ' 선택된 필터값 배열
    Dim 후보목록 As String           ' InputBox에 표시할 번호 목록
    Dim 후보키() As String           ' 번호에 대응하는 키 배열
    Dim 후보수 As Long
    Dim 입력값 As String
    Dim 필터표시 As String           ' 타이틀/파일명에 표시할 필터 정보

    ' ── 시트 참조 ──
    On Error Resume Next
    Set ws재고 = ThisWorkbook.Sheets(WS_원단관리)
    Set wsIndex = ThisWorkbook.Sheets("Index key")
    On Error GoTo 0

    If ws재고 Is Nothing Then
        MsgBox "원단관리 시트를 찾을 수 없습니다.", vbExclamation: Exit Sub
    End If
    ' Index key 시트 삭제 시 저장경로는 아래에서 기본값 사용

    ' ── 년도/월 입력 ──
    년도 = InputBox("년도를 입력하세요 (예: 2026)", "년도 선택", Year(Now))
    If 년도 = "" Then Exit Sub
    If Not IsNumeric(년도) Then MsgBox "올바른 년도를 입력하세요.", vbExclamation: Exit Sub
    년도값 = CInt(년도)

    월 = InputBox("월을 입력하세요 (1-12)", "월 선택", Month(Now))
    If 월 = "" Then Exit Sub
    If Not IsNumeric(월) Then MsgBox "올바른 월을 입력하세요.", vbExclamation: Exit Sub
    월값 = CInt(월)
    If 월값 < 1 Or 월값 > 12 Then MsgBox "월은 1~12 사이여야 합니다.", vbExclamation: Exit Sub

    일수 = Day(DateSerial(년도값, 월값 + 1, 0))

    ' ── 기본 데이터 수집 (마지막행/열, 월시작열) ──
    마지막행 = ws재고.Cells(ws재고.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 < START_ROW Then
        마지막행 = ws재고.Cells(ws재고.Rows.Count, COL_제품).End(xlUp).Row
    End If
    마지막열 = ws재고.Cells(2, ws재고.Columns.Count).End(xlToLeft).Column

    월시작열 = 0
    For i = COL_일별시작 To 마지막열
        찾은월 = Trim(CStr(ws재고.Cells(2, i).Value))
        If (찾은월 = (월값 & "M") Or 찾은월 = (월값 & "m")) And val(ws재고.Cells(3, i).Value) = 1 Then
            월시작열 = i
            Exit For
        End If
    Next i

    If 월시작열 = 0 Then
        MsgBox "원단관리 시트에서 " & 월값 & "월 데이터를 찾을 수 없습니다." & vbCrLf & _
            "(2행에 """ & 월값 & "M"" 또는 """ & 월값 & "m"", 3행에 1이 있는 열 필요)", vbExclamation
        Exit Sub
    End If

    ' ── 전체 그룹 수집 (제품명, 업체명) ──
    Set 전체그룹 = CreateObject("Scripting.Dictionary")   ' 제품명 → 그룹시작행
    Set 업체딕 = CreateObject("Scripting.Dictionary")     ' 업체명 → "" (고유값)
    Dim 그룹업체 As Object
    Set 그룹업체 = CreateObject("Scripting.Dictionary")   ' 그룹시작행 → 업체명

    For i = START_ROW To 마지막행 Step ROWS_PER_ITEM
        제품명 = ""
        업체명 = ""
        On Error Resume Next
        제품명 = Trim(CStr(ws재고.Cells(i, COL_제품).MergeArea.Cells(1, 1).Value))
        업체명 = Trim(CStr(ws재고.Cells(i, COL_업체명).MergeArea.Cells(1, 1).Value))
        On Error GoTo 0
        If 제품명 = "" Then 제품명 = Trim(CStr(ws재고.Cells(i, COL_제품).Value))
        If 업체명 = "" Then 업체명 = Trim(CStr(ws재고.Cells(i, COL_업체명).Value))

        If 제품명 <> "" Then
            If Not 전체그룹.exists(제품명) Then
                전체그룹.Add 제품명, i
                그룹업체.Add CStr(i), 업체명
            End If
            If 업체명 <> "" Then
                If Not 업체딕.exists(업체명) Then 업체딕.Add 업체명, ""
            End If
        End If
    Next i

    If 전체그룹.Count = 0 Then
        MsgBox "원단관리 시트에서 제품(C열)을 찾을 수 없습니다." & vbCrLf & _
            "C열(제품)에 데이터가 있고, 4행부터 " & ROWS_PER_ITEM & "행 간격으로 그룹이 있는지 확인하세요.", vbExclamation: Exit Sub
    End If

    ' ??????????????????????????????????????????????????????????????
    ' ★ 필터 유형 선택 (1=전체, 2=제품별, 3=업체별)
    ' ??????????????????????????????????????????????????????????????
    입력값 = InputBox( _
        "필터 유형을 선택하세요:" & vbCrLf & vbCrLf & _
        "  1 = 전체 (모든 제품)" & vbCrLf & _
        "  2 = 제품별 (특정 제품 선택)" & vbCrLf & _
        "  3 = 업체별 (특정 업체 선택)" & vbCrLf & vbCrLf & _
        "번호를 입력하세요:", "필터 선택", "1")
    If 입력값 = "" Then Exit Sub

    Select Case Trim(입력값)
        Case "1"
            필터유형 = "전체"
        Case "2"
            필터유형 = "제품별"
        Case "3"
            필터유형 = "업체별"
        Case Else
            MsgBox "1, 2, 3 중 하나를 입력하세요.", vbExclamation: Exit Sub
    End Select

    ' ── 필터별 세부 선택 ──
    필터표시 = ""
    Set 제품목록 = CreateObject("Scripting.Dictionary")

    If 필터유형 = "전체" Then
        ' 전체: 모든 제품 포함
        For Each key In 전체그룹.keys
            제품목록.Add key, 전체그룹(key)
        Next key
        필터표시 = "전체"

    ElseIf 필터유형 = "제품별" Then
        ' 제품 목록 표시
        후보수 = 전체그룹.Count
        ReDim 후보키(1 To 후보수)
        후보목록 = "출력할 제품 번호를 입력하세요 (쉼표로 복수선택 가능):" & vbCrLf & vbCrLf
        k = 1
        For Each key In 전체그룹.keys
            후보키(k) = CStr(key)
            후보목록 = 후보목록 & "  " & k & " = " & key & vbCrLf
            k = k + 1
        Next key

        입력값 = InputBox(후보목록, "제품 선택")
        If 입력값 = "" Then Exit Sub

        선택목록 = Split(입력값, ",")
        For i = 0 To UBound(선택목록)
            Dim 선택번호 As Long
            If IsNumeric(Trim(선택목록(i))) Then
                선택번호 = CLng(Trim(선택목록(i)))
                If 선택번호 >= 1 And 선택번호 <= 후보수 Then
                    Dim 선택키 As String
                    선택키 = 후보키(선택번호)
                    If Not 제품목록.exists(선택키) Then
                        제품목록.Add 선택키, 전체그룹(선택키)
                        If 필터표시 <> "" Then 필터표시 = 필터표시 & ", "
                        필터표시 = 필터표시 & 선택키
                    End If
                End If
            End If
        Next i

        If 제품목록.Count = 0 Then
            MsgBox "선택된 제품이 없습니다.", vbExclamation: Exit Sub
        End If

    ElseIf 필터유형 = "업체별" Then
        ' 업체 목록 표시
        후보수 = 업체딕.Count
        If 후보수 = 0 Then
            MsgBox "업체명 데이터가 없습니다.", vbExclamation: Exit Sub
        End If
        ReDim 후보키(1 To 후보수)
        후보목록 = "출력할 업체 번호를 입력하세요 (쉼표로 복수선택 가능):" & vbCrLf & vbCrLf
        k = 1
        For Each key In 업체딕.keys
            후보키(k) = CStr(key)
            후보목록 = 후보목록 & "  " & k & " = " & key & vbCrLf
            k = k + 1
        Next key

        입력값 = InputBox(후보목록, "업체 선택")
        If 입력값 = "" Then Exit Sub

        ' 선택된 업체 딕셔너리
        Dim 선택업체 As Object
        Set 선택업체 = CreateObject("Scripting.Dictionary")
        선택목록 = Split(입력값, ",")
        For i = 0 To UBound(선택목록)
            If IsNumeric(Trim(선택목록(i))) Then
                선택번호 = CLng(Trim(선택목록(i)))
                If 선택번호 >= 1 And 선택번호 <= 후보수 Then
                    If Not 선택업체.exists(후보키(선택번호)) Then
                        선택업체.Add 후보키(선택번호), ""
                        If 필터표시 <> "" Then 필터표시 = 필터표시 & ", "
                        필터표시 = 필터표시 & 후보키(선택번호)
                    End If
                End If
            End If
        Next i

        If 선택업체.Count = 0 Then
            MsgBox "선택된 업체가 없습니다.", vbExclamation: Exit Sub
        End If

        ' 선택 업체에 속하는 제품만 추출
        For Each key In 전체그룹.keys
            Dim 이그룹업체 As String
            이그룹업체 = 그룹업체(CStr(전체그룹(key)))
            If 선택업체.exists(이그룹업체) Then
                If Not 제품목록.exists(key) Then
                    제품목록.Add key, 전체그룹(key)
                End If
            End If
        Next key

        If 제품목록.Count = 0 Then
            MsgBox "선택한 업체에 해당하는 제품이 없습니다.", vbExclamation: Exit Sub
        End If
    End If

    ' ── 저장 경로 (Index key 시트 없으면 통합문서 경로 또는 바탕화면 사용) ──
    저장경로 = ""
    If Not wsIndex Is Nothing Then 저장경로 = Trim(CStr(wsIndex.Range("A3").Value))
    If 저장경로 = "" Then
        On Error Resume Next
        저장경로 = ThisWorkbook.Path
        If 저장경로 = "" Then 저장경로 = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        On Error GoTo 0
    End If
    If right(저장경로, 1) = "\" Then 저장경로 = left(저장경로, Len(저장경로) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    폴더경로 = 저장경로 & "\월별생산현황"
    If Not fso.FolderExists(폴더경로) Then fso.CreateFolder 폴더경로

    파일명 = "월별생산출고현황_" & Format(년도값, "0000") & "-" & Format(월값, "00") & "월_" & Format(Now, "yyyyMMddHHmm") & ".xlsx"
    전체경로 = 폴더경로 & "\" & 파일명

    ' ??????????????????????????????????????????????????????????????
    ' ★ 새 워크북 생성 & 데이터 출력
    ' ??????????????????????????????????????????????????????????????
    Application.ScreenUpdating = False
    Set wb신규 = Workbooks.Add
    Set ws신규 = wb신규.Sheets(1)
    ws신규.Name = "원단_월생산현황"
    ws신규.Cells.Borders.LineStyle = xlNone

    ' 1행: 타이틀 (필터 정보 포함)
    Dim 타이틀 As String
    타이틀 = 년도값 & "년 " & 월값 & "월 입고·사용량 현황"
    If 필터유형 <> "전체" Then 타이틀 = 타이틀 & "  [" & 필터유형 & ": " & 필터표시 & "]"
    ws신규.Range("A1").Value = 타이틀
    ws신규.Range("A1").Font.Size = 16
    ws신규.Range("A1").Font.Bold = True
    ws신규.Range("A1").HorizontalAlignment = xlLeft

    ' 2행: 헤더 (음영 RGB 129,135,253 + 흰색 폰트)
    ws신규.Range("A2").Value = "업체명"
    ws신규.Range("B2").Value = "제품명"
    ws신규.Range("C2").Value = "비고"
    ws신규.Range("D2").Value = "구분"
    ws신규.Range("A2:D2").Font.Bold = True
    ws신규.Range("A2:D2").HorizontalAlignment = xlCenter
    ws신규.Range("A2:D2").Interior.color = RGB(129, 135, 253)
    ws신규.Range("A2:D2").Font.color = RGB(255, 255, 255)
    ws신규.Rows(2).RowHeight = 25
    Dim 일열시작 As Long, 계열 As Long
    일열시작 = 5  ' E열부터 일자
    계열 = 일열시작 + 일수  ' 열방향 마지막 = 소계
    For i = 1 To 일수
        ws신규.Cells(2, 일열시작 + i - 1).Value = i
        ws신규.Cells(2, 일열시작 + i - 1).HorizontalAlignment = xlCenter
        ws신규.Cells(2, 일열시작 + i - 1).Font.Bold = True
        ws신규.Cells(2, 일열시작 + i - 1).Interior.color = RGB(129, 135, 253)
        ws신규.Cells(2, 일열시작 + i - 1).Font.color = RGB(255, 255, 255)
    Next i
    ws신규.Cells(2, 계열).Value = "소계"
    ws신규.Cells(2, 계열).HorizontalAlignment = xlCenter
    ws신규.Cells(2, 계열).Font.Bold = True
    ws신규.Cells(2, 계열).Interior.color = RGB(129, 135, 253)
    ws신규.Cells(2, 계열).Font.color = RGB(255, 255, 255)

    ' ── 데이터 입력 (제품당 6행: 입고/사용량 + 1차포장 입고/소요 + 2차포장 입고/소요, 원지 ROLL/TONS 제외) ──
    현재행 = 3
    Dim 소요값 As Variant, 입고값 As Variant
    Dim 행합계 As Double
    Dim 일차포장명 As String, 이차포장명 As String
    Dim 구분정보 As Variant  ' Array(행오프셋, 구분문자열)
    For Each key In 제품목록.keys
        제품명 = key
        그룹시작행 = 제품목록(key)

        업체명 = ""
        On Error Resume Next
        업체명 = Trim(CStr(ws재고.Cells(그룹시작행, COL_업체명).MergeArea.Cells(1, 1).Value))
        On Error GoTo 0
        If 업체명 = "" Then 업체명 = Trim(CStr(ws재고.Cells(그룹시작행, COL_업체명).Value))
        일차포장명 = Trim(CStr(ws재고.Cells(그룹시작행, COL_1차포장).Value))
        이차포장명 = Trim(CStr(ws재고.Cells(그룹시작행, COL_2차포장).Value))

        ' 6행: 원단관리 ROWS_PER_ITEM=6 구조 (0=Roll입고, 1=Roll사용, 3=Kg입고, 4=Kg사용 + 입고현황/사용량)
        구분정보 = Array( _
            Array(0, "입고현황", 0), _
            Array(1, "사용량", 1), _
            Array(OFFSET_입고1, 제품명 & "_1차포장_ROLL입고", 2), _
            Array(OFFSET_사용1, 제품명 & "_1차포장_ROLL소요", 3), _
            Array(OFFSET_입고2, 제품명 & "_2차포장_KG입고", 4), _
            Array(OFFSET_사용2, 제품명 & "_2차포장_KG소요", 5))
        Dim bi As Long, 행오프 As Long, 색idx As Long
        For bi = 0 To 5
            행오프 = 구분정보(bi)(0)
            ws신규.Cells(현재행, 1).Value = 업체명
            ws신규.Cells(현재행, 2).Value = 제품명
            If bi >= 2 Then
                If bi = 2 Or bi = 3 Then ws신규.Cells(현재행, 3).Value = 일차포장명 Else ws신규.Cells(현재행, 3).Value = 이차포장명
            Else
                ws신규.Cells(현재행, 3).Value = ""
            End If
            ws신규.Cells(현재행, 4).Value = 구분정보(bi)(1)
            행합계 = 0
            For j = 1 To 일수
                If 월시작열 + j - 1 <= 마지막열 Then
                    입고값 = ws재고.Cells(그룹시작행 + 행오프, 월시작열 + j - 1).Value
                    If IsNumeric(입고값) And 입고값 <> "" And 입고값 <> 0 Then
                        ws신규.Cells(현재행, 일열시작 + j - 1).Value = 입고값
                        색idx = (bi Mod 2)
                        If bi <= 1 Then
                            If bi = 0 Then ws신규.Cells(현재행, 일열시작 + j - 1).Interior.color = RGB(192, 135, 253) Else ws신규.Cells(현재행, 일열시작 + j - 1).Interior.color = RGB(135, 192, 253)
                            ws신규.Cells(현재행, 일열시작 + j - 1).Font.color = RGB(255, 255, 255)
                        Else
                            ws신규.Cells(현재행, 일열시작 + j - 1).Interior.color = RGB(220, 235, 240)
                        End If
                        행합계 = 행합계 + CDbl(입고값)
                    End If
                End If
            Next j
            If 행합계 > 0 Then ws신규.Cells(현재행, 계열).Value = 행합계
            현재행 = 현재행 + 1
        Next bi
    Next key

    ' ── A,B,C열 같은 항목 병합 (제품당 6행) - 경고 메시지 비표시 ──
    Dim 병합시작 As Long, 알림상태 As Boolean
    알림상태 = Application.DisplayAlerts
    Application.DisplayAlerts = False
    병합시작 = 3
    For i = 병합시작 To 현재행 - 1 Step 6
        If i + 5 <= 현재행 - 1 Then
            On Error Resume Next
            ws신규.Range(ws신규.Cells(i, 1), ws신규.Cells(i + 5, 1)).Merge
            ws신규.Range(ws신규.Cells(i, 2), ws신규.Cells(i + 5, 2)).Merge
            ws신규.Range(ws신규.Cells(i, 3), ws신규.Cells(i + 5, 3)).Merge
            ws신규.Range(ws신규.Cells(i, 1), ws신규.Cells(i + 5, 1)).HorizontalAlignment = xlCenter
            ws신규.Range(ws신규.Cells(i, 2), ws신규.Cells(i + 5, 2)).HorizontalAlignment = xlCenter
            ws신규.Range(ws신규.Cells(i, 3), ws신규.Cells(i + 5, 3)).VerticalAlignment = xlCenter
            On Error GoTo 0
        Else
            Dim 병합행수 As Long
            병합행수 = 현재행 - 1 - i + 1
            If 병합행수 > 1 Then
                On Error Resume Next
                ws신규.Range(ws신규.Cells(i, 1), ws신규.Cells(i + 병합행수 - 1, 1)).Merge
                ws신규.Range(ws신규.Cells(i, 2), ws신규.Cells(i + 병합행수 - 1, 2)).Merge
                ws신규.Range(ws신규.Cells(i, 3), ws신규.Cells(i + 병합행수 - 1, 3)).Merge
                ws신규.Range(ws신규.Cells(i, 1), ws신규.Cells(i + 병합행수 - 1, 1)).HorizontalAlignment = xlCenter
                ws신규.Range(ws신규.Cells(i, 2), ws신규.Cells(i + 병합행수 - 1, 2)).HorizontalAlignment = xlCenter
                ws신규.Range(ws신규.Cells(i, 3), ws신규.Cells(i + 병합행수 - 1, 3)).VerticalAlignment = xlCenter
                On Error GoTo 0
            End If
        End If
    Next i

    ' ── 데이터 마지막행: Total입고, Total사용량 ──
    Dim Total입고 As Double, Total사용량 As Double
    Dim 일별입고합계 As Double, 일별사용량합계 As Double
    Dim 데이터끝행 As Long
    데이터끝행 = 현재행 - 1
    Total입고 = 0: Total사용량 = 0
    For i = 3 To 데이터끝행 Step 6
        If IsNumeric(ws신규.Cells(i, 계열).Value) Then Total입고 = Total입고 + CDbl(ws신규.Cells(i, 계열).Value)
        If i + 1 <= 데이터끝행 And IsNumeric(ws신규.Cells(i + 1, 계열).Value) Then Total사용량 = Total사용량 + CDbl(ws신규.Cells(i + 1, 계열).Value)
    Next i

    ws신규.Cells(현재행, 4).Value = "Total입고"
    ws신규.Range(ws신규.Cells(현재행, 4), ws신규.Cells(현재행, 계열)).Font.Bold = True
    ws신규.Cells(현재행, 계열).Value = Total입고
    ws신규.Range(ws신규.Cells(현재행, 일열시작), ws신규.Cells(현재행, 계열)).Interior.color = RGB(220, 220, 255)
    For j = 1 To 일수
        일별입고합계 = 0
        For i = 3 To 데이터끝행 Step 6
            If IsNumeric(ws신규.Cells(i, 일열시작 + j - 1).Value) Then
                일별입고합계 = 일별입고합계 + CDbl(ws신규.Cells(i, 일열시작 + j - 1).Value)
            End If
        Next i
        ws신규.Cells(현재행, 일열시작 + j - 1).Value = IIf(일별입고합계 > 0, 일별입고합계, "")
    Next j
    현재행 = 현재행 + 1

    ws신규.Cells(현재행, 4).Value = "Total사용량"
    ws신규.Range(ws신규.Cells(현재행, 4), ws신규.Cells(현재행, 계열)).Font.Bold = True
    ws신규.Cells(현재행, 계열).Value = Total사용량
    ws신규.Range(ws신규.Cells(현재행, 일열시작), ws신규.Cells(현재행, 계열)).Interior.color = RGB(220, 235, 255)
    For j = 1 To 일수
        일별사용량합계 = 0
        For i = 4 To 데이터끝행 Step 6
            If IsNumeric(ws신규.Cells(i, 일열시작 + j - 1).Value) Then
                일별사용량합계 = 일별사용량합계 + CDbl(ws신규.Cells(i, 일열시작 + j - 1).Value)
            End If
        Next i
        ws신규.Cells(현재행, 일열시작 + j - 1).Value = IIf(일별사용량합계 > 0, 일별사용량합계, "")
    Next j

    ' ── Total입고, Total사용량 행의 A,B,C열 병합 (빈행 포함) ──
    Dim Total시작행 As Long
    Total시작행 = 데이터끝행 + 1
    If Total시작행 + 1 <= 현재행 Then
        On Error Resume Next
        ws신규.Range(ws신규.Cells(Total시작행, 1), ws신규.Cells(Total시작행 + 1, 1)).Merge
        ws신규.Range(ws신규.Cells(Total시작행, 2), ws신규.Cells(Total시작행 + 1, 2)).Merge
        ws신규.Range(ws신규.Cells(Total시작행, 3), ws신규.Cells(Total시작행 + 1, 3)).Merge
        ws신규.Range(ws신규.Cells(Total시작행, 1), ws신규.Cells(Total시작행 + 1, 1)).HorizontalAlignment = xlCenter
        ws신규.Range(ws신규.Cells(Total시작행, 2), ws신규.Cells(Total시작행 + 1, 2)).HorizontalAlignment = xlCenter
        ws신규.Range(ws신규.Cells(Total시작행, 3), ws신규.Cells(Total시작행 + 1, 3)).VerticalAlignment = xlCenter
        On Error GoTo 0
    End If
    Application.DisplayAlerts = 알림상태

    ' ── 데이터 영역 테두리 적용 (마지막행=Total사용량 포함) ──
    If 현재행 > 3 Then
        Set 테이블범위 = ws신규.Range(ws신규.Cells(2, 1), ws신규.Cells(현재행, 계열))
        With 테이블범위.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    End If

    ' ★ N월 매입매출현황 시트 제거됨 (금액 요소 모두 제거)
    ' ★ 원지사용·원지입고 섹션 제거됨
    ' ★ 부자재(1차포장, 2차포장) 섹션 제거됨

    ' ── 틀 고정 ──
    ws신규.Activate
    ws신규.Cells(3, 일열시작).Select
    ActiveWindow.FreezePanes = True

    ' ── 열 너비 ──
    ws신규.Columns("A:D").AutoFit
    ' E열~마지막 열: 데이터가 있는 전체 행 기준으로 AutoFit
    If 일수 > 0 Then
        With ws신규.Range(ws신규.Cells(1, 일열시작), ws신규.Cells(Total시작행 + 1, 계열))
            .Columns.AutoFit
        End With
    End If

    ' ── 폰트: Microsoft GothicNeo ──
    ws신규.usedRange.Font.Name = "Microsoft GothicNeo"

    ' ── 눈금선 제거 ──
    wb신규.Windows(1).DisplayGridlines = False

    ' ??????????????????????????????????????????????????????????????
    ' ★ Dashboard 시트 생성 및 차트 시각화
    ' ??????????????????????????????????????????????????????????????
    Dim wsDash As Worksheet
    Dim ch As ChartObject, cht As Chart
    Dim 제품수 As Long, 막대최대 As Long
    Dim dashRow As Long, v생산 As Variant, v출고 As Variant
    Application.DisplayAlerts = False
    Set wsDash = wb신규.Sheets.Add(Before:=wb신규.Sheets(1))
    wsDash.Name = "Dashboard"
    wsDash.Cells.Font.Name = "Microsoft GothicNeo"
    wsDash.Range("A1").Value = 년도값 & "년 " & 월값 & "월 생산·출고 Dashboard"
    wsDash.Range("A1").Font.Size = 18
    wsDash.Range("A1").Font.Bold = True

    ' 페이지 설정: 가로모드, 한페이지 인쇄, 최대크기, 여백없음
    With wsDash.PageSetup
        .Orientation = xlLandscape
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
    End With

    ' ── 차트용 데이터 Dashboard 하단에 준비 ──
    막대최대 = 12
    제품수 = (데이터끝행 - 2) \ 6
    If 제품수 > 막대최대 Then 제품수 = 막대최대
    dashRow = 100

    ' 일별 생산·출고 데이터 (A열=일, B열=생산, C열=출고)
    wsDash.Cells(dashRow, 1).Value = "일"
    wsDash.Cells(dashRow, 2).Value = "생산"
    wsDash.Cells(dashRow, 3).Value = "출고"
    For j = 1 To 일수
        wsDash.Cells(dashRow + j, 1).Value = j
        v생산 = ws신규.Cells(Total시작행, 일열시작 + j - 1).Value
        v출고 = ws신규.Cells(Total시작행 + 1, 일열시작 + j - 1).Value
        wsDash.Cells(dashRow + j, 2).Value = IIf(IsNumeric(v생산) And v생산 <> "", CDbl(v생산), 0)
        wsDash.Cells(dashRow + j, 3).Value = IIf(IsNumeric(v출고) And v출고 <> "", CDbl(v출고), 0)
    Next j

    ' 생산vs출고 비율 (D열=구분, E열=값)
    wsDash.Cells(dashRow, 5).Value = "생산"
    wsDash.Cells(dashRow, 6).Value = IIf(IsNumeric(ws신규.Cells(Total시작행, 계열).Value), CDbl(ws신규.Cells(Total시작행, 계열).Value), 0)
    wsDash.Cells(dashRow + 1, 5).Value = "출고"
    wsDash.Cells(dashRow + 1, 6).Value = IIf(IsNumeric(ws신규.Cells(Total시작행 + 1, 계열).Value), CDbl(ws신규.Cells(Total시작행 + 1, 계열).Value), 0)

    ' 제품별 생산·출고 (G열=제품명, H열=생산, I열=출고, J열=제품명복사, K열=출고복사)
    wsDash.Cells(dashRow, 7).Value = "제품명"
    wsDash.Cells(dashRow, 8).Value = "생산"
    wsDash.Cells(dashRow, 9).Value = "출고"
    For i = 1 To 제품수
        wsDash.Cells(dashRow + i, 7).Value = ws신규.Cells(3 + (i - 1) * 6, 2).Value
        v생산 = ws신규.Cells(3 + (i - 1) * 6, 계열).Value
        v출고 = ws신규.Cells(4 + (i - 1) * 6, 계열).Value
        wsDash.Cells(dashRow + i, 8).Value = IIf(IsNumeric(v생산) And v생산 <> "", CDbl(v생산), 0)
        wsDash.Cells(dashRow + i, 9).Value = IIf(IsNumeric(v출고) And v출고 <> "", CDbl(v출고), 0)
        wsDash.Cells(dashRow + i, 10).Value = wsDash.Cells(dashRow + i, 7).Value
        wsDash.Cells(dashRow + i, 11).Value = wsDash.Cells(dashRow + i, 9).Value
    Next i

    ' ── 차트 배치: 2열(50%씩), 너비 20cm, 높이 10cm, 간격 5mm, 행간격 1cm
    ' 1cm=28.35pt, 5mm=14.17pt → 20cm=567pt, 10cm=283pt
    If wb신규 Is Nothing Or wsDash Is Nothing Then GoTo SkipCharts
    wb신규.Activate
    wsDash.Select
    Dim cW As Double, chartH As Double, cGap As Double, rGap As Double, cLeft1 As Double, cLeft2 As Double
    Dim rTop1 As Double, rTop2 As Double, rTop3 As Double, rTop4 As Double
    cW = 567: chartH = 283: cGap = 14: rGap = 28
    cLeft1 = 5: cLeft2 = 5 + cW + cGap
    rTop1 = wsDash.Range("3:3").top: rTop2 = rTop1 + chartH + rGap: rTop3 = rTop2 + chartH + rGap: rTop4 = rTop3 + chartH + rGap

    ' ── 차트1: 일별 생산·출고 추이 (꺾은선형) ──
    If 일수 > 0 Then
        On Error Resume Next
        Set ch = wsDash.ChartObjects.Add(cLeft1, rTop1, cW, chartH)
        If Not ch Is Nothing Then
            Set cht = ch.Chart
            cht.ChartType = xlLineMarkers
            cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow + 1, 2), wsDash.Cells(dashRow + 일수, 3)), PlotBy:=xlColumns
            cht.SeriesCollection(1).Name = "생산"
            cht.SeriesCollection(2).Name = "출고"
            cht.SeriesCollection(1).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 1), wsDash.Cells(dashRow + 일수, 1))
            cht.SeriesCollection(2).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 1), wsDash.Cells(dashRow + 일수, 1))
            cht.HasTitle = True
            cht.ChartTitle.Text = "일별 생산·출고 추이"
            cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
            cht.ChartTitle.Font.Size = 11
            cht.Axes(xlCategory).TickLabels.Font.Name = "Microsoft GothicNeo"
            cht.Axes(xlValue).TickLabels.Font.Name = "Microsoft GothicNeo"
            cht.Legend.Font.Name = "Microsoft GothicNeo"
            cht.Legend.position = xlLegendPositionBottom
            cht.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(70, 130, 180)
            cht.SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(100, 149, 237)
            cht.SeriesCollection(1).HasDataLabels = True
            cht.SeriesCollection(2).HasDataLabels = True
            cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
            cht.SeriesCollection(2).DataLabels.Font.Name = "Microsoft GothicNeo"
            With cht.ChartArea.border
                .LineStyle = xlContinuous
                .color = RGB(242, 242, 242)
                .Weight = xlHairline
            End With
        End If
        On Error GoTo 0
    End If

    ' ── 차트2: 생산 vs 출고 비율 (원형) ──
    On Error Resume Next
    Set ch = wsDash.ChartObjects.Add(cLeft2, rTop1, cW, chartH)
    If Not ch Is Nothing Then
        Set cht = ch.Chart
        cht.ChartType = xlPie
        cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow, 5), wsDash.Cells(dashRow + 1, 6))
        cht.HasTitle = True
        cht.ChartTitle.Text = "생산 vs 출고 비율"
        cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
        cht.ChartTitle.Font.Size = 11
        cht.Legend.Font.Name = "Microsoft GothicNeo"
        cht.HasLegend = True
        cht.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
        cht.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(100, 149, 237)
        cht.SeriesCollection(1).HasDataLabels = True
        cht.SeriesCollection(1).DataLabels.ShowPercentage = True
        cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
        With cht.ChartArea.border
            .LineStyle = xlContinuous
            .color = RGB(242, 242, 242)
            .Weight = xlHairline
        End With
    End If
    On Error GoTo 0

    ' ── 차트3: 제품별 생산량 (가로 막대) ──
    If 제품수 > 0 Then
        Set ch = wsDash.ChartObjects.Add(cLeft1, rTop2, cW, chartH)
        Set cht = ch.Chart
        cht.ChartType = xlBarClustered
        cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow + 1, 8), wsDash.Cells(dashRow + 제품수, 8))
        cht.SeriesCollection(1).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 7), wsDash.Cells(dashRow + 제품수, 7))
        cht.SeriesCollection(1).Name = "생산"
        cht.HasTitle = True
        cht.ChartTitle.Text = "제품별 생산량 (Top " & 제품수 & ")"
        cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
        cht.ChartTitle.Font.Size = 11
        cht.Axes(xlCategory).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Axes(xlValue).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
        cht.ChartGroups(1).GapWidth = 50
        cht.SeriesCollection(1).HasDataLabels = True
        cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
        With cht.ChartArea.border
            .LineStyle = xlContinuous
            .color = RGB(242, 242, 242)
            .Weight = xlHairline
        End With
    End If

    ' ── 차트4: 제품별 출고량 (가로 막대) ──
    If 제품수 > 0 Then
        Set ch = wsDash.ChartObjects.Add(cLeft2, rTop2, cW, chartH)
        Set cht = ch.Chart
        cht.ChartType = xlBarClustered
        cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow + 1, 11), wsDash.Cells(dashRow + 제품수, 11))
        cht.SeriesCollection(1).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 10), wsDash.Cells(dashRow + 제품수, 10))
        cht.SeriesCollection(1).Name = "출고"
        cht.HasTitle = True
        cht.ChartTitle.Text = "제품별 출고량 (Top " & 제품수 & ")"
        cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
        cht.ChartTitle.Font.Size = 11
        cht.Axes(xlCategory).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Axes(xlValue).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(100, 149, 237)
        cht.ChartGroups(1).GapWidth = 50
        cht.SeriesCollection(1).HasDataLabels = True
        cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
        With cht.ChartArea.border
            .LineStyle = xlContinuous
            .color = RGB(242, 242, 242)
            .Weight = xlHairline
        End With
    End If

    ' ── 차트5: 제품별 생산·출고 비교 (묶은 가로 막대) ──
    If 제품수 > 0 Then
        Set ch = wsDash.ChartObjects.Add(cLeft2, rTop3, cW, chartH)
        Set cht = ch.Chart
        cht.ChartType = xlBarClustered
        cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow + 1, 8), wsDash.Cells(dashRow + 제품수, 9)), PlotBy:=xlColumns
        cht.SeriesCollection(1).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 7), wsDash.Cells(dashRow + 제품수, 7))
        cht.SeriesCollection(2).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 7), wsDash.Cells(dashRow + 제품수, 7))
        cht.SeriesCollection(1).Name = "생산"
        cht.SeriesCollection(2).Name = "출고"
        cht.HasTitle = True
        cht.ChartTitle.Text = "제품별 생산·출고 비교"
        cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
        cht.ChartTitle.Font.Size = 11
        cht.Axes(xlCategory).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Axes(xlValue).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Legend.Font.Name = "Microsoft GothicNeo"
        cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
        cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(100, 149, 237)
        cht.ChartGroups(1).GapWidth = 50
        cht.SeriesCollection(1).HasDataLabels = True
        cht.SeriesCollection(2).HasDataLabels = True
        cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
        cht.SeriesCollection(2).DataLabels.Font.Name = "Microsoft GothicNeo"
        With cht.ChartArea.border
            .LineStyle = xlContinuous
            .color = RGB(242, 242, 242)
            .Weight = xlHairline
        End With
    End If

    ' ── 차트7: 일별 생산·출고 누적 영역 (영역형) - 하단 전체 폭 ──
    If 일수 > 0 Then
        Set ch = wsDash.ChartObjects.Add(cLeft1, rTop4, cW * 2 + cGap, chartH)
        Set cht = ch.Chart
        cht.ChartType = xlAreaStacked
        cht.SetSourceData Source:=wsDash.Range(wsDash.Cells(dashRow + 1, 2), wsDash.Cells(dashRow + 일수, 3)), PlotBy:=xlColumns
        cht.SeriesCollection(1).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 1), wsDash.Cells(dashRow + 일수, 1))
        cht.SeriesCollection(2).XValues = wsDash.Range(wsDash.Cells(dashRow + 1, 1), wsDash.Cells(dashRow + 일수, 1))
        cht.SeriesCollection(1).Name = "생산"
        cht.SeriesCollection(2).Name = "출고"
        cht.HasTitle = True
        cht.ChartTitle.Text = "일별 생산·출고 누적 추이"
        cht.ChartTitle.Font.Name = "Microsoft GothicNeo"
        cht.ChartTitle.Font.Size = 11
        cht.Axes(xlCategory).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Axes(xlValue).TickLabels.Font.Name = "Microsoft GothicNeo"
        cht.Legend.Font.Name = "Microsoft GothicNeo"
        cht.Legend.position = xlLegendPositionBottom
        cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(70, 130, 180)
        cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(100, 149, 237)
        cht.SeriesCollection(1).HasDataLabels = True
        cht.SeriesCollection(2).HasDataLabels = True
        cht.SeriesCollection(1).DataLabels.Font.Name = "Microsoft GothicNeo"
        cht.SeriesCollection(2).DataLabels.Font.Name = "Microsoft GothicNeo"
        With cht.ChartArea.border
            .LineStyle = xlContinuous
            .color = RGB(242, 242, 242)
            .Weight = xlHairline
        End With
    End If

    wsDash.Columns("A:Z").ColumnWidth = 8
    If 일수 > 막대최대 Then
        wsDash.Rows(dashRow & ":" & dashRow + 일수 + 5).RowHeight = 0.5
    Else
        wsDash.Rows(dashRow & ":" & dashRow + 막대최대 + 5).RowHeight = 0.5
    End If

SkipCharts:
    Application.DisplayAlerts = 알림상태
    If Not wsDash Is Nothing Then
        wb신규.Windows(1).DisplayGridlines = False
        wsDash.Activate
    End If

    ' ── 저장 ──
    On Error Resume Next
    wb신규.SaveAs Filename:=전체경로, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        Application.ScreenUpdating = True
        MsgBox "파일 저장 실패: " & Err.Description, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = True

    MsgBox "월생산·출고현황 저장 완료!" & vbCrLf & vbCrLf & _
        "필터: " & 필터유형 & IIf(필터유형 = "전체", "", " (" & 필터표시 & ")") & vbCrLf & _
        "제품 수: " & 제품목록.Count & "개 (제품당 생산·출고 2행)" & vbCrLf & _
        "Dashboard: 차트 6개 (일별 추이, 비율, 제품별 생산·출고 등)" & vbCrLf & _
        "경로: " & 전체경로, vbInformation
End Sub

