Attribute VB_Name = "Module2"
' ============================================================
' 생산일보 목업데이터 생성 (속도·최적화)
' ============================================================
' 1. 메뉴초기화 실행 후, 지워지는 영역만 랜덤 목업 데이터로 채움
' 2. BK4 = 현재날짜 +1일씩 증가 (실행할 때마다 다른 날짜)
' 3. 유효성 목록 열 단위 캐시로 Validation 접근 최소화
' 4. 배열 일괄 입력 + Transpose 제거(2D 배열 직접 사용)
'
' [수식 영역 절대 미쓰기]
'   T03 주간소계(30~32행), T06 야간소계(60~62행)
'   T07/T08 완제품 수식열 X,AB,AD,AF,AH(주간) / 야간측
'   T09 종별수량(80~82행), T10 원자재수율 = 나머지 셀 수식
'   [T10 값입력 영역만] G88:K91,G93:K96,G98:K101, AP88:AT91,AP93:AT96,AP98:AT101
'   T11 가중시간(109행), 주야휴합계(110행)
'
' [셀 매핑 - 생산일보_셀매핑.json 기준]
' 원지 6그룹 × 12열 패턴: base=3,15,27,39,51,63
'   Co=base(3,15,27,39,51,63)=C,O,AA,AM,AY,BK, ID=base+2, 중량=base+9
'   Co열: Raw data 시트 H3~마지막(I열 기준) 넘버만 기록 (없으면 R+랜덤 폴백)
'   T01 직기 KT-1~5,FT (7~16행) / T02 직기 JRT-1~2,BT-1~3,PRE (20~29행)
'   T04 봉제 KT-1~5,FT (35~44행) / T05 봉제 JRT-1~2,BT-1~3,PRE (48~57행)
'
' [셀 매핑 - 완제품 가공내역]
'   T07 주간(68~77행): B(2)=제품, Q(17)=설비, T(20)=생산량BOX
'   T08 야간(68~77행): AK(37)=제품, AZ(52)=설비, BC(55)=생산량BOX
'
' [셀 매핑 - 작업시간/인력투입] (생산일보_셀매핑.json)
'   T11 108행: 설비별 base열 KT=7, FT=43, JRT=19, BT=31, PRE=55
'     오프셋 0,2,4,6,8 → 주간정규/연장, 야간정규/심야, 휴일
'   T11 111행: Direct 인력  /  112행: Indirect 인력 (동일 열 구조)
'
' [셀 매핑 - 비가동현황]
'   T12: A121~A126=내용, AJ121~AJ126=해결책
' ============================================================

Option Explicit

Private Const 시트명일보 As String = "생산일보"
Private Const 시트명RawIndex As String = "Raw data_index"
Private Const 시트명RawData As String = "Raw data"   ' Co(ROLLID) 넘버: H3~마지막 (I열 기준)
Private Const 시트명표준사양서 As String = "제품표준사양서"
Private Const 목업상태시트 As String = "목업상태"
Private Const 표준_G열 As Long = 7
Private Const 표준_AG열 As Long = 33
Private Const 표준_AJ열 As Long = 36
Private Const 표준_데이터시작행 As Long = 9
Private Const 목업날짜셀 As String = "A1"
Private Const RawIndex_헤더행 As Long = 2
Private Const RawIndex_데이터시작행 As Long = 3

Private Const 원지그룹수 As Long = 6
Private Const 원지열간격 As Long = 12

' ============================================================
' 메인: 메뉴초기화 호출 후 목업 데이터 생성
' ============================================================
Public Sub 생산일보목업데이터_생성()
    Dim ws As Worksheet
    Dim 원래ScreenUpdating As Boolean
    Dim 원래EnableEvents As Boolean
    Dim 원래Calculation As XlCalculation
    Dim 원래DisplayAlerts As Boolean
    Dim 다음날짜 As Date
    Dim prevSheet As Worksheet
    Dim prevScrollRow As Long
    Dim prevScrollCol As Long
    Dim prevAddr As String

    On Error GoTo ErrHandler

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(시트명일보)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        MsgBox "시트 '" & 시트명일보 & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set prevSheet = ActiveSheet
    prevScrollRow = ActiveWindow.ScrollRow
    prevScrollCol = ActiveWindow.ScrollColumn
    prevAddr = ActiveCell.Address
    If prevAddr = "" Then prevAddr = "A1"
    On Error GoTo ErrHandler

    원래ScreenUpdating = Application.ScreenUpdating
    원래EnableEvents = Application.EnableEvents
    원래Calculation = Application.Calculation
    원래DisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.CutCopyMode = False

    Randomize
    Application.Run "메뉴초기화"
    On Error Resume Next
    Application.Run "생산일보_제품유효성_적용"
    Err.Clear
    On Error GoTo ErrHandler

    Dim 유효성캐시 As Object
    Set 유효성캐시 = CreateObject("Scripting.Dictionary")

    ' BK4 날짜 +1일 진행
    If Not IsEmpty(ws.Range("BK4").Value) And IsDate(ws.Range("BK4").Value) Then
        다음날짜 = CDate(ws.Range("BK4").Value) + 1
    Else
        다음날짜 = Date + 1
    End If
    ws.Range("BK4").Value = 다음날짜

    ' 원지사용내역: 주간/야간 × KT계열/JRT계열
    Call 목업_원지_테이블(ws, 유효성캐시, 7, Array("KT-1", "KT-2", "KT-3", "KT-4", "KT-5", "FT"))
    Call 목업_원지_테이블(ws, 유효성캐시, 20, Array("JRT-1", "JRT-2", "BT-1", "BT-2", "BT-3", "PRE"))
    Call 목업_원지_테이블(ws, 유효성캐시, 35, Array("KT-1", "KT-2", "KT-3", "KT-4", "KT-5", "FT"))
    Call 목업_원지_테이블(ws, 유효성캐시, 48, Array("JRT-1", "JRT-2", "BT-1", "BT-2", "BT-3", "PRE"))

    ' 완제품 가공내역
    Call 목업_완제품_주간(ws, 유효성캐시)
    Call 목업_완제품_야간(ws, 유효성캐시)

    ' 수식 계산 (T10 원자재수율 등 수식 의존)
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate

    ' T10 주간/월간 누계·평균 계산 (일일생산DB 이력 기반) - 실패해도 DB전송 진행
    On Error Resume Next: Err.Clear
    Application.Run "주월_누적_재계산"
    Err.Clear
    On Error GoTo ErrHandler

    ' 작업시간 / 인력투입 (실패해도 DB전송 진행)
    On Error Resume Next: Call 목업_작업시간(ws): Err.Clear: On Error GoTo ErrHandler
    On Error Resume Next: Call 목업_인력투입(ws): Err.Clear: On Error GoTo ErrHandler

    ' 비가동현황
    On Error Resume Next: Call 목업_비가동현황(ws): Err.Clear: On Error GoTo ErrHandler

    ' 일일생산DB에 저장 (BK4 날짜 포함) ← 반드시 실행
    On Error Resume Next: Err.Clear
    Application.Run "일일생산일보_최적화DB_전송"
    If Err.Number <> 0 Then
        If Not g_무음모드 Then MsgBox "DB 저장 오류: " & Err.Description, vbCritical
        Err.Raise Err.Number, "목업_DB전송", "일일생산일보_최적화DB_전송 실패: " & Err.Description
    End If
    On Error GoTo ErrHandler

    If Not g_무음모드 Then MsgBox "목업 데이터 생성 완료." & vbCrLf & "날짜: " & Format(다음날짜, "yyyy-mm-dd"), vbInformation
    GoTo exitSub
ErrHandler:
    If Not g_무음모드 Then MsgBox "목업 데이터 생성 중 오류: " & Err.Description, vbCritical
exitSub:
    Application.Calculation = 원래Calculation
    Application.EnableEvents = 원래EnableEvents
    Application.DisplayAlerts = 원래DisplayAlerts
    Application.ScreenUpdating = 원래ScreenUpdating
    Application.CutCopyMode = False
    Application.StatusBar = False
    On Error Resume Next
    If Not prevSheet Is Nothing Then
        prevSheet.Activate
        ActiveWindow.ScrollRow = prevScrollRow
        ActiveWindow.ScrollColumn = prevScrollCol
        Application.GoTo prevSheet.Range(prevAddr), False
    End If
    On Error GoTo 0
End Sub

' 목업상태 시트 1회 조회 후 날짜 조회·저장
Private Function 목업_날짜_조회및저장(ws일보 As Worksheet, 저장할다음날짜 As Boolean) As Date
    Static wsSt As Worksheet
    Dim 다음날짜 As Date
    If wsSt Is Nothing Then
        On Error Resume Next
        Set wsSt = ThisWorkbook.Worksheets(목업상태시트)
        If wsSt Is Nothing Then
            Set wsSt = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wsSt.Name = 목업상태시트
            wsSt.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0
    End If
    If IsEmpty(wsSt.Range(목업날짜셀).Value) Or Not IsDate(wsSt.Range(목업날짜셀).Value) Then
        다음날짜 = Date + 1
    Else
        다음날짜 = CDate(wsSt.Range(목업날짜셀).Value)
    End If
    If 저장할다음날짜 Then wsSt.Range(목업날짜셀).Value = 다음날짜 + 1
    목업_날짜_조회및저장 = 다음날짜
End Function

' Raw data_index 2행 헤더에서 헤더명과 일치하는 열 번호 반환
Private Function RawIndex_헤더열찾기(wsRaw As Worksheet, 헤더명 As String) As Long
    Dim lastCol As Long, c As Long
    Dim t As String, findVal As String
    If wsRaw Is Nothing Then RawIndex_헤더열찾기 = 0: Exit Function
    findVal = LCase(Trim(Replace(Replace(헤더명, " ", ""), chR(10), "")))
    lastCol = wsRaw.Cells(RawIndex_헤더행, wsRaw.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 100
    For c = 1 To lastCol
        t = Trim(Replace(Replace(CStr(wsRaw.Cells(RawIndex_헤더행, c).Value), " ", ""), chR(10), ""))
        If LCase(t) = findVal Then RawIndex_헤더열찾기 = c: Exit Function
    Next c
    RawIndex_헤더열찾기 = 0
End Function

' 제품표준사양서 G9~마지막 행 값 (중복·빈셀 제외) → 0-based 배열 반환
Private Function 표준사양서_G9_제품목록(ws사양 As Worksheet) As Variant
    Dim lastRow As Long, r As Long, n As Long, v As String
    Dim out() As String
    If ws사양 Is Nothing Then 표준사양서_G9_제품목록 = Empty: Exit Function
    lastRow = ws사양.Cells(ws사양.Rows.Count, 표준_G열).End(xlUp).Row
    If lastRow < 표준_데이터시작행 Then 표준사양서_G9_제품목록 = Empty: Exit Function
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    For r = 표준_데이터시작행 To lastRow
        v = Trim(CStr(ws사양.Cells(r, 표준_G열).Value))
        If v <> "" And Not dict.exists(v) Then dict.Add v, True
    Next r
    If dict.Count = 0 Then 표준사양서_G9_제품목록 = Empty: Exit Function
    ReDim out(0 To dict.Count - 1)
    n = 0
    Dim k As Variant
    For Each k In dict.keys
        out(n) = CStr(k)
        n = n + 1
    Next k
    표준사양서_G9_제품목록 = out
End Function

' Raw data_index C열(제품) 3행~마지막 행 값 → 0-based 배열 반환 (폴백)
Private Function RawDataIndex_제품목록(wsRaw As Worksheet) As Variant
    Dim lastRow As Long, r As Long, n As Long, v As String
    Dim arr As Variant, out() As String
    If wsRaw Is Nothing Then RawDataIndex_제품목록 = Empty: Exit Function
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 3).End(xlUp).Row
    If lastRow < RawIndex_데이터시작행 Then RawDataIndex_제품목록 = Empty: Exit Function
    arr = wsRaw.Range(wsRaw.Cells(RawIndex_데이터시작행, 3), wsRaw.Cells(lastRow, 3)).Value
    n = 0
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            v = Trim(CStr(arr(r, 1)))
            If v <> "" Then n = n + 1
        Next r
    Else
        v = Trim(CStr(arr))
        If v <> "" Then n = 1
    End If
    If n = 0 Then RawDataIndex_제품목록 = Empty: Exit Function
    ReDim out(0 To n - 1)
    n = 0
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            v = Trim(CStr(arr(r, 1)))
            If v <> "" Then out(n) = v: n = n + 1
        Next r
    Else
        out(0) = Trim(CStr(arr)): n = 1
    End If
    RawDataIndex_제품목록 = out
End Function

' Raw data 시트 H3~마지막행(I열 기준) → Co(ROLLID) 넘버 0-based 배열 (원지사용내역 Co열용)
Private Function RawData_Co넘버목록() As Variant
    Dim wsRaw As Worksheet
    Dim lastRowI As Long, r As Long, n As Long, v As Variant
    Dim out() As Variant
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Worksheets(시트명RawData)
    If wsRaw Is Nothing Then Set wsRaw = ThisWorkbook.Worksheets(시트명RawIndex)
    On Error GoTo 0
    If wsRaw Is Nothing Then RawData_Co넘버목록 = Empty: Exit Function
    ' I열(9) 기준 마지막 행
    lastRowI = wsRaw.Cells(wsRaw.Rows.Count, 9).End(xlUp).Row
    If lastRowI < 3 Then RawData_Co넘버목록 = Empty: Exit Function
    n = 0
    For r = 3 To lastRowI
        v = wsRaw.Cells(r, 8).Value  ' H열
        If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then n = n + 1
    Next r
    If n = 0 Then RawData_Co넘버목록 = Empty: Exit Function
    ReDim out(0 To n - 1)
    n = 0
    For r = 3 To lastRowI
        v = wsRaw.Cells(r, 8).Value
        If Not IsEmpty(v) And Trim(CStr(v)) <> "" Then
            out(n) = v
            n = n + 1
        End If
    Next r
    RawData_Co넘버목록 = out
End Function

' Raw data_index 설비호기 열 3행~마지막 → 0-based 배열
Private Function RawDataIndex_설비호기목록(wsRaw As Worksheet) As Variant
    Dim col As Long, lastRow As Long, r As Long, n As Long, v As String
    Dim arr As Variant, out() As String
    If wsRaw Is Nothing Then RawDataIndex_설비호기목록 = Empty: Exit Function
    col = RawIndex_헤더열찾기(wsRaw, "설비호기")
    If col = 0 Then col = RawIndex_헤더열찾기(wsRaw, "가공라인호기")
    If col = 0 Then col = RawIndex_헤더열찾기(wsRaw, "가공라인")
    If col = 0 Then RawDataIndex_설비호기목록 = Empty: Exit Function
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, col).End(xlUp).Row
    If lastRow < RawIndex_데이터시작행 Then RawDataIndex_설비호기목록 = Empty: Exit Function
    arr = wsRaw.Range(wsRaw.Cells(RawIndex_데이터시작행, col), wsRaw.Cells(lastRow, col)).Value
    n = 0
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            v = Trim(CStr(arr(r, 1)))
            If v <> "" Then n = n + 1
        Next r
    Else
        v = Trim(CStr(arr))
        If v <> "" Then n = 1
    End If
    If n = 0 Then RawDataIndex_설비호기목록 = Empty: Exit Function
    ReDim out(0 To n - 1)
    n = 0
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            v = Trim(CStr(arr(r, 1)))
            If v <> "" Then out(n) = v: n = n + 1
        Next r
    Else
        out(0) = Trim(CStr(arr))
    End If
    RawDataIndex_설비호기목록 = out
End Function

' 유효성 목록 열 단위 캐시. 병합 셀은 왼쪽 위 셀에서 유효성 읽기/쓰기
Private Function 셀_유효성목록_랜덤선택(ws As Worksheet, r As Long, c As Long, ByRef cache As Object) As Boolean
    Dim key As String, listArr, i As Long, ub As Long, v As Validation, s As String, rngList As Range
    Dim cell As Range, writeCell As Range
    Set cell = ws.Cells(r, c)
    If cell.MergeCells Then Set writeCell = cell.MergeArea.Cells(1, 1) Else Set writeCell = cell
    key = "C" & c
    If cache.exists(key) Then
        listArr = cache(key)
        If IsEmpty(listArr) Then 셀_유효성목록_랜덤선택 = False: Exit Function
        ub = UBound(listArr)
        i = Int(Rnd * (ub + 1))
        writeCell.Value = listArr(i)
        셀_유효성목록_랜덤선택 = True
        Exit Function
    End If
    On Error Resume Next
    Set v = writeCell.Validation
    If v Is Nothing Or v.Type <> xlValidateList Then cache.Add key, Empty: 셀_유효성목록_랜덤선택 = False: Exit Function
    s = v.Formula1
    If left(Trim(s), 1) = "=" Then
        Set rngList = Application.Evaluate(Mid(Trim(s), 2))
        If rngList Is Nothing Then cache.Add key, Empty: 셀_유효성목록_랜덤선택 = False: Exit Function
        ub = rngList.Cells.Count
        If ub < 1 Then cache.Add key, Empty: 셀_유효성목록_랜덤선택 = False: Exit Function
        ReDim listArr(0 To ub - 1)
        For i = 0 To ub - 1
            listArr(i) = rngList.Cells(i + 1).Value
        Next i
    Else
        listArr = Split(Replace(Replace(s, ",", "|"), ";", "|"), "|")
        ub = UBound(listArr)
        For i = 0 To ub
            listArr(i) = Trim(listArr(i))
        Next i
    End If
    cache.Add key, listArr
    ub = UBound(listArr)
    i = Int(Rnd * (ub + 1))
    writeCell.Value = listArr(i)
    셀_유효성목록_랜덤선택 = True
    On Error GoTo 0
End Function

' Co(ROLLID) 열에 Raw data H열 넘버 기록 (없으면 R+랜덤 폴백)
Private Sub 목업_원지_Co넘버쓰기(ws As Worksheet, r As Long, colID As Long)
    Dim arr As Variant
    arr = RawData_Co넘버목록()
    If IsEmpty(arr) Then
        Call 목업_셀쓰기(ws, r, colID, "R" & Int(랜덤숫자(1000, 9999, 0)))
    Else
        Call 목업_셀쓰기(ws, r, colID, arr(Int(Rnd * (UBound(arr) + 1))))
    End If
End Sub

Private Function 랜덤숫자(최소 As Double, 최대 As Double, Optional 소수자리 As Long = 0) As Double
    Dim v As Double
    v = 최소 + Rnd * (최대 - 최소 + 0.0000001)
    If 소수자리 >= 0 Then v = Round(v, 소수자리)
    랜덤숫자 = v
End Function

Private Function 랜덤_25퍼() As Boolean
    랜덤_25퍼 = (Rnd < 0.25)
End Function

Private Sub 목업_셀쓰기(ws As Worksheet, r As Long, c As Long, val As Variant)
    Dim cell As Range
    Set cell = ws.Cells(r, c)
    If cell.MergeCells Then
        If cell.MergeArea.Cells(1, 1).Address <> cell.Address Then Exit Sub
    End If
    cell.Value = val
End Sub

Private Function 목업_셀에_데이터있음(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then 목업_셀에_데이터있음 = (CDbl(v) > 0): Exit Function
    목업_셀에_데이터있음 = (Trim(CStr(v)) <> "")
End Function

' 원지 그룹의 base 열 배열 반환 (3,15,27,39,51,63)
Private Function 원지_그룹_Base열() As Variant
    원지_그룹_Base열 = Array(3, 15, 27, 39, 51, 63)
End Function

' ============================================================
' 원지사용내역 목업 (테이블 공용)
' 시작행: T01=7, T02=20, T04=35, T05=48 (각 10행 데이터)
' 그룹명배열: 6개 (KT-1~5+FT 또는 JRT-1~2+BT-1~3+PRE)
' 각 그룹: 원지=base, ID=base+2, 중량=base+9
' 1~3개 그룹만 선택, 선택된 그룹의 25% 행만 채움
' ============================================================
Private Sub 목업_원지_테이블(ws As Worksheet, ByRef 유효성캐시 As Object, _
    시작행 As Long, 그룹명 As Variant)

    Dim baseCols As Variant
    baseCols = 원지_그룹_Base열()

    Dim 사용호기수 As Long, i As Long, j As Long, n As Long
    Dim 순서(0 To 5) As Long
    Dim 사용인덱스() As Long

    사용호기수 = 1 + Int(Rnd * 3)
    For i = 0 To 5: 순서(i) = i: Next i
    For i = 0 To 4
        j = i + Int(Rnd * (6 - i))
        n = 순서(i): 순서(i) = 순서(j): 순서(j) = n
    Next i
    ReDim 사용인덱스(0 To 사용호기수 - 1)
    For i = 0 To 사용호기수 - 1: 사용인덱스(i) = 순서(i): Next i

    Dim r As Long, idx As Long
    Dim bc As Long, col원지 As Long, colID As Long, col중량 As Long
    For r = 시작행 To 시작행 + 9
        If Not 랜덤_25퍼() Then GoTo Next원지행
        For i = 0 To 사용호기수 - 1
            idx = 사용인덱스(i)
            bc = baseCols(idx)
            col원지 = bc
            colID = bc + 2
            col중량 = bc + 9

            ' Co열(C,O,AA,AM,AY,BK=3,15,27,39,51,63)에 Raw data H3~마지막(I열기준) 넘버만 기록
            Call 목업_원지_Co넘버쓰기(ws, r, bc)
            Call 목업_셀쓰기(ws, r, colID, "R" & Int(랜덤숫자(1000, 9999, 0)))
            Call 목업_셀쓰기(ws, r, col중량, 랜덤숫자(50, 1500, 0))
        Next i
Next원지행:
    Next r
End Sub

' T01에서 실제 데이터가 입력된 KT 그룹 목록 반환 (완제품 설비호기 매핑용)
Private Function 목업_원지_KT목록(ws As Worksheet) As Variant
    Dim baseCols As Variant
    baseCols = 원지_그룹_Base열()
    Dim ktNames As Variant
    ktNames = Array("KT-1", "KT-2", "KT-3", "KT-4", "KT-5", "FT")
    Dim idx As Long, r As Long, n As Long, v As Variant
    Dim result() As String
    n = 0
    For idx = 0 To 5
        For r = 7 To 16
            v = ws.Cells(r, baseCols(idx)).Value
            If 목업_셀에_데이터있음(v) Then n = n + 1: Exit For
        Next r
    Next idx
    If n = 0 Then 목업_원지_KT목록 = Empty: Exit Function
    ReDim result(0 To n - 1)
    n = 0
    For idx = 0 To 5
        For r = 7 To 16
            v = ws.Cells(r, baseCols(idx)).Value
            If 목업_셀에_데이터있음(v) Then result(n) = ktNames(idx): n = n + 1: Exit For
        Next r
    Next idx
    목업_원지_KT목록 = result
End Function

' ============================================================
' 완제품 주간 목업: T07 rows 68-77
' B(2)=제품, Q(17)=설비호기, T(20)=생산량BOX
' X,AB,AD,AF,AH = 수식(자동계산, 절대 미쓰기)
' ============================================================
Private Sub 목업_완제품_주간(ws As Worksheet, ByRef 유효성캐시 As Object)
    Dim ws사양 As Worksheet, wsRaw As Worksheet
    Dim 제품목록 As Variant, ub제품 As Long, i As Long
    Dim kt목록 As Variant, ubKT As Long
    Dim r As Long, k As Long

    On Error Resume Next
    Set ws사양 = ThisWorkbook.Worksheets(시트명표준사양서)
    Set wsRaw = ThisWorkbook.Worksheets(시트명RawIndex)
    On Error GoTo 0

    kt목록 = 목업_원지_KT목록(ws)
    If IsEmpty(kt목록) Then
        ReDim kt목록(0 To 0): kt목록(0) = "KT-1"
    End If
    ubKT = UBound(kt목록)
    k = ubKT + 1

    제품목록 = Empty
    If Not ws사양 Is Nothing Then 제품목록 = 표준사양서_G9_제품목록(ws사양)
    If IsEmpty(제품목록) And Not wsRaw Is Nothing Then 제품목록 = RawDataIndex_제품목록(wsRaw)

    ' 원지에 사용된 설비는 완제품에 반드시 모두 표시 (1~k행)
    For r = 1 To k
        If Not IsEmpty(제품목록) Then
            ub제품 = UBound(제품목록)
            i = Int(Rnd * (ub제품 + 1))
            Call 목업_셀쓰기(ws, r + 67, 2, 제품목록(i))
        Else
            Call 셀_유효성목록_랜덤선택(ws, r + 67, 2, 유효성캐시)
        End If
        Call 목업_셀쓰기(ws, r + 67, 17, kt목록(r - 1))
        Call 목업_셀쓰기(ws, r + 67, 20, 랜덤숫자(50, 350, 0))
    Next r

    ' 나머지 행(k+1~10)은 25% 랜덤
    For r = k + 1 To 10
        If Not 랜덤_25퍼() Then GoTo Next완주간
        If Not IsEmpty(제품목록) Then
            ub제품 = UBound(제품목록)
            i = Int(Rnd * (ub제품 + 1))
            Call 목업_셀쓰기(ws, r + 67, 2, 제품목록(i))
        Else
            Call 셀_유효성목록_랜덤선택(ws, r + 67, 2, 유효성캐시)
        End If
        Call 목업_셀쓰기(ws, r + 67, 17, kt목록(Int(Rnd * (ubKT + 1))))
        Call 목업_셀쓰기(ws, r + 67, 20, 랜덤숫자(50, 350, 0))
Next완주간:
    Next r
End Sub

' ============================================================
' 완제품 야간 목업: T08 rows 68-77 (우측)
' AK(37)=제품, AZ(52)=설비호기, BC(55)=생산량BOX
' ============================================================
Private Sub 목업_완제품_야간(ws As Worksheet, ByRef 유효성캐시 As Object)
    Dim ws사양 As Worksheet, wsRaw As Worksheet
    Dim 제품목록 As Variant, ub제품 As Long, i As Long
    Dim r As Long
    Dim 설비목록 As Variant
    설비목록 = Array("KT-1", "KT-2", "KT-3", "FT", "JRT-1", "BT-1")

    On Error Resume Next
    Set ws사양 = ThisWorkbook.Worksheets(시트명표준사양서)
    Set wsRaw = ThisWorkbook.Worksheets(시트명RawIndex)
    On Error GoTo 0
    제품목록 = Empty
    If Not ws사양 Is Nothing Then 제품목록 = 표준사양서_G9_제품목록(ws사양)
    If IsEmpty(제품목록) And Not wsRaw Is Nothing Then 제품목록 = RawDataIndex_제품목록(wsRaw)

    For r = 68 To 77
        If Not 랜덤_25퍼() Then GoTo Next완야간
        If Not IsEmpty(제품목록) Then
            ub제품 = UBound(제품목록)
            i = Int(Rnd * (ub제품 + 1))
            Call 목업_셀쓰기(ws, r, 37, 제품목록(i))
        Else
            Call 목업_셀쓰기(ws, r, 37, "야간품" & (r - 67))
        End If
        Call 목업_셀쓰기(ws, r, 52, 설비목록(Int(Rnd * (UBound(설비목록) + 1))))
        Call 목업_셀쓰기(ws, r, 55, 랜덤숫자(30, 250, 0))
Next완야간:
    Next r
End Sub

' ============================================================
' 작업시간 목업: T11 row 108
' 5개 설비(KT,FT,JRT,BT,PRE) × 5열(주간정규,주간연장,야간정규,야간심야,휴일)
' ============================================================
Private Sub 목업_작업시간(ws As Worksheet)
    Dim 설비Base As Variant
    설비Base = Array(7, 43, 19, 31, 55)
    Dim i As Long, bc As Long

    For i = 0 To 4
        bc = 설비Base(i)
        ws.Cells(108, bc).Value = 랜덤숫자(6, 8, 1)
        ws.Cells(108, bc + 2).Value = 랜덤숫자(0, 2, 1)
        ws.Cells(108, bc + 4).Value = 랜덤숫자(6, 8, 1)
        ws.Cells(108, bc + 6).Value = 랜덤숫자(0, 2, 1)
        If 랜덤_25퍼() Then ws.Cells(108, bc + 8).Value = 랜덤숫자(4, 8, 1)
    Next i
End Sub

' ============================================================
' 인력투입 목업: T11 rows 111-112
' Direct(111행), Indirect(112행) - 동일 열 구조
' ============================================================
Private Sub 목업_인력투입(ws As Worksheet)
    Dim 설비Base As Variant
    설비Base = Array(7, 43, 19, 31, 55)
    Dim i As Long, bc As Long

    For i = 0 To 4
        bc = 설비Base(i)
        ws.Cells(111, bc).Value = 랜덤숫자(2, 6, 0)
        ws.Cells(111, bc + 2).Value = 랜덤숫자(0, 2, 0)
        ws.Cells(111, bc + 4).Value = 랜덤숫자(2, 5, 0)
        ws.Cells(111, bc + 6).Value = 랜덤숫자(0, 1, 0)
        ws.Cells(111, bc + 8).Value = IIf(랜덤_25퍼(), 랜덤숫자(1, 3, 0), 0)
        ws.Cells(112, bc).Value = 랜덤숫자(0, 2, 0)
        ws.Cells(112, bc + 2).Value = 0
        ws.Cells(112, bc + 4).Value = 랜덤숫자(0, 1, 0)
        ws.Cells(112, bc + 6).Value = 0
        ws.Cells(112, bc + 8).Value = 0
    Next i
End Sub

' ============================================================
' 비가동현황 목업: T12 rows 121-126
' A121=비가동 내용, AJ121=해결책
' ============================================================
Private Sub 목업_비가동현황(ws As Worksheet)
    Dim 특이목록, 해결목록
    특이목록 = Array("원자재 입고 지연", "설비 점검", "품질 이슈 없음", "일부 라인 가동률 저하", "전원 점검 완료", "")
    해결목록 = Array("다음 일정 조정", "당일 가동 연장", "별도 조치 없음", "예비 설비 활용", "")
    Dim r As Long
    For r = 121 To 123
        If 랜덤_25퍼() Then
            ws.Cells(r, 1).Value = 특이목록(Int(Rnd * (UBound(특이목록) + 1)))
            ws.Cells(r, 36).Value = 해결목록(Int(Rnd * (UBound(해결목록) + 1)))
        End If
    Next r
End Sub





