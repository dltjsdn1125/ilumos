Attribute VB_Name = "Module23"
' ============================================================
' Module23: 검증 보고서 출력 (독립 매크로)
' ============================================================
' [사용법]
'   1. 보고서_출력 매크로 실행
'   2. 일일생산DB 시트에서 데이터 범위 선택 (또는 취소 시 전체 사용)
'   3. 새 워크북에 검증 보고서 6시트 출력
'
' [오버플로우 방지]
'   - 모든 카운터 Long 타입 사용
'   - 최대 처리 행수 제한 (기본 80,000행)
'   - 배열 크기 사전 검증
'   - 날짜 루프 일수 제한 (3660일 = 10년)
' ============================================================

Option Explicit

Public g_보고서_마지막저장경로 As String

Private Const PA_허용오차 As Double = 1#
Private Const MAX_ROWS As Long = 80000
Private Const MAX_일수 As Long = 3660
Private Const DB_MIN_COLS As Long = 126

Private Function M23_SDbl(v As Variant) As Double
    On Error Resume Next
    M23_SDbl = 0
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then M23_SDbl = CDbl(v)
    If Err.Number <> 0 Then M23_SDbl = 0: Err.Clear
End Function

Private Function M23_주차번호(d As Date) As Long
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    If d < 첫일요일 Then M23_주차번호 = 1: Exit Function
    M23_주차번호 = 2 + CLng((d - 첫일요일) \ 7)
End Function

Private Function M23_주차키(d As Date) As String
    M23_주차키 = Month(d) & "월" & M23_주차번호(d) & "주"
End Function

Private Function M23_주차_첫날(d As Date) As Date
    Dim w As Long: w = M23_주차번호(d)
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    If w = 1 Then M23_주차_첫날 = 월1일: Exit Function
    M23_주차_첫날 = 첫일요일 + CLng(w - 2) * 7
End Function

Private Function M23_주차_마지막일(d As Date) As Date
    Dim w As Long: w = M23_주차번호(d)
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    Dim 말일 As Date: 말일 = DateSerial(Year(d), Month(d) + 1, 0)
    If w = 1 Then M23_주차_마지막일 = 첫일요일 - 1: Exit Function
    Dim 끝 As Date: 끝 = 첫일요일 + CLng(w - 2) * 7 + 6
    If 끝 > 말일 Then 끝 = 말일
    M23_주차_마지막일 = 끝
End Function

Private Function M23_월키(d As Date) As String
    M23_월키 = Month(d) & "월"
End Function

Private Sub M23_헤더서식(rng As Range)
    With rng
        .Font.Bold = True
        .Font.color = vbWhite
        .Interior.color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ============================================================
' ★ 메인 진입점: 유저가 시작일·종료일 입력
' ============================================================
Public Sub 보고서_출력()
    Dim wsDB As Worksheet
    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    On Error GoTo 0
    If wsDB Is Nothing Then
        MsgBox "'일일생산DB' 시트가 없습니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim dbLastRow As Long, dbMaxCol As Long
    dbLastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If dbLastRow < 2 Then
        MsgBox "일일생산DB에 데이터가 없습니다. (2행 이상 필요)", vbCritical, "보고서 출력"
        Exit Sub
    End If

    dbMaxCol = Application.WorksheetFunction.Min(126, wsDB.Cells(2, wsDB.Columns.Count).End(xlToLeft).Column)
    If dbMaxCol < 126 Then dbMaxCol = 126

    Dim dbFull As Variant
    On Error GoTo ErrRead
    dbFull = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(dbLastRow, dbMaxCol)).Value
    On Error GoTo 0

    Dim dMinDB As Date, dMaxDB As Date, r As Long, vDate As Variant
    dMinDB = DateSerial(9999, 12, 31): dMaxDB = DateSerial(1900, 1, 1)
    For r = 1 To UBound(dbFull, 1)
        vDate = dbFull(r, 1)
        If Not IsEmpty(vDate) And IsDate(vDate) Then
            If CDate(vDate) < dMinDB Then dMinDB = CDate(vDate)
            If CDate(vDate) > dMaxDB Then dMaxDB = CDate(vDate)
        End If
    Next r
    If dMinDB > dMaxDB Then
        MsgBox "일일생산DB에 유효한 날짜(1열)가 없습니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim 시작일 As Date, 종료일 As Date
    Dim s시작 As String, s종료 As String
    s시작 = Format(dMinDB, "yyyy-mm-dd")
    s종료 = Format(dMaxDB, "yyyy-mm-dd")

    s시작 = InputBox("보고서 시작일을 입력하세요 (yyyy-mm-dd)" & vbCrLf & "DB 기간: " & s시작 & " ~ " & s종료, "시작일 입력", s시작)
    If Trim(s시작) = "" Then Exit Sub
    On Error Resume Next
    시작일 = CDate(s시작)
    If Err.Number <> 0 Then MsgBox "시작일 형식이 잘못되었습니다. (yyyy-mm-dd)", vbCritical, "보고서 출력": Exit Sub
    On Error GoTo 0

    s종료 = InputBox("보고서 종료일을 입력하세요 (yyyy-mm-dd)", "종료일 입력", s종료)
    If Trim(s종료) = "" Then Exit Sub
    On Error Resume Next
    종료일 = CDate(s종료)
    If Err.Number <> 0 Then MsgBox "종료일 형식이 잘못되었습니다. (yyyy-mm-dd)", vbCritical, "보고서 출력": Exit Sub
    On Error GoTo 0

    If 시작일 > 종료일 Then
        MsgBox "시작일이 종료일보다 늦습니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim 일수 As Long: 일수 = CLng(종료일 - 시작일) + 1
    If 일수 > MAX_일수 Then
        MsgBox "날짜 범위가 " & MAX_일수 & "일을 초과합니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim outRows As Long: outRows = 0
    Dim r2 As Long
    For r = 1 To UBound(dbFull, 1)
        vDate = dbFull(r, 1)
        If Not IsEmpty(vDate) And IsDate(vDate) Then
            If CDate(vDate) >= 시작일 And CDate(vDate) <= 종료일 Then outRows = outRows + 1
        End If
    Next r

    If outRows = 0 Then
        MsgBox "입력한 기간(" & Format(시작일, "yyyy-mm-dd") & " ~ " & Format(종료일, "yyyy-mm-dd") & ")에 해당하는 데이터가 없습니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim arrDB As Variant
    ReDim arrDB(1 To outRows, 1 To UBound(dbFull, 2))
    outRows = 0
    For r = 1 To UBound(dbFull, 1)
        vDate = dbFull(r, 1)
        If Not IsEmpty(vDate) And IsDate(vDate) Then
            If CDate(vDate) >= 시작일 And CDate(vDate) <= 종료일 Then
                outRows = outRows + 1
                For r2 = 1 To UBound(dbFull, 2)
                    arrDB(outRows, r2) = dbFull(r, r2)
                Next r2
            End If
        End If
    Next r

    Dim ws재고 As Worksheet, ws원단 As Worksheet
    On Error Resume Next
    Set ws재고 = ThisWorkbook.Worksheets("부자재관리")
    Set ws원단 = ThisWorkbook.Worksheets("원단관리")
    On Error GoTo 0
    Call 보고서_출력_내부(arrDB, 시작일, 종료일, ws재고, ws원단, Empty, Empty, Empty, Empty, 0)
    Exit Sub
ErrRead:
    MsgBox "데이터 읽기 오류: " & Err.Description, vbCritical, "보고서 출력"
End Sub

' ============================================================
' 내부 보고서 생성 (Module3 검수테스트에서 호출 가능)
' 일별날짜,일별DB행,일별성공,일별오류,총일수 가 유효하면 Sheet1=일별처리결과
' SuppressMsg=True, FilePrefix 지정 시 검수테스트용 (MsgBox 생략, 파일명 접두사)
' 보호영역OK, phase2Err 등 전달 시 Sheet1에 Phase2/보호영역 검증 결과 추가
' ============================================================
Public Sub 보고서_출력_내부(arrDB As Variant, 시작일 As Date, 종료일 As Date, _
    Optional ws재고 As Worksheet = Nothing, Optional ws원단 As Worksheet = Nothing, _
    Optional 일별날짜 As Variant, Optional 일별DB행 As Variant, Optional 일별성공 As Variant, _
    Optional 일별오류 As Variant, Optional 총일수 As Long = 0, _
    Optional SuppressMsg As Boolean = False, Optional FilePrefix As String = "검증보고서", _
    Optional 보호영역OK As Boolean = True, Optional phase2Err As Long = 0, Optional phase2Desc As String = "", _
    Optional phase2원단Err As Long = 0, Optional phase2원단Desc As String = "", Optional ws원단없음 As Boolean = False)
    Dim 일수 As Long: 일수 = CLng(종료일 - 시작일) + 1
    If 일수 > MAX_일수 Then
        MsgBox "날짜 범위가 " & MAX_일수 & "일을 초과합니다.", vbCritical, "보고서 출력"
        Exit Sub
    End If

    Dim wbR As Workbook
    Set wbR = Workbooks.Add(xlWBATWorksheet)
    Do While wbR.Worksheets.Count < 6
        wbR.Worksheets.Add After:=wbR.Worksheets(wbR.Worksheets.Count)
    Loop
    wbR.Worksheets(1).Name = "일별처리결과"
    wbR.Worksheets(2).Name = "주차별PA검증"
    wbR.Worksheets(3).Name = "월별PA검증"
    wbR.Worksheets(4).Name = "부자재관리검증"
    wbR.Worksheets(5).Name = "원단관리검증"
    wbR.Worksheets(6).Name = "설비별_원지_기준소요량"

    Dim wsR As Worksheet
    For Each wsR In wbR.Worksheets
        wsR.Cells.Font.Name = "Microsoft GothicNeo"
        wsR.Cells.Font.Size = 9
    Next wsR

    If 총일수 > 0 And Not IsEmpty(일별날짜) And IsArray(일별날짜) And Not IsEmpty(일별DB행) And IsArray(일별DB행) Then
        M23_보고서_일별 wbR.Worksheets(1), 일별날짜, 일별DB행, 일별성공, 일별오류, 총일수
        Dim rPh As Long: rPh = 4 + 총일수 + 3
        With wbR.Worksheets(1)
            .Cells(rPh - 1, 1).Value = "보호영역(T03:G30/T10:G87 수식)"
            .Cells(rPh - 1, 1).Font.Bold = True
            If 보호영역OK Then
                .Cells(rPh - 1, 2).Value = "PASS"
                .Cells(rPh - 1, 2).Font.color = RGB(0, 128, 0)
            Else
                .Cells(rPh - 1, 2).Value = "FAIL (수식 덮어쓰기됨)"
                .Cells(rPh - 1, 2).Font.color = vbRed
            End If
            .Cells(rPh, 1).Value = "Phase 2: 부자재관리 데이터업데이트"
            .Cells(rPh, 1).Font.Bold = True
            If phase2Err = 0 Then
                .Cells(rPh, 2).Value = "PASS"
                .Cells(rPh, 2).Font.color = RGB(0, 128, 0)
            Else
                .Cells(rPh, 2).Value = "FAIL (Err " & phase2Err & ")"
                .Cells(rPh, 2).Font.color = vbRed
                .Cells(rPh, 3).Value = phase2Desc
            End If
            .Cells(rPh + 1, 1).Value = "Phase 2: 원단관리 데이터업데이트"
            .Cells(rPh + 1, 1).Font.Bold = True
            If phase2원단Err <> 0 Then
                .Cells(rPh + 1, 2).Value = "FAIL (Err " & phase2원단Err & ")"
                .Cells(rPh + 1, 2).Font.color = vbRed
                .Cells(rPh + 1, 3).Value = phase2원단Desc
            ElseIf ws원단없음 Then
                .Cells(rPh + 1, 2).Value = "SKIP(시트없음)"
                .Cells(rPh + 1, 2).Font.color = RGB(0, 128, 0)
            Else
                .Cells(rPh + 1, 2).Value = "PASS"
                .Cells(rPh + 1, 2).Font.color = RGB(0, 128, 0)
            End If
        End With
    Else
        M23_보고서_데이터범위요약 wbR.Worksheets(1), arrDB, 시작일, 종료일
    End If

    Dim paIdx As Object: Set paIdx = M23_PA_인덱스_빌드(arrDB)
    Application.StatusBar = "주차별PA 검증..."
    DoEvents
    M23_보고서_주차별PA wbR.Worksheets(2), arrDB, paIdx, 시작일, 종료일
    Application.StatusBar = "월별PA 검증..."
    DoEvents
    M23_보고서_월별PA wbR.Worksheets(3), arrDB, paIdx, 시작일, 종료일
    Application.StatusBar = "부자재관리 검증..."
    DoEvents
    M23_보고서_부자재관리 wbR.Worksheets(4), ws재고, 시작일, 종료일
    Application.StatusBar = "원단관리 검증..."
    DoEvents
    M23_보고서_원단관리 wbR.Worksheets(5), ws원단, 시작일, 종료일
    Application.StatusBar = "설비별 원지·기준소요량..."
    DoEvents
    M23_보고서_설비별 wbR.Worksheets(6), arrDB, paIdx, 시작일, 종료일
    Application.StatusBar = False

    Dim 파일경로 As String
    파일경로 = ThisWorkbook.Path & "\" & FilePrefix & "_" & Format(Now, "yyyymmdd_HHmmss") & ".xlsx"
    On Error Resume Next
    wbR.SaveAs 파일경로, xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        파일경로 = Environ("TEMP") & "\" & FilePrefix & "_" & Format(Now, "yyyymmdd_HHmmss") & ".xlsx"
        Err.Clear
        wbR.SaveAs 파일경로, xlOpenXMLWorkbook
    End If
    On Error GoTo 0
    g_보고서_마지막저장경로 = 파일경로
    If Not SuppressMsg Then MsgBox "보고서 생성 완료!" & vbCrLf & vbCrLf & "저장: " & 파일경로, vbInformation, "보고서 출력"
    wbR.Activate
    wbR.Worksheets(1).Activate
    Dim ws As Worksheet
    For Each ws In wbR.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws
    wbR.Worksheets(1).Activate
End Sub

Private Sub M23_보고서_데이터범위요약(ws As Worksheet, arr As Variant, 시작일 As Date, 종료일 As Date)
    ws.Cells(1, 1).Value = "데이터범위 요약"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Cells(2, 1).Value = "기간: " & Format(시작일, "yyyy-mm-dd") & " ~ " & Format(종료일, "yyyy-mm-dd")
    Dim n As Long: n = 0
    If IsArray(arr) Then n = UBound(arr, 1)
    ws.Cells(4, 1).Value = "선택 행수": ws.Cells(4, 2).Value = n
    ws.Cells(5, 1).Value = "날짜 범위": ws.Cells(5, 2).Value = Format(시작일, "yyyy-mm-dd") & " ~ " & Format(종료일, "yyyy-mm-dd")
    ws.Cells(6, 1).Value = "일수": ws.Cells(6, 2).Value = CLng(종료일 - 시작일) + 1
    ws.Columns("A").ColumnWidth = 14: ws.Columns("B").ColumnWidth = 30
End Sub

Private Function M23_PA_인덱스_빌드(arr As Variant) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Set M23_PA_인덱스_빌드 = idx
    If Not IsArray(arr) Then Exit Function
    Dim r As Long, 키 As String, 구분 As String, 유형 As String, 순번 As Variant, vDate As Variant
    Dim maxR As Long: maxR = UBound(arr, 1)
    For r = 1 To maxR
        If UBound(arr, 2) < 5 Then GoTo NextPA23
        구분 = Trim(CStr(arr(r, 3))): 유형 = Trim(CStr(arr(r, 4))): 순번 = arr(r, 5)
        If 구분 <> "생산일보" Or 유형 <> "ProductionAnalysis" Then GoTo NextPA23
        If Not IsNumeric(순번) Then GoTo NextPA23
        If CLng(순번) <> 1 Then GoTo NextPA23
        vDate = arr(r, 1)
        If IsEmpty(vDate) Or Not IsDate(vDate) Then GoTo NextPA23
        키 = Format(CDate(vDate), "YYYYMMDD")
        If Not idx.exists(키) Then idx.Add 키, r
NextPA23:
    Next r
End Function

Private Function M23_PA_합산(arr As Variant, paIdx As Object, 열번호 As Long, d시작 As Date, d끝 As Date, ByRef out일수 As Long) As Double
    Dim 합계 As Double: 합계 = 0: out일수 = 0
    Dim d As Date, 키 As String, r As Long
    If d끝 < d시작 Then Exit Function
    If CLng(d끝 - d시작) + 1 > MAX_일수 Then Exit Function
    For d = d시작 To d끝
        키 = Format(d, "YYYYMMDD")
        If paIdx.exists(키) Then
            r = CLng(paIdx(키))
            If UBound(arr, 2) >= 열번호 Then 합계 = 합계 + M23_SDbl(arr(r, 열번호))
            out일수 = out일수 + 1
        End If
    Next d
    M23_PA_합산 = 합계
End Function

Private Function M23_PA_마지막값(arr As Variant, paIdx As Object, 열번호 As Long, d시작 As Date, d끝 As Date) As Double
    M23_PA_마지막값 = 0
    Dim d As Date, 키 As String
    If CLng(d끝 - d시작) + 1 > MAX_일수 Then Exit Function
    For d = d끝 To d시작 Step -1
        키 = Format(d, "YYYYMMDD")
        If paIdx.exists(키) Then
            If UBound(arr, 2) >= 열번호 Then M23_PA_마지막값 = M23_SDbl(arr(CLng(paIdx(키)), 열번호))
            Exit Function
        End If
    Next d
End Function

Private Sub M23_보고서_일별(ws As Worksheet, 날짜 As Variant, DB행 As Variant, 성공 As Variant, 오류 As Variant, n As Long)
    If n <= 0 Or n > MAX_일수 Then Exit Sub
    ws.Cells(1, 1).Value = "일별 처리 결과"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Cells(2, 1).Value = "기간: " & Format(날짜(1), "yyyy-mm-dd") & " ~ " & Format(날짜(n), "yyyy-mm-dd") & " (" & n & "일)"
    Dim 헤더 As Variant: 헤더 = Array("#", "날짜", "요일", "월-주차", "DB 누적행", "결과", "오류내용")
    Dim i As Long
    For i = 0 To 6: ws.Cells(4, i + 1).Value = 헤더(i): Next i
    Call M23_헤더서식(ws.Range("A4:G4"))
    Dim out() As Variant: ReDim out(1 To n, 1 To 7)
    Dim 성공수 As Long: 성공수 = 0
    For i = 1 To n
        out(i, 1) = i
        out(i, 2) = 날짜(i)
        out(i, 3) = Choose(Weekday(날짜(i), vbMonday), "월", "화", "수", "목", "금", "토", "일")
        out(i, 4) = M23_주차키(날짜(i))
        out(i, 5) = DB행(i)
        If IsArray(성공) And UBound(성공) >= i Then
            If 성공(i) Then out(i, 6) = "PASS": 성공수 = 성공수 + 1 Else out(i, 6) = "FAIL"
        Else
            out(i, 6) = "PASS": 성공수 = 성공수 + 1
        End If
        If IsArray(오류) And UBound(오류) >= i And Not IsEmpty(오류(i)) Then out(i, 7) = CStr(오류(i))
    Next i
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + n, 7)).Value = out
    Dim r As Long
    For i = 1 To n
        r = 4 + i
        ws.Cells(r, 2).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, 5).NumberFormat = "#,##0"
        If out(i, 6) = "PASS" Then ws.Cells(r, 6).Font.color = RGB(0, 128, 0) Else ws.Cells(r, 6).Font.color = vbRed
        If i Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).Interior.color = RGB(242, 242, 242)
    Next i
    r = 4 + n + 2
    ws.Cells(r, 1).Value = "요약": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "PASS " & 성공수 & " / 전체 " & n
    ws.Cells(r, 5).Value = DB행(n): ws.Cells(r, 5).NumberFormat = "#,##0": ws.Cells(r, 5).Font.Bold = True
    ws.Columns("A").ColumnWidth = 5: ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 5: ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 8
    ws.Columns("G").ColumnWidth = 40
End Sub

Private Sub M23_보고서_주차별PA(ws As Worksheet, arr As Variant, paIdx As Object, 시작일 As Date, 종료일 As Date)
    If Not IsArray(arr) Then ws.Cells(1, 1).Value = "DB 데이터 없음": Exit Sub
    ws.Cells(1, 1).Value = "주차별 PA 누적 검증"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Cells(2, 1).Value = "검증: 주차 내 일별 합산 vs 마지막일 주간누계 비교 (허용오차: " & PA_허용오차 & ")"
    Dim hdr As Variant: hdr = Array("검증항목", "주차", "일수", "일별합산", "주간누계(DB)", "주간평균(계산)", "주간평균(DB)", "차이", "판정", "기간", "1일차", "2일차", "3일차", "4일차", "5일차", "6일차", "7일차")
    Dim i As Long: For i = 0 To 16: ws.Cells(4, i + 1).Value = hdr(i): Next i
    Call M23_헤더서식(ws.Range("A4:Q4"))
    ws.Range("J3:Q3").Merge: ws.Cells(3, 10).Value = "일별 상세"
    Dim 항목명 As Variant: 항목명 = Array("KT원지투입", "KT기준소요량", "FT원지투입", "FT기준소요량", "JRT원지투입", "JRT기준소요량", "BT원지투입", "BT기준소요량", "PRE원지투입", "PRE기준소요량")
    Dim 일별열 As Variant: 일별열 = Array(28, 31, 40, 123, 43, 124, 46, 125, 49, 126)
    Dim 주간열 As Variant: 주간열 = Array(34, 103, 54, 107, 58, 111, 64, 115, 68, 119)
    Dim 주간평균열 As Variant: 주간평균열 = Array(73, 105, 79, 109, 85, 113, 91, 117, 97, 121)
    Dim 주차순서 As Object: Set 주차순서 = CreateObject("Scripting.Dictionary")
    Dim d As Date, wk As String
    For d = 시작일 To 종료일
        wk = M23_주차키(d)
        If Not 주차순서.exists(wk) Then 주차순서.Add wk, Array(M23_주차_첫날(d), M23_주차_마지막일(d))
    Next d
    Dim 항목수 As Long: 항목수 = 10
    Dim 총행 As Long: 총행 = 항목수 * 주차순서.Count
    If 총행 = 0 Then ws.Cells(5, 1).Value = "주차 데이터 없음": Exit Sub
    Dim outArr() As Variant: ReDim outArr(1 To 총행, 1 To 17)
    Dim 합격배열() As Boolean: ReDim 합격배열(1 To 총행)
    Dim rw As Long, 합격수 As Long, 총검증 As Long
    rw = 0: 합격수 = 0: 총검증 = 0
    Dim wKeys As Variant: wKeys = 주차순서.keys
    Dim hi As Long, wi As Long, 주시작 As Date, 주끝 As Date, 일수 As Long, 합산 As Double, DB누계 As Double, DB주평균 As Double, 차이 As Double, dayOfs As Long, dayDate As Date, dayKey As String
    For hi = 0 To 9
        For wi = 0 To 주차순서.Count - 1
            주시작 = CDate(주차순서(wKeys(wi))(0))
            주끝 = CDate(주차순서(wKeys(wi))(1))
            If 주끝 > 종료일 Then 주끝 = 종료일
            합산 = M23_PA_합산(arr, paIdx, CLng(일별열(hi)), 주시작, 주끝, 일수)
            DB누계 = M23_PA_마지막값(arr, paIdx, CLng(주간열(hi)), 주시작, 주끝)
            DB주평균 = M23_PA_마지막값(arr, paIdx, CLng(주간평균열(hi)), 주시작, 주끝)
            차이 = Abs(합산 - DB누계)
            rw = rw + 1
            outArr(rw, 1) = CStr(항목명(hi)): outArr(rw, 2) = wKeys(wi): outArr(rw, 3) = 일수
            outArr(rw, 4) = Round(합산, 1): outArr(rw, 5) = Round(DB누계, 1): outArr(rw, 6) = Round(IIf(일수 > 0, 합산 / 일수, 0), 1)
            outArr(rw, 7) = Round(DB주평균, 1): outArr(rw, 8) = Round(차이, 2)
            합격배열(rw) = (차이 <= PA_허용오차) Or (일수 = 0 And DB누계 = 0)
            If 합격배열(rw) Then outArr(rw, 9) = "PASS": 합격수 = 합격수 + 1 Else outArr(rw, 9) = "FAIL"
            총검증 = 총검증 + 1
            outArr(rw, 10) = Format(주시작, "M/D") & "~" & Format(주끝, "M/D")
            For dayOfs = 0 To 6
                dayDate = 주시작 + dayOfs
                If dayDate > 주끝 Then Exit For
                dayKey = Format(dayDate, "YYYYMMDD")
                If paIdx.exists(dayKey) And UBound(arr, 2) >= CLng(일별열(hi)) Then
                    outArr(rw, 11 + dayOfs) = Round(M23_SDbl(arr(CLng(paIdx(dayKey)), CLng(일별열(hi)))), 1)
                End If
            Next dayOfs
        Next wi
    Next hi
    If rw > 0 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + rw, 17)).Value = outArr
        Dim r As Long
        For r = 1 To rw
            ws.Range(ws.Cells(4 + r, 4), ws.Cells(4 + r, 8)).NumberFormat = "#,##0.0"
            ws.Range(ws.Cells(4 + r, 11), ws.Cells(4 + r, 17)).NumberFormat = "#,##0.0"
            If 합격배열(r) Then ws.Cells(4 + r, 9).Font.color = RGB(0, 128, 0)
            If (4 + r) Mod 2 = 0 And 합격배열(r) Then ws.Range(ws.Cells(4 + r, 1), ws.Cells(4 + r, 17)).Interior.color = RGB(242, 242, 242)
            If Not 합격배열(r) Then ws.Cells(4 + r, 9).Font.color = vbRed: ws.Range(ws.Cells(4 + r, 1), ws.Cells(4 + r, 17)).Interior.color = RGB(255, 230, 230)
        Next r
    End If
    Dim 전체날수 As Long: 전체날수 = CLng(종료일 - 시작일) + 1
    Dim s설비 As Variant: s설비 = Array("KT", "FT", "JRT", "BT", "PRE")
    Dim s열 As Variant: s열 = Array(28, 40, 43, 46, 49)
    Dim 설비존재행 As Long: 설비존재행 = 5 + rw + 2
    ws.Cells(설비존재행 - 1, 1).Value = "설비별 일별 데이터 존재 확인": ws.Cells(설비존재행 - 1, 1).Font.Bold = True
    Dim shdr As Variant: shdr = Array("설비", "데이터있는날", "전체날", "판정")
    For i = 0 To 3: ws.Cells(설비존재행, i + 1).Value = shdr(i): Next i
    Call M23_헤더서식(ws.Range(ws.Cells(설비존재행, 1), ws.Cells(설비존재행, 4)))
    Dim si As Long, 있는날 As Long, sKey As String
    For si = 0 To 4
        있는날 = 0
        For d = 시작일 To 종료일
            sKey = Format(d, "YYYYMMDD")
            If paIdx.exists(sKey) And UBound(arr, 2) >= s열(si) Then
                If M23_SDbl(arr(CLng(paIdx(sKey)), s열(si))) <> 0 Then 있는날 = 있는날 + 1
            End If
        Next d
        ws.Cells(설비존재행 + 1 + si, 1).Value = s설비(si)
        ws.Cells(설비존재행 + 1 + si, 2).Value = 있는날
        ws.Cells(설비존재행 + 1 + si, 3).Value = 전체날수
        If 있는날 >= 전체날수 Then
            ws.Cells(설비존재행 + 1 + si, 4).Value = "PASS"
            ws.Cells(설비존재행 + 1 + si, 4).Font.color = RGB(0, 128, 0)
        ElseIf 있는날 > 0 Then
            ws.Cells(설비존재행 + 1 + si, 4).Value = "PARTIAL (" & 있는날 & "/" & 전체날수 & ")"
            ws.Cells(설비존재행 + 1 + si, 4).Font.color = RGB(200, 120, 0)
        Else
            ws.Cells(설비존재행 + 1 + si, 4).Value = "FAIL (없음)"
            ws.Cells(설비존재행 + 1 + si, 4).Font.color = vbRed
        End If
    Next si
    Dim sumRow As Long: sumRow = 설비존재행 + 6
    ws.Cells(sumRow, 1).Value = "요약": ws.Cells(sumRow, 1).Font.Bold = True
    ws.Cells(sumRow, 2).Value = "PASS " & 합격수 & " / FAIL " & (총검증 - 합격수) & " / 전체 " & 총검증
    If 합격수 = 총검증 Then
        ws.Cells(sumRow, 7).Value = "ALL PASS"
        ws.Cells(sumRow, 7).Font.color = RGB(0, 128, 0)
    Else
        ws.Cells(sumRow, 7).Value = "FAIL EXISTS"
        ws.Cells(sumRow, 7).Font.color = vbRed
    End If
    ws.Cells(sumRow, 7).Font.Bold = True
End Sub

Private Sub M23_보고서_월별PA(ws As Worksheet, arr As Variant, paIdx As Object, 시작일 As Date, 종료일 As Date)
    If Not IsArray(arr) Then ws.Cells(1, 1).Value = "DB 데이터 없음": Exit Sub
    ws.Cells(1, 1).Value = "월별 PA 누적 검증"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Cells(2, 1).Value = "검증: 월 내 일별 합산 vs 마지막일 월간누계 비교 (허용오차: " & PA_허용오차 & ")"
    Dim hdr As Variant: hdr = Array("검증항목", "월", "일수", "일별합산", "월간누계(DB)", "차이", "판정")
    Dim i As Long: For i = 0 To 6: ws.Cells(4, i + 1).Value = hdr(i): Next i
    Call M23_헤더서식(ws.Range("A4:G4"))
    Dim 항목명 As Variant: 항목명 = Array("KT원지투입", "KT기준소요량", "FT원지투입", "FT기준소요량", "JRT원지투입", "JRT기준소요량", "BT원지투입", "BT기준소요량", "PRE원지투입", "PRE기준소요량")
    Dim 일별열 As Variant: 일별열 = Array(28, 31, 40, 123, 43, 124, 46, 125, 49, 126)
    Dim 월간열 As Variant: 월간열 = Array(37, 104, 56, 108, 60, 112, 66, 116, 70, 120)
    Dim 월목록 As Object: Set 월목록 = CreateObject("Scripting.Dictionary")
    Dim d As Date, mk As String
    For d = 시작일 To 종료일
        mk = M23_월키(d)
        If Not 월목록.exists(mk) Then 월목록.Add mk, Array(DateSerial(Year(d), Month(d), 1), DateSerial(Year(d), Month(d) + 1, 0))
    Next d
    Dim 총행 As Long: 총행 = 10 * 월목록.Count
    Dim outArr() As Variant: ReDim outArr(1 To 총행, 1 To 7)
    Dim 합격배열() As Boolean: ReDim 합격배열(1 To 총행)
    Dim rw As Long, 합격수 As Long, 총검증 As Long
    rw = 0: 합격수 = 0: 총검증 = 0
    Dim mKeys As Variant: mKeys = 월목록.keys
    Dim hi As Long, mi As Long, m시작 As Date, m끝 As Date, 일수 As Long, 합산 As Double, DB누계 As Double, 차이 As Double
    For hi = 0 To 9
        For mi = 0 To 월목록.Count - 1
            m시작 = CDate(월목록(mKeys(mi))(0))
            m끝 = CDate(월목록(mKeys(mi))(1))
            If m끝 > 종료일 Then m끝 = 종료일
            합산 = M23_PA_합산(arr, paIdx, CLng(일별열(hi)), m시작, m끝, 일수)
            DB누계 = M23_PA_마지막값(arr, paIdx, CLng(월간열(hi)), m시작, m끝)
            차이 = Abs(합산 - DB누계)
            rw = rw + 1
            outArr(rw, 1) = CStr(항목명(hi)): outArr(rw, 2) = mKeys(mi): outArr(rw, 3) = 일수
            outArr(rw, 4) = Round(합산, 1): outArr(rw, 5) = Round(DB누계, 1): outArr(rw, 6) = Round(차이, 2)
            합격배열(rw) = (차이 <= PA_허용오차) Or (일수 = 0 And DB누계 = 0)
            If 합격배열(rw) Then outArr(rw, 7) = "PASS": 합격수 = 합격수 + 1 Else outArr(rw, 7) = "FAIL"
            총검증 = 총검증 + 1
        Next mi
    Next hi
    If rw > 0 Then
        ws.Range(ws.Cells(5, 1), ws.Cells(4 + rw, 7)).Value = outArr
        Dim r As Long
        For r = 1 To rw
            ws.Range(ws.Cells(4 + r, 4), ws.Cells(4 + r, 6)).NumberFormat = "#,##0.0"
            If 합격배열(r) Then ws.Cells(4 + r, 7).Font.color = RGB(0, 128, 0)
            If (4 + r) Mod 2 = 0 And 합격배열(r) Then ws.Range(ws.Cells(4 + r, 1), ws.Cells(4 + r, 7)).Interior.color = RGB(242, 242, 242)
            If Not 합격배열(r) Then ws.Cells(4 + r, 7).Font.color = vbRed: ws.Range(ws.Cells(4 + r, 1), ws.Cells(4 + r, 7)).Interior.color = RGB(255, 230, 230)
        Next r
    End If
    Dim sumRow As Long: sumRow = 5 + rw + 2
    ws.Cells(sumRow, 1).Value = "요약": ws.Cells(sumRow, 1).Font.Bold = True
    ws.Cells(sumRow, 2).Value = "PASS " & 합격수 & " / FAIL " & (총검증 - 합격수) & " / 전체 " & 총검증
    If 합격수 = 총검증 Then
        ws.Cells(sumRow, 7).Value = "ALL PASS"
        ws.Cells(sumRow, 7).Font.color = RGB(0, 128, 0)
    Else
        ws.Cells(sumRow, 7).Value = "FAIL EXISTS"
        ws.Cells(sumRow, 7).Font.color = vbRed
    End If
    ws.Cells(sumRow, 7).Font.Bold = True
    ws.Columns("A").ColumnWidth = 18: ws.Columns("B").ColumnWidth = 7: ws.Columns("C").ColumnWidth = 6
    ws.Columns("D").ColumnWidth = 14: ws.Columns("E").ColumnWidth = 16: ws.Columns("F").ColumnWidth = 10: ws.Columns("G").ColumnWidth = 10
End Sub

Private Sub M23_보고서_부자재관리(ws As Worksheet, ws재고 As Worksheet, 시작일 As Date, 종료일 As Date)
    ws.Cells(1, 1).Value = "부자재관리 시트 검증"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    If ws재고 Is Nothing Then ws.Cells(3, 1).Value = "부자재관리 시트가 존재하지 않습니다.": ws.Cells(3, 1).Font.color = vbRed: Exit Sub
    Dim COL_구분 As Long: COL_구분 = 9
    Dim COL_월별시작 As Long: COL_월별시작 = 10
    Dim COL_기초재고 As Long: COL_기초재고 = 5
    Dim START_ROW As Long: START_ROW = 4
    Dim ROWS_PER_ITEM As Long: ROWS_PER_ITEM = 6
    Dim OFFSET_입고1 As Long, OFFSET_사용1 As Long, OFFSET_재고1 As Long, OFFSET_입고2 As Long, OFFSET_사용2 As Long, OFFSET_재고2 As Long
    OFFSET_입고1 = 0: OFFSET_사용1 = 1: OFFSET_재고1 = 2
    OFFSET_입고2 = 3: OFFSET_사용2 = 4: OFFSET_재고2 = 5
    Dim rw As Long: rw = 3
    ws.Cells(rw, 1).Value = "1. 기본 구조": ws.Cells(rw, 1).Font.Bold = True: rw = rw + 1
    Dim hdr1 As Variant: hdr1 = Array("검증항목", "결과값", "판정")
    Dim i As Long: For i = 0 To 2: ws.Cells(rw, i + 1).Value = hdr1(i): Next i
    Call M23_헤더서식(ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 3))): rw = rw + 1
    Dim 마지막행 As Long
    마지막행 = ws재고.Cells(ws재고.Rows.Count, 1).End(xlUp).Row
    If ws재고.Cells(ws재고.Rows.Count, 2).End(xlUp).Row > 마지막행 Then 마지막행 = ws재고.Cells(ws재고.Rows.Count, 2).End(xlUp).Row
    ws.Cells(rw, 1).Value = "데이터 마지막행": ws.Cells(rw, 2).Value = 마지막행
    ws.Cells(rw, 3).Value = IIf(마지막행 >= START_ROW, "OK", "데이터 없음"): ws.Cells(rw, 3).Font.color = IIf(마지막행 >= START_ROW, RGB(0, 128, 0), vbRed): rw = rw + 1
    Dim 마지막일열 As Long
    마지막일열 = ws재고.Cells(2, ws재고.Columns.Count).End(xlToLeft).Column
    If 마지막일열 < COL_월별시작 Then 마지막일열 = COL_월별시작
    Dim arrInfo As Variant, arrDaily As Variant
    If 마지막행 >= START_ROW Then
        arrInfo = ws재고.Range(ws재고.Cells(START_ROW, 1), ws재고.Cells(마지막행, COL_구분)).Value
        arrDaily = ws재고.Range(ws재고.Cells(START_ROW, COL_월별시작), ws재고.Cells(마지막행, 마지막일열)).Value
    End If
    Dim 그룹수 As Long: 그룹수 = 0
    Dim r As Long
    If 마지막행 >= START_ROW And IsArray(arrInfo) Then
        For r = 1 To UBound(arrInfo, 1) Step ROWS_PER_ITEM
            If Trim(CStr(arrInfo(r, 2))) <> "" Then 그룹수 = 그룹수 + 1
        Next r
    End If
    ws.Cells(rw, 1).Value = "제품 그룹 수": ws.Cells(rw, 2).Value = 그룹수: ws.Cells(rw, 3).Value = IIf(그룹수 > 0, "OK (" & 그룹수 & "개)", "그룹 없음"): rw = rw + 1
    Dim 월열수 As Long: If 마지막일열 >= COL_월별시작 Then 월열수 = 마지막일열 - COL_월별시작 + 1 Else 월열수 = 0
    ws.Cells(rw, 1).Value = "월별 열 수": ws.Cells(rw, 2).Value = 월열수: rw = rw + 2
    ws.Cells(rw, 1).Value = "2. 그룹별 상세 (상위 20개)": ws.Cells(rw, 1).Font.Bold = True: rw = rw + 1
    Dim hdr2 As Variant: hdr2 = Array("#", "발주처(A)", "제품명(B)", "구분(I)", "월별데이터유무", "재고체인판정")
    For i = 0 To 5: ws.Cells(rw, i + 1).Value = hdr2(i): Next i
    Call M23_헤더서식(ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6))): rw = rw + 1
    Dim 그룹idx As Long: 그룹idx = 0
    Dim 체인OK수 As Long, 체인총수 As Long
    체인OK수 = 0: 체인총수 = 0
    Dim maxDailyCols As Long: If IsArray(arrDaily) Then maxDailyCols = UBound(arrDaily, 2) Else maxDailyCols = 0
    Dim 검증끝열 As Long: 검증끝열 = Application.WorksheetFunction.Min(maxDailyCols, 24)
    If 마지막행 >= START_ROW And IsArray(arrInfo) Then
        For r = 1 To UBound(arrInfo, 1) Step ROWS_PER_ITEM
            Dim 제품명 As String: 제품명 = Trim(CStr(arrInfo(r, 2)))
            If 제품명 = "" Then GoTo NextGroup23
            그룹idx = 그룹idx + 1
            If 그룹idx > 20 Then Exit For
            ws.Cells(rw, 1).Value = 그룹idx: ws.Cells(rw, 2).Value = Trim(CStr(arrInfo(r, 1))): ws.Cells(rw, 3).Value = 제품명
            ws.Cells(rw, 4).Value = IIf(UBound(arrInfo, 2) >= COL_구분, Trim(CStr(arrInfo(r, COL_구분))), "")
            Dim 데이터열수 As Long: 데이터열수 = 0
            Dim c As Long, hasVal As Boolean
            If maxDailyCols > 0 And r + OFFSET_재고2 <= UBound(arrDaily, 1) Then
                For c = 1 To 검증끝열
                    hasVal = False
                    If Not IsEmpty(arrDaily(r + OFFSET_사용1, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_재고1, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_사용2, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_재고2, c)) Then hasVal = True
                    If hasVal Then 데이터열수 = 데이터열수 + 1
                Next c
            End If
            ws.Cells(rw, 5).Value = 데이터열수 & "월": ws.Cells(rw, 5).Font.color = IIf(데이터열수 > 0, RGB(0, 128, 0), vbRed)
            Dim 체인OK As Boolean: 체인OK = True
            Dim 기초1 As Double: 기초1 = M23_SDbl(arrInfo(r, COL_기초재고))
            Dim 기초2 As Double: 기초2 = 0
            If r + 3 <= UBound(arrInfo, 1) Then 기초2 = M23_SDbl(arrInfo(r + 3, COL_기초재고))
            Dim 누적1 As Double, 누적2 As Double
            누적1 = 기초1: 누적2 = 기초2
            Dim 체크수 As Long: 체크수 = 0
            If maxDailyCols > 0 And r + OFFSET_재고2 <= UBound(arrDaily, 1) Then
                For c = 1 To 검증끝열
                    Dim 입고1v As Double: 입고1v = M23_SDbl(arrDaily(r + OFFSET_입고1, c))
                    Dim 사용1v As Double: 사용1v = M23_SDbl(arrDaily(r + OFFSET_사용1, c))
                    Dim 재고1v As Double: 재고1v = M23_SDbl(arrDaily(r + OFFSET_재고1, c))
                    Dim 입고2v As Double: 입고2v = M23_SDbl(arrDaily(r + OFFSET_입고2, c))
                    Dim 사용2v As Double: 사용2v = M23_SDbl(arrDaily(r + OFFSET_사용2, c))
                    Dim 재고2v As Double: 재고2v = M23_SDbl(arrDaily(r + OFFSET_재고2, c))
                    If 입고1v <> 0 Or 사용1v <> 0 Or 재고1v <> 0 Or 입고2v <> 0 Or 사용2v <> 0 Or 재고2v <> 0 Then
                        If c = 1 Then
                            누적1 = 기초1 + 입고1v - 사용1v: 누적2 = 기초2 + 입고2v - 사용2v
                        Else
                            누적1 = M23_SDbl(arrDaily(r + OFFSET_재고1, c - 1)) + 입고1v - 사용1v: 누적2 = M23_SDbl(arrDaily(r + OFFSET_재고2, c - 1)) + 입고2v - 사용2v
                        End If
                        체크수 = 체크수 + 1
                        If Abs(누적1 - 재고1v) > 1 Or Abs(누적2 - 재고2v) > 1 Then 체인OK = False
                    End If
                Next c
            End If
            체인총수 = 체인총수 + 1
            If 체인OK And 체크수 > 0 Then
                ws.Cells(rw, 6).Value = "PASS (" & 체크수 & "월)"
                ws.Cells(rw, 6).Font.color = RGB(0, 128, 0)
                체인OK수 = 체인OK수 + 1
            ElseIf 체크수 = 0 Then
                ws.Cells(rw, 6).Value = "데이터없음"
                ws.Cells(rw, 6).Font.color = RGB(150, 150, 150)
            Else
                ws.Cells(rw, 6).Value = "FAIL"
                ws.Cells(rw, 6).Font.color = vbRed
                ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6)).Interior.color = RGB(255, 230, 230)
            End If
            If rw Mod 2 = 0 And (체인OK Or 체크수 = 0) Then ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6)).Interior.color = RGB(242, 242, 242)
            rw = rw + 1
NextGroup23:
        Next r
    End If
    rw = rw + 1
    ws.Cells(rw, 1).Value = "요약": ws.Cells(rw, 1).Font.Bold = True
    ws.Cells(rw, 2).Value = "그룹: " & 그룹수 & "개, 재고체인: PASS " & 체인OK수 & "/" & 체인총수
    If 체인OK수 = 체인총수 And 체인총수 > 0 Then
        ws.Cells(rw, 6).Value = "ALL PASS"
        ws.Cells(rw, 6).Font.color = RGB(0, 128, 0)
    ElseIf 체인총수 = 0 Then
        ws.Cells(rw, 6).Value = "검증불가"
    Else
        ws.Cells(rw, 6).Value = "FAIL EXISTS"
        ws.Cells(rw, 6).Font.color = vbRed
    End If
    ws.Cells(rw, 6).Font.Bold = True
    ws.Columns("A").ColumnWidth = 5: ws.Columns("B").ColumnWidth = 14: ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 12: ws.Columns("E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 16
End Sub

Private Sub M23_보고서_원단관리(ws As Worksheet, ws원단 As Worksheet, 시작일 As Date, 종료일 As Date)
    ws.Cells(1, 1).Value = "원단관리 시트 검증"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    If ws원단 Is Nothing Then ws.Cells(3, 1).Value = "원단관리 시트가 존재하지 않습니다.": ws.Cells(3, 1).Font.color = vbRed: Exit Sub
    Dim COL_구분 As Long, COL_월별시작 As Long, COL_기초재고 As Long
    COL_구분 = 9: COL_월별시작 = 10: COL_기초재고 = 5
    Dim START_ROW As Long, ROWS_PER_ITEM As Long
    START_ROW = 4: ROWS_PER_ITEM = 6
    Dim OFFSET_입고1 As Long, OFFSET_사용1 As Long, OFFSET_재고1 As Long, OFFSET_입고2 As Long, OFFSET_사용2 As Long, OFFSET_재고2 As Long
    OFFSET_입고1 = 0: OFFSET_사용1 = 1: OFFSET_재고1 = 2
    OFFSET_입고2 = 3: OFFSET_사용2 = 4: OFFSET_재고2 = 5
    Dim rw As Long: rw = 3
    ws.Cells(rw, 1).Value = "1. 기본 구조": ws.Cells(rw, 1).Font.Bold = True: rw = rw + 1
    Dim hdr1 As Variant: hdr1 = Array("검증항목", "결과값", "판정")
    Dim i As Long: For i = 0 To 2: ws.Cells(rw, i + 1).Value = hdr1(i): Next i
    Call M23_헤더서식(ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 3))): rw = rw + 1
    Dim 마지막행 As Long
    마지막행 = ws원단.Cells(ws원단.Rows.Count, 1).End(xlUp).Row
    If ws원단.Cells(ws원단.Rows.Count, 2).End(xlUp).Row > 마지막행 Then 마지막행 = ws원단.Cells(ws원단.Rows.Count, 2).End(xlUp).Row
    ws.Cells(rw, 1).Value = "데이터 마지막행": ws.Cells(rw, 2).Value = 마지막행: ws.Cells(rw, 3).Value = IIf(마지막행 >= START_ROW, "OK", "데이터 없음"): rw = rw + 1
    Dim 마지막일열 As Long: 마지막일열 = ws원단.Cells(2, ws원단.Columns.Count).End(xlToLeft).Column
    If 마지막일열 < COL_월별시작 Then 마지막일열 = COL_월별시작
    Dim arrInfo As Variant, arrDaily As Variant
    If 마지막행 >= START_ROW Then
        arrInfo = ws원단.Range(ws원단.Cells(START_ROW, 1), ws원단.Cells(마지막행, COL_구분)).Value
        arrDaily = ws원단.Range(ws원단.Cells(START_ROW, COL_월별시작), ws원단.Cells(마지막행, 마지막일열)).Value
    End If
    Dim 그룹수 As Long: 그룹수 = 0
    Dim r As Long
    If 마지막행 >= START_ROW And IsArray(arrInfo) Then
        For r = 1 To UBound(arrInfo, 1) Step ROWS_PER_ITEM
            If Trim(CStr(arrInfo(r, 2))) <> "" Then 그룹수 = 그룹수 + 1
        Next r
    End If
    ws.Cells(rw, 1).Value = "원단 그룹 수": ws.Cells(rw, 2).Value = 그룹수: ws.Cells(rw, 3).Value = IIf(그룹수 > 0, "OK (" & 그룹수 & "개)", "그룹 없음"): rw = rw + 1
    Dim 월열수 As Long: If 마지막일열 >= COL_월별시작 Then 월열수 = 마지막일열 - COL_월별시작 + 1 Else 월열수 = 0
    ws.Cells(rw, 1).Value = "월별 열 수": ws.Cells(rw, 2).Value = 월열수: rw = rw + 2
    ws.Cells(rw, 1).Value = "2. 그룹별 상세 (상위 20개)": ws.Cells(rw, 1).Font.Bold = True: rw = rw + 1
    Dim hdr2 As Variant: hdr2 = Array("#", "Co(A)", "원단명(B)", "구분(I)", "월별데이터유무", "재고체인판정")
    For i = 0 To 5: ws.Cells(rw, i + 1).Value = hdr2(i): Next i
    Call M23_헤더서식(ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6))): rw = rw + 1
    Dim 그룹idx As Long, 체인OK수 As Long, 체인총수 As Long
    그룹idx = 0: 체인OK수 = 0: 체인총수 = 0
    Dim maxDailyCols As Long: If IsArray(arrDaily) Then maxDailyCols = UBound(arrDaily, 2) Else maxDailyCols = 0
    Dim 검증끝열 As Long: 검증끝열 = Application.WorksheetFunction.Min(maxDailyCols, 24)
    If 마지막행 >= START_ROW And IsArray(arrInfo) Then
        For r = 1 To UBound(arrInfo, 1) Step ROWS_PER_ITEM
            Dim 원단명 As String: 원단명 = Trim(CStr(arrInfo(r, 2)))
            If 원단명 = "" Then GoTo Next원단23
            그룹idx = 그룹idx + 1
            If 그룹idx > 20 Then Exit For
            ws.Cells(rw, 1).Value = 그룹idx: ws.Cells(rw, 2).Value = Trim(CStr(arrInfo(r, 1))): ws.Cells(rw, 3).Value = 원단명
            ws.Cells(rw, 4).Value = IIf(UBound(arrInfo, 2) >= COL_구분, Trim(CStr(arrInfo(r, COL_구분))), "")
            Dim 데이터열수 As Long: 데이터열수 = 0
            Dim c As Long, hasVal As Boolean
            If maxDailyCols > 0 And r + OFFSET_재고2 <= UBound(arrDaily, 1) Then
                For c = 1 To 검증끝열
                    hasVal = False
                    If Not IsEmpty(arrDaily(r + OFFSET_사용1, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_재고1, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_사용2, c)) Then hasVal = True
                    If Not hasVal And Not IsEmpty(arrDaily(r + OFFSET_재고2, c)) Then hasVal = True
                    If hasVal Then 데이터열수 = 데이터열수 + 1
                Next c
            End If
            ws.Cells(rw, 5).Value = 데이터열수 & "월": ws.Cells(rw, 5).Font.color = IIf(데이터열수 > 0, RGB(0, 128, 0), vbRed)
            Dim 체인OK As Boolean: 체인OK = True
            Dim 기초1 As Double: 기초1 = M23_SDbl(arrInfo(r, COL_기초재고))
            Dim 기초2 As Double: 기초2 = 0
            If r + 3 <= UBound(arrInfo, 1) Then 기초2 = M23_SDbl(arrInfo(r + 3, COL_기초재고))
            Dim 누적1 As Double, 누적2 As Double
            누적1 = 기초1: 누적2 = 기초2
            Dim 체크수 As Long: 체크수 = 0
            If maxDailyCols > 0 And r + OFFSET_재고2 <= UBound(arrDaily, 1) Then
                For c = 1 To 검증끝열
                    Dim 입고1v As Double: 입고1v = M23_SDbl(arrDaily(r + OFFSET_입고1, c))
                    Dim 사용1v As Double: 사용1v = M23_SDbl(arrDaily(r + OFFSET_사용1, c))
                    Dim 재고1v As Double: 재고1v = M23_SDbl(arrDaily(r + OFFSET_재고1, c))
                    Dim 입고2v As Double: 입고2v = M23_SDbl(arrDaily(r + OFFSET_입고2, c))
                    Dim 사용2v As Double: 사용2v = M23_SDbl(arrDaily(r + OFFSET_사용2, c))
                    Dim 재고2v As Double: 재고2v = M23_SDbl(arrDaily(r + OFFSET_재고2, c))
                    If 입고1v <> 0 Or 사용1v <> 0 Or 재고1v <> 0 Or 입고2v <> 0 Or 사용2v <> 0 Or 재고2v <> 0 Then
                        If c = 1 Then
                            누적1 = 기초1 + 입고1v - 사용1v: 누적2 = 기초2 + 입고2v - 사용2v
                        Else
                            누적1 = M23_SDbl(arrDaily(r + OFFSET_재고1, c - 1)) + 입고1v - 사용1v: 누적2 = M23_SDbl(arrDaily(r + OFFSET_재고2, c - 1)) + 입고2v - 사용2v
                        End If
                        체크수 = 체크수 + 1
                        If Abs(누적1 - 재고1v) > 1 Or Abs(누적2 - 재고2v) > 1 Then 체인OK = False
                    End If
                Next c
            End If
            체인총수 = 체인총수 + 1
            If 체인OK And 체크수 > 0 Then
                ws.Cells(rw, 6).Value = "PASS (" & 체크수 & "월)"
                ws.Cells(rw, 6).Font.color = RGB(0, 128, 0)
                체인OK수 = 체인OK수 + 1
            ElseIf 체크수 = 0 Then
                ws.Cells(rw, 6).Value = "데이터없음"
                ws.Cells(rw, 6).Font.color = RGB(150, 150, 150)
            Else
                ws.Cells(rw, 6).Value = "FAIL"
                ws.Cells(rw, 6).Font.color = vbRed
                ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6)).Interior.color = RGB(255, 230, 230)
            End If
            If rw Mod 2 = 0 And (체인OK Or 체크수 = 0) Then ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 6)).Interior.color = RGB(242, 242, 242)
            rw = rw + 1
Next원단23:
        Next r
    End If
    rw = rw + 1
    ws.Cells(rw, 1).Value = "요약": ws.Cells(rw, 1).Font.Bold = True
    ws.Cells(rw, 2).Value = "그룹: " & 그룹수 & "개, 재고체인: PASS " & 체인OK수 & "/" & 체인총수
    If 체인OK수 = 체인총수 And 체인총수 > 0 Then
        ws.Cells(rw, 6).Value = "ALL PASS"
        ws.Cells(rw, 6).Font.color = RGB(0, 128, 0)
    ElseIf 체인총수 = 0 Then
        ws.Cells(rw, 6).Value = "검증불가"
    Else
        ws.Cells(rw, 6).Value = "FAIL EXISTS"
        ws.Cells(rw, 6).Font.color = vbRed
    End If
    ws.Cells(rw, 6).Font.Bold = True
    ws.Columns("A").ColumnWidth = 5: ws.Columns("B").ColumnWidth = 14: ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 12: ws.Columns("E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 16
End Sub

Private Sub M23_보고서_설비별(ws As Worksheet, arr As Variant, paIdx As Object, 시작일 As Date, 종료일 As Date)
    If Not IsArray(arr) Then ws.Cells(1, 1).Value = "DB 데이터 없음": Exit Sub
    ws.Cells(1, 1).Value = "설비별 원지투입·기준소요량 상세 레포트"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Cells(2, 1).Value = "기간: " & Format(시작일, "yyyy-mm-dd") & " ~ " & Format(종료일, "yyyy-mm-dd")
    Dim 설비 As Variant: 설비 = Array("KT", "FT", "JRT", "BT", "PRE")
    Dim col원투 As Variant: col원투 = Array(28, 40, 43, 46, 49)
    Dim col기소 As Variant: col기소 = Array(31, 123, 124, 125, 126)
    Dim col주원 As Variant: col주원 = Array(34, 54, 58, 64, 68)
    Dim col월원 As Variant: col월원 = Array(37, 56, 60, 66, 70)
    Dim col주평원 As Variant: col주평원 = Array(73, 79, 85, 91, 97)
    Dim col월평원 As Variant: col월평원 = Array(76, 82, 88, 94, 100)
    Dim col주기소 As Variant: col주기소 = Array(103, 107, 111, 115, 119)
    Dim col월기소 As Variant: col월기소 = Array(104, 108, 112, 116, 120)
    Dim col주평기 As Variant: col주평기 = Array(105, 109, 113, 117, 121)
    Dim col월평기 As Variant: col월평기 = Array(106, 110, 114, 118, 122)
    Dim hdr As Variant: hdr = Array("설비", "원지투입총량(일)", "기준소요량(일)", "주간누계_원투", "주간누계_기소", "주간평균_원투", "주간평균_기소", "월누계_원투", "월누계_기소", "월평균_원투", "월평균_기소")
    Dim i As Long: For i = 0 To 10: ws.Cells(4, i + 1).Value = hdr(i): Next i
    Call M23_헤더서식(ws.Range("A4:K4"))
    Dim lastKey As String: lastKey = Format(종료일, "YYYYMMDD")
    Dim d As Date
    For d = 종료일 To 시작일 Step -1
        lastKey = Format(d, "YYYYMMDD")
        If paIdx.exists(lastKey) Then Exit For
    Next d
    Dim rw As Long: rw = 5
    Dim si As Long, r As Long, maxCol As Long
    For si = 0 To 4
        ws.Cells(rw, 1).Value = 설비(si)
        If paIdx.exists(lastKey) Then
            r = CLng(paIdx(lastKey))
            maxCol = UBound(arr, 2)
            If col원투(si) <= maxCol Then ws.Cells(rw, 2).Value = Round(M23_SDbl(arr(r, col원투(si))), 1)
            If col기소(si) <= maxCol Then ws.Cells(rw, 3).Value = Round(M23_SDbl(arr(r, col기소(si))), 1)
            If col주원(si) <= maxCol Then ws.Cells(rw, 4).Value = Round(M23_SDbl(arr(r, col주원(si))), 1)
            If col주기소(si) <= maxCol Then ws.Cells(rw, 5).Value = Round(M23_SDbl(arr(r, col주기소(si))), 1)
            If col주평원(si) <= maxCol Then ws.Cells(rw, 6).Value = Round(M23_SDbl(arr(r, col주평원(si))), 1)
            If col주평기(si) <= maxCol Then ws.Cells(rw, 7).Value = Round(M23_SDbl(arr(r, col주평기(si))), 1)
            If col월원(si) <= maxCol Then ws.Cells(rw, 8).Value = Round(M23_SDbl(arr(r, col월원(si))), 1)
            If col월기소(si) <= maxCol Then ws.Cells(rw, 9).Value = Round(M23_SDbl(arr(r, col월기소(si))), 1)
            If col월평원(si) <= maxCol Then ws.Cells(rw, 10).Value = Round(M23_SDbl(arr(r, col월평원(si))), 1)
            If col월평기(si) <= maxCol Then ws.Cells(rw, 11).Value = Round(M23_SDbl(arr(r, col월평기(si))), 1)
        End If
        For i = 2 To 11: ws.Cells(rw, i).NumberFormat = "#,##0.0": Next i
        If rw Mod 2 = 0 Then ws.Range(ws.Cells(rw, 1), ws.Cells(rw, 11)).Interior.color = RGB(242, 242, 242)
        rw = rw + 1
    Next si
    ws.Cells(rw, 1).Value = "마지막 일자(" & lastKey & ") PA 기준": ws.Cells(rw, 1).Font.Italic = True
    ws.Columns("A").ColumnWidth = 8: ws.Columns("B").ColumnWidth = 14: ws.Columns("C").ColumnWidth = 14
End Sub


