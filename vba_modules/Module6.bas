Attribute VB_Name = "Module6"
'================================================================================
' Stock 시트 현재고 표시 모듈
'
' ▶ 시트 활성화 시 자동 갱신 (ThisWorkbook Workbook_SheetActivate에서 호출)
'
' ★ H열(기말재고) 병합 구조
'   각 6행 그룹: D~H열은 상하 2블록 병합
'   · 블록1(+0~+2) 값 저장 위치: 그룹시작+0  ← Roll / 1차포장
'   · 블록2(+3~+5) 값 저장 위치: 그룹시작+3  ← KG  / 2차포장
'   ⇒ +2, +5 는 병합 안쪽 → 항상 빈값 반환 (반드시 +0, +3 으로 읽을 것)
'
' ▶ 값만 기록, 테두리만 자동 설정 (데이터 증감 시 동적 재설정, 나머지 서식 변경 없음)
'================================================================================

Option Explicit

' ── 시트명 ──────────────────────────────────────────────────────────────────
Private Const C_STOCK      As String = "Stock"
Private Const C_원단관리   As String = "원단관리"
Private Const C_부자재관리 As String = "부자재관리"
Private Const C_RawData    As String = "Raw data"
Private Const C_사양서     As String = "제품표준사양서"
Private Const C_PW         As String = "0205"

' ── 원단관리 / 부자재관리 공통 레이아웃 ─────────────────────────────────────
Private Const C_ITEM_START As Long = 4   ' 데이터 시작 행
Private Const C_ITEM_ROWS  As Long = 6   ' 그룹당 행 수
' ★ H열 병합: 값은 블록 상단에 저장됨 (+0, +3). +2/+5 는 병합 안쪽 → 빈값
Private Const C_OFF_H1     As Long = 0   ' Roll / 1차포장 기말재고 (그룹시작+0)
Private Const C_OFF_H2     As Long = 3   ' KG  / 2차포장 기말재고 (그룹시작+3)
Private Const C_COL_A      As Long = 1   ' A열 ? Co / 업체명
Private Const C_COL_B      As Long = 2   ' B열 ? 원단명 / 제품명
Private Const C_COL_H      As Long = 8   ' H열 ? 기말재고(현재고)
Private Const C_COL_I      As Long = 9   ' I열 ? 구분(입고/사용/재고)

' ── 제품표준사양서 ───────────────────────────────────────────────────────────
Private Const C_사양서_START As Long = 9
Private Const C_사양서_헤더행 As Long = 8   ' 8행→7행 순 검색
Private Const C_사양서_D     As Long = 4   ' D열 (기본)
Private Const C_사양서_E     As Long = 5
Private Const C_사양서_F     As Long = 6
Private Const C_사양서_G     As Long = 7   ' G열 ? 제품명 (기본)
Private Const C_현재고       As String = "Stock"  ' Stock 시트 ? 일일생산DB 기반 1차/2차포장 재고

'================================================================================
' ★ 외부 진입점
'================================================================================
Public Sub Stock_시트_갱신()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(C_STOCK)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Call Stock_갱신_내부(ws)
End Sub

'================================================================================
' 내부 갱신 로직
'================================================================================
Private Sub Stock_갱신_내부(ws As Worksheet)
    Dim 원래Screen As Boolean, 원래Events As Boolean
    Dim 시보호 As Boolean, wb보호 As Boolean

    원래Screen = Application.ScreenUpdating
    원래Events = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wb보호 = ThisWorkbook.ProtectStructure
    On Error Resume Next
    If wb보호 Then ThisWorkbook.Unprotect C_PW
    시보호 = ws.ProtectContents
    If 시보호 Then
        ws.Unprotect C_PW
        If ws.ProtectContents Then ws.Unprotect
    End If
    On Error GoTo 0

    On Error GoTo ErrCleanup

    ' 관리자 설정 서식 유지: 4행 이하 값만 Clear 후 갱신 (A23:D23 헤더 제외)
    On Error Resume Next
    Dim clearLast As Long
    clearLast = ws.usedRange.Row + ws.usedRange.Rows.Count - 1
    If clearLast < 4 Then clearLast = 200
    If clearLast >= 22 Then
        ws.Range(ws.Rows(4), ws.Rows(22)).ClearContents
        Stock_테두리_제거 ws, 4, 22, 6
    End If
    If clearLast >= 24 Then
        ws.Range(ws.Rows(24), ws.Rows(clearLast)).ClearContents
        Stock_테두리_제거 ws, 24, clearLast, 6
    End If
    If clearLast < 22 Then
        ws.Range(ws.Rows(4), ws.Rows(clearLast)).ClearContents
        Stock_테두리_제거 ws, 4, clearLast, 6
    End If
    On Error GoTo ErrCleanup

    ' ── 타이틀/조회시간: 값만 기록 (서식 유지) ──
    ws.Cells(1, 1).Value = "현재고 현황"
    ws.Cells(2, 1).Value = "조회 시간"
    ws.Cells(2, 2).Value = Now()

    Dim 다음행 As Long
    다음행 = 4
    다음행 = Stock_원단섹션(ws, 다음행)
    다음행 = 다음행 + 2
    If 다음행 < 22 Then 다음행 = 22   ' 완제품 22행·헤더 23행 고정
    다음행 = Stock_완제품섹션(ws, 다음행)

ErrCleanup:
    If Err.Number <> 0 Then Debug.Print "[Stock갱신] 오류: " & Err.Description
    On Error Resume Next
    If 시보호 Then
        ws.Protect Password:=C_PW, DrawingObjects:=True, _
                  Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    If wb보호 Then ThisWorkbook.Protect Password:=C_PW, Structure:=True, Windows:=False
    Application.ScreenUpdating = 원래Screen
    Application.EnableEvents = 원래Events
    On Error GoTo 0
End Sub

'================================================================================
' 테두리 제거 (firstRow~lastRow, 1~colCount열)
'================================================================================
Private Sub Stock_테두리_제거(ws As Worksheet, firstRow As Long, lastRow As Long, colCount As Long)
    If lastRow < firstRow Or colCount < 1 Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, colCount))
    On Error Resume Next
    rng.Borders.LineStyle = xlLineStyleNone
    On Error GoTo 0
End Sub

'================================================================================
' 테두리 적용 (firstRow~lastRow, 1~colCount열) ? 테두리만, 기타 서식 없음
'================================================================================
Private Sub Stock_테두리_적용(ws As Worksheet, firstRow As Long, lastRow As Long, colCount As Long)
    If lastRow < firstRow Or colCount < 1 Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, colCount))
    On Error Resume Next
    With rng.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeLeft): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlInsideHorizontal): .LineStyle = xlContinuous: .Weight = xlHairline: End With
    If colCount > 1 Then With rng.Borders(xlInsideVertical): .LineStyle = xlContinuous: .Weight = xlHairline: End With
    On Error GoTo 0
End Sub

'================================================================================
' 섹션 1 ? 원단 현재고
'   · 원단 목록 : Raw data H3:Ilast (Co 인덱스 / 원단명) ? H·I 동시 비어있으면 제외
'   · 현재고    : 원단관리 H열(기말재고) ? Roll = 그룹시작+0, KG = 그룹시작+3
'   · 반환값    : 다음 섹션 시작 행 번호
'================================================================================
Private Function Stock_원단섹션(ws As Worksheet, sRow As Long) As Long
    Const COLS As Long = 6   ' Co | 원단명 | Roll현재고 | 중량(KG) | 중량(TON) E·F병합
    Dim wsRaw  As Worksheet, ws원단 As Worksheet
    Dim lastRawRow As Long, lastWonRow As Long
    Dim r As Long, g As Long, dRow As Long, ci As Long
    Dim CoVal As String, 원단명 As String, wsCoVal As String
    Dim Roll재고 As Double, KG재고 As Double
    Dim 합계Roll As Double, 합계KG As Double
    Dim hRow As Long, totalRow As Long

    ' ── 섹션 제목 (값만, A:F 병합) ──
    ws.Cells(sRow, 1).Value = "원단 현재고"
    On Error Resume Next
    ws.Range(ws.Cells(sRow, 1), ws.Cells(sRow, COLS)).UnMerge
    ws.Range(ws.Cells(sRow, 1), ws.Cells(sRow, COLS)).Merge
    On Error GoTo 0

    ' ── 컬럼 헤더 (값만, E·F 병합) ──
    hRow = sRow + 1
    Dim hData As Variant
    hData = Array("Co", "원단명", "Roll 현재고", "중량 (KG)", "중량 (TON)")
    For ci = 0 To 3
        ws.Cells(hRow, ci + 1).Value = hData(ci)
    Next ci
    ws.Cells(hRow, 5).Value = hData(4)
    On Error Resume Next
    ws.Range(ws.Cells(hRow, 5), ws.Cells(hRow, 6)).UnMerge
    ws.Range(ws.Cells(hRow, 5), ws.Cells(hRow, 6)).Merge
    On Error GoTo 0

    dRow = sRow + 2
    합계Roll = 0
    합계KG = 0

    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets(C_RawData)
    Set ws원단 = ThisWorkbook.Sheets(C_원단관리)
    On Error GoTo 0

    If wsRaw Is Nothing Then
        ws.Cells(dRow, 1).Value = "(Raw data 시트 없음)"
        Stock_원단섹션 = dRow + 2
        Exit Function
    End If
    If ws원단 Is Nothing Then
        ws.Cells(dRow, 1).Value = "(원단관리 시트 없음)"
        Stock_원단섹션 = dRow + 2
        Exit Function
    End If

    lastRawRow = wsRaw.Cells(wsRaw.Rows.Count, 8).End(xlUp).Row   ' H열 마지막
    If lastRawRow < 3 Then
        ws.Cells(dRow, 1).Value = "(Raw data 데이터 없음)"
        Stock_원단섹션 = dRow + 2
        Exit Function
    End If
    lastWonRow = ws원단.Cells(ws원단.Rows.Count, C_COL_I).End(xlUp).Row

    ' ── Raw data H:I 순회 → 데이터 행 출력 ──
    For r = 3 To lastRawRow
        CoVal = Trim(CStr(wsRaw.Cells(r, 8).Value))
        원단명 = Trim(CStr(wsRaw.Cells(r, 9).Value))
        If CoVal = "" Or 원단명 = "" Then GoTo NextRawRow

        If IsNumeric(CoVal) Then
            If CLng(val(CoVal)) > 0 Then CoVal = Format(CLng(val(CoVal)), "00")
        End If

        Roll재고 = 0: KG재고 = 0
        If lastWonRow >= C_ITEM_START Then
            For g = C_ITEM_START To lastWonRow - (C_ITEM_ROWS - 1) Step C_ITEM_ROWS
                wsCoVal = Trim(CStr(ws원단.Cells(g, C_COL_A).Value))
                If IsNumeric(wsCoVal) Then
                    If CLng(val(wsCoVal)) > 0 Then wsCoVal = Format(CLng(val(wsCoVal)), "00")
                End If
                If wsCoVal = CoVal Then
                    On Error Resume Next
                    If IsNumeric(ws원단.Cells(g + C_OFF_H1, C_COL_H).Value) Then
                        Roll재고 = CDbl(ws원단.Cells(g + C_OFF_H1, C_COL_H).Value)
                    End If
                    If IsNumeric(ws원단.Cells(g + C_OFF_H2, C_COL_H).Value) Then
                        KG재고 = CDbl(ws원단.Cells(g + C_OFF_H2, C_COL_H).Value)
                    End If
                    On Error GoTo 0
                    Exit For
                End If
            Next g
        End If

        ws.Cells(dRow, 1).Value = CoVal
        ws.Cells(dRow, 2).Value = 원단명
        ws.Cells(dRow, 3).Value = Roll재고
        ws.Cells(dRow, 4).Value = KG재고
        ws.Cells(dRow, 5).Value = KG재고 / 1000
        On Error Resume Next
        ws.Range(ws.Cells(dRow, 5), ws.Cells(dRow, 6)).UnMerge
        ws.Range(ws.Cells(dRow, 5), ws.Cells(dRow, 6)).Merge
        On Error GoTo 0

        합계Roll = 합계Roll + Roll재고
        합계KG = 합계KG + KG재고
        dRow = dRow + 1

NextRawRow:
    Next r

    ' ── 합계 행 (값만, E·F 병합) ──
    totalRow = dRow
    ws.Cells(totalRow, 1).Value = "합  계"
    ws.Cells(totalRow, 3).Value = 합계Roll
    ws.Cells(totalRow, 4).Value = 합계KG
    ws.Cells(totalRow, 5).Value = 합계KG / 1000
    On Error Resume Next
    ws.Range(ws.Cells(totalRow, 5), ws.Cells(totalRow, 6)).UnMerge
    ws.Range(ws.Cells(totalRow, 5), ws.Cells(totalRow, 6)).Merge
    On Error GoTo 0

    Call Stock_테두리_적용(ws, hRow, totalRow, COLS)
    Stock_원단섹션 = totalRow + 1
End Function

'================================================================================
' 제품표준사양서 헤더에서 열 번호 찾기 (8행→7행 순)
'================================================================================
Private Function Stock_사양서_열찾기(ws As Worksheet, 헤더명 As String) As Long
    Dim hr As Long, c As Long, lastCol As Long
    Dim t As String, findVal As String
    If ws Is Nothing Then Stock_사양서_열찾기 = 0: Exit Function
    findVal = LCase(Trim(Replace(Replace(헤더명, " ", ""), chR(10), "")))
    For hr = C_사양서_헤더행 To 1 Step -1
        lastCol = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 20
        For c = 1 To lastCol
            t = Trim(Replace(Replace(CStr(ws.Cells(hr, c).Value), " ", ""), chR(10), ""))
            If LCase(t) = findVal Then Stock_사양서_열찾기 = c: Exit Function
        Next c
    Next hr
    Stock_사양서_열찾기 = 0
End Function

'================================================================================
' 섹션 2 ? 완제품 현재고
'   · 품목      : 제품표준사양서 D9:G열 ? D,E,F,G 그대로 복사 (공백 유지)
'   · 현재고    : 부자재관리 H열 ? 1차포장(그룹+0), 2차포장(그룹+3)
'   · 매칭      : G열(제품명) ↔ 부자재관리 B열(COL_제품) ? 대소문자 무시
'================================================================================
Private Function Stock_완제품섹션(ws As Worksheet, sRow As Long) As Long
    Const COLS As Long = 6   ' D|E|F|G|1차포장|2차포장
    Dim ws사양서 As Worksheet, ws부자재 As Worksheet
    Dim lastRow사 As Long, last부자재Row As Long
    Dim lastD As Long, lastE As Long, lastF As Long, lastG As Long
    Dim r As Long, g As Long, dRow As Long
    Dim vd As Variant, vE As Variant, vF As Variant, vG As Variant
    Dim ws제품명 As String
    Dim 포장1재고 As Double, 포장2재고 As Double
    Dim vGStr As String
    Dim hRow As Long

    ' ── 섹션 제목 (값만) 22행 고정 ──
    ws.Cells(sRow, 1).Value = "완제품 현재고"

    hRow = sRow + 1   ' 23행 = 헤더 (A23:D23 관리자 설정, 미기록)
    dRow = sRow + 2   ' 24행부터 데이터

    On Error Resume Next
    Set ws사양서 = ThisWorkbook.Sheets(C_사양서)
    Set ws부자재 = ThisWorkbook.Sheets(C_부자재관리)
    On Error GoTo 0

    If ws사양서 Is Nothing Then
        ws.Cells(dRow, 1).Value = "(제품표준사양서 시트 없음)"
        Stock_완제품섹션 = dRow + 2
        Exit Function
    End If

    ' ── 23행 헤더: A23:D23 관리자 설정 유지, 기록하지 않음 ──

    ' 마지막 행: D·E·F·G 중 가장 아래 ──
    lastD = ws사양서.Cells(ws사양서.Rows.Count, C_사양서_D).End(xlUp).Row
    lastE = ws사양서.Cells(ws사양서.Rows.Count, C_사양서_E).End(xlUp).Row
    lastF = ws사양서.Cells(ws사양서.Rows.Count, C_사양서_F).End(xlUp).Row
    lastG = ws사양서.Cells(ws사양서.Rows.Count, C_사양서_G).End(xlUp).Row
    lastRow사 = lastD
    If lastE > lastRow사 Then lastRow사 = lastE
    If lastF > lastRow사 Then lastRow사 = lastF
    If lastG > lastRow사 Then lastRow사 = lastG

    If lastRow사 < C_사양서_START Then
        ws.Cells(dRow, 1).Value = "(제품표준사양서 D9:G 데이터 없음)"
        Stock_완제품섹션 = dRow + 2
        Exit Function
    End If

    last부자재Row = 0
    If Not ws부자재 Is Nothing Then
        last부자재Row = ws부자재.Cells(ws부자재.Rows.Count, C_COL_I).End(xlUp).Row
    End If

    ' ── 제품표준사양서 D9:Glast 순회 ? D~G 그대로 복사(공백 유지), 재고는 부자재관리에서 조회 ──
    For r = C_사양서_START To lastRow사
        vd = ws사양서.Cells(r, C_사양서_D).Value
        vE = ws사양서.Cells(r, C_사양서_E).Value
        vF = ws사양서.Cells(r, C_사양서_F).Value
        vG = ws사양서.Cells(r, C_사양서_G).Value

        ' 품목: D~G 그대로 (공백 유지)
        ws.Cells(dRow, 1).Value = vd
        ws.Cells(dRow, 2).Value = vE
        ws.Cells(dRow, 3).Value = vF
        ws.Cells(dRow, 4).Value = vG

        ' 1차/2차포장 현재고: 부자재관리 H열
        ' 매칭: G열(제품명) ↔ 부자재관리 B열(COL_제품=2) ? 대소문자 무시, 공백 제거
        포장1재고 = 0: 포장2재고 = 0
        vGStr = LCase(Trim(Replace(Replace(CStr(vG), " ", ""), chR(10), "")))
        If vGStr <> "" And last부자재Row >= C_ITEM_START Then
            For g = C_ITEM_START To last부자재Row - (C_ITEM_ROWS - 1) Step C_ITEM_ROWS
                ws제품명 = LCase(Trim(Replace(Replace(CStr(ws부자재.Cells(g, C_COL_B).Value), " ", ""), chR(10), "")))
                If ws제품명 = vGStr Then
                    On Error Resume Next
                    If IsNumeric(ws부자재.Cells(g + C_OFF_H1, C_COL_H).Value) Then
                        포장1재고 = CDbl(ws부자재.Cells(g + C_OFF_H1, C_COL_H).Value)
                    End If
                    If IsNumeric(ws부자재.Cells(g + C_OFF_H2, C_COL_H).Value) Then
                        포장2재고 = CDbl(ws부자재.Cells(g + C_OFF_H2, C_COL_H).Value)
                    End If
                    On Error GoTo 0
                    Exit For
                End If
            Next g
        End If

        ws.Cells(dRow, 5).Value = 포장1재고
        ws.Cells(dRow, 6).Value = 포장2재고
        dRow = dRow + 1
    Next r

    If dRow > hRow + 1 Then Call Stock_테두리_적용(ws, hRow, dRow - 1, COLS)
    Stock_완제품섹션 = dRow + 1
End Function




