Attribute VB_Name = "Module5"
'================================================================================
' 매뉴얼 시트 - 일괄 양식 버튼 생성 + 클릭 시 해당 설명으로 자동 스크롤
' - 유저매뉴얼.txt 내용을 [매뉴얼] 시트 B3부터 붙여 넣은 뒤 사용하세요.
' - 버튼: 폰트 9, Microsoft GothicNeo, 높이 0.8cm, 너비 8.5cm
' - 버튼 클릭 시 매뉴얼 시트에서 해당 【】 제목이 있는 행으로 스크롤 이동
' - 매뉴얼 내용 입력 후 매뉴얼_시트보호 실행 → 사용자 편집 불가 (선택·스크롤·버튼 클릭은 가능)
'================================================================================

Option Explicit

Private Const WS_매뉴얼 As String = "매뉴얼"
' 1cm ? 28.35 포인트
Private Const PT_PER_CM As Double = 28.35
Private Const BTN_W_CM As Double = 8.5
Private Const BTN_H_CM As Double = 0.8
Private Const BTN_GAP_CM As Double = 0.15

' 버튼 캡션 목록 (유저매뉴얼.txt의 【제목】 순서와 동일하게)
Private Function 매뉴얼_섹션목록() As Variant
    매뉴얼_섹션목록 = Array( _
        "Home menu", _
        "생산일보", _
        "부자재관리", _
        "원단관리", _
        "Stock", _
        "매뉴얼", _
        "제품표준사양서" _
    )
End Function

' 매뉴얼 시트에서 B열에 【검색문자】 제목이 있는 행 번호만 반환 (3행~). 없으면 0
' ※ "부자재관리"가 "일일생산DB"나 본문 문장에 겹치지 않도록 【제목】 형식으로만 검색
Private Function 매뉴얼_섹션행찾기(ws As Worksheet, 검색문자 As String) As Long
    Dim r As Long, lastRow As Long, v As String, 제목패턴 As String
    매뉴얼_섹션행찾기 = 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    ' 【제목】 형태로 시작하는 행만 매칭 (부자재관리 vs 일일생산DB 등 키워드 겹침 방지)
    제목패턴 = "【" & 검색문자 & "】"
    For r = 3 To lastRow
        v = Trim(CStr(ws.Cells(r, 2).Value))
        If Len(v) >= Len(제목패턴) And StrComp(left(v, Len(제목패턴)), 제목패턴, vbTextCompare) = 0 Then
            매뉴얼_섹션행찾기 = r
            Exit Function
        End If
    Next r
End Function

' 버튼 클릭 시 호출: Caller(버튼 이름 Btn_M1~Btn_M7)로 해당 섹션 행으로 스크롤
Public Sub 매뉴얼_스크롤_클릭()
    Dim callerName As String, idx As Long, r As Long
    Dim ws As Worksheet
    Dim arr As Variant, i As Long, 검색문자 As String

    On Error Resume Next
    callerName = CStr(Application.Caller)
    On Error GoTo 0
    If Len(callerName) = 0 Then Exit Sub

    ' Btn_M1 ~ Btn_M7 에서 숫자 추출
    If left(callerName, 5) <> "Btn_M" Then Exit Sub
    idx = CLng(val(Mid(callerName, 6)))
    arr = 매뉴얼_섹션목록()
    If idx < 1 Or idx > UBound(arr) + 1 Then Exit Sub
    검색문자 = arr(idx - 1)

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_매뉴얼)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "매뉴얼 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    r = 매뉴얼_섹션행찾기(ws, 검색문자)
    If r = 0 Then
        MsgBox "해당 항목을 매뉴얼 시트 B열에서 찾을 수 없습니다." & vbCrLf & "검색: " & 검색문자, vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error Resume Next
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ThisWorkbook.Sheets(WS_매뉴얼).Select
    ws.Range("B" & r).Select
    Application.GoTo Reference:=ws.Range("B" & r), Scroll:=True
    ActiveWindow.ScrollRow = Application.Max(1, r - 2)
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 매뉴얼 시트에 양식 버튼 일괄 생성 (폰트 9, Microsoft GothicNeo, 높이 0.8cm, 너비 8.5cm)
Public Sub 매뉴얼_일괄버튼_생성()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim arr As Variant, i As Long, n As Long
    Dim leftPt As Double, topPt As Double
    Dim btnW As Double, btnH As Double, gapPt As Double
    Dim s As Shape, 기존버튼있음 As Boolean

    btnW = BTN_W_CM * PT_PER_CM
    btnH = BTN_H_CM * PT_PER_CM
    gapPt = BTN_GAP_CM * PT_PER_CM

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_매뉴얼)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = WS_매뉴얼
    End If

    ' 기존 매뉴얼 버튼 있으면 제거 후 재생성 (선택: 유지하려면 아래 블록 삭제)
    기존버튼있음 = False
    For Each s In ws.Shapes
        If left(s.Name, 5) = "Btn_M" Then
            기존버튼있음 = True
            s.Delete
        End If
    Next s
    If 기존버튼있음 Then
        ' 그룹 해제 후 삭제했을 수 있음. 이미 삭제 처리됨.
    End If

    arr = 매뉴얼_섹션목록()
    n = UBound(arr) + 1
    leftPt = 10
    topPt = 10

    Application.ScreenUpdating = False
    For i = 0 To n - 1
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, btnW, btnH)
        shp.Name = "Btn_M" & (i + 1)
        shp.OnAction = "매뉴얼_스크롤_클릭"
        shp.TextFrame2.TextRange.Text = arr(i)
        With shp.TextFrame2.TextRange.Font
            .Size = 9
            .Name = "Microsoft GothicNeo"
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        shp.Fill.ForeColor.RGB = RGB(129, 135, 253)
        shp.Line.Visible = msoFalse
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        topPt = topPt + btnH + gapPt
    Next i
    Application.ScreenUpdating = True

    MsgBox "매뉴얼 버튼 " & n & "개를 생성했습니다." & vbCrLf & _
           "B3부터 유저매뉴얼.txt 내용을 붙여 넣으면 버튼 클릭 시 해당 위치로 스크롤됩니다." & vbCrLf & _
           "편집 잠금이 필요하면 매뉴얼_시트보호를 실행하세요.", vbInformation, "매뉴얼 버튼 생성"
End Sub

'================================================================================
' 매뉴얼 시트 보호: 사용자 편집 불가 (셀 선택·스크롤·버튼 클릭은 가능)
' - 비밀번호 없음. 해제는 시트 우클릭 → 시트 보호 해제.
' - Workbook_Open에서 이 서브를 호출하면 매번 열 때마다 자동 보호됩니다.
'================================================================================
Public Sub 매뉴얼_시트보호()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_매뉴얼)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "매뉴얼 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    If ws.ProtectContents Then
        MsgBox "매뉴얼 시트가 이미 보호되어 있습니다.", vbInformation
        Exit Sub
    End If
    ' 암호 0205 (통합문서 보호와 동일), 셀 선택·스크롤 가능, 도형(버튼) 클릭 가능
    ws.Protect Password:="0205", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False
    'MsgBox "매뉴얼 시트를 보호했습니다." & vbCrLf & _
           "사용자는 셀을 편집·삽입·삭제할 수 없으며, 선택·스크롤·버튼 클릭은 가능합니다.", vbInformation, "매뉴얼 시트 보호"
End Sub




