Attribute VB_Name = "Module7"
'================================================================================
' 시트 이동/닫기(완전 감추기) 버튼 생성
' - 모든 시트에 대해 "시트명 열기", "시트명 닫기" 버튼 쌍 생성
' - 버튼은 시트 "이름"으로 열기/닫기 (시트 순서가 바뀌어도 Index key→Index key, 부자재관리→부자재관리 정확히 동작)
' - 열기: 시트 표시 후 Home menu 유지. 닫기: 시트 완전 감추기 (xlSheetVeryHidden)
'
' 사용법: 표준 모듈에 붙여넣은 뒤 "시트이동_버튼_생성" 매크로 실행 (기존 버튼 있으면 제거 후 재생성)
' 통합문서 보호(비밀번호 0205) 상태에서도 매크로 실행 가능하도록, 시트/구조 변경 시에만 잠시 해제 후 재설정
'================================================================================

Option Explicit

Private Const WS_HOME As String = "Home menu"
Private Const BTN_R As Long = 233
Private Const BTN_G As Long = 246
Private Const BTN_B As Long = 168
Private Const WB_PASSWORD As String = "0205"

' 통합문서 보호 해제 (매크로에서 시트 추가/삭제/표시 변경 시 호출). 반환: 해제했으면 True
Private Function 워크북_보호해제() As Boolean
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect WB_PASSWORD
        워크북_보호해제 = True
    End If
    On Error GoTo 0
End Function

' 통합문서 보호 설정 (비밀번호 0205, 구조 보호)
Private Sub 워크북_보호설정()
    On Error Resume Next
    ThisWorkbook.Protect Password:=WB_PASSWORD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

' 시트 이름으로 열기 (버튼 클릭 시 Application.Caller에서 시트명 추출 → 시트 순서와 무관하게 정확히 열림)
Public Sub 시트_열기_버튼()
    Dim callerName As String, 시트명 As String, sep As String
    Dim ws As Worksheet, wsHome As Worksheet
    Dim wasProtected As Boolean
    On Error GoTo ErrOut
    callerName = CStr(Application.Caller)
    sep = "$$"
    If InStr(1, callerName, sep) = 0 Then Exit Sub
    시트명 = Trim(Split(callerName, sep)(1))
    If 시트명 = "" Then Exit Sub
    Set ws = ThisWorkbook.Sheets(시트명)
    Set wsHome = ThisWorkbook.Sheets(WS_HOME)
    wasProtected = 워크북_보호해제
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Visible = xlSheetVisible
    If Not wsHome Is Nothing Then wsHome.Activate
ErrOut:
    If wasProtected Then 워크북_보호설정
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' 시트 이름으로 닫기 (감춘 뒤 화면은 항상 Home menu 유지)
Public Sub 시트_닫기_버튼()
    Dim callerName As String, 시트명 As String, sep As String
    Dim ws As Worksheet, wsHome As Worksheet, visCount As Long, w As Object
    Dim wasProtected As Boolean
    On Error GoTo ErrOut
    callerName = CStr(Application.Caller)
    sep = "$$"
    If InStr(1, callerName, sep) = 0 Then Exit Sub
    시트명 = Trim(Split(callerName, sep)(1))
    If 시트명 = "" Then Exit Sub
    Set ws = ThisWorkbook.Sheets(시트명)
    Set wsHome = ThisWorkbook.Sheets(WS_HOME)
    If ws Is Nothing Then GoTo ErrOut
    visCount = 0
    For Each w In ThisWorkbook.Sheets
        If w.Visible = xlSheetVisible Then visCount = visCount + 1
    Next w
    If visCount <= 1 Then
        MsgBox "표시 중인 시트가 1개뿐이면 감출 수 없습니다.", vbExclamation
        GoTo ErrOut
    End If
    wasProtected = 워크북_보호해제
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Visible = xlSheetVeryHidden
    If Not wsHome Is Nothing Then wsHome.Activate
ErrOut:
    If wasProtected Then 워크북_보호설정
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Stock 시트 열기 + 갱신 (현재고_시트_갱신 실행 후 Stock 시트로 이동)
Public Sub Stock_열기_버튼()
    Dim wsCur As Worksheet, wsHome As Worksheet
    Dim wasProtected As Boolean
    On Error Resume Next
    Set wsCur = ThisWorkbook.Sheets("Stock")
    On Error GoTo 0
    
    wasProtected = 워크북_보호해제
    ' Stock 시트가 없으면 갱신 매크로 실행 (호출 전용 매크로 사용)
    On Error Resume Next
    Application.Run "Run_현재고_시트_갱신"
    On Error GoTo 0
    
    On Error Resume Next
    Set wsCur = ThisWorkbook.Sheets("Stock")
    On Error GoTo 0
    If wsCur Is Nothing Then
        If wasProtected Then 워크북_보호설정
        MsgBox "Stock 시트를 생성할 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsCur.Visible = xlSheetVisible
    Set wsHome = Nothing
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets("Home menu")
    On Error GoTo 0
    If Not wsHome Is Nothing Then wsHome.Activate
    If wasProtected Then 워크북_보호설정
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Stock 시트 닫기 (감추기)
Public Sub Stock_닫기_버튼()
    Dim wsStock As Worksheet, wsHome As Worksheet
    Dim visCount As Long, w As Object
    Dim wasProtected As Boolean
    On Error Resume Next
    Set wsStock = ThisWorkbook.Sheets("Stock")
    On Error GoTo 0
    If wsStock Is Nothing Then Exit Sub
    visCount = 0
    For Each w In ThisWorkbook.Sheets
        If w.Visible = xlSheetVisible Then visCount = visCount + 1
    Next w
    If visCount <= 1 Then
        MsgBox "표시 중인 시트가 1개뿐이면 감출 수 없습니다.", vbExclamation
        Exit Sub
    End If
    wasProtected = 워크북_보호해제
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsStock.Visible = xlSheetVeryHidden
    Set wsHome = Nothing
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets("Home menu")
    On Error GoTo 0
    If Not wsHome Is Nothing Then wsHome.Activate
    If wasProtected Then 워크북_보호설정
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 매뉴얼 시트 열기
Public Sub 매뉴얼_열기_버튼()
    Dim wsManual As Worksheet, wsHome As Worksheet
    Dim wasProtected As Boolean
    On Error Resume Next
    Set wsManual = ThisWorkbook.Sheets("매뉴얼")
    On Error GoTo 0
    If wsManual Is Nothing Then
        MsgBox "매뉴얼 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    wasProtected = 워크북_보호해제
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsManual.Visible = xlSheetVisible
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(WS_HOME)
    If Not wsHome Is Nothing Then wsHome.Activate
    On Error GoTo 0
    If wasProtected Then 워크북_보호설정
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 매뉴얼 시트 닫기 (감추기)
Public Sub 매뉴얼_닫기_버튼()
    Dim wsManual As Worksheet, wsHome As Worksheet, visCount As Long, w As Object
    Dim wasProtected As Boolean
    On Error Resume Next
    Set wsManual = ThisWorkbook.Sheets("매뉴얼")
    On Error GoTo 0
    If wsManual Is Nothing Then Exit Sub
    visCount = 0
    For Each w In ThisWorkbook.Sheets
        If w.Visible = xlSheetVisible Then visCount = visCount + 1
    Next w
    If visCount <= 1 Then
        MsgBox "표시 중인 시트가 1개뿐이면 감출 수 없습니다.", vbExclamation
        Exit Sub
    End If
    wasProtected = 워크북_보호해제
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsManual.Visible = xlSheetVeryHidden
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(WS_HOME)
    If Not wsHome Is Nothing Then wsHome.Activate
    On Error GoTo 0
    If wasProtected Then 워크북_보호설정
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Home menu 시트에 시트별 "열기/닫기" 버튼 생성. 버튼 이름에 시트명 포함 → 시트 순서가 바뀌어도 올바른 시트 열기/닫기
Public Sub 시트이동_버튼_생성()
    Const btnW As Double = 73.71
    Const btnH As Double = 17.01
    Const colGap As Double = 8
    Const rowGap As Double = 4

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim btn As Shape
    Dim rngArea As Range
    Dim blockWidth As Double, blockLeft As Double, startTop As Double
    Dim c1Left As Double, c2Left As Double
    Dim x As Double, y As Double
    Dim i As Long, rowIdx As Long, colIdx As Long
    Dim grpCount As Long
    Dim shtName As String

    Set wb = ThisWorkbook
    If wb.ReadOnly Then
        MsgBox "통합문서가 읽기 전용이어서 버튼을 생성할 수 없습니다." & vbCrLf & _
               "다른 이름으로 저장하거나 읽기 전용을 해제한 뒤 다시 시도하세요.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    If wb.MultiUserEditing Then
        MsgBox "공유 통합문서에서는 시트이동 버튼을 생성할 수 없습니다." & vbCrLf & _
               "공유를 해제한 뒤 다시 시도하세요.", vbExclamation
        If wasProtected Then 워크북_보호설정
        Exit Sub
    End If
    On Error GoTo 0
    wasProtected = 워크북_보호해제
    On Error Resume Next
    Set ws = wb.Sheets(WS_HOME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = WS_HOME
    End If

    ' 시트 보호 해제 (Error 70 방지: 양식버튼 추가/삭제 시 권한 필요)
    Dim sheetWasProtected As Boolean
    sheetWasProtected = ws.ProtectContents
    If sheetWasProtected Then
        On Error Resume Next
        ws.Unprotect WB_PASSWORD
        If ws.ProtectContents Then ws.Unprotect ""
        On Error GoTo 0
    End If
    ws.Activate  ' 작업 시트 활성화

    ' 기존 시트이동 버튼 제거 (Open$$, Close$$ 이름)
    Dim ug As Long, nm As String
    For ug = 1 To 10
        For i = ws.Shapes.Count To 1 Step -1
            On Error Resume Next
            If ws.Shapes(i).Type = msoGroup And left(ws.Shapes(i).Name, 11) = "SheetBtnGroup" Then ws.Shapes(i).Ungroup
            On Error GoTo 0
        Next i
    Next ug
    For i = ws.Shapes.Count To 1 Step -1
        On Error Resume Next
        nm = ws.Shapes(i).Name
        If Err.Number = 0 Then
            If left(nm, 6) = "Open$$" Or left(nm, 7) = "Close$$" Or left(nm, 8) = "SheetBtn_" Or left(nm, 11) = "SheetBtnGroup" Then ws.Shapes(i).Delete
        End If
        On Error GoTo 0
    Next i

    ' 배치 영역: 2열 중앙 (1열=열기, 2열=닫기 → 시트별 1행이 1세트)
    Set rngArea = ws.Range("B5:N35")
    blockWidth = 2 * btnW + colGap
    blockLeft = rngArea.left + (rngArea.width - blockWidth) / 2
    startTop = rngArea.top
    c1Left = blockLeft
    c2Left = blockLeft + btnW + colGap

    Dim ns As Long
    ns = wb.Sheets.Count
    If ns > 60 Then
        MsgBox "시트가 60개를 초과하면 버튼을 60개까지만 만듭니다.", vbExclamation
        ns = 60
    End If
    grpCount = 0

    For i = 1 To ns
        shtName = wb.Sheets(i).Name
        rowIdx = i - 1
        y = startTop + rowIdx * (btnH + rowGap)

        ' 열기 버튼 (1열) - 도형버튼 (AddShape, Error 70 회피)
        x = c1Left
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, btnW, btnH)
        btn.Name = "Open$$" & shtName
        btn.OnAction = "시트_열기_버튼"
        On Error Resume Next: btn.TextFrame2.TextRange.Text = shtName: On Error GoTo 0
        grpCount = grpCount + 1

        ' 닫기 버튼 (2열) - 도형버튼
        x = c2Left
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, btnW, btnH)
        btn.Name = "Close$$" & shtName
        btn.OnAction = "시트_닫기_버튼"
        On Error Resume Next: btn.TextFrame2.TextRange.Text = shtName: On Error GoTo 0
        grpCount = grpCount + 1
    Next i

    ' ── Stock 전용 버튼 (열기=갱신+이동, 닫기=감추기) ──
    y = startTop + ns * (btnH + rowGap) + rowGap * 2  ' 시트 버튼들 아래에 간격 두고 배치

    ' Stock 열기 버튼 (1열)
    x = c1Left
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, btnW, btnH)
    btn.Name = "Open$$Stock"
    btn.OnAction = "Stock_열기_버튼"
    On Error Resume Next: btn.TextFrame2.TextRange.Text = "Stock": On Error GoTo 0
    grpCount = grpCount + 1

    ' Stock 닫기 버튼 (2열)
    x = c2Left
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, btnW, btnH)
    btn.Name = "Close$$Stock"
    btn.OnAction = "Stock_닫기_버튼"
    On Error Resume Next: btn.TextFrame2.TextRange.Text = "Stock": On Error GoTo 0
    grpCount = grpCount + 1

    ' 시트 보호 재설정 (UserInterfaceOnly:=True → VBA는 수정 가능)
    On Error Resume Next
    If sheetWasProtected Then
        ws.Protect Password:=WB_PASSWORD, UserInterfaceOnly:=True
    End If
    On Error GoTo 0

    ws.Activate
    Application.GoTo ws.Range("A1"), True
    If wasProtected Then 워크북_보호설정
    MsgBox "시트 이동/닫기 버튼이 " & grpCount & "개 생성되었습니다." & vbCrLf & _
           "(시트당 열기·닫기 1세트 + Stock 갱신/열기 버튼)", vbInformation, "시트이동 버튼"
End Sub




