Attribute VB_Name = "Module12"
Public Sub 초기화()
    Dim ws부자재관리 As Worksheet
    Dim ws원단관리 As Worksheet
    Dim ws생산일보 As Worksheet
    Dim ws일일생산DB As Worksheet
    Dim lastRow As Long
    Dim wb As Workbook
    Dim 원래Screen As Boolean, 원래Events As Boolean, 원래Alerts As Boolean

    Set wb = ThisWorkbook
    Set ws부자재관리 = wb.Sheets("부자재관리")
    On Error Resume Next
    Set ws원단관리 = wb.Sheets("원단관리")
    On Error GoTo 0
    Set ws생산일보 = wb.Sheets("생산일보")
    Set ws일일생산DB = wb.Sheets("일일생산DB")

    Dim pw As String: pw = "0205"
    Dim 원래시트 As Object, 원래셀 As Range
    Set 원래시트 = ActiveSheet
    Set 원래셀 = ActiveCell

    원래Screen = Application.ScreenUpdating
    원래Events = Application.EnableEvents
    원래Alerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo 초기화끝

    ' ── 통합문서 보호 해제 ──
    wb.Unprotect Password:=pw

    ' ── 각 시트 보호 해제 ──
    ws부자재관리.Unprotect Password:=pw
    ws생산일보.Unprotect Password:=pw
    ws일일생산DB.Unprotect Password:=pw
    If Not ws원단관리 Is Nothing Then ws원단관리.Unprotect Password:=pw

    ' ── 1. 부자재관리 시트: 4행~마지막행 서식·수식·값 모두 삭제 ──
    lastRow = ws부자재관리.Cells(ws부자재관리.Rows.Count, 1).End(xlUp).Row
    If ws부자재관리.Cells(ws부자재관리.Rows.Count, 2).End(xlUp).Row > lastRow Then lastRow = ws부자재관리.Cells(ws부자재관리.Rows.Count, 2).End(xlUp).Row
    If ws부자재관리.Cells(ws부자재관리.Rows.Count, 10).End(xlUp).Row > lastRow Then lastRow = ws부자재관리.Cells(ws부자재관리.Rows.Count, 10).End(xlUp).Row
    On Error Resume Next
    If ws부자재관리.usedRange.Rows.Count > 0 Then
        If ws부자재관리.usedRange.Row + ws부자재관리.usedRange.Rows.Count - 1 > lastRow Then lastRow = ws부자재관리.usedRange.Row + ws부자재관리.usedRange.Rows.Count - 1
    End If
    On Error GoTo 초기화끝
    If lastRow < 6 Then lastRow = 6
    ws부자재관리.Range(ws부자재관리.Rows(4), ws부자재관리.Rows(lastRow)).Clear
    ws부자재관리.Rows("4:" & lastRow).Delete Shift:=xlUp

    ' ── 1-2. 원단관리 시트: 4행~마지막행 서식·수식·값 모두 삭제 (있을 경우) ──
    If Not ws원단관리 Is Nothing Then
        lastRow = ws원단관리.Cells(ws원단관리.Rows.Count, 1).End(xlUp).Row
        If ws원단관리.Cells(ws원단관리.Rows.Count, 2).End(xlUp).Row > lastRow Then lastRow = ws원단관리.Cells(ws원단관리.Rows.Count, 2).End(xlUp).Row
        If ws원단관리.Cells(ws원단관리.Rows.Count, 10).End(xlUp).Row > lastRow Then lastRow = ws원단관리.Cells(ws원단관리.Rows.Count, 10).End(xlUp).Row
        On Error Resume Next
        If ws원단관리.usedRange.Rows.Count > 0 Then
            If ws원단관리.usedRange.Row + ws원단관리.usedRange.Rows.Count - 1 > lastRow Then lastRow = ws원단관리.usedRange.Row + ws원단관리.usedRange.Rows.Count - 1
        End If
        On Error GoTo 초기화끝
        If lastRow < 6 Then lastRow = 6
        ws원단관리.Range(ws원단관리.Rows(4), ws원단관리.Rows(lastRow)).Clear
        ws원단관리.Rows("4:" & lastRow).Delete Shift:=xlUp
    End If

    ' ── 2. 생산일보 시트: 메뉴초기화 매크로 실행 ──
    Application.Run "메뉴초기화"

    ' ── 3. 일일생산DB 시트: 2행~마지막행 전체 삭제 후 70열 헤더 생성 ──
    lastRow = ws일일생산DB.Cells(ws일일생산DB.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws일일생산DB.Rows("2:" & lastRow).Delete Shift:=xlUp
    End If
    Call 헤더_생성(ws일일생산DB)

    ' ── 4. 원단관리_동기화상태 시트 초기화 ──
    On Error Resume Next
    Dim ws동기화 As Worksheet
    Set ws동기화 = wb.Sheets("원단관리_동기화상태")
    If Not ws동기화 Is Nothing Then ws동기화.Range("A1").Value = 0
    On Error GoTo 0

    ' ── 5. 부자재DB 시트 삭제됨: 부자재 데이터는 일일생산DB 71~81열에 저장 (삭제 생략) ──

    ' ── 각 시트 보호 복원 ──
    ws부자재관리.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    If Not ws원단관리 Is Nothing Then ws원단관리.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws생산일보.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ws일일생산DB.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' ── 통합문서 보호 복원 ──
    wb.Protect Password:=pw, Structure:=True, Windows:=False

초기화끝:
    Application.ScreenUpdating = 원래Screen
    Application.EnableEvents = 원래Events
    Application.DisplayAlerts = 원래Alerts
    Application.StatusBar = False
    On Error Resume Next
    If Not 원래시트 Is Nothing Then 원래시트.Activate
    If Not 원래셀 Is Nothing Then 원래셀.Select
    On Error GoTo 0
End Sub

