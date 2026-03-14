Attribute VB_Name = "Module26"
' ============================================================
' 초기화_원지유지 - 생산일보 시트 지정 영역 제거 (주/야간 원지사용 내역만 유지)
' ============================================================
' 주/야간 원지사용 내역(C7:BV16, C20:BV29, C35:BV44, C48:BV57)은 유지하고,
' 나머지 기존 영역(완제품 가공내역, 작업시간, T10 원자재수율, 비가동현황 등)만 삭제.
' Module1 메뉴초기화 구조를 기반으로 원지사용내역 영역을 제외한 삭제 로직.
' ============================================================

Option Explicit

Private Sub 영역지우기_병합허용(ws As Worksheet, 영역주소 As String, Optional 서식메모포함 As Boolean = False)
    Dim rng As Range, c As Range, ma As Range
    Dim clearedKeys As Object
    Dim key As String

    Set rng = ws.Range(영역주소)
    Set clearedKeys = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each c In rng.Cells
        If c.MergeCells Then
            Set ma = c.MergeArea
            key = ma.Address(False, False)
            If Not clearedKeys.exists(key) Then
                clearedKeys.Add key, True
                ma.ClearContents
                If 서식메모포함 Then
                    ma.Interior.Pattern = xlPatternNone
                    ma.ClearComments
                End If
            End If
        Else
            c.ClearContents
            If 서식메모포함 Then
                c.Interior.Pattern = xlPatternNone
                c.ClearComments
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

Public Sub 초기화_원지유지()
    Dim ws As Worksheet
    Dim 시트명 As String
    Dim 원래ScreenUpdating As Boolean
    Dim 원래EnableEvents As Boolean
    Dim 원래Calculation As XlCalculation

    시트명 = "생산일보"

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(시트명)
    If ws Is Nothing Then
        MsgBox "시트 '" & 시트명 & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    원래ScreenUpdating = Application.ScreenUpdating
    원래EnableEvents = Application.EnableEvents
    원래Calculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 시트 보호 해제 (암호 0205) 후 작업, 완료 후 재설정
    Dim 시보호 As Boolean: 시보호 = ws.ProtectContents
    If 시보호 Then On Error Resume Next: ws.Unprotect "0205": If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0

    ' [유지] 주/야간 원지사용 내역 - 삭제하지 않음
    ' C7:BV16, C20:BV29, C35:BV44, C48:BV57

    ' === 2. 완제품 가공내역 ===
    Call 영역지우기_병합허용(ws, "B68:T77", True)
    Call 영역지우기_병합허용(ws, "AK68:BF77", True)

    ' === 3. 작업시간/인력투입 ===
    Call 영역지우기_병합허용(ws, "G108:BM108", True)
    Call 영역지우기_병합허용(ws, "G109:Q109", True)
    Call 영역지우기_병합허용(ws, "G111:BM111", True)
    Call 영역지우기_병합허용(ws, "G112:BM112", True)
    Call 영역지우기_병합허용(ws, "G116:K116", True)

    ' === 4. T10 원자재수율 - 값입력 영역만 삭제 ===
    Call 영역지우기_병합허용(ws, "G88:K91", True)
    Call 영역지우기_병합허용(ws, "G93:K96", True)
    Call 영역지우기_병합허용(ws, "G98:K101", True)
    Call 영역지우기_병합허용(ws, "O88:O91", True)
    Call 영역지우기_병합허용(ws, "O93:O96", True)
    Call 영역지우기_병합허용(ws, "O98:O101", True)
    Call 영역지우기_병합허용(ws, "AP88:AT91", True)
    Call 영역지우기_병합허용(ws, "AP93:AT96", True)
    Call 영역지우기_병합허용(ws, "AP98:AT101", True)
    Call 영역지우기_병합허용(ws, "AX88:AX91", True)
    Call 영역지우기_병합허용(ws, "AX93:AX96", True)

    ' === 5. 비가동현황 ===
    Call 영역지우기_병합허용(ws, "A122:AG126", True)
    Call 영역지우기_병합허용(ws, "AL122:BV126", True)

exitSub:
    If Not ws Is Nothing And 시보호 Then On Error Resume Next: ws.Protect Password:="0205", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True: On Error GoTo 0
    Application.Calculation = 원래Calculation
    Application.EnableEvents = 원래EnableEvents
    Application.ScreenUpdating = 원래ScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "초기화_원지유지 중 오류: " & Err.Description, vbCritical
    Resume exitSub
End Sub
