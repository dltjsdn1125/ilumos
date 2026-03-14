Attribute VB_Name = "Module1"

' ============================================================
' 메뉴초기화 - 생산일보 시트 지정 영역 제거 (생산일보_셀매핑.json 기준)
' ============================================================
' 값 제거 + 음영(채우기)·메모만 제거 (테두리 등 다른 서식 유지). 매크로 실행 시 화면 갱신/이벤트/계산 일시 중지.
' [삭제 금지] T03 직기소계(30~32행), T06 봉제소계(60~62행) = 수식
' [삭제 금지] T07/T08 완제품 수식열 X68:AF78(직기), BG68:BO78(봉제) = 자동 계산
' [삭제 금지] T09 종별수량(80~82행), T10 원자재수율 = 나머지 셀 수식(절대 삭제 금지)
' [T10 값입력 영역만 삭제] G88:K91,G93:K96,G98:K101, O88:O91,O93:O96,O98:O101, AP88:AT91,AP93:AT96,AP98:AT101, AX88:AX91,AX93:AX96
' [T10 삭제금지 수식행] G87,O87, G92,O92, G97,O97, AP87,AX87, AP92,AX92, AP97,AX97
' [삭제 금지] T11 가중시간(109행), 직기봉제합계(110행) = 수식
' [유효성 보존] 제품 유효성 목록 셀 → ClearContents만 사용
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

Public Sub 메뉴초기화()
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

    ' === 1. 원지사용내역 === (서식·메모 포함 제거)
    Call 영역지우기_병합허용(ws, "C7:BV16", True)
    Call 영역지우기_병합허용(ws, "C20:BV29", True)
    Call 영역지우기_병합허용(ws, "C35:BV44", True)
    Call 영역지우기_병합허용(ws, "C48:BV57", True)

    ' === 2. 완제품 가공내역 ===
    Call 영역지우기_병합허용(ws, "B68:T77", True)
    Call 영역지우기_병합허용(ws, "AK68:BF77", True)

    ' === 4. 작업시간/인력투입 ===
    Call 영역지우기_병합허용(ws, "G108:BM108", True)
    Call 영역지우기_병합허용(ws, "G109:Q109", True)
    Call 영역지우기_병합허용(ws, "G111:BM111", True)
    Call 영역지우기_병합허용(ws, "G112:BM112", True)
    Call 영역지우기_병합허용(ws, "G116:K116", True)

    ' === 3. T10 원자재수율 - 값입력 영역만 삭제 ===
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
    MsgBox "메뉴초기화 중 오류: " & Err.Description, vbCritical
    Resume exitSub
End Sub

' ============================================================
' 생산일보_완전초기화 - 지정 시트를 초기화 상태로 (역변환용)
' 메뉴초기화와 동일한 영역 적용, 시트 인자로 전달받음
' M1_FullInitSheet: Application.Run용 영문별칭
' ============================================================
Public Sub M1_FullInitSheet(ws As Worksheet)
    Call 생산일보_완전초기화(ws)
End Sub
Public Sub 생산일보_완전초기화(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then On Error Resume Next: ws.Unprotect "0205": If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0

    ' 1. 원지사용내역
    Call 영역지우기_병합허용(ws, "C7:BV16")
    Call 영역지우기_병합허용(ws, "C20:BV29")
    Call 영역지우기_병합허용(ws, "C35:BV44")
    Call 영역지우기_병합허용(ws, "C48:BV57")

    ' 2. 완제품 가공내역
    Call 영역지우기_병합허용(ws, "B68:T77")
    Call 영역지우기_병합허용(ws, "AK68:BF77")

    ' 3. 작업시간/인력투입
    Call 영역지우기_병합허용(ws, "G108:BM108")
    Call 영역지우기_병합허용(ws, "G109:Q109")
    Call 영역지우기_병합허용(ws, "G111:BM111")
    Call 영역지우기_병합허용(ws, "G112:BM112")
    Call 영역지우기_병합허용(ws, "G116:K116")

    ' 4. T10 원자재수율 - 값입력 영역
    Call 영역지우기_병합허용(ws, "G88:K91")
    Call 영역지우기_병합허용(ws, "G93:K96")
    Call 영역지우기_병합허용(ws, "G98:K101")
    Call 영역지우기_병합허용(ws, "O88:O91")
    Call 영역지우기_병합허용(ws, "O93:O96")
    Call 영역지우기_병합허용(ws, "O98:O101")
    Call 영역지우기_병합허용(ws, "AP88:AT91")
    Call 영역지우기_병합허용(ws, "AP93:AT96")
    Call 영역지우기_병합허용(ws, "AP98:AT101")
    Call 영역지우기_병합허용(ws, "AX88:AX91")
    Call 영역지우기_병합허용(ws, "AX93:AX96")

    ' 5. 비가동현황
    Call 영역지우기_병합허용(ws, "A122:AG126")
    Call 영역지우기_병합허용(ws, "AL122:BV126")

    ' 6. 작성일
    Call 영역지우기_병합허용(ws, "BK4")

    ' 역변환용: 이후 DB데이터_입력이 이어지므로 재보호 안 함
End Sub

