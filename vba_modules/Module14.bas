Attribute VB_Name = "Module14"
'================================================================================
' 부자재관리 시트 - 그룹행 삭제 매크로
' 커서가 있는 그룹행(15행 단위) 전체를 삭제. 예=완전삭제, 취소=동작 없음.
' 부자재관리.txt 모듈과 동일 상수 사용 시 해당 모듈에 붙여넣기 후 호출 가능.
'================================================================================

Option Explicit

' 부자재관리 시트 그룹 구조 (부자재관리.txt와 동일)
Private Const WS_부자재관리 As String = "부자재관리"
Private Const WS_일일생산DB As String = "일일생산DB"
Private Const START_ROW As Integer = 4
Private Const ROWS_PER_ITEM As Integer = 15
Private Const COL_구분 As Integer = 13
Private Const COL_일별시작 As Integer = 15
Private Const COL_합계 As Integer = 45
' 부자재관리 A~I (삭제 그룹 키)
Private Const COL_업체명 As Integer = 2
Private Const COL_제품 As Integer = 3
Private Const COL_제조사 As Integer = 4
Private Const COL_지관 As Integer = 5
Private Const COL_지폭 As Integer = 6
Private Const COL_총량 As Integer = 7
Private Const COL_1차포장 As Integer = 8
Private Const COL_2차포장 As Integer = 9
' 일일생산DB B~I (동일 키)
Private Const RDB_업체명 As Integer = 2
Private Const WS_시트보호암호 As String = "0205"
Private Const RDB_제품 As Integer = 3
Private Const RDB_제조사 As Integer = 4
Private Const RDB_지관 As Integer = 5
Private Const RDB_지폭 As Integer = 6
Private Const RDB_총량 As Integer = 7
Private Const RDB_1차포장 As Integer = 8
Private Const RDB_2차포장 As Integer = 9

' 마지막 일 열 번호 (2·3행 기준)
Private Function 마지막일열_찾기(ws As Worksheet) As Long
    Dim c2 As Long, c3 As Long
    On Error Resume Next
    c2 = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    c3 = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    마지막일열_찾기 = WorksheetFunction.Max(c2, c3)
    If 마지막일열_찾기 < COL_일별시작 Then 마지막일열_찾기 = COL_합계
    On Error GoTo 0
End Function

' 그룹행 교대 음영: 1·3·5번째 그룹 RGB(242,247,252) + 아래 굵은 검정 실선, 2·4번째 무음영
Private Sub 그룹행_교대음영_적용(ws As Worksheet, 마지막행 As Long)
    Dim 시작행 As Long, grpIdx As Long, 마지막일열 As Long
    Dim rngOdd As Range, rngEven As Range, rngBorder As Range
    If ws Is Nothing Or 마지막행 < START_ROW Then Exit Sub
    마지막일열 = 마지막일열_찾기(ws)
    If 마지막일열 < COL_일별시작 Then Exit Sub
    ws.Range(ws.Cells(START_ROW, COL_일별시작), ws.Cells(마지막행, 마지막일열)).Font.color = RGB(0, 0, 0)
    For 시작행 = START_ROW To 마지막행 - (ROWS_PER_ITEM - 1) Step ROWS_PER_ITEM
        grpIdx = ((시작행 - START_ROW) \ ROWS_PER_ITEM) + 1
        If grpIdx Mod 2 = 1 Then
            If rngOdd Is Nothing Then Set rngOdd = ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)) Else Set rngOdd = Union(rngOdd, ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)))
        Else
            If rngEven Is Nothing Then Set rngEven = ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)) Else Set rngEven = Union(rngEven, ws.Range(ws.Cells(시작행, 1), ws.Cells(시작행 + ROWS_PER_ITEM - 1, 마지막일열)))
        End If
        If rngBorder Is Nothing Then Set rngBorder = ws.Cells(시작행 + ROWS_PER_ITEM - 1, 1) Else Set rngBorder = Union(rngBorder, ws.Cells(시작행 + ROWS_PER_ITEM - 1, 1))
    Next 시작행
    If Not rngOdd Is Nothing Then With rngOdd.Interior: .Pattern = xlSolid: .color = RGB(242, 247, 252): End With
    If Not rngEven Is Nothing Then rngEven.Interior.Pattern = xlPatternNone
    If Not rngBorder Is Nothing Then
        Dim area As Range
        For Each area In rngBorder.Areas
            With ws.Range(ws.Cells(area.Row, 1), ws.Cells(area.Row, 마지막일열)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThick: .color = RGB(0, 0, 0)
            End With
        Next area
    End If
End Sub

' 일일생산DB에서 동일 키(업체·제품·제조사·지관·지폭·총량·1차·2차) 행 전부 삭제 후 현재고 시트 갱신
Private Sub 일일생산DB_현재고_삭제동기화(ByVal v업체 As String, ByVal v제품 As String, ByVal v제조사 As String, _
    ByVal v지관 As String, ByVal v지폭 As String, ByVal v총량 As String, ByVal v1차 As String, ByVal v2차 As String)
    Dim wsDB As Worksheet
    Dim r As Long, lastRow As Long
    Dim db보호 As Boolean
    On Error Resume Next
    Set wsDB = ThisWorkbook.Sheets(WS_일일생산DB)
    On Error GoTo 0
    If wsDB Is Nothing Then Exit Sub
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    db보호 = wsDB.ProtectContents
    If db보호 Then On Error Resume Next: wsDB.Unprotect WS_시트보호암호: If wsDB.ProtectContents Then wsDB.Unprotect: On Error GoTo 0
    ' 아래에서 위로 순회하며 일치 행 삭제
    For r = lastRow To 2 Step -1
        If Trim(CStr(wsDB.Cells(r, RDB_업체명).Value)) = v업체 And _
           Trim(CStr(wsDB.Cells(r, RDB_제품).Value)) = v제품 And _
           Trim(CStr(wsDB.Cells(r, RDB_제조사).Value)) = v제조사 And _
           Trim(CStr(wsDB.Cells(r, RDB_지관).Value)) = v지관 And _
           Trim(CStr(wsDB.Cells(r, RDB_지폭).Value)) = v지폭 And _
           Trim(CStr(wsDB.Cells(r, RDB_총량).Value)) = v총량 And _
           Trim(CStr(wsDB.Cells(r, RDB_1차포장).Value)) = v1차 And _
           Trim(CStr(wsDB.Cells(r, RDB_2차포장).Value)) = v2차 Then
            wsDB.Rows(r).Delete Shift:=xlUp
        End If
    Next r
    If db보호 Then wsDB.Protect Password:=WS_시트보호암호, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ' 현재고 시트는 DB 기준으로 다시 채우므로 갱신 (같은 통합문서에 부자재관리 모듈·현재고_시트_갱신 있으면 실행)
    On Error Resume Next
    Application.Run "현재고_시트_갱신"
    On Error GoTo 0
End Sub

'================================================================================
' 그룹행 삭제 (부자재관리 시트 전용)
' - 활성 셀 위치의 그룹(15행) 전체를 삭제
' - 경고: "삭제하면 되돌릴 수 없습니다" → 예=삭제, 취소/아니오=취소
' - 일일생산DB·현재고 시트 즉시 동기화
'================================================================================
Public Sub 그룹행_삭제()
    Dim ws As Worksheet
    Dim 현재행 As Long, 그룹시작행 As Long, 그룹끝행 As Long
    Dim 마지막행 As Long
    Dim 응답 As VbMsgBoxResult
    Dim 시트보호상태 As Boolean
    Dim 화면갱신 As Boolean, 자동계산 As Long
    Dim v업체 As String, v제품 As String, v제조사 As String, v지관 As String, v지폭 As String, v총량 As String, v1차 As String, v2차 As String
    
    On Error GoTo ErrHandler
    
    ' 부자재관리 시트인지 확인
    If ActiveSheet.Name <> WS_부자재관리 Then
        MsgBox "부자재관리 시트에서만 실행할 수 있습니다.", vbExclamation, "그룹행 삭제"
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    현재행 = ActiveCell.Row
    
    ' 데이터 영역(4행 이하)인지 확인
    If 현재행 < START_ROW Then
        MsgBox "데이터 영역(4행 이하)에서 커서를 놓고 실행해 주세요.", vbExclamation, "그룹행 삭제"
        Exit Sub
    End If
    
    ' 현재 행이 속한 그룹의 시작행·끝행 계산 (그룹당 15행)
    그룹시작행 = START_ROW + ((현재행 - START_ROW) \ ROWS_PER_ITEM) * ROWS_PER_ITEM
    그룹끝행 = 그룹시작행 + ROWS_PER_ITEM - 1
    
    ' 시트 마지막 데이터 행 확인 (M열 기준)
    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 그룹시작행 > 마지막행 Then
        MsgBox "데이터가 없는 영역입니다.", vbExclamation, "그룹행 삭제"
        Exit Sub
    End If
    
    ' 경고 메시지: 되돌릴 수 없음 → 예=완전삭제, 취소=취소
    응답 = MsgBox("삭제하면 되돌릴 수 없습니다." & vbCrLf & vbCrLf & _
                  "해당 그룹행(" & 그룹시작행 & "~" & 그룹끝행 & "행)을 완전히 삭제하시겠습니까?", _
                  vbYesNoCancel + vbExclamation, "그룹행 삭제")
    
    If 응답 <> vbYes Then
        Exit Sub
    End If
    
    ' 삭제할 그룹의 키 저장 (현재고/DB 동기화용)
    v업체 = Trim(CStr(ws.Cells(그룹시작행, COL_업체명).Value))
    v제품 = Trim(CStr(ws.Cells(그룹시작행, COL_제품).Value))
    v제조사 = Trim(CStr(ws.Cells(그룹시작행, COL_제조사).Value))
    v지관 = Trim(CStr(ws.Cells(그룹시작행, COL_지관).Value))
    v지폭 = Trim(CStr(ws.Cells(그룹시작행, COL_지폭).Value))
    v총량 = Trim(CStr(ws.Cells(그룹시작행, COL_총량).Value))
    v1차 = Trim(CStr(ws.Cells(그룹시작행, COL_1차포장).Value))
    v2차 = Trim(CStr(ws.Cells(그룹시작행, COL_2차포장).Value))
    
    ' 실행 동작 안 보이게 처리
    화면갱신 = Application.ScreenUpdating
    Application.ScreenUpdating = False
    자동계산 = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' 시트 보호 해제 (음영 적용도 보호 해제 상태에서 해야 함)
    시트보호상태 = ws.ProtectContents
    If 시트보호상태 Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
    
    ' 15행 통째로 삭제 (위로 밀기)
    ws.Range(ws.Cells(그룹시작행, 1), ws.Cells(그룹끝행, 1)).EntireRow.Delete Shift:=xlUp
    
    ' 삭제 후 그룹행별 교대 음영 재적용 (보호 복원 전에 적용)
    마지막행 = ws.Cells(ws.Rows.Count, COL_구분).End(xlUp).Row
    If 마지막행 >= START_ROW Then Call 그룹행_교대음영_적용(ws, 마지막행)
    
    If 시트보호상태 Then ws.Protect Password:=WS_시트보호암호, DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
    
    ' 일일생산DB에서 동일 품목 행 전부 삭제 후 현재고 시트 갱신 (즉시 동기화)
    Call 일일생산DB_현재고_삭제동기화(v업체, v제품, v제조사, v지관, v지폭, v총량, v1차, v2차)
    
    ' 화면/계산 복원
    Application.ScreenUpdating = 화면갱신
    Application.Calculation = 자동계산
    
    MsgBox "해당 그룹행이 삭제되었습니다." & vbCrLf & "일일생산DB·현재고 시트에 반영되었습니다.", vbInformation, "그룹행 삭제"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical, "그룹행 삭제"
End Sub




