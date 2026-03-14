Attribute VB_Name = "Module15"
'================================================================================
' 부자재관리 시트 키워드 검색 → 해당 셀로 이동 + 연두색 하이라이트
' - 실행 시 UserForm2 열림 → 유저폼에서 키워드 입력 후 검색
' - 검색 범위: A~I열 (제품발주처, 제품명, 포장구분 등). 그룹 단위 6행씩
' - 일치하는 셀로 커서 이동, 옅은 연두색 표시 (커서 이동 시 SelectionChange에서 해제)
' - 중복 일치 시: 검색할 때마다 다음 결과로 순서대로 이동(순환)
'================================================================================
' 설치 방법:
' 1) 아래 [모듈 코드] 를 표준 모듈(예: 모듈1)에 붙여넣기
' 2) 부자재관리 시트 코드 창에 [시트 이벤트 코드] 붙여넣기
'================================================================================

'---------- [모듈 코드] ----------
Option Explicit

' 부자재관리 시트 검색용 상수 (부자재관리.txt와 동일한 값 사용)
Private Const WS_부자재관리 As String = "부자재관리"
Private Const START_ROW As Long = 4
Private Const ROWS_PER_ITEM As Long = 6
Private Const COL_설비라인 As Long = 1    ' A
Private Const COL_업체명 As Long = 2      ' B
Private Const COL_제품 As Long = 3        ' C
Private Const COL_제조사 As Long = 4      ' D
Private Const COL_지관 As Long = 5        ' E
Private Const COL_지폭 As Long = 6        ' F
Private Const COL_평량 As Long = 7        ' G
Private Const COL_1차포장 As Long = 8     ' H
Private Const COL_2차포장 As Long = 9     ' I

' 검색 키워드 셀: 기본 J2. 이름 "검색키워드" 정의 시 해당 셀 사용
Private Const KEYWORD_CELL_DEFAULT As String = "J2"

' 옅은 연두색 (하이라이트용)
Private Const HIGHLIGHT_COLOR As Long = &HCCFFCC   ' RGB(204,255,204)
Private Const WS_시트보호암호 As String = "0205"

' 하이라이트 해제용: 부자재관리 시트 SelectionChange에서 호출
Public g_검색하이라이트셀 As Range
' 검색 매크로 실행 중 플래그 (진행 중에는 SelectionChange에서 하이라이트 해제 방지)
Public 검색_진행중 As Boolean

'================================================================================
' 실행 시 UserForm2 열기 (유저폼에서 키워드 입력 후 검색)
'================================================================================
Public Sub 키워드검색_실행()
    On Error Resume Next
    UserForm2.Show vbModal
    If Err.Number <> 0 Then
        UserForm2.Show
    End If
    On Error GoTo 0
End Sub

' 검색 키워드 입력 셀 반환 (이름 "검색키워드" 있으면 해당 셀, 없으면 부자재관리!J2)
Public Function 부자재관리_검색키워드셀() As Range
    On Error Resume Next
    Set 부자재관리_검색키워드셀 = ThisWorkbook.Names("검색키워드").RefersToRange
    If 부자재관리_검색키워드셀 Is Nothing Then
        Set 부자재관리_검색키워드셀 = ThisWorkbook.Sheets(WS_부자재관리).Range(KEYWORD_CELL_DEFAULT)
    End If
    On Error GoTo 0
End Function

' 0.3초 후 진행중 플래그 해제 (OnTime 예약)
Public Sub 검색플래그_해제()
    검색_진행중 = False
End Sub

'================================================================================
' 하이라이트 해제 (커서 이동 시 시트 이벤트에서 호출)
' 오늘이동 매크로 실행 중에만 해제 방지. 검색 후 커서 이동 시 즉시 음영 해제
'================================================================================
Public Sub 검색_하이라이트_해제()
    Dim wsH As Worksheet, 시보호H As Boolean
    On Error Resume Next
    If 오늘이동_진행중 Then Exit Sub
    On Error GoTo 0
    If g_검색하이라이트셀 Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsH = g_검색하이라이트셀.Parent
    시보호H = wsH.ProtectContents
    If 시보호H Then On Error Resume Next: wsH.Unprotect WS_시트보호암호: If wsH.ProtectContents Then wsH.Unprotect: On Error GoTo 0
    g_검색하이라이트셀.Interior.ColorIndex = xlNone
    Set g_검색하이라이트셀 = Nothing
    If 시보호H Then On Error Resume Next: Application.Run "셀보호_적용_무음", wsH: On Error GoTo 0
    On Error GoTo 0
End Sub

'================================================================================
' 키워드로 A~I 열 검색 → 해당 셀로 이동, 연두색 표시. 중복 시 다음 결과로 순환 이동
' (UserForm2 또는 J2 셀에서 호출)
'================================================================================
Public Sub 부자재관리_키워드검색_실행(Optional ByVal 키워드 As String = "")
    Dim ws As Worksheet
    If Len(키워드) = 0 Then
        키워드 = Trim(CStr(부자재관리_검색키워드셀().Value))
    End If
    부자재관리_키워드검색_내부 키워드
End Sub

' J2 키워드 셀 기준 검색 (Worksheet_Change에서 호출용)
Public Sub 부자재관리_키워드검색_스크롤()
    부자재관리_키워드검색_실행
End Sub

Private Sub 부자재관리_키워드검색_내부(ByVal 키워드 As String)
    Static 검색_마지막키워드 As String
    Static 검색_다음인덱스 As Long
    Dim ws As Worksheet
    Dim 마지막행 As Long
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    Dim MatchRows() As Long
    Dim MatchCols() As Long
    Dim nMatches As Long
    Dim idx As Long
    Dim prevHighlight As Range

    On Error GoTo ErrSearch

    검색_진행중 = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Sheets(WS_부자재관리)
    키워드 = Trim(키워드)
    If Len(키워드) = 0 Then
        검색_진행중 = False
        Call 검색_하이라이트_해제
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If

    ' 키워드가 바뀌면 다음 인덱스 1로 리셋
    If StrComp(키워드, 검색_마지막키워드, vbTextCompare) <> 0 Then
        검색_마지막키워드 = 키워드
        검색_다음인덱스 = 1
    End If

    ' 이전 하이라이트 셀 제거(색만 제거, Select 시 SelectionChange에서 해제되므로 여기서는 다음 하이라이트 전 정리)
    Set prevHighlight = g_검색하이라이트셀
    If Not prevHighlight Is Nothing Then
        On Error Resume Next
        If prevHighlight.Parent.Name = ws.Name Then
            Dim 시보호Prev As Boolean
            시보호Prev = prevHighlight.Parent.ProtectContents
            If 시보호Prev Then On Error Resume Next: prevHighlight.Parent.Unprotect WS_시트보호암호: If prevHighlight.Parent.ProtectContents Then prevHighlight.Parent.Unprotect: On Error GoTo 0
            prevHighlight.Interior.ColorIndex = xlNone
            If 시보호Prev Then On Error Resume Next: Application.Run "셀보호_적용_무음", prevHighlight.Parent: On Error GoTo ErrSearch
        End If
        On Error GoTo ErrSearch
        Set g_검색하이라이트셀 = Nothing
    End If

    마지막행 = ws.Cells(ws.Rows.Count, COL_설비라인).End(xlUp).Row
    If 마지막행 < START_ROW Then 마지막행 = ws.Cells(ws.Rows.Count, COL_제품).End(xlUp).Row
    If 마지막행 < START_ROW Then
        검색_진행중 = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If

    ' 모든 일치 (그룹 첫 행의 A~I) 수집
    nMatches = 0
    For r = START_ROW To 마지막행 Step ROWS_PER_ITEM
        For c = COL_설비라인 To COL_2차포장
            v = ws.Cells(r, c).Value
            If Not IsEmpty(v) Then
                If InStr(1, CStr(v), 키워드, vbTextCompare) > 0 Then
                    nMatches = nMatches + 1
                    If nMatches = 1 Then
                        ReDim MatchRows(1 To 1)
                        ReDim MatchCols(1 To 1)
                    Else
                        ReDim Preserve MatchRows(1 To nMatches)
                        ReDim Preserve MatchCols(1 To nMatches)
                    End If
                    MatchRows(nMatches) = r
                    MatchCols(nMatches) = c
                End If
            End If
        Next c
    Next r

    If nMatches = 0 Then
        검색_진행중 = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "검색어 '" & 키워드 & "'(를)을 포함하는 그룹이 없습니다.", vbInformation, "검색"
        Exit Sub
    End If

    ' 다음에 보여줄 인덱스(1~nMatches, 순환)
    idx = 검색_다음인덱스
    If idx < 1 Or idx > nMatches Then idx = 1
    검색_다음인덱스 = idx + 1
    If 검색_다음인덱스 > nMatches Then 검색_다음인덱스 = 1

    r = MatchRows(idx)
    c = MatchCols(idx)

    On Error Resume Next
    ws.Activate
    Application.GoTo ws.Cells(r, c), True
    If ActiveWindow.Visible Then ActiveWindow.ScrollRow = Application.Max(1, r - 2)
    On Error GoTo ErrSearch
    ' 옅은 연두색 하이라이트 (셀보호 시 일시 해제 후 적용)
    Dim 시보호New As Boolean
    시보호New = ws.ProtectContents
    If 시보호New Then On Error Resume Next: ws.Unprotect WS_시트보호암호: If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
    On Error Resume Next
    With ws.Cells(r, c).Interior
        .color = HIGHLIGHT_COLOR
        .Pattern = xlSolid
    End With
    On Error GoTo ErrSearch
    If 시보호New Then On Error Resume Next: Application.Run "셀보호_적용_무음", ws: On Error GoTo ErrSearch
    Set g_검색하이라이트셀 = ws.Cells(r, c)

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.OnTime Now + (0.3 / 86400), "검색플래그_해제", , True
    If Err.Number <> 0 Then 검색_진행중 = False
    On Error GoTo 0
    Exit Sub
ErrSearch:
    검색_진행중 = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Set g_검색하이라이트셀 = Nothing
    MsgBox "검색 중 오류: " & Err.Description, vbExclamation, "오류"
End Sub

'---------- [시트 이벤트 코드] (부자재관리 시트 모듈에 붙여넣기) ----------
' 기존 Worksheet_Change + J2 키워드 검색 통합 버전
'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim r As Range
    If Target.Cells.Count > 1 Then Set r = Target.Cells(1, 1) Else Set r = Target
    ' B·C열 변경 → D~I 자동 채우기
    If r.Column = 2 Or r.Column = 3 Then
        If r.Row >= 4 Then Call 부자재관리_BC_변경시_DtoI_채우기(r)
        Exit Sub
    End If
    ' 검색 키워드 셀 변경 시 검색 실행 (기본 J2, 이름 "검색키워드" 정의 시 해당 셀)
    If Not Intersect(Target, 부자재관리_검색키워드셀()) Is Nothing Then
        Application.EnableEvents = False
        On Error GoTo SearchErr
        부자재관리_키워드검색_스크롤
        Application.EnableEvents = True
        Exit Sub
SearchErr:
        Application.EnableEvents = True
        MsgBox "키워드 검색 오류: " & Err.Description, vbExclamation, "검색"
        Exit Sub
    End If
    ' J열(전월재고) 또는 N열 이후 → DB 동기화
    If Target.Cells.Count > 1 Then If r.Column <> 10 Then Exit Sub
    If r.Column < 14 And r.Column <> 10 Then Exit Sub
    On Error GoTo ErrorHandler
    Call 부자재관리_셀변경_DB동기화_예약(r)
    Exit Sub
ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "값 변경 처리 중 오류: " & Err.Description, vbExclamation
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next: Call 부자재관리_C열유효성_갱신대기처리: On Error GoTo 0
    Call 오늘이동_음영해제
    On Error Resume Next: Call 검색_하이라이트_해제: On Error GoTo 0
End Sub



