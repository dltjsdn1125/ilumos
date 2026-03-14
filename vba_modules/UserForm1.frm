




' ============================================================
' UserForm1 - 일일생산일보 역변환용 날짜 선택
' ============================================================
' [설치]
' 1. VBA 편집기(Alt+F11) → 삽입 → 사용자 정의 폼 → 이름을 반드시 "UserForm1"로
' 2. 도구상자에서 ListBox 추가 → 속성 (이름) = ListBox1
' 3. CommandButton 3개 추가 (기본 이름 사용):
'    - CommandButton1: Caption = Run macro
'    - CommandButton2: Caption = Cancel
'    - CommandButton3: Caption = 삭제
' 4. UserForm1 더블클릭 → 아래 코드 전체 붙여넣기
'
' [동작]
' - 폼 로드 시 일일생산DB 시트 A열(생산일)에서 출력 가능한 날짜만 수집
' - ListBox1에 "2026년2월11일 월요일" 형식으로 가로 표시 (최신순)
' - 날짜 선택 후 [Run macro] 클릭 → 해당 날짜로 역변환 매크로 실행
' - [삭제] 클릭 → 선택한 날짜의 일일생산DB 데이터 완전 삭제
' ============================================================

Option Explicit

' 선택된 날짜 (실행 시 설정, 취소 시 Empty)
Public SelectedDate As Variant
Public Cancelled As Boolean

' ListBox 인덱스 → 실제 날짜 매핑용
Private 날짜목록() As Date

Private Const 요일한글 As String = "월요일 화요일 수요일 목요일 금요일 토요일 일요일"

' ============================================================
' 날짜 → "2026년2월11일 월요일" 형식
' ============================================================
Private Function 날짜표시형식(d As Date) As String
    Dim 요일인덱스 As Long
    Dim 요일배열 As Variant
    요일배열 = Split(요일한글, " ")
    요일인덱스 = Weekday(d, vbMonday)  ' 1=월 ~ 7=일
    날짜표시형식 = Year(d) & "년" & Month(d) & "월" & Day(d) & "일 " & 요일배열(요일인덱스 - 1)
End Function

' ============================================================
' 폼 로드: 일일생산DB A열(생산일)에서 유일한 날짜 수집 → ListBox1 채우기
' ============================================================
Private Sub UserForm_Initialize()
    Dim wsDB As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dict As Object
    Dim key As Variant
    Dim d As Date
    Dim i As Long, n As Long
    Dim 정렬날짜() As Date

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    On Error GoTo 0

    ' 버튼 캡션 고정
    CommandButton1.Caption = "Run macro"
    CommandButton2.Caption = "Cancel"
    On Error Resume Next
    CommandButton3.Caption = "Delete"
    On Error GoTo 0

    ' ListBox 문자가 세로로 쓰이지 않도록: @폰트(세로) → @ 없는 폰트(가로)로 변경
    On Error Resume Next
    ListBox1.Font.Name = "맑은 고딕"
    If Err.Number <> 0 Then ListBox1.Font.Name = "Gulim"
    If Err.Number <> 0 Then ListBox1.Font.Name = "Arial"
    Err.Clear
    On Error GoTo 0
    ListBox1.ColumnWidths = "280"
    If ListBox1.width < 300 Then ListBox1.width = 300
    ListBox1.ColumnCount = 1

    ListBox1.Clear
    SelectedDate = Empty
    Cancelled = True

    If wsDB Is Nothing Then
        ListBox1.AddItem "(일일생산DB 시트 없음)"
        Exit Sub
    End If

    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ListBox1.AddItem "(데이터 없음)"
        Exit Sub
    End If

    ' A열(1열) 생산일에서 유일한 날짜만 Dictionary로 수집 (B열=ID 이므로 A열 사용)
    Set dict = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        If IsEmpty(wsDB.Cells(r, 1).Value) Then GoTo NextR
        On Error Resume Next
        d = CDate(wsDB.Cells(r, 1).Value)
        On Error GoTo 0
        If Not IsDate(d) Then GoTo NextR
        dict(CStr(CLng(d))) = d   ' 중복 제거 (날짜 직렬값으로 키)
NextR:
    Next r

    If dict.Count = 0 Then
        ListBox1.AddItem "(출력 가능한 날짜 없음)"
        Exit Sub
    End If

    ' 날짜 배열로 옮겨서 정렬 (최신순)
    ReDim 정렬날짜(0 To dict.Count - 1)
    n = 0
    For Each key In dict.keys
        정렬날짜(n) = dict(key)
        n = n + 1
    Next key
    Call 정렬_날짜_내림차순(정렬날짜)

    ReDim 날짜목록(0 To n - 1)
    For i = 0 To n - 1
        날짜목록(i) = 정렬날짜(i)
        ListBox1.AddItem 날짜표시형식(정렬날짜(i))
    Next i
End Sub

' 버블 정렬: 날짜 내림차순 (최신이 위로)
Private Sub 정렬_날짜_내림차순(arr() As Date)
    Dim i As Long, j As Long
    Dim t As Date
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
End Sub

' ============================================================
' [Run macro] 클릭: 선택한 날짜 저장 후 폼 숨김 → 매크로에서 실행
' ============================================================
Private Sub CommandButton1_Click()
    If ListBox1.ListCount = 0 Then
        MsgBox "표시된 날짜가 없습니다.", vbExclamation
        Exit Sub
    End If
    If ListBox1.ListIndex < 0 Then
        MsgBox "날짜를 선택한 뒤 [Run macro]를 눌러 주세요.", vbExclamation
        Exit Sub
    End If
    ' 괄호로 시작하면 안내 문구(데이터 없음 등) → 실행 불가
    If left(Trim(ListBox1.list(ListBox1.ListIndex)), 1) = "(" Then
        MsgBox "출력 가능한 날짜를 선택해 주세요.", vbExclamation
        Exit Sub
    End If
    SelectedDate = 날짜목록(ListBox1.ListIndex)
    Cancelled = False
    Me.Hide
End Sub

' ============================================================
' [Cancel] 클릭
' ============================================================
Private Sub CommandButton2_Click()
    Cancelled = True
    SelectedDate = Empty
    Me.Hide
End Sub

' ============================================================
' [삭제] 클릭 (CommandButton3): 선택한 날짜의 일일생산DB 데이터 완전 삭제
' ============================================================
Private Sub CommandButton3_Click()
    If ListBox1.ListCount = 0 Then
        MsgBox "표시된 날짜가 없습니다.", vbExclamation
        Exit Sub
    End If
    If ListBox1.ListIndex < 0 Then
        MsgBox "날짜를 선택한 뒤 [삭제]를 눌러 주세요.", vbExclamation
        Exit Sub
    End If
    If left(Trim(ListBox1.list(ListBox1.ListIndex)), 1) = "(" Then
        MsgBox "삭제할 날짜를 선택해 주세요.", vbExclamation
        Exit Sub
    End If
    Application.Run "선택일자_데이터_완전삭제", 날짜목록(ListBox1.ListIndex)
    UserForm_Initialize
End Sub
