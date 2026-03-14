
Option Explicit

Private Sub CommandButton1_Click()
    Dim 키워드 As String
    키워드 = Trim(CStr(Me.TextBox1.Value))
    If Len(키워드) = 0 Then
        MsgBox "키워드를 입력하세요.", vbExclamation, "키워드 검색"
        Exit Sub
    End If
    On Error Resume Next
    If g_키워드검색_원단모드 Then
        원단관리_키워드검색_셀기반 키워드
    Else
        부자재관리_키워드검색_실행 키워드
    End If
    If Err.Number <> 0 Then
        MsgBox "검색 오류: " & Err.Description, vbExclamation, "키워드 검색"
    End If
    On Error GoTo 0
    ' 폼 유지: 검색 버튼 누를 때마다 다음 일치 결과로 이동
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "키워드 검색"
    Me.TextBox1.Value = ""
End Sub

