Attribute VB_Name = "Module24"
' ============================================================
' Module24: 일일생산DB → 생산일보 시트 역변환 (현재 시트에만 채우기)
' ============================================================
' [사용법]
'   - 매크로 "역변환_생산일보시트_채우기" 실행
'   - UserForm1에서 날짜 선택 후 [Run macro]
'   - 선택한 날짜의 DB 데이터를 현재 통합문서의 생산일보 시트에 채움 (새 워크북 생성 없음)
'
' [참조] Module10.M10_FillSheetFromDB → 일일생산DB_역변환_생산일보시트_채우기
' ============================================================

Option Explicit

Public Sub 역변환_생산일보시트_채우기()
    UserForm1.Show vbModal
    If UserForm1.Cancelled Then Exit Sub
    Application.Run "M10_FillSheetFromDB", UserForm1.SelectedDate
End Sub

' ============================================================
' 선택 일자 데이터 완전 삭제 (CommandButton3 호출)
' ============================================================
' 일일생산DB에서 선택한 날짜의 모든 행 삭제. 2단계 확인 후 실행.
Public Sub 선택일자_데이터_완전삭제(선택날짜 As Date)
    Const 시트명DB As String = "일일생산DB"
    Const 열_생산일 As Long = 1
    Dim wsDB As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim 삭제건수 As Long
    Dim db보호 As Boolean

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets(시트명DB)
    On Error GoTo 0
    If wsDB Is Nothing Then
        MsgBox "시트 '" & 시트명DB & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = wsDB.Cells(wsDB.Rows.Count, 열_생산일).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "일일생산DB에 삭제할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' 1차 확인: 정말 삭제하는지
    If MsgBox("선택한 날짜(" & Format(선택날짜, "yyyy-mm-dd") & ")의 데이터를 정말 삭제하시겠습니까?", _
              vbYesNo + vbQuestion, "삭제 확인") <> vbYes Then
        Exit Sub
    End If

    ' 2차 확인: 되돌릴 수 없음 경고
    If MsgBox("경고: 이 작업은 되돌릴 수 없습니다." & vbCrLf & vbCrLf & _
              "정말 삭제를 진행하시겠습니까?", vbYesNo + vbExclamation, "최종 확인") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrDel

    ' 시트 보호 해제 (암호 0205)
    db보호 = wsDB.ProtectContents
    If db보호 Then
        On Error Resume Next
        wsDB.Unprotect "0205"
        If wsDB.ProtectContents Then wsDB.Unprotect
        On Error GoTo ErrDel
    End If

    삭제건수 = 0
    For r = lastRow To 2 Step -1
        If Not IsEmpty(wsDB.Cells(r, 열_생산일).Value) Then
            If IsDate(wsDB.Cells(r, 열_생산일).Value) Then
                If CLng(CDate(wsDB.Cells(r, 열_생산일).Value)) = CLng(선택날짜) Then
                    wsDB.Rows(r).Delete Shift:=xlUp
                    삭제건수 = 삭제건수 + 1
                End If
            End If
        End If
    Next r

    ' 시트 보호 복원
    If db보호 Then
        wsDB.Protect Password:="0205", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If

    Application.ScreenUpdating = True
    MsgBox "삭제 완료: " & 삭제건수 & "건이 삭제되었습니다.", vbInformation, "데이터 삭제"
    Exit Sub

ErrDel:
    Application.ScreenUpdating = True
    If db보호 And Not wsDB Is Nothing Then
        On Error Resume Next
        wsDB.Protect Password:="0205", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    MsgBox "삭제 중 오류: " & Err.Description, vbCritical, "오류"
End Sub
