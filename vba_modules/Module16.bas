Attribute VB_Name = "Module16"
' ============================================================
' 일일생산DB 데이터 정리(백업) 모듈
' ============================================================
' [기능]
' 매크로 실행 → 안내 메시지 → 기록된 모든 데이터 CSV로 바탕화면에 저장 → "엑셀에서 내용을 비우겠습니까?" 확인 → 예 시 일일생산DB 전체 삭제
'
' [일일생산DB] 1행=헤더, 2행~=데이터
' [실행] 매크로 "일일생산DB_백업후_데이터삭제" 실행
' ============================================================

Option Explicit

Private Const 시트명DB As String = "일일생산DB"
Private Const 열_작성일 As Long = 1   ' A열(생산일)
Private Const DB열수 As Long = 70

' ============================================================
' 메인: 안내 → 전체 데이터 CSV 바탕화면 저장 → 내용 비우기 여부 확인 → 삭제
' ============================================================
Public Sub 일일생산DB_백업후_데이터삭제()
    Dim wsDB As Worksheet
    Dim lastRow As Long
    Dim savePath As String
    Dim exportCount As Long

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets(시트명DB)
    On Error GoTo 0
    If wsDB Is Nothing Then
        MsgBox "시트 '" & 시트명DB & "'을(를) 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = wsDB.Cells(wsDB.Rows.Count, 열_작성일).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "일일생산DB에 백업할 데이터가 없습니다.", vbInformation
        Exit Sub
    End If

    ' 1) 안내 메시지
    MsgBox "기록된 모든 데이터를 CSV 파일로 바탕화면에 저장합니다." & vbCrLf & vbCrLf & _
        "이후 엑셀에서 일일생산DB 내용을 비울지 선택할 수 있습니다.", vbInformation, "데이터 백업"

    Application.ScreenUpdating = False
    On Error GoTo ErrBackup

    ' 2) 전체 데이터 CSV 내보내기 (바탕화면에 바로 저장)
    exportCount = 백업_CSV내보내기_전체(wsDB, lastRow, savePath)
    If exportCount <= 0 Then
        MsgBox "내보낼 데이터가 없습니다.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 3) 엑셀에서 내용 비우기 여부 확인
    If MsgBox("CSV 파일을 바탕화면에 저장했습니다." & vbCrLf & vbCrLf & _
        "엑셀에서 일일생산DB 내용을 모두 비우시겠습니까?" & vbCrLf & _
        "(1행 헤더는 유지됩니다)", vbYesNo + vbQuestion, "데이터 비우기") <> vbYes Then
        Application.ScreenUpdating = True
        MsgBox "백업 완료. 저장 위치: " & savePath, vbInformation
        Exit Sub
    End If

    ' 4) 일일생산DB 데이터 행 전체 삭제 (암호 0205 해제 후 삭제, 삭제 후 재설정)
    Dim db보호 As Boolean: db보호 = wsDB.ProtectContents
    If db보호 Then On Error Resume Next: wsDB.Unprotect "0205": If wsDB.ProtectContents Then wsDB.Unprotect: On Error GoTo 0
    wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, DB열수)).Delete Shift:=xlUp
    If db보호 Then wsDB.Protect Password:="0205", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    MsgBox "일일생산DB 데이터를 모두 비웠습니다." & vbCrLf & "백업 파일: " & savePath, vbInformation

    Application.ScreenUpdating = True
    Exit Sub
ErrBackup:
    Application.ScreenUpdating = True
    MsgBox "백업 중 오류: " & Err.Description, vbCritical
End Sub

' 전체 데이터 CSV 내보내기 (바탕화면에 저장). 반환: 행 수. savePath에 저장된 전체 경로 반환.
Private Function 백업_CSV내보내기_전체(wsDB As Worksheet, lastRow As Long, ByRef savePath As String) As Long
    Dim desktopPath As String
    Dim fname As String
    Dim arr As Variant
    Dim f As Long, r As Long

    desktopPath = Environ("USERPROFILE") & "\Desktop"
    If right(desktopPath, 1) <> "\" Then desktopPath = desktopPath & "\"
    fname = "일일생산DB_전체_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    savePath = desktopPath & fname

    arr = wsDB.Range(wsDB.Cells(1, 1), wsDB.Cells(lastRow, DB열수)).Value
    f = FreeFile
    On Error GoTo CSVExportErr
    Open savePath For Output As #f
    For r = 1 To UBound(arr, 1)
        Print #f, 백업_CSV한행만들기(arr, r)
    Next r
    Close #f
    백업_CSV내보내기_전체 = lastRow - 1
    Exit Function
CSVExportErr:
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise Err.Number, , Err.Description
End Function

' 배열의 r행을 CSV 한 줄 문자열로 (쉼표/줄바꿈/따옴표 이스케이프)
Private Function 백업_CSV한행만들기(arr As Variant, r As Long) As String
    Dim c As Long
    Dim s As String, out As String
    out = ""
    For c = 1 To UBound(arr, 2)
        s = CStr(arr(r, c))
        If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Or InStr(s, """") > 0 Then
            s = """" & Replace(s, """", """""") & """"
        End If
        If c > 1 Then out = out & ","
        out = out & s
    Next c
    백업_CSV한행만들기 = out
End Function



