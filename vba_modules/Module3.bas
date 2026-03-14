Attribute VB_Name = "Module3"
' ============================================================
' 최종 검수테스트 - 1/1~2/2 목업 데이터 생성 (보고서는 Module23)
' ============================================================
' [사용법]
'   1. 이 모듈을 VBA 프로젝트에 추가 (표준 모듈)
'   2. '최종_검수테스트' 매크로 실행
'
' [수행 단계]
'   Phase 0: 일일생산DB 초기화 (부자재관리·원단관리 시트는 유지, 기존 데이터·날짜 보존)
'   Phase 1: 1/1~2/2 (33일) 목업 데이터 생성 → DB 저장 반복
'   Phase 2: 부자재관리 + 원단관리 데이터업데이트 (무음)
'   Phase 3: (제거됨 - 보고서는 Module23.보고서_출력 사용)
'
' [전제 조건]
'   - 생산일보 시트 존재
'   - 생산일보목업데이터 모듈 (T03/T10 수식 보호)
'   - 일일생산일보_최적화DB 모듈
' ============================================================

Option Explicit

Public g_보호생략 As Boolean
Public g_무음모드 As Boolean
Public g_검수테스트_DB로드행제한 As Long

Private Const T_년 As Long = 2026
Private Const T_시작월 As Long = 1
Private Const T_시작일수 As Long = 1
Private Const T_종료월 As Long = 2
Private Const T_종료일수 As Long = 2
Private Const PA_허용오차 As Double = 1#

Private Function SDbl(v As Variant) As Double
    On Error Resume Next
    SDbl = 0
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then SDbl = CDbl(v)
    If Err.Number <> 0 Then SDbl = 0: Err.Clear
End Function

Private Function 주차번호(d As Date) As Long
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    If d < 첫일요일 Then 주차번호 = 1: Exit Function
    주차번호 = 2 + CLng((d - 첫일요일) \ 7)
End Function

Private Function 주차키(d As Date) As String
    주차키 = Month(d) & "월" & 주차번호(d) & "주"
End Function

Private Function 주차_첫날(d As Date) As Date
    Dim w As Long: w = 주차번호(d)
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    If w = 1 Then 주차_첫날 = 월1일: Exit Function
    주차_첫날 = 첫일요일 + CLng(w - 2) * 7
End Function

Private Function 주차_마지막일(d As Date) As Date
    Dim w As Long: w = 주차번호(d)
    Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
    Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
    Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
    Dim 말일 As Date: 말일 = DateSerial(Year(d), Month(d) + 1, 0)
    If w = 1 Then 주차_마지막일 = 첫일요일 - 1: Exit Function
    Dim 끝 As Date: 끝 = 첫일요일 + CLng(w - 2) * 7 + 6
    If 끝 > 말일 Then 끝 = 말일
    주차_마지막일 = 끝
End Function

Private Function 월키(d As Date) As String
    월키 = Month(d) & "월"
End Function

Private Sub 헤더서식(rng As Range)
    With rng
        .Font.Bold = True
        .Font.color = vbWhite
        .Interior.color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' 일일생산DB에 완제품(레코드유형="완제품") 행이 1건 이상 있는지 확인
' DB 구조: A(1)=생산일, B(2)=ID, C(3)=생산일보구분, D(4)=레코드_유형
Private Function DB에완제품있음(wsDB As Worksheet) As Boolean
    Dim lastRow As Long, r As Long, v As Variant
    If wsDB Is Nothing Then Exit Function
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    For r = 2 To lastRow
        v = wsDB.Cells(r, 4).Value
        If Not IsEmpty(v) Then
            If Trim(Replace(Replace(CStr(v), vbLf, ""), vbTab, "")) = "완제품" Then
                DB에완제품있음 = True
                Exit Function
            End If
        End If
    Next r
End Function

' ============================================================
' ★ 메인 진입점
' ============================================================
Public Sub 최종_검수테스트()
    Dim t0 As Double: t0 = Timer
    Dim tPhase As Double

    Dim ws일보 As Worksheet, wsDB As Worksheet, ws재고 As Worksheet, ws원단 As Worksheet
    On Error Resume Next
    Set ws일보 = ThisWorkbook.Worksheets("생산일보")
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    Set ws재고 = ThisWorkbook.Worksheets("부자재관리")
    Set ws원단 = ThisWorkbook.Worksheets("원단관리")
    On Error GoTo 0

    If ws일보 Is Nothing Then
        MsgBox "'생산일보' 시트가 없습니다.", vbCritical, "검수테스트"
        Exit Sub
    End If

    Dim 시작일 As Date: 시작일 = DateSerial(T_년, T_시작월, T_시작일수)
    Dim 종료일 As Date: 종료일 = DateSerial(T_년, T_종료월, T_종료일수)
    Dim 총일수 As Long: 총일수 = CLng(종료일 - 시작일) + 1

    If MsgBox("최종 검수테스트를 시작합니다." & vbCrLf & vbCrLf & _
        chR(9654) & " 기간: " & Format(시작일, "yyyy-mm-dd") & " ~ " & Format(종료일, "yyyy-mm-dd") & " (" & 총일수 & "일)" & vbCrLf & _
        chR(9654) & " 일일생산DB가 초기화됩니다" & vbCrLf & _
        chR(9654) & " 목업 데이터 생성 + DB 저장 (" & 총일수 & "회)" & vbCrLf & _
        chR(9654) & " 부자재관리 시트 자동 업데이트" & vbCrLf & vbCrLf & _
        "계속하시겠습니까?", vbYesNo + vbExclamation, "최종 검수테스트") = vbNo Then
        Exit Sub
    End If

    Dim 보호영역OK As Boolean: 보호영역OK = True
    Dim 원래BK4 As Variant: 원래BK4 = ws일보.Range("BK4").Value
    Dim 원래Screen As Boolean: 원래Screen = Application.ScreenUpdating
    Dim 원래Events As Boolean: 원래Events = Application.EnableEvents
    Dim 원래Calc As Long: 원래Calc = Application.Calculation
    Dim 원래Alerts As Boolean: 원래Alerts = Application.DisplayAlerts

    On Error GoTo TestErr
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    g_보호생략 = True
    g_무음모드 = True

    ' ???????????????????????????????????????????
    ' Phase 0: 전체 초기화 (일일생산DB + 부자재관리 + 관련 시트)
    ' ???????????????????????????????????????????
    tPhase = Timer
    Application.StatusBar = "Phase 0: 전체 초기화..."

    ' ── 0-1. 일일생산DB 초기화 ──
    If wsDB Is Nothing Then
        Dim wb As Workbook: Set wb = ThisWorkbook
        Set wsDB = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDB.Name = "일일생산DB"
        On Error Resume Next
        Application.Run "최적화DB시트_생성"
        On Error GoTo TestErr
    Else
        Dim lr As Long: lr = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lr >= 2 Then wsDB.Rows("2:" & lr).Delete
    End If

    ' ── 0-2. 부자재관리 시트 초기화 제거 (기존 데이터·날짜 유지, Phase 2 데이터업데이트가 증분 반영)
    Dim wsClean As Worksheet
    On Error Resume Next
    Set ws재고 = ThisWorkbook.Worksheets("부자재관리")
    On Error GoTo TestErr
    ' 일일생산DB 시트 초기화 + 헤더 생성 (암호 0205 해제 후 삭제, 삭제 후 재설정)
    Set wsClean = Nothing
    Set wsClean = ThisWorkbook.Worksheets("일일생산DB")
    If Not wsClean Is Nothing Then
        Dim db보호테스트 As Boolean: db보호테스트 = wsClean.ProtectContents
        If db보호테스트 Then On Error Resume Next: wsClean.Unprotect "0205": If wsClean.ProtectContents Then wsClean.Unprotect: On Error GoTo 0
        Dim lrDB2 As Long: lrDB2 = wsClean.Cells(wsClean.Rows.Count, 1).End(xlUp).Row
        If lrDB2 >= 2 Then wsClean.Rows("2:" & lrDB2).Delete
        If db보호테스트 Then wsClean.Protect Password:="0205", DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    ' 일일생산DB 1행 헤더 생성 (최초 실행 또는 헤더 누락 시)
    Application.Run "일일생산DB_헤더갱신"
    ' 동기화상태 시트 초기화 (마지막 동기화 ID 리셋)
    Set wsClean = Nothing
    Set wsClean = ThisWorkbook.Worksheets("부자재관리_동기화상태")
    If Not wsClean Is Nothing Then wsClean.Range("A1").Value = 0
    Set wsClean = ThisWorkbook.Worksheets("원단관리_동기화상태")
    If Not wsClean Is Nothing Then wsClean.Range("A1").Value = 0
    On Error GoTo TestErr

    Debug.Print "Phase 0: " & Format(Timer - tPhase, "0.0") & "s (DB+동기화 초기화, 부자재관리·원단관리 유지)"

    ' ???????????????????????????????????????????
    ' Phase 1: 33일 목업 + DB 저장
    ' ???????????????????????????????????????????
    tPhase = Timer
    Dim 일별날짜() As Date:    ReDim 일별날짜(1 To 총일수)
    Dim 일별DB행() As Long:    ReDim 일별DB행(1 To 총일수)
    Dim 일별성공() As Boolean: ReDim 일별성공(1 To 총일수)
    Dim 일별오류() As String:  ReDim 일별오류(1 To 총일수)

    Dim d As Date, idx As Long: idx = 0
    For d = 시작일 To 종료일
        idx = idx + 1
        일별날짜(idx) = d

        Application.StatusBar = "Phase 1: " & Format(d, "mm/dd") & " (" & idx & "/" & 총일수 & ")"
        DoEvents

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False

        ' 다음 데이터 생성 전 셀 비우기 (이전 일자 데이터 잔여 방지)
        On Error Resume Next: Application.Run "메뉴초기화": Err.Clear: On Error GoTo TestErr

        ' 목업은 BK4를 읽어 +1일 하므로, d-1 설정 → 목업이 d로 생성
        ws일보.Range("BK4").Value = d - 1

        On Error Resume Next: Err.Clear
        Application.Run "생산일보목업데이터_생성_v2"
        If Err.Number <> 0 Then
            일별오류(idx) = "목업: " & Err.Description
            일별성공(idx) = False: Err.Clear
            On Error GoTo TestErr: GoTo NextDay
        End If
        On Error GoTo TestErr

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

        일별DB행(idx) = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row - 1
        일별성공(idx) = True
NextDay:
    Next d
    ' ★ Phase 1 완료 후: DB 전체 날짜 열 NumberFormat 강제 적용 (안전장치)
    Dim dbForceLastRow As Long: dbForceLastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If dbForceLastRow >= 2 Then
        wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(dbForceLastRow, 1)).NumberFormat = "yyyy-mm-dd"
    End If

    Debug.Print "Phase 1: " & Format(Timer - tPhase, "0.0") & "s (" & 총일수 & "일)"

    ' ★ T03(G30) / T10(G87) 수식 보호 검증 (목업이 수식을 덮어쓰지 않았는지 확인)
    If ws일보.Range("G30").HasFormula = False Then 보호영역OK = False
    If ws일보.Range("G87").HasFormula = False Then 보호영역OK = False
    If 보호영역OK = False Then Debug.Print "경고: T03(G30) 또는 T10(G87) 수식이 덮어쓰기됨"

    Dim dbLastRow As Long: dbLastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Phase 1: DB " & (dbLastRow - 1) & "건"

    ' ???????????????????????????????????????????
    ' Phase 2: 부자재관리 + 원단관리 데이터업데이트
    ' ???????????????????????????????????????????
    tPhase = Timer
    DoEvents

    Application.EnableEvents = True
    Application.ScreenUpdating = False
    Application.StatusBar = "Phase 2: 부자재관리·원단관리 데이터업데이트..."

    Dim phase2Err As Long, phase2Desc As String
    Dim phase2원단Err As Long, phase2원단Desc As String
    phase2Err = 0: phase2Desc = ""
    phase2원단Err = 0: phase2원단Desc = ""
    If Not DB에완제품있음(wsDB) Then
        phase2Err = 1002
        phase2Desc = "일일생산DB에 완제품이 없어 Phase 2 건너뜀 (목업→DB저장 확인 필요)"
        Debug.Print "Phase 2 사전검사: " & phase2Desc
    Else
        On Error Resume Next: Err.Clear
        Application.Run "데이터업데이트_무음실행"
        phase2Err = Err.Number
        phase2Desc = Err.Description
        Err.Clear
        On Error GoTo TestErr
        ' 원단관리 데이터업데이트 (원단관리 시트 존재 시)
        If Not ws원단 Is Nothing Then
            On Error Resume Next: Err.Clear
            Application.Run "원단관리_데이터업데이트_무음실행"
            phase2원단Err = Err.Number
            phase2원단Desc = Err.Description
            Err.Clear
            On Error GoTo TestErr
            If phase2원단Err <> 0 Then Debug.Print "Phase 2 원단관리 오류(" & phase2원단Err & "): " & phase2원단Desc
        End If
    End If

    If phase2Err <> 0 Then
        Debug.Print "Phase 2 오류(" & phase2Err & "): " & phase2Desc
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Debug.Print "Phase 2: " & Format(Timer - tPhase, "0.0") & "s"
    If phase2Err <> 0 Then Debug.Print "Phase 2 오류: " & phase2Desc

    ws일보.Range("BK4").Value = 원래BK4
    GoTo TestDone

TestErr:
    g_보호생략 = False: g_무음모드 = False
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    ws일보.Range("BK4").Value = 원래BK4
    On Error GoTo 0
    MsgBox "검수테스트 중 오류:" & vbCrLf & errMsg, vbCritical, "검수테스트"

TestDone:
    g_보호생략 = False: g_무음모드 = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = 원래Events
    Application.DisplayAlerts = 원래Alerts
    Application.ScreenUpdating = 원래Screen

    Dim 소요시간 As String: 소요시간 = Format(Timer - t0, "#,##0.0")
    Dim 성공수 As Long: 성공수 = 0
    For idx = 1 To 총일수
        If 일별성공(idx) Then 성공수 = 성공수 + 1
    Next idx

    Dim 보호메시지 As String
    If 보호영역OK Then 보호메시지 = "" Else 보호메시지 = vbCrLf & "경고: T03(G30)/T10(G87) 수식 덮어쓰기됨"
    MsgBox "최종 검수테스트 완료!" & vbCrLf & vbCrLf & _
        "소요시간: " & 소요시간 & "초" & vbCrLf & _
        "성공: " & 성공수 & " / " & 총일수 & "일" & vbCrLf & _
        "DB 총 레코드: " & 일별DB행(총일수) & vbCrLf & _
        "보호영역(T03/T10 수식): " & IIf(보호영역OK, "PASS", "FAIL") & 보호메시지 & vbCrLf & vbCrLf & _
        "※ 보고서는 Module23.보고서_출력 매크로로 별도 실행", vbInformation, "검수테스트 완료"
End Sub
