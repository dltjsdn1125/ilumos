Attribute VB_Name = "Module10"
' ============================================================
' 일일생산DB → 일일생산일보 역변환 시스템
' ============================================================
' [주요 기능]
' 1. 일일생산DB 1행 헤더 삭제 후 제공 양식에 따라 재구성 (일일생산DB_헤더_재구성)
' 2. 일일생산DB에서 특정 날짜 데이터 추출 → 일일생산일보 셀주소에 맞게 입력
' 3. 수식 값 변환, 파일 저장 (일일생산일지_YYYYMMDD_HHMMSS.xlsb)
'
' [일일생산DB 70열 구조]
' A(1)=ID, B(2)=작성일, C(3)=생산일보구분, D(4)=레코드_유형, E(5)=순번
' F(6)=원지_종별, G(7)=원지_ROLL_ID, H(8)=원지_KG
' I(9)=완제품_구분, J(10)=완제품_품명, K(11)=가공라인호기, L(12)=생산량_BOX,
' M(13)=원지기준소요량, N(14)=PLT1, O(15)=PLT2, P(16)=인원
' Q(17)=인원_작업장, R(18)=인원_구분, S(19)=인원_인원수, T(20)=인원_작업시간,
' U(21)=인원_총작업시간, V(22)=실투입원/야간심야, W(23)=환산투입원/휴일,
' X(24)=인당생산량_실투, Y(25)=인당생산량_환산투
' Z~AI(26~35)=PA_일10개, AJ~AS(36~45)=PA_주10개, AT~BC(46~55)=PA_월10개
' BD(56)=Todays, BE(57)=특이사항, BF(58)=해결책, BG(59)=담당, BH(60)=이사
' BI~BR(61~70)=PA2_구분/기간/원지투입총량~목표인당PLT_월
'
' [생산일보 시트 매핑 - 생산일보_셀매핑.json 기준]
' BK4 = 작성일(날짜)
' T01 원지 주간 KT: rows 7-16, 6그룹 base(3,15,27,39,51,63), ROLLID=base+2, KG=base+9
' T02 원지 직기 JRT: rows 20-29, 동일 구조
' T04 원지 봉제 KT: rows 35-44
' T05 원지 봉제 JRT: rows 48-57
' T07 완제품 주간(68~77행): B(2)=제품, Q(17)=설비, T(20)=생산량BOX
' T08 완제품 야간(68~77행): AK(37)=제품, AZ(52)=설비, BC(55)=생산량BOX
' T11 작업시간(108행), Direct(111행), Indirect(112행)
'     설비 base: KT=7, FT=43, JRT=19, BT=31, PRE=55 (생산일보_셀매핑.json)
'     각 설비: 주간정규=base, 주간연장=base+2, 야간정규=base+4, 야간심야=base+6, 휴일=base+8
' T10 원자재수율(84~101행): 수식값 역변환 입력
' T12 비가동현황: rows 121-126, A(1)=내용, AJ(36)=해결책
'
' DB 생산일보구분: "생산일보_주간"/"생산일보_야간"/"생산일보"
' ============================================================

Option Explicit

' ============================================================
' 일일생산DB 시트 1행 헤더 삭제 후 제공 양식에 따라 재구성
' ============================================================
Public Sub 일일생산DB_헤더_재구성()
    Dim wsDB As Worksheet
    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    On Error GoTo 0
    If wsDB Is Nothing Then
        MsgBox "'일일생산DB' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    wsDB.Rows(1).Clear
    wsDB.Cells(1, 1).Value = "생산일"
    wsDB.Cells(1, 2).Value = "ID"
    wsDB.Cells(1, 3).Value = "생산일보구분"
    wsDB.Cells(1, 4).Value = "레코드_유형"
    wsDB.Cells(1, 5).Value = "순번"
    wsDB.Cells(1, 6).Value = "원지_종별"
    wsDB.Cells(1, 7).Value = "원지_ROLL_ID"
    wsDB.Cells(1, 8).Value = "원지_KG"
    wsDB.Cells(1, 9).Value = "완제품_구분"
    wsDB.Cells(1, 10).Value = "완제품_품명"
    wsDB.Cells(1, 11).Value = "가공라인호기"
    wsDB.Cells(1, 12).Value = "생산량_BOX"
    wsDB.Cells(1, 13).Value = "원지기준소요량"
    wsDB.Cells(1, 14).Value = "PLT1"
    wsDB.Cells(1, 15).Value = "PLT2"
    wsDB.Cells(1, 16).Value = "인원"
    wsDB.Cells(1, 17).Value = "인원_작업장"
    wsDB.Cells(1, 18).Value = "인원_구분"
    wsDB.Cells(1, 19).Value = "인원_인원수"
    wsDB.Cells(1, 20).Value = "인원_작업시간"
    wsDB.Cells(1, 21).Value = "인원_총작업시간"
    wsDB.Cells(1, 22).Value = "실투입원"
    wsDB.Cells(1, 23).Value = "환산투입원"
    wsDB.Cells(1, 24).Value = "인당생산량_실투"
    wsDB.Cells(1, 25).Value = "인당생산량_환산투"
    wsDB.Cells(1, 26).Value = "PA_원지투입_일"
    wsDB.Cells(1, 27).Value = "PA_원지투입_일평균"
    wsDB.Cells(1, 28).Value = "PA_기준소요량_일총량"
    wsDB.Cells(1, 29).Value = "PA_기준소요량_일평균"
    wsDB.Cells(1, 30).Value = "PA_원지수율_일"
    wsDB.Cells(1, 31).Value = "PA_LOSS율_일"
    wsDB.Cells(1, 32).Value = "PA_양품율_일"
    wsDB.Cells(1, 33).Value = "PA_pallet_일"
    wsDB.Cells(1, 34).Value = "PA_목표양품율_일"
    wsDB.Cells(1, 35).Value = "PA_목표인당PLT_일"
    wsDB.Cells(1, 36).Value = "PA_원지투입_주"
    wsDB.Cells(1, 37).Value = "PA_원지투입_주평균"
    wsDB.Cells(1, 38).Value = "PA_기준소요량_주총량"
    wsDB.Cells(1, 39).Value = "PA_기준소요량_주평균"
    wsDB.Cells(1, 40).Value = "PA_원지수율_주"
    wsDB.Cells(1, 41).Value = "PA_LOSS율_주"
    wsDB.Cells(1, 42).Value = "PA_양품율_주"
    wsDB.Cells(1, 43).Value = "PA_pallet_주"
    wsDB.Cells(1, 44).Value = "PA_목표양품율_주"
    wsDB.Cells(1, 45).Value = "PA_목표인당PLT_주"
    wsDB.Cells(1, 46).Value = "PA_원지투입_월"
    wsDB.Cells(1, 47).Value = "PA_원지투입_월평균"
    wsDB.Cells(1, 48).Value = "PA_기준소요량_월총량"
    wsDB.Cells(1, 49).Value = "PA_기준소요량_월평균"
    wsDB.Cells(1, 50).Value = "PA_원지수율_월"
    wsDB.Cells(1, 51).Value = "PA_LOSS율_월"
    wsDB.Cells(1, 52).Value = "PA_양품율_월"
    wsDB.Cells(1, 53).Value = "PA_pallet_월"
    wsDB.Cells(1, 54).Value = "PA_목표양품율_월"
    wsDB.Cells(1, 55).Value = "PA_목표인당PLT_월"
    wsDB.Cells(1, 56).Value = "Todays"
    wsDB.Cells(1, 57).Value = "특이사항"
    wsDB.Cells(1, 58).Value = "해결책"
    wsDB.Cells(1, 59).Value = "담당"
    wsDB.Cells(1, 60).Value = "이사"
    wsDB.Cells(1, 61).Value = "PA2_구분"
    wsDB.Cells(1, 62).Value = "PA2_기간"
    wsDB.Cells(1, 63).Value = "PA2_원지투입총량"
    wsDB.Cells(1, 64).Value = "PA2_양품환산총량"
    wsDB.Cells(1, 65).Value = "PA2_원지수율"
    wsDB.Cells(1, 66).Value = "PA2_LOSS율"
    wsDB.Cells(1, 67).Value = "PA2_목표대비양품율"
    wsDB.Cells(1, 68).Value = "PA2_목표대비PLT"
    wsDB.Cells(1, 69).Value = "PA2_목표양품율_월"
    wsDB.Cells(1, 70).Value = "PA2_목표인당PLT_월"

    MsgBox "일일생산DB 헤더를 재구성했습니다 (1행).", vbInformation
End Sub

' ============================================================
' 일일생산DB: 날짜(작성일) 기준 그룹화 → 교차 음영
' ============================================================
Public Sub 일일생산DB_일별_그룹_색상()
    Dim wsDB As Worksheet
    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    On Error GoTo 0
    If wsDB Is Nothing Then
        MsgBox "'일일생산DB' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim lastCol As Long
    lastCol = wsDB.Cells(1, wsDB.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 70

    Dim 옅은하늘색 As Long
    옅은하늘색 = RGB(230, 245, 255)

    Dim r As Long, grpStart As Long
    Dim prevDate As Variant, thisDate As Variant, useSky As Boolean

    grpStart = 2
    prevDate = wsDB.Cells(2, 1).Value
    useSky = True

    For r = 2 To lastRow
        thisDate = wsDB.Cells(r, 1).Value
        If r > 2 Then
            If IsEmpty(prevDate) And IsEmpty(thisDate) Then
            ElseIf IsEmpty(prevDate) Or IsEmpty(thisDate) Or (thisDate <> prevDate) Then
                wsDB.Range(wsDB.Cells(grpStart, 1), wsDB.Cells(r - 1, lastCol)).Interior.color = IIf(useSky, 옅은하늘색, RGB(255, 255, 255))
                useSky = Not useSky
                grpStart = r
            End If
        End If
        prevDate = thisDate
    Next r
    wsDB.Range(wsDB.Cells(grpStart, 1), wsDB.Cells(lastRow, lastCol)).Interior.color = IIf(useSky, 옅은하늘색, RGB(255, 255, 255))

    MsgBox "날짜별 그룹 음영(옅은 하늘색/흰색 교차)을 적용했습니다.", vbInformation
End Sub

' ============================================================
' 메인 프로시저: UserForm1으로 날짜 선택 후 역변환 실행
' ============================================================
Public Sub 일일생산DB_역변환_일일생산일보생성()
    Dim wsDB As Worksheet
    Dim ws템플릿 As Worksheet

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    Set ws템플릿 = ThisWorkbook.Worksheets("생산일보")
    On Error GoTo 0

    If wsDB Is Nothing Then
        MsgBox "'일일생산DB' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If
    If ws템플릿 Is Nothing Then
        MsgBox "'생산일보' 템플릿 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    UserForm1.Show vbModal
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "UserForm1을 찾을 수 없습니다. UserForm1_일일생산일보날짜선택.txt 안내대로 폼을 추가한 뒤 다시 실행하세요.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If UserForm1.Cancelled Or Not IsDate(UserForm1.SelectedDate) Then
        Unload UserForm1
        Exit Sub
    End If

    Call 일일생산DB_역변환_일일생산일보_실행(UserForm1.SelectedDate)
    Unload UserForm1
End Sub

' ============================================================
' 선택한 날짜로 일일생산일보 역변환 실행
' ============================================================
Public Sub 일일생산DB_역변환_일일생산일보_실행(선택날짜 As Variant)
    Dim wsDB As Worksheet
    Dim ws템플릿 As Worksheet
    Dim visDB As XlSheetVisibility, vis템플릿 As XlSheetVisibility
    Dim lastRow As Long
    Dim r As Long

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    Set ws템플릿 = ThisWorkbook.Worksheets("생산일보")
    On Error GoTo 0

    If wsDB Is Nothing Or ws템플릿 Is Nothing Then
        MsgBox "필요한 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    ' 기존 DB 열 보완 (col 125/126 헤더 없으면 자동 추가)


    visDB = wsDB.Visible
    vis템플릿 = ws템플릿.Visible

    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "일일생산DB에 데이터가 없습니다.", vbCritical
        Exit Sub
    End If

    Dim 데이터존재 As Boolean
    데이터존재 = False
    For r = 2 To lastRow
        Dim v1 As Variant: v1 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(v1) Then
            If v1 = 선택날짜 Then 데이터존재 = True: Exit For
            If IsDate(v1) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(v1), CDate(선택날짜)) = 0 Then 데이터존재 = True: Exit For
            End If
        End If
    Next r
    If Not 데이터존재 Then
        MsgBox "선택한 날짜(" & Format(선택날짜, "yyyy-mm-dd") & ")의 데이터가 일일생산DB에 없습니다.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wb보호해제 As Boolean
    wb보호해제 = False
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect "0205"
        wb보호해제 = True
    End If
    wsDB.Visible = xlSheetVisible
    ws템플릿.Visible = xlSheetVisible
    On Error GoTo 0

    On Error GoTo ErrCopy
    ws템플릿.Copy
    Dim wb신규 As Workbook
    Dim ws신규 As Worksheet
    Set wb신규 = ActiveWorkbook
    Application.DisplayAlerts = False
    Do While wb신규.Worksheets.Count > 1
        If InStr(1, wb신규.Worksheets(1).Name, "생산일보", vbTextCompare) > 0 Then
            wb신규.Worksheets(2).Delete
        Else
            wb신규.Worksheets(1).Delete
        End If
    Loop
    Application.DisplayAlerts = True
    Set ws신규 = wb신규.Worksheets(1)
    ws신규.Name = "생산일보"

    Call 시트가시성_안전복원(wsDB, visDB)
    Call 시트가시성_안전복원(ws템플릿, vis템플릿)
    If wb보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False

    Call 매크로버튼_삭제(ws신규)
    Call 생산일보_역변환용_초기화(ws신규)    ' 템플릿 복사본 → 완전 초기화 후 DB 입력
    Call DB데이터_일일생산일보_입력(wsDB, ws신규, 선택날짜)
    Call 일일생산일보_주월_계산후_시트반영(ws신규, wsDB, 선택날짜)   ' 먼저 DB 이력 기반 주/월 재계산
    Call PA데이터_T10_재입력(wsDB, ws신규, 선택날짜)                 ' DB 저장값으로 최종 덮어쓰기 (우선순위 최고)
    Call 투입인원_T11_재입력(wsDB, ws신규, 선택날짜)
    ' T10 자동수식(원지수율,LOSS율,목표대비,절감/초과,Conversion)은 PA·주월 데이터 반영 후 계산되어야 함
    Application.Calculation = xlCalculationAutomatic
    ws신규.Calculate
    DoEvents
    Call 수식을값으로변환(ws신규)

    Dim 파일명 As String
    Dim 저장경로 As String
    파일명 = "일일생산일지_" & Format(선택날짜, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".xlsb"
    저장경로 = ThisWorkbook.Path & "\" & 파일명

    On Error Resume Next
    wb신규.SaveAs Filename:=저장경로, FileFormat:=xlExcel12
    On Error GoTo 0

    On Error Resume Next
    Dim wsHome As Worksheet
    Set wsHome = ThisWorkbook.Sheets("Home menu")
    If Not wsHome Is Nothing Then wsHome.Activate
    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "일일생산일보가 생성되었습니다!" & vbCrLf & vbCrLf & "파일명: " & 파일명 & vbCrLf & "저장 위치: " & ThisWorkbook.Path, vbInformation
    GoTo End실행

ErrCopy:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    Call 시트가시성_안전복원(wsDB, visDB)
    Call 시트가시성_안전복원(ws템플릿, vis템플릿)
    If wb보호해제 Then ThisWorkbook.Protect Password:="0205", Structure:=True, Windows:=False
    On Error Resume Next
    Dim wsHomeErr As Worksheet
    Set wsHomeErr = ThisWorkbook.Sheets("Home menu")
    If Not wsHomeErr Is Nothing Then wsHomeErr.Activate
    On Error GoTo 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "일일생산일보 시트 복사 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류(" & errNum & "): " & errDesc, vbCritical
End실행:
End Sub

' ============================================================
' 생산일보 역변환용 초기화 (Module1 의존 없음 - 영역지우기 로직 인라인)
' ============================================================
Private Sub 생산일보_역변환용_초기화(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then On Error Resume Next: ws.Unprotect "0205": If ws.ProtectContents Then ws.Unprotect: On Error GoTo 0
    On Error Resume Next
    ws.Range("C7:BV16").ClearContents
    ws.Range("C20:BV29").ClearContents
    ws.Range("C35:BV44").ClearContents
    ws.Range("C48:BV57").ClearContents
    ws.Range("B68:T77").ClearContents
    ws.Range("AK68:BF77").ClearContents
    ws.Range("G108:BM108").ClearContents
    ws.Range("G109:Q109").ClearContents
    ws.Range("G111:BM111").ClearContents
    ws.Range("G112:BM112").ClearContents
    ws.Range("G116:K116").ClearContents
    ws.Range("G88:K91").ClearContents
    ws.Range("G93:K96").ClearContents
    ws.Range("G98:K101").ClearContents
    ws.Range("O88:O91").ClearContents
    ws.Range("O93:O96").ClearContents
    ws.Range("O98:O101").ClearContents
    ws.Range("AP88:AT91").ClearContents
    ws.Range("AP93:AT96").ClearContents
    ws.Range("AP98:AT101").ClearContents
    ws.Range("AX88:AX91").ClearContents
    ws.Range("AX93:AX96").ClearContents
    ws.Range("A122:AG126").ClearContents
    ws.Range("AL122:BV126").ClearContents
    ws.Range("BK4").ClearContents
    On Error GoTo 0
End Sub

' ============================================================
' 생산일보 시트에 역변환 데이터 직접 채우기 (Module24용)
' - 새 워크북 생성/파일 저장 없이 현재 생산일보 시트에 DB 데이터 입력
' - 날짜 선택은 호출자(UserForm1)에서 처리
' - M10_FillSheetFromDB: Application.Run용 영문별칭 (1004 오류 회피)
' ============================================================
Public Sub M10_FillSheetFromDB(선택날짜 As Variant)
    Call 일일생산DB_역변환_생산일보시트_채우기(선택날짜)
End Sub
Public Sub 일일생산DB_역변환_생산일보시트_채우기(선택날짜 As Variant)
    Dim wsDB As Worksheet
    Dim ws생산일보 As Worksheet
    Dim lastRow As Long
    Dim r As Long

    On Error Resume Next
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    Set ws생산일보 = ThisWorkbook.Worksheets("생산일보")
    On Error GoTo 0

    If wsDB Is Nothing Or ws생산일보 Is Nothing Then
        MsgBox "'일일생산DB' 또는 '생산일보' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "일일생산DB에 데이터가 없습니다.", vbCritical
        Exit Sub
    End If

    Dim 데이터존재 As Boolean
    Dim v1 As Variant
    데이터존재 = False
    For r = 2 To lastRow
        v1 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(v1) Then
            If v1 = 선택날짜 Then 데이터존재 = True: Exit For
            If IsDate(v1) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(v1), CDate(선택날짜)) = 0 Then 데이터존재 = True: Exit For
            End If
        End If
    Next r
    If Not 데이터존재 Then
        MsgBox "선택한 날짜(" & Format(선택날짜, "yyyy-mm-dd") & ")의 데이터가 일일생산DB에 없습니다.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrSheetFill
    ' 초기화: Module1 생산일보_완전초기화 로직 사용 (이전 데이터 가져오기 전 전체 초기화)
    Call 생산일보_완전초기화(ws생산일보)
    Call DB데이터_일일생산일보_입력(wsDB, ws생산일보, 선택날짜)
    Call 일일생산일보_주월_계산후_시트반영(ws생산일보, wsDB, 선택날짜)
    Call PA데이터_T10_재입력(wsDB, ws생산일보, 선택날짜)
    Call 투입인원_T11_재입력(wsDB, ws생산일보, 선택날짜)
    Application.Calculation = xlCalculationAutomatic
    ws생산일보.Calculate
    DoEvents
    ' 수식 유지: Module24 역변환 채우기는 초기화 영역에만 데이터 입력, 수식은 자동 계산 유지

    ws생산일보.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "생산일보 시트에 " & Format(선택날짜, "yyyy-mm-dd") & " 데이터를 채웠습니다.", vbInformation
    Exit Sub

ErrSheetFill:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "생산일보 시트 채우기 중 오류:" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================
' 역변환 후 수정된 생산일보 → 일일생산DB 위치기반 갱신 (기존 행 값 교체)
' - 새 행 추가/삭제 없이, DB 기존 행의 정확한 셀위치를 찾아 시트 변경값으로 교체
' - Module25에서 호출
' ============================================================
Public Sub 역변환_생산일보_변경값_DB반영()
    Dim ws생산일보 As Worksheet, wsDB As Worksheet
    Dim 작성일 As Variant, lastRow As Long, r As Long
    Dim DB날짜 As Variant, 마지막유효날짜 As Variant, 날짜일치 As Boolean
    Dim 생산일보구분 As String, 레코드유형 As String, 종별 As String, 순번 As Long

    On Error Resume Next
    Set ws생산일보 = ThisWorkbook.Worksheets("생산일보")
    Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
    On Error GoTo 0

    If ws생산일보 Is Nothing Or wsDB Is Nothing Then
        MsgBox "'생산일보' 또는 '일일생산DB' 시트를 찾을 수 없습니다.", vbCritical
        Exit Sub
    End If

    작성일 = ws생산일보.Range("BK4").Value
    If IsEmpty(작성일) Or Not IsDate(작성일) Then
        MsgBox "BK4에 유효한 날짜가 없습니다.", vbCritical
        Exit Sub
    End If

    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrDB반영

    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 작성일 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(작성일) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(작성일)) = 0 Then 날짜일치 = True
            End If
        End If
        If Not 날짜일치 Then GoTo NextDB행

        생산일보구분 = Trim(CStr(wsDB.Cells(r, 3).Value))
        레코드유형 = Trim(CStr(wsDB.Cells(r, 4).Value))
        순번 = 0
        On Error Resume Next: 순번 = CLng(wsDB.Cells(r, 5).Value): On Error GoTo ErrDB반영
        종별 = Trim(CStr(wsDB.Cells(r, 6).Value))

        If StrComp(레코드유형, "원지", vbTextCompare) = 0 Then
            Call 역변환_원지_시트값_DB반영(ws생산일보, wsDB, r, 생산일보구분, 종별, 순번)
        ElseIf StrComp(레코드유형, "완제품", vbTextCompare) = 0 Then
            Call 역변환_완제품_시트값_DB반영(ws생산일보, wsDB, r, 생산일보구분, 순번)
        ElseIf StrComp(레코드유형, "투입인원", vbTextCompare) = 0 Then
            Call 역변환_투입인원_시트값_DB반영(ws생산일보, wsDB, r, 생산일보구분, 순번)
        ElseIf StrComp(레코드유형, "ProductionAnalysis", vbTextCompare) = 0 Or StrComp(레코드유형, "PA", vbTextCompare) = 0 Then
            Call 역변환_PA_시트값_DB반영(ws생산일보, wsDB, r, 생산일보구분, 순번)
        End If
NextDB행:
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "역변환으로 불러온 " & Format(작성일, "yyyy-mm-dd") & " 데이터의" & vbCrLf & _
        "변경값이 일일생산DB 기존 행에 반영되었습니다.", vbInformation
    Exit Sub

ErrDB반영:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DB 반영 중 오류:" & vbCrLf & Err.Description, vbCritical
End Sub

' 원지: 시트 셀값 → DB 해당 행에 직접 쓰기 (위치 기반)
Private Sub 역변환_원지_시트값_DB반영(ws As Worksheet, wsDB As Worksheet, DB행 As Long, 생산일보구분 As String, 종별 As String, 순번 As Long)
    Dim 시작행 As Long, 끝행 As Long, ROLLID열 As Long, KG열 As Long, Co열 As Long, 대상행 As Long
    If Not 원지_매핑_구하기(생산일보구분, 종별, 시작행, 끝행, ROLLID열, KG열) Then Exit Sub
    Co열 = ROLLID열 - 2
    대상행 = 시작행 + 순번 - 1
    If 대상행 < 시작행 Or 대상행 > 끝행 Then Exit Sub

    Dim vCo As Variant, vROLL As Variant, vKG As Variant
    vCo = 셀값_병합처리(ws, 대상행, Co열)
    vROLL = 셀값_병합처리(ws, 대상행, ROLLID열)
    vKG = 셀값_병합처리(ws, 대상행, KG열)

    wsDB.Cells(DB행, 7).Value = vROLL
    wsDB.Cells(DB행, 8).Value = vKG
    Dim colCo As Long: colCo = M18_원지Co열(wsDB)
    If colCo > 0 And colCo <= 126 Then wsDB.Cells(DB행, colCo).Value = vCo
End Sub

' 완제품: 시트 셀값 → DB 해당 행에 직접 쓰기
Private Sub 역변환_완제품_시트값_DB반영(ws As Worksheet, wsDB As Worksheet, DB행 As Long, 생산일보구분 As String, 순번 As Long)
    If 순번 < 1 Or 순번 > 10 Then Exit Sub
    Dim 대상행 As Long: 대상행 = 67 + 순번

    Dim v품명 As Variant, v설비 As Variant, v생산량 As Variant
    If InStr(1, 생산일보구분, "야간", vbTextCompare) > 0 Then
        v품명 = 셀값_병합처리(ws, 대상행, 37): v설비 = 셀값_병합처리(ws, 대상행, 52): v생산량 = 셀값_병합처리(ws, 대상행, 55)
    Else
        v품명 = 셀값_병합처리(ws, 대상행, 2): v설비 = 셀값_병합처리(ws, 대상행, 17): v생산량 = 셀값_병합처리(ws, 대상행, 20)
    End If
    wsDB.Cells(DB행, 10).Value = v품명
    wsDB.Cells(DB행, 11).Value = v설비
    wsDB.Cells(DB행, 12).Value = v생산량
End Sub

' 투입인원: 시트 셀값 → DB 해당 행에 직접 쓰기 (간략 처리)
Private Sub 역변환_투입인원_시트값_DB반영(ws As Worksheet, wsDB As Worksheet, DB행 As Long, 생산일보구분 As String, 순번 As Long)
    Dim 설비명 As String, 구분 As String, base열 As Long
    설비명 = UCase(Trim(CStr(wsDB.Cells(DB행, 19).Value)))
    구분 = Trim(CStr(wsDB.Cells(DB행, 20).Value))

    Select Case 설비명
        Case "KT":  base열 = 7
        Case "FT":  base열 = 43
        Case "JRT": base열 = 19
        Case "BT":  base열 = 31
        Case "PRE": base열 = 55
        Case Else: Exit Sub
    End Select

    Dim v As Variant
    If 구분 = "작업시간" Then
        v = 셀값_병합처리(ws, 108, base열): If IsNumeric(v) Then wsDB.Cells(DB행, 21).Value = CDbl(v)
        v = 셀값_병합처리(ws, 109, base열): If IsNumeric(v) Then wsDB.Cells(DB행, 22).Value = CDbl(v)
    ElseIf 구분 = "연장근로" Then
        v = 셀값_병합처리(ws, 108, base열 + 2): If IsNumeric(v) Then wsDB.Cells(DB행, 21).Value = CDbl(v)
        v = 셀값_병합처리(ws, 109, base열 + 2): If IsNumeric(v) Then wsDB.Cells(DB행, 22).Value = CDbl(v)
    ElseIf 구분 = "인력Direct" Or 구분 = "인력Indirect" Then
        v = 셀값_병합처리(ws, 108, base열): If IsNumeric(v) Then wsDB.Cells(DB행, 21).Value = CDbl(v)
        v = 셀값_병합처리(ws, 108, base열 + 2): If IsNumeric(v) Then wsDB.Cells(DB행, 22).Value = CDbl(v)
        v = 셀값_병합처리(ws, 108, base열 + 4): If IsNumeric(v) Then wsDB.Cells(DB행, 23).Value = CDbl(v)
        v = 셀값_병합처리(ws, 108, base열 + 6): If IsNumeric(v) Then wsDB.Cells(DB행, 24).Value = CDbl(v)
        v = 셀값_병합처리(ws, 108, base열 + 8): If IsNumeric(v) Then wsDB.Cells(DB행, 25).Value = CDbl(v)
    End If
End Sub

' PA: 시트 셀값 → DB 해당 행에 직접 쓰기 (PA2 비가동 제외, PA1 T10 일계만 간략 반영)
Private Sub 역변환_PA_시트값_DB반영(ws As Worksheet, wsDB As Worksheet, DB행 As Long, 생산일보구분 As String, 순번 As Long)
    Dim PA순번 As Long
    PA순번 = 0
    On Error Resume Next: PA순번 = CLng(wsDB.Cells(DB행, 5).Value): On Error GoTo 0
    If PA순번 = 2 Then Exit Sub
    ' PA1: KT/FT/JRT/BT/PRE 일계_원투·양품 (DB 28,29 / 40,41 / 43,44 / 46,47 / 49,50)
    Dim DB원열 As Variant, DB양열 As Variant, 시트행 As Variant, 시트base As Variant
    DB원열 = Array(28, 40, 43, 46, 49)
    DB양열 = Array(29, 41, 44, 47, 50)
    시트행 = Array(87, 87, 92, 97, 92)
    시트base = Array(7, 46, 11, 11, 46)
    Dim idx As Long, v As Variant
    For idx = 0 To 4
        v = 셀값_병합처리(ws, CLng(시트행(idx)), CLng(시트base(idx)))
        If IsNumeric(v) Then wsDB.Cells(DB행, CLng(DB원열(idx))).Value = CDbl(v)
        v = 셀값_병합처리(ws, CLng(시트행(idx)), CLng(시트base(idx)) + 2)
        If IsNumeric(v) Then wsDB.Cells(DB행, CLng(DB양열(idx))).Value = CDbl(v)
    Next idx
End Sub

' 병합셀 포함 안전 읽기
Private Function 셀값_병합처리(ws As Worksheet, r As Long, c As Long) As Variant
    Dim rng As Range
    Set rng = ws.Cells(r, c)
    On Error Resume Next
    If rng.MergeCells And rng.MergeArea.Rows.Count = 1 Then Set rng = rng.MergeArea.Cells(1, 1)
    On Error GoTo 0
    셀값_병합처리 = rng.Value
End Function

' ============================================================
' DB 데이터를 일일생산일보 셀주소에 따라 입력 (새 매핑)
' DB 생산일보구분: "생산일보_주간"/"생산일보_야간"/"생산일보"
' ============================================================
Private Sub DB데이터_일일생산일보_입력(wsDB As Worksheet, ws신규 As Worksheet, 선택날짜 As Variant)
    Dim lastRow As Long, r As Long, 종Tmp As String
    lastRow = wsDB.Cells(wsDB.Rows.Count, 4).End(xlUp).Row

    ' BK4 = 작성일 (날짜 형식 차이 대응: DateDiff 기반 비교)
    For r = 2 To lastRow
        Dim vd As Variant: vd = wsDB.Cells(r, 1).Value
        If Not IsEmpty(vd) Then
            If vd = 선택날짜 Then ws신규.Range("BK4").Value = vd: Exit For
            If IsDate(vd) And IsDate(선택날짜) And DateDiff("d", CDate(vd), CDate(선택날짜)) = 0 Then ws신규.Range("BK4").Value = vd: Exit For
        End If
    Next r

    Dim 완주간순번 As Long, 완야간순번 As Long
    완주간순번 = 0
    완야간순번 = 0

    ' (생산일보_완전초기화가 호출자에서 이미 실행됨 - 원지 포함 전체 초기화 완료)

    Dim DB날짜 As Variant, 날짜일치 As Boolean, 마지막유효날짜 As Variant
    Dim 생산일보구분 As String, 레코드유형 As String
    마지막유효날짜 = Empty

    ' ★ 원지 누락방지: KT계(rows 7/35)와 JRT계(rows 20/48)가 row 20~26에서 겹침.
    '   JRT가 나중에 처리되면 KT를 덮어써 누락됨 → 1차 JRT계, 2차 KT계 순으로 처리
    ' ★ 날짜 빈 행 상속: DB col1(생산일) 빈 행은 이전 행 날짜 적용 (그룹 연속성)
    ' ★ 원지 순차입력: 블록별 다음행 딕셔너리 초기화 (첫행부터 연속 입력, 중간 빈행 없음)
    Dim 원지_다음행딕 As Object: Set 원지_다음행딕 = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If Not 날짜일치 Then GoTo NextRow1
        생산일보구분 = Trim(CStr(wsDB.Cells(r, 3).Value))
        레코드유형 = Trim(CStr(wsDB.Cells(r, 4).Value))
        If StrComp(레코드유형, "원지", vbTextCompare) = 0 Then
            종Tmp = 원지_종별_정규화(Trim(CStr(wsDB.Cells(r, 6).Value)))
            ' JRT계(JRT/BT/PRE) 먼저 처리
            If 종Tmp = "JRT-1" Or 종Tmp = "JRT-2" Or 종Tmp = "BT-1" Or 종Tmp = "BT-2" Or 종Tmp = "BT-3" Or 종Tmp = "PRE" Then
                Call 원지데이터_입력(wsDB, ws신규, r, 생산일보구분, 원지_다음행딕)
            End If
        End If
NextRow1:
    Next r

    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If Not 날짜일치 Then GoTo NextRow2
        생산일보구분 = Trim(CStr(wsDB.Cells(r, 3).Value))
        레코드유형 = Trim(CStr(wsDB.Cells(r, 4).Value))
        If StrComp(레코드유형, "원지", vbTextCompare) = 0 Then
            종Tmp = 원지_종별_정규화(Trim(CStr(wsDB.Cells(r, 6).Value)))
            ' KT계(KT-1~5, FT) 나중 처리 → 겹치는 row 20~26에 KT가 덮어써짐
            If 종Tmp = "KT-1" Or 종Tmp = "KT-2" Or 종Tmp = "KT-3" Or 종Tmp = "KT-4" Or 종Tmp = "KT-5" Or 종Tmp = "FT" Then
                Call 원지데이터_입력(wsDB, ws신규, r, 생산일보구분, 원지_다음행딕)
            End If
        End If
NextRow2:
    Next r

    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If Not 날짜일치 Then GoTo NextRow

        생산일보구분 = Trim(CStr(wsDB.Cells(r, 3).Value))
        레코드유형 = Trim(CStr(wsDB.Cells(r, 4).Value))

        If StrComp(레코드유형, "원지", vbTextCompare) = 0 Then
            ' 이미 위 2패스에서 처리됨
        ElseIf StrComp(레코드유형, "완제품", vbTextCompare) = 0 Then
                ' 주야 구분: "주간" 포함→좌측(T07), "야간" 포함→우측(T08). 미지정("생산일보" 등)→주간
                If InStr(1, 생산일보구분, "야간", vbTextCompare) > 0 Then
                    완야간순번 = 완야간순번 + 1
                    Call 완제품_야간_입력(wsDB, ws신규, r, 완야간순번)
                Else
                    완주간순번 = 완주간순번 + 1
                    Call 완제품_주간_입력(wsDB, ws신규, r, 완주간순번)
                End If
        ElseIf StrComp(레코드유형, "투입인원", vbTextCompare) = 0 Then
                Call 투입인원_입력(wsDB, ws신규, r)
        ElseIf StrComp(레코드유형, "ProductionAnalysis", vbTextCompare) = 0 Or StrComp(레코드유형, "PA", vbTextCompare) = 0 Then
                Dim PA순번 As Long
                PA순번 = 0
                On Error Resume Next
                PA순번 = CLng(wsDB.Cells(r, 5).Value)
                On Error GoTo 0
                If PA순번 = 2 Then
                    Call 비가동현황_입력(wsDB, ws신규, r)
                Else
                    Call PA데이터_입력(wsDB, ws신규, r)
                End If
        End If
NextRow:
    Next r

    ' ★ 3패스: 오버플로우(seq > 블록끝행) 처리 - 정규 기록 완료 후 빈 셀 스캔 배치
    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If Not 날짜일치 Then GoTo NextOverflow
        생산일보구분 = Trim(CStr(wsDB.Cells(r, 3).Value))
        레코드유형 = Trim(CStr(wsDB.Cells(r, 4).Value))
        If StrComp(레코드유형, "원지", vbTextCompare) = 0 Then
            Call 원지_오버플로우_입력(wsDB, ws신규, r, 생산일보구분)
        End If
NextOverflow:
    Next r
End Sub

' 원지 영역 초기화 (ROLLID·KG 셀) - 역변환 전 템플릿/기존데이터 제거
Private Sub 원지영역_초기화(ws As Worksheet)
    Dim base열 As Variant, r As Long, c As Long, ROLLID열 As Long, KG열 As Long
    base열 = Array(3, 15, 27, 39, 51, 63)
    Dim 블록시작 As Variant, 블록끝 As Variant
    블록시작 = Array(7, 20, 35, 48)
    블록끝 = Array(16, 29, 44, 57)
    Dim b As Long
    For b = 0 To 3
        For r = 블록시작(b) To 블록끝(b)
            For c = 0 To 5
                ROLLID열 = base열(c) + 2
                KG열 = base열(c) + 9
                On Error Resume Next
                ws.Cells(r, ROLLID열).ClearContents
                ws.Cells(r, KG열).ClearContents
                On Error GoTo 0
            Next c
        Next r
    Next b
End Sub

' ============================================================
' 원지 데이터 입력 (종별 + 생산일보구분 → 시작행/열 매핑)
' ★ 순차입력: 첫행부터 연속 입력, 중간 빈행 없음 (다음행딕으로 블록별 카운터 유지)
' seq > 끝행인 오버플로우는 3패스(원지_오버플로우_입력)에서 처리
' ============================================================
Private Sub 원지데이터_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 생산일보구분 As String, ByRef 다음행딕 As Object)
    Dim 종별 As String
    종별 = 원지_종별_정규화(Trim(CStr(wsDB.Cells(DB행, 6).Value)))
    Dim ROLLID As String, KG As Variant, vCo As String, colCo As Long
    ROLLID = Trim(CStr(wsDB.Cells(DB행, 7).Value))
    KG = wsDB.Cells(DB행, 8).Value
    If ROLLID = "" And IsEmpty(KG) Then Exit Sub

    colCo = M18_원지Co열(wsDB)
    vCo = Trim(CStr(wsDB.Cells(DB행, colCo).Value))

    Dim 시작행 As Long, 끝행 As Long, ROLLID열 As Long, KG열 As Long, Co열 As Long
    If Not 원지_매핑_구하기(생산일보구분, 종별, 시작행, 끝행, ROLLID열, KG열) Then Exit Sub
    Co열 = ROLLID열 - 2   ' 원지사용내역: Co열(id왼쪽)=base, id=base+2, KG=base+9

    Dim 맵키 As String: 맵키 = 생산일보구분 & "|" & 종별
    If Not 다음행딕.exists(맵키) Then 다음행딕.Add 맵키, 0
    Dim nextOffset As Long: nextOffset = CLng(다음행딕(맵키))

    Dim 대상행 As Long
    대상행 = 시작행 + nextOffset
    If 대상행 > 끝행 Then Exit Sub  ' 오버플로우: 3패스에서 별도 처리 (카운터 증가 안 함)

    다음행딕(맵키) = nextOffset + 1

    ' ★ 역변환 시 셀병합 절대 해제하지 않음. 병합된 경우 가로병합(같은 행)만 MergeArea 첫 셀에 기록
    Dim rngCo As Range, rngROLL As Range, rngKG As Range
    Set rngCo = ws신규.Cells(대상행, Co열)
    Set rngROLL = ws신규.Cells(대상행, ROLLID열)
    Set rngKG = ws신규.Cells(대상행, KG열)
    On Error Resume Next
    If rngCo.MergeCells And rngCo.MergeArea.Rows.Count = 1 Then Set rngCo = rngCo.MergeArea.Cells(1, 1)
    If rngROLL.MergeCells And rngROLL.MergeArea.Rows.Count = 1 Then Set rngROLL = rngROLL.MergeArea.Cells(1, 1)
    If rngKG.MergeCells And rngKG.MergeArea.Rows.Count = 1 Then Set rngKG = rngKG.MergeArea.Cells(1, 1)
    On Error GoTo 0

    If vCo <> "" Then rngCo.Value = vCo
    If ROLLID <> "" Then rngROLL.Value = wsDB.Cells(DB행, 7).Value
    If Not IsEmpty(KG) Then rngKG.Value = KG
End Sub

' ============================================================
' 원지 오버플로우 입력 (3패스에서 호출)
' 정규 기록이 모두 끝난 후 ws신규 빈 셀을 스캔해서 배치
' KT-1→KT-2~5, JRT-1→JRT-2,BT-1~3 등 설비군 내 이관
' ============================================================
Private Sub 원지_오버플로우_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 생산일보구분 As String)
    Dim 종별 As String, 순번 As Long
    종별 = 원지_종별_정규화(Trim(CStr(wsDB.Cells(DB행, 6).Value)))
    순번 = 0
    On Error Resume Next
    순번 = CLng(wsDB.Cells(DB행, 5).Value)
    On Error GoTo 0
    If 순번 < 1 Or 순번 > 20 Then Exit Sub

    Dim 시작행 As Long, 끝행 As Long, ROLLID열 As Long, KG열 As Long
    If Not 원지_매핑_구하기(생산일보구분, 종별, 시작행, 끝행, ROLLID열, KG열) Then Exit Sub

    Dim 대상행 As Long
    대상행 = 시작행 + 순번 - 1
    If 대상행 <= 끝행 Then Exit Sub  ' 정규 레코드, 오버플로우 아님

    ' 형제 컬럼 목록 (이관 대상)
    Dim siblingBases As Variant
    If left(종별, 2) = "KT" Then
        Select Case 종별
            Case "KT-1": siblingBases = Array(15, 27, 39, 51)
            Case "KT-2": siblingBases = Array(27, 39, 51)
            Case "KT-3": siblingBases = Array(39, 51)
            Case "KT-4": siblingBases = Array(51)
            Case Else: Exit Sub
        End Select
    Else
        Select Case 종별
            Case "JRT-1": siblingBases = Array(15, 27, 39, 51)
            Case "JRT-2": siblingBases = Array(27, 39, 51)
            Case "BT-1": siblingBases = Array(39, 51)
            Case "BT-2": siblingBases = Array(51)
            Case Else: Exit Sub
        End Select
    End If

    ' ws신규 빈 셀 스캔: 정규 기록 완료 후이므로 정확한 빈 위치 파악 가능
    Dim baseCol As Long, r As Long
    Dim targetRow As Long, targetCo As Long, targetROLL As Long, targetKG As Long
    Dim found As Boolean
    found = False
    Dim i As Long
    For i = LBound(siblingBases) To UBound(siblingBases)
        baseCol = siblingBases(i)
        For r = 시작행 To 끝행
            Dim rngTest As Range
            Set rngTest = ws신규.Cells(r, baseCol + 2)
            On Error Resume Next
            If rngTest.MergeCells And rngTest.MergeArea.Rows.Count = 1 Then Set rngTest = rngTest.MergeArea.Cells(1, 1)
            On Error GoTo 0
            If IsEmpty(rngTest.Value) Then
                targetRow = r
                targetCo = baseCol
                targetROLL = baseCol + 2
                targetKG = baseCol + 9
                found = True
                Exit For
            End If
        Next r
        If found Then Exit For
    Next i
    If Not found Then Exit Sub

    Dim rngCo As Range, rngROLL As Range, rngKG As Range
    Set rngCo = ws신규.Cells(targetRow, targetCo)
    Set rngROLL = ws신규.Cells(targetRow, targetROLL)
    Set rngKG = ws신규.Cells(targetRow, targetKG)
    On Error Resume Next
    If rngCo.MergeCells And rngCo.MergeArea.Rows.Count = 1 Then Set rngCo = rngCo.MergeArea.Cells(1, 1)
    If rngROLL.MergeCells And rngROLL.MergeArea.Rows.Count = 1 Then Set rngROLL = rngROLL.MergeArea.Cells(1, 1)
    If rngKG.MergeCells And rngKG.MergeArea.Rows.Count = 1 Then Set rngKG = rngKG.MergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim ROLLID As String, KG As Variant, vCo As String, colCo As Long
    ROLLID = Trim(CStr(wsDB.Cells(DB행, 7).Value))
    KG = wsDB.Cells(DB행, 8).Value
    colCo = M18_원지Co열(wsDB)
    vCo = Trim(CStr(wsDB.Cells(DB행, colCo).Value))
    If vCo <> "" Then rngCo.Value = vCo
    If ROLLID <> "" Then rngROLL.Value = wsDB.Cells(DB행, 7).Value
    If Not IsEmpty(KG) Then rngKG.Value = KG

    Dim memo As String
    memo = 종별 & " 원본 seq" & 순번 & ", 중량 " & IIf(IsEmpty(KG), "-", CStr(KG)) & " kg (분산배치)"
    On Error Resume Next
    rngCo.ClearComments: rngCo.AddComment: rngCo.Comment.Text memo: rngCo.Comment.Visible = False
    rngROLL.ClearComments: rngROLL.AddComment: rngROLL.Comment.Text memo: rngROLL.Comment.Visible = False
    rngKG.ClearComments: rngKG.AddComment: rngKG.Comment.Text memo: rngKG.Comment.Visible = False
    On Error GoTo 0
    rngCo.Interior.color = RGB(255, 182, 193)
    rngROLL.Interior.color = RGB(255, 182, 193)
    rngKG.Interior.color = RGB(255, 182, 193)
End Sub

' 원지 종별 + 생산일보구분 → (시작행, 끝행, ROLLID열, KG열) 반환
' base열 = 3,15,27,39,51,63 → ROLLID = base+2, KG = base+9
' ★ 블록별 행범위: KT 7-16/35-44, JRT 20-29/48-57. 순번 11+가 다른 블록 덮어쓰지 않도록 끝행 검사
Private Function 원지_매핑_구하기(생산일보구분 As String, 종별 As String, _
    ByRef 시작행 As Long, ByRef 끝행 As Long, ByRef ROLLID열 As Long, ByRef KG열 As Long) As Boolean

    원지_매핑_구하기 = False
    Dim base열 As Long
    base열 = -1

    ' 종별 정규화 (2월 마이그레이션·기존 데이터 변형 대응)
    종별 = 원지_종별_정규화(종별)

    Select Case 종별
        Case "KT-1", "JRT-1": base열 = 3
        Case "KT-2", "JRT-2": base열 = 15
        Case "KT-3", "BT-1":  base열 = 27
        Case "KT-4", "BT-2":  base열 = 39
        Case "KT-5", "BT-3":  base열 = 51
        Case "FT", "PRE":     base열 = 63
        Case Else: Exit Function
    End Select

    ROLLID열 = base열 + 2
    KG열 = base열 + 9

    Dim isKT계 As Boolean
    isKT계 = (left(종별, 2) = "KT" Or 종별 = "FT")

    ' 주간/야간 구분 + 블록별 끝행 (순번 11~20이 JRT 블록 덮어쓰는 것 방지)
    If InStr(1, 생산일보구분, "주간", vbTextCompare) > 0 Then
        If isKT계 Then 시작행 = 7: 끝행 = 16 Else 시작행 = 20: 끝행 = 29
    ElseIf InStr(1, 생산일보구분, "야간", vbTextCompare) > 0 Then
        If isKT계 Then 시작행 = 35: 끝행 = 44 Else 시작행 = 48: 끝행 = 57
    ElseIf InStr(1, 생산일보구분, "생산일보", vbTextCompare) > 0 Then
        If isKT계 Then 시작행 = 7: 끝행 = 16 Else 시작행 = 20: 끝행 = 29
    Else
        Exit Function
    End If

    원지_매핑_구하기 = True
End Function

' 원지 종별 정규화: KT①→KT-1, 지관→PRE, KT1→KT-1 등
Private Function 원지_종별_정규화(ByVal s As String) As String
    s = Trim(s): If s = "" Then 원지_종별_정규화 = s: Exit Function
    s = Replace(s, "KT-①", "KT-1"): s = Replace(s, "KT-②", "KT-2")
    s = Replace(s, "KT-③", "KT-3"): s = Replace(s, "KT-④", "KT-4")
    s = Replace(s, "BT-⑤", "KT-5"):  s = Replace(s, "KT-⑤", "KT-5")
    s = Replace(s, "KT-⑥", "FT"):    s = Replace(s, "지관", "PRE")
    ' 숫자만 붙은 변형 (KT1→KT-1)
    If Len(s) >= 3 And UCase(left(s, 2)) = "KT" And Mid(s, 3, 1) <> "-" Then
        If IsNumeric(Mid(s, 3, 1)) Then s = "KT-" & Mid(s, 3, 1)
    End If
    If Len(s) >= 4 And UCase(left(s, 3)) = "JRT" And Mid(s, 4, 1) <> "-" Then
        If IsNumeric(Mid(s, 4, 1)) Then s = "JRT-" & Mid(s, 4, 1)
    End If
    If Len(s) >= 3 And UCase(left(s, 2)) = "BT" And Mid(s, 3, 1) <> "-" Then
        If IsNumeric(Mid(s, 3, 1)) Then s = "BT-" & Mid(s, 3, 1)
    End If
    원지_종별_정규화 = s
End Function

' ============================================================
' 완제품 주간 입력: T07 rows 68-77
' B(2)=품명, Q(17)=설비호기, T(20)=생산량BOX 만 값입력
' X68:AF78 = 수식구간(원지기준소요량~AJ13) → 덮어쓰기 금지, 자동 계산
' ============================================================
Private Sub 완제품_주간_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 순번 As Long)
    If 순번 < 1 Or 순번 > 10 Then Exit Sub
    Dim 대상행 As Long
    대상행 = 67 + 순번

    ws신규.Cells(대상행, 2).Value = wsDB.Cells(DB행, 10).Value   ' 품명
    ws신규.Cells(대상행, 17).Value = wsDB.Cells(DB행, 11).Value  ' 설비호기
    ws신규.Cells(대상행, 20).Value = wsDB.Cells(DB행, 12).Value  ' 생산량BOX
End Sub

' ============================================================
' 완제품 야간 입력: T08 rows 68-77 (우측)
' AK(37)=품명, AZ(52)=설비호기, BC(55)=생산량BOX 만 값입력
' BG68:BO78 = 수식구간(원지기준소요량~AJ13) → 덮어쓰기 금지, 자동 계산
' ============================================================
Private Sub 완제품_야간_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 순번 As Long)
    If 순번 < 1 Or 순번 > 10 Then Exit Sub
    Dim 대상행 As Long
    대상행 = 67 + 순번

    ws신규.Cells(대상행, 37).Value = wsDB.Cells(DB행, 10).Value   ' 품명
    ws신규.Cells(대상행, 52).Value = wsDB.Cells(DB행, 11).Value   ' 설비호기
    ws신규.Cells(대상행, 55).Value = wsDB.Cells(DB행, 12).Value   ' 생산량BOX
End Sub

' ============================================================
' 투입인원 T11 재입력 (수식변환 후 호출 - G108/G109 덮어쓰기 방지)
' ============================================================
Private Sub PA데이터_T10_재입력(wsDB As Worksheet, ws신규 As Worksheet, 선택날짜 As Variant)
    Dim lastRow As Long, r As Long, DB날짜 As Variant, 날짜일치 As Boolean, PA순번 As Long, 마지막유효날짜 As Variant
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If 날짜일치 Then
            If StrComp(Trim(CStr(wsDB.Cells(r, 4).Value)), "ProductionAnalysis", vbTextCompare) = 0 Or StrComp(Trim(CStr(wsDB.Cells(r, 4).Value)), "PA", vbTextCompare) = 0 Then
                PA순번 = 0
                On Error Resume Next
                PA순번 = CLng(wsDB.Cells(r, 5).Value)
                On Error GoTo 0
                If PA순번 <> 2 Then Call PA데이터_입력(wsDB, ws신규, r)
            End If
        End If
    Next r
End Sub

Private Sub 투입인원_T11_재입력(wsDB As Worksheet, ws신규 As Worksheet, 선택날짜 As Variant)
    Dim lastRow As Long, r As Long, DB날짜 As Variant, 마지막유효날짜 As Variant, 날짜일치 As Boolean
    lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    마지막유효날짜 = Empty
    For r = 2 To lastRow
        DB날짜 = wsDB.Cells(r, 1).Value
        If Not IsEmpty(DB날짜) Then 마지막유효날짜 = DB날짜
        날짜일치 = False
        If Not IsEmpty(마지막유효날짜) Then
            If 마지막유효날짜 = 선택날짜 Then
                날짜일치 = True
            ElseIf IsDate(마지막유효날짜) And IsDate(선택날짜) Then
                If DateDiff("d", CDate(마지막유효날짜), CDate(선택날짜)) = 0 Then 날짜일치 = True
            End If
        End If
        If 날짜일치 And StrComp(Trim(CStr(wsDB.Cells(r, 4).Value)), "투입인원", vbTextCompare) = 0 Then Call 투입인원_입력(wsDB, ws신규, r)
    Next r
End Sub

' ============================================================
' 투입인원 입력
' 02시트는 주간(정규/연장)만 있음 → base(정규), base+2(연장)만 입력
' KT: F40,F41=인력정규/연장 J40,J41=작업정규/연장 → G108,I108, G109,I109
' JRT: G68,G69=인력 I68,I69=작업 → S108,U108, S109,U109
' DB: 작업시간=정규레코드, 연장근로=연장레코드
' ============================================================
Private Sub 투입인원_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long)
    Dim 설비명 As String, 구분 As String
    설비명 = UCase(Trim(CStr(wsDB.Cells(DB행, 19).Value)))
    구분 = Trim(CStr(wsDB.Cells(DB행, 20).Value))

    Dim base열 As Long
    Select Case 설비명
        Case "KT":  base열 = 7
        Case "FT":  base열 = 43
        Case "JRT": base열 = 19
        Case "BT":  base열 = 31
        Case "PRE": base열 = 55
        Case Else: Exit Sub
    End Select

    Dim v As Variant
    If 구분 = "작업시간" Then
        ' 주간정규: F21→108 base(인력), F22→109 base(작업시간)
        v = wsDB.Cells(DB행, 21).Value: If IsNumeric(v) Then ws신규.Cells(108, base열).Value = CDbl(v)
        v = wsDB.Cells(DB행, 22).Value
        If Not IsNumeric(v) Then v = wsDB.Cells(DB행, 21).Value
        If IsNumeric(v) Then ws신규.Cells(109, base열).Value = CDbl(v)
        ' B67:Z76 (JRT/FT/BT/PRE) 정규행
        If 설비명 <> "KT" Then
            Dim 행정규 As Long
            Select Case 설비명
                Case "JRT": 행정규 = 68
                Case "FT":  행정규 = 72
                Case "BT":  행정규 = 74
                Case "PRE": 행정규 = 76
                Case Else: Exit Sub
            End Select
            v = wsDB.Cells(DB행, 22).Value
            If Not IsNumeric(v) Then v = wsDB.Cells(DB행, 21).Value
            If IsNumeric(v) Then ws신규.Cells(행정규, 9).Value = CDbl(v)
        End If
    ElseIf 구분 = "연장근로" Then
        ' 주간연장: F21→108 base+2(인력), F22→109 base+2(작업시간)
        v = wsDB.Cells(DB행, 21).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 2).Value = CDbl(v)
        v = wsDB.Cells(DB행, 22).Value
        If IsNumeric(v) Then ws신규.Cells(109, base열 + 2).Value = CDbl(v)
        ' B67:Z76 연장행
        Select Case 설비명
            Case "JRT": v = wsDB.Cells(DB행, 22).Value: If IsNumeric(v) Then ws신규.Cells(69, 9).Value = CDbl(v)
            Case "FT":  v = wsDB.Cells(DB행, 22).Value: If IsNumeric(v) Then ws신규.Cells(73, 9).Value = CDbl(v)
            Case "BT":  v = wsDB.Cells(DB행, 22).Value: If IsNumeric(v) Then ws신규.Cells(75, 9).Value = CDbl(v)
        End Select
    ElseIf 구분 = "인력Direct" Or 구분 = "인력Indirect" Then
        v = wsDB.Cells(DB행, 21).Value: If IsNumeric(v) Then ws신규.Cells(108, base열).Value = CDbl(v)
        v = wsDB.Cells(DB행, 22).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 2).Value = CDbl(v)
        v = wsDB.Cells(DB행, 23).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 4).Value = CDbl(v)
        v = wsDB.Cells(DB행, 24).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 6).Value = CDbl(v)
        v = wsDB.Cells(DB행, 25).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 8).Value = CDbl(v)
    End If
End Sub

' Overflow(6) 방지: 비수치/초대형값/에러값 시 0 반환
Private Function PA_SafeD(v As Variant) As Double
    PA_SafeD = 0
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    On Error Resume Next
    Dim d As Double: d = CDbl(v)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    If d > 1E+15 Or d < -1E+15 Then Exit Function
    PA_SafeD = d: Err.Clear
End Function

' ============================================================
' PA(원자재수율) 데이터 입력: T10 rows 84-101
' DB F26~F124 (최적화DB v3 124열) → 시트 셀 매핑
'   좌측: G(7)=원투, L(12)=양품, K(11)=기준소요량, O(15)=원지수율, Y(25)=LOSS율
'   우측: AP(42)=원투, AU(47)=양품, AT(46)=기준소요량, AX(50)=원지수율, BH(60)=LOSS율
'   KT:87-91 / JRT:92-96 / BT:97-101 / FT:87-91(right) / PRE:92-96(right)
' ============================================================
Private Sub PA데이터_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long)
    If wsDB.Cells(DB행, wsDB.Columns.Count).End(xlToLeft).Column < 26 Then Exit Sub

    Dim v원 As Double, v양 As Double, vLOSS As Double

    ' KT 일계 (row 87) = 수식구간(G87,O87) → 덮어쓰기 금지, 자동 계산
    ' DB 28~33은 row 88~91 입력 후 수식이 자동 반영

    ' KT 주간누계/월간누계 (row 88,89) → DB 34~39, 기소 103/104  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 34).Value) <> 0 Then ws신규.Cells(88, 7).Value = wsDB.Cells(DB행, 34).Value
    If PA_SafeD(wsDB.Cells(DB행, 35).Value) <> 0 Then ws신규.Cells(88, 12).Value = wsDB.Cells(DB행, 35).Value
    If PA_SafeD(wsDB.Cells(DB행, 36).Value) <> 0 Then ws신규.Cells(88, 15).Value = wsDB.Cells(DB행, 36).Value
    If PA_SafeD(wsDB.Cells(DB행, 103).Value) <> 0 Then
        ws신규.Cells(88, 11).Value = wsDB.Cells(DB행, 103).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(88, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 34).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 35).Value)
    If v원 <> 0 Then ws신규.Cells(88, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    If PA_SafeD(wsDB.Cells(DB행, 37).Value) <> 0 Then ws신규.Cells(89, 7).Value = wsDB.Cells(DB행, 37).Value
    If PA_SafeD(wsDB.Cells(DB행, 38).Value) <> 0 Then ws신규.Cells(89, 12).Value = wsDB.Cells(DB행, 38).Value
    If PA_SafeD(wsDB.Cells(DB행, 39).Value) <> 0 Then ws신규.Cells(89, 15).Value = wsDB.Cells(DB행, 39).Value
    If PA_SafeD(wsDB.Cells(DB행, 104).Value) <> 0 Then
        ws신규.Cells(89, 11).Value = wsDB.Cells(DB행, 104).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(89, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 37).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 38).Value)
    If v원 <> 0 Then ws신규.Cells(89, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' KT 주간평균/월간평균 (row 90,91) → DB 73~78, 기소 105/106  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 73).Value) <> 0 Then ws신규.Cells(90, 7).Value = wsDB.Cells(DB행, 73).Value
    If PA_SafeD(wsDB.Cells(DB행, 74).Value) <> 0 Then ws신규.Cells(90, 12).Value = wsDB.Cells(DB행, 74).Value
    If PA_SafeD(wsDB.Cells(DB행, 75).Value) <> 0 Then ws신규.Cells(90, 15).Value = wsDB.Cells(DB행, 75).Value
    If PA_SafeD(wsDB.Cells(DB행, 105).Value) <> 0 Then
        ws신규.Cells(90, 11).Value = wsDB.Cells(DB행, 105).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(90, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = 0: v양 = 0
    v원 = PA_SafeD(wsDB.Cells(DB행, 73).Value)
    If v원 = 0 Then v원 = PA_SafeD(wsDB.Cells(DB행, 34).Value) / 7
    v양 = PA_SafeD(wsDB.Cells(DB행, 74).Value)
    If v양 = 0 Then v양 = PA_SafeD(wsDB.Cells(DB행, 35).Value) / 7
    If v원 <> 0 Then ws신규.Cells(90, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 76).Value) <> 0 Then ws신규.Cells(91, 7).Value = wsDB.Cells(DB행, 76).Value
    If PA_SafeD(wsDB.Cells(DB행, 77).Value) <> 0 Then ws신규.Cells(91, 12).Value = wsDB.Cells(DB행, 77).Value
    If PA_SafeD(wsDB.Cells(DB행, 78).Value) <> 0 Then ws신규.Cells(91, 15).Value = wsDB.Cells(DB행, 78).Value
    If PA_SafeD(wsDB.Cells(DB행, 106).Value) <> 0 Then
        ws신규.Cells(91, 11).Value = wsDB.Cells(DB행, 106).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(91, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 76).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 77).Value)
    If v원 <> 0 Then ws신규.Cells(91, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' FT 일계 (Row87 우측) = 수식구간(AP87,AX87) → 덮어쓰기 금지

    ' JRT 일계 (Row92 좌측) = 수식구간(G92,O92) → 덮어쓰기 금지
    ' JRT row 88~91 입력 후 수식 자동 반영

    ' BT 일계 (Row97) = 수식구간(G97,O97) → 덮어쓰기 금지
    ' BT row 98~101 입력 후 수식 자동 반영

    ' PRE 일계 (Row92 우측) = 수식구간(AP92,AX92) → 덮어쓰기 금지
    ' PRE row 93~96 입력 후 수식 자동 반영

    ' 목표원지수율(BK84), 목표생산량(BK85) ? DB 52~53
    If Not IsEmpty(wsDB.Cells(DB행, 52).Value) Then ws신규.Cells(84, 63).Value = wsDB.Cells(DB행, 52).Value
    If Not IsEmpty(wsDB.Cells(DB행, 53).Value) Then ws신규.Cells(85, 63).Value = wsDB.Cells(DB행, 53).Value

    ' FT 주간누계/월간누계 (Row88,89 우측) → DB 54~57, 기소 107/108  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 54).Value) <> 0 Then ws신규.Cells(88, 42).Value = wsDB.Cells(DB행, 54).Value
    If PA_SafeD(wsDB.Cells(DB행, 55).Value) <> 0 Then ws신규.Cells(88, 47).Value = wsDB.Cells(DB행, 55).Value
    If PA_SafeD(wsDB.Cells(DB행, 107).Value) <> 0 Then
        ws신규.Cells(88, 46).Value = wsDB.Cells(DB행, 107).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(88, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 54).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 55).Value)
    If v원 <> 0 Then ws신규.Cells(88, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(88, 60).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 56).Value) <> 0 Then ws신규.Cells(89, 42).Value = wsDB.Cells(DB행, 56).Value
    If PA_SafeD(wsDB.Cells(DB행, 57).Value) <> 0 Then ws신규.Cells(89, 47).Value = wsDB.Cells(DB행, 57).Value
    If PA_SafeD(wsDB.Cells(DB행, 108).Value) <> 0 Then
        ws신규.Cells(89, 46).Value = wsDB.Cells(DB행, 108).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(89, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 56).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 57).Value)
    If v원 <> 0 Then ws신규.Cells(89, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(89, 60).Value = Round((v원 - v양) / v원 * 100, 1)

    ' FT 주간평균/월간평균 (Row90,91 우측) → DB 79~84, 기소 109/110  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 79).Value) <> 0 Then ws신규.Cells(90, 42).Value = wsDB.Cells(DB행, 79).Value
    If PA_SafeD(wsDB.Cells(DB행, 80).Value) <> 0 Then ws신규.Cells(90, 47).Value = wsDB.Cells(DB행, 80).Value
    If PA_SafeD(wsDB.Cells(DB행, 81).Value) <> 0 Then ws신규.Cells(90, 50).Value = wsDB.Cells(DB행, 81).Value
    If PA_SafeD(wsDB.Cells(DB행, 109).Value) <> 0 Then
        ws신규.Cells(90, 46).Value = wsDB.Cells(DB행, 109).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(90, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 82).Value) <> 0 Then ws신규.Cells(91, 42).Value = wsDB.Cells(DB행, 82).Value
    If PA_SafeD(wsDB.Cells(DB행, 83).Value) <> 0 Then ws신규.Cells(91, 47).Value = wsDB.Cells(DB행, 83).Value
    If PA_SafeD(wsDB.Cells(DB행, 84).Value) <> 0 Then ws신규.Cells(91, 50).Value = wsDB.Cells(DB행, 84).Value
    If PA_SafeD(wsDB.Cells(DB행, 110).Value) <> 0 Then
        ws신규.Cells(91, 46).Value = wsDB.Cells(DB행, 110).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(91, 46).Value = wsDB.Cells(DB행, 123).Value
    End If

    ' JRT 주간누계/월간누계 (Row93,94 좌측) → DB 58~63, 기소 111/112  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 58).Value) <> 0 Then ws신규.Cells(93, 7).Value = wsDB.Cells(DB행, 58).Value
    If PA_SafeD(wsDB.Cells(DB행, 59).Value) <> 0 Then ws신규.Cells(93, 12).Value = wsDB.Cells(DB행, 59).Value
    If PA_SafeD(wsDB.Cells(DB행, 111).Value) <> 0 Then
        ws신규.Cells(93, 11).Value = wsDB.Cells(DB행, 111).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(93, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 58).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 59).Value)
    If v원 <> 0 Then ws신규.Cells(93, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(93, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 60).Value) <> 0 Then ws신규.Cells(94, 7).Value = wsDB.Cells(DB행, 60).Value
    If PA_SafeD(wsDB.Cells(DB행, 63).Value) <> 0 Then ws신규.Cells(94, 12).Value = wsDB.Cells(DB행, 63).Value
    If PA_SafeD(wsDB.Cells(DB행, 112).Value) <> 0 Then
        ws신규.Cells(94, 11).Value = wsDB.Cells(DB행, 112).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(94, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 60).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 63).Value)
    If v원 <> 0 Then ws신규.Cells(94, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(94, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' JRT 주간평균/월간평균 (Row95,96 좌측) → DB 85~90, 기소 113/114  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 85).Value) <> 0 Then ws신규.Cells(95, 7).Value = wsDB.Cells(DB행, 85).Value
    If PA_SafeD(wsDB.Cells(DB행, 86).Value) <> 0 Then ws신규.Cells(95, 12).Value = wsDB.Cells(DB행, 86).Value
    If PA_SafeD(wsDB.Cells(DB행, 87).Value) <> 0 Then ws신규.Cells(95, 15).Value = wsDB.Cells(DB행, 87).Value
    If PA_SafeD(wsDB.Cells(DB행, 113).Value) <> 0 Then
        ws신규.Cells(95, 11).Value = wsDB.Cells(DB행, 113).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(95, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 88).Value) <> 0 Then ws신규.Cells(96, 7).Value = wsDB.Cells(DB행, 88).Value
    If PA_SafeD(wsDB.Cells(DB행, 89).Value) <> 0 Then ws신규.Cells(96, 12).Value = wsDB.Cells(DB행, 89).Value
    If PA_SafeD(wsDB.Cells(DB행, 90).Value) <> 0 Then ws신규.Cells(96, 15).Value = wsDB.Cells(DB행, 90).Value
    If PA_SafeD(wsDB.Cells(DB행, 114).Value) <> 0 Then
        ws신규.Cells(96, 11).Value = wsDB.Cells(DB행, 114).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(96, 11).Value = wsDB.Cells(DB행, 124).Value
    End If

    ' BT 주간누계/월간누계 (Row98,99 좌측) → DB 64~67, 기소 115/116  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 64).Value) <> 0 Then ws신규.Cells(98, 7).Value = wsDB.Cells(DB행, 64).Value
    If PA_SafeD(wsDB.Cells(DB행, 65).Value) <> 0 Then ws신규.Cells(98, 12).Value = wsDB.Cells(DB행, 65).Value
    If PA_SafeD(wsDB.Cells(DB행, 115).Value) <> 0 Then
        ws신규.Cells(98, 11).Value = wsDB.Cells(DB행, 115).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(98, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 64).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 65).Value)
    If v원 <> 0 Then ws신규.Cells(98, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(98, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 66).Value) <> 0 Then ws신규.Cells(99, 7).Value = wsDB.Cells(DB행, 66).Value
    If PA_SafeD(wsDB.Cells(DB행, 67).Value) <> 0 Then ws신규.Cells(99, 12).Value = wsDB.Cells(DB행, 67).Value
    If PA_SafeD(wsDB.Cells(DB행, 116).Value) <> 0 Then
        ws신규.Cells(99, 11).Value = wsDB.Cells(DB행, 116).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(99, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 66).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 67).Value)
    If v원 <> 0 Then ws신규.Cells(99, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(99, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' BT 주간평균/월간평균 (Row100,101 좌측) → DB 91~96, 기소 117/118  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 91).Value) <> 0 Then ws신규.Cells(100, 7).Value = wsDB.Cells(DB행, 91).Value
    If PA_SafeD(wsDB.Cells(DB행, 92).Value) <> 0 Then ws신규.Cells(100, 12).Value = wsDB.Cells(DB행, 92).Value
    If PA_SafeD(wsDB.Cells(DB행, 93).Value) <> 0 Then ws신규.Cells(100, 15).Value = wsDB.Cells(DB행, 93).Value
    If PA_SafeD(wsDB.Cells(DB행, 117).Value) <> 0 Then
        ws신규.Cells(100, 11).Value = wsDB.Cells(DB행, 117).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(100, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 94).Value) <> 0 Then ws신규.Cells(101, 7).Value = wsDB.Cells(DB행, 94).Value
    If PA_SafeD(wsDB.Cells(DB행, 95).Value) <> 0 Then ws신규.Cells(101, 12).Value = wsDB.Cells(DB행, 95).Value
    If PA_SafeD(wsDB.Cells(DB행, 96).Value) <> 0 Then ws신규.Cells(101, 15).Value = wsDB.Cells(DB행, 96).Value
    If PA_SafeD(wsDB.Cells(DB행, 118).Value) <> 0 Then
        ws신규.Cells(101, 11).Value = wsDB.Cells(DB행, 118).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(101, 11).Value = wsDB.Cells(DB행, 125).Value
    End If

    ' PRE 주간누계/월간누계 (Row93,94 우측) → DB 68~71, 기소 119/120  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 68).Value) <> 0 Then ws신규.Cells(93, 42).Value = wsDB.Cells(DB행, 68).Value
    If PA_SafeD(wsDB.Cells(DB행, 69).Value) <> 0 Then ws신규.Cells(93, 47).Value = wsDB.Cells(DB행, 69).Value
    If PA_SafeD(wsDB.Cells(DB행, 119).Value) <> 0 Then
        ws신규.Cells(93, 46).Value = wsDB.Cells(DB행, 119).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(93, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 68).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 69).Value)
    If v원 <> 0 Then ws신규.Cells(93, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(93, 60).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 70).Value) <> 0 Then ws신규.Cells(94, 42).Value = wsDB.Cells(DB행, 70).Value
    If PA_SafeD(wsDB.Cells(DB행, 71).Value) <> 0 Then ws신규.Cells(94, 47).Value = wsDB.Cells(DB행, 71).Value
    If PA_SafeD(wsDB.Cells(DB행, 120).Value) <> 0 Then
        ws신규.Cells(94, 46).Value = wsDB.Cells(DB행, 120).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(94, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 70).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 71).Value)
    If v원 <> 0 Then ws신규.Cells(94, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(94, 60).Value = Round((v원 - v양) / v원 * 100, 1)

    ' PRE 주간평균/월간평균 (Row95,96 우측) → DB 97~102, 기소 121/122  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 97).Value) <> 0 Then ws신규.Cells(95, 42).Value = wsDB.Cells(DB행, 97).Value
    If PA_SafeD(wsDB.Cells(DB행, 98).Value) <> 0 Then ws신규.Cells(95, 47).Value = wsDB.Cells(DB행, 98).Value
    If PA_SafeD(wsDB.Cells(DB행, 99).Value) <> 0 Then ws신규.Cells(95, 50).Value = wsDB.Cells(DB행, 99).Value
    If PA_SafeD(wsDB.Cells(DB행, 121).Value) <> 0 Then
        ws신규.Cells(95, 46).Value = wsDB.Cells(DB행, 121).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(95, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 100).Value) <> 0 Then ws신규.Cells(96, 42).Value = wsDB.Cells(DB행, 100).Value
    If PA_SafeD(wsDB.Cells(DB행, 101).Value) <> 0 Then ws신규.Cells(96, 47).Value = wsDB.Cells(DB행, 101).Value
    If PA_SafeD(wsDB.Cells(DB행, 102).Value) <> 0 Then ws신규.Cells(96, 50).Value = wsDB.Cells(DB행, 102).Value
    If PA_SafeD(wsDB.Cells(DB행, 122).Value) <> 0 Then
        ws신규.Cells(96, 46).Value = wsDB.Cells(DB행, 122).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(96, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
End Sub

' ============================================================
' 비가동현황 입력: T12 rows 121-126
' DB F59=특이사항(줄바꿈 구분), F60=해결책(줄바꿈 구분)
' → rows 121~126, A(1)=내용, AJ(36)=해결책
' ============================================================
Private Sub 비가동현황_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long)
    Dim 특이사항 As String, 해결책 As String
    특이사항 = Trim(CStr(wsDB.Cells(DB행, 59).Value))
    해결책 = Trim(CStr(wsDB.Cells(DB행, 60).Value))
    ' vbCr→vbLf 통일 후 Split (Excel/PDF 복붙 시 vbCr/vbCrLf 혼합 대응)
    특이사항 = Replace(특이사항, vbCr, vbLf)
    해결책 = Replace(해결책, vbCr, vbLf)

    Dim 특이줄() As String, 해결줄() As String
    Dim i As Long

    If 특이사항 <> "" Then
        특이줄 = Split(특이사항, vbLf)
        For i = 0 To UBound(특이줄)
            If i > 5 Then Exit For
            Dim v1 As String
            v1 = Trim(특이줄(i))
            If v1 <> "" Then ws신규.Cells(121 + i, 1).Value = v1
        Next i
    End If

    If 해결책 <> "" Then
        해결줄 = Split(해결책, vbLf)
        For i = 0 To UBound(해결줄)
            If i > 5 Then Exit For
            Dim v2 As String
            v2 = Trim(해결줄(i))
            If v2 <> "" Then ws신규.Cells(121 + i, 36).Value = v2
        Next i
    End If
End Sub

' 시트 가시성 안전 복원 (완전숨김/숨김 상태일 수 있음 - 오류 시 대체 시도)
Private Sub 시트가시성_안전복원(ws As Worksheet, 원래상태 As XlSheetVisibility)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Visible = 원래상태
    If Err.Number <> 0 Then
        Err.Clear
        If 원래상태 = xlSheetVeryHidden Then ws.Visible = xlSheetHidden
        If Err.Number <> 0 Then Err.Clear: ws.Visible = xlSheetVisible
    End If
    On Error GoTo 0
End Sub

Private Sub 매크로버튼_삭제(ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub 수식을값으로변환(ws As Worksheet)
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    Dim cell As Range
    On Error Resume Next
    For Each cell In usedRange
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    On Error GoTo 0
End Sub






