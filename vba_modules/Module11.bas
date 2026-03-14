Attribute VB_Name = "Module11"
' ============================================================
' 일일생산DB → 일일생산일보 역변환 시스템 (PDF 출력)
' ============================================================
' [주요 기능]
' 1. 일일생산DB에서 특정 날짜 데이터 추출 → 제공 셀주소에 맞게 입력
' 2. 수식 값 변환
' 3. PDF로 출력 (파일명: 일일생산일지_YYYYMMDD_HHMMSS.pdf)
'
' [일일생산DB 70열 구조] - 역변환_생성과 동일
' [생산일보 시트 매핑 - 생산일보_셀매핑.json 기준]
' BK4 = 작성일(날짜)
' T01 원지 주간 KT: rows 7-16, 6그룹 base(3,15,27,39,51,63), ROLLID=base+2, KG=base+9
' T02 원지 직기 JRT: rows 20-29, 동일 구조
' T04 원지 봉제 KT: rows 35-44
' T05 원지 봉제 JRT: rows 48-57
' T07 완제품 주간(68~77행): B(2)=제품, Q(17)=설비, T(20)=생산량BOX
' T08 완제품 야간(68~77행): AK(37)=제품, AZ(52)=설비, BC(55)=생산량BOX
' T11 작업시간(108행), Direct(111행), Indirect(112행)
'     설비 base: KT=7, FT=43, JRT=19, BT=31, PRE=55
' T10 원자재수율(84~101행): 수식값 역변환 입력
' T12 비가동현황: rows 121-126, A(1)=내용, AJ(36)=해결책
'
' DB 생산일보구분: "생산일보_주간"/"생산일보_야간"/"생산일보"
' ============================================================

Option Explicit

' ============================================================
' 생산일보 역변환용 초기화 (Module1 의존 없음)
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
' 메인 프로시저: UserForm1으로 날짜 선택 후 PDF 출력 실행
' ============================================================
Public Sub 일일생산DB_역변환_일일생산일보_PDF출력()
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
        MsgBox "UserForm1을 찾을 수 없습니다." & vbCrLf & _
               "UserForm1_일일생산일보날짜선택.txt 안내대로 폼을 추가한 뒤 다시 실행하세요.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If UserForm1.Cancelled Or Not IsDate(UserForm1.SelectedDate) Then
        Unload UserForm1
        Exit Sub
    End If

    Call 일일생산DB_역변환_일일생산일보_PDF출력_실행(UserForm1.SelectedDate)
    Unload UserForm1
End Sub

' ============================================================
' 선택한 날짜로 일일생산일보 PDF 출력 실행
' ============================================================
Public Sub 일일생산DB_역변환_일일생산일보_PDF출력_실행(선택날짜 As Variant)
    Dim wsDB As Worksheet
    Dim ws템플릿 As Worksheet
    Dim visDB As XlSheetVisibility, vis템플릿 As XlSheetVisibility
    Dim lastRow As Long
    Dim 데이터존재 As Boolean
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

    Dim wb신규 As Workbook
    Dim ws신규 As Worksheet
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

    파일명 = "일일생산일지_" & Format(선택날짜, "yyyymmdd") & "_" & Format(Now, "hhmmss") & ".pdf"
    저장경로 = ThisWorkbook.Path & "\" & 파일명

    On Error Resume Next
    ws신규.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=저장경로, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    On Error GoTo 0

    wb신규.Close SaveChanges:=False

    On Error Resume Next
    Dim wsHome As Worksheet
    Set wsHome = ThisWorkbook.Sheets("Home menu")
    If Not wsHome Is Nothing Then wsHome.Activate
    On Error GoTo 0

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "일일생산일보가 PDF로 출력되었습니다!" & vbCrLf & vbCrLf & _
           "파일명: " & 파일명 & vbCrLf & _
           "저장 위치: " & ThisWorkbook.Path, vbInformation
    Exit Sub

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
End Sub

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

    Dim DB날짜 As Variant, 날짜일치 As Boolean, 마지막유효날짜 As Variant, 생산일보구분 As String, 레코드유형 As String
    마지막유효날짜 = Empty

    ' ★ 원지 누락방지: KT계(rows 7/35)와 JRT계(rows 20/48)가 row 20~26에서 겹침. JRT 먼저, KT 나중 처리
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
                ' 주야 구분: "야간" 포함→우측(T08), 그 외("주간"/"생산일보" 등)→좌측(T07)
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
    Dim ROLLID As String, KG As Variant
    ROLLID = Trim(CStr(wsDB.Cells(DB행, 7).Value))
    KG = wsDB.Cells(DB행, 8).Value
    If ROLLID = "" And IsEmpty(KG) Then Exit Sub

    Dim 시작행 As Long, 끝행 As Long, ROLLID열 As Long, KG열 As Long
    If Not 원지_매핑_구하기(생산일보구분, 종별, 시작행, 끝행, ROLLID열, KG열) Then Exit Sub

    Dim 맵키 As String: 맵키 = 생산일보구분 & "|" & 종별
    If Not 다음행딕.exists(맵키) Then 다음행딕.Add 맵키, 0
    Dim nextOffset As Long: nextOffset = CLng(다음행딕(맵키))

    Dim 대상행 As Long
    대상행 = 시작행 + nextOffset
    If 대상행 > 끝행 Then Exit Sub  ' 오버플로우: 3패스에서 별도 처리 (카운터 증가 안 함)

    다음행딕(맵키) = nextOffset + 1

    Dim rngROLL As Range, rngKG As Range
    Set rngROLL = ws신규.Cells(대상행, ROLLID열)
    Set rngKG = ws신규.Cells(대상행, KG열)
    On Error Resume Next
    If rngROLL.MergeCells And rngROLL.MergeArea.Rows.Count = 1 Then Set rngROLL = rngROLL.MergeArea.Cells(1, 1)
    If rngKG.MergeCells And rngKG.MergeArea.Rows.Count = 1 Then Set rngKG = rngKG.MergeArea.Cells(1, 1)
    On Error GoTo 0

    If ROLLID <> "" Then rngROLL.Value = wsDB.Cells(DB행, 7).Value
    If Not IsEmpty(KG) Then rngKG.Value = KG
End Sub

' ============================================================
' 원지 오버플로우 입력 (3패스에서 호출) - Module10과 동일
' 정규 기록 완료 후 ws신규 빈 셀 스캔해서 배치
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
    If 대상행 <= 끝행 Then Exit Sub

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

    Dim baseCol As Long, r As Long
    Dim targetRow As Long, targetROLL As Long, targetKG As Long
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
                targetROLL = baseCol + 2
                targetKG = baseCol + 9
                found = True
                Exit For
            End If
        Next r
        If found Then Exit For
    Next i
    If Not found Then Exit Sub

    Dim rngROLL As Range, rngKG As Range
    Set rngROLL = ws신규.Cells(targetRow, targetROLL)
    Set rngKG = ws신규.Cells(targetRow, targetKG)
    On Error Resume Next
    If rngROLL.MergeCells And rngROLL.MergeArea.Rows.Count = 1 Then Set rngROLL = rngROLL.MergeArea.Cells(1, 1)
    If rngKG.MergeCells And rngKG.MergeArea.Rows.Count = 1 Then Set rngKG = rngKG.MergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim ROLLID As String, KG As Variant
    ROLLID = Trim(CStr(wsDB.Cells(DB행, 7).Value))
    KG = wsDB.Cells(DB행, 8).Value
    If ROLLID <> "" Then rngROLL.Value = wsDB.Cells(DB행, 7).Value
    If Not IsEmpty(KG) Then rngKG.Value = KG

    Dim memo As String
    memo = 종별 & " 원본 seq" & 순번 & ", 중량 " & IIf(IsEmpty(KG), "-", CStr(KG)) & " kg (분산배치)"
    On Error Resume Next
    rngROLL.ClearComments: rngROLL.AddComment: rngROLL.Comment.Text memo: rngROLL.Comment.Visible = False
    rngKG.ClearComments: rngKG.AddComment: rngKG.Comment.Text memo: rngKG.Comment.Visible = False
    On Error GoTo 0
    rngROLL.Interior.color = RGB(255, 182, 193)
    rngKG.Interior.color = RGB(255, 182, 193)
End Sub

' 원지 종별 + 생산일보구분 → (시작행, 끝행, ROLLID열, KG열) 반환
' ★ 블록별 행범위: KT 7-16/35-44, JRT 20-29/48-57. 순번 11+가 다른 블록 덮어쓰지 않도록 끝행 검사
Private Function 원지_매핑_구하기(생산일보구분 As String, 종별 As String, _
    ByRef 시작행 As Long, ByRef 끝행 As Long, ByRef ROLLID열 As Long, ByRef KG열 As Long) As Boolean

    원지_매핑_구하기 = False
    Dim base열 As Long
    base열 = -1

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
' B(2)=품명, Q(17)=설비호기, T(20)=생산량BOX, X(24)=소요량, AB(28)=PLT1, AD(30)=PLT2
' AF(32)=AJ11, AH(34)=AJ13: DB F16/F17에 저장된 값 우선, 없으면 시트 수식 자동계산
' ============================================================
' 완제품 주간: B,Q,T만 값입력, X68:AF78=수식구간 덮어쓰기 금지
Private Sub 완제품_주간_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 순번 As Long)
    If 순번 < 1 Or 순번 > 10 Then Exit Sub
    Dim 대상행 As Long
    대상행 = 67 + 순번

    If Not IsEmpty(wsDB.Cells(DB행, 10).Value) Then ws신규.Cells(대상행, 2).Value = wsDB.Cells(DB행, 10).Value
    If Not IsEmpty(wsDB.Cells(DB행, 11).Value) Then ws신규.Cells(대상행, 17).Value = wsDB.Cells(DB행, 11).Value
    If Not IsEmpty(wsDB.Cells(DB행, 12).Value) Then ws신규.Cells(대상행, 20).Value = wsDB.Cells(DB행, 12).Value
End Sub

' ============================================================
' 완제품 야간: AK,AZ,BC만 값입력, BG68:BO78=수식구간 덮어쓰기 금지
Private Sub 완제품_야간_입력(wsDB As Worksheet, ws신규 As Worksheet, DB행 As Long, 순번 As Long)
    If 순번 < 1 Or 순번 > 10 Then Exit Sub
    Dim 대상행 As Long
    대상행 = 67 + 순번

    ws신규.Cells(대상행, 37).Value = wsDB.Cells(DB행, 10).Value
    ws신규.Cells(대상행, 52).Value = wsDB.Cells(DB행, 11).Value
    ws신규.Cells(대상행, 55).Value = wsDB.Cells(DB행, 12).Value
End Sub

' ============================================================
' 투입인원 T11 재입력 (수식변환 후 호출)
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
    Dim lastRow As Long, r As Long, DB날짜 As Variant, 날짜일치 As Boolean, 마지막유효날짜 As Variant
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
' 02시트는 주간(정규/연장)만 → base(정규), base+2(연장)만
' KT: G108,I108(인력), G109,I109(작업) / JRT: S108,U108, S109,U109
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
        v = wsDB.Cells(DB행, 21).Value: If IsNumeric(v) Then ws신규.Cells(108, base열).Value = CDbl(v)
        v = wsDB.Cells(DB행, 22).Value
        If Not IsNumeric(v) Then v = wsDB.Cells(DB행, 21).Value
        If IsNumeric(v) Then ws신규.Cells(109, base열).Value = CDbl(v)
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
        v = wsDB.Cells(DB행, 21).Value: If IsNumeric(v) Then ws신규.Cells(108, base열 + 2).Value = CDbl(v)
        v = wsDB.Cells(DB행, 22).Value
        If IsNumeric(v) Then ws신규.Cells(109, base열 + 2).Value = CDbl(v)
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

    ' KT 일계 (row 87) = 수식구간(G87,O87) → 덮어쓰기 금지

    ' KT 주간누계/월간누계 (row 88,89) ? DB 34~39, 기소 103/104  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 34).Value) <> 0 Then ws신규.Cells(88, 7).Value = wsDB.Cells(DB행, 34).Value
    If PA_SafeD(wsDB.Cells(DB행, 35).Value) <> 0 Then ws신규.Cells(88, 12).Value = wsDB.Cells(DB행, 35).Value
    If PA_SafeD(wsDB.Cells(DB행, 36).Value) <> 0 Then ws신규.Cells(88, 15).Value = wsDB.Cells(DB행, 36).Value
    If Not IsEmpty(wsDB.Cells(DB행, 103).Value) And IsNumeric(wsDB.Cells(DB행, 103).Value) Then
        ws신규.Cells(88, 11).Value = wsDB.Cells(DB행, 103).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(88, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 34).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 35).Value)
    If v원 <> 0 Then ws신규.Cells(88, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    If PA_SafeD(wsDB.Cells(DB행, 37).Value) <> 0 Then ws신규.Cells(89, 7).Value = wsDB.Cells(DB행, 37).Value
    If PA_SafeD(wsDB.Cells(DB행, 38).Value) <> 0 Then ws신규.Cells(89, 12).Value = wsDB.Cells(DB행, 38).Value
    If PA_SafeD(wsDB.Cells(DB행, 39).Value) <> 0 Then ws신규.Cells(89, 15).Value = wsDB.Cells(DB행, 39).Value
    If Not IsEmpty(wsDB.Cells(DB행, 104).Value) And IsNumeric(wsDB.Cells(DB행, 104).Value) Then
        ws신규.Cells(89, 11).Value = wsDB.Cells(DB행, 104).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(89, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 37).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 38).Value)
    If v원 <> 0 Then ws신규.Cells(89, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' KT 주간평균/월간평균 (row 90,91) ? DB 73~78, 기소 105/106  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 73).Value) <> 0 Then ws신규.Cells(90, 7).Value = wsDB.Cells(DB행, 73).Value
    If PA_SafeD(wsDB.Cells(DB행, 74).Value) <> 0 Then ws신규.Cells(90, 12).Value = wsDB.Cells(DB행, 74).Value
    If PA_SafeD(wsDB.Cells(DB행, 75).Value) <> 0 Then ws신규.Cells(90, 15).Value = wsDB.Cells(DB행, 75).Value
    If Not IsEmpty(wsDB.Cells(DB행, 105).Value) And IsNumeric(wsDB.Cells(DB행, 105).Value) Then
        ws신규.Cells(90, 11).Value = wsDB.Cells(DB행, 105).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(90, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 73).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 74).Value)
    If v원 <> 0 Then ws신규.Cells(90, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 76).Value) <> 0 Then ws신규.Cells(91, 7).Value = wsDB.Cells(DB행, 76).Value
    If PA_SafeD(wsDB.Cells(DB행, 77).Value) <> 0 Then ws신규.Cells(91, 12).Value = wsDB.Cells(DB행, 77).Value
    If PA_SafeD(wsDB.Cells(DB행, 78).Value) <> 0 Then ws신규.Cells(91, 15).Value = wsDB.Cells(DB행, 78).Value
    If Not IsEmpty(wsDB.Cells(DB행, 106).Value) And IsNumeric(wsDB.Cells(DB행, 106).Value) Then
        ws신규.Cells(91, 11).Value = wsDB.Cells(DB행, 106).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 31).Value) And IsNumeric(wsDB.Cells(DB행, 31).Value) Then
        ws신규.Cells(91, 11).Value = wsDB.Cells(DB행, 31).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 76).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 77).Value)
    If v원 <> 0 Then ws신규.Cells(91, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' FT 일계 (Row87 우측) = 수식구간(AP87,AX87) → 덮어쓰기 금지

    ' JRT 일계 (Row92 좌측) = 수식구간(G92,O92) → 덮어쓰기 금지
    ' JRT row 93~96 입력 후 수식 자동 반영

    ' BT 일계 (Row97) = 수식구간(G97,O97) → 덮어쓰기 금지
    ' BT row 98~101 입력 후 수식 자동 반영

    ' PRE 일계 (Row92 우측) = 수식구간(AP92,AX92) → 덮어쓰기 금지
    ' PRE row 93~96 입력 후 수식 자동 반영

    ' 목표원지수율(BK84), 목표생산량(BK85) ? DB 52~53
    If Not IsEmpty(wsDB.Cells(DB행, 52).Value) Then ws신규.Cells(84, 63).Value = wsDB.Cells(DB행, 52).Value
    If Not IsEmpty(wsDB.Cells(DB행, 53).Value) Then ws신규.Cells(85, 63).Value = wsDB.Cells(DB행, 53).Value

    ' FT 주간누계/월간누계 (Row88,89 우측) ? DB 54~57, 기소 107/108  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 54).Value) <> 0 Then ws신규.Cells(88, 42).Value = wsDB.Cells(DB행, 54).Value
    If PA_SafeD(wsDB.Cells(DB행, 55).Value) <> 0 Then ws신규.Cells(88, 47).Value = wsDB.Cells(DB행, 55).Value
    If Not IsEmpty(wsDB.Cells(DB행, 107).Value) And IsNumeric(wsDB.Cells(DB행, 107).Value) Then
        ws신규.Cells(88, 46).Value = wsDB.Cells(DB행, 107).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(88, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 54).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 55).Value)
    If v원 <> 0 Then ws신규.Cells(88, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(88, 60).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 56).Value) <> 0 Then ws신규.Cells(89, 42).Value = wsDB.Cells(DB행, 56).Value
    If PA_SafeD(wsDB.Cells(DB행, 57).Value) <> 0 Then ws신규.Cells(89, 47).Value = wsDB.Cells(DB행, 57).Value
    If Not IsEmpty(wsDB.Cells(DB행, 108).Value) And IsNumeric(wsDB.Cells(DB행, 108).Value) Then
        ws신규.Cells(89, 46).Value = wsDB.Cells(DB행, 108).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(89, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 56).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 57).Value)
    If v원 <> 0 Then ws신규.Cells(89, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(89, 60).Value = Round((v원 - v양) / v원 * 100, 1)

    ' FT 주간평균/월간평균 (Row90,91 우측) ? DB 79~84, 기소 109/110  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 79).Value) <> 0 Then ws신규.Cells(90, 42).Value = wsDB.Cells(DB행, 79).Value
    If PA_SafeD(wsDB.Cells(DB행, 80).Value) <> 0 Then ws신규.Cells(90, 47).Value = wsDB.Cells(DB행, 80).Value
    If PA_SafeD(wsDB.Cells(DB행, 81).Value) <> 0 Then ws신규.Cells(90, 50).Value = wsDB.Cells(DB행, 81).Value
    If Not IsEmpty(wsDB.Cells(DB행, 109).Value) And IsNumeric(wsDB.Cells(DB행, 109).Value) Then
        ws신규.Cells(90, 46).Value = wsDB.Cells(DB행, 109).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(90, 46).Value = wsDB.Cells(DB행, 123).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 82).Value) <> 0 Then ws신규.Cells(91, 42).Value = wsDB.Cells(DB행, 82).Value
    If PA_SafeD(wsDB.Cells(DB행, 83).Value) <> 0 Then ws신규.Cells(91, 47).Value = wsDB.Cells(DB행, 83).Value
    If PA_SafeD(wsDB.Cells(DB행, 84).Value) <> 0 Then ws신규.Cells(91, 50).Value = wsDB.Cells(DB행, 84).Value
    If Not IsEmpty(wsDB.Cells(DB행, 110).Value) And IsNumeric(wsDB.Cells(DB행, 110).Value) Then
        ws신규.Cells(91, 46).Value = wsDB.Cells(DB행, 110).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 123).Value) And IsNumeric(wsDB.Cells(DB행, 123).Value) Then
        ws신규.Cells(91, 46).Value = wsDB.Cells(DB행, 123).Value
    End If

    ' JRT 주간누계/월간누계 (Row93,94 좌측) ? DB 58~63, 기소 111/112  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 58).Value) <> 0 Then ws신규.Cells(93, 7).Value = wsDB.Cells(DB행, 58).Value
    If PA_SafeD(wsDB.Cells(DB행, 59).Value) <> 0 Then ws신규.Cells(93, 12).Value = wsDB.Cells(DB행, 59).Value
    If Not IsEmpty(wsDB.Cells(DB행, 111).Value) And IsNumeric(wsDB.Cells(DB행, 111).Value) Then
        ws신규.Cells(93, 11).Value = wsDB.Cells(DB행, 111).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(93, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 58).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 59).Value)
    If v원 <> 0 Then ws신규.Cells(93, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(93, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 60).Value) <> 0 Then ws신규.Cells(94, 7).Value = wsDB.Cells(DB행, 60).Value
    If PA_SafeD(wsDB.Cells(DB행, 63).Value) <> 0 Then ws신규.Cells(94, 12).Value = wsDB.Cells(DB행, 63).Value
    If Not IsEmpty(wsDB.Cells(DB행, 112).Value) And IsNumeric(wsDB.Cells(DB행, 112).Value) Then
        ws신규.Cells(94, 11).Value = wsDB.Cells(DB행, 112).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(94, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 60).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 63).Value)
    If v원 <> 0 Then ws신규.Cells(94, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(94, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' JRT 주간평균/월간평균 (Row95,96 좌측) ? DB 85~90, 기소 113/114  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 85).Value) <> 0 Then ws신규.Cells(95, 7).Value = wsDB.Cells(DB행, 85).Value
    If PA_SafeD(wsDB.Cells(DB행, 86).Value) <> 0 Then ws신규.Cells(95, 12).Value = wsDB.Cells(DB행, 86).Value
    If PA_SafeD(wsDB.Cells(DB행, 87).Value) <> 0 Then ws신규.Cells(95, 15).Value = wsDB.Cells(DB행, 87).Value
    If Not IsEmpty(wsDB.Cells(DB행, 113).Value) And IsNumeric(wsDB.Cells(DB행, 113).Value) Then
        ws신규.Cells(95, 11).Value = wsDB.Cells(DB행, 113).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(95, 11).Value = wsDB.Cells(DB행, 124).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 88).Value) <> 0 Then ws신규.Cells(96, 7).Value = wsDB.Cells(DB행, 88).Value
    If PA_SafeD(wsDB.Cells(DB행, 89).Value) <> 0 Then ws신규.Cells(96, 12).Value = wsDB.Cells(DB행, 89).Value
    If PA_SafeD(wsDB.Cells(DB행, 90).Value) <> 0 Then ws신규.Cells(96, 15).Value = wsDB.Cells(DB행, 90).Value
    If Not IsEmpty(wsDB.Cells(DB행, 114).Value) And IsNumeric(wsDB.Cells(DB행, 114).Value) Then
        ws신규.Cells(96, 11).Value = wsDB.Cells(DB행, 114).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 124).Value) And IsNumeric(wsDB.Cells(DB행, 124).Value) Then
        ws신규.Cells(96, 11).Value = wsDB.Cells(DB행, 124).Value
    End If

    ' BT 주간누계/월간누계 (Row98,99 좌측) ? DB 64~67, 기소 115/116  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 64).Value) <> 0 Then ws신규.Cells(98, 7).Value = wsDB.Cells(DB행, 64).Value
    If PA_SafeD(wsDB.Cells(DB행, 65).Value) <> 0 Then ws신규.Cells(98, 12).Value = wsDB.Cells(DB행, 65).Value
    If Not IsEmpty(wsDB.Cells(DB행, 115).Value) And IsNumeric(wsDB.Cells(DB행, 115).Value) Then
        ws신규.Cells(98, 11).Value = wsDB.Cells(DB행, 115).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(98, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 64).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 65).Value)
    If v원 <> 0 Then ws신규.Cells(98, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(98, 25).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 66).Value) <> 0 Then ws신규.Cells(99, 7).Value = wsDB.Cells(DB행, 66).Value
    If PA_SafeD(wsDB.Cells(DB행, 67).Value) <> 0 Then ws신규.Cells(99, 12).Value = wsDB.Cells(DB행, 67).Value
    If Not IsEmpty(wsDB.Cells(DB행, 116).Value) And IsNumeric(wsDB.Cells(DB행, 116).Value) Then
        ws신규.Cells(99, 11).Value = wsDB.Cells(DB행, 116).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(99, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 66).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 67).Value)
    If v원 <> 0 Then ws신규.Cells(99, 15).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(99, 25).Value = Round((v원 - v양) / v원 * 100, 1)

    ' BT 주간평균/월간평균 (Row100,101 좌측) ? DB 91~96, 기소 117/118  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 91).Value) <> 0 Then ws신규.Cells(100, 7).Value = wsDB.Cells(DB행, 91).Value
    If PA_SafeD(wsDB.Cells(DB행, 92).Value) <> 0 Then ws신규.Cells(100, 12).Value = wsDB.Cells(DB행, 92).Value
    If PA_SafeD(wsDB.Cells(DB행, 93).Value) <> 0 Then ws신규.Cells(100, 15).Value = wsDB.Cells(DB행, 93).Value
    If Not IsEmpty(wsDB.Cells(DB행, 117).Value) And IsNumeric(wsDB.Cells(DB행, 117).Value) Then
        ws신규.Cells(100, 11).Value = wsDB.Cells(DB행, 117).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(100, 11).Value = wsDB.Cells(DB행, 125).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 94).Value) <> 0 Then ws신규.Cells(101, 7).Value = wsDB.Cells(DB행, 94).Value
    If PA_SafeD(wsDB.Cells(DB행, 95).Value) <> 0 Then ws신규.Cells(101, 12).Value = wsDB.Cells(DB행, 95).Value
    If PA_SafeD(wsDB.Cells(DB행, 96).Value) <> 0 Then ws신규.Cells(101, 15).Value = wsDB.Cells(DB행, 96).Value
    If Not IsEmpty(wsDB.Cells(DB행, 118).Value) And IsNumeric(wsDB.Cells(DB행, 118).Value) Then
        ws신규.Cells(101, 11).Value = wsDB.Cells(DB행, 118).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 125).Value) And IsNumeric(wsDB.Cells(DB행, 125).Value) Then
        ws신규.Cells(101, 11).Value = wsDB.Cells(DB행, 125).Value
    End If

    ' PRE 주간누계/월간누계 (Row93,94 우측) ? DB 68~71, 기소 119/120  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 68).Value) <> 0 Then ws신규.Cells(93, 42).Value = wsDB.Cells(DB행, 68).Value
    If PA_SafeD(wsDB.Cells(DB행, 69).Value) <> 0 Then ws신규.Cells(93, 47).Value = wsDB.Cells(DB행, 69).Value
    If Not IsEmpty(wsDB.Cells(DB행, 119).Value) And IsNumeric(wsDB.Cells(DB행, 119).Value) Then
        ws신규.Cells(93, 46).Value = wsDB.Cells(DB행, 119).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(93, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 68).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 69).Value)
    If v원 <> 0 Then ws신규.Cells(93, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(93, 60).Value = Round((v원 - v양) / v원 * 100, 1)
    If PA_SafeD(wsDB.Cells(DB행, 70).Value) <> 0 Then ws신규.Cells(94, 42).Value = wsDB.Cells(DB행, 70).Value
    If PA_SafeD(wsDB.Cells(DB행, 71).Value) <> 0 Then ws신규.Cells(94, 47).Value = wsDB.Cells(DB행, 71).Value
    If Not IsEmpty(wsDB.Cells(DB행, 120).Value) And IsNumeric(wsDB.Cells(DB행, 120).Value) Then
        ws신규.Cells(94, 46).Value = wsDB.Cells(DB행, 120).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(94, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    v원 = PA_SafeD(wsDB.Cells(DB행, 70).Value): v양 = PA_SafeD(wsDB.Cells(DB행, 71).Value)
    If v원 <> 0 Then ws신규.Cells(94, 50).Value = Round(v양 / v원 * 100, 1): ws신규.Cells(94, 60).Value = Round((v원 - v양) / v원 * 100, 1)

    ' PRE 주간평균/월간평균 (Row95,96 우측) ? DB 97~102, 기소 121/122  ★ DB값=0이면 주월계산값 유지
    If PA_SafeD(wsDB.Cells(DB행, 97).Value) <> 0 Then ws신규.Cells(95, 42).Value = wsDB.Cells(DB행, 97).Value
    If PA_SafeD(wsDB.Cells(DB행, 98).Value) <> 0 Then ws신규.Cells(95, 47).Value = wsDB.Cells(DB행, 98).Value
    If PA_SafeD(wsDB.Cells(DB행, 99).Value) <> 0 Then ws신규.Cells(95, 50).Value = wsDB.Cells(DB행, 99).Value
    If Not IsEmpty(wsDB.Cells(DB행, 121).Value) And IsNumeric(wsDB.Cells(DB행, 121).Value) Then
        ws신규.Cells(95, 46).Value = wsDB.Cells(DB행, 121).Value
    ElseIf Not IsEmpty(wsDB.Cells(DB행, 126).Value) And IsNumeric(wsDB.Cells(DB행, 126).Value) Then
        ws신규.Cells(95, 46).Value = wsDB.Cells(DB행, 126).Value
    End If
    If PA_SafeD(wsDB.Cells(DB행, 100).Value) <> 0 Then ws신규.Cells(96, 42).Value = wsDB.Cells(DB행, 100).Value
    If PA_SafeD(wsDB.Cells(DB행, 101).Value) <> 0 Then ws신규.Cells(96, 47).Value = wsDB.Cells(DB행, 101).Value
    If PA_SafeD(wsDB.Cells(DB행, 102).Value) <> 0 Then ws신규.Cells(96, 50).Value = wsDB.Cells(DB행, 102).Value
    If Not IsEmpty(wsDB.Cells(DB행, 122).Value) And IsNumeric(wsDB.Cells(DB행, 122).Value) Then
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

    Dim 특이줄() As String, 해결줄() As String
    Dim i As Long

    If 특이사항 <> "" Then
        특이줄 = Split(특이사항, vbLf)
        For i = 0 To UBound(특이줄)
            If i > 5 Then Exit For
            Dim v1 As String
            v1 = Trim(Replace(특이줄(i), vbCr, ""))
            If v1 <> "" Then ws신규.Cells(121 + i, 1).Value = v1
        Next i
    End If

    If 해결책 <> "" Then
        해결줄 = Split(해결책, vbLf)
        For i = 0 To UBound(해결줄)
            If i > 5 Then Exit For
            Dim v2 As String
            v2 = Trim(Replace(해결줄(i), vbCr, ""))
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
    For Each cell In usedRange
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
End Sub






