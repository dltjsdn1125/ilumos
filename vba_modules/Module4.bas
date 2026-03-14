Attribute VB_Name = "Module4"
    ' ============================================================
    ' 일일생산일보 → 일일생산DB 최적화 데이터 전송 시스템 (최종판 v3)
    ' ============================================================
    ' 작성일: 2025-01-XX
    '
    ' [주요 기능]
    ' 1. "일일생산DB" 시트 자동 생성 (없을 경우)
    ' 2. 일일생산일보 시트의 모든 데이터를 최적화된 flat DB 구조로 전송
    ' 3. 생산일보 I (원지, 완제품 KT, 투입인원) 및 생산일보 II (JRT/WT/FT/BT/지관) 모두 포함
    ' 4. Production Analysis 데이터 포함 (일/주/월 세분화)
    ' 5. 투입인원 데이터 정규/연장 분리 저장
    ' 6. 증분 업데이트 - 추가/변경/삭제된 레코드만 정확히 처리
    '
    ' [DB 구조] - 100열 (1~70 역변환 호환, 71~100 T10 주간/월간 평균)
    ' A=생산일, B=ID, C~E 공통, F~H 원지, I~P 완제품, Q~Y 투입인원
    ' Z~AK(26~55) PA_일/주/월 각 10개, AL~AP(56~60) Todays/특이사항/해결책/담당/이사
    ' AQ~AX(61~70) PA2_구분/기간/원지투입~목표인당PLT_월/원지_Co, 71~100 T10 주·월 평균
    ' ============================================================

    Option Explicit

    ' ★ 프로파일링용 모듈 변수 (재고반영 → 증분_업데이트로 시간 전달)
    Private m_재고프로파일 As String

    ' 제품표준사양서: DB 저장 전 계산 반영용 (Y/AC, AW/AZ = KP11→RoundUp(Q or AO/AJ), AJ13→동일)
    Private Const 시트명표준사양서 As String = "제품표준사양서"
    Private Const 표준_G열 As Long = 7
    Private Const 표준_AG열 As Long = 33
    Private Const 표준_AJ열 As Long = 36
    Private Const 표준_데이터시작행 As Long = 9

    ' ============================================================
    ' 메인 프로시저: 최적화DB 시트 생성
    ' ============================================================
    Public Sub 최적화DB시트_생성()
        Dim wsDB As Worksheet
        Dim 시트명 As String
        시트명 = "일일생산DB"

        On Error Resume Next
        Set wsDB = ThisWorkbook.Worksheets(시트명)
        On Error GoTo 0

        If Not wsDB Is Nothing Then
            MsgBox "'" & 시트명 & "' 시트가 이미 존재합니다." & vbCrLf & _
                "기존 시트를 사용합니다.", vbInformation
            Exit Sub
        End If

        Set wsDB = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsDB.Name = 시트명

        Call 헤더_생성(wsDB)

            MsgBox "'" & 시트명 & "' 시트가 생성되었습니다." & vbCrLf & _
            "총 124개의 헤더가 설정되었습니다.", vbInformation
    End Sub

    ' ============================================================
    ' 헤더 생성 (100열: 1~70 역변환 호환, 71~100 T10 주·월 평균)
    ' A=ID, B=작성일, C=생산일보구분, D=레코드_유형, E=순번
    ' F~H 원지, I~P 완제품, Q~Y 인원, Z~AP(26~60) PA/Todays/특이사항/해결책/담당/이사
    ' AQ~AX(61~70) PA2/원지_Co, 71~100 KT/FT/JRT/BT/PRE 주평균·월평균(원투/양품/수율)
    ' ============================================================
    Public Sub 헤더_생성(wsDB As Worksheet)
        ' ★ 컬럼 구조 (126열):
        '  1~5:공통  6~8:원지  9~15:완제품  16:AJ11  17:AJ13
        '  18:인원   19:인원_작업장  20:인원_구분
        '  21:주간정규  22:주간연장  23:야간정규  24:야간심야  25:휴일
        '  26:인당생산량_실투  27:인당생산량_환산투
        '  28~33:KT일계  34~36:KT주누계  37~39:KT월누계
        '  40~42:FT일계  43~45:JRT일계  46~48:BT일계  49~51:PRE일계
        '  52~53:목표  54~57:FT누계  58~60:JRT누계  61~62:담당이사
        '  63~67:JRT월+BT누계  68~72:PRE누계+Co
        '  73~102:주평균/월평균  103~126:기소
        Dim col As Long
        col = 1
        wsDB.Cells(1, col).Value = "생산일": col = col + 1          ' 1
        wsDB.Cells(1, col).Value = "ID": col = col + 1               ' 2
        wsDB.Cells(1, col).Value = "생산일보구분": col = col + 1     ' 3
        wsDB.Cells(1, col).Value = "레코드_유형": col = col + 1      ' 4
        wsDB.Cells(1, col).Value = "순번": col = col + 1             ' 5
        wsDB.Cells(1, col).Value = "원지_종별": col = col + 1        ' 6
        wsDB.Cells(1, col).Value = "원지_ROLL_ID": col = col + 1     ' 7
        wsDB.Cells(1, col).Value = "원지_KG": col = col + 1          ' 8
        wsDB.Cells(1, col).Value = "완제품_구분": col = col + 1      ' 9
        wsDB.Cells(1, col).Value = "품명": col = col + 1             ' 10 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "설비호기": col = col + 1         ' 11 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "생산량_BOX": col = col + 1       ' 12
        wsDB.Cells(1, col).Value = "원지기준소요량": col = col + 1   ' 13
        wsDB.Cells(1, col).Value = "PLT1": col = col + 1             ' 14
        wsDB.Cells(1, col).Value = "PLT2": col = col + 1             ' 15
        wsDB.Cells(1, col).Value = "AJ11": col = col + 1             ' 16 ← O열 우측 (AF32)
        wsDB.Cells(1, col).Value = "AJ13": col = col + 1             ' 17 ← (AH34)
        wsDB.Cells(1, col).Value = "인원": col = col + 1             ' 18
        wsDB.Cells(1, col).Value = "인원_작업장": col = col + 1      ' 19
        wsDB.Cells(1, col).Value = "인원_구분": col = col + 1        ' 20
        wsDB.Cells(1, col).Value = "주간정규": col = col + 1         ' 21 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "주간연장": col = col + 1         ' 22 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "야간정규": col = col + 1         ' 23 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "야간심야": col = col + 1         ' 24 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "휴일": col = col + 1             ' 25 ← 생산일보 일치
        wsDB.Cells(1, col).Value = "인당생산량_실투": col = col + 1  ' 26
        wsDB.Cells(1, col).Value = "인당생산량_환산투": col = col + 1 ' 27
        ' KT 일계 6열 (28~33)
        wsDB.Cells(1, col).Value = "KT_일계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "KT_일계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "KT_일계_수율": col = col + 1
        wsDB.Cells(1, col).Value = "KT_일계_기소평균": col = col + 1
        wsDB.Cells(1, col).Value = "KT_일계_LOSS": col = col + 1
        wsDB.Cells(1, col).Value = "KT_일계_양품율": col = col + 1
        ' KT 주간누계 3열 (34~36)
        wsDB.Cells(1, col).Value = "KT_주누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "KT_주누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "KT_주누계_수율": col = col + 1
        ' KT 월간누계 3열 (37~39)
        wsDB.Cells(1, col).Value = "KT_월누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월누계_수율": col = col + 1
        ' FT 일계 3열 (40~42)
        wsDB.Cells(1, col).Value = "FT_일계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "FT_일계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "FT_일계_수율": col = col + 1
        ' JRT 일계 3열 (43~45)
        wsDB.Cells(1, col).Value = "JRT_일계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_일계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_일계_수율": col = col + 1
        ' BT 일계 3열 (46~48)
        wsDB.Cells(1, col).Value = "BT_일계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "BT_일계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "BT_일계_수율": col = col + 1
        ' PRE 일계 3열 (49~51)
        wsDB.Cells(1, col).Value = "PRE_일계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_일계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_일계_수율": col = col + 1
        ' 목표 2열 (52~53)
        wsDB.Cells(1, col).Value = "목표원지수율": col = col + 1
        wsDB.Cells(1, col).Value = "목표생산량": col = col + 1
        ' FT 주간/월간 누계 4열 (54~57)
        wsDB.Cells(1, col).Value = "FT_주누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월누계_양품": col = col + 1
        ' JRT 주간누계 2열 + 월간누계 원투 1열 (58~60)
        wsDB.Cells(1, col).Value = "JRT_주누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월누계_원투": col = col + 1
        ' 담당/이사 (61~62)
        wsDB.Cells(1, col).Value = "담당": col = col + 1
        wsDB.Cells(1, col).Value = "이사": col = col + 1
        ' JRT 월간누계 양품 (63)
        wsDB.Cells(1, col).Value = "JRT_월누계_양품": col = col + 1
        ' BT 주간/월간 누계 4열 (64~67)
        wsDB.Cells(1, col).Value = "BT_주누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월누계_양품": col = col + 1
        ' PRE 주간/월간 누계 4열 (68~71)
        wsDB.Cells(1, col).Value = "PRE_주누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주누계_양품": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월누계_원투": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월누계_양품": col = col + 1
        ' 원지_Co (72)
        wsDB.Cells(1, col).Value = "원지_Co": col = col + 1
        ' T10 주간평균/월간평균 30열 (73~102)
        wsDB.Cells(1, col).Value = "KT_주평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "KT_주평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "KT_주평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주평균_수율": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월평균_원투": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월평균_양품": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월평균_수율": col = col + 1
        ' 기준소요량 주간/월간 누계·평균 (103~122): KT,FT,JRT,BT,PRE 각 4열
        wsDB.Cells(1, col).Value = "KT_주누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "KT_주평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "KT_월평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "FT_주평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "FT_월평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_주평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_월평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "BT_주평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "BT_월평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월누계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_주평균_기소": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_월평균_기소": col = col + 1
        ' FT,JRT,BT,PRE 일계_기소 (123~126)
        wsDB.Cells(1, col).Value = "FT_일계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "JRT_일계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "BT_일계_기소": col = col + 1
        wsDB.Cells(1, col).Value = "PRE_일계_기소": col = col + 1

        With wsDB.Range(wsDB.Cells(1, 1), wsDB.Cells(1, 126))
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        ' wsDB.Columns.AutoFit  ' 열너비 자동맞춤 제거 (사용자 설정 열너비 유지)
    End Sub

    ' 제품표준사양서에서 제품명(G열 G9~마지막)과 일치하는 행의 AG값·AJ값 반환
    Private Sub 표준사양서_제품조회(ws표준 As Worksheet, 제품명 As String, ByRef outAG As String, ByRef outAJ As Double)
        Dim lastRow As Long, r As Long, vG As String, vAG As Variant, vAJ As Variant
        outAG = "": outAJ = 0
        If ws표준 Is Nothing Then Exit Sub
        lastRow = ws표준.Cells(ws표준.Rows.Count, 표준_G열).End(xlUp).Row
        If lastRow < 표준_데이터시작행 Then Exit Sub
        제품명 = Trim(제품명)
        If 제품명 = "" Then Exit Sub
        For r = 표준_데이터시작행 To lastRow
            vG = Trim(CStr(ws표준.Cells(r, 표준_G열).Value))
            If vG = 제품명 Then
                vAG = ws표준.Cells(r, 표준_AG열).Value
                outAG = Trim(CStr(vAG))
                vAJ = ws표준.Cells(r, 표준_AJ열).Value
                If IsNumeric(vAJ) And CDbl(vAJ) <> 0 Then outAJ = CDbl(vAJ) Else outAJ = 1
                Exit Sub
            End If
        Next r
    End Sub

    ' ============================================================
    ' [독립 실행용] DB 헤더 확인 매크로
    ' - Excel 매크로 목록(Alt+F8)에서 직접 실행 가능
    ' - 일일생산DB 시트에 col 16(AJ11), 17(AJ13) 헤더 확인
    ' ============================================================
    Public Sub DB헤더_확인_실행()
        Dim wsDB As Worksheet
        On Error Resume Next
        Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
        On Error GoTo 0
        If wsDB Is Nothing Then
            MsgBox "'일일생산DB' 시트를 찾을 수 없습니다." & vbCrLf & _
                   "먼저 일일생산일보_최적화DB_전송을 실행하여 DB 시트를 생성하세요.", vbCritical
            Exit Sub
        End If
        MsgBox "일일생산DB 헤더 확인" & vbCrLf & _
               "col 16: " & wsDB.Cells(1, 16).Value & vbCrLf & _
               "col 17: " & wsDB.Cells(1, 17).Value & vbCrLf & _
               "col 18: " & wsDB.Cells(1, 18).Value, vbInformation
    End Sub

    ' 생산일보 시트의 완제품 수식(X,AB,AD,AF,AH)이 자동 계산하므로 VBA 선반영 불필요
    Private Sub 표준사양서_계산_일일생산일보_반영(ws일보 As Worksheet)
        ' T07/T08 완제품 영역의 원지기준소요량·PLT 열은 시트 수식으로 자동 계산됨
        ' 더 이상 VBA에서 표준사양서를 조회하여 직접 셀에 쓸 필요 없음
    End Sub

    ' ============================================================
    ' 메인 프로시저: 일일생산일보 → 최적화DB 전송 (증분 업데이트)
    ' ★ 목업데이터와 동일: 수기 입력 데이터에 대해 1)시트 수식 계산 2)주·월 누적 반영 3)표준사양서 계산 반영
    '    진행 후, 그 결과를 일일생산DB 시트에만 저장. 일일생산DB는 부자재관리 시트 "데이터 업데이트"로 별도 반영.
    ' ============================================================
    Public Sub 일일생산일보_최적화DB_전송()
        Dim ws일보 As Worksheet, wsDB As Worksheet
        Dim 작성일 As Variant
        Dim 담당 As String, 이사 As String

        On Error Resume Next
        Set ws일보 = ThisWorkbook.Worksheets("생산일보")
        Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
        On Error GoTo 0

        If ws일보 Is Nothing Then
            MsgBox "'생산일보' 시트를 찾을 수 없습니다.", vbCritical
            Exit Sub
        End If

        If wsDB Is Nothing Then
            MsgBox "'일일생산DB' 시트가 없습니다." & vbCrLf & "자동으로 생성합니다.", vbInformation
            Set wsDB = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            wsDB.Name = "일일생산DB"
            Call 헤더_생성(wsDB)
        End If

        작성일 = ws일보.Range("BK4").Value
        If IsEmpty(작성일) Or Not IsDate(작성일) Then
            MsgBox "BK4 셀에 유효한 날짜가 없습니다.", vbCritical
            Exit Sub
        End If

        담당 = ""
        이사 = ""

        ' 0단계: DB 헤더 확인 (헤더_생성이 126열 전체 처리하므로 별도 보완 불필요)

        ' ★ 목업데이터 로직과 동일: 수기 입력 데이터에 대해 먼저 계산을 진행한 뒤 시트에 반영·DB 저장
        ' 1단계: 시트 수식 계산 (AS22, AS24, AI62 등 수식 영역이 수기 입력값 기준으로 계산되도록)
        Dim 원래Calculation As Long
        원래Calculation = Application.Calculation
        Application.Calculation = xlCalculationAutomatic
        ws일보.Calculate
        Application.Calculation = 원래Calculation

        ' 2단계: 주·월 누적 계산 후 일일생산일보 시트에 반영 (AJ4=작성일 기준, 일일생산DB 참조 가산)
        Call 일일생산일보_주월_계산후_시트반영(ws일보, wsDB, 작성일)

        ' 3단계: 제품표준사양서 기준 Y/AC(생산일보1), AW/AZ(생산일보2) 계산 후 일일생산일보에 선반영
        Call 표준사양서_계산_일일생산일보_반영(ws일보)

        ' 4단계: 계산·반영이 완료된 시트 기준으로 일일생산DB 시트에만 증분 저장 (일일생산DB는 별도 데이터 업데이트)
        Call 증분_업데이트(wsDB, ws일보, 작성일, 담당, 이사)

        ' 부자재관리 반영은 분리: 부자재관리 시트의 "데이터 업데이트" 버튼으로 일일생산DB 변경분만 증분 반영

        If Not g_무음모드 Then
            MsgBox "일일생산DB 저장이 완료되었습니다." & vbCrLf & vbCrLf & _
                   "· 일일생산일보 → 일일생산DB 시트만 업데이트되었습니다." & vbCrLf & _
                   "· 부자재관리에 반영하려면 [부자재관리] 시트에서 '데이터 업데이트' 버튼을 실행하세요." & vbCrLf & _
                   "  (변경·추가된 일일생산DB 내용만 빠르게 반영됩니다.)", vbInformation, "DB 저장"
        End If
    End Sub

    ' ============================================================
    ' 주/월 누적 재계산 (AJ4 날짜만 바꾼 뒤 누적만 갱신할 때 사용, DB 저장 없음)
    ' ※ 수식 계산 후 주·월 반영 (목업데이터와 동일). 일일생산DB 시트 참조로 해당 주차·월만 합산.
    ' ============================================================
    Public Sub 주월_누적_재계산()
        Dim ws일보 As Worksheet, wsDB As Worksheet
        Dim 작성일 As Variant
        Dim 원래Calculation As Long
        On Error Resume Next
        Set ws일보 = ThisWorkbook.Worksheets("생산일보")
        Set wsDB = ThisWorkbook.Worksheets("일일생산DB")
        On Error GoTo 0
        If ws일보 Is Nothing Or wsDB Is Nothing Then
            MsgBox "'생산일보' 또는 '일일생산DB' 시트가 없습니다.", vbCritical
            Exit Sub
        End If
        작성일 = ws일보.Range("BK4").Value
        If IsEmpty(작성일) Or Not IsDate(작성일) Then
            MsgBox "BK4 셀에 유효한 날짜를 입력한 뒤 실행하세요.", vbCritical
            Exit Sub
        End If
        원래Calculation = Application.Calculation
        Application.Calculation = xlCalculationAutomatic
        ws일보.Calculate
        Application.Calculation = 원래Calculation
        Call 일일생산일보_주월_계산후_시트반영(ws일보, wsDB, 작성일)
    End Sub

    ' ============================================================
    ' DB 저장 전: KT/JRT/WT/FT/BT 주차·월 누적 계산 후 일일생산일보 시트에 반영
' ※ 누적 = 일일생산DB 이전 생산기록 참조 + 작성일(오늘) 시트값 가산.
' 주간 = 월요일 기준 주차 (최종검수테스트·PA 검증과 동일 정의)
' 월 누적 = 해당 월 1일~말일 전부 합산.
    ' ============================================================
    ' 날짜를 Long 시리얼로 안전 변환 (오버플로우·잘못된 값 시 0 반환)
    Private Function 주월_SafeDateSerial(v As Variant) As Long
        On Error Resume Next
        Dim d As Double
        주월_SafeDateSerial = 0
        If IsEmpty(v) Then Exit Function
        d = CDbl(CDate(v))
        If Err.Number <> 0 Then Err.Clear: Exit Function
        If d < 1 Or d > 2958465 Then Exit Function
        If d < -2147483648# Or d > 2147483647# Then Exit Function
        주월_SafeDateSerial = CLng(Int(d))
        If Err.Number <> 0 Then 주월_SafeDateSerial = 0
        Err.Clear
    End Function

    ' 숫자 셀 안전 변환 (오버플로우·에러값·비수치 시 0 반환, 합산용 상한 적용)
    Private Function 주월_SafeDouble(v As Variant) As Double
        주월_SafeDouble = 0
        If IsEmpty(v) Then Exit Function
        If IsError(v) Then Exit Function
        On Error Resume Next
        Dim s As String: s = CStr(v)
        If Len(s) > 50 Then Exit Function
        Dim d As Double
        d = val(s)
        If Err.Number <> 0 Then Err.Clear: Exit Function
        d = CDbl(d)
        If Err.Number <> 0 Then Err.Clear: Exit Function
        If d > 1E+15 Or d < -1E+15 Then Exit Function
        주월_SafeDouble = d
        Err.Clear
    End Function

    ' 주차 경계: 1주=월1일~첫일요일 전일(예: 1월 1~3일), 2주=첫일요일~+6일(4~10일), 둘째주부터 초기화
    Private Function 주월_주차번호(d As Date) As Long
        Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
        Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
        Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
        If d < 첫일요일 Then 주월_주차번호 = 1: Exit Function
        주월_주차번호 = 2 + CLng((d - 첫일요일) \ 7)
    End Function
    Private Function 주월_주차첫날(d As Date) As Date
        Dim w As Long: w = 주월_주차번호(d)
        Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
        Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
        Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
        If w = 1 Then 주월_주차첫날 = 월1일: Exit Function
        주월_주차첫날 = 첫일요일 + CLng(w - 2) * 7
    End Function
    Private Function 주월_주차마지막일(d As Date) As Date
        Dim w As Long: w = 주월_주차번호(d)
        Dim 월1일 As Date: 월1일 = DateSerial(Year(d), Month(d), 1)
        Dim 요일 As Long: 요일 = Weekday(월1일, vbMonday)
        Dim 첫일요일 As Date: 첫일요일 = 월1일 + (7 - 요일)
        Dim 말일 As Date: 말일 = DateSerial(Year(d), Month(d) + 1, 0)
        If w = 1 Then 주월_주차마지막일 = 첫일요일 - 1: Exit Function
        Dim 끝 As Date: 끝 = 첫일요일 + CLng(w - 2) * 7 + 6
        If 끝 > 말일 Then 끝 = 말일
        주월_주차마지막일 = 끝
    End Function

    ' T10 원자재수율: DB 이력 기반으로 주간/월간 누계·평균을 계산하여 시트에 반영
    ' ★ 값입력 영역만 쓰기 (나머지 T10 셀은 수식, 절대 덮어쓰기 금지):
    '   좌측 G88:K91, G93:K96, G98:K101 / 우측 AP88:AT91, AP93:AT96, AP98:AT101
    ' 주차 = 월요일 기준 (최종검수테스트와 동일). 월 = 해당 월 전체.
    ' ★ 주일수/월일수 = 실제 생산(원투>0) 있는 날만 카운트 (최종검수테스트 PA_합산과 동일)
    ' 설비별: KT(row87) 좌측, FT(row87) 우측, JRT(row92) 좌측, BT(row97) 좌측, PRE(row92) 우측
    ' Public: 역변환 모듈에서 동일 로직 사용 시 호출
    Public Sub 일일생산일보_주월_계산후_시트반영(ws일보 As Worksheet, wsDB As Worksheet, 작성일 As Variant)
        If ws일보 Is Nothing Then Exit Sub
        If IsEmpty(작성일) Or Not IsDate(작성일) Then Exit Sub

        Dim d작성 As Date: d작성 = CDate(작성일)
        Dim 연 As Long, 월 As Long
        연 = Year(d작성): 월 = Month(d작성)

        Dim 주첫날 As Date: 주첫날 = 주월_주차첫날(d작성)
        Dim 주마지막 As Date: 주마지막 = 주월_주차마지막일(d작성)

        ' 5개 설비: 0=KT, 1=FT, 2=JRT, 3=BT, 4=PRE
        ' KT=좌측Row87, FT=우측Row87, JRT=좌측Row92, BT=좌측Row97, PRE=우측Row92
        Dim DB원열 As Variant: DB원열 = Array(28, 40, 43, 46, 49)   ' KT/FT/JRT/BT/PRE 일계_원투 (스키마 +2 반영)
        Dim DB양열 As Variant: DB양열 = Array(29, 41, 44, 47, 50)   ' KT/FT/JRT/BT/PRE 일계_양품 (스키마 +2 반영)
        Dim DB기소열 As Variant: DB기소열 = Array(31, 123, 124, 125, 126)  ' KT,FT,JRT,BT,PRE 일계_기소
        Dim 시트일계행 As Variant: 시트일계행 = Array(87, 87, 92, 97, 92)
        Dim 시트원열 As Variant: 시트원열 = Array(7, 42, 7, 7, 42)
        Dim 시트양열 As Variant: 시트양열 = Array(12, 47, 12, 12, 47)
        Dim 시트수열 As Variant: 시트수열 = Array(15, 50, 15, 15, 50)   ' O(15) 원지수율 left, AX(50) right
        Dim 시트기소열 As Variant: 시트기소열 = Array(11, 46, 11, 11, 46)  ' K(11) 기준소요량 left, AT(46) right
        Dim 시트LOSS열 As Variant: 시트LOSS열 = Array(25, 60, 25, 25, 60)  ' Y(25) left, BH(60) right

        Dim 주원(0 To 4) As Double, 주양(0 To 4) As Double, 주기소(0 To 4) As Double
        Dim 월원(0 To 4) As Double, 월양(0 To 4) As Double, 월기소(0 To 4) As Double
        Dim 주일수 As Long, 월일수 As Long
        주일수 = 0: 월일수 = 0

        ' DB에서 같은 월의 PA 이력 읽기 (오늘 이전만, 124열까지 기소 포함)
        ' ★ 최종검수테스트와 동일: 생산일보구분="생산일보", 유형="ProductionAnalysis", 순번=1(T10만, 비가동 제외)
        If Not wsDB Is Nothing Then
            Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 2 Then
                Dim arrPA As Variant
                arrPA = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, 126)).Value
                Dim r As Long
                For r = 1 To UBound(arrPA, 1)
                    If InStr(1, Trim(CStr(arrPA(r, 3))), "생산일보", vbTextCompare) < 1 Then GoTo NxtPA
                    If Trim(CStr(arrPA(r, 4))) <> "ProductionAnalysis" And Trim(CStr(arrPA(r, 4))) <> "PA" Then GoTo NxtPA
                    If Not IsNumeric(arrPA(r, 5)) Or CLng(arrPA(r, 5)) <> 1 Then GoTo NxtPA
                    If Not IsDate(arrPA(r, 1)) Then GoTo NxtPA
                    Dim dd As Date: dd = CDate(arrPA(r, 1))
                    If Year(dd) <> 연 Or Month(dd) <> 월 Then GoTo NxtPA
                    If dd >= d작성 Then GoTo NxtPA

                    Dim eq As Long, v1 As Double, v2 As Double, v3 As Double
                    Dim 주포함 As Boolean: 주포함 = (dd >= 주첫날 And dd <= 주마지막)
                    Dim 행원투합 As Double: 행원투합 = 0

                    For eq = 0 To 4
                        v1 = 주월_SafeDouble(arrPA(r, CLng(DB원열(eq))))
                        v2 = 주월_SafeDouble(arrPA(r, CLng(DB양열(eq))))
                        v3 = 0
                        If CLng(DB기소열(eq)) <= UBound(arrPA, 2) Then v3 = 주월_SafeDouble(arrPA(r, CLng(DB기소열(eq))))
                        행원투합 = 행원투합 + v1
                        월원(eq) = 월원(eq) + v1
                        월양(eq) = 월양(eq) + v2
                        월기소(eq) = 월기소(eq) + v3
                        If 주포함 Then
                            주원(eq) = 주원(eq) + v1
                            주양(eq) = 주양(eq) + v2
                            주기소(eq) = 주기소(eq) + v3
                        End If
                    Next eq
                    ' ★ 원투 합계 0인 행(데이터 없는 날)은 주일수/월일수에서 제외 (역변환 시 평균 오류 방지)
                    If 행원투합 > 0.0000000001 Then
                        월일수 = 월일수 + 1
                        If 주포함 Then 주일수 = 주일수 + 1
                    End If
NxtPA:
                Next r
            End If
        End If

        ' 오늘 일계 추가: 시트 우선, 시트 0/빈값이면 DB 해당일 PA 행에서 보완 (FT/JRT/BT/PRE 누락 방지)
        Dim c기소 As Long, v원 As Double, v양 As Double, v기 As Double, 오늘원투합 As Double
        Dim 오늘DB원(0 To 4) As Double, 오늘DB양(0 To 4) As Double, 오늘DB기(0 To 4) As Double
        Dim rPA As Long, foundPA As Boolean
        For eq = 0 To 4: 오늘DB원(eq) = 0: 오늘DB양(eq) = 0: 오늘DB기(eq) = 0: Next eq
        foundPA = False
        If Not wsDB Is Nothing Then
            Dim lastR2 As Long: lastR2 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
            Dim 마지막날짜 As Variant: 마지막날짜 = Empty
            For rPA = 2 To lastR2
                If Not IsEmpty(wsDB.Cells(rPA, 1).Value) Then 마지막날짜 = wsDB.Cells(rPA, 1).Value
                If Not IsDate(마지막날짜) Then GoTo NextPArow
                If Year(CDate(마지막날짜)) <> 연 Or Month(CDate(마지막날짜)) <> 월 Then GoTo NextPArow
                If DateDiff("d", CDate(마지막날짜), d작성) <> 0 Then GoTo NextPArow
                If InStr(1, Trim(CStr(wsDB.Cells(rPA, 3).Value)), "생산일보", vbTextCompare) < 1 Then GoTo NextPArow
                If Trim(CStr(wsDB.Cells(rPA, 4).Value)) <> "ProductionAnalysis" And Trim(CStr(wsDB.Cells(rPA, 4).Value)) <> "PA" Then GoTo NextPArow
                If Not IsNumeric(wsDB.Cells(rPA, 5).Value) Or CLng(wsDB.Cells(rPA, 5).Value) <> 1 Then GoTo NextPArow
                foundPA = True
                For eq = 0 To 4
                    오늘DB원(eq) = 주월_SafeDouble(wsDB.Cells(rPA, CLng(DB원열(eq))).Value)
                    오늘DB양(eq) = 주월_SafeDouble(wsDB.Cells(rPA, CLng(DB양열(eq))).Value)
                    If CLng(DB기소열(eq)) <= 126 Then 오늘DB기(eq) = 주월_SafeDouble(wsDB.Cells(rPA, CLng(DB기소열(eq))).Value)
                Next eq
                Exit For
NextPArow:
            Next rPA
        End If
        오늘원투합 = 0
        For eq = 0 To 4
            c기소 = CLng(시트기소열(eq))
            v원 = 주월_SafeDouble(ws일보.Cells(CLng(시트일계행(eq)), CLng(시트원열(eq))).Value)
            v양 = 주월_SafeDouble(ws일보.Cells(CLng(시트일계행(eq)), CLng(시트양열(eq))).Value)
            v기 = 주월_SafeDouble(ws일보.Cells(CLng(시트일계행(eq)), c기소).Value)
            If v원 <= 0.0000000001 And foundPA And 오늘DB원(eq) > 0.0000000001 Then v원 = 오늘DB원(eq)
            If v양 <= 0.0000000001 And foundPA And 오늘DB양(eq) > 0.0000000001 Then v양 = 오늘DB양(eq)
            If v기 <= 0.0000000001 And foundPA And 오늘DB기(eq) > 0.0000000001 Then v기 = 오늘DB기(eq)
            오늘원투합 = 오늘원투합 + v원
            주원(eq) = 주원(eq) + v원: 주양(eq) = 주양(eq) + v양: 주기소(eq) = 주기소(eq) + v기
            월원(eq) = 월원(eq) + v원: 월양(eq) = 월양(eq) + v양: 월기소(eq) = 월기소(eq) + v기
        Next eq
        If 오늘원투합 > 0.0000000001 Then 주일수 = 주일수 + 1: 월일수 = 월일수 + 1

        ' 시트에 반영: 좌측 G(7)=원투 L(12)=양품 O(15)=수율 Y(25)=LOSS K(11)=기소
        '            우측 AP(42)=원투 AU(47)=양품 AX(50)=수율 BH(60)=LOSS AT(46)=기소
        Dim 기행 As Long, rw As Long
        Dim v원투 As Double, v양품 As Double, v수율 As Double, v기소 As Double, vLOSS As Double
        Dim 쓰기좌 As Boolean, 쓰기우 As Boolean
        On Error Resume Next
        For eq = 0 To 4
            기행 = CLng(시트일계행(eq))
            쓰기좌 = (CLng(시트원열(eq)) = 7)
            쓰기우 = (CLng(시트원열(eq)) = 42)
            For rw = 1 To 4
                If rw = 1 Then
                    v원투 = 주원(eq): v양품 = 주양(eq)
                    v수율 = 0: If 주원(eq) > 0.0000000001 Then v수율 = Round(주양(eq) / 주원(eq) * 100, 1)
                    v기소 = Round(주기소(eq), 1)
                    vLOSS = 0: If 주원(eq) > 0.0000000001 Then vLOSS = Round((주원(eq) - 주양(eq)) / 주원(eq) * 100, 1)
                ElseIf rw = 2 Then
                    v원투 = 월원(eq): v양품 = 월양(eq)
                    v수율 = 0: If 월원(eq) > 0.0000000001 Then v수율 = Round(월양(eq) / 월원(eq) * 100, 1)
                    v기소 = Round(월기소(eq), 1)
                    vLOSS = 0: If 월원(eq) > 0.0000000001 Then vLOSS = Round((월원(eq) - 월양(eq)) / 월원(eq) * 100, 1)
                ElseIf rw = 3 Then
                    v원투 = 0: v양품 = 0: v수율 = 0: v기소 = 0: vLOSS = 0
                    If 주일수 > 0 Then v원투 = Round(주원(eq) / 주일수, 1): v양품 = Round(주양(eq) / 주일수, 1): v기소 = Round(주기소(eq) / 주일수, 1)
                    If 주원(eq) > 0.0000000001 Then v수율 = Round(주양(eq) / 주원(eq) * 100, 1): vLOSS = Round((주원(eq) - 주양(eq)) / 주원(eq) * 100, 1)
                Else
                    v원투 = 0: v양품 = 0: v수율 = 0: v기소 = 0: vLOSS = 0
                    If 월일수 > 0 Then v원투 = Round(월원(eq) / 월일수, 1): v양품 = Round(월양(eq) / 월일수, 1): v기소 = Round(월기소(eq) / 월일수, 1)
                    If 월원(eq) > 0.0000000001 Then v수율 = Round(월양(eq) / 월원(eq) * 100, 1): vLOSS = Round((월원(eq) - 월양(eq)) / 월원(eq) * 100, 1)
                End If
                ' 계산값이 0일 때 PA가 입력한 값 보존 (덮어쓰기 방지)
                Dim 기쓸 As Double, 양쓸 As Double, 수쓸 As Double, L쓸 As Double, 소쓸 As Double
                기쓸 = v원투: 양쓸 = v양품: 수쓸 = v수율: L쓸 = vLOSS: 소쓸 = v기소
                If 쓰기좌 Then
                    If 기쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 7).Value) Then 기쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 7).Value)
                    If 양쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 12).Value) Then 양쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 12).Value)
                    If 소쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 11).Value) Then 소쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 11).Value)
                    If 수쓸 = 0 And 기쓸 > 0.0000000001 Then 수쓸 = Round(양쓸 / 기쓸 * 100, 1): L쓸 = Round((기쓸 - 양쓸) / 기쓸 * 100, 1)
                    ws일보.Cells(기행 + rw, 7).Value = 기쓸: ws일보.Cells(기행 + rw, 12).Value = 양쓸
                    ws일보.Cells(기행 + rw, 15).Value = 수쓸: ws일보.Cells(기행 + rw, 25).Value = L쓸
                    ws일보.Cells(기행 + rw, 11).Value = 소쓸
                End If
                If 쓰기우 Then
                    If 기쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 42).Value) Then 기쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 42).Value)
                    If 양쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 47).Value) Then 양쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 47).Value)
                    If 소쓸 = 0 And IsNumeric(ws일보.Cells(기행 + rw, 46).Value) Then 소쓸 = 주월_SafeDouble(ws일보.Cells(기행 + rw, 46).Value)
                    If 수쓸 = 0 And 기쓸 > 0.0000000001 Then 수쓸 = Round(양쓸 / 기쓸 * 100, 1): L쓸 = Round((기쓸 - 양쓸) / 기쓸 * 100, 1)
                    ws일보.Cells(기행 + rw, 42).Value = 기쓸: ws일보.Cells(기행 + rw, 47).Value = 양쓸
                    ws일보.Cells(기행 + rw, 50).Value = 수쓸: ws일보.Cells(기행 + rw, 60).Value = L쓸
                    ws일보.Cells(기행 + rw, 46).Value = 소쓸
                End If
            Next rw
        Next eq
        Err.Clear
    End Sub

    ' ============================================================
    ' 일일생산DB PA 행: DB 적재된 일계를 기반으로 주/월 누계·평균 실제 계산
    ' ※ 이관(Module21/22) 등에서 일계만 수집 후 호출. 가져온 값 표시가 아닌 DB 누적 기반 계산.
    ' ============================================================
    Public Sub 일일생산DB_PA누계평균_행반영(wsDB As Worksheet, fromRow As Long, toRow As Long)
        If wsDB Is Nothing Then Exit Sub
        If fromRow > toRow Then Exit Sub
        Dim lastRow As Long: lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then Exit Sub
        Dim arrPA As Variant
        arrPA = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, 126)).Value2
        Dim DB원열 As Variant: DB원열 = Array(28, 40, 43, 46, 49)
        Dim DB양열 As Variant: DB양열 = Array(29, 41, 44, 47, 50)
        Dim DB기소열 As Variant: DB기소열 = Array(31, 123, 124, 125, 126)
        Dim r As Long, eq As Long, ri As Long
        Dim d작성 As Date, 연 As Long, 월 As Long
        Dim 주첫날 As Date, 주마지막 As Date
        Dim 주원(0 To 4) As Double, 주양(0 To 4) As Double, 주기소(0 To 4) As Double
        Dim 월원(0 To 4) As Double, 월양(0 To 4) As Double, 월기소(0 To 4) As Double
        Dim 주일수 As Long, 월일수 As Long
        Dim dd As Date, 주포함 As Boolean, 행원투합 As Double
        Dim v1 As Double, v2 As Double, v3 As Double
        Dim v수율 As Double
        For r = fromRow To toRow
            If r < 2 Or r > lastRow Then GoTo NextR
            If Not IsDate(wsDB.Cells(r, 1).Value) Then GoTo NextR
            If InStr(1, Trim(CStr(wsDB.Cells(r, 3).Value)), "생산일보", vbTextCompare) < 1 Then GoTo NextR
            If Trim(CStr(wsDB.Cells(r, 4).Value)) <> "ProductionAnalysis" And Trim(CStr(wsDB.Cells(r, 4).Value)) <> "PA" Then GoTo NextR
            If Not IsNumeric(wsDB.Cells(r, 5).Value) Or CLng(wsDB.Cells(r, 5).Value) <> 1 Then GoTo NextR
            d작성 = CDate(wsDB.Cells(r, 1).Value)
            연 = Year(d작성): 월 = Month(d작성)
            주첫날 = 주월_주차첫날(d작성): 주마지막 = 주월_주차마지막일(d작성)
            For eq = 0 To 4: 주원(eq) = 0: 주양(eq) = 0: 주기소(eq) = 0: 월원(eq) = 0: 월양(eq) = 0: 월기소(eq) = 0: Next eq
            주일수 = 0: 월일수 = 0
            For ri = 1 To UBound(arrPA, 1)
                If InStr(1, Trim(CStr(arrPA(ri, 3))), "생산일보", vbTextCompare) < 1 Then GoTo NxtRi
                If Trim(CStr(arrPA(ri, 4))) <> "ProductionAnalysis" And Trim(CStr(arrPA(ri, 4))) <> "PA" Then GoTo NxtRi
                If Not IsNumeric(arrPA(ri, 5)) Or CLng(arrPA(ri, 5)) <> 1 Then GoTo NxtRi
                If Not IsDate(arrPA(ri, 1)) Then GoTo NxtRi
                dd = CDate(arrPA(ri, 1))
                If Year(dd) <> 연 Or Month(dd) <> 월 Or dd > d작성 Then GoTo NxtRi
                주포함 = (dd >= 주첫날 And dd <= 주마지막)
                행원투합 = 0
                For eq = 0 To 4
                    v1 = 주월_SafeDouble(arrPA(ri, CLng(DB원열(eq))))
                    v2 = 주월_SafeDouble(arrPA(ri, CLng(DB양열(eq))))
                    v3 = 0
                    If CLng(DB기소열(eq)) <= UBound(arrPA, 2) Then v3 = 주월_SafeDouble(arrPA(ri, CLng(DB기소열(eq))))
                    행원투합 = 행원투합 + v1
                    월원(eq) = 월원(eq) + v1: 월양(eq) = 월양(eq) + v2: 월기소(eq) = 월기소(eq) + v3
                    If 주포함 Then 주원(eq) = 주원(eq) + v1: 주양(eq) = 주양(eq) + v2: 주기소(eq) = 주기소(eq) + v3
                Next eq
                If 행원투합 > 0.0000000001 Then 월일수 = 월일수 + 1: If 주포함 Then 주일수 = 주일수 + 1
NxtRi:
            Next ri
            ' DB 행 반영: 34~122
            For eq = 0 To 4
                v수율 = 0
                If 주원(eq) > 0.0000000001 Then v수율 = Round(주양(eq) / 주원(eq) * 100, 1)
                Select Case eq
                    Case 0
                        wsDB.Cells(r, 34).Value = Round(주원(0), 1): wsDB.Cells(r, 35).Value = Round(주양(0), 1): wsDB.Cells(r, 36).Value = v수율
                        wsDB.Cells(r, 37).Value = Round(월원(0), 1): wsDB.Cells(r, 38).Value = Round(월양(0), 1): wsDB.Cells(r, 39).Value = IIf(월원(0) > 0.0000000001, Round(월양(0) / 월원(0) * 100, 1), 0)
                        wsDB.Cells(r, 73).Value = IIf(주일수 > 0, Round(주원(0) / 주일수, 1), 0): wsDB.Cells(r, 74).Value = IIf(주일수 > 0, Round(주양(0) / 주일수, 1), 0): wsDB.Cells(r, 75).Value = v수율
                        wsDB.Cells(r, 76).Value = IIf(월일수 > 0, Round(월원(0) / 월일수, 1), 0): wsDB.Cells(r, 77).Value = IIf(월일수 > 0, Round(월양(0) / 월일수, 1), 0): wsDB.Cells(r, 78).Value = v수율
                        wsDB.Cells(r, 103).Value = Round(주기소(0), 1): wsDB.Cells(r, 104).Value = Round(월기소(0), 1)
                        wsDB.Cells(r, 105).Value = IIf(주일수 > 0, Round(주기소(0) / 주일수, 1), 0): wsDB.Cells(r, 106).Value = IIf(월일수 > 0, Round(월기소(0) / 월일수, 1), 0)
                    Case 1
                        wsDB.Cells(r, 54).Value = Round(주원(1), 1): wsDB.Cells(r, 55).Value = Round(주양(1), 1): wsDB.Cells(r, 56).Value = Round(월원(1), 1): wsDB.Cells(r, 57).Value = Round(월양(1), 1)
                        wsDB.Cells(r, 79).Value = IIf(주일수 > 0, Round(주원(1) / 주일수, 1), 0): wsDB.Cells(r, 80).Value = IIf(주일수 > 0, Round(주양(1) / 주일수, 1), 0): wsDB.Cells(r, 81).Value = IIf(주원(1) > 0.0000000001, Round(주양(1) / 주원(1) * 100, 1), 0)
                        wsDB.Cells(r, 82).Value = IIf(월일수 > 0, Round(월원(1) / 월일수, 1), 0): wsDB.Cells(r, 83).Value = IIf(월일수 > 0, Round(월양(1) / 월일수, 1), 0): wsDB.Cells(r, 84).Value = IIf(월원(1) > 0.0000000001, Round(월양(1) / 월원(1) * 100, 1), 0)
                        wsDB.Cells(r, 107).Value = Round(주기소(1), 1): wsDB.Cells(r, 108).Value = Round(월기소(1), 1): wsDB.Cells(r, 109).Value = IIf(주일수 > 0, Round(주기소(1) / 주일수, 1), 0): wsDB.Cells(r, 110).Value = IIf(월일수 > 0, Round(월기소(1) / 월일수, 1), 0)
                    Case 2
                        wsDB.Cells(r, 58).Value = Round(주원(2), 1): wsDB.Cells(r, 59).Value = Round(주양(2), 1): wsDB.Cells(r, 60).Value = Round(월원(2), 1): wsDB.Cells(r, 63).Value = Round(월양(2), 1)
                        wsDB.Cells(r, 85).Value = IIf(주일수 > 0, Round(주원(2) / 주일수, 1), 0): wsDB.Cells(r, 86).Value = IIf(주일수 > 0, Round(주양(2) / 주일수, 1), 0): wsDB.Cells(r, 87).Value = IIf(주원(2) > 0.0000000001, Round(주양(2) / 주원(2) * 100, 1), 0)
                        wsDB.Cells(r, 88).Value = IIf(월일수 > 0, Round(월원(2) / 월일수, 1), 0): wsDB.Cells(r, 89).Value = IIf(월일수 > 0, Round(월양(2) / 월일수, 1), 0): wsDB.Cells(r, 90).Value = IIf(월원(2) > 0.0000000001, Round(월양(2) / 월원(2) * 100, 1), 0)
                        wsDB.Cells(r, 111).Value = Round(주기소(2), 1): wsDB.Cells(r, 112).Value = Round(월기소(2), 1): wsDB.Cells(r, 113).Value = IIf(주일수 > 0, Round(주기소(2) / 주일수, 1), 0): wsDB.Cells(r, 114).Value = IIf(월일수 > 0, Round(월기소(2) / 월일수, 1), 0)
                    Case 3
                        wsDB.Cells(r, 64).Value = Round(주원(3), 1): wsDB.Cells(r, 65).Value = Round(주양(3), 1): wsDB.Cells(r, 66).Value = Round(월원(3), 1): wsDB.Cells(r, 67).Value = Round(월양(3), 1)
                        wsDB.Cells(r, 91).Value = IIf(주일수 > 0, Round(주원(3) / 주일수, 1), 0): wsDB.Cells(r, 92).Value = IIf(주일수 > 0, Round(주양(3) / 주일수, 1), 0): wsDB.Cells(r, 93).Value = IIf(주원(3) > 0.0000000001, Round(주양(3) / 주원(3) * 100, 1), 0)
                        wsDB.Cells(r, 94).Value = IIf(월일수 > 0, Round(월원(3) / 월일수, 1), 0): wsDB.Cells(r, 95).Value = IIf(월일수 > 0, Round(월양(3) / 월일수, 1), 0): wsDB.Cells(r, 96).Value = IIf(월원(3) > 0.0000000001, Round(월양(3) / 월원(3) * 100, 1), 0)
                        wsDB.Cells(r, 115).Value = Round(주기소(3), 1): wsDB.Cells(r, 116).Value = Round(월기소(3), 1): wsDB.Cells(r, 117).Value = IIf(주일수 > 0, Round(주기소(3) / 주일수, 1), 0): wsDB.Cells(r, 118).Value = IIf(월일수 > 0, Round(월기소(3) / 월일수, 1), 0)
                    Case 4
                        wsDB.Cells(r, 68).Value = Round(주원(4), 1): wsDB.Cells(r, 69).Value = Round(주양(4), 1): wsDB.Cells(r, 70).Value = Round(월원(4), 1): wsDB.Cells(r, 71).Value = Round(월양(4), 1)
                        wsDB.Cells(r, 97).Value = IIf(주일수 > 0, Round(주원(4) / 주일수, 1), 0): wsDB.Cells(r, 98).Value = IIf(주일수 > 0, Round(주양(4) / 주일수, 1), 0): wsDB.Cells(r, 99).Value = IIf(주원(4) > 0.0000000001, Round(주양(4) / 주원(4) * 100, 1), 0)
                        wsDB.Cells(r, 100).Value = IIf(월일수 > 0, Round(월원(4) / 월일수, 1), 0): wsDB.Cells(r, 101).Value = IIf(월일수 > 0, Round(월양(4) / 월일수, 1), 0): wsDB.Cells(r, 102).Value = IIf(월원(4) > 0.0000000001, Round(월양(4) / 월원(4) * 100, 1), 0)
                        wsDB.Cells(r, 119).Value = Round(주기소(4), 1): wsDB.Cells(r, 120).Value = Round(월기소(4), 1): wsDB.Cells(r, 121).Value = IIf(주일수 > 0, Round(주기소(4) / 주일수, 1), 0): wsDB.Cells(r, 122).Value = IIf(월일수 > 0, Round(월기소(4) / 월일수, 1), 0)
                End Select
            Next eq
NextR:
        Next r
    End Sub

    ' ============================================================
    ' 증분 업데이트 (추가/변경/삭제만 처리)
    ' ============================================================
    Private Sub 증분_업데이트(wsDB As Worksheet, ws일보 As Worksheet, 작성일 As Variant, 담당 As String, 이사 As String)
        Dim dictOld As Object, dictNew As Object
        Set dictOld = CreateObject("Scripting.Dictionary")
        Set dictNew = CreateObject("Scripting.Dictionary")
        
        ' ★★ 프로파일러: 각 단계 소요시간 측정 ★★
        Dim tStart As Double, tStep As Double, 프로파일 As String
        tStart = Timer: tStep = Timer
        프로파일 = "=== 속도 프로파일 ===" & vbCrLf

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        프로파일 = 프로파일 & "[DB] App설정: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer

        ' 1. 기존 데이터 로드 (작성일이 같은 것만)
        Call 기존데이터_로드(wsDB, 작성일, dictOld)
        프로파일 = 프로파일 & "[DB] 기존데이터_로드: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer

        ' 2. 새 데이터 수집
        Dim ID카운터 As Long
        Dim 마지막ID값 As Variant

        마지막ID값 = wsDB.Cells(wsDB.Rows.Count, 2).End(xlUp).Value

        If IsEmpty(마지막ID값) Then
            ID카운터 = 0
        ElseIf VarType(마지막ID값) = vbString And Trim(CStr(마지막ID값)) = "ID" Then
            ID카운터 = 0
        ElseIf IsNumeric(마지막ID값) Then
            ID카운터 = CLng(마지막ID값)
        Else
            ID카운터 = 0
        End If

        Call 새데이터_수집_v3(dictNew, ID카운터, 작성일, ws일보, 담당, 이사)
        프로파일 = 프로파일 & "[DB] 새데이터_수집: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer

        ' 3. 비교 및 처리
        Dim 추가수 As Long, 변경수 As Long, 삭제수 As Long
        추가수 = 0: 변경수 = 0: 삭제수 = 0
        
        ' ★ 신규 레코드를 Collection에 수집 (나중에 배치 쓰기)
        Dim 신규목록 As Collection: Set 신규목록 = New Collection

        ' 3-1. 새 데이터 처리 (추가 또는 업데이트)
        Dim key As Variant
        For Each key In dictNew.keys
            If dictOld.exists(key) Then
                If 레코드_비교(dictOld(key), dictNew(key)) Then
                    Call 레코드_업데이트(wsDB, dictOld(key), dictNew(key))
                    변경수 = 변경수 + 1
                End If
            Else
                ' ★ 즉시 쓰지 않고 수집만
                신규목록.Add dictNew(key)
                추가수 = 추가수 + 1
            End If
        Next key
        
        ' ★ 3-1b. 신규 레코드 일괄 배치 쓰기 (1회 Range.Value)
        If 추가수 > 0 Then
            Dim 새행시작 As Long
            새행시작 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
            Dim arrBatch() As Variant
            ReDim arrBatch(1 To 추가수, 1 To 126)
            Dim bi As Long, col As Long
            For bi = 1 To 추가수
                Dim rec As Object: Set rec = 신규목록(bi)
                For col = 1 To 126
                    If rec.exists("F" & col) Then arrBatch(bi, col) = rec("F" & col) Else arrBatch(bi, col) = Empty
                Next col
            Next bi
            wsDB.Range(wsDB.Cells(새행시작, 1), wsDB.Cells(새행시작 + 추가수 - 1, 126)).Value = arrBatch
        End If

        프로파일 = 프로파일 & "[DB] 비교+추가+변경: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer
        
        ' 3-2. 삭제된 레코드 처리 (Union으로 일괄 삭제)
        Dim 삭제행목록 As Object
        Set 삭제행목록 = CreateObject("Scripting.Dictionary")
        For Each key In dictOld.keys
            If Not dictNew.exists(key) Then
                Set 삭제행목록(CStr(dictOld(key)("행번호"))) = dictOld(key)
            End If
        Next key
        If 삭제행목록.Count > 0 Then
            Dim 행번호Arr() As Long
            ReDim 행번호Arr(삭제행목록.Count - 1)
            Dim idx As Long: idx = 0
            For Each key In 삭제행목록.keys
                행번호Arr(idx) = CLng(key): idx = idx + 1
            Next key
            Dim i As Long, j As Long, temp As Long
            For i = LBound(행번호Arr) To UBound(행번호Arr) - 1
                For j = i + 1 To UBound(행번호Arr)
                    If 행번호Arr(i) < 행번호Arr(j) Then
                        temp = 행번호Arr(i): 행번호Arr(i) = 행번호Arr(j): 행번호Arr(j) = temp
                    End If
                Next j
            Next i
            Dim delRng As Range
            Set delRng = wsDB.Rows(행번호Arr(LBound(행번호Arr)))
            For i = LBound(행번호Arr) + 1 To UBound(행번호Arr)
                Set delRng = Union(delRng, wsDB.Rows(행번호Arr(i)))
            Next i
            delRng.Delete
            삭제수 = 삭제행목록.Count
        End If

        프로파일 = 프로파일 & "[DB] 삭제처리: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer
        
        ' 부자재관리 반영은 분리: 부자재관리 시트 "데이터 업데이트" 버튼으로만 실행 (DB저장 시 병목 제거)
        
        Application.Calculation = xlCalculationAutomatic
        프로파일 = 프로파일 & "[복원] Calculation=Auto: " & Format(Timer - tStep, "0.000") & "s" & vbCrLf: tStep = Timer
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        프로파일 = 프로파일 & "▶ 총 소요: " & Format(Timer - tStart, "0.000") & "s" & vbCrLf & _
                "DB: 추가=" & 추가수 & " 변경=" & 변경수 & " 삭제=" & 삭제수
        
        If Not g_무음모드 Then MsgBox 프로파일, vbInformation, "속도 프로파일 결과"
    End Sub

    ' StatusBar 리셋용
    Public Sub StatusBar_초기화()
        Application.StatusBar = False
    End Sub

    ' ============================================================
    ' 기존 데이터 로드 (배열 일괄 읽기로 속도 개선)
    ' ============================================================
    Private Sub 기존데이터_로드(wsDB As Worksheet, 작성일 As Variant, dict As Object)
        Dim lastRow As Long, r As Long
        Dim dataArr As Variant
        lastRow = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then Exit Sub

        ' 전체 데이터를 배열로 일괄 읽기 (시트 접근 1회)
        dataArr = wsDB.Range(wsDB.Cells(2, 1), wsDB.Cells(lastRow, 126)).Value

        For r = 1 To UBound(dataArr, 1)
            If dataArr(r, 1) = 작성일 Then
                Dim key As String
                key = 레코드키_생성_배열(dataArr, r)

                Dim 레코드Info As Object
                Set 레코드Info = CreateObject("Scripting.Dictionary")
                레코드Info("행번호") = r + 1
                레코드Info("ID") = dataArr(r, 2)
                Dim col As Long
                For col = 1 To 126
                    레코드Info("F" & col) = dataArr(r, col)
                Next col
                Set dict(key) = 레코드Info
            End If
        Next r
    End Sub

    ' 레코드 키 생성 - 배열용 (dataArr의 1-based 행 인덱스)
    Private Function 레코드키_생성_배열(dataArr As Variant, r As Long) As String
        Dim 생산일보구분 As String, 레코드유형 As String
        Dim 원지종별 As String, ROLL_ID As String, KG As String
        Dim 인원구분 As String, 작업장 As String
        Dim 구분 As String, 품명 As String, 가공라인 As String, 가공규격 As String, 수량 As String

        생산일보구분 = CStr(dataArr(r, 3))
        레코드유형 = CStr(dataArr(r, 4))

        Select Case 레코드유형
            Case "원지"
                원지종별 = CStr(dataArr(r, 6))
                ROLL_ID = Replace(Trim(CStr(dataArr(r, 7))), "|", "_")
                KG = Replace(CStr(dataArr(r, 8)), "|", "_")
                레코드키_생성_배열 = 생산일보구분 & "|원지|" & 원지종별 & "|" & ROLL_ID & "|" & KG
            Case "완제품"
                구분 = Replace(Trim(CStr(dataArr(r, 9))), "|", "_")
                품명 = Replace(Trim(CStr(dataArr(r, 10))), "|", "_")
                가공라인 = Replace(Trim(CStr(dataArr(r, 11))), "|", "_")
                가공규격 = Replace(Trim(CStr(dataArr(r, 12))), "|", "_")
                수량 = Replace(CStr(dataArr(r, 13)), "|", "_")
                레코드키_생성_배열 = 생산일보구분 & "|완제품|" & 구분 & "|" & 품명 & "|" & 가공라인 & "|" & 가공규격 & "|" & 수량
            Case "투입인원"
                작업장 = Trim(CStr(dataArr(r, 19)))
                인원구분 = CStr(dataArr(r, 20))
                If 작업장 <> "" And 작업장 <> "전체" Then
                    레코드키_생성_배열 = 생산일보구분 & "|투입인원|" & 작업장 & "|" & 인원구분
                Else
                    레코드키_생성_배열 = 생산일보구분 & "|투입인원|" & 인원구분
                End If
            Case "ProductionAnalysis"
                레코드키_생성_배열 = 생산일보구분 & "|ProductionAnalysis"
            Case Else
                레코드키_생성_배열 = 생산일보구분 & "|" & 레코드유형 & "|" & CStr(dataArr(r, 5))
        End Select
    End Function

    ' ============================================================
    ' 레코드 키 생성 (내용 기반 - 순번 대신 실제 데이터로 식별)
    ' DB 열 기준: F9=완제품_구분, F10=품명, F11=가공라인, F12=생산량, F13=소요량
    '             F17=인원_작업장, F18=인원_구분
    ' ============================================================
    Private Function 레코드키_생성(ws As Worksheet, 행 As Long) As String
        Dim 생산일보구분 As String, 레코드유형 As String
        Dim 원지종별 As String, ROLL_ID As String, KG As String
        Dim 인원구분 As String, 작업장 As String
        Dim 구분 As String, 품명 As String, 가공라인 As String, 가공규격 As String, 수량 As String

        생산일보구분 = CStr(ws.Cells(행, 3).Value)
        레코드유형 = CStr(ws.Cells(행, 4).Value)

        Select Case 레코드유형
            Case "원지"
                원지종별 = CStr(ws.Cells(행, 6).Value)
                ROLL_ID = Replace(Trim(CStr(ws.Cells(행, 7).Value)), "|", "_")
                KG = Replace(CStr(ws.Cells(행, 8).Value), "|", "_")
                레코드키_생성 = 생산일보구분 & "|원지|" & 원지종별 & "|" & ROLL_ID & "|" & KG
            Case "완제품"
                구분 = Replace(Trim(CStr(ws.Cells(행, 9).Value)), "|", "_")
                품명 = Replace(Trim(CStr(ws.Cells(행, 10).Value)), "|", "_")
                가공라인 = Replace(Trim(CStr(ws.Cells(행, 11).Value)), "|", "_")
                가공규격 = Replace(Trim(CStr(ws.Cells(행, 12).Value)), "|", "_")
                수량 = Replace(CStr(ws.Cells(행, 13).Value), "|", "_")
                레코드키_생성 = 생산일보구분 & "|완제품|" & 구분 & "|" & 품명 & "|" & 가공라인 & "|" & 가공규격 & "|" & 수량
            Case "투입인원"
                작업장 = Trim(CStr(ws.Cells(행, 19).Value))
                인원구분 = CStr(ws.Cells(행, 20).Value)
                If 작업장 <> "" And 작업장 <> "전체" Then
                    레코드키_생성 = 생산일보구분 & "|투입인원|" & 작업장 & "|" & 인원구분
                Else
                    레코드키_생성 = 생산일보구분 & "|투입인원|" & 인원구분
                End If
            Case "ProductionAnalysis"
                레코드키_생성 = 생산일보구분 & "|ProductionAnalysis"
            Case Else
                레코드키_생성 = 생산일보구분 & "|" & 레코드유형 & "|" & CStr(ws.Cells(행, 5).Value)
        End Select
    End Function

    ' ============================================================
    ' 새 데이터 수집 (Dictionary 사용, 배열 일괄 읽기로 속도 개선)
    ' ============================================================
    Private Sub 새데이터_수집_v3(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        ws일보 As Worksheet, 담당 As String, 이사 As String)

        ' 생산일보 시트 전체를 배열로 일괄 읽기 (A1:BV133 = 생산일보_셀매핑.json data_range)
        Dim 일보Arr As Variant
        일보Arr = ws일보.Range(ws일보.Cells(1, 1), ws일보.Cells(133, 74)).Value

        ' ──────────────────────────────────────────────────
        ' T01 원지 주간 KT-1~5, FT (데이터행 7~16)
        ' Co=C,O,AA,AM,AY,BK(3,15,27,39,51,63), ID=base+2, KG=base+9
        ' ──────────────────────────────────────────────────
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "KT-1", 3, 5, 12, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "KT-2", 15, 17, 24, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "KT-3", 27, 29, 36, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "KT-4", 39, 41, 48, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "KT-5", 51, 53, 60, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 7, 16, "생산일보_주간", "FT", 63, 65, 72, 담당, 이사)

        ' T02 원지 직기 JRT-1~2, BT-1~3, PRE (데이터행 20~29)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "JRT-1", 3, 5, 12, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "JRT-2", 15, 17, 24, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "BT-1", 27, 29, 36, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "BT-2", 39, 41, 48, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "BT-3", 51, 53, 60, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 20, 29, "생산일보_주간", "PRE", 63, 65, 72, 담당, 이사)

        ' T04 원지 봉제 KT-1~5, FT (데이터행 35~44)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "KT-1", 3, 5, 12, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "KT-2", 15, 17, 24, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "KT-3", 27, 29, 36, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "KT-4", 39, 41, 48, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "KT-5", 51, 53, 60, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 35, 44, "생산일보_야간", "FT", 63, 65, 72, 담당, 이사)

        ' T05 원지 봉제 JRT-1~2, BT-1~3, PRE (데이터행 48~57)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "JRT-1", 3, 5, 12, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "JRT-2", 15, 17, 24, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "BT-1", 27, 29, 36, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "BT-2", 39, 41, 48, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "BT-3", 51, 53, 60, 담당, 이사)
        Call 원지_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 48, 57, "생산일보_야간", "PRE", 63, 65, 72, 담당, 이사)

        ' T07 완제품 주간 (데이터행 68~77): B(2)=제품, Q(17)=설비, T(20)=생산량, X(24)=소요량
        Call 완제품_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 68, 77, "생산일보_주간", 담당, 이사)

        ' T08 완제품 야간 (데이터행 68~77): AK(37)=제품, AZ(52)=설비, BC(55)=생산량
        Call 완제품_수집_II_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 담당, 이사)

        ' T11 투입인원: 108행(작업시간), 111행(Direct), 112행(Indirect)
        Call 투입인원_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 108, 112, "생산일보", 담당, 이사)

        ' T10 원자재수율 (84~101행, 수식 자동계산 값 읽기)
        Call ProductionAnalysis_수집_v3_Arr(dict, ID카운터, 작성일, 일보Arr, 담당, 이사)

        ' T12 비가동현황 (121~126행)
        Call 비가동현황_수집_Arr(dict, ID카운터, 작성일, 일보Arr, 담당, 이사)
    End Sub

    ' ============================================================
    ' 원지 데이터 수집 v3 (배열 버전 - 시트 접근 없음)
    ' Co열 = C,O,AA,AM,AY,BK(3,15,27,39,51,63), ROLLID열 = base+2, KG열 = base+9
    Private Sub 원지_수집_v3_Arr(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        일보Arr As Variant, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, _
        종별 As String, Co열 As Long, ROLLID열 As Long, KG열 As Long, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        For r = 시작행 To 끝행
            Dim CoVal As String, ROLL_ID As String, KG As Variant
            CoVal = Trim(CStr(IIf(IsEmpty(일보Arr(r, Co열)), "", 일보Arr(r, Co열))))
            ROLL_ID = Trim(CStr(일보Arr(r, ROLLID열)))
            KG = 일보Arr(r, KG열)
            If ROLL_ID <> "" And ROLL_ID <> "계" And ROLL_ID <> "TOTAL" And Not IsNumeric(ROLL_ID) Then
                Dim key As String
                key = 생산일보구분 & "|원지|" & 종별 & "|" & Replace(ROLL_ID, "|", "_") & "|" & Replace(CStr(IIf(IsEmpty(KG), "", KG)), "|", "_")
                If CoVal = "" Then CoVal = RawData_원단명ToCo(ROLL_ID)
                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일: rec("F3") = 생산일보구분: rec("F4") = "원지": rec("F5") = 순번
                rec("F6") = 종별: rec("F7") = ROLL_ID: rec("F8") = KG: rec("F61") = 담당: rec("F62") = 이사
                rec("F72") = CoVal   ' 원지_Co (C,O,AA,AM,AY,BK) 생산일보 Co열 또는 Raw data 원단명→Co fallback
                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' ============================================================
    ' 원지 데이터 수집 v3 (ROLLID열/KG열 파라미터로 6개 KT그룹 지원)
    ' 실제 시트 구조:
    '   KT-①: ROLLID열=3,  KG열=9
    '   KT-②: ROLLID열=12, KG열=18
    '   KT-③: ROLLID열=21, KG열=27
    '   KT-④: ROLLID열=30, KG열=36
    '   BT-⑤: ROLLID열=39, KG열=45
    '   KT-⑥: ROLLID열=48, KG열=54
    ' ============================================================
    Private Sub 원지_수집_v3(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        ws일보 As Worksheet, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, _
        종별 As String, ROLLID열 As Long, KG열 As Long, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1

        For r = 시작행 To 끝행
            Dim ROLL_ID As String, KG As Variant
            ROLL_ID = Trim(CStr(ws일보.Cells(r, ROLLID열).Value))
            KG = ws일보.Cells(r, KG열).Value

            ' ROLL_ID가 유효한 데이터인 경우만 처리 (계/TOTAL/숫자 제외)
            If ROLL_ID <> "" And ROLL_ID <> "계" And ROLL_ID <> "TOTAL" And Not IsNumeric(ROLL_ID) Then
                Dim key As String
                key = 생산일보구분 & "|원지|" & 종별 & "|" & Replace(ROLL_ID, "|", "_") & "|" & Replace(CStr(IIf(IsEmpty(KG), "", KG)), "|", "_")

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = 생산일보구분
                rec("F4") = "원지"
                rec("F5") = 순번
                rec("F6") = 종별
                rec("F7") = ROLL_ID
                rec("F8") = KG
                rec("F61") = 담당
                rec("F62") = 이사
                rec("F72") = RawData_원단명ToCo(ROLL_ID)

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' 완제품 데이터 수집 v3 - 배열 버전
    ' T07 주간 완제품: B(2)=제품, Q(17)=설비호기, T(20)=생산량BOX, X(24)=소요량(수식)
    ' AB(28)=KP11(수식), AD(30)=KP13(수식), AF(32)=AJ11(수식), AH(34)=AJ13(수식)
    Private Sub 완제품_수집_v3_Arr(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        일보Arr As Variant, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        For r = 시작행 To 끝행
            Dim 품명 As String
            품명 = Trim(CStr(일보Arr(r, 2)))
            If 품명 <> "" And 품명 <> "계" And 품명 <> "TOTAL" And Not IsNumeric(품명) Then
                Dim key As String
                Dim 구분값 As String, 가공라인 As String, 생산량 As String, 소요량값 As Variant
                가공라인 = Replace(Trim(CStr(일보Arr(r, 17))), "|", "_")
                구분값 = 가공라인
                If 구분값 = "" Then 구분값 = "KT"
                생산량 = Replace(Trim(CStr(일보Arr(r, 20))), "|", "_")
                소요량값 = 일보Arr(r, 24)
                key = 생산일보구분 & "|완제품|" & 구분값 & "|" & Replace(품명, "|", "_") & "|" & 가공라인 & "|" & 생산량 & "|" & Replace(CStr(IIf(IsEmpty(소요량값), "", 소요량값)), "|", "_")
                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일: rec("F3") = 생산일보구분: rec("F4") = "완제품": rec("F5") = 순번
                rec("F9") = Trim(CStr(일보Arr(r, 17))): rec("F10") = 품명: rec("F11") = Trim(CStr(일보Arr(r, 17)))
                rec("F12") = 일보Arr(r, 20): rec("F13") = 일보Arr(r, 24):                 rec("F14") = 일보Arr(r, 28)
                rec("F15") = 일보Arr(r, 30): rec("F18") = Empty: rec("F61") = 담당: rec("F62") = 이사
                ' AJ11(AF=32수식)→col16, AJ13(AH=34수식)→col17: O열 우측 저장
                If Not IsEmpty(일보Arr(r, 32)) Then rec("F16") = 일보Arr(r, 32)
                If Not IsEmpty(일보Arr(r, 34)) Then rec("F17") = 일보Arr(r, 34)
                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' ============================================================
    ' 완제품 데이터 수집 v3 (생산일보 I)
    ' 시트 열 매핑:
    '   col 2  → F9  완제품_구분 (순번/No)
    '   col 3  → F10 완제품_품명
    '   col 13 → F11 완제품_가공라인호기
    '   col 17 → F12 완제품_생산량BOX
    '   col 21 → F13 완제품_원지기준소요량
    '   col 25 → F14 완제품_PLT11
    '   col 29 → F15 완제품_PLT13
    '   col 33 → F16 완제품_인원
    ' ============================================================
    Private Sub 완제품_수집_v3(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        ws일보 As Worksheet, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, _
        담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1

        For r = 시작행 To 끝행
            Dim 품명 As String
            품명 = Trim(CStr(ws일보.Cells(r, 3).Value))

            If 품명 <> "" And 품명 <> "계" And 품명 <> "TOTAL" And Not IsNumeric(품명) Then
                Dim key As String
                Dim 구분값 As String, 가공라인 As String, 생산량 As String, 소요량값 As Variant
                가공라인 = Replace(Trim(CStr(ws일보.Cells(r, 13).Value)), "|", "_")
                구분값 = 가공라인: If 구분값 = "" Then 구분값 = "KT"
                생산량 = Replace(Trim(CStr(ws일보.Cells(r, 17).Value)), "|", "_")
                소요량값 = ws일보.Cells(r, 21).Value
                key = 생산일보구분 & "|완제품|" & 구분값 & "|" & Replace(품명, "|", "_") & "|" & 가공라인 & "|" & 생산량 & "|" & Replace(CStr(IIf(IsEmpty(소요량값), "", 소요량값)), "|", "_")

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = 생산일보구분
                rec("F4") = "완제품"
                rec("F5") = 순번
                rec("F9") = Trim(CStr(ws일보.Cells(r, 13).Value))         ' 완제품_구분 = 가공라인호기(설비호수 KT①·KT-②)
                rec("F10") = 품명                                         ' 품명
                rec("F11") = Trim(CStr(ws일보.Cells(r, 13).Value))        ' 가공라인호기
                rec("F12") = ws일보.Cells(r, 17).Value                    ' 생산량(BOX)
                rec("F13") = ws일보.Cells(r, 21).Value                    ' 원지기준소요량
                rec("F14") = ws일보.Cells(r, 25).Value                    ' PLT(11)
                rec("F15") = ws일보.Cells(r, 29).Value                    ' PLT(13)
                rec("F18") = ws일보.Cells(r, 33).Value                    ' 인원
                rec("F61") = 담당
                rec("F62") = 이사

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' 투입인원 데이터 수집 v3 - 배열 버전
    ' T11 작업시간(108행), 인력투입 Direct(111행), Indirect(112행)
    ' 설비별 base열: KT=7, FT=43, JRT=19, BT=31, PRE=55 (생산일보_셀매핑.json)
    ' 각 설비: 주간정규=base, 주간연장=base+2, 야간정규=base+4, 야간심야=base+6, 휴일=base+8
    Private Sub 투입인원_수집_v3_Arr(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        일보Arr As Variant, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, 담당 As String, 이사 As String)

        Dim 순번 As Long, 설비idx As Long
        Dim 설비명 As Variant, 설비Base As Variant
        Dim bc As Long, rType As Long
        순번 = 1
        설비명 = Array("KT", "FT", "JRT", "BT", "PRE")
        설비Base = Array(7, 43, 19, 31, 55)
        Dim 행라벨 As Variant
        행라벨 = Array(Array(108, "작업시간"), Array(111, "인력Direct"), Array(112, "인력Indirect"))

        Dim h As Long
        For h = 0 To 2
            Dim 행번호 As Long, 구분명 As String
            행번호 = 행라벨(h)(0)
            구분명 = 행라벨(h)(1)
            For 설비idx = 0 To 4
                bc = 설비Base(설비idx)
                Dim hasData As Boolean
                hasData = False
                Dim c As Long
                For c = bc To bc + 8 Step 2
                    If Not IsEmpty(일보Arr(행번호, c)) Then
                        If IsNumeric(일보Arr(행번호, c)) Then
                            If CDbl(일보Arr(행번호, c)) <> 0 Then hasData = True: Exit For
                        End If
                    End If
                Next c
                If hasData Then
                    Dim key As String
                    key = 생산일보구분 & "|투입인원|" & CStr(설비명(설비idx)) & "|" & 구분명
                    Dim rec As Object
                    Set rec = CreateObject("Scripting.Dictionary")
                    rec("F2") = ID카운터 + 1
                    rec("F1") = 작성일: rec("F3") = 생산일보구분: rec("F4") = "투입인원": rec("F5") = 순번
                    rec("F19") = CStr(설비명(설비idx)): rec("F20") = 구분명
                    rec("F21") = 일보Arr(행번호, bc)
                    rec("F22") = 일보Arr(행번호, bc + 2)
                    rec("F23") = 일보Arr(행번호, bc + 4)
                    rec("F24") = 일보Arr(행번호, bc + 6)
                    rec("F25") = 일보Arr(행번호, bc + 8)
                    rec("F61") = 담당: rec("F62") = 이사
                    Set dict(key) = rec
                    ID카운터 = ID카운터 + 1
                    순번 = 순번 + 1
                End If
            Next 설비idx
        Next h
    End Sub

    ' ============================================================
    ' 투입인원 데이터 수집 v3 (생산일보 I)
    ' 시트 열 매핑:
    '   col 2  → F17 인원_작업장 ("전체" 고정)
    '   col 2  → F18 인원_구분 (정규시간/연장근로)
    '   col 6  → F19 인원_인원수
    '   col 10 → F20 인원_작업시간
    '   col 13 → F21 인원_총작업시간
    '   col 17 → F22 인원_실투입원
    '   col 21 → F23 인원_환산투입원
    '   col 25 → F24 인원_인당생산량
    ' ============================================================
    Private Sub 투입인원_수집_v3(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        ws일보 As Worksheet, 시작행 As Long, 끝행 As Long, 생산일보구분 As String, _
        담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1

        For r = 시작행 To 끝행
            Dim 구분 As String
            구분 = Trim(CStr(ws일보.Cells(r, 2).Value))

            If 구분 = "정규시간" Or 구분 = "연장근로" Then
                Dim key As String
                key = 생산일보구분 & "|투입인원|" & 구분

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = 생산일보구분
                rec("F4") = "투입인원"
                rec("F5") = 순번
                rec("F19") = "전체"
                rec("F20") = 구분
                rec("F21") = ws일보.Cells(r, 6).Value     ' 주간정규
                rec("F22") = ws일보.Cells(r, 10).Value    ' 주간연장
                rec("F23") = ws일보.Cells(r, 13).Value    ' 야간정규
                rec("F24") = ws일보.Cells(r, 17).Value    ' 야간심야
                rec("F25") = ws일보.Cells(r, 21).Value    ' 휴일
                rec("F26") = ws일보.Cells(r, 25).Value    ' 인당생산량_실투 (Y)
                rec("F27") = ws일보.Cells(r, 29).Value    ' 인당생산량_환산투 (AC)
                rec("F61") = 담당
                rec("F62") = 이사

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' ============================================================
    ' Production Analysis 수집 v3 (생산일보 I)
    ' 실제 시트 구조 (우측 PA영역):
    '   R25: 원지투입(kg) 총량   → col45=일, col49=주, col53=월
    '   R26: 원지투입(kg) 평균   → col45=일, col49=주, col53=월
    '   R27: 양품환산(kg) 총량   → col45=일, col49=주, col53=월
    '   R28: 양품환산(kg) 평균   → col45=일, col49=주, col53=월
    '   R29: 원지수율 율          → col45=일, col49=주, col53=월
    '   R30: LOSS율 율            → col45=일, col49=주, col53=월
    '   R31: 양품율 목표대비      → col45=일, col49=주, col53=월
    '   R32: pallet              → col45=일, col49=주, col53=월
    '   R33: Today's             → col41
    '   R34: 특이사항(col36), 해결책(col44)
    ' ============================================================
    Private Sub ProductionAnalysis_수집_v3(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, ws일보 As Worksheet, 담당 As String, 이사 As String)

        Dim key As String
        key = "생산일보I|ProductionAnalysis"

        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("F2") = ID카운터 + 1
        rec("F1") = 작성일
        rec("F3") = "생산일보I"
        rec("F4") = "ProductionAnalysis"
        rec("F5") = 1

        ' ── PA 생산일보1: AS22:AS31(일), AW22:AW31(주), BA22:BA31(월) → F28~F57 ──
        rec("F28") = ws일보.Cells(22, 45).Value    ' PA_원지투입_일
        rec("F29") = ws일보.Cells(23, 45).Value    ' PA_원지투입_일평균
        rec("F30") = ws일보.Cells(24, 45).Value    ' PA_기준소요량_일총량
        rec("F31") = ws일보.Cells(25, 45).Value    ' PA_기준소요량_일평균
        rec("F32") = ws일보.Cells(26, 45).Value    ' PA_원지수율_일
        rec("F33") = ws일보.Cells(27, 45).Value    ' PA_LOSS율_일
        rec("F34") = ws일보.Cells(28, 45).Value    ' PA_양품율_일
        rec("F35") = ws일보.Cells(29, 45).Value    ' PA_pallet_일
        rec("F36") = ws일보.Cells(30, 45).Value    ' PA_목표양품율_일
        rec("F37") = ws일보.Cells(31, 45).Value    ' PA_목표인당PLT_일
        rec("F38") = ws일보.Cells(22, 49).Value: rec("F39") = ws일보.Cells(23, 49).Value
        rec("F40") = ws일보.Cells(24, 49).Value: rec("F41") = ws일보.Cells(25, 49).Value
        rec("F42") = ws일보.Cells(26, 49).Value: rec("F43") = ws일보.Cells(27, 49).Value
        rec("F44") = ws일보.Cells(28, 49).Value: rec("F45") = ws일보.Cells(29, 49).Value
        rec("F46") = ws일보.Cells(30, 49).Value: rec("F47") = ws일보.Cells(31, 49).Value
        rec("F48") = ws일보.Cells(22, 53).Value: rec("F49") = ws일보.Cells(23, 53).Value
        rec("F50") = ws일보.Cells(24, 53).Value: rec("F51") = ws일보.Cells(25, 53).Value
        rec("F52") = ws일보.Cells(26, 53).Value: rec("F53") = ws일보.Cells(27, 53).Value
        rec("F54") = ws일보.Cells(28, 53).Value: rec("F55") = ws일보.Cells(29, 53).Value
        rec("F56") = ws일보.Cells(30, 53).Value: rec("F57") = ws일보.Cells(31, 53).Value

        rec("F58") = ws일보.Cells(33, 41).Value    ' Todays (R33 col41)
        Dim 특이사항 As String
        특이사항 = Trim(CStr(ws일보.Cells(37, 36).Value))   ' AJ37
        If Trim(CStr(ws일보.Cells(35, 36).Value)) <> "" Then 특이사항 = Trim(CStr(ws일보.Cells(35, 36).Value)) & vbCrLf & 특이사항
        rec("F59") = 특이사항
        rec("F60") = Trim(CStr(ws일보.Cells(37, 44).Value))   ' AR37 해결책
        rec("F61") = 담당
        rec("F62") = 이사

        Set dict(key) = rec
        ID카운터 = ID카운터 + 1
    End Sub

    ' Production Analysis 수집 v3 - 배열 버전
    ' T10 원자재수율 (84~101행, 수식 자동계산 값 읽기)
    ' KT(left): row87=일계,88~91=주간누계/월간누계/주간평균/월간평균
    '   G(7)=원투, L(12)=양품, K(11)=기준소요량, O(15)=원지수율, Y(25)=LOSS
    ' FT(right): rows 87-91  AT(46)=기준소요량, AX(50)=원지수율
    ' JRT(left): rows 92-96  K(11), O(15)  / BT(left): rows 97-101
    ' PRE(right): rows 92-96  AT(46), AX(50)
    ' BK84(col63)=목표원지수율, BK85(col63)=목표생산량
    Private Sub ProductionAnalysis_수집_v3_Arr(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, 일보Arr As Variant, 담당 As String, 이사 As String)

        Dim key As String
        key = "생산일보|ProductionAnalysis"
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("F2") = ID카운터 + 1
        rec("F1") = 작성일: rec("F3") = "생산일보": rec("F4") = "ProductionAnalysis": rec("F5") = 1

        ' T10 KT 일계~월간평균 (rows 87-91, left: G=7, L=12, K=11 기소, O=15 수율, Y=25 LOSS)
        rec("F28") = 일보Arr(87, 7): rec("F29") = 일보Arr(87, 12): rec("F30") = 일보Arr(87, 15)
        rec("F31") = 일보Arr(87, 11): rec("F32") = 일보Arr(87, 25): rec("F33") = 일보Arr(87, 30)
        ' KT 주간누계 (34~36)
        rec("F34") = 일보Arr(88, 7): rec("F35") = 일보Arr(88, 12): rec("F36") = 일보Arr(88, 15)
        ' KT 월간누계 (37~39)
        rec("F37") = 일보Arr(89, 7): rec("F38") = 일보Arr(89, 12): rec("F39") = 일보Arr(89, 15)
        ' FT 일계 (우측 Row87: AP=42, AU=47, AT=46 기소, AX=50 수율) (40~42)
        rec("F40") = 일보Arr(87, 42): rec("F41") = 일보Arr(87, 47): rec("F42") = 일보Arr(87, 50)
        ' JRT 일계 (좌측 Row92: G=7, L=12, O=15) (43~45)
        rec("F43") = 일보Arr(92, 7): rec("F44") = 일보Arr(92, 12): rec("F45") = 일보Arr(92, 15)
        ' BT 일계 (좌측 Row97: G=7, L=12, O=15) (46~48)
        rec("F46") = 일보Arr(97, 7): rec("F47") = 일보Arr(97, 12): rec("F48") = 일보Arr(97, 15)
        ' PRE 일계 (우측 Row92: AP=42, AU=47, AX=50) (49~51)
        rec("F49") = 일보Arr(92, 42): rec("F50") = 일보Arr(92, 47): rec("F51") = 일보Arr(92, 50)
        ' 목표원지수율(BK84=col63), 목표생산량(BK85=col63) (52~53)
        rec("F52") = 일보Arr(84, 63): rec("F53") = 일보Arr(85, 63)
        ' FT 주간누계 (우측 Row88: AP=42, AU=47), 월간누계 (우측 Row89) (54~57)
        rec("F54") = 일보Arr(88, 42): rec("F55") = 일보Arr(88, 47)
        rec("F56") = 일보Arr(89, 42): rec("F57") = 일보Arr(89, 47)
        ' JRT 주간누계 (좌측 Row93: G=7, L=12), 월간누계 (좌측 Row94) (58~60)
        rec("F58") = 일보Arr(93, 7): rec("F59") = 일보Arr(93, 12)
        rec("F60") = 일보Arr(94, 7)
        rec("F61") = 담당: rec("F62") = 이사
        rec("F63") = 일보Arr(94, 12)
        ' BT 주간누계 (좌측 Row98: G=7, L=12), 월간누계 (좌측 Row99) (64~67)
        rec("F64") = 일보Arr(98, 7): rec("F65") = 일보Arr(98, 12)
        rec("F66") = 일보Arr(99, 7): rec("F67") = 일보Arr(99, 12)
        ' PRE 주간누계 (우측 Row93: AP=42, AU=47), 월간누계 (우측 Row94) (68~71)
        rec("F68") = 일보Arr(93, 42): rec("F69") = 일보Arr(93, 47)
        rec("F70") = 일보Arr(94, 42): rec("F71") = 일보Arr(94, 47)
        ' T10 주간평균/월간평균 (73~102)
        rec("F73") = 일보Arr(90, 7): rec("F74") = 일보Arr(90, 12): rec("F75") = 일보Arr(90, 15)
        rec("F76") = 일보Arr(91, 7): rec("F77") = 일보Arr(91, 12): rec("F78") = 일보Arr(91, 15)
        rec("F79") = 일보Arr(90, 42): rec("F80") = 일보Arr(90, 47): rec("F81") = 일보Arr(90, 50)
        rec("F82") = 일보Arr(91, 42): rec("F83") = 일보Arr(91, 47): rec("F84") = 일보Arr(91, 50)
        rec("F85") = 일보Arr(95, 7): rec("F86") = 일보Arr(95, 12): rec("F87") = 일보Arr(95, 15)
        rec("F88") = 일보Arr(96, 7): rec("F89") = 일보Arr(96, 12): rec("F90") = 일보Arr(96, 15)
        rec("F91") = 일보Arr(100, 7): rec("F92") = 일보Arr(100, 12): rec("F93") = 일보Arr(100, 15)
        rec("F94") = 일보Arr(101, 7): rec("F95") = 일보Arr(101, 12): rec("F96") = 일보Arr(101, 15)
        rec("F97") = 일보Arr(95, 42): rec("F98") = 일보Arr(95, 47): rec("F99") = 일보Arr(95, 50)
        rec("F100") = 일보Arr(96, 42): rec("F101") = 일보Arr(96, 47): rec("F102") = 일보Arr(96, 50)
        ' 기준소요량 주간/월간 누계·평균 (103~122)
        rec("F103") = 일보Arr(88, 11): rec("F104") = 일보Arr(89, 11): rec("F105") = 일보Arr(90, 11): rec("F106") = 일보Arr(91, 11)
        rec("F107") = 일보Arr(88, 46): rec("F108") = 일보Arr(89, 46): rec("F109") = 일보Arr(90, 46): rec("F110") = 일보Arr(91, 46)
        rec("F111") = 일보Arr(93, 11): rec("F112") = 일보Arr(94, 11): rec("F113") = 일보Arr(95, 11): rec("F114") = 일보Arr(96, 11)
        rec("F115") = 일보Arr(98, 11): rec("F116") = 일보Arr(99, 11): rec("F117") = 일보Arr(100, 11): rec("F118") = 일보Arr(101, 11)
        rec("F119") = 일보Arr(93, 46): rec("F120") = 일보Arr(94, 46): rec("F121") = 일보Arr(95, 46): rec("F122") = 일보Arr(96, 46)
        ' FT,JRT,BT,PRE 일계_기소 (123~126)
        rec("F123") = 일보Arr(87, 46): rec("F124") = 일보Arr(92, 11): rec("F125") = 일보Arr(97, 11): rec("F126") = 일보Arr(92, 46)
        Set dict(key) = rec
        ID카운터 = ID카운터 + 1
    End Sub

    ' ============================================================
    ' 원지 데이터 수집 II v3 (생산일보 II)
    ' ============================================================
    Private Sub 원지_수집_II_v3(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        ws일보 As Worksheet, 시작행 As Long, 생산일보구분 As String, 종별 As String, _
        ROLLID열 As Long, KG열 As Long, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1

        For r = 시작행 To 시작행 + 20
            Dim 품명 As String, KG As Variant
            품명 = Trim(CStr(ws일보.Cells(r, ROLLID열).Value))
            KG = ws일보.Cells(r, KG열).Value

            If 품명 <> "" And 품명 <> "계" And Not IsNumeric(품명) And Not IsEmpty(KG) Then
                Dim key As String
                key = 생산일보구분 & "|원지|" & 종별 & "|" & Replace(품명, "|", "_") & "|" & Replace(CStr(IIf(IsEmpty(KG), "", KG)), "|", "_")

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = 생산일보구분
                rec("F4") = "원지"
                rec("F5") = 순번
                rec("F6") = 종별
                rec("F7") = 품명
                rec("F8") = KG
                rec("F61") = 담당
                rec("F62") = 이사
                rec("F72") = RawData_원단명ToCo(품명)

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            ElseIf 품명 = "" And IsEmpty(KG) Then
                Exit For
            End If
        Next r
    End Sub

    ' 원지 데이터 수집 II v3 - 배열 버전
    Private Sub 원지_수집_II_v3_Arr(dict As Object, ByRef ID카운터 As Long, 작성일 As Variant, _
        일보Arr As Variant, 시작행 As Long, 생산일보구분 As String, 종별 As String, _
        ROLLID열 As Long, KG열 As Long, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        For r = 시작행 To 시작행 + 20
            Dim 품명 As String, KG As Variant
            품명 = Trim(CStr(일보Arr(r, ROLLID열)))
            KG = 일보Arr(r, KG열)
            If 품명 <> "" And 품명 <> "계" And Not IsNumeric(품명) And Not IsEmpty(KG) Then
                Dim key As String
                key = 생산일보구분 & "|원지|" & 종별 & "|" & Replace(품명, "|", "_") & "|" & Replace(CStr(IIf(IsEmpty(KG), "", KG)), "|", "_")
                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일: rec("F3") = 생산일보구분: rec("F4") = "원지": rec("F5") = 순번
                rec("F6") = 종별: rec("F7") = 품명: rec("F8") = KG: rec("F61") = 담당: rec("F62") = 이사
                rec("F72") = RawData_원단명ToCo(품명)
                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            ElseIf 품명 = "" And IsEmpty(KG) Then
                Exit For
            End If
        Next r
    End Sub

    ' ============================================================
    ' 완제품 데이터 수집 II v3 (생산일보 II, 우측 col31~55)
    ' 시트 구조 (역변환 코드 확인 기준):
    '   품명(col33=AG), 생산량BOX(col41=AO), 원지기준소요량(col45=AS),
    '   PLT11(col49=AW), PLT13(col52=AZ), 인원(col55=BC)
    '   구분라벨: col31(AE) - 첫 데이터행에 JRT/WT/FT/BT, TOTAL행은 별도
    '   데이터 구분별:
    '     JRT : R44~R47 (AE44=JRT, R48=TOTAL)
    '     WT  : R49~R50 (AE49=WT,  R51=TOTAL)
    '     FT  : R52~R53 (AE52=FT,  R54=TOTAL)
    '     BT  : R55~R56 (AE55=BT,  R57=TOTAL)
    '   ★ R58~ 는 완제품 아님 (PA2 요약 등) → 수집 제외
    ' ============================================================
    Private Sub 완제품_수집_II_v3(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, ws일보 As Worksheet, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        Dim 현재구분 As String
        현재구분 = ""

        For r = 44 To 57
            Dim 구분셀 As String
            구분셀 = Trim(CStr(ws일보.Cells(r, 31).Value))

            ' 구분 라벨 업데이트 (JRT/WT/FT/BT 등)
            If 구분셀 <> "" And UCase(구분셀) <> "TOTAL" Then
                현재구분 = 구분셀
            End If

            ' TOTAL행 건너뛰기
            If UCase(구분셀) = "TOTAL" Then GoTo NextRow완제품II

            ' 데이터 존재 여부 확인
            Dim 품명 As String
            품명 = Trim(CStr(ws일보.Cells(r, 33).Value))
            If 품명 = "" And IsEmpty(ws일보.Cells(r, 41).Value) And _
                IsEmpty(ws일보.Cells(r, 45).Value) And IsEmpty(ws일보.Cells(r, 49).Value) Then
                GoTo NextRow완제품II
            End If

            If 현재구분 <> "" Then
                ' 키 생성: 생산일보II|완제품|구분|품명|가공라인(빈값)|생산량|소요량
                Dim key As String
                Dim 생산량Val As String, 소요량Val As String
                On Error Resume Next
                생산량Val = Replace(CStr(IIf(IsEmpty(ws일보.Cells(r, 41).Value), "", ws일보.Cells(r, 41).Value)), "|", "_")
                If Err.Number <> 0 Then 생산량Val = "": Err.Clear
                소요량Val = Replace(CStr(IIf(IsEmpty(ws일보.Cells(r, 45).Value), "", ws일보.Cells(r, 45).Value)), "|", "_")
                If Err.Number <> 0 Then 소요량Val = "": Err.Clear
                On Error GoTo 0

                key = "생산일보II|완제품|" & Replace(현재구분, "|", "_") & "|" & _
                    Replace(품명, "|", "_") & "||" & 생산량Val & "|" & 소요량Val

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = "생산일보II"
                rec("F4") = "완제품"
                rec("F5") = 순번
                rec("F9") = 현재구분                               ' 구분 (JRT/WT/FT/BT)
                rec("F10") = 품명                                  ' 품명
                ' F11 (가공라인호기) = 생산일보II에는 해당 없음
                rec("F12") = ws일보.Cells(r, 41).Value             ' 생산량(BOX)
                rec("F13") = ws일보.Cells(r, 45).Value             ' 원지기준소요량
                rec("F14") = ws일보.Cells(r, 49).Value             ' PLT(11)
                rec("F15") = ws일보.Cells(r, 52).Value             ' PLT(13)
                rec("F18") = ws일보.Cells(r, 55).Value             ' 인원
                rec("F61") = 담당
                rec("F62") = 이사

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
NextRow완제품II:
        Next r
    End Sub

    ' 완제품 데이터 수집 II v3 - 배열 버전
    ' T08 야간 완제품 (68~77행): AK(37)=제품, AZ(52)=설비호기, BC(55)=생산량BOX, BG(59)=원지기준소요량
    ' BI(63)=PLT1, BK(65)=PLT2, BM(67)=AJ11, BO(69)=AJ13 (소요59+4,6,8,10)
    Private Sub 완제품_수집_II_v3_Arr(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, 일보Arr As Variant, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        For r = 68 To 77
            Dim 품명 As String
            품명 = Trim(CStr(일보Arr(r, 37)))
            If 품명 = "" Or 품명 = "계" Or 품명 = "TOTAL" Or IsNumeric(품명) Then GoTo NextRow완야간
            Dim 생산량Val As String
            On Error Resume Next
            생산량Val = Replace(CStr(IIf(IsEmpty(일보Arr(r, 55)), "", 일보Arr(r, 55))), "|", "_")
            If Err.Number <> 0 Then 생산량Val = "": Err.Clear
            On Error GoTo 0
            Dim 설비 As String
            설비 = Replace(Trim(CStr(일보Arr(r, 52))), "|", "_")
            If 설비 = "" Then 설비 = "야간"
            Dim key As String
            key = "생산일보_야간|완제품|" & 설비 & "|" & Replace(품명, "|", "_") & "|" & 설비 & "|" & 생산량Val & "|"
            Dim rec As Object
            Set rec = CreateObject("Scripting.Dictionary")
            rec("F2") = ID카운터 + 1
            rec("F1") = 작성일: rec("F3") = "생산일보_야간": rec("F4") = "완제품": rec("F5") = 순번
            rec("F9") = 설비: rec("F10") = 품명: rec("F11") = 설비
            rec("F12") = 일보Arr(r, 55): rec("F13") = 일보Arr(r, 59)
            If Not IsEmpty(일보Arr(r, 63)) Then rec("F14") = 일보Arr(r, 63)  ' PLT1(BI)
            If Not IsEmpty(일보Arr(r, 65)) Then rec("F15") = 일보Arr(r, 65)  ' PLT2(BK)
            If Not IsEmpty(일보Arr(r, 67)) Then rec("F16") = 일보Arr(r, 67)  ' AJ11(BM)
            If Not IsEmpty(일보Arr(r, 69)) Then rec("F17") = 일보Arr(r, 69)  ' AJ13(BO)
            rec("F18") = Empty: rec("F61") = 담당: rec("F62") = 이사
            Set dict(key) = rec
            ID카운터 = ID카운터 + 1
            순번 = 순번 + 1
NextRow완야간:
        Next r
    End Sub

    ' ============================================================
    ' 투입인원 데이터 수집 II v3 (생산일보 II, R68~R77)
    ' 시트 구조:
    '   R67(좌측): 헤더 - 구분(col2), 인원(col4), 작업시간(col9),
    '              실투입원(col14), 환산투입(col18), 인당생산량(col22)
    '   R68~R77: 5품목 × 2행(정규/연장)
    '     JRT: R68(정규), R69(연장)
    '     WT : R70(정규), R71(연장)
    '     FT : R72(정규), R73(연장)
    '     BT : R74(정규), R75(연장)
    '     지관: R76(인원=정규), R77(연장)
    '   열 매핑:
    '     col 2  = 작업장 (JRT/WT/FT/BT/지관, 병합)
    '     col 4  = 구분 (정규/연장/인원)
    '     col 7  = 인원수
    '     col 9  = 작업시간
    '     col 11 = 총작업시간
    '     col 14 = 실투입원 (2행 병합 → 정규행에만 값)
    '     col 18 = 환산투입원 (2행 병합 → 정규행에만 값)
    '     col 22 = 인당생산량 (2행 병합 → 정규행에만 값)
    ' ============================================================
    Private Sub 투입인원_수집_II_v3(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, ws일보 As Worksheet, 담당 As String, 이사 As String)

        Dim r As Long, 순번 As Long
        순번 = 1
        Dim 현재작업장 As String
        현재작업장 = ""

        For r = 68 To 77
            Dim 작업장셀 As String
            작업장셀 = Trim(CStr(ws일보.Cells(r, 2).Value))

            ' 작업장 라벨 업데이트 (JRT, WT, FT, BT, 지관)
            If 작업장셀 <> "" Then
                현재작업장 = 작업장셀
            End If

            ' 구분 (정규/연장/인원)
            Dim 구분 As String
            구분 = Trim(CStr(ws일보.Cells(r, 4).Value))

            ' 유효한 구분값인 경우만 처리
            If (구분 = "정규" Or 구분 = "연장" Or 구분 = "인원") And 현재작업장 <> "" Then
                Dim key As String
                key = "생산일보II|투입인원|" & 현재작업장 & "|" & 구분

                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("F2") = ID카운터 + 1
                rec("F1") = 작성일
                rec("F3") = "생산일보II"
                rec("F4") = "투입인원"
                rec("F5") = 순번
                rec("F19") = 현재작업장                            ' 인원_작업장
                rec("F20") = 구분                                  ' 인원_구분 (정규/연장/인원)
                rec("F21") = ws일보.Cells(r, 7).Value              ' 주간정규
                rec("F22") = ws일보.Cells(r, 9).Value              ' 주간연장
                rec("F23") = ws일보.Cells(r, 11).Value             ' 야간정규
                rec("F24") = ws일보.Cells(r, 14).Value             ' 야간심야
                rec("F25") = ws일보.Cells(r, 18).Value             ' 휴일
                rec("F26") = ws일보.Cells(r, 22).Value             ' 인당생산량_실투 (V)
                rec("F27") = ws일보.Cells(r, 26).Value             ' 인당생산량_환산투 (Z)
                rec("F61") = 담당
                rec("F62") = 이사

                Set dict(key) = rec
                ID카운터 = ID카운터 + 1
                순번 = 순번 + 1
            End If
        Next r
    End Sub

    ' 비가동현황 수집 - 배열 버전
    ' T12: A(1)=비가동 내용, AJ(36)=해결책 (121~126행)
    Private Sub 비가동현황_수집_Arr(dict As Object, ByRef ID카운터 As Long, _
        작성일 As Variant, 일보Arr As Variant, 담당 As String, 이사 As String)

        Dim 특이사항 As String, 해결책 As String
        Dim r As Long
        특이사항 = ""
        해결책 = ""
        For r = 121 To 126
            Dim v1 As String, v2 As String
            v1 = Trim(CStr(일보Arr(r, 1)))
            v2 = Trim(CStr(일보Arr(r, 36)))
            If v1 <> "" Then
                If 특이사항 <> "" Then 특이사항 = 특이사항 & vbCrLf
                특이사항 = 특이사항 & v1
            End If
            If v2 <> "" Then
                If 해결책 <> "" Then 해결책 = 해결책 & vbCrLf
                해결책 = 해결책 & v2
            End If
        Next r
        If 특이사항 = "" And 해결책 = "" Then Exit Sub

        Dim key As String
        key = "생산일보|비가동현황"
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("F2") = ID카운터 + 1
        rec("F1") = 작성일: rec("F3") = "생산일보": rec("F4") = "ProductionAnalysis": rec("F5") = 2
        rec("F59") = 특이사항: rec("F60") = 해결책
        rec("F61") = 담당: rec("F62") = 이사
        Set dict(key) = rec
        ID카운터 = ID카운터 + 1
    End Sub

    ' ============================================================
    ' 레코드 비교 (True: 변경됨, False: 동일)
    ' 순번(F5)은 제외 - 순서 변경만으로는 업데이트하지 않음
    ' ============================================================
    Private Function 레코드_비교(oldRec As Object, newRec As Object) As Boolean
        Dim col As Long
        For col = 3 To 124
            If col = 5 Then GoTo SkipCompare  ' 순번 제외
            Dim oldVal As Variant, newVal As Variant
            If oldRec.exists("F" & col) Then oldVal = oldRec("F" & col) Else oldVal = Empty
            If newRec.exists("F" & col) Then newVal = newRec("F" & col) Else newVal = Empty

            If IsEmpty(oldVal) Then oldVal = ""
            If IsEmpty(newVal) Then newVal = ""

            If CStr(oldVal) <> CStr(newVal) Then
                레코드_비교 = True
                Exit Function
            End If
SkipCompare:
        Next col
        레코드_비교 = False
    End Function

    ' ============================================================
    ' 레코드 업데이트 (ID 유지, F1·F3~F126 벌크 쓰기)
    ' ============================================================
    Private Sub 레코드_업데이트(wsDB As Worksheet, oldRec As Object, newRec As Object)
        Dim 행 As Long, col As Long
        Dim arr(1 To 1, 1 To 126) As Variant
        행 = oldRec("행번호")
        For col = 1 To 126
            If newRec.exists("F" & col) Then arr(1, col) = newRec("F" & col) Else arr(1, col) = Empty
        Next col
        arr(1, 2) = oldRec("F2")
        wsDB.Range(wsDB.Cells(행, 1), wsDB.Cells(행, 126)).Value = arr
    End Sub

    ' ============================================================
    ' 레코드 추가 (벌크 쓰기)
    ' ============================================================
    Private Sub 레코드_추가(wsDB As Worksheet, newRec As Object)
        Dim 새행 As Long
        Dim arr(1 To 1, 1 To 126) As Variant
        Dim col As Long
        새행 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
        For col = 1 To 126
            If newRec.exists("F" & col) Then arr(1, col) = newRec("F" & col) Else arr(1, col) = Empty
        Next col
        wsDB.Range(wsDB.Cells(새행, 1), wsDB.Cells(새행, 126)).Value = arr
    End Sub

    ' ============================================================
    ' 레코드 삭제
    ' ============================================================
    Private Sub 레코드_삭제(wsDB As Worksheet, oldRec As Object)
        Dim 행 As Long
        행 = oldRec("행번호")
        wsDB.Rows(행).Delete
    End Sub

    ' ============================================================
    ' ★★★ DB저장 + 완제품 생산이력 업데이트 (메인 진입점) ★★★
    ' 기존 일일생산DB 전송(증분) + Raw data_index 신규제품 추가
    ' + 부자재관리 시트 생산행에 값 자동 입력 + 업데이트 내역 출력
    ' ※ DB저장 버튼에 이 매크로를 연결하세요: DB저장_및_재고반영
    ' ============================================================
    Public Sub DB저장_및_재고반영()
        ' ★ 이 매크로를 DB저장 버튼에 연결하세요!
        ' (버튼 우클릭 → 매크로 지정 → DB저장_및_재고반영)
        '
        ' ※ 완제품_생산이력_재고반영은 증분_업데이트 내부에서 자동 호출됨
        '   (무음모드 = True, MsgBox/워크북 없이 고속 처리)
        '   여기서 별도 호출하면 2중 실행이므로 호출하지 않음!
        
        On Error GoTo ErrDB
        Call 일일생산일보_최적화DB_전송
        On Error GoTo 0
        Exit Sub
        
ErrDB:
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "일일생산DB 전송 중 오류:" & vbCrLf & vbCrLf & _
            "오류 번호: " & Err.Number & vbCrLf & _
            "오류 내용: " & Err.Description, vbCritical, "DB전송 오류"
        Exit Sub
        
    End Sub

    ' ============================================================
    ' ★ 진단용: 완제품 생산이력만 단독 실행 (DB저장 없이 테스트)
    ' ============================================================
    Public Sub 완제품_생산이력_단독테스트()
        MsgBox "완제품_생산이력_재고반영을 단독 실행합니다." & vbCrLf & _
            "(일일생산DB 전송은 생략)", vbInformation, "단독 테스트"
        Call 완제품_생산이력_재고반영
    End Sub

    ' ============================================================
    ' ★ 완제품 생산이력 업데이트 (최적화 v3)
    ' - 배열 일괄 읽기/쓰기로 워크시트 I/O 최소화
    ' - D열 사전 Dictionary 빌드로 반복 검색 제거
    ' - Union 기반 일괄 서식 적용
    ' - 로그 워크북 배열 기반 일괄 출력
    ' ============================================================
    Public Sub 완제품_생산이력_재고반영(Optional 무음모드 As Boolean = False)
        Dim ws일보 As Worksheet, wsRaw As Worksheet, ws재고 As Worksheet
        Dim 작성일 As Date
        Dim r As Long, c As Long
        Dim 품명 As String, 생산량 As Double
        Dim key As Variant, i As Long, j As Long
        Dim 원래Calculation As Long
        Dim 진행단계 As String
        
        진행단계 = "초기화"
        
        ' ── 시트 참조 ──
        On Error Resume Next
        Set ws일보 = ThisWorkbook.Worksheets("생산일보")
        Set wsRaw = ThisWorkbook.Worksheets("Raw data_index")
        Set ws재고 = ThisWorkbook.Worksheets("부자재관리")
        On Error GoTo 0
        
        If ws일보 Is Nothing Then
            MsgBox "'생산일보' 시트를 찾을 수 없습니다.", vbCritical
            Exit Sub
        End If
        
        ' ── 날짜 확인 (BK4) ──
        On Error Resume Next
        Dim boVal As Variant
        boVal = ws일보.Range("BK4").Value
        On Error GoTo 0
        If IsEmpty(boVal) Or Not IsDate(boVal) Then
            MsgBox "생산일보 시트 BK4 셀에 유효한 날짜가 없습니다." & vbCrLf & _
                "현재값: " & CStr(boVal), vbCritical
            Exit Sub
        End If
        작성일 = CDate(boVal)
        
        ' ── Application 상태 설정 ──
        원래Calculation = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        ' ★★ 내부 프로파일러 ★★
        Dim tp0 As Double, tp1 As Double, 내부프로파일 As String
        tp0 = Timer: tp1 = Timer
        내부프로파일 = ""
        
        On Error GoTo ErrorHandler
        
        ' ??????????????????????????????????????????????????????
        ' 1단계: 완제품 구분|품명 + 생산량 수집 (★ 행별 고유 키 → 6건이면 6개 그룹, 중복 그룹 생성 금지)
        ' 키 = 구분|품명|순번 (순번으로 건별 구분. 부자재관리 A=제품발주처(업체명), B=제품명)
        ' ??????????????????????????????????????????????????????
        진행단계 = "1단계: 제품 수집"
        Dim 제품딕 As Object
        Set 제품딕 = CreateObject("Scripting.Dictionary")
        Dim 완제품순번 As Long: 완제품순번 = 0
        
        ' --- T07 주간 완제품: B(2)=제품, Q(17)=설비, T(20)=생산량 (rows 68-77) ---
        Dim arr1설비 As Variant, arr1제품 As Variant, arr1생산 As Variant
        arr1설비 = ws일보.Range("Q68:Q77").Value
        arr1제품 = ws일보.Range("B68:B77").Value
        arr1생산 = ws일보.Range("T68:T77").Value
        Dim 구분I As String
        For r = 1 To 10
            품명 = Trim(CStr(arr1제품(r, 1)))
            If InStr(품명, "|") > 0 Then 품명 = Replace(품명, "|", "*")
            If 품명 <> "" And UCase(품명) <> "TOTAL" And 품명 <> "계" And Not IsNumeric(품명) Then
                생산량 = 0
                If IsNumeric(arr1생산(r, 1)) Then 생산량 = CDbl(arr1생산(r, 1))
                If 생산량 <> 0 Then
                    구분I = Trim(CStr(arr1설비(r, 1)))
                    If 구분I = "" Then 구분I = "KT"
                    완제품순번 = 완제품순번 + 1
                    Dim key1 As String: key1 = 구분I & "|" & 품명 & "|" & 완제품순번
                    제품딕.Add key1, 생산량
                End If
            End If
        Next r
        
        ' --- T08 야간 완제품: AK(37)=제품, AZ(52)=설비, BC(55)=생산량 (rows 68-77) ---
        Dim arr2제품 As Variant, arr2생산 As Variant, arr2설비 As Variant
        arr2제품 = ws일보.Range("AK68:AK77").Value
        arr2생산 = ws일보.Range("BC68:BC77").Value
        arr2설비 = ws일보.Range("AZ68:AZ77").Value
        For r = 1 To 10
            품명 = Trim(CStr(arr2제품(r, 1)))
            If InStr(품명, "|") > 0 Then 품명 = Replace(품명, "|", "*")
            If 품명 <> "" And UCase(품명) <> "TOTAL" And 품명 <> "계" And Not IsNumeric(품명) Then
                생산량 = 0
                If IsNumeric(arr2생산(r, 1)) Then 생산량 = CDbl(arr2생산(r, 1))
                If 생산량 <> 0 Then
                    Dim 구분야 As String
                    구분야 = Trim(CStr(arr2설비(r, 1)))
                    If 구분야 = "" Then 구분야 = "야간"
                    완제품순번 = 완제품순번 + 1
                    key1 = 구분야 & "|" & 품명 & "|" & 완제품순번
                    제품딕.Add key1, 생산량
                End If
            End If
        Next r
        
        If 제품딕.Count = 0 Then
            Application.Calculation = 원래Calculation
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "완제품 생산 데이터가 없습니다." & vbCrLf & _
                "(T07 B68~T77, T08 AK68~BC77에서 유효한 값이 없음)", vbInformation, "생산이력 업데이트"
            Exit Sub
        End If
        
        진행단계 = "1단계 완료: " & 제품딕.Count & "건"
        내부프로파일 = 내부프로파일 & "  1.제품수집: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
        
        ' ??????????????????????????????????????????????????????
        ' 2단계: Raw data_index C열 (★ 배열 일괄 읽기 + 일괄 쓰기)
        ' ??????????????????????????????????????????????????????
        진행단계 = "2단계: Raw data_index"
        Dim Raw추가목록 As Collection: Set Raw추가목록 = New Collection
        
        If Not wsRaw Is Nothing Then
            Dim rawDict As Object: Set rawDict = CreateObject("Scripting.Dictionary")
            Dim rawLastRow As Long
            rawLastRow = wsRaw.Cells(wsRaw.Rows.Count, 3).End(xlUp).Row
            
            ' ★ 배열 일괄 읽기
            If rawLastRow >= 3 Then
                Dim arrRaw As Variant
                arrRaw = wsRaw.Range(wsRaw.Cells(3, 3), wsRaw.Cells(rawLastRow, 3)).Value
                Dim rawVal As String
                If IsArray(arrRaw) Then
                    For r = 1 To UBound(arrRaw, 1)
                        rawVal = Trim(CStr(arrRaw(r, 1)))
                        If rawVal <> "" Then
                            If Not rawDict.exists(rawVal) Then rawDict.Add rawVal, r + 2
                        End If
                    Next r
                Else
                    rawVal = Trim(CStr(arrRaw))
                    If rawVal <> "" Then rawDict.Add rawVal, 3
                End If
            End If
            
            ' ★ 신규 제품 수집(품명만, 중복 제거) 후 Raw data_index C열 일괄 쓰기
            Dim 신규Raw목록() As String
            Dim 신규Raw수 As Long: 신규Raw수 = 0
            ReDim 신규Raw목록(1 To 제품딕.Count)
            Dim 품명만 As String
            For Each key In 제품딕.keys
                ' ★ 키 형식 "구분|품명|순번" → 품명만 추출 (첫 | 와 마지막 | 사이). 기존 Mid는 순번(|1)까지 포함해 *→| 혼동 발생
                품명만 = CStr(key)
                Dim fp As Long, lp As Long
                fp = InStr(품명만, "|")
                lp = InStrRev(품명만, "|")
                If fp > 0 And lp > fp Then
                    품명만 = Trim(Mid(품명만, fp + 1, lp - fp - 1))
                ElseIf fp > 0 Then
                    품명만 = Trim(Mid(품명만, fp + 1))
                End If
                If 품명만 <> "" And Not rawDict.exists(품명만) Then
                    rawDict.Add 품명만, 0
                    신규Raw수 = 신규Raw수 + 1
                    신규Raw목록(신규Raw수) = 품명만
                    Raw추가목록.Add 품명만
                End If
            Next key
            
            If 신규Raw수 > 0 Then
                Dim newRow As Long
                newRow = wsRaw.Cells(wsRaw.Rows.Count, 3).End(xlUp).Row + 1
                ' 배열로 일괄 쓰기
                Dim arrWrite() As Variant
                ReDim arrWrite(1 To 신규Raw수, 1 To 1)
                For i = 1 To 신규Raw수
                    ' ★ C열 *→| 변질 방지: 쓰기 직전 한 번 더 | 를 * 로 복원
                    arrWrite(i, 1) = Replace(신규Raw목록(i), "|", "*")
                Next i
                wsRaw.Cells(newRow, 3).Resize(신규Raw수, 1).Value = arrWrite
            End If
        End If
        
        내부프로파일 = 내부프로파일 & "  2.Raw data: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
        
        ' ??????????????????????????????????????????????????????
        ' 3단계: 부자재관리 시트 (★ 사전 Dictionary + Union 서식)
        ' ??????????????????????????????????????????????????????
        진행단계 = "3단계: 부자재관리"
        Dim 재고업데이트목록 As Collection: Set 재고업데이트목록 = New Collection
        Dim 재고신규추가목록 As Collection: Set 재고신규추가목록 = New Collection
        Dim 재고실패목록 As Collection: Set 재고실패목록 = New Collection
        Dim 제품상태딕 As Object: Set 제품상태딕 = CreateObject("Scripting.Dictionary")
        
        ' 부자재관리 시트 구조: A=제품발주처(업체명), B=제품명, C=포장구분, D~I=제조사·지관·지폭·평량·1차·2차, J=전월재고, K=현재고, L=안전재고, M=구분, N=자재유형, O~=일별
        Const INV_START_ROW As Long = 4
        Const INV_ROWS_PER_ITEM As Long = 15
        Const INV_COL_제품발주처 As Long = 1  ' A열 (업체명)
        Const INV_COL_제품명 As Long = 2      ' B열
        Const INV_COL_구분 As Long = 13
        Const INV_COL_일별시작 As Long = 15
        
        ' ★ 서식 적용할 셀을 Union으로 모음 (루프 후 1회 적용)
        Dim rngFormat As Range
        
        If Not ws재고 Is Nothing Then
            Dim 재고lastRow As Long
            재고lastRow = ws재고.Cells(ws재고.Rows.Count, INV_COL_구분).End(xlUp).Row
            If 재고lastRow < INV_START_ROW Then 재고lastRow = INV_START_ROW
            Dim 마지막열 As Long
            마지막열 = ws재고.Cells(3, ws재고.Columns.Count).End(xlToLeft).Column
            
            ' ── 부자재관리 시트 초기화 상태면 부자재관리 모듈의 데이터업데이트 먼저 실행 (일일생산DB 기준 그룹 생성) ──
            If 재고lastRow < INV_START_ROW Or 마지막열 < INV_COL_일별시작 Then
                On Error Resume Next
                Application.Run "데이터업데이트"
                On Error GoTo 0
                재고lastRow = ws재고.Cells(ws재고.Rows.Count, INV_COL_구분).End(xlUp).Row
                If 재고lastRow < INV_START_ROW Then 재고lastRow = INV_START_ROW
                마지막열 = ws재고.Cells(3, ws재고.Columns.Count).End(xlToLeft).Column
            End If
            
            ' ── 날짜 매칭: 2행/3행 배열 일괄 읽기 ──
            Dim 대상월 As Integer: 대상월 = Month(작성일)
            Dim 대상일 As Integer: 대상일 = Day(작성일)
            Dim 날짜열 As Long: 날짜열 = 0
            
            If 마지막열 >= INV_COL_일별시작 Then
                Dim arrRow2 As Variant, arrRow3 As Variant
                arrRow2 = ws재고.Range(ws재고.Cells(2, INV_COL_일별시작), ws재고.Cells(2, 마지막열)).Value
                arrRow3 = ws재고.Range(ws재고.Cells(3, INV_COL_일별시작), ws재고.Cells(3, 마지막열)).Value
                Dim v월str As String, v일val As Variant, 매칭월 As Long, 매칭일 As Long
                Dim colCnt As Long
                If IsArray(arrRow2) Then
                    colCnt = UBound(arrRow2, 2) - LBound(arrRow2, 2) + 1
                Else
                    colCnt = 1
                End If
                
                For c = 1 To colCnt
                    If colCnt = 1 Then
                        v월str = Replace(Replace(CStr(arrRow2), "m", ""), "M", "")
                        v일val = arrRow3
                    Else
                        v월str = Replace(Replace(CStr(arrRow2(1, c)), "m", ""), "M", "")
                        v일val = arrRow3(1, c)
                    End If
                    매칭월 = 0: 매칭일 = 0
                    If IsNumeric(v월str) Then 매칭월 = CLng(val(CStr(v월str)))
                    If Not IsEmpty(v일val) Then
                        If IsDate(v일val) Then
                            매칭일 = Day(CDate(v일val))
                        ElseIf IsNumeric(v일val) Then
                            매칭일 = CLng(CDbl(v일val))
                        End If
                    End If
                    If 매칭월 = 대상월 And 매칭일 = 대상일 Then
                        날짜열 = INV_COL_일별시작 + c - 1
                        Exit For
                    End If
                Next c
            End If
            
            If 날짜열 = 0 Then
                Application.Calculation = 원래Calculation
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                Dim 재고초기화안내 As String
                재고초기화안내 = "부자재관리 시트에서 해당 날짜를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
                    "작성일: " & Format(작성일, "YYYY-MM-DD") & vbCrLf & _
                    "검색 조건: " & 대상월 & "m / " & 대상일 & "일" & vbCrLf & vbCrLf
                If 마지막열 < INV_COL_일별시작 Or 재고lastRow <= INV_START_ROW Then
                    재고초기화안내 = 재고초기화안내 & "부자재관리 시트가 초기화 상태입니다." & vbCrLf & _
                        "부자재관리 시트의 [데이터 업데이트] 버튼을 먼저 실행한 뒤, DB저장을 다시 실행해 주세요." & vbCrLf & vbCrLf
                Else
                    재고초기화안내 = 재고초기화안내 & "'다음달열 생성' 버튼으로 해당 월의 열을 먼저 생성해 주세요." & vbCrLf & vbCrLf
                End If
                MsgBox 재고초기화안내, vbExclamation, "날짜 매칭 실패"
                Exit Sub
            End If
            
            ' ★★ 재계산용 상수 1회만 계산 (블록_재고_재계산 호출 제거) ★★
            Const INV_COL_전월재고 As Long = 10  ' J열
            Const INV_COL_현재고 As Long = 11    ' K열
            Dim 재계산_열수 As Long: 재계산_열수 = 마지막열 - INV_COL_일별시작 + 1
            ' 현재고 인덱스 계산 (1회만)
            Dim 오늘일 As Long: 오늘일 = Day(Date)
            Dim 현재고_열idx As Long: 현재고_열idx = INV_COL_일별시작 + 오늘일 - 1
            If 현재고_열idx > 마지막열 Then 현재고_열idx = 마지막열
            
            ' ── ★ 품명→업체명 조회 (제품표준사양서) + A열(제품발주처)|B열(제품명) 키로 사전 빌드 → 키 "업체|제품" ──
            Dim 업체명딕 As Object: Set 업체명딕 = CreateObject("Scripting.Dictionary")
            Dim ws표준 As Worksheet
            On Error Resume Next
            Set ws표준 = ThisWorkbook.Sheets(시트명표준사양서)
            On Error GoTo 0
            If Not ws표준 Is Nothing Then
                Dim col업체 As Long, col제품 As Long, lastRow표 As Long, ri As Long
                col업체 = 0: col제품 = 0
                For ri = 1 To 20
                    Dim hdrVal As String: hdrVal = Trim(CStr(ws표준.Cells(8, ri).Value))
                    If InStr(1, hdrVal, "업체", vbTextCompare) > 0 Or InStr(1, hdrVal, "발주처", vbTextCompare) > 0 Then col업체 = ri: Exit For
                Next ri
                If col업체 = 0 Then For ri = 1 To 20: If InStr(1, Trim(CStr(ws표준.Cells(7, ri).Value)), "업체", vbTextCompare) > 0 Then col업체 = ri: Exit For: Next ri
                For ri = 1 To 20
                    If InStr(1, Trim(CStr(ws표준.Cells(8, ri).Value)), "제품", vbTextCompare) > 0 Then col제품 = ri: Exit For
                Next ri
                If col제품 = 0 Then For ri = 1 To 20: If InStr(1, Trim(CStr(ws표준.Cells(7, ri).Value)), "제품", vbTextCompare) > 0 Then col제품 = ri: Exit For: Next ri
                If col업체 > 0 And col제품 > 0 Then
                    lastRow표 = ws표준.Cells(ws표준.Rows.Count, col제품).End(xlUp).Row
                    For ri = 9 To lastRow표
                        Dim tmp품 As String, tmp업 As String
                        tmp품 = Trim(CStr(ws표준.Cells(ri, col제품).Value))
                        tmp업 = Trim(CStr(ws표준.Cells(ri, col업체).Value))
                        If tmp품 <> "" And Not 업체명딕.exists(tmp품) Then 업체명딕.Add tmp품, IIf(tmp업 = "", "-", tmp업)
                    Next ri
                End If
            End If
            Dim 제품맵 As Object: Set 제품맵 = CreateObject("Scripting.Dictionary")
            Dim 그룹시작 As Long, a업체 As String, d제품 As String
            Dim arrA As Variant, arrB As Variant
            arrA = ws재고.Range(ws재고.Cells(INV_START_ROW, INV_COL_제품발주처), ws재고.Cells(재고lastRow, INV_COL_제품발주처)).Value
            arrB = ws재고.Range(ws재고.Cells(INV_START_ROW, INV_COL_제품명), ws재고.Cells(재고lastRow, INV_COL_제품명)).Value
            If Not IsArray(arrA) Then: Dim tmpA(1 To 1, 1 To 1) As Variant: tmpA(1, 1) = arrA: arrA = tmpA
            If Not IsArray(arrB) Then: Dim tmpB(1 To 1, 1 To 1) As Variant: tmpB(1, 1) = arrB: arrB = tmpB
            For 그룹시작 = INV_START_ROW To 재고lastRow Step INV_ROWS_PER_ITEM
                Dim arrIdx As Long: arrIdx = 그룹시작 - INV_START_ROW + 1
                If arrIdx >= 1 And arrIdx <= UBound(arrB, 1) Then
                    d제품 = 재고_셀값문자열(arrB(arrIdx, 1))
                    If d제품 <> "" Then
                        a업체 = 재고_셀값문자열(arrA(arrIdx, 1))
                        Dim 맵키 As String
                        If a업체 <> "" Then 맵키 = a업체 & "|" & d제품 Else 맵키 = "|" & d제품
                        If Not 제품맵.exists(맵키) Then Set 제품맵(맵키) = New Collection
                        제품맵(맵키).Add Array(그룹시작, a업체)
                    End If
                End If
            Next 그룹시작
            
            내부프로파일 = 내부프로파일 & "  3a.Dict빌드+날짜열: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
            
            Dim 원래보호 As Boolean: 원래보호 = ws재고.ProtectContents
            If 원래보호 Then On Error Resume Next: ws재고.Unprotect "0205": If ws재고.ProtectContents Then ws재고.Unprotect: On Error GoTo 0
            내부프로파일 = 내부프로파일 & "  3b.Unprotect: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
            
            ' ── 각 제품 처리 (19초 버전: 제품마다 셀쓰기 + 빠른_제품재고_재계산) ──
            For Each key In 제품딕.keys
                Dim 제품키 As String: 제품키 = CStr(key)
                Dim 총생산량 As Double: 총생산량 = 0
                If IsNumeric(제품딕(key)) Then 총생산량 = CDbl(제품딕(key))
                ' 맵키 = 업체|품명 (재고 시트 그룹 매칭용). 제품키 형식: 구분|품명|순번 → 품명으로 업체 조회
                Dim 구분값 As String, 품명값 As String
                Dim parts() As String: parts = Split(제품키, "|")
                If UBound(parts) >= 1 Then
                    구분값 = Trim(parts(0))
                    품명값 = Trim(parts(1))
                Else
                    구분값 = "": 품명값 = 제품키
                End If
                Dim v업체 As String
                If 업체명딕.exists(품명값) Then v업체 = Trim(CStr(업체명딕(품명값))) Else v업체 = ""
                If v업체 <> "" Then 맵키 = v업체 & "|" & 품명값 Else 맵키 = "|" & 품명값
                Dim 매칭수 As Long
                If 제품맵.exists(맵키) Then: 매칭수 = 제품맵(맵키).Count: Else: 매칭수 = 0
                
                ' ── CASE A: 미등록 → 신규 그룹 1개 (A=제품발주처(업체명), B=제품명은 빈행그룹_추가_무음에서 이미 입력) ──
                If 매칭수 = 0 Then
                    Dim 새그룹시작 As Long
                    새그룹시작 = 빈행그룹_추가_무음(품명값)
                    If 새그룹시작 > 0 Then
                        Dim 신규업체 As String
                        If 업체명딕.exists(품명값) Then 신규업체 = Trim(CStr(업체명딕(품명값))) Else 신규업체 = "-"
                        ws재고.Cells(새그룹시작, INV_COL_제품발주처).Value = 신규업체
                        재고lastRow = ws재고.Cells(ws재고.Rows.Count, INV_COL_구분).End(xlUp).Row
                        ws재고.Cells(새그룹시작, 날짜열).Value = 총생산량
                        If rngFormat Is Nothing Then: Set rngFormat = ws재고.Cells(새그룹시작, 날짜열): Else: Set rngFormat = Union(rngFormat, ws재고.Cells(새그룹시작, 날짜열))
                        Call 빠른_제품재고_재계산(ws재고, 새그룹시작, INV_COL_일별시작, 마지막열, INV_COL_전월재고, INV_COL_현재고, 현재고_열idx, 날짜열)
                        재고신규추가목록.Add 제품키 & " -> " & Format(총생산량, "#,##0") & " BOX (신규그룹)"
                        제품상태딕.Add 제품키, "신규"
                    Else
                        재고실패목록.Add 제품키 & " (그룹 추가 실패)"
                        제품상태딕.Add 제품키, "실패"
                    End If
                
                ' ── CASE B: 1개 → 바로 ──
                ElseIf 매칭수 = 1 Then
                    Dim 단일정보 As Variant: 단일정보 = 제품맵(맵키)(1)
                    Dim 단일그룹 As Long: 단일그룹 = 0
                    Dim 첫요소 As Variant: 첫요소 = Empty
                    If IsArray(단일정보) Then 첫요소 = 단일정보(LBound(단일정보))
                    If IsNumeric(첫요소) Then 단일그룹 = CLng(CDbl(첫요소))
                    If 단일그룹 < INV_START_ROW Then GoTo NextProduct
                    ws재고.Cells(단일그룹, 날짜열).Value = 총생산량
                    If rngFormat Is Nothing Then: Set rngFormat = ws재고.Cells(단일그룹, 날짜열): Else: Set rngFormat = Union(rngFormat, ws재고.Cells(단일그룹, 날짜열))
                    Call 빠른_제품재고_재계산(ws재고, 단일그룹, INV_COL_일별시작, 마지막열, INV_COL_전월재고, INV_COL_현재고, 현재고_열idx, 날짜열)
                    재고업데이트목록.Add 제품키 & " -> " & Format(총생산량, "#,##0") & " BOX (" & Format(작성일, "MM/DD") & ")"
                    제품상태딕.Add 제품키, "자동"
                
                ' ── CASE C: 2개+ → InputBox ──
                Else
                    Application.ScreenUpdating = True
                    Dim 매칭그룹() As Long, 매칭업체() As String
                    ReDim 매칭그룹(1 To 매칭수): ReDim 매칭업체(1 To 매칭수)
                    For j = 1 To 매칭수
                        Dim tmpInfo As Variant: tmpInfo = 제품맵(맵키)(j)
                        Dim tmpFirst As Variant, tmpSecond As Variant
                        If IsArray(tmpInfo) Then
                            tmpFirst = tmpInfo(LBound(tmpInfo))
                            tmpSecond = tmpInfo(UBound(tmpInfo))
                        Else
                            tmpFirst = Empty: tmpSecond = ""
                        End If
                        If IsNumeric(tmpFirst) Then 매칭그룹(j) = CLng(CDbl(tmpFirst)) Else 매칭그룹(j) = 0
                        매칭업체(j) = 재고_셀값문자열(tmpSecond)
                    Next j
                    
                    Dim 선택안내 As String
                    선택안내 = "[" & 제품키 & "] 생산량: " & Format(총생산량, "#,##0") & " BOX" & vbCrLf & vbCrLf & _
                            "동일 제품이 " & 매칭수 & "개 그룹:" & vbCrLf & vbCrLf
                    For j = 1 To 매칭수
                        선택안내 = 선택안내 & "  " & j & ". " & 매칭업체(j) & vbCrLf
                    Next j
                    선택안내 = 선택안내 & vbCrLf & "선택:번호 | 분할:1=300,2=200 | 전체:A"
                    
                    Dim 사용자입력 As String
                    사용자입력 = InputBox(선택안내, "생산값 입력 - " & 제품키, "1")
                    Application.ScreenUpdating = False
                    
                    If StrPtr(사용자입력) = 0 Then GoTo NextProduct
                    사용자입력 = Trim(사용자입력)
                    If 사용자입력 = "" Then GoTo NextProduct
                    
                    If UCase(사용자입력) = "A" Then
                        For j = 1 To 매칭수
                            Dim grpA As Long: grpA = 매칭그룹(j)
                            ws재고.Cells(grpA, 날짜열).Value = 총생산량
                            If rngFormat Is Nothing Then: Set rngFormat = ws재고.Cells(grpA, 날짜열): Else: Set rngFormat = Union(rngFormat, ws재고.Cells(grpA, 날짜열))
                            Call 빠른_제품재고_재계산(ws재고, grpA, INV_COL_일별시작, 마지막열, INV_COL_전월재고, INV_COL_현재고, 현재고_열idx, 날짜열)
                        Next j
                        재고업데이트목록.Add 제품키 & " -> " & Format(총생산량, "#,##0") & " BOX x" & 매칭수 & "그룹 (" & Format(작성일, "MM/DD") & ")"
                        제품상태딕.Add 제품키, "전체"
                    
                    ElseIf InStr(사용자입력, "=") > 0 Then
                        Dim 분할항목() As String: 분할항목 = Split(사용자입력, ",")
                        Dim 분할건수 As Long: 분할건수 = 0
                        For j = 0 To UBound(분할항목)
                            Dim 분할조각() As String: 분할조각 = Split(Trim(분할항목(j)), "=")
                            If UBound(분할조각) >= 1 Then
                                Dim 분할번호 As Long, 분할수량 As Double
                                If IsNumeric(Trim(분할조각(0))) And IsNumeric(Trim(분할조각(1))) Then
                                    분할번호 = CLng(Trim(분할조각(0))): 분할수량 = CDbl(Trim(분할조각(1)))
                                    If 분할번호 >= 1 And 분할번호 <= 매칭수 Then
                                        Dim grpS As Long: grpS = 매칭그룹(분할번호)
                                        ws재고.Cells(grpS, 날짜열).Value = 분할수량
                                        If rngFormat Is Nothing Then: Set rngFormat = ws재고.Cells(grpS, 날짜열): Else: Set rngFormat = Union(rngFormat, ws재고.Cells(grpS, 날짜열))
                                        Call 빠른_제품재고_재계산(ws재고, grpS, INV_COL_일별시작, 마지막열, INV_COL_전월재고, INV_COL_현재고, 현재고_열idx, 날짜열)
                                        재고업데이트목록.Add 제품키 & "[" & 매칭업체(분할번호) & "] -> " & Format(분할수량, "#,##0") & " BOX (" & Format(작성일, "MM/DD") & ")"
                                        분할건수 = 분할건수 + 1
                                    End If
                                End If
                            End If
                        Next j
                        If 분할건수 > 0 Then 제품상태딕.Add 제품키, "분할"
                    
                    ElseIf IsNumeric(사용자입력) Then
                        Dim 선택번호 As Long: 선택번호 = CLng(사용자입력)
                        If 선택번호 >= 1 And 선택번호 <= 매칭수 Then
                            Dim grpC As Long: grpC = 매칭그룹(선택번호)
                            ws재고.Cells(grpC, 날짜열).Value = 총생산량
                            If rngFormat Is Nothing Then: Set rngFormat = ws재고.Cells(grpC, 날짜열): Else: Set rngFormat = Union(rngFormat, ws재고.Cells(grpC, 날짜열))
                            Call 빠른_제품재고_재계산(ws재고, grpC, INV_COL_일별시작, 마지막열, INV_COL_전월재고, INV_COL_현재고, 현재고_열idx, 날짜열)
                            재고업데이트목록.Add 제품키 & "[" & 매칭업체(선택번호) & "] -> " & Format(총생산량, "#,##0") & " BOX (" & Format(작성일, "MM/DD") & ")"
                            제품상태딕.Add 제품키, "선택"
                        End If
                    End If
                End If
NextProduct:
            Next key
            
            내부프로파일 = 내부프로파일 & "  3c.제품루프(값+재계산): " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
            
            ' ★ 음영 서식 일괄 적용 (1회)
            If Not rngFormat Is Nothing Then
                With rngFormat
                    .Interior.color = RGB(192, 135, 253)
                    .Font.color = RGB(255, 255, 255)
                End With
            End If
            내부프로파일 = 내부프로파일 & "  3d.서식적용: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
            
            If 원래보호 Then
                On Error Resume Next
                Call 셀보호_적용_무음(ws재고)
                On Error GoTo ErrorHandler
            End If
            내부프로파일 = 내부프로파일 & "  3e.셀보호: " & Format(Timer - tp1, "0.000") & "s" & vbCrLf: tp1 = Timer
        End If
        
        ' ??????????????????????????????????????????????????????
        ' 4단계: 업데이트 내역 MsgBox
        ' ??????????????????????????????????????????????????????
        진행단계 = "4단계: 결과 메시지"
        Dim 메시지 As String
        메시지 = "[ 완제품 생산이력 업데이트 완료 ]" & vbCrLf & _
                "작성일: " & Format(작성일, "YYYY-MM-DD") & vbCrLf & _
                "수집 제품: " & 제품딕.Count & "건" & vbCrLf & vbCrLf
        
        If Raw추가목록.Count > 0 Then
            메시지 = 메시지 & "▶ Raw data_index 신규: " & Raw추가목록.Count & "건" & vbCrLf
            For i = 1 To Raw추가목록.Count
                메시지 = 메시지 & "  - " & Raw추가목록(i) & vbCrLf
            Next i
            메시지 = 메시지 & vbCrLf
        Else
            메시지 = 메시지 & "▶ Raw data_index: 기존 등록 완료" & vbCrLf & vbCrLf
        End If
        
        If 재고신규추가목록.Count > 0 Then
            메시지 = 메시지 & "▶ 부자재관리 신규 그룹: " & 재고신규추가목록.Count & "건" & vbCrLf
            For i = 1 To 재고신규추가목록.Count
                메시지 = 메시지 & "  - " & 재고신규추가목록(i) & vbCrLf
            Next i
            메시지 = 메시지 & vbCrLf
        End If
        
        If 재고업데이트목록.Count > 0 Then
            메시지 = 메시지 & "▶ 부자재관리 반영: " & 재고업데이트목록.Count & "건" & vbCrLf
            For i = 1 To 재고업데이트목록.Count
                메시지 = 메시지 & "  - " & 재고업데이트목록(i) & vbCrLf
            Next i
            메시지 = 메시지 & vbCrLf
        End If
        
        If 재고실패목록.Count > 0 Then
            메시지 = 메시지 & "▶ 처리 실패: " & 재고실패목록.Count & "건" & vbCrLf
            For i = 1 To 재고실패목록.Count
                메시지 = 메시지 & "  - " & 재고실패목록(i) & vbCrLf
            Next i
        End If
        
        ' ── Application 상태 복원 ──
        Application.Calculation = 원래Calculation
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        내부프로파일 = 내부프로파일 & "  전체: " & Format(Timer - tp0, "0.000") & "s"
        
        ' ★ 무음모드: MsgBox/워크북 생성 없이 즉시 종료 (증분_업데이트에서 체인 호출 시)
        If 무음모드 Then
            m_재고프로파일 = 내부프로파일  ' 모듈 변수로 전달
            Exit Sub
        End If
        
        ' ── 단독 실행 시: 프로파일 포함 MsgBox ──
        MsgBox 메시지 & vbCrLf & vbCrLf & "── 속도 프로파일 ──" & vbCrLf & 내부프로파일, _
            vbInformation, "완제품 생산이력 업데이트 완료"
        Exit Sub
        
ErrorHandler:
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox "완제품 생산이력 업데이트 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
            "진행 단계: " & 진행단계 & vbCrLf & _
            "오류 번호: " & Err.Number & vbCrLf & _
            "오류 내용: " & Err.Description, vbCritical, "생산이력 업데이트 오류"
    End Sub

    ' 셀 값 → 문자열 (Error/Empty 시 형식 불일치 방지)
    Private Function 재고_셀값문자열(v As Variant) As String
        If IsEmpty(v) Then 재고_셀값문자열 = "": Exit Function
        If IsError(v) Then 재고_셀값문자열 = "": Exit Function
        On Error Resume Next
        재고_셀값문자열 = Trim(CStr(v))
        If Err.Number <> 0 Then 재고_셀값문자열 = ""
        On Error GoTo 0
    End Function

    ' ============================================================
    ' ★★ 부분 재계산: 수정열~끝만 계산 (열 확장에 대응)
    '
    ' [핵심 원리] 생산값을 입력한 날짜열(수정열) 이전의 재고는
    ' 이미 정확하므로, 수정열부터만 계산하면 됨.
    '
    ' [성능 비교] 12개월(372열) 시트, 2월9일(40번째열) 수정 시:
    '   기존: 372열 전체 스캔 = 1,488 셀 처리
    '   최적: 40열~372열 = 332열만 처리 (11% 절약)
    '   BUT 실사용: 대부분 당월 마지막~수정열 = 20열 내외 (95% 절약)
    '   ※ 0값 연속 시 조기종료로 추가 절약
    ' ============================================================
    Private Sub 빠른_제품재고_재계산(ws As Worksheet, 그룹시작행 As Long, _
        COL_시작 As Long, COL_끝 As Long, COL_전월 As Long, COL_현재고 As Long, _
        현재고열 As Long, Optional 수정열 As Long = 0)
        
        On Error Resume Next
        
        ' ── 재계산 시작점 결정 ──
        Dim 계산시작 As Long
        If 수정열 > 0 And 수정열 >= COL_시작 Then
            계산시작 = 수정열
        Else
            계산시작 = COL_시작   ' 수정열 미지정 → 전체 재계산
        End If
        
        Dim 열수 As Long: 열수 = COL_끝 - 계산시작 + 1
        If 열수 < 1 Then Exit Sub
        
        ' ── 전일재고 결정 (1회 읽기) ──
        Dim 전일재고 As Double: 전일재고 = 0
        If 계산시작 = COL_시작 Then
            ' 첫 열 → 전월재고 사용
            If IsNumeric(ws.Cells(그룹시작행, COL_전월).Value) Then
                전일재고 = CDbl(ws.Cells(그룹시작행, COL_전월).Value)
            End If
        Else
            ' 이전 열의 기존 재고값 사용 (이미 정확한 값)
            If IsNumeric(ws.Cells(그룹시작행 + 3, 계산시작 - 1).Value) Then
                전일재고 = CDbl(ws.Cells(그룹시작행 + 3, 계산시작 - 1).Value)
            End If
        End If
        
        ' ── 수정열~끝 범위만 배열 읽기 (1회) ──
        Dim arr As Variant
        arr = ws.Range(ws.Cells(그룹시작행, 계산시작), ws.Cells(그룹시작행 + 3, COL_끝)).Value
        ' 단일 셀 시 스칼라 반환 → 형식 불일치(Error 13) 방지
        If Not IsArray(arr) Then
            Dim arr1(1 To 4, 1 To 1) As Variant
            arr1(1, 1) = arr: arr1(2, 1) = 0: arr1(3, 1) = 0: arr1(4, 1) = 0
            arr = arr1
        End If
        ' arr(1,c)=생산, arr(2,c)=입고, arr(3,c)=소요, arr(4,c)=재고
        
        ' ── 메모리에서 재고 계산 + 0값 연속 조기종료 ──
        Dim 생산 As Double, 입고 As Double, 소요 As Double, 재고 As Double
        Dim cc As Long, 변경없음연속 As Long, 기존재고 As Double
        Dim 쓰기끝 As Long: 쓰기끝 = 열수   ' 실제 쓰기 범위 (조기종료 시 축소)
        
        재고 = 전일재고
        For cc = 1 To 열수
            생산 = 0: 입고 = 0: 소요 = 0
            If IsNumeric(arr(1, cc)) Then 생산 = CDbl(arr(1, cc))
            If IsNumeric(arr(2, cc)) Then 입고 = CDbl(arr(2, cc))
            If IsNumeric(arr(3, cc)) Then 소요 = CDbl(arr(3, cc))
            
            재고 = 재고 + 생산 + 입고 - 소요
            
            ' 기존 재고값과 비교 → 변경 없으면 카운트
            기존재고 = 0
            If IsNumeric(arr(4, cc)) Then 기존재고 = CDbl(arr(4, cc))
            
            arr(4, cc) = 재고
            
            If Abs(재고 - 기존재고) < 0.001 Then
                변경없음연속 = 변경없음연속 + 1
                ' 3열 연속 변경 없으면 → 이후도 동일하므로 조기종료
                If 변경없음연속 >= 3 And cc > 1 Then
                    쓰기끝 = cc
                    Exit For
                End If
            Else
                변경없음연속 = 0
            End If
        Next cc
        
        ' ── 변경된 범위만 쓰기 ──
        If 쓰기끝 = 1 Then
            ws.Cells(그룹시작행 + 3, 계산시작).Value = arr(4, 1)
        Else
            Dim arrOut() As Variant
            ReDim arrOut(1 To 1, 1 To 쓰기끝)
            For cc = 1 To 쓰기끝
                arrOut(1, cc) = arr(4, cc)
            Next cc
            ws.Range(ws.Cells(그룹시작행 + 3, 계산시작), _
                    ws.Cells(그룹시작행 + 3, 계산시작 + 쓰기끝 - 1)).Value = arrOut
        End If
        
        ' ── 현재고 업데이트 ──
        Dim 현재고값 As Double: 현재고값 = 0
        If 현재고열 >= 계산시작 And 현재고열 <= 계산시작 + 쓰기끝 - 1 Then
            ' 계산 범위 내에 현재고열이 있음
            Dim idxCur As Long: idxCur = 현재고열 - 계산시작 + 1
            If idxCur >= 1 And idxCur <= 쓰기끝 And IsNumeric(arr(4, idxCur)) Then 현재고값 = CDbl(arr(4, idxCur))
        ElseIf 현재고열 < 계산시작 Then
            ' 현재고열이 계산 시작점 이전 → 기존값 그대로 (변경 불필요)
            GoTo SkipCurrent
        Else
            ' 현재고열이 계산 범위 이후 (조기종료) → 마지막 계산된 재고 사용
            현재고값 = 재고
        End If
        
        Dim arrCur(1 To 4, 1 To 1) As Variant
        arrCur(1, 1) = 현재고값: arrCur(2, 1) = 현재고값
        arrCur(3, 1) = 현재고값: arrCur(4, 1) = 현재고값
        ws.Range(ws.Cells(그룹시작행, COL_현재고), ws.Cells(그룹시작행 + 3, COL_현재고)).Value = arrCur

SkipCurrent:
        On Error GoTo 0
    End Sub

    ' ============================================================
    ' ★ 진단 함수: 각 단계별 정상 작동 여부 확인
    ' Alt+F8 → 생산이력_진단 실행
    ' ============================================================
    Public Sub 생산이력_진단()
        Dim 결과 As String
        결과 = "=== 완제품 생산이력 진단 ===" & vbCrLf & vbCrLf
        
        ' ── 1. 시트 존재 확인 ──
        Dim ws일보 As Worksheet, wsRaw As Worksheet, ws재고 As Worksheet
        On Error Resume Next
        Set ws일보 = ThisWorkbook.Worksheets("생산일보")
        Set wsRaw = ThisWorkbook.Worksheets("Raw data_index")
        Set ws재고 = ThisWorkbook.Worksheets("부자재관리")
        On Error GoTo 0
        
        결과 = 결과 & "[시트 확인]" & vbCrLf
        결과 = 결과 & "  일일생산일보: " & IIf(ws일보 Is Nothing, "X 없음!", "O 존재") & vbCrLf
        결과 = 결과 & "  Raw data_index: " & IIf(wsRaw Is Nothing, "X 없음!", "O 존재") & vbCrLf
        결과 = 결과 & "  부자재관리: " & IIf(ws재고 Is Nothing, "X 없음!", "O 존재") & vbCrLf & vbCrLf
        
        If ws일보 Is Nothing Then
            MsgBox 결과 & "일일생산일보 시트가 없어서 진단 중단", vbCritical, "진단 결과"
            Exit Sub
        End If
        
        ' ── 2. 날짜(BK4) 확인 ──
        Dim ajVal As Variant
        On Error Resume Next
        ajVal = ws일보.Range("BK4").Value
        On Error GoTo 0
        결과 = 결과 & "[날짜 확인]" & vbCrLf
        결과 = 결과 & "  BK4 값: " & CStr(ajVal) & vbCrLf
        결과 = 결과 & "  날짜여부: " & IIf(IsDate(ajVal), "O (" & Format(CDate(ajVal), "YYYY-MM-DD") & ")", "X 날짜아님!") & vbCrLf & vbCrLf
        
        ' ── 3. 제품 수집 테스트 ──
        결과 = 결과 & "[제품 수집 테스트]" & vbCrLf
        Dim 제품딕 As Object: Set 제품딕 = CreateObject("Scripting.Dictionary")
        Dim r As Long, 품명 As String, 생산량 As Double
        
        ' 영역 I: B21~ C열
        Dim cnt1 As Long: cnt1 = 0
        For r = 21 To 43
            품명 = Trim(CStr(ws일보.Cells(r, 3).Value))
            If 품명 <> "" And UCase(품명) <> "TOTAL" And 품명 <> "계" And Not IsNumeric(품명) Then
                생산량 = 0
                If IsNumeric(ws일보.Cells(r, 17).Value) Then 생산량 = CDbl(ws일보.Cells(r, 17).Value)
                If 생산량 <> 0 Then
                    If Not 제품딕.exists(품명) Then 제품딕.Add 품명, 생산량 Else 제품딕(품명) = CDbl(제품딕(품명)) + 생산량
                    cnt1 = cnt1 + 1
                End If
            End If
        Next r
        결과 = 결과 & "  영역I (B21~C43, Q열): " & cnt1 & "건" & vbCrLf
        
        ' 처음 몇 개 셀값 표시 (디버그용)
        결과 = 결과 & "    B21=" & CStr(ws일보.Cells(21, 2).Value) & ", C21=" & CStr(ws일보.Cells(21, 3).Value) & ", Q21=" & CStr(ws일보.Cells(21, 17).Value) & vbCrLf
        결과 = 결과 & "    B22=" & CStr(ws일보.Cells(22, 2).Value) & ", C22=" & CStr(ws일보.Cells(22, 3).Value) & ", Q22=" & CStr(ws일보.Cells(22, 17).Value) & vbCrLf
        
        ' 영역 II: AE44~ AG열
        Dim cnt2 As Long: cnt2 = 0
        For r = 44 To 70
            품명 = Trim(CStr(ws일보.Cells(r, 33).Value))
            If 품명 <> "" And UCase(품명) <> "TOTAL" And 품명 <> "계" And Not IsNumeric(품명) Then
                생산량 = 0
                If IsNumeric(ws일보.Cells(r, 41).Value) Then 생산량 = CDbl(ws일보.Cells(r, 41).Value)
                If 생산량 <> 0 Then
                    If Not 제품딕.exists(품명) Then 제품딕.Add 품명, 생산량 Else 제품딕(품명) = CDbl(제품딕(품명)) + 생산량
                    cnt2 = cnt2 + 1
                End If
            End If
        Next r
        결과 = 결과 & "  영역II (AE44~AG70, AO열): " & cnt2 & "건" & vbCrLf
        
        ' 처음 몇 개 셀값 표시 (디버그용)
        결과 = 결과 & "    AE44=" & CStr(ws일보.Cells(44, 31).Value) & ", AG44=" & CStr(ws일보.Cells(44, 33).Value) & ", AO44=" & CStr(ws일보.Cells(44, 41).Value) & vbCrLf
        
        결과 = 결과 & "  총 수집: " & 제품딕.Count & "건" & vbCrLf
        If 제품딕.Count > 0 Then
            Dim key As Variant
            For Each key In 제품딕.keys
                결과 = 결과 & "    " & CStr(key) & " = " & Format(CDbl(제품딕(key)), "#,##0") & " BOX" & vbCrLf
            Next key
        End If
        결과 = 결과 & vbCrLf
        
        ' ── 4. 부자재관리 날짜열 확인 ──
        If Not ws재고 Is Nothing And IsDate(ajVal) Then
            결과 = 결과 & "[부자재관리 날짜열 확인]" & vbCrLf
            Dim 작성일 As Date: 작성일 = CDate(ajVal)
            Dim 대상월 As Integer: 대상월 = Month(작성일)
            Dim 대상일 As Integer: 대상일 = Day(작성일)
            Dim 마지막열 As Long: 마지막열 = ws재고.Cells(3, ws재고.Columns.Count).End(xlToLeft).Column
            Dim 날짜열 As Long: 날짜열 = 0
            Dim c As Long
            
            결과 = 결과 & "  검색 조건: " & 대상월 & "m / " & 대상일 & "일" & vbCrLf
            결과 = 결과 & "  검색 범위: N열(14) ~ " & 마지막열 & "열" & vbCrLf
            
            For c = 14 To 마지막열
                Dim v월 As String, v일 As Variant
                v월 = Replace(Replace(CStr(ws재고.Cells(2, c).Value), "m", ""), "M", "")
                v일 = ws재고.Cells(3, c).Value
                If IsNumeric(v월) And IsNumeric(v일) Then
                    If CLng(val(v월)) = 대상월 And CLng(val(v일)) = 대상일 Then
                        날짜열 = c
                        Exit For
                    End If
                End If
            Next c
            
            결과 = 결과 & "  매칭 열: " & IIf(날짜열 > 0, "O (" & 날짜열 & "열)", "X 없음!") & vbCrLf
            
            ' 2행/3행의 처음 몇 열 표시
            결과 = 결과 & "  2행(월): "
            For c = 14 To Application.Min(마지막열, 20)
                결과 = 결과 & "[" & c & "]" & CStr(ws재고.Cells(2, c).Value) & " "
            Next c
            결과 = 결과 & vbCrLf
            결과 = 결과 & "  3행(일): "
            For c = 14 To Application.Min(마지막열, 20)
                결과 = 결과 & "[" & c & "]" & CStr(ws재고.Cells(3, c).Value) & " "
            Next c
            결과 = 결과 & vbCrLf & vbCrLf
        End If
        
        ' ── 5. 외부 함수 접근 확인 ──
        결과 = 결과 & "[외부 함수 접근 확인]" & vbCrLf
        On Error Resume Next
        Dim 테스트결과 As Long
        테스트결과 = 빈행그룹_추가_무음("")
        If Err.Number <> 0 Then
            결과 = 결과 & "  빈행그룹_추가_무음: X 오류 (" & Err.Number & ": " & Err.Description & ")" & vbCrLf
            Err.Clear
        Else
            결과 = 결과 & "  빈행그룹_추가_무음: O 접근가능 (반환=" & 테스트결과 & ")" & vbCrLf
        End If
        On Error GoTo 0
        결과 = 결과 & vbCrLf
        
        ' ── 6. 현재 버튼 매크로 확인 안내 ──
        결과 = 결과 & "[중요! 버튼 매크로 확인]" & vbCrLf
        결과 = 결과 & "  DB저장 버튼에 'DB저장_및_재고반영' 매크로가 연결되어 있는지 확인하세요." & vbCrLf
        결과 = 결과 & "  (버튼 우클릭 → 매크로 지정 → 'DB저장_및_재고반영' 선택)" & vbCrLf
        
        MsgBox 결과, vbInformation, "생산이력 진단 결과"
    End Sub









