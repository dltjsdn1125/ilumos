Attribute VB_Name = "Module21"
' ============================================================
' 이전 생산관리 파일 → 일일생산DB 이관 VBA (2월 전용)
' - 1월대비: 시트별 행 오프셋 동적 감지 (행 밀림 대응)
' - 주야간: 6:4 랜덤 대신, 원지 음영(RGB 255,230,153) 구간만 5:5 주/야, 비음영=전부 주간
' - v8: AJ11→F16, AJ13→F17. 인원F19~25, PA F28~
' ============================================================
Option Explicit

Public Sub 이전데이터_이관_실행_FEB()

    ' ──────────────────────────────────────────────────────
    ' [변수 선언 - 전부 프로시저 최상단]
    ' ──────────────────────────────────────────────────────
    Dim srcFile     As String
    Dim wbSrc       As Workbook
    Dim wbDst       As Workbook
    Dim wsDB        As Worksheet
    Dim ws          As Worksheet
    Dim 시트명       As String
    Dim sheetList(30) As String
    Dim sheetCount  As Long
    Dim si          As Long, sj As Long, temp As String

    Dim batchArr()  As Variant
    Dim batchSize   As Long
    Dim seedVal     As Long
    Dim ID카운터    As Long
    Dim 마지막행    As Long
    Dim lastID      As Variant
    Dim 새행시작    As Long
    Dim 추가수      As Long
    Dim errNum      As Long
    Dim errDesc     As String
    Dim i           As Long

    ' 루프/수집용 공통
    Dim r           As Long
    Dim t           As Long
    Dim 순번        As Long
    Dim 작성일      As Variant

    ' 원지 생산일보I
    Dim 종별_I(5)   As String
    Dim ROLLID_I(5) As Long
    Dim KG_I(5)     As Long
    Dim rollVal     As Variant, kgVal As Variant
    Dim rollStr     As String
    Dim 구분I       As String

    ' 원지 JRT-1/2
    Dim jrtSeq      As Long
    Dim jrt2Seq     As Long
    Dim 구분JRT     As String
    Dim roll2       As Variant, kg2 As Variant
    Dim roll2Str    As String
    Dim 구분JRT2    As String

    ' 원지 FT/BT/PRE
    Dim 종별_II(2)   As String
    Dim ROLLID_II(2) As Long
    Dim KG_II(2)     As Long
    Dim kgStr        As String
    Dim 구분II       As String

    ' 완제품I
    Dim prev_가공라인 As String
    Dim 품명          As String
    Dim 가공라인      As String
    Dim 구분완I       As String

    ' 완제품II
    Dim 현재구분     As String
    Dim 구분셀       As String
    Dim 생산량II     As Variant
    Dim 호기II       As String
    Dim 구분완II     As String

    ' 투입인원I
    Dim 구분인원(1)  As String
    Dim 행인원(1)    As Long
    Dim 인원수       As Variant
    Dim 작업시간     As Variant
    Dim 인원유효     As Boolean
    Dim 시간유효     As Boolean

    ' 투입인원II
    Dim 현재작업장   As String
    Dim 구분맵       As String
    Dim 작업장셀     As String
    Dim 구분4        As String
    Dim hasData2     As Boolean
    Dim c2           As Long
    Dim v정규        As Variant
    Dim v연장        As Variant
    Dim 인원II인원(1 To 10) As Variant
    Dim 인원II정규(1 To 10) As Variant
    Dim 인원II연장(1 To 10) As Variant
    Dim 인원II작업장(1 To 10) As String
    Dim 인원II총MH(1 To 10) As Variant
    Dim 인원II실투입(1 To 10) As Variant
    Dim 인원II환산(1 To 10) As Variant
    Dim 인원II인당실(1 To 10) As Variant
    Dim 인원II인당환(1 To 10) As Variant
    Dim 인원II갯수   As Long
    Dim idx          As Long

    ' PA (KT 수율 지표)
    Dim pa행(6)     As Long
    Dim vII         As Variant
    Dim 일원(0 To 4) As Double
    Dim 일양(0 To 4) As Double
    Dim 일기소(0 To 4) As Double
    Dim eqIdx As Long

    ' 2월 전용: 시트별 동적 행 오프셋·원지 확장
    Dim 행오프셋    As Long
    Dim 행확장      As Long   ' 원지I 추가행으로 인한 하단 블록 밀림
    Dim 원지I끝     As Long   ' 원지 생산일보I 실제 끝행
    Dim rBase       As Long

    ' ──────────────────────────────────────────────────────
    ' 파일 선택
    ' ──────────────────────────────────────────────────────
    srcFile = Application.GetOpenFilename("Excel 파일 (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "이전 생산관리 파일 선택")
    If srcFile = "False" Then Exit Sub

    Set wbSrc = Workbooks.Open(srcFile, ReadOnly:=True)
    Set wbDst = ThisWorkbook

    On Error Resume Next
    Set wsDB = wbDst.Worksheets("일일생산DB")
    On Error GoTo 0

    If wsDB Is Nothing Then
        ' DB 시트 없으면 자동 생성
        Dim ans As VbMsgBoxResult
        ans = MsgBox("현재 파일에 '일일생산DB' 시트가 없습니다." & vbCrLf & _
                     "지금 새로 만들겠습니까?", vbQuestion + vbYesNo, "DB 시트 없음")
        If ans = vbNo Then
            wbSrc.Close False
            Exit Sub
        End If
        Set wsDB = wbDst.Worksheets.Add(After:=wbDst.Worksheets(wbDst.Worksheets.Count))
        wsDB.Name = "일일생산DB"
        ' 기본 헤더 126열 생성
        Dim hdrArr(1 To 126) As String
        ' 공통
        hdrArr(1) = "작성일": hdrArr(2) = "ID": hdrArr(3) = "구분": hdrArr(4) = "레코드_유형": hdrArr(5) = "순번"
        ' 원지(6~8)
        hdrArr(6) = "원지_종별": hdrArr(7) = "원지_ROLLID": hdrArr(8) = "원지_KG"
        ' 완제품(9~16)
        hdrArr(9) = "완제품_구분": hdrArr(10) = "완제품_품명": hdrArr(11) = "가공라인호기"
        hdrArr(12) = "생산량_BOX": hdrArr(13) = "원지기준소요량": hdrArr(14) = "PLT1": hdrArr(15) = "PLT2": hdrArr(16) = "인원"
        ' 투입인원(17~25)
        hdrArr(17) = "작업장": hdrArr(18) = "구분_인원": hdrArr(19) = "인원수": hdrArr(20) = "작업시간"
        hdrArr(21) = "총작업시간": hdrArr(22) = "실투입원": hdrArr(23) = "환산원": hdrArr(24) = "인당실투입": hdrArr(25) = "인당환산"
        ' PA KT 일계(26~31)
        hdrArr(26) = "KT_일계_원투": hdrArr(27) = "KT_일계_양품": hdrArr(28) = "KT_일계_수율"
        hdrArr(29) = "KT_일계_기소": hdrArr(30) = "KT_일계_LOSS": hdrArr(31) = "KT_일계_양품율"
        ' PA KT 주누계(32~37)
        hdrArr(32) = "KT_주누계_원투": hdrArr(33) = "KT_주누계_양품": hdrArr(34) = "KT_주누계_수율"
        hdrArr(35) = "KT_월누계_원투": hdrArr(36) = "KT_월누계_양품": hdrArr(37) = "KT_월누계_수율"
        ' PA FT/JRT/BT 일계(38~46)
        hdrArr(38) = "FT_일계_원투": hdrArr(39) = "FT_일계_양품": hdrArr(40) = "FT_일계_수율"
        hdrArr(41) = "JRT_일계_원투": hdrArr(42) = "JRT_일계_양품": hdrArr(43) = "JRT_일계_수율"
        hdrArr(44) = "BT_일계_원투": hdrArr(45) = "BT_일계_양품": hdrArr(46) = "BT_일계_수율"
        ' 나머지 47~100 공통 헤더
        Dim hc As Long
        For hc = 47 To 100
            hdrArr(hc) = "F" & hc
        Next hc
        ' 기소 열 101~124
        hdrArr(101) = "KT_주누계_기소": hdrArr(102) = "KT_월누계_기소"
        hdrArr(103) = "KT_주평균_기소": hdrArr(104) = "KT_월평균_기소"
        hdrArr(105) = "FT_주누계_기소": hdrArr(106) = "FT_월누계_기소"
        hdrArr(107) = "FT_주평균_기소": hdrArr(108) = "FT_월평균_기소"
        hdrArr(109) = "JRT_주누계_기소": hdrArr(110) = "JRT_월누계_기소"
        hdrArr(111) = "JRT_주평균_기소": hdrArr(112) = "JRT_월평균_기소"
        hdrArr(113) = "BT_주누계_기소": hdrArr(114) = "BT_월누계_기소"
        hdrArr(115) = "BT_주평균_기소": hdrArr(116) = "BT_월평균_기소"
        hdrArr(117) = "PRE_주누계_기소": hdrArr(118) = "PRE_월누계_기소"
        hdrArr(119) = "PRE_주평균_기소": hdrArr(120) = "PRE_월평균_기소"
        hdrArr(121) = "FT_일계_기소": hdrArr(122) = "JRT_일계_기소"
        hdrArr(123) = "BT_일계_기소": hdrArr(124) = "PRE_일계_기소"
        hdrArr(125) = "완제품_AJ11": hdrArr(126) = "완제품_AJ13"
        ' 헤더 쓰기
        Dim hRng As Range
        Set hRng = wsDB.Range(wsDB.Cells(1, 1), wsDB.Cells(1, 126))
        Dim hv As Long
        For hv = 1 To 126: wsDB.Cells(1, hv).Value = hdrArr(hv): Next hv
        With hRng
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .HorizontalAlignment = xlCenter
        End With
        wsDB.Rows(1).AutoFit
    End If

    ' 역변환(Module4/10/11)과 완전히 동일한 126열 헤더를 항상 강제 적용
    ' 2/3월 모듈의 자체 헤더 정의는 Module4 스키마와 달라 누계/평균 역변환을 깨뜨린다.
    Call 헤더_생성(wsDB)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo MigErr

    ' DB 헤더 보완 (col 125=완제품_AJ11, col 126=완제품_AJ13 없으면 추가) - 인라인
    Dim hIdx As Long
    Dim hCols(1) As Long:   hCols(0) = 125:          hCols(1) = 126
    Dim hNames(1) As String: hNames(0) = "완제품_AJ11": hNames(1) = "완제품_AJ13"
    For hIdx = 0 To 1
        If Trim(CStr(wsDB.Cells(1, hCols(hIdx)).Value)) = "" Then
            wsDB.Cells(1, hCols(hIdx)).Value = hNames(hIdx)
            With wsDB.Cells(1, hCols(hIdx))
                .Font.Bold = True
                .Interior.color = RGB(200, 200, 200)
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next hIdx

    추가수 = 0

    ' ── ID 카운터 초기화 ──
    마지막행 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row
    If 마지막행 < 2 Then
        ID카운터 = 0
    Else
        lastID = wsDB.Cells(마지막행, 2).Value
        If IsNumeric(lastID) Then ID카운터 = CLng(lastID) Else ID카운터 = 0
    End If

    ' ── 시트 목록 수집 (숫자 2자리만) ──
    sheetCount = 0
    For Each ws In wbSrc.Worksheets
        시트명 = ws.Name
        If Len(시트명) = 2 And IsNumeric(시트명) Then
            sheetList(sheetCount) = 시트명
            sheetCount = sheetCount + 1
        End If
    Next ws

    ' 시트명 정렬 (01~31)
    For si = 0 To sheetCount - 2
        For sj = si + 1 To sheetCount - 1
            If CLng(sheetList(si)) > CLng(sheetList(sj)) Then
                temp = sheetList(si): sheetList(si) = sheetList(sj): sheetList(sj) = temp
            End If
        Next sj
    Next si

    ' ── 배치 배열 초기화 ──
    batchSize = 0
    ReDim batchArr(1 To 5000, 1 To 126)
    seedVal = 12345
    ' ──────────────────────────────────────────────────────
    ' 종별_I 배열 초기화 (생산일보I 원지)
    '   KT-1: c3/c9   KT-2: c12/c18  KT-3: c21/c27
    '   KT-4: c30/c36 KT-5: c39/c45  FT_I: c48/c54
    ' ──────────────────────────────────────────────────────
    종별_I(0) = "KT-1":  ROLLID_I(0) = 3:  KG_I(0) = 9
    종별_I(1) = "KT-2":  ROLLID_I(1) = 12: KG_I(1) = 18
    종별_I(2) = "KT-3":  ROLLID_I(2) = 21: KG_I(2) = 27
    종별_I(3) = "KT-4":  ROLLID_I(3) = 30: KG_I(3) = 36
    종별_I(4) = "KT-5":  ROLLID_I(4) = 39: KG_I(4) = 45
    종별_I(5) = "FT":    ROLLID_I(5) = 48: KG_I(5) = 54

    ' ──────────────────────────────────────────────────────
    ' 종별_II 배열 초기화 (생산일보II FT/BT/PRE 행57~63)
    ' ──────────────────────────────────────────────────────
    종별_II(0) = "FT":   ROLLID_II(0) = 3:  KG_II(0) = 9
    종별_II(1) = "BT-1": ROLLID_II(1) = 12: KG_II(1) = 18
    종별_II(2) = "PRE":  ROLLID_II(2) = 21: KG_II(2) = 27

    ' ──────────────────────────────────────────────────────
    ' 투입인원I 구분 배열
    ' ──────────────────────────────────────────────────────
    구분인원(0) = "작업시간":  행인원(0) = 40
    구분인원(1) = "연장근로":  행인원(1) = 41

    ' ??????????????????????????????????????????????????????
    ' 시트별 처리 루프
    ' ??????????????????????????????????????????????????????
    For i = 0 To sheetCount - 1
        시트명 = sheetList(i)
        Set ws = wbSrc.Worksheets(시트명)

        ' 시트별 행 오프셋 동적 감지 (2월 매핑대상: 행 밀림 대응)
        행오프셋 = 시트_행오프셋_감지(ws)
        ' 원지I 끝행 동적 감지 (계/빈행 스캔) → 행확장 = 하단 블록 밀림량
        원지I끝 = 시트_원지I_끝행_찾기(ws, 7 + 행오프셋, 3)
        행확장 = 원지I끝 - 19 - 행오프셋
        If 행확장 < 0 Then 행확장 = 0
        If 행확장 > 30 Then 행확장 = 0   ' 비정상치 방지

        ' 날짜 취득 (생산일보I: 기준행4 c36, 비어있으면 공휴일 → 스킵)
        작성일 = ws.Cells(4 + 행오프셋, 36).Value
        If IsEmpty(작성일) Or IsNull(작성일) Or 작성일 = "" Then GoTo NextSheet
        For eqIdx = 0 To 4
            일원(eqIdx) = 0
            일양(eqIdx) = 0
            일기소(eqIdx) = 0
        Next eqIdx

        ' ────────────────────────────────────────────────
        ' 1. 원지 수집 - 생산일보I (행7~끝 동적, 추가행 포함)
        '    2월: 음영(RGB 255,230,153) 구간 5:5 주/야, 비음영=주간
        ' ────────────────────────────────────────────────
        For t = 0 To 5
            순번 = 1
            For r = 7 + 행오프셋 To 원지I끝
                rollVal = ws.Cells(r, ROLLID_I(t)).Value
                kgVal = ws.Cells(r, KG_I(t)).Value
                rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
                If rollStr = "" Or rollStr = "계" Then GoTo SkipRI
                If IsNumeric(rollStr) Then GoTo SkipRI
                구분I = 구분_음영기반(ws, r, ROLLID_I(t), 순번, t)
                batchSize = batchSize + 1
                ID카운터 = ID카운터 + 1
                batchArr(batchSize, 1) = 작성일
                batchArr(batchSize, 2) = ID카운터
                batchArr(batchSize, 3) = 구분I
                batchArr(batchSize, 4) = "원지"
                batchArr(batchSize, 5) = 순번
                batchArr(batchSize, 6) = 종별_I(t)
                batchArr(batchSize, 7) = rollStr
                batchArr(batchSize, 8) = kgVal
                순번 = 순번 + 1
SkipRI:
            Next r
        Next t

        ' ────────────────────────────────────────────────
        ' 2. 원지 수집 - 생산일보II JRT (행48~54 + 오프셋 + 행확장)
        '    JRT-1: c3/c9   JRT-2: L48:R54 (col12~18)
        '    ★ L48:R54 전체 수집: ROLLID(L=12), KG(R=18) 및 M~Q(13~17) 포함
        '    ★ JRT-2: L 또는 R에 유효값 있으면 원지 기록 (필터 완화)
        ' ────────────────────────────────────────────────
        jrtSeq = 1
        jrt2Seq = 1
        For r = 48 + 행오프셋 + 행확장 To 54 + 행오프셋 + 행확장
            ' JRT-1: c3/c9
            rollVal = ws.Cells(r, 3).Value
            kgVal = ws.Cells(r, 9).Value
            rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
            If rollStr <> "" And Not IsNumeric(rollStr) And rollStr <> "계" _
               And rollStr <> "FT" And rollStr <> "BT" And rollStr <> "지관" Then
                구분JRT = 구분_음영기반(ws, r, 3, jrtSeq, 10)
                batchSize = batchSize + 1
                ID카운터 = ID카운터 + 1
                batchArr(batchSize, 1) = 작성일
                batchArr(batchSize, 2) = ID카운터
                batchArr(batchSize, 3) = 구분JRT
                batchArr(batchSize, 4) = "원지"
                batchArr(batchSize, 5) = jrtSeq
                batchArr(batchSize, 6) = "JRT-1"
                batchArr(batchSize, 7) = rollStr
                batchArr(batchSize, 8) = kgVal
                jrtSeq = jrtSeq + 1
            End If

            ' JRT-2: L48:R54 (col12~18) - L=ROLLID, R=KG, M~Q 중량/보조값
            roll2 = ws.Cells(r, 12).Value     ' L
            kg2 = ws.Cells(r, 18).Value       ' R
            roll2Str = Trim(CStr(IIf(IsEmpty(roll2), "", roll2)))
            ' L(col12) 또는 R(col18)에 값이 있으면 수집 (헤더/계 제외)
            If roll2Str <> "" And roll2Str <> "ROLL ID" And roll2Str <> "KG" _
               And roll2Str <> "계" And UCase(roll2Str) <> "TOTAL" Then
                구분JRT2 = 구분_음영기반(ws, r, 12, jrt2Seq, 11)
                batchSize = batchSize + 1
                ID카운터 = ID카운터 + 1
                batchArr(batchSize, 1) = 작성일
                batchArr(batchSize, 2) = ID카운터
                batchArr(batchSize, 3) = 구분JRT2
                batchArr(batchSize, 4) = "원지"
                batchArr(batchSize, 5) = jrt2Seq
                batchArr(batchSize, 6) = "JRT-2"
                batchArr(batchSize, 7) = roll2
                batchArr(batchSize, 8) = kg2
                ' M~Q(col13~17) 보조값: KG 없을 때 col15(O) 또는 col17(Q) 사용
                If IsEmpty(kg2) Or SafeCDbl(kg2) = 0 Then
                    If Not IsEmpty(ws.Cells(r, 15).Value) Then batchArr(batchSize, 8) = ws.Cells(r, 15).Value
                    If IsEmpty(batchArr(batchSize, 8)) And Not IsEmpty(ws.Cells(r, 17).Value) Then batchArr(batchSize, 8) = ws.Cells(r, 17).Value
                End If
                jrt2Seq = jrt2Seq + 1
            ElseIf (IsEmpty(roll2) Or roll2Str = "") And Not IsEmpty(kg2) _
                   And SafeCDbl(kg2) <> 0 Then
                구분JRT2 = 구분_음영기반(ws, r, 12, jrt2Seq, 11)
                batchSize = batchSize + 1
                ID카운터 = ID카운터 + 1
                batchArr(batchSize, 1) = 작성일
                batchArr(batchSize, 2) = ID카운터
                batchArr(batchSize, 3) = 구분JRT2
                batchArr(batchSize, 4) = "원지"
                batchArr(batchSize, 5) = jrt2Seq
                batchArr(batchSize, 6) = "JRT-2"
                batchArr(batchSize, 7) = ""                  ' ROLLID 없음
                batchArr(batchSize, 8) = kg2
                jrt2Seq = jrt2Seq + 1
            End If
        Next r

        ' ────────────────────────────────────────────────
        ' 3. 원지 수집 - 생산일보II FT/BT/PRE (행57~63 + 오프셋 + 행확장)
        ' ────────────────────────────────────────────────
        For t = 0 To 2
            순번 = 1
            For r = 57 + 행오프셋 + 행확장 To 63 + 행오프셋 + 행확장
                rollVal = ws.Cells(r, ROLLID_II(t)).Value
                kgVal = ws.Cells(r, KG_II(t)).Value
                rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
                kgStr = Trim(CStr(IIf(IsEmpty(kgVal), "", kgVal)))
                If rollStr <> "" And Not IsNumeric(rollStr) And rollStr <> "계" And kgStr <> "" Then
                    구분II = 구분_음영기반(ws, r, ROLLID_II(t), 순번, 12 + t)
                    batchSize = batchSize + 1
                    ID카운터 = ID카운터 + 1
                    batchArr(batchSize, 1) = 작성일
                    batchArr(batchSize, 2) = ID카운터
                    batchArr(batchSize, 3) = 구분II
                    batchArr(batchSize, 4) = "원지"
                    batchArr(batchSize, 5) = 순번
                    batchArr(batchSize, 6) = 종별_II(t)
                    batchArr(batchSize, 7) = rollStr
                    batchArr(batchSize, 8) = kgVal
                    순번 = 순번 + 1
                End If
            Next r
        Next t

        ' ────────────────────────────────────────────────
        ' 4. 완제품 수집 - 생산일보I (행25~34 + 오프셋 + 행확장)
        '    c3=품명, c13=가공라인(병합→None 시 prev 승계),
        '    c17=생산량BOX, c21=기소, c25=PLT1(Y), c29=PLT2(AC), c33=인원
        '    ※ KT PLT(11)=Y25:Y34, PLT(13)=AC25:AC34
        ' ────────────────────────────────────────────────
        순번 = 1
        prev_가공라인 = ""
        For r = 25 + 행오프셋 + 행확장 To 34 + 행오프셋 + 행확장
            품명 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 3).Value), "", ws.Cells(r, 3).Value)))
            If 품명 = "" Or 품명 = "계" Or 품명 = "TOTAL" Or IsNumeric(품명) Then GoTo SkipFin1
            가공라인 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 13).Value), "", ws.Cells(r, 13).Value)))
            If 가공라인 = "" Then
                가공라인 = prev_가공라인
            Else
                가공라인 = 종별변환(가공라인)
                prev_가공라인 = 가공라인
            End If
            If 가공라인 = "" Then 가공라인 = "KT-1"
            ' 완제품I(행25~34): 02시트 좌측/첫블록 = 주간 (랜덤구분 제거, 주야 혼용 방지)
            구분완I = "생산일보_주간"
            batchSize = batchSize + 1
            ID카운터 = ID카운터 + 1
            batchArr(batchSize, 1) = 작성일
            batchArr(batchSize, 2) = ID카운터
            batchArr(batchSize, 3) = 구분완I
            batchArr(batchSize, 4) = "완제품"
            batchArr(batchSize, 5) = 순번
            batchArr(batchSize, 9) = "KT"                              ' 완제품_구분
            batchArr(batchSize, 10) = 품명                              ' 완제품_품명
            batchArr(batchSize, 11) = 가공라인                          ' 가공라인호기
            batchArr(batchSize, 12) = ws.Cells(r, 17).Value            ' 생산량_BOX
            batchArr(batchSize, 13) = ws.Cells(r, 21).Value            ' 원지기준소요량
            batchArr(batchSize, 14) = ws.Cells(r, 25).Value            ' PLT1
            batchArr(batchSize, 15) = ws.Cells(r, 29).Value            ' PLT2
            batchArr(batchSize, 18) = ws.Cells(r, 33).Value            ' 인원
            eqIdx = PA_설비인덱스(가공라인)
            If eqIdx >= 0 Then 일기소(eqIdx) = 일기소(eqIdx) + SafeCDbl(ws.Cells(r, 21).Value)
            순번 = 순번 + 1
SkipFin1:
        Next r

        ' ────────────────────────────────────────────────
        ' 5. 완제품 수집 - 생산일보II (행48~63 + 오프셋 + 행확장)
        '    c31=구분(JRT/WT/FT/BT), c33=품명, c41=생산량
        '    c45=기소, c49=PLT1(AW), c52=PLT2(AZ), c55=인원
        '    ※ JRT PLT(11)=AW48:AW51, PLT(13)=AZ48:AZ51
        '    TOTAL행 스킵, 가공라인호기 = 구분에서 변환
        ' ────────────────────────────────────────────────
        순번 = 1
        현재구분 = ""
        For r = 48 + 행오프셋 + 행확장 To 63 + 행오프셋 + 행확장
            구분셀 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 31).Value), "", ws.Cells(r, 31).Value)))
            If 구분셀 <> "" And UCase(구분셀) <> "TOTAL" And 구분셀 <> "구분" Then
                현재구분 = 구분셀
            End If
            If UCase(구분셀) = "TOTAL" Then GoTo SkipFin2
            품명 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 33).Value), "", ws.Cells(r, 33).Value)))
            If 품명 = "" Or 품명 = "구분" Or 품명 = "품명" Or IsNumeric(품명) Then GoTo SkipFin2
            ' 가공라인호기: 현재구분 → JRT→JRT-1, BT→BT-1, 그외 그대로
            호기II = 현재구분
            If 호기II = "JRT" Then 호기II = "JRT-1"
            If 호기II = "BT" Then 호기II = "BT-1"
            생산량II = ws.Cells(r, 41).Value
            ' 완제품II(행48~63): 02시트 우측/둘째블록 = 야간 (랜덤구분 제거, 주야 혼용 방지)
            구분완II = "생산일보_야간"
            batchSize = batchSize + 1
            ID카운터 = ID카운터 + 1
            batchArr(batchSize, 1) = 작성일
            batchArr(batchSize, 2) = ID카운터
            batchArr(batchSize, 3) = 구분완II
            batchArr(batchSize, 4) = "완제품"
            batchArr(batchSize, 5) = 순번
            batchArr(batchSize, 9) = 호기II                            ' 완제품_구분(가공라인호기)
            batchArr(batchSize, 10) = 품명                              ' 완제품_품명
            batchArr(batchSize, 11) = 호기II                           ' 가공라인호기
            batchArr(batchSize, 12) = 생산량II                         ' 생산량_BOX
            batchArr(batchSize, 13) = ws.Cells(r, 45).Value            ' 원지기준소요량
            batchArr(batchSize, 14) = ws.Cells(r, 49).Value            ' PLT1
            batchArr(batchSize, 15) = ws.Cells(r, 52).Value            ' PLT2
            batchArr(batchSize, 18) = ws.Cells(r, 55).Value            ' 인원
            eqIdx = PA_설비인덱스(호기II)
            If eqIdx >= 0 Then 일기소(eqIdx) = 일기소(eqIdx) + SafeCDbl(ws.Cells(r, 45).Value)
            순번 = 순번 + 1
SkipFin2:
        Next r

        ' ────────────────────────────────────────────────
        ' 6. 투입인원 - 생산일보I (행40~41 + 오프셋 + 행확장) KT
        '    DB 컬럼: F19=작업장, F20=구분, F21=인원수, F22=작업시간, F23=총작업시간, F24=실투입원, F25=환산
        '    행40=정규 → "작업시간" 레코드, 행41=연장 → "연장근로" 레코드
        ' ────────────────────────────────────────────────
        rBase = 40 + 행오프셋 + 행확장
        v정규 = ws.Cells(rBase, 10).Value
        인원유효 = Not IsEmpty(v정규) And SafeCDbl(v정규) <> 0
        v연장 = ws.Cells(rBase, 6).Value
        시간유효 = Not IsEmpty(v연장) And SafeCDbl(v연장) <> 0
        If 인원유효 Or 시간유효 Then
            batchSize = batchSize + 1
            ID카운터 = ID카운터 + 1
            batchArr(batchSize, 1) = 작성일
            batchArr(batchSize, 2) = ID카운터
            batchArr(batchSize, 3) = "생산일보"
            batchArr(batchSize, 4) = "투입인원"
            batchArr(batchSize, 5) = 1
            batchArr(batchSize, 19) = "KT"
            batchArr(batchSize, 20) = "작업시간"
            batchArr(batchSize, 21) = SafeCDbl(ws.Cells(rBase, 6).Value)
            batchArr(batchSize, 22) = SafeCDbl(ws.Cells(rBase, 10).Value)
            batchArr(batchSize, 23) = SafeCDbl(ws.Cells(rBase, 13).Value)
            batchArr(batchSize, 24) = SafeCDbl(ws.Cells(rBase, 17).Value)
            batchArr(batchSize, 25) = SafeCDbl(ws.Cells(rBase, 21).Value)
        End If

        rBase = 41 + 행오프셋 + 행확장
        v정규 = ws.Cells(rBase, 10).Value
        인원유효 = Not IsEmpty(v정규) And SafeCDbl(v정규) <> 0
        v연장 = ws.Cells(rBase, 6).Value
        시간유효 = Not IsEmpty(v연장) And SafeCDbl(v연장) <> 0
        If 인원유효 Or 시간유효 Then
            batchSize = batchSize + 1
            ID카운터 = ID카운터 + 1
            batchArr(batchSize, 1) = 작성일
            batchArr(batchSize, 2) = ID카운터
            batchArr(batchSize, 3) = "생산일보"
            batchArr(batchSize, 4) = "투입인원"
            batchArr(batchSize, 5) = 1
            batchArr(batchSize, 19) = "KT"
            batchArr(batchSize, 20) = "연장근로"
            batchArr(batchSize, 21) = SafeCDbl(ws.Cells(rBase, 6).Value)
            batchArr(batchSize, 22) = SafeCDbl(ws.Cells(rBase, 10).Value)
            batchArr(batchSize, 23) = SafeCDbl(ws.Cells(rBase, 13).Value)
            batchArr(batchSize, 24) = SafeCDbl(ws.Cells(rBase, 17).Value)
            batchArr(batchSize, 25) = SafeCDbl(ws.Cells(rBase, 21).Value)
        End If

        ' ────────────────────────────────────────────────
        ' 7. 투입인원 - 생산일보II (행68~77 + 오프셋 + 행확장)
        '    DB 컬럼: F19=작업장, F20=구분, F21=인원수, F22=작업시간, F23=총작업시간, F24=실투입원, F25=환산
        '    정규행 → "작업시간", 연장행 → "연장근로" (별도 레코드)
        ' ────────────────────────────────────────────────
        For idx = 1 To 10
            인원II인원(idx) = Empty: 인원II정규(idx) = Empty: 인원II연장(idx) = Empty: 인원II작업장(idx) = ""
            인원II총MH(idx) = Empty: 인원II실투입(idx) = Empty: 인원II환산(idx) = Empty
            인원II인당실(idx) = Empty: 인원II인당환(idx) = Empty
        Next idx
        인원II갯수 = 0
        현재작업장 = ""
        For r = 68 + 행오프셋 + 행확장 + 행확장 To 77 + 행오프셋 + 행확장 + 행확장
            작업장셀 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 2).Value), "", ws.Cells(r, 2).Value)))
            If 작업장셀 <> "" And 작업장셀 <> "구분" And 작업장셀 <> "작업장" Then
                If 작업장셀 = "지관" Then
                    현재작업장 = "PRE"
                Else
                    현재작업장 = 작업장셀
                End If
            End If
            구분4 = Trim(CStr(IIf(IsEmpty(ws.Cells(r, 4).Value), "", ws.Cells(r, 4).Value)))
            hasData2 = False
            For c2 = 7 To 11
                If Not IsEmpty(ws.Cells(r, c2).Value) And IsNumeric(ws.Cells(r, c2).Value) Then
                    If SafeCDbl(ws.Cells(r, c2).Value) <> 0 Then hasData2 = True: Exit For
                End If
            Next c2
            If Not hasData2 Or 현재작업장 = "" Then GoTo SkipInw2
            For idx = 1 To 인원II갯수
                If 인원II작업장(idx) = 현재작업장 Then Exit For
            Next idx
            If idx > 인원II갯수 Then
                인원II갯수 = 인원II갯수 + 1
                인원II작업장(인원II갯수) = 현재작업장
                인원II인원(인원II갯수) = Empty
                인원II정규(인원II갯수) = Empty
                인원II연장(인원II갯수) = Empty
                인원II총MH(인원II갯수) = Empty
                인원II실투입(인원II갯수) = Empty
                인원II환산(인원II갯수) = Empty
                인원II인당실(인원II갯수) = Empty
                인원II인당환(인원II갯수) = Empty
                idx = 인원II갯수
            End If
            If 구분4 = "정규" Then
                인원II인원(idx) = ws.Cells(r, 7).Value
                인원II정규(idx) = ws.Cells(r, 9).Value
                인원II총MH(idx) = ws.Cells(r, 11).Value
                인원II실투입(idx) = ws.Cells(r, 14).Value
                인원II환산(idx) = ws.Cells(r, 18).Value
                인원II인당실(idx) = ws.Cells(r, 22).Value
                인원II인당환(idx) = ws.Cells(r, 26).Value
            ElseIf 구분4 = "연장" Then
                인원II연장(idx) = ws.Cells(r, 9).Value
            End If
SkipInw2:
        Next r
        For idx = 1 To 인원II갯수
            v정규 = 인원II인원(idx)
            인원유효 = Not IsEmpty(v정규) And SafeCDbl(v정규) <> 0
            v연장 = 인원II정규(idx)
            시간유효 = Not IsEmpty(v연장) And SafeCDbl(v연장) <> 0
            If 인원유효 Or 시간유효 Then
                batchSize = batchSize + 1
                ID카운터 = ID카운터 + 1
                batchArr(batchSize, 1) = 작성일
                batchArr(batchSize, 2) = ID카운터
                batchArr(batchSize, 3) = "생산일보"
                batchArr(batchSize, 4) = "투입인원"
                batchArr(batchSize, 5) = idx
                batchArr(batchSize, 19) = 인원II작업장(idx)
                batchArr(batchSize, 20) = "작업시간"
                batchArr(batchSize, 21) = SafeCDbl(인원II인원(idx))
                batchArr(batchSize, 22) = SafeCDbl(인원II정규(idx))
                batchArr(batchSize, 23) = SafeCDbl(인원II총MH(idx))
                batchArr(batchSize, 24) = SafeCDbl(인원II실투입(idx))
                batchArr(batchSize, 25) = SafeCDbl(인원II환산(idx))
            End If
        Next idx

        ' ────────────────────────────────────────────────
        ' 8. PA 수집 - 생산일보I KT 수율 지표 (1레코드)
        '    행25=원투총량, 행27=양품총량, 행29=수율
        '    행30=LOSS율, 행31=목표대비양품율
        '    일계:c45, 주누계:c49, 월누계:c53
        '    ※ 02시트: 행24 열49/53은 구분("주"/"월")헤더일 수 있음 → 행26 기준소요량 시도, IsNumeric 체크
        ' ────────────────────────────────────────────────
        batchSize = batchSize + 1
        ID카운터 = ID카운터 + 1
        batchArr(batchSize, 1) = 작성일
        batchArr(batchSize, 2) = ID카운터
        batchArr(batchSize, 3) = "생산일보"
        batchArr(batchSize, 4) = "PA"
        batchArr(batchSize, 5) = 1
        ' KT 일계 - 다중 셀 fallback (2/3월 시트 구조 변형 대응)
        Dim v원투 As Variant, v양품 As Variant, v수율 As Variant, vLOSS As Variant, v양품율 As Variant
        v원투 = PA_KT셀읽기(ws, 25, 행오프셋, 행확장)
        v양품 = PA_KT셀읽기(ws, 27, 행오프셋, 행확장)
        v수율 = PA_KT셀읽기(ws, 29, 행오프셋, 행확장)
        vLOSS = PA_KT셀읽기(ws, 30, 행오프셋, 행확장)
        v양품율 = PA_KT셀읽기(ws, 31, 행오프셋, 행확장)
        ' 원투 fallback: 시트 셀 빈 경우 완제품 KT 원지기준소요량 합계 사용
        If (IsEmpty(v원투) Or Not IsNumeric(v원투) Or SafeCDbl(v원투) = 0) And 일기소(0) > 0 Then
            v원투 = 일기소(0)
        End If
        batchArr(batchSize, 28) = v원투
        batchArr(batchSize, 29) = v양품
        batchArr(batchSize, 30) = PA_수율보정(v수율, v원투, v양품)
        batchArr(batchSize, 32) = vLOSS
        batchArr(batchSize, 33) = v양품율
        일원(0) = SafeCDbl(v원투)
        일양(0) = SafeCDbl(v양품)
        batchArr(batchSize, 31) = 일기소(0)

        ' ────────────────────────────────────────────────
        ' 8-2. PA 수집 - 생산일보II (AE66:BA79) 일계만 수집. 누계/평균은 DB 기반 계산(일일생산DB_PA누계평균_행반영)
        '    67행 빈행, 시트별 구조 변형 → 동적 스캔. 열: AI(35)=원지투입, AM(39)=양품환산, AQ(43)=원지수율
        '    설비순: JRT(0), WT/PRE(1), FT(2), BT(3) × 일계(rowIdx=0)만
        ' ────────────────────────────────────────────────
        Dim blockRows As Variant
        blockRows = PA_II_블록행_동적찾기(ws, 행오프셋, 행확장)
        ' JRT/FT/BT/PRE 일계만 (43~45, 40~42, 46~48, 49~51)
        vII = PA_II_블록값읽기(ws, blockRows, 0, 0, 35): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 43) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 0, 0, 39): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 44) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 0, 0, 43): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 45) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 2, 0, 35): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 40) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 2, 0, 39): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 41) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 2, 0, 43): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 42) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 3, 0, 35): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 46) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 3, 0, 39): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 47) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 3, 0, 43): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 48) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 1, 0, 35): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 49) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 1, 0, 39): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 50) = SafeCDbl(vII)
        vII = PA_II_블록값읽기(ws, blockRows, 1, 0, 43): If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 51) = SafeCDbl(vII)
        일원(2) = SafeCDbl(batchArr(batchSize, 43))
        일양(2) = SafeCDbl(batchArr(batchSize, 44))
        일원(1) = SafeCDbl(batchArr(batchSize, 40))
        일양(1) = SafeCDbl(batchArr(batchSize, 41))
        일원(3) = SafeCDbl(batchArr(batchSize, 46))
        일양(3) = SafeCDbl(batchArr(batchSize, 47))
        일원(4) = SafeCDbl(batchArr(batchSize, 49))
        일양(4) = SafeCDbl(batchArr(batchSize, 50))
        batchArr(batchSize, 123) = 일기소(1)
        batchArr(batchSize, 124) = 일기소(2)
        batchArr(batchSize, 125) = 일기소(3)
        batchArr(batchSize, 126) = 일기소(4)
        ' 누계/평균(34~122)은 DB 쓰기 후 Module4.일일생산DB_PA누계평균_행반영에서 DB 적재값 기반 계산

        추가수 = 추가수 + 1

NextSheet:
    Next i

    ' ──────────────────────────────────────────────────────
    ' DB에 일괄 쓰기
    ' ──────────────────────────────────────────────────────
    If batchSize > 0 Then
        새행시작 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
        Dim writeArr As Variant, rw As Long, cl As Long
        ReDim writeArr(1 To batchSize, 1 To 126)
        For rw = 1 To batchSize
            For cl = 1 To 126
                writeArr(rw, cl) = batchArr(rw, cl)
            Next cl
        Next rw
        wsDB.Range(wsDB.Cells(새행시작, 1), wsDB.Cells(새행시작 + batchSize - 1, 126)).Value2 = writeArr
        On Error Resume Next
        Call Module4.일일생산DB_PA누계평균_행반영(wsDB, 새행시작, 새행시작 + batchSize - 1)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo MigErr
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    wbSrc.Close False

    MsgBox "이관 완료!" & vbCrLf & _
           "총 " & batchSize & "건의 레코드가 일일생산DB에 추가되었습니다.", vbInformation, "이관 결과"
    Exit Sub

MigErr:
    errNum = Err.Number: errDesc = Err.Description   ' 즉시 보존 (다른 코드가 Err 덮어쓰기 전)
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next: wbSrc.Close False: On Error GoTo 0
    MsgBox "오류 발생!" & vbCrLf & _
           "오류 번호: " & errNum & vbCrLf & _
           "오류 내용: " & errDesc & vbCrLf & vbCrLf & _
           "처리 중 시트: " & 시트명 & vbCrLf & _
           "현재 batchSize: " & batchSize, vbCritical, "마이그레이션 오류"
End Sub


' ──────────────────────────────────────────────────────────
' 헬퍼: 안전 Double 변환 (형식불일치 오류 13 방지)
' ──────────────────────────────────────────────────────────
Private Function SafeCDbl(v As Variant) As Double
    SafeCDbl = 0
    On Error Resume Next
    If Not IsEmpty(v) And Not IsNull(v) Then SafeCDbl = CDbl(v)
    If Err.Number <> 0 Then SafeCDbl = 0
    Err.Clear
End Function

' ──────────────────────────────────────────────────────────
' 헬퍼: 종별 변환 (KT-① → KT-1 등)
' ──────────────────────────────────────────────────────────
Private Function 종별변환(s As String) As String
    Dim result As String: result = s
    result = Replace(result, "KT-①", "KT-1")
    result = Replace(result, "KT-②", "KT-2")
    result = Replace(result, "KT-③", "KT-3")
    result = Replace(result, "KT-④", "KT-4")
    result = Replace(result, "BT-⑤", "KT-5")
    result = Replace(result, "KT-⑤", "KT-5")
    result = Replace(result, "KT-⑥", "FT")
    result = Replace(result, "지관", "PRE")
    종별변환 = result
End Function


' ──────────────────────────────────────────────────────────
' 2월 전용: 원지 생산일보I 끝행 동적 감지 (추가행 포함)
'   col열(startRow~) 스캔, "계" 또는 연속 2행 빈값 시 직전행 반환
' ──────────────────────────────────────────────────────────
Private Function 시트_원지I_끝행_찾기(ws As Worksheet, startRow As Long, col As Long) As Long
    Dim rr As Long
    Dim v As Variant
    Dim s As String
    Dim emptyCnt As Long
    시트_원지I_끝행_찾기 = startRow
    emptyCnt = 0
    On Error Resume Next
    For rr = startRow To startRow + 40
        If rr > ws.usedRange.Rows.Count + 5 Then Exit For
        v = ws.Cells(rr, col).Value
        s = Trim(CStr(IIf(IsEmpty(v), "", v)))
        If s = "계" Or UCase(s) = "TOTAL" Then
            시트_원지I_끝행_찾기 = rr - 1
            Exit Function
        End If
        If s = "" Or IsNumeric(s) Then
            emptyCnt = emptyCnt + 1
            If emptyCnt >= 2 Then
                시트_원지I_끝행_찾기 = rr - 2
                If 시트_원지I_끝행_찾기 < startRow Then 시트_원지I_끝행_찾기 = startRow
                Exit Function
            End If
        Else
            emptyCnt = 0
            시트_원지I_끝행_찾기 = rr
        End If
    Next rr
    Err.Clear
End Function

' ──────────────────────────────────────────────────────────
' 2월 전용: 시트별 행 오프셋 동적 감지 (날짜 셀 위치 기반)
'   열36(AJ) 1~20행 스캔, IsDate인 첫 행 - 4 = 오프셋
' ──────────────────────────────────────────────────────────
Private Function 시트_행오프셋_감지(ws As Worksheet) As Long
    Dim rr As Long, v As Variant
    시트_행오프셋_감지 = 0
    On Error Resume Next
    For rr = 1 To 20
        v = ws.Cells(rr, 36).Value
        If Not IsEmpty(v) And IsDate(v) Then
            시트_행오프셋_감지 = rr - 4
            If 시트_행오프셋_감지 < -3 Then 시트_행오프셋_감지 = 0
            If 시트_행오프셋_감지 > 10 Then 시트_행오프셋_감지 = 0
            Exit Function
        End If
    Next rr
    Err.Clear
End Function

' 생산일보II PA 셀 - 행확장 2배/1배/0배 다중 시도 (레거시)
Private Function PA_II셀읽기(ws As Worksheet, baseRow As Long, col As Long, 행오프셋 As Long, 행확장 As Long) As Variant
    Dim r As Long, v As Variant, mult As Long
    For mult = 2 To 0 Step -1
        r = baseRow + 행오프셋 + (행확장 * mult)
        If r >= 1 And r <= 200 Then
            v = ws.Cells(r, col).Value
            If Not IsEmpty(v) And IsNumeric(v) Then PA_II셀읽기 = v: Exit Function
        End If
    Next mult
    PA_II셀읽기 = Empty
End Function

' AE66:BA79 영역 동적 스캔 - 67행 빈행 등 시트별 구조 변형 대응
' 4설비(JRT, WT/PRE, FT, BT) × 3행(일계, 주평균, 월평균) = 12개 데이터행, 빈행 스킵
' col35(AI)=원지투입에 숫자 있는 행만 수집 → 순서대로 블록 할당
' 반환: blockRows(0 To 3, 0 To 2) - blockRows(설비, 0)=일계행, 1=주행, 2=월행 (0=미발견)
Private Function PA_II_블록행_동적찾기(ws As Worksheet, 행오프셋 As Long, 행확장 As Long) As Variant
    Dim out(0 To 3, 0 To 2) As Long
    Dim rr As Long, cnt As Long, v As Variant
    Dim startScan As Long, mult As Long, bestCnt As Long, bestOut(0 To 3, 0 To 2) As Long
    Dim i As Long, j As Long
    For i = 0 To 3: For j = 0 To 2: out(i, j) = 0: bestOut(i, j) = 0: Next j: Next i
    bestCnt = 0
    For mult = 2 To 0 Step -1
        startScan = 64 + 행오프셋 + (행확장 * mult)
        cnt = 0
        For i = 0 To 3: For j = 0 To 2: out(i, j) = 0: Next j: Next i
        For rr = startScan To startScan + 25
            If rr < 1 Or rr > 200 Then GoTo Nx
            v = ws.Cells(rr, 35).Value
            If Not IsEmpty(v) And IsNumeric(v) Then
                If cnt < 12 Then
                    out(cnt \ 3, cnt Mod 3) = rr
                    cnt = cnt + 1
                End If
            End If
Nx:
        Next rr
        If cnt > bestCnt Then
            bestCnt = cnt
            For i = 0 To 3: For j = 0 To 2: bestOut(i, j) = out(i, j): Next j: Next i
        End If
        If bestCnt >= 12 Then Exit For
    Next mult
    PA_II_블록행_동적찾기 = bestOut
End Function

' blockRows에서 특정 설비/구분의 셀값 읽기 (col: 35=원투, 39=양품, 43=수율)
Private Function PA_II_블록값읽기(ws As Worksheet, blockRows As Variant, blockIdx As Long, rowIdx As Long, col As Long) As Variant
    Dim r As Long
    If blockIdx < 0 Or blockIdx > 3 Or rowIdx < 0 Or rowIdx > 2 Then PA_II_블록값읽기 = Empty: Exit Function
    r = blockRows(blockIdx, rowIdx)
    If r <= 0 Then PA_II_블록값읽기 = Empty: Exit Function
    Dim v As Variant: v = ws.Cells(r, col).Value
    If Not IsEmpty(v) And IsNumeric(v) Then PA_II_블록값읽기 = v Else PA_II_블록값읽기 = Empty
End Function

' KT PA 셀 다중 시도 (행/열 변형 시트 대응)
Private Function PA_KT셀읽기(ws As Worksheet, baseRow As Long, 행오프셋 As Long, 행확장 As Long) As Variant
    Dim r As Long, c As Long, v As Variant
    Dim rowOffs() As Variant, colOffs() As Variant
    rowOffs = Array(0, -1, 1, -2, 2)
    colOffs = Array(45, 43, 47, 41, 49)
    Dim i As Long, j As Long
    For i = 0 To UBound(rowOffs)
        For j = 0 To UBound(colOffs)
            r = baseRow + 행오프셋 + 행확장 + CLng(rowOffs(i))
            c = CLng(colOffs(j))
            If r >= 1 And r <= 200 And c >= 1 And c <= 100 Then
                v = ws.Cells(r, c).Value
                If Not IsEmpty(v) And IsNumeric(v) Then
                    PA_KT셀읽기 = v
                    Exit Function
                End If
            End If
        Next j
    Next i
    PA_KT셀읽기 = Empty
End Function

' 수율 0~150% 범위 검증, 초과 시 원투/양품으로 재계산 (6583 등 오기록 방지)
Private Function PA_수율보정(v수율 As Variant, v원투 As Variant, v양품 As Variant) As Variant
    Dim d As Double, 원 As Double, 양 As Double
    If IsEmpty(v수율) Or Not IsNumeric(v수율) Then
        If Not IsEmpty(v원투) And IsNumeric(v원투) And Not IsEmpty(v양품) And IsNumeric(v양품) Then
            원 = SafeCDbl(v원투): 양 = SafeCDbl(v양품)
            If 원 > 0.0000000001 Then PA_수율보정 = Round(양 / 원 * 100, 1): Exit Function
        End If
        PA_수율보정 = Empty
        Exit Function
    End If
    d = SafeCDbl(v수율)
    If d >= 0 And d <= 150 Then PA_수율보정 = v수율: Exit Function
    If Not IsEmpty(v원투) And IsNumeric(v원투) And Not IsEmpty(v양품) And IsNumeric(v양품) Then
        원 = SafeCDbl(v원투): 양 = SafeCDbl(v양품)
        If 원 > 0.0000000001 Then PA_수율보정 = Round(양 / 원 * 100, 1): Exit Function
    End If
    PA_수율보정 = Empty
End Function

Private Function PA_설비인덱스(ByVal 설비명 As String) As Long
    Dim s As String
    s = UCase(Trim(설비명))
    PA_설비인덱스 = -1
    If s = "" Then Exit Function
    If left$(s, 2) = "KT" Then PA_설비인덱스 = 0: Exit Function
    If left$(s, 2) = "FT" Then PA_설비인덱스 = 1: Exit Function
    If left$(s, 3) = "JRT" Then PA_설비인덱스 = 2: Exit Function
    If left$(s, 2) = "BT" Then PA_설비인덱스 = 3: Exit Function
    If left$(s, 3) = "PRE" Or left$(s, 3) = "PNL" Or left$(s, 2) = "WT" Then PA_설비인덱스 = 4
End Function

Private Function PA_주차번호(ByVal d As Date) As Long
    Dim 월1일 As Date, 요일 As Long, 첫일요일 As Date
    월1일 = DateSerial(Year(d), Month(d), 1)
    요일 = Weekday(월1일, vbMonday)
    첫일요일 = 월1일 + (7 - 요일)
    If d < 첫일요일 Then
        PA_주차번호 = 1
    Else
        PA_주차번호 = 2 + CLng((d - 첫일요일) \ 7)
    End If
End Function

Private Function PA_주키(ByVal d As Date) As String
    PA_주키 = Format$(d, "yyyymm") & "_" & CStr(PA_주차번호(d))
End Function

Private Function PA_월키(ByVal d As Date) As String
    PA_월키 = Format$(d, "yyyymm")
End Function

Private Function PA_누적배열_가져오기(dict As Object, ByVal key As String) As Variant
    If dict.exists(key) Then
        PA_누적배열_가져오기 = dict(key)
    Else
        PA_누적배열_가져오기 = Array(0#, 0#, 0#, 0#, 0#)
    End If
End Function

Private Sub PA_누적배열_더하기(dict As Object, ByVal key As String, ByRef values() As Double)
    Dim arr As Variant, i As Long
    arr = PA_누적배열_가져오기(dict, key)
    For i = 0 To 4
        arr(i) = CDbl(arr(i)) + values(i)
    Next i
    If dict.exists(key) Then
        dict(key) = arr
    Else
        dict.Add key, arr
    End If
End Sub

Private Function PA_평균(ByVal 합계 As Double, ByVal 일수 As Long) As Double
    If 일수 <= 0 Then Exit Function
    PA_평균 = Round(합계 / 일수, 1)
End Function

Private Function PA_비율(ByVal 분자 As Double, ByVal 분모 As Double) As Double
    If Abs(분모) <= 0.0000000001 Then Exit Function
    PA_비율 = Round(분자 / 분모 * 100, 1)
End Function

' ──────────────────────────────────────────────────────────
' 2월 전용: 음영(RGB 255,230,153) 기반 주야간 구분
'   음영 구간: 5:5 비율 주간/야간 (순번+groupId 기반)
'   비음영: 전부 주간
' ──────────────────────────────────────────────────────────
Private Function 구분_음영기반(ws As Worksheet, r As Long, col As Long, 순번 As Long, ByVal groupId As Long) As String
    구분_음영기반 = "생산일보_주간"
    On Error Resume Next
    Dim c As Long: c = ws.Cells(r, col).Interior.color
    If Err.Number <> 0 Then Err.Clear: Exit Function
    ' RGB(255,230,153) = 10086443 (VBA Long)
    If c = 10086443 Or (c = 153 + 230 * 256 + 255 * 65536) Then
        ' 음영 구간: 5:5 주/야 (순번+groupId 짝수=주간, 홀수=야간)
        If ((순번 + groupId) Mod 2) = 1 Then 구분_음영기반 = "생산일보_야간"
    End If
    Err.Clear
End Function





