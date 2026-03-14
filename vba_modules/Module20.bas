Attribute VB_Name = "Module20"
' ============================================================
' 이전 생산관리 파일 → 일일생산DB 이관 VBA v6
' - v5: Dim 최상단 이동, JRT-2 원지 수집 정상화
' - v6: 4.작업시간/인력투입 매핑 수정 (현재 형식에 맞춤)
'   설비별 정규+연장 1레코드 통합. F19=주간정규, F20=주간연장
' - v7: 정규=주간(F19), 연장=야간(F21). 휴일 없음. 첨부 F22~25
' - v8: AJ11→F16, AJ13→F17, 16번 이후 컬럼 +2 시프트. 인원F19~25, PA F28~
' ============================================================
Option Explicit

Public Sub 이전데이터_이관_실행()

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
    Dim sheetCount  As Integer
    Dim si          As Integer, sj As Integer, temp As String

    Dim batchArr()  As Variant
    Dim batchSize   As Long
    Dim seedVal     As Long
    Dim ID카운터    As Long
    Dim 마지막행    As Long
    Dim lastID      As Variant
    Dim 새행시작    As Long
    Dim 추가수      As Long
    Dim i           As Integer

    ' 루프/수집용 공통
    Dim r           As Long
    Dim t           As Integer
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
    Dim 인원II갯수   As Integer
    Dim idx          As Integer

    ' PA (KT 수율 지표)
    Dim pa행(6)     As Long
    Dim vII         As Variant

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
        Dim ans As Integer
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

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo MigErr

    ' DB 헤더 보완 (col 125=완제품_AJ11, col 126=완제품_AJ13 없으면 추가) - 인라인
    Dim hIdx As Integer
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

        ' 날짜 취득 (생산일보I: 행4 c36, 비어있으면 공휴일 → 스킵)
        작성일 = ws.Cells(4, 36).Value
        If IsEmpty(작성일) Or IsNull(작성일) Or 작성일 = "" Then GoTo NextSheet

        ' ────────────────────────────────────────────────
        ' 1. 원지 수집 - 생산일보I (행7~19)
        '    종별 6가지, ROLLID열/KG열 쌍으로 처리
        ' ────────────────────────────────────────────────
        For t = 0 To 5
            순번 = 1
            For r = 7 To 19
                rollVal = ws.Cells(r, ROLLID_I(t)).Value
                kgVal = ws.Cells(r, KG_I(t)).Value
                rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
                If rollStr = "" Or rollStr = "계" Then GoTo SkipRI
                If IsNumeric(rollStr) Then GoTo SkipRI
                구분I = 랜덤구분(seedVal, 순번 * (t + 1) + 10)
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
        ' 2. 원지 수집 - 생산일보II JRT (행48~54)
        '    JRT-1: c3/c9   JRT-2: L48:R54 (col12~18)
        '    ★ L48:R54 전체 수집: ROLLID(L=12), KG(R=18) 및 M~Q(13~17) 포함
        '    ★ JRT-2: L 또는 R에 유효값 있으면 원지 기록 (필터 완화)
        ' ────────────────────────────────────────────────
        jrtSeq = 1
        jrt2Seq = 1
        For r = 48 To 54
            ' JRT-1: c3/c9
            rollVal = ws.Cells(r, 3).Value
            kgVal = ws.Cells(r, 9).Value
            rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
            If rollStr <> "" And Not IsNumeric(rollStr) And rollStr <> "계" _
               And rollStr <> "FT" And rollStr <> "BT" And rollStr <> "지관" Then
                구분JRT = 랜덤구분(seedVal, jrtSeq + 100)
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
                ' 숫자 ROLLID도 허용 (이전 필터 완화)
                구분JRT2 = 랜덤구분(seedVal, jrt2Seq + 150)
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
                If IsEmpty(kg2) Or (IsNumeric(kg2) And CDbl(kg2) = 0) Then
                    If Not IsEmpty(ws.Cells(r, 15).Value) Then batchArr(batchSize, 8) = ws.Cells(r, 15).Value
                    If IsEmpty(batchArr(batchSize, 8)) And Not IsEmpty(ws.Cells(r, 17).Value) Then batchArr(batchSize, 8) = ws.Cells(r, 17).Value
                End If
                jrt2Seq = jrt2Seq + 1
            ElseIf (IsEmpty(roll2) Or roll2Str = "") And Not IsEmpty(kg2) _
                   And IsNumeric(kg2) And CDbl(kg2) <> 0 Then
                ' R에만 KG 있고 L 비어있는 경우 (R이 주값)
                구분JRT2 = 랜덤구분(seedVal, jrt2Seq + 150)
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
        ' 3. 원지 수집 - 생산일보II FT/BT/PRE (행57~63)
        ' ────────────────────────────────────────────────
        For t = 0 To 2
            순번 = 1
            For r = 57 To 63
                rollVal = ws.Cells(r, ROLLID_II(t)).Value
                kgVal = ws.Cells(r, KG_II(t)).Value
                rollStr = Trim(CStr(IIf(IsEmpty(rollVal), "", rollVal)))
                kgStr = Trim(CStr(IIf(IsEmpty(kgVal), "", kgVal)))
                If rollStr <> "" And Not IsNumeric(rollStr) And rollStr <> "계" And kgStr <> "" Then
                    구분II = 랜덤구분(seedVal, 순번 * (t + 10) + 200)
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
        ' 4. 완제품 수집 - 생산일보I (행25~34)
        '    c3=품명, c13=가공라인(병합→None 시 prev 승계),
        '    c17=생산량BOX, c21=기소, c25=PLT1(Y), c29=PLT2(AC), c33=인원
        '    ※ KT PLT(11)=Y25:Y34, PLT(13)=AC25:AC34
        ' ────────────────────────────────────────────────
        순번 = 1
        prev_가공라인 = ""
        For r = 25 To 34
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
            순번 = 순번 + 1
SkipFin1:
        Next r

        ' ────────────────────────────────────────────────
        ' 5. 완제품 수집 - 생산일보II (행48~63)
        '    c31=구분(JRT/WT/FT/BT), c33=품명, c41=생산량
        '    c45=기소, c49=PLT1(AW), c52=PLT2(AZ), c55=인원
        '    ※ JRT PLT(11)=AW48:AW51, PLT(13)=AZ48:AZ51
        '    TOTAL행 스킵, 가공라인호기 = 구분에서 변환
        ' ────────────────────────────────────────────────
        순번 = 1
        현재구분 = ""
        For r = 48 To 63
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
            순번 = 순번 + 1
SkipFin2:
        Next r

        ' ────────────────────────────────────────────────
        ' 6. 투입인원 - 생산일보I (행40~41) KT
        '    DB 컬럼: F19=작업장, F20=구분, F21=인원수, F22=작업시간, F23=총작업시간, F24=실투입원, F25=환산
        '    행40=정규 → "작업시간" 레코드, 행41=연장 → "연장근로" 레코드
        ' ────────────────────────────────────────────────
        v정규 = ws.Cells(40, 10).Value
        인원유효 = Not IsEmpty(v정규) And IsNumeric(v정규) And CDbl(v정규) <> 0
        v연장 = ws.Cells(40, 6).Value
        시간유효 = Not IsEmpty(v연장) And IsNumeric(v연장) And CDbl(v연장) <> 0
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
            batchArr(batchSize, 21) = IIf(Not IsEmpty(ws.Cells(40, 6).Value) And IsNumeric(ws.Cells(40, 6).Value), CDbl(ws.Cells(40, 6).Value), 0)
            batchArr(batchSize, 22) = IIf(Not IsEmpty(ws.Cells(40, 10).Value) And IsNumeric(ws.Cells(40, 10).Value), CDbl(ws.Cells(40, 10).Value), 0)
            batchArr(batchSize, 23) = IIf(Not IsEmpty(ws.Cells(40, 13).Value) And IsNumeric(ws.Cells(40, 13).Value), CDbl(ws.Cells(40, 13).Value), 0)
            batchArr(batchSize, 24) = IIf(Not IsEmpty(ws.Cells(40, 17).Value) And IsNumeric(ws.Cells(40, 17).Value), CDbl(ws.Cells(40, 17).Value), 0)
            batchArr(batchSize, 25) = IIf(Not IsEmpty(ws.Cells(40, 21).Value) And IsNumeric(ws.Cells(40, 21).Value), CDbl(ws.Cells(40, 21).Value), 0)
        End If

        v정규 = ws.Cells(41, 10).Value
        인원유효 = Not IsEmpty(v정규) And IsNumeric(v정규) And CDbl(v정규) <> 0
        v연장 = ws.Cells(41, 6).Value
        시간유효 = Not IsEmpty(v연장) And IsNumeric(v연장) And CDbl(v연장) <> 0
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
            batchArr(batchSize, 21) = IIf(Not IsEmpty(ws.Cells(41, 6).Value) And IsNumeric(ws.Cells(41, 6).Value), CDbl(ws.Cells(41, 6).Value), 0)
            batchArr(batchSize, 22) = IIf(Not IsEmpty(ws.Cells(41, 10).Value) And IsNumeric(ws.Cells(41, 10).Value), CDbl(ws.Cells(41, 10).Value), 0)
            batchArr(batchSize, 23) = IIf(Not IsEmpty(ws.Cells(41, 13).Value) And IsNumeric(ws.Cells(41, 13).Value), CDbl(ws.Cells(41, 13).Value), 0)
            batchArr(batchSize, 24) = IIf(Not IsEmpty(ws.Cells(41, 17).Value) And IsNumeric(ws.Cells(41, 17).Value), CDbl(ws.Cells(41, 17).Value), 0)
            batchArr(batchSize, 25) = IIf(Not IsEmpty(ws.Cells(41, 21).Value) And IsNumeric(ws.Cells(41, 21).Value), CDbl(ws.Cells(41, 21).Value), 0)
        End If

        ' ────────────────────────────────────────────────
        ' 7. 투입인원 - 생산일보II (행68~77)
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
        For r = 68 To 77
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
                    If CDbl(ws.Cells(r, c2).Value) <> 0 Then hasData2 = True: Exit For
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
            인원유효 = Not IsEmpty(v정규) And IsNumeric(v정규) And CDbl(v정규) <> 0
            v연장 = 인원II정규(idx)
            시간유효 = Not IsEmpty(v연장) And IsNumeric(v연장) And CDbl(v연장) <> 0
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
                batchArr(batchSize, 21) = IIf(Not IsEmpty(인원II인원(idx)) And IsNumeric(인원II인원(idx)), CDbl(인원II인원(idx)), 0)
                batchArr(batchSize, 22) = IIf(Not IsEmpty(인원II정규(idx)) And IsNumeric(인원II정규(idx)), CDbl(인원II정규(idx)), 0)
                batchArr(batchSize, 23) = IIf(Not IsEmpty(인원II총MH(idx)) And IsNumeric(인원II총MH(idx)), CDbl(인원II총MH(idx)), 0)
                batchArr(batchSize, 24) = IIf(Not IsEmpty(인원II실투입(idx)) And IsNumeric(인원II실투입(idx)), CDbl(인원II실투입(idx)), 0)
                batchArr(batchSize, 25) = IIf(Not IsEmpty(인원II환산(idx)) And IsNumeric(인원II환산(idx)), CDbl(인원II환산(idx)), 0)
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
        ' KT 일계
        batchArr(batchSize, 28) = ws.Cells(25, 45).Value   ' KT_일계_원투
        batchArr(batchSize, 29) = ws.Cells(27, 45).Value   ' KT_일계_양품
        batchArr(batchSize, 30) = ws.Cells(29, 45).Value   ' KT_일계_수율
        batchArr(batchSize, 32) = ws.Cells(30, 45).Value   ' KT_일계_LOSS
        batchArr(batchSize, 33) = ws.Cells(31, 45).Value   ' KT_일계_양품율
        ' KT 주누계
        batchArr(batchSize, 34) = ws.Cells(25, 49).Value   ' KT_주누계_원투 (G88)
        batchArr(batchSize, 35) = ws.Cells(27, 49).Value   ' KT_주누계_양품 (L88)
        batchArr(batchSize, 36) = ws.Cells(29, 49).Value   ' KT_주누계_수율 (O88)
        ' KT_주누계_기소(F103): 행24=구분헤더("주")일 수 있음 → 행26 시도 후 IsNumeric 체크
        Dim v101 As Variant: v101 = ws.Cells(26, 49).Value
        If Not IsNumeric(v101) Then v101 = ws.Cells(24, 49).Value
        If IsNumeric(v101) Then batchArr(batchSize, 103) = v101
        ' KT 월누계
        batchArr(batchSize, 37) = ws.Cells(25, 53).Value   ' KT_월누계_원투 (G89)
        batchArr(batchSize, 38) = ws.Cells(27, 53).Value   ' KT_월누계_양품 (L89)
        batchArr(batchSize, 39) = ws.Cells(29, 53).Value   ' KT_월누계_수율 (O89)
        ' KT_월누계_기소(F104): 행24=구분헤더("월")일 수 있음 → 행26 시도 후 IsNumeric 체크
        Dim v102 As Variant: v102 = ws.Cells(26, 53).Value
        If Not IsNumeric(v102) Then v102 = ws.Cells(24, 53).Value
        If IsNumeric(v102) Then batchArr(batchSize, 104) = v102
        ' KT 주/월간평균_기소 (F105,F106): IsNumeric 체크로 "주"/"월" 라벨 입력 방지
        Dim v103 As Variant: v103 = ws.Cells(23, 49).Value
        Dim v104 As Variant: v104 = ws.Cells(23, 53).Value
        If IsNumeric(v103) Then batchArr(batchSize, 105) = v103
        If IsNumeric(v104) Then batchArr(batchSize, 106) = v104

        ' ────────────────────────────────────────────────
        ' 8-2. PA 수집 - 생산일보II FT/JRT/BT 수율 지표
        '    범위: AE64:BA77 (행64=헤더, 행65=빈행, 행66부터 데이터)
        '    열구조: AI(35)=원지투입총량, AM(39)=양품환산총량, AQ(43)=원지수율
        '    JRT: 행66(일계), 행67(주평균), 행68(월평균)
        '    WT:  행69~71 (DB 매핑 없음, 스킵)
        '    FT:  행72(일계), 행73(주평균), 행74(월평균)
        '    BT:  행75(일계), 행76(주평균), 행77(월평균)
        ' ────────────────────────────────────────────────
        ' JRT 일계 (DB F43=원투, F44=양품, F45=수율)
        vII = ws.Cells(66, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 43) = CDbl(vII)
        vII = ws.Cells(66, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 44) = CDbl(vII)
        vII = ws.Cells(66, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 45) = CDbl(vII)
        ' JRT 주평균 (DB F85=원투, F86=양품, F87=수율)
        vII = ws.Cells(67, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 85) = CDbl(vII)
        vII = ws.Cells(67, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 86) = CDbl(vII)
        vII = ws.Cells(67, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 87) = CDbl(vII)
        ' JRT 월평균 (DB F88=원투, F89=양품, F90=수율)
        vII = ws.Cells(68, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 88) = CDbl(vII)
        vII = ws.Cells(68, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 89) = CDbl(vII)
        vII = ws.Cells(68, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 90) = CDbl(vII)

        ' FT 일계 (DB F40=원투, F41=양품, F42=수율)
        vII = ws.Cells(72, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 40) = CDbl(vII)
        vII = ws.Cells(72, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 41) = CDbl(vII)
        vII = ws.Cells(72, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 42) = CDbl(vII)
        ' FT 주평균 (DB F79=원투, F80=양품, F81=수율)
        vII = ws.Cells(73, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 79) = CDbl(vII)
        vII = ws.Cells(73, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 80) = CDbl(vII)
        vII = ws.Cells(73, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 81) = CDbl(vII)
        ' FT 월평균 (DB F82=원투, F83=양품, F84=수율)
        vII = ws.Cells(74, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 82) = CDbl(vII)
        vII = ws.Cells(74, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 83) = CDbl(vII)
        vII = ws.Cells(74, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 84) = CDbl(vII)

        ' BT 일계 (DB F46=원투, F47=양품, F48=수율)
        vII = ws.Cells(75, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 46) = CDbl(vII)
        vII = ws.Cells(75, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 47) = CDbl(vII)
        vII = ws.Cells(75, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 48) = CDbl(vII)
        ' BT 주평균 (DB F91=원투, F92=양품, F93=수율)
        vII = ws.Cells(76, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 91) = CDbl(vII)
        vII = ws.Cells(76, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 92) = CDbl(vII)
        vII = ws.Cells(76, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 93) = CDbl(vII)
        ' BT 월평균 (DB F94=원투, F95=양품, F96=수율)
        vII = ws.Cells(77, 35).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 94) = CDbl(vII)
        vII = ws.Cells(77, 39).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 95) = CDbl(vII)
        vII = ws.Cells(77, 43).Value: If IsNumeric(vII) And Not IsEmpty(vII) Then batchArr(batchSize, 96) = CDbl(vII)

        추가수 = 추가수 + 1

NextSheet:
    Next i

    ' ──────────────────────────────────────────────────────
    ' DB에 일괄 쓰기
    ' ──────────────────────────────────────────────────────
    If batchSize > 0 Then
        새행시작 = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
        wsDB.Range(wsDB.Cells(새행시작, 1), wsDB.Cells(새행시작 + batchSize - 1, 126)).Value = batchArr
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    wbSrc.Close False

    MsgBox "이관 완료!" & vbCrLf & _
           "총 " & batchSize & "건의 레코드가 일일생산DB에 추가되었습니다.", vbInformation, "이관 결과"
    Exit Sub

MigErr:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next: wbSrc.Close False: On Error GoTo 0
    MsgBox "오류 발생!" & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "오류 내용: " & Err.Description & vbCrLf & vbCrLf & _
           "처리 중 시트: " & 시트명 & vbCrLf & _
           "현재 batchSize: " & batchSize, vbCritical, "마이그레이션 오류"
End Sub


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
' 헬퍼: 60%/40% 주간/야간 랜덤 구분 (seed 기반 재현 가능)
' ──────────────────────────────────────────────────────────
Private Function 랜덤구분(seed As Long, n As Long) As String
    Dim v As Double
    v = ((CDbl(seed) * 17 + CDbl(n) * 31) Mod 100)
    If v < 0 Then v = -v
    If v < 60 Then
        랜덤구분 = "생산일보_주간"
    Else
        랜덤구분 = "생산일보_야간"
    End If
End Function



