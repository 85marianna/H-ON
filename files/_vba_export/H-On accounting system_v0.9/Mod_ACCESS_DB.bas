
Function Get_DB_Access(ByVal TableName As String) As Variant
'기존 Get_DB 엑세스버전  (헤더행 없음)


    Dim Conn As Object
    Dim Rs As Object
    Dim arr As Variant
    Dim tmp()
    Dim i As Long, j As Long

    ' DB 연결
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    ' 테이블 동적 조회
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open "SELECT * FROM " & TableName, Conn

    If Not Rs.EOF Then
        arr = Rs.GetRows
        
        ' 배열 방향 맞추기 (안전 버전)
        ReDim tmp(1 To UBound(arr, 2) + 1, 1 To UBound(arr, 1) + 1)
        
        For i = 0 To UBound(arr, 1)
            For j = 0 To UBound(arr, 2)
                tmp(j + 1, i + 1) = arr(i, j)
            Next j
        Next i
        
        Get_DB_Access = tmp
    End If

    Rs.Close
    Conn.Close

End Function



Function Get_DB_Access_Where(TableName As String, whereSql As String) As Variant

'이 함수는 WHERE절을 통째로 넘기는 함수
'Get_DB_Access_Where(테이블명, "WHERE 뒤에 들어갈 문장")
'사용샘플예시(db = Get_DB_Access_Where("즐겨찾기", "즐겨찾기명='" & FavName & "' ORDER BY 명세번호")
'Get_DB_Access_Where("분개장", "전표번호='2026-07-01-001'")
'Get_DB_Access_Where("분개장", "계정명='미지급금'")
'Get_DB_Access_Where("분개장", "거래처='삼성전자' ORDER BY 회계일자 DESC")


    Dim Conn As Object
    Dim Rs As Object
    Dim arr As Variant
    Dim tmp()
    Dim i As Long, j As Long

     db 연결
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

     '테이블 동적 조회
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open "SELECT * FROM " & TableName & " WHERE " & whereSql, Conn
    
    If Not Rs.EOF Then
        arr = Rs.GetRows
        
         '배열 방향 맞추기 (안전 버전)
        ReDim tmp(1 To UBound(arr, 2) + 1, 1 To UBound(arr, 1) + 1)
        
        For i = 0 To UBound(arr, 1)
            For j = 0 To UBound(arr, 2)
                tmp(j + 1, i + 1) = arr(i, j)
            Next j
        Next i
        
        Get_DB_Access_Where = tmp
    End If

    Rs.Close
    Conn.Close

End Function




''Function 원장데이터가져오기_DB( _
''    Optional filterCode As String = "", _
''    Optional dateFrom As Variant = "", _
''    Optional dateTo As Variant = "") As Variant
''
''    '=====================
''    ' 1. DB 연결
''    '=====================
''    Dim Conn As Object
''    Set Conn = CreateObject("ADODB.Connection")
''    Conn.Open Get_ConnStr()
''
''
''    '=====================
''    ' 1. 계정코드마스터 → Dictionary (DB 버전)
''    '=====================
''    Dim dictMaster As Object
''    Set dictMaster = CreateObject("Scripting.Dictionary")
''
''    Dim rsMaster As Object
''    Set rsMaster = CreateObject("ADODB.Recordset")
''    rsMaster.Open "SELECT 세목코드, 계정명, 잔액방향 FROM 계정마스터", Conn, 1, 1
''
''    Do While Not rsMaster.EOF
''        Dim sMasterCode As String
''        sMasterCode = CStr(rsMaster.Fields("세목코드").Value)
''        If Not dictMaster.Exists(sMasterCode) Then
''            dictMaster.Add sMasterCode, Array(rsMaster.Fields("계정명").Value, rsMaster.Fields("잔액방향").Value)
''        End If
''        rsMaster.MoveNext
''    Loop
''
''    rsMaster.Close
''    Set rsMaster = Nothing
''
''
''    '=====================
''    ' 3. SQL 구성 → Recordset 조회
''    '    전기이월 + 조회기간 데이터를 한 번에 가져옴
''    '=====================
''    Dim Sql As String
''    Sql = "SELECT 회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액 " & _
''          "FROM 분개장 "
''
''    ' WHERE 조건 구성
''    Dim conditions As String
''    conditions = ""
''
''    ' 계정코드 필터
''    If filterCode <> "" Then
''        conditions = "세목코드 = '" & filterCode & "'"
''    End If
''
''    ' dateTo 조건 (조회일자 이하 + 전기이월 포함)
''    ' → 전기이월(1/1, 적요='전기이월')은 dateFrom 관계없이 항상 포함해야 하므로
''    '   dateTo만 SQL에서 거르고, dateFrom 미만 처리는 VBA에서 담당
''    If IsDate(dateTo) Then
''        If conditions <> "" Then conditions = conditions & " AND "
''        conditions = conditions & "회계일자 <= #" & Format(CDate(dateTo), "yyyy-mm-dd") & "#"
''    End If
''
''    If conditions <> "" Then Sql = Sql & "WHERE " & conditions & " "
''
''    ' 정렬: 세목코드 → 회계일자 → 전표번호 → 명세번호
''    Sql = Sql & "ORDER BY 세목코드, 회계일자, 전표번호, 명세번호"
''
''    Debug.Print Sql  ' 쿼리 확인용
''
''    Dim Rs As Object
''    Set Rs = CreateObject("ADODB.Recordset")
''    Rs.Open Sql, Conn, 1, 1  ' adOpenKeyset, adLockReadOnly
''
''    '=====================
''    ' 4. Recordset → 배열로 변환
''    '=====================
''    Dim rawList() As Variant
''    Dim rawCount As Long
''    rawCount = 0
''
''    If Not Rs.EOF Then
''        ' 일단 충분한 크기로 확보
''        ReDim rawList(1 To 10000, 1 To 12)
''
''        Do While Not Rs.EOF
''            rawCount = rawCount + 1
''            rawList(rawCount, 1) = Rs.Fields("회계일자").Value
''            rawList(rawCount, 2) = Rs.Fields("전표번호").Value
''            rawList(rawCount, 3) = Rs.Fields("명세번호").Value
''            rawList(rawCount, 4) = Rs.Fields("적요").Value
''            rawList(rawCount, 5) = Rs.Fields("거래처").Value
''            rawList(rawCount, 6) = Rs.Fields("차변").Value
''            rawList(rawCount, 7) = Rs.Fields("대변").Value
''            rawList(rawCount, 8) = Rs.Fields("세목코드").Value
''            rawList(rawCount, 9) = Rs.Fields("계정명").Value
''            rawList(rawCount, 10) = Rs.Fields("부서명").Value
''            rawList(rawCount, 11) = Rs.Fields("통화").Value
''            rawList(rawCount, 12) = Rs.Fields("외화금액").Value
''            Rs.MoveNext
''        Loop
''    End If
''
''    Rs.Close
''    Conn.Close
''    Set Rs = Nothing
''    Set Conn = Nothing
''
''    If rawCount = 0 Then
''        원장데이터가져오기_DB = Empty
''        Exit Function
''    End If
''
''    '=====================
''    ' 5. 행 분류 (Excel 버전과 동일 로직)
''    '    - 전기이월행: 적요="전기이월" AND 1월1일
''    '    - dateFrom 미만: 잔액합산용
''    '    - dateFrom ~ dateTo: 출력용
''    '=====================
''    Dim dictIW    As Object
''    Dim dictPre   As Object
''    Dim dictIndex As Object
''    Set dictIW = CreateObject("Scripting.Dictionary")
''    Set dictPre = CreateObject("Scripting.Dictionary")
''    Set dictIndex = CreateObject("Scripting.Dictionary")
''
''    Dim bDateFrom As Boolean
''    Dim bDateTo   As Boolean
''    bDateFrom = IsDate(dateFrom)
''    bDateTo = IsDate(dateTo)
''
''    Dim rowDate As Date
''    Dim code As String
''
''    For i = 1 To rawCount
''        code = CStr(rawList(i, 8))  ' 세목코드
''
''        ' 전기이월 판단
''        Dim bIsIW As Boolean
''        bIsIW = False
''        If CStr(rawList(i, 4)) = "전기이월" Then
''            Dim iwDate As Date
''            iwDate = CDate(rawList(i, 1))
''            If Month(iwDate) = 1 And Day(iwDate) = 1 Then
''                bIsIW = True
''            End If
''        End If
''
''        If bIsIW Then
''            If Not dictIW.Exists(code) Then dictIW.Add code, New Collection
''            dictIW(code).Add i
''
''        ElseIf Not IsNull(rawList(i, 1)) And rawList(i, 1) <> "" Then
''            rowDate = CDate(rawList(i, 1))
''
''            ' dateFrom 미만 → 잔액합산용
''            If bDateFrom Then
''                If rowDate < CDate(dateFrom) Then
''                    If Not dictPre.Exists(code) Then dictPre.Add code, New Collection
''                    dictPre(code).Add i
''                    GoTo NextRow
''                End If
''            End If
''
''            ' 출력용
''            If Not dictIndex.Exists(code) Then dictIndex.Add code, New Collection
''            dictIndex(code).Add i
''        End If
''
''NextRow:
''    Next i
''
''    '=====================
''    ' 6. 계정코드 수집 + 정렬 (Excel 버전과 동일)
''    '=====================
''    Dim dictAll As Object
''    Set dictAll = CreateObject("Scripting.Dictionary")
''
''    Dim c As Variant
''    For Each c In dictIW.Keys
''        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
''    Next c
''    For Each c In dictPre.Keys
''        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
''    Next c
''    For Each c In dictIndex.Keys
''        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
''    Next c
''
''    Dim codeCount As Long
''    codeCount = dictAll.Count
''    If codeCount = 0 Then
''        원장데이터가져오기_DB = Empty
''        Exit Function
''    End If
''
''    Dim codes() As String
''    ReDim codes(0 To codeCount - 1)
''    Dim idx As Long
''    idx = 0
''    For Each c In dictAll.Keys
''        codes(idx) = CStr(c)
''        idx = idx + 1
''    Next c
''
''    ' 버블정렬 (계정코드 수 적으니 충분)
''    Dim j As Long, tmp As String
''    For i = 0 To codeCount - 2
''        For j = 0 To codeCount - 2 - i
''            If codes(j) > codes(j + 1) Then
''                tmp = codes(j)
''                codes(j) = codes(j + 1)
''                codes(j + 1) = tmp
''            End If
''        Next j
''    Next i
''
''    '=====================
''    ' 7. 출력 배열 준비 + 계정별 누적잔액 계산
''    '    (Excel 버전과 동일 로직 ? rawData 대신 rawList 사용)
''    '=====================
''    Dim totalRows As Long
''    totalRows = rawCount + (codeCount * 2) + 10
''
''    Dim result() As Variant
''    ReDim result(1 To totalRows, 1 To 12)
''
''    Dim outRow As Long
''    outRow = 1
''
''    For i = 0 To codeCount - 1
''        code = codes(i)
''
''        Dim 계정명   As String
''        Dim 잔액방향 As String
''        If dictMaster.Exists(code) Then
''            계정명 = dictMaster(code)(0)
''            잔액방향 = dictMaster(code)(1)
''        Else
''            계정명 = "미등록계정"
''            잔액방향 = "D"
''        End If
''
''        Dim 누적잔액     As Double
''        Dim 소계차변     As Double
''        Dim 소계대변     As Double
''        Dim 누적외화잔액 As Double
''        누적잔액 = 0: 소계차변 = 0: 소계대변 = 0: 누적외화잔액 = 0
''
''        Dim 차변 As Double, 대변 As Double
''        Dim 외화금액raw As Double, 외화증감 As Double
''        Dim rowIdx As Variant
''
''        ' --- 전기이월
''        Dim iw잔액 As Double, iw외화잔액 As Double
''        iw잔액 = 0: iw외화잔액 = 0
''        If dictIW.Exists(code) Then
''            For Each rowIdx In dictIW(code)
''
''                If IsNull(rawList(rowIdx, 6)) Or IsEmpty(rawList(rowIdx, 6)) Then
''                    차변 = 0
''                Else
''                    차변 = CDbl(rawList(rowIdx, 6))
''                End If
''
''                If IsNull(rawList(rowIdx, 7)) Or IsEmpty(rawList(rowIdx, 7)) Then
''                    대변 = 0
''                Else
''                    대변 = CDbl(rawList(rowIdx, 7))
''                End If
''
''                If 잔액방향 = "D" Then
''                    iw잔액 = iw잔액 + 차변 - 대변
''                Else
''                    iw잔액 = iw잔액 + 대변 - 차변
''                End If
''
''                If IsNull(rawList(rowIdx, 12)) Or IsEmpty(rawList(rowIdx, 12)) Then
''                    외화금액raw = 0
''                Else
''                    외화금액raw = CDbl(rawList(rowIdx, 12))
''                End If
''
''                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
''                If 잔액방향 = "D" Then
''                    iw외화잔액 = iw외화잔액 + 외화증감
''                Else
''                    iw외화잔액 = iw외화잔액 - 외화증감
''                End If
''            Next rowIdx
''        End If
''
''        ' --- dateFrom 미만
''        Dim pre잔액 As Double, pre외화잔액 As Double
''        pre잔액 = 0: pre외화잔액 = 0
''        If dictPre.Exists(code) Then
''            For Each rowIdx In dictPre(code)
''
''                If IsNull(rawList(rowIdx, 6)) Or IsEmpty(rawList(rowIdx, 6)) Then
''                    차변 = 0
''                Else
''                    차변 = CDbl(rawList(rowIdx, 6))
''                End If
''
''                If IsNull(rawList(rowIdx, 7)) Or IsEmpty(rawList(rowIdx, 7)) Then
''                    대변 = 0
''                Else
''                    대변 = CDbl(rawList(rowIdx, 7))
''                End If
''
''                If 잔액방향 = "D" Then
''                    pre잔액 = pre잔액 + 차변 - 대변
''                Else
''                    pre잔액 = pre잔액 + 대변 - 차변
''                End If
''
''                If IsNull(rawList(rowIdx, 12)) Or IsEmpty(rawList(rowIdx, 12)) Then
''                    외화금액raw = 0
''                Else
''                    외화금액raw = CDbl(rawList(rowIdx, 12))
''                End If
''
''                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
''                If 잔액방향 = "D" Then
''                    pre외화잔액 = pre외화잔액 + 외화증감
''                Else
''                    pre외화잔액 = pre외화잔액 - 외화증감
''                End If
''            Next rowIdx
''        End If
''
''        ' --- 전기이월 출력행
''        누적잔액 = iw잔액 + pre잔액
''        누적외화잔액 = iw외화잔액 + pre외화잔액
''        result(outRow, 1) = code
''        result(outRow, 2) = 계정명
''        If bDateFrom Then
''            result(outRow, 3) = CDate(dateFrom)
''        Else
''            result(outRow, 3) = ""
''        End If
''        result(outRow, 6) = "전기이월"
''        result(outRow, 10) = 누적잔액
''        result(outRow, 11) = "IW"
''        result(outRow, 12) = 누적외화잔액
''        outRow = outRow + 1
''
''        ' --- 출력용 거래행
''        If dictIndex.Exists(code) Then
''            For Each rowIdx In dictIndex(code)
''
''                If IsNull(rawList(rowIdx, 6)) Or IsEmpty(rawList(rowIdx, 6)) Then
''                    차변 = 0
''                Else
''                    차변 = CDbl(rawList(rowIdx, 6))
''                End If
''
''                If IsNull(rawList(rowIdx, 7)) Or IsEmpty(rawList(rowIdx, 7)) Then
''                    대변 = 0
''                Else
''                    대변 = CDbl(rawList(rowIdx, 7))
''                End If
''
''                If 잔액방향 = "D" Then
''                    누적잔액 = 누적잔액 + 차변 - 대변
''                Else
''                    누적잔액 = 누적잔액 + 대변 - 차변
''                End If
''                소계차변 = 소계차변 + 차변
''                소계대변 = 소계대변 + 대변
''
''                If IsNull(rawList(rowIdx, 12)) Or IsEmpty(rawList(rowIdx, 12)) Then
''                    외화금액raw = 0
''                Else
''                    외화금액raw = CDbl(rawList(rowIdx, 12))
''                End If
''
''                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
''                If 잔액방향 = "D" Then
''                    누적외화잔액 = 누적외화잔액 + 외화증감
''                Else
''                    누적외화잔액 = 누적외화잔액 - 외화증감
''                End If
''                result(outRow, 1) = code
''                result(outRow, 2) = 계정명
''                result(outRow, 3) = rawList(rowIdx, 1)    ' 회계일자
''                result(outRow, 4) = rawList(rowIdx, 2)    ' 전표번호
''                result(outRow, 5) = rawList(rowIdx, 3)    ' 명세번호
''                result(outRow, 6) = rawList(rowIdx, 4)    ' 적요
''                result(outRow, 7) = rawList(rowIdx, 5)    ' 거래처
''                result(outRow, 8) = 차변
''                result(outRow, 9) = 대변
''                result(outRow, 10) = 누적잔액
''                result(outRow, 11) = "D"
''                result(outRow, 12) = 누적외화잔액
''                outRow = outRow + 1
''            Next rowIdx
''        End If
''
''        ' --- 소계행
''        result(outRow, 1) = code
''        result(outRow, 2) = 계정명
''        result(outRow, 6) = "소계"
''        result(outRow, 8) = 소계차변
''        result(outRow, 9) = 소계대변
''        result(outRow, 10) = 누적잔액
''        result(outRow, 11) = "S"
''        result(outRow, 12) = 누적외화잔액
''        outRow = outRow + 1
''
''    Next i
''
''    '=====================
''    ' 8. 실제 사용 행만 잘라서 반환
''    '=====================
''    Dim finalResult() As Variant
''    ReDim finalResult(1 To outRow - 1, 1 To 12)
''    Dim k As Long
''    For i = 1 To outRow - 1
''        For k = 1 To 12
''            finalResult(i, k) = result(i, k)
''        Next k
''    Next i
''
''    원장데이터가져오기_DB = finalResult
''
''End Function



Function 원장데이터가져오기_DB( _
    Optional filterCode As String = "", _
    Optional dateFrom As Variant = "", _
    Optional dateTo As Variant = "") As Variant

    '=====================
    ' 1. DB 연결
    '=====================
    Dim Conn As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()


    '=====================
    ' 2. 계정코드마스터 → Dictionary (DB 버전)
    '=====================
    Dim dictMaster As Object
    Set dictMaster = CreateObject("Scripting.Dictionary")

    Dim rsMaster As Object
    Set rsMaster = CreateObject("ADODB.Recordset")
    rsMaster.Open "SELECT 세목코드, 계정명, 잔액방향 FROM 계정마스터", Conn, 1, 1

    Do While Not rsMaster.EOF
        Dim sMasterCode As String
        sMasterCode = CStr(rsMaster.Fields("세목코드").Value)
        If Not dictMaster.Exists(sMasterCode) Then
            dictMaster.Add sMasterCode, Array(rsMaster.Fields("계정명").Value, rsMaster.Fields("잔액방향").Value)
        End If
        rsMaster.MoveNext
    Loop

    rsMaster.Close
    Set rsMaster = Nothing


    '=====================
    ' 3. SQL 구성 → Recordset 조회
    '    targetYear 기반으로 해당 연도 데이터만 가져옴
    '=====================
    Dim Sql As String
    Sql = "SELECT 연도, 결산이월잔액, 회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, " & _
          "세목코드, 계정명, 부서명, 통화, 외화금액 " & _
          "FROM 분개장 "

    ' WHERE 조건 구성
    Dim conditions As String
    conditions = ""

    ' 연도 필터 (필수) - dateFrom의 year를 기준으로
    Dim targetYear As Long
    If IsDate(dateFrom) Then
        targetYear = Year(CDate(dateFrom))
        conditions = "연도 = " & targetYear
    Else
        원장데이터가져오기_DB = Empty
        Exit Function
    End If

    ' 계정코드 필터 (선택)
    If filterCode <> "" Then
        conditions = conditions & " AND 세목코드 = '" & filterCode & "'"
    End If

    ' dateTo 조건 (선택)
    If IsDate(dateTo) Then
        conditions = conditions & " AND 회계일자 <= #" & Format(CDate(dateTo), "yyyy-mm-dd") & "#"
    End If

    If conditions <> "" Then Sql = Sql & "WHERE " & conditions & " "

    ' 정렬: 세목코드 → 회계일자 → 전표번호 → 명세번호
    Sql = Sql & "ORDER BY 세목코드, 회계일자, 전표번호, 명세번호"

    Debug.Print Sql  ' 쿼리 확인용

    Dim Rs As Object
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Sql, Conn, 1, 1  ' adOpenKeyset, adLockReadOnly

    '=====================
    ' 4. Recordset → 배열로 변환
    '    rawList: (1) 연도 (2) 결산이월잔액 (3) 회계일자 (4) 전표번호 (5) 명세번호
    '             (6) 적요 (7) 거래처 (8) 차변 (9) 대변 (10) 세목코드
    '             (11) 계정명 (12) 부서명 (13) 통화 (14) 외화금액
    '=====================
    Dim rawList() As Variant
    Dim rawCount As Long
    rawCount = 0

    If Not Rs.EOF Then
        ' 일단 충분한 크기로 확보
        ReDim rawList(1 To 10000, 1 To 14)

        Do While Not Rs.EOF
            rawCount = rawCount + 1
            rawList(rawCount, 1) = Rs.Fields("연도").Value
            rawList(rawCount, 2) = Rs.Fields("결산이월잔액").Value
            rawList(rawCount, 3) = Rs.Fields("회계일자").Value
            rawList(rawCount, 4) = Rs.Fields("전표번호").Value
            rawList(rawCount, 5) = Rs.Fields("명세번호").Value
            rawList(rawCount, 6) = Rs.Fields("적요").Value
            rawList(rawCount, 7) = Rs.Fields("거래처").Value
            rawList(rawCount, 8) = Rs.Fields("차변").Value
            rawList(rawCount, 9) = Rs.Fields("대변").Value
            rawList(rawCount, 10) = Rs.Fields("세목코드").Value
            rawList(rawCount, 11) = Rs.Fields("계정명").Value
            rawList(rawCount, 12) = Rs.Fields("부서명").Value
            rawList(rawCount, 13) = Rs.Fields("통화").Value
            rawList(rawCount, 14) = Rs.Fields("외화금액").Value
            Rs.MoveNext
        Loop
    End If

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing

    If rawCount = 0 Then
        원장데이터가져오기_DB = Empty
        Exit Function
    End If

    '=====================
    ' 5. 행 분류 (결산이월잔액 컬럼 기반)
    '    - 전기이월행: 결산이월잔액='Y'
    '    - dateFrom 미만: 잔액합산용 (결산이월잔액='N' AND 회계일자 < dateFrom)
    '    - dateFrom ~ dateTo: 출력용 (결산이월잔액='N' AND 회계일자 >= dateFrom)
    '=====================
    Dim dictIW    As Object
    Dim dictPre   As Object
    Dim dictIndex As Object
    Set dictIW = CreateObject("Scripting.Dictionary")
    Set dictPre = CreateObject("Scripting.Dictionary")
    Set dictIndex = CreateObject("Scripting.Dictionary")

    Dim bDateFrom As Boolean
    Dim bDateTo   As Boolean
    bDateFrom = IsDate(dateFrom)
    bDateTo = IsDate(dateTo)

    Dim rowDate As Date
    Dim code As String

    For i = 1 To rawCount
        code = CStr(rawList(i, 10))  ' 세목코드

        ' 전기이월 판단: 결산이월잔액='Y'
        Dim bIsIW As Boolean
        bIsIW = False
        If CStr(rawList(i, 2)) = "Y" Then
            bIsIW = True
        End If

        If bIsIW Then
            If Not dictIW.Exists(code) Then dictIW.Add code, New Collection
            dictIW(code).Add i

        ElseIf Not IsNull(rawList(i, 3)) And rawList(i, 3) <> "" Then
            rowDate = CDate(rawList(i, 3))

            ' dateFrom 미만 → 잔액합산용
            If bDateFrom Then
                If rowDate < CDate(dateFrom) Then
                    If Not dictPre.Exists(code) Then dictPre.Add code, New Collection
                    dictPre(code).Add i
                    GoTo NextRow
                End If
            End If

            ' 출력용
            If Not dictIndex.Exists(code) Then dictIndex.Add code, New Collection
            dictIndex(code).Add i
        End If

NextRow:
    Next i

    '=====================
    ' 6. 계정코드 수집 + 정렬
    '=====================
    Dim dictAll As Object
    Set dictAll = CreateObject("Scripting.Dictionary")

    Dim c As Variant
    For Each c In dictIW.Keys
        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
    Next c
    For Each c In dictPre.Keys
        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
    Next c
    For Each c In dictIndex.Keys
        If Not dictAll.Exists(CStr(c)) Then dictAll.Add CStr(c), 1
    Next c

    Dim codeCount As Long
    codeCount = dictAll.Count
    If codeCount = 0 Then
        원장데이터가져오기_DB = Empty
        Exit Function
    End If

    Dim codes() As String
    ReDim codes(0 To codeCount - 1)
    Dim idx As Long
    idx = 0
    For Each c In dictAll.Keys
        codes(idx) = CStr(c)
        idx = idx + 1
    Next c

    ' 버블정렬 (계정코드 수 적으니 충분)
    Dim j As Long, tmp As String
    For i = 0 To codeCount - 2
        For j = 0 To codeCount - 2 - i
            If codes(j) > codes(j + 1) Then
                tmp = codes(j)
                codes(j) = codes(j + 1)
                codes(j + 1) = tmp
            End If
        Next j
    Next i

    '=====================
    ' 7. 출력 배열 준비 + 계정별 누적잔액 계산
    '    result 배열: (1) 세목코드 (2) 계정명 (3) 회계일자 (4) 전표번호 (5) 명세번호
    '                 (6) 적요 (7) 거래처 (8) 차변 (9) 대변 (10) 누적잔액
    '                 (11) 행유형(IW/PRE/D/S) (12) 외화잔액
    '=====================
    Dim totalRows As Long
    totalRows = rawCount + (codeCount * 3) + 10

    Dim result() As Variant
    ReDim result(1 To totalRows, 1 To 12)

    Dim outRow As Long
    outRow = 1

    For i = 0 To codeCount - 1
        code = codes(i)

        Dim 계정명   As String
        Dim 잔액방향 As String
        If dictMaster.Exists(code) Then
            계정명 = dictMaster(code)(0)
            잔액방향 = dictMaster(code)(1)
        Else
            계정명 = "미등록계정"
            잔액방향 = "D"
        End If

        Dim 소계차변     As Double
        Dim 소계대변     As Double
        소계차변 = 0: 소계대변 = 0

        Dim 차변 As Double, 대변 As Double
        Dim 외화금액raw As Double, 외화증감 As Double
        Dim rowIdx As Variant

        ' --- 전기이월
        Dim iw잔액 As Double, iw외화잔액 As Double
        iw잔액 = 0: iw외화잔액 = 0
        If dictIW.Exists(code) Then
            For Each rowIdx In dictIW(code)

                If IsNull(rawList(rowIdx, 8)) Or IsEmpty(rawList(rowIdx, 8)) Then
                    차변 = 0
                Else
                    차변 = CDbl(rawList(rowIdx, 8))
                End If

                If IsNull(rawList(rowIdx, 9)) Or IsEmpty(rawList(rowIdx, 9)) Then
                    대변 = 0
                Else
                    대변 = CDbl(rawList(rowIdx, 9))
                End If

                If 잔액방향 = "D" Then
                    iw잔액 = iw잔액 + 차변 - 대변
                Else
                    iw잔액 = iw잔액 + 대변 - 차변
                End If

                If IsNull(rawList(rowIdx, 14)) Or IsEmpty(rawList(rowIdx, 14)) Then
                    외화금액raw = 0
                Else
                    외화금액raw = CDbl(rawList(rowIdx, 14))
                End If

                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    iw외화잔액 = iw외화잔액 + 외화증감
                Else
                    iw외화잔액 = iw외화잔액 - 외화증감
                End If
            Next rowIdx
        End If

        ' --- dateFrom 미만
        Dim pre잔액 As Double, pre외화잔액 As Double
        pre잔액 = 0: pre외화잔액 = 0
        If dictPre.Exists(code) Then
            For Each rowIdx In dictPre(code)

                If IsNull(rawList(rowIdx, 8)) Or IsEmpty(rawList(rowIdx, 8)) Then
                    차변 = 0
                Else
                    차변 = CDbl(rawList(rowIdx, 8))
                End If

                If IsNull(rawList(rowIdx, 9)) Or IsEmpty(rawList(rowIdx, 9)) Then
                    대변 = 0
                Else
                    대변 = CDbl(rawList(rowIdx, 9))
                End If

                If 잔액방향 = "D" Then
                    pre잔액 = pre잔액 + 차변 - 대변
                Else
                    pre잔액 = pre잔액 + 대변 - 차변
                End If

                If IsNull(rawList(rowIdx, 14)) Or IsEmpty(rawList(rowIdx, 14)) Then
                    외화금액raw = 0
                Else
                    외화금액raw = CDbl(rawList(rowIdx, 14))
                End If

                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    pre외화잔액 = pre외화잔액 + 외화증감
                Else
                    pre외화잔액 = pre외화잔액 - 외화증감
                End If
            Next rowIdx
        End If

        ' --- 전기이월 출력행
        Dim iw누적잔액 As Double
        Dim iw누적외화잔액 As Double
        iw누적잔액 = iw잔액
        iw누적외화잔액 = iw외화잔액
        result(outRow, 1) = code
        result(outRow, 2) = 계정명
        result(outRow, 3) = DateSerial(targetYear, 1, 1)  ' 당해년도 1월1일 (전기이월 고정)
        result(outRow, 6) = "전기이월"
        result(outRow, 10) = iw누적잔액
        result(outRow, 11) = "IW"
        result(outRow, 12) = iw누적외화잔액
        outRow = outRow + 1

        ' --- 전월이월 출력행
        Dim pre누적잔액 As Double
        Dim pre누적외화잔액 As Double
        pre누적잔액 = pre잔액
        pre누적외화잔액 = pre외화잔액
        result(outRow, 1) = code
        result(outRow, 2) = 계정명
        result(outRow, 3) = CDate(dateFrom)
        result(outRow, 6) = "전월이월"
        result(outRow, 10) = pre누적잔액
        result(outRow, 11) = "PRE"
        result(outRow, 12) = pre누적외화잔액
        outRow = outRow + 1

        ' --- 출력용 거래행
        Dim 거래누적잔액 As Double
        Dim 거래누적외화잔액 As Double
        거래누적잔액 = iw누적잔액 + pre누적잔액
        거래누적외화잔액 = iw누적외화잔액 + pre누적외화잔액

        If dictIndex.Exists(code) Then
            For Each rowIdx In dictIndex(code)

                If IsNull(rawList(rowIdx, 8)) Or IsEmpty(rawList(rowIdx, 8)) Then
                    차변 = 0
                Else
                    차변 = CDbl(rawList(rowIdx, 8))
                End If

                If IsNull(rawList(rowIdx, 9)) Or IsEmpty(rawList(rowIdx, 9)) Then
                    대변 = 0
                Else
                    대변 = CDbl(rawList(rowIdx, 9))
                End If

                If 잔액방향 = "D" Then
                    거래누적잔액 = 거래누적잔액 + 차변 - 대변
                Else
                    거래누적잔액 = 거래누적잔액 + 대변 - 차변
                End If
                소계차변 = 소계차변 + 차변
                소계대변 = 소계대변 + 대변

                If IsNull(rawList(rowIdx, 14)) Or IsEmpty(rawList(rowIdx, 14)) Then
                    외화금액raw = 0
                Else
                    외화금액raw = CDbl(rawList(rowIdx, 14))
                End If

                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    거래누적외화잔액 = 거래누적외화잔액 + 외화증감
                Else
                    거래누적외화잔액 = 거래누적외화잔액 - 외화증감
                End If
                result(outRow, 1) = code
                result(outRow, 2) = 계정명
                result(outRow, 3) = rawList(rowIdx, 3)    ' 회계일자
                result(outRow, 4) = rawList(rowIdx, 4)    ' 전표번호
                result(outRow, 5) = rawList(rowIdx, 5)    ' 명세번호
                result(outRow, 6) = rawList(rowIdx, 6)    ' 적요
                result(outRow, 7) = rawList(rowIdx, 7)    ' 거래처
                result(outRow, 8) = 차변
                result(outRow, 9) = 대변
                result(outRow, 10) = 거래누적잔액
                result(outRow, 11) = "D"
                result(outRow, 12) = 거래누적외화잔액
                outRow = outRow + 1
            Next rowIdx
        End If

        ' --- 소계행
        result(outRow, 1) = code
        result(outRow, 2) = 계정명
        result(outRow, 6) = "소계"
        result(outRow, 8) = 소계차변
        result(outRow, 9) = 소계대변
        result(outRow, 10) = 거래누적잔액
        result(outRow, 11) = "S"
        result(outRow, 12) = 거래누적외화잔액
        outRow = outRow + 1

    Next i

    '=====================
    ' 8. 실제 사용 행만 잘라서 반환
    '=====================
    Dim finalResult() As Variant
    ReDim finalResult(1 To outRow - 1, 1 To 12)
    Dim k As Long
    For i = 1 To outRow - 1
        For k = 1 To 12
            finalResult(i, k) = result(i, k)
        Next k
    Next i

    원장데이터가져오기_DB = finalResult

End Function






Public Function 외화평가전표존재(evalDate As Date, Optional ByRef voucherNo As String = "") As Boolean

    Dim Conn As Object
    Dim Rs As Object
    Dim Sql As String
    Dim remark As String

    remark = Format(evalDate, "yyyy년 m월") & " 외화예금평가"

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Sql = "SELECT TOP 1 전표번호 " & _
          "FROM 분개장 " & _
          "WHERE 회계일자 = #" & Format(evalDate, "yyyy-mm-dd") & "# " & _
          "AND 적요 = '" & remark & "'"

    Set Rs = Conn.Execute(Sql)

    If Not Rs.EOF Then

        voucherNo = Rs.Fields("전표번호").Value
        외화평가전표존재 = True

    Else

        외화평가전표존재 = False

    End If

    Rs.Close
    Conn.Close

    Set Rs = Nothing
    Set Conn = Nothing

End Function


Public Function CreateReverseVoucher(Conn As Object, OrgVoucherNo As String, ReverseDate As Date) As String
'반대분개 전표 생성 (현재는 외화평가 취소 전용, 금액을 음수로 넣어서 반대분개 처리!!)

    Dim Rs As Object
    Dim Sql As String

    Dim NewVoucherNo As String
    Dim debit As Currency
    Dim credit As Currency

    '==========================
    '1. 새 전표번호 생성
    '==========================
    NewVoucherNo = Get_NextVoucherNo(Conn, ReverseDate)

    '==========================
    '2. 원전표 조회
    '==========================
    Sql = "SELECT * FROM 분개장 " & _
          "WHERE 전표번호='" & OrgVoucherNo & "' " & _
          "ORDER BY 명세번호"

    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Sql, Conn, 1, 1

    '==========================
    '3. 반대분개 생성
    '==========================
    Do Until Rs.EOF

        debit = Rs("차변").Value
        credit = Rs("대변").Value

        Sql = "INSERT INTO 분개장 " & _
              "(회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 입력일시, 입력자) " & _
              "VALUES (" & _
              "#" & Format(ReverseDate, "yyyy-mm-dd") & "#," & _
              "'" & NewVoucherNo & "'," & _
              Rs("명세번호").Value & "," & _
              "'" & Rs("적요").Value & " 취소'," & _
              "'" & Rs("거래처").Value & "'," & _
              (-debit) & "," & _
              (-credit) & "," & _
              "'" & Rs("세목코드").Value & "'," & _
              "'" & Rs("계정명").Value & "'," & _
              "'" & Rs("부서명").Value & "'," & _
              "'" & Rs("통화").Value & "'," & _
              Rs("외화금액").Value & "," & _
              Rs("환율").Value & "," & _
              "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#," & _
              "'" & Application.UserName & "')"

        Conn.Execute Sql

        Rs.MoveNext

    Loop

    Rs.Close
    Set Rs = Nothing

    CreateReverseVoucher = NewVoucherNo

End Function


Function 연결_DB_Access(db As Variant, ForeignID_Fields As Variant, FromTable As String, Fields As String, Optional IncludeHeader As Boolean = False)

Dim cRow As Long: Dim cCol As Long
Dim vForeignID_Fields As Variant: Dim vForeignID_Field As Variant
Dim ForeignID As Variant
Dim vFields As Variant: Dim vField As Variant
Dim vID As Variant: Dim vFieldNo As Variant
Dim Dict As Object
Dim i As Long: Dim j As Long
Dim AddCols As Long

cRow = UBound(db, 1)
cCol = UBound(db, 2)

If InStr(1, Fields, ",") > 1 Then
    AddCols = Len(Fields) - Len(Replace(Fields, ",", "")) + 1
    vFields = Split(Fields, ",")
Else
    AddCols = 1
    vFields = Array(Fields)
End If

ReDim Preserve db(1 To cRow, 1 To cCol + AddCols)

Set Dict = 생성_Dict_Access(FromTable, "거래처")   ' ← 여기만 교체
vID = Dict("거래처")

ReDim vFieldNo(0 To UBound(vFields))
For Each vField In vFields
    For i = 1 To UBound(vID)
        If vID(i) = Trim(vField) Then vFieldNo(j) = i: j = j + 1
    Next
Next

If InStr(1, ForeignID_Fields, ",") > 0 Then vForeignID_Fields = Split(ForeignID_Fields, ",") Else vForeignID_Fields = Array(ForeignID_Fields)
For Each vForeignID_Field In vForeignID_Fields
    For i = 1 To cRow
        If IncludeHeader = True And i = 1 Then ForeignID = "ID" Else ForeignID = db(i, Trim(vForeignID_Field))
        If Dict.Exists(ForeignID) Then
            For j = 1 To AddCols
                db(i, cCol + j) = Dict(ForeignID)(vFieldNo(j - 1))
            Next
        End If
    Next
Next

연결_DB_Access = db

End Function


Function 생성_Dict_Access(TableName As String, KeyFieldName As String) As Object

    Dim Conn As Object, Rs As Object
    Dim Dict As Object
    Dim fieldNames() As String
    Dim arr As Variant
    Dim i As Long
    Dim keyColIdx As Long

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open "SELECT * FROM [" & TableName & "]", Conn, 3, 1

    ' 필드명 배열 (1-based) ? vFieldNo 매칭용, 원본 로직에서 vID 역할
    ReDim fieldNames(1 To Rs.Fields.Count)
    For i = 1 To Rs.Fields.Count
        fieldNames(i) = Rs.Fields(i - 1).Name
        If Rs.Fields(i - 1).Name = KeyFieldName Then keyColIdx = i
    Next i

    Set Dict = CreateObject("Scripting.Dictionary")
    Dict("거래처") = fieldNames   ' 원본과 동일하게 헤더는 "거래처" 키에 저장

    Do While Not Rs.EOF
        ReDim arr(1 To Rs.Fields.Count)
        For i = 1 To Rs.Fields.Count
            arr(i) = Rs.Fields(i - 1).Value & ""   ' Null-safe 문자열 변환
        Next i
        If Not Dict.Exists(arr(keyColIdx)) Then
            Dict(arr(keyColIdx)) = arr
        End If
        Rs.MoveNext
    Loop

    Rs.Close: Set Rs = Nothing
    Conn.Close: Set Conn = Nothing

    Set 생성_Dict_Access = Dict

End Function
