Option Explicit

Sub 총계정원장생성_DB(dateFrom As Date, dateTo As Date)

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    '=====================
    ' 1. 데이터 가져오기 (공용함수 호출)
    '=====================
    Dim data As Variant
    data = 원장데이터가져오기_DB("", dateFrom, dateTo)

    If IsEmpty(data) Then
        MsgBox "데이터가 없습니다.", vbExclamation
        GoTo CleanUp
    End If

    '=====================
    ' 2. 출력 시트 준비
    '=====================
    Dim wbNew As Workbook
    Dim ws원장 As Worksheet

    Set wbNew = Workbooks.Add
    Set ws원장 = wbNew.Worksheets(1)

    On Error Resume Next
    ws원장.Name = "총계정원장"
    On Error GoTo ErrorHandler

    '=====================
    ' 3. 타이틀 출력
    '=====================
    With ws원장.Range("A1:J1")
        .Merge
        .Value = "총계정원장   " & Format(dateFrom, "yyyy-mm-dd") & " ~ " & Format(dateTo, "yyyy-mm-dd")
        .Font.Bold = True
        .Font.Size = 13
    End With

    '=====================
    ' 4. 헤더 출력
    '=====================
    Dim headers As Variant
    headers = Array("계정코드", "계정명", "회계일자", "전표번호", "명세번호", _
                    "적요", "거래처", "차변", "대변", "잔액")
    Dim h As Integer
    For h = 0 To 9
        ws원장.Cells(2, h + 1).Value = headers(h)
    Next h

    '=====================
    ' 5. 데이터 한방 출력
    '=====================
    Dim totalRows As Long
    totalRows = UBound(data, 1)

    Dim outData() As Variant
    ReDim outData(1 To totalRows, 1 To 10)

    Dim i As Long
    For i = 1 To totalRows
        Dim k As Integer
        For k = 1 To 10
            outData(i, k) = data(i, k)
        Next k
    Next i

    ws원장.Range("A3").Resize(totalRows, 10).Value = outData

    '=====================
    ' 6. 서식 처리
    '=====================
    With ws원장.Range("A2:J2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws원장.Columns("C").NumberFormat = "yyyy-mm-dd"
    ws원장.Columns("H").NumberFormat = "#,##0"
    ws원장.Columns("I").NumberFormat = "#,##0"
    ws원장.Columns("J").NumberFormat = "#,##0"

    ' 소계행 스타일
    For i = 1 To totalRows
        If data(i, 11) = "S" Then
            With ws원장.Range(ws원장.Cells(i + 2, 1), ws원장.Cells(i + 2, 10))
                .Interior.Color = RGB(242, 242, 242)
                .Font.Bold = True
            End With
        End If
    Next i

    ws원장.Columns("A:J").AutoFit
    ws원장.Activate
    
    ' 서식 처리 끝난 후 인쇄페이지 설정
    Call SetPrintOptions(ws원장, "$1:$2")

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "총계정원장 생성 완료!" & vbNewLine & _
           "소요시간: " & Format(Timer - startTime, "0.00") & "초", vbInformation
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "오류 발생: " & Err.Description, vbCritical

End Sub


Sub 계정별명세출력_DB(filterCode As String, dateFrom As Variant, dateTo As Variant)

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    '=====================
    ' 1. 데이터 가져오기
    '=====================
    Dim data As Variant
    data = 원장데이터가져오기_DB(filterCode, dateFrom, dateTo)

    If IsEmpty(data) Then
        MsgBox "조회된 데이터가 없습니다.", vbExclamation
        GoTo CleanUp
    End If
    

    '=====================
    ' 2. 출력 시트 준비
    '    시트명: 계정별명세
    '=====================
    Dim wbNew As Workbook
    Dim ws명세 As Worksheet

    Set wbNew = Workbooks.Add
    Set ws명세 = wbNew.Worksheets(1)

    On Error Resume Next
    ws명세.Name = "계정명세"
    On Error GoTo ErrorHandler
      

    '=====================
    ' 3. 타이틀 출력
    '=====================
    Dim 계정명 As String
    계정명 = data(1, 2)  ' 첫번째 행에서 계정명 가져오기

    With ws명세.Range("A1:J1")
        .Merge
        .Value = 계정명 & " (" & filterCode & ")  " & _
                 Format(CDate(dateFrom), "yyyy-mm-dd") & " ~ " & _
                 Format(CDate(dateTo), "yyyy-mm-dd")
        .Font.Bold = True
        .Font.Size = 13
    End With

    '=====================
    ' 4. 헤더 출력
    '=====================
    Dim headers As Variant
    headers = Array("계정코드", "계정명", "회계일자", "전표번호", "명세번호", _
                    "적요", "거래처", "차변", "대변", "잔액")
    Dim h As Integer
    For h = 0 To 9
        ws명세.Cells(2, h + 1).Value = headers(h)
    Next h

    '=====================
    ' 5. 데이터 출력
    '=====================
    Dim totalRows As Long
    totalRows = UBound(data, 1)

    Dim outData() As Variant
    ReDim outData(1 To totalRows, 1 To 10)

    Dim i As Long, k As Integer
    For i = 1 To totalRows
        For k = 1 To 10
            outData(i, k) = data(i, k)
        Next k
    Next i

    ws명세.Range("A3").Resize(totalRows, 10).Value = outData

    '=====================
    ' 6. 서식 처리
    '=====================
    ' 헤더행
    With ws명세.Range("A2:J2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws명세.Columns("C").NumberFormat = "yyyy-mm-dd"
    ws명세.Columns("H").NumberFormat = "#,##0"
    ws명세.Columns("I").NumberFormat = "#,##0"
    ws명세.Columns("J").NumberFormat = "#,##0"

    ' 소계행 스타일
    For i = 1 To totalRows
        If data(i, 11) = "S" Then
            With ws명세.Range(ws명세.Cells(i + 2, 1), ws명세.Cells(i + 2, 10))
                .Interior.Color = RGB(68, 114, 196)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            End With
        End If
    Next i

    ws명세.Columns("A:J").AutoFit
    ws명세.Activate
    
   ' 서식 처리 끝난 후 인쇄페이지 설정
    Call SetPrintOptions(ws명세, "$1:$2")

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "계정별명세 생성 완료!" & vbNewLine & _
           "소요시간: " & Format(Timer - startTime, "0.00") & "초", vbInformation
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "오류 발생: " & Err.Description, vbCritical

End Sub


Sub 분개장출력_DB(dateFrom As Date, dateTo As Date)

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    '=====================
    ' 1. 데이터 가져오기 (분개장 테이블) : 전기이월 데이터는 쿼리 단계에서 제외.
    '====================
    Dim data As Variant
    Dim Conn As Object, Rs As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()
    
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open "SELECT 회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명 " & _
            "FROM 분개장 " & _
            "WHERE 회계일자 >= #" & Format(dateFrom, "yyyy-mm-dd") & "# " & _
            "AND 회계일자 <= #" & Format(dateTo, "yyyy-mm-dd") & "# " & _
            "AND 적요 <> '전기이월' " & _
            "AND 전표번호 IS NOT NULL " & _
            "ORDER BY 전표번호, 명세번호", Conn, 1, 1
    
    If Rs.RecordCount = 0 Then
        MsgBox "데이터가 없습니다.", vbExclamation
        Rs.Close
        Conn.Close
        GoTo CleanUp
    End If
    
    ' 배열로 변환
    data = Rs.GetRows()
    Rs.Close
    Conn.Close

    '=====================
    ' 2. 출력 시트 준비
    '=====================
    Dim wbNew As Workbook
    Dim ws분개장 As Worksheet

    Set wbNew = Workbooks.Add
    Set ws분개장 = wbNew.Worksheets(1)

    On Error Resume Next
    ws분개장.Name = "분개장"
    On Error GoTo ErrorHandler

    '=====================
    ' 3. 타이틀 출력
    '=====================
    With ws분개장.Range("A1:I1")
        .Merge
        .Value = "분개장   " & Format(dateFrom, "yyyy-mm-dd") & " ~ " & Format(dateTo, "yyyy-mm-dd")
        .Font.Bold = True
        .Font.Size = 13
    End With

    '=====================
    ' 4. 헤더 출력
    '=====================
    Dim headers As Variant
    headers = Array("회계일자", "전표번호", "명세번호", "적요", "거래처", "차변", "대변", "세목코드", "계정명")
    Dim h As Integer
    For h = 0 To 8
        ws분개장.Cells(2, h + 1).Value = headers(h)
    Next h
    

''    '=====================
''    ' 5. 데이터 출력
''    '=====================
''    Dim totalRows As Long
''    Dim i As Long, j As Integer
''
''    totalRows = UBound(Data, 2) + 1
''''
''    ' 행 단위로 데이터 넣기
''    For i = 0 To UBound(Data, 2)
''        For j = 0 To 8
''            ws분개장.Cells(i + 3, j + 1).Value = Data(j, i)
''        Next j
''    Next i
        
    
    '=====================
    ' 5. 데이터 출력 + 전표번호별 소계
    '=====================
    Dim totalRows As Long
    Dim i As Long, j As Integer
    Dim currentRow As Long
    Dim lastVoucher As String
    Dim sumDebit As Double, sumCredit As Double

    totalRows = UBound(data, 2) + 1
    currentRow = 3
    lastVoucher = ""
    sumDebit = 0
    sumCredit = 0
    
    For i = 0 To UBound(data, 2)
        Dim currentVoucher As String
        currentVoucher = CStr(data(1, i))
        
        ' 전표번호가 바뀌면 이전 전표의 소계 출력
        If i > 0 And currentVoucher <> lastVoucher Then
            ws분개장.Cells(currentRow, 2).Value = lastVoucher & " 소계"
            ws분개장.Cells(currentRow, 6).Value = sumDebit
            ws분개장.Cells(currentRow, 7).Value = sumCredit
            
            With ws분개장.Range(ws분개장.Cells(currentRow, 1), ws분개장.Cells(currentRow, 9))
                .Interior.Color = RGB(242, 242, 242)
                .Font.Bold = True
            End With
            
            currentRow = currentRow + 1
            sumDebit = 0
            sumCredit = 0
        End If
        
        ' 데이터행 출력
        For j = 0 To 8
            ws분개장.Cells(currentRow, j + 1).Value = data(j, i)
        Next j
        
        ' 차변, 대변 합계에 더하기 (Null 체크)
        If Not IsNull(data(5, i)) Then
            sumDebit = sumDebit + CDbl(data(5, i))
        End If
        
        If Not IsNull(data(6, i)) Then
            sumCredit = sumCredit + CDbl(data(6, i))
        End If
        
        lastVoucher = currentVoucher
        currentRow = currentRow + 1
    Next i
    
    ' 마지막 전표 소계 출력
    ws분개장.Cells(currentRow, 2).Value = lastVoucher & " 소계"
    ws분개장.Cells(currentRow, 6).Value = sumDebit
    ws분개장.Cells(currentRow, 7).Value = sumCredit
    
    With ws분개장.Range(ws분개장.Cells(currentRow, 1), ws분개장.Cells(currentRow, 9))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
    End With


    '=====================
    ' 6. 서식 처리
    '=====================
    With ws분개장.Range("A2:I2")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' 날짜 서식
    ws분개장.Columns("A").NumberFormat = "yyyy-mm-dd"
    ws분개장.Columns("N").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    ' 금액 서식 (차변, 대변, 외화금액)
    ws분개장.Columns("F").NumberFormat = "#,##0"
    ws분개장.Columns("G").NumberFormat = "#,##0"
    ws분개장.Columns("L").NumberFormat = "#,##0.00"
    
    ' 환율 서식
    ws분개장.Columns("M").NumberFormat = "0.0000"

    ws분개장.Columns("A:J").AutoFit
    ws분개장.Activate
    
    Call SetPrintOptions(ws분개장, "$1:$2")

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "분개장 출력 완료!" & vbNewLine & _
           "소요시간: " & Format(Timer - startTime, "0.00") & "초", vbInformation
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "오류 발생: " & Err.Description, vbCritical

End Sub


'===========================================
' 합계잔액시산표 생성 (최종 완성본)
' 모듈: HJ_Mod_Report
'
' [최종 기능]
' 1) 유형 헤더를 맨 처음에 출력 (【자산】-> 유동자산 -> 당좌자산 ... 계층적)
' 2) 잔액(음수)은 그대로, 합계(차변/대변)는 양쪽 모두 표시
' 3) 계정코드 열 제거 (A/B/C/D/E 5열 구조)
' 4) 셀 색상 제거
'===========================================
Public Sub 합계잔액시산표생성_DB(dateFrom As Date, dateTo As Date)

    Dim wsResult As Worksheet
    Dim arrData() As Variant
    Dim tStart As Double
    tStart = Timer
    
    Dim pdfFile As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ------------------------------------------
    ' 1. 원장 데이터 가져오기
    ' ------------------------------------------
    arrData = 원장데이터가져오기_DB("", dateFrom, dateTo)
    If Not IsArray(arrData) Then
        MsgBox "원장 데이터를 가져올 수 없습니다.", vbExclamation
        GoTo CleanExit
    End If

    ' ------------------------------------------
    ' 2. 계정마스터 로드
    ' ------------------------------------------
    Dim Conn As Object, rsMaster As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set rsMaster = CreateObject("ADODB.Recordset")
''    rsMaster.Open "SELECT 세목코드, 계정명, 계정유형, 재무제표, 잔액방향, " & _
''                  "계정구분, 계정구분명, 중분류2, 중분류3 " & _
''                  "FROM 계정마스터 WHERE 사용여부='Y' " & _
''                  "ORDER BY 계정구분, 세목코드", Conn, 1, 1

    Set rsMaster = CreateObject("ADODB.Recordset")
    rsMaster.Open "SELECT 세목코드, 계정명, 계정유형, 재무제표, 잔액방향, " & _
                  "계정구분, 계정구분명, 중분류2, 중분류3 " & _
                  "FROM 계정마스터 " & _
                  "ORDER BY 계정구분, 세목코드", Conn, 1, 1

    Dim nMaster As Long
    nMaster = 0
    Do While Not rsMaster.EOF
        nMaster = nMaster + 1
        rsMaster.MoveNext
    Loop

    If nMaster = 0 Then
        MsgBox "사용여부=Y인 계정이 없습니다.", vbExclamation
        rsMaster.Close: Conn.Close
        GoTo CleanExit
    End If
    rsMaster.MoveFirst

    Dim arrM() As Variant
    ReDim arrM(1 To nMaster, 1 To 9)
    Dim m As Long
    m = 0
    Do While Not rsMaster.EOF
        m = m + 1
        arrM(m, 1) = CStr(rsMaster.Fields("세목코드").Value)
        arrM(m, 2) = rsMaster.Fields("계정명").Value
        arrM(m, 3) = rsMaster.Fields("계정유형").Value
        arrM(m, 4) = rsMaster.Fields("재무제표").Value
        arrM(m, 5) = CStr(rsMaster.Fields("잔액방향").Value)
        arrM(m, 6) = CStr(rsMaster.Fields("계정구분").Value)
        arrM(m, 7) = rsMaster.Fields("계정구분명").Value

        If IsNull(rsMaster.Fields("중분류2").Value) Then
            arrM(m, 8) = ""
        Else
            arrM(m, 8) = CStr(rsMaster.Fields("중분류2").Value)
        End If

        If IsNull(rsMaster.Fields("중분류3").Value) Then
            arrM(m, 9) = ""
        Else
            arrM(m, 9) = CStr(rsMaster.Fields("중분류3").Value)
        End If

        rsMaster.MoveNext
    Loop
    rsMaster.Close
    Conn.Close
    Set rsMaster = Nothing: Set Conn = Nothing

    ' ------------------------------------------
    ' 3. S행 / IW행 수집
    ' ------------------------------------------
    Dim nS As Long, i As Long
    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then nS = nS + 1
    Next i

    Dim arrSCode()   As String
    Dim arrSDebit()  As Double
    Dim arrSCredit() As Double
    Dim arrSBal()    As Double
    ReDim arrSCode(1 To nS)
    ReDim arrSDebit(1 To nS)
    ReDim arrSCredit(1 To nS)
    ReDim arrSBal(1 To nS)

    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then
            nS = nS + 1
            arrSCode(nS) = CStr(arrData(i, 1))
            arrSDebit(nS) = CDbl(IIf(IsNumeric(arrData(i, 8)), arrData(i, 8), 0))
            arrSCredit(nS) = CDbl(IIf(IsNumeric(arrData(i, 9)), arrData(i, 9), 0))
            arrSBal(nS) = CDbl(IIf(IsNumeric(arrData(i, 10)), arrData(i, 10), 0))
        End If
    Next i

    Dim nIW As Long
    nIW = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "IW" Then nIW = nIW + 1
    Next i

    Dim arrIWCode() As String
    Dim arrIWBal()  As Double
    ReDim arrIWCode(1 To nIW)
    ReDim arrIWBal(1 To nIW)

    nIW = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "IW" Then
            nIW = nIW + 1
            arrIWCode(nIW) = CStr(arrData(i, 1))
            arrIWBal(nIW) = CDbl(IIf(IsNumeric(arrData(i, 10)), arrData(i, 10), 0))
        End If
    Next i

    ' ------------------------------------------
    ' 4. 신규 워크북 생성 및 헤더 추가
    ' ------------------------------------------
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add
    Set wsResult = wbNew.Worksheets(1)
    On Error Resume Next
    wsResult.Name = "합계잔액시산표"
    On Error GoTo 0

    ' 제목
    wsResult.Cells(1, 3).Value = "【합계잔액시산표】"
    wsResult.Cells(1, 3).Font.Bold = True
    wsResult.Cells(1, 3).Font.Size = 14

    ' 기준년월
    Dim maxDate As Date
    maxDate = DateSerial(1900, 1, 1)
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "D" Then
            If IsDate(arrData(i, 3)) Then
                If CDate(arrData(i, 3)) > maxDate Then maxDate = CDate(arrData(i, 3))
            End If
        End If
    Next i
    wsResult.Cells(2, 1).Value = "기준년월"
    wsResult.Cells(2, 2).Value = Format(maxDate, "YYYY년 MM월")

    ' 구간 헤더
    wsResult.Cells(3, 1).Value = "차변"
    wsResult.Cells(3, 4).Value = "대변"
    
    With wsResult.Range("C3:C4")
        .Merge
        .Value = "계정과목"
    End With

    ' 열 헤더
    wsResult.Cells(4, 1).Value = "잔액"
    wsResult.Cells(4, 2).Value = "합계"
    wsResult.Cells(4, 4).Value = "합계"
    wsResult.Cells(4, 5).Value = "잔액"
    With wsResult.Range("A3:E4")
        .Font.Bold = True
    End With

    Dim outRow As Long
    outRow = 5

    Dim typeList(1 To 6) As String
    typeList(1) = "자산"
    typeList(2) = "부채"
    typeList(3) = "자본"
    typeList(4) = "수익"
    typeList(5) = "비용"
    typeList(6) = "제조비용"

    ' ------------------------------------------
    ' 5. 유형 -> 중분류2 -> 중분류3(있으면) -> 계정구분명
    ' ------------------------------------------
    Dim t As Integer
    Dim tSumD As Double, tSumC As Double, tBalD As Double, tBalC As Double
    Dim totalSumD As Double, totalSumC As Double, totalBalD As Double, totalBalC As Double
    totalSumD = 0: totalSumC = 0: totalBalD = 0: totalBalC = 0

    Dim lvlCode(1 To 3) As String
    Dim lvlName(1 To 3) As String
    Dim lvlHeaderRow(1 To 3) As Long
    Dim lvlSumD(1 To 3) As Double
    Dim lvlSumC(1 To 3) As Double
    Dim lvlBalD(1 To 3) As Double
    Dim lvlBalC(1 To 3) As Double

    Dim typeHeaderRow As Long

    For t = 1 To 6
        Dim tLabel As String
        tLabel = typeList(t)

        Dim hasType As Boolean
        hasType = False
        For m = 1 To nMaster
            If CStr(arrM(m, 3)) = tLabel Then hasType = True: Exit For
        Next m
        If Not hasType Then GoTo NextType

        ' *** 유형 헤더를 맨 처음에 출력 ***
        typeHeaderRow = outRow
        wsResult.Cells(outRow, 3).Value = "【" & tLabel & "】"
        wsResult.Cells(outRow, 3).Font.Bold = True
        outRow = outRow + 1

        tSumD = 0: tSumC = 0: tBalD = 0: tBalC = 0
        Dim lv As Integer
        For lv = 1 To 3
            lvlCode(lv) = "|__초기값__|"
            lvlHeaderRow(lv) = 0
            lvlSumD(lv) = 0: lvlSumC(lv) = 0: lvlBalD(lv) = 0: lvlBalC(lv) = 0
        Next lv

        For m = 1 To nMaster
            If CStr(arrM(m, 3)) <> tLabel Then GoTo NextMaster

            Dim newCode(1 To 3) As String
            Dim newName(1 To 3) As String
            newCode(1) = CStr(arrM(m, 8)): newName(1) = CStr(arrM(m, 8))
            newCode(2) = CStr(arrM(m, 9)): newName(2) = CStr(arrM(m, 9))
            newCode(3) = CStr(arrM(m, 6)): newName(3) = CStr(arrM(m, 7))

            Dim changeLv As Integer
            changeLv = 4
            For lv = 1 To 3
                If newCode(lv) <> lvlCode(lv) Then
                    changeLv = lv
                    Exit For
                End If
            Next lv

            If changeLv <= 3 Then
                For lv = 3 To changeLv Step -1
                    If lvlHeaderRow(lv) > 0 Then
                        Call 헤더마감(wsResult, lvlHeaderRow(lv), lvlBalD(lv), lvlSumD(lv), lvlSumC(lv), lvlBalC(lv))
                        lvlSumD(lv) = 0: lvlSumC(lv) = 0: lvlBalD(lv) = 0: lvlBalC(lv) = 0
                        lvlHeaderRow(lv) = 0
                    End If
                Next lv

                For lv = changeLv To 3
                    lvlCode(lv) = newCode(lv)
                    lvlName(lv) = newName(lv)
                    If lv = 3 Then
                        lvlHeaderRow(lv) = outRow
                        wsResult.Cells(outRow, 3).Value = Space(lv * 3) & newName(lv)
                        outRow = outRow + 1
                    ElseIf newCode(lv) <> "" Then
                        lvlHeaderRow(lv) = outRow
                        wsResult.Cells(outRow, 3).Value = Space(lv * 3) & newName(lv)
                        outRow = outRow + 1
                    Else
                        lvlHeaderRow(lv) = 0
                    End If
                Next lv
            End If

            ' ---- 개별 세목 금액 계산 ----
            Dim dSumD As Double, dSumC As Double, dBal As Double, dBalD As Double, dBalC As Double
            Dim sBalDir As String
            dSumD = 0: dSumC = 0: dBal = 0: dBalD = 0: dBalC = 0
            sBalDir = CStr(arrM(m, 5))

            ' 원장에서 차변/대변/잔액 수집 (둘 다!)
            For i = 1 To nS
                If arrSCode(i) = CStr(arrM(m, 1)) Then
                    dSumD = arrSDebit(i)
                    dSumC = arrSCredit(i)
                    dBal = arrSBal(i)
                    Exit For
                End If
            Next i
            
            
            Dim dIWBal As Double
            dIWBal = 0
            For i = 1 To nIW
                If arrIWCode(i) = CStr(arrM(m, 1)) Then
                    dIWBal = arrIWBal(i)
                    Exit For
                End If
            Next i
            

            ' *** 잔액(dBalD/dBalC): 음수 그대로 처리 ***
            If sBalDir = "D" Then
                dBalD = dBal
                dBalC = 0
            Else
                dBalC = dBal
                dBalD = 0
            End If

            ' *** 합계(dSumD/dSumC): 원래 로직 (IW 반영) ***
            If sBalDir = "D" Then
                dSumD = dSumD + dIWBal
            Else
                dSumC = dSumC + dIWBal
            End If

            ' 열려있는 레벨에 누산
            For lv = 1 To 3
                lvlSumD(lv) = lvlSumD(lv) + dSumD: lvlSumC(lv) = lvlSumC(lv) + dSumC
                lvlBalD(lv) = lvlBalD(lv) + dBalD: lvlBalC(lv) = lvlBalC(lv) + dBalC
            Next lv
            tSumD = tSumD + dSumD: tSumC = tSumC + dSumC
            tBalD = tBalD + dBalD: tBalC = tBalC + dBalC

NextMaster:
        Next m

        ' 유형 끝날 때 열려있던 헤더 전부 마감
        For lv = 3 To 1 Step -1
            If lvlHeaderRow(lv) > 0 Then
                Call 헤더마감(wsResult, lvlHeaderRow(lv), lvlBalD(lv), lvlSumD(lv), lvlSumC(lv), lvlBalC(lv))
            End If
        Next lv

        ' *** 유형 헤더 행에 소계 금액 채우기 ***
        If tBalD <> 0 Then wsResult.Cells(typeHeaderRow, 1).Value = tBalD
        If tSumD <> 0 Then wsResult.Cells(typeHeaderRow, 2).Value = tSumD
        If tSumC <> 0 Then wsResult.Cells(typeHeaderRow, 4).Value = tSumC
        If tBalC <> 0 Then wsResult.Cells(typeHeaderRow, 5).Value = tBalC
        wsResult.Range("A" & typeHeaderRow & ",B" & typeHeaderRow & ",D" & typeHeaderRow & ",E" & typeHeaderRow).NumberFormat = "#,##0"

        totalSumD = totalSumD + tSumD: totalSumC = totalSumC + tSumC
        totalBalD = totalBalD + tBalD: totalBalC = totalBalC + tBalC

NextType:
    Next t

    ' 최종 합계
    If totalBalD <> 0 Then wsResult.Cells(outRow, 1).Value = totalBalD
    If totalSumD <> 0 Then wsResult.Cells(outRow, 2).Value = totalSumD
    wsResult.Cells(outRow, 3).Value = "합   계"
    If totalSumC <> 0 Then wsResult.Cells(outRow, 4).Value = totalSumC
    If totalBalC <> 0 Then wsResult.Cells(outRow, 5).Value = totalBalC
    wsResult.Cells(outRow, 3).Font.Bold = True
    wsResult.Range("A" & outRow & ",B" & outRow & ",D" & outRow & ",E" & outRow).NumberFormat = "#,##0"

    ' ------------------------------------------
    ' 6. 서식
    ' ------------------------------------------
    Dim dataEnd As Long
    dataEnd = outRow

    wsResult.Range("A5:B" & dataEnd).NumberFormat = "#,##0"
    wsResult.Range("D5:E" & dataEnd).NumberFormat = "#,##0"
    
    With wsResult.Range("A3:E" & dataEnd)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    ' 계정과목 컬럼 구분선
    With wsResult.Range("C3:C" & dataEnd)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    With wsResult.Range("B4:B" & dataEnd)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(220, 220, 220)
    End With
    
    With wsResult.Range("E4:E" & dataEnd)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(220, 220, 220)
    End With
        
    
    ' 제목
    With wsResult.Range("A1:E1")
        .HorizontalAlignment = -4108
    End With
    wsResult.Rows(1).RowHeight = 40
    
    With wsResult.Range("A3:E4")
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
        .HorizontalAlignment = -4108
    End With
    
    
    ' 헤더 병합
    wsResult.Range("A3:B3").Merge
    wsResult.Range("D3:E3").Merge


    ' 열 너비 조정
    wsResult.Columns("A:E").AutoFit
    wsResult.Activate
    
   Call SetPrintOptions(wsResult, "$1:$4")
    

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Dim tElapsed As Double
    tElapsed = Timer - tStart

    MsgBox "합계잔액시산표 생성 완료!" & vbCrLf & _
       "계정 수 : " & nMaster & "개" & vbCrLf & _
       "소요 시간 : " & Format(tElapsed, "0.00") & " 초", _
       vbInformation

End Sub


'===========================================
' 중분류 헤더행 마감 헬퍼
'===========================================
Private Sub 헤더마감(ws As Worksheet, headerRow As Long, _
                     balD As Double, sumD As Double, sumC As Double, balC As Double)
    With ws
        If balD <> 0 Then .Cells(headerRow, 1).Value = balD
        If sumD <> 0 Then .Cells(headerRow, 2).Value = sumD
        If sumC <> 0 Then .Cells(headerRow, 4).Value = sumC
        If balC <> 0 Then .Cells(headerRow, 5).Value = balC
        .Range("A" & headerRow & ",B" & headerRow & _
               ",D" & headerRow & ",E" & headerRow).NumberFormat = "#,##0"
    End With
End Sub

