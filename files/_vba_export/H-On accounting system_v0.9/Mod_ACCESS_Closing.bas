Sub 결산마감_2025()
    
    Dim Conn As Object
    Dim Rs As Object
    Dim Sql As String
    Dim closingYear As Integer
    Dim nextYear As Integer
    Dim inputDateTime As String
    
    closingYear = 2025
    nextYear = 2026
    inputDateTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()
    
    On Error GoTo ErrorHandler
    
''    ' ================= Step 1: 손익계정 =================
''    MsgBox "Step 1: 손익계정 합계 계산...", vbInformation
''
''    Dim netIncome As Currency
''    Dim revenue As Currency
''    Dim expense As Currency
''
''    Sql = "SELECT SUM(대변 - 차변) as 수익금액 FROM 분개장 WHERE 연도=" & closingYear & " AND 결산이월잔액='N' AND 세목코드 LIKE '4%'"
''    Set Rs = Conn.Execute(Sql)
''    revenue = 0
''    If Not Rs.EOF And Not IsNull(Rs.Fields("수익금액").Value) Then
''        revenue = Rs.Fields("수익금액").Value
''    End If
''    Rs.Close
''
''    Sql = "SELECT SUM(차변 - 대변) as 비용금액 FROM 분개장 WHERE 연도=" & closingYear & " AND 결산이월잔액='N' AND 세목코드 LIKE '5%'"
''    Set Rs = Conn.Execute(Sql)
''    expense = 0
''    If Not Rs.EOF And Not IsNull(Rs.Fields("비용금액").Value) Then
''        expense = Rs.Fields("비용금액").Value
''    End If
''    Rs.Close
''
''    netIncome = revenue - expense
''
''    MsgBox "수익: " & Format(revenue, "#,##0") & vbCrLf & _
''           "비용: " & Format(expense, "#,##0") & vbCrLf & _
''           "당기순이익: " & Format(netIncome, "#,##0"), vbInformation

    Dim db As Currency, cr As Currency, 순액 As Currency

    ' Step 1 시작 전에 추가
    Debug.Print "==============================================="
    Debug.Print "[4% 계정 중 Y 항목 확인]"
    
    Sql = "SELECT 세목코드, 계정명, 결산이월잔액, SUM(차변) as 차변, SUM(대변) as 대변 FROM 분개장 " & _
          "WHERE 연도=" & closingYear & " AND 세목코드 LIKE '4%' AND 결산이월잔액='Y' " & _
          "GROUP BY 세목코드, 계정명, 결산이월잔액 ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    If Rs.RecordCount = 0 Then
        Debug.Print "4% Y 항목: 없음"
    Else
        Do While Not Rs.EOF
            ''Dim db As Currency, cr As Currency
            db = 0: cr = 0
            If Not IsNull(Rs.Fields("차변").Value) Then db = Rs.Fields("차변").Value
            If Not IsNull(Rs.Fields("대변").Value) Then cr = Rs.Fields("대변").Value
            Debug.Print Rs.Fields("세목코드") & " " & Rs.Fields("계정명") & " | 차변=" & Format(db, "#,##0") & " / 대변=" & Format(cr, "#,##0")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
    
    Debug.Print ""
    Debug.Print "[5% 계정 중 Y 항목 확인]"
    
    Sql = "SELECT 세목코드, 계정명, 결산이월잔액, SUM(차변) as 차변, SUM(대변) as 대변 FROM 분개장 " & _
          "WHERE 연도=" & closingYear & " AND 세목코드 LIKE '5%' AND 결산이월잔액='Y' " & _
          "GROUP BY 세목코드, 계정명, 결산이월잔액 ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    If Rs.RecordCount = 0 Then
        Debug.Print "5% Y 항목: 없음"
    Else
        Do While Not Rs.EOF
            db = 0: cr = 0
            If Not IsNull(Rs.Fields("차변").Value) Then db = Rs.Fields("차변").Value
            If Not IsNull(Rs.Fields("대변").Value) Then cr = Rs.Fields("대변").Value
            Debug.Print Rs.Fields("세목코드") & " " & Rs.Fields("계정명") & " | 차변=" & Format(db, "#,##0") & " / 대변=" & Format(cr, "#,##0")
            Rs.MoveNext
        Loop
    End If
    Rs.Close
    Debug.Print "==============================================="


    
    ' ================= Step 1: 손익계정 상세 디버그 =================
    MsgBox "Step 1: 손익계정 상세 분석...", vbInformation
    
    Debug.Print "==============================================="
    Debug.Print "[4% 계정 상세]"
    
    Sql = "SELECT 세목코드, 계정명, SUM(차변) as 차변, SUM(대변) as 대변 FROM 분개장 " & _
          "WHERE 연도=" & closingYear & " AND 세목코드 LIKE '4%' " & _
          "GROUP BY 세목코드, 계정명 ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    Dim revenue As Currency
    revenue = 0
    
    Do While Not Rs.EOF
        ''Dim db As Currency, cr As Currency, 순액 As Currency
        db = 0
        cr = 0
        If Not IsNull(Rs.Fields("차변").Value) Then db = Rs.Fields("차변").Value
        If Not IsNull(Rs.Fields("대변").Value) Then cr = Rs.Fields("대변").Value
        순액 = cr - db
        revenue = revenue + 순액
        
        Debug.Print Rs.Fields("세목코드") & " " & Rs.Fields("계정명") & " | 차변=" & Format(db, "#,##0") & " / 대변=" & Format(cr, "#,##0") & " → " & Format(순액, "#,##0")
        Rs.MoveNext
    Loop
    Rs.Close
    
    Debug.Print "4% 합계: " & Format(revenue, "#,##0")
    Debug.Print ""
    
    Debug.Print "==============================================="
    Debug.Print "[5% 계정 상세]"
    
    Sql = "SELECT 세목코드, 계정명, SUM(차변) as 차변, SUM(대변) as 대변 FROM 분개장 " & _
          "WHERE 연도=" & closingYear & " AND 세목코드 LIKE '5%' " & _
          "GROUP BY 세목코드, 계정명 ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    Dim expense As Currency
    expense = 0
    
    Do While Not Rs.EOF
        db = 0
        cr = 0
        If Not IsNull(Rs.Fields("차변").Value) Then db = Rs.Fields("차변").Value
        If Not IsNull(Rs.Fields("대변").Value) Then cr = Rs.Fields("대변").Value
        순액 = db - cr
        expense = expense + 순액
        
        Debug.Print Rs.Fields("세목코드") & " " & Rs.Fields("계정명") & " | 차변=" & Format(db, "#,##0") & " / 대변=" & Format(cr, "#,##0") & " → " & Format(순액, "#,##0")
        Rs.MoveNext
    Loop
    Rs.Close
    
    Debug.Print "5% 합계: " & Format(expense, "#,##0")
    Debug.Print ""
    
    Dim netIncome As Currency
    netIncome = revenue - expense
    
    Debug.Print "==============================================="
    Debug.Print "당기순이익: " & Format(netIncome, "#,##0")
    Debug.Print "==============================================="
    
    MsgBox "수익: " & Format(revenue, "#,##0") & vbCrLf & _
           "비용: " & Format(expense, "#,##0") & vbCrLf & _
           "당기순이익: " & Format(netIncome, "#,##0"), vbInformation
    

    
''    ' ================= Step 2: 이월이익 =================
''    MsgBox "Step 2: 이월이익잉여금 계산...", vbInformation
''
''    Dim 기존이월이익 As Currency
''    Dim 신규이월이익 As Currency
''
''    sql = "SELECT SUM(대변) - SUM(차변) as 이월이익 FROM 분개장 WHERE 연도=" & closingYear & " AND 세목코드='3311200'"
''    Set rs = conn.Execute(sql)
''    기존이월이익 = 0
''    If Not rs.EOF And Not IsNull(rs.Fields("이월이익").Value) Then
''        기존이월이익 = rs.Fields("이월이익").Value
''    End If
''    rs.Close
''
''    신규이월이익 = 기존이월이익 + netIncome
''
''    MsgBox "기존: " & Format(기존이월이익, "#,##0") & vbCrLf & _
''           "신규: " & Format(신규이월이익, "#,##0"), vbInformation
''
           
           
''    '================= Step 2: 이월이익 =================
''    MsgBox "Step 2: 이월이익잉여금 계산...", vbInformation
''
''    Dim 기존이월이익 As Currency
''    Dim 신규이월이익 As Currency
''
''    Sql = "SELECT SUM(대변) - SUM(차변) as 이월이익 FROM 분개장 WHERE 연도=" & closingYear & " AND 세목코드='3311200'"
''    Set Rs = Conn.Execute(Sql)
''    기존이월이익 = 0
''    If Not Rs.EOF And Not IsNull(Rs.Fields("이월이익").Value) Then
''       기존이월이익 = Rs.Fields("이월이익").Value
''    End If
''    Rs.Close
''
''    Debug.Print "=== 이월이익잉여금 계산 ==="
''    Debug.Print "기존 이월이익: " & Format(기존이월이익, "#,##0")
''    Debug.Print "당기순이익: " & Format(netIncome, "#,##0")
''
''    신규이월이익 = 기존이월이익 + netIncome
''
''    Debug.Print "신규 이월이익: " & Format(신규이월이익, "#,##0")
''    Debug.Print ""
''
''    MsgBox "기존: " & Format(기존이월이익, "#,##0") & vbCrLf & _
''          "신규: " & Format(신규이월이익, "#,##0"), vbInformation

           

    ' ================= Step 2: 이월이익 디버그 =================
    MsgBox "Step 2: 이월이익잉여금 계산...", vbInformation
    
    Dim 기존이월이익 As Currency
    Dim 신규이월이익 As Currency
    Dim 차변합계 As Currency
    Dim 대변합계 As Currency
    
    ' 1단계: 기존 이월이익 조회 (상세)
    Sql = "SELECT " & _
          "SUM(차변) as 총차변, " & _
          "SUM(대변) as 총대변 " & _
          "FROM 분개장 WHERE 연도=" & closingYear & " AND 세목코드='3311200'"
    
    Set Rs = Conn.Execute(Sql)
    
    차변합계 = 0
    대변합계 = 0
    
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields("총차변").Value) Then 차변합계 = Rs.Fields("총차변").Value
        If Not IsNull(Rs.Fields("총대변").Value) Then 대변합계 = Rs.Fields("총대변").Value
    End If
    Rs.Close
    
    기존이월이익 = 대변합계 - 차변합계
    
    ' 2단계: 신규 이월이익 계산
    신규이월이익 = 기존이월이익 + netIncome
    
    ' 3단계: 상세 로그
    Debug.Print "==============================================="
    Debug.Print "[이월이익잉여금 계산 상세]"
    Debug.Print "연도: " & closingYear
    Debug.Print ""
    Debug.Print "당기순이익 계산:"
    Debug.Print "  ├─ 수익(4%): " & Format(revenue, "#,##0")
    Debug.Print "  ├─ 비용(5%): " & Format(expense, "#,##0")
    Debug.Print "  └─ 당기순이익: " & Format(netIncome, "#,##0")
    Debug.Print ""
    Debug.Print "이월이익잉여금(3311200):"
    Debug.Print "  ├─ 차변합계: " & Format(차변합계, "#,##0")
    Debug.Print "  ├─ 대변합계: " & Format(대변합계, "#,##0")
    Debug.Print "  ├─ 기존이월이익(대변-차변): " & Format(기존이월이익, "#,##0")
    Debug.Print "  └─ 신규이월이익(기존+당기): " & Format(신규이월이익, "#,##0")
    Debug.Print "==============================================="
    
    MsgBox "기존이월이익: " & Format(기존이월이익, "#,##0") & vbCrLf & _
           "당기순이익: " & Format(netIncome, "#,##0") & vbCrLf & _
           "신규이월이익: " & Format(신규이월이익, "#,##0"), vbInformation
            
    
    ' ================= Step 3: 개시잔액 분개 생성 =================
    MsgBox "Step 3: 개시잔액 분개 생성...", vbInformation
    
    Dim sequenceNo As Integer
    Dim balanceAmount As Currency
    
    sequenceNo = 0
    
    ' B/S 계정 조회 (계정마스터 조인으로 잔액방향 가져오기)
    Sql = "SELECT j.세목코드, j.계정명, SUM(j.차변) as 총차변, SUM(j.대변) as 총대변, m.잔액방향 " & _
          "FROM 분개장 j " & _
          "LEFT JOIN 계정마스터 m ON j.세목코드 = m.세목코드 " & _
          "WHERE j.연도=" & closingYear & " AND (j.세목코드 LIKE '1%' OR j.세목코드 LIKE '2%' OR j.세목코드 LIKE '3%') " & _
          "GROUP BY j.세목코드, j.계정명, m.잔액방향 " & _
          "ORDER BY j.세목코드"
    
    Set Rs = Conn.Execute(Sql)
    
    Do While Not Rs.EOF
        
        Dim accountCode As String
        Dim accountName As String
        Dim balanceDirection As String
        Dim debit As Currency
        Dim credit As Currency
        Dim netBalance As Currency
        Dim totalDebit As Currency
        Dim totalCredit As Currency
        
        accountCode = CStr(Rs.Fields("세목코드").Value)
        accountName = CStr(Rs.Fields("계정명").Value)
        balanceDirection = ""
        
        If Not IsNull(Rs.Fields("잔액방향").Value) Then
            balanceDirection = CStr(Rs.Fields("잔액방향").Value)
        End If
        
        Debug.Print "계정: " & accountCode & " / 계정명: " & accountName & " / 잔액방향: [" & balanceDirection & "]"
        
        ' 총차변/총대변
        totalDebit = 0
        totalCredit = 0
        
        If Not IsNull(Rs.Fields("총차변").Value) Then
            totalDebit = Rs.Fields("총차변").Value
        End If
        If Not IsNull(Rs.Fields("총대변").Value) Then
            totalCredit = Rs.Fields("총대변").Value
        End If
        
        ' 잔액방향 기반 정규화
        debit = 0
        credit = 0
        
        If balanceDirection = "D" Then
            ' 자산/비용 (차변 정상): 차변 - 대변
            balanceAmount = totalDebit - totalCredit
            Debug.Print "  → D (차변 정상) / 순잔액: " & balanceAmount
            debit = balanceAmount
            credit = 0
            
        ElseIf balanceDirection = "C" Then
            ' 부채/자본 (대변 정상): 대변 - 차변
            balanceAmount = totalCredit - totalDebit
            Debug.Print "  → C (대변 정상) / 순잔액: " & balanceAmount
            debit = 0
            credit = balanceAmount  '＇ 음수도 그대로!
            
        Else
            Debug.Print "  → 기타/NULL: [" & balanceDirection & "] / 순잔액: " & (totalDebit - totalCredit)
            ' NULL이면 기본값 (자산처럼 처리)
            balanceAmount = totalDebit - totalCredit
            debit = balanceAmount
            credit = 0
        End If
        
    
''        ' 이월이익잉여금 특별 처리
''        If accountCode = "3311200" Then
''            debit = 0
''            credit = 신규이월이익
''            Debug.Print "  → 이월이익잉여금 적용"
''        End If

        ' 이월이익잉여금 특별 처리
        If accountCode = "3311200" Then
            debit = 0
            credit = 신규이월이익
            Debug.Print "★★★ 이월이익잉여금 항목 ★★★"
            Debug.Print "  INSERT: 차변=" & debit & " / 대변=" & credit
            Debug.Print ""
        End If

        
        sequenceNo = sequenceNo + 1
        
        accountName = Replace(accountName, "'", "''")
        
        ' ? 전표번호 없음 (빈 값), 명세번호만 순번
        Sql = "INSERT INTO 분개장 " & _
              "(회계일자, 전표번호, 명세번호, 적요, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 입력일시, 입력자, 연도, 결산이월잔액) " & _
              "VALUES " & _
              "(#" & nextYear & "-01-01#, '', " & sequenceNo & ", '전기이월', " & _
              debit & ", " & credit & ", " & _
              "'" & accountCode & "', '" & accountName & "', '회계팀', 'KRW', " & _
              "'" & inputDateTime & "', '" & Application.UserName & "', " & nextYear & ", 'Y')"
        
        Conn.Execute Sql
        
        Debug.Print "  ? INSERT: 차변=" & debit & " / 대변=" & credit & vbCrLf
        
        Rs.MoveNext
    Loop

    
    Rs.Close
    
    MsgBox "마감 완료!" & vbCrLf & _
           "생성된 이월 항목: " & sequenceNo & "개", vbInformation

CleanUp:
    Conn.Close
    Set Conn = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "에러: " & Err.Description, vbCritical
    Conn.Close
    Set Conn = Nothing
    
End Sub




Sub 마감검증_2025_2026(Conn As Object, closingYear As Integer, nextYear As Integer)
    
    Dim Rs As Object
    Dim Sql As String
    Dim result As String
    
    result = "=== 마감 검증 결과 ===" & vbCrLf & vbCrLf
    
    ' 25년 말 B/S 계정 잔액 (개시잔액 포함)
    result = result & "[25년도 말 B/S 잔액]" & vbCrLf
    
    Sql = "SELECT 세목코드, 계정명, " & _
          "SUM(차변) as 차변, SUM(대변) as 대변, " & _
          "SUM(차변-대변) as 잔액 " & _
          "FROM 분개장 " & _
          "WHERE 연도=" & closingYear & " AND (세목코드 LIKE '1%' OR 세목코드 LIKE '2%' OR 세목코드 LIKE '3%') " & _
          "GROUP BY 세목코드, 계정명 " & _
          "ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    
    Dim balance2025 As Currency
    balance2025 = 0
    
    Do While Not Rs.EOF
        result = result & Rs.Fields("세목코드").Value & " " & _
                 Rs.Fields("계정명").Value & ": " & _
                 Format(Rs.Fields("잔액").Value, "#,##0") & vbCrLf
        
        balance2025 = balance2025 + Rs.Fields("잔액").Value
        
        Rs.MoveNext
    Loop
    
    result = result & "총 잔액: " & Format(balance2025, "#,##0") & vbCrLf & vbCrLf
    Rs.Close
    
    
    ' 26년 1월1일 개시잔액 (방금 생성된 것)
    result = result & "[26년 1월1일 개시잔액]" & vbCrLf
    
    Sql = "SELECT 세목코드, 계정명, " & _
          "SUM(차변) as 차변, SUM(대변) as 대변, " & _
          "SUM(차변-대변) as 잔액 " & _
          "FROM 분개장 " & _
          "WHERE 연도=" & nextYear & " AND 결산이월잔액='Y' AND (세목코드 LIKE '1%' OR 세목코드 LIKE '2%' OR 세목코드 LIKE '3%') " & _
          "GROUP BY 세목코드, 계정명 " & _
          "ORDER BY 세목코드"
    
    Set Rs = Conn.Execute(Sql)
    
    Dim balance2026 As Currency
    balance2026 = 0
    
    Do While Not Rs.EOF
        result = result & Rs.Fields("세목코드").Value & " " & _
                 Rs.Fields("계정명").Value & ": " & _
                 Format(Rs.Fields("잔액").Value, "#,##0") & vbCrLf
        
        balance2026 = balance2026 + Rs.Fields("잔액").Value
        
        Rs.MoveNext
    Loop
    
    result = result & "총 잔액: " & Format(balance2026, "#,##0") & vbCrLf & vbCrLf
    Rs.Close
    
    
    ' 비교
    result = result & "=== 비교 ===" & vbCrLf
    result = result & "25년 말 잔액: " & Format(balance2025, "#,##0") & vbCrLf
    result = result & "26년 초 잔액: " & Format(balance2026, "#,##0") & vbCrLf
    
    If balance2025 = balance2026 Then
        result = result & " 일치! 마감이 정확합니다."
    Else
        result = result & " 불일치! 차이: " & Format(balance2025 - balance2026, "#,##0")
    End If
    
    MsgBox result, vbInformation, "마감 검증 결과"

End Sub


