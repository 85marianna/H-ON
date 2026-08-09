Option Explicit


'===========================================
' 당기순이익 계산 (BS/PL 공용, 화면 출력 없이 값만 반환)
'===========================================

Public Function 당기순이익_계산(Optional dateFrom As Variant = "", Optional dateTo As Variant = "") As Double

    Dim arrData() As Variant
    arrData = 원장데이터가져오기_DB("", dateFrom, dateTo)

    If Not IsArray(arrData) Then
        당기순이익_계산 = 0
        Exit Function
    End If

    Dim Conn As Object, rsMaster As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set rsMaster = CreateObject("ADODB.Recordset")
    rsMaster.Open "SELECT 세목코드, 잔액방향 FROM 계정마스터 " & _
                  "WHERE 사용여부='Y' AND 재무제표='PL'", Conn, 1, 1

    Dim nMaster As Long
    nMaster = 0
    Do While Not rsMaster.EOF
        nMaster = nMaster + 1
        rsMaster.MoveNext
    Loop
    If nMaster = 0 Then
        rsMaster.Close: Conn.Close
        당기순이익_계산 = 0
        Exit Function
    End If
    rsMaster.MoveFirst

    Dim arrM() As Variant
    ReDim arrM(1 To nMaster, 1 To 2)
    Dim m As Long
    m = 0
    Do While Not rsMaster.EOF
        m = m + 1
        arrM(m, 1) = CStr(rsMaster.Fields("세목코드").Value)
        arrM(m, 2) = CStr(rsMaster.Fields("잔액방향").Value)
        rsMaster.MoveNext
    Loop
    rsMaster.Close
    Conn.Close
    Set rsMaster = Nothing: Set Conn = Nothing

    Dim nS As Long, i As Long
    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then nS = nS + 1
    Next i

    Dim arrSCode() As String
    Dim arrSBal() As Double
    ReDim arrSCode(1 To nS)
    ReDim arrSBal(1 To nS)

    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then
            nS = nS + 1
            arrSCode(nS) = CStr(arrData(i, 1))
            arrSBal(nS) = CDbl(IIf(IsNumeric(arrData(i, 10)), arrData(i, 10), 0))
        End If
    Next i

    Dim 순이익 As Double
    순이익 = 0

    For m = 1 To nMaster
        Dim dBal As Double, dBalD As Double, dBalC As Double
        Dim sBalDir As String
        dBal = 0: dBalD = 0: dBalC = 0
        sBalDir = arrM(m, 2)

        For i = 1 To nS
            If arrSCode(i) = arrM(m, 1) Then
                dBal = arrSBal(i)
                Exit For
            End If
        Next i

        If sBalDir = "D" Then
            dBalD = IIf(dBal >= 0, dBal, 0)
            dBalC = IIf(dBal < 0, Abs(dBal), 0)
        Else
            dBalC = IIf(dBal >= 0, dBal, 0)
            dBalD = IIf(dBal < 0, Abs(dBal), 0)
        End If

        순이익 = 순이익 + (dBalC - dBalD)
    Next m

    당기순이익_계산 = 순이익

End Function




'===========================================
' 손익계산서 생성 (신규 워크북 출력)
'===========================================

Public Sub 손익계산서_생성()
    ' 날짜 설정
    Dim dtEnd As Date
    Dim dtStart As Date
    dtEnd = frm보고서조회.txtDate.Value
    dtStart = DateSerial(Year(dtEnd), 1, 1)  ' 1월 1일
    
    Dim arrData() As Variant
    arrData = 원장데이터가져오기_DB("", dtStart, dtEnd)

    If Not IsArray(arrData) Then
        MsgBox "원장 데이터를 가져올 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' ------------------------------------------
    ' 1. 계정마스터 로드
    ' ------------------------------------------
    Dim Conn As Object, rsMaster As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set rsMaster = CreateObject("ADODB.Recordset")
    rsMaster.Open "SELECT 세목코드, 계정명, 계정유형, 계정구분, 계정구분명, 중분류2, 잔액방향 " & _
                  "FROM 계정마스터 WHERE 사용여부='Y' AND 재무제표='PL' " & _
                  "ORDER BY 계정유형, 세목코드", Conn, 1, 1

    Dim nMaster As Long
    nMaster = 0
    Do While Not rsMaster.EOF
        nMaster = nMaster + 1
        rsMaster.MoveNext
    Loop

    If nMaster = 0 Then
        MsgBox "재무제표=PL 인 계정이 없습니다.", vbExclamation
        rsMaster.Close: Conn.Close
        Exit Sub
    End If
    rsMaster.MoveFirst

    Dim arrM() As Variant
    ReDim arrM(1 To nMaster, 1 To 7)
    Dim m As Long
    m = 0
    Do While Not rsMaster.EOF
        m = m + 1
        arrM(m, 1) = CStr(rsMaster.Fields("세목코드").Value)
        arrM(m, 2) = rsMaster.Fields("계정명").Value
        arrM(m, 3) = rsMaster.Fields("계정유형").Value
        arrM(m, 4) = CStr(rsMaster.Fields("계정구분").Value)
        arrM(m, 5) = rsMaster.Fields("계정구분명").Value
        If IsNull(rsMaster.Fields("중분류2").Value) Then
            arrM(m, 6) = ""
        Else
            arrM(m, 6) = CStr(rsMaster.Fields("중분류2").Value)
        End If
        arrM(m, 7) = CStr(rsMaster.Fields("잔액방향").Value)
        rsMaster.MoveNext
    Loop
    rsMaster.Close
    Conn.Close
    Set rsMaster = Nothing: Set Conn = Nothing

    ' ------------------------------------------
    ' 2. S행 / IW행 수집
    ' ------------------------------------------
    Dim nS As Long, i As Long
    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then nS = nS + 1
    Next i

    Dim arrSCode() As String
    Dim arrSBal() As Double
    ReDim arrSCode(1 To nS)
    ReDim arrSBal(1 To nS)

    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then
            nS = nS + 1
            arrSCode(nS) = CStr(arrData(i, 1))
            arrSBal(nS) = CDbl(IIf(IsNumeric(arrData(i, 10)), arrData(i, 10), 0))
        End If
    Next i

    ' ------------------------------------------
    ' 3. 신규 워크북 생성
    ' ------------------------------------------
    Dim wbNew As Workbook
    Dim wsResult As Worksheet
    Set wbNew = Workbooks.Add
    Set wsResult = wbNew.Worksheets(1)
    On Error Resume Next
    wsResult.Name = "손익계산서"
    On Error GoTo 0

    wsResult.Cells(1, 1).Value = "【 손   익   계   산   서 】"
    wsResult.Cells(1, 1).Font.Bold = True
    wsResult.Cells(1, 1).Font.Size = 15
    
    With wsResult.Range("A1:C1")
         .HorizontalAlignment = xlCenterAcrossSelection
    End With
    
    Dim maxDate As Date
    maxDate = DateSerial(1900, 1, 1)
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "D" Then
            If IsDate(arrData(i, 3)) Then
                If CDate(arrData(i, 3)) > maxDate Then maxDate = CDate(arrData(i, 3))
            End If
        End If
    Next i
    
    wsResult.Cells(3, 1).Value = "韓國電氣硝子(株)"
    wsResult.Cells(2, 2).Value = "第 53 期 " & Year(maxDate) & "年 01月 01日 부터 " & Format(maxDate, "YYYY年 MM月 DD日") & " 까지"
    With wsResult.Range("B2:C2")
        .HorizontalAlignment = xlCenterAcrossSelection
        .ShrinkToFit = True
    End With
        
    wsResult.Cells(3, 5).Value = "(單位:원)"
    wsResult.Cells(3, 5).HorizontalAlignment = -4152

    wsResult.Cells(4, 1).Value = "科  目"
    With wsResult.Range("A4:A5")
        .Merge
        .HorizontalAlignment = -4108
    End With

    wsResult.Cells(4, 2).Value = "第 53 (當) 期 "
    With wsResult.Range("B4:C4")
        .Merge
        .HorizontalAlignment = -4108
    End With

    wsResult.Cells(5, 2).Value = "金  額"
    With wsResult.Range("B5:C5")
        .Merge
        .HorizontalAlignment = -4108
    End With
    
    
        wsResult.Cells(4, 4).Value = "第 52 (當) 期 "
    With wsResult.Range("D4:E4")
        .Merge
        .HorizontalAlignment = -4108
    End With

    wsResult.Cells(5, 4).Value = "金  額"
    With wsResult.Range("D5:E5")
        .Merge
        .HorizontalAlignment = -4108
    End With


    wsResult.Range("A4:E5").Font.Bold = True
    wsResult.Range("A4:E5").Interior.Color = RGB(217, 217, 217)

    Dim outRow As Long
    outRow = 6

    ' ------------------------------------------
    ' 매출/매출원가 템플릿 섹션
    ' ------------------------------------------
    wsResult.Cells(outRow, 1).Value = "Ⅰ. 純賣出額"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   1. 賣出額"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "      輸出"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "      國內賣出"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "      商品賣出"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "      其他賣出"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   2. 賣出에누리"
    outRow = outRow + 2

    wsResult.Cells(outRow, 1).Value = "Ⅱ. 賣出原價"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   1. 期初製品在庫額"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   2. 當期製品製造原價"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   3. 當期商品賣出原價"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   4. 其他賣出原價"
    outRow = outRow + 2
    wsResult.Cells(outRow, 1).Value = "           計"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   5. 關稅還給金"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   6. 他計定對替額"
    outRow = outRow + 1
    wsResult.Cells(outRow, 1).Value = "   7. 期末製品在庫額"
    outRow = outRow + 2

    wsResult.Cells(outRow, 1).Value = "Ⅲ. 賣出總利益"
    outRow = outRow + 1

    ' ------------------------------------------
    ' 판관비/영업외수익/영업외비용 섹션
    ' ------------------------------------------
    ' 타입별 합계 저장 변수 선언
    Dim 판관비합계 As Double
    Dim 영업외수익합계 As Double
    Dim 영업외비용합계 As Double
    Dim 법인세비용합계 As Double
    
    판관비합계 = 0
    영업외수익합계 = 0
    영업외비용합계 = 0
    법인세비용합계 = 0

    ' 로마자 숫자 매핑
    Dim 로마자(1 To 4) As String
    로마자(1) = "Ⅳ"
    로마자(2) = "Ⅵ"
    로마자(3) = "Ⅶ"
    로마자(4) = "Ⅹ"

    Dim typeList(1 To 4) As String
    typeList(1) = "판매비와관리비"
    typeList(2) = "영업외수익"
    typeList(3) = "영업외비용"
    typeList(4) = "법인세비용"

    Dim lvlCode(1 To 2) As String
    Dim lvlName(1 To 2) As String
    Dim lvlHeaderRow(1 To 2) As Long
    Dim lvlSum(1 To 2) As Double

    Dim lvlColor(1 To 2) As Long
    lvlColor(1) = RGB(198, 217, 241)
    lvlColor(2) = RGB(235, 241, 222)

    Dim t As Integer
    For t = 1 To 4
        Dim tLabel As String
        tLabel = typeList(t)

        ' hasType 체크 - 중분류2 기준
        Dim hasType As Boolean
        hasType = False
        For m = 1 To nMaster
            If CStr(arrM(m, 6)) = tLabel Then hasType = True: Exit For
        Next m
        If Not hasType Then GoTo NextType

        Dim tSum As Double
        tSum = 0
        Dim lv As Integer
        For lv = 1 To 2
            lvlCode(lv) = "|__초기값__|"
            lvlHeaderRow(lv) = 0
            lvlSum(lv) = 0
        Next lv

        ' 계정구분 번호 초기화 (각 중분류2마다)
        Dim 계정구분번호 As Integer
        계정구분번호 = 0

        For m = 1 To nMaster
            ' 중분류2 기준으로 필터링
            If CStr(arrM(m, 6)) <> tLabel Then GoTo NextMaster

            Dim newCode(1 To 2) As String
            Dim newName(1 To 2) As String
            newCode(1) = CStr(arrM(m, 6)): newName(1) = CStr(arrM(m, 6))
            newCode(2) = CStr(arrM(m, 4)): newName(2) = CStr(arrM(m, 5))

            Dim changeLv As Integer
            changeLv = 3
            For lv = 1 To 2
                If newCode(lv) <> lvlCode(lv) Then
                    changeLv = lv
                    Exit For
                End If
            Next lv

            If changeLv <= 2 Then
                For lv = 2 To changeLv Step -1
                    If lvlHeaderRow(lv) > 0 Then
                        Call PL_헤더마감(wsResult, lvlHeaderRow(lv), lvlSum(lv), lvlColor(lv), lv)
                        lvlSum(lv) = 0
                        lvlHeaderRow(lv) = 0
                    End If
                Next lv

                For lv = changeLv To 2
                    lvlCode(lv) = newCode(lv)
                    lvlName(lv) = newName(lv)
                    lvlHeaderRow(lv) = outRow
                    
                    If lv = 1 Then
                        ' 중분류2 (로마자 붙이기)
                        wsResult.Cells(outRow, 1).Value = Space((lv - 1) * 3) & 로마자(t) & ". " & newName(lv)
                        계정구분번호 = 0  ' 새로운 중분류2이므로 번호 초기화
                    Else
                        ' 계정구분 (번호 붙이기)
                        계정구분번호 = 계정구분번호 + 1
                        wsResult.Cells(outRow, 1).Value = Space((lv - 1) * 3) & 계정구분번호 & ". " & newName(lv)
                    End If
                    
                    outRow = outRow + 1
                Next lv
            End If

            ' 개별 세목 금액 계산
            Dim dBal As Double, dBalD As Double, dBalC As Double
            Dim sBalDir As String
            dBal = 0: dBalD = 0: dBalC = 0
            sBalDir = CStr(arrM(m, 7))

            For i = 1 To nS
                If arrSCode(i) = CStr(arrM(m, 1)) Then
                    dBal = arrSBal(i)
                    Exit For
                End If
            Next i

            If sBalDir = "D" Then
                dBalD = IIf(dBal >= 0, dBal, 0)
                dBalC = IIf(dBal < 0, Abs(dBal), 0)
            Else
                dBalC = IIf(dBal >= 0, dBal, 0)
                dBalD = IIf(dBal < 0, Abs(dBal), 0)
            End If

            ' 타입별 계산
            Dim dAmt As Double
            If tLabel = "영업외수익" Then
                dAmt = dBalC - dBalD
            Else
                dAmt = dBalD - dBalC
            End If

            lvlSum(1) = lvlSum(1) + dAmt
            lvlSum(2) = lvlSum(2) + dAmt
            tSum = tSum + dAmt

NextMaster:
        Next m

        For lv = 2 To 1 Step -1
            If lvlHeaderRow(lv) > 0 Then
                Call PL_헤더마감(wsResult, lvlHeaderRow(lv), lvlSum(lv), lvlColor(lv), lv)
            End If
        Next lv
        
        ' tSum을 타입별 변수에 저장
        If t = 1 Then
            판관비합계 = tSum
        ElseIf t = 2 Then
            영업외수익합계 = tSum
        ElseIf t = 3 Then
            영업외비용합계 = tSum
        ElseIf t = 4 Then
            법인세비용합계 = tSum
        End If
        
        ' 강제 행 삽입 (계산된 값 포함)
        If t = 1 Then  ' 판매비 끝
            Dim 매출총이익 As Double
            매출총이익 = 0
            Dim 영업이익 As Double
            영업이익 = 매출총이익 - 판관비합계
            outRow = outRow + 1
            Call PL_소계행출력(wsResult, outRow, "Ⅴ. 營業利益", 영업이익, RGB(220, 230, 241))
            outRow = outRow + 1
        ElseIf t = 3 Then  ' 영업외비용 끝
            Dim 경상이익 As Double
            경상이익 = 영업이익 + (영업외수익합계 - 영업외비용합계)
            outRow = outRow + 1
            Call PL_소계행출력(wsResult, outRow, "Ⅷ. 經常利益", 경상이익, RGB(220, 230, 241))
            outRow = outRow + 2
            Call PL_소계행출력(wsResult, outRow, "Ⅸ. 法人稅費用差減前順利益", 경상이익, RGB(220, 230, 241))
            outRow = outRow + 2
        End If

NextType:
    Next t

    
    ' 당기순이익 (경상이익 - 법인세비용)
    Dim 당기순이익 As Double
    Dim nIPResult As Double
    nIPResult = 당기순이익_계산(dtStart, dtEnd)
    당기순이익 = nIPResult
    Call PL_소계행출력(wsResult, outRow, "ⅩⅠ. 當期純利益", 당기순이익, RGB(180, 198, 231))
    

    ' ------------------------------------------
    ' 4. 서식
    ' ------------------------------------------
    wsResult.Range("B6:C" & outRow).NumberFormat = "#,##0;(#,##0);"""""
    wsResult.Columns("A").ColumnWidth = 35
    wsResult.Columns("B").ColumnWidth = 17
    wsResult.Columns("C").ColumnWidth = 17
    wsResult.Columns("D").ColumnWidth = 17
    wsResult.Columns("E").ColumnWidth = 17

    With wsResult.Range("A4:E" & outRow)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    Call SetPrintOptions(wsResult, "$1:$5")

    MsgBox "손익계산서 생성 완료" & vbCrLf & vbCrLf & _
           "당기순이익 : " & Format(당기순이익, "#,##0"), vbInformation

End Sub


'===========================================
' PL용 소계행 출력 헬퍼
'===========================================
Private Sub PL_소계행출력(ws As Worksheet, r As Long, sLabel As String, sumVal As Double, bgColor As Long)
    With ws
        .Cells(r, 1).Value = sLabel
        .Cells(r, 3).Value = sumVal
        .Range("A" & r & ":E" & r).Interior.Color = bgColor
        .Cells(r, 3).NumberFormat = "#,##0"
    End With
End Sub


'===========================================
' PL용 헤더 마감 헬퍼 (B/C열 분리)
'===========================================
Private Sub PL_헤더마감(ws As Worksheet, headerRow As Long, sumVal As Double, bgColor As Long, lv As Integer)
    With ws
        Dim col As Integer
        If lv = 1 Then
            col = 3  ' 중분류2 합계 → C열
        Else
            col = 2  ' 계정구분명 합계 → B열
        End If
        
        .Cells(headerRow, col).Value = sumVal
        .Cells(headerRow, col).NumberFormat = "#,##0"
    End With
End Sub

