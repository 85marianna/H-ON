Option Explicit

'''===========================================
''' 재무상태표 생성 (신규 워크북 출력, 3단계 소계)
''' 모듈: HJ_Mod_Report
'''
''' [주요 기능]
''' - 중분류2: 로마자 순번 (Ⅰ. Ⅱ. Ⅲ.)
''' - 중분류3: 괄호 순번 ((1) (2) (3)...)
''' - 계정구분명: 숫자 순번 (1. 2. 3...)
''' - 각 대분류(자산/부채/자본) 기준으로 순번 리셋
''' - 금액을 C열에 표시 (B열은 참고값 전용)
''' - 이월이익잉여금(3311200) 다음에 당기순이익 행 자동 삽입
''' - 감가상각대상 유형자산 순자산가액 자동 계산 & 누적
''' - 감가상각누계액 행 별도 표시
'''===========================================
Public Sub 재무상태표_생성()
 
''    ' 날짜 설정
''    Dim dtEnd As Date
''    Dim dtStart As Date
''    dtEnd = frm보고서조회.txtDate.Value
''    dtStart = DateSerial(Year(dtEnd), 1, 1)
''
''    Dim arrData() As Variant
''    arrData = 원장데이터가져오기_DB("", dtStart, dtEnd)
''    If Not IsArray(arrData) Then
''        MsgBox "원장 데이터를 가져올 수 없습니다.", vbExclamation
''        Exit Sub
''    End If
''
''    ' 당기순이익 먼저 계산 (이월이익잉여금에 반영할 값)
''    Dim 당기순이익 As Double
''    당기순이익 = 당기순이익_계산(dtStart, dtEnd)

    
    ' 날짜 설정
    Dim dtEnd As Date
    Dim dtStart As Date
    dtEnd = frm보고서조회.txtDate.Value              ' 선택 날짜 (예: 2025.06.30)
    dtStart = DateSerial(Year(dtEnd), 1, 1)          ' 그 해의 1월1일 (예: 2025.01.01)
    
    ' 당기 데이터
    Dim arrData() As Variant
    arrData = 원장데이터가져오기_DB("", dtStart, dtEnd)
        
    If Not IsArray(arrData) Then
        MsgBox "당기 원장 데이터를 가져올 수 없습니다.", vbExclamation
        Exit Sub
    End If
    

''    ' 당기순이익
''    Dim 당기순이익 As Double
''    당기순이익 = 당기순이익_계산(Year(dtEnd))


    ' 당기순이익 (dtStart, dtEnd 날짜를 그대로 넘기기)
    Dim 당기순이익 As Double
    당기순이익 = 당기순이익_계산(dtStart, dtEnd)
    
    


    ' ------------------------------------------
    ' 1. 계정마스터 로드 (BS만, 중분류2/3 포함)
    ' ★ NOT IN으로 감가상각누계액 세부 계정 제외
    '    (부채 섹션에서 중복 제거용)
    '    실제 금액은 원장데이터(arrSCode)에서 별도 조회
    ' ------------------------------------------
    Dim Conn As Object, rsMaster As Object
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set rsMaster = CreateObject("ADODB.Recordset")
    rsMaster.Open "SELECT 세목코드, 계정명, 계정유형, 계정구분, 계정구분명, " & _
                  "중분류2, 중분류3, 잔액방향 " & _
                  "FROM 계정마스터 WHERE 사용여부='Y' AND 재무제표='BS' " & _
                  "AND 세목코드 NOT IN ('2210701', '2210702', '2210703', '2210704', '2210706') " & _
                  "ORDER BY 계정유형, 세목코드", Conn, 1, 1

    Dim nMaster As Long
    nMaster = 0
    Do While Not rsMaster.EOF
        nMaster = nMaster + 1
        rsMaster.MoveNext
    Loop

    If nMaster = 0 Then
        MsgBox "재무제표=BS 인 계정이 없습니다.", vbExclamation
        rsMaster.Close: Conn.Close
        Exit Sub
    End If
    rsMaster.MoveFirst

    ' 마스터 배열: 1세목코드,2계정명,3계정유형,4계정구분,5계정구분명,
    '              6중분류2,7중분류3,8잔액방향
    Dim arrM() As Variant
    ReDim arrM(1 To nMaster, 1 To 8)
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

        If IsNull(rsMaster.Fields("중분류3").Value) Then
            arrM(m, 7) = ""
        Else
            arrM(m, 7) = CStr(rsMaster.Fields("중분류3").Value)
        End If

        arrM(m, 8) = CStr(rsMaster.Fields("잔액방향").Value)
        rsMaster.MoveNext
    Loop
    rsMaster.Close
    Conn.Close
    Set rsMaster = Nothing: Set Conn = Nothing

    ' ------------------------------------------
    ' 2. S행 수집 (기존과 동일)
    ' ------------------------------------------
    Dim nS As Long, i As Long
    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then nS = nS + 1
    Next i

    Dim arrSCode() As String
    Dim arrSBal()  As Double
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
    wsResult.Name = "재무상태표"
    On Error GoTo 0

    wsResult.Cells(1, 1).Value = "【 재   무   상   태   표 】"
    wsResult.Cells(1, 1).Font.Bold = True
    wsResult.Cells(1, 1).Font.Size = 15

    With wsResult.Range("A1:C1")
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = -4108
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

    wsResult.Cells(2, 2).Value = "第 53 期 " & Format(maxDate, "YYYY年 MM月 DD日") & " 現在"
    With wsResult.Range("B2:C2")
        .HorizontalAlignment = xlCenterAcrossSelection
    End With

    wsResult.Cells(3, 5).Value = "(單位:원)"
    wsResult.Cells(3, 5).HorizontalAlignment = -4152 'xlRight

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
    
    
    wsResult.Cells(2, 4).Value = "第 52 期 " & Format(DateAdd("yyyy", -1, maxDate), "yyyy年 MM月 dd日") & " 現在"
    With wsResult.Range("D2:E2")
        .HorizontalAlignment = xlCenterAcrossSelection
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
    wsResult.Range("A4:E5").Interior.Color = RGB(217, 217, 217) ' 연회색


    Dim outRow As Long
    outRow = 6

    Dim typeList(1 To 3) As String
    typeList(1) = "자산"
    typeList(2) = "부채"
    typeList(3) = "자본"

    ' 레벨: 1=중분류2, 2=중분류3(없을 수 있음), 3=계정구분명
    Dim lvlCode(1 To 3) As String
    Dim lvlName(1 To 3) As String
    Dim lvlHeaderRow(1 To 3) As Long
    Dim lvlSum(1 To 3) As Double
    Dim lvlAcctCode(1 To 3) As String  ' 계정코드 저장 (감가상각 차감용)

    Dim lvlColor(1 To 3) As Long
    lvlColor(1) = RGB(198, 217, 241)
    lvlColor(2) = RGB(222, 235, 250)
    lvlColor(3) = RGB(235, 241, 222)

    Dim cnt2 As Integer, cnt3 As Integer, cnt4 As Integer

    Dim 자산합계 As Double, 부채합계 As Double, 자본합계 As Double
    자산합계 = 0: 부채합계 = 0: 자본합계 = 0

    Dim t As Integer
    For t = 1 To 3
        Dim tLabel As String
        tLabel = typeList(t)

        Dim hasType As Boolean
        hasType = False
        For m = 1 To nMaster
            If CStr(arrM(m, 3)) = tLabel Then hasType = True: Exit For
        Next m
        If Not hasType Then GoTo NextType

        wsResult.Cells(outRow, 1).Value = tLabel
        outRow = outRow + 1

        Dim tSum As Double
        tSum = 0
        Dim lv As Integer
        For lv = 1 To 3
            lvlCode(lv) = "|__초기값__|"
            lvlHeaderRow(lv) = 0
            lvlSum(lv) = 0
            lvlAcctCode(lv) = ""
        Next lv

        cnt2 = 0: cnt3 = 0: cnt4 = 0

        Dim bPrev3311200 As Boolean
        bPrev3311200 = False

        For m = 1 To nMaster
            If CStr(arrM(m, 3)) <> tLabel Then GoTo NextMaster

            ' 이전이 3311200이고 현재가 다르면 → 당기순이익 행 삽입
            If bPrev3311200 And CStr(arrM(m, 1)) <> "3311200" Then
                wsResult.Cells(outRow, 1).Value = Space(12) & "당기순이익"
                wsResult.Cells(outRow, 2).Value = 당기순이익
                wsResult.Cells(outRow, 2).NumberFormat = "#,##0"

                outRow = outRow + 1
                bPrev3311200 = False
            End If

            Dim newCode(1 To 3) As String
            Dim newName(1 To 3) As String
            newCode(1) = CStr(arrM(m, 6)): newName(1) = CStr(arrM(m, 6))
            newCode(2) = CStr(arrM(m, 7)): newName(2) = CStr(arrM(m, 7))
            newCode(3) = CStr(arrM(m, 4)): newName(3) = CStr(arrM(m, 5))

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
                        Call BS_헤더마감(wsResult, lvlHeaderRow(lv), lvlSum(lv), lvlColor(lv), lvlAcctCode(lv), arrSCode, arrSBal, nS)
                        lvlSum(lv) = 0
                        lvlHeaderRow(lv) = 0
                        lvlAcctCode(lv) = ""
                    End If
                Next lv

                For lv = changeLv To 3
                    lvlCode(lv) = newCode(lv)
                    lvlName(lv) = newName(lv)

                    If lv = 1 Then
                        cnt2 = cnt2 + 1
                        cnt3 = 0: cnt4 = 0
                    ElseIf lv = 2 Then
                        cnt3 = cnt3 + 1
                        cnt4 = 0
                    ElseIf lv = 3 Then
                        cnt4 = cnt4 + 1
                    End If

                    If lv = 3 Then
                        lvlHeaderRow(lv) = outRow
                        lvlAcctCode(lv) = CStr(arrM(m, 1))  ' 계정코드 저장
                        wsResult.Cells(outRow, 1).Value = Space(lv * 3) & cnt4 & ". " & newName(lv)
                        outRow = outRow + 1
                    ElseIf newCode(lv) <> "" Then
                        lvlHeaderRow(lv) = outRow
                        lvlAcctCode(lv) = ""
                        If lv = 1 Then
                            wsResult.Cells(outRow, 1).Value = Space(lv * 3) & Excel.WorksheetFunction.Roman(cnt2) & ". " & newName(lv)
                        ElseIf lv = 2 Then
                            wsResult.Cells(outRow, 1).Value = Space(lv * 3) & "(" & cnt3 & ") " & newName(lv)
                        End If
                        outRow = outRow + 1
                    Else
                        lvlHeaderRow(lv) = 0
                        lvlAcctCode(lv) = ""
                    End If
                Next lv
            End If

            ' ---- 개별 세목 금액 계산 ----
            Dim dBal As Double, dBalD As Double, dBalC As Double
            Dim sBalDir As String
            dBal = 0: dBalD = 0: dBalC = 0
            sBalDir = CStr(arrM(m, 8))

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

            Dim dAmt As Double
            If sBalDir = "D" Then
                dAmt = dBalD - dBalC
            Else
                dAmt = dBalC - dBalD
            End If

            ' 이월이익잉여금(3311200)은 당기순이익 반영
            If CStr(arrM(m, 1)) = "3311200" Then
                dAmt = dAmt + 당기순이익
            End If

            ' ★ 감가상각 처리 & lvlSum 누적 금액 결정
            Dim sDeprCode As String
            sDeprCode = Get감가상각충당금코드(CStr(arrM(m, 1)))

            Dim dAcctAmount As Double  ' lvlSum에 실제 누적할 금액
            dAcctAmount = dAmt  ' 기본값은 dAmt

            If sDeprCode <> "" Then
                ' 감가상각충당금 금액 조회
                Dim dDeprAmt As Double
                dDeprAmt = 0
                For i = 1 To nS
                    If arrSCode(i) = sDeprCode Then
                        dDeprAmt = arrSBal(i)
                        Exit For
                    End If
                Next i

                ' ★ 순자산가액으로 누적 금액 변경!
                dAcctAmount = dAmt - dDeprAmt

                ' 취득금액을 B열에 표시
                wsResult.Cells(outRow - 1, 2).Value = dAmt
                wsResult.Cells(outRow - 1, 2).NumberFormat = "#,##0"

                ' 감가상각누계액 행 추가
                ' B열: 감가상각충당금, C열: 비움 (상위에서 이미 차감됨)
                wsResult.Cells(outRow, 1).Value = Space(12) & "감가상각누계액"
                wsResult.Cells(outRow, 2).Value = dDeprAmt
                wsResult.Cells(outRow, 2).NumberFormat = "#,##0"

                outRow = outRow + 1
            End If

            ' ★ lvlSum에 누적 (감가상각대상이면 순자산가액, 아니면 원래 금액)
            lvlSum(1) = lvlSum(1) + dAcctAmount
            lvlSum(2) = lvlSum(2) + dAcctAmount
            lvlSum(3) = lvlSum(3) + dAcctAmount
            tSum = tSum + dAcctAmount

            ' 현재 계정이 3311200인가 체크
            If CStr(arrM(m, 1)) = "3311200" Then
                bPrev3311200 = True
            Else
                bPrev3311200 = False
            End If

NextMaster:
        Next m

        ' 루프 끝난 후: 마지막이 3311200이었으면 당기순이익 추가
        If bPrev3311200 Then
            wsResult.Cells(outRow, 1).Value = Space(12) & "당기순이익"
            wsResult.Cells(outRow, 2).Value = 당기순이익
            wsResult.Cells(outRow, 2).NumberFormat = "#,##0"

            outRow = outRow + 1
        End If

        For lv = 3 To 1 Step -1
            If lvlHeaderRow(lv) > 0 Then
                Call BS_헤더마감(wsResult, lvlHeaderRow(lv), lvlSum(lv), lvlColor(lv), lvlAcctCode(lv), arrSCode, arrSBal, nS)
            End If
        Next lv

        Call BS_소계행출력(wsResult, outRow, tLabel & " 총계", tSum, RGB(220, 230, 241))
        outRow = outRow + 1

        If tLabel = "자산" Then 자산합계 = tSum
        If tLabel = "부채" Then 부채합계 = tSum
        If tLabel = "자본" Then 자본합계 = tSum

NextType:
    Next t

    Call BS_소계행출력(wsResult, outRow, "부채와자본 총계", 부채합계 + 자본합계, RGB(180, 198, 231))

    ' ------------------------------------------
    ' 4. 서식
    ' ------------------------------------------
    wsResult.Range("B4:E" & outRow).NumberFormat = "#,##0;(#,##0);"""""  '음수는 괄호 표시, 0은 공백

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

    ' ------------------------------------------
    ' 5. 대차평형 검증
    ' ------------------------------------------
    Dim diff As Double
    diff = 자산합계 - (부채합계 + 자본합계)

    Dim msg As String
    msg = "재무상태표 생성 완료" & vbCrLf & vbCrLf & _
          "자산합계 : " & Format(자산합계, "#,##0") & vbCrLf & _
          "부채합계 : " & Format(부채합계, "#,##0") & vbCrLf & _
          "자본합계 : " & Format(자본합계, "#,##0") & " (당기순이익 " & Format(당기순이익, "#,##0") & " 반영)"

    If Abs(diff) < 1 Then
        msg = msg & vbCrLf & vbCrLf & "대차 일치 OK!"
    Else
        msg = msg & vbCrLf & vbCrLf & "대차 불일치! NG!  차이 : " & Format(diff, "#,##0")
    End If

    MsgBox msg, vbInformation

End Sub


'''===========================================
''' 유형자산 코드 → 감가상각충당금 코드 매칭
''' ★ 유지보수: 새로운 감가상각 계정 추가 시 여기만 수정
'''===========================================
''Private Function Get감가상각충당금코드(sAssetCode As String) As String
''
''    ' 마스터에서는 제외된 감가상각 세부 계정 코드 반환
''    ' 원장에서 실제 금액을 조회할 때 사용
''    Select Case sAssetCode
''        Case "1310200"
''            Get감가상각충당금코드 = "2210701"  ' 건물
''        Case "1310300"
''            Get감가상각충당금코드 = "2210702"  ' 구축물
''        Case "1310400"
''            Get감가상각충당금코드 = "2210703"  ' 기계장치
''        Case "1310500"
''            Get감가상각충당금코드 = "2210704"  ' 차량운반구
''        Case "1310700"
''            Get감가상각충당금코드 = "2210706"  ' 공기구비품
''        Case Else
''            Get감가상각충당금코드 = ""
''    End Select
''
''End Function


Private Function Get감가상각충당금코드(sAssetCode As String) As String

    ' ★ TODO: 결산마감처리 후 개별 코드로 변경
    ' 현재는 외부시스템에서 2210700(통합코드)로 이월되므로
    ' 공기구비품에만 임시 매핑
    
    Select Case sAssetCode
        Case "1310200"
            Get감가상각충당금코드 = ""  ' 건물 - 향후 "2210701" 사용 예정
        Case "1310300"
            Get감가상각충당금코드 = ""  ' 구축물 - 향후 "2210702" 사용 예정
        Case "1310400"
            Get감가상각충당금코드 = ""  ' 기계장치 - 향후 "2210703" 사용 예정
        Case "1310500"
            Get감가상각충당금코드 = ""  ' 차량운반구 - 향후 "2210704" 사용 예정
        Case "1310700"
            Get감가상각충당금코드 = "2210700"   '공기구비품 - 향후 "2210706" 사용 예정
        Case Else
            Get감가상각충당금코드 = ""
    End Select

End Function



'===========================================
' BS용 소계행 출력 헬퍼
'===========================================
Private Sub BS_소계행출력(ws As Worksheet, r As Long, sLabel As String, sumVal As Double, bgColor As Long)
    With ws
        .Cells(r, 1).Value = sLabel
        .Cells(r, 3).Value = sumVal
        .Cells(r, 3).NumberFormat = "#,##0"
        .Range("A" & r & ":E" & r).Interior.Color = bgColor
    End With
End Sub


'===========================================
' BS용 중분류2/3/계정구분명 헤더 마감 헬퍼
' ★ 감가상각 유형자산의 경우 순자산가액 자동 계산
'    C열 = 취득금액 - 감가상각누계액
'===========================================

Private Sub BS_헤더마감(ws As Worksheet, headerRow As Long, sumVal As Double, bgColor As Long, sAcctCode As String, arrSCode() As String, arrSBal() As Double, nS As Long)

    ' ★ 이미 메인 루프에서 순자산가액으로 계산되었으므로
    '    감가상각 차감 없이 그대로 표시
    Dim dDisplayVal As Double
    dDisplayVal = sumVal

    With ws
         .Cells(headerRow, 3).Value = dDisplayVal
         .Cells(headerRow, 3).NumberFormat = "#,##0"
    End With
End Sub
