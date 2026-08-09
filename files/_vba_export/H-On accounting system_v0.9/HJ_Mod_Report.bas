Option Explicit

Function 원장데이터가져오기( _
    Optional filterCode As String = "", _
    Optional dateFrom As Variant = "", _
    Optional dateTo As Variant = "") As Variant

    Dim ws분개   As Worksheet
    Dim ws마스터 As Worksheet
    Set ws분개 = sht분개장
    Set ws마스터 = sht계정코드

    '=====================
    ' 1. 계정코드마스터 → Dictionary
    '=====================
    Dim dictMaster As Object
    Set dictMaster = CreateObject("Scripting.Dictionary")

    Dim masterLastRow As Long
    masterLastRow = ws마스터.Cells(ws마스터.Rows.Count, 1).End(xlUp).Row

    Dim masterData As Variant
    masterData = ws마스터.Range("A2:J" & masterLastRow).Value

    Dim i As Long
    For i = 1 To UBound(masterData, 1)
        Dim sCode As String
        sCode = CStr(masterData(i, 1))
        If Not dictMaster.Exists(sCode) Then
            dictMaster.Add sCode, Array(masterData(i, 2), masterData(i, 6))
        End If
    Next i

    '=====================
    ' 2. 분개장 읽기 → 정렬
    '=====================
    Dim LastRow As Long
    LastRow = ws분개.Cells(ws분개.Rows.Count, 1).End(xlUp).Row

    Dim rawData As Variant
    rawData = ws분개.Range("A2:M" & LastRow).Value   ' ★ J → M (외화금액 컬럼까지 포함)
    rawData = 분개장정렬(rawData)

    '=====================
    ' 3. 계정코드별 행 분류
    '    - 전기이월행 (적요="전기이월" AND 1월1일)
    '    - dateFrom 미만 행 (잔액합산용)
    '    - dateFrom ~ dateTo 행 (출력용)
    '=====================
    Dim dictIW     As Object  ' 전기이월 행 인덱스 (계정코드별)
    Dim dictPre    As Object  ' dateFrom 미만 행 (잔액합산용)
    Dim dictIndex  As Object  ' 출력용 행
    Set dictIW = CreateObject("Scripting.Dictionary")
    Set dictPre = CreateObject("Scripting.Dictionary")
    Set dictIndex = CreateObject("Scripting.Dictionary")

    Dim rowDate As Date
    Dim bDateFrom As Boolean
    Dim bDateTo   As Boolean
    bDateFrom = IsDate(dateFrom)
    bDateTo = IsDate(dateTo)

    For i = 1 To UBound(rawData, 1)
        Dim code As String
        code = CStr(rawData(i, 9))

        ' 계정코드 필터
        If filterCode <> "" And code <> filterCode Then GoTo NextRow

        Dim bHasDate As Boolean
        bHasDate = Not IsEmpty(rawData(i, 2))

        ' 전기이월 판단: 적요="전기이월" AND 회계일자 1월1일
        Dim bIsIW As Boolean
        bIsIW = False
        If CStr(rawData(i, 5)) = "전기이월" And bHasDate Then
            Dim iwDate As Date
            iwDate = CDate(rawData(i, 2))
            If Month(iwDate) = 1 And Day(iwDate) = 1 Then
                bIsIW = True
            End If
        End If

        If bIsIW Then
            ' 전기이월 행
            If Not dictIW.Exists(code) Then dictIW.Add code, New Collection
            dictIW(code).Add i
      ElseIf bHasDate Then
          rowDate = CDate(rawData(i, 2))
          
          ' dateFrom 미만 → 잔액합산용
          If bDateFrom Then
              If rowDate < CDate(dateFrom) Then
                  If Not dictPre.Exists(code) Then dictPre.Add code, New Collection
                  dictPre(code).Add i
                  GoTo NextRow
              End If
          End If
          
          ' dateTo 초과 → 스킵
          If bDateTo Then
              If rowDate > CDate(dateTo) Then GoTo NextRow
          End If
          
          ' 출력용
          If Not dictIndex.Exists(code) Then dictIndex.Add code, New Collection
          dictIndex(code).Add i
      
      End If

NextRow:
    Next i

    '=====================
    ' 4. 계정코드 수집 (세 딕셔너리 합산)
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
        원장데이터가져오기 = Empty
        Exit Function
    End If

    ' 계정코드 정렬
    Dim codes() As String
    ReDim codes(0 To codeCount - 1)
    Dim idx As Long
    idx = 0
    For Each c In dictAll.Keys
        codes(idx) = CStr(c)
        idx = idx + 1
    Next c

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
    ' 5. 출력 배열 준비
    '=====================
    Dim totalRows As Long
    totalRows = (LastRow - 1) + (codeCount * 2) + 10  ' 여유있게

    Dim result() As Variant
    ReDim result(1 To totalRows, 1 To 12)   ' ★ 11 → 12

    Dim outRow As Long
    outRow = 1

    '=====================
    ' 6. 계정별 집계 → 배열 저장
    '=====================
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

        Dim 누적잔액 As Double
        Dim 소계차변 As Double
        Dim 소계대변 As Double
        누적잔액 = 0: 소계차변 = 0: 소계대변 = 0

        Dim 누적외화잔액 As Double   ' ★ 추가
        누적외화잔액 = 0             ' ★ 추가

        Dim 차변 As Double
        Dim 대변 As Double
        Dim 외화금액raw As Double    ' ★ 추가
        Dim 외화증감   As Double     ' ★ 추가
        Dim rowIdx As Variant

        '--- 전기이월 잔액 계산
        Dim iw잔액 As Double
        Dim iw외화잔액 As Double      ' ★ 추가
        iw잔액 = 0
        iw외화잔액 = 0                ' ★ 추가
        If dictIW.Exists(code) Then
            For Each rowIdx In dictIW(code)
                차변 = IIf(IsNumeric(rawData(rowIdx, 7)), CDbl(rawData(rowIdx, 7)), 0)
                대변 = IIf(IsNumeric(rawData(rowIdx, 8)), CDbl(rawData(rowIdx, 8)), 0)
                If 잔액방향 = "D" Then
                    iw잔액 = iw잔액 + 차변 - 대변
                Else
                    iw잔액 = iw잔액 + 대변 - 차변
                End If

                ' ★ 외화잔액 누적 (전기이월)
                외화금액raw = IIf(IsNumeric(rawData(rowIdx, 13)), CDbl(rawData(rowIdx, 13)), 0)
                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    iw외화잔액 = iw외화잔액 + 외화증감
                Else
                    iw외화잔액 = iw외화잔액 - 외화증감
                End If
            Next rowIdx
        End If

        '--- dateFrom 미만 잔액 합산
        Dim pre잔액 As Double
        Dim pre외화잔액 As Double      ' ★ 추가
        pre잔액 = 0
        pre외화잔액 = 0                ' ★ 추가
        If dictPre.Exists(code) Then
            For Each rowIdx In dictPre(code)
                차변 = IIf(IsNumeric(rawData(rowIdx, 7)), CDbl(rawData(rowIdx, 7)), 0)
                대변 = IIf(IsNumeric(rawData(rowIdx, 8)), CDbl(rawData(rowIdx, 8)), 0)
                If 잔액방향 = "D" Then
                    pre잔액 = pre잔액 + 차변 - 대변
                Else
                    pre잔액 = pre잔액 + 대변 - 차변
                End If

                ' ★ 외화잔액 누적 (dateFrom 미만)
                외화금액raw = IIf(IsNumeric(rawData(rowIdx, 13)), CDbl(rawData(rowIdx, 13)), 0)
                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    pre외화잔액 = pre외화잔액 + 외화증감
                Else
                    pre외화잔액 = pre외화잔액 - 외화증감
                End If
            Next rowIdx
        End If

        '--- 전기이월 행 출력 (iw잔액 + pre잔액 합산)
        누적잔액 = iw잔액 + pre잔액
        누적외화잔액 = iw외화잔액 + pre외화잔액   ' ★ 추가

        result(outRow, 1) = code
        result(outRow, 2) = 계정명
       '' result(outRow, 3) = IIf(bDateFrom, CDate(dateFrom), "")
        
        If bDateFrom Then
            result(outRow, 3) = CDate(dateFrom)
        Else
            result(outRow, 3) = ""
        End If
        
        result(outRow, 6) = "전기이월"
        result(outRow, 10) = 누적잔액
        result(outRow, 11) = "IW"  ' 행유형: 전기이월
        result(outRow, 12) = 누적외화잔액   ' ★ 추가
        outRow = outRow + 1

        '--- 출력용 거래 행
        If dictIndex.Exists(code) Then
            For Each rowIdx In dictIndex(code)
                차변 = IIf(IsNumeric(rawData(rowIdx, 7)), CDbl(rawData(rowIdx, 7)), 0)
                대변 = IIf(IsNumeric(rawData(rowIdx, 8)), CDbl(rawData(rowIdx, 8)), 0)

                If 잔액방향 = "D" Then
                    누적잔액 = 누적잔액 + 차변 - 대변
                Else
                    누적잔액 = 누적잔액 + 대변 - 차변
                End If
                소계차변 = 소계차변 + 차변
                소계대변 = 소계대변 + 대변

                ' ★ 외화잔액 누적 (출력용 거래)
                외화금액raw = IIf(IsNumeric(rawData(rowIdx, 13)), CDbl(rawData(rowIdx, 13)), 0)
                외화증감 = IIf(차변 > 0, 외화금액raw, IIf(대변 > 0, -외화금액raw, 0))
                If 잔액방향 = "D" Then
                    누적외화잔액 = 누적외화잔액 + 외화증감
                Else
                    누적외화잔액 = 누적외화잔액 - 외화증감
                End If

                result(outRow, 1) = code
                result(outRow, 2) = 계정명
                result(outRow, 3) = rawData(rowIdx, 2)
                result(outRow, 4) = rawData(rowIdx, 3)
                result(outRow, 5) = rawData(rowIdx, 4)
                result(outRow, 6) = rawData(rowIdx, 5)
                result(outRow, 7) = rawData(rowIdx, 6)
                result(outRow, 8) = 차변
                result(outRow, 9) = 대변
                result(outRow, 10) = 누적잔액
                result(outRow, 11) = "D"
                result(outRow, 12) = 누적외화잔액   ' ★ 추가
                outRow = outRow + 1
            Next rowIdx
        End If

        '--- 소계행
        result(outRow, 1) = code
        result(outRow, 2) = 계정명
        result(outRow, 6) = "소계"
        result(outRow, 8) = 소계차변
        result(outRow, 9) = 소계대변
        result(outRow, 10) = 누적잔액
        result(outRow, 11) = "S"
        result(outRow, 12) = 누적외화잔액   ' ★ 추가
        outRow = outRow + 1

    Next i

    '=====================
    ' 7. 실제 사용 행만 잘라서 반환
    '=====================
    Dim finalResult() As Variant
    ReDim finalResult(1 To outRow - 1, 1 To 12)   ' ★ 11 → 12
    Dim k As Long
    For i = 1 To outRow - 1
        For k = 1 To 12   ' ★ 11 → 12
            finalResult(i, k) = result(i, k)
        Next k
    Next i

    원장데이터가져오기 = finalResult

End Function


Sub 총계정원장생성()

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    '=====================
    ' 1. 데이터 가져오기 (공용함수 호출)
    '=====================
    Dim data As Variant
    data = 원장데이터가져오기()  ' filterCode="", 날짜 전체

    If IsEmpty(data) Then
        MsgBox "데이터가 없습니다.", vbExclamation
        GoTo CleanUp
    End If

    '=====================
    ' 2. 출력 시트 준비
    '=====================
    Dim ws원장 As Worksheet
    On Error Resume Next
    Set ws원장 = ThisWorkbook.Sheets("총계정원장_클로드")
    On Error GoTo ErrorHandler

    If ws원장 Is Nothing Then
        Set ws원장 = ThisWorkbook.Sheets.Add
        ws원장.Name = "총계정원장_클로드"
    Else
        ws원장.Cells.Clear
    End If

'=====================
    ' 3. 타이틀 출력
    '=====================
    Dim 최초일 As Date
    Dim 최종일 As Date
    최초일 = sht분개장.Cells(2, 2).Value
    최종일 = sht분개장.Cells(sht분개장.Rows.Count, 2).End(xlUp).Value

    With ws원장.Range("A1:J1")
        .Merge
        .Value = "총계정원장   " & Format(최초일, "yyyy-mm-dd") & " ~ " & Format(최종일, "yyyy-mm-dd")
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


Sub 계정별명세출력(filterCode As String, dateFrom As Variant, dateTo As Variant)

    Dim startTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    '=====================
    ' 1. 데이터 가져오기
    '=====================
    Dim data As Variant
    data = 원장데이터가져오기(filterCode, dateFrom, dateTo)

    If IsEmpty(data) Then
        MsgBox "조회된 데이터가 없습니다.", vbExclamation
        GoTo CleanUp
    End If

        '=====================
      ' 2. 출력 시트 준비
      '    시트명: 계정별명세_계정코드
      '=====================
      Dim wsName As String
      wsName = "계정별명세_" & filterCode
      
      Dim ws명세 As Worksheet
      Dim ws명세존재 As Boolean
      ws명세존재 = False
      
      Dim ws As Worksheet
      For Each ws In ThisWorkbook.Sheets
          If ws.Name = wsName Then
              ws명세존재 = True
              Exit For
          End If
      Next ws
      
      If ws명세존재 Then
          Set ws명세 = ThisWorkbook.Sheets(wsName)
          ws명세.Cells.Clear
      Else
          Set ws명세 = ThisWorkbook.Sheets.Add
          ws명세.Name = wsName
      End If

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



'===========================================
' 합계잔액시산표 생성
' 모듈: HJ_Mod_Report 에 추가
'===========================================
Public Sub 합계잔액시산표생성()

    Dim wsResult  As Worksheet
    Dim wsMaster  As Worksheet
    Dim arrData() As Variant

    Dim tStart As Double '소요시간 체크하기 위해 코파일럿에게 물어서 변수 추가함
    tStart = Timer
    
    
    ' 엑셀 화면/자동계산/이벤트 중지 (시산표 생성 속도 개선) - 코파일럿
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ' -------------------------------
    '  여기부터가 무거운 작업이므로 엑셀 이벤트는 그 전단계에서 작성 - 코파일럿
    ' -------------------------------
    
    ' ------------------------------------------
    ' 1. 원장 데이터 가져오기 (전체기간, 인자없이)
    ' ------------------------------------------
    arrData = 원장데이터가져오기()
    If Not IsArray(arrData) Then
        MsgBox "원장 데이터를 가져올 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' ------------------------------------------
    ' 2. 계정코드마스터 로드 (사용여부=Y만)
    ' ------------------------------------------
    Set wsMaster = sht계정코드  ' 기존 시트변수 활용
    
    Dim lastRowM As Long
    lastRowM = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    
    Dim masterData As Variant
    masterData = wsMaster.Range("A2:J" & lastRowM).Value
    ' 컬럼: 1=세목코드, 2=계정명, 3=계정유형, 4=재무제표,
    '        5=출력순서, 6=잔액방향, 7=계정구분, 8=계정구분명, 9=계정구분명_JP, 10=사용여부
    
    ' 사용여부=Y인 계정만 수집
    Dim nMaster As Long
    nMaster = 0
    Dim i As Long
    For i = 1 To UBound(masterData, 1)
        If CStr(masterData(i, 10)) = "Y" Then nMaster = nMaster + 1
    Next i
    
    ' 마스터 배열: 코드/계정명/계정유형/재무제표/출력순서/잔액방향
    Dim arrM() As Variant
    ReDim arrM(1 To nMaster, 1 To 6)
    Dim m As Long
    m = 0
    For i = 1 To UBound(masterData, 1)
        If CStr(masterData(i, 10)) = "Y" Then
            m = m + 1
            arrM(m, 1) = CStr(masterData(i, 1))  ' 세목코드
            arrM(m, 2) = masterData(i, 2)          ' 계정명
            arrM(m, 3) = masterData(i, 3)          ' 계정유형
            arrM(m, 4) = masterData(i, 4)          ' 재무제표
            arrM(m, 5) = CLng(masterData(i, 5))    ' 출력순서
            arrM(m, 6) = CStr(masterData(i, 6))    ' 잔액방향
        End If
    Next i
    
    ' ------------------------------------------
    ' 3. 정렬: 계정유형순(자산→부채→자본→수익→비용) + 출력순서
    ' ------------------------------------------
    Dim j As Long, k As Long
    Dim tmpRow(1 To 6) As Variant
    
    For i = 1 To nMaster - 1
        For j = 1 To nMaster - i
            If 계정정렬키(CStr(arrM(j, 3)), CLng(arrM(j, 5))) > _
               계정정렬키(CStr(arrM(j + 1, 3)), CLng(arrM(j + 1, 5))) Then
                For k = 1 To 6
                    tmpRow(k) = arrM(j, k)
                    arrM(j, k) = arrM(j + 1, k)
                    arrM(j + 1, k) = tmpRow(k)
                Next k
            End If
        Next j
    Next i
    
    ' ------------------------------------------
    ' 4. 소계행(S) + IW행에서 계정별 합계/잔액 수집
    '    arrData 컬럼:
    '    1=계정코드, 2=계정명, 8=차변, 9=대변, 10=누적잔액, 11=행유형
    ' ------------------------------------------
    ' 소계행: 기간 거래 차변합계/대변합계, 누적잔액
    ' IW행:   전기이월 차변/대변 (잔액방향에 따라 역산)
    
    ' 계정코드별로 Dictionary 대신 배열로 처리
    ' 먼저 소계행 수집
    Dim nS As Long
    nS = 0
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "S" Then nS = nS + 1
    Next i
    
    Dim arrSCode()   As String
    Dim arrSDebit()  As Double   ' 소계 차변합계 (기간 거래분)
    Dim arrSCredit() As Double   ' 소계 대변합계 (기간 거래분)
    Dim arrSBal()    As Double   ' 누적잔액
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
    
    ' IW행에서 전기이월 차변/대변 수집
    ' 잔액방향 D: iw잔액>0이면 차변이월, <0이면 대변이월
    ' → 합계차변 = 소계차변 + iw차변이월, 합계대변 = 소계대변 + iw대변이월
    ' IW행의 누적잔액을 잔액방향으로 역산
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
    ' 5. 시산표 시트 출력
    ' ------------------------------------------
    Set wsResult = sht시산표

    ' 헤더 4행 보존, 5행부터 내용+서식 초기화
    Dim lastRowR As Long
    lastRowR = wsResult.Cells(wsResult.Rows.Count, 3).End(xlUp).Row
    If lastRowR > 4 Then
        With wsResult.Rows("5:" & lastRowR)
            .ClearContents
            .ClearFormats
        End With
    End If

    ' 기준년월 (B2)
    Dim maxDate As Date
    maxDate = DateSerial(1900, 1, 1)
    For i = 1 To UBound(arrData, 1)
        If arrData(i, 11) = "D" Then
            If IsDate(arrData(i, 3)) Then
                If CDate(arrData(i, 3)) > maxDate Then maxDate = CDate(arrData(i, 3))
            End If
        End If
    Next i
    wsResult.Cells(2, 2).Value = Format(maxDate, "YYYY년 MM월")

    Dim outRow As Long
    outRow = 5

    Dim typeList(1 To 5) As String
    typeList(1) = "자산"
    typeList(2) = "부채"
    typeList(3) = "자본"
    typeList(4) = "수익"
    typeList(5) = "비용"

    ' ------------------------------------------
    ' 6. 유형 → 중분류(상단 헤더) → 개별계정 순 출력
    ' ------------------------------------------
    Dim t         As Integer
    Dim tSumD     As Double, tSumC As Double
    Dim tBalD     As Double, tBalC As Double
    Dim grpSumD   As Double, grpSumC As Double
    Dim grpBalD   As Double, grpBalC As Double
    Dim prevGrp   As String
    Dim grpName   As String
    Dim grpHeaderRow As Long   ' ★ 중분류 헤더 행 번호 기억

    Dim totalSumD As Double, totalSumC As Double
    Dim totalBalD As Double, totalBalC As Double
    totalSumD = 0: totalSumC = 0: totalBalD = 0: totalBalC = 0

    For t = 1 To 5
        Dim tLabel As String
        tLabel = typeList(t)

        Dim hasType As Boolean
        hasType = False
        For m = 1 To nMaster
            If CStr(arrM(m, 3)) = tLabel Then hasType = True: Exit For
        Next m
        If Not hasType Then GoTo NextType

        ' [ 유형 ] 헤더
        wsResult.Cells(outRow, 4).Value = "[ " & tLabel & " ]"
        wsResult.Cells(outRow, 4).Font.Bold = True
        wsResult.Cells(outRow, 4).Font.Color = RGB(0, 70, 127)
        outRow = outRow + 1

        tSumD = 0: tSumC = 0: tBalD = 0: tBalC = 0
        prevGrp = ""
        grpHeaderRow = 0

        For m = 1 To nMaster
            If CStr(arrM(m, 3)) <> tLabel Then GoTo NextMaster

            ' masterData에서 중분류 코드/명 찾기
            Dim curGrpCode As String, curGrpName As String
            curGrpCode = "": curGrpName = ""
            Dim mi As Long
            For mi = 1 To UBound(masterData, 1)
                If CStr(masterData(mi, 1)) = CStr(arrM(m, 1)) Then
                    curGrpCode = CStr(masterData(mi, 7))
                    curGrpName = CStr(masterData(mi, 8))
                    Exit For
                End If
            Next mi

            ' 중분류 바뀌면 직전 중분류 헤더행에 합계 금액 채우기
            If curGrpCode <> prevGrp And prevGrp <> "" Then
                ' ★ 헤더행으로 돌아가서 금액 채우기
                With wsResult
                    If grpBalD <> 0 Then .Cells(grpHeaderRow, 1).Value = grpBalD
                    If grpSumD <> 0 Then .Cells(grpHeaderRow, 2).Value = grpSumD
                    If grpSumC <> 0 Then .Cells(grpHeaderRow, 5).Value = grpSumC
                    If grpBalC <> 0 Then .Cells(grpHeaderRow, 6).Value = grpBalC
                    .Range("A" & grpHeaderRow & ":F" & grpHeaderRow).Font.Bold = True
                    .Range("A" & grpHeaderRow & ":F" & grpHeaderRow).Interior.Color = RGB(235, 241, 222)
                    .Range("A" & grpHeaderRow & ",B" & grpHeaderRow & _
                           ",E" & grpHeaderRow & ",F" & grpHeaderRow).NumberFormat = "#,##0"
                End With
                grpSumD = 0: grpSumC = 0: grpBalD = 0: grpBalC = 0
            End If

            ' 새 중분류 헤더행 출력 (금액은 나중에 채움)
            If curGrpCode <> prevGrp Then
                prevGrp = curGrpCode
                grpName = curGrpName
                grpHeaderRow = outRow   ' ★ 행 번호 기억
                wsResult.Cells(outRow, 4).Value = curGrpName
                outRow = outRow + 1
                grpSumD = 0: grpSumC = 0: grpBalD = 0: grpBalC = 0
            End If

            ' 소계행에서 데이터 찾기
            Dim dSumD As Double, dSumC As Double
            Dim dBal  As Double, dBalD As Double, dBalC As Double
            Dim sBalDir As String
            dSumD = 0: dSumC = 0: dBal = 0: dBalD = 0: dBalC = 0
            sBalDir = CStr(arrM(m, 6))

            For i = 1 To nS
                If arrSCode(i) = CStr(arrM(m, 1)) Then
                    dSumD = arrSDebit(i)
                    dSumC = arrSCredit(i)
                    dBal = arrSBal(i)
                    Exit For
                End If
            Next i

            ' IW 잔액 가산
            Dim dIWBal As Double
            dIWBal = 0
            For i = 1 To nIW
                If arrIWCode(i) = CStr(arrM(m, 1)) Then
                    dIWBal = arrIWBal(i)
                    Exit For
                End If
            Next i

            If sBalDir = "D" Then
                dSumD = dSumD + IIf(dIWBal > 0, dIWBal, 0)
                dSumC = dSumC + IIf(dIWBal < 0, Abs(dIWBal), 0)
                dBalD = IIf(dBal >= 0, dBal, 0)
                dBalC = IIf(dBal < 0, Abs(dBal), 0)
            Else
                dSumC = dSumC + IIf(dIWBal > 0, dIWBal, 0)
                dSumD = dSumD + IIf(dIWBal < 0, Abs(dIWBal), 0)
                dBalC = IIf(dBal >= 0, dBal, 0)
                dBalD = IIf(dBal < 0, Abs(dBal), 0)
            End If

            ' 개별 계정 출력
            With wsResult
                If dBalD <> 0 Then .Cells(outRow, 1).Value = dBalD
                If dSumD <> 0 Then .Cells(outRow, 2).Value = dSumD
                .Cells(outRow, 3).Value = CStr(arrM(m, 1))
                .Cells(outRow, 4).Value = CStr(arrM(m, 2))
                If dSumC <> 0 Then .Cells(outRow, 5).Value = dSumC
                If dBalC <> 0 Then .Cells(outRow, 6).Value = dBalC
            End With
            outRow = outRow + 1

            ' 중분류/유형 누산
            grpSumD = grpSumD + dSumD: grpSumC = grpSumC + dSumC
            grpBalD = grpBalD + dBalD: grpBalC = grpBalC + dBalC
            tSumD = tSumD + dSumD: tSumC = tSumC + dSumC
            tBalD = tBalD + dBalD: tBalC = tBalC + dBalC

NextMaster:
        Next m

        ' 마지막 중분류 헤더행에 금액 채우기
        If grpHeaderRow > 0 Then
            With wsResult
                If grpBalD <> 0 Then .Cells(grpHeaderRow, 1).Value = grpBalD
                If grpSumD <> 0 Then .Cells(grpHeaderRow, 2).Value = grpSumD
                If grpSumC <> 0 Then .Cells(grpHeaderRow, 5).Value = grpSumC
                If grpBalC <> 0 Then .Cells(grpHeaderRow, 6).Value = grpBalC
                .Range("A" & grpHeaderRow & ":F" & grpHeaderRow).Font.Bold = True
                .Range("A" & grpHeaderRow & ":F" & grpHeaderRow).Interior.Color = RGB(235, 241, 222)
                .Range("A" & grpHeaderRow & ",B" & grpHeaderRow & _
                       ",E" & grpHeaderRow & ",F" & grpHeaderRow).NumberFormat = "#,##0"
            End With
        End If

        ' 유형 소계행
        Call 소계행출력(wsResult, outRow, "소계: [" & tLabel & "]", tSumD, tSumC, tBalD, tBalC, RGB(220, 230, 241))
        outRow = outRow + 1

        totalSumD = totalSumD + tSumD: totalSumC = totalSumC + tSumC
        totalBalD = totalBalD + tBalD: totalBalC = totalBalC + tBalC

NextType:
    Next t

    ' 전체 합계행
    Call 소계행출력(wsResult, outRow, "합   계", totalSumD, totalSumC, totalBalD, totalBalC, RGB(180, 198, 231))

    ' ------------------------------------------
    ' 7. 숫자 서식 + 전체 테두리
    ' ------------------------------------------
    Dim dataEnd As Long
    dataEnd = outRow

    wsResult.Range("A5:B" & dataEnd).NumberFormat = "#,##0"
    wsResult.Range("E5:F" & dataEnd).NumberFormat = "#,##0"

    With wsResult.Range("A5:F" & dataEnd)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
      
    ' ===== 엑셀 원상복구 (반드시!!)  =====
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ' ===== 엑셀 원상복구 (반드시!!)  =====

    
    Dim tElapsed As Double       ''소요시간 체크하기 위해 코파일럿에게 물어서 시간변수 추가함, 프로시저 상단에도 tStart 선언
    tElapsed = Timer - tStart

    MsgBox "합계잔액시산표 생성 완료!" & vbCrLf & _
       "계정 수 : " & nMaster & "개" & vbCrLf & _
       "소요 시간 : " & Format(tElapsed, "0.00") & " 초", _
       vbInformation
 
    
End Sub


''Mod_ACCESS_DB에다가 옮겨놓음!!
''
'''===========================================
''' 소계/합계 행 출력 헬퍼
'''===========================================
''Private Sub 소계행출력(ws As Worksheet, r As Long, sLabel As String, _
''                       dSumD As Double, dSumC As Double, _
''                       dBalD As Double, dBalC As Double, _
''                       bgColor As Long)
''    With ws
''        If dBalD <> 0 Then .Cells(r, 1).Value = dBalD
''        If dSumD <> 0 Then .Cells(r, 2).Value = dSumD
''        .Cells(r, 4).Value = sLabel
''        If dSumC <> 0 Then .Cells(r, 5).Value = dSumC
''        If dBalC <> 0 Then .Cells(r, 6).Value = dBalC
''        .Range("A" & r & ":F" & r).Font.Bold = True
''        .Range("A" & r & ":F" & r).Interior.Color = bgColor
''        .Range("A" & r & ",B" & r & ",E" & r & ",F" & r).NumberFormat = "#,##0"
''    End With
''End Sub
''
''
'''===========================================
''' 계정 정렬 키
'''===========================================
''Private Function 계정정렬키(sType As String, nOrder As Long) As String
''    Dim nSeq As Integer
''    Select Case sType
''        Case "자산": nSeq = 1
''        Case "부채": nSeq = 2
''        Case "자본": nSeq = 3
''        Case "수익": nSeq = 4
''        Case "비용": nSeq = 5
''        Case Else:   nSeq = 9
''    End Select
''    계정정렬키 = Format(nSeq, "0") & Format(nOrder, "000")
''End Function
''