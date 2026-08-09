Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt조회일자.Value = frmCalendar.GetDate

End Sub

Private Sub cmd닫기_Click()

Unload Me

End Sub

Private Sub cmd반제_Click()

Dim i As Long


' 1. ★ 입구 컷: 리스트박스에 줄이 하나도 없으면 메시지 띄우고 종료! ★
    If Me.lstMain.ListCount = 0 Then
        MsgBox "반제 처리할 명세가 없습니다.", vbExclamation, "알림"
        Exit Sub
    End If


ReDim arrIDs(0 To Me.lstMain.ListCount - 1)

For i = 0 To Me.lstMain.ListCount - 1
    arrIDs(i) = Me.lstMain.List(i, 0)
Next i

' ★ 핵심: 입력폼의 날짜 칸에 현재 조회일자를 미리 넣어주기 ★
' 폼을 열기 전에 값을 먼저 전달해!
frm반제전표.txt전표일자.Value = Me.txt조회일자.Value
' 날짜를 마음대로 못 바꾸게 잠그기 (선택사항이지만 혜정이 의도라면 추천!)
frm반제전표.txt전표일자.Locked = True
frm반제전표.btnDate.Enabled = False

    
Call 반제전표(frm채권채무조회)

Call 차대합계_일치_Public(frm반제전표)

'반제전표입력폼이 열려있으면 다시 안열어도 돼!!
If frm반제전표.Visible = False Then
    frm반제전표.Show 1    ''->반제전표입력폼 띄우기
    
End If

Unload Me

End Sub

Private Sub cmd인쇄_Click()

Dim ws As Worksheet
Dim newWb As Workbook
Dim i As Long, j As Long
Dim colCount As Integer
Dim lastCol As Integer
Dim optName As String

' ---------------------------------------------------------
' ★ [추가] 채권/채무 선택 여부 확인 ★
' ---------------------------------------------------------
' opt채권과 opt채무가 둘 다 False(선택 안 됨)라면 메시지 출력
If Me.opt채권.Value = False And Me.opt채무.Value = False Then
    MsgBox "채권 또는 채무를 먼저 선택해 주세요.", vbExclamation, "선택 확인"
    Exit Sub ' 코드 실행 중단
End If
' ---------------------------------------------------------

' 1. 분류 텍스트 설정
If Me.opt채권.Value = True Then
    optName = "채권"
Else
    optName = "채무"
End If


' 1. 리스트박스 데이터 확인
If Me.lstMain.ListCount = 0 Then
    MsgBox "인쇄할 내역이 없습니다.", vbInformation
    Exit Sub
End If

Application.ScreenUpdating = False


Set newWb = Workbooks.Add
Set ws = newWb.Worksheets(1)

ws.Name = "채권채무내역"



' 4. 리스트박스 값을 통채로 시트에 넣기
colCount = Me.lstMain.ColumnCount
For i = 0 To Me.lstMain.ListCount - 1
    For j = 0 To colCount - 1
        ws.Cells(i + 5, j + 1).Value = Me.lstMain.List(i, j)
    Next j
Next i


 '5. 불필요한 열 삭제 (B~N열만 남기기 위해 A, O~Q 삭제)
ws.Columns("O:Q").Delete
ws.Columns("A:A").Delete


' ---------------------------------------------------------
' ★ [추가] 거래처(E열) 기준 정렬 - 소계 넣기 전 필수 ★
' ---------------------------------------------------------
Dim lastDataRow As Long
lastDataRow = 4 + Me.lstMain.ListCount  ' 데이터가 5행부터 시작하므로

If lastDataRow >= 5 Then
    ws.Range(ws.Cells(5, 1), ws.Cells(lastDataRow, 13)).Sort _
        Key1:=ws.Range(ws.Cells(5, 5), ws.Cells(lastDataRow, 5)), _
        Order1:=xlAscending, Header:=xlNo
End If


' ---------------------------------------------------------
' ★ 제목 및 조회 조건 설정 ★
' ---------------------------------------------------------

' ★ 헤더 먼저 강제 입력
ws.Range("A4:M4").Value = Array( _
    "회계일자", "전표번호", "명세번호", "적요", "거래처", _
    "차변", "대변", "세목코드", "계정명", "부서명", _
    "통화", "외화금액", "환율")

lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column ' 현재 데이터 열 개수 (M이나 N쯤 되겠지?)


'  헤더 색상 변경 (이제 헤더는 4행에 있음)
With ws.Range(ws.Cells(4, 1), ws.Cells(4, lastCol))
    .Interior.Color = RGB(217, 217, 217)
End With


' 제목 [채권채무 대장현황]
With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    .Merge ' 셀 병합
    .Value = "【 채 권 채 무 대 장 현 황 】"
    .Font.Size = 16
    .Font.Bold = True
    '.HorizontalAlignment = xlCenter
    ' ★ xlCenter 대신 숫자 -4108을 써봐. 이게 제일 확실해!
    .HorizontalAlignment = -4108 ' xlCenter의 실제 값
End With

' 조회 조건 줄
' 발생일자는 유저폼의 값을 참조하거나 상황에 맞게 수정해줘!
Dim infoText As String
infoText = "■ 채권채무분류 : " & optName & "     ■ 발생일 : " & "     ■ 조회일 : " & Format(Me.txt조회일자.Value, "yyyy-mm-dd")

With ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol))
    .Merge
    .Value = infoText
    .HorizontalAlignment = xlLeft ' 조건은 보통 왼쪽 정렬이 보기 좋아
End With


' ---------------------------------------------------------
' ★ [추가] 거래처별 소계 삽입 (아래→위 방향으로 삽입해야 인덱스 안 꼬임) ★
' ---------------------------------------------------------
Dim vendorCol As Integer, debitCol As Integer, creditCol As Integer
vendorCol = 5 ' 거래처(E)
debitCol = 6  ' 차변(F)
creditCol = 7 ' 대변(G)

Dim curRow As Long, groupStart As Long
Dim sumDebit As Double, sumCredit As Double

curRow = lastDataRow
Do While curRow >= 5
    groupStart = curRow
    Do While groupStart > 5 And ws.Cells(groupStart - 1, vendorCol).Value = ws.Cells(curRow, vendorCol).Value
        groupStart = groupStart - 1
    Loop

    sumDebit = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(groupStart, debitCol), ws.Cells(curRow, debitCol)))
    sumCredit = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(groupStart, creditCol), ws.Cells(curRow, creditCol)))

    ws.Rows(curRow + 1).Insert Shift:=xlDown

    With ws.Cells(curRow + 1, vendorCol)
        .Value = ws.Cells(curRow, vendorCol).Value & " 소계"
    End With
    With ws.Cells(curRow + 1, debitCol)
        .Value = sumDebit
        .NumberFormatLocal = "#,##0"
    End With
    With ws.Cells(curRow + 1, creditCol)
        .Value = sumCredit
        .NumberFormatLocal = "#,##0"
    End With
    With ws.Range(ws.Cells(curRow + 1, 1), ws.Cells(curRow + 1, lastCol))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
    End With

    curRow = groupStart - 1
Loop


' ---------------------------------------------------------
' ★ [수정] 마지막 줄에 총계 금액 넣기 (병합 없이) ★
' ---------------------------------------------------------
Dim totalRow As Long
totalRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row + 1

' 1. '거래처' 열(5번째 열로 가정)에 "총  계"라고 넣기
' 만약 열 위치가 다르면 숫자를 조정해줘! (E열이면 5)
With ws.Cells(totalRow, 5)
    .Value = "총      계"
    .HorizontalAlignment = -4108 ' 가운데 정렬
End With

' 2. 총 금액 넣기 (유저폼의 txt총금액 값을 G열에 입력)
' 금액(대변)이 7번째 열이 맞는지 확인!
With ws.Cells(totalRow, 7)
    .Value = Me.txt총금액.Value
    .NumberFormatLocal = "#,##0"
    .HorizontalAlignment = -4152 ' 오른쪽 정렬
End With

' 3. 총계 라인 전체 배경색 및 굵게 설정
With ws.Range(ws.Cells(totalRow, 1), ws.Cells(totalRow, lastCol))
    .Interior.Color = RGB(217, 217, 217)
    .Font.Bold = True
End With
' ---------------------------------------------------------


' 8. 마무리 정리 (테두리는 데이터가 있는 4행부터 'totalRow'까지 적용)
' 범위를 totalRow까지 잡아주면 합계줄까지 테두리가 깔끔하게 그려져!
With ws.Range(ws.Cells(4, 1), ws.Cells(totalRow, lastCol))
    .Borders.LineStyle = 1 ' xlContinuous 대신 1 사용
End With

ws.Columns.AutoFit

Application.ScreenUpdating = True


Unload Me


End Sub


Private Sub opt채권_Click()

Call 조회_실행

End Sub


Private Sub opt채무_Click()

Call 조회_실행

End Sub

Private Sub 리스트박스_총금액()

Dim i As Long
Dim total As Double
Dim rowCount As Integer

rowCount = Me.lstMain.ListCount
total = 0

If opt채권 = True Then
' 금액이 2번째 열(인덱스 1)에 있다고 가정
    For i = 0 To rowCount - 1
        total = total + Me.lstMain.List(i, 6)    '--> CDbl(Me.lstMain.List(i, 6)) : 따라서 CDbl(값) → 그 값을 숫자형(Double) 으
    Next i

ElseIf opt채무 = True Then
' 금액이 2번째 열(인덱스 1)에 있다고 가정
    For i = 0 To rowCount - 1
        total = total + Me.lstMain.List(i, 7)   '--> CDbl(Me.lstMain.List(i, 7))
    Next i

End If
    
' TextBox에 합계 표시
Me.txt총금액.Value = Format(total, "#,##0")
Me.txt총건수.Value = Me.lstMain.ListCount

End Sub


Private Sub txt조회일자_Change()

Call 조회_실행

End Sub

Private Sub UserForm_Initialize()

txt조회일자.Value = Format(Date, "yyyy-mm-dd")

End Sub


Sub 조회_실행()
    Dim i As Integer
    Dim db As Variant
    Dim filterType As String
    Dim colIdx As Integer
    
    ' 1. 현재 어떤 옵션이 선택되어 있는지 확인
    If Me.opt채권.Value = True Then
        filterType = "채권"
        colIdx = 7 ' 차변 열
    ElseIf Me.opt채무.Value = True Then
        filterType = "채무"
        colIdx = 8 ' 대변 열
    Else
        Exit Sub ' 아무것도 선택 안 됐으면 종료
    End If
    
    ' 2. 데이터 가져오기 및 1차 필터링(채권/채무)
    db = Get_DB_Access("채권채무명세")
    db = Filtered_DB(db, filterType, "15", True)
    
    ' 3. 날짜 필터링 (txt조회일자 기준)
    If Not IsEmpty(db) Then
        db = Filtered_DB(db, "<=" & Me.txt조회일자.Value, "2", True)
        
        ' 4. 금액 포맷팅 (콤마 찍기)
        If Not IsEmpty(db) Then
            For i = LBound(db, 1) To UBound(db, 1)
                db(i, colIdx) = Format(db(i, colIdx), "#,##0")
                db(i, colIdx) = Format(db(i, colIdx), "@@@@@@@@@@@@@@@@@")
            Next i
        End If
    End If
    
    ' 5. 리스트박스 업데이트
    Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;150pt;110pt;110pt;0pt;0pt;0pt;30pt;0pt;0pt;0pt;0pt;0pt;"
    Call 리스트박스_총금액
End Sub


'Private Sub lstMain_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstMain
'End Sub
'