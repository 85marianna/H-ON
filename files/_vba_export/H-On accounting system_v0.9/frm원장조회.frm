
Private Sub btnDateFrom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txtFrom.Value = frmCalendar.GetDate

End Sub

Private Sub btnDateTo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txtTo.Value = frmCalendar.GetDate

End Sub

Private Sub cbo계정코드_Change()

Dim i As Long
Dim strInput As String

strInput = Me.cbo계정코드.Text ' 현재 입력한 글자 (숫자든 한글이든 상관없음)

' 1. 아무것도 없으면 이름 지우고 종료
If strInput = "" Then
    Me.txt계정명.Value = ""
    Exit Sub
End If

' 2. [통합 검색 구간] 목록을 돌면서 코드(0열)나 이름(1열)과 일치하는지 확인
For i = 0 To Me.cbo계정코드.ListCount - 1
    ' 조건: 입력한 값이 '계정코드'와 같거나, '계정명'과 정확히 같을 때
    If Me.cbo계정코드.List(i, 0) = strInput Or Me.cbo계정코드.List(i, 1) = strInput Then
        
        ' 찾았다면! 텍스트박스에는 항상 '이름'을 뿌려주고
        Me.txt계정명.Value = Me.cbo계정코드.List(i, 1)
        
        ' [중요] 만약 이름을 쳐서 찾았다면, 콤보박스 값을 해당 '코드'로 자동 변경!
        ' (그래야 나중에 시트에 저장할 때 코드로 깔끔하게 저장돼)
        If Me.cbo계정코드.List(i, 1) = strInput Then
            ' 이벤트를 잠시 피하기 위해 이 코드가 실행될 때 주의가 필요할 수 있지만,
            ' 단순 입력 시에는 아래 한 줄로 충분해.
            Me.cbo계정코드.Text = Me.cbo계정코드.List(i, 0)
        End If
        
        Exit Sub
    End If
Next i

' 3. 일치하는 게 없으면 이름을 비워둠
Me.txt계정명.Value = ""

End Sub


Private Sub cmd조회_Click()

   Dim sCode As String
   sCode = CStr(cbo계정코드.Value)

   '=====================
   ' 날짜 유효성 검사 (공통)
   '=====================
   If Not IsDate(txtFrom.Value) Then
       MsgBox "시작일을 선택해주세요.", , "원장 조회", vbExclamation
       txtFrom.SetFocus
       Exit Sub
   End If
   If Not IsDate(txtTo.Value) Then
       MsgBox "종료일을 선택해주세요.", , "원장 조회", vbExclamation
       txtTo.SetFocus
       Exit Sub
   End If
   If CDate(txtFrom.Value) > CDate(txtTo.Value) Then
       MsgBox "시작일이 종료일보다 클 수 없습니다.", , "원장 조회", vbExclamation
       Exit Sub
   End If

   '=====================
   ' 조회구분별 실행
   '=====================
   If opt총계정원장.Value Then
       Call 총계정원장생성_DB(CDate(txtFrom.Value), CDate(txtTo.Value))

   ElseIf opt계정명세.Value Then
       If CStr(cbo계정코드.Value) = "" Then
           MsgBox "계정코드를 선택해주세요.", , "계정명세", vbExclamation
           cbo계정코드.SetFocus
           Exit Sub
       End If
       Call 계정별명세출력_DB(CStr(cbo계정코드.Value), CDate(txtFrom.Value), CDate(txtTo.Value))

   ElseIf opt분개장.Value Then
       Call 분개장출력_DB(CDate(txtFrom.Value), CDate(txtTo.Value))

   End If

End Sub


Private Sub optLastMonth_Click()

    Me.txtFrom.Value = DateSerial(Year(Date), Month(Date) - 1, 1)
    Me.txtTo.Value = DateSerial(Year(Date), Month(Date), 0)

End Sub

Private Sub optThisMonth_Click()

    Me.txtFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    Me.txtTo.Value = DateSerial(Year(Date), Month(Date) + 1, 0)

End Sub

Private Sub optThisYear_Click()

    Me.txtFrom.Value = DateSerial(Year(Date), 1, 1)
    Me.txtTo.Value = DateSerial(Year(Date), 12, 31)

End Sub

Private Sub optToday_Click()

Me.txtFrom.Value = Date
Me.txtTo.Value = Date

End Sub

Private Sub opt계정명세_Click()

Me.cmd조회.Caption = "계정명세 조회"

End Sub

Private Sub opt분개장_Click()

Me.cbo계정코드.Value = ""
Me.cmd조회.Caption = "분개장 조회"

End Sub

Private Sub opt총계정원장_Click()

Me.cbo계정코드.Value = ""
Me.cmd조회.Caption = "총계정원장 조회"

End Sub

Private Sub UserForm_Initialize()

Update_cbo계정코드 Me.cbo계정코드

End Sub