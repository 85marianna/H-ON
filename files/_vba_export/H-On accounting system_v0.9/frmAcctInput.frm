Option Explicit

'"이건 사용자 입력이 아니라 코드 작업이니까 이벤트 무시해"는 스위치＝＞bUpdating = True → 이벤트 무시　／　bUpdating = False → 이벤트 실행
Dim bUpdating As Boolean
Public CreateReverse As Boolean  '반대분개 변수(현재는 외화평가 취소에 사용됨)

'
Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt전표일자.Value = frmCalendar.GetDate

End Sub
'
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

'
Private Sub cbo차대_Change()

If Me.cbo차대 = "차변" Then
   Me.txt대변 = ""
Else
   Me.txt차변 = ""
End If

CalculateAmountFx

End Sub

Private Sub cbo통화_Change()

If Me.cbo통화.Value <> "KRW" Then   '->외화전표
   lbl외화.Caption = "(외화 금액)"
   Me.txt차변.Locked = True
   Me.txt대변.Locked = True
   Call ToggleControl(Me.txt환율, True)
   Call ToggleControl(Me.cbo차대, True)
   Call ToggleControl(Me.txt외화금액, True)
   
   Me.chk즐겨찾기.Visible = False
   Me.cmd즐겨찾기.Visible = False
   Me.txt즐겨찾기.Visible = False

Else  '->원화전표
   Me.txt차변.Locked = False
   Me.txt대변.Locked = False
   Call ToggleControl(Me.txt환율, False)
   Call ToggleControl(Me.cbo차대, False)
   Call ToggleControl(Me.txt외화금액, False)
   
   Me.chk즐겨찾기.Visible = True
   Me.cmd즐겨찾기.Visible = True
   Me.txt즐겨찾기.Visible = True
   
End If

End Sub

Private Sub cmd닫기_Click()

Unload Me

End Sub
'
Private Sub cmd전표입력_Click()

유저폼_전표입력

End Sub


Private Sub chk즐겨찾기_Click()


    If Me.chk즐겨찾기.Value = True Then

        Me.txt즐겨찾기.Value = InputBox("즐겨찾기 이름을 입력하세요.", "자주 사용하는 분개 등록")

    Else

        Me.txt즐겨찾기.Value = ""

    End If


End Sub

Private Sub cmd전표저장_Click()

'''엑세스 DB용

    Dim Conn As Object
    Dim Sql As String
    Dim i As Integer
    Dim printNo As String
    Dim voucherNo As String
    
    Dim ReverseVoucherNo As String


        ' 검증 먼저
    If Me.txt전표일자.Value = "" Then MsgBox "전표날짜를 입력하세요.", vbExclamation, "주의!": Exit Sub
    If Me.cbo부서명.Value = "" Then MsgBox "부서를 선택하세요.", vbExclamation, "주의!": Exit Sub
    If lstMain.ListCount = 0 Then MsgBox "전표 내용을 입력하세요.", vbExclamation, "주의!": Exit Sub
    If Me.txt차변합계 <> Me.txt대변합계 Then MsgBox "차/대 금액이 일치하지 않습니다.", vbExclamation, "주의!": Exit Sub
    

    '  DB 연결(경로 공용함수)
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()


    ' 신규 전표번호 생성 (Access에서 MAX 조회하는 함수)
    voucherNo = Get_NextVoucherNo(Conn, DateValue(Me.txt전표일자.Value))

        For i = 0 To lstMain.ListCount - 1


''        '  테스트 데이터 INSERT
''        sql = "INSERT INTO 분개장 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 입력일시, 입력자) " & _
''              "VALUES (#2026-05-21#, '2026-05-21-001', 99, '테스트입력', 'AGC화인테크노한국㈜', 1000, 2000, '5212000', '지급수수료', '관리팀', 'KRW', #2026-05-01#, '홍길동')"


            '  분개장 INSERT
            Sql = "INSERT INTO 분개장 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 입력일시, 입력자)" & _
                  " VALUES (" & _
                  "#" & Format(Me.txt전표일자.Value, "yyyy-mm-dd") & "#, " & _
                  "'" & voucherNo & "', " & _
                  (i + 1) & ", " & _
                  "'" & lstMain.List(i, 3) & "', " & _
                  "'" & lstMain.List(i, 4) & "', " & _
                  CLng(Replace(lstMain.List(i, 5), ",", "")) & ", " & _
                  CLng(Replace(lstMain.List(i, 6), ",", "")) & ", " & _
                  "'" & lstMain.List(i, 1) & "', " & _
                  "'" & lstMain.List(i, 2) & "', " & _
                  "'" & Me.cbo부서명.Column(1) & "', " & _
                  "'" & Me.cbo통화.Value & "', " & _
                  Val(lstMain.List(i, 8) & "") & ", " & _
                  Val(lstMain.List(i, 9) & "") & ", " & _
                  "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, " & _
                  "'" & Application.UserName & "')"

            'Debug.Print sql
            Conn.Execute Sql


            '  채권채무 조건 처리
            If lstMain.List(i, 1) = "2110500" Then

                     ' 채권채무명세 Insert
                    Sql = "INSERT INTO 채권채무명세 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 상태)" & _
                          " VALUES (" & _
                          "#" & Format(Me.txt전표일자.Value, "yyyy-mm-dd") & "#, " & _
                          "'" & voucherNo & "', " & _
                          (i + 1) & ", " & _
                          "'" & lstMain.List(i, 3) & "', " & _
                          "'" & lstMain.List(i, 4) & "', " & _
                          CLng(Replace(lstMain.List(i, 5), ",", "")) & ", " & _
                          CLng(Replace(lstMain.List(i, 6), ",", "")) & ", " & _
                          "'" & lstMain.List(i, 1) & "', " & _
                          "'" & lstMain.List(i, 2) & "', " & _
                          "'" & Me.cbo부서명.Column(1) & "', " & _
                          "'" & Me.cbo통화.Value & "', " & _
                          Val(lstMain.List(i, 8) & "") & ", " & _
                          Val(lstMain.List(i, 9) & "") & ", " & _
                          "'채무')"


                Conn.Execute Sql

            ElseIf lstMain.List(i, 1) = "1110700" Then

                    Sql = "INSERT INTO 채권채무명세 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 상태)" & _
                          " VALUES (" & _
                          "#" & Format(Me.txt전표일자.Value, "yyyy-mm-dd") & "#, " & _
                          "'" & voucherNo & "', " & _
                          (i + 1) & ", " & _
                          "'" & lstMain.List(i, 3) & "', " & _
                          "'" & lstMain.List(i, 4) & "', " & _
                          CLng(Replace(lstMain.List(i, 5), ",", "")) & ", " & _
                          CLng(Replace(lstMain.List(i, 6), ",", "")) & ", " & _
                          "'" & lstMain.List(i, 1) & "', " & _
                          "'" & lstMain.List(i, 2) & "', " & _
                          "'" & Me.cbo부서명.Column(1) & "', " & _
                          "'" & Me.cbo통화.Value & "', " & _
                          Val(lstMain.List(i, 8) & "") & ", " & _
                          Val(lstMain.List(i, 9) & "") & ", " & _
                          "'채권')"
                Conn.Execute Sql
                
        End If

        Next i
        
        
        If Me.chk즐겨찾기.Value = True And Trim(Me.txt즐겨찾기.Value) <> "" Then

        Call 즐겨찾기저장_DB(Conn)

        End If
        

        MsgBox "전표가 DB에 저장되었습니다!" & vbCrLf & "전표번호 : " & voucherNo, vbOKOnly + vbInformation, "전표저장 완료"
               

        If CreateReverse Then  '반대분개 (현재는 외화평가 취소에만 사용됨)
        
            ReverseVoucherNo = CreateReverseVoucher(Conn, voucherNo, DateSerial(Year(Me.txt전표일자.Value), Month(Me.txt전표일자.Value) + 1, 1))

            MsgBox "반대분개 전표가 생성되었습니다!" & vbCrLf & "전표번호 : " & ReverseVoucherNo, vbInformation, "반대분개"

        End If
        

        Unload Me


    Conn.Close
    Set Conn = Nothing

End Sub


Sub 즐겨찾기저장_DB(Conn As Object)

    Dim i As Long
    Dim Sql As String

    For i = 0 To lstMain.ListCount - 1

        Sql = "INSERT INTO 즐겨찾기 (즐겨찾기명, 명세번호, 세목코드, 계정명, 적요, 거래처, 차변, 대변, 등록자) " & _
              "VALUES (" & _
              "'" & Me.txt즐겨찾기.Value & "', " & _
              (i + 1) & ", " & _
              "'" & lstMain.List(i, 1) & "', " & _
              "'" & lstMain.List(i, 2) & "', " & _
              "'" & lstMain.List(i, 3) & "', " & _
              "'" & lstMain.List(i, 4) & "', " & _
              CLng(Replace(lstMain.List(i, 5), ",", "")) & ", " & _
              CLng(Replace(lstMain.List(i, 6), ",", "")) & ", " & _
              "'" & Application.UserName & "')"

        Conn.Execute Sql

    Next i

End Sub


Private Sub cmd즐겨찾기_Click()

frm즐겨찾기.Show 1

End Sub

Private Sub cmd추가_Click()

'행추가와 행수정이 가능

If Me.cbo계정코드.Value = "" Then MsgBox "계정을 입력하세요.", vbExclamation: Exit Sub
If Me.txt적요.Value = "" Then MsgBox "전표 내용을 입력하세요.", vbExclamation: Exit Sub
If Me.txt계정명.Value = "" Then MsgBox "계정코드를 입력하세요.", vbExclamation: Exit Sub

'거래처코드가　필수인　계정
If (Me.cbo계정코드.Value = "2110500" Or Me.cbo계정코드.Value = "1110700") _
   And Me.txt거래처.Value = "" Then
    MsgBox "해당 계정은 거래처 입력이 필수입니다.", vbExclamation
    Exit Sub
End If


Dim r As Integer
r = Me.lstMain.ListIndex

If r <> -1 Then     '-----> 수정하기 위해 행을 선택했다면 행수정

    Me.lstMain.List(r, 1) = Me.cbo계정코드.Value
    Me.lstMain.List(r, 2) = Me.txt계정명.Value
    Me.lstMain.List(r, 3) = Me.txt적요.Value
    Me.lstMain.List(r, 4) = Me.txt거래처.Value
    Me.lstMain.List(r, 5) = Me.txt차변.Value
    Me.lstMain.List(r, 6) = Me.txt대변.Value
    Me.lstMain.List(r, 7) = Me.cbo통화.Value
    Me.lstMain.List(r, 8) = Me.txt외화금액.Value
    Me.lstMain.List(r, 9) = Me.txt환율.Value

    Call 차대합계_일치_Public(Me)

    Me.lstMain.ListIndex = -1   '선택 해제
    MsgBox r + 1 & "행의 내용을 수정했습니다.", vbExclamation

Else
    ' 새 행 추가
    lstMain.AddItem

    ' 순번 넣기 (첫 번째 열)  ---> 중간행 수정/삭제로 인해 순번넣기는 해제함.
    'lstMain.List(lstMain.ListCount - 1, 0) = lstMain.ListCount    'ListBox.List(RowIndex, ColumnIndex)

    '리스트박스에 분개내역 추가
    lstMain.List(lstMain.ListCount - 1, 1) = Me.cbo계정코드.Value
    lstMain.List(lstMain.ListCount - 1, 2) = Me.txt계정명.Value
    lstMain.List(lstMain.ListCount - 1, 3) = Me.txt적요.Value
    lstMain.List(lstMain.ListCount - 1, 4) = Me.txt거래처.Value
    lstMain.List(lstMain.ListCount - 1, 5) = Me.txt차변.Value
    lstMain.List(lstMain.ListCount - 1, 6) = Me.txt대변.Value
    lstMain.List(lstMain.ListCount - 1, 7) = Me.cbo통화.Value
    lstMain.List(lstMain.ListCount - 1, 8) = Me.txt외화금액.Value
    lstMain.List(lstMain.ListCount - 1, 9) = Me.txt환율.Value

    Call 차대합계_일치_Public(Me)

End If


bUpdating = True   ' 이벤트 OFF

'입력 컨트롤 지우기
Me.cbo계정코드.Value = ""
Me.txt계정명.Value = ""
Me.txt차변.Value = ""
Me.txt대변.Value = ""
Me.txt외화금액.Value = ""

bUpdating = False  ' 이벤트 ON


' 이벤트 끝난 뒤 최종 상태 강제 지정
ToggleControl Me.txt차변, True
ToggleControl Me.txt대변, True


Call ToggleControl(Me.cbo통화, False)

End Sub


Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

frm거래처검색.Show 1

End Sub


Private Sub lstMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim answer As VbMsgBoxResult
Dim r As Integer
Dim accCode As Long
Dim i As Integer

r = Me.lstMain.ListIndex
accCode = lstMain.List(r, 1)


answer = MsgBox("선택한 행을 삭제하시려면 [예], 수정하시려면 [아니오]를 눌러주세요.", vbYesNo + vbInformation, "명세 삭제/수정")

    If answer = vbYes Then
        '행 삭제
        Me.lstMain.RemoveItem r
        Me.lstMain.ListIndex = -1
        차대합계_일치_Public Me

        If Me.lstMain.ListCount = 0 Then
            Call ToggleControl(Me.cbo통화, True)
        End If
    Exit Sub

    ElseIf answer = vbNo Then
        '행 수정을 위해 입력컨트롤에 기존 데이터 표시
        
        bUpdating = True   '  이벤트 차단
        
        For i = 0 To Me.cbo계정코드.ListCount - 1
            If Me.cbo계정코드.List(i, 0) = accCode Then   ' 첫 번째 열에서 일치코드 행 찾기
                Me.cbo계정코드.ListIndex = i             ' 해당 행 선택
                Exit For
            End If
        Next i

        Me.txt적요.Value = lstMain.List(r, 3)
        Me.txt거래처.Value = lstMain.List(r, 4)
        Me.txt차변.Value = lstMain.List(r, 5)
        Me.txt대변.Value = lstMain.List(r, 6)
        
        bUpdating = False  '  이벤트 다시 허용
        
        
        '  상태를 마지막에 강제 지정 (핵심)
        If Val(Me.txt차변.Value) > 0 Then
            ToggleControl Me.txt차변, True
            ToggleControl Me.txt대변, False
    
        ElseIf Val(Me.txt대변.Value) > 0 Then
            ToggleControl Me.txt차변, False
            ToggleControl Me.txt대변, True
    
        Else
            ToggleControl Me.txt차변, True
            ToggleControl Me.txt대변, True
        End If
        
        Exit Sub
    End If

End Sub


Private Sub txt대변_AfterUpdate()

   Me.lbl메세지.Caption = ""
   Me.txt대변.BackColor = vbWhite

End Sub

Private Sub txt대변_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    ' 공백이면 그냥 나가도록 허용
    If Me.txt대변.Value = "" Then
        Cancel = False
        Exit Sub
    End If

    ' 숫자가 아니면 막기
    If Not IsNumeric(Me.txt대변.Value) Then
        Me.lbl메세지.Caption = "금액은 숫자로만 입력해주세요."
        Me.txt대변.BackColor = RGB(255, 224, 224)
        Me.txt대변.SelStart = 0
        Me.txt대변.SelLength = Len(Me.txt대변.Value)
        Cancel = True

'    ' 음수면 막기
'    ElseIf Me.txt대변.Value < 0 Then
'        Me.lbl메세지.Caption = "금액은 양수로만 입력해주세요."
'        Me.txt대변.BackColor = RGB(255, 224, 224)
'        Me.txt대변.SelStart = 0
'        Me.txt대변.SelLength = Len(Me.txt대변.Value)
'        Cancel = True
    End If

End Sub


Private Sub txt대변_Change()

    If bUpdating Then Exit Sub
    bUpdating = True

    Dim s As String
    Dim v As Currency

    s = Me.txt대변.Text

    If IsNumeric(Replace(s, ",", "")) And s <> "" Then
        v = CDbl(Replace(s, ",", ""))
        Me.txt대변.Text = Format(v, "#,##0")
        ' 차변 비활성화
        Me.txt차변.Text = "0"
        ToggleControl Me.txt차변, False
    ElseIf s = "" Then
        ' 대변 지우면 차변 다시 활성화
        Me.txt차변.Text = ""
        ToggleControl Me.txt차변, True
    End If

    bUpdating = False

End Sub


Private Sub txt외화금액_Change()
   
    Call CalculateAmountFx

End Sub

Private Sub txt전표일자_Change()

If Len(Me.txt전표일자.Value) = 0 Then Exit Sub  '----> 입력컨트롤을 지웠을 때, 메세지박스 중복 발생을 방지.

If Not Me.txt전표일자.Value Like "####-##-##" Then
    MsgBox "날짜 형식이 아닙니다. 올바른 날짜를 입력하세요.", vbCritical, "입력 오류"
    Me.txt전표일자.Value = ""
    Exit Sub

End If

End Sub


Private Sub txt차변_AfterUpdate()

   Me.lbl메세지.Caption = ""
   Me.txt차변.BackColor = vbWhite

End Sub

Private Sub txt차변_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    ' 공백이면 그냥 나가도록 허용
    If Me.txt차변.Value = "" Then
        Cancel = False
        Exit Sub
    End If

    ' 숫자가 아니면 막기
    If Not IsNumeric(Me.txt차변.Value) Then
        Me.lbl메세지.Caption = "금액은 숫자로만 입력해주세요."
        Me.txt차변.BackColor = RGB(255, 224, 224)
        Me.txt차변.SelStart = 0
        Me.txt차변.SelLength = Len(Me.txt차변.Value)
        Cancel = True
'
'    ' 음수면 막기
'    ElseIf Me.txt차변.Value < 0 Then
'        Me.lbl메세지.Caption = "금액은 양수로만 입력해주세요."
'        Me.txt차변.BackColor = RGB(255, 224, 224)
'        Me.txt차변.SelStart = 0
'        Me.txt차변.SelLength = Len(Me.txt차변.Value)
'        Cancel = True
    End If

End Sub


Private Sub txt차변_Change()


    If bUpdating Then Exit Sub
    bUpdating = True

    Dim s As String
    Dim v As Currency

    s = Me.txt차변.Text

    If IsNumeric(Replace(s, ",", "")) And s <> "" Then
        v = CDbl(Replace(s, ",", ""))
        Me.txt차변.Text = Format(v, "#,##0")
        ' 대변 비활성화
        Me.txt대변.Text = "0"
        ToggleControl Me.txt대변, False
    ElseIf s = "" Then
        ' 차변 지우면 대변 다시 활성화
        Me.txt대변.Text = ""
        ToggleControl Me.txt대변, True
    End If

    bUpdating = False


End Sub


Private Sub txt환율_Change()

Call CalculateAmountFx

End Sub

Private Sub UserForm_Initialize()

Update_cbo부서명
Update_cbo계정코드 Me.cbo계정코드  'frm공용함수 호출

Me.cbo통화.Value = "KRW"

With Me.cbo차대
    .AddItem "차변"
    .AddItem "대변"
End With

With Me.cbo통화
    .AddItem "KRW"
    .AddItem "JPY"
    .AddItem "USD"
End With

Me.txt전표일자.SetFocus

End Sub

Sub Update_cbo부서명()

Dim db As Variant
db = Get_DB_Access("부서마스터")

Update_Cbo Me.cbo부서명, db, 2
''2열을 표시

End Sub


Private Sub CalculateAmountFx()
   On Error Resume Next ' 입력 중 오류 방지
   ' 외화금액 * 환율 -> 원화 금액칸에 입력
   If Me.cbo차대.Value = "차변" Then
       txt차변.Value = Format(Val(txt외화금액.Value) * Val(txt환율.Value), "#,##0")
   ElseIf Me.cbo차대.Value = "대변" Then
       txt대변.Value = Format(Val(txt외화금액.Value) * Val(txt환율.Value), "#,##0")
   End If
End Sub
