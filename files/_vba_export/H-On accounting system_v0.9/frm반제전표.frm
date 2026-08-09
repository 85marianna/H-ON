Option Explicit

'"이건 사용자 입력이 아니라 코드 작업이니까 이벤트 무시해"는 스위치＝＞bUpdating = True → 이벤트 무시　／　bUpdating = False → 이벤트 실행
Dim bUpdating As Boolean


Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt전표일자.Value = frmCalendar.GetDate

End Sub

Private Sub cbo계좌코드_Change()

If Me.cbo계좌코드.Value = "" Then
    Me.txt계정명.Value = ""
    Exit Sub
    Else
    Me.txt계정명.Value = Me.cbo계좌코드.Column(1)
    Me.txt계좌정보.Value = Me.cbo계좌코드.Column(2) & " | " & Me.cbo계좌코드.Column(4)
    Me.txt거래처.Value = "외환은행(구미)"

End If

End Sub

Private Sub cmd닫기_Click()

Unload Me

End Sub

Private Sub cbo통화_Change()

If Me.cbo통화.Value <> "KRW" Then   '->외화전표
   lbl외화.Caption = "(외화 금액)"
   Me.txt차변.Locked = True
   Me.txt대변.Locked = True
   Call ToggleControl(Me.txt환율, True)
   Call ToggleControl(Me.cbo차대, True)
   Call ToggleControl(Me.txt외화금액, True)

Else  '->원화전표
   Me.txt차변.Locked = False
   Me.txt대변.Locked = False
   Call ToggleControl(Me.txt환율, False)
   Call ToggleControl(Me.cbo차대, False)
   Call ToggleControl(Me.txt외화금액, False)

End If

End Sub

Private Sub cmd전표입력_Click()

유저폼_반제전표입력

End Sub


Private Sub cmd채권채무검색_Click()

frm채권채무조회.Show 1

End Sub

Private Sub cmd추가_Click()

    ' 새 행 추가
    lstMain.AddItem

    ' 순번 넣기 (첫 번째 열)  ---> 중간행 수정/삭제로 인해 순번넣기는 해제함.
    'lstMain.List(lstMain.ListCount - 1, 0) = lstMain.ListCount    'ListBox.List(RowIndex, ColumnIndex)

    '리스트박스에 분개내역 추가
    lstMain.List(lstMain.ListCount - 1, 1) = Me.cbo계좌코드.Value
    lstMain.List(lstMain.ListCount - 1, 2) = Me.txt계정명.Value
    lstMain.List(lstMain.ListCount - 1, 3) = Me.txt적요.Value
    lstMain.List(lstMain.ListCount - 1, 4) = Me.txt거래처.Value
    lstMain.List(lstMain.ListCount - 1, 5) = Me.txt차변.Value
    lstMain.List(lstMain.ListCount - 1, 6) = Me.txt대변.Value
    lstMain.List(lstMain.ListCount - 1, 7) = Me.cbo통화.Value
    lstMain.List(lstMain.ListCount - 1, 8) = Me.txt외화금액.Value
    lstMain.List(lstMain.ListCount - 1, 9) = Me.txt환율.Value

    Call 차대합계_일치_Public(Me)

    '입력 컨트롤 지우기
    Me.cbo계좌코드.Value = ""
    Me.txt계좌정보.Value = ""
    Me.txt계정명.Value = ""
    Me.txt적요.Value = ""
    Me.txt거래처.Value = ""
    Me.txt차변.Value = ""
    Me.txt대변.Value = ""
    Me.txt외화금액.Value = ""


End Sub


Private Sub lstMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim answer As VbMsgBoxResult
Dim r As Integer
Dim accCode As Long
Dim i As Integer

r = Me.lstMain.ListIndex
accCode = lstMain.List(r, 1)


answer = MsgBox("선택한 행을 삭제하시겠습니까? ", vbYesNo + vbInformation, "명세 삭제")

    If answer = vbYes Then
        '행 삭제
        Me.lstMain.RemoveItem r
        Me.lstMain.ListIndex = -1  ' 리스트박스에서 선택된 줄을 해제해라
        차대합계_일치_Public Me

        If Me.lstMain.ListCount = 0 Then
            Call ToggleControl(Me.cbo통화, True)
        End If
    Exit Sub

    ElseIf answer = vbNo Then
        Exit Sub
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


Private Sub txt전표일자_Change()

If Len(Me.txt전표일자.Value) = 0 Then Exit Sub

If Not Me.txt전표일자.Value Like "####-##-##" Then
    MsgBox "날짜 형식이 아닙니다. 올바른 날짜를 입력하세요.", vbCritical, "입력 오류"
    Me.txt전표일자.Value = ""
    Exit Sub

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


Private Sub UserForm_Initialize()

Update_cbo계좌코드

With Me.cbo차대
    .AddItem "차변"
    .AddItem "대변"
End With

With Me.cbo통화
    .AddItem "JPY"
    .AddItem "USD"
End With

Me.txt전표일자.SetFocus
'
End Sub


Sub Update_cbo계좌코드()

Dim db As Variant
db = Get_DB_Access("제예금마스터")

Update계정코드검색 Me.cbo계좌코드, db

End Sub


''Sub 유저폼_반제전표입력()
''
''Dim db As Variant
''Dim ws As Worksheet
''Dim i As Integer
''Dim printNo As String
''
''Set ws = ThisWorkbook.Sheets("분개장")
''
''With ws
''printNo = Format(Get_MaxPrintNo(ws), "000")
''End With
''
''If Me.txt전표일자.Value = "" Then MsgBox "전표날짜를 입력하세요.": Exit Sub
''
''
''If Me.txt차변합계 = Me.txt대변합계 Then
''
''    For i = 0 To lstMain.ListCount - 1
''
''        Insert_Record sht분개장, _
''            DateValue(Me.txt전표일자.Value), _
''            Me.txt전표일자.Value & "-" & printNo, _
''            i + 1, _
''            lstMain.List(i, 3), _
''            lstMain.List(i, 4), _
''            lstMain.List(i, 5), _
''            lstMain.List(i, 6), _
''            lstMain.List(i, 1), _
''            lstMain.List(i, 2), _
''            "회계팀", _
''            lstMain.List(i, 7), _
''            lstMain.List(i, 8), _
''            lstMain.List(i, 9), _
''            Now(), _
''            Application.UserName
''
''
''            If lstMain.List(i, 1) = "2110500" Then    '---> 미지급금이면 거래처명세에 추가 입력(거래처명세 관리할 계정 추가하기!!)
''
''                Insert_Record sht채권채무명세, _
''                Me.txt전표일자.Value, _
''                Me.txt전표일자.Value & "-" & printNo, _
''                i + 1, _
''                lstMain.List(i, 3), _
''                lstMain.List(i, 4), _
''                lstMain.List(i, 5), _
''                lstMain.List(i, 6), _
''                lstMain.List(i, 1), _
''                lstMain.List(i, 2), _
''                "회계팀", _
''                lstMain.List(i, 7), _
''                lstMain.List(i, 8), _
''                lstMain.List(i, 9), _
''                "채무반제", _
''                Me.txt전표일자.Value & "-" & printNo
''
''            ElseIf lstMain.List(i, 1) = "1110700" Then    '---> 미수금이면 거래처명세에 추가 입력(거래처명세 관리할 계정 추가하기!!)
''
''                Insert_Record sht채권채무명세, _
''                Me.txt전표일자.Value, _
''                Me.txt전표일자.Value & "-" & printNo, _
''                i + 1, _
''                lstMain.List(i, 3), _
''                lstMain.List(i, 4), _
''                lstMain.List(i, 5), _
''                lstMain.List(i, 6), _
''                lstMain.List(i, 1), _
''                lstMain.List(i, 2), _
''                "회계팀", _
''                lstMain.List(i, 7), _
''                lstMain.List(i, 8), _
''                lstMain.List(i, 9), _
''                "채권반제", _
''                Me.txt전표일자.Value & "-" & printNo
''
''            End If
''
''
''    Next i
''
''
''     ' ========================================================
''    ' ★ 바로 여기! 예금계정 빼고 진짜 ID가 있는 명세만 골라 담기 ★　－＞변수　새로　선언　validCount
''    'Call 반제_업데이트를 실행하기 직전에 리스트박스에 남은 데이터만 다시 담기 -> 반제 전표 입력 시에 명세를 삭제하는 경우 대비하여 배열을 새로 갱신!
''    ' ========================================================
''' ========================================================
''    ' 1. 진짜 반제 처리할 명세(ID가 있는 줄)가 몇 개인지 먼저 세기
''    ' ========================================================
''    Dim validCount As Long
''    Dim j As Long
''
''    validCount = 0
''    If Me.lstMain.ListCount > 0 Then
''        For j = 0 To Me.lstMain.ListCount - 1
''            ' ID(0번 열)가 비어있지 않은 것만 카운트
''            If Trim(Me.lstMain.List(j, 0)) <> "" Then
''                validCount = validCount + 1
''            End If
''        Next j
''    End If
''
''    ' ========================================================
''    ' 2. ★ 입구 컷! ★ 반제할 명세가 없으면 여기서 중단
''    ' ========================================================
''    If validCount = 0 Then
''        MsgBox "반제 처리할 명세가 없으므로 전표를 입력할 수 없습니다.", vbExclamation, "입력 불가"
''        Unload Me  ' 폼 닫기
''        Exit Sub   ' 매크로 종료
''    End If
''
''    ' ========================================================
''    ' 3. 알맹이가 있을 때만 배열 만들고 담기
''    ' ========================================================
''    ReDim arrIDs(0 To validCount - 1)
''    Dim k As Long: k = 0
''
''    For j = 0 To Me.lstMain.ListCount - 1
''        If Trim(Me.lstMain.List(j, 0)) <> "" Then
''            arrIDs(k) = Me.lstMain.List(j, 0)
''            k = k + 1
''        End If
''    Next j
''
''
''Call 반제_업데이트(arrIDs, Me.txt전표일자.Value, printNo)
''
''Unload Me
''
''ThisWorkbook.Save
''
''MsgBox "전표가 입력되었습니다. 전표번호를 확인한 후 인쇄해주세요.", vbInformation
''
''
''Else
''    MsgBox "차/대 금액이 일치하지 않습니다. 금액을 확인해주세요.", vbInformation
''    Exit Sub
''End If
''
''
''End Sub


Private Sub 유저폼_반제전표입력()

    Dim Conn As Object
    Dim Sql As String
    Dim i As Integer
    Dim voucherNo As String
    Dim j As Long, k As Long
    Dim validCount As Long
    Dim arrIDs() As Variant

    ' ===== 검증 =====
    If Me.txt전표일자.Value = "" Then MsgBox "전표날짜를 입력하세요.", vbExclamation, "주의!": Exit Sub
    If Me.txt차변합계 <> Me.txt대변합계 Then
        MsgBox "차/대 금액이 일치하지 않습니다. 금액을 확인해주세요.", vbInformation
        Exit Sub
    End If

    ' ===== 반제 처리할 유효 명세(ID 있는 줄)만 먼저 카운트 =====
    validCount = 0
    If Me.lstMain.ListCount > 0 Then
        For j = 0 To Me.lstMain.ListCount - 1
            If Trim(Me.lstMain.List(j, 0)) <> "" Then
                validCount = validCount + 1
            End If
        Next j
    End If

    ' ===== 입구 컷! 반제할 명세가 없으면 여기서 중단 =====
    If validCount = 0 Then
        MsgBox "반제 처리할 명세가 없으므로 전표를 입력할 수 없습니다.", vbExclamation, "입력 불가"
        Exit Sub
    End If

    ' ===== 유효한 ID만 배열에 담기 =====
    ReDim arrIDs(0 To validCount - 1)
    k = 0
    For j = 0 To Me.lstMain.ListCount - 1
        If Trim(Me.lstMain.List(j, 0)) <> "" Then
            arrIDs(k) = Me.lstMain.List(j, 0)
            k = k + 1
        End If
    Next j


    ' ===== 지금 리스트박스에 몇 줄 있는지 먼저 확인 (디버그용) =====
    Debug.Print "저장 시점 lstMain 줄 수: " & lstMain.ListCount
    For i = 0 To lstMain.ListCount - 1
        Debug.Print i & " / ID:" & lstMain.List(i, 0) & " / 계정:" & lstMain.List(i, 2) & _
                     " / 차변:" & lstMain.List(i, 5) & " / 대변:" & lstMain.List(i, 6)
    Next i



    ' ===== DB 연결 =====
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    ' ===== 신규 전표번호 채번 =====
    voucherNo = Get_NextVoucherNo(Conn, DateValue(Me.txt전표일자.Value))

    ' ===== 분개장 INSERT + 채권채무명세 신규 등록 =====
    For i = 0 To lstMain.ListCount - 1

        ' 분개장 INSERT
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
              "'회계팀', " & _
              "'" & lstMain.List(i, 7) & "', " & _
              Val(lstMain.List(i, 8) & "") & ", " & _
              Val(lstMain.List(i, 9) & "") & ", " & _
              "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, " & _
              "'" & Application.UserName & "')"

        Debug.Print Sql
        Conn.Execute Sql

        ' 상대계정이 미지급금(2110500)이면 → 채권채무명세에 "채무반제" 상태로 신규 등록
        If lstMain.List(i, 1) = "2110500" Then

            Sql = "INSERT INTO 채권채무명세 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 상태, 반제전표)" & _
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
                  "'회계팀', " & _
                  "'" & lstMain.List(i, 7) & "', " & _
                  Val(lstMain.List(i, 8) & "") & ", " & _
                  Val(lstMain.List(i, 9) & "") & ", " & _
                  "'채무반제', " & _
                  "'" & voucherNo & "')"

            Debug.Print Sql
            Conn.Execute Sql

        ' 상대계정이 미수금(1110700)이면 → 채권채무명세에 "채권반제" 상태로 신규 등록
        ElseIf lstMain.List(i, 1) = "1110700" Then

            Sql = "INSERT INTO 채권채무명세 (회계일자, 전표번호, 명세번호, 적요, 거래처, 차변, 대변, 세목코드, 계정명, 부서명, 통화, 외화금액, 환율, 상태, 반제전표)" & _
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
                  "'회계팀', " & _
                  "'" & Me.cbo통화.Value & "', " & _
                  Val(lstMain.List(i, 8) & "") & ", " & _
                  Val(lstMain.List(i, 9) & "") & ", " & _
                  "'채권반제', " & _
                  "'" & voucherNo & "')"

            Debug.Print Sql
            Conn.Execute Sql

        End If

    Next i

    ' ===== 기존 미결 채권채무명세 상태 업데이트 (진짜 반제 처리) =====
    Call 반제_업데이트_DB(Conn, arrIDs, Me.txt전표일자.Value, voucherNo)

    MsgBox "전표가 입력되었습니다." & vbCrLf & "전표번호 : " & voucherNo, vbInformation

    Unload Me

    Conn.Close
    Set Conn = Nothing

End Sub
