Option Explicit


Private Sub cmdLogin_Click()

    '입력 정보 확인
    If Trim(txtEmpNo.Value) = "" Or Trim(txtPassword.Value) = "" Then
        MsgBox "사번과 비밀번호를 입력해주세요.", vbExclamation
        Exit Sub
    End If

    Dim Conn As Object
    Dim Rs As Object
    Dim Sql As String

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    '1. 사번 확인
    Sql = "SELECT * FROM 로그인정보 " & _
          "WHERE 사번 = '" & txtEmpNo.Value & "'"

    Set Rs = Conn.Execute(Sql)

    If Rs.EOF Then
        MsgBox "등록되지 않은 사번입니다.", vbExclamation
        GoTo ExitProc
    End If

    '2. 비밀번호 확인
    If Rs("비밀번호") <> txtPassword.Value Then
        MsgBox "비밀번호가 일치하지 않습니다.", vbExclamation
        GoTo ExitProc
    End If
    
    '3. 로그인 성공

    gUserID = Rs("ID")
    gEmpNo = Rs("사번")
    gUserName = Rs("이름")
    gAuthority = Rs("권한")
    gLoginTime = Now
    
    IsLogin = True
    
    ''MsgBox gUserName & "님 환영합니다.", vbInformation
    
    Unload Me
    frmMain.Show
    

ExitProc:

    If Not Rs Is Nothing Then
        If Rs.State = 1 Then Rs.Close
    End If

    If Not Conn Is Nothing Then
        If Conn.State = 1 Then Conn.Close
    End If

    Set Rs = Nothing
    Set Conn = Nothing

End Sub


Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'엔터만 눌러도 로그인가능하도록

    If KeyCode = vbKeyReturn Then
        cmdLogin_Click
    End If

End Sub
