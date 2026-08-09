

Private Sub btnExit_Click()

    Dim ret As VbMsgBoxResult

    ret = MsgBox("시스템을 종료하시겠습니까?", vbQuestion + vbYesNo, "H-On Accounting System 종료")

    If ret = vbYes Then
        Application.DisplayAlerts = False
        Application.Quit   ' 여기서 종료 루트 진입
        ' DisplayAlerts 복구는 필요 없음 (프로세스 종료)
    End If

End Sub


Private Sub CommandButton1_Click()

frmAcctInput.Show 1

End Sub

Private Sub CommandButton2_Click()

frm전표인쇄.Show 1

End Sub

Private Sub CommandButton3_Click()

frm보고서조회.Show 1

End Sub

'Private Sub lstNotice_Click()
'
'    If Me.lstNotice.ListIndex = -1 Then Exit Sub
'
'    Me.txtNotice.Value = LoadNoticeContent(Me.lstNotice.List(Me.lstNotice.ListIndex, 0))
'
'    Me.lstNotice.SetFocus
'
'End Sub



Private Sub lstNotice_Click()

    If Me.lstNotice.ListIndex = -1 Then Exit Sub

    Me.txtNotice.Value = _
        Me.lstNotice.List(Me.lstNotice.ListIndex, 3) & vbCrLf & _
        String(70, "=") & vbCrLf & vbCrLf & _
        LoadNoticeContent(Me.lstNotice.List(Me.lstNotice.ListIndex, 0))

    Me.lstNotice.SetFocus

End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)


Dim inputPass As String   ' 맨 위로 이동

    ' 클릭한 메뉴의 이름을 메시지 박스로 확인 (테스트용)
    ' MsgBox Node.Text & "를 선택했습니다."
' 뒤에 1을 붙이면 이 폼은 '모달'로 떠서 엑셀을 못 건드리게 돼!
    Select Case Node.key
        Case "CHILD1" ' 일반전표
            frmAcctInput.Show 1
            shtMAIN.Activate
            
        Case "CHILD2"
            frm반제전표.Show 1
            shtMAIN.Activate
            
        Case "CHILD3"
            frm외화평가.Show 1
            shtMAIN.Activate
            
        Case "ROOT2"
            frm전표인쇄.Show 1
            shtMAIN.Activate
            
        Case "CHILD4"
            frm채권채무조회.Show 1
            shtMAIN.Activate
        
        Case "CHILD5"
            frm송금리스트.Show 1
            shtMAIN.Activate
            
        Case "ROOT4"
            frm제예금명세.Show 1
            shtMAIN.Activate
            
        Case "CHILD6" ' 총계정원장/계정별명세
            frm원장조회.Show 1
            
        Case "CHILD7"
            frm보고서조회.Show 1

            
        Case "CHILD11"
            frm계정검색.Show 1
            
        Case "CHILD13"
        
            inputPass = InputBox("거래처 정보를 관리하려면" & vbCrLf & "관리자 비밀번호를 입력하세요.", "거래처관리 보안 확인")

            If inputPass = "" Then Exit Sub
            If inputPass <> ADMIN_PW Then
                MsgBox "비밀번호가 일치하지 않습니다.", vbCritical, "오류"
                Exit Sub
            End If

            frm거래처관리.Show 1
            
        Case "CHILD15"   ''개발자모드
            Call ToggleSheetTabs

            
            
'        Case "CHILD5" ' 대차대조표 (BS)
'            ' 폼이 아니라 시트로 이동할 때
'            Sheets("BS").Activate
'
'        Case "CHILD6" ' 손익계산서 (IS)
'            Sheets("IS").Activate
            
    End Select

End Sub



Private Sub UserForm_Initialize()

    Dim Node As Node
    
    LoadNotice Me.lstNotice
    
    Label1 = Application.UserName & " 님 환영합니다!  ● 로그인일시: " & Format(gLoginTime, "yyyy-mm-dd hh:mm:ss")
    TreeView1.Nodes.Clear
    
    이미지리스트.ListImages.Add key:="이미지1", Picture:=LoadPicture(ThisWorkbook.Path & "\메뉴아이콘.jpg")
    이미지리스트.ListImages.Add key:="이미지2", Picture:=LoadPicture(ThisWorkbook.Path & "\점아이콘.jpg")
    이미지리스트.ListImages.Add key:="이미지3", Picture:=LoadPicture(ThisWorkbook.Path & "\설정이미지.jpg")

    Set TreeView1.ImageList = 이미지리스트
    
    ' ===== 전표입력 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT1", " 전표입력 ", Image:="이미지1")
    Set Node = TreeView1.Nodes.Add("ROOT1", tvwChild, "CHILD1", "일반전표", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT1", tvwChild, "CHILD2", "반제전표", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT1", tvwChild, "CHILD3", "외화평가전표", Image:="이미지2")

    ' 공백줄
    TreeView1.Nodes.Add , , "SP1", " "
    
    ' ===== 전표조회 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT2", " 전표조회 ", Image:="이미지1")
    
    TreeView1.Nodes.Add , , "SP2", " "

    ' ===== 채권채무 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT3", " 채권채무 ", Image:="이미지1")
    Set Node = TreeView1.Nodes.Add("ROOT3", tvwChild, "CHILD4", "채권채무명세", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT3", tvwChild, "CHILD5", "송금리스트", Image:="이미지2")

    TreeView1.Nodes.Add , , "SP3", " "

    ' ===== 제예금명세 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT4", " 제예금명세 ", Image:="이미지1")

    TreeView1.Nodes.Add , , "SP4", " "

    ' ===== 재무제표 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT5", " 보고서 조회 ", Image:="이미지1")
    Set Node = TreeView1.Nodes.Add("ROOT5", tvwChild, "CHILD6", "계정원장 조회", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT5", tvwChild, "CHILD7", "재무재표 조회", Image:="이미지2")

    TreeView1.Nodes.Add , , "SP5", " "

    ' ===== 마스터관리 =====
    Set Node = TreeView1.Nodes.Add(, , "ROOT6", " 마스터관리", Image:="이미지1")
    Set Node = TreeView1.Nodes.Add("ROOT6", tvwChild, "CHILD11", "계정코드관리", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT6", tvwChild, "CHILD12", "부서코드관리", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT6", tvwChild, "CHILD13", "거래처코드관리", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT6", tvwChild, "CHILD14", "은행코드관리", Image:="이미지2")
    Set Node = TreeView1.Nodes.Add("ROOT6", tvwChild, "CHILD15", "개발자환경", Image:="이미지3")

End Sub


Private Sub UserForm_Terminate()
' 이건: [X] 버튼으로 닫아도 무조건 실행됨 (MAIN유저폼 닫으면 MAIN시트로 이동!)

    ActivateSheetSafely shtMAIN

End Sub


Public Sub LoadNotice(lst As MSForms.ListBox)

    Dim Conn As Object
    Dim Rs As Object
    Dim Sql As String

    Set Conn = ConnectAccessDB()

    Sql = "SELECT ID, 등록일, 구분, 제목, 등록자 " & _
        "FROM 공지사항 " & _
        "ORDER BY 등록일 DESC;"

    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Sql, Conn, 1, 1

    lst.Clear

    Do Until Rs.EOF

        lst.AddItem Rs!ID
        lst.List(lst.ListCount - 1, 1) = Format(Rs!등록일, "yyyy-mm-dd")
        lst.List(lst.ListCount - 1, 2) = Rs!구분
        lst.List(lst.ListCount - 1, 3) = Rs!제목
        lst.List(lst.ListCount - 1, 4) = Rs!등록자

        Rs.MoveNext
    Loop

    Rs.Close
    Conn.Close

End Sub

Public Function LoadNoticeContent(ByVal NoticeID As Long) As String

    Dim Conn As Object
    Dim Rs As Object
    Dim Sql As String

    Set Conn = ConnectAccessDB()

    Sql = "SELECT 내용 FROM 공지사항 " & _
          "WHERE ID = " & NoticeID

    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Sql, Conn, 1, 1

    If Not Rs.EOF Then

        If IsNull(Rs!내용) Then
            LoadNoticeContent = ""
        Else
            LoadNoticeContent = CStr(Rs!내용)
        End If

    End If

    Rs.Close
    Conn.Close

    Set Rs = Nothing
    Set Conn = Nothing

End Function
