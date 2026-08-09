Private Sub btnClose_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then UnloadMe
End Sub

Private Sub btnClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
UnloadMe
End Sub

Private Sub btnDelete_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then DeleteCustomer
End Sub

Private Sub btnDelete_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DeleteCustomer
End Sub

Private Sub btnEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then EditCustomer
End Sub

Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
EditCustomer
End Sub

Private Sub btnInit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Initialize
End Sub

Private Sub btnInit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Initialize
End Sub

Private Sub btnRegister_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then RegisterCustomer
End Sub

Private Sub btnRegister_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
RegisterCustomer
End Sub


Private Sub cbo은행명_Change()

End Sub

Private Sub lstMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
UnloadMe
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub txtSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Filter_ListBox
End Sub

Private Sub lstMain_Click()

Dim vArr As Variant
'Get_ListItm 보조함수
'리스트박스에서 선택된 항목을 배열로 바꿔줍니다.
'Get_ListItm (리스트박스)
vArr = Get_ListItm(Me.lstMain)

Me.txtID.Value = vArr(0)
Me.txtCustomer.Value = vArr(1)
Me.txtContact.Value = vArr(2)
Me.txtPIC.Value = vArr(3)
Me.txtAddress.Value = vArr(4)
Me.txt예금주.Value = vArr(5)
Me.cbo은행명.Value = vArr(7)
Me.txt계좌번호.Value = vArr(8)

End Sub


Private Sub UserForm_Initialize()

Dim db As Variant
db = Get_DB_Access("거래처마스터")

Update_cbo은행명

'Update_list 보조함수
'지정한 리스트박의 값을 배열에서 받아와 갱신합니다.
Update_List Me.lstMain, db, "0pt;170pt;100pt;100pt;250pt;0pt;0pt;100pt;100pt;"

End Sub

Sub Update_cbo은행명()

Dim db As Variant
db = Get_DB_Access("은행코드마스터")

Update은행코드검색 cbo은행명, db

End Sub


Sub UnloadMe()

Unload Me

End Sub


Sub EditCustomer()

    Dim Conn As Object
    Dim Sql As String
    Dim varBankName As Variant
    Dim varBankCode As Variant

    If Me.txtID.Value = "" Then MsgBox "수정할 거래처를 먼저 선택하세요.": Exit Sub
    If Me.txtCustomer.Value = "" Then MsgBox "거래처명을 입력하세요.": Exit Sub
    If Me.txtAddress.Value = "" Then MsgBox "주소를 입력하세요.": Exit Sub

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()


    Sql = "UPDATE 거래처마스터 SET " & _
          "거래처 = '" & Me.txtCustomer.Value & "', " & _
          "사업자등록번호 = '" & Me.txtContact.Value & "', " & _
          "대표 = '" & Me.txtPIC.Value & "', " & _
          "주소 = '" & Me.txtAddress.Value & "', " & _
          "예금주 = '" & Me.txt예금주.Value & "', " & _
          "은행코드 = '" & Me.cbo은행명.Column(1) & "', " & _
          "은행명 = '" & Me.cbo은행명.Column(0) & "', " & _
          "계좌번호 = '" & Me.txt계좌번호.Value & "'" & _
          " WHERE ID = " & Val(Me.txtID.Value)

    'Debug.Print sql
    Conn.Execute Sql

    Conn.Close
    Set Conn = Nothing

    Filter_ListBox
    Select_ListItm Me.lstMain, Me.txtID.Value

    MsgBox "고객정보가 수정되었습니다.", vbInformation, "거래처 수정"

End Sub


Sub Initialize()

'Clear_Ctrls 보조함수
'유저폼 내 특정 컨트롤의 값을 초기화합니다. 와일드카드를 사용할 수 있습니다.
'Clear_Ctrls Me, "초기화컨트롤", "제외할컨트롤"
Clear_Ctrls Me, "txt*", "txtSearch,txtID"
Me.cbo은행명.Value = ""
'Me.txtContact.Value = ""
'Me.txtAddress.Value = ""
'Me.txtPIC.Value = ""

End Sub


Sub RegisterCustomer()

    Dim Conn As Object
    Dim Sql As String
    Dim db As Variant
    Dim varBankName As Variant
    Dim varBankCode As Variant

    If Me.txtCustomer.Value = "" Then MsgBox "고객이름을 입력하세요.": Exit Sub
    If Me.txtAddress.Value = "" Then MsgBox "주소를 입력하세요.": Exit Sub

    ' DB 연결
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()
    

    Sql = "INSERT INTO 거래처마스터 (거래처, 사업자등록번호, 대표, 주소, 예금주, 은행코드, 은행명, 계좌번호)" & _
          " VALUES (" & _
          "'" & Me.txtCustomer.Value & "', " & _
          "'" & Me.txtContact.Value & "', " & _
          "'" & Me.txtPIC.Value & "', " & _
          "'" & Me.txtAddress.Value & "', " & _
          "'" & Me.txt예금주.Value & "', " & _
          "'" & Me.cbo은행명.Column(1) & "', " & _
          "'" & Me.cbo은행명.Column(0) & "', " & _
          "'" & Me.txt계좌번호.Value & "')"

    'Debug.Print sql
    Conn.Execute Sql

    Conn.Close
    Set Conn = Nothing

    db = Get_DB_Access("거래처마스터")
    Update_List Me.lstMain, db, "0pt;170pt;100pt;100pt;250pt;0pt;0pt;100pt;100pt;"

    Initialize

    MsgBox "신규 고객 정보가 등록되었습니다.", vbInformation

End Sub


Sub DeleteCustomer()

    Dim Conn As Object
    Dim Sql As String
    Dim db As Variant
    Dim YN As VbMsgBoxResult

    If Me.txtID.Value = "" Then MsgBox "삭제할 거래처를 선택하세요.", , "거래처 삭제": Exit Sub

    YN = MsgBox("거래처를 정말로 삭제하시겠습니까? 한번 삭제된 정보는 복구가 불가능합니다.", vbYesNo, "거래처 삭제")
    If YN = vbNo Then Exit Sub

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Sql = "DELETE FROM 거래처마스터 WHERE ID = " & Val(Me.txtID.Value)
    Conn.Execute Sql

    Conn.Close
    Set Conn = Nothing

    db = Get_DB_Access("거래처마스터")
    Update_List Me.lstMain, db, "0pt;170pt;100pt;100pt;250pt;0pt;0pt;100pt;100pt;"
    Initialize

    MsgBox "고객정보가 삭제되었습니다.", vbInformation

End Sub


Sub Filter_ListBox()

Dim db As Variant

db = Get_DB_Access("거래처마스터")
db = Filtered_DB(db, Me.txtSearch.Value)

Update_List Me.lstMain, db, "0pt;170pt;100pt;100pt;250pt;0pt;0pt;100pt;100pt;"

End Sub

'---------------------------------------------------
'유저폼 스타일 꾸미기
'-----------------------------------------------------

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btninit 를 버튼 이름으로 변경합니다.
Private Sub btnInit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnInit
End Sub

Private Sub btnInit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnInit
End Sub

Private Sub btnInit_Enter()
OnHover_Css Me.btnInit
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnDelete 를 버튼 이름으로 변경합니다.
Private Sub btnDelete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_Enter()
OnHover_Css Me.btnDelete
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnEdit 를 버튼 이름으로 변경합니다.
Private Sub btnEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_Enter()
OnHover_Css Me.btnEdit
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnRegister 를 버튼 이름으로 변경합니다.
Private Sub btnRegister_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_Enter()
OnHover_Css Me.btnRegister
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose 를 버튼 이름으로 변경합니다.
Private Sub btnClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnClose
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnClose
End Sub

Private Sub btnClose_Enter()
OnHover_Css Me.btnClose
End Sub


'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnInit, btnDelete, btnEdit, btnClose, btnRegister" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
 For Each vList In vLists
 If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
 Next
Next
End Sub
'커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
Private Sub OnHover_Css(lbl As Control): With lbl: .BackColor = RGB(211, 240, 224): .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl: .BackColor = &H8000000E: .BorderColor = -2147483638: End With: End Sub

Private Sub lstMain_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UnhookListBoxScroll
End Sub
Private Sub lstMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.lstMain
End Sub










