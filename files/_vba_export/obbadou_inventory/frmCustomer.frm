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

End Sub

Private Sub UserForm_Initialize()

Dim DB As Variant
DB = Get_DB(shtCustomer)
'Update_list 보조함수
'지정한 리스트박의 값을 배열에서 받아와 갱신합니다.
'Update_list Listbox, DB, 열넓이
'열넓이 "0pt;50pt...."
Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"

Me.txtSearch.SetFocus

If frmProduct.Visible = True Then Me.btnClose.Text = "선택"

End Sub

'-------------------------------------------------------------

Sub UnloadMe()

If frmProduct.Visible = False Then
    Unload Me
Else
    If isListBoxSelected(Me.lstMain) = False Then MsgBox "거래처를 선택하세요.", vbInformation: Exit Sub
    arrCustomer = Get_ListItm(Me.lstMain)
    Unload Me
End If

End Sub
Sub EditCustomer()

Dim DB As Variant

Update_Record shtCustomer, Me.txtID.Value, _
Me.txtCustomer.Value, Me.txtContact.Value, _
Me.txtPIC.Value, Me.txtAddress.Value

Filter_ListBox

Select_ListItm Me.lstMain, Me.txtID.Value

MsgBox "고객정보가 수정되었습니다.", vbInformation

End Sub


Sub Initialize()

'Clear_Ctrls 보조함수
'유저폼 내 특정 컨트롤의 값을 초기화합니다. 와일드카드를 사용할 수 있습니다.
'Clear_Ctrls Me, "초기화컨트롤", "제외할컨트롤"
Clear_Ctrls Me, "txt*", "txtSearch,txtID"

'Me.txtCustomer.Value = ""
'Me.txtContact.Value = ""
'Me.txtAddress.Value = ""
'Me.txtPIC.Value = ""

End Sub


Sub RegisterCustomer()

Dim DB As Variant

If Me.txtCustomer.Value = "" Then MsgBox "고객이름을 입력하세요.": Exit Sub
If Me.txtContact.Value = "" Then MsgBox "연락처를 입력하세요.": Exit Sub
If Me.txtPIC.Value = "" Then MsgBox "담당자를 입력하세요.": Exit Sub
If Me.txtAddress.Value = "" Then MsgBox "주소를 입력하세요.": Exit Sub

Insert_Record shtCustomer, Me.txtCustomer.Value, _
Me.txtContact.Value, Me.txtPIC.Value, Me.txtAddress.Value

DB = Get_DB(shtCustomer)

Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"

Initialize

MsgBox "신규 고객 정보가 등록되었습니다.", vbInformation

End Sub

Sub DeleteCustomer()

Dim DB As Variant
Dim YN As VbMsgBoxResult

YN = MsgBox("고객정보를 정말로 삭제하시겠습니까? 한번 삭제된 정보는 복구가 불가능합니다.", vbYesNo)
If YN = vbNo Then Exit Sub

Delete_Record shtCustomer, Me.txtID.Value

DB = Get_DB(shtCustomer)
Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"
Initialize

MsgBox "고객정보가 삭제되었습니다.", vbInformation

End Sub

Sub Filter_ListBox()

Dim DB As Variant

DB = Get_DB(shtCustomer)
DB = Filtered_DB(DB, Me.txtSearch.Value)

Update_List Me.lstMain, DB, "0pt;120pt;100pt;80pt;150pt;"

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












