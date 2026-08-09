Dim UnLockTime As Date

Private Sub btnCategory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
frmCategory.Show
Update_cboCategory
End Sub

Private Sub btnClose_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Unload Me
End Sub

Private Sub btnClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload Me
End Sub

Private Sub btnCustomer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.lstMain.Enabled = False

frmCustomer.Show

Me.txtCustomerID.Value = arrCustomer(0)
Me.txtCustomer.Value = arrCustomer(1)
Me.txtContact.Value = arrCustomer(2)

UnLockTime = Now() + TimeSerial(0, 0, 1)

End Sub

Private Sub btnDelete_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Delete_Product
End Sub

Private Sub btnDelete_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Delete_Product
End Sub

Private Sub btnEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Edit_Product
End Sub

Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Edit_Product
End Sub

Private Sub btnInit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Initialize
End Sub

Private Sub btnInit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Initialize
End Sub

Private Sub btnRegister_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Register_Product
End Sub

Private Sub btnRegister_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Register_Product
End Sub

Private Sub btnUnit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
frmUnit.Show
Update_cboUnit
End Sub

Private Sub lstMain_Click()
Click_lstMain
End Sub

Private Sub txtSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Update_lstMain
End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Initialize()
Update_lstMain
Update_cboUnit
Update_cboCategory
Me.txtSearch.SetFocus
End Sub

'------------------------------------------------

Sub Delete_Product()
Dim YN As VbMsgBoxResult

YN = MsgBox("정말로 삭제하시겠습니까? 한번 삭제된 데이터는 복구가 불가능합니다.", vbYesNo)
If YN = vbNo Then Exit Sub

Delete_Record shtItem, Me.txtID.Value
Update_lstMain
Initialize

MsgBox "제품 정보가 삭제 되었습니다", vbInformation

End Sub
Sub Edit_Product()

Dim DB As Variant
Dim blnUnique As Boolean

DB = Get_DB(shtItem)
blnUnique = IsUnique(DB, Me.txtSKU.Value, 4, Get_ListItm(Me.lstMain)(3))
If blnUnique = False Then MsgBox "동일한 제품 코드가 존재합니다. 다시 확인해주세요.", vbExclamation:   Exit Sub

Update_Record shtItem, Me.txtID.Value, Me.txtCustomerID.Value, Me.cboCategory.Value, _
Me.txtSKU.Value, Me.txtProductName.Value, Me.cboUnit.Value, _
Me.txtRemark.Value

Update_lstMain
Select_ListItm Me.lstMain, Me.txtID.Value

MsgBox "제품 정보 수정이 완료되었습니다.", vbInformation

End Sub
Sub Register_Product()
Dim DB As Variant
Dim blnUnique As Boolean  'True/False

DB = Get_DB(shtItem)
blnUnique = IsUnique(DB, Me.txtSKU.Value, 4)
If blnUnique = False Then MsgBox "동일한 제품 코드가 존재합니다. 다시 확인해주세요.", vbExclamation:   Exit Sub

Insert_Record shtItem, Me.txtCustomerID.Value, Me.cboCategory.Value, _
Me.txtSKU.Value, Me.txtProductName.Value, Me.cboUnit.Value, _
Me.txtRemark.Value

Update_lstMain
Initialize
MsgBox "신규 제품 정보가 등록 되었습니다.", vbInformation

End Sub
Sub Initialize()

'대소문자구분합니다.
Clear_Ctrls Me, "txt*,cbo*", "txtCustomerID,txtID"

End Sub
Sub Click_lstMain()
Dim vArr As Variant

vArr = Get_ListItm(Me.lstMain)


Me.txtID.Value = vArr(0)
Me.txtCustomerID.Value = vArr(1)
Me.txtSKU.Value = vArr(3)
Me.txtProductName.Value = vArr(4)
Select_CboItm Me.cboUnit, vArr(5)
Select_CboItm Me.cboCategory, vArr(2)
Me.txtRemark.Value = vArr(6)
Me.txtCustomer.Value = vArr(8)
Me.txtContact.Value = vArr(9)

End Sub

Sub Update_lstMain()
Dim DB As Variant
DB = Get_DB(shtItem)
DB = Connect_DB(DB, 3, shtCategory, "제품구분")
DB = Connect_DB(DB, 2, shtCustomer, "거래처명,연락처")
DB = Filtered_DB(DB, Me.txtSearch.Value)

Update_List Me.lstMain, DB, _
"0,0,0,100,150,50,120,100,120,0"

End Sub

Sub Update_cboUnit()
Dim DB As Variant
DB = Get_DB(shtUnit, True)

Update_Cbo Me.cboUnit, DB

End Sub

Sub Update_cboCategory()
Dim DB As Variant
DB = Get_DB(shtCategory)

Update_Cbo Me.cboCategory, DB, 2

End Sub

'---------------------------------------

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnEdit를 버튼 이름으로 변경합니다.
Private Sub btnEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnEdit
End Sub

Private Sub btnEdit_Enter()
OnHover_Css Me.btnEdit
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose를 버튼 이름으로 변경합니다.
Private Sub btnClose_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnClose
End Sub

Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnClose
End Sub

Private Sub btnClose_Enter()
OnHover_Css Me.btnClose
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnInit를 버튼 이름으로 변경합니다.
Private Sub btnInit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnInit
End Sub

Private Sub btnInit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnInit
End Sub

Private Sub btnInit_Enter()
OnHover_Css Me.btnInit
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnRegister를 버튼 이름으로 변경합니다.
Private Sub btnRegister_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnRegister
End Sub

Private Sub btnRegister_Enter()
OnHover_Css Me.btnRegister
End Sub


'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnDelete를 버튼 이름으로 변경합니다.
Private Sub btnDelete_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnDelete
End Sub

Private Sub btnDelete_Enter()
OnHover_Css Me.btnDelete
End Sub

'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnRegister,btnInit,btnEdit,btnClose,btnDelete" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
    For Each vList In vLists
        If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
    Next
Next
If Now() >= UnLockTime And Me.lstMain.Enabled = False Then lstMain.Enabled = True

End Sub

'커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
Private Sub OnHover_Css(lbl As Control): With lbl:   .BackColor = RGB(211, 240, 224):    .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl:   .BackColor = &H8000000E:    .BorderColor = &H8000000A: End With: End Sub
