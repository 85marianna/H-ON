Private Sub btnAdd_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
InsertUnit
End Sub

Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
EditUnit
End Sub

Private Sub btnClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload Me
End Sub

Private Sub btnDelete_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
DeleteUnit
End Sub

Private Sub btnDown_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim WS As Worksheet
Dim Rng As Range: Dim Target As Range

If isListBoxSelected(Me.lstMain) = False Then Exit Sub

Set WS = shtUnit
Set Rng = WS.UsedRange
Set Target = Rng.Find(Get_ListItm(Me.lstMain)(0))
If Target.Row >= Rng.Rows.Count Then Exit Sub

Target.Offset(1, 0).Cut
Target.Insert

UpdateList
Select_ListItm Me.lstMain, Me.txtUnit.Value

End Sub

Private Sub btnUp_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim WS As Worksheet
Dim Rng As Range: Dim Target As Range

If isListBoxSelected(Me.lstMain) = False Then Exit Sub

Set WS = shtUnit
Set Rng = WS.UsedRange
Set Target = Rng.Find(Get_ListItm(Me.lstMain)(0))
If Target.Row <= 2 Then Exit Sub

Target.Cut
Target.Offset(-1, 0).Insert

UpdateList
Select_ListItm Me.lstMain, Me.txtUnit.Value

End Sub

Private Sub lstMain_Click()
ClickListMain
End Sub

Private Sub UserForm_Initialize()
UpdateList
End Sub

'---------------------------

Sub UpdateList()
Dim DB As Variant
DB = Get_DB(shtUnit, True)
Update_List Me.lstMain, DB, "80"

End Sub

Sub ClickListMain()
Me.txtUnit.Value = Get_ListItm(Me.lstMain)(0)
End Sub

Sub InsertUnit()
Dim DB As Variant
DB = Get_DB(shtUnit, True)
If IsUnique(DB, Me.txtUnit.Value) = False Then MsgBox "중복된 값이 존재합니다. 다시 확인해주세요.", vbInformation: Exit Sub

Insert_Record shtUnit, Me.txtUnit.Value
UpdateList

MsgBox "단위가 등록되었습니다.", vbInformation

End Sub

Sub DeleteUnit()
Dim YN As VbMsgBoxResult
YN = MsgBox("정말로 삭제하겠습니까? 삭제된 데이터는 복구가 불가능합니다.", vbYesNo)
If YN = vbNo Then Exit Sub

Delete_Record shtUnit, Get_ListItm(Me.lstMain)(0)
UpdateList
MsgBox "단위가 삭제되었습니다.", vbInformation
End Sub

'-------------------------------------

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnAdd를 버튼 이름으로 변경합니다.
Private Sub btnAdd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnAdd
End Sub

Private Sub btnAdd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnAdd
End Sub

Private Sub btnAdd_Enter()
OnHover_Css Me.btnAdd
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

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnUp를 버튼 이름으로 변경합니다.
Private Sub btnUp_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnUp
End Sub

Private Sub btnUp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnUp
End Sub

Private Sub btnUp_Enter()
OnHover_Css Me.btnUp
End Sub

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnDown를 버튼 이름으로 변경합니다.
Private Sub btnDown_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnDown
End Sub

Private Sub btnDown_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnDown
End Sub

Private Sub btnDown_Enter()
OnHover_Css Me.btnDown
End Sub

'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnAdd,btnDelete,btnClose,btnUp,btnDown" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
Dim vLists As Variant: Dim vList As Variant
If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
For Each ctl In Me.Controls
    For Each vList In vLists
        If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
    Next
Next
If Now() >= UnLockTime And lstMain.Enabled = False Then lstMain.Enabled = True

End Sub

'커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
Private Sub OnHover_Css(lbl As Control): With lbl:   .BackColor = RGB(211, 240, 224):    .BorderColor = RGB(134, 191, 160): End With: End Sub
Private Sub OutHover_Css(lbl As Control): With lbl:   .BackColor = &H8000000E:    .BorderColor = &H8000000A: End With: End Sub
