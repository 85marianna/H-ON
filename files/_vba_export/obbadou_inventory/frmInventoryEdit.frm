Dim shp As Shape
Dim cRow As Long

Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Me.txtDate.Value = frmCalendar.GetDate
End Sub

Private Sub btnEdit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'UpdateInventory
End Sub

Private Sub UserForm_Initialize()

'cRow = Selection.Row
'
'If shtMain.Range("F" & cRow).Value = "" And shtMain.Range("J" & cRow).Value > 0 Then MsgBox "요약된 정보는 수정할 수 없습니다.", vbInformation: End
'If shtMain.Range("F" & cRow).Value = "" Then MsgBox "올바른 값을 선택하세요.", vbInformation: End
'
'With shtMain
'    Update_Cbo Me.cboUnit, Get_DB(shtUnit, True)
'    Me.txtID.Value = .Range("F" & cRow).Value
'    Me.txtProductName.Value = .Range("O" & cRow).Value
'    Me.txtProductCode.Value = .Range("N" & cRow).Value
'    Me.txtDate.Value = .Range("K" & cRow).Value
'    Me.txtQtyIN.Value = .Range("R" & cRow).Value
'    Me.txtQtyOUT.Value = .Range("S" & cRow).Value
'    Me.cboUnit.Value = .Range("P" & cRow).Value
'    Me.txtRemark.Value = .Range("Q" & cRow).Value
'    Me.txtMemo.Value = .Range("U" & cRow).Value
'End With
'
'Set Shp = ShapeInRange(shtMain.Range("F" & cRow & ":U" & cRow))

End Sub

Private Sub UserForm_Terminate()
'Shp.Delete
End Sub

'----------------------------
Sub UpdateInventory()

'Update_Record shtInventory, Me.txtID.Value, , Me.txtDate.Value, Me.txtQtyIN.Value, Me.txtQtyOUT.Value, Me.cboUnit.Value, _
'Me.txtRemark.Value, Me.txtMemo.Value
'
'With shtMain
'    .Range("K" & cRow).Value = Me.txtDate.Value
'    .Range("P" & cRow).Value = Me.cboUnit.Value
'    .Range("Q" & cRow).Value = Me.txtRemark.Value
'    .Range("R" & cRow).Value = Me.txtQtyIN.Value
'    .Range("S" & cRow).Value = Me.txtQtyOUT.Value
'    .Range("U" & cRow).Value = Me.txtMemo.Value
'End With
'
'MsgBox "재고의 입출고 정보가 수정되었습니다.", vbInformation
'
'Unload Me

End Sub

''------------------------------------------------------------------------------------------

''유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose를 버튼 이름으로 변경합니다.
'Private Sub btnEdit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'OutHover_Css Me.btnEdit
'End Sub
'
'Private Sub btnEdit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'OnHover_Css Me.btnEdit
'End Sub
'
'Private Sub btnEdit_Enter()
'OnHover_Css Me.btnEdit
'End Sub
'
''아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Dim ctl As Control
'Dim btnList As String: btnList = "btnEdit" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
'Dim vLists As Variant: Dim vList As Variant
'If InStr(1, btnList, ",") > 0 Then vLists = Split(btnList, ",") Else vLists = Array(btnList)
'For Each ctl In Me.Controls
'    For Each vList In vLists
'        If InStr(1, ctl.Name, Trim(vList)) > 0 Then OutHover_Css ctl
'    Next
'Next
'
'End Sub
'
''커서 이동시 버튼 색깔을 변경하는 보조명령문을 유저폼에 추가합니다.
'Private Sub OnHover_Css(lbl As Control): With lbl:   .BackColor = RGB(211, 240, 224):    .BorderColor = RGB(134, 191, 160): End With: End Sub
'Private Sub OutHover_Css(lbl As Control): With lbl:   .BackColor = &H8000000E:    .BorderColor = &H8000000A: End With: End Sub
