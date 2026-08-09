
Private Sub btnSelect_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SelectProduct
End Sub

Private Sub lstMain_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SelectProduct
End Sub

Private Sub txtSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Update_lstMain
End Sub

Private Sub UserForm_Initialize()
Update_lstMain
End Sub

'------------------------------------------------
Sub SelectProduct()
Dim Arr As Variant
Arr = Get_ListItm(Me.lstMain)

With shtMain
    .Range("A7").Value = Arr(0)
    .Range("A9").Value = Arr(2)
    .Range("C7").Value = Arr(3)
    .Range("C8").Value = Arr(5)
    .Range("C9").Value = Arr(8)
    .Range("C6").Value = Arr(4)
    .Range("C12").Value = Arr(6)
End With

Unload Me

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

'---------------------------------------

'유저폼에 추가한 버튼에 개수만큼 아래 명령문을 유저폼에 추가한 뒤, btnClose를 버튼 이름으로 변경합니다.
Private Sub btnSelect_Exit(ByVal Cancel As MSForms.ReturnBoolean)
OutHover_Css Me.btnSelect
End Sub

Private Sub btnSelect_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
OnHover_Css Me.btnSelect
End Sub

Private Sub btnSelect_Enter()
OnHover_Css Me.btnSelect
End Sub

'아래 코드를 유저폼에 추가한 뒤, "btnXXX, btnYYY"를 버튼이름을 쉼표로 구분한 값으로 변경합니다.
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim ctl As Control
Dim btnList As String: btnList = "btnSelect" ' 버튼 이름을 쉼표로 구분하여 입력하세요.
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
