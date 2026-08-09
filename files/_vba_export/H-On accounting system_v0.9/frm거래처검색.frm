
Private Sub btnClose_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Unload Me

End Sub


Private Sub lstMain_Click()

Dim vArr As Variant
'Get_ListItm 보조함수
'리스트박스에서 선택된 항목을 배열로 바꿔줍니다.
'Get_ListItm (리스트박스)

vArr = Get_ListItm(Me.lstMain)

frmAcctInput.txt거래처.Value = vArr(1)

Unload Me

End Sub

Private Sub txtSearch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Filter_ListBox

End Sub

Private Sub UserForm_Initialize()

Dim db As Variant

db = Get_DB_Access("거래처마스터")

'Update_list 보조함수
'지정한 리스트박의 값을 배열에서 받아와 갱신합니다.
'Update_list Listbox, DB, 열넓이
'열넓이 "0pt;50pt...."
Update_List Me.lstMain, db, "0pt;150pt;100pt;150pt;200pt;0pt;0pt;0pt;0pt;"

Me.txtSearch.SetFocus

End Sub

Sub Filter_ListBox()

Dim db As Variant

db = Get_DB_Access("거래처마스터")

db = Filtered_DB(db, Me.txtSearch.Value)

Update_List Me.lstMain, db, "0pt;120pt;100pt;80pt;150pt;"

End Sub

Private Sub lstMain_Exit(ByVal Cancel As MSForms.ReturnBoolean)
UnhookListBoxScroll
End Sub

Private Sub lstMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.lstMain
End Sub

Private Sub UserForm_Click()

End Sub