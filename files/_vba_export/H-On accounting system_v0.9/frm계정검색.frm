
Private Sub CommandButton1_Click()
Unload Me

End Sub

Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Unload Me

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim Index   As Long
Index = ListBox1.ListIndex
data_1 = ListBox1.List(Index, 0)
data_2 = ListBox1.List(Index, 1)
data_3 = ListBox1.List(Index, 2)

'MsgBox data_1 & "-" & data_2 & "-" & data_3
Unload Me
End Sub

Private Sub TextBox1_Change()

Dim i As Long
Dim searchWord As String '검색어를 저장할 변수
Dim found As Boolean
'Dim q(1, 3)


searchWord = Me.TextBox1.Value '검색어를 searchWord 변수에 저장합니다.
Me.ListBox1.Clear '검색 결과를 출력할 리스트 박스를 초기화합니다.

'데이터 소스에서 검색합니다.
For i = 1 To WorksheetFunction.CountA(Sheets("기준").Range("C:C"))
    If InStr(Sheets("기준").Range("C" & i + 2).Value, searchWord) > 0 Then '검색어가 소스 문자열에 포함되어 있으면
            
        With Me.ListBox1
        .AddItem Sheets("기준").Range("A" & i + 2).Value
        .List(ListBox1.ListCount - 1, 1) = Sheets("기준").Range("B" & i + 2).Value
        .List(ListBox1.ListCount - 1, 2) = Sheets("기준").Range("C" & i + 2).Value
        End With
        
    found = True
End If
Next

'검색 결과가 없으면 알림 메시지를 출력합니다.
If found = False Then
MsgBox "검색 결과가 없습니다.", vbInformation
End If


End Sub



Private Sub UserForm_Initialize()

TextBox1.IMEMode = vbIMEModeHangul '한글

Me.ListBox1.Clear '검색 결과를 출력할 리스트 박스를 초기화합니다.

End Sub