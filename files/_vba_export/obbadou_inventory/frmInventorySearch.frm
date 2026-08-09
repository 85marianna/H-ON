Private Sub btnDateFrom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.txtFrom.Value = frmCalendar.GetDate
End Sub

Private Sub btnDateTo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Me.txtTo.Value = frmCalendar.GetDate
End Sub

Private Sub btnSearch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim DB As Variant
DB = Get_DB(shtInventory, False)
DB = Connect_DB(DB, 2, shtItem, "ID_거래처,ID_제품구분,제품코드,제품명")
DB = Connect_DB(DB, 10, shtCategory, "제품구분")
DB = Connect_DB(DB, 9, shtCustomer, "거래처명,연락처,담당자,주소")

If Me.txtFrom.Value <> "" Then DB = Filtered_DB(DB, ">=" & Me.txtFrom.Value, 3)
If Me.txtTo.Value <> "" Then DB = Filtered_DB(DB, "<=" & Me.txtTo.Value, 3)
If Me.txtProductName.Value <> "" Then DB = Filtered_DB(DB, Me.txtProductName.Value, 12)
If Me.txtProductCode.Value <> "" Then DB = Filtered_DB(DB, Me.txtProductCode.Value, 11)
If Me.txtCustomer.Value <> "" Then DB = Filtered_DB(DB, Me.txtCustomer.Value, 14)

If IsEmpty(DB) Then MsgBox "조건을 만족하는 재고 데이터가 없습니다. 다시 확인하세요.", vbInformation: Exit Sub

DB = Sort2DArray(DB, 3, 1)

ClearInventoryData
ArrayToRng shtMain.Range("F5"), DB, "1,9,2,10,,3,14,16,11,12,6,7,4,5,,8"
SequenceToRng shtMain.Range("J5"), UBound(DB, 1)

shtMain.Columns("R:S").EntireColumn.Hidden = False

If IsUniqueArray(DB, 2) = True Then
    shtMain.Range("t:t").EntireColumn.Hidden = False
    RunningSumRng shtMain.Range("t5"), UBound(DB, 1), -2, -1
Else
    shtMain.Range("t:t").EntireColumn.Hidden = True
End If
    
MsgBox "재고 정보를 조회하였습니다.", vbInformation
Unload Me

End Sub

Private Sub optAll_Click()
ChangeDate
End Sub

Private Sub optLastMonth_Click()
ChangeDate
End Sub

Private Sub optThisMonth_Click()
ChangeDate
End Sub

Private Sub optThisYear_Click()
ChangeDate
End Sub

Private Sub optToday_Click()
ChangeDate
End Sub

Private Sub UserForm_Initialize()
optThisMonth.Value = True
End Sub

Sub ChangeDate()

Dim Y As Long: Y = Year(Date)
Dim M As Long: M = Month(Date)

If optThisMonth.Value = True Then
Me.txtFrom.Value = DateSerial(Y, M, 1)
Me.txtTo.Value = DateSerial(Y, M + 1, 0)
End If

If optLastMonth.Value = True Then
Me.txtFrom.Value = DateSerial(Y, M - 1, 1)
Me.txtTo.Value = DateSerial(Y, M, 0)
End If

If optToday.Value = True Then
Me.txtFrom.Value = Date
Me.txtTo.Value = Date
End If

If optAll.Value = True Then
Me.txtFrom.Value = ""
Me.txtTo.Value = ""
End If

If optThisYear.Value = True Then
Me.txtFrom.Value = DateSerial(Y, 1, 1)
Me.txtTo.Value = DateSerial(Y, 12, 31)
End If

End Sub

