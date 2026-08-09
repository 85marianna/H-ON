Public arrCustomer As Variant

Sub OpenfrmCustomer()
frmCustomer.Show
End Sub

Sub OpenfrmProduct()
frmProduct.Show
End Sub

Sub OpenfrmProductSelect()
frmProductSelect.Show
End Sub

Sub OpenfrmInventorySearch()
frmInventorySearch.Show
End Sub

Sub OpenfrmInventoryEdit()
frmInventoryEdit.Show
End Sub

Sub ClearProductData()
shtMain.Range("C5:D15").ClearContents
shtMain.Range("A7:A9").ClearContents

End Sub

Sub ClearInventoryData()
'<------ 재고정보 초기화 (기준열 : 순번)  ------->
ClearContentsBelow shtMain.Range("F5"), "U", 5
End Sub

Sub InventoryDate()
shtMain.Range("C5").Value = frmCalendar.GetDate
End Sub

Sub Insert_Inventory()

With shtMain
    '<------ 재고 정보 누락되었을 경우 안내메시지 출력 후 명령문 종료  ------->
    If .Range("C5").Value = "" Then MsgBox "날짜를 입력하세요.", vbInformation: Exit Sub
    If .Range("C7").Value = "" Then MsgBox "등록할 제품을 선택하세요.", vbInformation: Exit Sub
    If Not .Range("C10").Value + .Range("C11").Value > 0 Then MsgBox "입/출고 수량을 입력하세요.", vbInformation: Exit Sub
    '---------------------------------------------------------------------------------------
    Insert_Record shtInventory, .Range("A7").Value, .Range("C5").Value, CDbl(.Range("C10").Value), CDbl(.Range("C11").Value), .Range("C8").Value, .Range("C12").Value, .Range("C14").Value
End With

MsgBox "재고 입출고 정보가 등록되었습니다.", vbInformation
ClearProductData

End Sub

Sub GetCurrentBalance()

Dim DB As Variant
DB = Get_DB(shtItem)
DB = Connect_DB(DB, 2, shtCustomer, "거래처명,담당자")
DB = Get_Balance(DB, shtInventory, 4, 5, 2)
DB = Filtered_DB(DB, ">" & 0, 10)

'ArrayToRng Sheet1.Range("a1"), DB  '<- DB결과 출력 확인

'오류처리
If IsEmpty(DB) Then MsgBox "현재 보관중인 재고가 없습니다.", vbInformation: Exit Sub

ClearInventoryData
ArrayToRng shtMain.Range("f5"), DB, ",2,1,3,,,8,9,4,5,6,7,,,10"

SequenceToRng shtMain.Range("j5"), UBound(DB, 1)
ValueToRng shtMain.Range("k5"), UBound(DB, 1), Date

'불필요한 열 숨김
shtMain.Columns("R:S").EntireColumn.Hidden = True

shtMain.Range("T:T").EntireColumn.Hidden = False

MsgBox "현재 재고 정보 조회를 완료하였습니다", vbInformation

End Sub

Sub DeleteInventory()

Dim shp As Shape
Dim cRow As Long
Dim YN As VbMsgBoxResult

cRow = Selection.Row

If shtMain.Range("f" & cRow).Value = "" And shtMain.Range("J" & cRow).Value > 0 Then
    MsgBox "요약된 정보는 삭제할 수 없습니다.", vbRetryCancel
    Exit Sub
End If

If shtMain.Range("f" & cRow).Value = "" And shtMain.Range("J" & cRow).Value = "" Then
    MsgBox "삭제할 행을 선택하세요.", vbMsgBoxRight
    Exit Sub
End If

Set shp = ShapeInRange(shtMain.Range("F" & cRow & ":U" & cRow))

YN = MsgBox("선택된 재고 정보를 삭제하시겠습니까?", vbYesNo)
If YN = vbNo Then shp.Delete: Exit Sub

Delete_Record shtInventory, shtMain.Range("F" & cRow).Value
shp.Delete
shtMain.Range(("F" & cRow & ":U" & cRow)).Delete
MsgBox "재고 정보가 삭제되었습니다."

End Sub























