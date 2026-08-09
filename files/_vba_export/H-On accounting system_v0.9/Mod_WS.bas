''Option Explicit
''
''Public arrCustomer As Variant
''
''Sub OpenfrmCustomer()
''frmCustomer.Show
''End Sub
''
''Sub OpenfrmProduct()
''frmProduct.Show
''End Sub
''
''Sub OpenfrmProductSelect()
''frmProductSelect.Show
''End Sub
''
''Sub OpenfrmInventorySearch()
''frmInventorySearch.Show
''End Sub
''
''Sub OpenfrmInventoryEdit()
''frmInventoryEdit.Show
''End Sub
''
''Sub ClearProductData()
''shtMAIN.Range("C5:D15").ClearContents
''shtMAIN.Range("A7:A9").ClearContents
''End Sub
''
''Sub ClearInventoryData()
'''<------ 재고정보 초기화 (기준열 : 순번)  ------->
''ClearContentsBelow shtMAIN.Range("F5"), "U", 5
''End Sub
''
''Sub InventoryDate()
''shtMAIN.Range("C5").Value = frmCalendar.GetDate
''End Sub
''
''Sub Insert_Inventory()
''
''With shtMAIN
''    '<------ 재고 정보 누락되었을 경우 안내메시지 출력 후 명령문 종료  ------->
''    If .Range("C5").Value = "" Then MsgBox "날짜를 입력하세요.", vbInformation: Exit Sub
''    If .Range("C7").Value = "" Then MsgBox "등록할 제품을 선택하세요.", vbInformation: Exit Sub
''    If Not .Range("C10").Value + .Range("C11").Value > 0 Then MsgBox "입/출고 수량을 입력하세요.", vbInformation: Exit Sub
''    '---------------------------------------------------------------------------------------
''    Insert_Record shtInventory, .Range("A7").Value, .Range("C5").Value, CDbl(.Range("C10").Value), CDbl(.Range("C11").Value), .Range("C8").Value, .Range("C12").Value, .Range("C14").Value
''End With
''
''MsgBox "재고 입출고 정보가 등록되었습니다.", vbInformation
''ClearProductData
''
''End Sub
''
''Sub GetCurrentBalance()
''
''' 변수 만들기
''Dim db As Variant
''
''db = Get_DB(shtItem)
''db = Connect_DB(db, 2, shtCustomer, "거래처명,담당자")
''db = Get_Balance(db, shtInventory, 4, 5, 2)
''db = Filtered_DB(db, ">" & 0, 10)
''
''' 오류처리
''If IsEmpty(db) Then MsgBox "현재 보관중인 재고가 없습니다.", vbInformation: Exit Sub
''
''' 범위로 출력
''ClearInventoryData
''ArrayToRng shtMAIN.Range("F5"), db, ",2,1,3,,,8,9,4,5,6,7,,,10"
''SequenceToRng shtMAIN.Range("J5"), UBound(db, 1)
''ValueToRng shtMAIN.Range("K5"), UBound(db, 1), Date
''
'''불필요한 열 숨겨줌
''shtMAIN.Range("R:S").EntireColumn.Hidden = True
''shtMAIN.Range("T:T").EntireColumn.Hidden = False
''
'''안내메시지 출력 후 명령문 종료
''MsgBox "현재 재고 정보 조회를 완료하였습니다.", vbInformation
''
''End Sub
''
''Sub Delete_Inventory()
''
'''변수
''Dim Shp As Shape
''Dim cRow As Long: Dim eRow As Long
''Dim i As Long
''
''Dim YN As VbMsgBoxResult
''
'''변수
''cRow = Selection.Row
''eRow = cRow + Selection.Rows.Count - 1
''
'''오류처리
''If shtMAIN.Range("F" & cRow).Value = "" And shtMAIN.Range("J" & cRow).Value > 0 Then
''    MsgBox "요약된 정보는 삭제 할 수 없습니다.", vbInformation
''    Exit Sub
''End If
''
''If shtMAIN.Range("F" & cRow).Value = "" And shtMAIN.Range("J" & cRow).Value = "" Then
''    MsgBox "올바른 값을 선택하세요.", vbInformation
''    Exit Sub
''End If
''
'''도형 삽입
''Set Shp = ShapeInRange(shtMAIN.Range("F" & cRow & ":U" & eRow))
''
'''안내 문구 출력
''YN = MsgBox("선택된 재고 정보를 정말로 삭제하시겠습니까? 삭제된 정보는 복구가 불가능합니다.", vbYesNo)
''If YN = vbNo Then Shp.Delete: Exit Sub
''
''For i = cRow To eRow
'''데이터 삭제
''Delete_Record shtInventory, shtMAIN.Range("F" & i).Value
''Next
''
'''도형 삭제
''Shp.Delete
''
'''범위 삭제
''shtMAIN.Range("F" & cRow & ":U" & eRow).Delete
''
'''안내 문구 출력
''MsgBox "재고 정보가 삭제되었습니다.", vbInformation
''
''End Sub
''
''
''
'''--------------------혜정----------------------------------------------------------------
''Sub 전표입력폼()
''
''frmAcctInput.Show
''End Sub
''
''Sub 거래처관리()
''
''frm거래처관리.Show
''End Sub
''
''Sub 거래처검색()
''
''frm거래처검색.Show
''End Sub
''
''Sub 채권채무조회()
''
''frm채권채무조회.Show
''End Sub
''
''Sub 송금리스트조회()
''
''frm송금리스트.Show
''End Sub
''
''Sub 반제전표입력폼()
''
''frm반제전표.Show
''
''End Sub
''
''Sub 날짜입력폼()
''
''sht전표.Range("B2").Value = frmCalendar.GetDate
''
''End Sub
''
''Sub 전표입력()    '----> 유저폼말고 엑셀에서 입력된 값이 원장에 전기됨
''
''Dim db As Variant
''Dim ws As Worksheet
''Dim answer As VbMsgBoxResult
''Dim printNo As String
''Dim i As Long
''
''
''If sht전표.Range("B5").Value <> "" Then
''
''    answer = MsgBox("새로운 전표를 입력하시겠습니까?", vbYesNo + vbQuestion, "전표 입력")
''
''    If answer = vbYes Then
''        ' 입력 셀 지우기 (예: B5:D10 범위)
''        sht전표.Range("B2,B5,D5,F5,G5,C8:G17").ClearContents
''        Exit Sub
''    ElseIf answer = vbNo Then
''        Exit Sub
''    End If
''End If
''
''
''Set ws = ThisWorkbook.Sheets("총계정원장")
''
''With ws
''printNo = Format(Get_MaxPrintNo(ws), "000")
''End With
''
''If sht전표.Range("B2") = "" Then MsgBox "전표날짜를 입력하세요.": Exit Sub
''If sht전표.Range("D5") = "" Then MsgBox "부서를 선택하세요.": Exit Sub
''If sht전표.Range("C8") = "" Then MsgBox "전표를 입력하세요.": Exit Sub
''If sht전표.Range("F18") <> sht전표.Range("G18") Then MsgBox "차대변 합계가 일치하지 않습니다.": Exit Sub
''
''
''For i = 8 To 17
''
''    If sht전표.Range("C" & i).Value <> "" Then
''
''
''Insert_Record sht분개장, sht전표.Range("B2"), sht전표.Range("B2") & "-" & printNo, sht전표.Range("A" & i), sht전표.Range("D" & i), sht전표.Range("E" & i), _
''sht전표.Range("F" & i), sht전표.Range("G" & i), , sht전표.Range("B" & i), sht전표.Range("B" & i), sht전표.Range("C" & i), sht전표.Range("D5"), _
''Now(), Application.UserName
''
''    End If
''
''Next i
''
''
''db = Get_DB(sht분개장)
''
''
''
''' A열 기준으로 마지막 행 찾기
''LastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
''
''' 마지막 값 가져오기
''lastValue = ws.Cells(LastRow, "A").Value
''
''
''sht전표.Range("B5") = ws.Cells(LastRow, "C").Value
''sht전표.Range("F5") = ws.Cells(LastRow, "N").Value
''sht전표.Range("G5") = ws.Cells(LastRow, "O").Value
''
''
''answer = MsgBox("전표가 입력되었습니다. " & Chr(10) & sht전표.Range("B2") & "-" & printNo & " 전표를 인쇄하시겠습니까?", vbOKCancel + vbQuestion, "전표 입력 완료")
''
''    ' 확인 버튼을 눌렀을 때만 인쇄 실행
''    If answer = vbOK Then
''        ActiveSheet.PrintOut
''        MsgBox "인쇄가 완료되었습니다.", vbInformation, "알림"
''    End If
''
''End Sub
''
''
''Sub 전표인쇄()
''
''frm전표인쇄.Show
''
''End Sub
''
''
''Sub 전표삭제(ws As Worksheet, 전표번호)
''
''Dim cRow As Long
''
''With ws
''    cRow = get_UpdateRow(ws, 전표번호)
''    MsgBox "삭제할 행은 " & cRow
''
''
''    '.Cells(cRow, 3).EntireRow.Delete
''End With
''
''End Sub

