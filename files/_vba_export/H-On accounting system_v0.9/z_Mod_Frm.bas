Option Explicit

Public arrIDs() As Variant    '---> 반제정보 갱신을 위한 전역 배열변수 선언  '(Public Sub 반제_업데이트)


'########################
' 콤보박스를 DB 값으로 갱신
' Update_Cbo cboBox, DB, "1"   '지금 함수는 DisplayCol에 지정된 컬럼만 보이고 나머지는 0 폭으로 숨기는 구조야.
'########################
Sub Update_Cbo(cboBox As MSForms.ComboBox, db As Variant, Optional DisplayCol As Long = 1, Optional SetDefault As Boolean = False)

Dim colCount As Long
Dim colWidths As String
Dim i As Long

colCount = UBound(db, 2)

With cboBox
    .ColumnCount = colCount
    For i = 1 To colCount
        If DisplayCol = i Then colWidths = colWidths & .Width - 15 & "," Else colWidths = colWidths & "0,"
    Next
    colWidths = left(colWidths, Len(colWidths) - 1)
    .List = db
    .ColumnWidths = colWidths
    If SetDefault = True Then .ListIndex = 0
End With

End Sub

'########################
' 콤보박스의 특정 필드 값을 참조하여 값을 선택
' Select_CboItm cboBox, 1, 1
'########################
Sub Select_CboItm(cboBox As MSForms.ComboBox, ID, Optional ColNo As Long = 1)

Dim i As Long

If IsNumeric(ID) Then ID = CLng(ID)

With cboBox
    For i = 0 To .ListCount - 1
        If .List(i, ColNo - 1) = ID Then .ListIndex = i
    Next
End With

End Sub

'########################
' 리스트박스를 DB 값으로 갱신
' Update_List ListBox, DB, "0pt; 80pt; 50pt"
'########################
Sub Update_List(lstBox As MSForms.ListBox, db As Variant, Widths As String)

With lstBox
    .Clear
    .ColumnWidths = Widths
    If Not IsEmpty(db) Then
        .ColumnCount = UBound(db, 2)
        .List = db
    End If
End With

End Sub

'########################
' 리스트박스의 아이템 목록을 배열로 반환
' Array = Get_ListItm listbox
'########################
Function Get_ListItm(lstBox As Control) As Variant

Dim i As Long: Dim j As Long
Dim vaArr As Variant

With lstBox
    If .ListIndex <> -1 Then
    ReDim vaArr(0 To .ColumnCount - 1)
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                For j = 0 To .ColumnCount - 1
                    vaArr(j) = .List(i, j)
                Next
                Exit For
            End If
        Next
    End If
End With

Get_ListItm = vaArr

End Function

'########################
' 리스트박스의 첫번째필드 ID를 참조하여 해당 ID 값을 선택
' Select_ListItm ListBox, ID
'########################
Function Select_ListItm(lstBox As Control, ID, Optional ColNo As Long = 1)

Dim i As Long

If IsNumeric(ID) Then ID = CLng(ID)

With lstBox
    For i = 0 To .ListCount - 1
        If .List(i, ColNo - 1) = ID Then .Selected(i) = True: Exit For
    Next
End With

End Function

'########################
' 리스트박스 활성화
' Active_ListBox ( ListBox, Select_ListItm(ListBox, ID) )
'########################
Function Active_ListBox(lstBox As Control, Optional Index As Long = 0)

If lstBox.ListCount > 0 Then lstBox.Selected(Index) = True

End Function

'########################
' 현재 선택된 값의 순번 확인
' i = Get_ListIndex(ListBox)
'########################
Function Get_ListIndex(lstBox As Control)

Dim i As Long
With lstBox
    If .ListIndex <> -1 Then
        For i = 0 To .ListCount - 1
            If .Selected(i) Then Get_ListIndex = i: Exit For
        Next
    End If
End With

End Function

'########################
' 리스트 박스가 선택되어 있는지 여부 확인
' boolean = isListBoxSelected(ListBox1)
'########################

Function isListBoxSelected(ListBox As MSForms.ListBox) As Boolean
 
Dim i As Long
 
For i = 0 To ListBox.ListCount - 1
If ListBox.Selected(i) Then isListBoxSelected = True: Exit Function
Next
 
isListBoxSelected = False
 
End Function
 
 '########################
' 유저폼에서 해당 컨트롤 버튼 값 초기화
' Clear_Ctrls ( Userform1, "Label", "이름" )  ' 유저폼에서 "이름"이 들어가는 라벨 제외 모든 label 제거
' 컨트롤 이름에는 와일드카드(*,?) 사용가능 (예: txt* 는 txt로 시작하는 모든 버튼을 의미)
' 컨트롤 종류 :
' Label, Frame, TextBox, CommandButton, ComboBox, TabStrip, ListBox,
' MultiPage, CheckBox, ScrollBar, OptionButton, SpinButton, ToggleButton, Image
'########################
Sub Clear_Ctrls(frm As UserForm, CtlType As String, Optional Exclude As String)

Dim ctl As Control
Dim Excs As Variant: Dim Exc As Variant
Dim blnPass As Boolean
Dim vaType As Variant: Dim vType As Variant

If InStr(1, Exclude, ",") > 0 Then: Excs = Split(Exclude, ","): Else Excs = Array(Exclude)
If InStr(1, CtlType, ",") > 0 Then: vaType = Split(CtlType, ","): Else vaType = Array(CtlType)

For Each vType In vaType
    For Each ctl In frm.Controls
        If ctl.Name Like Trim(vType) Then
            blnPass = False
            For Each Exc In Excs
                If ctl.Name Like Trim(Exc) Then blnPass = True: Exit For
            Next
            If blnPass = False Then ctl.Value = ""
        End If
    Next
Next

End Sub

 '########################
' 유저폼의 컨트롤 중 비어있는 컨트롤이 있는지 확인(오류방지)
' blnCheck = IsEmpty_Ctrls ( Userform1, "Label", "이름" )  ' 유저폼에서 "이름"이 들어가는 라벨 제외 모든 label 제거
' 컨트롤 이름에는 와일드카드(*,?) 사용가능 (예: txt* 는 txt로 시작하는 모든 버튼을 의미)
' 컨트롤 종류 :
' Label, Frame, TextBox, CommandButton, ComboBox, TabStrip, ListBox,
' MultiPage, CheckBox, ScrollBar, OptionButton, SpinButton, ToggleButton, Image
'########################
Function IsEmpty_Ctrls(frm As UserForm, CtlType As String, Optional Exclude As String)

Dim ctl As Control
Dim vaType As Variant: Dim vType As Variant

If InStr(1, CtlType, ",") > 0 Then: vaType = Split(CtlType, ","): Else vaType = Array(CtlType)

For Each vType In vaType
    For Each ctl In frm.Controls
        If ctl.Name Like Trim(vType) And ctl.Name <> Exclude Then
            If ctl.Value = "" Then IsEmpty_Ctrls = True: Exit Function
        End If
    Next
Next

IsEmpty_Ctrls = False

End Function

'---------------------------------혜정-----------------------------------

Sub Update계정코드검색(cboBox As MSForms.ComboBox, db As Variant, Optional DisplayCol As Long = 1, Optional SetDefault As Boolean = False)
''Access DB의 ID 컬럼은 자동으로 제외하여 콤보박스에 표시한다.

    Dim newDB() As Variant
    Dim r As Long, c As Long
    Dim colCount As Long
    Dim colWidths As String
    Dim i As Long

    '=========================
    ' ID 컬럼(1열) 제거
    '=========================
    ReDim newDB(1 To UBound(db, 1), 1 To UBound(db, 2) - 1)

    For r = 1 To UBound(db, 1)
        For c = 2 To UBound(db, 2)
            newDB(r, c - 1) = db(r, c)
        Next c
    Next r

    colCount = UBound(newDB, 2)

    With cboBox
        .ColumnCount = colCount

        For i = 1 To colCount
            If i = 1 Or i = 2 Then
                colWidths = colWidths & "60,"
            Else
                colWidths = colWidths & "0,"
            End If
        Next i

        colWidths = left(colWidths, Len(colWidths) - 1)

        .List = newDB
        .ListWidth = 195
        .ColumnWidths = colWidths

        If SetDefault Then .ListIndex = 0
    End With

End Sub


'---------------------------------혜정-----------------------------------

Sub Update은행코드검색(cboBox As MSForms.ComboBox, db As Variant, Optional DisplayCol As Long = 1, Optional SetDefault As Boolean = False)
''Access DB의 ID 컬럼은 자동으로 제외하여 콤보박스에 표시한다.

    Dim newDB() As Variant
    Dim r As Long, c As Long

    ' ID 컬럼(1열) 제거
    ReDim newDB(1 To UBound(db, 1), 1 To UBound(db, 2) - 1)

    For r = 1 To UBound(db, 1)
        For c = 2 To UBound(db, 2)
            newDB(r, c - 1) = db(r, c)
        Next c
    Next r

    With cboBox
        .ColumnCount = 2
        .List = newDB
        .ListWidth = 190
        .ColumnWidths = "140;40"

        If SetDefault Then .ListIndex = 0
    End With

End Sub


'-------------------------------------------혜정--------------------------


'전표입력 화면의 차대합계 일치를 확인
Public Sub 차대합계_일치_Public(ByVal targetForm As Object) '“그냥 Sub와 Public Sub는 똑같이 동작하지만, Public Sub는 의도를 더 명확히 보여주는 표현이다.”

Dim DebitAmount As Currency
Dim CreditAmount As Currency
Dim DebitTotal As Currency
Dim CreditTotal As Currency

Dim DebitTotalFx As Double
Dim CreditTotalFx As Double

Dim fxAmount As Double
Dim i As Long


With targetForm       ' ---->frm인쇄(전포복사)에서 쓰기 위해 with문으로 지정하고 공용 모듈에 저장.


    For i = 0 To .lstMain.ListCount - 1
    
        ' 외화금액 (공통 필드)
        'fxAmount = Val(.lstMain.List(i, 8))  '-> 94 런타임 에러(Null값)
        fxAmount = Val(.lstMain.List(i, 8) & "")
        
        
        ' ? 차변 처리
        If .lstMain.List(i, 5) <> "" And Val(.lstMain.List(i, 5)) <> 0 Then
            
            DebitAmount = CDbl(.lstMain.List(i, 5))
            DebitTotal = DebitTotal + DebitAmount
            
            ' 외화도 차변으로 누적
            DebitTotalFx = DebitTotalFx + fxAmount

        ' ? 대변 처리
        ElseIf .lstMain.List(i, 6) <> "" And Val(.lstMain.List(i, 6)) <> 0 Then
            
            CreditAmount = CDbl(.lstMain.List(i, 6))
            CreditTotal = CreditTotal + CreditAmount
            
            ' 외화도 대변으로 누적
            CreditTotalFx = CreditTotalFx + fxAmount

        End If

    Next i

    ' ? 원화 합계 표시
    .txt차변합계.Value = Format(DebitTotal, "#,##0")
    .txt대변합계.Value = Format(CreditTotal, "#,##0")

    ' ? 외화 합계 표시
    If DebitTotalFx <> 0 Then
        .lbl외화차변.Caption = "(" & .cbo통화.Value & " " & Format(DebitTotalFx, "#,##0") & ")"
    Else
        .lbl외화차변.Caption = ""
    End If

    If CreditTotalFx <> 0 Then
        .lbl외화대변.Caption = "(" & .cbo통화.Value & " " & Format(CreditTotalFx, "#,##0") & ")"
    Else
        .lbl외화대변.Caption = ""
    End If


End With

End Sub


Public Sub 반제전표(frmSource As UserForm)

    Dim r As Integer
    Dim c As Integer
    Dim maxRow As Integer
    Dim LastRow As Integer

    r = frmSource.lstMain.ListIndex
    maxRow = frmSource.lstMain.ListCount - 1

    For c = 0 To maxRow
        frm반제전표.lstMain.AddItem
        LastRow = frm반제전표.lstMain.ListCount - 1
        
        ' ★★★ 바로 여기! 0번째 열(ID)을 넘겨주는 코드를 추가했어! ★★★　＇－＞반제전표입력폼으로 넘기고 명세를 삭제하는 경우, arrIDs변수를 갱신할 때 ID행열이 필요했음.(추가수정)
        frm반제전표.lstMain.List(LastRow, 0) = frmSource.lstMain.List(c, 0)
        frm반제전표.lstMain.List(LastRow, 1) = frmSource.lstMain.List(c, 8)
        frm반제전표.lstMain.List(LastRow, 2) = frmSource.lstMain.List(c, 9)
        frm반제전표.lstMain.List(LastRow, 3) = frmSource.lstMain.List(c, 4)
        frm반제전표.lstMain.List(LastRow, 4) = frmSource.lstMain.List(c, 5)
        frm반제전표.lstMain.List(LastRow, 8) = frmSource.lstMain.List(c, 12)   '->외화금액도 넘기기


        If frm채권채무조회.opt채무.Value = True Then  '채무라면
            frm반제전표.lstMain.List(LastRow, 5) = Format(frmSource.lstMain.List(c, 7), "#,##0")
            frm반제전표.lstMain.List(LastRow, 6) = 0

        ElseIf frm채권채무조회.opt채권.Value = True Then  '채권이라면
            frm반제전표.lstMain.List(LastRow, 5) = 0
            frm반제전표.lstMain.List(LastRow, 6) = Format(frmSource.lstMain.List(c, 6), "#,##0")
        
        End If
        
        frm반제전표.lstMain.List(LastRow, 7) = frmSource.lstMain.List(c, 11)

        frm반제전표.cbo통화.Value = frmSource.lstMain.List(c, 11)
    
    Next c


End Sub


Public Sub 반제_업데이트(arrIDs() As Variant, voucherDate As String, printNo As String)

Dim ws As Worksheet
Dim i As Long
Dim findCell As Range

Set ws = sht채권채무명세

For i = LBound(arrIDs) To UBound(arrIDs)
    Set findCell = ws.Range("A:A").Find(What:=arrIDs(i), LookAt:=xlWhole)

    If Not findCell Is Nothing Then
        ws.Cells(findCell.Row, 15).Value = ws.Cells(findCell.Row, 15).Value & "반제"
        ws.Cells(findCell.Row, 16).Value = voucherDate & "-" & printNo   ' P열 = 16번째 열 / 반제전표번호를 갱신!
    End If
Next i

End Sub



Public Sub 반제_업데이트_DB(Conn As Object, arrIDs() As Variant, voucherDate As String, voucherNo As String)

    Dim i As Long
    Dim Sql As String

    For i = LBound(arrIDs) To UBound(arrIDs)

        Sql = "UPDATE 채권채무명세 SET " & _
              "상태 = 상태 & '반제', " & _
              "반제전표 = '" & voucherNo & "' " & _
              "WHERE ID = " & arrIDs(i)

        Debug.Print Sql
        Conn.Execute Sql

    Next i

End Sub




Sub 즐겨찾기불러오기(FavName As String)
'즐겨찾기 전표 전표입력유저폼 리스트 박스에 넣기  (즐겨찾기 전표는 Only KRW)

    Dim db As Variant
    Dim i As Long

    db = Get_DB_Access_Where("즐겨찾기", _
                             "즐겨찾기명='" & FavName & "' ORDER BY 명세번호")

    frmAcctInput.lstMain.Clear

    For i = 1 To UBound(db, 1)

        frmAcctInput.lstMain.AddItem

        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 1) = db(i, 4)   '세목코드
        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 2) = db(i, 5)   '계정명
        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 3) = db(i, 6)   '적요
        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 4) = db(i, 7)   '거래처
        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 5) = db(i, 8)   '차변
        frmAcctInput.lstMain.List(frmAcctInput.lstMain.ListCount - 1, 6) = db(i, 9)   '대변

    Next i

    frmAcctInput.txt즐겨찾기.Value = FavName
  

End Sub


' 공용 함수: 컨트롤 활성/비활성 + 색상 변경
Public Sub ToggleControl(ctrl As MSForms.Control, enableFlag As Boolean)
    If enableFlag Then
        ctrl.Enabled = True
        ctrl.BackColor = vbWhite          ' 활성화 시 흰색
        ctrl.ForeColor = vbBlack          ' 글자색 검정
    Else
        ctrl.Enabled = False
        ctrl.BackColor = &H80000004       ' 비활성화 시 시스템 회색
        ctrl.ForeColor = vbGrayText       ' 글자색 회색
    End If
End Sub

' 표준 모듈(Module)에 작성하여 공용함수로 사용
Public Sub Update_cbo계정코드(cbo As MSForms.ComboBox)

    Dim db As Variant
    db = Get_DB_Access("계정마스터")
    

''    ' "사용여부"가 "Y"인 것만 필터링
''    db = Filtered_DB(db, "Y", 11, True)

    ' 콤보박스 업데이트
    Update계정코드검색 cbo, db

End Sub


Public Sub ShowStatus(frm As Object, msg As String)
'공용 메시지 함수 만들기
    frm.lblStatus.Caption = msg
    
End Sub