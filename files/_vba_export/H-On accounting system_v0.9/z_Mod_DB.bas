Option Explicit
Option Compare Text
'########################
' 특정 워크시트에서 앞으로 추가해야 할 최대 ID번호 리턴 (시트 DB 우측 첫번째 머릿글)
' i = Get_MaxID(Sheet1)
'########################
Function Get_MaxID(ws As Worksheet) As Long
With ws
    Get_MaxID = .Cells(1, .Columns.Count).End(xlToLeft).Value
    .Cells(1, .Columns.Count).End(xlToLeft).Value = .Cells(1, .Columns.Count).End(xlToLeft).Value + 1
End With
End Function

'########################
' 총계정원장 시트에서 전표번호 생성 (시트 DB 우측 두번째 머릿글)
''########################
Function Get_MaxPrintNo(ws As Worksheet) As Long
    With ws
        Get_MaxPrintNo = .Cells(1, .Columns.Count).End(xlToLeft).Offset(0, -1).Value
    .Cells(1, .Columns.Count).End(xlToLeft).Offset(0, -1).Value = .Cells(1, .Columns.Count).End(xlToLeft).Offset(0, -1).Value + 1
    End With
End Function


'########################
' 워크시트에 새로운 데이터를 추가해야 할 열번호 반환
' i = Get_InsertRow(Sheet1)
'########################
Function Get_InsertRow(ws As Worksheet) As Long
With ws:    Get_InsertRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1: End With
End Function
'########################
' 시트의 열 개수 반환 (이번 예제파일에서만 사용)
' i  = Get_ColumnCnt(Sheet1)
'########################
Function Get_ColumnCnt(ws As Worksheet, Optional Offset As Long = -1) As Long
With ws:    Get_ColumnCnt = .Cells(1, .Columns.Count).End(xlToLeft).Column + Offset: End With
End Function
'########################
' 시트에서 특정 ID 의 행 번호 반환 (-> 해당 행 번호 데이터 업데이트)
' i = get_UpdateRow(Sheet1, ID)
'########################
Function get_UpdateRow(ws As Worksheet, ID)
Dim i As Long
Dim cRow As Long
With ws
    cRow = Get_InsertRow(ws) - 1
    For i = 1 To cRow
        If .Cells(i, 1).Value = ID Then get_UpdateRow = i: Exit For
    Next
End With
End Function


'########################
' 특정 시트의 DB 정보를 배열로 반환 (이번 예제파일에서만 사용)
' Array = Get_DB(Sheet1)
' 제목행 제외 (기본값 False)
'########################
Function Get_DB(ws As Worksheet, Optional NoID As Boolean = False, Optional IncludeHeader As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim offCol As Long

If NoID = False Then offCol = -1

With ws
    cRow = Get_InsertRow(ws) - 1
    cCol = Get_ColumnCnt(ws, offCol)
    Get_DB = .Range(.Cells(2 + Sgn(IncludeHeader), 1), .Cells(cRow, cCol))
End With
    
End Function
'########################
'특정 시트에서 지정한 ID의 필드 값 반환 (이번 예제파일 전용)
' Value = Get_Records(Sheet1, ID, "필드명")
'########################
Function Get_Records(ws As Worksheet, ID, Fields)

Dim cRow As Long: Dim cCol As Long
Dim vFields As Variant: Dim vField As Variant
Dim vFieldNo As Variant
Dim i As Long: Dim j As Long


cRow = Get_InsertRow(ws) - 1
cCol = Get_ColumnCnt(ws)

If InStr(1, Fields, ",") > 0 Then vFields = Split(Fields, ",") Else vFields = Array(Fields)
ReDim vFieldNo(0 To UBound(vFields))

With ws
    For Each vField In vFields
        For i = 1 To cCol
            If .Cells(1, i).Value = Trim(vField) Then vFieldNo(j) = i: j = j + 1
        Next
    Next

    For i = 2 To cRow
        If .Cells(i, 1).Value = ID Then
            For j = 0 To UBound(vFieldNo)
                vFieldNo(j) = .Cells(i, vFieldNo(j))
            Next
            Exit For
        End If
    Next
    
Get_Records = vFieldNo

End With

End Function

'########################
' 시트에 새로운 레코드 추가 (반드시 첫번째 값은 ID, 나머지 값 순서대로 입력)
' Insert_Record Sheet1, ID, 필드1, 필드2, 필드3, ..
'########################
Sub Insert_Record(ws As Worksheet, ParamArray vaParamArr() As Variant)

Dim cID As Long
Dim cRow As Long
Dim vaArr As Variant: Dim i As Long: i = 2

With ws
    cRow = Get_InsertRow(ws)
    If InStr(1, .Cells(1, 1).Value, "ID") > 0 Then
        cID = Get_MaxID(ws)
        .Cells(cRow, 1).Value = cID
        For Each vaArr In vaParamArr
            .Cells(cRow, i).Value = vaArr
            i = i + 1
        Next
    Else
        For Each vaArr In vaParamArr
            .Cells(cRow, i - 1).Value = vaArr
            i = i + 1
        Next
    End If
    
End With

End Sub
'########################
' 시트에서 ID 를 갖는 레코드의 모든 값 업데이트 (반드시 첫번째 값은 ID여야 하며, 나머지 값을 순서대로 입력)
' Update_Record Sheet1, ID, 필드1, 필드2, 필드3, ...
'########################
Sub Update_Record(ws As Worksheet, ParamArray vaParamArr() As Variant)

Dim cRow As Long
Dim i As Long
Dim ID As Variant

If IsNumeric(vaParamArr(0)) = True Then ID = CLng(vaParamArr(0)) Else ID = vaParamArr(0)

With ws
    cRow = get_UpdateRow(ws, ID)
    
    For i = 1 To UBound(vaParamArr)
        If Not IsMissing(vaParamArr(i)) Then .Cells(cRow, i + 1).Value = vaParamArr(i)
    Next
    
End With

End Sub
'########################
' 시트에서 ID 를 갖는 레코드 삭제
' Delete_Record Sheet1, ID
'########################
Sub Delete_Record(ws As Worksheet, ID)

Dim cRow As Long

If IsNumeric(ID) = True Then ID = CLng(ID)

With ws
    cRow = get_UpdateRow(ws, ID)
    .Cells(cRow, 1).EntireRow.Delete
End With

End Sub

'########################
' 배열의 외부ID키 필드를 본 시트DB와 연결하여 해당 외부ID키의 연관된 값을 배열로 반환
' Array = Connect_DB(Get_DB(Sheet1),2,Sheet2, "필드1, 필드2, 필드3")
'########################
Function Connect_DB(db As Variant, ForeignID_Fields As Variant, FromWS As Worksheet, Fields As String, Optional IncludeHeader As Boolean = False)

Dim cRow As Long: Dim cCol As Long
Dim vForeignID_Fields As Variant: Dim vForeignID_Field As Variant
Dim ForeignID As Variant
Dim vFields As Variant: Dim vField As Variant
Dim vID As Variant: Dim vFieldNo As Variant
Dim Dict As Object
Dim i As Long: Dim j As Long
Dim AddCols As Long


cRow = UBound(db, 1)
cCol = UBound(db, 2)

If InStr(1, Fields, ",") > 1 Then
    AddCols = Len(Fields) - Len(Replace(Fields, ",", "")) + 1
    vFields = Split(Fields, ",")
Else
    AddCols = 1
    vFields = Array(Fields)
End If

ReDim Preserve db(1 To cRow, 1 To cCol + AddCols)
        

Set Dict = Get_Dict(FromWS)
vID = Dict("ID")


ReDim vFieldNo(0 To UBound(vFields))

For Each vField In vFields
    For i = 1 To UBound(vID)
        If vID(i) = Trim(vField) Then vFieldNo(j) = i: j = j + 1
    Next
Next


If InStr(1, ForeignID_Fields, ",") > 0 Then vForeignID_Fields = Split(ForeignID_Fields, ",") Else vForeignID_Fields = Array(ForeignID_Fields)


For Each vForeignID_Field In vForeignID_Fields
    For i = 1 To cRow
        If IncludeHeader = True And i = 1 Then ForeignID = "ID" Else ForeignID = db(i, Trim(vForeignID_Field))
        If Dict.Exists(ForeignID) Then
            For j = 1 To AddCols
                db(i, cCol + j) = Dict(ForeignID)(vFieldNo(j - 1))

            Next
        End If
    Next
Next

Connect_DB = db
    
End Function
'########################
' 특정 배열에서 Value를 포함하는 레코드만 찾아 다시 배열로 반환
' Array = Filtered_DB(Array, "검색값", False)
'########################
Function Filtered_DB(db, Value, Optional FilterCol, Optional ExactMatch As Boolean = False) As Variant

Dim cRow As Long
Dim cCol As Long
Dim vArr As Variant: Dim s As String: Dim filterArr As Variant:  Dim Cols As Variant: Dim col As Variant: Dim Colcnt As Long
Dim isDateVal As Boolean
Dim vReturn As Variant: Dim vResult As Variant
Dim Dict As Object: Dim dictKey As Variant
Dim i As Long: Dim j As Long
Dim Operator As String

Set Dict = CreateObject("Scripting.Dictionary")

If Value <> "" Then
    cRow = UBound(db, 1)     ' DB 배열의 행 개수
    cCol = UBound(db, 2)     ' DB 배열의 열 개수
    ReDim vArr(1 To cRow)
    For i = 1 To cRow
        s = ""
        For j = 1 To cCol
            s = s & db(i, j) & "|^"
        Next
        vArr(i) = s
    Next
    
    If IsMissing(FilterCol) Then
        filterArr = vArr
    Else
        Cols = Split(FilterCol, ",")
        ReDim filterArr(1 To cRow)
        For i = 1 To cRow
            s = ""
            For Each col In Cols
                s = s & db(i, Trim(col)) & "|^"
            Next
            filterArr(i) = s
        Next
    End If
    
    If left(Value, 2) = ">=" Or left(Value, 2) = "<=" Or left(Value, 2) = "=>" Or left(Value, 2) = "=<" Then
        Operator = left(Value, 2)
        If IsDate(Right(Value, Len(Value) - 2)) Then isDateVal = True
    ElseIf left(Value, 1) = ">" Or left(Value, 1) = "<" Then
        Operator = left(Value, 1)
        If IsDate(Right(Value, Len(Value) - 1)) Then isDateVal = True
    Else: End If
    
    If Operator <> "" Then
        If isDateVal = False Then
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDbl(left(filterArr(i), Len(filterArr(i)) - 2)) > CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDbl(left(filterArr(i), Len(filterArr(i)) - 2)) < CDbl(Right(Value, Len(Value) - 1)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDbl(left(filterArr(i), Len(filterArr(i)) - 2)) >= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDbl(left(filterArr(i), Len(filterArr(i)) - 2)) <= CDbl(Right(Value, Len(Value) - 2)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
            End Select
        Else
            Select Case Operator
                Case ">"
                    For i = 1 To cRow
                        If CDate(left(filterArr(i), Len(filterArr(i)) - 2)) > CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case "<"
                    For i = 1 To cRow
                        If CDate(left(filterArr(i), Len(filterArr(i)) - 2)) < CDate(Right(Value, Len(Value) - 1)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                Case ">=", "=>"
                    For i = 1 To cRow
                        If CDate(left(filterArr(i), Len(filterArr(i)) - 2)) >= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
                 Case "<=", "=<"
                    For i = 1 To cRow
                        If CDate(left(filterArr(i), Len(filterArr(i)) - 2)) <= CDate(Right(Value, Len(Value) - 2)) Then: vArr(i) = left(vArr(i), Len(vArr(i)) - 2): vReturn = Split(vArr(i), "|^"): Dict.Add i, vReturn
                    Next
            End Select
        End If
    Else
        If ExactMatch = False Then
            For i = 1 To cRow
                If filterArr(i) Like "*" & Value & "*" Then
                    vArr(i) = left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    Dict.Add i, vReturn
                End If
            Next
        Else
            For i = 1 To cRow
                If filterArr(i) Like Value & "|^" Then
                    vArr(i) = left(vArr(i), Len(vArr(i)) - 2)
                    vReturn = Split(vArr(i), "|^")
                    Dict.Add i, vReturn
                End If
            Next
        End If
    End If
        
    If Dict.Count > 0 Then
        ReDim vResult(1 To Dict.Count, 1 To cCol)
        i = 1
        For Each dictKey In Dict.Keys
            For j = 1 To cCol
                vResult(i, j) = Dict(dictKey)(j - 1)
            Next
            i = i + 1
        Next
    End If
    
    Filtered_DB = vResult
Else
    Filtered_DB = db
End If

End Function

''---------------------------------------혜정-------------------------------------------
'########################
' 특정 값과 완전일치하는 레코드 제외 (해당 검색값 DB필터링 값에서 제외)
' Array = Filtered_DB_Exclude(Array, "검색값", "컬럼번호")
'########################

Function Filtered_DB_Exclude(db, Value, Optional FilterCol) As Variant

    Dim cRow As Long
    Dim cCol As Long
    Dim vArr As Variant
    Dim filterArr As Variant
    Dim Cols As Variant
    Dim col As Variant
    Dim vReturn As Variant
    Dim vResult As Variant
    Dim Dict As Object
    Dim dictKey As Variant
    Dim i As Long
    Dim j As Long
    Dim s As String

    Set Dict = CreateObject("Scripting.Dictionary")

    If Value <> "" Then

        cRow = UBound(db, 1)
        cCol = UBound(db, 2)

        ReDim vArr(1 To cRow)

        For i = 1 To cRow

            s = ""

            For j = 1 To cCol
                s = s & db(i, j) & "|^"
            Next

            vArr(i) = s

        Next

        If IsMissing(FilterCol) Then

            filterArr = vArr

        Else

            Cols = Split(FilterCol, ",")

            ReDim filterArr(1 To cRow)

            For i = 1 To cRow

                s = ""

                For Each col In Cols
                    s = s & db(i, Trim(col)) & "|^"
                Next

                filterArr(i) = s

            Next

        End If

        For i = 1 To cRow

            ' 완전일치하는 값 제외
            If Not filterArr(i) Like Value & "|^" Then

                vArr(i) = left(vArr(i), Len(vArr(i)) - 2)

                vReturn = Split(vArr(i), "|^")

                Dict.Add i, vReturn

            End If

        Next

        If Dict.Count > 0 Then

            ReDim vResult(1 To Dict.Count, 1 To cCol)

            i = 1

            For Each dictKey In Dict.Keys

                For j = 1 To cCol
                    vResult(i, j) = Dict(dictKey)(j - 1)
                Next

                i = i + 1

            Next

        End If

        Filtered_DB_Exclude = vResult

    Else

        Filtered_DB_Exclude = db

    End If

End Function



'########################
' 각 제품별 잔고수량을 계산합니다.
' DB = Get_Balance(DB, shtInventory, 입고수량열번호, 출고수량열번호, 제품ID열번호)
'########################

Function Get_Balance(db, InventoryWS As Worksheet, ColumnIN, ColumnOUT, ColumnID) As Variant

Dim InventoryDB As Variant
Dim Dict As Dictionary
Dim cRow As Long: Dim cCol As Long
Dim i As Long: Dim cID

If Not IsNumeric(ColumnOUT) Then ColumnOUT = Range(ColumnOUT & 1).Column
If Not IsNumeric(ColumnIN) Then ColumnIN = Range(ColumnIN & 1).Column
If Not IsNumeric(ColumnID) Then ColumnID = Range(ColumnID & 1).Column

cRow = UBound(db, 1)
cCol = UBound(db, 2)
Set Dict = CreateObject("Scripting.Dictionary")

ReDim Preserve db(1 To cRow, 1 To cCol + 1)

For i = 1 To cRow:    Dict.Add db(i, 1), 0: Next
InventoryDB = Get_DB(InventoryWS)

For i = LBound(InventoryDB, 1) To UBound(InventoryDB, 1)
    cID = InventoryDB(i, ColumnID)
    If Dict.Exists(cID) Then
        Dict(cID) = Dict(cID) + InventoryDB(i, CLng(ColumnIN)) - InventoryDB(i, CLng(ColumnOUT))
    End If
Next

For i = LBound(db, 1) To UBound(db, 1)
    db(i, cCol + 1) = Dict(db(i, 1))
Next

Get_Balance = db

End Function

'########################
' 특정 시트의 DB 정보를 Dictionary로 반환 (이번 예제파일에서만 사용)
' Dict = GetDict(Sheet1)
'########################
Function Get_Dict(ws As Worksheet) As Object

Dim cRow As Long: Dim cCol As Long
Dim Dict As Object
Dim vArr As Variant
Dim i As Long: Dim j As Long

Set Dict = CreateObject("Scripting.Dictionary")

With ws
    cRow = Get_InsertRow(ws) - 1
    cCol = Get_ColumnCnt(ws)
    
    For i = 1 To cRow
        ReDim vArr(1 To cCol - 1)
        For j = 2 To cCol
            vArr(j - 1) = .Cells(i, j)
        Next
        Dict.Add .Cells(i, 1).Value, vArr
    Next
End With

Set Get_Dict = Dict


End Function

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Arr_To_Dict 함수
'▶ 범위를 Dictionary 로 변환합니다.
'▶ 인수 설명
'_____________Arr       : Dictionary로 변환할 배열입니다.
'▶ 사용 예제
'Dict = Arr_To_Dict(Arr)
'##############################################################
Function Arr_To_Dict(arr As Variant) As Object

Dim Dict As Object: Dim vArr As Variant
Dim cCol As Long
Dim i As Long: Dim j As Long

Set Dict = CreateObject("Scripting.Dictionary")
cCol = UBound(arr, 2)

For i = LBound(arr, 1) To UBound(arr, 1)
        ReDim vArr(1 To cCol - 1)
        For j = 2 To cCol
            vArr(j - 1) = arr(i, j)
        Next
        Dict.Add arr(i, 1), vArr
Next

Set Arr_To_Dict = Dict

End Function

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Dict_To_Arr 함수
'▶ Dictionary를 범위로 변환합니다.
'▶ 인수 설명
'_____________Dict       : 배열로 변환할 Dictionary 입니다.
'▶ 사용 예제
'Arr = Dict_To_Arr(Dict)
'##############################################################
Function Dict_To_Arr(Dict As Object) As Variant

Dim i As Long: Dim j As Long: Dim dictKey As Variant: Dim cCol As Long
Dim vTest As Variant
i = 1

If Dict.Count > 0 Then
    If IsObject(Dict(Dict.Keys()(0))) Then cCol = UBound(Dict(Dict.Keys()(0))) Else cCol = 1
    ReDim vResult(1 To Dict.Count, 1 To cCol + 1)
    For Each dictKey In Dict.Keys
        vResult(i, 1) = dictKey
        If cCol = 1 Then
            vResult(i, 2) = Dict(dictKey)
        Else
            For j = 2 To cCol + 1
                vResult(i, j) = Dict(dictKey)(j - 1)
            Next
        End If
        i = i + 1
    Next
End If

Dict_To_Arr = vResult
    
End Function
'########################
' 시트의 특정 필드 내에서 추가되는 값이 고유값인지 확인. 고유값일 경우 TRUE를 반환
' boolean = IsUnique(Sheet1, "사과", 1)
'########################
Function IsUnique(db As Variant, uniqueVal, Optional ColNo As Long = 1, Optional Exclude) As Boolean

Dim endRow As Long
Dim i As Long

For i = LBound(db, 1) To UBound(db, 1)
    If db(i, ColNo) = uniqueVal Then
        If Not IsMissing(Exclude) Then
            If Exclude <> uniqueVal Then
                IsUnique = False
                Exit Function
            End If
        Else
            IsUnique = False: Exit Function
        End If
    End If
Next

IsUnique = True

End Function

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Extract_Column 함수
'▶ 배열에서 지정한 열을 추출합니다.
'▶ 인수 설명
'_____________DB        : 특정 열을 추출할 배열입니다.
'_____________Col       : 배열에서 추출할 열의 열번호입니다.
'▶ 사용 예제
'Arr = Extract_Column(Arr, 3) '<- 3번째 열을 추출합니다.
'##############################################################

Function Extract_Column(db As Variant, col As Long) As Variant

Dim i As Long
Dim vArr As Variant

ReDim vArr(LBound(db) To UBound(db), 1 To 1)
For i = LBound(db) To UBound(db)
        vArr(i, 1) = db(i, col)
Next

Extract_Column = vArr

End Function
