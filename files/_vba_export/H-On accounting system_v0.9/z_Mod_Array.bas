Option Explicit

Sub ArrayToRng(startRng As Range, arr As Variant, Optional ColumnNo As String = "")

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ArrayToRng 함수
'▶ 배열을 범위 위로 반환합니다.
'▶ 인수 설명
'_____________startRng      : 배열을 반환할 기준 범위(셀) 입니다.
'_____________Arr               : 반환할 배열입니다.
'_____________ColumnNo   : [선택인수] 배열의 특정 열을 선택하여 범위로 반환합니다. 여러개 열을 반환할 경우 열 번호를 쉼표로 구분하여 입력합니다.
'                                               값으로 공란을 입력하면 열을 건너뜁니다.
'▶ 사용 예제
'Dim v As Variant
'ReDim v(0 to 1)
''v(0) = "a" : v(1) = "b"
'ArrayToRng Sheet1.Range("A1"), v
'▶ 사용된 보조 명령문
'Extract_Column 함수
'##############################################################

On Error GoTo SingleDimension:

Dim Cols As Variant: Dim col As Variant
Dim X As Long: X = 1
If ColumnNo = "" Then
    startRng.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1) = arr
Else
    Cols = Split(ColumnNo, ",")
    For Each col In Cols
        If Trim(col) <> "" Then
            startRng.Cells(1, X).Resize(UBound(arr, 1) - LBound(arr, 1) + 1) = Extract_Column(arr, CLng(Trim(col)))
        End If
        X = X + 1
    Next
End If
Exit Sub

SingleDimension:
Dim tempArr As Variant: Dim i As Long
ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), 1 To 1)
For i = LBound(arr, 1) To UBound(arr, 1)
    tempArr(i, 1) = arr(i)
Next
startRng.Cells(1, 1).Resize(UBound(arr, 1) - LBound(arr, 1) + 1, 1) = tempArr

End Sub

Sub SequenceToRng(startRng As Range, Count As Long, Optional StartNo As Double = 1, Optional Increment As Double = 1, Optional ToRight As Boolean = False)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ SequenceToRng 함수
'▶ 순번을 범위로 반환합니다.
'▶ 인수 설명
'_____________startRng      : 배열을 반환할 기준 범위(셀) 입니다.
'_____________Count          : 배열로 반환할 순번의 갯수입니다.
'_____________StartNo        : [선택인수] 순번의 시작 번호입니다. 기본값은 1 입니다.
'_____________Increment    : [선택인수]순번이 증가 또는 감소하는 차이값입니다. 기본값은 1 입니다.
'_____________ToRight        : [선택인수] True일 경우 순번을 오른쪽 방향으로 반환합니다. 기본값은 False(=아래방향)입니다.
'▶ 사용 예제
'SequenceToRng Range("A1")
'##############################################################

Dim arr As Variant: Dim v As Double: v = StartNo - Increment

If ToRight = False Then ReDim arr(1 To Count, 1 To 1) Else ReDim arr(1 To 1, 1 To Count)

If ToRight = False Then
    For i = 1 To Count
        v = v + Increment
        arr(i, 1) = v
    Next
Else
    For i = 1 To Count
        v = v + Increment
        arr(1, i) = v
    Next
End If

If ToRight = False Then startRng.Cells(1, 1).Resize(Count) = arr Else startRng.Cells(1, 1).Resize(1, Count) = arr

End Sub

Sub ValueToRng(startRng As Range, Count As Long, Value, Optional ToRight As Boolean = False)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ValueToRng 함수
'▶ 고정된 값을 범위에 뿌려줍니다.
'▶ 인수 설명
'_____________startRng      : 값을 뿌려줄 기준 셀 입니다.
'_____________Count          : 뿌려줄 값의 갯수입니다.
'_____________Value           : 뿌려줄 값 입니다.
'_____________ToRight        : [선택인수] True일 경우 값을 오른쪽 방향으로 뿌려줍니다. 기본값은 False(=아래방향)입니다.
'▶ 사용 예제
'ValueToRng Range("A1"), 10, "A"  '<- A1:A10 범위에 "A"를 출력합니다.
'##############################################################

If ToRight = False Then startRng.Cells(1, 1).Resize(Count) = Value Else startRng.Cells(1, 1).Resize(1, Count) = Value

End Sub

Sub RunningSumRng(startRng As Range, Count As Long, _
                                    Optional Offset_Add As Long = -1, Optional Offset_Deduct As Long = 0, _
                                    Optional blnReverse As Boolean = False)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ RunningSumRng 함수
'▶ 시작셀을 기준으로 더할위치/뺄위치를 지정하여 누계를 계산합니다.
'▶ 인수 설명
'_____________startRng               : 누계를 계산할 기준 셀입니다.
'_____________Count                   : 누계를 계산할 범위의 갯수입니다.
'_____________Offset_Add          : 더할 값이 입력된 열의 상대 위치입니다. (음수는 왼쪽방향, 양수는 오른쪽방향입니다. 기본값은 -1, 한칸 왼쪽에 있는 값을 더합니다.)
'_____________Offset_Deduct     : [선택인수] 뺄 값이 입력된 열의 상대위치입니다. 기본값은 없음입니다.
'_____________blnReverse            : [선택인수] True일 경우 시작셀을 맨 아래로 부터 시작하여 위로 올라가며 누계를 계산합니다. 기본값은 False 입니다.
'▶ 사용 예제
'RunningSumRng Range("C1"), 10,   '<- C1:C10 범위에 B1:B10을 참조하여 누계를 계산합니다.
'##############################################################

Dim t As Double
Dim vArr As Variant
Dim fR As Single: Dim fS As Long: Dim fE As Long

If blnReverse = False Then fR = 1 Else fR = -1
If Count < 1 Then Count = 1

ReDim vArr(1 To Count, 1 To 1)
    If fR = 1 Then fS = 1: fE = Count Else fS = Count: fE = 1
    
    If Offset_Deduct <> 0 Then
        For i = 1 To Count
            t = t + startRng.Offset((i - 1) * fR, Offset_Add).Value - startRng.Offset((i - 1) * fR, Offset_Deduct).Value
            vArr(i, 1) = t
        Next
    Else
        For i = 1 To Count
            t = t + startRng.Offset((i - 1) * fR, Offset_Add).Value
            vArr(i, 1) = t
        Next
    End If

If fR = 1 Then
    startRng.Resize(Count) = vArr
Else
    fE = UBound(vArr, 1)
    fS = LBound(vArr, 1)
    For i = fS To (fE - fS) \ 2 + fS
        t = vArr(fE, 1)
        vArr(fE, 1) = vArr(i, 1)
        vArr(i, 1) = t
        fE = fE - 1
    Next
    startRng.Offset(-Count + 1).Resize(Count) = vArr
End If

End Sub

Function IsUniqueArray(arr As Variant, Optional ColNo As String = "") As Boolean

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ IsUniqueArray 함수
'▶ 배열에서 지정한 열이 하나의 값으로만 입력되어 있는지 확인합니다.
'▶ 인수 설명
'_____________Arr               : 배열입니다.
'_____________ColNo          : [선택인수] 하나의 값만 입력되어 있는지 확인할 열 번호입니다. 쉼표로 구분하여 여러개의 열을 지정할 수 있습니다. AND 조건으로 고유여부를 조회합니다. 기본값은 배열의 최초 열번호입니다.
'▶ 사용 예제
'Debug.Print IsUniqueArray (DB, 2)  '<- 배열의 두번째 열의 값이 하나로만 이루어져있는지 확인합니다.
'▶ 사용된 보조함수
'ArrayDimension 함수
'##############################################################

Dim D As Long: Dim i As Long
Dim c As Variant
Dim vCols As Variant: Dim vCol As Variant
Dim sTemp As String

D = ArrayDimension(arr)

If ColNo = "" Then If D > 1 Then ColNo = LBound(arr, 2)
vCols = Split(ColNo, ",")

If D = 1 Then
    c = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> c Then IsUniqueArray = False: Exit Function
    Next
Else
    For Each vCol In vCols
        sTemp = sTemp & arr(LBound(arr, 1), CLng(Trim(vCol)))
    Next
    c = sTemp
    For i = LBound(arr, 1) To UBound(arr, 1)
        sTemp = ""
        For Each vCol In vCols
            sTemp = sTemp & arr(i, CLng(Trim(vCol)))
        Next
        If c <> sTemp Then IsUniqueArray = False: Exit Function
    Next
End If

IsUniqueArray = True

End Function
Function IsDistinctArray(arr As Variant, Optional ColNo As String = "") As Boolean

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ IsDistinctArray 함수
'▶ 배열에서 지정한 열의 값들이 모두 고유한 값들만 입력되어 있는지 확인합니다.
'▶ 인수 설명
'_____________Arr               : 배열입니다.
'_____________ColNo          : [선택인수] 값들이 고유한지 여부를 조회할 열 번호입니다. 쉼표로 구분하여 여러개의 열을 지정할 수 있습니다. AND 조건으로 고유여부를 조회합니다. 기본값은 배열의 최초 열번호입니다.
'▶ 사용 예제
'Debug.Print IsDistinctArray(DB, 2)  '<- 배열의 두번째 열의 값들이 고유한지 검사합니다.
'▶ 사용된 보조함수
'ArrayDimension 함수
'##############################################################

Dim Dict As Dictionary
Dim vCols As Variant: Dim vCol As Variant
Dim sTemp As String: Dim i As Long

Set Dict = New Dictionary

If ColNo = "" Then If ArrayDimension(arr) > 1 Then ColNo = LBound(arr, 2)

vCols = Split(ColNo, ",")

On Error GoTo Duplicate:

If ArrayDimension(arr) = 1 Then
    For i = LBound(arr) To UBound(arr)
        Dict.Add arr(i), 0
    Next
Else
    For i = LBound(arr, 1) To UBound(arr, 1)
        sTemp = ""
        For Each vCol In vCols
            sTemp = sTemp & arr(i, CLng(Trim(vCol)))
        Next
        Dict.Add sTemp, 0
    Next
End If

IsDistinctArray = True

Exit Function

Duplicate:
    IsDistinctArray = False

End Function

Sub ClearContentsBelow(startRng As Range, Optional ColNo, Optional BaseCol As Long = 0)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ClearContentsBelow 함수
'▶ 기준셀부터 특정열까지 아래로 입력된 데이터를 모두 초기화합니다. (값만 초기화되며 서식은 그대로 유지됩니다.)
'▶ 인수 설명
'_____________startRng      : 기준셀입니다.
'_____________ColNo          : [선택인수] 기준셀로부터 삭제될 열번호(또는 알파벳)입니다. 기본값은 기준셀로부터 연속된 범위의 우측 마지막셀 열번호를 반환합니다.
'_____________BaseCol       : [선택인수] 기준셀 아래로 연속된 데이터를 참조할 기준열번호 입니다. 기본값은 기준셀의 열번호입니다.
'▶ 사용 예제
'ClearContentsBelow Range("A5"), "F"   '<- A5~F열까지 아래로 입력된 데이터를 초기화합니다.
'##############################################################

Dim ws As Worksheet: Dim LastRow As Long: Set ws = startRng.Parent
If IsMissing(ColNo) Then ColNo = ws.Cells(startRng.Row, ws.Columns.Count).End(xlToLeft).Column
If Not IsNumeric(ColNo) Then ColNo = Range(ColNo & 1).Column
If BaseCol = 0 Then BaseCol = startRng.Column Else BaseCol = startRng.Column + BaseCol - 1
LastRow = ws.Cells(ws.Rows.Count, BaseCol).End(xlUp).Row
If LastRow < startRng.Row Then Exit Sub
ws.Range(startRng, ws.Cells(LastRow, ColNo)).ClearContents

End Sub

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ArrayDimension 함수
'▶ 배열의 차원수를 반환합니다.
'▶ 인수 설명
'_____________vaArray     : 차원을 검토할 배열을 입력합니다.
'###############################################################
Function ArrayDimension(vaArray As Variant) As Integer
 
Dim i As Integer: Dim X As Integer
 
On Error Resume Next
 
Do
    i = i + 1
    X = UBound(vaArray, i)
Loop Until Err.Number <> 0
 
Err.Clear
 
ArrayDimension = i - 1
 
End Function
