Option Explicit

'=============================================
' QuickSort - 배열의 마지막 컬럼을 정렬키로 사용 (오름차순)
'=============================================
Sub 퀵정렬(arr As Variant, ByVal lngStart As Long, ByVal lngEnd As Long)
    If lngStart >= lngEnd Then Exit Sub
    
    Dim i As Long, j As Long, k As Long
    Dim keyCol As Long
    keyCol = UBound(arr, 2)  ' ★ 정렬키 = 항상 마지막 컬럼
    
    Dim pivot As String
    Dim temp As Variant
    
    i = lngStart
    j = lngEnd
    pivot = arr((lngStart + lngEnd) \ 2, keyCol)
    
    Do
        Do While arr(i, keyCol) < pivot: i = i + 1: Loop
        Do While arr(j, keyCol) > pivot: j = j - 1: Loop
        
        If i <= j Then
            For k = 1 To keyCol   ' ★ 마지막 컬럼(정렬키)까지 전부 swap
                temp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = temp
            Next k
            i = i + 1
            j = j - 1
        End If
    Loop Until i > j
    
    If lngStart < j Then Call 퀵정렬(arr, lngStart, j)
    If i < lngEnd Then Call 퀵정렬(arr, i, lngEnd)
End Sub

'=============================================
' 분개장정렬 - 입력 컬럼 수에 비종속적으로 정렬
' 정렬키: 세목코드(9) + 회계일자(2) + 전표번호(3) + 명세번호(4)
'=============================================
Function 분개장정렬(db As Variant) As Variant
    Dim i As Long, k As Long
    Dim nCols As Long
    nCols = UBound(db, 2)   ' ★ 입력 데이터의 실제 컬럼 수 (10이든 13이든)
    
    Dim sortData() As Variant
    ReDim sortData(1 To UBound(db, 1), 1 To nCols + 1)  ' ★ 마지막 컬럼은 정렬키용
    
    For i = 1 To UBound(db, 1)
        For k = 1 To nCols
            sortData(i, k) = db(i, k)
        Next k
        sortData(i, nCols + 1) = Format(db(i, 9), "0") & _
                          Format(db(i, 2), "00000") & _
                          IIf(db(i, 3) = "", "00000000000000", CStr(db(i, 3))) & _
                          Format(db(i, 4), "000")
    Next i
    
    Call 퀵정렬(sortData, 1, UBound(sortData, 1))
    
    분개장정렬 = sortData
End Function
