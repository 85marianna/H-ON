Option Explicit

Function Sort2DArray(DB, ByVal Index As Long, Optional ByVal Order As Integer = -1, Optional ByVal ByColumn As Boolean = False, Optional ByVal lngStart As Long = 0, Optional ByVal lngEnd As Long = 0, Optional THRESHOLD As Long = 20)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'수정 및 배포 시 출처를 반드시 명시해야 합니다.
'
'■ Sort2DArray 명령문
'■ 2차원 배열을 오름차순/내림차순 또는 내림차순으로 정렬합니다. 한계점을 설정하여 QuickSort 또는 InsertionSort로 보다 빠르게 정렬할 수 있습니다. 기본 한계점은 20입니다.
'■ 인수 설명
'_____________DB                : 정렬할 배열입니다.
'_____________Index            : DB를 정렬할 기준 순번입니다.
'_____________Order           : [선택인수] 1 이면 내림차순 정렬합니다. 기본값은 -1 (=오름차순) 정렬입니다.
'_____________ByColumn    : [선택인수] True 면 열방향(가로방향) 정렬입니다. 기본값은 FALSE 입니다.
'_____________lngStart        : [선택인수] 정렬을 시작할 시작점입니다. 기본값은 배열의 시작지점입니다.
'_____________lngEnd         : [선택인수] 정렬을 종료할 마지막점입니다. 기본값은 배열의 마지막지점입니다.
'_____________Threshold    : [선택인수] QuickSort와 InsertionSort를 구분할 한계점입니다. 기본값은 20 입니다. Threshold 는 사용되는 데이터의 구성에 따라 다르지만 대부분 10-20 사이의 정수가 사용됩니다.
'■ 반환값
'_____________정렬된 배열을 반환합니다.
'본 명령문은 아래 링크를 참조하여 작성된 명령문입니다.
'https://www.vbforums.com/showthread.php?631366-RESOLVED-Quick-Sort-2D-Array
'###############################################################

Dim i As Long: Dim j As Long: Dim k As Long
Dim Pivot: Dim Temp
Dim Stack(1 To 64) As Long: Dim StackPtr As Long

If lngStart = 0 Then
    If ByColumn = False Then lngStart = LBound(DB, 1) Else lngStart = LBound(DB, 2)
End If

If lngEnd = 0 Then
    If ByColumn = False Then lngEnd = UBound(DB, 1) Else lngEnd = UBound(DB, 2)
End If

'가로방향 정렬
  If ByColumn Then
    ReDim Temp(LBound(DB, 1) To UBound(DB, 1))
    Stack(StackPtr + 1) = lngStart
    Stack(StackPtr + 2) = lngEnd
    StackPtr = StackPtr + 2
    Do
      StackPtr = StackPtr - 2
      lngStart = Stack(StackPtr + 1)
      lngEnd = Stack(StackPtr + 2)
      If lngEnd - lngStart < THRESHOLD Then
        ' 비교 대상의 첫번째 값과 마지막값 차이가 20 미만일 경우 Insertion Sort
        For j = lngStart + 1 To lngEnd
          For k = LBound(DB, 1) To UBound(DB, 1)
            Temp(k) = DB(k, j)
          Next
          Pivot = DB(Index, j)
          For i = j - 1 To lngStart Step -1
            If Order >= 0 Then
              If DB(Index, i) <= Pivot Then Exit For
            Else
              If DB(Index, i) >= Pivot Then Exit For
            End If
            For k = LBound(DB) To UBound(DB)
              DB(k, i + 1) = DB(k, i)
            Next
          Next
          For k = LBound(DB) To UBound(DB)
            DB(k, i + 1) = Temp(k)
          Next
        Next
      Else
        ' 비교 대상의 첫번째 값과 마지막값 차이가 20 이상일 경우 Quick Sort
        i = lngStart: j = lngEnd
        Pivot = DB(Index, (lngStart + lngEnd) \ 2)
        Do
          If Order >= 0 Then
            Do While (DB(Index, i) < Pivot): i = i + 1: Loop
            Do While (DB(Index, j) > Pivot): j = j - 1: Loop
          Else
            Do While (DB(Index, i) > Pivot): i = i + 1: Loop
            Do While (DB(Index, j) < Pivot): j = j - 1: Loop
          End If
          If i <= j Then
            If i < j Then
              For k = LBound(DB) To UBound(DB)
                Temp(k) = DB(k, i)
                DB(k, i) = DB(k, j)
                DB(k, j) = Temp(k)
              Next
            End If
            i = i + 1: j = j - 1
          End If
        Loop Until i > j
        If (lngStart < j) Then
          Stack(StackPtr + 1) = lngStart
          Stack(StackPtr + 2) = j
          StackPtr = StackPtr + 2
        End If
        If (i < lngEnd) Then
          Stack(StackPtr + 1) = i
          Stack(StackPtr + 2) = lngEnd
          StackPtr = StackPtr + 2
        End If
      End If
    Loop Until StackPtr = 0
'세로방향 정렬
Else
    ReDim Temp(LBound(DB, 2) To UBound(DB, 2))
        ' Stack 설정
        Stack(StackPtr + 1) = lngStart
        Stack(StackPtr + 2) = lngEnd
        StackPtr = StackPtr + 2
            Do
                StackPtr = StackPtr - 2
                lngStart = Stack(StackPtr + 1)
                lngEnd = Stack(StackPtr + 2)
                    ' 비교 대상의 첫번째 값과 마지막값 차이가 20 미만일 경우 Insertion Sort
                    If lngEnd - lngStart < THRESHOLD Then
                          For j = lngStart + 1 To lngEnd
                            For k = LBound(DB, 2) To UBound(DB, 2)
                              Temp(k) = DB(j, k)
                            Next
                            Pivot = DB(j, Index)
                            For i = j - 1 To lngStart Step -1
                              If Order >= 0 Then
                                If DB(i, Index) <= Pivot Then Exit For
                              Else
                                If DB(i, Index) >= Pivot Then Exit For
                              End If
                              For k = LBound(DB, 2) To UBound(DB, 2)
                                DB(i + 1, k) = DB(i, k)
                              Next
                            Next
                            For k = LBound(DB, 2) To UBound(DB, 2)
                              DB(i + 1, k) = Temp(k)
                            Next
                          Next
                Else
                    ' 비교 대상의 첫번째 값과 마지막값 차이가 20 이상일 경우 Quick Sort
                    i = lngStart: j = lngEnd
                    Pivot = DB((lngStart + lngEnd) \ 2, Index)
                        Do
                            If Order >= 0 Then
                                Do While (DB(i, Index) < Pivot): i = i + 1: Loop
                                Do While (DB(j, Index) > Pivot): j = j - 1: Loop
                            Else
                                Do While (DB(i, Index) > Pivot): i = i + 1: Loop
                                Do While (DB(j, Index) < Pivot): j = j - 1: Loop
                            End If
                            If i <= j Then
                                  If i < j Then
                                        For k = LBound(DB, 2) To UBound(DB, 2)
                                            Temp(k) = DB(i, k)
                                            DB(i, k) = DB(j, k)
                                            DB(j, k) = Temp(k)
                                        Next
                                  End If
                                    i = i + 1: j = j - 1
                            End If
                        Loop Until i > j
                            If (lngStart < j) Then
                                  Stack(StackPtr + 1) = lngStart
                                  Stack(StackPtr + 2) = j
                                  StackPtr = StackPtr + 2
                            End If
                            If (i < lngEnd) Then
                                  Stack(StackPtr + 1) = i
                                  Stack(StackPtr + 2) = lngEnd
                                  StackPtr = StackPtr + 2
                            End If
                End If
            Loop Until StackPtr = 0
End If
  
Sort2DArray = DB
  
End Function
