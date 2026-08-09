Dim db As Variant

Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt송금일자.Value = frmCalendar.GetDate

End Sub

Private Sub cmd닫기_Click()

Unload Me
End Sub

Private Sub cmd송금리스트인쇄_Click()

Dim ws As Worksheet
Dim newWb As Workbook
Dim remitDate As String
Dim i As Integer
Dim LastRow As Long
Dim rng As Range
Dim Dict As Object
Dim key As Variant
Dim arr
Dim r As Long


' 1. 입력값 유효성 검사 (날짜가 있는지)
If Trim(Me.txt송금일자.Value) = "" Then
    MsgBox "송금일자를 선택해주세요.", vbCritical, "날짜 선택"
    Me.txt송금일자.SetFocus ' 커서를 다시 날짜 입력창으로 보내주는 센스!
    Exit Sub
End If


' 2. 데이터 유효성 검사 (가져온 내역이 있는지)
If Not IsArray(db) Then
    MsgBox "해당일자에 송금 내역이 없습니다.", vbInformation, "내역 없음"
    Exit Sub
End If


db = 연결_DB_Access(db, 6, "거래처마스터", "예금주,은행코드,은행명,계좌번호")


Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;150pt;110pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"


Set newWb = Workbooks.Add
Set ws = newWb.Worksheets(1)

ws.Name = "송금리스트"


'헤더 작성
ws.Range("A1:G1").Value = Array("No.", "이체일", "예금주", "은행코드", "은행명", "계좌번호", "금액(원)")

'헤더 서식
With ws.Range("A1:G1")
    .Interior.Color = RGB(217, 217, 217)
End With

' ★ 은행코드，은행계좌 열을 텍스트로 고정 ★
ws.Columns("D:F").NumberFormat = "@"

Set Dict = CreateObject("Scripting.Dictionary")

'거래처별 소계 집계
For i = 0 To lstMain.ListCount - 1

    key = lstMain.List(i, 16) & "|" & _
          lstMain.List(i, 17) & "|" & _
          lstMain.List(i, 18) & "|" & _
          lstMain.List(i, 19)

    If Dict.Exists(key) Then

        arr = Dict(key)
        arr(0) = arr(0) + CDbl(lstMain.List(i, 6))
        Dict(key) = arr

    Else

        Dict.Add key, Array( _
            CDbl(lstMain.List(i, 6)), _
            lstMain.List(i, 16), _
            lstMain.List(i, 17), _
            lstMain.List(i, 18), _
            lstMain.List(i, 19))

    End If

Next i

'출력
r = 2

For Each key In Dict.Keys

    arr = Dict(key)

    ws.Cells(r, 1).Value = r - 1                  'No.
    ws.Cells(r, 2).Value = Me.txt송금일자.Value    '이체일
    ws.Cells(r, 3).Value = arr(1)                 '예금주
    ws.Cells(r, 4).Value = arr(2)                 '은행코드
    ws.Cells(r, 5).Value = arr(3)                 '은행명
    ws.Cells(r, 6).Value = arr(4)           '계좌번호
    ws.Cells(r, 7).Value = arr(0)                 '합계금액

    r = r + 1

Next key


' A열에서 빈셀이 있는 행 찾기
LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

ws.Range("G2:G" & LastRow).NumberFormat = "#,##0"

' 그 행의 G열에 값 넣기
ws.Cells(LastRow, "G").Value = Me.txt송금액.Value


' ★ 여기서 열 너비 자동 조절! ★
ws.Columns("A:G").AutoFit

    
Set rng = ws.Range("A1:G" & LastRow) ' 표 범위 지정


With rng.Borders
    .LineStyle = xlContinuous   ' 실선
    .Weight = xlThin            ' 얇은 선
    .Color = RGB(0, 0, 0)       ' 검정색
End With

ws.Range("A" & LastRow & ":G" & LastRow).Interior.Color = RGB(217, 217, 217)

remitDate = Replace(ws.Range("B2").Value, "-", "")

Unload Me

End Sub


Private Sub txt송금일자_Change()

Dim total As Double
Dim i As Integer

db = Get_DB_Access("채권채무명세")
db = Filtered_DB(db, "채무반제", "15", True)
db = Filtered_DB(db, "*?", "7", True)   ' "*?" → 최소 한 글자 이상 있는 문자열을 의미  '7열은 차변
db = Filtered_DB(db, Me.txt송금일자.Value, "2", True)

If Not IsEmpty(db) Then
    For i = LBound(db, 1) To UBound(db, 1)
        ' 7번째 열(차변) 값에 콤마 표시 적용
        db(i, 7) = Format(db(i, 7), "#,##0")
        db(i, 7) = Format(db(i, 7), "@@@@@@@@@@@@@@@@@")
        total = total + db(i, 7)
    Next i
End If

Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;150pt;110pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

Me.txt송금액.Value = Format(total, "#,##0")
Me.txt송금건수.Value = Me.lstMain.ListCount


End Sub
