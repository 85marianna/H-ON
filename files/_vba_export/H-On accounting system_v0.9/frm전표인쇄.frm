Option Explicit


Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt전표일자.Value = frmCalendar.GetDate

If Me.lstMain.ListCount > 0 Then
    Me.lstMain.ListIndex = 0
End If

End Sub

Private Sub cmd닫기_Click()

sht전표.Range("B2,B5,D5,F5,G5,C8:G17").ClearContents
Unload Me
End Sub


Private Sub cmd복사_Click()

Dim db As Variant
r = Me.lstMain.ListIndex

If r = -1 Then
        MsgBox "복사할 전표를 선택해주세요", vbExclamation, "알림"
        Exit Sub
End If

Me.Label1.Caption = "  전표번호                 적요                                                                    거래처                                 차        변                  대        변"

db = Get_DB_Access("분개장")
db = Filtered_DB(db, Me.lstMain.List(r, 2))

Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;160pt;110pt;110pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

If Me.lstMain.ListCount > 0 Then    '' ListBox에서 현재 선택된 항목의 인덱스 번호를 반환(아무 것도 선택되지 않으면 -1을 반환)
    Me.lstMain.ListIndex = 0        '젤 첫번째 줄은 강제 선택
    End If
   
'Me.txt전표번호.Value = Me.lstMain.List(r, 2)
   
   
전표복사
Call 차대합계_일치_Public(frmAcctInput)
전표입력폼

End Sub


Private Sub cmd삭제_Click()

    Dim Conn As Object
    Dim db As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim voucherNo As String
    Dim inputPass As String
    Dim YN As VbMsgBoxResult
    Dim r As Long


    If Me.lstMain.ListIndex = -1 Then
            MsgBox "삭제할 전표를 선택해주세요", vbExclamation, "알림"
            Exit Sub
    End If
    
    ' 2. 바로 비밀번호 입력창 띄우기
    inputPass = InputBox("선택하신 전표를 삭제하시려면" & vbCrLf & "관리자 비밀번호를 입력하세요.", "전표 삭제 보안 확인")
    
    ' 취소를 눌렀거나 아무것도 입력 안 했을 때 처리
    If inputPass = "" Then
        Exit Sub
    End If
    
    ' 전역 변수 ADMIN_PW와 비교
    If inputPass <> ADMIN_PW Then
        MsgBox "비밀번호가 일치하지 않습니다.", vbCritical, "오류"
        Exit Sub
    End If
    
    Me.Label1.Caption = "   적요                                                                     거래처                                차        변                대        변                 계정명"


    r = Me.lstMain.ListIndex
    voucherNo = Me.lstMain.List(r, 2)

    ' 전표 전체 조회
    db = Get_DB_Access("분개장")
    db = Filtered_DB(db, voucherNo)

    Update_List Me.lstMain, db, "0pt;0pt;0pt;0pt;280pt;160pt;110pt;110pt;0pt;100pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

    YN = MsgBox("선택하신 전표를 삭제하시겠습니까? ", vbYesNo, "알림")
    If YN = vbNo Then Exit Sub

    ReDim arr(0 To Me.lstMain.ListCount - 1)

    For i = 0 To Me.lstMain.ListCount - 1
        arr(i) = Me.lstMain.List(i, 0) ' ID
    Next i

    ' DB 연결
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    '  분개장 삭제 + 백업
    For i = LBound(arr) To UBound(arr)

        Conn.Execute "INSERT INTO 삭제전표 SELECT * FROM 분개장 WHERE ID=" & arr(i)
        Conn.Execute "DELETE FROM 분개장 WHERE ID=" & arr(i)

    Next i

    '  채권채무명세 삭제
    db = Get_DB_Access("채권채무명세")
    db = Filtered_DB(db, voucherNo, "3", True)

    If Not IsEmpty(db) Then

        ReDim arr(1 To UBound(db, 1))

        For i = 1 To UBound(db, 1)
            arr(i) = db(i, 1)
        Next i

        For i = LBound(arr) To UBound(arr)
            Conn.Execute "DELETE FROM 채권채무명세 WHERE ID=" & arr(i)
        Next i

    End If
    
    
        ' ===== 채권채무명세 상태 반제전으로 원상복구 =====
    Dim Rs As Object
    Dim Sql As String
    
    Sql = "SELECT ID, 상태 FROM 채권채무명세 WHERE 반제전표 = '" & voucherNo & "'"
    
    Set Rs = CreateObject("ADODB.Recordset")
    Rs.Open Sql, Conn, 1, 3   ' adOpenKeyset, adLockOptimistic
    
    Do While Not Rs.EOF
    
        Dim origStatus As String
        origStatus = Rs.Fields("상태").Value
    
        ' 뒤에 붙은 "반제" 글자 제거 (2글자)
        If Len(origStatus) >= 2 Then
            If Right(origStatus, 2) = "반제" Then
                origStatus = left(origStatus, Len(origStatus) - 2)
            End If
        End If
    
        Conn.Execute "UPDATE 채권채무명세 SET 상태 = '" & origStatus & "', 반제전표 = NULL WHERE ID = " & Rs.Fields("ID").Value
    
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    

    Conn.Close
    Set Conn = Nothing

    MsgBox "전표 DB 삭제 완료"

    UserForm_Initialize


End Sub


Private Sub cmd인쇄_Click()

    Dim ws As Worksheet
    Dim db As Variant
    Dim r As Long
    Dim i As Long
    Dim answer As VbMsgBoxResult
    Dim finalDescription As String
    
    Dim borderRange As Variant
    Dim pdfFile As String


    r = Me.lstMain.ListIndex

    If r = -1 Then
        MsgBox "인쇄할 전표를 선택해주세요", vbExclamation, "알림"
        Exit Sub
    End If
    
    Me.Label1.Caption = "   적요                                                                     거래처                                차        변                대        변                 계정명"

    Set ws = sht전표

    db = Get_DB_Access("분개장")
    db = Filtered_DB(db, Me.lstMain.List(r, 2))
    
    Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;160pt;110pt;110pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"
    Me.txt전표번호.Value = Me.lstMain.List(r, 2)

    answer = MsgBox("선택하신 전표를 인쇄하시겠습니까?", vbYesNo + vbQuestion, "전표 인쇄")

    If answer <> vbYes Then Exit Sub

    If db(1, 12) = "KRW" Then
        finalDescription = ""   ' 외화전표가 아니면 적요는 한줄
    Else
        finalDescription = vbLf & "( " & db(1, 12) & " " & Format(db(1, 13), "#,##0") & " / 환율: " & db(1, 14) & " )"
    End If

    With ws
        ' 기존 데이터 삭제
        .Range("A1:K25").Clear

        ' 서식
        .Columns("A:E").NumberFormat = "@"
        .Columns("F").NumberFormat = "#,##0"
        .Columns("G").NumberFormat = "#,##0"
        .Range("G5").NumberFormat = "@"   ' G5는 텍스트(입력자/입력일시)라서 별도 지정
        
        ' A열 가운데 정렬
        .Columns("A").HorizontalAlignment = -4108

        ' 라벨 + 회색 배경 (한번에 처리)
        Dim labelCell As Variant
        For Each labelCell In Array("A2", "A5", "C5", "E5")
            .Range(labelCell).Interior.Color = RGB(217, 217, 217)
            .Range(labelCell).Font.Bold = True
            .Range(labelCell).HorizontalAlignment = -4108
        Next labelCell
        .Range("A2").Value = "회계일자"
        .Range("A5").Value = "부서명"
        .Range("C5").Value = "전표번호"
        .Range("E5").Value = "입력자/입력일시"

        .Range("A7:G7").Value = Array("No.", "계정코드", "계정명", "적    요", "거래처", "차    변", "대    변")

        ' 공통 헤더 값
        .Range("B2").Value = db(1, 2)
        .Range("B2").HorizontalAlignment = -4108
        .Range("B5").Value = db(1, 11)
        .Range("B5").HorizontalAlignment = -4108
        .Range("D5").Value = db(1, 3)
        .Range("D5").HorizontalAlignment = -4108
        .Range("F5").Value = db(1, 16)
        .Range("G5").Value = db(1, 15)
        .Range("G5").HorizontalAlignment = -4108

        ' 데이터 행 반복
        For i = 1 To UBound(db, 1)
            .Cells(i + 7, 1).Value = i
            .Cells(i + 7, 2).Value = db(i, 9)
            .Cells(i + 7, 3).Value = db(i, 10)
            .Cells(i + 7, 4).Value = db(i, 5) & finalDescription
            .Cells(i + 7, 4).EntireRow.AutoFit
            .Cells(i + 7, 4).RowHeight = .Cells(i + 7, 4).RowHeight + 5
            .Cells(i + 7, 5).Value = db(i, 6)
            .Cells(i + 7, 6).Value = Format(db(i, 7), "#,##0")
            .Cells(i + 7, 7).Value = Format(db(i, 8), "#,##0")
        Next i

        With .Range("A3:G3")
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = -4108 'xlCenter
        End With

        With .Range("A3")
            .Font.Size = 17
            .Font.Bold = True
            .Value = "회  계  전  표"
        End With

        .Rows(3).RowHeight = 60
        .Rows(4).RowHeight = 17
        .Rows(7).RowHeight = 24
        .Rows(7).Font.Bold = True
        .Rows(7).HorizontalAlignment = -4108    'xlCenter
        .Range("A7:G7").Interior.Color = RGB(217, 217, 217)
    End With

    ' 테두리 (3군데 한번에 처리)
    For Each borderRange In Array(ws.Range("A2:B2"), ws.Range("A5:G5"), ws.Range(ws.Cells(7, 1), ws.Cells(i + 7, 7)))
        With borderRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next borderRange

    ' 합계 행
    With ws
        .Cells(i + 7, 5).Value = "합   계"
        .Cells(i + 7, 5).RowHeight = 24
        .Cells(i + 7, 6).Formula = "=SUM(F8:F" & i + 6 & ")"
        .Cells(i + 7, 6).NumberFormat = "#,##0"
        .Cells(i + 7, 7).Formula = "=SUM(G8:G" & i + 6 & ")"
        .Cells(i + 7, 7).NumberFormat = "#,##0"

        With .Range(.Cells(i + 7, 1), .Cells(i + 7, 7))
            .Interior.Color = RGB(217, 217, 217)
            .Font.Bold = True
        End With

        .Cells(i + 10, 1).Value = "한국전기초자㈜"
        .Cells(i + 10, 1).HorizontalAlignment = -4131    'xlLeft

        
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 40
        .Columns("E").ColumnWidth = 25
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 20

        .PageSetup.PrintArea = "A1:G" & i + 10
     
    End With
    
    'Temp 폴더에 생성
    pdfFile = Environ("TEMP") & "\H-ON.pdf"
    
    ' PDF 생성 후 자동 열기
    ExportSheetToPDF sht전표, pdfFile, True
     
End Sub



Private Sub cmd조회_Click()

Dim r As Long

Dim db As Variant
r = Me.lstMain.ListIndex

If r = -1 Then
        MsgBox "조회할 전표를 선택해주세요", vbExclamation, "알림"
        Exit Sub
End If

Me.Label1.Caption = "   적요                                                                     거래처                                차        변                대        변                 계정명"

db = Get_DB_Access("분개장")
db = Filtered_DB(db, Me.lstMain.List(r, 2))

Update_List Me.lstMain, db, "0pt;0pt;0pt;0pt;280pt;160pt;110pt;110pt;0pt;100pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

If Me.lstMain.ListCount > 0 Then    '' ListBox에서 현재 선택된 항목의 인덱스 번호를 반환(아무 것도 선택되지 않으면 -1을 반환)
    Me.lstMain.ListIndex = 0
    End If
    
'Me.txt전표번호.Value = Me.lstMain.List(r, 2)
    
End Sub


Private Sub CommandButton1_Click()


    Dim Conn As Object
    Dim db As Variant
    Dim i As Long
    Dim arr() As Variant
    Dim voucherNo As String
    Dim inputPass As String
    Dim YN As VbMsgBoxResult
    Dim r As Long


    If Me.lstMain.ListIndex = -1 Then
            MsgBox "삭제할 전표를 선택해주세요", vbExclamation, "알림"
            Exit Sub
    End If
    
    ' 2. 바로 비밀번호 입력창 띄우기
    inputPass = InputBox("선택하신 전표를 삭제하시려면" & vbCrLf & "관리자 비밀번호를 입력하세요.", "전표 삭제 보안 확인")
    
    ' 취소를 눌렀거나 아무것도 입력 안 했을 때 처리
    If inputPass = "" Then
        Exit Sub
    End If
    
    ' 전역 변수 ADMIN_PW와 비교
    If inputPass <> ADMIN_PW Then
        MsgBox "비밀번호가 일치하지 않습니다.", vbCritical, "오류"
        Exit Sub
    End If
    
    Me.Label1.Caption = "   적요                                                                     거래처                                차        변                대        변                 계정명"


    r = Me.lstMain.ListIndex
    voucherNo = Me.lstMain.List(r, 2)

    ' 전표 전체 조회
    db = Get_DB_Access("분개장")
    db = Filtered_DB(db, voucherNo)

    Update_List Me.lstMain, db, "0pt;0pt;0pt;0pt;280pt;160pt;110pt;110pt;0pt;100pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

    YN = MsgBox("선택하신 전표를 삭제하시겠습니까? ", vbYesNo, "알림")
    If YN = vbNo Then Exit Sub

    ReDim arr(0 To Me.lstMain.ListCount - 1)

    For i = 0 To Me.lstMain.ListCount - 1
        arr(i) = Me.lstMain.List(i, 0) ' ID
    Next i

    ' DB 연결
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    '  분개장 삭제 + 백업
    For i = LBound(arr) To UBound(arr)

        Conn.Execute "INSERT INTO 삭제전표 SELECT * FROM 분개장 WHERE ID=" & arr(i)
        Conn.Execute "DELETE FROM 분개장 WHERE ID=" & arr(i)

    Next i

    '  채권채무명세 삭제
    db = Get_DB_Access("채권채무명세")
    db = Filtered_DB(db, voucherNo, "3", True)

    If Not IsEmpty(db) Then

        ReDim arr(1 To UBound(db, 1))

        For i = 1 To UBound(db, 1)
            arr(i) = db(i, 1)
        Next i

        For i = LBound(arr) To UBound(arr)
            Conn.Execute "DELETE FROM 채권채무명세 WHERE ID=" & arr(i)
        Next i

    End If

    Conn.Close
    Set Conn = Nothing

    MsgBox "전표 DB 삭제 완료"

    UserForm_Initialize


End Sub

Private Sub txt전표일자_Change()

Me.Label1.Caption = "  전표번호                    적요                                                                                      거래처"

Dim db As Variant

db = Get_DB_Access("분개장")
db = Filtered_DB(db, "1", "4", True) '---->4열에서 명세번호 1번만 필터링
db = Filtered_DB(db, Me.txt전표일자.Value, "2", True)  '-----> 2열에서 완전 일치하는 값 찾기

Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;160pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"


Me.txt전표번호.Value = ""

End Sub

Private Sub UserForm_Initialize()

Me.Label1.Caption = "  전표번호             　적요                                                             거래처                           　　  계정코드               계정명"

Dim db As Variant
' 제목행 제외 (기본값 False)
'db = Get_DB(sht분개장)

db = Get_DB_Access("분개장")
db = Filtered_DB(db, "1", "4", True) '명세번호 1번만 필터링 'true값을 쓰면 완전일치
db = Filtered_DB_Exclude(db, "전기이월", "5") '"전기이월" 이라는 텍스트를 5번열에서 제외시켜줌

'Update_list 보조함수
'지정한 리스트박의 값을 배열에서 받아와 갱신합니다.
'Update_list Listbox, DB, 열넓이
'열넓이 "0pt;50pt...."
Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;250pt;160pt;0pt;0pt;100pt;100pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"


Me.txt전표일자.SetFocus


End Sub

Sub Update_lstMain()

Dim db As Variant

db = Get_DB_Access("분개장")
db = Filtered_DB(db, Me.txt전표일자.Value)

Update_List Me.lstMain, db, "0pt;0pt;100pt;0pt;280pt;160pt;110pt;110pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;"

End Sub

'########################
' 시트에서 ID 를 갖는 레코드 복사붙여넣기 후 삭제   '--->오빠두 엑셀 Delet 함수 활용하여 만듬
' Delete_Record Sheet1, ID
'########################

Sub CopyBackup_Delete_Record(ws As Worksheet, ID)

Dim cRow As Long

If IsNumeric(ID) = True Then ID = CLng(ID)

With ws
    cRow = get_UpdateRow(ws, ID)
    ' 해당 행을 "삭제리스트" 시트의 마지막 행 아래에 복사
    .Rows(cRow).Copy Destination:=sht삭제백업.Rows(sht삭제백업.Cells(Rows.Count, 1).End(xlUp).Row + 1)
    .Rows(cRow).Delete

End With

End Sub

Private Sub 전표복사()

Dim r As Integer
Dim c As Integer

r = Me.lstMain.ListIndex
maxRow = Me.lstMain.ListCount - 1

If Me.lstMain.List(maxRow, 11) = "KRW" Then   '원화전표 복사일 땐, 금액까지 복사

   For c = 0 To maxRow
       frmAcctInput.lstMain.AddItem                ' 새 행 추가
       LastRow = frmAcctInput.lstMain.ListCount - 1 ' 새 행 번호 갱신
   
       frmAcctInput.lstMain.List(LastRow, 1) = Me.lstMain.List(c, 8)
       frmAcctInput.lstMain.List(LastRow, 2) = Me.lstMain.List(c, 9)
       frmAcctInput.lstMain.List(LastRow, 3) = Me.lstMain.List(c, 4)
       frmAcctInput.lstMain.List(LastRow, 4) = Me.lstMain.List(c, 5)
       frmAcctInput.lstMain.List(LastRow, 5) = Format(Me.lstMain.List(c, 6), "#,##0")
       frmAcctInput.lstMain.List(LastRow, 6) = Format(Me.lstMain.List(c, 7), "#,##0")
   Next c
   
Else

   MsgBox "외화 전표는 복사가 불가합니다." & vbCrLf & "전표 입력창에서 직접 입력해주세요.", vbCritical, "외화 전표"

End If

End Sub

'Private Sub lstMain_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub lstMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.lstMain
'End Sub