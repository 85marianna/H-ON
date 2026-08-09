
Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt조회일자.Value = frmCalendar.GetDate

End Sub

Private Sub cmd닫기_Click()

Unload Me

End Sub


Private Sub cmd인쇄_Click()

    Dim pdfFile As String
    Dim r As Long
    Dim c As Long

    ' 날짜 입력 체크
    If Trim(Me.txt조회일자.Value) = "" Then
        MsgBox "인쇄할 날짜를 선택하세요.", vbExclamation, "날짜 선택"
        Me.txt조회일자.SetFocus
        Exit Sub
    End If

    Call cmd조회_Click

    With sht템플릿
    
        ' 기존 데이터 삭제
        .Range("A1:K25").Clear
        
        '텍스트
        .Columns("A").NumberFormat = "@"
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "@"
        .Columns("D").NumberFormat = "@"
        .Columns("E").NumberFormat = "@"
        .Columns("G").NumberFormat = "@"
    
        '날짜
        .Columns("F").NumberFormat = "yyyy-mm-dd"
    
        '금액
        .Columns("H").NumberFormat = "#,##0"    '외화금액
        .Columns("I").NumberFormat = "#,##0"    '원화금액
    

        ' 제목
        With .Range("A1")
            .Font.Size = 17
            .Font.Bold = True
        End With
    
        ' 컬럼행(4행)
        .Rows(3).RowHeight = 20
        .Rows(4).RowHeight = 35
        .Rows(4).Font.Bold = True
        .Rows(4).HorizontalAlignment = -4108 'xlCenter
    
        ' 컬럼행 회색 음영
        .Range("A4:I4").Interior.Color = RGB(217, 217, 217)
    
        ' 전체 테두리
        With .Range("A4:I6").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        

        ' 리스트박스 → 시트
        For r = 0 To Me.lstMain.ListCount - 1

            ' 0번 컬럼(숨김 계정코드)은 제외
            For c = 1 To Me.lstMain.ColumnCount - 1

                .Cells(r + 5, c).Value = Me.lstMain.List(r, c)

            Next c

        Next r
        
        With .Range("A1:I1")
            .HorizontalAlignment = xlCenterAcrossSelection   ''텍스트맞춤- 선택영역의 가운데로 (병합x)
            .VerticalAlignment = -4108    'xlCenter
        End With

        .Range("A1").Value = "【 제예금 명세서 】"
        .Range("A3").Value = Me.txt조회일자.Value
   
        .Range("A4:I4").Value = Array("세목코드", "계정명", "은행명", "은행코드", "계좌번호", "개설일", "환종", "외화잔액", "원화잔액")
        
        .Columns("B").ColumnWidth = 10
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 23
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 7
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 17

        .PageSetup.PrintArea = "A1:I7"

    End With

    pdfFile = Environ("TEMP") & "\H-ON.pdf"

    ' PDF 생성 후 자동 열기
    ExportSheetToPDF sht템플릿, pdfFile, True
    

    Unload Me

End Sub


Private Sub cmd조회_Click()

    Dim ws As Worksheet
    
    Dim ledger As Variant
    Dim LastRow As Long
    
    Dim wonBal1110105 As Double
    Dim fxBal1110105 As Double
    Dim wonBal1110107 As Double
    
    Dim i As Long

    ' 날짜 입력 체크
    If Trim(Me.txt조회일자.Value) = "" Then
        MsgBox "조회할 날짜를 선택하세요.", vbExclamation, " 날짜 선택"
        Me.txt조회일자.SetFocus
        Exit Sub
    End If
   
   
    '==========================
    ' 외화예금
    '==========================
    ledger = 원장데이터가져오기_DB("1110105", DateSerial(Year(Me.txt조회일자.Value), 1, 1), Me.txt조회일자.Value)
    
    If Not IsEmpty(ledger) Then
        LastRow = UBound(ledger, 1)
        wonBal1110105 = ledger(LastRow, 10)
        fxBal1110105 = ledger(LastRow, 12)
    End If
    
    '==========================
    ' 기업자유예금
    '==========================
    ledger = 원장데이터가져오기_DB("1110107", DateSerial(Year(Me.txt조회일자.Value), 1, 1), Me.txt조회일자.Value)
    
    If Not IsEmpty(ledger) Then
        LastRow = UBound(ledger, 1)
        wonBal1110107 = ledger(LastRow, 10)
    End If
    
    
    For i = 0 To Me.lstMain.ListCount - 1
    
        Select Case Me.lstMain.List(i, 1)
    
            Case "1110105"    ' 외화예금
                Me.lstMain.List(i, 8) = Format(fxBal1110105, "#,##0")
                Me.lstMain.List(i, 9) = Format(wonBal1110105, "#,##0")
    
            Case "1110107"    ' 기업자유예금
                Me.lstMain.List(i, 9) = Format(wonBal1110107, "#,##0")
    
        End Select
    
    Next i
    
End Sub


Private Sub UserForm_Initialize()

    Dim db As Variant

    db = Get_DB_Access("제예금마스터")

    With Me.lstMain
        .ColumnCount = 10
        .ColumnWidths = "0pt;70pt;100pt;70pt;50pt;140pt;90pt;60pt;100pt;100pt"
        If IsArray(db) Then
            .List = db
        End If
    End With

End Sub
