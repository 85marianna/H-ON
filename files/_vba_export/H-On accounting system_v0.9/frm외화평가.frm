Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txt조회일자.Value = frmCalendar.GetDate

End Sub

Private Sub cmd닫기_Click()

Unload Me

End Sub


Private Sub cmd외화평가_Click()

Dim voucherNo As String
Dim remark As String

If Me.lstMain.List(0, 12) = "" Then    ''평가차액이 공란일 경우
    MsgBox "외화 평가할 내역 조회를 실행해주세요.", , "외화평가"
    Exit Sub
End If


If 외화평가전표존재(Me.txt조회일자.Value, voucherNo) Then

    MsgBox "외화평가 전표가 존재합니다." & vbCrLf & vbCrLf & _
           "전표번호 : " & voucherNo, _
           vbExclamation, "외화평가"
    Exit Sub
End If


If Me.lstMain.List(0, 12) = "0" Or 0 Then
    MsgBox "평가차액이 0원이므로 외화평가가 불가합니다.", , "외화평가"
    Exit Sub
End If


If MsgBox("외화평가 전표를 생성하시겠습니까?", _
          vbQuestion + vbYesNo, "외화평가") = vbNo Then Exit Sub


remark = Format(Me.txt조회일자.Value, "yyyy년 m월") & " 외화예금평가"

'외화평가전표 생성

If Me.lstMain.List(0, 12) > 0 Then  '평가차익이면

       frmAcctInput.lstMain.AddItem                ' 새 행 추가
       LastRow = frmAcctInput.lstMain.ListCount - 1 ' 새 행 번호 갱신
   
       frmAcctInput.lstMain.List(LastRow, 1) = "4210900"
       frmAcctInput.lstMain.List(LastRow, 2) = "외화환산이익"
       frmAcctInput.lstMain.List(LastRow, 3) = remark
       frmAcctInput.lstMain.List(LastRow, 4) = ""
       frmAcctInput.lstMain.List(LastRow, 5) = "0"
       frmAcctInput.lstMain.List(LastRow, 6) = Me.lstMain.List(0, 12)
       
       frmAcctInput.lstMain.AddItem                ' 새 행 추가
       LastRow = frmAcctInput.lstMain.ListCount - 1 ' 새 행 번호 갱신
   
       frmAcctInput.lstMain.List(LastRow, 1) = "1110105"
       frmAcctInput.lstMain.List(LastRow, 2) = "외화예금"
       frmAcctInput.lstMain.List(LastRow, 3) = remark
       frmAcctInput.lstMain.List(LastRow, 4) = ""
       frmAcctInput.lstMain.List(LastRow, 5) = Me.lstMain.List(0, 12)   '차변
       frmAcctInput.lstMain.List(LastRow, 6) = "0"                      '대변

   
Else     ' 평가차손이면

       frmAcctInput.lstMain.AddItem                ' 새 행 추가
       LastRow = frmAcctInput.lstMain.ListCount - 1 ' 새 행 번호 갱신
   
       frmAcctInput.lstMain.List(LastRow, 1) = "5311100"
       frmAcctInput.lstMain.List(LastRow, 2) = "외화환산손실"
       frmAcctInput.lstMain.List(LastRow, 3) = remark
       frmAcctInput.lstMain.List(LastRow, 4) = ""
       frmAcctInput.lstMain.List(LastRow, 5) = Me.lstMain.List(0, 12)
       frmAcctInput.lstMain.List(LastRow, 6) = "0"
       
       frmAcctInput.lstMain.AddItem                ' 새 행 추가
       LastRow = frmAcctInput.lstMain.ListCount - 1 ' 새 행 번호 갱신
   
       frmAcctInput.lstMain.List(LastRow, 1) = "1110105"
       frmAcctInput.lstMain.List(LastRow, 2) = "외화예금"
       frmAcctInput.lstMain.List(LastRow, 3) = remark
       frmAcctInput.lstMain.List(LastRow, 4) = ""
       frmAcctInput.lstMain.List(LastRow, 5) = ""
       frmAcctInput.lstMain.List(LastRow, 6) = Me.lstMain.List(0, 12)

   
End If


' ★ 핵심: 입력폼의 날짜 칸에 현재 조회일자를 미리 넣어주기 ★
' 폼을 열기 전에 값을 먼저 전달해!
frmAcctInput.txt전표일자.Value = Me.txt조회일자.Value
' 날짜를 마음대로 못 바꾸게 잠그기 (선택사항이지만 혜정이 의도라면 추천!)
frmAcctInput.txt전표일자.Locked = True
frmAcctInput.btnDate.Enabled = False
'frmAcctInput.cbo부서명.Value = "회계팀"

frmAcctInput.CreateReverse = Me.chk반대분개.Value

  
Call 차대합계_일치_Public(frmAcctInput)

Unload Me

전표입력폼


End Sub

Private Sub cmd인쇄_Click()

    Dim pdfFile As String

    ' 날짜 입력 체크
    If Trim(Me.txt조회일자.Value) = "" Then
        MsgBox "인쇄할 날짜를 선택하세요.", vbExclamation, " 날짜 선택"
        Me.txt조회일자.SetFocus
        Exit Sub
    End If
  
    With sht템플릿
    
        ' 기존 데이터 삭제
        .Range("A1:K25").Clear

        With .Range("A1:K1")
            .HorizontalAlignment = xlCenterAcrossSelection   ''텍스트맞춤- 선택영역의 가운데로 (병합x)
            .VerticalAlignment = -4108    'xlCenter
        End With
        
        '제목 서식
        With .Range("A1")
            .Font.Size = 17
            .Font.Bold = True
            .Value = "【 외화 평가 내역 】"
        End With
        
        
        ' 컬럼행(4행)
        .Rows(3).RowHeight = 20
        .Rows(4).RowHeight = 35
        .Rows(4).Font.Bold = True
        .Rows(4).HorizontalAlignment = -4108 'xlCenter
    
        ' 컬럼행 회색 음영
        .Range("A4:K4").Interior.Color = RGB(217, 217, 217)
    
        ' 전체 테두리
        With .Range("A4:K5").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With

                    
        '텍스트
        .Columns("A").NumberFormat = "@"
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "@"
        .Columns("D").NumberFormat = "@"
        .Columns("E").NumberFormat = "@"
        
        '금액
        .Columns("F").NumberFormat = "#,##0"
        .Columns("G").NumberFormat = "#,##0"
        
        '환율 (소수점 4자리)
        .Columns("H").NumberFormat = "#,##0.0000"
        .Columns("I").NumberFormat = "#,##0.0000"
        
        '금액
        .Columns("J").NumberFormat = "#,##0"
        .Columns("K").NumberFormat = "#,##0"
           
    
        '리스트박스 → 시트 출력
        .Range("A5").Value = Me.lstMain.List(0, 1)
        .Range("B5").Value = Me.lstMain.List(0, 2)
        .Range("C5").Value = Me.lstMain.List(0, 3)
        .Range("D5").Value = Me.lstMain.List(0, 5)
        .Range("E5").Value = Me.lstMain.List(0, 7)
        .Range("F5").Value = Me.lstMain.List(0, 8)
        .Range("G5").Value = Me.lstMain.List(0, 9)
        .Range("H5").Value = Me.lstMain.List(0, 10)
        .Range("I5").Value = Me.lstMain.List(0, 11)
        .Range("J5").Value = Me.lstMain.List(0, 12)
        .Range("K5").Value = Me.lstMain.List(0, 13)
        
        
        .Range("A3").Value = Me.txt조회일자.Value
        .Range("A4:K4").Value = Array("세목코드", "계정명", "은행명", "계좌번호", "환종", "외화잔액", "원화잔액", "기존환율", "평가환율", "평가후 원화잔액", "평가차액")
        
        
        .Columns("A:E").AutoFit
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 10
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 15

        .PageSetup.PrintArea = "A1:K5"

    End With

''    ' 다운로드 폴더 저장
''    pdfFile = Environ("USERPROFILE") & _
''              "\Downloads\외화평가내역_" & _
''              Format(Me.txt조회일자.Value, "yyyymmdd") & ".pdf"
    
    pdfFile = Environ("TEMP") & "\H-ON.pdf"

    ' PDF 생성 후 자동 열기
    ExportSheetToPDF sht템플릿, pdfFile, True

    Unload Me
    
End Sub


Private Sub txt조회일자_Change()

If Month(CDate(Me.txt조회일자.Value)) = 12 Then
    Me.chk반대분개.Value = False
Else
    Me.chk반대분개.Value = True
End If

End Sub

Private Sub txt환율_AfterUpdate()

   Me.lbl메세지.Caption = ""
   Me.txt환율.BackColor = vbWhite


End Sub

Private Sub txt환율_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    ' 숫자가 아니면 막기
    If Not IsNumeric(txt환율.Value) Then
        Me.lbl메세지.Caption = "환율을 정확하게 입력하세요."
        Me.txt환율.BackColor = RGB(255, 224, 224)
        Me.txt환율.SelStart = 0
        Me.txt환율.SelLength = Len(Me.txt환율.Value)
        Cancel = True
        
    End If

End Sub



Private Sub UserForm_Initialize()

    Dim db As Variant

    db = Get_DB_Access("제예금마스터")
    db = Filtered_DB(db, "1110105", "2")  '외화예금만 필터링


    With Me.lstMain
        .ColumnCount = 14
        .ColumnWidths = "0pt;70pt;70pt;0pt;0pt;140pt;0pt;50pt;100pt;110pt;70pt;70pt;110pt;80pt"
        If IsArray(db) Then
            .List = db
        End If
    End With

End Sub


Private Sub cmd조회_Click()

    Dim ws As Worksheet
    Dim ledger As Variant

    Dim LastRow As Long

    Dim wonBal As Double
    Dim fxBal As Double

    Dim avgRate As Double
    Dim rate As Double
    Dim evalAmt As Double
    Dim gainLoss As Double

    '============================
    ' 조회일자 입력 체크
    '============================
    If Trim(Me.txt조회일자.Value) = "" Then
        MsgBox "조회할 날짜를 선택하세요.", vbExclamation, "날짜 선택"
        Me.txt조회일자.SetFocus
        Exit Sub
    End If

    '============================
    ' 평가환율 입력 체크
    '============================
    If Trim(Me.txt환율.Value) = "" Then
        MsgBox "평가환율을 입력해주세요.", vbExclamation, "평가환율 입력"
        Me.txt환율.SetFocus
        Exit Sub
    End If

    '============================
    ' 외화예금 원장 조회
    '============================
    ledger = 원장데이터가져오기_DB("1110105", "", Me.txt조회일자.Value)

    If IsEmpty(ledger) Then
        MsgBox "조회된 외화예금 내역이 없습니다.", vbInformation
        Exit Sub
    End If

    LastRow = UBound(ledger, 1)

    wonBal = ledger(LastRow, 10)
    fxBal = ledger(LastRow, 12)

    '============================
    ' 외화평가 계산
    '============================
    rate = CDbl(Me.txt환율.Value)

    If fxBal <> 0 Then
        avgRate = wonBal / fxBal
    Else
        avgRate = 0
    End If

    evalAmt = Int(rate * fxBal)
    gainLoss = evalAmt - wonBal

    
    '============================
    ' 리스트박스 출력
    '============================
    With Me.lstMain
    
        .List(0, 8) = Format(fxBal, "#,##0")           '외화잔액
        .List(0, 9) = Format(wonBal, "#,##0")          '원화잔액
        .List(0, 10) = Format(avgRate, "#,##0.0000")   '평균환율
        .List(0, 11) = Format(rate, "#,##0.0000")      '평가환율
        .List(0, 12) = Format(evalAmt, "#,##0")        '평가후 원화잔액
        .List(0, 13) = Format(gainLoss, "#,##0")       '평가차액
    
    End With
    

End Sub

