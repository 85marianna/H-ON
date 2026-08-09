
Private Sub btnDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Me.txtDate.Value = frmCalendar.GetDate

End Sub

Private Sub cboReport_Change()

    'cmd조회.Enabled = (cboReport.ListIndex <> -1)   '보고서를 선택하면 조회버튼 활성화
    CheckSearchEnable

End Sub

Private Sub cmd조회_Click()

    ' 1. 보고서 선택 여부 체크
    If cboReport.ListIndex = -1 Then
        MsgBox "보고서 종류를 선택해 주세요.", vbExclamation
        cboReport.SetFocus
        Exit Sub
    End If

    ' 2. 선택된 보고서에 따라 생성 + 시트 이동
    Select Case cboReport.Value

        Case "합계잔액시산표"
            
            Dim dtEnd As Date
            Dim dtStart As Date
        
            dtEnd = frm보고서조회.txtDate.Value              ' 선택 날짜 (예: 2025.06.30)
            dtStart = DateSerial(Year(dtEnd), 1, 1)          ' 그 해의 1월1일 (예: 2025.01.01)
        
            Call 합계잔액시산표생성_DB(dtStart, dtEnd)

        Case "재무상태표"
            Call 재무상태표_생성

        Case "손익계산서"
            Call 손익계산서_생성

        Case Else
            MsgBox "정의되지 않은 보고서입니다.", vbCritical
            Exit Sub

    End Select
    
   Me.Hide

End Sub


Private Sub txtDate_Change()

    CheckSearchEnable
    
End Sub

Private Sub UserForm_Initialize()

    With cboReport
        .Clear
        .List = Array("합계잔액시산표", "재무상태표", "손익계산서")
    End With
    
    txtDate.Value = ""
    cmd조회.Enabled = False

End Sub


Private Sub CheckSearchEnable()

    Dim isReady As Boolean

    ' 날짜, 보고서 둘 다 선택되었는지
    isReady = _
        (cboReport.ListIndex <> -1) _
        And (Trim(txtDate.Value) <> "")

    ' 조회 버튼 활성화
    cmd조회.Enabled = isReady

    ' 안내문구 표시/숨김
    lblGuide.Visible = Not isReady

End Sub

