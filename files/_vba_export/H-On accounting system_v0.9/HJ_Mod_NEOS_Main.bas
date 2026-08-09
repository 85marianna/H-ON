Option Explicit

' 이 시스템의 마스터 비밀번호 통합 관리
Public Const ADMIN_PW As String = "1234"

'로그인한 사람 정보를 전역변수 저장
Public gUserID As Long
Public gEmpNo As String
Public gUserName As String
Public gAuthority As String
Public gLoginTime As Date

Public IsLogin As Boolean


Sub ToggleSheetTabs()

    Dim ws As Worksheet
    Dim pwd As String
    Dim isDevMode As Boolean

    ' 개발자 모드 여부 판단 (MAIN 외 시트가 하나라도 보이면 개발자 모드로 간주)
    isDevMode = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MAIN" And ws.Visible = xlSheetVisible Then
            isDevMode = True
            Exit For
        End If
    Next ws

    ' 일반 모드 → 개발자 모드 진입 시 비밀번호 요구
    If isDevMode = False Then
        pwd = InputBox("개발자 모드에 진입하려면 비밀번호를 입력해주세요", "보안 확인")
        If pwd <> ADMIN_PW Then
            MsgBox "비밀번호 오류!", vbCritical, "경고"
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False

    If isDevMode = False Then
        ' =========================
        ' 개발자 모드 ON
        ' =========================
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws

        With ThisWorkbook.Sheets("MAIN")
            .Unprotect Password:=ADMIN_PW
            .EnableSelection = xlNoRestrictions
            .ScrollArea = ""
        End With

        MsgBox "개발자 모드 ON: 모든 시트가 표시됩니다.", vbInformation

    Else
        ' =========================
        ' 개발자 모드 OFF (시스템 모드)
        ' =========================
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = "MAIN" Then
                ws.Visible = xlSheetVisible
            Else
                ws.Visible = xlSheetVeryHidden
            End If
        Next ws

        With ThisWorkbook.Sheets("MAIN")
            .Protect Password:=ADMIN_PW, UserInterfaceOnly:=True
            .EnableSelection = xlNoSelection
            .ScrollArea = "A1:K25"
        End With

        ThisWorkbook.Sheets("MAIN").Activate
        ThisWorkbook.Sheets("MAIN").Range("A1").Select

        MsgBox "시스템 모드로 전환되었습니다.", vbInformation
    End If

    Application.ScreenUpdating = True

End Sub


Sub ActivateSheetSafely(ws As Worksheet)
''이거 한 번 만들어두면 모든 시트 이동 문제 끝남

    Application.ScreenUpdating = False

    ' 시트가 숨겨져 있으면 임시로 보이게
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.Activate

    Application.ScreenUpdating = True

End Sub


'공용 PDF 생성 함수 (표준모듈)
Public Sub ExportSheetToPDF( _
            ws As Worksheet, _
            pdfFile As String, _
            Optional OpenPDF As Boolean = True)

    Dim oldVisible As XlSheetVisibility

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    oldVisible = ws.Visible

    ' 숨김 시트면 임시로 표시
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfFile, _
        Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=OpenPDF

ExitHandler:

    ws.Visible = oldVisible
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:

    MsgBox "PDF 생성 중 오류가 발생했습니다." & vbCrLf & _
           Err.Description, vbExclamation

    Resume ExitHandler

End Sub


Sub 메인화면열기()

    ' 메인 유저폼 열기 '' 뒤에 0을 붙이면 '비모달(Modeless)'로 실행
    frmMain.Show 0

End Sub


Sub 로그인창열기()

        If IsLogin Then
        frmMain.Show 0
    Else
        frm로그인.Show 0
    End If

End Sub
