
Function Get_DB_Path() As String

'1. 경로 전용 함수
    ''Get_DB_Path = Environ("USERPROFILE") & "\Box\AFK_14_Accounting\경리재무_14_HEG\★H-On 회계시스템★\H-On accounting_accDB.accdb"

    Get_DB_Path = "C:\Users\hyejung.jo\Downloads\한국전기초자회계시스템\H-On accounting_DB(로컬테스트용).accdb"
    
    ''Get_DB_Path = "C:\Users\85mar\Downloads\H-On accounting_DB(로컬테스트용).accdb"

End Function

Function Get_ConnStr() As String
'2. 연결문 함수

    Get_ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=" & Get_DB_Path()
                  
End Function

Public Function ConnectAccessDB() As Object

    Dim Conn As Object

    ' DB 존재 여부 확인
    If Dir(Get_DB_Path()) = "" Then
        MsgBox "DB 파일을 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               Get_DB_Path(), vbCritical, "H-On Accounting"
        Exit Function
    End If

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Set ConnectAccessDB = Conn

End Function


Function Get_NextVoucherNo(Conn As Object, inputDate As Date) As String

    Dim Sql As String
    Dim Rs As Object
    Dim thisYear As String
    Dim nextNo As Long
    
    thisYear = Format(inputDate, "yyyy")
    
    Sql = "SELECT MAX(CInt(Right(전표번호, 3))) FROM 분개장 " & _
          "WHERE Left(전표번호, 4) = '" & thisYear & "'"
    
    Set Rs = Conn.Execute(Sql)
    
    If IsNull(Rs.Fields(0).Value) Then
        nextNo = 1          ' 해당 연도 첫 전표
    Else
        nextNo = Rs.Fields(0).Value + 1
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Get_NextVoucherNo = Format(inputDate, "yyyy-mm-dd") & "-" & Format(nextNo, "000")

End Function


'=== 공용 세로 인쇄 설정 함수 ===
Public Sub SetPrintOptions(ws As Worksheet, Optional titleRow As String = "")

    With ws.PageSetup
        ' 여백 (1cm)
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)

        ' 페이지 방향: 세로
        .Orientation = xlPortrait

        ' 모든 열을 한 페이지에 맞춤 (세로는 여러 페이지 가능)
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False

        ' 페이지 가운데 맞춤 (가로)
        .CenterHorizontally = True

        ' 머리글/바닥글
        .CenterFooter = "Page &P of &N"

        ' 인쇄 제목 행 설정 (예: "$2:$2")
        If titleRow <> "" Then
            .PrintTitleRows = titleRow
        End If
    End With

End Sub
