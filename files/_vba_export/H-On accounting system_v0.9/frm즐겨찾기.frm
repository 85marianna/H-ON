
Option Explicit

Private Sub cmd닫기_Click()

Unload Me

End Sub

Private Sub cmd불러오기_Click()

    ''Call 즐겨찾기불러오기(Me.lst즐겨찾기.Value)
    Call 즐겨찾기불러오기(Me.lst즐겨찾기.List(Me.lst즐겨찾기.ListIndex, 1))
    Call 차대합계_일치_Public(frmAcctInput)
    Call ToggleControl(frmAcctInput.cbo통화, False)
    
    'Unload Me

End Sub

Private Sub cmd삭제_Click()

    Dim Conn As Object
    Dim Sql As String
    Dim FavName As String

    'FavName = Me.lst즐겨찾기.Value  ''ID행번호가 나옴
    FavName = Me.lst즐겨찾기.List(Me.lst즐겨찾기.ListIndex, 1)   '여기서 1은 두 번째 컬럼(즐겨찾기명)

    If FavName = "" Then Exit Sub

    If MsgBox("[" & FavName & "] 항목을 삭제하시겠습니까?", _
              vbYesNo + vbQuestion, "즐겨찾기 삭제") = vbNo Then Exit Sub

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Sql = "DELETE FROM 즐겨찾기 WHERE 즐겨찾기명 = '" & FavName & "'"

    Conn.Execute Sql

    Conn.Close
    Set Conn = Nothing

    Me.lst즐겨찾기.Clear
    UserForm_Initialize  ' 리스트박스 새로고침

    MsgBox "삭제되었습니다."

End Sub

Private Sub cmd수정_Click()

    Dim Conn As Object
    Dim Sql As String
    Dim FavName As String
    Dim newName As String

    FavName = Me.lst즐겨찾기.List(Me.lst즐겨찾기.ListIndex, 1)

    If FavName = "" Then Exit Sub

    newName = InputBox("수정할 즐겨찾기명을 입력하세요.", "즐겨찾기명 수정")

    If Trim(newName) = "" Then Exit Sub

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()

    Sql = "UPDATE 즐겨찾기 " & _
          "SET 즐겨찾기명 = '" & newName & "' " & _
          "WHERE 즐겨찾기명 = '" & FavName & "'"

    Conn.Execute Sql

    Conn.Close
    Set Conn = Nothing

    Me.lst즐겨찾기.Clear
    UserForm_Initialize

    MsgBox "수정되었습니다."


End Sub



Private Sub UserForm_Initialize()

    Dim Conn As Object
    Dim db As Variant

    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open Get_ConnStr()
    
    
    db = Get_DB_Access("즐겨찾기")
    db = Filtered_DB(db, "1", "3")  '명세번호 1번만 필터링
    
    Update_List Me.lst즐겨찾기, db, "0pt;250pt;0pt;0pt;0pt;0pt;0pt;0pt;0pt;50pt;"

    Conn.Close
    Set Conn = Nothing
    

End Sub
