Function ShapeInRange(rng As Range, _
                Optional iRed As Long = 255, _
                Optional iGreen As Long = 0, _
                Optional iBlue As Long = 0, _
                Optional FillVisible As MsoTriState = msoTrue, _
                Optional LineVisible As MsoTriState = msoTrue, _
                Optional Transparent As Double = 0.95, _
                Optional LineWeight As Double = 0.5, _
                Optional DashType As MsoLineDashStyle = msoLineDash, _
                Optional ShapeType As MsoAutoShapeType = msoShapeRectangle, _
                Optional ActivateSheet As Boolean = True) As Shape

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'수정 및 배포 시 출처를 반드시 명시해야 합니다.
'
'■ ShapeInRange 명령문
'■ 선택된 범위 안에 도형을 삽입합니다.
'■ 사용방법
''아래 명령문을 유저폼 모듈안에 붙여넣기 한 뒤, '리스트박스' 를 실제 적용할 리스트박스 이름으로 변경합니다.
''----------------------------------------------------------------------------------------------------
'Dim Shp As Shape
'Set Shp = ShapeInRange(Range("A1"))
''-----------------------------------------------------------------------------------------------------
'▶ 인수 설명
'_____________Rng               : 도형을 삽입할 범위입니다.
'_____________iRed               : [선택인수] 삽입할 도형의 RGB, R 값입니다. 기본값은 255 입니다.
'_____________iGreen           : [선택인수] 삽입할 도형의 RGB, G 값입니다. 기본값은 0 입니다.
'_____________iBlue              : [선택인수] 삽입할 도형의 RGB, B 값입니다. 기본값은 0 입니다.
'_____________FillVisible       : [선택인수] 채우기 여부입니다. 기본값은 TRUE 입니다.
'_____________LineVisible     : [선택인수] 윤곽선 여부입니다. 기본값은 TRUE 입니다.
'_____________Transparent  : [선택인수] 채우기 투명도입니다. 기본값은 0.25 입니다.
'_____________LineWeight   : [선택인수] 윤곽선 두께입니다. 기본값은 0.5 입니다.
'_____________DashType      : [선택인수] 윤곽선 스타일입니다. 기본값은 점선입니다.
'_____________ShapeType    : [선택인수] 도형 모양입니다. 기본값은 직사각형입니다.
'_____________AvtiveSheet   : [선택인수] 도형삽입 후 삽입된 시트 활성화여부입니다. 기본값은 True 입니다.
'###############################################################

Dim Shp As Shape
Dim ws As Worksheet

Set ws = rng.Parent

With rng
    Set Shp = ws.Shapes.AddShape(ShapeType, .left, .top, .Width, .Height)
End With

With Shp
    With .Fill
        .Visible = FillVisible
        .ForeColor.RGB = RGB(iRed, iGreen, iBlue)
        .Transparency = Transparent
    End With
    With .Line
        .Visible = LineVisible
        .ForeColor.RGB = RGB(iRed, iGreen, iBlue)
        .Weight = LineWeight
        .DashStyle = DashType
    End With
End With

Set ShapeInRange = Shp
If ActivateSheet = True Then Shp.Parent.Activate

End Function
