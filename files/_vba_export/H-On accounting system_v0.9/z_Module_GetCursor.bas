Option Explicit

'####################################################
'■ 현재 마우스 커서의 위치를 픽셀로 반환하는 반환하는User32 윈도우 API 모듈입니다.
'    수정 및 배포는 자유로우나, 배포 시 출처를 반드시 명시해야합니다.
'    https://www.oppadu.com (오빠두엑셀)
'-------------------------------------------------------
'■ 사용방법
'Dim MousePOS As POINTAPI
'POINTAPI = convertMouseToForm
'유저폼.Top = POINTAPI.Y
'유저폼.Left = POINTAPI.X
'--------------------------------------------------------
'■ 본 모듈은 아래 링크를 일부 참조하여 작성되었습니다.
'https://sites.google.com/a/mcpher.com/share/Home/excelquirks/snippets/mouseposition
'####################################################

#If VBA7 Then
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#Else
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Type POINTAPI
    X As Long
    Y As Long
End Type

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Public Function pointsPerPixelX() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function pointsPerPixelY() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
End Function

Public Function WhereIsTheMouseAt() As POINTAPI
    Dim mPos As POINTAPI
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function

Public Function convertMouseToForm() As POINTAPI
    Dim mPos As POINTAPI
    mPos = WhereIsTheMouseAt
    mPos.X = pointsPerPixelY * mPos.X
    mPos.Y = pointsPerPixelX * mPos.Y
    convertMouseToForm = mPos
End Function