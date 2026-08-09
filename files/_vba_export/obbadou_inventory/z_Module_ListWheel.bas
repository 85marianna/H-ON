Option Explicit

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'수정 및 배포 시 출처를 반드시 명시해야 합니다.
'
'■ HookListBoxScroll & UnHookListBoxScroll 명령문
'■ 유저폼에 사용된 리스트박스가 마우스 휠로 스크롤 되도록 후킹 / 후킹해제하는 명령문입니다.
'■ 사용방법
''아래 명령문을 유저폼 모듈안에 붙여넣기 한 뒤, '리스트박스' 를 실제 적용할 리스트박스 이름으로 변경합니다.
''----------------------------------------------------------------------------------------------------
'Private Sub 리스트박스_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'UnhookListBoxScroll
'End Sub
'Private Sub 리스트박스_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'HookListBoxScroll Me, Me.리스트박스
'End Sub
''-----------------------------------------------------------------------------------------------------
'
'본 명령문은 아래 링크를 참조하여 작성된 명령문입니다.
'https://social.msdn.microsoft.com/Forums/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba
'###############################################################

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As LongPtr
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As LongPtr = &H20A
Private Const HC_ACTION As LongPtr = 0
Private Const GWL_HINSTANCE As LongPtr = (-6)

Private mLngMouseHook As LongPtr
Private mListBoxHwnd As LongPtr
Private mbHook As Boolean
Private mCtl As MSForms.Control
Dim n As LongPtr

Sub HookListBoxScroll(frm As Object, ctl As MSForms.Control)
Dim lngAppInst As LongPtr
Dim hwndUnderCursor As LongPtr
Dim tPT As POINTAPI
     GetCursorPos tPT
     hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
     If Not frm.ActiveControl Is ctl Then
             ctl.SetFocus
     End If
     If mListBoxHwnd <> hwndUnderCursor Then
             UnhookListBoxScroll
             Set mCtl = ctl
             mListBoxHwnd = hwndUnderCursor
             lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
             If Not mbHook Then
                     mLngMouseHook = SetWindowsHookEx( _
                                                     WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
                     mbHook = mLngMouseHook <> 0
             End If
     End If
End Sub

Sub UnhookListBoxScroll()
     If mbHook Then
                Set mCtl = Nothing
             UnhookWindowsHookEx mLngMouseHook
             mLngMouseHook = 0
             mListBoxHwnd = 0
             mbHook = False
        End If
End Sub

Private Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
Dim idx As LongPtr
        On Error GoTo errH
     If (nCode = HC_ACTION) Then
             If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = mListBoxHwnd Then
                     If wParam = WM_MOUSEWHEEL Then
                                MouseProc = True
                                If lParam.hwnd > 0 Then idx = -1 Else idx = 1
                             idx = idx + mCtl.TopIndex
                             If idx >= 0 Then mCtl.TopIndex = idx
                                Exit Function
                     End If
             Else
                     UnhookListBoxScroll
             End If
     End If
     MouseProc = CallNextHookEx( _
                             mLngMouseHook, nCode, wParam, ByVal lParam)
     Exit Function
errH:
     UnhookListBoxScroll
End Function