Attribute VB_Name = "modWndProc"
Option Explicit
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_DESTROY As Long = &H2

Private Const WM_HSCROLL As Long = &H114
Private Const WM_VSCROLL As Long = &H115
Private Const SB_LINEDOWN As Long = 1
Private Const SB_LINEUP As Long = 0
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
   ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function GetParent Lib "user32.dll" ( _
   ByVal hWnd As Long) As Long

Private Declare Function GetClientRect Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   lpRect As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function ScreenToClient Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   lpPoint As POINTAPI) As Long

Private Declare Function PtInRect Lib "user32.dll" ( _
   lpRect As RECT, _
   ByVal x As Long, _
   ByVal y As Long) As Long

Private Declare Function GetWindowPlacement Lib "user32.dll" ( _
   ByVal hWnd As Long, _
   lpwndpl As WINDOWPLACEMENT) As Long
Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
     ByVal lpPrevWndFunc As Long, _
     ByVal hWnd As Long, _
     ByVal msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
     ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = -4

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const OLDWNDPROC = "OldWndProc"

Private lOldWndProc As Long
Private lHwnd As Long

Public Function MouseWheelProc(ByVal hWnd As Long, _
   ByVal iMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long
   
   Select Case iMsg
     Case WM_MOUSEWHEEL
        Dim lMove As Integer
        Dim lNewlParam As Long
        Dim lNewwParam As Long
        Dim rRect As RECT
        Dim pt As POINTAPI
        
        Dim wp As WINDOWPLACEMENT
        
        Call GetWindowPlacement(hWnd, wp)
        
        Debug.Print "Window Placement"
        Debug.Print "Max:", wp.ptMaxPosition.x, wp.ptMaxPosition.y
        Debug.Print "Min:", wp.ptMinPosition.x, wp.ptMinPosition.y
        Debug.Print "rcNormal:", wp.rcNormalPosition.Left, wp.rcNormalPosition.Top, wp.rcNormalPosition.Right, wp.rcNormalPosition.Bottom
        
        
        
        Call GetClientRect(hWnd, rRect)
        Debug.Print rRect.Left, rRect.Top, rRect.Right, rRect.Bottom
        Debug.Print "Wheel Motion", HIWORD(wParam)
        Debug.Print "Mouse Pos:", LOWORD(lParam), HIWORD(lParam)
        
        pt.x = LOWORD(lParam)
        pt.y = HIWORD(lParam)
        
        Call ScreenToClient(GetParent(hWnd), pt)
        
        If PtInRect(wp.rcNormalPosition, pt.x, pt.y) Then
           Debug.Print "in Rect"
          
          
          
          lMove = HIWORD(wParam)
          If lMove > 0 Then
            'up
            'multiples of 120
            If lMove = 120 Then
              lNewlParam = MAKELPARAM(SB_LINEUP, 0) '0, SB_LINEUP)
            Else
              'full page
              lNewlParam = MAKELPARAM(SB_PAGEUP, 0)
            End If
            
            Call SendMessage(hWnd, WM_VSCROLL, lNewlParam, 0)
          Else
            'down
            If lMove = -120 Then
              lNewlParam = MAKELPARAM(SB_LINEDOWN, 0) ', SB_LINEDOWN)
            Else
              'full page
              lNewlParam = MAKELPARAM(SB_PAGEDOWN, 0) ', SB_LINEDOWN)
            End If
            
            Call SendMessage(hWnd, WM_VSCROLL, lNewlParam, 0)
          End If
        
        End If
     Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, iMsg, wParam, lParam)
      Call UnTrackMouseWheel(hWnd)
      Exit Function
   
   End Select
   
   MouseWheelProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, iMsg, wParam, lParam)
   
End Function
   

Public Function TrackMouseWheel(ByVal hWnd As Long) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out

  If GetProp(hWnd, OLDWNDPROC) Then
    TrackMouseWheel = True
    Exit Function
  End If
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MouseWheelProc)

#Else
    Dim objWPHook As WindowProcHook
    
    Set objWPHook = CreateWindowProcHook
    m_colWPHooks.Add objWPHook, CStr(hWnd)
    
    With objWPHook
      Call .SetMainProc(AddressOf MouseWheelProc)
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
      Call .SetDebugProc(lpfnOld)
    End With

#End If
  
  If lpfnOld Then
    fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
  End If
  
Out:
  If fSuccess Then
    TrackMouseWheel = True
  
  Else
    If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    'MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
    '              "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
End Function

Public Sub UnTrackMouseWheel(ByVal hWnd As Long)
 Call SetWindowLong(hWnd, GWL_WNDPROC, GetProp(hWnd, OLDWNDPROC))
 Call RemoveProp(hWnd, OLDWNDPROC)
End Sub
