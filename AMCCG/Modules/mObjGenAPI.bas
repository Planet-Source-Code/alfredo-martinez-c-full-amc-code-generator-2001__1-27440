Attribute VB_Name = "mObjGenAPI"
Option Explicit

#If UNICODE Then
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2


'API Const for Editor

Public Const WM_USER = &H400
Public Const WM_PASTE = &H302
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_CLEAR = &H303

Public Const EM_LINEINDEX = &HBB&
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_LINEFROMCHAR = &HC9&
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_HIDESELECTION = (WM_USER + 63)
Public Const EM_REQUESTRESIZE = (WM_USER + 65)
Public Const EM_SETUNDOLIMIT = (WM_USER + 82)
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const EM_STOPGROUPTYPING = (WM_USER + 88)
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)
Public Const EM_AUTOURLDETECT = (WM_USER + 91)

Public Const EM_GETFIRSTVISIBLELINE = &HCE

'Declares with Editor CodeGen
Public Const LB_FINDSTRING = &H18F
Public Const LB_ERR = (-1)
Public Const EM_POSFROMCHAR = &HD6&
'Private Const EM_LINEFROMCHAR = &HC9
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Functions with Find Text
Public Declare Function GetFocus Lib "user32" () As Long
'Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Const CB_SHOWDROPDOWN = WM_USER + 15



Public Sub FlatBorder(ByVal hwnd As Long)
Dim TFlat As Long
    TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong hwnd, GWL_EXSTYLE, TFlat
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub


