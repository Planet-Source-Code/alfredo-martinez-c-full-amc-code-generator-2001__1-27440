VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variables y constantes para el progresbar
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


Private hProgBar As Long

Const WM_USER = &H400
Const PBM_SETRANGE = (WM_USER + 1)
Const PBM_SETPOS = (WM_USER + 2)
Const PBM_DELTAPOS = (WM_USER + 3)
Const PBM_SETSTEP = (WM_USER + 4)
Const PBM_STEPIT = (WM_USER + 5)
Const PBM_SETRANGE32 = (WM_USER + 6)
Const PBM_GETRANGE = (WM_USER + 7)
Private Type PPBRANGE
   iLow As Integer
   iHigh As Integer
End Type
Const PBM_GETPOS = (WM_USER + 8)
Const PROGRESS_CLASS = "msctls_progress32"
Private Const SB_GETRECT = (WM_USER + 10)
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000

'local variable(s) to hold property value(s)
Private mvarPanel As Long 'local copy
Private mvarMin As Long 'local copy
Private mvarMax As Long 'local copy
Private mvarValue As Long 'local copy



Public Function CreateProgress(StatusBar As StatusBar, Panel As Long) As Boolean
    
    On Error GoTo CreateProgressError
    
    Dim hStatBar As Long
    Dim rctPart As RECT
    
    mvarPanel = Panel
    CreateProgress = False
    
    ' Get the bounding rectangle dimesions of the status
    ' bar's 1st part, in status bar client coords.
    ' (replace the 0 part index for the index you want to use)
    hStatBar = StatusBar.hwnd
    
    SendMessage hStatBar, SB_GETRECT, ByVal Panel, rctPart
    hProgBar = CreateWindowEx(0, PROGRESS_CLASS, _
                                              vbNullString, WS_CHILD Or WS_VISIBLE, _
                                              rctPart.Left, rctPart.Top, _
                                              rctPart.Right, rctPart.Bottom - 2, _
                                              hStatBar, 0, _
                                              App.hInstance, ByVal 0)
    
    CreateProgress = True
    
    Exit Function

CreateProgressError:
    If IsWindow(hProgBar) Then DestroyWindow hProgBar
    Exit Function
End Function

Public Property Let Value(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
'
'
'  ' Let's do some stuff...
'  For dwIdx = 1 To dwItins
'    DoEvents
'
'                                                                                             '\___/'
'    If dwIdx Mod dwIncrement = 0 Then
'      ' Advance the current position of the progress bar by the step increment.
'      SendMessage hProgBar, PBM_STEPIT, 0, 0
'    End If
    
    ' Either of these could be used instead of PBM_STEPIT above
    ' but the progress bar would be hit and redrawn on every itineration.
    SendMessage hProgBar, PBM_SETPOS, ByVal vData, 0
'    SendMessage m_hProgBar, PBM_DELTAPOS, ByVal 1, 0

    If vData = mvarMax Then
        If IsWindow(hProgBar) Then DestroyWindow hProgBar
    End If
End Property


Public Property Get Value() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let Max(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Max = 5

    mvarMax = vData
    
    ' Set the range of the progess bar.
    ' (Minimum range = low word, Maximum range = high word).
    SendMessage hProgBar, PBM_SETRANGE, 1, ByVal (vData * &H10000)

    ' Set the value of the highlight increment. We''ll set it to 100
    ' itins here for the example even though it's the default value.
    ' SendMessage hProgBar, PBM_SETSTEP, ByVal (vData \ 100), 0
    
End Property


Public Property Get Max() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Max
    Max = mvarMax
End Property

Public Property Get Panel() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Panel
    Panel = mvarPanel
End Property

Private Sub Class_Terminate()
    DestroyWindow hProgBar
End Sub


