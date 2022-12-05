VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl HighlightSintax 
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ScaleHeight     =   2490
   ScaleWidth      =   4080
   ToolboxBitmap   =   "HighlightSintax.ctx":0000
   Begin RichTextLib.RichTextBox rich 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      RightMargin     =   20000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HighlightSintax.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "HighlightSintax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event SelChange()
Public Event Change()
Public Event Click()
Public Event DblClick()

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
Public Event Validate(Cancel As Boolean)

Public Enum ItemCodeType
    enumKeyword = 1
    enumOperator = 2
    enumFunction = 3
    enumDelimiter = 4
End Enum

Public Enum ProgrammingLanguage
    hlNOHighLight = 0
    hlVisualBasic = 1
    hlJava = 2
    hlhtml = 3
    [SQL Server] = 4
    [AMC Script Languaje] = 5
End Enum

Public Enum enumHighlightCode
    hlOnNewLine = 0
    hlAsType = 1
End Enum

Public CompareCase As VbCompareMethod
Public GiveCorrectCase As Boolean

Private bFireSelectionChange As Boolean
Private bListenForChange As Boolean
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private strSeparator(14) As String
Private iSeparatorCount As Integer

Private m_Language As ProgrammingLanguage
Dim HighLightWords() As HightlightedWord
Dim mHighlightCode As enumHighlightCode

Private Type HightlightedWord
    Word As String
    WordType As ItemCodeType
End Type

Private Type CommentTag
    CommentStart As String
    CommentEnd As String
End Type

Private m_Comment() As CommentTag
Private m_CommentCount As Integer

Dim WordCount As Integer

Dim mKeywordColor As OLE_COLOR
Dim mOperatorColor As OLE_COLOR
Dim mDelimiterColor As OLE_COLOR
Dim mForeColor As OLE_COLOR
Dim mFunctionColor As OLE_COLOR

Dim strKeywordColor As String
Dim strOperatorColor As String
Dim strDelimiterColor As String
Dim strForeColor As String
Dim strFunctionColor As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_FMTLINES = &HC8
Private Const EM_LINELENGTH = &HC1
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2
Private Const EC_USEFONTINFO = &HFFFF
Private Const EM_SETMARGINS = &HD3
Private Const EM_GETMARGINS = &HD4
Private Const EM_CANUNDO = &HC6
Private Const EM_EMPTYUNDOBUFFER = &HCD
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETHANDLE = &HBD
Private Const EM_GETMODIFY = &HB8
Private Const EM_GETPASSWORDCHAR = &HD2
Private Const EM_GETRECT = &HB2
Private Const EM_GETSEL = &HB0
Private Const EM_GETTHUMB = &HBE
Private Const EM_GETWORDBREAKPROC = &HD1
Private Const EM_LIMITTEXT = &HC5
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB

Private Const EM_LINESCROLL = &HB6
Private Const EM_REPLACESEL = &HC2
Private Const EM_SCROLL = &HB5
Private Const EM_SCROLLCARET = &HB7
Private Const EM_SETHANDLE = &HBC
Private Const EM_SETMODIFY = &HB9
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETRECT = &HB3
Private Const EM_SETRECTNP = &HB4
Private Const EM_SETSEL = &HB1
Private Const EM_SETTABSTOPS = &HCB
Private Const EM_SETWORDBREAKPROC = &HD0
Private Const EM_UNDO = &HC7

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Function ColorWord(ByVal sWord As String) As String
Dim iWord As Integer
    For iWord = 0 To WordCount - 1
        If StrComp(sWord, HighLightWords(iWord).Word, CompareCase) = 0 Then
            If GiveCorrectCase Then sWord = HighLightWords(iWord).Word
            ColorWord = "\cf" & HighLightWords(iWord).WordType & " " & sWord & "\cf0 "
            Exit Function
        End If
    Next
    ColorWord = "\cf0 " & sWord & "\cf0 "
End Function

Private Function GetRTFColor(Color As OLE_COLOR) As String
    Dim lrgb As Long
    lrgb = TranslateColor(Color)
    GetRTFColor = "\red" & (lrgb And &HFF&) & "\green" & (lrgb And &HFF00&) \ &H100 & "\blue" & (lrgb And &HFF0000) \ &H10000 & ";"
End Function

Private Function GetWord(sBlock As String, lngWordStart As Long, lngCharPos As Long, sSep As String) As String
    Dim sWord As String
On Error GoTo en

    sWord = Mid$(sBlock, lngWordStart, lngCharPos - lngWordStart)
        If sSep = vbCrLf Then
            sSep = "\par " & vbCrLf
        ElseIf sSep = vbTab Then
                sSep = "\tab "
        ElseIf sSep = "\" Then
                sSep = "\cf2 \\\cf0 "
        ElseIf sSep = "{" Then
                sSep = "\cf2 \{\cf0 "
        ElseIf sSep = "}" Then
                sSep = "\cf2 \}\cf0 "
        ElseIf sSep <> " " And Len(sSep) Then
            sSep = "\cf2 " & sSep & "\cf0 "
        End If
        If lngCharPos - lngWordStart > 0 Then
            GetWord = ColorWord(sWord) & sSep
        Else
            GetWord = sSep
        End If
en:
End Function

Private Function HighlightComment(sComment As String, sEndofComment As String) As String
    sComment = Replace(sComment, "\", "\\")
    sComment = Replace(sComment, "{", "\{")
    sComment = Replace(sComment, "}", "\}")
    sComment = Replace(sComment, vbCrLf, "\par ")
    If sEndofComment = vbCrLf Then
        sComment = sComment & "\par" & vbCrLf
    Else
        If sEndofComment = vbTab Then
            sComment = sComment & "\tab "
        Else
            sComment = sComment & sEndofComment
        End If
    End If
    HighlightComment = "\cf4 " & sComment & "\cf0 "
End Function




Private Function StartOfComment(sBlock As String, lngCharPos As Long) As Integer
Dim sChar As String
Dim i As Byte
    For i = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(i).CommentStart))
        If sChar = m_Comment(i).CommentStart Then
            StartOfComment = i
            Exit Function
        End If
    Next
    StartOfComment = -1
End Function

Private Function isSeparator(sBlock As String, lngCharPos As Long) As String
Dim sChar As String
Dim i As Byte
    For i = 0 To iSeparatorCount
        sChar = Mid$(sBlock, lngCharPos, Len(strSeparator(i)))
        If sChar = strSeparator(i) Then
            isSeparator = sChar
            Exit Function
        End If
    Next
End Function

Private Function EndOfComment(sBlock As String, lngCharPos As Long) As Integer
    Dim sChar As String
    Dim i As Byte
    For i = 0 To m_CommentCount - 1
        sChar = Mid$(sBlock, lngCharPos, Len(m_Comment(i).CommentEnd))
        If sChar = m_Comment(i).CommentEnd Then
            EndOfComment = i
            Exit Function
        End If
    Next
    EndOfComment = -1
End Function


Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Sub AddCommentTag(ByVal CommentTagStart As String, ByVal CommentTagEnd As String)
    ReDim Preserve m_Comment(m_CommentCount)
    With m_Comment(m_CommentCount)
        .CommentStart = CommentTagStart
        .CommentEnd = CommentTagEnd
    End With
    m_CommentCount = m_CommentCount + 1
End Sub

Public Property Let BackColor(newColor As OLE_COLOR)
    rich.BackColor = newColor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = rich.BackColor
End Property

Public Property Get Font() As StdFont
    Set Font = rich.Font
End Property
Public Property Set Font(newFont As StdFont)
    Set rich.Font = newFont
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Get FunctionColor() As OLE_COLOR
    FunctionColor = mFunctionColor
End Property


Public Property Let ForeColor(newForeColor As OLE_COLOR)
    mForeColor = newForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Let FunctionColor(newFunctionColor As OLE_COLOR)
    mFunctionColor = newFunctionColor
    strFunctionColor = GetRTFColor(mFunctionColor)
    PropertyChanged "FunctionColor"
End Property

Function HighlightBlock(sBlock As String) As String
    Dim lngCharPos As Long
    Dim lngBlockLength As Long
    Dim sWord As String
    Dim lngCommentStartPos As Long
    Dim byteStartOfComment As Integer
    Dim byteEndOfComment As Integer
    Dim sSep As String
    Dim lngWordStart As Long
    Dim sHighlighted As String
    Dim T As Integer
    Dim bWordFound As Boolean
    Dim bLastStepWasComment As Boolean

    If m_Language = hlNOHighLight Then
        HighlightBlock = sBlock
        Exit Function
    End If
    lngBlockLength = Len(sBlock)
    lngWordStart = 1
    byteStartOfComment = -1
    For lngCharPos = 1 To lngBlockLength
        
        T = StartOfComment(sBlock, lngCharPos)
        If T > -1 And byteStartOfComment = -1 Then
            lngCommentStartPos = lngCharPos
            byteStartOfComment = T
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
            
           If byteStartOfComment > -1 Then
                byteEndOfComment = EndOfComment(sBlock, lngCharPos)
                If byteEndOfComment > -1 And byteEndOfComment = byteStartOfComment Then
                    
                    sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), m_Comment(byteEndOfComment).CommentEnd)

                    byteStartOfComment = -1
                    bLastStepWasComment = True
                    lngWordStart = lngCharPos + Len(m_Comment(byteEndOfComment).CommentEnd)
                End If
            Else
                If byteStartOfComment = -1 Then
                    
                    sSep = isSeparator(sBlock, lngCharPos)
                    Dim SepLength As ItemCodeType
                    SepLength = Len(sSep)
                    If SepLength > 0 Then
                        If lngCharPos <= lngBlockLength Then
                            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, sSep)
                        End If
                        lngWordStart = lngCharPos + SepLength
                            bLastStepWasComment = False
                    End If
                End If
            End If
        End If
    Next
    If byteStartOfComment > -1 Then
        
        
        Dim lngCommentEndPos As Long
        lngCommentEndPos = InStr(lngCharPos, rich.Text, m_Comment(byteStartOfComment).CommentEnd)
        If lngCommentEndPos = 0 Then lngCommentEndPos = Len(rich.Text)
        sHighlighted = sHighlighted & HighlightComment(Mid$(sBlock, lngCommentStartPos, (lngCharPos - lngCommentStartPos)), "")
    Else
        If bLastStepWasComment Then
            sHighlighted = sHighlighted & GetWord(sBlock, lngWordStart, lngCharPos, "")
        Else
            If lngBlockLength - lngWordStart >= 0 Then
                sWord = Mid$(sBlock, lngWordStart, (lngBlockLength - lngWordStart) + 1)
                sHighlighted = sHighlighted & ColorWord(sWord)
            End If
        End If
    End If
    If Len(sHighlighted) = 0 Then Exit Function
    HighlightBlock = "{{\colortbl ;" & strKeywordColor & strOperatorColor & strFunctionColor & strDelimiterColor & "}" & sHighlighted & "}"
End Function

Public Property Get HighlightCode() As enumHighlightCode
    HighlightCode = mHighlightCode
End Property

Public Property Let HighlightCode(newHighlightCode As enumHighlightCode)
    mHighlightCode = newHighlightCode
    PropertyChanged "HighlightCode"
End Property



Public Property Get KeywordColor() As OLE_COLOR
    KeywordColor = mKeywordColor
End Property

Public Property Get DelimiterColor() As OLE_COLOR
    DelimiterColor = mDelimiterColor
End Property


Public Property Let DelimiterColor(newDelimiterColor As OLE_COLOR)
    mDelimiterColor = newDelimiterColor
    strDelimiterColor = GetRTFColor(mDelimiterColor)
    PropertyChanged "DelimiterColor"
End Property


Public Property Get Line(lngLine As Long) As String
Dim bReturnedLineBuffer() As Byte
Dim LengthOfLine As Long
Dim LineStart As Long

    LineStart = LineStartPos(LineIndex)
    If LineStart = -1 Then Exit Function
    
    LengthOfLine = LineLength(LineStart)
    If LengthOfLine < 1 Then Exit Function
    
    
    ReDim bReturnedLineBuffer(LengthOfLine)

    bReturnedLineBuffer(0) = LengthOfLine And 255
    bReturnedLineBuffer(1) = LengthOfLine \ 256


    SendMessage rich.hwnd, EM_GETLINE, LineIndex, bReturnedLineBuffer(0)

    Line = Left$(StrConv(bReturnedLineBuffer, vbUnicode), LengthOfLine)
End Property

Public Property Let LineIndex(lngNewLineIndex As Long)
    rich.SelStart = Abs(LineStartPos(lngNewLineIndex))
End Property

Public Property Get LineLength(CharacterIndex As Long) As Long
    LineLength = SendMessage(rich.hwnd, EM_LINELENGTH, CharacterIndex, 0&)
End Property

Public Property Get LineStartPos(ByVal LineIndex As Long) As Long
    LineStartPos = SendMessage(rich.hwnd, EM_LINEINDEX, LineIndex, 0&)
End Property


Public Property Get LineIndex() As Long
    LineIndex = SendMessage(rich.hwnd, EM_LINEFROMCHAR, ByVal -1, 0&)
End Property


Public Sub LoadFile(strFilename)
    Dim FileNum As Integer
    Dim sData As String
    Dim bListen As Boolean
    bListen = bListenForChange
    bListenForChange = False
    
    FileNum = FreeFile
    Open strFilename For Input As FileNum
        sData = Input(LOF(FileNum), FileNum)
    Close FileNum
    bFireSelectionChange = False
    rich.TextRTF = ""
    rich.SelRTF = HighlightBlock(sData)
    bFireSelectionChange = True
bListenForChange = bListen
End Sub

Public Property Get OperatorColor() As OLE_COLOR
    OperatorColor = mOperatorColor
End Property


Public Property Let KeywordColor(newKeywordColor As OLE_COLOR)
    mKeywordColor = newKeywordColor
    strKeywordColor = GetRTFColor(mKeywordColor)
    PropertyChanged "KeywordColor"
End Property


Public Property Let OperatorColor(newOperatorColor As OLE_COLOR)
    mOperatorColor = newOperatorColor
    strOperatorColor = GetRTFColor(mOperatorColor)
    PropertyChanged "OperatorColor"
End Property



Public Sub SaveFile(strFilename As String)
    rich.SaveFile strFilename, rtfText
End Sub

Public Property Let SelLength(lngNewSelLength As Long)
    rich.SelLength = lngNewSelLength
End Property
Public Property Get SelLength() As Long
    SelLength = rich.SelLength
End Property
Public Property Get SelStart() As Long
    SelStart = rich.SelStart
End Property

Public Property Let SelStart(lngNewSelStart As Long)
    rich.SelStart = lngNewSelStart
End Property

Private Sub SetJava()
    WordCount = 0
    AddWord "abstract"
    AddWord "boolean"
    AddWord "break"
    AddWord "byte"
    AddWord "case"
    AddWord "catch"
    AddWord "char"
    AddWord "class"
    AddWord "const"
    AddWord "continue"
    AddWord "default"
    AddWord "do"
    AddWord "double"
    AddWord "else"
    AddWord "extends"
    AddWord "final"
    AddWord "finally"
    AddWord "float"
    AddWord "for"
    AddWord "goto"
    AddWord "if"
    AddWord "implements"
    AddWord "import"
    AddWord "instanceof"
    AddWord "int"
    AddWord "interface"
    AddWord "long"
    AddWord "native"
    AddWord "new"
    AddWord "package"
    AddWord "private"
    AddWord "protected"
    AddWord "public"
    AddWord "return"
    AddWord "short"
    AddWord "static"
    AddWord "super"
    AddWord "switch"
    AddWord "synchronized"
    AddWord "this"
    AddWord "throw"
    AddWord "throws"
    AddWord "transient"
    AddWord "try"
    AddWord "void"
    AddWord "volatitle"
    AddWord "while"


    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "*", enumOperator
    AddWord "/", enumOperator
    AddWord "%", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord ">=", enumOperator
    AddWord "<=", enumOperator
    AddWord "!=", enumOperator
    AddWord "==", enumOperator
    AddWord "!", enumOperator
    AddWord "&&", enumOperator
    AddWord "||", enumOperator
    AddWord "-", enumOperator
    AddWord "&", enumOperator
    AddWord "|", enumOperator
    AddWord "^", enumOperator
    AddWord "<<", enumOperator
    AddWord ">>", enumOperator
    AddWord ">>>", enumOperator
    
    AddWord "=", enumOperator
    AddWord "++", enumOperator
    AddWord "--", enumOperator
    AddWord "+=", enumOperator
    AddWord "-=", enumOperator
    AddWord "*=", enumOperator
    AddWord "/=", enumOperator
    AddWord "%=", enumOperator
    AddWord "|=", enumOperator
    AddWord "&=", enumOperator
    AddWord "^=", enumOperator
    AddWord "<<=", enumOperator
    AddWord ">>=", enumOperator
    AddWord ">>>=", enumOperator
    AddWord "new", enumOperator
    AddWord "?", enumOperator
    AddWord ":", enumOperator
    AddWord "(", enumOperator
    AddWord ")", enumOperator
    AddWord "{", enumOperator
    AddWord "}", enumOperator
    
    AddWord "true", enumOperator
    AddWord "false", enumOperator

    CompareCase = vbBinaryCompare
    ReDim Preserve m_Comment(0)
    m_CommentCount = 0
    AddWord """", enumDelimiter
    AddCommentTag "//", vbCrLf
    AddCommentTag "/*", "*/"
    AddCommentTag "/**", "*/"
    GiveCorrectCase = False
End Sub


Private Sub SetAMCPLT()
    
    WordCount = 0
    Erase m_Comment
    m_CommentCount = 0
    
    AddWord "db.ConnectionString", enumKeyword
    AddWord "db.Svr.Name", enumKeyword
    AddWord "db.Name", enumKeyword
    AddWord "db.Type", enumKeyword
    AddWord "db.CreateDate", enumKeyword
    AddWord "db.Tbl.Name", enumKeyword
    AddWord "db.Tbl.Owner", enumKeyword
    AddWord "db.Tbl.Type", enumKeyword
    AddWord "db.Tbl.CreateDate", enumKeyword
    AddWord "db.Tbl.FieldsCount", enumKeyword
    
    AddWord "Tbl.Fld.ADOConstantValue", enumKeyword
    AddWord "Tbl.Fld.DefineSize", enumKeyword
    AddWord "Tbl.Fld.IsIdentity", enumKeyword
    AddWord "Tbl.Fld.IsNull", enumKeyword
    AddWord "Tbl.Fld.IsPK", enumKeyword
    AddWord "Tbl.Fld.Name", enumKeyword
    AddWord "Tbl.Fld.NumericScale", enumKeyword
    AddWord "Tbl.Fld.Precision", enumKeyword
    AddWord "Tbl.Fld.Type", enumKeyword
    AddWord "Tbl.Fld.StringType", enumKeyword
    AddWord "Tbl.Fld.PosDataBase", enumKeyword
    
    AddWord "vb.Fld.Prefix", enumKeyword
    AddWord "vb.Fld.StringType", enumKeyword
    
    
    AddWord "Str.CrLf", enumKeyword
    AddWord "Str.Tab", enumKeyword
    AddWord "Str.Comma", enumKeyword
    
    

    AddWord "CG.Legal", enumKeyword
    AddWord "CG.Author", enumKeyword
    
    
    AddWord "Sys.Time", enumKeyword
    AddWord "Sys.Date", enumKeyword
    
    
    
    AddWord "If", enumKeyword
    AddWord "Then", enumKeyword
    AddWord "Else", enumKeyword
    AddWord "EndIf", enumKeyword
    AddWord "Do", enumKeyword
    AddWord "Until", enumKeyword
    AddWord "Loop", enumKeyword
    AddWord "Set", enumKeyword
    AddWord "Var", enumKeyword
    AddWord ";", enumKeyword
    
    AddWord "db.Fld.First", enumKeyword
    AddWord "db.Fld.Next", enumKeyword
    AddWord "db.Fld.Previous", enumKeyword
    AddWord "db.Fld.Last", enumKeyword
    
    AddWord "Write", enumFunction
    AddWord "Bracket", enumFunction
    AddWord "CompressName", enumFunction
    AddWord "Trim", enumFunction
    AddWord "Tab", enumFunction
    AddWord "CrLf", enumFunction
    AddWord "SkipLastChar", enumFunction
    AddWord "LCase", enumFunction
    AddWord "UCase", enumFunction
    AddWord "Len", enumFunction
    AddWord "IsNumeric", enumFunction
    AddWord "IsOperator", enumFunction
    AddWord "InStr", enumFunction
    AddWord "Right", enumFunction
    AddWord "Mid", enumFunction
    
    AddWord "=", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "<=", enumOperator
    AddWord ">=", enumOperator
    AddWord "=<", enumOperator
    AddWord "=>", enumOperator
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "/", enumOperator
    AddWord "*", enumOperator
    AddWord "<>", enumOperator
    AddWord "&", enumOperator
    
    AddWord """", enumDelimiter
    
    CompareCase = vbTextCompare
    AddCommentTag "/-", "-/"
    
    CompareCase = vbTextCompare
    GiveCorrectCase = True
    
End Sub


Private Sub SetVB()
    WordCount = 0
    Erase m_Comment
    m_CommentCount = 0
    
    AddWord "#Const"
    AddWord "#Else"
    AddWord "#ElseIf"
    AddWord "#End If"
    AddWord "#If"
    AddWord "Alias"
    AddWord "And"
    AddWord "As"
    AddWord "Base"
    AddWord "Binary"
    AddWord "Boolean"
    AddWord "Byte"
    AddWord "ByVal"
    AddWord "Call"
    AddWord "Case"
    AddWord "CBool"
    AddWord "CByte"
    AddWord "CCur"
    AddWord "CDate"
    AddWord "CDbl"
    AddWord "CDec"
    AddWord "CInt"
    AddWord "CLng"
    AddWord "Close"
    AddWord "Compare"
    AddWord "Const"
    AddWord "CSng"
    AddWord "CStr"
    AddWord "Currency"
    AddWord "CVar"
    AddWord "CVErr"
    AddWord "Decimal"
    AddWord "Declare"
    AddWord "DefBool"
    AddWord "DefByte"
    AddWord "DefCur"
    AddWord "DefDate"
    AddWord "DefDbl"
    AddWord "DefDec"
    AddWord "DefInt"
    AddWord "DefLng"
    AddWord "DefObj"
    AddWord "DefSng"
    AddWord "DefStr"
    AddWord "DefVar"
    AddWord "Dim"
    AddWord "Do"
    AddWord "Double"
    AddWord "Each"
    AddWord "Else"
    AddWord "ElseIf"
    AddWord "End"
    AddWord "Enum"
    AddWord "Eqv"
    AddWord "Erase"
    AddWord "Error"
    AddWord "Exit"
    AddWord "Explicit"
    AddWord "False"
    AddWord "For"
    AddWord "Function"
    AddWord "Get"
    AddWord "Global"
    AddWord "GoSub"
    AddWord "GoTo"
    AddWord "If"
    AddWord "Imp"
    AddWord "In"

    AddWord "Integer"
    AddWord "Is"
    AddWord "LBound"
    AddWord "Let"
    AddWord "Lib"
    AddWord "Like"
    AddWord "Line"
    AddWord "Lock"
    AddWord "Long"
    AddWord "Loop"
    AddWord "LSet"
    AddWord "Name"
    AddWord "New"
    AddWord "Next"
    AddWord "Not"
    AddWord "Object"
    AddWord "On"
    AddWord "Open"
    AddWord "Option"
    AddWord "Optional"
    AddWord "Or"
    AddWord "Output"
    AddWord "Print"
    AddWord "Private"
    AddWord "Property"
    AddWord "Public"
    AddWord "Put"
    AddWord "Random"
    AddWord "Read"
    AddWord "ReDim"
    AddWord "Resume"
    AddWord "Return"
    AddWord "RSet"
    AddWord "Seek"
    AddWord "Select"
    AddWord "Set"
    AddWord "Single"
    AddWord "Spc"
    AddWord "Static"
    AddWord "String"
    AddWord "Stop"
    AddWord "Sub"
    AddWord "Tab"
    AddWord "Then"
    AddWord "True"
    AddWord "Type"
    AddWord "UBound"
    AddWord "Unlock"
    AddWord "Variant"
    AddWord "Wend"
    AddWord "While"
    AddWord "With"
    AddWord "Nothing"
    AddWord "To"
    
    AddWord "Input"

    AddWord "MsgBox", enumFunction
    
    AddWord "Xor", enumOperator
    AddWord "=", enumOperator
    AddWord ">", enumOperator
    AddWord "<", enumOperator
    AddWord "<=", enumOperator
    AddWord ">=", enumOperator
    AddWord "=<", enumOperator
    AddWord "=>", enumOperator
    AddWord "+", enumOperator
    AddWord "-", enumOperator
    AddWord "/", enumOperator
    AddWord "*", enumOperator
    AddWord "<>", enumOperator
    AddWord "&", enumOperator

    AddWord """", enumDelimiter
    CompareCase = vbTextCompare
    AddCommentTag "'", vbCrLf
    GiveCorrectCase = True
End Sub

Private Sub SetSQL()
    WordCount = 0
    Erase m_Comment
    m_CommentCount = 0

   AddWord "Add"
   AddWord "All"
   AddWord "Alter"
   AddWord "And"
   AddWord "Any"
   AddWord "As"
   AddWord "Asc"
   AddWord "Authorization"
   AddWord "Avg"
   AddWord "Backup"
   AddWord "Begin"
   AddWord "Between"
   AddWord "Break"
   AddWord "Browse"
   AddWord "Bulk"
   AddWord "By"
   AddWord "Cascade"
   AddWord "Case"
   AddWord "Check"
   AddWord "Checkpoint"
   AddWord "Close"
   AddWord "Clustered"
   AddWord "Coalesce"
   AddWord "Column"
   AddWord "Commit"
   AddWord "Committed"
   AddWord "Compute"
   AddWord "Confirm"
   AddWord "Constraint"
   AddWord "Contains"
   AddWord "Containstable"
   AddWord "Continue"
   AddWord "Controlrow"
   AddWord "Convert"
   AddWord "Count"
   AddWord "Create"
   AddWord "Cross"
   AddWord "Current"
   AddWord "Current_date"
   AddWord "Current_time"
   AddWord "Current_timestamp"
   AddWord "Current_user"
   AddWord "Cursor"
   AddWord "Database"
   AddWord "Dbcc"
   AddWord "Deallocate"
   AddWord "Declare"
   AddWord "Default"
   AddWord "Delete"
   AddWord "Deny"
   AddWord "Desc"
   AddWord "Disk"
   AddWord "Distinct"
   AddWord "Distributed"
   AddWord "Double"
   AddWord "Drop"
   AddWord "Dummy"
   AddWord "Dump"
   AddWord "Else"
   AddWord "End"
   AddWord "Errlvl"
   AddWord "Errorexit"
   AddWord "Escape"
   AddWord "Except"
   AddWord "Exec"
   AddWord "Execute"
   AddWord "Exists"
   AddWord "Exit"
   AddWord "Fetch"
   AddWord "File"
   AddWord "Fillfactor"
   AddWord "Floppy"
   AddWord "For"
   AddWord "Foreign"
   AddWord "Freetext"
   AddWord "Freetexttable"
   AddWord "From"
   AddWord "Full"
   AddWord "Goto"
   AddWord "Grant"
   AddWord "Group"
   AddWord "Having"
   AddWord "Holdlock"
   AddWord "Identity"
   AddWord "Identity_insert"
   AddWord "Identitycol"
   AddWord "If"
   AddWord "In"
   AddWord "Index"
   AddWord "Inner"
   AddWord "Insert"
   AddWord "Intersect"
   AddWord "Into"
   AddWord "Is"
   AddWord "Isolation"
   AddWord "Join"
   AddWord "Key"
   AddWord "Kill"
   AddWord "Left"
   AddWord "Level"
   AddWord "Like"
   AddWord "Lineno"
   AddWord "Load"
   AddWord "Max"
   AddWord "Min"
   AddWord "Mirrorexit"
   AddWord "National"
   AddWord "Nocheck"
   AddWord "Nonclustered"
   AddWord "Not"
   AddWord "Null"
   AddWord "Nullif"
   AddWord "Of"
   AddWord "Off"
   AddWord "Offsets"
   AddWord "On"
   AddWord "Once"
   AddWord "Only"
   AddWord "Open"
   AddWord "Opendatasource"
   AddWord "Openquery"
   AddWord "Openrowset"
   AddWord "Option"
   AddWord "Or"
   AddWord "Order"
   AddWord "Outer"
   AddWord "Over"
   AddWord "Percent"
   AddWord "Perm"
   AddWord "Permanent"
   AddWord "Pipe"
   AddWord "Plan"
   AddWord "Precision"
   AddWord "Prepare"
   AddWord "Primary"
   AddWord "Print"
   AddWord "Privileges"
   AddWord "Proc"
   AddWord "Procedure"
   AddWord "Processexit"
   AddWord "Public"
   AddWord "Raiserror"
   AddWord "Read"
   AddWord "Readtext"
   AddWord "Reconfigure"
   AddWord "References"
   AddWord "Repeatable"
   AddWord "Replication"
   AddWord "Restore"
   AddWord "Restrict"
   AddWord "Return"
   AddWord "Revoke"
   AddWord "Right"
   AddWord "Rollback"
   AddWord "Rowcount"
   AddWord "Rowguidcol"
   AddWord "Rule"
   AddWord "Save"
   AddWord "Schema"
   AddWord "Select"
   AddWord "Serializable"
   AddWord "Session_user"
   AddWord "Set"
   AddWord "Setuser"
   AddWord "Shutdown"
   AddWord "Some"
   AddWord "Statistics"
   AddWord "Sum"
   AddWord "System_user"
   AddWord "Table"
   AddWord "Tape"
   AddWord "Temp"
   AddWord "Temporary"
   AddWord "Textsize"
   AddWord "Then"
   AddWord "To"
   AddWord "Top"
   AddWord "Tran"
   AddWord "Transaction"
   AddWord "Trigger"
   AddWord "Truncate"
   AddWord "Tsequal"
   AddWord "Uncommitted"
   AddWord "Union"
   AddWord "Unique"
   AddWord "Update"
   AddWord "Updatetext"
   AddWord "Use"
   AddWord "User"
   AddWord "Values"
   AddWord "Varying"
   AddWord "View"
   AddWord "Waitfor"
   AddWord "When"
   AddWord "Where"
   AddWord "While"
   AddWord "With"
   AddWord "Work"
   AddWord "Writetext"
   
   
   AddWord "SmallInt"
   AddWord "Int"
   AddWord "Real"
   AddWord "Float"
   AddWord "Money"
   AddWord "Bit"
   AddWord "TinyInt"
   AddWord "Binary"
   AddWord "Char"
   AddWord "Numeric"
   AddWord "Datetime"
   AddWord "SmallDateTime"
   AddWord "Varchar"
   AddWord "Text"
   AddWord "VarBinary"
   AddWord "Image"
   AddWord "Decimal"
   AddWord "nChar"
   AddWord "nText"
   AddWord "nVarChar"

   AddWord "+", enumOperator
   AddWord "-", enumOperator
   AddWord "*", enumOperator
   AddWord "/", enumOperator
   AddWord "%", enumOperator
   AddWord ">", enumOperator
   AddWord "<", enumOperator
   AddWord ">=", enumOperator
   AddWord "<=", enumOperator
   AddWord "!=", enumOperator
   AddWord "==", enumOperator
   AddWord "!", enumOperator
   AddWord "&&", enumOperator
   AddWord "||", enumOperator
   AddWord "-", enumOperator
   AddWord "&", enumOperator
   AddWord "|", enumOperator
   AddWord "^", enumOperator
   AddWord "<<", enumOperator
   AddWord ">>", enumOperator
   AddWord ">>>", enumOperator
   
   AddWord "=", enumOperator
   AddWord "++", enumOperator
   AddWord "*=", enumOperator
   AddWord "/=", enumOperator
   AddWord "%=", enumOperator
   AddWord "&=", enumOperator
   AddWord "^=", enumOperator
   AddWord "?", enumOperator
   AddWord ":", enumOperator
   AddWord "(", enumOperator
   AddWord ")", enumOperator
   AddWord "@", enumOperator
   
   AddWord "true", enumOperator
   AddWord "false", enumOperator
   AddWord "@@Error", enumFunction
   AddWord "@@Identity", enumFunction
   AddWord "Count", enumFunction
   AddWord "Convert", enumFunction

   AddWord "'", enumDelimiter
   
   AddCommentTag "//", vbCrLf
   AddCommentTag "--", vbCrLf
   AddCommentTag "/*", "*/"
   
    
   CompareCase = vbTextCompare
   GiveCorrectCase = True


End Sub



Private Sub rich_Change()
    RaiseEvent Change
End Sub

Private Sub rich_Click()
    RaiseEvent Click
End Sub


Private Sub rich_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rich_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeyTab Then ' Indent
        Dim SelStart As Long
        If rich.SelLength > 0 Then
            Dim strLines() As String
            Dim LineCount As Long, i As Long
            Dim strResult As String
            strLines = Split(rich.SelText, vbCrLf)
            LineCount = UBound(strLines)
            If LineCount > 0 Then
                SelStart = rich.SelStart
                For i = 0 To LineCount - 1
                    strResult = strResult & vbTab & strLines(i) & vbCrLf
                Next
                strResult = strResult & vbTab & strLines(i)
                rich.SelText = strResult
                rich.SelStart = SelStart
                rich.SelLength = Len(strResult)
                KeyCode = 0
            End If
        End If
    End If

End Sub

Private Sub rich_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim i As Byte
    If mHighlightCode = hlAsType Then
        For i = 0 To iSeparatorCount
            If KeyAscii = Asc(strSeparator(i)) Then
                    LockWindowUpdate rich.hwnd
                    bFireSelectionChange = False
                    Dim TheStart As Long
                    TheStart = rich.SelStart
                    rich.SelStart = Me.LineStartPos(Me.LineIndex)
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(Me.LineIndex))
                    rich.SelStart = TheStart
                    LockWindowUpdate 0
                    bFireSelectionChange = True
                Exit Sub
            End If
        Next
    End If
End Sub


Private Sub rich_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rich_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub rich_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub rich_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub rich_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub


Private Sub rich_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub rich_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub rich_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub rich_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub rich_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub rich_SelChange()
    Static lngLastLine As Long
    Dim lngNewLine As Long
    Dim TheStart As Long
    
    If bFireSelectionChange Then
        If rich.SelLength = 0 Then
                bFireSelectionChange = False
                lngNewLine = Me.LineIndex
                If lngNewLine <> lngLastLine Then
                    On Error GoTo en:
                    LockWindowUpdate rich.hwnd
                    TheStart = rich.SelStart
                    rich.SelStart = Me.LineStartPos(lngLastLine)
                    rich.SelLength = Me.LineLength(rich.SelStart)
                    rich.SelRTF = HighlightBlock(Line(lngLastLine))
en:
                    rich.SelStart = TheStart
                    rich.SelLength = SelLength
                    LockWindowUpdate 0
                End If
                lngLastLine = lngNewLine
                bFireSelectionChange = True
        End If
    RaiseEvent SelChange
    End If
End Sub

Private Sub rich_Validate(Cancel As Boolean)
    'RaiseEvent Validate(Cancel)
End Sub

Private Sub UserControl_Initialize()
    strSeparator(0) = " "
    strSeparator(1) = vbCrLf
    strSeparator(2) = vbTab
    strSeparator(3) = "("
    strSeparator(4) = ")"
    strSeparator(5) = "="
    strSeparator(6) = "+"
    strSeparator(7) = "-"
    strSeparator(8) = "*"
    strSeparator(9) = ">"
    strSeparator(10) = "<"
    strSeparator(11) = "\"
    strSeparator(12) = "/"
    strSeparator(13) = "{"
    strSeparator(14) = "}"
    iSeparatorCount = 14
    bFireSelectionChange = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    rich.Text = PropBag.ReadProperty("Text", "")
    Language = PropBag.ReadProperty("Language", hlNOHighLight)

    rich.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    KeywordColor = PropBag.ReadProperty("KeywordColor", vbBlue)
    OperatorColor = PropBag.ReadProperty("OperatorColor", vbYellow)
    DelimiterColor = PropBag.ReadProperty("DelimiterColor", vbCyan)
    mForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    FunctionColor = PropBag.ReadProperty("FunctionColor", vbMagenta)
    HighlightCode = PropBag.ReadProperty("HighlightCode", 1)
    
    Set rich.Font = PropBag.ReadProperty("Font", rich.Font)
End Sub

Private Sub UserControl_Resize()
    rich.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Sub AddWord(ByVal Word As String, Optional WordType As ItemCodeType = enumKeyword)
    ReDim Preserve HighLightWords(WordCount)
    If WordType = enumDelimiter Then
        AddCommentTag Word, Word
    Else
        With HighLightWords(WordCount)
            .Word = Word
            .WordType = WordType
        End With
        WordCount = WordCount + 1
    End If
End Sub


Public Property Get Text() As String
    Text = rich.Text
End Property

Public Property Get SelText() As String
    SelText = rich.SelText
End Property


Public Property Let SelText(newSelText As String)
bFireSelectionChange = False
    rich.SelRTF = HighlightBlock(newSelText)
    bFireSelectionChange = True
End Property

Public Property Let Text(ByVal vNewValue As String)
    rich.TextRTF = HighlightBlock(vNewValue)
    PropertyChanged "Text"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", rich.Text, ""
    PropBag.WriteProperty "Language", m_Language, hlNOHighLight
    
    PropBag.WriteProperty "BackColor", rich.BackColor, vbWindowBackground
    PropBag.WriteProperty "KeywordColor", mKeywordColor, vbBlue
    PropBag.WriteProperty "OperatorColor", mOperatorColor, vbYellow
    PropBag.WriteProperty "DelimiterColor", mDelimiterColor, vbCyan
    PropBag.WriteProperty "ForeColor", mForeColor, vbWindowText
    PropBag.WriteProperty "FunctionColor", mFunctionColor, vbMagenta
    PropBag.WriteProperty "HighlightCode", mHighlightCode, 1
    
    PropBag.WriteProperty "Font", rich.Font
End Sub


Public Property Get Language() As ProgrammingLanguage
    Language = m_Language
End Property


Public Property Let Language(ByVal vNewValue As ProgrammingLanguage)
Dim sData As String
    If m_Language <> vNewValue Then
        Select Case vNewValue
            Case hlVisualBasic
                SetVB
            Case hlJava
                SetJava
            Case hlhtml
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
                AddCommentTag "<", ">"
             Case [AMC Script Languaje]
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
                SetAMCPLT
            Case hlNOHighLight
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
            Case [SQL Server]
                WordCount = 0
                Erase HighLightWords
                m_CommentCount = 0
                Erase m_Comment
                SetSQL
        End Select
        m_Language = vNewValue

        sData = rich.Text
        rich.TextRTF = ""
        rich.SelRTF = HighlightBlock(sData)
        PropertyChanged "Language"
    End If
End Property



Public Sub Refresh()
   rich.Refresh
End Sub
