VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSCGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Vars of generic properties
Private m_sScript As String
Private m_sCode As String

'vars of database
Private m_sServer As String
Private m_sDataBase As String
Private m_eTypeDB As enmDBType
Private m_sOwner As String
Private m_sConnectionString As String
Private m_sAuthor As String
Private m_sLegal As String

Private m_oTable As CTable
Private iPosField As Integer

Private m_iTabWidth As Integer


'Enumeracion para el tipo de base de datos a utilizar
Public Enum enmDBType
   [None] = 0
   [Microsoft Access] = 1
   [SQL Server] = 2
   [ODBC Database] = 3
End Enum


Public Event Progress(ByVal nPos As Long)


'Vars for internal use in the class
Dim lCount As Long, lTemp As Long
Dim lSetCount As Long, lSetX As Long

Dim bInElse As Boolean, bInIf As Boolean
Dim bInLoop As Boolean, bEndTrad As Boolean

Dim ifstate, loopstate, LoopEnd As Long, LoopStart As Long

'Collections for used variables
Private VarNames As Collection, VarData As Collection, NewVariable
Private SetI As Long, exvar, exset, titleapp, ifstate2, SetX As Long



Public Property Get LengthScript() As Long
   LengthScript = Len(m_sScript)
End Property


Public Property Get Script() As String
   Script = m_sScript
End Property
Public Property Let Script(ByVal vNewValue As String)
   m_sScript = vNewValue
End Property

Public Property Get Code() As String
   Code = m_sCode
End Property
Public Property Let Code(ByVal vNewValue As String)
   m_sCode = vNewValue
End Property

Public Property Get Server() As String
   Server = m_sServer
End Property
Public Property Let Server(ByVal vNewValue As String)
   m_sServer = vNewValue
End Property

Public Property Get DataBase() As String
   DataBase = m_sDataBase
End Property
Public Property Let DataBase(ByVal vNewValue As String)
   m_sDataBase = vNewValue
End Property

Public Property Get TypeDataBase() As enmDBType
   TypeDataBase = m_eTypeDB
End Property
Public Property Let TypeDataBase(ByVal vNewValue As enmDBType)
   m_eTypeDB = vNewValue
End Property



Public Property Get Table() As CTable
   Set Table = m_oTable
End Property
Public Property Let Table(ByVal vNewValue As CTable)
   Set m_oTable = vNewValue
   If m_oTable.Fields.Count > 0 Then
      iPosField = 1
   End If
End Property



Public Property Get Owner() As String
   Owner = m_sOwner
End Property
Public Property Let Owner(ByVal vNewValue As String)
   m_sOwner = vNewValue
End Property

Public Property Get ConnectionString() As String
   ConnectionString = m_sConnectionString
End Property
Public Property Let ConnectionString(ByVal vNewValue As String)
   m_sConnectionString = vNewValue
End Property

Public Property Get Author() As String
   Author = m_sAuthor
End Property
Public Property Let Author(ByVal vNewValue As String)
   m_sAuthor = vNewValue
End Property

Public Property Get Legal() As String
   Legal = m_sLegal
End Property
Public Property Let Legal(ByVal vNewValue As String)
   m_sLegal = vNewValue
End Property

Public Property Get TabWidth() As Integer
   TabWidth = m_iTabWidth
End Property
Public Property Let TabWidth(ByVal vNewValue As Integer)
   m_iTabWidth = vNewValue
End Property




'Funcion para iniciar la generacion de codigo
Public Function Generate() As Boolean
   If Trim(m_sScript) = "" Then
   Else
      If Traduct = True Then
         Generate = True
      Else
         m_sScript = ""
         Generate = False
      End If
   End If
End Function


'Function of init traduction.
Private Function Traduct() As Boolean
   
   AddVariable "True", True
   AddVariable "False", False
   
   For lCount = 1 To Len(m_sScript)
      RaiseEvent Progress(lCount)
      lTemp = 0
      
      'Validate en Traducction
      If bEndTrad = True Then Traduct = True: Exit Function
      
      'Validate If
      If bInElse = True Then
          If LCase(Mid(m_sScript, lCount, 5)) = "endif" Then
              bInElse = False
              lCount = lCount + 5
          End If
      End If
      If bInIf = True Then
          If LCase(Mid(m_sScript, lCount, 5)) = "endif" Then
              bInIf = False
              lCount = lCount + 5
          ElseIf LCase(Mid(m_sScript, lCount, 4)) = "else" Then
              If CodeInStr(lCount, m_sScript, "endif") = 0 Then
                  FlagError "Block if without endif.", True
              Else
                  lCount = CodeInStr(lCount, m_sScript, "endif")
                  bInIf = False
                  bInElse = False
                  lCount = lCount + 5
              End If
          End If
      Else
          If LCase(Mid(m_sScript, lCount, 4)) = "else" Then
              bInElse = True
              bInIf = False
              lCount = lCount + 4
          End If
      End If
      
      
      'For Loops
      If bInLoop = True Then
          If lCount = LoopEnd Then
              lCount = LoopStart
              bInLoop = True
          End If
      End If
      
      
      'Valid comment
      If LCase(Mid(m_sScript, lCount, 2)) = "/-" Then
         lCount = lCount + 2
         'Call WriteText(GetValue(lCount, "-/"))
         lCount = InStr(lCount, m_sScript, "-/", vbBinaryCompare)
         lCount = lCount + 2
         
         
         
      'Move first field
      ElseIf LCase(Mid(m_sScript, lCount, 13)) = "db.fld.first;" Then
         lCount = lCount + 13
         iPosField = 1
      'Move next field
      ElseIf LCase(Mid(m_sScript, lCount, 12)) = "db.fld.next;" Then
         lCount = lCount + 12
         If iPosField < m_oTable.Fields.Count Then
            iPosField = iPosField + 1
         End If
         
      'Move previous field
      ElseIf LCase(Mid(m_sScript, lCount, 16)) = "db.fld.previous;" Then
         lCount = lCount + 15
         If iPosField > 1 Then
            iPosField = iPosField - 1
         End If
      'Move last field
      ElseIf LCase(Mid(m_sScript, lCount, 12)) = "db.fld.last;" Then
         lCount = lCount + 12
         iPosField = m_oTable.Fields.Count
      
      'Write text
      ElseIf LCase(Mid(m_sScript, lCount, 6)) = "write " Then
         lCount = lCount + 6
         Call WriteText(GetValue(lCount, ";"))
         lCount = lSetCount
      'create a new variable
      ElseIf LCase(Mid(m_sScript, lCount, 4)) = "var " Then
         lCount = lCount + 4
         NewVariable = GetValue2(lCount, ";")
         lCount = lSetCount
         AddVariable NewVariable, ""
         NewVariable = ""
      'Set values on Variables
      ElseIf LCase(Mid(m_sScript, lCount, 4)) = "set " Then
        lCount = lCount + 4
        exvar = GetValue2(lCount, "=")
        lCount = lSetCount
        exset = GetValue(lCount, ";")
        lCount = lSetCount
        SetVar exvar, exset
        exvar = ""
        exset = ""
        
      'if statments
      ElseIf LCase(Mid(m_sScript, lCount, 3)) = "if " Then
         lCount = lCount + 3
         'Get expression
         ifstate = GetValue2(lCount, "then")
         lCount = lSetCount
         'If it's true, InIf = True
         If Eval(ifstate) = True Then
             bInIf = True
             bInElse = False
         'If not, we have to look for else
         ElseIf Eval(ifstate) = False Then
             If StringExist(lCount, m_sScript, "else") = False Then
                 'If there is not else, we need to find endif
                 If StringExist(lCount, m_sScript, "endif") = False Then
                     'No endif?  Flag an error
                     FlagError "Syntax Error: Block if without endif.", True
                 Else
                     'Set location at endif
                     lCount = CodeInStr(lCount, m_sScript, "endif")
                     bInElse = False
                     bInIf = False
                 End If
             Else
                 'Set location at else
                 lCount = CodeInStr(lCount, m_sScript, "else")
                 bInElse = True
                 bInIf = False
             End If
         End If
         ifstate = ""
        
        
      'Loops
      ElseIf LCase(Mid(m_sScript, lCount, 9)) = "do until " Then
          LoopStart = lCount
          lCount = lCount + 9
          
          If bEndTrad = True Then Exit Function
          
          loopstate = GetValue2(lCount, ";")
          lCount = lSetCount
          
          If StringExist(lCount, m_sScript, "loop") = False Then
              FlagError "Syntax Error: Block do without loop: " & lCount, True
          End If
          
          If bEndTrad = True Then Exit Function
          
          LoopEnd = CodeInStr(lCount, m_sScript, "loop")
          
          If Eval(loopstate) = True Then
              lCount = LoopEnd + 4
              bInLoop = False
          Else
              bInLoop = True
          End If
          loopstate = ""
        
      End If
   Next lCount
End Function




'This function write Text in var generic of result
Private Sub WriteText(ByVal vsText As String)
   m_sCode = m_sCode & vsText
End Sub



'This will search for the paramter, starting at the
'specified location, skipping things in quotes, then
'evalutate it.
Private Function GetValue(ByVal vlStart As Long, ByVal vsParam As String, Optional ByVal vsInWath As String) As Variant
Dim sEnCode As String, lWhereIsIt As String
On Error GoTo ErrorGetValue
   If Trim(vsInWath) = "" Then vsInWath = m_sScript
   
   lWhereIsIt = CodeInStr(vlStart, vsInWath, vsParam)
   
   If lWhereIsIt = 0 Then
      FlagError "Expected '" & vsParam & "', but found end of template.", True
   Else
    'Get the code between the starting location and the paramter
    sEnCode = Mid(vsInWath, vlStart, lWhereIsIt - vlStart)
    'Evalutate it
    GetValue = Solve(sEnCode)
    'Set the location
    lSetCount = vlStart + Len(sEnCode)
      
   End If
Exit Function
ErrorGetValue:
   FlagError "Sintax Error, '" & Error & "'", True
   Exit Function
End Function




Private Function CodeInStr(StartPos As Long, SourceText As String, ToFind As String) As Long
'This is just like InStr() only it skips over ""'s
Dim iPend As Long, i As Long
Dim sTemp As String

For i = StartPos To Len(SourceText)
    If LCase(Mid(SourceText, i, Len(ToFind))) = LCase(ToFind) Then
        CodeInStr = i
        Exit Function
    End If
                          
    If Mid(SourceText, i, 1) = Chr(34) Then
        i = InStr(i + 1, SourceText, Chr(34))
    End If
Next i
End Function


'Solve Function
Private Function Solve(vsFunction As String) As Variant
'This is the base solver
   vsFunction = GetFunctions(vsFunction) 'First clear all functions out
   vsFunction = Equation(vsFunction)
   Solve = SolveFunction(vsFunction)    'Then solve
End Function



Private Function GetFunctions(vsFunction As String) As Variant
'This sub searches for functions and solves them
'On Error Resume Next
Dim TempFunction As String
Dim Arg1 As String, Arg2 As String, Arg3 As String
Dim X As Long, sChar As String
Dim iCounter As Integer, iCicle As Integer, sTempCad As String

   sTempCad = ""
   iCounter = 0
   iCicle = 0
   
   For X = 1 To Len(vsFunction)
      sChar = Mid(vsFunction, X, 1)
      'Bracket
      If LCase(Mid(vsFunction, X, 8)) = "bracket(" Then
          X = X + 8
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & Bracket(Arg1) & Chr(34)
      'CompressName
      ElseIf LCase(Mid(vsFunction, X, 13)) = "compressname(" Then
          X = X + 13
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & CompressName(Arg1) & Chr(34)
      
      'Trim
      ElseIf LCase(Mid(vsFunction, X, 5)) = "trim(" Then
          X = X + 5
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & Trim(Arg1) & Chr(34)
      
      'Tab
      ElseIf LCase(Mid(vsFunction, X, 4)) = "tab(" Then
         X = X + 4
         Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
         X = lSetX: Arg1 = Solve(Arg1)
         If IsNumeric(Arg1) = True Then
            iCicle = Val(Arg1)
            For iCounter = 1 To iCicle
               sTempCad = sTempCad & vbTab
            Next iCounter
            TempFunction = TempFunction & Chr(34) & Trim(sTempCad) & Chr(34)
         Else
            TempFunction = TempFunction & Chr(34) & " " & Chr(34)
         End If
      'CrLf
      ElseIf LCase(Mid(vsFunction, X, 5)) = "crlf(" Then
         X = X + 5
         Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
         X = lSetX: Arg1 = Solve(Arg1)
         If IsNumeric(Arg1) = True Then
            iCicle = Val(Arg1)
            For iCounter = 1 To iCicle
               sTempCad = sTempCad & vbCrLf
            Next iCounter
            TempFunction = TempFunction & Chr(34) & Trim(sTempCad) & Chr(34)
         Else
            TempFunction = TempFunction & Chr(34) & " " & Chr(34)
         End If
      
      'SkipLastChar
      ElseIf LCase(Mid(vsFunction, X, 13)) = "skiplastchar(" Then
          X = X + 13
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & SkipLastChar(Arg1) & Chr(34)
      
      'Lower Case
      ElseIf LCase(Mid(vsFunction, X, 6)) = "lcase(" Then
          X = X + 6
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & LCase(Arg1) & Chr(34)
      'Upper Case
      ElseIf LCase(Mid(vsFunction, X, 6)) = "ucase(" Then
          X = X + 6
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & UCase(Arg1) & Chr(34)
      'Length
      ElseIf LCase(Mid(vsFunction, X, 4)) = "len(" Then
          X = X + 4
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & Len(Arg1) & Chr(34)
      'Is a number
      ElseIf LCase(Mid(vsFunction, X, 10)) = "isnumeric(" Then
          X = X + 10
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & IsNumeric(Arg1) & Chr(34)
      'Is an operator
      ElseIf LCase(Mid(vsFunction, X, 5)) = "isop(" Then
          X = X + 5
          Arg1 = GetValue3(X, ")", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          TempFunction = TempFunction & Chr(34) & IsOperator(Arg1) & Chr(34)
      'InStr
      ElseIf LCase(Mid(vsFunction, X, 6)) = "instr(" Then
          X = X + 6
          Arg1 = GetValue3(X, ",", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          Arg2 = GetValue3(X, ",", vsFunction) 'Argument 2
          X = lSetX: Arg2 = Solve(Arg2)
          Arg3 = GetValue3(X, ")", vsFunction) 'Argument 3
          X = lSetX: Arg3 = Solve(Arg3)
          TempFunction = TempFunction & Chr(34) & InStr(Arg1, Arg2, Arg3) & Chr(34)
      'Right
      ElseIf LCase(Mid(vsFunction, X, 6)) = "right(" Then
          X = X + 6
          Arg1 = GetValue3(X, ",", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          Arg2 = GetValue3(X, ")", vsFunction) 'Argument 2
          X = lSetX: Arg2 = Solve(Arg2)
          TempFunction = TempFunction & Chr(34) & Right(Arg1, Arg2) & Chr(34)
          
      'Left
      ElseIf LCase(Mid(vsFunction, X, 5)) = "left(" Then
          X = X + 5
          Arg1 = GetValue3(X, ",", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          Arg2 = GetValue3(X, ")", vsFunction) 'Argument 2
          X = lSetX: Arg2 = Solve(Arg2)
          TempFunction = TempFunction & Chr(34) & Left(Arg1, Arg2) & Chr(34)
      'Mid
      ElseIf LCase(Mid(vsFunction, X, 4)) = "mid(" Then
          X = X + 4
          Arg1 = GetValue3(X, ",", vsFunction) 'Argument 1
          X = lSetX: Arg1 = Solve(Arg1)
          Arg2 = GetValue3(X, ",", vsFunction) 'Argument 2
          X = lSetX: Arg2 = Solve(Arg2)
          Arg3 = GetValue3(X, ")", vsFunction) 'Argument 3
          X = lSetX: Arg3 = Solve(Arg3)
          TempFunction = TempFunction & Chr(34) & Mid(Arg1, Arg2, Arg3) & Chr(34)
      Else
          TempFunction = TempFunction & sChar
      End If
   Next X
   
   GetFunctions = TempFunction

End Function


Private Function RepeatChar(ByVal sCharRepeat As String, ByVal iCicles As Integer) As String
Dim ic As Integer, sTempCad As String
   For ic = 1 To iCicles
      sTempCad = sTempCad & sCharRepeat
   Next ic
   RepeatChar = sTempCad
End Function


Private Function SkipLastChar(ByVal vsCharSkip As String) As String
Dim X As Long
   'On Error Resume Next
   For X = Len(m_sCode) To 1 Step -1
      If X > Len(vsCharSkip) Then
         If LCase(Mid(m_sCode, X - Len(vsCharSkip), Len(vsCharSkip))) = LCase(vsCharSkip) Then
            m_sCode = Mid(m_sCode, 1, X - Len(vsCharSkip) - 1) & Mid(m_sCode, X, Len(m_sCode))
            SkipLastChar = " ": Exit Function
         End If
      End If
   Next X
   SkipLastChar = " ": Exit Function
End Function




'Evaluate ecuations
Private Function Equation(ByVal vsFunction As String) As Variant
'This sub basically looks for parentheses, and solves what's in them
Dim Paren1 As Integer, Paren2 As Integer, sChar As String, X
Do
    DoEvents
    For X = 1 To Len(vsFunction)
        sChar = Mid(vsFunction, X, 1)
        Select Case sChar
            Case Chr(34) 'Character 34 is the "
                X = InStr(X + 1, vsFunction, Chr(34))
            Case "("
                Paren1 = X
            Case ")"
                Paren2 = X
                Exit For
        End Select
    Next X
    If Paren1 = 0 Then
        Exit Do
    Else
        vsFunction = Mid(vsFunction, 1, Paren1 - 1) & " " & Chr(34) & SolveFunction(Mid(vsFunction, Paren1 + 1, Paren2 - (Paren1 + 1))) & Chr(34) & " " & Mid(vsFunction, Paren2 + 1)
        Paren1 = 0
        Paren2 = 0
    End If
Loop
Equation = vsFunction
End Function




Private Function SolveFunction(vsFunction As String) As Variant
'This sub solves equations like 5 + (num1 * 3)
Dim Quote As Integer, sChar As String, variable As String
Dim Num2 As Variant, SolveOp As String, Num1, X
vsFunction = Trim(vsFunction)
For X = 1 To Len(vsFunction)
    sChar = Mid(vsFunction, X, 1)
    If sChar = Chr(34) Then                             'thats the "
        Quote = InStr(X + 1, vsFunction, Chr(34))
        Num2 = Mid(vsFunction, X + 1, Quote - (X + 1))
        X = Quote
        If SolveOp <> "" Then
            SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
            SolveOp = ""
        Else
            SolveFunction = Num2
        End If
    ElseIf IsOperator(sChar) = True Then
        If Num1 <> 0 Then
            Num2 = GetVar(Trim(Mid(vsFunction, Num1, X - (Num1 + 1))))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Num1 = 0
        End If
        SolveOp = sChar
    Else
        If Asc(sChar) <> 32 And Num1 = 0 Then Num1 = X
        If X >= Len(vsFunction) Then
            Num2 = GetVar(Trim(Mid(vsFunction, Num1, X)))
            If SolveOp <> "" Then
                SolveFunction = DoOperation(SolveFunction, SolveOp, Num2)
                SolveOp = ""
            Else
                SolveFunction = Num2
            End If
            Exit For
        End If
    End If
    GoTo NextX
NextX:
Next X
End Function



Private Function DoOperation(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Variant
'Solves an equation
Select Case Operator
    Case "+"
        DoOperation = Val(LeftVal) + Val(RightVal)
    Case "-"
        DoOperation = Val(LeftVal) - Val(RightVal)
    Case "/"
        DoOperation = Val(LeftVal) / Val(RightVal)
    Case "\"
        DoOperation = Val(LeftVal) \ Val(RightVal)
    Case "^"
        DoOperation = Val(LeftVal) ^ Val(RightVal)
    Case "*"
        DoOperation = Val(LeftVal) * Val(RightVal)
    Case "&"
        DoOperation = LeftVal & RightVal
    Case Else
        FlagError "Invalid Operator, '" & Operator & "'", True
End Select
End Function



Private Function DoOperation2(ByVal LeftVal As Variant, ByVal Operator As Variant, ByVal RightVal As Variant) As Boolean
'Determines if an expression is True or False
Select Case Operator
    Case ">"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) > Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal > RightVal Then DoOperation2 = True
        End If
    Case "<"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) < Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal < RightVal Then DoOperation2 = True
        End If
    Case "="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) = Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal = RightVal Then DoOperation2 = True
        End If
    Case "<>"
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <> Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <> RightVal Then DoOperation2 = True
        End If
    Case ">="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) >= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal >= RightVal Then DoOperation2 = True
        End If
    Case "<="
        If IsNumeric(LeftVal) Or IsNumeric(RightVal) Then
            If Val(LeftVal) <= Val(RightVal) Then DoOperation2 = True
        Else
            If LeftVal <= RightVal Then DoOperation2 = True
        End If
    Case Else
        FlagError "Invalid Operator, '" & Operator & "'", True
End Select
End Function



Private Function GetValue2(Start As Long, Parameter As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
Dim FinalCode As String, WhereItIs As Long
'Determine the location of the parameter
WhereItIs = CodeInStr(Start, m_sScript, Parameter)
If WhereItIs = 0 Then
    'Not there? Flag an error.
    FlagError "Expected '" & Parameter & "', but found end of Script.", True
Else
    'Get the code between the starting location and the paramter
    FinalCode = Mid(m_sScript, Start, WhereItIs - Start)
    GetValue2 = Trim(FinalCode)
    'Set the location
    lSetCount = Start + Len(FinalCode) + 1
End If
End Function



Private Function GetValue3(ByVal vlStart As Long, vsParameter As String, vsInWhat As String) As Variant
'This will search for the paramter, starting at the
'specified location, skipping things in quotes, but
'DOESN'T evaluate it.
Dim sEndCode As Variant, X As Long, lQuote As Long, lParen As Long
   If vsInWhat = "" Then vsInWhat = m_sScript
   If CodeInStr(vlStart, vsInWhat, vsParameter) = 0 Then
       'Not there?  Flag an error.
       MsgBox "Expected '" & vsParameter, vbCritical, App.Title
   End If
   For X = vlStart To Len(vsInWhat)
       If Mid(vsInWhat, X, Len(vsParameter)) = vsParameter Then
           sEndCode = Mid(vsInWhat, vlStart, X - vlStart)
           GetValue3 = Trim(sEndCode)
           lSetX = vlStart + Len(sEndCode) + 1
           Exit Function
       ElseIf Mid(vsInWhat, X, 1) = Chr(34) Then
           lQuote = InStr(X, vsInWhat, Chr(34))
           If lQuote <> 0 Then
               X = lQuote
           End If
       ElseIf Mid(vsInWhat, X, 1) = "(" Then
           lParen = InStr(X, vsInWhat, ")")
           If lParen <> 0 Then
               X = lParen
           End If
       End If
   Next X
End Function



Private Function IsOperator(NumVal As Variant) As Boolean
'Determines if a character is an operator
Select Case NumVal
Case "+", "-", "*", "\", "/", "&", ">", "<", "="
    IsOperator = True
End Select
End Function



Private Function GetVar(TheVar As Variant) As Variant
'Gets a variables value
Dim X As Long
   If IsNumeric(TheVar) Then GetVar = TheVar: Exit Function
   
   If FindVar(TheVar) = False Then
      Select Case LCase(TheVar)
         
         
         'Properties of conection and server
         Case "db.connectionstring"
            GetVar = m_sConnectionString: Exit Function
         Case "db.svr.name"
            GetVar = m_sServer: Exit Function
         Case "db.svr.type"
            GetVar = m_sServer: Exit Function
            
         
         'Properties of database
         Case "db.name"
            GetVar = m_sDataBase: Exit Function
         Case "db.type"
            GetVar = m_sDataBase: Exit Function
         Case "db.createdate"
            GetVar = m_sDataBase: Exit Function
            
         
         'Properties of table
         Case "db.tbl.name"
            GetVar = m_oTable.Name: Exit Function
         Case "db.tbl.owner"
            GetVar = m_oTable.Owner: Exit Function
         Case "db.tbl.type"
            GetVar = m_oTable.TypeTable: Exit Function
         Case "db.tbl.createdate"
            GetVar = m_oTable.CreateDate: Exit Function
         Case "db.tbl.fieldscount"
            GetVar = m_oTable.Fields.Count: Exit Function
            
         
         
         '//Properties of selected fields in table
         Case "tbl.fld.adoconstantvalue"
            GetVar = m_oTable.Fields.Item(iPosField).ADOConstantValue: Exit Function
         Case "tbl.fld.definedsize"
            GetVar = m_oTable.Fields.Item(iPosField).DefinedSize: Exit Function
         Case "tbl.fld.isidentity"
            GetVar = m_oTable.Fields.Item(iPosField).IsIdentity: Exit Function
         Case "tbl.fld.isnull"
            GetVar = m_oTable.Fields.Item(iPosField).IsNull: Exit Function
         Case "tbl.fld.ispk"
            GetVar = m_oTable.Fields.Item(iPosField).IsPK: Exit Function
         Case "tbl.fld.name"
            GetVar = m_oTable.Fields.Item(iPosField).Name: Exit Function
         Case "tbl.fld.numericscale"
            GetVar = m_oTable.Fields.Item(iPosField).NumericScale: Exit Function
         Case "tbl.fld.precision"
            GetVar = m_oTable.Fields.Item(iPosField).Precision: Exit Function
         Case "tbl.fld.type"
            GetVar = m_oTable.Fields.Item(iPosField).StringType: Exit Function
         Case "tbl.fld.stringtype"
            GetVar = TranslateTypeSQL(m_oTable): Exit Function
         Case "tbl.fld.posdatabase"
            GetVar = (iPosField - 1): Exit Function
            
            
            
         
         
         '// Visual Basic variables based in SQL Datatypes
         Case "vb.fld.prefix"
            GetVar = PrefixSQLToVB(m_oTable.Fields.Item(iPosField).StringType): Exit Function
         Case "vb.fld.stringtype"
            GetVar = ConvertTypeSQLToVB(m_oTable.Fields.Item(iPosField).StringType): Exit Function
            
            
         'Return values of string
         Case "str.crlf"
            GetVar = vbCrLf: Exit Function
         Case "str.tab"
            GetVar = Space(m_iTabWidth): Exit Function
         Case "str.comma"
            GetVar = """": Exit Function
         
         'Return information of author an legal
         Case "cg.legal"
            GetVar = m_sLegal: Exit Function
         Case "cg.author"
            GetVar = m_sAuthor: Exit Function
         
         
         'Return values of system
         Case "sys.time"
            GetVar = Format(Time, "hh:mm:ss"): Exit Function
         Case "sys.date"
            GetVar = Format(Date, "mm/dd/yyyy"): Exit Function
            
         
         Case Else
            FlagError "Variable not found, '" & TheVar & "'", True
      End Select
   End If
   
   For X = 1 To VarNames.Count
       If VarNames(X) = TheVar Then
          GetVar = VarData(X)
           Exit Function
       End If
   Next X
   
End Function


Private Sub AddVariable(VarName As Variant, VariableData As Variant)
   If IsNumeric(VarName) Then FlagError "Numbers are constants, " & VarName & " cannot be a variable.", True
   If FindVar(VarName) = True Then
       FlagError "Variable already exits, '" & VarName & "'.", True
   Else
       VarNames.Add VarName
       VarData.Add VariableData
   End If
End Sub


Private Function FindVar(TheVar As Variant) As Boolean
Dim X As Long
   'Determines if a variable exists
   For X = 1 To VarNames.Count
      If VarNames(X) = TheVar Then
         FindVar = True
         Exit Function
      End If
   Next X
   FindVar = False
   
End Function


Private Sub SetVar(TheVar As Variant, NewVal As Variant)
Dim X As Long
   'Sets the value of a variable
   'If a number...
   If IsNumeric(TheVar) Then FlagError "Numbers are constants, cannnot change the value of " & TheVar & ".", True
   'If variable doesn't exist...
   If FindVar(TheVar) = False Then FlagError "Variable not found, '" & TheVar & "'", True
   
   If TheVar = "dpp.ip" Or TheVar = "dpp.host" Or TheVar = "dpp.systemfolder" Or TheVar = "dpp.crlf" Or TheVar = "dpp.path" Or TheVar = "True" Or TheVar = "False" Then
      FlagError "Warning! You cannot edit D++ preset variables!  You tried to edit " & TheVar & ".", True
   End If
   
   For X = VarNames.Count To 1 Step -1
      If VarNames(X) = TheVar Then
         VarNames.Remove X
         VarData.Remove X
         VarNames.Add TheVar
         VarData.Add NewVal
         Exit Sub
      End If
   Next X
   
End Sub


Private Function Eval(ByVal sFunction As String) As Boolean
'This parses a string into a left value, operator, and right value
Dim LeftVal As String, RightVal As String, Operator
Dim sChar, OpFound As Boolean
Dim X As Long
   
   OpFound = False

   For X = 1 To Len(sFunction)
       sChar = Mid(sFunction, X, 1)
       If sChar = ">" Or sChar = "<" Or sChar = "=" Then
           Operator = Operator & sChar
           OpFound = True
       Else
           If OpFound = True Then
               RightVal = RightVal & sChar
           Else
               LeftVal = LeftVal & sChar
           End If
       End If
   Next X
   
   LeftVal = Solve(LeftVal)
   RightVal = Solve(RightVal)
   
   Eval = DoOperation2(LeftVal, Operator, RightVal)
   
End Function



Private Function StringExist(StartPos As Long, SourceText As String, ToFind As String) As Boolean
    'This determines if a string exists in another string
    If CodeInStr(StartPos, LCase(SourceText), LCase(ToFind)) = 0 Then
        StringExist = False
    Else
        StringExist = True
    End If
End Function



Private Sub FlagError(ErrorText As String, Optional EndProgram As Boolean = True)
   MsgBox ErrorText, vbCritical, App.Title
   If EndProgram = True Then
       bEndTrad = True
   Else
       Exit Sub
   End If
End Sub



'Traslacion de campos SQL Server
Public Function TranslateTypeSQL(ByVal voTblViewSP As Object) As String
   
   With voTblViewSP.Fields.Item(iPosField)
      Select Case UCase(.StringType)
         Case UCase("bit"), UCase("datetime"), UCase("float"), UCase("image"), _
              UCase("int"), UCase("money"), UCase("ntext"), UCase("real"), _
              UCase("smalldatetime"), UCase("smallint"), UCase("smallmoney"), _
              UCase("text"), UCase("timestamp"), UCase("tinyint"), _
              UCase("uniqueidentifier")
              TranslateTypeSQL = .StringType
         
         Case UCase("binary"), UCase("char"), UCase("nchar"), UCase("nvarchar"), _
              UCase("varbinary"), UCase("varchar")
              TranslateTypeSQL = .StringType & "(" & .DefinedSize & ")"
         
         Case UCase("decimal"), UCase("numeric")
            If .NumericScale = 0 Then
               TranslateTypeSQL = .StringType & "(" & .DefinedSize & ")"
            Else
              TranslateTypeSQL = .StringType & "(" & .NumericScale & "," & .Precision & ")"
            End If
      End Select
   
   End With
   
End Function




'Funtion to convert prefixed SQL to Visual Basic
Private Function ConvertTypeSQLToVB(ByVal vsTypeField As String) As String
   Select Case LCase(vsTypeField)
      Case "binary", "image", "varbinary", "text", "ntext"
         ConvertTypeSQLToVB = "Variant"
      Case "bit"
         ConvertTypeSQLToVB = "Byte"
      Case "char", "nchar", "varchar", "nvarchar"
         ConvertTypeSQLToVB = "String"
      Case "datetime", "smalldatetime", "timestamp"
         ConvertTypeSQLToVB = "Date"
      Case "float", "real", "numeric", "uniqueidentifier"
         ConvertTypeSQLToVB = "Double"
      Case "decimal"
         ConvertTypeSQLToVB = "Single"
      Case "tinyint", "smallint"
         ConvertTypeSQLToVB = "Integer"
      Case "int"
         ConvertTypeSQLToVB = "Long"
      Case "money", "smallmoney"
         ConvertTypeSQLToVB = "Currency"
   End Select
End Function


'Funtion to generate prefix visual basic datatypes based on SQL datatypes
Private Function PrefixSQLToVB(ByVal vsTypeField As String) As String
   Select Case LCase(Trim(vsTypeField))
      Case "binary", "image", "varbinary", "text", "ntext"
         PrefixSQLToVB = "var"
      Case "bit"
         PrefixSQLToVB = "byt"
      Case "char", "nchar", "varchar", "nvarchar"
         PrefixSQLToVB = "str"
      Case "datetime", "smalldatetime", "timestamp"
         PrefixSQLToVB = "dat"
      Case "float", "real", "numeric", "uniqueidentifier"
         PrefixSQLToVB = "dbl"
      Case "decimal"
         PrefixSQLToVB = "sgl"
      Case "tinyint", "smallint"
         PrefixSQLToVB = "int"
      Case "int"
         PrefixSQLToVB = "lng"
      Case "money", "smallmoney"
         PrefixSQLToVB = "cur"
   End Select
End Function




'Procedimiento para el tratamiento de nombres de campos con espacios intermedios
Private Function Bracket(ByVal strValue As String) As String
   If InStr(strValue, " ") > 0 Then
      Bracket = "[" & strValue & "]"
   Else
      Bracket = strValue
   End If
End Function

'Procedimiento para el tratamiento de nombres de campos con espacios intermedios
Private Function CompressName(ByVal strValue As String) As String
   If InStr(strValue, " ") > 0 Then
      CompressName = Replace(strValue, " ", "", 1, 1, vbTextCompare)
   Else
      CompressName = strValue
   End If
End Function





Public Sub CancelTrad()

   Set VarNames = Nothing
   Set VarData = Nothing
   Set m_oTable = Nothing
   iPosField = 0
   bEndTrad = True
   m_sScript = ""
   m_sCode = ""
   m_sServer = ""
   m_sDataBase = ""
   m_sOwner = ""
   m_sConnectionString = ""
   m_sAuthor = ""
   m_sLegal = ""
   
End Sub


Private Sub Class_Initialize()
   
   Set VarNames = New Collection
   Set VarData = New Collection
   Set m_oTable = New CTable
   m_eTypeDB = [SQL Server]
   iPosField = 1
   bEndTrad = False
   m_sScript = ""
   m_sCode = ""
   m_sServer = ""
   m_sDataBase = ""
   m_sOwner = ""
   m_sConnectionString = ""
   m_sAuthor = ""
   m_sLegal = ""
   m_iTabWidth = 3
   
End Sub


Private Sub Class_Terminate()
   Set VarNames = Nothing
   Set VarData = Nothing
   Set m_oTable = Nothing
   iPosField = 0
   bEndTrad = True
   m_sScript = ""
   m_sCode = ""
   m_sServer = ""
   m_sDataBase = ""
   m_sOwner = ""
   m_sConnectionString = ""
   m_sAuthor = ""
   m_sLegal = ""
End Sub


