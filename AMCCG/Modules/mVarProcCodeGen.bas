Attribute VB_Name = "mVarProcCodeGen"
Option Explicit

Public oMPG As New CMPG
Public oDBs As New CADOXConnect
Public oCnnTmp As New CCnnTmp

Public sysDB As New ADODB.Connection
Public sSQLcmd As String



Public Function OpenFileTemplate(ByVal sFile As String) As String
Dim vResult, sLine As String, sText As String
   vResult = Dir(sFile)
   If vResult <> "" Then
      Open sFile For Input As #1
         Do While Not EOF(1)
            Line Input #1, sLine
            sText = sText & sLine & vbCrLf
         Loop
      Close #1
   End If
   OpenFileTemplate = sText
End Function



Sub Main()
   
   Load frmMain
      If ConnectSysDB = False Then
         Unload frmMain
         End
      End If
   Call frmMain.LoadLanguajes
   Call frmMain.LoadLibraries
   frmMain.Show
End Sub


Private Function ConnectSysDB() As Boolean
Dim vResult, bExistDB As Boolean

On Error GoTo ErrorConnectSysDB

   vResult = Dir(App.Path & "\CFGOG.mdb")
   
   If Trim(vResult) = "" Then
      MsgBox "The configuration database not encountered", vbInformation, App.Title
      If MsgBox("You locate database? ", vbQuestion + vbYesNoCancel) = vbYes Then
         With frmMain.cdlgGen
            .Filter = "Microsoft Access Database *.mdb|*.mdb"
            .FileName = App.Path & "\CFGOG.mdb"
            .ShowOpen
            If Trim(.FileTitle) = "" Then
               ConnectSysDB = False
               Exit Function
            Else
               If LCase(Trim(.FileTitle)) <> "cfgog.mdb" Then
                  MsgBox "The selected database not is " & App.Path & " configuration database", vbExclamation, App.Title
                  MsgBox App.Title & " it's finished", vbCritical, App.Title
                  ConnectSysDB = False: Exit Function
               End If
            End If
         End With
      Else
         ConnectSysDB = False: Exit Function
      End If
   Else
      bExistDB = True
   End If
   
   With sysDB
      If bExistDB = True Then
         .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & App.Path & "\CFGOG.mdb"
      Else
         .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " & frmMain.cdlgGen.FileName
      End If
   End With
   
   ConnectSysDB = True
   
Exit Function
ErrorConnectSysDB:
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
   MsgBox App.Title & " it's finished", vbInformation, App.Title
   ConnectSysDB = False
End Function




