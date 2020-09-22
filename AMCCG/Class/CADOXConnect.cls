VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CADOXConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//-----------------------------------------------------------------------
'//   Name: CADOXConnect
'//   Use:  Administrator of ADODB connections used in AMC-CodeAssist
'//   Created: 10/March/2001
'//   Autor:   Alfredo Martínez C.
'//-----------------------------------------------------------------------


'Private vars member of propertys
Private mcConnections As Collection
Private iCountCnn As Integer

'Public method of add connection in then collection
Public Sub Add(ByVal vNewADODBConnection As ADODB.Connection, ByVal sKey As String)
   With mcConnections
      .Add vNewADODBConnection, sKey
   End With
End Sub

'Public method of access properties ADODB connections
Public Property Get Item(Index) As ADODB.Connection
   Set Item = mcConnections(Index)
End Property

'Public method of remove ADODB connection with collection
Public Sub Remove(Index)
   Call mcConnections.Remove(Index)
End Sub

'Public method count numer of ADODB Connections in the collection
Public Property Get Count() As Integer
   Count = mcConnections.Count
End Property

'Public Method for destroy ADODB connections in the object collection.
Public Sub Clear()
   Set mcConnections = Nothing
   Set mcConnections = New Collection
End Sub

'Verify connection is included in then collection
Public Function VerifyKey(ByVal sKey As String) As Boolean
   With mcConnections
      For iCountCnn = 1 To .Count
         If Trim(.Item(iCountCnn).ConnectionString) = Trim(sKey) Then
            VerifyKey = True: Exit Function
         End If
      Next iCountCnn
      VerifyKey = False
   End With
End Function

'Initialize object collecion
Private Sub Class_Initialize()
   Set mcConnections = New Collection
End Sub

'Destroy object collecion
Private Sub Class_Terminate()
   Set mcConnections = Nothing
End Sub
