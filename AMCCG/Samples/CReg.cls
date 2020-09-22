VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Use with SQL
'Based in the table Region'The properties is based in the class objRegion
'The template requiered function to validate errors
'This template requiered one connection based in ADO 2.5, 2.6
'Date and Time: 06/16/2001 - 04:21:28
'Author: Alfredo Martínez C.
'Legal information: Codigo solo de ejemplo por: AMC



'// Variables of properties
Private m_lng_RegionID As Long
Private m_str_RegionDescription As String
Public adoConn As ADODB.Connection


'// Methods the read/write on properties

'Property of read based in field RegionID
Public Property Get RegionID() As Long
   RegionID = m_lng_RegionID
End Property
'Property of write based in field RegionID
Public Property Let RegionID(ByVal vlngRegionID As Long)
   m_lng_RegionID = vlngRegionID
End Property

'Property of read based in field RegionDescription
Public Property Get RegionDescription() As String
   RegionDescription = m_str_RegionDescription
End Property
'Property of write based in field RegionDescription
Public Property Let RegionDescription(ByVal vstrRegionDescription As String)
   m_str_RegionDescription = vstrRegionDescription
End Property



'// Procedure for deleted records
Public Sub Delete()
Dim sSQL As String
Dim CMD As ADODB.Command
Dim nRowsAffected As Long

On Error GoTo DeleteErrorRegion

   sSQL = "DELETE FROM " & Region
   sSQL = sSQL & " WHERE "
   sSQL = sSQL & "RegionID" = m_lng_RegionID & "  "
 
   Set CMD = New ADODB.Command
   With CMD
      .ActiveConnection = adoConn
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute nRowsAffected
   End With

DeleteExitRegion:
   Exit Sub
DeleteErrorRegion:
   Call GenError
   Resume DeleteExitRegion
End Sub


'Procedimiento para insertar un registro
Public Sub Insert()
Dim sSQL As String
Dim CMD As ADODB.Command
Dim nRowsAffected As Long
On Error GoTo InsertErrorRegion

   sSQL = "INSERT INTO Region"
   sSQL = sSQL & " ("
   sSQL = sSQL & "RegionID , "
   sSQL = sSQL & "RegionDescription  "
   sSQL = sSQL & ") "
   sSQL = sSQL & "VALUES "
   sSQL = sSQL & "("
   sSQL = sSQL & m_lng_RegionID & " , "
   sSQL = sSQL & m_str_RegionDescription & "  "
   sSQL = sSQL & ") "
 
   Set CMD = New ADODB.Command

   With CMD
      .ActiveConnection = adoConn
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute nRowsAffected
   End With


InsertExitRegion:
   Exit Sub

InsertErrorRegion:
   Call GenError
   Resume InsertExitRegion
End Sub


'Procedure for Select Records
Public Sub GetData()
Dim sSQL As String
Dim RS As ADODB.Recordset
On Error GoTo GetDataErrorRegion

   Set RS = New ADODB.Recordset

   RS.CursorLocation = adUseClient

   sSQL = "SELECT "
   sSQL = sSQL & "RegionID, "
   sSQL = sSQL & "RegionDescription "
    sSQL = sSQL & "FROM Region WHERE "
   sSQL = sSQL & "RegionID = " & m_lng_RegionID & "  "
 
   Set RS.ActiveConnection = adoConn
   RS.Open sSQL, , adOpenForwardOnly, adLockReadOnly

   If RS.RecordCount > 0 Then
      If Not IsNull(RS("RegionID")) Then m_lng_RegionID = RS(0)
      If Not IsNull(RS("RegionDescription")) Then m_str_RegionDescription = RS(1)
   End If

GetDataExitRegion:
   If Not RS Is Nothing Then
      If RS.State = adStateOpen Then
         RS.Close
      End If
      Set RS = Nothing
   End If
   Exit Sub

GetDataErrorRegion:
   Call GenError
   Resume GetDataExitRegion
End Sub


'Procedure for Update Record
Public Sub Update()
Dim sSQL As String
Dim CMD As ADODB.Command
Dim nRowsAffected As Long
On Error GoTo UpdateErrorRegion

   sSQL = "UPDATE Region SET "
   sSQL = sSQL & "RegionDescription = " & m_str_RegionDescription & " "
    sSQL = sSQL & " WHERE "
   sSQL = sSQL & "RegionID" = m_lng_RegionID & "  "
 
   Set CMD = New ADODB.Command
   With CMD
      .ActiveConnection = adoConn
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute nRowsAffected
   End With

UpdateExitRegion:
   Exit Sub

UpdateErrorRegion:
   Call GenError
   Resume UpdateExitRegion
End Sub


'Procedure for display message error
Private Sub GenError()
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
End Sub








