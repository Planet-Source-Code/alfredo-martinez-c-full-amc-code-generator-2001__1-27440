VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_CFields As Collection


Public Sub Add(Optional ByVal voField As CField)
   If Not voField Is Nothing Then
      m_CFields.Add voField
   Else
      Dim oNewField As New CField
      m_CFields.Add oNewField
      Set oNewField = Nothing
   End If
End Sub

Public Sub Remove(Index)
   m_CFields.Remove (Index)
End Sub


Public Sub Clear()
   Set m_CFields = Nothing
   Set m_CFields = New Collection
End Sub

Public Function Item(Index) As CField
   Set Item = m_CFields(Index)
End Function

Public Property Get Count() As Integer
   Count = m_CFields.Count
End Property

Private Sub Class_Initialize()
   Set m_CFields = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_CFields = Nothing
End Sub
