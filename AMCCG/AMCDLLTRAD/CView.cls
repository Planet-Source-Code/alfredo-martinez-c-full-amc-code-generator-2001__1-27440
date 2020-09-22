VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_sName As String
Private m_sTypeView As String
Private m_sOwner As String
Private m_sCreateDate As String

Private m_oFields As CFields



Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Let Name(ByVal vNewValue As String)
   m_sName = vNewValue
End Property

Public Property Get TypeView() As String
   TypeView = m_sTypeView
End Property
Public Property Let TypeView(ByVal vNewValue As String)
   m_sTypeView = vNewValue
End Property

Public Property Get Owner() As String
   Owner = m_sOwner
End Property
Public Property Let Owner(ByVal vNewValue As String)
   m_sOwner = vNewValue
End Property

Public Property Get CreateDate() As String
   CreateDate = m_sCreateDate
End Property
Public Property Let CreateDate(ByVal vNewValue As String)
   m_sCreateDate = vNewValue
End Property




Public Property Get Fields() As CFields
   Set Fields = m_oFields
End Property
Public Property Let Fields(ByVal vNewValue As CFields)
   Set m_oFields = vNewValue
End Property




Private Sub Class_Initialize()
   Set m_oFields = New CFields
End Sub
Private Sub Class_Terminate()
   Set m_oFields = Nothing
End Sub




