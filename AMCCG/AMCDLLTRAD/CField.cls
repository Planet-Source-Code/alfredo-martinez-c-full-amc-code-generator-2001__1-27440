VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'// Variables of Properties
Private m_sName As String
Private m_sStringType As String
Private m_iADOConstantValue As Integer
Private m_lDefinedSize As Long
Private m_iPrecision As Integer
Private m_iNumericScale As Integer
Private m_bIsPK As Boolean
Private m_bIsNull As Boolean
Private m_bIsIdentity As Boolean
Private m_bIsParamProc As Boolean
Private m_bTypeParam As enmTypeParam

'// Enumetarions
Enum enmTypeParam
   [NoParam] = 0
   [ParamInput] = 1
   [ParamOutput] = 2
   [ParamInputOutput] = 3
End Enum


'// Properties Class
Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Let Name(ByVal vNewValue As String)
   m_sName = vNewValue
End Property

Public Property Get StringType() As String
   StringType = m_sStringType
End Property
Public Property Let StringType(ByVal vNewValue As String)
   m_sStringType = vNewValue
End Property

Public Property Get ADOConstantValue() As Integer
   ADOConstantValue = m_iADOConstantValue
End Property
Public Property Let ADOConstantValue(ByVal vNewValue As Integer)
   m_iADOConstantValue = vNewValue
End Property

Public Property Get DefinedSize() As Long
   DefinedSize = m_lDefinedSize
End Property
Public Property Let DefinedSize(ByVal vNewValue As Long)
   m_lDefinedSize = vNewValue
End Property

Public Property Get Precision() As Integer
   Precision = m_iPrecision
End Property
Public Property Let Precision(ByVal vNewValue As Integer)
   m_iPrecision = vNewValue
End Property

Public Property Get NumericScale() As Integer
   NumericScale = m_iNumericScale
End Property
Public Property Let NumericScale(ByVal vNewValue As Integer)
   m_iNumericScale = vNewValue
End Property

Public Property Get IsPK() As Boolean
   IsPK = m_bIsPK
End Property
Public Property Let IsPK(ByVal vNewValue As Boolean)
   m_bIsPK = vNewValue
   If vNewValue = True Then
      m_bIsNull = False
   Else
      m_bIsNull = True
   End If
End Property

Public Property Get IsNull() As Boolean
   IsNull = m_bIsNull
End Property
Public Property Let IsNull(ByVal vNewValue As Boolean)
   m_bIsNull = vNewValue
   If m_bIsPK = True Then m_bIsNull = False
End Property

Public Property Get IsIdentity() As Boolean
   IsIdentity = m_bIsIdentity
End Property
Public Property Let IsIdentity(ByVal vNewValue As Boolean)
   m_bIsIdentity = vNewValue
End Property

Public Property Get IsParamProc() As Boolean
   IsParamProc = m_bIsParamProc
End Property
Public Property Let IsParamProc(ByVal vNewValue As Boolean)
   m_bIsParamProc = vNewValue
End Property

Public Property Get TypeParam() As enmTypeParam
   TypeParam = m_bTypeParam
End Property
Public Property Let TypeParam(ByVal vNewValue As enmTypeParam)
   m_bTypeParam = vNewValue
End Property


Private Sub Class_Initialize()
   m_bTypeParam = [NoParam]
End Sub
