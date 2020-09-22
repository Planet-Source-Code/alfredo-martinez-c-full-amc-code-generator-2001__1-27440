VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCnnTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lCounter As Long




Enum enmDisplayFieldsMode
   [Table] = 0
   [View] = 1
End Enum


Private oPBar As CProgressBar


Public Sub DisplayConnection(ByVal vsKey As String, ByVal voTV As TreeView, ByVal sDatabaseName As String, ByVal eDBType As enmDBType)
Dim xNode As Node, lCountFields As Long, lCountVars As Long
Dim oCatalog As New ADOX.Catalog, lIndexNode As Long
   
   'Set xNode = voTV.Nodes.Add(, , vsKey, sDatabaseName, Image:=1, SelectedImage:=1)
   Set xNode = voTV.Nodes.Add("databases", tvwChild, vsKey, sDatabaseName, Image:=1, SelectedImage:=1)
   
   xNode.Tag = eDBType
   xNode.Expanded = True
   Set xNode = voTV.Nodes.Add(vsKey, tvwChild, "TABL" & vsKey, "Tables", Image:=2, SelectedImage:=2)
   xNode.Tag = "Tables"
   Set xNode = voTV.Nodes.Add(vsKey, tvwChild, "VIEW" & vsKey, "Views", Image:=3, SelectedImage:=3)
   xNode.Tag = "Views"
   If Not eDBType = [Microsoft Access] Then
      Set xNode = voTV.Nodes.Add(vsKey, tvwChild, "SPRO" & vsKey, "Store Procedures", Image:=4, SelectedImage:=4)
      xNode.Tag = "Procedures"
   End If

End Sub


Private Sub ConfigLVTblViewSP(ByVal voLV As ListView)
   With voLV
      .ListItems.Clear
      .ColumnHeaders.Clear
      .Checkboxes = True
      .ColumnHeaders.Add , , "Name", 3000
      .ColumnHeaders.Add , , "Owner", 700
      .ColumnHeaders.Add , , "Type", 1500
      .ColumnHeaders.Add , , "Date Created", 2000
      .ColumnHeaders.Add , , "Date Modified", 2000
   End With
End Sub

Private Sub ConfigLVDatabase(ByVal voLV As ListView)
   With voLV
      .Checkboxes = False
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Property Name", 2500
      .ColumnHeaders.Add , , "Value", 2850
      .ColumnHeaders.Add , , "Type", 600
      .ColumnHeaders.Add , , "Attributes", 900
   End With
End Sub

Private Function CfgTblViewFields(ByVal voLV As ListView) As String
   With voLV
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Name", 3000
      .ColumnHeaders.Add , , "Type", 1200
      .ColumnHeaders.Add , , "Size", 1200
      .ColumnHeaders.Add , , "Numeric Scale", 1300
      .ColumnHeaders.Add , , "Precision", 1000
   End With
End Function




Public Sub DisplayDBProperties(ByVal vsKey As String, ByVal voLV As ListView, ByVal voSB As StatusBar)
Dim oCatalog As New ADOX.Catalog, iCount As Integer
   Set oPBar = New CProgressBar
   oPBar.CreateProgress voSB, 5
   Call ConfigLVDatabase(voLV)
   With oDBs.Item(vsKey).Properties
      oPBar.Max = .Count + 1
      For iCount = 0 To .Count - 1
         oPBar.Value = iCount + 1
         voLV.ListItems.Add , , .Item(iCount).Name, , SmallIcon:=11
         voLV.ListItems(voLV.ListItems.Count).SubItems(1) = .Item(iCount).Value
         voLV.ListItems(voLV.ListItems.Count).SubItems(2) = .Item(iCount).Type
         voLV.ListItems(voLV.ListItems.Count).SubItems(3) = .Item(iCount).Attributes
      Next iCount
   End With
   oPBar.Value = 0
   Set oPBar = Nothing

End Sub



Private Function CfgSPParams(ByVal voLV As ListView) As String
   With voLV
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "Param Type", 1500
      .ColumnHeaders.Add , , "Name", 2000
      .ColumnHeaders.Add , , "ADO Type", 1200
      .ColumnHeaders.Add , , "Size", 800
      .ColumnHeaders.Add , , "Numeric Scale", 1300
      .ColumnHeaders.Add , , "Precision", 1000
   End With
End Function


Public Function DisplayParamsSP(ByVal vsKey As String, ByVal vsObjectBDName As String, ByVal voLV As ListView, ByVal voSB As StatusBar) As CFields
Dim iCount As Integer, oCmd As New ADODB.Command, oFields As New CFields
Dim RsParams As New ADODB.Recordset, nIndexParam As Integer, bErrExist As Boolean
On Error GoTo ErrorDisplayParamsSP
   
   Call CfgSPParams(voLV)
   
   With oCmd
      .CommandText = vsObjectBDName
      .CommandType = adCmdStoredProc
      .ActiveConnection = oDBs.Item(vsKey)
      
      With .Parameters
         For iCount = 0 To .Count - 1
            Select Case .Item(iCount).Direction
               Case adParamInput
                  oFields.Add
                  nIndexParam = nIndexParam + 1
                  oFields.Item(nIndexParam).Name = .Item(iCount).Name
                  oFields.Item(nIndexParam).StringType = ConvertValueADONameConstant(.Item(iCount).Type)
                  oFields.Item(nIndexParam).DefinedSize = .Item(iCount).Size
                  oFields.Item(nIndexParam).Precision = .Item(iCount).Precision
                  oFields.Item(nIndexParam).NumericScale = .Item(iCount).NumericScale
                  oFields.Item(nIndexParam).ADOConstantValue = .Item(iCount).Type
                  oFields.Item(nIndexParam).IsParamProc = True
                  oFields.Item(nIndexParam).TypeParam = ParamInput
               Case adParamOutput
                  oFields.Add
                  nIndexParam = nIndexParam + 1
                  oFields.Item(nIndexParam).Name = .Item(iCount).Name
                  oFields.Item(nIndexParam).StringType = ConvertValueADONameConstant(.Item(iCount).Type)
                  oFields.Item(nIndexParam).DefinedSize = .Item(iCount).Size
                  oFields.Item(nIndexParam).Precision = .Item(iCount).Precision
                  oFields.Item(nIndexParam).NumericScale = .Item(iCount).NumericScale
                  oFields.Item(nIndexParam).ADOConstantValue = .Item(iCount).Type
                  oFields.Item(nIndexParam).IsParamProc = True
                  oFields.Item(nIndexParam).TypeParam = ParamOutput
               Case adParamInputOutput
                  oFields.Add
                  nIndexParam = nIndexParam + 1
                  oFields.Item(nIndexParam).Name = .Item(iCount).Name
                  oFields.Item(nIndexParam).StringType = ConvertValueADONameConstant(.Item(iCount).Type)
                  oFields.Item(nIndexParam).DefinedSize = .Item(iCount).Size
                  oFields.Item(nIndexParam).Precision = .Item(iCount).Precision
                  oFields.Item(nIndexParam).NumericScale = .Item(iCount).NumericScale
                  oFields.Item(nIndexParam).ADOConstantValue = .Item(iCount).Type
                  oFields.Item(nIndexParam).IsParamProc = True
                  oFields.Item(nIndexParam).TypeParam = ParamInputOutput
            End Select
         Next iCount
         
      End With
   End With
   iCount = 0
   nIndexParam = oFields.Count
   With RsParams
      .Source = oCmd.CommandText
      .ActiveConnection = oDBs.Item(vsKey)
      With .Fields
         For iCount = 0 To .Count - 1
            If bErrExist = True Then Exit For
            oFields.Add
            nIndexParam = nIndexParam + 1
            oFields.Item(nIndexParam).TypeParam = ParamOutput
            oFields.Item(nIndexParam).IsParamProc = True
            oFields.Item(nIndexParam).Name = .Item(iCount).Name
            oFields.Item(nIndexParam).StringType = ConvertValueADONameConstant(.Item(iCount).Type)
            oFields.Item(nIndexParam).DefinedSize = .Item(iCount).DefinedSize
            oFields.Item(nIndexParam).Precision = .Item(iCount).Precision
            oFields.Item(nIndexParam).NumericScale = .Item(iCount).NumericScale
            oFields.Item(nIndexParam).ADOConstantValue = .Item(iCount).Type
            oFields.Item(nIndexParam).TypeParam = ParamOutput
         Next iCount
      End With
   End With
         
         
   With oFields
      For iCount = 1 To oFields.Count
         Select Case .Item(iCount).TypeParam
            Case [ParamInput]
               voLV.ListItems.Add , , "Param Input", , SmallIcon:=15
            Case [ParamInputOutput]
               voLV.ListItems.Add , , "Param Input/Output"
            Case [ParamOutput]
               voLV.ListItems.Add , , "Param Output", , SmallIcon:=5
         End Select
            voLV.ListItems(voLV.ListItems.Count).SubItems(1) = .Item(iCount).Name
            voLV.ListItems(voLV.ListItems.Count).SubItems(2) = .Item(iCount).StringType
            voLV.ListItems(voLV.ListItems.Count).SubItems(3) = .Item(iCount).DefinedSize
            voLV.ListItems(voLV.ListItems.Count).SubItems(4) = .Item(iCount).NumericScale
            voLV.ListItems(voLV.ListItems.Count).SubItems(5) = .Item(iCount).Precision
      Next iCount
   End With
   
   
      

   Set DisplayParamsSP = oFields

Exit Function
ErrorDisplayParamsSP:
   If Err = -2147217900 Then bErrExist = True: Resume Next
   MsgBox "Error: [" & Err & "] " & Error, vbCritical, App.Title
End Function




Public Function DisplayFields(ByVal vsKey As String, ByVal vsObjectBDName As String, ByVal voLV As ListView, ByVal eDBType As enmDBType, ByVal eDisplayMode As enmDisplayFieldsMode, ByVal voSB As StatusBar) As CFields
Dim oCatalog As New ADOX.Catalog, iCount As Long
Dim iCountKeys As Integer, iCountKeyFields As Integer, bIsKey As Boolean
Dim iColumsIndex As Integer, xNode As Node, iIndex As Integer
Dim oIndexes As Indexes, oRelations As Keys, iRel As Integer
Dim objsSQLServer As Object, RsFields As New ADODB.Recordset
Dim oFields As New CFields, oField As New CField
On Error Resume Next

   Set oPBar = New CProgressBar
   oPBar.CreateProgress voSB, 5

   Select Case eDBType
      Case [None]
         Exit Function
      
      Case [SQL Server]
         Call CfgTblViewFields(voLV)
            voSB.Panels(1).Text = "Loading Fields"
            With RsFields
               .CursorLocation = adUseClient
               .ActiveConnection = oDBs.Item(vsKey)
               .Source = "select * from " & oMPG.BracketField(vsObjectBDName)
               oPBar.Max = .Fields.Count + 1
               For iCount = 0 To .Fields.Count - 1
                  oFields.Add
                  oFields.Item(iCount + 1).Name = .Fields(iCount).Name
                  oFields.Item(iCount + 1).ADOConstantValue = .Fields(iCount).Type
                  oFields.Item(iCount + 1).DefinedSize = .Fields(iCount).DefinedSize
                  oFields.Item(iCount + 1).Precision = .Fields(iCount).Precision
                  oFields.Item(iCount + 1).NumericScale = .Fields(iCount).NumericScale
                  oPBar.Value = iCount + 1
               Next iCount
               .Source = "sp_columns '" & vsObjectBDName & "'"
               .Open
               oPBar.Max = .RecordCount + 1
               voSB.Panels(1).Text = "Formating Datatypes"
               For iCount = 0 To .RecordCount - 1
                  oPBar.Value = iCount + 1
                  For iCountKeys = 1 To oFields.Count
                     If oFields.Item(iCountKeys).Name = .Fields("COLUMN_NAME") Then
                        '// Valid Identity Field
                        If Len(.Fields("TYPE_NAME")) <= 7 Then
                           If UCase(.Fields("TYPE_NAME")) = "TID" Then
                              oFields.Item(iCountKeys).StringType = "varchar"
                           Else
                              oFields.Item(iCountKeys).StringType = .Fields("TYPE_NAME")
                           End If
                        Else
                           If UCase(Mid(.Fields("TYPE_NAME"), Len(.Fields("TYPE_NAME")) - 7, Len(.Fields("TYPE_NAME")))) = "IDENTITY" Then
                              If LCase(Mid(.Fields("TYPE_NAME"), 1, 3)) = "int" Then
                                 oFields.Item(iCountKeys).StringType = Trim(Mid(.Fields("TYPE_NAME"), 1, Len(.Fields("TYPE_NAME")) - 8))
                              Else
                                 oFields.Item(iCountKeys).StringType = Trim(Mid(.Fields("TYPE_NAME"), 1, Len(.Fields("TYPE_NAME")) - 11))
                              End If
                           Else
                              oFields.Item(iCountKeys).StringType = .Fields("TYPE_NAME")
                           End If
                        End If
                        'Valid is Nullable Field
                        If .Fields("NULLABLE") = 0 Then
                           oFields.Item(iCountKeys).IsNull = True
                        Else
                           oFields.Item(iCountKeys).IsNull = False
                        End If
                        'Valid is Identity Field
                        If UCase(Mid(.Fields("TYPE_NAME"), Len(.Fields("TYPE_NAME")) - 7, Len(.Fields("TYPE_NAME")))) <> "IDENTITY" Then
                           oFields.Item(iCountKeys).IsIdentity = False
                        Else
                           oFields.Item(iCountKeys).IsIdentity = True
                        End If
                        Exit For
                     End If
                  Next iCountKeys
                  .MoveNext
               Next iCount
               .Close
               
               
               If eDisplayMode = Table Then
                  voSB.Panels(1).Text = "Loading Indexes"
                  .Open "sp_pkeys '" & vsObjectBDName & "'"
                  oPBar.Max = .RecordCount + 1
                  For iCount = 0 To .RecordCount - 1
                     oPBar.Max = iCount + 1
                     For iCountKeys = 1 To oFields.Count
                        If UCase(oFields.Item(iCountKeys).Name) = UCase(.Fields("COLUMN_NAME")) Then
                           oFields.Item(iCountKeys).IsPK = True
                        End If
                     Next iCountKeys
                     .MoveNext
                  Next iCount
                  .Close
               End If
               
               
            End With
            
            voSB.Panels(1).Text = "Write Fields in Screen"
            
            With oFields
               oPBar.Max = .Count + 1
               For iCount = 1 To .Count
                  oPBar.Value = iCount + 1
                  If .Item(iCount).IsPK = True Then
                     voLV.ListItems.Add , , .Item(iCount).Name, , SmallIcon:=6
                  Else
                     voLV.ListItems.Add , , .Item(iCount).Name, , SmallIcon:=5
                  End If
                  voLV.ListItems(iCount).SubItems(1) = .Item(iCount).StringType
                  voLV.ListItems(iCount).SubItems(2) = .Item(iCount).DefinedSize
                  voLV.ListItems(iCount).SubItems(3) = .Item(iCount).NumericScale
                  voLV.ListItems(iCount).SubItems(4) = .Item(iCount).Precision
               Next iCount
            End With
      
      Case [Microsoft Access]
         Call CfgTblViewFields(voLV)
         voSB.Panels(1).Text = "Loading Fields"
         With oCatalog
            .ActiveConnection = oDBs.Item(vsKey)
            RsFields.Source = "select * from " & oMPG.BracketField(vsObjectBDName)
            RsFields.ActiveConnection = oDBs.Item(vsKey)
            
            With .Tables(vsObjectBDName)
               Set oIndexes = .Indexes
               Set oRelations = .Keys
               oPBar.Max = RsFields.Fields.Count + 1
               With RsFields
                  For iCount = 0 To .Fields.Count - 1
                  oFields.Add
                  oFields.Item(iCount + 1).Name = .Fields(iCount).Name
                  oFields.Item(iCount + 1).ADOConstantValue = .Fields(iCount).Type
                  oFields.Item(iCount + 1).DefinedSize = .Fields(iCount).DefinedSize
                  oFields.Item(iCount + 1).Precision = .Fields(iCount).Precision
                  oFields.Item(iCount + 1).NumericScale = .Fields(iCount).NumericScale
                  oPBar.Value = iCount + 1
                  Next iCount
               End With
               
               voSB.Panels(1).Text = "Loading Indexes"
               
               If eDisplayMode = Table Then
                  With oIndexes
                     oPBar.Max = .Count + 1
                     For iCountKeys = 0 To .Count - 1
                        oPBar.Value = iCountKeys + 1
                        iIndex = iIndex + 1
                        If .Item(iCountKeys).PrimaryKey = True Then
                           For iCountKeyFields = 1 To oFields.Count
                              For iColumsIndex = 0 To .Item(iCountKeys).Columns.Count - 1
                                 If oFields.Item(iCountKeyFields).Name = .Item(iCountKeys).Columns(iColumsIndex).Name Then
                                    oFields.Item(iCountKeyFields).IsPK = True
                                 End If
                              Next iColumsIndex
                           Next iCountKeyFields
                        End If
                     Next iCountKeys
                  End With
               End If
               
               voSB.Panels(1).Text = "Write Fields in Screen"
               
               For iCount = 1 To oFields.Count
                  If oFields.Item(iCount).IsPK = True Then
                     voLV.ListItems.Add , , oFields.Item(iCount).Name, , SmallIcon:=6
                  Else
                     voLV.ListItems.Add , , oFields.Item(iCount).Name, , SmallIcon:=5
                  End If
                  voLV.ListItems(iCount).SubItems(1) = TranslateADOToAccess(oFields.Item(iCount).ADOConstantValue)
                  voLV.ListItems(iCount).SubItems(2) = oFields.Item(iCount).DefinedSize
                  voLV.ListItems(iCount).SubItems(3) = oFields.Item(iCount).NumericScale
                  voLV.ListItems(iCount).SubItems(4) = oFields.Item(iCount).Precision
               Next iCount
            End With
         End With

      Case [ODBC Database]
      
   End Select
   Set oCatalog = Nothing
   Set RsFields = Nothing
   Set DisplayFields = oFields
   oPBar.Value = 0
   Set oPBar = Nothing
   voSB.Panels(1).Text = "Listen..."
End Function



Public Function DisplayInternalCodeSQLServer(ByVal vsKey As String, ByVal vsNameDBObject As String, ByVal voSB As StatusBar) As String
Dim RsCode As New ADODB.Recordset
Dim iCount As Integer, sInternalCode As String
On Error GoTo ErrorDisplayInternalCodeSQLServer
   voSB.Panels(1).Text = "Loading Internal Code"
   sInternalCode = "Not Internal Code Display"
   With RsCode
      .ActiveConnection = oDBs.Item(vsKey)
      .CursorLocation = adUseClient
      .Open "sp_HelpText '" & vsNameDBObject & "'"
      If .RecordCount > 0 Then
         sInternalCode = ""
         Set oPBar = New CProgressBar
         oPBar.CreateProgress voSB, 5
         oPBar.Max = .RecordCount + 1
         For iCount = 0 To .RecordCount - 1
            oPBar.Value = iCount
            sInternalCode = sInternalCode & .Fields(0)
            .MoveNext
         Next iCount
      End If
      .Close
   End With
   Set RsCode = Nothing
   oPBar.Value = 0
   Set oPBar = Nothing
   DisplayInternalCodeSQLServer = sInternalCode
   voSB.Panels(1).Text = "Listen..."
Exit Function
ErrorDisplayInternalCodeSQLServer:
   Set RsCode = Nothing
   Set oPBar = Nothing
   DisplayInternalCodeSQLServer = sInternalCode
   voSB.Panels(1).Text = "Listen..."
End Function


Public Function DisplayInternalCodeAccessODBC(ByVal vsKey As String, ByVal vsNameDBObject As String) As String
'Dim RsCode As New ADODB.Recordset
   DisplayInternalCodeAccessODBC = "Not' Internal Code Display"
   'With RsCode
   '   .CursorLocation = adUseClient
   '   .Open "sp_HelpText"
   'End With

End Function





Public Sub DisplayTables(ByVal vsKey As String, ByVal voLV As ListView, ByVal voSB As StatusBar)
Dim lCount As Long
Dim oCatalog As New ADOX.Catalog
Dim abc As New ADOX.Index, RsTables As New ADODB.Recordset
On Error Resume Next
   
   Set oPBar = New CProgressBar
   oPBar.CreateProgress voSB, 5
   voSB.Panels(1).Text = "Loading Tables"
   
   With oCatalog
      
      'Assing connection with object catalog
      .ActiveConnection = oDBs.Item(vsKey)
      'Tables and Views
      Call ConfigLVTblViewSP(voLV)
      With .Tables
         oPBar.Max = .Count + 1
         For lCount = 0 To .Count - 1
            oPBar.Value = lCount
            If .Item(lCount).Type = "TABLE" Or .Item(lCount).Type = "SYSTEM TABLE" Then
               voLV.ListItems.Add , , .Item(lCount).Name, , SmallIcon:=2
               voLV.ListItems(voLV.ListItems.Count).Tag = "TABLE"
               voLV.ListItems(voLV.ListItems.Count).SubItems(1) = oCatalog.GetObjectOwner(.Item(lCount).Name, adPermObjTable)
               voLV.ListItems(voLV.ListItems.Count).SubItems(2) = .Item(lCount).Type
               voLV.ListItems(voLV.ListItems.Count).SubItems(3) = .Item(lCount).DateCreated
               voLV.ListItems(voLV.ListItems.Count).SubItems(4) = .Item(lCount).DateModified
            End If
         Next lCount
      End With
   End With
   
   oPBar.Value = 0
   Set oPBar = Nothing
      
   voSB.Panels(1).Text = "Listen..."
      
End Sub


Public Sub DisplayViews(ByVal vsKey As String, ByVal voLV As ListView, ByVal voSB As StatusBar)
Dim lCount As Long
Dim oCatalog As New ADOX.Catalog
On Error Resume Next
   
   Set oPBar = New CProgressBar
   oPBar.CreateProgress voSB, 5
   voSB.Panels(1).Text = "Loading Views"
   
   With oCatalog
      'Assing connection with object catalog
      .ActiveConnection = oDBs.Item(vsKey)
      'Tables and Views
      Call ConfigLVTblViewSP(voLV)
      With .Tables
         oPBar.Max = .Count + 1
         For lCount = 0 To .Count - 1
            oPBar.Value = lCount
            If .Item(lCount).Type = "VIEW" Or .Item(lCount).Type = "SYSTEM VIEW" Then
               voLV.ListItems.Add , , .Item(lCount).Name, , SmallIcon:=3
               voLV.ListItems(voLV.ListItems.Count).Tag = "VIEW"
               voLV.ListItems(voLV.ListItems.Count).SubItems(1) = oCatalog.GetObjectOwner(.Item(lCount).Name, adPermObjTable)
               voLV.ListItems(voLV.ListItems.Count).SubItems(2) = .Item(lCount).Type
               voLV.ListItems(voLV.ListItems.Count).SubItems(3) = .Item(lCount).DateCreated
               voLV.ListItems(voLV.ListItems.Count).SubItems(4) = .Item(lCount).DateModified
            End If
         Next lCount
      End With
   End With
   
   oPBar.Value = 0
   Set oPBar = Nothing
   
   voSB.Panels(1).Text = "Listen..."
End Sub


Public Sub DisplayProcedures(ByVal vsKey As String, ByVal voLV As ListView, ByVal voSB As StatusBar)
Dim lCount As Long
Dim oCatalog As New ADOX.Catalog
On Error Resume Next
   
   Set oPBar = New CProgressBar
   oPBar.CreateProgress voSB, 5
   voSB.Panels(1).Text = "Loading Stored Procedures"
   
   With oCatalog
      'Assing connection with object catalog
      .ActiveConnection = oDBs.Item(vsKey)
      'Tables and Views
      Call ConfigLVTblViewSP(voLV)
      With .Procedures
         oPBar.Max = .Count + 1
         For lCount = 0 To .Count - 1
            oPBar.Value = lCount + 1
            voLV.ListItems.Add , , Mid(.Item(lCount).Name, 1, Len(.Item(lCount).Name) - 2), , SmallIcon:=4
            voLV.ListItems(voLV.ListItems.Count).Tag = "PROCEDURE"
            voLV.ListItems(voLV.ListItems.Count).SubItems(1) = oCatalog.GetObjectOwner(.Item(lCount).Name, adPermObjTable)
            voLV.ListItems(voLV.ListItems.Count).SubItems(2) = "PROCEDURE"
            voLV.ListItems(voLV.ListItems.Count).SubItems(3) = .Item(lCount).DateCreated
            voLV.ListItems(voLV.ListItems.Count).SubItems(4) = .Item(lCount).DateModified
         Next lCount
      End With
   End With
   oPBar.Value = 0
   Set oPBar = Nothing
   voSB.Panels(1).Text = "Listen..."
End Sub




'Open template function
Public Function OpenLibrary(ByVal vsFilePath As String, ByVal vsFileName As String, ByVal voTV As TreeView, ByVal volvProps As ListView)
Dim xNode As Node, lNodesCount As Long, iCountLibrarys As Integer
Dim strResult As String, strResult2 As String, sTitleLib As String
   
   
   strResult = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Titulo")
   With volvProps.ListItems
      .Add
      .Item(.Count).Text = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Titulo")
      .Item(.Count).SubItems(1) = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Autor")
      .Item(.Count).SubItems(2) = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Fecha")
      .Item(.Count).SubItems(3) = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Tipo")
      .Item(.Count).SubItems(4) = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Plantillas")
      .Item(.Count).SubItems(5) = vsFilePath
      .Item(.Count).SubItems(6) = vsFileName
   End With
   
   With voTV.Nodes
      'Verify if library is open
      For lNodesCount = 1 To .Count
         If .Item(lNodesCount).Text = strResult Then
            MsgBox "Library already is open...", vbExclamation, "AMC-CodeAssist"
            Exit Function
         End If
      Next lNodesCount
      
      Set xNode = .Add(, , vsFilePath & vsFileName, strResult, Image:=10, SelectedImage:=10)
      xNode.Expanded = True
      strResult = oMPG.ReadINI(vsFilePath & vsFileName, "Configuracion", "Plantillas")
      iCountLibrarys = CLng(strResult)
      
      For lNodesCount = 1 To iCountLibrarys
         strResult = oMPG.ReadINI(vsFilePath & vsFileName, "PLT" & Format(lNodesCount, "00"), "Titulo")
         strResult2 = oMPG.ReadINI(vsFilePath & vsFileName, "PLT" & Format(lNodesCount, "00"), "Nombre")
         Set xNode = .Add(vsFilePath & vsFileName, tvwChild, vsFilePath & strResult2, strResult, Image:=11, SelectedImage:=11)
      Next lNodesCount
      
      
   End With

End Function


Public Sub LoadTemplates(ByVal vsFileName As String, ByVal iTemplates As Integer, ByVal voListTemp As ListView, sPath As String)
Dim iCountTemp As Integer
Dim sResult As String, sResult2 As String
   With voListTemp.ListItems
      .Clear
      For iCountTemp = 1 To iTemplates
         sResult = oMPG.ReadINI(vsFileName, "PLT" & Format(iCountTemp, "00"), "Titulo")
         sResult2 = oMPG.ReadINI(vsFileName, "PLT" & Format(iCountTemp, "00"), "Nombre")
         .Add
         .Item(iCountTemp).Text = sResult
         .Item(iCountTemp).SubItems(1) = sPath
         .Item(iCountTemp).SubItems(2) = sResult2
      Next iCountTemp
   End With
End Sub



'Public Method with Display information in txtResult
Public Sub DisplayInfo(ByVal sModule As String, ByVal sInfo As String)
   
   'With MDIGen.txtBuild
   '   .Text = .Text & sModule & " ->   " & sInfo & vbCrLf
   '   .SelStart = Len(.Text)
   '   .SelLength = Len(.Text)
   'End With
   
End Sub

Public Sub InfoClear()
   'MDIGen.txtBuild.Text = ""
End Sub



'Function para transladar la constante del tipo de campo ADO
'a descripcion de campo Microsoft Access
Public Function TranslateADOToAccess(ByVal eDataType As DataTypeEnum) As String
   Select Case eDataType
      Case 11: TranslateADOToAccess = "Boolean"
      Case 17: TranslateADOToAccess = "Byte"
      Case 2: TranslateADOToAccess = "Integer"
      Case 3: TranslateADOToAccess = "Long"
      Case 6: TranslateADOToAccess = "Currency"
      Case 4: TranslateADOToAccess = "Single"
      Case 5: TranslateADOToAccess = "Double"
      Case 7: TranslateADOToAccess = "Date/Time"
      Case 129: TranslateADOToAccess = "Text"
      Case 130: TranslateADOToAccess = "Text"
      Case 205: TranslateADOToAccess = "Binary"
      Case 201: TranslateADOToAccess = "Memo"
      Case 202: TranslateADOToAccess = "Text"
      Case 203: TranslateADOToAccess = "Memo"
      Case Else: TranslateADOToAccess = "!!Error!!"
   End Select
End Function


Private Sub Class_Initialize()
   'oPBar
End Sub

Private Sub Class_Terminate()
   Set oPBar = Nothing
End Sub



'Funcion de conversion de tipo de campos SQL Server
Public Function ConvertValueADONameConstant(ByVal nTipo As Long) As String
   Select Case nTipo
      Case 20: ConvertValueADONameConstant = "adBigInt"
      Case 128: ConvertValueADONameConstant = "adBinary"
      Case 11: ConvertValueADONameConstant = "adBoolean"
      Case 8: ConvertValueADONameConstant = "adBSTR"
      Case 136: ConvertValueADONameConstant = "adChapter"
      Case 129: ConvertValueADONameConstant = "adChar"
      Case 6: ConvertValueADONameConstant = "adCurrency"
      Case 7: ConvertValueADONameConstant = "adDate"
      Case 133: ConvertValueADONameConstant = "adDBDate"
      Case 134: ConvertValueADONameConstant = "adDBTime"
      Case 135: ConvertValueADONameConstant = "adDBTimeStamp"
      Case 14: ConvertValueADONameConstant = "adDecimal"
      Case 5: ConvertValueADONameConstant = "adDouble"
      Case 0: ConvertValueADONameConstant = "adEmpty"
      Case 10: ConvertValueADONameConstant = "adError"
      Case 64: ConvertValueADONameConstant = "adFileTime"
      Case 72: ConvertValueADONameConstant = "adGUID"
      Case 9: ConvertValueADONameConstant = "adIDispatch"
      Case 3: ConvertValueADONameConstant = "adInteger"
      Case 13: ConvertValueADONameConstant = "adIUnknown"
      Case 205: ConvertValueADONameConstant = "adLongVarBinary"
      Case 201: ConvertValueADONameConstant = "adLongVarChar"
      Case 203: ConvertValueADONameConstant = "adLongVarWChar"
      Case 131: ConvertValueADONameConstant = "adNumeric"
      Case 138: ConvertValueADONameConstant = "adPropVariant"
      Case 4: ConvertValueADONameConstant = "adSingle"
      Case 2: ConvertValueADONameConstant = "adSmallInt"
      Case 16: ConvertValueADONameConstant = "adTinyInt"
      Case 21: ConvertValueADONameConstant = "adUnsignedBigInt"
      Case 19: ConvertValueADONameConstant = "adUnsignedInt"
      Case 18: ConvertValueADONameConstant = "adUnsignedSmallInt"
      Case 17: ConvertValueADONameConstant = "adUnsignedTinyInt"
      Case 132: ConvertValueADONameConstant = "adUserDefined"
      Case 204: ConvertValueADONameConstant = "adVarBinary"
      Case 200: ConvertValueADONameConstant = "adVarChar"
      Case 12: ConvertValueADONameConstant = "adVariant"
      Case 139: ConvertValueADONameConstant = "adVarNumeric"
      Case 202: ConvertValueADONameConstant = "adVarWChar"
      Case 130: ConvertValueADONameConstant = "adWChar"
      Case Else: ConvertValueADONameConstant = "UserDefined"
   End Select
End Function


