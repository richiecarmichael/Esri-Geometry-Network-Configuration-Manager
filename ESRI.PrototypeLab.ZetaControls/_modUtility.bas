'Attribute VB_Name = "modUtility"
'Option Explicit

'Private Const MODULE_NAME As String = "modUtility.bas"

''==================================================================================

''--------------------------
'' RESET ObjectClass GUID's
''--------------------------

'Public Sub GUIDReset(ByRef pWorkspace As esriGeoDatabase.IWorkspace)
'    On Error GoTo ErrorHandler
'    '
'    Dim pMouseCursor As esriFramework.IMouseCursor
'    '
'    Dim pEnumDatasetNameDT As IEnumDatasetName
'    Dim pEnumDatasetNameFD As IEnumDatasetName
'    Dim pFeatureDatasetName As IFeatureDatasetName
'    Dim pDatasetNameDT As IDatasetName
'    Dim pDatasetNameFD As IDatasetName
'    Dim pDatasetDT As IDataset
'    Dim pDatasetFD As IDataset
'    Dim pName As IName
'    '---------------------
'    ' Change Mouse Cursor
'    '---------------------
'27:     Set pMouseCursor = New esriFramework.MouseCursor
'28:     Call pMouseCursor.SetCursor(2)
'    '
'30:     Set pEnumDatasetNameDT = pWorkspace.DatasetNames(esriDTAny)
'31:     Call pEnumDatasetNameDT.Reset
'32:     Set pDatasetNameDT = pEnumDatasetNameDT.Next
'33:     Do Until pDatasetNameDT Is Nothing
'        Select Case pDatasetNameDT.Type
'        Case esriDatasetType.esriDTFeatureDataset
'36:             Set pFeatureDatasetName = pDatasetNameDT
'37:             Set pEnumDatasetNameFD = pFeatureDatasetName.FeatureClassNames
'38:             Call pEnumDatasetNameFD.Reset
'39:             Set pDatasetNameFD = pEnumDatasetNameFD.Next
'40:             Do Until pDatasetNameFD Is Nothing
'41:                 If pDatasetNameFD.Type = esriDatasetType.esriDTFeatureClass Then
'42:                     Call frmOutputWindow.AddOutputMessage("Updating GUID's for " & pDatasetNameFD.Name)
'43:                     Set pName = pDatasetNameFD
'44:                     Set pDatasetFD = pName.Open
'45:                     Call SetGUID(pDatasetFD)
'46:                 End If
'47:                 Set pDatasetNameFD = pEnumDatasetNameFD.Next
'48:             Loop
'        Case esriDatasetType.esriDTFeatureClass, _
'             esriDatasetType.esriDTTable
'51:             Call frmOutputWindow.AddOutputMessage("Updating GUID's for " & pDatasetNameDT.Name)
'52:             Set pName = pDatasetNameDT
'53:             Set pDatasetDT = pName.Open
'54:             Call SetGUID(pDatasetDT)
'55:         End Select
'56:         Set pDatasetNameDT = pEnumDatasetNameDT.Next
'57:     Loop
'    '---------------------
'    ' Restore MouseCursor
'    '---------------------
'61:     Set pMouseCursor = Nothing
'    '
'    Exit Sub
'ErrorHandler:
'65:     Call HandleError(False, "GUIDReset " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub SetGUID(ByRef pDataset As IDataset)
'    On Error GoTo ErrorHandler
'    '
'    Dim pClassSchemaEdit As IClassSchemaEdit
'    Dim pFeatureClass As IFeatureClass
'    Dim pCLSID As UID
'    Dim pEXTCLSID As UID
'    '
'76:     Set pClassSchemaEdit = pDataset
'77:     Set pCLSID = New UID
'78:     Set pEXTCLSID = New UID
'    '
'    Select Case pDataset.Type
'    Case esriDTTable
'82:         pCLSID.Value = GUID_TABLE_CLSID
'83:         Set pEXTCLSID = Nothing
'    Case esriDTFeatureClass
'85:         Set pFeatureClass = pDataset
'        Select Case pFeatureClass.FeatureType
'        Case esriFTAnnotation
'88:             pCLSID.Value = GUID_ANNOTATION_CLSID
'89:             pEXTCLSID.Value = GUID_ANNOTATION_EXTCLSID
'        Case esriFTDimension
'91:             pCLSID.Value = GUID_DIMENSION_CLSID
'92:             pEXTCLSID.Value = GUID_DIMENSION_EXTCLSID
'        Case esriFTComplexEdge
'94:             pCLSID.Value = GUID_COMPLEXEDGE_CLSID
'95:             Set pEXTCLSID = Nothing
'        Case esriFTSimpleEdge
'97:             pCLSID.Value = GUID_SIMPLEEDGE_CLSID
'98:             Set pEXTCLSID = Nothing
'        Case esriFTSimpleJunction
'100:             pCLSID.Value = GUID_SIMPLEJUNCTION_CLSID
'101:             Set pEXTCLSID = Nothing
'        Case esriFTSimple
'103:             pCLSID.Value = GUID_FEATURECLASS_CLSID
'104:             Set pEXTCLSID = Nothing
'105:         End Select
'106:     End Select
'    '
'108:     Call pClassSchemaEdit.AlterInstanceCLSID(pCLSID)
'    '
'110:     If pEXTCLSID Is Nothing Then
'111:         Call pClassSchemaEdit.AlterClassExtensionCLSID(Nothing, Nothing)
'112:     Else
'113:         Call pClassSchemaEdit.AlterClassExtensionCLSID(pEXTCLSID, Nothing)
'114:     End If
'    '
'    Exit Sub
'ErrorHandler:
'118:     Call HandleError(False, "SetGUID " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub TruncateGeodatabase(ByRef pDataset As esriGeoDatabase.IDataset)
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnumDataset As esriGeoDatabase.IEnumDataset
'    Dim pDatasetCH As esriGeoDatabase.IDataset
'    '
'    Select Case pDataset.Type
'    Case esriDatasetType.esriDTContainer, _
'         esriDatasetType.esriDTFeatureDataset
'        '----------------------------
'        ' Iterate Through Each Child
'        '----------------------------
'133:         Set pEnumDataset = pDataset.Subsets
'134:         If Not (pEnumDataset Is Nothing) Then
'135:             Call pEnumDataset.Reset
'136:             Set pDatasetCH = pEnumDataset.Next
'137:             Do Until pDatasetCH Is Nothing
'138:                 Call TruncateGeodatabase(pDatasetCH)
'139:                 Set pDatasetCH = pEnumDataset.Next
'140:             Loop
'141:         End If
'    Case esriDatasetType.esriDTTable, _
'         esriDatasetType.esriDTFeatureClass, _
'         esriDatasetType.esriDTRelationshipClass
'        '-------------------------------------
'        ' Call Truncate ObjectClass Procedure
'        '-------------------------------------
'148:         Call TruncateObjectClass(pDataset)
'    Case Else
'        '------------------------------
'        ' Not Supported for Truncation
'        '------------------------------
'153:     End Select
'    '
'    Exit Sub
'ErrorHandler:
'157:     Call HandleError(False, "TruncateGeodatabase " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub TruncateObjectClass(ByRef pDataset As esriGeoDatabase.IDataset)
'    On Error GoTo ErrorHandler
'    '
'163:     Call frmOutputWindow.AddOutputMessage("Truncating FeatureClass/Table/Relationship: " & pDataset.Name)
'    '
'    Dim pTable As esriGeoDatabase.ITable
'    Dim pQueryFilter As esriGeoDatabase.IQueryFilter
'    Dim pWorkspace As esriGeoDatabase.IWorkspace
'    Dim pWorkspaceEdit As esriGeoDatabase.IWorkspaceEdit
'    '
'    Dim lngRowCount As Long
'    '
'172:     Set pTable = pDataset
'173:     Set pQueryFilter = New esriGeoDatabase.QueryFilter
'    '
'175:     lngRowCount = pTable.RowCount(pQueryFilter)
'    '
'177:     If lngRowCount > 0 Then
'178:         Set pWorkspace = pDataset.Workspace
'179:         Set pWorkspaceEdit = pWorkspace
'        '
'181:         Call pWorkspaceEdit.StartEditing(False)
'182:         Call pWorkspaceEdit.StartEditOperation
'        '
'184:         Call pTable.DeleteSearchedRows(pQueryFilter)
'        '
'186:         Call pWorkspaceEdit.StopEditOperation
'187:         Call pWorkspaceEdit.StopEditing(True)
'188:     End If
'    '
'190:     Call frmOutputWindow.AddOutputMessage("   " & CStr(lngRowCount) & " Rows/Features Removed")
'    '
'    Exit Sub
'ErrorHandler:
'194:     Call HandleError(False, "TruncateObjectClass " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

