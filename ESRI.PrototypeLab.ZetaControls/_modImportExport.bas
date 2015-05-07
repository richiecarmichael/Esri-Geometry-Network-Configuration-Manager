'Attribute VB_Name = "modImportExport"
'Option Explicit

'Private Const MODULE_NAME As String = "modImportExport.bas"
'Private Const MISSING As Long = -999999999

''==================================================================================

''-----------------
'' EXPORT - DOMAIN
''-----------------

'Public Sub ExportDomain(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                        ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    '------------------------------
'    ' Get GeodatabaseDesigner Node
'    '------------------------------
'20:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '
'22:     Call ExportDomain2(pXMLDOMNodeGeodatabaseDesigner, pWorkspace)
'    '
'    Exit Sub
'ErrorHandler:
'26:     Call HandleError(False, "ExportDomain " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ExportDomain2(ByRef pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode, _
'                         ByRef pWorkspace As esriGeoDatabase.IWorkspace)
'    On Error GoTo ErrorHandler
'    '
'    Dim pWorkspaceDomains As esriGeoDatabase.IWorkspaceDomains
'    Dim pEnumDomain As esriGeoDatabase.IEnumDomain
'    Dim pDomain As esriGeoDatabase.IDomain
'    Dim pRangeDomain As esriGeoDatabase.IRangeDomain
'    Dim pCodedValueDomain As esriGeoDatabase.ICodedValueDomain
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pEnumDataset As esriGeoDatabase.IEnumDataset
'    '
'    Dim pDataset2 As esriGeoDatabase.IDataset
'    Dim pEnumDataset2 As esriGeoDatabase.IEnumDataset
'    Dim pFeatureClassContainer2 As esriGeoDatabase.IFeatureClassContainer
'    '
'    Dim pIndexMember As Long
'    '
'    'Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    Dim pXMLDOMNodeDomain As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeMember As MSXML2.IXMLDOMNode
'    '------------------
'    ' Set DOM Document
'    '------------------
'54:     Set pDOMDocument = pXMLDOMNodeGeodatabaseDesigner.ownerDocument
''    '------------------------------
''    ' Get GeodatabaseDesigner Node
''    '------------------------------
''    Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '------------------------
'    ' Get Domains Enumerator
'    '------------------------
'62:     Set pWorkspaceDomains = pWorkspace
'63:     Set pEnumDomain = pWorkspaceDomains.Domains
'64:     If pEnumDomain Is Nothing Then
'65:         Call frmOutputWindow.AddOutputMessage("No Domains Found in Target Geodatabase")
'        '
'        Exit Sub
'68:     End If
'    '----------------------
'    ' Loop for each Domain
'    '----------------------
'72:     Set pDomain = pEnumDomain.Next
'73:     Do Until pDomain Is Nothing
'        '----------------------
'        ' Update Output Window
'        '----------------------
'77:         Call frmOutputWindow.AddOutputMessage("Exporting Domain: " & pDomain.Name, , , vbMagenta)
'        '------------------------
'        ' Create new Domain Node
'        '------------------------
'81:         Set pXMLDOMNodeDomain = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "domain", "")
'82:         Set pXMLDOMNodeDomain = pXMLDOMNodeGeodatabaseDesigner.appendChild(pXMLDOMNodeDomain)
'        '-----------------------
'        ' Add Domain Attributes
'        '-----------------------
'86:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "name", pDomain.Name)
'87:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "description", pDomain.Description)
'88:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "owner", pDomain.Owner)
'89:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "esriDomainType", CStr(pDomain.Type))
'90:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "esriFieldType", CStr(pDomain.FieldType))
'91:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "esriMergePolicyType", CStr(pDomain.MergePolicy))
'92:         Call modCommon.AddNodeAttribute(pXMLDOMNodeDomain, "esriSplitPolicyType", CStr(pDomain.SplitPolicy))
'        Select Case pDomain.Type
'        Case esriDTCodedValue
'            '----------------------------------
'            ' Add CodedValue Domain Attributes
'            '----------------------------------
'98:             Set pCodedValueDomain = pDomain
'99:             For pIndexMember = 0 To pCodedValueDomain.CodeCount - 1 Step 1
'100:                 Set pXMLDOMNodeMember = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "member", "")
'101:                 Set pXMLDOMNodeMember = pXMLDOMNodeDomain.appendChild(pXMLDOMNodeMember)
'102:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeMember, "name", pCodedValueDomain.Name(pIndexMember))
'103:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeMember, "value", CStr(pCodedValueDomain.Value(pIndexMember)))
'104:             Next pIndexMember
'        Case esriDTRange
'            '-----------------------------
'            ' Add Range Domain Attributes
'            '-----------------------------
'109:             Set pRangeDomain = pDomain
'110:             Set pXMLDOMNodeMember = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "member", "")
'111:             Set pXMLDOMNodeMember = pXMLDOMNodeDomain.appendChild(pXMLDOMNodeMember)
'112:             Call modCommon.AddNodeAttribute(pXMLDOMNodeMember, "minValue", CStr(pRangeDomain.MinValue))
'113:             Call modCommon.AddNodeAttribute(pXMLDOMNodeMember, "maxValue", CStr(pRangeDomain.MaxValue))
'        Case esriDTString
'            '--------------------------------------
'            ' Not Supported - What is this anyway?
'            '--------------------------------------
'118:         End Select
'119:         Set pDomain = pEnumDomain.Next
'120:     Loop
'    '
'122:     If mDMExportOCAssocation Then
'        '-----------------------------
'        ' Add all domain dependancies
'        '-----------------------------
'126:         Set pEnumDataset = pWorkspace.Datasets(esriDTAny)
'127:         Set pDataset = pEnumDataset.Next
'        '
'129:         Do Until pDataset Is Nothing
'130:             If TypeOf pDataset Is IFeatureDataset Then
'131:                 Set pFeatureClassContainer2 = pDataset
'132:                 Set pEnumDataset2 = pFeatureClassContainer2.Classes
'133:                 Set pDataset2 = pEnumDataset2.Next
'134:                 Do Until pDataset2 Is Nothing
'135:                     If TypeOf pDataset2 Is esriGeoDatabase.IObjectClass Then
'136:                         Call frmOutputWindow.AddOutputMessage("Exporting Domain Assocations From: " & pDataset2.Name)
'                        '
'138:                         Call ExportObjectClassAssociation(pDOMDocument, pDataset2)
'139:                     End If
'140:                     Set pDataset2 = pEnumDataset2.Next
'141:                 Loop
'142:             Else
'143:                 If TypeOf pDataset Is esriGeoDatabase.IObjectClass Then
'144:                     Call frmOutputWindow.AddOutputMessage("Exporting Domain Assocations From: " & pDataset.Name)
'                    '
'146:                     Call ExportObjectClassAssociation(pDOMDocument, pDataset)
'147:                 End If
'148:             End If
'149:             Set pDataset = pEnumDataset.Next
'150:         Loop
'151:     Else
'        '---------------------------
'        ' Dependancies not exported
'        '---------------------------
'155:         Call frmOutputWindow.AddOutputMessage("ObjectClass Dependencies Not Exported")
'156:     End If
'    '
'    Exit Sub
'ErrorHandler:
'160:     Call HandleError(False, "ExportDomain2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportObjectClassAssociation(ByRef pDOMDocument As MSXML2.DOMDocument40, _
'                                         ByRef pObjectClass As esriGeoDatabase.IObjectClass)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pSubtypes As esriGeoDatabase.ISubtypes
'    Dim pSQLSyntax As esriGeoDatabase.ISQLSyntax
'    Dim pDomain As esriGeoDatabase.IDomain
'    Dim pEnumSubtype As esriGeoDatabase.IEnumSubtype
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pField As esriGeoDatabase.IField

'    Dim pDatabaseName As String
'    Dim pOwnerName As String
'    Dim pTableName As String
'    Dim pSubtypeName As String
'    Dim pSubtypeCode As Long
'    Dim pIndexField As Long
'    '-------------------------------------------
'    ' Interate throught each subtytpe and field
'    '-------------------------------------------
'184:     Set pDataset = pObjectClass
'185:     Set pSQLSyntax = pDataset.Workspace
'186:     Set pSubtypes = pObjectClass
'187:     Set pFields = pObjectClass.fields
'    '
'189:     Call pSQLSyntax.ParseTableName(CStr(pDataset.Name), pDatabaseName, pOwnerName, pTableName)
'    '
'    '--------------------------------
'    ' ObjectClasses without Subtypes
'    '--------------------------------
'194:     For pIndexField = 0 To pFields.FieldCount - 1 Step 1
'195:         Set pField = pFields.Field(pIndexField)
'196:         Set pDomain = pField.Domain
'197:         If pDomain Is Nothing Then
'            '---------------------
'            ' Field has no domain
'            '---------------------
'201:         Else
'202:             Call ExportObjectClassAssociationNode(pDOMDocument, pDatabaseName, pOwnerName, pTableName, "", pField.Name, pDomain.Name)
'203:         End If
'204:     Next pIndexField
'    '
'206:     If pSubtypes.HasSubtype Then
'       '-----------------------------
'       ' ObjectClasses with Subtypes
'       '-----------------------------
'210:        Set pEnumSubtype = pSubtypes.Subtypes
'211:        pSubtypeName = pEnumSubtype.Next(pSubtypeCode)
'212:        Do Until pSubtypeName = ""
'213:            For pIndexField = 0 To pFields.FieldCount - 1 Step 1
'214:                Set pField = pFields.Field(pIndexField)
'215:                Set pDomain = pSubtypes.Domain(pSubtypeCode, pField.Name)
'216:                If pDomain Is Nothing Then
'                   '--------------------------
'                   ' This field has no Domain
'                   '--------------------------
'220:                Else
'221:                    Call ExportObjectClassAssociationNode(pDOMDocument, pDatabaseName, pOwnerName, pTableName, pSubtypeName, pField.Name, pDomain.Name)
'222:                End If
'223:            Next pIndexField
'224:            pSubtypeName = pEnumSubtype.Next(pSubtypeCode)
'225:        Loop
'226:     End If
'    '
'    Exit Sub
'ErrorHandler:
'230:     Call HandleError(False, "ExportObjectClassAssociation " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportObjectClassAssociationNode(ByRef pDOMDocument As MSXML2.DOMDocument40, _
'                                             ByRef pDatabaseName As String, _
'                                             ByRef pOwnerName As String, _
'                                             ByRef pTableName As String, _
'                                             ByRef pSubtypeName As String, _
'                                             ByRef pFieldName As String, _
'                                             ByRef pDomainName As String)
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeDomain As MSXML2.IXMLDOMNode
'    '
'    Dim pXMLDOMNode As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeParent As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeList As MSXML2.IXMLDOMNodeList
'    '----------------------------
'    ' Get Domain Collection Node
'    '----------------------------
'251:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '--------------------------------------------
'    ' Get Domain Node from DomainCollection Node
'    ' <domain name="" ...>
'    '--------------------------------------------
'256:     Set pXMLDOMNodeDomain = pXMLDOMNodeGeodatabaseDesigner.selectNodes("domain[@name='" & pDomainName & "']").nextNode
'    '------------------------------------------------
'    ' Get Child FeatureClass
'    ' <domain name="" ...>
'    '    <objectClass database="" owner="" table="">
'    '-------------------------------------------------
'262:     Set pXMLDOMNodeList = pXMLDOMNodeDomain.selectNodes("objectClass[@database='" & pDatabaseName & "'" & _
'                                                                   " and @owner='" & pOwnerName & "'" & _
'                                                                   " and @table='" & pTableName & "']")
'265:     If pXMLDOMNodeList.length = 0 Then
'        '--------------
'        ' Add New Node
'        '--------------
'269:         Set pXMLDOMNode = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "objectClass", "")
'270:         Set pXMLDOMNode = pXMLDOMNodeDomain.appendChild(pXMLDOMNode)
'        '
'272:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "database", pDatabaseName)
'273:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "owner", pOwnerName)
'274:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "table", pTableName)
'        '
'276:         Set pXMLDOMNodeParent = pXMLDOMNode
'277:     Else
'278:         Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'279:     End If
'    '------------------------------------------------
'    ' Get Child Subtype
'    ' <domain name="" ...>
'    '    <objectClass database="" owner="" table="">
'    '       <subtype name="">
'    '------------------------------------------------
'286:     Set pXMLDOMNodeList = pXMLDOMNodeParent.selectNodes("subtype[@name='" & pSubtypeName & "']")
'287:     If pXMLDOMNodeList.length = 0 Then
'        '--------------
'        ' Add New Node
'        '--------------
'291:         Set pXMLDOMNode = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'292:         Set pXMLDOMNode = pXMLDOMNodeParent.appendChild(pXMLDOMNode)
'        '
'294:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "name", pSubtypeName)
'        '
'296:         Set pXMLDOMNodeParent = pXMLDOMNode
'297:     Else
'298:         Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'299:     End If
'    '------------------------------------------------
'    ' Get Child Field
'    ' <domain name="" ...>
'    '    <objectClass database="" owner="" table="">
'    '       <subtype name="">
'    '          <field name="">
'    '------------------------------------------------
'    '--------------
'    ' Add New Node
'    '--------------
'310:     Set pXMLDOMNode = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "field", "")
'311:     Set pXMLDOMNode = pXMLDOMNodeParent.appendChild(pXMLDOMNode)
'    '
'313:     Call modCommon.AddNodeAttribute(pXMLDOMNode, "name", pFieldName)
'    '
'    Exit Sub
'ErrorHandler:
'317:     Call HandleError(False, "ExportObjectClassAssociationNode " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

''==================================================================================

''---------------
'' Domain Import
''---------------

'Public Sub ImportDomain(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                        ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDatabaseConnectionInfo As esriGeoDatabase.IDatabaseConnectionInfo
'    Dim pWorkspaceDomains2 As esriGeoDatabase.IWorkspaceDomains2
'    Dim pDomain As esriGeoDatabase.IDomain
'    Dim pRangeDomain As esriGeoDatabase.IRangeDomain
'    Dim pCodedValueDomain As esriGeoDatabase.ICodedValueDomain
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListDomain As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeDomain As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListMember As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeMember As MSXML2.IXMLDOMNode
'    '
'    Dim pDomainName As String
'    Dim pIndexDomain As Long
'    Dim pIndexMember As Long
'    Dim blnNewDomain As Boolean
'    Dim pDomainOwnerOriginal As String
'    Dim pMemberName As String
'    Dim pMemberValue As String
'    Dim pMinValue As String
'    Dim pMaxValue As String
'    '-------------------------
'    ' QI to IWorkspaceDomains
'    '-------------------------
'354:     Set pWorkspaceDomains2 = pWorkspace
'355:     If pWorkspace.Type = esriGeoDatabase.esriWorkspaceType.esriRemoteDatabaseWorkspace Then
'356:         Set pDatabaseConnectionInfo = pWorkspace
'357:     End If
'    '---------------------------
'    ' Get New Domain Collection
'    '---------------------------
'361:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'362:     Set pXMLDOMNodeListDomain = pXMLDOMNodeGeodatabaseDesigner.selectNodes("domain")
'363:     For pIndexDomain = 0 To pXMLDOMNodeListDomain.length - 1 Step 1
'        '-----------------
'        ' Get Domain Node
'        '-----------------
'367:         Set pXMLDOMNodeDomain = pXMLDOMNodeListDomain.Item(pIndexDomain)
'        '-----------------
'        ' Get Domain Name
'        '-----------------
'371:         pDomainName = pXMLDOMNodeDomain.Attributes.getNamedItem("name").Text
'        '--------------------
'        ' Add Output Message
'        '--------------------
'375:         Call frmOutputWindow.AddOutputMessage("Importing Domain: " & pDomainName, , , vbMagenta)
'        '--------------------------------------
'        ' Get Old Domain (if any) by same name
'        '--------------------------------------
'379:         Set pDomain = pWorkspaceDomains2.DomainByName(pDomainName)
'380:         If pDomain Is Nothing Then
'            '----------------------------------------------------
'            ' An existing domain by the same name does not exist
'            '----------------------------------------------------
'384:             blnNewDomain = True
'            '---------------------
'            ' Create a NEW domain
'            '---------------------
'            Select Case CLng(pXMLDOMNodeDomain.Attributes.getNamedItem("esriDomainType").Text)
'            Case esriGeoDatabase.esriDomainType.esriDTCodedValue
'390:                 Set pDomain = New esriGeoDatabase.CodedValueDomain
'            Case esriGeoDatabase.esriDomainType.esriDTRange
'392:                 Set pDomain = New esriGeoDatabase.RangeDomain
'393:             End Select
'            '------------------------
'            ' Set Name of NEW domain
'            '------------------------
'397:             pDomain.Name = pDomainName
'398:         Else
'            '-----------------------
'            ' Domain Already Exists
'            '-----------------------
'402:             Call frmOutputWindow.AddOutputMessage("Domain Already Exists", , , , , , True, 1)
'            '
'404:             blnNewDomain = False
'            '-----------------------------
'            ' Remove all existing members
'            '-----------------------------
'408:             If pDomain.Type = esriDTCodedValue Then
'409:                 Call frmOutputWindow.AddOutputMessage("Removing Existing Domain Members", , , , , , True, 1)
'                '
'411:                 Set pCodedValueDomain = pDomain
'412:                 For pIndexMember = pCodedValueDomain.CodeCount - 1 To 0 Step -1
'413:                     Call pCodedValueDomain.DeleteCode(pCodedValueDomain.Value(pIndexMember))
'414:                 Next pIndexMember
'415:             End If
'416:         End If
'        '----------------
'        ' Set Properties
'        '----------------
'420:         pDomain.Description = CStr(pXMLDOMNodeDomain.Attributes.getNamedItem("description").Text)
'421:         pDomain.FieldType = CLng(pXMLDOMNodeDomain.Attributes.getNamedItem("esriFieldType").Text)
'422:         pDomain.MergePolicy = CLng(pXMLDOMNodeDomain.Attributes.getNamedItem("esriMergePolicyType").Text)
'423:         pDomain.SplitPolicy = CLng(pXMLDOMNodeDomain.Attributes.getNamedItem("esriSplitPolicyType").Text)
'        '-------------
'        ' Set Members
'        '-------------
'427:         Set pXMLDOMNodeListMember = pXMLDOMNodeDomain.selectNodes("member")
'        Select Case pDomain.Type
'        Case esriGeoDatabase.esriDomainType.esriDTCodedValue
'430:             Set pCodedValueDomain = pDomain
'431:             For pIndexMember = 0 To pXMLDOMNodeListMember.length - 1 Step 1
'432:                 Set pXMLDOMNodeMember = pXMLDOMNodeListMember.Item(pIndexMember)
'                '
'434:                 pMemberValue = CStr(pXMLDOMNodeMember.Attributes.getNamedItem("value").Text)
'435:                 pMemberName = CStr(pXMLDOMNodeMember.Attributes.getNamedItem("name").Text)
'                '
'437:                 Call frmOutputWindow.AddOutputMessage("Importing Member. Name[" & pMemberName & "] Value[" & pMemberValue & "]", , , , , , True, 1)
'                Select Case pDomain.FieldType
'                Case esriFieldType.esriFieldTypeSmallInteger:
'440:                     Call pCodedValueDomain.AddCode(CVar(CInt(pMemberValue)), CStr(pMemberName))
'                Case esriFieldType.esriFieldTypeInteger:
'442:                     Call pCodedValueDomain.AddCode(CVar(CLng(pMemberValue)), CStr(pMemberName))
'                Case esriFieldType.esriFieldTypeSingle:
'444:                     Call pCodedValueDomain.AddCode(CVar(CSng(pMemberValue)), CStr(pMemberName))
'                Case esriFieldType.esriFieldTypeDouble:
'446:                     Call pCodedValueDomain.AddCode(CVar(CDbl(pMemberValue)), CStr(pMemberName))
'                Case esriFieldType.esriFieldTypeString:
'448:                     Call pCodedValueDomain.AddCode(CVar(CStr(pMemberValue)), CStr(pMemberName))
'                Case esriFieldType.esriFieldTypeDate:
'450:                     Call pCodedValueDomain.AddCode(CVar(CDate(pMemberValue)), CStr(pMemberName))
'451:                 End Select
'452:             Next pIndexMember
'        Case esriGeoDatabase.esriDomainType.esriDTRange
'454:             Set pRangeDomain = pDomain
'455:             Set pXMLDOMNodeMember = pXMLDOMNodeListMember.Item(0)
'            '
'457:             pMinValue = CStr(pXMLDOMNodeMember.Attributes.getNamedItem("minValue").Text)
'458:             pMaxValue = CStr(pXMLDOMNodeMember.Attributes.getNamedItem("maxValue").Text)
'            '
'460:             Call frmOutputWindow.AddOutputMessage("Importing Range Values. Min[" & pMinValue & "] Max[" & pMaxValue & "]", , , , , , True, 1)
'            '
'462:             pRangeDomain.MinValue = CVar(pMinValue)
'463:             pRangeDomain.MaxValue = CVar(pMaxValue)
'464:         End Select
'        '----------------------------------------
'        ' Either ADD or ALTER an existing domain
'        '----------------------------------------
'468:         If blnNewDomain Then
'            '------------
'            ' Add Domain
'            '------------
'472:             Call pWorkspaceDomains2.AddDomain(pDomain)
'473:             If pWorkspace.Type = esriRemoteDatabaseWorkspace Then
'474:                 If CStr(pXMLDOMNodeDomain.Attributes.getNamedItem("owner").Text) <> "" Then
'475:                     pDomain.Owner = CStr(pXMLDOMNodeDomain.Attributes.getNamedItem("owner").Text)
'476:                 End If
'477:             End If
'478:         Else
'            '--------------
'            ' Alter Domain
'            '--------------
'482:             pDomainOwnerOriginal = pDomain.Owner
'            '
'484:             If pWorkspace.Type = esriRemoteDatabaseWorkspace Then
'485:                 pDomain.Owner = pDatabaseConnectionInfo.ConnectedUser
'486:             End If
'487:             Call pWorkspaceDomains2.AlterDomain(pDomain)
'            '--------------------------------
'            ' Change the Owner of the Domain
'            '--------------------------------
'491:             If pWorkspace.Type = esriRemoteDatabaseWorkspace Then
'492:                 If CStr(pXMLDOMNodeDomain.Attributes.getNamedItem("owner").Text) = "" Then
'493:                     pDomain.Owner = pDomainOwnerOriginal
'494:                 Else
'495:                     pDomain.Owner = CStr(pXMLDOMNodeDomain.Attributes.getNamedItem("owner").Text)
'496:                 End If
'497:             End If
'498:         End If
'499:     Next pIndexDomain
'    '
'    Exit Sub
'ErrorHandler:
'    Select Case Err.Number
'    Case FDO_E_CODED_VALUE_DOMAIN_VALUE_NOT_COMPATIBLE
'505:         Call frmOutputWindow.AddOutputMessage("The value is not compatible with the field types", , , vbRed, , , True, 1)
'506:         Resume Next
'    Case Else
'508:         Call HandleError(False, "ImportDomain " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'509:     End Select
'End Sub

''==================================================================================

''--------------------
'' ObjectClass Import
''--------------------

'Public Sub ImportGDB(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                     ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListDataset As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeDataset As MSXML2.IXMLDOMNode
'    '
'    Dim pIndex As Long
'    '-----------------------------------------------
'    ' Get GeodatabaseDesigner Node and Dataset list
'    '-----------------------------------------------
'530:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'531:     Set pXMLDOMNodeListDataset = pXMLDOMNodeGeodatabaseDesigner.childNodes
'    '-------------------------------------------------------
'    ' Launch procedure for each type of dataset encountered
'    '-------------------------------------------------------
'535:     For pIndex = 0 To pXMLDOMNodeListDataset.length - 1 Step 1
'536:         Set pXMLDOMNodeDataset = pXMLDOMNodeListDataset.Item(pIndex)
'        Select Case UCase(pXMLDOMNodeDataset.baseName)
'        Case "METADATA"
'            '------------
'            ' Do Nothing
'            '------------
'        Case "FEATUREDATASET"
'            '----------------------------------------
'            ' Create FeatureDataset and sub-datasets
'            '----------------------------------------
'546:             Call ImportGDB_FeatureDataset(pXMLDOMNodeDataset, pWorkspace)
'        Case "OBJECTCLASS"
'            '---------------------------
'            ' Create Table/FeatureClass
'            '---------------------------
'551:             Call ImportGDB_ObjectClass(pXMLDOMNodeDataset, Nothing, pWorkspace)
'        Case Else
'            '------------
'            ' Do Nothing
'            '------------
'556:         End Select
'557:     Next pIndex
'    '
'    Exit Sub
'ErrorHandler:
'561:     Call HandleError(False, "ImportGDB " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ImportGDB_FeatureDataset(ByRef pXMLDOMNodeFeatureDataset As MSXML2.IXMLDOMNode, _
'                                     ByRef pWorkspace As esriGeoDatabase.IWorkspace)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------
'    ' <featureDataset database"" owner="" table="">
'    '-----------------------------------------------
'    Dim pXMLDOMNodeSpatialReference As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeObjectClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeObjectClassList As MSXML2.IXMLDOMNodeList
'    '
'    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
'    Dim pFeatureDataset As esriGeoDatabase.IFeatureDataset
'    Dim pSpatialReference As esriGeometry.ISpatialReference
'    Dim pIndex As Long
'    Dim pName As String
'    '----------
'    ' Get Name
'    '----------
'582:     pName = CStr(pXMLDOMNodeFeatureDataset.Attributes.getNamedItem("table").Text)
'583:     Call frmOutputWindow.AddOutputMessage("Importing FeatureDataset: " & pName, , , vbMagenta)
'    '-----------------
'    ' Already Exists?
'    '-----------------
'587:     Set pFeatureDataset = modCommon.GetDataset(pWorkspace, pName, esriDTFeatureDataset)
'588:     If pFeatureDataset Is Nothing Then
'        '----------------------
'        ' Get SpatialReference
'        '----------------------
'592:         Set pXMLDOMNodeSpatialReference = pXMLDOMNodeFeatureDataset.selectNodes("spatialReference").nextNode
'593:         Set pSpatialReference = ImportGDB_SpatialReference(pXMLDOMNodeSpatialReference)
'        '-----------------------
'        ' Create FeatureDataset
'        '-----------------------
'597:         Set pFeatureWorkspace = pWorkspace
'598:         Set pFeatureDataset = pFeatureWorkspace.CreateFeatureDataset(pName, pSpatialReference)
'599:     Else
'        '-------------------------------
'        ' FeatureDataset Already Exists
'        '-------------------------------
'603:         Call frmOutputWindow.AddOutputMessage("FeatureDataset Already Exists", , , , , , True, 1)
'604:     End If
'605:     If pFeatureDataset Is Nothing Then
'        '-----------------------------------------
'        ' Skip Creation Underlying FeatureClasses
'        '-----------------------------------------
'609:     Else
'        '----------------------------------
'        ' Create Underlying FeatureClasses
'        '----------------------------------
'613:         Set pXMLDOMNodeObjectClassList = pXMLDOMNodeFeatureDataset.selectNodes("objectClass")
'614:         For pIndex = 0 To pXMLDOMNodeObjectClassList.length - 1 Step 1
'615:             Set pXMLDOMNodeObjectClass = pXMLDOMNodeObjectClassList.Item(pIndex)
'616:             Call ImportGDB_ObjectClass(pXMLDOMNodeObjectClass, pFeatureDataset, pWorkspace)
'617:         Next pIndex
'618:     End If
'    '
'    Exit Sub
'ErrorHandler:
'    Select Case Err.Number
'    Case FDO_E_TABLE_INVALID_NAME
'624:         Call frmOutputWindow.AddOutputMessage("Invalid FeatureDataset Name", , , vbRed, , , True, 1)
'625:         Resume Next
'    Case Else
'627:         Call HandleError(False, "ImportGDB_FeatureDataset " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'628:     End Select
'End Sub

'Private Sub ImportGDB_ObjectClass(ByRef pXMLDOMNodeObjectClass As MSXML2.IXMLDOMNode, _
'                                  ByRef pFeatureDataset As esriGeoDatabase.IFeatureDataset, _
'                                  ByRef pWorkspace As esriGeoDatabase.IWorkspace)
'    On Error GoTo ErrorHandler
'    '----------------------------------------
'    '    <objectClass  database="" owner="" table=""
'    '                  aliasName=""
'    '                  esriDatasetType=""
'    '                  esriFeatureType=""
'    '                  oidField=""
'    '                  shapeField=""
'    '                  subtypeField=""
'    '                  defaultSubtypeCode=""
'    '                  modelName=""
'    '                  configKeyword="">
'    '----------------------------------------
'    Dim pXMLDOMNodeIndex As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeIndexList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeFieldList As MSXML2.IXMLDOMNodeList
'    '
'    Dim pXMLDOMNodeDimension As MSXML2.IXMLDOMNode
'    '
'    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pObjectClass As esriGeoDatabase.IObjectClass
'    Dim pClassSchemaEdit As esriGeoDatabase.IClassSchemaEdit
'    Dim pTable As esriGeoDatabase.ITable
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pFieldsEdit As esriGeoDatabase.IFieldsEdit
'    Dim pField As esriGeoDatabase.IField
'    Dim pIndex As esriGeoDatabase.IIndex
'    Dim pIndexEdit As esriGeoDatabase.IIndexEdit
'    '
'    Dim pDimensionClassExtension As esriCarto.IDimensionClassExtension
'    '
'    Dim pName As String
'    Dim pAliasName As String
'    Dim pEsriFeatureType As Long
'    Dim pEsriDatasetType As Long
'    Dim pShapeFieldName As String
'    Dim pSubtypeFieldName As String
'    Dim pDefaultSubtypeCode As String
'    Dim pConfigKeyword As String
'    Dim pModelName As String
'    '
'    Dim pFieldName As String
'    Dim pIndexName As String
'    Dim pFieldIndex As Long
'    Dim pIndexIndex As Long
'    Dim pIndexField As Long
'    Dim pIndexClass As Long
'    '
'    Dim pCLSID As esriSystem.UID
'    Dim pEXTCLSID As esriSystem.UID
'    '
'    Dim pFieldChecker As esriGeoDatabase.IFieldChecker
'    Dim pEnumFieldError As esriGeoDatabase.IEnumFieldError
'    Dim pFixedFields As esriGeoDatabase.IFields
'    Dim pFixedTableName As String
'    Dim pFieldError As esriGeoDatabase.IFieldError
'    Dim pEsriTableNameErrorType As esriGeoDatabase.esriTableNameErrorType
'    '
'    Dim pXMLDOMNodeAnnotation As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListClass As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeAnnotateLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeExtent As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotateLayerTransformationProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLabelEngineLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeBasicOverposterLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLineLabelPlacementPriorities As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLineLabelPosition As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodePointPlacementPriorities As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOverposterLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotationExpressionEngine As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeTextSymbol As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeColor As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFont As MSXML2.IXMLDOMNode
'    '
'    Dim pAnnoClassAdmin As esriCarto.IAnnoClassAdmin
'    Dim pAnnotateLayerPropertiesCollection As esriCarto.IAnnotateLayerPropertiesCollection
'    Dim pAnnotateLayerTransformationProperties  As esriCarto.IAnnotateLayerTransformationProperties
'    Dim pAnnotateLayerProperties As esriCarto.IAnnotateLayerProperties
'    Dim pEnvelope As esriGeometry.IEnvelope
'    Dim pAnnotationExpressionEngine As esriCarto.IAnnotationExpressionEngine
'    Dim pLabelEngineLayerProperties As esriCarto.ILabelEngineLayerProperties
'    Dim pLineLabelPlacementPriorities As esriCarto.ILineLabelPlacementPriorities
'    Dim pLineLabelPosition As esriCarto.ILineLabelPosition
'    Dim pBasicOverposterLayerProperties3 As esriCarto.IBasicOverposterLayerProperties3
'    Dim pPointPlacementPriorities  As esriCarto.IPointPlacementPriorities
'    Dim pOverposterLayerProperties As esriCarto.IOverposterLayerProperties
'    Dim pTextSymbol As esriDisplay.ITextSymbol
'    Dim pColor As esriDisplay.IColor
'    '
'    Dim strAngleList As String
'    Dim varAngleList As Variant
'    '
'729:     Set pFeatureWorkspace = pWorkspace
'    '---------------------------------------------
'    ' Extract the ObjectClass properties from XML
'    '---------------------------------------------
'733:     pName = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("table").Text)
'734:     pConfigKeyword = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("configKeyword").Text)
'735:     pShapeFieldName = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("shapeField").Text)
'736:     pEsriDatasetType = CLng(pXMLDOMNodeObjectClass.Attributes.getNamedItem("esriDatasetType").Text)
'    '
'738:     Call frmOutputWindow.AddOutputMessage("Importing ObjectClass: " & pName, , , vbMagenta)
'    '----------------------------------------
'    ' Get FeatureClass (if already exists?)
'    '----------------------------------------
'742:     Set pObjectClass = modCommon.GetDataset(pWorkspace, pName, esriDTTable)
'743:     If pObjectClass Is Nothing Then
'        '--------------------------------------------------------
'        ' Validation Input Table Name against target Geodatabase
'        '--------------------------------------------------------
'747:         Set pFieldChecker = New FieldChecker
'748:         Set pFieldChecker.ValidateWorkspace = pWorkspace
'749:         pEsriTableNameErrorType = pFieldChecker.ValidateTableName(pName, pFixedTableName)
'750:         If pEsriTableNameErrorType < 1 Then
'            '-----------------
'            ' Table Name Pass
'            '-----------------
'754:             Call frmOutputWindow.AddOutputMessage("Table name passes validator", , , , , , True, 1)
'755:         Else
'756:             Call frmOutputWindow.AddOutputMessage("Table Error encountered in " & pName, , , vbRed, , , True, 1)
'757:             Call frmOutputWindow.AddOutputMessage("Error: " & ReturnTableErrorText(pEsriTableNameErrorType), , , vbRed, , , , 2)
'758:             Call frmOutputWindow.AddOutputMessage("ObjectClass Creation Skipped", , , vbRed, , , , 2)
'            '----------------------------------
'            ' Bail out of ObjectClass creation
'            '----------------------------------
'            Exit Sub
'763:         End If
'        '---------------------------
'        ' Create the Fields Object.
'        '---------------------------
'767:         Set pFields = ImportGDB_Field(pXMLDOMNodeObjectClass.selectNodes("field"))
'        '--------------------------------------------------
'        ' Validate Input Fields against target Geodatabase
'        '--------------------------------------------------
'771:         Call pFieldChecker.Validate(pFields, pEnumFieldError, pFixedFields)
'772:         If pEnumFieldError Is Nothing Then
'            '------------------------
'            ' Fields Pass Validation
'            '------------------------
'776:             Call frmOutputWindow.AddOutputMessage("Field Names pass validator", , , , , , True, 1)
'777:         Else
'            '------------------------------------
'            ' Iterate through the error messages
'            '------------------------------------
'781:             Call frmOutputWindow.AddOutputMessage("Field Error(s) encountered in " & pName & vbCrLf, , , vbRed, , , True, 1)
'782:             Call pEnumFieldError.Reset
'783:             Set pFieldError = pEnumFieldError.Next
'784:             Do Until pFieldError Is Nothing
'785:                 Call frmOutputWindow.AddOutputMessage("[FieldName(" & pFields.Field(pFieldError.FieldIndex).Name & ") Error(" & ReturnFieldErrorText(pFieldError.FieldError) & ")]", , , vbRed, , , , 2)
'786:                 Set pFieldError = pEnumFieldError.Next
'787:             Loop
'788:             Call frmOutputWindow.AddOutputMessage("ObjectClass Creation Skipped", , , vbRed, , , True, 1)
'            '----------------------------------
'            ' Bail out of ObjectClass creation
'            '----------------------------------
'            Exit Sub
'793:         End If
'        '--------------
'        ' Assign GUIDs
'        '--------------
'797:         Set pCLSID = New UID
'798:         Set pEXTCLSID = New UID
'        '
'        Select Case pEsriDatasetType
'        Case esriGeoDatabase.esriDatasetType.esriDTTable
'            '---------------------
'            ' Table (no EXTCLSID)
'            '---------------------
'805:             pCLSID.Value = GUID_TABLE_CLSID
'806:             Set pEXTCLSID = Nothing
'        Case esriGeoDatabase.esriDatasetType.esriDTFeatureClass
'            '--------------
'            ' FeatureClass
'            '--------------
'811:             pEsriFeatureType = CLng(pXMLDOMNodeObjectClass.Attributes.getNamedItem("esriFeatureType").Text)
'            '
'            Select Case pEsriFeatureType
'            Case esriGeoDatabase.esriFeatureType.esriFTAnnotation
'                '-------------------------------------
'                ' Set Annotation Class Extension GUID
'                '-------------------------------------
'818:                 pCLSID.Value = GUID_ANNOTATION_CLSID
'819:                 pEXTCLSID.Value = GUID_ANNOTATION_EXTCLSID
'                '------------------------------------
'                ' Do NOT Add Annotation FeatureClass
'                '------------------------------------
'                Exit Sub
'            Case esriGeoDatabase.esriFeatureType.esriFTDimension
'                '------------------------------------
'                ' Set Dimension Class Extension GUID
'                '------------------------------------
'828:                 pCLSID.Value = GUID_DIMENSION_CLSID
'829:                 pEXTCLSID.Value = GUID_DIMENSION_EXTCLSID
'            Case esriGeoDatabase.esriFeatureType.esriFTComplexEdge, esriGeoDatabase.esriFeatureType.esriFTComplexJunction, esriGeoDatabase.esriFeatureType.esriFTSimpleEdge, esriGeoDatabase.esriFeatureType.esriFTSimpleJunction
'                '--------------------------------------------------
'                ' Downgrade Simple/Complex Edge/Junction to Simple
'                '--------------------------------------------------
'834:                 pEsriFeatureType = esriGeoDatabase.esriFeatureType.esriFTSimple
'835:                 pCLSID.Value = GUID_FEATURECLASS_CLSID
'836:                 Set pEXTCLSID = Nothing
'            Case esriGeoDatabase.esriFeatureType.esriFTSimple
'                '---------------------
'                ' Simple FeatureClass
'                '---------------------
'841:                 pCLSID.Value = GUID_FEATURECLASS_CLSID
'842:                 Set pEXTCLSID = Nothing
'843:             End Select
'844:         End Select
'        '
'        Select Case pEsriDatasetType
'        Case esriGeoDatabase.esriDatasetType.esriDTTable
'            '------------------
'            ' Create the Table
'            '------------------
'851:             Set pTable = pFeatureWorkspace.CreateTable(pName, pFields, pCLSID, pEXTCLSID, pConfigKeyword)
'852:             Set pObjectClass = pTable
'        Case esriGeoDatabase.esriDatasetType.esriDTFeatureClass
'            '-------------------------
'            ' Create the FeatureClass
'            '-------------------------
'857:             If pFeatureDataset Is Nothing Then
'                '-------------------------
'                ' Standalone FeatureClass
'                '-------------------------
'861:                 Set pFeatureClass = pFeatureWorkspace.CreateFeatureClass(pName, pFields, pCLSID, pEXTCLSID, pEsriFeatureType, pShapeFieldName, pConfigKeyword)
'862:             Else
'                '----------------------------------------------
'                ' FeatureClass is a subset of a FeatureDataset
'                '----------------------------------------------
'866:                 Set pFeatureClass = pFeatureDataset.CreateFeatureClass(pName, pFields, pCLSID, pEXTCLSID, pEsriFeatureType, pShapeFieldName, pConfigKeyword)
'867:             End If
'            '--------------------------------------------------
'            ' Set Annotation and Dimension Specific Properties
'            '--------------------------------------------------
'            Select Case pEsriFeatureType
'            Case esriGeoDatabase.esriFeatureType.esriFTAnnotation
'                '----------------------------------------------------------------------
'                ' <annotation referenceScale="" esriUnits="" version="" autoCreate="">
'                '----------------------------------------------------------------------
'876:                 Set pXMLDOMNodeAnnotation = pXMLDOMNodeObjectClass.selectNodes("annotation").nextNode
'877:                 Set pAnnoClassAdmin = pFeatureClass.Extension
'878:                 pAnnoClassAdmin.ReferenceScale = CDbl(pXMLDOMNodeAnnotation.Attributes.getNamedItem("referenceScale").Text)
'879:                 pAnnoClassAdmin.ReferenceScaleUnits = CLng(pXMLDOMNodeAnnotation.Attributes.getNamedItem("esriUnits").Text)
'880:                 pAnnoClassAdmin.AutoCreate = CBool(pXMLDOMNodeAnnotation.Attributes.getNamedItem("autoCreate").Text)
'                '
'882:                 Set pAnnotateLayerPropertiesCollection = New esriCarto.AnnotateLayerPropertiesCollection
'883:                 Set pXMLDOMNodeListClass = pXMLDOMNodeAnnotation.selectNodes("class")
'884:                 For pIndexClass = 0 To pXMLDOMNodeListClass.length - 1 Step 1
'                    '----------------
'                    ' <class name=""
'                    '----------------
'888:                     Set pXMLDOMNodeClass = pXMLDOMNodeListClass.Item(pIndexClass)
'                    '------------------------------------------------------------
'                    ' <annotateLayerProperties addUnplacedToGraphicsContainer=""
'                    '                          annotationMinimumScale=""
'                    '                          annotationMaximumScale=""
'                    '                          createUnplacedElements=""
'                    '                          displayAnnotation=""
'                    '                          esriLabelWhichFeatures=""
'                    '                          featureLinked=""
'                    '                          priority=""
'                    '                          useOutput=""
'                    '                          whereClause="">
'                    '------------------------------------------------------------
'901:                     Set pAnnotateLayerProperties = New esriCarto.LabelEngineLayerProperties
'902:                     pAnnotateLayerProperties.Class = CStr(pXMLDOMNodeClass.Attributes.getNamedItem("name").Text)
'                    '
'904:                     If pXMLDOMNodeClass.selectNodes("annotateLayerProperties").length = 1 Then
'905:                         Set pXMLDOMNodeAnnotateLayerProperties = pXMLDOMNodeClass.selectSingleNode("annotateLayerProperties")
'906:                         pAnnotateLayerProperties.AddUnplacedToGraphicsContainer = CBool(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("addUnplacedToGraphicsContainer").Text)
'907:                         pAnnotateLayerProperties.AnnotationMinimumScale = CDbl(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("annotationMinimumScale").Text)
'908:                         pAnnotateLayerProperties.AnnotationMaximumScale = CDbl(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("annotationMaximumScale").Text)
'909:                         pAnnotateLayerProperties.CreateUnplacedElements = CBool(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("createUnplacedElements").Text)
'910:                         pAnnotateLayerProperties.DisplayAnnotation = CBool(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("displayAnnotation").Text)
'911:                         pAnnotateLayerProperties.LabelWhichFeatures = CLng(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("esriLabelWhichFeatures").Text)
'912:                         pAnnotateLayerProperties.FeatureLinked = CBool(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("featureLinked").Text)
'913:                         pAnnotateLayerProperties.Priority = CLng(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("priority").Text)
'914:                         pAnnotateLayerProperties.UseOutput = CBool(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("useOutput").Text)
'915:                         pAnnotateLayerProperties.WhereClause = CStr(pXMLDOMNodeAnnotateLayerProperties.Attributes.getNamedItem("whereClause").Text)
'916:                         If pXMLDOMNodeAnnotateLayerProperties.selectNodes("extent").length = 1 Then
'917:                             Set pXMLDOMNodeExtent = pXMLDOMNodeAnnotateLayerProperties.selectSingleNode("extent")
'918:                             Set pEnvelope = NodeEnvelope(pXMLDOMNodeExtent)
'919:                             pAnnotateLayerProperties.Extent = pEnvelope
'920:                         End If
'921:                     End If
'                    '--------------------------------------------------------------------------------------
'                    ' <annotateLayerTransformationProperties referenceScale="" scaleRatio="" esriUnits="">
'                    '     <extent xMin="" xMax="" yMin="" yMax=""/>
'                    ' </annotateLayerTransformationProperties>
'                    '--------------------------------------------------------------------------------------
'927:                     If pXMLDOMNodeClass.selectNodes("annotateLayerTransformationProperties").length = 1 Then
'928:                         Set pAnnotateLayerTransformationProperties = pAnnotateLayerProperties
'929:                         Set pXMLDOMNodeAnnotateLayerTransformationProperties = pXMLDOMNodeClass.selectSingleNode("annotateLayerTransformationProperties")
'930:                         pAnnotateLayerTransformationProperties.ReferenceScale = CDbl(pXMLDOMNodeAnnotateLayerTransformationProperties.Attributes.getNamedItem("referenceScale").Text)
'931:                         pAnnotateLayerTransformationProperties.ScaleRatio = CDbl(pXMLDOMNodeAnnotateLayerTransformationProperties.Attributes.getNamedItem("scaleRatio").Text)
'932:                         pAnnotateLayerTransformationProperties.Units = CLng(pXMLDOMNodeAnnotateLayerTransformationProperties.Attributes.getNamedItem("esriUnits").Text)
'933:                         If pXMLDOMNodeAnnotateLayerTransformationProperties.selectNodes("extent").length = 1 Then
'934:                             Set pXMLDOMNodeExtent = pXMLDOMNodeAnnotateLayerTransformationProperties.selectSingleNode("extent")
'935:                             Set pEnvelope = NodeEnvelope(pXMLDOMNodeExtent)
'936:                             pAnnotateLayerTransformationProperties.Bounds = pEnvelope
'937:                         End If
'938:                     End If
'                    '----------------------------------------------------------------------------------------
'                    ' <labelEngineLayerProperties expression="" isExpressionSimple="" offset="" symbolId="">
'                    '----------------------------------------------------------------------------------------
'942:                     If pXMLDOMNodeClass.selectNodes("labelEngineLayerProperties").length = 1 Then
'943:                         Set pXMLDOMNodeLabelEngineLayerProperties = pXMLDOMNodeClass.selectSingleNode("labelEngineLayerProperties")
'944:                         Set pLabelEngineLayerProperties = pAnnotateLayerProperties
'945:                         pLabelEngineLayerProperties.Expression = CStr(pXMLDOMNodeLabelEngineLayerProperties.Attributes.getNamedItem("expression").Text)
'946:                         pLabelEngineLayerProperties.IsExpressionSimple = CBool(pXMLDOMNodeLabelEngineLayerProperties.Attributes.getNamedItem("isExpressionSimple").Text)
'947:                         pLabelEngineLayerProperties.Offset = CDbl(pXMLDOMNodeLabelEngineLayerProperties.Attributes.getNamedItem("offset").Text)
'948:                         pLabelEngineLayerProperties.SymbolID = CLng(pXMLDOMNodeLabelEngineLayerProperties.Attributes.getNamedItem("symbolId").Text)
'                        '------------------------------------------------------------------------
'                        ' <basicOverposterLayerProperties bufferRatio=""
'                        '                                 esriBasicOverposterFeatureType=""
'                        '                                 featureWeight=""
'                        '                                 generateUnplacedLabels=""
'                        '                                 labelWeight=""
'                        '                                 lineOffset=""
'                        '                                 maxDistanceFromTarget=""
'                        '                                 esriBasicNumLabelsOption=""
'                        '                                 perpendicularToAngle=""
'                        '                                 pointPlacementAngles=""
'                        '                                 esriOverposterPointPlacementMethod=""
'                        '                                 pointPlacementOnTop=""
'                        '                                 rotationField=""
'                        '                                 esriLabelRotationType="">
'                        '------------------------------------------------------------------------
'965:                         Set pBasicOverposterLayerProperties3 = New esriCarto.BasicOverposterLayerProperties
'966:                         If pXMLDOMNodeLabelEngineLayerProperties.selectNodes("basicOverposterLayerProperties").length = 1 Then
'967:                             Set pXMLDOMNodeBasicOverposterLayerProperties = pXMLDOMNodeLabelEngineLayerProperties.selectSingleNode("basicOverposterLayerProperties")
'968:                             pBasicOverposterLayerProperties3.BufferRatio = CDbl(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("bufferRatio").Text)
'969:                             pBasicOverposterLayerProperties3.FeatureType = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("esriBasicOverposterFeatureType").Text)
'970:                             pBasicOverposterLayerProperties3.FeatureWeight = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("featureWeight").Text)
'971:                             pBasicOverposterLayerProperties3.GenerateUnplacedLabels = CBool(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("generateUnplacedLabels").Text)
'972:                             pBasicOverposterLayerProperties3.LabelWeight = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("labelWeight").Text)
'973:                             pBasicOverposterLayerProperties3.LineOffset = CDbl(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("lineOffset").Text)
'974:                             pBasicOverposterLayerProperties3.MaxDistanceFromTarget = CDbl(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("maxDistanceFromTarget").Text)
'975:                             pBasicOverposterLayerProperties3.NumLabelsOption = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("esriBasicNumLabelsOption").Text)
'976:                             pBasicOverposterLayerProperties3.PerpendicularToAngle = CBool(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("perpendicularToAngle").Text)
'977:                             strAngleList = CStr(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("pointPlacementAngles").Text)
'978:                             If strAngleList <> "" Then
'979:                                 varAngleList = Split(strAngleList, ",")
'980:                                 pBasicOverposterLayerProperties3.PointPlacementAngles = varAngleList
'981:                             End If
'982:                             pBasicOverposterLayerProperties3.PointPlacementMethod = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("esriOverposterPointPlacementMethod").Text)
'983:                             pBasicOverposterLayerProperties3.PointPlacementOnTop = CBool(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("pointPlacementOnTop").Text)
'984:                             pBasicOverposterLayerProperties3.RotationField = CStr(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("rotationField").Text)
'985:                             pBasicOverposterLayerProperties3.RotationType = CLng(pXMLDOMNodeBasicOverposterLayerProperties.Attributes.getNamedItem("esriLabelRotationType").Text)
'                            '------------------------------------------------
'                            ' <lineLabelPlacementPriorities aboveAfter=""
'                            '                               aboveAlong=""
'                            '                               aboveEnd=""
'                            '                               aboveStart=""
'                            '                               centerAfter=""
'                            '                               centerAlong=""
'                            '                               centerEnd=""
'                            '                               centerStart=""
'                            '                               belowAfter=""
'                            '                               belowAlong=""
'                            '                               belowEnd=""
'                            '                               belowStart="" />
'                            '------------------------------------------------
'1000:                             If pXMLDOMNodeBasicOverposterLayerProperties.selectNodes("lineLabelPlacementPriorities").length = 1 Then
'1001:                                 Set pXMLDOMNodeLineLabelPlacementPriorities = pXMLDOMNodeBasicOverposterLayerProperties.selectSingleNode("lineLabelPlacementPriorities")
'1002:                                 Set pLineLabelPlacementPriorities = New esriCarto.LineLabelPlacementPriorities
'1003:                                 pLineLabelPlacementPriorities.AboveAfter = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("aboveAfter").Text)
'1004:                                 pLineLabelPlacementPriorities.AboveAlong = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("aboveAlong").Text)
'1005:                                 pLineLabelPlacementPriorities.AboveEnd = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("aboveEnd").Text)
'1006:                                 pLineLabelPlacementPriorities.AboveStart = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("aboveStart").Text)
'1007:                                 pLineLabelPlacementPriorities.CenterAfter = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("centerAfter").Text)
'1008:                                 pLineLabelPlacementPriorities.CenterAlong = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("centerAlong").Text)
'1009:                                 pLineLabelPlacementPriorities.CenterEnd = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("centerEnd").Text)
'1010:                                 pLineLabelPlacementPriorities.CenterStart = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("centerStart").Text)
'1011:                                 pLineLabelPlacementPriorities.BelowAfter = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("belowAfter").Text)
'1012:                                 pLineLabelPlacementPriorities.BelowAlong = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("belowAlong").Text)
'1013:                                 pLineLabelPlacementPriorities.BelowEnd = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("belowEnd").Text)
'1014:                                 pLineLabelPlacementPriorities.BelowStart = CLng(pXMLDOMNodeLineLabelPlacementPriorities.Attributes.getNamedItem("belowStart").Text)
'                                '
'1016:                                 pBasicOverposterLayerProperties3.LineLabelPlacementPriorities = pLineLabelPlacementPriorities
'1017:                             End If
'                            '--------------------------------------------
'                            ' <lineLabelPosition above=""
'                            '                    atEnd=""
'                            '                    atStart=""
'                            '                    below=""
'                            '                    horizontal=""
'                            '                    inLine=""
'                            '                    left=""
'                            '                    offset=""
'                            '                    onTop=""
'                            '                    parallel=""
'                            '                    perpendicular=""
'                            '                    produceCurvedLabels=""
'                            '                    right="" />
'                            '--------------------------------------------
'1033:                             If pXMLDOMNodeBasicOverposterLayerProperties.selectNodes("lineLabelPosition").length = 1 Then
'1034:                                 Set pXMLDOMNodeLineLabelPosition = pXMLDOMNodeBasicOverposterLayerProperties.selectSingleNode("lineLabelPosition")
'1035:                                 Set pLineLabelPosition = New esriCarto.LineLabelPosition
'1036:                                 pLineLabelPosition.Above = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("above").Text)
'1037:                                 pLineLabelPosition.AtEnd = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("atEnd").Text)
'1038:                                 pLineLabelPosition.AtStart = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("atStart").Text)
'1039:                                 pLineLabelPosition.Below = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("below").Text)
'1040:                                 pLineLabelPosition.Horizontal = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("horizontal").Text)
'1041:                                 pLineLabelPosition.InLine = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("inLine").Text)
'1042:                                 pLineLabelPosition.Left = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("left").Text)
'1043:                                 pLineLabelPosition.Offset = CDbl(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("offset").Text)
'1044:                                 pLineLabelPosition.OnTop = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("onTop").Text)
'1045:                                 pLineLabelPosition.Parallel = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("parallel").Text)
'1046:                                 pLineLabelPosition.Perpendicular = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("perpendicular").Text)
'1047:                                 pLineLabelPosition.ProduceCurvedLabels = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("produceCurvedLabels").Text)
'1048:                                 pLineLabelPosition.Right = CBool(pXMLDOMNodeLineLabelPosition.Attributes.getNamedItem("right").Text)
'                                '
'1050:                                 pBasicOverposterLayerProperties3.LineLabelPosition = pLineLabelPosition
'1051:                             End If
'                            '----------------------------------------------
'                            ' <pointPlacementPriorities aboveCenter=""
'                            '                           aboveLeft=""
'                            '                           aboveRight=""
'                            '                           belowCenter=""
'                            '                           belowLeft=""
'                            '                           belowRight=""
'                            '                           centerLeft=""
'                            '                           centerRight="" />
'                            '----------------------------------------------
'1062:                             If pXMLDOMNodeBasicOverposterLayerProperties.selectNodes("pointPlacementPriorities").length = 1 Then
'1063:                                 Set pXMLDOMNodePointPlacementPriorities = pXMLDOMNodeBasicOverposterLayerProperties.selectSingleNode("pointPlacementPriorities")
'1064:                                 Set pPointPlacementPriorities = New esriCarto.PointPlacementPriorities
'1065:                                 pPointPlacementPriorities.AboveCenter = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("aboveCenter").Text)
'1066:                                 pPointPlacementPriorities.AboveLeft = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("aboveLeft").Text)
'1067:                                 pPointPlacementPriorities.AboveRight = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("aboveRight").Text)
'1068:                                 pPointPlacementPriorities.BelowCenter = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("belowCenter").Text)
'1069:                                 pPointPlacementPriorities.BelowLeft = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("belowLeft").Text)
'1070:                                 pPointPlacementPriorities.BelowRight = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("belowRight").Text)
'1071:                                 pPointPlacementPriorities.CenterLeft = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("centerLeft").Text)
'1072:                                 pPointPlacementPriorities.CenterRight = CLng(pXMLDOMNodePointPlacementPriorities.Attributes.getNamedItem("centerRight").Text)
'                                '
'1074:                                 pBasicOverposterLayerProperties3.PointPlacementPriorities = pPointPlacementPriorities
'1075:                             End If
'1076:                         End If
'                        '---------------------------------------------------------------------------
'                        ' <overposterLayerProperties isBarrier="" placeLabels="" placeSymbols="" />
'                        '---------------------------------------------------------------------------
'1080:                         If pXMLDOMNodeLabelEngineLayerProperties.selectNodes("overposterLayerProperties").length = 1 Then
'1081:                             Set pXMLDOMNodeOverposterLayerProperties = pXMLDOMNodeLabelEngineLayerProperties.selectSingleNode("overposterLayerProperties")
'1082:                             Set pOverposterLayerProperties = pBasicOverposterLayerProperties3
'1083:                             pOverposterLayerProperties.IsBarrier = CBool(pXMLDOMNodeOverposterLayerProperties.Attributes.getNamedItem("isBarrier").Text)
'1084:                             pOverposterLayerProperties.PlaceLabels = CBool(pXMLDOMNodeOverposterLayerProperties.Attributes.getNamedItem("placeLabels").Text)
'1085:                             pOverposterLayerProperties.PlaceSymbols = CBool(pXMLDOMNodeOverposterLayerProperties.Attributes.getNamedItem("placeSymbols").Text)
'1086:                         End If
'                        '
'1088:                         If Not (pBasicOverposterLayerProperties3 Is Nothing) Then
'1089:                             Set pLabelEngineLayerProperties.BasicOverposterLayerProperties = pBasicOverposterLayerProperties3
'1090:                         End If
'                        '------------------------------------------------------
'                        ' <annotationExpressionEngine name="" appendCode="" />
'                        '------------------------------------------------------
'1094:                         If pXMLDOMNodeLabelEngineLayerProperties.selectNodes("annotationExpressionEngine").length = 1 Then
'1095:                             Set pXMLDOMNodeAnnotationExpressionEngine = pXMLDOMNodeLabelEngineLayerProperties.selectSingleNode("annotationExpressionEngine")
'                            Select Case UCase(CStr(pXMLDOMNodeAnnotationExpressionEngine.Attributes.getNamedItem("name").Text))
'                            Case "VB SCRIPT"
'1098:                                 Set pAnnotationExpressionEngine = New esriCarto.AnnotationVBScriptEngine
'                            Case "JAVA SCRIPT"
'1100:                                 Set pAnnotationExpressionEngine = New esriCarto.AnnotationJScriptEngine
'                            Case Else
'                                '
'1103:                             End Select
'                            '
'1105:                             If Not (pAnnotationExpressionEngine Is Nothing) Then
'1106:                                 Set pLabelEngineLayerProperties.ExpressionParser = pAnnotationExpressionEngine
'1107:                             End If
'1108:                         End If
'                        '---------------------------------------------
'                        ' <textSymbol angle=""
'                        '             esriTextHorizontalAlignment=""
'                        '             rightToLeft=""
'                        '             size=""
'                        '             text=""
'                        '             esriTextVerticalAlignment="">
'                        '---------------------------------------------
'1117:                         If pXMLDOMNodeLabelEngineLayerProperties.selectNodes("textSymbol").length = 1 Then
'1118:                             Set pXMLDOMNodeTextSymbol = pXMLDOMNodeLabelEngineLayerProperties.selectSingleNode("textSymbol")
'1119:                             Set pTextSymbol = New esriDisplay.TextSymbol
'1120:                             pTextSymbol.Angle = CDbl(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("angle").Text)
'1121:                             pTextSymbol.HorizontalAlignment = CLng(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("esriTextHorizontalAlignment").Text)
'1122:                             pTextSymbol.RightToLeft = CBool(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("rightToLeft").Text)
'1123:                             pTextSymbol.Size = CDbl(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("size").Text)
'1124:                             pTextSymbol.Text = CStr(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("text").Text)
'1125:                             pTextSymbol.VerticalAlignment = CLng(pXMLDOMNodeTextSymbol.Attributes.getNamedItem("esriTextVerticalAlignment").Text)
'                            '------------------
'                            ' <color rgb="" />
'                            '------------------
'1129:                             If pXMLDOMNodeTextSymbol.selectNodes("color").length = 1 Then
'1130:                                 Set pXMLDOMNodeColor = pXMLDOMNodeTextSymbol.selectSingleNode("color")
'1131:                                 Set pColor = New esriDisplay.RgbColor
'1132:                                 pColor.RGB = CLng(pXMLDOMNodeColor.Attributes.getNamedItem("rgb").Text)
'                                '
'1134:                                 pTextSymbol.Color = pColor
'1135:                             End If
'                            '-------------------------
'                            ' <font name=""
'                            '       bold=""
'                            '       charset=""
'                            '       italic=""
'                            '       size=""
'                            '       strikeThrough=""
'                            '       underline=""
'                            '       weight="" />
'                            '-------------------------
'1146:                             If pXMLDOMNodeTextSymbol.selectNodes("font").length = 1 Then
'1147:                                 Set pXMLDOMNodeFont = pXMLDOMNodeTextSymbol.selectSingleNode("font")
'1148:                                 pTextSymbol.Font.Name = CStr(pXMLDOMNodeFont.Attributes.getNamedItem("name").Text)
'1149:                                 pTextSymbol.Font.Bold = CBool(pXMLDOMNodeFont.Attributes.getNamedItem("bold").Text)
'1150:                                 pTextSymbol.Font.Charset = CInt(pXMLDOMNodeFont.Attributes.getNamedItem("charset").Text)
'1151:                                 pTextSymbol.Font.Italic = CBool(pXMLDOMNodeFont.Attributes.getNamedItem("italic").Text)
'1152:                                 pTextSymbol.Font.Size = CInt(pXMLDOMNodeFont.Attributes.getNamedItem("size").Text)
'1153:                                 pTextSymbol.Font.Strikethrough = CBool(pXMLDOMNodeFont.Attributes.getNamedItem("strikeThrough").Text)
'1154:                                 pTextSymbol.Font.Underline = CBool(pXMLDOMNodeFont.Attributes.getNamedItem("underline").Text)
'1155:                                 pTextSymbol.Font.Weight = CInt(pXMLDOMNodeFont.Attributes.getNamedItem("weight").Text)
'1156:                             End If
'                            '
'1158:                             Set pLabelEngineLayerProperties.Symbol = pTextSymbol
'1159:                         End If
'1160:                     End If
'                    '--------------------------------------------------------------------------------
'                    ' Last Step: Add LabelEngineLayerProperties to AnnotateLayerPropertiesCollection
'                    '--------------------------------------------------------------------------------
'1164:                     Call pAnnotateLayerPropertiesCollection.Add(pAnnotateLayerProperties)
'1165:                 Next pIndexClass
'                '--------------------------------------------------------------------------
'                ' Assign Annotation Layer Properties Collection to Annotation FeatureClass
'                '--------------------------------------------------------------------------
'1169:                 If pAnnotateLayerPropertiesCollection.Count > 0 Then
'1170:                     pAnnoClassAdmin.AnnoProperties = pAnnotateLayerPropertiesCollection
'1171:                     Call pAnnoClassAdmin.UpdateProperties
'1172:                 End If
'            Case esriGeoDatabase.esriFeatureType.esriFTDimension
'                '---------------------------------------------
'                ' <dimension referenceScale="" esriUnits="">
'                '---------------------------------------------
'1177:                 Set pXMLDOMNodeDimension = pXMLDOMNodeObjectClass.selectNodes("dimension").nextNode
'1178:                 Set pDimensionClassExtension = pFeatureClass.Extension
'1179:                 pDimensionClassExtension.ReferenceScale = CDbl(pXMLDOMNodeDimension.Attributes.getNamedItem("referenceScale").Text)
'1180:                 pDimensionClassExtension.ReferenceScaleUnits = CLng(pXMLDOMNodeDimension.Attributes.getNamedItem("esriUnits").Text)
'1181:             End Select
'            '-------------------------------------------------------------
'            ' Set the IObjectClass interface for downstream compatiablity
'            '-------------------------------------------------------------
'1185:             Set pObjectClass = pFeatureClass
'1186:         End Select
'        '------------
'        ' AliasName
'        '------------
'1190:         Set pClassSchemaEdit = pObjectClass
'1191:         pAliasName = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("aliasName").Text)
'1192:         Call pClassSchemaEdit.AlterAliasName(pAliasName)
'        '-----------
'        ' ModelInfo
'        '-----------
'1196:         pModelName = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("modelName").Text)
'1197:         Call pClassSchemaEdit.AlterModelName(pModelName)
'        '-----------
'        ' Subtytpes
'        '-----------
'1201:         pSubtypeFieldName = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("subtypeField").Text)
'1202:         pDefaultSubtypeCode = CStr(pXMLDOMNodeObjectClass.Attributes.getNamedItem("defaultSubtypeCode").Text)
'        '
'1204:         Call ImportGDB_Subtype(pXMLDOMNodeObjectClass.selectNodes("subtype"), _
'                               pObjectClass, _
'                               pSubtypeFieldName, _
'                               pDefaultSubtypeCode)
'        '-------------------
'        ' Add Field Indexes
'        '-------------------
'1211:         If mOCImportFieldIndex Then
'1212:             Set pXMLDOMNodeIndexList = pXMLDOMNodeObjectClass.selectNodes("index")
'            '
'1214:             If pXMLDOMNodeIndexList.length > 0 Then
'1215:                 For pIndexIndex = 0 To pXMLDOMNodeIndexList.length - 1 Step 1
'1216:                     Set pXMLDOMNodeIndex = pXMLDOMNodeIndexList(pIndexIndex)
'                    '
'1218:                     pIndexName = CStr(pXMLDOMNodeIndex.Attributes.getNamedItem("name").Text)
'                    '
'1220:                     Call frmOutputWindow.AddOutputMessage("Adding Index: " & pIndexName)
'1221:                     If IndexExist(pObjectClass, pIndexName) Then
'                        '----------------------------
'                        ' Index Already Exists! Skip
'                        '----------------------------
'1225:                     Else
'                        '---------------------------
'                        ' Index Does Not Exist. Add
'                        '---------------------------
'1229:                         Set pXMLDOMNodeFieldList = pXMLDOMNodeIndex.selectNodes("field")
'1230:                         If pXMLDOMNodeFieldList.length > 0 Then
'                            '--------------------------------------------
'                            ' Only Add Indexes with one or field assigned
'                            '--------------------------------------------
'1234:                             Set pFields = New esriGeoDatabase.fields
'1235:                             Set pFieldsEdit = pFields
'1236:                             pFieldsEdit.FieldCount = pXMLDOMNodeFieldList.length
'                            '
'1238:                             For pIndexField = 0 To pXMLDOMNodeFieldList.length - 1 Step 1
'1239:                                 Set pXMLDOMNodeField = pXMLDOMNodeFieldList.Item(pIndexField)
'1240:                                 pFieldName = CStr(pXMLDOMNodeField.Attributes.getNamedItem("name").Text)
'1241:                                 pFieldIndex = pObjectClass.FindField(pFieldName)
'1242:                                 If pFieldIndex = -1 Then
'                                    '-------------------------
'                                    ' ERROR Cannot Find Field
'                                    '-------------------------
'1246:                                     Set pFields = Nothing
'1247:                                     Exit For
'1248:                                 Else
'1249:                                     Set pField = pObjectClass.fields.Field(pFieldIndex)
'                                    Select Case CLng(pField.Type)
'                                    Case esriGeoDatabase.esriFieldType.esriFieldTypeGeometry, esriGeoDatabase.esriFieldType.esriFieldTypeOID, esriGeoDatabase.esriFieldType.esriFieldTypeBlob
'                                        '------------------------------------------------------------------------
'                                        ' ERROR: Cannot create an index of using a OID, Shape or Blob field type
'                                        '------------------------------------------------------------------------
'1255:                                         Set pFields = Nothing
'1256:                                         Exit For
'                                    Case Else
'                                        '---------------------
'                                        ' Add Field to Fields
'                                        '---------------------
'1261:                                         Set pFieldsEdit.Field(pIndexField) = pField
'1262:                                     End Select
'1263:                                 End If
'1264:                             Next pIndexField
'                            '
'1266:                             If pFields Is Nothing Then
'                                '----------------------------------
'                                ' Error Occurred, do not add index
'                                '----------------------------------
'1270:                             Else
'1271:                                 Set pIndex = New esriGeoDatabase.Index
'1272:                                 Set pIndexEdit = pIndex
'1273:                                 Set pIndexEdit.fields = pFields
'1274:                                 pIndexEdit.Name = pIndexName
'1275:                                 pIndexEdit.IsAscending = CBool(pXMLDOMNodeIndex.Attributes.getNamedItem("isAscending").Text)
'1276:                                 pIndexEdit.IsUnique = CBool(pXMLDOMNodeIndex.Attributes.getNamedItem("isUnique").Text)
'                                '
'1278:                                 Call pObjectClass.AddIndex(pIndex)
'1279:                             End If
'1280:                         End If
'1281:                     End If
'1282:                 Next pIndexIndex
'1283:             End If
'1284:         End If
'1285:     Else
'        '-----------------------------
'        ' FeatureClass Already Exists
'        '-----------------------------
'1289:         Call frmOutputWindow.AddOutputMessage("FeatureClass Already Exists - Skipping", , , , , , True, 1)
'1290:     End If
'    '
'    Exit Sub
'ErrorHandler:
'    Select Case Err.Number
'    Case FDO_E_OBJECTCLASS_MODEL_NAME_ALREADY_EXISTS
'1296:         Call frmOutputWindow.AddOutputMessage("Another class with the specified model name already exists", , , vbRed, , , True, 1)
'1297:         Resume Next
'    Case Else
'1299:         Call HandleError(False, "ImportGDB_ObjectClass " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'1300:     End Select
'End Sub

'Private Function IndexExist(ByRef pObjectClass As esriGeoDatabase.IObjectClass, ByRef pIndexName As String) As Boolean
'    On Error GoTo ErrorHandler
'    '
'    Dim pIndexes As esriGeoDatabase.IIndexes
'    Dim pIndex As esriGeoDatabase.IIndex
'    '
'    Dim pIndexIndex As Long
'    Dim pIndexExist As Boolean
'    '
'1312:     pIndexExist = False
'    '
'1314:     Set pIndexes = pObjectClass.Indexes
'    '
'1316:     For pIndexIndex = 0 To pIndexes.IndexCount - 1 Step 1
'1317:         Set pIndex = pIndexes.Index(pIndexIndex)
'1318:         If UCase(pIndex.Name) = UCase(pIndexName) Then
'1319:             pIndexExist = True
'1320:             Exit For
'1321:         End If
'1322:     Next pIndexIndex
'    '
'1324:     IndexExist = pIndexExist
'    '
'    Exit Function
'ErrorHandler:
'1328:     Call HandleError(False, "FindIndexPosition " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Function ImportGDB_Field(ByRef pXMLDOMNodeListField As MSXML2.IXMLDOMNodeList) As esriGeoDatabase.IFields
'    On Error GoTo ErrorHandler
'    '-----------------------------------------
'    '            <field name=""
'    '                   aliasName=""
'    '                   domainFixed = ""
'    '                   editable = ""
'    '                   isNullable = ""
'    '                   length = ""
'    '                   precision = ""
'    '                   required = ""
'    '                   scale=""
'    '                   esriFieldType = ""
'    '                   modelName=""/>
'    '-----------------------------------------
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeGeometryDef As MSXML2.IXMLDOMNode
'    '
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pFieldsEdit As esriGeoDatabase.IFieldsEdit
'    Dim pField As esriGeoDatabase.IField
'    Dim pIndex As Long
'    '
'1354:     Set pFields = New esriGeoDatabase.fields
'1355:     Set pFieldsEdit = pFields
'    '
'1357:     pFieldsEdit.FieldCount = pXMLDOMNodeListField.length
'    '
'1359:     For pIndex = 0 To pXMLDOMNodeListField.length - 1 Step 1
'        '----------------
'        ' Get Field Node
'        '----------------
'1363:         Set pXMLDOMNodeField = pXMLDOMNodeListField.Item(pIndex)
'        '-------------------
'        ' Create new IField
'        '-------------------
'1367:         Set pField = modImportExport.MakeField(pXMLDOMNodeField)
'        '------------------------------------
'        ' Add Field to the Fields Collection
'        '------------------------------------
'1371:         Set pFieldsEdit.Field(pIndex) = pField
'1372:     Next pIndex
'    '
'1374:     Set ImportGDB_Field = pFields
'    '
'    Exit Function
'ErrorHandler:
'1378:     Call HandleError(False, "ImportGDB_Field " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Sub ImportGDB_Subtype(ByRef pXMLDOMNodeListSubtype As MSXML2.IXMLDOMNodeList, _
'                              ByRef pObjectClass As esriGeoDatabase.IObjectClass, _
'                              ByRef pSubtypeFieldName As String, _
'                              ByRef pSubtypeDefaultCode As String)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------------
'    '            <subtype name="" code="">
'    '               <field name="" defaultValue="" domain=""/>
'    '            </subtype>
'    '-----------------------------------------------------------
'    Dim pXMLDOMNodeSubtype As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListField As MSXML2.IXMLDOMNodeList
'    '
'    Dim pIndexSubtype As Long
'    Dim pIndexField As Long
'    Dim pSubtypes As esriGeoDatabase.ISubtypes
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pWorkspaceDomain As esriGeoDatabase.IWorkspaceDomains
'    Dim pDomain As esriGeoDatabase.IDomain
'    Dim pField As esriGeoDatabase.IField
'    Dim pFieldEdit As esriGeoDatabase.IFieldEdit
'    Dim pFieldIndex As Long
'    Dim pDomainFixed As Boolean
'    '
'    Dim pSubtypeName As String
'    Dim pSubtypeCode As Long
'    Dim pFieldName As String
'    Dim pDefaultValue As Variant
'    Dim pDomainName As String
'    '
'1412:     Set pDataset = pObjectClass
'1413:     Set pSubtypes = pObjectClass
'1414:     Set pWorkspaceDomain = pDataset.Workspace
'    '----------------------------
'    ' Loop for each Subtype Node
'    '----------------------------
'1418:     For pIndexSubtype = 0 To pXMLDOMNodeListSubtype.length - 1 Step 1
'1419:         Set pXMLDOMNodeSubtype = pXMLDOMNodeListSubtype.Item(pIndexSubtype)
'        '--------------------------------
'        ' Get Subtype Name (Description)
'        '--------------------------------
'1423:         pSubtypeName = CStr(pXMLDOMNodeSubtype.Attributes.getNamedItem("name").Text)
'1424:         If pSubtypeName = "" Then
'            '--------------------------------------------
'            ' Base ObjectClass Domains and DefaultValues
'            '--------------------------------------------
'1428:             Call frmOutputWindow.AddOutputMessage("Assigning Domains and Default Values to Base ObjectClass", , , , , , True, 1)
'1429:         Else
'            '----------------------
'            ' Assign Subtype Field
'            '----------------------
'1433:             If pSubtypes.SubtypeFieldName = "" Then
'1434:                 pSubtypes.SubtypeFieldName = pSubtypeFieldName
'1435:             End If
'            '------------------
'            ' Get Subtype Code
'            '------------------
'1439:             pSubtypeCode = CLng(pXMLDOMNodeSubtype.Attributes.getNamedItem("code").Text)
'            '-----------------
'            ' Add new Subtype
'            '-----------------
'1443:             Call frmOutputWindow.AddOutputMessage("Importing Subtype: " & pSubtypeName, , , , , , True, 1)
'1444:             Call pSubtypes.AddSubtype(pSubtypeCode, pSubtypeName)
'1445:         End If
'1446:         Set pXMLDOMNodeListField = pXMLDOMNodeSubtype.selectNodes("field")
'        '--------------------------
'        ' Loop for each Field Node
'        '--------------------------
'1450:         For pIndexField = 0 To pXMLDOMNodeListField.length - 1 Step 1
'1451:             Set pXMLDOMNodeField = pXMLDOMNodeListField.Item(pIndexField)
'            '
'1453:             pFieldName = CStr(pXMLDOMNodeField.Attributes.getNamedItem("name").Text)
'1454:             pDefaultValue = CStr(pXMLDOMNodeField.Attributes.getNamedItem("defaultValue").Text)
'1455:             pDomainName = CStr(pXMLDOMNodeField.Attributes.getNamedItem("domain").Text)
'            '
'1457:             Call frmOutputWindow.AddOutputMessage("Assigning Domain [" & pDomainName & "] and DefaultValue [" & pDefaultValue & "] to Field [" & pFieldName & "]", , , , , , True, 2)
'            '----------------------
'            ' Get the Field Object
'            '----------------------
'1461:             pFieldIndex = pObjectClass.FindField(pFieldName)
'1462:             Set pField = pObjectClass.fields.Field(pFieldIndex)
'1463:             Set pFieldEdit = pField
'            '---------------
'            ' Assign Domain
'            '---------------
'1467:             Set pDomain = Nothing
'1468:             If pDomainName <> "" Then
'                '------------
'                ' Get Domain
'                '------------
'1472:                 Set pDomain = pWorkspaceDomain.DomainByName(pDomainName)
'1473:                 If Not (pDomain Is Nothing) Then
'                    '-------------------------
'                    ' De-Activate DOMAINFIXED
'                    '-------------------------
'1477:                     pDomainFixed = False
'1478:                     If pField.DomainFixed Then
'1479:                         pFieldEdit.DomainFixed = False
'1480:                         pDomainFixed = True
'1481:                     End If
'                    '-------------------------------------------
'                    ' Assign Domain to Field (or Subtype/Field)
'                    '-------------------------------------------
'1485:                     If pSubtypeName = "" Then
'1486:                         Set pFieldEdit.Domain = pDomain
'1487:                     Else
'1488:                         Set pSubtypes.Domain(pSubtypeCode, pFieldName) = pDomain
'1489:                     End If
'                    '-------------------------
'                    ' Re-Activate DOMAINFIXED
'                    '-------------------------
'1493:                     If pDomainFixed Then
'1494:                         pFieldEdit.DomainFixed = True
'1495:                     End If
'1496:                 End If
'1497:             End If
'            '----------------------
'            ' Assign Default Value
'            '----------------------
'1501:             If pDefaultValue = "" Then
'1502:                 If pDomain Is Nothing Then
'                    '
'1504:                 Else
'1505:                     If pSubtypeName = "" Then
'1506:                         pSubtypes.defaultValue(0, pFieldName) = Null
'1507:                     Else
'                        '*************************************************************************************
'                        ' Richie: November 3, 2003
'                        ' The line below caused extra lines to be added to the GDB_DEFAULTS table when the
'                        ' field type was numeric. ArcCatalog/ArcMap would also use an errorous default value.
'                        '*************************************************************************************
'                        ' pSubtypes.DefaultValue(pSubtypeCode, pFieldName) = Null
'1514:                     End If
'1515:                 End If
'1516:             Else
'                Select Case pField.Type
'                Case esriFieldTypeDate
'1519:                     If IsDate(pDefaultValue) Then
'1520:                         If pDomain Is Nothing Then
'1521:                             If pSubtypeName = "" Then
'1522:                                 pSubtypes.defaultValue(0, pFieldName) = CDate(pDefaultValue)
'1523:                             Else
'1524:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CDate(pDefaultValue)
'1525:                             End If
'1526:                         Else
'1527:                             If pDomain.MemberOf(CDate(pDefaultValue)) Then
'1528:                                 If pSubtypeName = "" Then
'1529:                                     pSubtypes.defaultValue(0, pFieldName) = CDate(pDefaultValue)
'1530:                                 Else
'1531:                                     pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CDate(pDefaultValue)
'1532:                                 End If
'1533:                             Else
'1534:                                 Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1535:                             End If
'1536:                         End If
'1537:                     Else
'1538:                         Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] cannot be converted to a DATE", , , vbRed, , , True, 2)
'1539:                     End If
'                Case esriFieldTypeDouble
'1541:                     If IsNumeric(pDefaultValue) Then
'1542:                         If pDomain Is Nothing Then
'1543:                             If pSubtypeName = "" Then
'1544:                                 pSubtypes.defaultValue(0, pFieldName) = CDbl(pDefaultValue)
'1545:                             Else
'1546:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CDbl(pDefaultValue)
'1547:                             End If
'1548:                         Else
'1549:                             If pDomain.MemberOf(CDbl(pDefaultValue)) Then
'1550:                                 If pSubtypeName = "" Then
'1551:                                     pSubtypes.defaultValue(0, pFieldName) = CDbl(pDefaultValue)
'1552:                                 Else
'1553:                                     pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CDbl(pDefaultValue)
'1554:                                 End If
'1555:                             Else
'1556:                                 Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1557:                             End If
'1558:                         End If
'1559:                     End If
'                Case esriFieldTypeInteger
'1561:                     If IsNumeric(pDefaultValue) Then
'1562:                         If pDomain Is Nothing Then
'1563:                             If pSubtypeName = "" Then
'1564:                                 pSubtypes.defaultValue(0, pFieldName) = CLng(pDefaultValue)
'1565:                             Else
'1566:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CLng(pDefaultValue)
'1567:                             End If
'1568:                         Else
'1569:                             If pDomain.MemberOf(CLng(pDefaultValue)) Then
'1570:                                 If pSubtypeName = "" Then
'1571:                                     pSubtypes.defaultValue(0, pFieldName) = CLng(pDefaultValue)
'1572:                                 Else
'1573:                                     pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CLng(pDefaultValue)
'1574:                                 End If
'1575:                             Else
'1576:                                 Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1577:                             End If
'1578:                         End If
'1579:                     End If
'                Case esriFieldTypeSingle
'1581:                     If IsNumeric(pDefaultValue) Then
'1582:                         If pDomain Is Nothing Then
'1583:                             If pSubtypeName = "" Then
'1584:                                 pSubtypes.defaultValue(0, pFieldName) = CSng(pDefaultValue)
'1585:                             Else
'1586:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CSng(pDefaultValue)
'1587:                             End If
'1588:                         Else
'1589:                             If pDomain.MemberOf(CSng(pDefaultValue)) Then
'1590:                                 If pSubtypeName = "" Then
'1591:                                     pSubtypes.defaultValue(0, pFieldName) = CSng(pDefaultValue)
'1592:                                 Else
'1593:                                     pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CSng(pDefaultValue)
'1594:                                 End If
'1595:                             Else
'1596:                                 Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1597:                             End If
'1598:                         End If
'1599:                     End If
'                Case esriFieldTypeSmallInteger
'1601:                     If IsNumeric(pDefaultValue) Then
'1602:                         If pDomain Is Nothing Then
'1603:                             If pSubtypeName = "" Then
'1604:                                 pSubtypes.defaultValue(0, pFieldName) = CInt(pDefaultValue)
'1605:                             Else
'1606:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CInt(pDefaultValue)
'1607:                             End If
'1608:                         Else
'1609:                             If pDomain.MemberOf(CInt(pDefaultValue)) Then
'1610:                                 If pSubtypeName = "" Then
'1611:                                     pSubtypes.defaultValue(0, pFieldName) = CInt(pDefaultValue)
'1612:                                 Else
'1613:                                     pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CInt(pDefaultValue)
'1614:                                 End If
'1615:                             Else
'1616:                                 Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1617:                             End If
'1618:                         End If
'1619:                     End If
'                Case esriFieldTypeString
'1621:                     If pDomain Is Nothing Then
'1622:                         If pSubtypeName = "" Then
'1623:                             pSubtypes.defaultValue(0, pFieldName) = CStr(pDefaultValue)
'1624:                         Else
'1625:                             pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CStr(pDefaultValue)
'1626:                         End If
'1627:                     Else
'1628:                         If pDomain.MemberOf(CStr(pDefaultValue)) Then
'1629:                             If pSubtypeName = "" Then
'1630:                                 pSubtypes.defaultValue(0, pFieldName) = CStr(pDefaultValue)
'1631:                             Else
'1632:                                 pSubtypes.defaultValue(pSubtypeCode, pFieldName) = CStr(pDefaultValue)
'1633:                             End If
'1634:                         Else
'1635:                             Call frmOutputWindow.AddOutputMessage("Default Value [" & CStr(pDefaultValue) & "] does not belong to domain [" & pDomain.Name & "]", , , vbRed, , , True, 2)
'1636:                         End If
'1637:                     End If
'1638:                 End Select
'1639:             End If
'1640:         Next pIndexField
'1641:     Next pIndexSubtype
'    '--------------------------
'    ' Set Default Subtype Code
'    '--------------------------
'1645:     If pSubtypeFieldName <> "" Then
'1646:         pSubtypes.DefaultSubtypeCode = CLng(pSubtypeDefaultCode)
'1647:     End If
'    '
'    Exit Sub
'ErrorHandler:
'1651:     Call HandleError(False, "ImportGDB_Subtype " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function ImportGDB_GeometryDef(ByRef pXMLDOMNodeGeometryDef As MSXML2.IXMLDOMNode) As esriGeoDatabase.IGeometryDef
'    On Error GoTo ErrorHandler
'    '--------------------------------------------
'    '        <geometryDef esriGeometryType = ""
'    '                     avgNumPoints = ""
'    '                     gridCount = ""
'    '                     gridSize[0-n] = ""
'    '                     hasM=""
'    '                     hasZ=""/>
'    '--------------------------------------------
'    Dim pXMLDOMNodeFeatureDataset As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeSpatialReference As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListGrid As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeGrid As MSXML2.IXMLDOMNode
'    '
'    Dim pSpatialReference As esriGeometry.ISpatialReference
'    Dim pGeometryDef As esriGeoDatabase.IGeometryDef
'    Dim pGeometryDefEdit  As esriGeoDatabase.IGeometryDefEdit
'    Dim pSize As String
'    Dim pIndex As Long
'    '
'1675:     Set pGeometryDef = New GeometryDef
'1676:     Set pGeometryDefEdit = pGeometryDef
'1677:     pGeometryDefEdit.AvgNumPoints = CLng(pXMLDOMNodeGeometryDef.Attributes.getNamedItem("avgNumPoints").Text)
'1678:     pGeometryDefEdit.GeometryType = CLng(pXMLDOMNodeGeometryDef.Attributes.getNamedItem("esriGeometryType").Text)
'1679:     pGeometryDefEdit.HasM = CBool(pXMLDOMNodeGeometryDef.Attributes.getNamedItem("hasM").Text)
'1680:     pGeometryDefEdit.HasZ = CBool(pXMLDOMNodeGeometryDef.Attributes.getNamedItem("hasZ").Text)
'    '
'1682:     Set pXMLDOMNodeListGrid = pXMLDOMNodeGeometryDef.selectNodes("grid")
'1683:     pGeometryDefEdit.GridCount = pXMLDOMNodeListGrid.length
'1684:     For pIndex = 0 To pXMLDOMNodeListGrid.length - 1 Step 1
'1685:         pSize = CStr(pXMLDOMNodeListGrid.Item(pIndex).Attributes.getNamedItem("size").Text)
'        '
'1687:         Call frmOutputWindow.AddOutputMessage("Assigning Grid [" & CStr(pIndex) & "] with size [" & pSize & "]", , , , , , True, 1)
'        '
'1689:         pGeometryDefEdit.GridSize(pIndex) = CDbl(pSize)
'1690:     Next pIndex
'    '----------------------------------------------------------
'    ' If there is a child "spatialReference" node then add it.
'    '----------------------------------------------------------
'1694:     If pXMLDOMNodeGeometryDef.selectNodes("spatialReference").length = 1 Then
'        '--------------------------------------
'        ' Get SpatialReference from child node
'        '--------------------------------------
'1698:         Set pXMLDOMNodeSpatialReference = pXMLDOMNodeGeometryDef.selectNodes("spatialReference").nextNode
'1699:     Else
'        '-------------------------------------------------
'        ' Get SpatialReference from parent FeatureDataset
'        '-------------------------------------------------
'1703:         Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeGeometryDef.parentNode.parentNode.parentNode
'1704:         Set pXMLDOMNodeSpatialReference = pXMLDOMNodeFeatureDataset.selectNodes("spatialReference").nextNode
'1705:     End If
'1706:     Set pSpatialReference = ImportGDB_SpatialReference(pXMLDOMNodeSpatialReference)
'1707:     Set pGeometryDefEdit.SpatialReference = pSpatialReference
'    '----------------------------
'    ' Return Geometry Definition
'    '----------------------------
'1711:     Set ImportGDB_GeometryDef = pGeometryDef
'    '
'    Exit Function
'ErrorHandler:
'1715:     Call HandleError(False, "ImportGDB_GeometryDef " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Function ImportGDB_SpatialReference(ByRef pXMLDOMNodeSpatialReference As MSXML2.IXMLDOMNode) As esriGeometry.ISpatialReference
'    On Error GoTo ErrorHandler
'    '---------------------------------------------------------------------
'    '    <spatialReference minX="" minY="" precisionXY=""
'    '                      minM=""         precisionM=""
'    '                      minZ=""         precisionZ=""
'    '                      coordinateSystemDescription=""/>
'    '---------------------------------------------------------------------
'    Dim pSpatialReference As esriGeometry.ISpatialReference
'    Dim pESRISpatialReference As esriGeometry.IESRISpatialReference
'    Dim pSpatialReferenceFactory As ISpatialReferenceFactory
'    Dim pCoordinateSystemDescription As String
'    ' Minimum
'    Dim pXMin As Double
'    Dim pYMin As Double
'    Dim pMMin As Double
'    Dim pZMin As Double
'    ' Precision
'    Dim pXYPrecision As Double
'    Dim pMPrecision As Double
'    Dim pZPrecision As Double
'    Dim pBytes As Long
'    '---------------------------------
'    ' Get CoordinateSystemDescription
'    '---------------------------------
'1743:     pCoordinateSystemDescription = pXMLDOMNodeSpatialReference.Attributes.getNamedItem("coordinateSystemDescription").Text
'1744:     If pCoordinateSystemDescription = "" Then
'1745:         Set pSpatialReference = New esriGeometry.UnknownCoordinateSystem
'1746:     Else
'1747:         Set pSpatialReferenceFactory = New esriGeometry.SpatialReferenceEnvironment
'1748:         Call pSpatialReferenceFactory.CreateESRISpatialReference(pCoordinateSystemDescription, pSpatialReference, pBytes)
'1749:     End If
'    '------------------
'    ' Import XY Domain
'    '------------------
'1753:     If pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minX").Text = "" Or _
'       pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minY").Text = "" Or _
'       pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionXY").Text = "" Then
'        '---------------------------
'        ' XY Domain is Blank. Skip.
'        '---------------------------
'1759:     Else
'1760:         pXMin = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minX").Text)
'1761:         pYMin = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minY").Text)
'1762:         pXYPrecision = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionXY").Text)
'1763:         If pXYPrecision <> 0 Then
'1764:             Call pSpatialReference.SetFalseOriginAndUnits(pXMin, pYMin, pXYPrecision)
'1765:         End If
'1766:     End If
'    '-----------------
'    ' Import M Domain
'    '-----------------
'1770:     If pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minM").Text = "" Or _
'       pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionM").Text = "" Then
'        '-------------------------
'        ' M Domain is Blank. Skip
'        '-------------------------
'1775:     Else
'1776:         pMMin = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minM").Text)
'1777:         pMPrecision = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionM").Text)
'1778:         If pMPrecision <> 0 Then
'1779:             Call pSpatialReference.SetMFalseOriginAndUnits(pMMin, pMPrecision)
'1780:         End If
'1781:     End If
'    '-----------------
'    ' Import Z Domain
'    '-----------------
'1785:     If pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minZ").Text = "" Or _
'       pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionZ").Text = "" Then
'        '-------------------
'        ' Z Domain is blank
'        '-------------------
'1790:     Else
'1791:         pZMin = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("minZ").Text)
'1792:         pZPrecision = CDbl(pXMLDOMNodeSpatialReference.Attributes.getNamedItem("precisionZ").Text)
'1793:         If pZPrecision <> 0 Then
'1794:             Call pSpatialReference.SetZFalseOriginAndUnits(pZMin, pZPrecision)
'1795:         End If
'1796:     End If
'    '---------------------------
'    ' Return Spatial Reference
'    '--------------------------
'1800:     Set ImportGDB_SpatialReference = pSpatialReference
'    '
'    Exit Function
'ErrorHandler:
'1804:     Call HandleError(False, "ImportGDB_SpatialReference " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Function MakeField(ByRef pXMLDOMNodeField As MSXML2.IXMLDOMNode) As esriGeoDatabase.IField
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNodeGeometryDef As MSXML2.IXMLDOMNode
'    Dim pField As esriGeoDatabase.IField
'    Dim pFieldEdit As esriGeoDatabase.IFieldEdit
'    Dim pModelInfo As esriGeoDatabase.IModelInfo
'    '--------------------
'    ' Create a new field
'    '--------------------
'1817:     Set pField = New esriGeoDatabase.Field
'1818:     Set pFieldEdit = pField
'1819:     Set pModelInfo = pField
'    '------------------------------
'    ' Edit Common Field Attributes
'    '------------------------------
'1823:     Call frmOutputWindow.AddOutputMessage("Importing Field: " & pXMLDOMNodeField.Attributes.getNamedItem("name").Text, , , , , , True, 1)
'1824:     With pFieldEdit
'1825:         .Name = CStr(pXMLDOMNodeField.Attributes.getNamedItem("name").Text)
'1826:         .AliasName = CStr(pXMLDOMNodeField.Attributes.getNamedItem("aliasName").Text)
'1827:         .Type = CLng(pXMLDOMNodeField.Attributes.getNamedItem("esriFieldType").Text)
'1828:         .IsNullable = CBool(pXMLDOMNodeField.Attributes.getNamedItem("isNullable").Text)
'1829:         .Required = CBool(pXMLDOMNodeField.Attributes.getNamedItem("required").Text)
'1830:         .Editable = CBool(pXMLDOMNodeField.Attributes.getNamedItem("editable").Text)
'1831:     End With
'    '-------------------------------------
'    ' Edit Field-Type specific properties
'    '-------------------------------------
'    Select Case pField.Type
'    Case esriFieldTypeGeometry
'        '-------------------------
'        ' Set Geometry Definition
'        '-------------------------
'1840:         Set pXMLDOMNodeGeometryDef = pXMLDOMNodeField.childNodes.nextNode
'1841:         Set pFieldEdit.GeometryDef = ImportGDB_GeometryDef(pXMLDOMNodeGeometryDef)
'    Case esriFieldTypeOID
'        '
'    Case esriFieldTypeBlob
'1845:         With pFieldEdit
'1846:             .length = CLng(pXMLDOMNodeField.Attributes.getNamedItem("length").Text)
'1847:         End With
'    Case esriFieldTypeDate
'1849:         With pFieldEdit
'1850:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1851:         End With
'    Case esriFieldTypeDouble
'1853:         With pFieldEdit
'1854:             .Precision = CLng(pXMLDOMNodeField.Attributes.getNamedItem("precision").Text)
'1855:             .Scale = CLng(pXMLDOMNodeField.Attributes.getNamedItem("scale").Text)
'1856:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1857:         End With
'    Case esriFieldTypeInteger
'1859:         With pFieldEdit
'1860:             .Precision = CLng(pXMLDOMNodeField.Attributes.getNamedItem("precision").Text)
'1861:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1862:         End With
'    Case esriFieldTypeSingle
'1864:         With pFieldEdit
'1865:             .Precision = CLng(pXMLDOMNodeField.Attributes.getNamedItem("precision").Text)
'1866:             .Scale = CLng(pXMLDOMNodeField.Attributes.getNamedItem("scale").Text)
'1867:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1868:         End With
'    Case esriFieldTypeSmallInteger
'1870:         With pFieldEdit
'1871:             .Precision = CLng(pXMLDOMNodeField.Attributes.getNamedItem("precision").Text)
'1872:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1873:         End With
'    Case esriFieldTypeString
'1875:         With pFieldEdit
'1876:             .length = CLng(pXMLDOMNodeField.Attributes.getNamedItem("length").Text)
'1877:             .DomainFixed = CBool(pXMLDOMNodeField.Attributes.getNamedItem("domainFixed").Text)
'1878:         End With
'1879:     End Select
'    '----------------
'    ' Set Model Name
'    '----------------
'1883:     pModelInfo.ModelName = CStr(pXMLDOMNodeField.Attributes.getNamedItem("modelName").Text)
'    '--------------------------------
'    ' Return reference to new IField
'    '--------------------------------
'1887:     Set MakeField = pField
'    '
'    Exit Function
'ErrorHandler:
'1891:     Call HandleError(False, "MakeField " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''==================================================================================

''-----------------------
'' RELATIONSHIP - Import
''-----------------------

'Public Sub ImportRelationship(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                              ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNodeListRelationship As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeListField As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeListRule As MSXML2.IXMLDOMNodeList

'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeRelationship As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFeatureDataset As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOrigin As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeDestination As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeRule As MSXML2.IXMLDOMNode
'    '
'    Dim pFeatureDataset As esriGeoDatabase.IFeatureDataset
'    Dim pOriginObjectClass As esriGeoDatabase.IObjectClass
'    Dim pDestinationObjectClass As esriGeoDatabase.IObjectClass
'    '
'    Dim pName As String
'    Dim pFeatureDatasetName As String
'    Dim pOriginObjectClassName As String
'    Dim pDestinationObjectClassName As String
'    Dim pForwardLabel As String
'    Dim pBackwardLabel As String
'    Dim pEsriRelCardinality As Long
'    Dim pEsriRelNotification As Long
'    Dim pIsComposite As Boolean
'    Dim pIsAttributed As Boolean
'    Dim pOriginPrimaryKey As String
'    Dim pOriginForeignKey As String
'    Dim pDestinationPrimaryKey As String
'    Dim pDestinationForeignKey As String
'    '
'    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
'    Dim pRelationshipClassContainer As esriGeoDatabase.IRelationshipClassContainer
'    Dim pRelationshipClass As esriGeoDatabase.IRelationshipClass
'    Dim pRelationshipRule As IRelationshipRule
'    Dim pTable As esriGeoDatabase.ITable
'    Dim pField As esriGeoDatabase.IField
'    '
'    Dim pIndexRelationship As Long
'    Dim pIndexField As Long
'    Dim pIndexRule As Long
'    '-----------------------------------------------
'    ' Get GeodatabaseDesigner Node and Dataset list
'    '-----------------------------------------------
'1948:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'1949:     Set pXMLDOMNodeListRelationship = pXMLDOMNodeGeodatabaseDesigner.selectNodes("relationshipClass")
'    '-------------------------------------------------------
'    ' Launch procedure for each type of dataset encountered
'    '-------------------------------------------------------
'1953:     For pIndexRelationship = 0 To pXMLDOMNodeListRelationship.length - 1 Step 1
'1954:         Set pXMLDOMNodeRelationship = pXMLDOMNodeListRelationship.Item(pIndexRelationship)
'1955:         Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeRelationship.selectSingleNode("featureDataset")
'1956:         Set pXMLDOMNodeOrigin = pXMLDOMNodeRelationship.selectSingleNode("origin")
'1957:         Set pXMLDOMNodeDestination = pXMLDOMNodeRelationship.selectSingleNode("destination")
'        '----------------------------------------------------------------------------
'        '    <relationshipClass  database="" owner="" table="" esriRelCardinality="" esriRelNotification="" isComposite="" isAttributed="" originPrimaryKey="" destinationPrimaryKey="" originForeignKey="" destinationForeignKey="">
'        '        <featureDataset database="" owner="" table="" />
'        '        <origin         database="" owner="" table="" label="" />
'        '        <destination    database="" owner="" table="" label="" />
'        '        <!-- Relationship Attribute Fields  -->
'        '        <field />
'        '        <!-- Relationship Rules  -->
'        '        <rule destinationSubtype="" destinationMinimum="" destinationMaximum="" originSubtype="" originMinimum="" originMaximum="" />
'        '    </relationshipClass>
'        '----------------------------------------------------------------------------
'        '   CreateRelationshipClass (in relClassName: String,
'        '                            in OriginClass: IObjectClass,
'        '                            in DestinationClass: IObjectClass,
'        '                            in forwardLabel: String,
'        '                            in backwardLabel: String,
'        '                            in Cardinality: esriRelCardinality,
'        '                            in Notification: esriRelNotification,
'        '                            in IsComposite: Boolean,
'        '                            in IsAttributed: Boolean,
'        '                            in relAttrFields: IFields,
'        '                            in OriginPrimaryKey: String,
'        '                            in destPrimaryKey: String,
'        '                            in OriginForeignKey: String,
'        '                            in destForeignKey: String): IRelationshipClass
'        '----------------------------------------------------------------------------
'        '------------------------------------------------
'        ' Get Relationship Propertied from the XML Node.
'        '------------------------------------------------
'1987:         pName = CStr(pXMLDOMNodeRelationship.Attributes.getNamedItem("table").Text)
'        '
'1989:         Call frmOutputWindow.AddOutputMessage("Importing Relationship: " & pName, , , vbMagenta)
'        '--------------------------------------
'        ' Check if Relationship already exists
'        '--------------------------------------
'1993:         Set pRelationshipClass = modCommon.GetDataset(pWorkspace, pName, esriDTRelationshipClass)
'        '
'1995:         If pRelationshipClass Is Nothing Then
'1996:             pEsriRelCardinality = CLng(pXMLDOMNodeRelationship.Attributes.getNamedItem("esriRelCardinality").Text)
'1997:             pEsriRelNotification = CLng(pXMLDOMNodeRelationship.Attributes.getNamedItem("esriRelNotification").Text)
'1998:             pIsComposite = CBool(pXMLDOMNodeRelationship.Attributes.getNamedItem("isComposite").Text)
'1999:             pIsAttributed = CBool(pXMLDOMNodeRelationship.Attributes.getNamedItem("isAttributed").Text)
'2000:             pOriginPrimaryKey = CStr(pXMLDOMNodeRelationship.Attributes.getNamedItem("originPrimaryKey").Text)
'2001:             pOriginForeignKey = CStr(pXMLDOMNodeRelationship.Attributes.getNamedItem("originForeignKey").Text)
'2002:             pDestinationPrimaryKey = CStr(pXMLDOMNodeRelationship.Attributes.getNamedItem("destinationPrimaryKey").Text)
'2003:             pDestinationForeignKey = CStr(pXMLDOMNodeRelationship.Attributes.getNamedItem("destinationForeignKey").Text)
'            '
'2005:             pFeatureDatasetName = CStr(pXMLDOMNodeFeatureDataset.Attributes.getNamedItem("table").Text)
'            '
'2007:             pOriginObjectClassName = CStr(pXMLDOMNodeOrigin.Attributes.getNamedItem("table").Text)
'2008:             pBackwardLabel = CStr(pXMLDOMNodeOrigin.Attributes.getNamedItem("label").Text)
'            '
'2010:             pDestinationObjectClassName = CStr(pXMLDOMNodeDestination.Attributes.getNamedItem("table").Text)
'2011:             pForwardLabel = CStr(pXMLDOMNodeDestination.Attributes.getNamedItem("label").Text)
'            '
'2013:             Set pOriginObjectClass = modCommon.GetDataset(pWorkspace, pOriginObjectClassName, esriDTTable)
'2014:             Set pDestinationObjectClass = modCommon.GetDataset(pWorkspace, pDestinationObjectClassName, esriDTTable)
'2015:             If (pOriginObjectClass Is Nothing) Or (pDestinationObjectClass Is Nothing) Then
'                '----------------------------
'                ' Missing Origin ObjectClass
'                '----------------------------
'2019:                 If pOriginObjectClass Is Nothing Then
'2020:                     Call frmOutputWindow.AddOutputMessage("The Origin ObjectClass [" & pOriginObjectClassName & "] Does NOT Exist", , , vbRed, , , , 1)
'2021:                 End If
'                '---------------------------------
'                ' Missing Destination ObjectClass
'                '---------------------------------
'2025:                 If pDestinationObjectClass Is Nothing Then
'2026:                     Call frmOutputWindow.AddOutputMessage("The Origin ObjectClass [" & pDestinationObjectClassName & "] Does NOT Exist", , , vbRed, , , , 1)
'2027:                 End If
'2028:             Else
'                '---------------------------------------
'                ' Add non-key attribute fields (if any)
'                '---------------------------------------
'                Dim pFields As IFields
'                Dim pFieldsEdit As IFieldsEdit
'2034:                 Set pFields = Nothing
'2035:                 Set pXMLDOMNodeListField = pXMLDOMNodeRelationship.selectNodes("field")
'2036:                 If pXMLDOMNodeListField.length > 0 Then
'                    '------------------------
'                    ' QI to ITable Interface
'                    '------------------------
'2040:                     Set pFields = New fields
'2041:                     Set pFieldsEdit = pFields
'                    'Set pTable = pRelationshipClass
'2043:                     For pIndexField = 0 To pXMLDOMNodeListField.length - 1 Step 1
'                        '----------------
'                        ' Get Field Node
'                        '----------------
'2047:                         Set pXMLDOMNodeField = pXMLDOMNodeListField.Item(pIndexField)
'                        '---------------------
'                        ' Create a new IField
'                        '---------------------
'2051:                         Set pField = modImportExport.MakeField(pXMLDOMNodeField)
'                        '-----------------------------------------------
'                        ' Add Field to the Relationship Attribute Table
'                        '-----------------------------------------------
'2055:                         Call pFieldsEdit.AddField(pField)
'                        'Call pTable.AddField(pField)
'2057:                     Next pIndexField
'2058:                 End If
'                '-------------------------
'                ' Create the Relationship
'                '-------------------------
'2062:                 Set pFeatureWorkspace = pWorkspace
'2063:                 Set pRelationshipClass = pFeatureWorkspace.CreateRelationshipClass(pName, _
'                                                                                   pOriginObjectClass, _
'                                                                                   pDestinationObjectClass, _
'                                                                                   pForwardLabel, _
'                                                                                   pBackwardLabel, _
'                                                                                   pEsriRelCardinality, _
'                                                                                   pEsriRelNotification, _
'                                                                                   pIsComposite, _
'                                                                                   pIsAttributed, _
'                                                                                   pFieldsEdit, _
'                                                                                   pOriginPrimaryKey, _
'                                                                                   pDestinationPrimaryKey, _
'                                                                                   pOriginForeignKey, _
'                                                                                   pDestinationForeignKey)
'                '----------------------------------------------------------------------------
'                ' Transfer ownership of Relationship to a FeatureDataset (if not standalone)
'                '----------------------------------------------------------------------------
'2080:                 If pFeatureDatasetName <> "" Then
'2081:                     Set pRelationshipClassContainer = pFeatureWorkspace.OpenFeatureDataset(pFeatureDatasetName)
'2082:                     Call pRelationshipClassContainer.AddRelationshipClass(pRelationshipClass)
'2083:                 End If
'                '----------------------------------------
'                ' Add Rules to the Relationship (if any)
'                ' <rule destinationSubtype=""
'                '       destinationMinimum=""
'                '       destinationMaximum=""
'                '       originSubtype=""
'                '       originMinimum=""
'                '       originMaximum=""/>
'                '----------------------------------------
'2093:                 Set pXMLDOMNodeListRule = pXMLDOMNodeRelationship.selectNodes("rule")
'2094:                 If pXMLDOMNodeListRule.length > 0 Then
'                    '
'2096:                     For pIndexRule = 0 To pXMLDOMNodeListRule.length - 1 Step 1
'                        '---------------
'                        ' Get Rule Node
'                        '---------------
'2100:                         Set pXMLDOMNodeRule = pXMLDOMNodeListRule.Item(pIndexRule)
'                        '------------------------
'                        ' Create new Rule Object
'                        '------------------------
'2104:                         Set pRelationshipRule = New RelationshipRule
'2105:                         With pRelationshipRule
'2106:                             .DestinationClassID = pDestinationObjectClass.ObjectClassID
'2107:                             .DestinationSubtypeCode = GetSubtypeCodeFromName(pDestinationObjectClass, _
'                                                      CStr(pXMLDOMNodeRule.Attributes.getNamedItem("destinationSubtype").Text))
'2109:                             .DestinationMaximumCardinality = CLng(pXMLDOMNodeRule.Attributes.getNamedItem("destinationMaximum").Text)
'2110:                             .DestinationMinimumCardinality = CLng(pXMLDOMNodeRule.Attributes.getNamedItem("destinationMinimum").Text)
'2111:                             .OriginClassID = pOriginObjectClass.ObjectClassID
'2112:                             .OriginSubtypeCode = GetSubtypeCodeFromName(pOriginObjectClass, _
'                                                 CStr(pXMLDOMNodeRule.Attributes.getNamedItem("originSubtype").Text))
'2114:                             .OriginMaximumCardinality = CLng(pXMLDOMNodeRule.Attributes.getNamedItem("originMaximum").Text)
'2115:                             .OriginMinimumCardinality = CLng(pXMLDOMNodeRule.Attributes.getNamedItem("originMinimum").Text)
'2116:                         End With
'                        '
'2118:                         Call pRelationshipClass.AddRelationshipRule(pRelationshipRule)
'2119:                     Next pIndexRule
'2120:                 End If
'2121:             End If
'2122:         Else
'            '-----------------------------
'            ' Relationship Already Exists
'            '-----------------------------
'2126:             Call frmOutputWindow.AddOutputMessage("Relationship Already Exists - Skipping", , , , , , True, 1)
'2127:         End If
'2128:     Next pIndexRelationship
'    '
'    Exit Sub
'ErrorHandler:
'2132:     Call HandleError(False, "ImportRelationship " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetSubtypeCodeFromName(ByRef pSubtypes As esriGeoDatabase.ISubtypes, _
'                                       ByRef pSubtypeNameIn As String) As Long
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnumSubtype As esriGeoDatabase.IEnumSubtype
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pSubtypeNameOut As String
'    Dim pSubtypeCodeOut As Long
'    Dim pSubtypeFound As Boolean
'    '
'2145:     pSubtypeFound = False
'2146:     pSubtypeCodeOut = 0
'    '
'2148:     If pSubtypeNameIn = "" Then
'        '------------------------------------------
'        ' ObjectClass has no subtypes. Return '0'.
'        '------------------------------------------
'2152:         pSubtypeFound = True
'2153:     Else
'2154:         Set pEnumSubtype = pSubtypes.Subtypes
'2155:         pSubtypeNameOut = pEnumSubtype.Next(pSubtypeCodeOut)
'2156:         Do Until pSubtypeNameOut = ""
'2157:             If UCase(Trim(pSubtypeNameOut)) = UCase(Trim(pSubtypeNameIn)) Then
'2158:                 pSubtypeFound = True
'2159:                 Exit Do
'2160:             End If
'2161:             pSubtypeNameOut = pEnumSubtype.Next(pSubtypeCodeOut)
'2162:         Loop
'2163:     End If
'    '
'2165:     If pSubtypeFound Then
'        '----------------------------------------------
'        ' Subtype Found OK - Return Valid Subtype Code
'        '----------------------------------------------
'2169:         GetSubtypeCodeFromName = pSubtypeCodeOut
'2170:     Else
'        '-------------------------------------
'        ' Subtype NOT Found - Display Message
'        '-------------------------------------
'2174:         Set pDataset = pSubtypes
'2175:         Call frmOutputWindow.AddOutputMessage("Geodatabase Inconsistancy. Subtype NOT found in Dataset", , , vbRed, True, , , 0)
'2176:         Call frmOutputWindow.AddOutputMessage("Dataset Name: [" & pDataset.Name & "]", , , vbRed, True, , , 1)
'2177:         Call frmOutputWindow.AddOutputMessage("Subtype Name: [" & pSubtypeNameIn & "]", , , vbRed, True, , , 1)
'2178:         Call frmOutputWindow.AddOutputMessage("Returning Subtype Code '0'", , , vbRed, False, , True, 1)
'        '----------
'        ' Return 0
'        '----------
'2182:         GetSubtypeCodeFromName = 0
'2183:     End If
'    '
'    Exit Function
'ErrorHandler:
'2187:     Call HandleError(False, "GetSubtypeCodeFromName " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''==================================================================================

''-----------------------
'' EXPORT - RELATIONSHIP
''-----------------------

'Public Sub ExportRelationship(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                              ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnumDatasetName As esriGeoDatabase.IEnumDatasetName
'    Dim pEnumDatasetNameRL As esriGeoDatabase.IEnumDatasetName
'    Dim pDatasetName As esriGeoDatabase.IDatasetName
'    Dim pDatasetNameRL As esriGeoDatabase.IDatasetName
'    Dim pFeatureDatasetName As esriGeoDatabase.IFeatureDatasetName
'    Dim pName As IName
'    '
'    Dim pRelationshipClass As esriGeoDatabase.IRelationshipClass
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    '----------------------------
'    ' Get GeodatabaseDesigner Node
'    '------------------------------
'2213:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '-------------------------------------
'    ' Export FeatureDataset Relationships
'    '-------------------------------------
'2217:     Set pEnumDatasetName = pWorkspace.DatasetNames(esriDTFeatureDataset)
'2218:     Set pFeatureDatasetName = pEnumDatasetName.Next
'2219:     Do Until pFeatureDatasetName Is Nothing
'2220:         Set pEnumDatasetNameRL = pFeatureDatasetName.RelationshipClassNames
'2221:         Set pDatasetNameRL = pEnumDatasetNameRL.Next
'2222:         Do Until pDatasetNameRL Is Nothing
'2223:             Call frmOutputWindow.AddOutputMessage("Exporting Relationship: " & pDatasetNameRL.Name, , , vbMagenta)
'2224:             Set pName = pDatasetNameRL
'2225:             Set pRelationshipClass = Nothing
'            On Error GoTo ErrorRecovery
'2227:             Set pRelationshipClass = pName.Open
'            On Error GoTo ErrorHandler
'2229:             If pRelationshipClass Is Nothing Then
'2230:                 Call frmOutputWindow.AddOutputMessage("Cannot Open Relationship", , , vbRed)
'2231:             Else
'2232:                 Call ExportRelationship2(pXMLDOMNodeGeodatabaseDesigner, pRelationshipClass)
'2233:             End If
'            '
'2235:             Set pDatasetNameRL = pEnumDatasetNameRL.Next
'2236:         Loop
'2237:         Set pFeatureDatasetName = pEnumDatasetName.Next
'2238:     Loop
'    '---------------------------------
'    ' Export Standalone Relationships
'    '---------------------------------
'2242:     Set pEnumDatasetNameRL = pWorkspace.DatasetNames(esriDTRelationshipClass)
'2243:     Set pDatasetNameRL = pEnumDatasetNameRL.Next
'2244:     Do Until pDatasetNameRL Is Nothing
'2245:         Call frmOutputWindow.AddOutputMessage("Exporting Relationship: " & pDatasetNameRL.Name, , , vbMagenta)
'2246:         Set pName = pDatasetNameRL
'2247:         Set pRelationshipClass = Nothing
'        On Error GoTo ErrorRecovery
'2249:         Set pRelationshipClass = pName.Open
'        On Error GoTo ErrorHandler
'2251:         If pRelationshipClass Is Nothing Then
'2252:             Call frmOutputWindow.AddOutputMessage("Cannot Open Relationship", , , vbRed)
'2253:         Else
'2254:             Call ExportRelationship2(pXMLDOMNodeGeodatabaseDesigner, pRelationshipClass)
'2255:         End If
'        '
'2257:         Set pDatasetNameRL = pEnumDatasetNameRL.Next
'2258:     Loop
'    '
'    Exit Sub
'ErrorRecovery:
'2262:     Call ErrorRecovery("ExportRelationship " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'2263:     Resume Next
'ErrorHandler:
'2265:     Call HandleError(False, "ExportRelationship " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ExportRelationship2(ByRef pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode, _
'                               ByRef pRelationshipClass As esriGeoDatabase.IRelationshipClass)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------
'    '    <relationshipClass database=""
'    '                       owner=""
'    '                       table=""
'    '                       esriRelCardinality=""
'    '                       esriRelNotification=""
'    '                       isComposite=""
'    '                       isAttributed=""
'    '                       originPrimaryKey=""
'    '                       destinationPrimaryKey=""
'    '                       originForeignKey=""
'    '                       destinationForeignKey="">
'    '       <featureDataset database="" owner="" table="" />
'    '       <origin         database="" owner="" table="" label="" />
'    '       <destination    database="" owner="" table="" label="" />
'    '       <field />
'    '       <rule  destinationSubtype=""
'    '              destinationMinimum=""
'    '              destinationMaximum=""
'    '              originSubtype=""
'    '              originMinimum=""
'    '              originMaximum=""/>
'    '    </relationshipClass>
'    '----------------------------------------------------
'    Dim pXMLDOMNodeRelationshipClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFeatureDataset As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOrigin As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeDestination As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeRule As MSXML2.IXMLDOMNode
'    '
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pTable As esriGeoDatabase.ITable
'    Dim pField As esriGeoDatabase.IField
'    Dim pModelInfo As esriGeoDatabase.IModelInfo
'    Dim pEnumRule As esriGeoDatabase.IEnumRule
'    Dim pRelationshipRule As esriGeoDatabase.IRelationshipRule
'    Dim pDatasetDestination As esriGeoDatabase.IDataset
'    Dim pDatasetOrigin As esriGeoDatabase.IDataset
'    Dim pSubtypesDestination As esriGeoDatabase.ISubtypes
'    Dim pSubtypesOrigin As esriGeoDatabase.ISubtypes
'    '
'    Dim pIndex As Long
'    '
'2315:     Set pDataset = pRelationshipClass
'    '-----------------------------
'    ' Add Relationship properties
'    '-----------------------------
'    On Error GoTo ErrorRecovery
'2320:     Set pDatasetOrigin = pRelationshipClass.OriginClass
'2321:     If pDatasetOrigin Is Nothing Then
'2322:         Call frmOutputWindow.AddOutputMessage("Cannot Open Origin FeatureClass!", , , vbRed)
'        Exit Sub
'2324:     End If
'    On Error GoTo ErrorHandler
'    On Error GoTo ErrorRecovery
'2327:     Set pDatasetDestination = pRelationshipClass.DestinationClass
'2328:     If pDatasetDestination Is Nothing Then
'2329:         Call frmOutputWindow.AddOutputMessage("Cannot Open Destination FeatureClass!", , , vbRed)
'        Exit Sub
'2331:     End If
'    On Error GoTo ErrorHandler
'    '---------------------------
'    ' Add new Relationship Node
'    '---------------------------
'2336:     Set pXMLDOMNodeRelationshipClass = pXMLDOMNodeGeodatabaseDesigner.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "relationshipClass", "")
'2337:     Set pXMLDOMNodeRelationshipClass = pXMLDOMNodeGeodatabaseDesigner.appendChild(pXMLDOMNodeRelationshipClass)
'    '
'2339:     Call AddQualifiedTableNameParts(pXMLDOMNodeRelationshipClass, pDataset)
'    '
'2341:     Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "esriRelCardinality", CStr(pRelationshipClass.Cardinality))
'2342:     Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "esriRelNotification", CStr(pRelationshipClass.Notification))
'2343:     Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "isComposite", CStr(pRelationshipClass.IsComposite))
'2344:     Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "isAttributed", CStr(pRelationshipClass.IsAttributed))
'2345:     If pRelationshipClass.Cardinality = esriRelCardinalityManyToMany Then
'2346:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "originPrimaryKey", CStr(pRelationshipClass.OriginPrimaryKey))
'2347:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "originForeignKey", CStr(pRelationshipClass.OriginForeignKey))
'2348:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "destinationPrimaryKey", CStr(pRelationshipClass.DestinationPrimaryKey))
'2349:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "destinationForeignKey", CStr(pRelationshipClass.DestinationForeignKey))
'2350:     Else
'2351:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "originPrimaryKey", CStr(pRelationshipClass.OriginPrimaryKey))
'2352:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "originForeignKey", CStr(pRelationshipClass.OriginForeignKey))
'2353:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "destinationPrimaryKey", "")
'2354:         Call modCommon.AddNodeAttribute(pXMLDOMNodeRelationshipClass, "destinationForeignKey", "")
'2355:     End If
'    '-------------------------
'    ' Add FeatureDataset Node
'    '-------------------------
'2359:     Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeRelationshipClass.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "featureDataset", "")
'2360:     Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeRelationshipClass.appendChild(pXMLDOMNodeFeatureDataset)
'    '
'2362:     If pRelationshipClass.FeatureDataset Is Nothing Then
'2363:         Call AddQualifiedTableNameParts(pXMLDOMNodeFeatureDataset, Nothing)
'2364:     Else
'2365:         Call AddQualifiedTableNameParts(pXMLDOMNodeFeatureDataset, pRelationshipClass.FeatureDataset)
'2366:     End If
'    '-----------------
'    ' Add Origin Node
'    '-----------------
'2370:     Set pXMLDOMNodeOrigin = pXMLDOMNodeRelationshipClass.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "origin", "")
'2371:     Set pXMLDOMNodeOrigin = pXMLDOMNodeRelationshipClass.appendChild(pXMLDOMNodeOrigin)
'    '
'2373:     Call AddQualifiedTableNameParts(pXMLDOMNodeOrigin, pDatasetOrigin)
'2374:     Call modCommon.AddNodeAttribute(pXMLDOMNodeOrigin, "label", CStr(pRelationshipClass.BackwardPathLabel))
'    '----------------------
'    ' Add Destination Node
'    '----------------------
'2378:     Set pXMLDOMNodeDestination = pXMLDOMNodeRelationshipClass.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "destination", "")
'2379:     Set pXMLDOMNodeDestination = pXMLDOMNodeRelationshipClass.appendChild(pXMLDOMNodeDestination)
'    '
'2381:     Call AddQualifiedTableNameParts(pXMLDOMNodeDestination, pDatasetDestination)
'2382:     Call modCommon.AddNodeAttribute(pXMLDOMNodeDestination, "label", CStr(pRelationshipClass.ForwardPathLabel))
'    '--------------------------------------------------------
'    ' Add non-key fields in Attributed Relationship (if any)
'    '--------------------------------------------------------
'2386:     If pRelationshipClass.IsAttributed Then
'2387:         Set pTable = pRelationshipClass
'2388:         For pIndex = 0 To pTable.fields.FieldCount - 1 Step 1
'2389:             Set pField = pTable.fields.Field(pIndex)
'2390:             If UCase(pField.Name) <> UCase(pRelationshipClass.OriginForeignKey) And _
'               UCase(pField.Name) <> UCase(pRelationshipClass.DestinationForeignKey) And _
'               pField.Type <> esriFieldTypeOID Then
'                '----------------
'                ' Add Field Node
'                '----------------
'2396:                 Set pModelInfo = pField
'2397:                 Set pXMLDOMNodeField = pXMLDOMNodeRelationshipClass.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "field", "")
'2398:                 Set pXMLDOMNodeField = pXMLDOMNodeRelationshipClass.appendChild(pXMLDOMNodeField)
'2399:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "name", CStr(pField.Name))
'2400:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "aliasName", CStr(pField.AliasName))
'2401:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "esriFieldType", CStr(pField.Type))
'2402:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "length", CStr(pField.length))
'2403:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "precision", CStr(pField.Precision))
'2404:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "required", CStr(pField.Required))
'2405:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "scale", CStr(pField.Scale))
'2406:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "domainFixed", CStr(pField.DomainFixed))
'2407:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "editable", CStr(pField.Editable))
'2408:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "isNullable", CStr(pField.IsNullable))
'2409:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "modelName", CStr(pModelInfo.ModelName))
'2410:             End If
'2411:         Next pIndex
'2412:     End If
'    '------------------------
'    ' Add Relationship Rules
'    '------------------------
'2416:     Set pSubtypesDestination = pDatasetDestination
'2417:     Set pSubtypesOrigin = pDatasetOrigin
'2418:     Set pEnumRule = pRelationshipClass.RelationshipRules
'2419:     Set pRelationshipRule = pEnumRule.Next
'2420:     Do Until pRelationshipRule Is Nothing
'        '------------------------------------
'        '       <rule destinationSubtype=""
'        '             destinationMinimum=""
'        '             destinationMaximum=""
'        '             originSubtype=""
'        '             originMinimum=""
'        '             originMaximum=""/>
'        '------------------------------------
'2429:         Call frmOutputWindow.AddOutputMessage("Exporting Rule [" & CStr(pRelationshipRule.id) & "]", , , , , , , 1)
'2430:         If IsSubtypeCodeValid(pSubtypesOrigin, pRelationshipRule.OriginSubtypeCode) Then
'2431:             If IsSubtypeCodeValid(pSubtypesDestination, pRelationshipRule.DestinationSubtypeCode) Then
'2432:                 Set pXMLDOMNodeRule = pXMLDOMNodeRelationshipClass.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "rule", "")
'2433:                 Set pXMLDOMNodeRule = pXMLDOMNodeRelationshipClass.appendChild(pXMLDOMNodeRule)
'2434:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "originSubtype", CStr(pSubtypesOrigin.SubtypeName(pRelationshipRule.OriginSubtypeCode)))
'2435:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "destinationSubtype", CStr(pSubtypesDestination.SubtypeName(pRelationshipRule.DestinationSubtypeCode)))
'2436:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "originMinimum", CStr(pRelationshipRule.OriginMinimumCardinality))
'2437:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "originMaximum", CStr(pRelationshipRule.OriginMaximumCardinality))
'2438:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "destinationMinimum", CStr(pRelationshipRule.DestinationMinimumCardinality))
'2439:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeRule, "destinationMaximum", CStr(pRelationshipRule.DestinationMaximumCardinality))
'2440:             Else
'                ' ************ Destination Subtype Code is Invalid **************
'2442:                 Call frmOutputWindow.AddOutputMessage("Destination Subtype Code is invalid", , , vbRed, , , , 2)
'2443:             End If
'2444:         Else
'            ' ************ Origin Subtype Code is Invalid **************
'2446:             Call frmOutputWindow.AddOutputMessage("Origin Subtype Code is invalid", , , vbRed, , , , 2)
'2447:         End If
'        '
'2449:         Set pRelationshipRule = pEnumRule.Next
'2450:     Loop
'    '
'    Exit Sub
'ErrorRecovery:
'2454:     Call ErrorRecovery("ExportRelationship2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'2455:     Resume Next
'ErrorHandler:
'2457:     Call HandleError(False, "ExportRelationship2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function IsSubtypeCodeValid(ByRef pSubtypes As esriGeoDatabase.ISubtypes, _
'                                    ByRef pSubtypeCodeIn As Long) As Boolean
'On Error GoTo ErrorHandler
'    '
'    Dim pEnumSubtype As esriGeoDatabase.IEnumSubtype
'    Dim pSubtypeName As String
'    Dim pSubtypeCodeOut As Long
'    Dim pValidateSubtypeCode As Boolean
'    '
'2469:     pValidateSubtypeCode = False
'    '
'2471:     Set pEnumSubtype = pSubtypes.Subtypes
'2472:     pSubtypeName = pEnumSubtype.Next(pSubtypeCodeOut)
'2473:     Do Until pSubtypeName = ""
'2474:         If UCase(Trim(pSubtypeCodeIn)) = UCase(Trim(pSubtypeCodeOut)) Then
'2475:             pValidateSubtypeCode = True
'2476:             Exit Do
'2477:         End If
'2478:         pSubtypeName = pEnumSubtype.Next(pSubtypeCodeOut)
'2479:     Loop
'    '
'2481:     IsSubtypeCodeValid = pValidateSubtypeCode
'    '
'    Exit Function
'ErrorHandler:
'2485:     Call HandleError(False, "IsSubtypeCodeValid " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''=====================================================================

''-------------------------
'' Export GeometricNetwork
''-------------------------

'Public Sub ExportGeometricNetwork(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                                  ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '---------------------------------------------------------------------------
'    ' This routine will export both the GeometricNetwork and Connectivity Rules
'    '---------------------------------------------------------------------------
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    '
'    Dim pEnumDatasetNameFD As IEnumDatasetName
'    Dim pEnumDatasetNameGN As IEnumDatasetName
'    Dim pFeatureDatasetName As IFeatureDatasetName
'    Dim pGeometricNetworkName As IGeometricNetworkName
'    '------------------------
'    ' Add header to XML file
'    '------------------------
'2509:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '-----------------------
'    ' Get GeometricNetworks
'    '-----------------------
'2513:     Set pEnumDatasetNameFD = pWorkspace.DatasetNames(esriDTFeatureDataset)
'2514:     If Not (pEnumDatasetNameFD Is Nothing) Then
'2515:         Call pEnumDatasetNameFD.Reset
'2516:         Set pFeatureDatasetName = pEnumDatasetNameFD.Next
'2517:         Do Until pFeatureDatasetName Is Nothing
'2518:             Set pEnumDatasetNameGN = pFeatureDatasetName.GeometricNetworkNames
'2519:             If Not (pEnumDatasetNameGN Is Nothing) Then
'2520:                 Call pEnumDatasetNameGN.Reset
'2521:                 Set pGeometricNetworkName = pEnumDatasetNameGN.Next
'2522:                 Do Until pGeometricNetworkName Is Nothing
'2523:                     Call ExportGeometricNetwork2(pXMLDOMNodeGeodatabaseDesigner, pGeometricNetworkName)
'                    '
'2525:                     Set pGeometricNetworkName = pEnumDatasetNameGN.Next
'2526:                 Loop
'2527:             End If
'2528:             Set pFeatureDatasetName = pEnumDatasetNameFD.Next
'2529:         Loop
'2530:     End If
'    '
'    Exit Sub
'ErrorHandler:
'2534:     Call HandleError(False, "ExportGeometricNework " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ExportGeometricNetwork2(ByRef pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode, _
'                                   ByRef pGeometricNetworkName As esriGeoDatabase.IGeometricNetworkName)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    Dim pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeJunctionConnRule As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdgeConnRule As MSXML2.IXMLDOMNode
'    '
'    Dim pXMLDOMNodeChild As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeParent As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeParent2 As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeWeight As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeWeightAssociation As MSXML2.IXMLDOMNode
'    '
'    Dim pGeometricNetwork As IGeometricNetwork
'    Dim pDatasetName2 As IDatasetName
'    Dim pName As IName
'    Dim pIndexWeight As Long
'    Dim pNetwork As esriGeoDatabase.INetwork
'    Dim pNetSchema As esriGeoDatabase.INetSchema
'    Dim pNetWeight As esriGeoDatabase.INetWeight
'    Dim pNetWeightAssociation As esriGeoDatabase.INetWeightAssociation
'    Dim pEnumNetWeightAssociation As esriGeoDatabase.IEnumNetWeightAssociation
'    '
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pEnumRule As esriGeoDatabase.IEnumRule
'    Dim pRule As IRule
'    Dim pJunctionConnectivityRule2 As esriGeoDatabase.IJunctionConnectivityRule2
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
'    '
'    Dim pFeatureClassEdge As esriGeoDatabase.IFeatureClass
'    Dim pDatasetEdge As esriGeoDatabase.IDataset
'    Dim pSubtypesEdge As esriGeoDatabase.ISubtypes
'    Dim pSubtypeNameEdge As String
'    '
'    Dim pFeatureClassJunction As esriGeoDatabase.IFeatureClass
'    Dim pDatasetJunction As esriGeoDatabase.IDataset
'    Dim pSubtypesJunction As esriGeoDatabase.ISubtypes
'    Dim pSubtypeNameJunction As String
'    '
'    Dim pEdgeConnectivityRule As esriGeoDatabase.IEdgeConnectivityRule
'    Dim pFeatureClassEdge1 As esriGeoDatabase.IFeatureClass
'    Dim pFeatureClassEdge2 As esriGeoDatabase.IFeatureClass
'    Dim pDatasetEdge1 As esriGeoDatabase.IDataset
'    Dim pDatasetEdge2 As esriGeoDatabase.IDataset
'    Dim pSubtypesEdge1 As esriGeoDatabase.ISubtypes
'    Dim pSubtypesEdge2 As esriGeoDatabase.ISubtypes
'    Dim pSubtypeEdgeName1 As String
'    Dim pSubtypeEdgeName2 As String
'    Dim pSubtypeJunctionName As String
'    Dim pJunctionCounter As Long
'    '
'    Dim pSQLSyntax As esriGeoDatabase.ISQLSyntax
'    Dim pDatabaseName As String
'    Dim pOwnerName As String
'    Dim pTableName As String
'    '
'2597:     Set pDOMDocument = pXMLDOMNodeGeodatabaseDesigner.ownerDocument
'    '
'2599:     Set pDatasetName2 = pGeometricNetworkName
'2600:     Call frmOutputWindow.AddOutputMessage("Exporting GeometricNetwork: " & pDatasetName2.Name, , , vbMagenta)
'    '
'2602:     Set pName = pGeometricNetworkName
'    On Error GoTo ErrorRecovery
'2604:     Set pGeometricNetwork = pName.Open
'    On Error GoTo ErrorHandler
'2606:     If pGeometricNetwork Is Nothing Then
'2607:         Call frmOutputWindow.AddOutputMessage("Failed To GeometricNetwork Dataset", , , vbRed)
'        Exit Sub
'2609:     End If
'    '
'2611:     Set pDataset = pGeometricNetwork
'2612:     Set pFeatureWorkspace = pDataset.Workspace
'    '--------------------------------------------------------------------
'    ' Add GeometricNetwork to XML
'    ' <geometricNetwork database="" owner="" table="" esriNetworkType=""
'    '--------------------------------------------------------------------
'2617:     Set pXMLDOMNodeGeometricNetwork = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "geometricNetwork", "")
'2618:     Set pXMLDOMNodeGeometricNetwork = pXMLDOMNodeGeodatabaseDesigner.appendChild(pXMLDOMNodeGeometricNetwork)
'2619:     Call AddQualifiedTableNameParts(pXMLDOMNodeGeometricNetwork, pDataset)
'2620:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometricNetwork, "esriNetworkType", CStr(pGeometricNetwork.NetworkType))
'2621:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometricNetwork, "configKeyword", "")
'    '----------------------------------------------------------------------------------------------------
'    ' Add GN FeatureClasses to XML
'    ' <featureClass database="" owner="" table="" esriFeatureType="" esriNetworkClassAncillaryRole="" />
'    '----------------------------------------------------------------------------------------------------
'2626:     If mGNExportFeatureClass Then
'2627:         Call frmOutputWindow.AddOutputMessage("Exporting Simple Junctions", , , , , , True, 1)
'2628:         Call WriteNetworkClassToXML(pXMLDOMNodeGeometricNetwork, pGeometricNetwork, esriFTSimpleJunction)
'2629:         Call frmOutputWindow.AddOutputMessage("Exporting Simple Edges", , , , , , True, 1)
'2630:         Call WriteNetworkClassToXML(pXMLDOMNodeGeometricNetwork, pGeometricNetwork, esriFTSimpleEdge)
'2631:         Call frmOutputWindow.AddOutputMessage("Exporting Complex Edges", , , , , , True, 1)
'2632:         Call WriteNetworkClassToXML(pXMLDOMNodeGeometricNetwork, pGeometricNetwork, esriFTComplexEdge)
'2633:     End If
'    '-----------------------------------------------------------
'    ' Add all GeometricNetwork Weights
'    ' <weight name="" esriweighttype="" bitGateSize="">
'    '    <association database="" owner="" table="" field="" />
'    ' </weight>
'    '-----------------------------------------------------------
'2640:     Set pNetwork = pGeometricNetwork.Network
'2641:     Set pNetSchema = pNetwork
'    '----------------------------------
'    ' Loop through all Network Weights
'    '-----------------------------------
'2645:     For pIndexWeight = 0 To pNetSchema.WeightCount - 1 Step 1
'2646:         Call frmOutputWindow.AddOutputMessage("Writing Weights [" & CStr(pIndexWeight) & "]", , , , , , True, 1)
'2647:         Set pNetWeight = pNetSchema.Weight(pIndexWeight)
'        '---------------------------------------------------
'        ' Add Weight Node and Attributes
'        ' <weight name="" esriweighttype="" bitGateSize="">
'        '---------------------------------------------------
'2652:         Set pXMLDOMNodeWeight = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "weight", "")
'2653:         Set pXMLDOMNodeWeight = pXMLDOMNodeGeometricNetwork.appendChild(pXMLDOMNodeWeight)
'        '
'2655:         Call modCommon.AddNodeAttribute(pXMLDOMNodeWeight, "name", CStr(pNetWeight.WeightName))
'2656:         Call modCommon.AddNodeAttribute(pXMLDOMNodeWeight, "esriWeightType", CStr(pNetWeight.WeightType))
'2657:         If pNetWeight.WeightType = esriGeoDatabase.esriWeightType.esriWTBitGate Then
'2658:             Call modCommon.AddNodeAttribute(pXMLDOMNodeWeight, "bitGateSize", CStr(pNetWeight.BitGateSize))
'2659:         Else
'2660:             Call modCommon.AddNodeAttribute(pXMLDOMNodeWeight, "bitGateSize", "")
'2661:         End If
'        '--------------------------------------
'        ' Loop through all Weight Associations
'        '--------------------------------------
'2665:         Set pEnumNetWeightAssociation = pNetSchema.WeightAssociations(pIndexWeight)
'2666:         Set pNetWeightAssociation = pEnumNetWeightAssociation.Next
'2667:         Do Until pNetWeightAssociation Is Nothing
'            '-------------------------------------------------------
'            ' Add WeightAssociation
'            ' <association database="" owner="" table="" field=""/>
'            '-------------------------------------------------------
'2672:             Set pXMLDOMNodeWeightAssociation = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "association", "")
'2673:             Set pXMLDOMNodeWeightAssociation = pXMLDOMNodeWeight.appendChild(pXMLDOMNodeWeightAssociation)
'2674:             Call AddQualifiedTableNameParts(pXMLDOMNodeWeightAssociation, pFeatureWorkspace.OpenFeatureClass(pNetWeightAssociation.TableName))
'2675:             Call modCommon.AddNodeAttribute(pXMLDOMNodeWeightAssociation, "field", pNetWeightAssociation.FieldName)
'            '
'2677:             Set pNetWeightAssociation = pEnumNetWeightAssociation.Next
'2678:         Loop
'2679:     Next pIndexWeight
'    '-----------------------------------------------------
'    ' Add ALL Connectivity Rules
'    ' <junctionConnectivityRule> & <edgeConnectivityRule>
'    '-----------------------------------------------------
'2684:     Set pXMLDOMNodeJunctionConnRule = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "junctionConnectivityRule", "")
'2685:     Set pXMLDOMNodeJunctionConnRule = pXMLDOMNodeGeometricNetwork.appendChild(pXMLDOMNodeJunctionConnRule)
'2686:     Set pXMLDOMNodeEdgeConnRule = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "edgeConnectivityRule", "")
'2687:     Set pXMLDOMNodeEdgeConnRule = pXMLDOMNodeGeometricNetwork.appendChild(pXMLDOMNodeEdgeConnRule)
'    '
'2689:     Set pEnumRule = pGeometricNetwork.Rules
'2690:     Set pRule = pEnumRule.Next
'2691:     Set pFeatureClassContainer = pGeometricNetwork
'    '
'2693:     Call frmOutputWindow.AddOutputMessage("Export GeometricNetwork: Writing Connectivity Rules")
'    '
'2695:     Do Until pRule Is Nothing
'2696:         If TypeOf pRule Is esriGeoDatabase.IJunctionConnectivityRule Then
'2697:             If mGNExportJunctionConnRule Then
'                '-----------------------------
'                ' Junction Connectivity Rules
'                '-----------------------------
'2701:                 Set pJunctionConnectivityRule2 = pRule
'                '---------------------
'                ' Get Rule Properties
'                '---------------------
'2705:                 Set pFeatureClassEdge = pFeatureClassContainer.ClassByID(pJunctionConnectivityRule2.EdgeClassID)
'2706:                 Set pDatasetEdge = pFeatureClassEdge
'2707:                 Set pSubtypesEdge = pDatasetEdge
'2708:                 If pSubtypesEdge.HasSubtype Then
'2709:                     pSubtypeNameEdge = pSubtypesEdge.SubtypeName(pJunctionConnectivityRule2.EdgeSubtypeCode)
'2710:                 Else
'2711:                     pSubtypeNameEdge = ""
'2712:                 End If
'                '
'2714:                 Set pFeatureClassJunction = pFeatureClassContainer.ClassByID(pJunctionConnectivityRule2.JunctionClassID)
'2715:                 Set pDatasetJunction = pFeatureClassJunction
'2716:                 Set pSubtypesJunction = pDatasetJunction
'2717:                 If pSubtypesJunction.HasSubtype Then
'2718:                     pSubtypeNameJunction = pSubtypesJunction.SubtypeName(pJunctionConnectivityRule2.JunctionSubtypeCode)
'2719:                 Else
'2720:                     pSubtypeNameJunction = ""
'2721:                 End If
'                '-----------------------
'                ' Add entry to log file
'                '-----------------------
'2725:                 Call frmOutputWindow.AddOutputMessage("EJ: " & pDatasetEdge.Name & "|" & pSubtypeNameEdge & " > " & pDatasetJunction.Name & "|" & pSubtypeNameJunction, , , , , , True, 1)
'                '----------------------------------------------
'                ' Edge FC <edge database="" owner="" table="">
'                '----------------------------------------------
'2729:                 Set pSQLSyntax = pDatasetEdge.Workspace
'2730:                 Call pSQLSyntax.ParseTableName(CStr(pDatasetEdge.Name), pDatabaseName, pOwnerName, pTableName)
'2731:                 Set pXMLDOMNodeList = pXMLDOMNodeJunctionConnRule.selectNodes("edge[@database='" & pDatabaseName & "'" & _
'                                                                              " and @owner='" & pOwnerName & "'" & _
'                                                                              " and @table='" & pTableName & "']")
'2734:                 If pXMLDOMNodeList.length = 0 Then
'                    '--------------
'                    ' Add New Node
'                    '--------------
'2738:                     Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "edge", "")
'2739:                     Set pXMLDOMNodeChild = pXMLDOMNodeJunctionConnRule.appendChild(pXMLDOMNodeChild)
'                    '
'2741:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "database", pDatabaseName)
'2742:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "owner", pOwnerName)
'2743:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "table", pTableName)
'                    '
'2745:                     Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2746:                 Else
'2747:                     Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2748:                 End If
'                '--------------------------------
'                ' Edge Subtype <subtype name="">
'                '--------------------------------
'2752:                 Set pXMLDOMNodeList = pXMLDOMNodeParent.selectNodes("subtype[@name='" & CStr(pSubtypeNameEdge) & "']")
'2753:                 If pXMLDOMNodeList.length = 0 Then
'                    '--------------
'                    ' Add New Node
'                    '--------------
'2757:                     Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'2758:                     Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                    '
'2760:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "name", pSubtypeNameEdge)
'                    '
'2762:                     Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2763:                 Else
'2764:                     Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2765:                 End If
'                '------------------------------------------------------
'                ' Junction FC <junction database="" owner="" table="">
'                '------------------------------------------------------
'2769:                 Set pSQLSyntax = pDatasetJunction.Workspace
'2770:                 Call pSQLSyntax.ParseTableName(CStr(pDatasetJunction.Name), pDatabaseName, pOwnerName, pTableName)
'2771:                 Set pXMLDOMNodeList = pXMLDOMNodeJunctionConnRule.selectNodes("junction[@database='" & pDatabaseName & "'" & _
'                                                                              " and @owner='" & pOwnerName & "'" & _
'                                                                              " and @table='" & pTableName & "']")
'2774:                 If pXMLDOMNodeList.length = 0 Then
'                    '--------------
'                    ' Add New Node
'                    '--------------
'2778:                     Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "junction", "")
'2779:                     Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                    '
'2781:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "database", pDatabaseName)
'2782:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "owner", pOwnerName)
'2783:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "table", pTableName)
'                    '
'2785:                     Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2786:                 Else
'2787:                     Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2788:                 End If
'                '----------------------------------------------------------------------------------
'                ' Junction Subytype <subtype name="" eMin="" eMax="" jMin="" jMax="" default="" />
'                ' Duplication not possible, hence no test.
'                '----------------------------------------------------------------------------------
'2793:                 Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'2794:                 Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                '
'2796:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "name", pSubtypeNameJunction)
'2797:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "eMin", CStr(pJunctionConnectivityRule2.EdgeMinimumCardinality))
'2798:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "eMax", CStr(pJunctionConnectivityRule2.EdgeMaximumCardinality))
'2799:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "jMin", CStr(pJunctionConnectivityRule2.JunctionMinimumCardinality))
'2800:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "jMax", CStr(pJunctionConnectivityRule2.JunctionMaximumCardinality))
'2801:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "default", CStr(pJunctionConnectivityRule2.DefaultJunction))
'2802:             End If
'2803:         Else
'2804:             If TypeOf pRule Is esriGeoDatabase.IEdgeConnectivityRule Then
'2805:                 If mGNExportEdgeConnRule Then
'                    '-------------------------
'                    ' Edge Connectivity Rules
'                    '-------------------------
'2809:                     Set pEdgeConnectivityRule = pRule
'                    '---------------------
'                    ' Get Rule Properties
'                    '---------------------
'2813:                     Set pFeatureClassEdge1 = pFeatureClassContainer.ClassByID(pEdgeConnectivityRule.FromEdgeClassID)
'2814:                     Set pFeatureClassEdge2 = pFeatureClassContainer.ClassByID(pEdgeConnectivityRule.ToEdgeClassID)
'2815:                     Set pDatasetEdge1 = pFeatureClassEdge1
'2816:                     Set pDatasetEdge2 = pFeatureClassEdge2
'2817:                     Set pSubtypesEdge1 = pFeatureClassEdge1
'2818:                     Set pSubtypesEdge2 = pFeatureClassEdge2
'2819:                     If pSubtypesEdge1.HasSubtype Then
'2820:                         pSubtypeEdgeName1 = pSubtypesEdge1.SubtypeName(pEdgeConnectivityRule.FromEdgeSubtypeCode)
'2821:                     Else
'2822:                         pSubtypeEdgeName1 = ""
'2823:                     End If
'2824:                     If pSubtypesEdge2.HasSubtype Then
'2825:                         pSubtypeEdgeName2 = pSubtypesEdge2.SubtypeName(pEdgeConnectivityRule.ToEdgeSubtypeCode)
'2826:                     Else
'2827:                         pSubtypeEdgeName2 = ""
'2828:                     End If
'                    '-----------------------
'                    ' Add entry to log file
'                    '-----------------------
'2832:                     Call frmOutputWindow.AddOutputMessage("EE: " & pDatasetEdge1.Name & "|" & pSubtypeEdgeName1 & " > " & pDatasetEdge2.Name & "|" & pSubtypeEdgeName2, , , , , , True, 1)
'                    '--------------------------------------
'                    ' Add FROM Edge FC
'                    ' <edge database="" owner="" table="">
'                    '--------------------------------------
'2837:                     Set pSQLSyntax = pDatasetEdge1.Workspace
'2838:                     Call pSQLSyntax.ParseTableName(CStr(pDatasetEdge1.Name), pDatabaseName, pOwnerName, pTableName)
'2839:                     Set pXMLDOMNodeList = pXMLDOMNodeJunctionConnRule.selectNodes("edge[@database='" & pDatabaseName & "'" & _
'                                                                                  " and @owner='" & pOwnerName & "'" & _
'                                                                                  " and @table='" & pTableName & "']")
'2842:                     If pXMLDOMNodeList.length = 0 Then
'                        '--------------
'                        ' Add New Node
'                        '--------------
'2846:                         Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "edge", "")
'2847:                         Set pXMLDOMNodeChild = pXMLDOMNodeEdgeConnRule.appendChild(pXMLDOMNodeChild)
'                        '
'2849:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "database", pDatabaseName)
'2850:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "owner", pOwnerName)
'2851:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "table", pTableName)
'                        '
'2853:                         Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2854:                     Else
'2855:                         Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2856:                     End If
'                    '---------------------------
'                    ' Add FROM Edge Subtype
'                    ' <subtype name="">
'                    '---------------------------
'2861:                     Set pXMLDOMNodeList = pXMLDOMNodeParent.selectNodes("subtype[@name='" & CStr(pSubtypeEdgeName1) & "']")
'2862:                     If pXMLDOMNodeList.length = 0 Then
'                        '--------------
'                        ' Add New Node
'                        '--------------
'2866:                         Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'2867:                         Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                        '
'2869:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "name", pSubtypeEdgeName1)
'                        '
'2871:                         Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2872:                     Else
'2873:                         Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2874:                     End If
'                    '--------------------------------------
'                    ' Add TO Edge FC
'                    ' <edge database="" owner="" table="">
'                    '--------------------------------------
'2879:                     Set pSQLSyntax = pDatasetEdge2.Workspace
'2880:                     Call pSQLSyntax.ParseTableName(CStr(pDatasetEdge2.Name), pDatabaseName, pOwnerName, pTableName)
'2881:                     Set pXMLDOMNodeList = pXMLDOMNodeJunctionConnRule.selectNodes("edge[@database='" & pDatabaseName & "'" & _
'                                                                                  " and @owner='" & pOwnerName & "'" & _
'                                                                                  " and @table='" & pTableName & "']")
'2884:                     If pXMLDOMNodeList.length = 0 Then
'                        '--------------
'                        ' Add New Node
'                        '--------------
'2888:                         Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "edge", "")
'2889:                         Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                        '
'2891:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "database", pDatabaseName)
'2892:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "owner", pOwnerName)
'2893:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "table", pTableName)
'                        '
'2895:                         Set pXMLDOMNodeParent = pXMLDOMNodeChild
'2896:                     Else
'2897:                         Set pXMLDOMNodeParent = pXMLDOMNodeList.nextNode
'2898:                     End If
'                    '---------------------------
'                    ' Add TO Edge Subtype
'                    ' <subtype name="">
'                    '---------------------------
'2903:                     Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'2904:                     Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                    '
'2906:                     Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "name", pSubtypeEdgeName2)
'                    '
'2908:                     Set pXMLDOMNodeParent = pXMLDOMNodeChild
'                    '----------------------------------
'                    ' Loop Through ALL Valid Junctions
'                    '----------------------------------
'2912:                     For pJunctionCounter = 0 To pEdgeConnectivityRule.JunctionCount - 1 Step 1
'2913:                         Set pFeatureClassJunction = pFeatureClassContainer.ClassByID(pEdgeConnectivityRule.JunctionClassID(pJunctionCounter))
'2914:                         Set pDatasetJunction = pFeatureClassJunction
'2915:                         Set pSubtypesJunction = pFeatureClassJunction
'2916:                         pSubtypeJunctionName = pSubtypesJunction.SubtypeName(pEdgeConnectivityRule.JunctionSubtypeCode(pJunctionCounter))
'                        '------------------------------------------
'                        ' Add VIA Junction FC
'                        ' <junction database="" owner="" table="">
'                        '------------------------------------------
'2921:                         Set pSQLSyntax = pDatasetJunction.Workspace
'2922:                         Call pSQLSyntax.ParseTableName(CStr(pDatasetJunction.Name), pDatabaseName, pOwnerName, pTableName)
'2923:                         Set pXMLDOMNodeList = pXMLDOMNodeJunctionConnRule.selectNodes("junction[@database='" & pDatabaseName & "'" & _
'                                                                                      " and @owner='" & pOwnerName & "'" & _
'                                                                                      " and @table='" & pTableName & "']")
'2926:                         If pXMLDOMNodeList.length = 0 Then
'                            '--------------
'                            ' Add New Node
'                            '--------------
'2930:                             Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "junction", "")
'2931:                             Set pXMLDOMNodeChild = pXMLDOMNodeParent.appendChild(pXMLDOMNodeChild)
'                            '
'2933:                             Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "database", pDatabaseName)
'2934:                             Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "owner", pOwnerName)
'2935:                             Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "table", pTableName)
'                            '
'2937:                             Set pXMLDOMNodeParent2 = pXMLDOMNodeChild
'2938:                         Else
'2939:                             Set pXMLDOMNodeParent2 = pXMLDOMNodeList.nextNode
'2940:                         End If
'                        '----------------------------------------
'                        ' Add VIA Junction Subtype
'                        ' <subtype name="" default="" />
'                        '----------------------------------------
'2945:                         Set pXMLDOMNodeChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'2946:                         Set pXMLDOMNodeChild = pXMLDOMNodeParent2.appendChild(pXMLDOMNodeChild)
'                        '
'2948:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "name", pSubtypeJunctionName)
'2949:                         If pEdgeConnectivityRule.JunctionClassID(pJunctionCounter) = pEdgeConnectivityRule.DefaultJunctionClassID And _
'                           pEdgeConnectivityRule.JunctionSubtypeCode(pJunctionCounter) = pEdgeConnectivityRule.DefaultJunctionSubtypeCode Then
'                            '--------------------------
'                            ' Default Junction/Subtype
'                            '--------------------------
'2954:                             Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "default", "True")
'2955:                         Else
'2956:                             Call modCommon.AddNodeAttribute(pXMLDOMNodeChild, "default", "False")
'2957:                         End If
'2958:                     Next pJunctionCounter
'2959:                 End If
'2960:             Else
'                ' ===>>> ERROR <<<===
'2962:             End If
'2963:         End If
'2964:         Set pRule = pEnumRule.Next
'2965:     Loop
'    '
'    Exit Sub
'ErrorRecovery:
'2969:     Call ErrorRecovery("ExportGeometricNework2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'2970:     Resume Next
'ErrorHandler:
'2972:     Call HandleError(False, "ExportGeometricNework2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub WriteNetworkClassToXML(ByRef pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode, _
'                                   ByRef pGeometricNetwork As esriGeoDatabase.IGeometricNetwork, _
'                                   ByRef pEsriFeatureType As esriGeoDatabase.esriFeatureType)
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnumFeatureClass As esriGeoDatabase.IEnumFeatureClass
'    Dim pNetworkClass As esriGeoDatabase.INetworkClass
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    '
'2985:     Set pEnumFeatureClass = pGeometricNetwork.ClassesByType(pEsriFeatureType)
'2986:     Set pNetworkClass = pEnumFeatureClass.Next
'    '
'2988:     Do Until pNetworkClass Is Nothing
'        '------------------------------------------------------------------------------
'        ' Add GN FeatureClasses to XML
'        ' <featureClass name="" esriFeatureType="" esriNetworkClassAncillaryRole="" />
'        '------------------------------------------------------------------------------
'2993:         If pNetworkClass.FeatureClassID = pGeometricNetwork.OrphanJunctionFeatureClass.FeatureClassID Then
'            '----------------------------------
'            ' Skip Writing Orphan FeatureClass
'            '----------------------------------
'2997:         Else
'2998:             Set pXMLDOMNodeFeatureClass = pXMLDOMNodeGeometricNetwork.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "featureClass", "")
'2999:             Set pXMLDOMNodeFeatureClass = pXMLDOMNodeGeometricNetwork.appendChild(pXMLDOMNodeFeatureClass)
'3000:             Set pDataset = pNetworkClass
'3001:             Call AddQualifiedTableNameParts(pXMLDOMNodeFeatureClass, pDataset)
'3002:             Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "esriFeatureType", CStr(pEsriFeatureType))
'3003:             If pEsriFeatureType = esriFTSimpleJunction Then
'3004:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "esriNetworkClassAncillaryRole", CStr(pNetworkClass.NetworkAncillaryRole))
'3005:             Else
'3006:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "esriNetworkClassAncillaryRole", "")
'3007:             End If
'3008:         End If
'        '
'3010:         Set pNetworkClass = pEnumFeatureClass.Next
'3011:     Loop
'    '
'    Exit Sub
'ErrorHandler:
'3015:     Call HandleError(False, "WriteNetworkClassToXML " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

''==================================================================================

''----------------------------
'' IMPORT - GEOMETRIC NETWORK
''----------------------------

'Public Sub ImportGeometricNetwork(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                                  ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '------------------------------------------------------------------------
'    ' Main Routine. This routine is called when the user clicks "Import GN"
'    ' If a FeatureDataset is selected:   (1) Create NEW GeometricNetwork
'    '                                    (2) Add NEW Connectivity Rules
'    ' If a GeometricNetwork is selected: (1) Add NEW Connectivity Rules
'    '------------------------------------------------------------------------
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListGeometricNetwork As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    '
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pGeometricNetwork As esriGeoDatabase.IGeometricNetwork
'    Dim pFeatureDataset As esriGeoDatabase.IFeatureDataset
'    '
'    Dim pIndex As Long
'    Dim strGeometricNetworkName As String
'    Dim strFeatureClassName As String
'    '
'3046:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'3047:     Set pXMLDOMNodeListGeometricNetwork = pXMLDOMNodeGeodatabaseDesigner.selectNodes("geometricNetwork")
'    '---------------------------------------------------
'    ' Loop Through Each GeometricNetwork in the XML DOM
'    '---------------------------------------------------
'3051:     For pIndex = 0 To pXMLDOMNodeListGeometricNetwork.length - 1 Step 1
'3052:         Set pXMLDOMNodeGeometricNetwork = pXMLDOMNodeListGeometricNetwork.Item(pIndex)
'3053:         Set pXMLDOMNodeFeatureClass = pXMLDOMNodeGeometricNetwork.selectNodes("featureClass").Item(0)
'        '
'3055:         strGeometricNetworkName = CStr(pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("table").Text)
'3056:         strFeatureClassName = CStr(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("table").Text)
'        '
'3058:         Call frmOutputWindow.AddOutputMessage("Importing GeometricNetwork: " & strGeometricNetworkName, , , vbMagenta)
'        '
'3060:         Set pGeometricNetwork = modCommon.GetDataset(pWorkspace, strGeometricNetworkName, esriDTGeometricNetwork)
'        '
'3062:         If pGeometricNetwork Is Nothing Then
'3063:             Set pFeatureClass = modCommon.GetDataset(pWorkspace, strFeatureClassName, esriDTFeatureClass)
'3064:             Set pFeatureDataset = pFeatureClass.FeatureDataset
'3065:             Set pGeometricNetwork = CreateGeometricNetwork(pXMLDOMNodeGeometricNetwork, pFeatureDataset)
'            '
'3067:             If pGeometricNetwork Is Nothing Then
'3068:                 Call frmOutputWindow.AddOutputMessage("GeometricNetwork Creation Failed", , , vbRed, , , , 1)
'3069:             Else
'3070:                 Call frmOutputWindow.AddOutputMessage("Importing Connectivity Rules", , , , , , True, 1)
'3071:                 Call ImportGeometricNetworkConnectivityRules(pXMLDOMNodeGeometricNetwork, pGeometricNetwork)
'                ' Tomasz Czerny:
'                ' I've moved it from NetworkEdit to NetworkLoader so it now can import existing weights data to network.
'                'Call frmOutputWindow.AddOutputMessage("Importing Network Weights", , , , , , True, 1)
'                'Call ImportGeometricNetworkWeights(pXMLDOMNodeGeometricNetwork, pGeometricNetwork)
'3076:             End If
'3077:         Else
'3078:             Call frmOutputWindow.AddOutputMessage("GeometricNetwork Already Exisits", , , , , , True, 1)
'3079:             Call frmOutputWindow.AddOutputMessage("Importing Connectivity Rules", , , , , , True, 1)
'3080:             Call ImportGeometricNetworkConnectivityRules(pXMLDOMNodeGeometricNetwork, pGeometricNetwork)
'3081:         End If
'3082:     Next pIndex
'    '
'    Exit Sub
'ErrorHandler:
'3086:     Call HandleError(False, "ImportGeometricNetwork " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function CreateGeometricNetwork(ByRef pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode, _
'                                        ByRef pFeatureDataset As esriGeoDatabase.IFeatureDataset) As esriGeoDatabase.IGeometricNetwork
'    On Error GoTo ErrorHandler
'    '----------------------------------------------------------
'    ' Creates a NEW GeometricNetwork inside the FeatureDataset
'    '----------------------------------------------------------
'    Dim pNetworkLoader As esriNetworkAnalysis.INetworkLoader
'    Dim pNetworkLoader2 As esriNetworkAnalysis.INetworkLoader2
'    Dim pNetworkLoaderProps As esriNetworkAnalysis.INetworkLoaderProps
'    '
'    Dim pGeometricNetwork As esriGeoDatabase.IGeometricNetwork
'    Dim pNetworkCollection As esriGeoDatabase.INetworkCollection
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pFeatureDatasetName As esriGeoDatabase.IFeatureDatasetName
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pDataset As esriGeoDatabase.IDataset
'    '
'    Dim pUIDComplexEdge As esriSystem.UID
'    Dim pUIDSimpleEdge As esriSystem.UID
'    Dim pUIDSimpleJunction As esriSystem.UID
'    '
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListFeatureClass As MSXML2.IXMLDOMNodeList
'    '
'    Dim pIndex As Long
'    Dim pConfigKeyword As String
'    '----------------
'    ' Set Interfaces
'    '----------------
'3118:     Set pGeometricNetwork = Nothing
'3119:     Set pNetworkCollection = pFeatureDataset
'3120:     Set pFeatureClassContainer = pFeatureDataset
'    '----------------------
'    ' Create a new network
'    '----------------------
'3124:     Set pNetworkLoader = New NetworkLoader
'3125:     Set pNetworkLoader2 = pNetworkLoader
'3126:     Set pNetworkLoaderProps = pNetworkLoader
'3127:     Set pFeatureDatasetName = pFeatureDataset.FullName
'3128:     Set pNetworkLoader.FeatureDatasetName = pFeatureDatasetName
'3129:     pNetworkLoader.NetworkName = CStr(pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("table").Text)
'3130:     pNetworkLoader.NetworkType = CLng(pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("esriNetworkType").Text)
'3131:     If pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("configKeyword") Is Nothing Then
'        '--------------------------
'        ' No Configuration Keyword
'        '--------------------------
'3135:     Else
'3136:         If pFeatureDataset.Workspace.Type = esriRemoteDatabaseWorkspace Then
'3137:             pConfigKeyword = CStr(pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("configKeyword").Text)
'3138:             If pConfigKeyword <> "" Then
'3139:                 pNetworkLoader2.ConfigurationKeyword = pConfigKeyword
'3140:                 Call frmOutputWindow.AddOutputMessage("Assigning Configuration Keyword [" & pConfigKeyword & "]", , , , , , True, 1)
'3141:             End If
'3142:         End If
'3143:     End If
'    '-----------------------
'    ' Create new Class ID's
'    '-----------------------
'3147:     Set pUIDComplexEdge = New esriSystem.UID
'3148:     Set pUIDSimpleEdge = New esriSystem.UID
'3149:     Set pUIDSimpleJunction = New esriSystem.UID
'3150:     pUIDComplexEdge.Value = CStr(GUID_COMPLEXEDGE_CLSID)
'3151:     pUIDSimpleEdge.Value = CStr(GUID_SIMPLEEDGE_CLSID)
'3152:     pUIDSimpleJunction.Value = CStr(GUID_SIMPLEJUNCTION_CLSID)
'    '----------------------------------
'    ' Set snapping tolerance (if used)
'    '----------------------------------
'    Select Case mGNImportSnapping
'    Case -1
'3158:         pNetworkLoader.SnapTolerance = pNetworkLoader2.MinSnapTolerance
'    Case 0
'3160:         pNetworkLoader.SnapTolerance = pNetworkLoader2.MinSnapTolerance
'    Case Else
'3162:         pNetworkLoader.SnapTolerance = mGNImportSnapping
'3163:     End Select
'     '--------------------------------
'    ' Loop through each FeatureClass
'    '--------------------------------
'3167:     Set pXMLDOMNodeListFeatureClass = pXMLDOMNodeGeometricNetwork.selectNodes("featureClass")
'3168:     For pIndex = 0 To pXMLDOMNodeListFeatureClass.length - 1 Step 1
'        '-------------------------
'        ' Get Source FeatureClass
'        '-------------------------
'3172:         Set pXMLDOMNodeFeatureClass = pXMLDOMNodeListFeatureClass.Item(pIndex)
'3173:         Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("table").Text)
'3174:         Set pDataset = pFeatureClass
'        '-------------
'        ' Add Message
'        '-------------
'3178:         Call frmOutputWindow.AddOutputMessage("Importing NetworkClass: " & pDataset.Name, , , , , , True, 1)
'        '-----------------------------------
'        ' Check if FeatureClass is suitable
'        '-----------------------------------
'        Select Case CLng(pNetworkLoader2.CanUseFeatureClass(pDataset.Name))
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCCannotOpen
'3184:             Call frmOutputWindow.AddOutputMessage("FeatureClass cannot be opened.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCInAnotherNetwork
'3187:             Call frmOutputWindow.AddOutputMessage("FeatureClass is already in another network.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCInvalidFeatureType
'3190:             Call frmOutputWindow.AddOutputMessage("FeatureClass has invalid type.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCInvalidShapeType
'3193:             Call frmOutputWindow.AddOutputMessage("FeatureClass does not have point or line geometry.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCRegisteredAsVersioned
'3196:             Call frmOutputWindow.AddOutputMessage("FeatureClass is registered as versioned.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCUnknownError
'3199:             Call frmOutputWindow.AddOutputMessage("An unknown error was encountered.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNetworkAnalysis.esriNetworkLoaderFeatureClassCheck.esriNLFCCValid
'            '-------------------------------------------------------
'            ' The given feature class can participate in a network.
'            '-------------------------------------------------------
'3205:             Call frmOutputWindow.AddOutputMessage("FeatureClass can be added to a Network", , , , , , True, 2)
'        Case Else
'            ' Do Nothing
'3208:         End Select
'        '-------------------------------------------------------
'        ' Add FeatureClass to GeometricNetwork
'        ' FeatureClasses will snap if (mGNImportSnapping <> -1)
'        '-------------------------------------------------------
'        Select Case CLng(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("esriFeatureType").Text)
'        Case esriFTSimpleJunction
'3215:             Call pNetworkLoader.AddFeatureClass(pDataset.Name, esriFTSimpleJunction, pUIDSimpleJunction, CBool(mGNImportSnapping <> -1))
'            '--------------------------------------------------------------------
'            ' Check if Junctions FeatureClasses have an existing Ancillary Field
'            '--------------------------------------------------------------------
'            Select Case CLng(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("esriNetworkClassAncillaryRole").Text)
'            Case esriNCARNone
'                '--------------------------------------------
'                ' This Junction FC will not be a Source/Sink
'                '--------------------------------------------
'            Case esriNCARSourceSink
'                '------------------------------------------------
'                ' Check if Junction FC already has a ROLE field.
'                '------------------------------------------------
'                Select Case pNetworkLoader2.CheckAncillaryRoleField(pDataset.Name, pNetworkLoaderProps.DefaultAncillaryRoleField)
'                Case esriNLFCValid
'                    '-------------------------------------
'                    ' AncillaryRole field found and is OK.
'                    '-------------------------------------
'3233:                     Call frmOutputWindow.AddOutputMessage("The AncillaryRole field already exists.", , , , , , True, 2)
'3234:                     Call pNetworkLoader.PutAncillaryRole(pDataset.Name, esriNCARSourceSink, pNetworkLoaderProps.DefaultAncillaryRoleField)
'                Case esriNLFCNotFound
'                    '------------------------------------------------------
'                    ' AncillaryRole field not found. Safe to add the field.
'                    '------------------------------------------------------
'3239:                     Call frmOutputWindow.AddOutputMessage("Adding AncillaryRole field to FeatureClass.", , , , , , True, 2)
'3240:                     Call pNetworkLoader.PutAncillaryRole(pDataset.Name, esriNCARSourceSink, pNetworkLoaderProps.DefaultAncillaryRoleField)
'                Case esriNLFCInvalidType
'3242:                     Call frmOutputWindow.AddOutputMessage("The AncillaryRole field has invalid type.", , , vbRed, , , True, 2)
'                    Exit Function
'                Case esriNLFCInvalidDomain
'3245:                     Call frmOutputWindow.AddOutputMessage("The AncillaryRole field has invalid domain.", , , vbRed, , , True, 2)
'                    Exit Function
'                Case esriNLFCUnknownError
'3248:                     Call frmOutputWindow.AddOutputMessage("(ROLE Field)- An unknown error was encountered.", , , vbRed, , , True, 2)
'                    Exit Function
'                Case Else
'3251:                     Call frmOutputWindow.AddOutputMessage("(ROLE Field)- An unknown error was encountered.", , , vbRed, , , True, 2)
'                    Exit Function
'3253:                 End Select
'            Case Else
'3255:                 Call frmOutputWindow.AddOutputMessage("XML Error: Invalid [esriNetworkClassAncillaryRole]", , , vbRed, , , True, 2)
'                Exit Function
'3257:             End Select
'        Case esriFTSimpleEdge
'3259:             Call pNetworkLoader.AddFeatureClass(pDataset.Name, esriFTSimpleEdge, pUIDSimpleEdge, CBool(mGNImportSnapping <> -1))
'        Case esriFTComplexEdge
'3261:             Call pNetworkLoader.AddFeatureClass(pDataset.Name, esriFTComplexEdge, pUIDComplexEdge, CBool(mGNImportSnapping <> -1))
'        Case Else
'3263:             Call frmOutputWindow.AddOutputMessage("XML Error: Invalid [esriFeatureType]", , , vbRed, , , True, 2)
'            Exit Function
'3265:         End Select
'        '--------------------------------------------------
'        ' Check if FeatureClass has a valid Enabled Field.
'        '--------------------------------------------------
'        Select Case pNetworkLoader2.CheckEnabledDisabledField(pDataset.Name, pNetworkLoaderProps.DefaultEnabledField)
'        Case esriNLFCValid
'            '----------------------------------------------------
'            ' Enabled field found and is OK: Action not required
'            '----------------------------------------------------
'3274:             Call frmOutputWindow.AddOutputMessage("ENABLED field already exists.", , , , , , True, 2)
'        Case esriNLFCNotFound
'            '-------------------------------------------------
'            ' Enabled field not found. Safe to add the field.
'            '-------------------------------------------------
'3279:             Call frmOutputWindow.AddOutputMessage("Adding ENABLED field.", , , , , , True, 2)
'        Case esriNLFCInvalidType
'3281:             Call frmOutputWindow.AddOutputMessage("The ENABLED field has invalid type.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNLFCInvalidDomain
'3284:             Call frmOutputWindow.AddOutputMessage("The ENABLED field has invalid domain.", , , vbRed, , , True, 2)
'            Exit Function
'        Case esriNLFCUnknownError
'3287:             Call frmOutputWindow.AddOutputMessage("(ENABLED Field)- An unknown error was encountered.", , , vbRed, , , True, 2)
'            Exit Function
'        Case Else
'3290:             Call frmOutputWindow.AddOutputMessage("(ENABLED Field)- An unknown error was encountered.", , , vbRed, , , True, 2)
'            Exit Function
'3292:         End Select
'3293:     Next pIndex
'    '------------------------------------------
'    ' If already then preserve existing values
'    '------------------------------------------
'3297:     pNetworkLoader2.PreserveEnabledValues = mGNImportPreserveEnabledValue
'    '---------------------------
'    ' Added by Tomasz Czerny
'    ' Import Network Weights....
'    '---------------------------
'3302:     Call frmOutputWindow.AddOutputMessage("Importing Network Weights", , , , , , True, 1)
'3303:     Call ImportGeometricNetworkWeights(pXMLDOMNodeGeometricNetwork, pNetworkLoader, pFeatureDataset)
'    '--------------------------
'    'Print some detail messages
'    '--------------------------
'3307:     Set frmOutputWindow.pNetworkLoaderEvents = pNetworkLoader
'3308:     Call frmOutputWindow.AddOutputMessage("Start loading network", , , , , , True, 1)
'    '------------------------------------------------------
'    ' OK, Lets load then Network with network loader events
'    '------------------------------------------------------
'3312:     Call pNetworkLoader.LoadNetwork
'    '--------------------------
'    ' Clear networkloaderevents
'    '--------------------------
'3316:     Set frmOutputWindow.pNetworkLoaderEvents = Nothing
'    '---------------------------------
'    ' Return the NEW GeometricNetwork
'    '---------------------------------
'3320:     Set CreateGeometricNetwork = pNetworkCollection.GeometricNetworkByName(pXMLDOMNodeGeometricNetwork.Attributes.getNamedItem("table").Text)
'    '-----------------
'    ' Clear Variables
'    '-----------------
'3324:     Set pNetworkLoader = Nothing
'3325:     Set pNetworkLoader2 = Nothing
'3326:     Set pNetworkLoaderProps = Nothing
'    '
'3328:     Set pGeometricNetwork = Nothing
'3329:     Set pNetworkCollection = Nothing
'3330:     Set pFeatureClassContainer = Nothing
'3331:     Set pFeatureDatasetName = Nothing
'3332:     Set pFeatureClass = Nothing
'3333:     Set pDataset = Nothing
'    '
'3335:     Set pUIDComplexEdge = Nothing
'3336:     Set pUIDSimpleEdge = Nothing
'3337:     Set pUIDSimpleJunction = Nothing
'    '
'    Exit Function
'ErrorHandler:
'3341:     Call HandleError(False, "CreateGeometricNetwork " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Sub ImportGeometricNetworkConnectivityRules(ByRef pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode, _
'                                                    ByRef pGeometricNetwork As esriGeoDatabase.IGeometricNetwork)
'    On Error GoTo ErrorHandler
'    '------------------------------------------------------------------------------------------
'    ' This Routine will recreate a GeometricNetwork (and Connectivity Rules) from an XML file.
'    '------------------------------------------------------------------------------------------
'    Dim pJunctionConnectivityRule2 As esriGeoDatabase.IJunctionConnectivityRule2
'    Dim pEdgeConnectivityRule As esriGeoDatabase.IEdgeConnectivityRule
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    '
'    Dim pXMLDOMNodeEdgeList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeEdgeSubtypeList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeJunctionList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeJunctionSubtypeList As MSXML2.IXMLDOMNodeList
'    '
'    Dim pXMLDOMNodeJunctionConnRule As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdgeConnRule As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdge As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdgeSubtype As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeJunction As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeJunctionSubtype As MSXML2.IXMLDOMNode
'    '
'    Dim pXMLDOMNodeEdge1 As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdge2 As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdgeSubtype1 As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeEdgeSubtype2 As MSXML2.IXMLDOMNode

'    Dim pXMLDOMNodeEdgeList1 As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeEdgeList2 As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeEdgeSubtypeList1 As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeEdgeSubtypeList2 As MSXML2.IXMLDOMNodeList
'    '
'    Dim pIndexEdge As Long
'    Dim pIndexEdgeSubtype As Long
'    Dim pIndexJunction As Long
'    Dim pIndexJunctionSubtype As Long
'    '
'    Dim pIndexEdge1 As Long
'    Dim pIndexEdge2 As Long
'    Dim pIndexEdgeSubtype1 As Long
'    Dim pIndexEdgeSubtype2 As Long
'    '----------------------------------------
'    ' Delete all existing Connectivity Rules
'    '----------------------------------------
'3389:     If mGNImportClearConnRule Then
'3390:         Call RemoveConnectivityRules(pGeometricNetwork)
'3391:     End If
'    '------------------------------------------------------------------------------------------------------
'    ' QI to IFeatureClassContainer. This will be used to get the ObjectClass ID from the FeatureClass name
'    '------------------------------------------------------------------------------------------------------
'3395:     Set pFeatureClassContainer = pGeometricNetwork
'    '-------------------------------------
'    ' Junction Connectivity Rules
'    ' Get <junctionConnectivityRule> Node
'    '-------------------------------------
'3400:     Set pXMLDOMNodeJunctionConnRule = pXMLDOMNodeGeometricNetwork.selectSingleNode("junctionConnectivityRule")
'3401:     Set pXMLDOMNodeEdgeList = pXMLDOMNodeJunctionConnRule.selectNodes("edge")
'    '--------------------------------------
'    ' Loop through each EDGE
'    ' <edge database="" owner="" table="">
'    '--------------------------------------
'3406:     For pIndexEdge = 0 To pXMLDOMNodeEdgeList.length - 1 Step 1
'3407:         Set pXMLDOMNodeEdge = pXMLDOMNodeEdgeList.Item(pIndexEdge)
'3408:         Set pXMLDOMNodeEdgeSubtypeList = pXMLDOMNodeEdge.childNodes
'        '--------------------------------
'        ' Loop through each EDGE SUBTYPE
'        ' <subtype name="Cable TV Line">
'        '--------------------------------
'3413:         For pIndexEdgeSubtype = 0 To pXMLDOMNodeEdgeSubtypeList.length - 1 Step 1
'3414:             Set pXMLDOMNodeEdgeSubtype = pXMLDOMNodeEdgeSubtypeList.Item(pIndexEdgeSubtype)
'3415:             Set pXMLDOMNodeJunctionList = pXMLDOMNodeEdgeSubtype.childNodes
'            '------------------------------------------
'            ' Loop through each JUNCTION
'            ' <junction database="" owner="" table="">
'            '------------------------------------------
'3420:             For pIndexJunction = 0 To pXMLDOMNodeJunctionList.length - 1 Step 1
'3421:                 Set pXMLDOMNodeJunction = pXMLDOMNodeJunctionList.Item(pIndexJunction)
'3422:                 Set pXMLDOMNodeJunctionSubtypeList = pXMLDOMNodeJunction.childNodes
'                '------------------------------------------------------------
'                ' Loop through each JUNCTION SUBTYPE
'                ' <subtype name="Service Pillar"
'                '  eMin="-1" eMax="-1" jMin="-1" jMax="-1" default="True" />
'                '------------------------------------------------------------
'3428:                 For pIndexJunctionSubtype = 0 To pXMLDOMNodeJunctionSubtypeList.length - 1 Step 1
'3429:                     Set pXMLDOMNodeJunctionSubtype = pXMLDOMNodeJunctionSubtypeList.Item(pIndexJunctionSubtype)
'                    '----------------------------------------------------
'                    ' Add Junction Connectivity Rule to GeometricNetwork
'                    '----------------------------------------------------
'3433:                     Set pJunctionConnectivityRule2 = New JunctionConnectivityRule
'                    '-----------------------------------
'                    ' Set Edge FeatureClass and Subtype
'                    '-----------------------------------
'3437:                     Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeEdge.Attributes.getNamedItem("table").Text)
'3438:                     pJunctionConnectivityRule2.EdgeClassID = pFeatureClass.FeatureClassID
'3439:                     pJunctionConnectivityRule2.EdgeSubtypeCode = GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeEdgeSubtype.Attributes.getNamedItem("name").Text)
'                    '---------------------------------------
'                    ' Set Junction FeatureClass and Subtype
'                    '---------------------------------------
'3443:                     Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeJunction.Attributes.getNamedItem("table").Text)
'3444:                     pJunctionConnectivityRule2.JunctionClassID = pFeatureClass.FeatureClassID
'3445:                     pJunctionConnectivityRule2.JunctionSubtypeCode = GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("name").Text)
'                    '-----------------------------------
'                    ' Set Edge and Junction Cardinality
'                    '-----------------------------------
'3449:                     pJunctionConnectivityRule2.EdgeMinimumCardinality = CLng(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("eMin").Text)
'3450:                     pJunctionConnectivityRule2.EdgeMaximumCardinality = CLng(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("eMax").Text)
'3451:                     pJunctionConnectivityRule2.JunctionMinimumCardinality = CLng(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("jMin").Text)
'3452:                     pJunctionConnectivityRule2.JunctionMaximumCardinality = CLng(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("jMax").Text)
'                    '----------------------
'                    ' Set Default Junction
'                    '----------------------
'3456:                     If CBool(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("default").Text) Then
'3457:                         pJunctionConnectivityRule2.DefaultJunction = True
'3458:                     End If
'                    '-----------------
'                    ' Display Message
'                    '-----------------
'3462:                     Call frmOutputWindow.AddOutputMessage("Importing Junction Rules:", , , , , , True, 2)
'3463:                     Call frmOutputWindow.AddOutputMessage("Edge FeatureClass[Subtype]: " & CStr(pXMLDOMNodeEdge.Attributes.getNamedItem("table").Text) & "[" & CStr(pXMLDOMNodeEdgeSubtype.Attributes.getNamedItem("name").Text) & "]", , , , , , True, 3)
'3464:                     Call frmOutputWindow.AddOutputMessage("Junction FeatureClass[Subtype]: " & CStr(pXMLDOMNodeJunction.Attributes.getNamedItem("table").Text) & "[" & CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("name").Text) & "]", , , , , , True, 3)
'3465:                     Call frmOutputWindow.AddOutputMessage("Rules [Emin|EMax|JMin|JMax]: " & _
'                        "[" & _
'                        CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("eMin").Text) & "|" & _
'                        CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("eMax").Text) & "|" & _
'                        CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("jMin").Text) & "|" & _
'                        CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("jMax").Text) & _
'                        "]" _
'                        , , , , , , True, 3)
'                    '------------------------------
'                    ' Add Rule to GeometricNetwork
'                    '------------------------------
'3476:                     Call pGeometricNetwork.AddRule(pJunctionConnectivityRule2)
'3477:                 Next pIndexJunctionSubtype
'3478:              Next pIndexJunction
'3479:         Next pIndexEdgeSubtype
'3480:     Next pIndexEdge
'    '-------------------------
'    ' Edge Connectivity Rules
'    ' <edgeConnectivityRule>
'    '-------------------------
'3485:     Set pXMLDOMNodeEdgeConnRule = pXMLDOMNodeGeometricNetwork.selectSingleNode("edgeConnectivityRule")
'3486:     Set pXMLDOMNodeEdgeList1 = pXMLDOMNodeEdgeConnRule.selectNodes("edge")
'    '--------------------------------------
'    ' Loop through each FROM EDGE 1
'    ' <edge database="" owner="" table="">
'    '--------------------------------------
'3491:     For pIndexEdge1 = 0 To pXMLDOMNodeEdgeList1.length - 1 Step 1
'3492:         Set pXMLDOMNodeEdge1 = pXMLDOMNodeEdgeList1.Item(pIndexEdge1)
'3493:         Set pXMLDOMNodeEdgeSubtypeList1 = pXMLDOMNodeEdge1.childNodes
'        '----------------------------------
'        ' Loop through each EDGE SUBTYPE 1
'        ' <subtype name="Cable TV Line">
'        '----------------------------------
'3498:         For pIndexEdgeSubtype1 = 0 To pXMLDOMNodeEdgeSubtypeList1.length - 1 Step 1
'3499:             Set pXMLDOMNodeEdgeSubtype1 = pXMLDOMNodeEdgeSubtypeList1.Item(pIndexEdgeSubtype1)
'3500:             Set pXMLDOMNodeEdgeList2 = pXMLDOMNodeEdgeSubtype1.childNodes
'            '------------------------------------------
'            ' Loop through each EDGE 2
'            ' <junction database="" owner="" table="">
'            '------------------------------------------
'3505:             For pIndexEdge2 = 0 To pXMLDOMNodeEdgeList2.length - 1 Step 1
'3506:                 Set pXMLDOMNodeEdge2 = pXMLDOMNodeEdgeList2.Item(pIndexEdge2)
'3507:                 Set pXMLDOMNodeEdgeSubtypeList2 = pXMLDOMNodeEdge2.childNodes
'                '----------------------------------
'                ' Loop through each EDGE SUBTYPE 2
'                ' <subtype name="Cable TV Line">
'                '----------------------------------
'3512:                 For pIndexEdgeSubtype2 = 0 To pXMLDOMNodeEdgeSubtypeList2.length - 1 Step 1
'3513:                     Set pXMLDOMNodeEdgeSubtype2 = pXMLDOMNodeEdgeSubtypeList2.Item(pIndexEdgeSubtype2)
'3514:                     Set pXMLDOMNodeJunctionList = pXMLDOMNodeEdgeSubtype2.childNodes
'                    '-------------------------------------------------------------
'                    ' Create a NEW Edge Connectivity Rule - Add FROM and TO Edges
'                    '-------------------------------------------------------------
'3518:                     Set pEdgeConnectivityRule = New esriGeoDatabase.EdgeConnectivityRule
'                    '----------------------------------------
'                    ' Set From Edge FeatureClass and Subtype
'                    '----------------------------------------
'3522:                     Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeEdge1.Attributes.getNamedItem("table").Text)
'3523:                     pEdgeConnectivityRule.FromEdgeClassID = pFeatureClass.FeatureClassID
'3524:                     pEdgeConnectivityRule.FromEdgeSubtypeCode = GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeEdgeSubtype1.Attributes.getNamedItem("name").Text)
'                    '--------------------------------------
'                    ' Set To Edge FeatureClass and Subtype
'                    '--------------------------------------
'3528:                     Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeEdge2.Attributes.getNamedItem("table").Text)
'3529:                     pEdgeConnectivityRule.ToEdgeClassID = pFeatureClass.FeatureClassID
'3530:                     pEdgeConnectivityRule.ToEdgeSubtypeCode = GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeEdgeSubtype2.Attributes.getNamedItem("name").Text)
'                    '-----------------
'                    ' Display Message
'                    '-----------------
'3534:                     Call frmOutputWindow.AddOutputMessage("Importing Edge Rules:", , , , , , True, 2)
'3535:                     Call frmOutputWindow.AddOutputMessage("From Edge FeatureClass[Subtype]: " & CStr(pXMLDOMNodeEdge1.Attributes.getNamedItem("table").Text) & "[" & CStr(pXMLDOMNodeEdgeSubtype1.Attributes.getNamedItem("name").Text) & "]", , , , , , True, 3)
'3536:                     Call frmOutputWindow.AddOutputMessage("To   Edge FeatureClass[Subtype]: " & CStr(pXMLDOMNodeEdge2.Attributes.getNamedItem("table").Text) & "[" & CStr(pXMLDOMNodeEdgeSubtype2.Attributes.getNamedItem("name").Text) & "]", , , , , , True, 3)
'                    '-----------------------------------------------
'                    ' Loop through each JUNCTION
'                    ' <junction database="" owner="" table="">
'                    '-----------------------------------------------
'3541:                     For pIndexJunction = 0 To pXMLDOMNodeJunctionList.length - 1 Step 1
'3542:                         Set pXMLDOMNodeJunction = pXMLDOMNodeJunctionList.Item(pIndexJunction)
'3543:                         Set pXMLDOMNodeJunctionSubtypeList = pXMLDOMNodeJunction.childNodes
'3544:                         Set pFeatureClass = pFeatureClassContainer.ClassByName(pXMLDOMNodeJunction.Attributes.getNamedItem("table").Text)
'                        '------------------------------------------------------
'                        ' Loop through each JUNCTION SUBTYPE
'                        ' <subtype name="EF_DetailedDrawing" default="True" />
'                        '------------------------------------------------------
'3549:                         For pIndexJunctionSubtype = 0 To pXMLDOMNodeJunctionSubtypeList.length - 1 Step 1
'3550:                             Set pXMLDOMNodeJunctionSubtype = pXMLDOMNodeJunctionSubtypeList.Item(pIndexJunctionSubtype)
'                            '
'3552:                             Call pEdgeConnectivityRule.AddJunction(pFeatureClass.FeatureClassID, GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("name").Text))
'                            '
'3554:                             If CBool(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("default").Text) Then
'3555:                                 pEdgeConnectivityRule.DefaultJunctionClassID = pFeatureClass.FeatureClassID
'3556:                                 pEdgeConnectivityRule.DefaultJunctionSubtypeCode = GetSubtypeCodeFromName(pFeatureClass, pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("name").Text)
'3557:                             End If
'                            '-----------------
'                            ' Display Message
'                            '-----------------
'3561:                             Call frmOutputWindow.AddOutputMessage("Via Junction FeatureClass[Subtype]: " & CStr(pXMLDOMNodeJunction.Attributes.getNamedItem("table").Text) & "[" & CStr(pXMLDOMNodeJunctionSubtype.Attributes.getNamedItem("name").Text) & "]", , , , , , True, 4)
'3562:                         Next pIndexJunctionSubtype
'3563:                     Next pIndexJunction
'                    '------------------------------------------------
'                    ' Add Edge Connectivity Rule to GeometricNetwork
'                    '------------------------------------------------
'3567:                     Call pGeometricNetwork.AddRule(pEdgeConnectivityRule)
'                    '-----------------
'                    ' Display Message
'                    '-----------------
'3571:                     Call frmOutputWindow.AddOutputMessage("Importing Edge Rule [" & CStr(pEdgeConnectivityRule.id) & "]", , , , , , True, 2)
'3572:                 Next pIndexEdgeSubtype2
'3573:             Next pIndexEdge2
'3574:         Next pIndexEdgeSubtype1
'3575:     Next pIndexEdge1
'    '
'    Exit Sub
'ErrorHandler:
'3579:     Call HandleError(False, "ImportGeometricNetworkConnectivityRules " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ImportGeometricNetworkWeights(ByRef pXMLDOMNodeGeometricNetwork As MSXML2.IXMLDOMNode, _
'                                          ByRef pNetworkLoader As INetworkLoader, _
'                                          ByRef pFeatureDataset As IFeatureDataset)
'    On Error GoTo ErrorHandler
'    '-------------------------------------------------------
'    ' This Routine will add Weights to the GeometricNetwork
'    '-------------------------------------------------------
'    Dim pFeatureClassContainer As IFeatureClassContainer
'    Dim pDataset As IDataset
'3591:     Set pFeatureClassContainer = pFeatureDataset

'    Dim pXMLDOMNodeListWeight As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeListWeightAssociation As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeWeight As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeWeightAssociation As MSXML2.IXMLDOMNode
'    '
'    Dim pIndexWeight As Long
'    Dim pIndexWeightAssociation As Long
'    '-----------------------
'    ' Start Schema Updating
'    '-----------------------
'3603:     Set pXMLDOMNodeListWeight = pXMLDOMNodeGeometricNetwork.selectNodes("weight")
'    '---------------------------------------------------
'    ' Loop through all WEIGHTS in the XML file.
'    ' <weight name="" esriWeightType="" bitGateSize="">
'    '---------------------------------------------------
'3608:     For pIndexWeight = 0 To pXMLDOMNodeListWeight.length - 1 Step 1
'3609:         Set pXMLDOMNodeWeight = pXMLDOMNodeListWeight(pIndexWeight)
'        Dim sWeightName As String
'        Dim lWeightType As Long
'        Dim lBitGateSize As Long
'3613:         sWeightName = CStr(pXMLDOMNodeWeight.Attributes.getNamedItem("name").Text)
'3614:         lWeightType = CLng(pXMLDOMNodeWeight.Attributes.getNamedItem("esriWeightType").Text)
'3615:         If lWeightType = esriGeoDatabase.esriWeightType.esriWTBitGate Then
'3616:             lBitGateSize = CLng(pXMLDOMNodeWeight.Attributes.getNamedItem("bitGateSize").Text)
'3617:         End If
'        '------------
'        ' Add WEIGHT
'        '------------
'        'Call pNetSchemaEdit.AddWeight(pNetWeight)
'3622:         pNetworkLoader.AddWeight sWeightName, lWeightType, lBitGateSize
'3623:         Call frmOutputWindow.AddOutputMessage("Importing Network Weight [" & CStr(sWeightName) & "]", , , , , , True, 2)
'        '-------------------------------------------------
'        ' Loop through all WEIGHT ASSOCIATIONS
'        ' <association featureClass="" field=""/>
'        '-------------------------------------------------
'3628:         Set pXMLDOMNodeListWeightAssociation = pXMLDOMNodeWeight.selectNodes("association")
'3629:         For pIndexWeightAssociation = 0 To pXMLDOMNodeListWeightAssociation.length - 1 Step 1
'3630:             Set pXMLDOMNodeWeightAssociation = pXMLDOMNodeListWeightAssociation.Item(pIndexWeightAssociation)
'            Dim sTableName As String
'            Dim sFullTableName As String
'            Dim sFieldName As String
'3634:             sTableName = CStr(pXMLDOMNodeWeightAssociation.Attributes.getNamedItem("table").Text)
'3635:             Set pDataset = pFeatureClassContainer.ClassByName(sTableName)

'3637:             sFullTableName = pDataset.Name
'            '''test->sFullTableName = sTableName
'3639:             sFieldName = CStr(pXMLDOMNodeWeightAssociation.Attributes.getNamedItem("field").Text)
'            '------------------------
'            ' Add WEIGHT ASSOCIATION
'            '------------------------
'3643:             Call frmOutputWindow.AddOutputMessage("Adding Weight Association to Table:" & sFullTableName & " Column:" & sFieldName, , , , , , True, 1)
'3644:             pNetworkLoader.AddWeightAssociation sWeightName, sFullTableName, sFieldName
'3645:         Next pIndexWeightAssociation
'3646:     Next pIndexWeight
'    Exit Sub
'ErrorHandler:
'3649:     Call HandleError(False, "ImportGeometricNetworkWeights " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description & " " & sWeightName & " " & sTableName & " " & sFieldName)
'End Sub

''==================================================================================

''------------------------------------
'' EXPORT - Geodatabase ObjectClasses
''------------------------------------

'Public Sub ExportGDB(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                     ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnumDatasetName As esriGeoDatabase.IEnumDatasetName
'    Dim pDatasetName As esriGeoDatabase.IDatasetName
'    Dim pName As esriSystem.IName
'    Dim pFeatureDataset As esriGeoDatabase.IFeatureDataset
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pTable As esriGeoDatabase.ITable
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    '----------------------------
'    ' Get GeodatabaseDesigner Node
'    '------------------------------
'3673:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '------------------------------------------------
'    ' Iterate through each IDataset under IWorkspace
'    '------------------------------------------------
'3677:     Set pEnumDatasetName = pWorkspace.DatasetNames(esriDTAny)
'3678:     Call pEnumDatasetName.Reset
'3679:     Set pDatasetName = pEnumDatasetName.Next
'3680:     Do Until pDatasetName Is Nothing
'        Select Case pDatasetName.Type
'        Case esriGeoDatabase.esriDatasetType.esriDTFeatureDataset
'            '-----------------------------
'            ' Document the FeatureDataset
'            '-----------------------------
'3686:             Call frmOutputWindow.AddOutputMessage("Exporting FeatureDataset: " & pDatasetName.Name, , , vbMagenta)
'3687:             Set pName = pDatasetName
'3688:             Set pFeatureDataset = Nothing
'            On Error GoTo ErrorRecovery
'3690:             Set pFeatureDataset = pName.Open
'            On Error GoTo ErrorHandler
'3692:             If pFeatureDataset Is Nothing Then
'3693:                 Call frmOutputWindow.AddOutputMessage("Exporting FeatureDataset: " & "Bad FeatureDataset. Skiped.", , , vbRed)
'3694:             Else
'3695:                 Call ExportGDB_FeatureDataset(pXMLDOMNodeGeodatabaseDesigner, pFeatureDataset)
'3696:             End If
'        Case esriGeoDatabase.esriDatasetType.esriDTFeatureClass
'            '---------------------------
'            ' Document the FeatureClass
'            '---------------------------
'3701:             Call frmOutputWindow.AddOutputMessage("Exporting FeatureClass: " & pDatasetName.Name, , , vbMagenta)
'3702:             Set pName = pDatasetName
'3703:             Set pFeatureClass = Nothing
'            On Error GoTo ErrorRecovery
'3705:             Set pFeatureClass = pName.Open
'            On Error GoTo ErrorHandler
'3707:             If pFeatureClass Is Nothing Then
'3708:                 Call frmOutputWindow.AddOutputMessage("Exporting FeatureClass: " & "Bad FeatureClass. Skiped.", , , vbRed)
'3709:             Else
'3710:                 Call ExportGDB_ObjectClass(pXMLDOMNodeGeodatabaseDesigner, pFeatureClass)
'3711:             End If
'        Case esriGeoDatabase.esriDatasetType.esriDTTable
'            '--------------------
'            ' Document the Table
'            '--------------------
'3716:             Call frmOutputWindow.AddOutputMessage("Exporting Table: " & pDatasetName.Name, , , vbMagenta)
'3717:             Set pName = pDatasetName
'3718:             Set pTable = Nothing
'            On Error GoTo ErrorRecovery
'3720:             Set pTable = pName.Open
'            On Error GoTo ErrorHandler
'3722:             If pTable Is Nothing Then
'3723:                 Call frmOutputWindow.AddOutputMessage("Exporting Table: " & "Bad Table. Skiped.", , , vbRed)
'3724:             Else
'3725:                 Call ExportGDB_ObjectClass(pXMLDOMNodeGeodatabaseDesigner, pTable)
'3726:             End If
'3727:         End Select
'        '
'3729:         Set pDatasetName = pEnumDatasetName.Next
'3730:     Loop
'    '
'    Exit Sub
'ErrorRecovery:
'3734:     Call ErrorRecovery("ExportGDB " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'3735:     Resume Next
'ErrorHandler:
'3737:     Call HandleError(False, "ExportGDB " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportGDB_FeatureDataset(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                                     ByRef pFeatureDataset As esriGeoDatabase.IFeatureDataset)
'    On Error GoTo ErrorHandler
'    '------------------------------------------------
'    ' <featureDataset database="" owner="" table="">
'    '------------------------------------------------
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pEnumFeatureClass As esriGeoDatabase.IEnumFeatureClass
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pGeoDataset As esriGeoDatabase.IGeoDataset
'    Dim pDatasetFeatureClass As esriGeoDatabase.IDataset
'    '
'    Dim pXMLDOMNodeFeatureDataset As MSXML2.IXMLDOMNode
'    '------------------------------
'    ' ADD the FeatureDataset Node
'    '------------------------------
'3756:     Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeParent.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "featureDataset", "")
'3757:     Set pXMLDOMNodeFeatureDataset = pXMLDOMNodeParent.appendChild(pXMLDOMNodeFeatureDataset)
'    '
'3759:     Call AddQualifiedTableNameParts(pXMLDOMNodeFeatureDataset, pFeatureDataset)
'    '-----------------------------------------------
'    ' Export the FeatureDataset's Spatial Reference
'    '-----------------------------------------------
'3763:     Set pGeoDataset = pFeatureDataset
'3764:     Call frmOutputWindow.AddOutputMessage("Exporting Spatial Reference...", , , , , , True, 1)
'3765:     Call ExportGDB_SpatialReference(pXMLDOMNodeFeatureDataset, pGeoDataset.SpatialReference)
'    '-----------------------------------------
'    ' Iterate through each child FeatureClass
'    '-----------------------------------------
'3769:     Set pFeatureClassContainer = pFeatureDataset
'3770:     Set pEnumFeatureClass = pFeatureClassContainer.Classes

'3772:     Set pFeatureClass = pEnumFeatureClass.Next
'3773:     Set pDatasetFeatureClass = pFeatureClass
'3774:     Do Until pFeatureClass Is Nothing
'3775:         Call frmOutputWindow.AddOutputMessage("Exporting ObjectClass: " & pDatasetFeatureClass.Name, , , vbMagenta)
'3776:         Call ExportGDB_ObjectClass(pXMLDOMNodeFeatureDataset, pFeatureClass)
'        '
'3778:         Set pFeatureClass = pEnumFeatureClass.Next
'3779:         Set pDatasetFeatureClass = pFeatureClass
'3780:     Loop
'    '
'    Exit Sub
'ErrorHandler:
'3784:     Call HandleError(False, "ExportGDB_FeatureDataset " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportGDB_SpatialReference(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                                       ByRef pSpatialReference As esriGeometry.ISpatialReference)
'    On Error GoTo ErrorHandler
'    '---------------------------------------------------------------------
'    '    <spatialReference minX="" minY="" maxX="" maxY="" precisionXY=""
'    '                      minM="" maxM="" precisionM=""
'    '                      minZ="" maxZ="" precisionZ=""
'    '                      coordinateSystemDescription=""/>
'    '---------------------------------------------------------------------
'    Dim pXMLDOMNodeSpatialReference As MSXML2.IXMLDOMNode
'    Dim pBuffer As String * 2048
'    Dim pBytes As Long
'    Dim pESRISpatialReference As esriGeometry.IESRISpatialReference
'    ' Minimum
'    Dim pXMin As Double
'    Dim pYMin As Double
'    Dim pMMin As Double
'    Dim pZMin As Double
'    ' Maximum
'    Dim pXMax As Double
'    Dim pYMax As Double
'    Dim pMMax As Double
'    Dim pZMax As Double
'    ' Precision
'    Dim pXYPrecision As Double
'    Dim pMPrecision As Double
'    Dim pZPrecision As Double
'    ' False Origin
'    Dim pXFalse As Double
'    Dim pYFalse As Double
'    Dim pMFalse As Double
'    Dim pZFalse As Double
'    '----------------------------
'    ' Add Spatial Reference Node
'    '----------------------------
'3822:     Set pXMLDOMNodeSpatialReference = pXMLDOMNodeParent.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "spatialReference", "")
'3823:     Set pXMLDOMNodeSpatialReference = pXMLDOMNodeParent.appendChild(pXMLDOMNodeSpatialReference)
'    '------------------------------------------------
'    ' Write "X" & "Y"
'    ' minX="" minY="" precisionXY=""
'    '------------------------------------------------
'3828:     If pSpatialReference.HasXYPrecision Then
'3829:         Call pSpatialReference.GetDomain(pXMin, pXMax, pYMin, pYMax)
'3830:         Call pSpatialReference.GetFalseOriginAndUnits(pXFalse, pYFalse, pXYPrecision)
'        '
'3832:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minX", CStr(pXMin))
'3833:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minY", CStr(pYMin))
'3834:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionXY", CStr(pXYPrecision))
'3835:     Else
'3836:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minX", "")
'3837:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minY", "")
'3838:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionXY", "")
'3839:     End If
'    '--------------------------------
'    ' Write "M"
'    ' minM="" precisionM=""
'    '--------------------------------
'3844:     If pSpatialReference.HasMPrecision Then
'3845:         Call pSpatialReference.GetMDomain(pMMin, pMMax)
'3846:         Call pSpatialReference.GetMFalseOriginAndUnits(pMFalse, pMPrecision)
'3847:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minM", CStr(pMMin))
'3848:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionM", CStr(pMPrecision))
'3849:     Else
'3850:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minM", "")
'3851:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionM", "")
'3852:     End If
'    '-------------------------------
'    ' Write "Z"
'    ' minZ="" precisionZ=""
'    '-------------------------------
'3857:     If pSpatialReference.HasZPrecision Then
'3858:         Call pSpatialReference.GetZDomain(pZMin, pZMax)
'3859:         Call pSpatialReference.GetZFalseOriginAndUnits(pZFalse, pZPrecision)
'        '
'3861:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minZ", CStr(pZMin))
'3862:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionZ", CStr(pZPrecision))
'3863:     Else
'3864:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "minZ", "")
'3865:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "precisionZ", "")
'3866:     End If
'    '--------------------------------
'    ' Coordinate System Description
'    ' coordinateSystemDescription=""
'    '--------------------------------
'3871:     If TypeOf pSpatialReference Is esriGeometry.IUnknownCoordinateSystem Then
'3872:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "coordinateSystemDescription", "")
'3873:     Else
'3874:         Set pESRISpatialReference = pSpatialReference
'3875:         Call pESRISpatialReference.ExportToESRISpatialReference(pBuffer, pBytes)
'3876:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSpatialReference, "coordinateSystemDescription", CStr(Left(pBuffer, InStr(1, pBuffer, Chr(0)) - 1)))
'3877:     End If
'    '
'    Exit Sub
'ErrorHandler:
'3881:     Call HandleError(False, "ExportGDB_SpatialReference " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ExportGDB_ObjectClass(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                                 ByRef pObjectClass As esriGeoDatabase.IObjectClass)
'    On Error GoTo ErrorHandler
'    '--------------------------------------
'    '    <objectClass  database=""
'    '                  owner=""
'    '                  table=""
'    '                  aliasName=""
'    '                  esriDatasetType=""
'    '                  esriFeatureType=""
'    '                  oidField=""
'    '                  shapeField=""
'    '                  subtypeField=""
'    '                  defaultSubtypeCode=""
'    '                  modelName=""
'    '                  configKeyword="">
'    '--------------------------------------
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    '
'    Dim pXMLDOMNodeObjectClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeIndex As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotation As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotateLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeExtent As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotateLayerTransformationProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLabelEngineLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeBasicOverposterLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLineLabelPlacementPriorities As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeLineLabelPosition As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodePointPlacementPriorities As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOverposterLayerProperties As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeAnnotationExpressionEngine As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeTextSymbol As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeColor As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFont As MSXML2.IXMLDOMNode
'    '
'    Dim pXMLDOMNodeDimension As MSXML2.IXMLDOMNode
'    '
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pModelInfo As esriGeoDatabase.IModelInfo
'    Dim pSubtypes As esriGeoDatabase.ISubtypes
'    Dim pEnumSubtype As esriGeoDatabase.IEnumSubtype
'    Dim pAnnoClass As esriCarto.IAnnoClass
'    Dim pDimensionClassExtension As esriCarto.IDimensionClassExtension
'    Dim pIndexes As esriGeoDatabase.IIndexes
'    Dim pIndex As esriGeoDatabase.IIndex
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pField As esriGeoDatabase.IField
'    '
'    Dim pAnnoClassAdmin As esriCarto.IAnnoClassAdmin
'    Dim pAnnotateLayerPropertiesCollection As esriCarto.IAnnotateLayerPropertiesCollection
'    Dim pAnnotateLayerTransformationProperties  As esriCarto.IAnnotateLayerTransformationProperties
'    Dim pAnnotateLayerProperties As esriCarto.IAnnotateLayerProperties
'    Dim pAnnotationExpressionEngine As esriCarto.IAnnotationExpressionEngine
'    Dim pLabelEngineLayerProperties2 As esriCarto.ILabelEngineLayerProperties2
'    Dim pLineLabelPlacementPriorities As esriCarto.ILineLabelPlacementPriorities
'    Dim pLineLabelPosition As esriCarto.ILineLabelPosition
'    Dim pBasicOverposterLayerProperties3 As esriCarto.IBasicOverposterLayerProperties3
'    Dim pPointPlacementPriorities  As esriCarto.IPointPlacementPriorities
'    Dim pOverposterLayerProperties As esriCarto.IOverposterLayerProperties
'    Dim pTextSymbol As esriDisplay.ITextSymbol
'    Dim pColor As esriDisplay.IColor
'    Dim pFontDisp  As stdole.IFontDisp
'    '
'    Dim pSubtypeCode As Long
'    Dim pSubtypeName As String
'    Dim pIndexField As Long
'    Dim pIndexIndex As Long
'    Dim strAngleList As String
'    Dim varAngleList As Variant
'    Dim pIndexClass As Long
'    Dim pIndexAngles As Long
'    '
'    Dim pNetworkClass As esriGeoDatabase.INetworkClass
'    Dim pGeometricNetwork As esriGeoDatabase.IGeometricNetwork
'    '------------------
'    ' Get DOM Document
'    '------------------
'3965:     Set pDOMDocument = pXMLDOMNodeParent.ownerDocument
'    '----------------------------------------------------
'    ' Do not export GeometricNetwork Orphan FeatureClass
'    '----------------------------------------------------
'3969:     If TypeOf pObjectClass Is esriGeoDatabase.IFeatureClass Then
'3970:         Set pFeatureClass = pObjectClass
'        Select Case pFeatureClass.FeatureType
'        Case esriFTComplexEdge, esriFTSimpleEdge, esriFTSimpleJunction
'3973:             Set pNetworkClass = pFeatureClass
'3974:             Set pGeometricNetwork = pNetworkClass.GeometricNetwork
'3975:             If Not (pGeometricNetwork Is Nothing) Then
'3976:                 If pFeatureClass.ObjectClassID = pGeometricNetwork.OrphanJunctionFeatureClass.ObjectClassID Then
'                    Exit Sub
'3978:                 End If
'3979:             End If
'        Case Else
'            '
'3982:         End Select
'3983:     End If
'    '--------------------------
'    ' ADD the ObjectClass Node
'    '--------------------------
'3987:     Set pXMLDOMNodeObjectClass = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "objectClass", "")
'3988:     Set pXMLDOMNodeObjectClass = pXMLDOMNodeParent.appendChild(pXMLDOMNodeObjectClass)
'    '
'3990:     Set pDataset = pObjectClass
'3991:     Set pModelInfo = pObjectClass
'3992:     Set pSubtypes = pObjectClass
'    '
'3994:     Call AddQualifiedTableNameParts(pXMLDOMNodeObjectClass, pDataset)
'3995:     Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "aliasName", CStr(pObjectClass.AliasName))
'3996:     Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "oidField", CStr(pObjectClass.OIDFieldName))
'3997:     Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "esriDatasetType", CStr(pDataset.Type))
'3998:     If pDataset.Type = esriDTFeatureClass Then
'3999:         Set pFeatureClass = pObjectClass
'4000:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "esriFeatureType", CStr(pFeatureClass.FeatureType))
'4001:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "shapeField", CStr(pFeatureClass.ShapeFieldName))
'4002:     Else
'4003:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "esriFeatureType", "")
'4004:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "shapeField", "")
'4005:     End If
'4006:     If pSubtypes.HasSubtype Then
'4007:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "subtypeField", CStr(pSubtypes.SubtypeFieldName))
'4008:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "defaultSubtypeCode", CStr(pSubtypes.DefaultSubtypeCode))
'4009:     Else
'4010:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "subtypeField", "")
'4011:         Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "defaultSubtypeCode", "")
'4012:     End If
'4013:     Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "modelName", CStr(pModelInfo.ModelName))
'4014:     Call modCommon.AddNodeAttribute(pXMLDOMNodeObjectClass, "configKeyword", "")
'    '-----------------
'    ' Add Field Nodes
'    '-----------------
'4018:     Call ExportGDB_Field(pXMLDOMNodeObjectClass, pObjectClass)
'    '----------------------------------------------------
'    ' Add Subtype Node for BASE Domains & Default Values
'    '----------------------------------------------------
'4022:     Call ExportGDB_Subtype(pXMLDOMNodeObjectClass, pObjectClass)
'    '--------------------------------
'    ' Add Subtype Nodes for SUBTYPES
'    '--------------------------------
'4026:     If pSubtypes.HasSubtype Then
'        '-----------------------------
'        ' ObjectClasses with Subtypes
'        '-----------------------------
'4030:         Set pEnumSubtype = pSubtypes.Subtypes
'4031:         pSubtypeName = pEnumSubtype.Next(pSubtypeCode)
'4032:         Do Until pSubtypeName = ""
'4033:             Call frmOutputWindow.AddOutputMessage("Exporting Subtype: " & pSubtypeName & "(" & CStr(pSubtypeCode) & ")", , , , , , True, 1)
'4034:             Call ExportGDB_Subtype(pXMLDOMNodeObjectClass, pObjectClass, pSubtypeCode)
'4035:             pSubtypeName = pEnumSubtype.Next(pSubtypeCode)
'4036:         Loop
'4037:     End If
'    '--------------------------------------------------------
'    ' Add additional information for FeatureClass Extensions
'    '--------------------------------------------------------
'4041:     If pDataset.Type = esriDTFeatureClass Then
'        Select Case pFeatureClass.FeatureType
'        Case esriFTAnnotation
'            '----------------------------------------------------------------------
'            ' <annotation referenceScale="" esriUnits="" version="" autoCreate="">
'            '----------------------------------------------------------------------
'4047:             Set pAnnoClass = pFeatureClass.Extension
'4048:             Set pAnnoClassAdmin = pAnnoClass
'4049:             Set pXMLDOMNodeAnnotation = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "annotation", "")
'4050:             Set pXMLDOMNodeAnnotation = pXMLDOMNodeObjectClass.appendChild(pXMLDOMNodeAnnotation)
'4051:             Call modCommon.AddNodeAttribute(pXMLDOMNodeAnnotation, "referenceScale", CStr(pAnnoClass.ReferenceScale))
'4052:             Call modCommon.AddNodeAttribute(pXMLDOMNodeAnnotation, "esriUnits", CStr(pAnnoClass.ReferenceScaleUnits))
'4053:             Call modCommon.AddNodeAttribute(pXMLDOMNodeAnnotation, "version", CStr(pAnnoClass.version))
'4054:             Call modCommon.AddNodeAttribute(pXMLDOMNodeAnnotation, "autoCreate", CStr(pAnnoClassAdmin.AutoCreate))
'        Case esriFTDimension
'            '---------------------------------------------
'            ' <dimension referenceScale="" esriUnits="">
'            '---------------------------------------------
'4059:             Set pDimensionClassExtension = pFeatureClass.Extension
'4060:             Set pXMLDOMNodeDimension = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "dimension", "")
'4061:             Set pXMLDOMNodeDimension = pXMLDOMNodeObjectClass.appendChild(pXMLDOMNodeDimension)
'4062:             Call modCommon.AddNodeAttribute(pXMLDOMNodeDimension, "referenceScale", CStr(pDimensionClassExtension.ReferenceScale))
'4063:             Call modCommon.AddNodeAttribute(pXMLDOMNodeDimension, "esriUnits", CStr(pDimensionClassExtension.ReferenceScaleUnits))
'4064:         End Select
'4065:     End If
'    '----------------------------------------------
'    ' Export Field Indexes
'    '   <index name="" isAscending="" isUnique="">
'    '         <field name=""/>
'    '   </index>
'    '-----------------------------------------------
'4072:     If mOCExportFieldIndex Then
'4073:         Set pIndexes = pObjectClass.Indexes
'4074:         If pIndexes.IndexCount > 0 Then
'4075:             For pIndexIndex = 0 To pIndexes.IndexCount - 1 Step 1
'4076:                 Set pIndex = pIndexes.Index(pIndexIndex)
'                '
'4078:                 Set pXMLDOMNodeIndex = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "index", "")
'4079:                 Set pXMLDOMNodeIndex = pXMLDOMNodeObjectClass.appendChild(pXMLDOMNodeIndex)
'4080:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeIndex, "name", CStr(pIndex.Name))
'4081:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeIndex, "isAscending", CStr(pIndex.IsAscending))
'4082:                 Call modCommon.AddNodeAttribute(pXMLDOMNodeIndex, "isUnique", CStr(pIndex.IsUnique))
'                '
'4084:                 Set pFields = pIndex.fields
'4085:                 If pFields.FieldCount > 0 Then
'4086:                     For pIndexField = 0 To pFields.FieldCount - 1 Step 1
'4087:                         Set pField = pFields.Field(pIndexField)
'                        '
'4089:                         Set pXMLDOMNodeField = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "field", "")
'4090:                         Set pXMLDOMNodeField = pXMLDOMNodeIndex.appendChild(pXMLDOMNodeField)
'4091:                         Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "name", CStr(pField.Name))
'4092:                     Next pIndexField
'4093:                 End If
'4094:             Next pIndexIndex
'4095:         End If
'4096:     End If
'    '
'    Exit Sub
'ErrorHandler:
'4100:     Call HandleError(False, "ExportGDB_ObjectClass " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportGDB_Field(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                            ByRef pObjectClass As esriGeoDatabase.IObjectClass)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------
'    '            <field name=""
'    '                   aliasName=""
'    '                   DomainFixed = ""
'    '                   Editable = ""
'    '                   IsNullable = ""
'    '                   Length = ""
'    '                   Precision = ""
'    '                   Required = ""
'    '                   scale=""
'    '                   esriFieldType = ""
'    '                   modelName=""/>
'    '-----------------------------------------
'    Dim pIndexField As Long
'    Dim pField As esriGeoDatabase.IField
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pModelInfo As esriGeoDatabase.IModelInfo
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    Dim pField_Length As String
'    Dim pField_Area As String
'    '
'4128:     If TypeOf pObjectClass Is esriGeoDatabase.IFeatureClass Then
'4129:         Set pFeatureClass = pObjectClass
'        '-----------------------
'        ' Get Length Field Name
'        '-----------------------
'4133:         Set pField = pFeatureClass.LengthField
'4134:         If pField Is Nothing Then
'4135:             pField_Length = ""
'4136:         Else
'4137:             pField_Length = pField.Name
'4138:         End If
'        '---------------------
'        ' Get Area Field Name
'        '---------------------
'4142:         Set pField = pFeatureClass.AreaField
'4143:         If pField Is Nothing Then
'4144:             pField_Area = ""
'4145:         Else
'4146:             pField_Area = pField.Name
'4147:         End If
'4148:     Else
'4149:         pField_Length = ""
'4150:         pField_Area = ""
'4151:     End If
'    '
'4153:     Set pFields = pObjectClass.fields
'4154:     For pIndexField = 0 To pFields.FieldCount - 1 Step 1
'4155:         Set pField = pFields.Field(pIndexField)
'4156:         Set pModelInfo = pField
'4157:         If pField.Name = pField_Length Or pField.Name = pField_Area Then
'            '---------------------------------------------
'            ' Skip "Shape_Area" and "Shape_Length" fields
'            '---------------------------------------------
'4161:         Else
'            '---------------------------------
'            ' Add Field Node (and attributes)
'            '---------------------------------
'4165:             Call frmOutputWindow.AddOutputMessage("Exporting Field: " & pField.Name, , , , , , True, 1)
'            '
'4167:             Set pXMLDOMNodeField = pXMLDOMNodeParent.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "field", "")
'4168:             Set pXMLDOMNodeField = pXMLDOMNodeParent.appendChild(pXMLDOMNodeField)
'4169:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "name", CStr(pField.Name))
'4170:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "aliasName", CStr(pField.AliasName))
'4171:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "esriFieldType", CStr(pField.Type))
'4172:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "length", CStr(pField.length))
'4173:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "precision", CStr(pField.Precision))
'4174:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "required", CStr(pField.Required))
'4175:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "scale", CStr(pField.Scale))
'4176:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "domainFixed", CStr(pField.DomainFixed))
'4177:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "editable", CStr(pField.Editable))
'4178:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "isNullable", CStr(pField.IsNullable))
'4179:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "modelName", CStr(pModelInfo.ModelName))
'            '-------------------------------------------------------------
'            ' If Geometry field then parse field to GeometryDef procedure
'            '-------------------------------------------------------------
'4183:             If pField.Type = esriFieldTypeGeometry Then
'4184:                 Call ExportGDB_GeometryDef(pXMLDOMNodeField, pField.GeometryDef, CBool(pFeatureClass.FeatureDataset Is Nothing))
'4185:             End If
'4186:         End If
'4187:     Next pIndexField
'    '
'    Exit Sub
'ErrorHandler:
'4191:     Call HandleError(False, "ExportGDB_Field " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportGDB_GeometryDef(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                                  ByRef pGeometryDef As esriGeoDatabase.IGeometryDef, _
'                                  ByRef pStandAlone As Boolean)
'    On Error GoTo ErrorHandler
'    '--------------------------------------------
'    '        <geometryDef esriGeometryType = ""
'    '                     AvgNumPoints = ""
'    '                     GridCount = ""
'    '                     GridSize = ""
'    '                     HasM=""
'    '                     HasZ=""/>
'    '--------------------------------------------
'    Dim pXMLDOMNodeGeometryDef As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeGrid As MSXML2.IXMLDOMNode
'    Dim pIndexGrid As Long
'    '---------------------------------------
'    ' Add GeometryDef Node (and attributes)
'    '---------------------------------------
'4212:     Set pXMLDOMNodeGeometryDef = pXMLDOMNodeParent.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "geometryDef", "")
'4213:     Set pXMLDOMNodeGeometryDef = pXMLDOMNodeParent.appendChild(pXMLDOMNodeGeometryDef)
'4214:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometryDef, "esriGeometryType", CStr(pGeometryDef.GeometryType))
'4215:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometryDef, "avgNumPoints", CStr(pGeometryDef.AvgNumPoints))
'4216:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometryDef, "hasM", CStr(pGeometryDef.HasM))
'4217:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeometryDef, "hasZ", CStr(pGeometryDef.HasZ))
'    '--------------------
'    ' Add Grid Sub-Nodes
'    '--------------------
'4221:     For pIndexGrid = 0 To pGeometryDef.GridCount - 1 Step 1
'4222:         Set pXMLDOMNodeGrid = pXMLDOMNodeGeometryDef.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "grid", "")
'4223:         Set pXMLDOMNodeGrid = pXMLDOMNodeGeometryDef.appendChild(pXMLDOMNodeGrid)
'4224:         Call modCommon.AddNodeAttribute(pXMLDOMNodeGrid, "size", CStr(pGeometryDef.GridSize(pIndexGrid)))
'4225:     Next pIndexGrid
'    '
'4227:     If pStandAlone Then
'4228:         Call ExportGDB_SpatialReference(pXMLDOMNodeGeometryDef, pGeometryDef.SpatialReference)
'4229:     End If
'    '
'    Exit Sub
'ErrorHandler:
'4233:     Call HandleError(False, "ExportGDB_GeometryDef " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub ExportGDB_Subtype(ByRef pXMLDOMNodeParent As MSXML2.IXMLDOMNode, _
'                              ByRef pObjectClass As esriGeoDatabase.IObjectClass, _
'                              Optional ByRef pSubtypeCode As Long = MISSING)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------------
'    '            <subtype name="" code="" default="">
'    '               <field name="" defaultValue="" domain=""/>
'    '            </subtype>
'    '-----------------------------------------------------------
'    Dim pXMLDOMNodeSubtype As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeField As MSXML2.IXMLDOMNode
'    '
'    Dim pSubtypes As esriGeoDatabase.ISubtypes
'    Dim pFields As esriGeoDatabase.IFields
'    Dim pField As esriGeoDatabase.IField
'    Dim pDomain As esriGeoDatabase.IDomain
'    Dim pDefaultValue As String
'    Dim pDomainName As String
'    Dim pIndexField As Long
'    '---------------
'    ' Get ISubtypes
'    '---------------
'4258:     Set pSubtypes = pObjectClass
'4259:     Set pFields = pObjectClass.fields
'    '-----------------------------------
'    ' Add Subtype Node (and attributes)
'    '-----------------------------------
'4263:     Set pXMLDOMNodeSubtype = pXMLDOMNodeParent.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "subtype", "")
'4264:     Set pXMLDOMNodeSubtype = pXMLDOMNodeParent.appendChild(pXMLDOMNodeSubtype)
'4265:     If pSubtypeCode = MISSING Then
'        '--------------------------------------------
'        ' Base ObjectClass Domains and DefaultValues
'        '--------------------------------------------
'4269:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSubtype, "name", "")
'4270:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSubtype, "code", "")
'4271:     Else
'        '------------------------------------
'        ' Subtype Domains and Default Values
'        '------------------------------------
'4275:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSubtype, "name", CStr(pSubtypes.SubtypeName(pSubtypeCode)))
'4276:         Call modCommon.AddNodeAttribute(pXMLDOMNodeSubtype, "code", CStr(pSubtypeCode))
'4277:     End If
'    '------------------
'    ' Add Fields Nodes
'    '------------------
'4281:     For pIndexField = 0 To pFields.FieldCount - 1 Step 1
'4282:         Set pField = pFields.Field(pIndexField)
'4283:         Set pDomain = Nothing
'4284:         pDefaultValue = ""
'4285:         pDomainName = ""
'        '
'4287:         If pSubtypeCode = MISSING Then
'4288:             Set pDomain = pField.Domain
'4289:             If UCase(pField.Name) <> UCase(pSubtypes.SubtypeFieldName) Then
'4290:                 If Not IsNull(pField.defaultValue) Then
'4291:                     pDefaultValue = CStr(pField.defaultValue)
'4292:                 End If
'4293:             End If
'4294:         Else
'4295:             If UCase(pField.Name) <> UCase(pSubtypes.SubtypeFieldName) Then
'4296:                 Set pDomain = pSubtypes.Domain(pSubtypeCode, pField.Name)
'4297:                 If Not IsNull(pSubtypes.defaultValue(pSubtypeCode, pField.Name)) Then
'4298:                     pDefaultValue = CStr(pSubtypes.defaultValue(pSubtypeCode, pField.Name))
'4299:                 End If
'4300:             End If
'4301:         End If
'4302:         If Not (pDomain Is Nothing) Then
'4303:             pDomainName = CStr(pDomain.Name)
'4304:         End If
'4305:         If pDefaultValue = "" And pDomainName = "" Then
'            '------------------------------
'            ' Skip adding this field node.
'            '------------------------------
'4309:         Else
'4310:             Set pXMLDOMNodeField = pXMLDOMNodeSubtype.ownerDocument.createNode(MSXML2.NODE_ELEMENT, "field", "")
'4311:             Set pXMLDOMNodeField = pXMLDOMNodeSubtype.appendChild(pXMLDOMNodeField)
'4312:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "name", CStr(pField.Name))
'4313:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "defaultValue", pDefaultValue)
'4314:             Call modCommon.AddNodeAttribute(pXMLDOMNodeField, "domain", pDomainName)
'4315:         End If
'4316:     Next pIndexField
'    '
'    Exit Sub
'ErrorHandler:
'4320:     Call HandleError(False, "ExportGDB_Subtype " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub AddQualifiedTableNameParts(ByRef pXMLDOMNode As MSXML2.IXMLDOMNode, _
'                                       ByRef pDataset As esriGeoDatabase.IDataset)
'    On Error GoTo ErrorHandler
'    '
'    Dim pSQLSyntax As esriGeoDatabase.ISQLSyntax
'    Dim pDatabaseName As String
'    Dim pOwnerName As String
'    Dim pTableName As String
'    '
'4332:     If pDataset Is Nothing Then
'4333:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "database", "")
'4334:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "owner", "")
'4335:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "table", "")
'4336:     Else
'4337:         Set pSQLSyntax = pDataset.Workspace
'        '---------------------------------------------
'        ' Parse Dataset Name - Split into RDBMS parts
'        '---------------------------------------------
'4341:         Call pSQLSyntax.ParseTableName(CStr(pDataset.Name), pDatabaseName, pOwnerName, pTableName)
'        '--------------------------
'        ' Add parts to parsed node
'        '--------------------------
'4345:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "database", CStr(pDatabaseName))
'4346:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "owner", CStr(pOwnerName))
'4347:         Call modCommon.AddNodeAttribute(pXMLDOMNode, "table", CStr(pTableName))
'4348:     End If
'    '
'    Exit Sub
'ErrorHandler:
'4352:     Call HandleError(False, "AddQualifiedTableNameParts " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetQualifiedTableNameSimple(ByRef pXMLDOMNode As MSXML2.IXMLDOMNode) As String
'    On Error GoTo ErrorHandler
'    '
'    Dim pDatabaseName As String
'    Dim pOwnerName As String
'    Dim pTableName As String
'    '
'4362:     pDatabaseName = CStr(pXMLDOMNode.Attributes.getNamedItem("database").Text)
'4363:     pOwnerName = CStr(pXMLDOMNode.Attributes.getNamedItem("owner").Text)
'4364:     pTableName = CStr(pXMLDOMNode.Attributes.getNamedItem("table").Text)
'    '
'4366:     If pDatabaseName = "" And pOwnerName = "" And pTableName <> "" Then
'4367:         GetQualifiedTableNameSimple = pTableName
'4368:     Else
'4369:         If pDatabaseName = "" And pOwnerName <> "" And pTableName <> "" Then
'4370:             GetQualifiedTableNameSimple = pOwnerName & "." & pTableName
'4371:         Else
'4372:             If pDatabaseName <> "" And pOwnerName <> "" And pTableName <> "" Then
'4373:                 GetQualifiedTableNameSimple = pDatabaseName & "." & pOwnerName & "." & pTableName
'4374:             Else
'4375:                 GetQualifiedTableNameSimple = "??"
'4376:             End If
'4377:         End If
'4378:     End If
'    '
'    Exit Function
'ErrorHandler:
'4382:     Call HandleError(False, "GetQualifiedTableNameSimple " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub RemoveConnectivityRules(ByRef pGeometricNetwork As esriGeoDatabase.IGeometricNetwork)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------------------
'    ' Removes all Connectivity Rules from the parsed GeometricNetwork
'    '-----------------------------------------------------------------
'    Dim pRule As esriGeoDatabase.IRule
'4391:     Set pRule = pGeometricNetwork.Rules.Next
'    '
'4393:     Do Until pRule Is Nothing
'4394:         If TypeOf pRule Is esriGeoDatabase.IConnectivityRule Then
'4395:             Call frmOutputWindow.AddOutputMessage("Removing GeometricNetwork Rule[ " & CStr(pRule.id) & "]", , , , , , True, 1)
'4396:             Call pGeometricNetwork.DeleteRule(pRule)
'4397:         End If
'4398:         Set pRule = pGeometricNetwork.Rules.Next
'4399:     Loop
'    '
'    Exit Sub
'ErrorHandler:
'4403:     Call HandleError(False, "RemoveConnectivityRules " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Sub RemoveTopologyRules(ByRef pTopology As esriGeoDatabase.ITopology)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------------------
'    ' Removes all Connectivity Rules from the parsed GeometricNetwork
'    '-----------------------------------------------------------------
'    Dim pTopologyRuleContainer As esriGeoDatabase.ITopologyRuleContainer
'    Dim pTopologyRule As esriGeoDatabase.IRule
'    '
'4414:     Set pTopologyRuleContainer = pTopology
'4415:     Set pTopologyRule = pTopologyRuleContainer.Rules.Next
'    '
'4417:     Do Until pTopologyRule Is Nothing
'4418:         Call frmOutputWindow.AddOutputMessage("Removing Topology Rule[ " & CStr(pTopologyRule.id) & "]", , , , , , True, 1)
'4419:         Call pTopologyRuleContainer.DeleteRule(pTopologyRule)
'        '
'4421:         Set pTopologyRule = pTopologyRuleContainer.Rules.Next
'4422:     Loop
'    '
'    Exit Sub
'ErrorHandler:
'4426:     Call HandleError(False, "RemoveTopologyRules " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function ReturnFieldErrorText(eFieldNameErrorType As esriFieldNameErrorType) As String
'  Select Case eFieldNameErrorType
'    Case esriDuplicatedFieldName
'4432:         ReturnFieldErrorText = "Field name is a duplicate of another field name."
'    Case esriInvalidCharacter
'4434:         ReturnFieldErrorText = "Field name containes invalid character."
'    Case esriInvalidFieldNameLength
'4436:         ReturnFieldErrorText = "Field name is too long."
'    Case esriNoFieldError
'4438:         ReturnFieldErrorText = "No field error."
'    Case esriSQLReservedWord
'4440:         ReturnFieldErrorText = "Field name is a SQL Reserved word."
'    Case Else
'4442:         ReturnFieldErrorText = "unknownErrorType" & ":" & eFieldNameErrorType
'4443:   End Select
'End Function

'Private Function ReturnTableErrorText(pEsriTableNameErrorType As esriTableNameErrorType) As String
'  Select Case pEsriTableNameErrorType
'    Case esriIsSQLReservedWord
'4449:         ReturnTableErrorText = "Table name is a SQL reserved word."
'    Case esriHasInvalidCharacter
'4451:         ReturnTableErrorText = "Table name contains an Invalid Character."
'    Case esriHasInvalidStartingCharacter
'4453:         ReturnTableErrorText = "Table name has an invalid starting character."
'    Case Else
'4455:         ReturnTableErrorText = "unknownErrorType" & ":" & pEsriTableNameErrorType
'4456:   End Select
'End Function

'Public Sub ExportTopology(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                          ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '---------------------------------------------------------------------------
'    ' This routine will export both the Topology Dataset and Topology Rules
'    '---------------------------------------------------------------------------
'    Dim pEnumDatasetNameFD As IEnumDatasetName
'    Dim pEnumDatasetNameTP As IEnumDatasetName
'    'Dim pName As IName
'    'Dim pDatasetName As IDatasetName
'    Dim pTopologyName As ITopologyName
'    'Dim pTopology As ITopology
'    Dim pFeatureDatasetName2 As IFeatureDatasetName2
'    '
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    '--------------------------------------------------------------------
'    '<topology database="" owner="" table=""
'    '          clusterTolerance=""
'    '          maxGeneratedErrorCount=""
'    '          nothingTrusted=""             REMOVED!!!!
'    '          esriTopologyState=""
'    '          configurationKeywork="">
'    '</topology>
'    '--------------------------------------------------------------------
'4483:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'4484:     Set pEnumDatasetNameFD = pWorkspace.DatasetNames(esriDTFeatureDataset)
'4485:     Call pEnumDatasetNameFD.Reset
'4486:     Set pFeatureDatasetName2 = pEnumDatasetNameFD.Next
'4487:     Do Until pFeatureDatasetName2 Is Nothing
'4488:         Set pEnumDatasetNameTP = pFeatureDatasetName2.TopologyNames
'4489:         Call pEnumDatasetNameTP.Reset
'4490:         Set pTopologyName = pEnumDatasetNameTP.Next
'4491:         Do Until pTopologyName Is Nothing
'4492:             Call ExportTopology2(pXMLDOMNodeGeodatabaseDesigner, pTopologyName)
'            '
'4494:             Set pTopologyName = pEnumDatasetNameTP.Next
'4495:         Loop
'4496:         Set pFeatureDatasetName2 = pEnumDatasetNameFD.Next
'4497:     Loop
'    '
'    Exit Sub
'ErrorHandler:
'4501:     Call HandleError(False, "ExportTopology " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ExportTopology2(ByRef pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode, _
'                           ByRef pTopologyName As esriGeoDatabase.ITopologyName)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDatasetName As esriGeoDatabase.IDatasetName
'    Dim pName As IName
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pTopology As esriGeoDatabase.ITopology
'    Dim pTopologyProperties As esriGeoDatabase.ITopologyProperties
'    Dim pEnumFeatureClass As esriGeoDatabase.IEnumFeatureClass
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pDatasetFC As esriGeoDatabase.IDataset
'    Dim pTopologyClass As esriGeoDatabase.ITopologyClass
'    Dim pEnumRule As esriGeoDatabase.IEnumRule
'    Dim pRule As esriGeoDatabase.IRule
'    Dim pTopologyRule As esriGeoDatabase.ITopologyRule
'    Dim pTopologyRuleContainer As esriGeoDatabase.ITopologyRuleContainer
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pDatasetOrigin As esriGeoDatabase.IDataset
'    Dim pDatasetDestination As esriGeoDatabase.IDataset
'    Dim pSubtypesOrigin As esriGeoDatabase.ISubtypes
'    Dim pSubtypesDestination As esriGeoDatabase.ISubtypes
'    Dim pSubtypeNameOrigin As String
'    Dim pSubtypeNameDestination As String
'    '
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    Dim pXMLDOMNodeTopology As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeTopologyRule As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOrigin As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeDestination As MSXML2.IXMLDOMNode
'    '------------------
'    ' Get IDatasetName
'    '------------------
'4538:     Set pDatasetName = pTopologyName
'4539:     Call frmOutputWindow.AddOutputMessage("Exporting Topology: " & pDatasetName.Name, , , vbMagenta)
'4540:     Set pName = pTopologyName
'    '----------------------------------
'    ' Attempt to Open Topology Dataset
'    '----------------------------------
'    On Error GoTo ErrorRecovery
'4545:     Set pTopology = pName.Open
'    On Error GoTo ErrorHandler
'    '-------------------------------------------
'    ' If Open Failed then Report Error and Exit
'    '-------------------------------------------
'4550:     If pTopology Is Nothing Then
'4551:         Call frmOutputWindow.AddOutputMessage("Failed To Open Topology Dataset", , , vbRed)
'        Exit Sub
'4553:     End If
'4554:     Set pDataset = pTopology
'    '------------------
'    ' Get DOM Document
'    '------------------
'4558:     Set pDOMDocument = pXMLDOMNodeGeodatabaseDesigner.ownerDocument
'    '
'4560:     Set pXMLDOMNodeTopology = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "topology", "")
'4561:     Set pXMLDOMNodeTopology = pXMLDOMNodeGeodatabaseDesigner.appendChild(pXMLDOMNodeTopology)
'4562:     Call AddQualifiedTableNameParts(pXMLDOMNodeTopology, pDataset)
'4563:     Call modCommon.AddNodeAttribute(pXMLDOMNodeTopology, "clusterTolerance", CStr(pTopology.ClusterTolerance))
'4564:     Call modCommon.AddNodeAttribute(pXMLDOMNodeTopology, "maxGeneratedErrorCount", CStr(pTopology.MaximumGeneratedErrorCount))
'4565:     Call modCommon.AddNodeAttribute(pXMLDOMNodeTopology, "esriTopologyState", CStr(pTopology.State))
'4566:     Call modCommon.AddNodeAttribute(pXMLDOMNodeTopology, "configurationKeyword", "")
'    '----------------------------------------------------
'    ' Write Topology FeatureClasses
'    '    <featureClass database="" owner="" table=""
'    '                  weight=""
'    '                  xyRank=""
'    '                  zRank=""
'    '                  eventNotificationOnValidate="" />
'    '----------------------------------------------------
'4575:     Call frmOutputWindow.AddOutputMessage("Exporting Topology Classes", , , , , , True, 1)

'4577:     Set pTopologyProperties = pTopology
'4578:     Set pEnumFeatureClass = pTopologyProperties.Classes
'4579:     Set pFeatureClass = pEnumFeatureClass.Next

'4581:     Do Until pFeatureClass Is Nothing
'4582:         Set pDatasetFC = pFeatureClass
'4583:         Set pTopologyClass = pFeatureClass
'        '
'4585:         Set pXMLDOMNodeFeatureClass = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "featureClass", "")
'4586:         Set pXMLDOMNodeFeatureClass = pXMLDOMNodeTopology.appendChild(pXMLDOMNodeFeatureClass)
'4587:         Call AddQualifiedTableNameParts(pXMLDOMNodeFeatureClass, pDatasetFC)
'4588:         Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "weight", CStr(pTopologyClass.Weight))
'4589:         Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "xyRank", CStr(pTopologyClass.XYRank))
'4590:         Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "zRank", CStr(pTopologyClass.ZRank))
'4591:         Call modCommon.AddNodeAttribute(pXMLDOMNodeFeatureClass, "eventNotificationOnValidate", CStr(pTopologyClass.EventNotificationOnValidate))
'        '
'4593:         Set pFeatureClass = pEnumFeatureClass.Next
'4594:     Loop
'    '------------------------------------------------------------------
'    ' Write Topology Rules
'    '    <rule name=""
'    '          allOriginSubtypes=""
'    '          allDestinationSubtypes=""
'    '          guid=""
'    '          esriTopologyRuleType=""
'    '          triggerErrorEvents="">
'    '        <origin      database="" owner="" table="" subtype="" />
'    '        <destination database="" owner="" table="" subtype="" />
'    '    </rule>
'    '------------------------------------------------------------------
'4607:     Set pFeatureClassContainer = pTopology
'4608:     Set pTopologyRuleContainer = pTopology
'4609:     Set pEnumRule = pTopologyRuleContainer.Rules
'4610:     Set pRule = pEnumRule.Next

'4612:     Do Until pRule Is Nothing
'4613:         If TypeOf pRule Is esriGeoDatabase.ITopologyRule Then
'4614:             Call frmOutputWindow.AddOutputMessage("Exporting Topology Rules [" & pRule.id & "]", , , , , , True, 2)
'            '
'4616:             Set pTopologyRule = pRule
'            '---------------
'            ' Add Rule Node
'            '---------------
'4620:             Set pXMLDOMNodeTopologyRule = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "rule", "")
'4621:             Set pXMLDOMNodeTopologyRule = pXMLDOMNodeTopology.appendChild(pXMLDOMNodeTopologyRule)
'4622:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "name", CStr(pTopologyRule.Name))
'4623:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "allOriginSubtypes", CStr(pTopologyRule.AllOriginSubtypes))
'4624:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "allDestinationSubtypes", CStr(pTopologyRule.AllDestinationSubtypes))
'4625:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "guid", CStr(pTopologyRule.Guid))
'4626:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "esriTopologyRuleType", CStr(pTopologyRule.TopologyRuleType))
'4627:             Call modCommon.AddNodeAttribute(pXMLDOMNodeTopologyRule, "triggerErrorEvents", CStr(pTopologyRule.TriggerErrorEvents))
'            '---------------------------------
'            ' Add Origin and Destination Node
'            '---------------------------------
'4631:             Set pDatasetOrigin = pFeatureClassContainer.ClassByID(pTopologyRule.OriginClassID)
'4632:             Set pSubtypesOrigin = pDatasetOrigin
'4633:             Set pXMLDOMNodeOrigin = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "origin", "")
'4634:             Set pXMLDOMNodeOrigin = pXMLDOMNodeTopologyRule.appendChild(pXMLDOMNodeOrigin)
'4635:             Call AddQualifiedTableNameParts(pXMLDOMNodeOrigin, pDatasetOrigin)
'4636:             If pTopologyRule.AllOriginSubtypes Then
'4637:                 pSubtypeNameOrigin = ""
'4638:             Else
'4639:                 pSubtypeNameOrigin = pSubtypesOrigin.SubtypeName(pTopologyRule.OriginSubtype)
'4640:             End If
'4641:             Call modCommon.AddNodeAttribute(pXMLDOMNodeOrigin, "subtype", pSubtypeNameOrigin)
'            '
'4643:             Set pDatasetDestination = pFeatureClassContainer.ClassByID(pTopologyRule.DestinationClassID)
'4644:             Set pSubtypesDestination = pDatasetDestination
'4645:             Set pXMLDOMNodeDestination = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "destination", "")
'4646:             Set pXMLDOMNodeDestination = pXMLDOMNodeTopologyRule.appendChild(pXMLDOMNodeDestination)
'4647:             Call AddQualifiedTableNameParts(pXMLDOMNodeDestination, pDatasetDestination)
'4648:             If pTopologyRule.AllDestinationSubtypes Then
'4649:                 pSubtypeNameDestination = ""
'4650:             Else
'4651:                 pSubtypeNameDestination = pSubtypesDestination.SubtypeName(pTopologyRule.DestinationSubtype)
'4652:             End If
'4653:             Call modCommon.AddNodeAttribute(pXMLDOMNodeDestination, "subtype", pSubtypeNameDestination)
'4654:         End If
'4655:         Set pRule = pEnumRule.Next
'4656:     Loop
'    '
'    Exit Sub
'ErrorRecovery:
'4660:     Call ErrorRecovery("ExportTopology2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'4661:     Resume Next
'ErrorHandler:
'4663:     Call HandleError(False, "ExportTopology2 " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub ImportTopology(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                          ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '------------------------------------------------------------------------
'    ' Main Routine. This routine is called when the user clicks "Import Topology"
'    ' If a FeatureDataset is selected:  (1) Create NEW Topology Dataset
'    '                                   (2) Add NEW Topology Rules
'    ' If a TopologyDataset is selected: (1) Add NEW Topology Rules
'    '------------------------------------------------------------------------
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListTopology As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeTopology As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    '
''    Dim pGeodatabaseRelease As esriGeoDatabase.IGeodatabaseRelease
'    Dim pTopology As esriGeoDatabase.ITopology
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pFeatureDataset As esriGeoDatabase.IFeatureDataset
'    '
'    Dim pIndexTopology As Long
'    Dim strName As String
'    Dim strFeatureClassName As String
''    '---------------------------
''    ' Check Geodatabase Version
''    '---------------------------
''    Set pGeodatabaseRelease = pWorkspace
''    Select Case pGeodatabaseRelease.MajorVersion
''    Case 0
''        '------------------
''        ' Geodatabase 8.0x
''        '------------------
''        Call frmOutputWindow.AddOutputMessage("Topology Datasets cannot be created on a pre-8.3 Geodatabase", , , vbRed)
''        MsgBox "This commands requires an 8.3+ Geodatabase", vbExclamation, App.ProductName
''        Exit Sub
''    Case 1
''        '---------------------------
''        ' Geodatabase 8.1, 8.2, 8.3
''        '---------------------------
''        If pGeodatabaseRelease.MinorVersion >= 3 Then
''            '--------------------------
''            ' OK (Geodatabase is 8.3+)
''            '--------------------------
''        Else
''            Call frmOutputWindow.AddOutputMessage("Topology Datasets cannot be created on a pre-8.3 Geodatabase", , , vbRed)
''            Exit Sub
''        End If
''    Case Else
''        '----------------------
''        ' OK (Geodatabase 9.x)
''        '----------------------
''    End Select
'    '----------------------------------
'    ' Get GeodatabaseDesigner XML Node
'    '----------------------------------
'4720:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.getElementsByTagName("geodatabaseDesigner").nextNode
'    '--------------------
'    ' Get Topology Nodes
'    '--------------------
'4724:     Set pXMLDOMNodeListTopology = pXMLDOMNodeGeodatabaseDesigner.selectNodes("topology")
'    '-------------------------------------------
'    ' Loop Through Each Topology in the XML DOM
'    '-------------------------------------------
'4728:     For pIndexTopology = 0 To pXMLDOMNodeListTopology.length - 1 Step 1
'4729:         Set pXMLDOMNodeTopology = pXMLDOMNodeListTopology.Item(pIndexTopology)
'4730:         Set pXMLDOMNodeFeatureClass = pXMLDOMNodeTopology.selectNodes("featureClass").Item(0)
'        '
'4732:         strName = CStr(pXMLDOMNodeTopology.Attributes.getNamedItem("table").Text)
'4733:         strFeatureClassName = CStr(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("table").Text)
'        '----------------------------------
'        ' Check if Topology already exists
'        '----------------------------------
'4737:         Set pTopology = modCommon.GetDataset(pWorkspace, strName, esriDTTopology)
'        '
'4739:         If pTopology Is Nothing Then
'4740:             Set pFeatureClass = modCommon.GetDataset(pWorkspace, strFeatureClassName, esriDTFeatureClass)
'4741:             Set pFeatureDataset = pFeatureClass.FeatureDataset
'4742:             Set pTopology = CreateTopology(pXMLDOMNodeTopology, pFeatureDataset)
'4743:             If pTopology Is Nothing Then
'4744:                 Call frmOutputWindow.AddOutputMessage("Topology Creation Unsuccessful", , , vbRed, , , True, 1)
'4745:             Else
'4746:                 Call frmOutputWindow.AddOutputMessage("Topology Creation Successful", , , , , , True, 1)
'4747:                 Call ImportTopologyRules(pXMLDOMNodeTopology, pTopology)
'4748:             End If
'4749:         Else
'4750:             Call frmOutputWindow.AddOutputMessage("Topology [" & strName & "] already exists", , , vbRed)
'4751:             Call ImportTopologyRules(pXMLDOMNodeTopology, pTopology)
'4752:         End If
'4753:     Next pIndexTopology
'    '
'    Exit Sub
'NoTopology:
'4757:     Set pTopology = Nothing
'4758:     Resume Next
'    '
'ErrorHandler:
'4761:     Call HandleError(False, "ImportTopology " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function CreateTopology(ByRef pXMLDOMNodeTopology As MSXML2.IXMLDOMNode, _
'                                ByRef pFeatureDataset As esriGeoDatabase.IFeatureDataset) As esriGeoDatabase.ITopology
'    On Error GoTo ErrorHandler
'    '----------------------------------------------------------
'    ' Creates a NEW Topology inside the FeatureDataset
'    '----------------------------------------------------------
'    Dim pTopology As esriGeoDatabase.ITopology
'    Dim pTopologyContainer As esriGeoDatabase.ITopologyContainer
'    Dim pFeatureClassContainer As IFeatureClassContainer
'    Dim pFeatureClass As IFeatureClass
'    '
'    Dim pXMLDOMNodeFeatureClass As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeListFeatureClass As MSXML2.IXMLDOMNodeList

'    Dim pIndex As Long
'    Dim pName As String
'    Dim pClusterTolerance As String
'    Dim pMaxGeneratedErrorCount As Long
'    Dim pConfigurationKeyword As String
'    Dim pWeight As Double
'    Dim pXYRank As Long
'    Dim pZRank As Long
'    Dim pEventNotificationOnValidate As Boolean
'    '----------------
'    ' Set Interfaces
'    '----------------
'4790:     Set pTopology = Nothing
'4791:     Set pTopologyContainer = pFeatureDataset
'4792:     Set pFeatureClassContainer = pFeatureDataset
'    '-------------------------
'    ' Get Topology properties
'    '-------------------------
'4796:     pName = CStr(pXMLDOMNodeTopology.Attributes.getNamedItem("table").Text)
'4797:     pClusterTolerance = CStr(pXMLDOMNodeTopology.Attributes.getNamedItem("clusterTolerance").Text)
'4798:     pMaxGeneratedErrorCount = CLng(pXMLDOMNodeTopology.Attributes.getNamedItem("maxGeneratedErrorCount").Text)
'4799:     pConfigurationKeyword = CStr(pXMLDOMNodeTopology.Attributes.getNamedItem("configurationKeyword").Text)
'    '
'4801:     Call frmOutputWindow.AddOutputMessage("Importing Topology: " & pName, , , vbMagenta)
'    '-----------------------------
'    ' Create new Topology Dataset
'    '-----------------------------
'4805:     Set pTopology = pTopologyContainer.CreateTopology(pName, pClusterTolerance, pMaxGeneratedErrorCount, pConfigurationKeyword)
'    '-------------------------------------------------------------
'    ' Add FeatureClasses (ITopologyClass) to the Topology Dataset
'    '-------------------------------------------------------------
'4809:     Set pXMLDOMNodeListFeatureClass = pXMLDOMNodeTopology.selectNodes("featureClass")
'4810:     For pIndex = 0 To pXMLDOMNodeListFeatureClass.length - 1 Step 1
'4811:         Set pXMLDOMNodeFeatureClass = pXMLDOMNodeListFeatureClass.Item(pIndex)
'        '
'4813:         pName = CStr(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("table").Text)
'4814:         pWeight = CDbl(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("weight").Text)
'4815:         pXYRank = CLng(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("xyRank").Text)
'4816:         pZRank = CLng(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("zRank").Text)
'4817:         pEventNotificationOnValidate = CStr(pXMLDOMNodeFeatureClass.Attributes.getNamedItem("eventNotificationOnValidate").Text)
'        '
'4819:         Set pFeatureClass = pFeatureClassContainer.ClassByName(pName)
'        '
'4821:         Call frmOutputWindow.AddOutputMessage("Importing TopologyClass: " & pName, , , , , , True, 1)
'4822:         Call pTopology.AddClass(pFeatureClass, pWeight, pXYRank, pZRank, pEventNotificationOnValidate)
'4823:     Next pIndex
'    '---------------------
'    ' Return New Topology
'    '---------------------
'4827:     Set CreateTopology = pTopology
'    '
'    Exit Function
'ErrorHandler:
'4831:     Call HandleError(False, "CreateTopology " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Sub ImportTopologyRules(ByRef pXMLDOMNodeTopology As MSXML2.IXMLDOMNode, _
'                                ByRef pTopology As esriGeoDatabase.ITopology)
'    On Error GoTo ErrorHandler
'    '-------------------------------------------------------------
'    ' This Routine will recreate Topology Rules from an XML file.
'    '-------------------------------------------------------------
'    Dim pXMLDOMNodeListRule As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeRule As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeOrigin As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeDestination As MSXML2.IXMLDOMNode
'    '
'    Dim pFeatureClassContainer As esriGeoDatabase.IFeatureClassContainer
'    Dim pFeatureClass As esriGeoDatabase.IFeatureClass
'    Dim pTopologyRuleContainer As esriGeoDatabase.ITopologyRuleContainer
'    Dim pTopologyRule As esriGeoDatabase.ITopologyRule
'    '
'    Dim pName As String
'    Dim pEsriTopologyRuleType As Long
'    Dim pTriggerErrorEvents As Boolean
'    Dim pAllDestinationSubtypes As Boolean
'    Dim pAllOriginSubtypes As Boolean
'    Dim pDestinationClassName As String
'    Dim pDestinationSubtypeName As String
'    Dim pOriginClassName As String
'    Dim pOriginSubtypeName As String
'    '
'    Dim pIndex As Long
'    '----------------------------------------------------
'    ' <rule name=""
'    '       allDestinationSubtypes=""
'    '       allOriginSubtypes=""
'    '       guid=""
'    '       esriTopologyRuleType=""
'    '       triggerErrorEvents="">
'    '    <origin      database="" owner="" subtype="" />
'    '    <destination database="" owner="" subtype="" />
'    ' </rule>
'    '----------------------------------------------------
'4872:     Set pTopologyRuleContainer = pTopology
'4873:     Set pFeatureClassContainer = pTopology
'    '------------------------------------------
'    ' First Delete all existing Topology Rules
'    '------------------------------------------
'4877:     Call RemoveTopologyRules(pTopology)
'    '
'4879:     Set pXMLDOMNodeListRule = pXMLDOMNodeTopology.selectNodes("rule")
'    '
'4881:     For pIndex = 0 To pXMLDOMNodeListRule.length - 1 Step 1
'4882:         Set pXMLDOMNodeRule = pXMLDOMNodeListRule.Item(pIndex)
'4883:         Set pXMLDOMNodeOrigin = pXMLDOMNodeRule.selectSingleNode("origin")
'4884:         Set pXMLDOMNodeDestination = pXMLDOMNodeRule.selectSingleNode("destination")
'        '
'4886:         pName = CStr(pXMLDOMNodeRule.Attributes.getNamedItem("name").Text)
'4887:         pEsriTopologyRuleType = CLng(pXMLDOMNodeRule.Attributes.getNamedItem("esriTopologyRuleType").Text)
'4888:         pTriggerErrorEvents = CBool(pXMLDOMNodeRule.Attributes.getNamedItem("triggerErrorEvents").Text)
'4889:         pAllDestinationSubtypes = CBool(pXMLDOMNodeRule.Attributes.getNamedItem("allDestinationSubtypes").Text)
'4890:         pAllOriginSubtypes = CBool(pXMLDOMNodeRule.Attributes.getNamedItem("allOriginSubtypes").Text)

'4892:         pDestinationClassName = CStr(pXMLDOMNodeDestination.Attributes.getNamedItem("table").Text)
'4893:         pDestinationSubtypeName = CStr(pXMLDOMNodeDestination.Attributes.getNamedItem("subtype").Text)
'4894:         pOriginClassName = CStr(pXMLDOMNodeOrigin.Attributes.getNamedItem("table").Text)
'4895:         pOriginSubtypeName = CStr(pXMLDOMNodeOrigin.Attributes.getNamedItem("subtype").Text)
'        '
'4897:         Set pTopologyRule = New esriGeoDatabase.TopologyRule
'4898:         pTopologyRule.Name = pName
'4899:         pTopologyRule.TopologyRuleType = pEsriTopologyRuleType
'4900:         pTopologyRule.AllDestinationSubtypes = pAllDestinationSubtypes
'4901:         pTopologyRule.AllOriginSubtypes = pAllOriginSubtypes
'4902:         pTopologyRule.DestinationClassID = pFeatureClassContainer.ClassByName(pDestinationClassName).FeatureClassID
'4903:         pTopologyRule.OriginClassID = pFeatureClassContainer.ClassByName(pOriginClassName).FeatureClassID
'        '
'4905:         If Not pAllDestinationSubtypes Then
'4906:             pTopologyRule.DestinationSubtype = GetSubtypeCodeFromName(pFeatureClassContainer.ClassByName(pDestinationClassName), pDestinationSubtypeName)
'4907:         End If
'4908:         If Not pAllOriginSubtypes Then
'4909:             pTopologyRule.OriginSubtype = GetSubtypeCodeFromName(pFeatureClassContainer.ClassByName(pOriginClassName), pOriginSubtypeName)
'4910:         End If
'        '
'4912:         Call frmOutputWindow.AddOutputMessage("Importing Topology Rule [ " & CStr(pIndex) & "]", , , , , , True, 1)
'4913:         Call pTopologyRuleContainer.AddRule(pTopologyRule)
'4914:     Next pIndex
'    '
'    Exit Sub
'ErrorHandler:
'4918:     Call HandleError(False, "ImportTopologyRules " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub StartGeodatabaseOperation(ByRef pApplication As IApplication, _
'                                     ByRef pOperationType As gdbOperationType)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    Dim pXMLDOMNodeListGD As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeList As MSXML2.IXMLDOMNodeList
'    Dim pXMLDOMNodeGD As MSXML2.IXMLDOMNode
'    '
'    Dim pGxApplication As esriCatalogUI.IGxApplication
'    Dim pGxObject As esriCatalog.IGxObject
'    Dim pGxDatabase As esriCatalog.IGxDatabase
'    Dim pGxDataset As esriCatalog.IGxDataset
'    '
'    Dim pWorkspace As esriGeoDatabase.IWorkspace
'    Dim pDataset As esriGeoDatabase.IDataset
'    Dim pGeodatabaseRelease As esriGeoDatabase.IGeodatabaseRelease
'    '
'    Dim blnIsGeodatabase83 As Boolean
'    '---------------------
'    ' Get Selected Object
'    '---------------------
'4943:     Set pGxApplication = pApplication
'4944:     Set pGxObject = pGxApplication.SelectedObject
'    '-------------------------
'    ' Check Selected GxObject
'    '-------------------------
'    Select Case pOperationType
'    Case gdbOperationType.gdbExportAll, _
'         gdbOperationType.gdbExportDomain, _
'         gdbOperationType.gdbExportGeometricNetwork, _
'         gdbOperationType.gdbExportObjectClass, _
'         gdbOperationType.gdbExportRelationship, _
'         gdbOperationType.gdbExportTopology, _
'         gdbOperationType.gdbImport, _
'         gdbOperationType.gdbUtilityResetGUID
'        '---------------
'        ' Get Workspace
'        '---------------
'4960:         If TypeOf pGxObject Is esriCatalog.IGxDatabase Then
'4961:             Set pGxDatabase = pGxObject
'4962:             Set pWorkspace = pGxDatabase.Workspace
'            '
'4964:             If pWorkspace.Type = esriFileSystemWorkspace Then
'4965:                 Call MsgBox("Please select a Personal or ArcSDE Geodatabase", vbInformation, App.FileDescription)
'                Exit Sub
'4967:             End If
'4968:         Else
'4969:             Call MsgBox("Please select a Personal or ArcSDE Geodatabase", vbInformation, App.FileDescription)
'            Exit Sub
'4971:         End If
'    Case gdbOperationType.gdbUtilityGeometricNetworkEditor
'        '----------------------
'        ' Get GeometricNetwork
'        '----------------------
'4976:         If TypeOf pGxObject Is esriCatalog.IGxDataset Then
'4977:             Set pGxDataset = pGxObject
'4978:             If pGxDataset.Type = esriDTGeometricNetwork Then
'4979:                 Set pDataset = pGxDataset.Dataset
'4980:                 Set pWorkspace = pDataset.Workspace
'4981:             Else
'4982:                 Call MsgBox("Please select a GeometricNetwork", vbExclamation, App.FileDescription)
'                Exit Sub
'4984:             End If
'4985:         Else
'4986:             Call MsgBox("Please select a GeometricNetwork", vbExclamation, App.FileDescription)
'            Exit Sub
'4988:         End If
'    Case gdbOperationType.gdbUtilityTruncate
'        '-------------
'        ' Get Dataset
'        '-------------
'4993:         If TypeOf pGxObject Is esriCatalog.IGxDatabase Then
'4994:             Set pGxDatabase = pGxObject
'4995:             Set pWorkspace = pGxDatabase.Workspace
'            '
'4997:             If pWorkspace.Type = esriFileSystemWorkspace Then
'4998:                 Call MsgBox("Please select a Personal or ArcSDE Geodatabase", vbInformation, App.FileDescription)
'                Exit Sub
'5000:             Else
'5001:                  Set pDataset = pWorkspace
'5002:             End If
'5003:         Else
'5004:             If TypeOf pGxObject Is esriCatalog.IGxDataset Then
'5005:                 Set pGxDataset = pGxObject
'5006:                 Set pDataset = pGxDataset.Dataset
'5007:                 Set pWorkspace = pDataset.Workspace
'5008:                 If pWorkspace.Type = esriFileSystemWorkspace Then
'5009:                     Call MsgBox("Please select an object from a Personal or ArcSDE Geodatabase", vbInformation, App.FileDescription)
'                    Exit Sub
'5011:                 End If
'                '
'                Select Case pDataset.Type
'                Case esriDatasetType.esriDTTable, _
'                     esriDatasetType.esriDTFeatureClass, _
'                     esriDatasetType.esriDTFeatureDataset
'                Case esriDatasetType.esriDTRelationshipClass
'5018:                     If Not TypeOf pDataset Is ITable Then
'5019:                         Call MsgBox("Only attributed relationships can be truncated", vbExclamation, App.FileDescription)
'                        Exit Sub
'5021:                     End If
'                Case Else
'5023:                     Call MsgBox("Please select a Geodatabase, Table, FeatureClass, FeatureDataset or Relationship", vbExclamation, App.FileDescription)
'                    Exit Sub
'5025:                 End Select
'5026:             Else
'5027:                 Call MsgBox("Please select a Geodatabase or Geodatabase Object", vbExclamation, App.FileDescription)
'                Exit Sub
'5029:             End If
'5030:         End If
'    Case Else
'5032:         Call MsgBox("Unknown Operation Type", vbInformation, App.FileDescription)
'        Exit Sub
'5034:     End Select
'    '---------------------------
'    ' Check Geodatabase Version
'    '---------------------------
'5038:     Set pGeodatabaseRelease = pWorkspace
'    Select Case pGeodatabaseRelease.MajorVersion
'    Case 0
'        '------------------
'        ' Geodatabase 8.0x
'        '------------------
'5044:         blnIsGeodatabase83 = False
'    Case 1
'        '---------------------------
'        ' Geodatabase 8.1, 8.2, 8.3
'        '---------------------------
'5049:         If pGeodatabaseRelease.MinorVersion >= 3 Then
'            '--------------------------
'            ' OK (Geodatabase is 8.3+)
'            '--------------------------
'5053:             blnIsGeodatabase83 = True
'5054:         Else
'5055:             blnIsGeodatabase83 = False
'5056:         End If
'    Case Else
'        '----------------------
'        ' OK (Geodatabase 9.x)
'        '----------------------
'5061:         blnIsGeodatabase83 = True
'5062:     End Select
'    '--------------------------
'    ' Add header to log window
'    '--------------------------
'5066:     Call frmOutputWindow.AddOutputWindowHeader(pWorkspace)
'    '----------------------------------
'    ' Change to hourglass mouse cursor
'    '----------------------------------
'    Select Case pOperationType
'    Case gdbOperationType.gdbExportAll, _
'         gdbOperationType.gdbExportDomain, _
'         gdbOperationType.gdbExportGeometricNetwork, _
'         gdbOperationType.gdbExportObjectClass, _
'         gdbOperationType.gdbExportRelationship, _
'         gdbOperationType.gdbExportTopology
'        '----------------------------------
'        ' Get new XML Document (in memory)
'        '----------------------------------
'5080:         Set pDOMDocument = modCommon.NewXMLDocument
'        '------------------------
'        ' Add header to XML file
'        '------------------------
'5084:          Call modCommon.WriteGeodatabaseDesignerHeader(pDOMDocument, pWorkspace)
'    Case gdbOperationType.gdbImport
'        '----------------------------------
'        ' Get new XML Document (in memory)
'        '----------------------------------
'5089:         Set pDOMDocument = modCommon.NewXMLDocument
'        '------------------------------
'        ' Load XML from rtfXML control
'        '------------------------------
'5093:         Call pDOMDocument.loadXML(frmOutputWindow.rtbXML.Text)
'5094:     End Select
'    '---------------------------------
'    ' Call Operation Specific Routine
'    '---------------------------------
'    Select Case pOperationType
'    Case gdbOperationType.gdbExportAll
'        '---------------------------
'        ' Export Entire Geodatabase
'        '---------------------------
'5103:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "Domain Exporter", , , vbBlue, True, True)
'5104:         Call modImportExport.ExportDomain(pWorkspace, pDOMDocument)
'        '
'5106:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "ObjectClass Exporter", , , vbBlue, True, True)
'5107:         Call modImportExport.ExportGDB(pWorkspace, pDOMDocument)
'        '
'5109:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "Relationship Exporter", , , vbBlue, True, True)
'5110:         Call modImportExport.ExportRelationship(pWorkspace, pDOMDocument)
'        '
'5112:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "GeometricNetwork Exporter", , , vbBlue, True, True)
'5113:         Call modImportExport.ExportGeometricNetwork(pWorkspace, pDOMDocument)
'        '
'5115:         If blnIsGeodatabase83 Then
'5116:             Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Exporter", , , vbBlue, True, True)
'5117:             Call modImportExport.ExportTopology(pWorkspace, pDOMDocument)
'5118:         Else
'5119:             Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Exporter - Skipped (non-83 Geodatabase)", , , vbBlue, True)
'5120:         End If
'    Case gdbOperationType.gdbExportDomain
'        '---------------
'        ' Export Domain
'        '---------------
'5125:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "Domain Exporter", , , vbBlue, True, True)
'5126:         Call modImportExport.ExportDomain(pWorkspace, pDOMDocument)
'    Case gdbOperationType.gdbExportObjectClass
'        '--------------------
'        ' Export ObjectClass
'        '--------------------
'5131:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "ObjectClass Exporter", , , vbBlue, True, True)
'5132:         Call modImportExport.ExportGDB(pWorkspace, pDOMDocument)
'    Case gdbOperationType.gdbExportRelationship
'        '---------------------
'        ' Export Relationship
'        '---------------------
'5137:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "Relationship Exporter", , , vbBlue, True, True)
'5138:         Call modImportExport.ExportRelationship(pWorkspace, pDOMDocument)
'    Case gdbOperationType.gdbExportGeometricNetwork
'        '-------------------------
'        ' Export GeometricNetwork
'        '-------------------------
'5143:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "GeometricNetwork Exporter", , , vbBlue, True, True)
'5144:         Call modImportExport.ExportGeometricNetwork(pWorkspace, pDOMDocument)
'    Case gdbOperationType.gdbExportTopology
'        '-----------------
'        ' Export Topology
'        '-----------------
'5149:         If blnIsGeodatabase83 Then
'5150:             Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Exporter", , , vbBlue, True, True)
'5151:             Call modImportExport.ExportTopology(pWorkspace, pDOMDocument)
'5152:         Else
'5153:             Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Exporter - Skipped (non-83 Geodatabase)", , , vbBlue, True)
'5154:         End If
'    Case gdbOperationType.gdbUtilityGeometricNetworkEditor
'        '-------------------------
'        ' GeometricNetwork Editor
'        '-------------------------
'5159:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "GeometricNetwork Editor Initiated", , , vbBlue, True, True)
'5160:         Call frmGeometricNetwork.Init(pDataset)
'    Case gdbOperationType.gdbUtilityTruncate
'        '---------------
'        ' Truncate Tool
'        '---------------
'5165:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "Truncate Tool", , , vbBlue, True, True)
'5166:         Call modUtility.TruncateGeodatabase(pDataset)
'    Case gdbOperationType.gdbUtilityResetGUID
'        '---------------
'        ' Reset GUID
'        '---------------
'5171:         Call frmOutputWindow.AddOutputMessage(vbCrLf & "ObjectClass GUID Initilizer", , , vbBlue, True, True)
'5172:         Call modUtility.GUIDReset(pWorkspace)
'    Case gdbOperationType.gdbImport
'        '------------------
'        ' Import Functions
'        '------------------
'5177:         Set pXMLDOMNodeListGD = pDOMDocument.selectNodes("geodatabaseDesigner")
'        '
'5179:         If pXMLDOMNodeListGD.length <> 1 Then
'5180:             Call frmOutputWindow.AddOutputMessage("Not a valid Geodatabase Designer XML File!")
'5181:         Else
'5182:             Set pXMLDOMNodeGD = pXMLDOMNodeListGD.Item(0)
'            '------------------
'            ' Domain XML File?
'            '------------------
'5186:             If pXMLDOMNodeGD.selectNodes("domain").length > 0 Then
'5187:                 Call frmOutputWindow.AddOutputMessage(vbCrLf & "Domain Importer", , , vbBlue, True, True)
'5188:                 Call ImportDomain(pWorkspace, pDOMDocument)
'5189:             End If
'            '-----------------------
'            ' ObjectClass XML File?
'            '-----------------------
'5193:             If pXMLDOMNodeGD.selectNodes("featureDataset").length > 0 Or _
'               pXMLDOMNodeGD.selectNodes("objectClass").length > 0 Then
'5195:                 Call frmOutputWindow.AddOutputMessage(vbCrLf & "ObjectClass Importer", , , vbBlue, True, True)
'5196:                 Call ImportGDB(pWorkspace, pDOMDocument)
'5197:             End If
'            '------------------------
'            ' Relationship XML File?
'            '------------------------
'5201:             If pXMLDOMNodeGD.selectNodes("relationshipClass").length > 0 Then
'5202:                 Call frmOutputWindow.AddOutputMessage(vbCrLf & "Relationship Importer", , , vbBlue, True, True)
'5203:                 Call ImportRelationship(pWorkspace, pDOMDocument)
'5204:             End If
'            '----------------------------
'            ' GeometricNetwork XML File?
'            '----------------------------
'5208:             If pXMLDOMNodeGD.selectNodes("geometricNetwork").length > 0 Then
'5209:                 Call frmOutputWindow.AddOutputMessage(vbCrLf & "GeometricNetwork Importer", , , vbBlue, True, True)
'5210:                 Call ImportGeometricNetwork(pWorkspace, pDOMDocument)
'5211:             End If
'            '--------------------
'            ' Topology XML File?
'            '--------------------
'5215:             If blnIsGeodatabase83 Then
'5216:                 If pXMLDOMNodeGD.selectNodes("topology").length > 0 Then
'5217:                     Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Importer", , , vbBlue, True, True)
'5218:                     Call ImportTopology(pWorkspace, pDOMDocument)
'5219:                 End If
'5220:             Else
'5221:                 Call frmOutputWindow.AddOutputMessage(vbCrLf & "Topology Importer - Skipped (non-83 Geodatabase)", , , vbBlue, True)
'5222:             End If
'            '--------------------------------------
'            ' Refresh ArcCatalog Table of Contents
'            '--------------------------------------
'5226:             Call frmOutputWindow.AddOutputMessage(vbCrLf & "Refreshing ArcCatalog Table of Contents")
'5227:             Call pGxApplication.Refresh(pGxObject.FullName)
'5228:         End If
'5229:     End Select
'    '
'    Select Case pOperationType
'    Case gdbOperationType.gdbExportAll, _
'         gdbOperationType.gdbExportDomain, _
'         gdbOperationType.gdbExportGeometricNetwork, _
'         gdbOperationType.gdbExportObjectClass, _
'         gdbOperationType.gdbExportRelationship, _
'         gdbOperationType.gdbExportTopology
'        '----------------------------------
'        ' Get new XML Document (in memory)
'        '----------------------------------
'5241:         Call modCommon.PutXMLDocument(pDOMDocument)
'5242:     End Select
'    '---------------------
'    ' Add Log File Footer
'    '---------------------
'5246:     Call frmOutputWindow.AddOutputWindowFooter
'    '---------------------------------------------
'    ' Update Dockable Window Button Enable Status
'    '---------------------------------------------
'5250:     Call frmOutputWindow.UpdateButtonEnableStatus
'    '-----------------
'    ' Clear Variables
'    '-----------------
'5254:     Set pWorkspace = Nothing
'5255:     Set pDataset = Nothing
'5256:     Set pDOMDocument = Nothing
'    '
'    Exit Sub
'ErrorHandler:
'5260:     Call HandleError(False, "StartGeodatabaseOperation " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Private Function EnvelopeNode(ByRef pDOMDocument As MSXML2.DOMDocument40, _
'                              ByRef pEnvelope As esriGeometry.IEnvelope, _
'                              ByRef strName As String) As MSXML2.IXMLDOMNode
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMNode As MSXML2.IXMLDOMNode
'    '
'5270:     Set pXMLDOMNode = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, strName, "")
'5271:     Call modCommon.AddNodeAttribute(pXMLDOMNode, "xMin", CStr(pEnvelope.XMin))
'5272:     Call modCommon.AddNodeAttribute(pXMLDOMNode, "xMax", CStr(pEnvelope.XMax))
'5273:     Call modCommon.AddNodeAttribute(pXMLDOMNode, "yMin", CStr(pEnvelope.YMin))
'5274:     Call modCommon.AddNodeAttribute(pXMLDOMNode, "yMax", CStr(pEnvelope.YMax))
'    '
'5276:     Set EnvelopeNode = pXMLDOMNode
'    '
'    Exit Function
'ErrorHandler:
'5280:     Call HandleError(False, "EnvelopeNode " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Private Function NodeEnvelope(ByRef pXMLDOMNode As MSXML2.IXMLDOMNode) As esriGeometry.IEnvelope
'    On Error GoTo ErrorHandler
'    '
'    Dim pEnvelope As esriGeometry.IEnvelope
'    '
'5288:     Set pEnvelope = New esriGeometry.Envelope
'5289:     Call pEnvelope.PutCoords(CDbl(pXMLDOMNode.Attributes.getNamedItem("xMin").Text), _
'                             CDbl(pXMLDOMNode.Attributes.getNamedItem("yMin").Text), _
'                             CDbl(pXMLDOMNode.Attributes.getNamedItem("xMax").Text), _
'                             CDbl(pXMLDOMNode.Attributes.getNamedItem("yMax").Text))
'    '
'5294:     Set NodeEnvelope = pEnvelope
'    '
'    Exit Function
'ErrorHandler:
'5298:     Call HandleError(False, "NodeEnvelope " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

