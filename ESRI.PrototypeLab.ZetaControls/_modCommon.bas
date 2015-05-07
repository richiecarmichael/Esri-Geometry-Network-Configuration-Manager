'Attribute VB_Name = "modCommon"
'Option Explicit

''-----------------
'' Public Contants
''-----------------
'Public Const COMMAND_CATEGORY      As String = "Geodatabase Designer 2"
'Public Const OUTPUT_WINDOW_CAPTION As String = "Geodatabase Designer Output Window"

'Public Const HELP_FOLDER As String = "AVI"                  ' "HELP"
'Public Const HELP_FILE   As String = "GD2_INTRODUCTION.avi" ' "HELP.pdf"

'Public Const GUID_FEATURECLASS_CLSID   As Variant = "{52353152-891A-11D0-BEC6-00805F7C4268}"
'Public Const GUID_TABLE_CLSID          As Variant = "{7A566981-C114-11D2-8A28-006097AFF44E}"
'Public Const GUID_ANNOTATION_CLSID     As Variant = "{E3676993-C682-11D2-8A2A-006097AFF44E}"
'Public Const GUID_ANNOTATION_EXTCLSID  As Variant = "{24429589-D711-11D2-9F41-00C04F6BC6A5}"
'Public Const GUID_DIMENSION_CLSID      As Variant = "{496764FC-E0C9-11D3-80CE-00C04F601565}"
'Public Const GUID_DIMENSION_EXTCLSID   As Variant = "{48F935E2-DA66-11D3-80CE-00C04F601565}"
'Public Const GUID_SIMPLEJUNCTION_CLSID As Variant = "{CEE8D6B8-55FE-11D1-AE55-0000F80372B4}"
'Public Const GUID_SIMPLEEDGE_CLSID     As Variant = "{E7031C90-55FE-11D1-AE55-0000F80372B4}"
'Public Const GUID_COMPLEXEDGE_CLSID    As Variant = "{A30E8A2A-C50B-11D1-AEA9-0000F80372B4}"

'Public Const ENUMERATOR_FIELDTYPE   As String = "Small Integer, Integer, Single, Double, String, Date, OID, Geometry, Blob"
'Public Const ENUMERATOR_MERGEPOLICY As String = "0, Sum Values, Area Weigthed, Default Value"
'Public Const ENUMERATOR_SPLITPOLICY As String = "0, Geometry Ratio, Duplicate, Default Value"
'Public Const ENUMERATOR_DOMAINTYPE  As String = "0, Range, Coded Value, String"
'Public Const ENUMERATOR_NETWORKCLASSANCILLARYROLE As String = "None, Source/Sink"
'Public Const ENUMERATOR_FEATURETYPE As String = "0, Simple Feature, 2, 3, 4, 5, 6, Simple Junction, Simple Edge, Complex Junction, Complex Edge, Annotation, Coverage Annotation, Dimension"

''------------------
'' Public Variables
''------------------
'Public mGNExportFeatureClass As Boolean
'Public mGNExportJunctionConnRule As Boolean
'Public mGNExportEdgeConnRule As Boolean
'Public mGNExportWeights As Boolean
'Public mGNImportSnapping As Double                  ' -1 = None | 0 = Minimum | >0 = Other
'Public mGNImportPreserveEnabledValue As Boolean
'Public mGNImportClearConnRule As Boolean

'Public mDMExportOCAssocation As Boolean

'Public mOCExportFieldIndex As Boolean
'Public mOCImportFieldIndex As Boolean
'Public mCurrentXMLFilename As String

''-----------------------------------
'' GDB Designer Operation Enumerator
''-----------------------------------
'Public Enum gdbOperationType
'    gdbExportDomain = 0
'    gdbExportObjectClass = 1
'    gdbExportRelationship = 2
'    gdbExportGeometricNetwork = 3
'    gdbExportTopology = 4
'    gdbExportAll = 5
'    gdbImport = 6
'    '
'    gdbUtilityGeometricNetworkEditor = 10
'    gdbUtilityTruncate = 11
'    gdbUtilityResetGUID = 12
'End Enum

''-------------------
'' Private Constants
''-------------------
'Private Const SW_SHOWNORMAL As Long = 1                            ' Used by ShellExecute Windows API
'Private Const MODULE_NAME As String = "modCommon.bas"              ' Error Handler
'Private Const TEMP_FILENAME As String = "_geodatabasedesigner.xml" ' Temp GD Filename

''----------------------
'' Windows Declarations
''----------------------
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFileName As OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFileName As OPENFILENAME) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

''-------------------
'' Custom Data Types
''-------------------
'Private Type OPENFILENAME
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type

''-------------------
'' Windows API Calls
''-------------------

'Public Function ShowOpen(ByRef pMessage As String, _
'                         ByRef pExtension As String) As String
'    On Error GoTo ErrorHandler
'    '
'    Dim pOpenFileName As OPENFILENAME
'    ' Set the structure size
'116:     pOpenFileName.lStructSize = Len(pOpenFileName)
'    ' Set the application's instance
'118:     pOpenFileName.hInstance = App.hInstance
'    ' Set the filet
'120:     pOpenFileName.lpstrFilter = UCase(pExtension) & " (*." & LCase(pExtension) & ")" & Chr$(0) & "*." & LCase(pExtension) & Chr$(0) & _
'                                "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'    ' Create a buffer
'123:     pOpenFileName.lpstrFile = Space$(254)
'    ' Set the maximum number of chars
'125:     pOpenFileName.nMaxFile = 255
'    ' Create a buffer
'127:     pOpenFileName.lpstrFileTitle = Space$(254)
'    ' Set the maximum number of chars
'129:     pOpenFileName.nMaxFileTitle = 255
'    ' Set the dialog title
'131:     pOpenFileName.lpstrTitle = pMessage
'    ' no extra flags
'133:     pOpenFileName.flags = 0
'    ' Show the 'Open File'-dialog
'135:     If GetOpenFileName(pOpenFileName) Then
'136:         ShowOpen = Trim(pOpenFileName.lpstrFile)
'137:     Else
'138:         ShowOpen = ""
'139:     End If
'    '
'    Exit Function
'ErrorHandler:
'143:     Call HandleError(False, "ShowOpen " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function ShowSave(ByRef pMessage As String, _
'                         ByRef pExtension As String) As String
'    On Error GoTo ErrorHandler
'    '
'    Dim pOpenFileName As OPENFILENAME
'    ' Set the structure size
'152:     pOpenFileName.lStructSize = Len(pOpenFileName)
'    ' Set the application's instance
'154:     pOpenFileName.hInstance = App.hInstance
'    ' Set the file filter
'156:     pOpenFileName.lpstrFilter = UCase(pExtension) & " (*." & LCase(pExtension) & ")" & Chr$(0) & "*." & LCase(pExtension) & Chr$(0) & _
'                                "All files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
'    ' Create a buffer
'159:     pOpenFileName.lpstrFile = Space$(254)
'    ' Set the maximum number of chars
'161:     pOpenFileName.nMaxFile = 255
'    ' Create a buffer
'163:     pOpenFileName.lpstrFileTitle = Space$(254)
'    ' Set the maximum number of chars
'165:     pOpenFileName.nMaxFileTitle = 255
'    ' Set the dialog title
'167:     pOpenFileName.lpstrTitle = pMessage
'    ' Set flags
'    ' cdlOFNPathMustExist    &H800
'    ' cdlOFNOverwritePrompt  &H2
'171:     pOpenFileName.flags = &H2 Or &H800
'    ' Specify the default extension (used when no ext is specified).
'173:     pOpenFileName.lpstrDefExt = "." & LCase(pExtension)
'    ' Show the 'Save File'-dialog
'175:     If GetSaveFileName(pOpenFileName) Then
'176:         ShowSave = Trim(pOpenFileName.lpstrFile)
'177:     Else
'178:         ShowSave = ""
'179:     End If
'    '
'    Exit Function
'ErrorHandler:
'183:     Call HandleError(False, "ShowSave " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetXMLDocument() As MSXML2.DOMDocument40
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLFileName As String
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    '----------------------------------------------------------------
'    ' Use Microsoft Common Dialog to prompt user for source XML file
'    '----------------------------------------------------------------
'194:     pXMLFileName = modCommon.ShowOpen("Open XML File", "xml")
'195:     If pXMLFileName = "" Then
'196:         Call MsgBox("No File Selected", vbInformation, App.FileDescription)
'        Exit Function
'198:     End If
'    '-------------------
'    ' Open XML Document
'    '-------------------
'202:     Set pDOMDocument = New MSXML2.DOMDocument40
'203:     pDOMDocument.async = False
'204:     pDOMDocument.validateOnParse = True
'205:     Call pDOMDocument.Load(pXMLFileName)
'206:     Call pDOMDocument.setProperty("SelectionLanguage", "XPath")
'    '--------------------------------------------
'    ' Return reference to XML Document to client
'    '--------------------------------------------
'210:     Set GetXMLDocument = pDOMDocument
'    '
'    Exit Function
'ErrorHandler:
'214:     Call HandleError(False, "GetXMLDocument " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function NewXMLDocument() As MSXML2.DOMDocument40
'On Error GoTo ErrorHandler
'    '------------------------------------------
'    ' Returns to the client a NEW XML document
'    '------------------------------------------
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    '
'224:     Set pDOMDocument = New MSXML2.DOMDocument40
'225:     Call pDOMDocument.setProperty("SelectionLanguage", "XPath")
'226:     Set NewXMLDocument = pDOMDocument
'    '
'    Exit Function
'ErrorHandler:
'230:     Call HandleError(False, "NewXMLDocument " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub PutXMLDocument(ByRef pDOMDocument As MSXML2.DOMDocument40)
'    On Error GoTo ErrorHandler
'    '---------------------------
'    ' Create a new XML document
'    '---------------------------
'    'Dim pTemporaryFileName As String
'    Dim pFSO As New Scripting.FileSystemObject
'    Dim pFile As Scripting.File
'    Dim pTextStream As Scripting.TextStream
'    '
'    Dim pVBSAXContentHandler As MSXML2.IVBSAXContentHandler
'    Dim pMXXMLWriter40 As New MSXML2.MXXMLWriter40
'    Dim pSAXXMLReader40 As New MSXML2.SAXXMLReader40
'    '
'247:     If pDOMDocument Is Nothing Then
'        '-----------------------------
'        ' The Export Operation Failed
'        '-----------------------------
'251:     Else
'        '--------------------------------
'        ' The Export Operation Succeeded
'        '--------------------------------
'255:         pMXXMLWriter40.indent = True
'        'pMXXMLWriter40.standalone = True
'257:         Set pVBSAXContentHandler = pMXXMLWriter40
'258:         Set pSAXXMLReader40.contentHandler = pVBSAXContentHandler
'259:         Call pSAXXMLReader40.parse(pDOMDocument)
'        '
'261:         mCurrentXMLFilename = GetTemporaryFileName
'262:         Set pTextStream = pFSO.CreateTextFile(mCurrentXMLFilename, True, True)
'263:         Call pTextStream.Write(pMXXMLWriter40.Output)
'264:         Call pTextStream.Close
'        '
'266:         Set pDOMDocument = Nothing
'        '
'268:         Call frmOutputWindow.webHTML.Navigate(mCurrentXMLFilename)
'        '
'270:         Set pFile = pFSO.GetFile(mCurrentXMLFilename)
'271:         Set pTextStream = pFile.OpenAsTextStream(ForReading, TristateTrue)
'272:         frmOutputWindow.rtbXML.Text = pTextStream.ReadAll
'273:         Call pTextStream.Close
'274:     End If
'    '
'    Exit Sub
'ErrorHandler:
'278:     Call HandleError(False, "PutXMLDocument " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetTemporaryFileName() As String
'On Error GoTo ErrorHandler
'    '
'    Dim pFSO As New FileSystemObject
'    '
'286:     GetTemporaryFileName = pFSO.BuildPath(pFSO.GetSpecialFolder(TemporaryFolder), TEMP_FILENAME)
'    '
'    Exit Function
'ErrorHandler:
'290:     Call HandleError(False, "GetTemporaryFileName " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetCurrentUserName() As String
'On Error GoTo ErrorHandler
'    '
'    Dim strUserName As String
'    '
'298:     strUserName = String(100, Chr(0))
'299:     Call GetUserName(strUserName, 100)
'300:     GetCurrentUserName = Left(strUserName, InStr(strUserName, Chr(0)) - 1)
'    '
'    Exit Function
'ErrorHandler:
'304:     Call HandleError(False, "GetCurrentUserName " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetPCName() As String
'On Error GoTo ErrorHandler
'    '
'    Dim strComputer As String
'    '
'312:     strComputer = String(255, Chr(0))
'313:     Call GetComputerName(strComputer, 255)
'314:     GetPCName = Left(strComputer, InStr(1, strComputer, Chr(0)) - 1)
'    '
'    Exit Function
'ErrorHandler:
'318:     Call HandleError(False, "GetPCName " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''----------------------
'' GEODATABASE ROUTINES
''----------------------

'Public Function GetGeodatabaseVersion(ByRef pWorkspace As esriGeoDatabase.IWorkspace) As String
'On Error GoTo ErrorHandler
'    '
'    Dim pGeodatabaseRelease As esriGeoDatabase.IGeodatabaseRelease
'    '
'330:     Set pGeodatabaseRelease = pWorkspace
'331:     GetGeodatabaseVersion = pGeodatabaseRelease.MajorVersion & "." & pGeodatabaseRelease.MinorVersion & "." & pGeodatabaseRelease.BugfixVersion
'    '
'    Exit Function
'ErrorHandler:
'335:     Call HandleError(False, "GetGeodatabaseVersion " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetConnectionProperties(ByRef pWorkspace As esriGeoDatabase.IWorkspace) As VBA.Collection
'    On Error GoTo ErrorHandler
'    '
'    Dim pCollection As VBA.Collection
'    Dim pPropertySet As esriSystem.IPropertySet
'    Dim pIndex As Long
'    Dim pNames As Variant
'    Dim pValues As Variant
'    '
'347:     Set pCollection = New VBA.Collection
'348:     Set pPropertySet = pWorkspace.ConnectionProperties
'349:     pPropertySet.GetAllProperties pNames, pValues
'350:     For pIndex = 0 To UBound(pNames) Step 1
'351:         If UCase(pNames(pIndex)) <> "PASSWORD" And _
'           UCase(pNames(pIndex)) <> "PROVIDERCLSID" Then
'353:            pCollection.Add pNames(pIndex), pValues(pIndex)
'354:         End If
'355:     Next pIndex
'    '
'    Exit Function
'ErrorHandler:
'359:     Call HandleError(False, "GetConnectionProperties " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetGeodatabaseFlavor(ByRef pWorkspace As esriGeoDatabase.IWorkspace) As String
'On Error GoTo ErrorHandler
'    '---------------------------------------------------
'    ' This function will return the Geodatabase's RDBMS
'    '---------------------------------------------------
'367:     Const pKeywordOracle As String = "ACCESS"
'368:     Const pKeywordSQLServer As String = "AUTHORIZATION"
'369:     Const pKeywordInformix As String = "?"
'370:     Const pKeywordDB2 As String = "?"
'    '
'    Select Case pWorkspace.Type
'    Case esriFileSystemWorkspace
'374:         GetGeodatabaseFlavor = "Invalid"
'    Case esriLocalDatabaseWorkspace
'376:         GetGeodatabaseFlavor = "Access"
'    Case esriRemoteDatabaseWorkspace
'        '--------------------------------------
'        ' Select RDBMS based on reserved words
'        '--------------------------------------
'        Dim pSQLSyntax As esriGeoDatabase.ISQLSyntax
'        Dim pEnumBSTR As esriSystem.IEnumBSTR
'        Dim pKeyword As String
'        Dim pTable As esriGeoDatabase.ITable
'        '
'386:         Set pSQLSyntax = pWorkspace
'387:         Set pEnumBSTR = pSQLSyntax.GetKeywords
'388:         pKeyword = pEnumBSTR.Next
'389:         GetGeodatabaseFlavor = "Unknown"
'390:         Do Until pKeyword = ""
'            Select Case pKeyword
'            Case pKeywordOracle
'393:                 GetGeodatabaseFlavor = "Oracle"
'            Case pKeywordSQLServer
'395:                 GetGeodatabaseFlavor = "SQLServer"
'            Case pKeywordInformix
'397:                 GetGeodatabaseFlavor = "Informix"
'            Case pKeywordDB2
'399:                 GetGeodatabaseFlavor = "DB2"
'400:             End Select
'401:             pKeyword = pEnumBSTR.Next
'402:         Loop
'403:     End Select
'    '
'    Exit Function
'ErrorHandler:
'407:     Call HandleError(False, "GetGeodatabaseFlavor " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''---------------------------
'' FileSystemObject Routines
''---------------------------

'Public Function AppendInstallationPath(ByRef pFileName As String) As String
'On Error GoTo ErrorHandler
'    '
'    Dim pFSO As Scripting.FileSystemObject
'418:     Set pFSO = New Scripting.FileSystemObject
'    '
'420:     AppendInstallationPath = pFSO.BuildPath(App.Path, pFileName)
'    '
'    Exit Function
'ErrorHandler:
'424:     Call HandleError(False, "AppendInstallationPath " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub SubmitFileToOS(ByRef pFileName As String)
'On Error GoTo ErrorHandler
'    '
'430:     Call ShellExecute(0, "open", pFileName, vbNullString, "", SW_SHOWNORMAL)
'    '
'    Exit Sub
'ErrorHandler:
'434:     Call HandleError(False, "SubmitFileToOS " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub WriteFile(ByRef pFileName As String, ByRef pLine As String)
'    '------------------------------------------------------
'    ' Write line to a text file.
'    ' Used to write the HTML document and Error Reporting.
'    '------------------------------------------------------
'442:     If Len(pLine) > 0 Then
'        Dim pFSO As Scripting.FileSystemObject
'        Dim pTextStream As Scripting.TextStream
'        '
'446:         Set pFSO = New Scripting.FileSystemObject
'        '---------------------------------------
'        ' Write the submitted line to the File.
'        '----------------------------------------
'450:         Set pTextStream = pFSO.OpenTextFile(pFileName, _
'                                            Scripting.ForAppending, _
'                                            True, _
'                                            Scripting.TristateFalse)
'454:         pTextStream.WriteLine pLine
'455:         pTextStream.Close
'456:     End If
'End Sub

''----------------------------
'' Geodatabase Designer Calls
''----------------------------

'Public Function QuoteSwap(ByRef pLineIn As String) As String
'On Error GoTo ErrorHandler
'    '------------------------------------
'    ' Swap single quote for double quote
'    '------------------------------------
'468:     QuoteSwap = Replace(pLineIn, Chr(39), Chr(34), 1, -1, vbTextCompare)
'    '
'    Exit Function
'ErrorHandler:
'472:     Call HandleError(False, "QuoteSwap " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub AddNodeAttribute(ByRef pXMLDOMNode As MSXML2.IXMLDOMNode, _
'                            ByRef pName As String, _
'                            ByRef pValue As String)
'    On Error GoTo ErrorHandler
'    '
'    Dim pXMLDOMAttribute As MSXML2.IXMLDOMAttribute
'    '
'482:     Set pXMLDOMAttribute = pXMLDOMNode.Attributes.setNamedItem(pXMLDOMNode.ownerDocument.createAttribute(pName))
'483:     pXMLDOMAttribute.Text = pValue
'    '
'    Exit Sub
'ErrorHandler:
'487:     Call HandleError(False, "AddNodeAttribute " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub WriteGeodatabaseDesignerHeader(ByRef pXMLDOMNode As MSXML2.IXMLDOMNode, _
'                                          ByRef pWorkspace As esriGeoDatabase.IWorkspace)
'    On Error GoTo ErrorHandler
'    '
'    Dim pDOMDocument As MSXML2.DOMDocument40
'    Dim pXMLDOMNodeXML As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeGeodatabaseDesigner As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeMetadata As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeMetadataChild As MSXML2.IXMLDOMNode
'    Dim pXMLDOMNodeConnectionProperty As MSXML2.IXMLDOMNode
'    '
'    Dim pPropertySet As esriSystem.IPropertySet
'    Dim pNames As Variant
'    Dim pValues As Variant
'    Dim pIndex As Long
'    '--------------------------------------------------------------------
'    ' Add the "Processing Instruction Header"
'    ' <?xml version="1.0" encoding='UTF-8'?>
'    ' <?xml-stylesheet type="text/xsl" href="Geodatabase Designer.xsl"?>
'    '--------------------------------------------------------------------
'510:     If pXMLDOMNode.nodeType = NODE_DOCUMENT Then
'511:         Set pDOMDocument = pXMLDOMNode
'        'Set pXMLDOMNodeXML = pDOMDocument.createProcessingInstruction("xml", QuoteSwap("version='1.0' encoding='UTF-8'"))
'513:         Set pXMLDOMNodeXML = pDOMDocument.createProcessingInstruction("xml", QuoteSwap("version='1.0'"))
'514:         Set pXMLDOMNodeXML = pDOMDocument.appendChild(pXMLDOMNodeXML)
'515:         Set pXMLDOMNodeXML = pDOMDocument.createProcessingInstruction("xml-stylesheet", QuoteSwap("type='text/xsl' href='" & GetXSLFilename & "'"))
'516:         Set pXMLDOMNodeXML = pDOMDocument.appendChild(pXMLDOMNodeXML)
'517:     Else
'518:         Set pDOMDocument = pXMLDOMNode.ownerDocument
'519:     End If
'    '--------------------------------------------
'    ' Add "GeodatabaseDesigner" Node (+ version)
'    ' <geodatabasedesigner version="1.0">
'    '--------------------------------------------
'524:     Set pXMLDOMNodeGeodatabaseDesigner = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "geodatabaseDesigner", "")
'525:     Set pXMLDOMNodeGeodatabaseDesigner = pXMLDOMNode.appendChild(pXMLDOMNodeGeodatabaseDesigner)
'526:     Call modCommon.AddNodeAttribute(pXMLDOMNodeGeodatabaseDesigner, "version", CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision))
'    '-------------------------------------------------
'    ' Add "Metadata" node under "GeodatabaseDesigner"
'    ' <metadata>
'    '-------------------------------------------------
'531:     Set pXMLDOMNodeMetadata = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "metadata", "")
'532:     Set pXMLDOMNodeMetadata = pXMLDOMNodeGeodatabaseDesigner.appendChild(pXMLDOMNodeMetadata)
'    '------------------------------------------------------------
'    ' Add "CreationDate" node under "Metadata"
'    ' <creationDate year="" month="" day="" hour="" second="" />
'    '------------------------------------------------------------
'537:     Set pXMLDOMNodeMetadataChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "creationDate", "")
'538:     Set pXMLDOMNodeMetadataChild = pXMLDOMNodeMetadata.appendChild(pXMLDOMNodeMetadataChild)
'539:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "year", CStr(Year(Date)))
'540:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "month", PadZero(CStr(Month(Date)), 2))
'541:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "day", PadZero(CStr(Day(Date)), 2))
'542:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "hour", PadZero(CStr(Hour(Time)), 2))
'543:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "minute", PadZero(CStr(Minute(Time)), 2))
'544:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "second", PadZero(CStr(Second(Time)), 2))
'    '--------------------------------------
'    ' Add "Creator" under "Metadata"
'    ' <creator      user="" computer="" />
'    '--------------------------------------
'549:     Set pXMLDOMNodeMetadataChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "creator", "")
'550:     Set pXMLDOMNodeMetadataChild = pXMLDOMNodeMetadata.appendChild(pXMLDOMNodeMetadataChild)
'551:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "user", modCommon.GetCurrentUserName)
'552:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "computer", modCommon.GetPCName)
'    '------------------------------------------------------------
'    ' Add "Geodatabase" under "Metadata"
'    ' <geodatabase  esriWorkspaceType="" flavor="" version="" />
'    '-------------------------------------------------------------
'557:     Set pXMLDOMNodeMetadataChild = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "geodatabase", "")
'558:     Set pXMLDOMNodeMetadataChild = pXMLDOMNodeMetadata.appendChild(pXMLDOMNodeMetadataChild)
'559:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "esriWorkspaceType", CStr(pWorkspace.Type))
'560:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "flavor", GetGeodatabaseFlavor(pWorkspace))
'561:     Call modCommon.AddNodeAttribute(pXMLDOMNodeMetadataChild, "version", GetGeodatabaseVersion(pWorkspace))
'    '------------------------------------------------------
'    ' Add "ConnectionProperties" under "Metadata"
'    ' <connectionProperty parameter1="" parameter2="" />
'    '------------------------------------------------------
'566:     Set pPropertySet = pWorkspace.ConnectionProperties
'567:     pPropertySet.GetAllProperties pNames, pValues
'568:     For pIndex = 0 To UBound(pNames) Step 1
'569:         If UCase(pNames(pIndex)) <> "PASSWORD" And _
'           UCase(pNames(pIndex)) <> "PROVIDERCLSID" Then
'            '-------------------------------------------------------------------
'            ' Add Connection Parameters to "ConnectionProperties" as Attributes
'            '-------------------------------------------------------------------
'574:             Set pXMLDOMNodeConnectionProperty = pDOMDocument.createNode(MSXML2.NODE_ELEMENT, "connectionProperty", "")
'575:             Set pXMLDOMNodeConnectionProperty = pXMLDOMNodeMetadataChild.appendChild(pXMLDOMNodeConnectionProperty)
'576:             Call modCommon.AddNodeAttribute(pXMLDOMNodeConnectionProperty, "name", CStr(pNames(pIndex)))
'577:             Call modCommon.AddNodeAttribute(pXMLDOMNodeConnectionProperty, "value", CStr(pValues(pIndex)))
'578:         End If
'579:     Next pIndex
'    '
'    Exit Sub
'ErrorHandler:
'583:     Call HandleError(False, "WriteGeodatabaseDesignerHeader " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetXSLFilename() As String
'    On Error GoTo ErrorHandler
'    '
'589:     Const pXSLFolderName As String = "XSL"
'590:     Const pXSLFileName As String = "Geodatabase Designer.xsl"
'    '
'    Dim pXSLPathName As String
'    Dim pFSO As Scripting.FileSystemObject
'    '
'595:     Set pFSO = New FileSystemObject
'596:     pXSLPathName = pFSO.BuildPath(pFSO.GetFolder(App.Path).ParentFolder.Path, pXSLFolderName)
'597:     pXSLPathName = pFSO.BuildPath(pXSLPathName, pXSLFileName)
'    '
'599:     GetXSLFilename = pXSLPathName
'    '
'    Exit Function
'ErrorHandler:
'603:     Call HandleError(False, "GetXSLFilename " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

''===============
'' FORM ROUTINES
''===============

'Public Sub EnableFrame(ByRef pFrame As VB.Frame, _
'                       ByRef pState As Boolean)
'    On Error GoTo ErrorHandler
'    '-----------------------------------------------------------
'    ' Enables or Disables all the controls in the parsed FRAME.
'    '-----------------------------------------------------------
'    Dim pControl As VB.Control
'    Dim pObject As Object
'    Dim pFrameTest As VB.Frame
'    Dim pForm As VB.Form
'    '
'621:     pFrame.Enabled = pState
'622:     Set pForm = pFrame.Parent
'    '
'624:     For Each pControl In pForm.Controls
'625:         If Not (TypeOf pControl Is MSComctlLib.ImageList) Then
'626:             Set pObject = pControl.Container
'627:             If TypeOf pObject Is VB.Frame Then
'628:                 Set pFrameTest = pObject
'629:                 If pFrameTest.Name = pFrame.Name Then
'630:                     If TypeOf pControl Is VB.Frame Then
'631:                         Call EnableFrame(pControl, pState)
'632:                     Else
'633:                         pControl.Enabled = pFrame.Enabled
'634:                     End If
'635:                 End If
'636:             End If
'637:         End If
'638:     Next
'    '
'    Exit Sub
'ErrorHandler:
'642:     Call HandleError(False, "EnableFrame " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Sub EnableForm(ByRef pForm As VB.Form, _
'                      ByRef blnState As Boolean)
'    On Error GoTo ErrorHandler
'    '
'    Dim pControl As VB.Control
'    '
'651:     For Each pControl In pForm.Controls
'652:         If (TypeOf pControl Is VB.CheckBox) Or _
'            (TypeOf pControl Is VB.ComboBox) Or _
'            (TypeOf pControl Is VB.CommandButton) Or _
'            (TypeOf pControl Is VB.Frame) Or _
'            (TypeOf pControl Is VB.Label) Or _
'            (TypeOf pControl Is VB.ListBox) Or _
'            (TypeOf pControl Is VB.OptionButton) Or _
'            (TypeOf pControl Is VB.PictureBox) Or _
'            (TypeOf pControl Is VB.TextBox) Or _
'            (TypeOf pControl Is RichTextLib.RichTextBox) Or _
'            (TypeOf pControl Is MSComctlLib.TabStrip) Then
'663:             If blnState Then
'664:                 pControl.Enabled = True
'665:                 pControl.MousePointer = VBRUN.MousePointerConstants.vbDefault
'666:             Else
'667:                 pControl.Enabled = False
'668:                 pControl.MousePointer = VBRUN.MousePointerConstants.vbHourglass
'669:             End If
'670:         End If
'671:     Next
'    '
'673:     Call pForm.Refresh
'674:     DoEvents
'    '
'    Exit Sub
'ErrorHandler:
'678:     Call HandleError(False, "EnableForm " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function PadZero(ByRef pLine As String, _
'                         ByRef pLength As Long) As String
'    On Error GoTo ErrorHandler
'    '
'685:     PadZero = String(CLng(pLength - Len(pLine)), "0") & pLine
'    '
'    Exit Function
'ErrorHandler:
'689:     Call HandleError(False, "PadZero " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub LaunchHelp()
'    On Error GoTo ErrorHandler
'    '
'    Dim pFSO As Scripting.FileSystemObject
'    Dim pFolder As Scripting.Folder
'    Dim pHelpLocation As String
'    '
'699:     Set pFSO = New Scripting.FileSystemObject
'    '
'701:     Set pFolder = pFSO.GetFolder(App.Path)
'702:     Set pFolder = pFolder.ParentFolder
'    '
'704:     If pFSO.FolderExists(pFSO.BuildPath(pFolder.Path, HELP_FOLDER)) Then
'705:         Set pFolder = pFSO.GetFolder(pFSO.BuildPath(pFolder.Path, HELP_FOLDER))
'        '
'707:         If pFSO.FileExists(pFSO.BuildPath(pFolder.Path, HELP_FILE)) Then
'708:             Call modCommon.SubmitFileToOS(pFSO.BuildPath(pFolder.Path, HELP_FILE))
'709:         Else
'710:             Call MsgBox("Error: Help file is missing from:" & vbCrLf & App.Path & "\" & HELP_FOLDER & " \ " & HELP_FILE)
'711:         End If
'712:     Else
'713:         Call MsgBox("Error: Help file is missing from:" & vbCrLf & App.Path & "\" & HELP_FOLDER & " \ " & HELP_FILE)
'714:     End If
'    '
'    Exit Sub
'ErrorHandler:
'718:     Call HandleError(False, "LaunchHelp " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetTimeStamp() As String
'    On Error GoTo ErrorHandler
'    '
'724:     GetTimeStamp = CStr(Year(Date)) & "/" & _
'                    PadZero(CStr(Month(Date)), 2) & "/" & _
'                    PadZero(CStr(Day(Date)), 2) & " " & _
'                    PadZero(CStr(Hour(Time)), 2) & ":" & _
'                    PadZero(CStr(Minute(Time)), 2) & ":" & _
'                    PadZero(CStr(Second(Time)), 2)
'    '
'    Exit Function
'ErrorHandler:
'733:     Call HandleError(False, "GetTimeStamp " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Sub KeyPress(ByRef pKey As Long)
'    On Error GoTo ErrorHandler
'    '
'739:     Call keybd_event(pKey, 0, 0, 0)
'    '
'    Exit Sub
'ErrorHandler:
'743:     Call HandleError(False, "KeyPress " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Sub

'Public Function GetDataset(ByRef pWorkspace As esriGeoDatabase.IWorkspace, _
'                           ByRef strDatasetName As String, _
'                           ByRef pEsriDataset As esriGeoDatabase.esriDatasetType) As esriGeoDatabase.IDataset
'    On Error GoTo ErrorHandler
'    '
'    Dim pFeatureWorkspace As esriGeoDatabase.IFeatureWorkspace
'    Dim pTopologyWorkspace As esriGeoDatabase.ITopologyWorkspace
'    '
'    Dim pEnumDataset As IEnumDataset
'    Dim pFeatureDataset As IFeatureDataset
'    Dim pNetworkCollection As INetworkCollection
'    '
'    Select Case pEsriDataset
'    Case esriGeoDatabase.esriDatasetType.esriDTFeatureDataset
'760:         Set pFeatureWorkspace = pWorkspace
'        '
'        On Error GoTo DatasetOpenError
'763:         Set GetDataset = pFeatureWorkspace.OpenFeatureDataset(strDatasetName)
'        On Error GoTo ErrorHandler
'    Case esriGeoDatabase.esriDatasetType.esriDTFeatureClass
'766:         Set pFeatureWorkspace = pWorkspace
'        '
'        On Error GoTo DatasetOpenError
'769:         Set GetDataset = pFeatureWorkspace.OpenFeatureClass(strDatasetName)
'        On Error GoTo ErrorHandler
'    Case esriGeoDatabase.esriDatasetType.esriDTGeometricNetwork
'772:         Set pEnumDataset = pWorkspace.Datasets(esriDTFeatureDataset)
'773:         Set pFeatureDataset = pEnumDataset.Next
'        '
'775:         Do Until pFeatureDataset Is Nothing
'776:             Set pNetworkCollection = pFeatureDataset
'            On Error GoTo DatasetOpenError
'778:             Set GetDataset = pNetworkCollection.GeometricNetworkByName(strDatasetName)
'            On Error GoTo ErrorHandler
'780:             If Not (GetDataset Is Nothing) Then
'781:                 Exit Do
'782:             End If
'            '
'784:             Set pFeatureDataset = pEnumDataset.Next
'785:         Loop
'    Case esriGeoDatabase.esriDatasetType.esriDTRelationshipClass
'787:         Set pFeatureWorkspace = pWorkspace
'        '
'        On Error GoTo DatasetOpenError
'790:         Set GetDataset = pFeatureWorkspace.OpenRelationshipClass(strDatasetName)
'        On Error GoTo ErrorHandler
'    Case esriGeoDatabase.esriDatasetType.esriDTTable
'793:         Set pFeatureWorkspace = pWorkspace
'        '
'        On Error GoTo DatasetOpenError
'796:         Set GetDataset = pFeatureWorkspace.OpenTable(strDatasetName)
'        On Error GoTo ErrorHandler
'    Case esriGeoDatabase.esriDatasetType.esriDTTopology
'799:         Set pTopologyWorkspace = pWorkspace
'        '
'        On Error GoTo DatasetOpenError
'802:         Set GetDataset = pTopologyWorkspace.OpenTopology(strDatasetName)
'        On Error GoTo ErrorHandler
'    Case Else
'        '-------------------------
'        ' Unsupported DatasetType
'        '-------------------------
'808:     End Select
'    '
'    Exit Function
'DatasetOpenError:
'812:     Call Err.Clear
'813:     Resume Next
'    '
'    Exit Function
'ErrorHandler:
'817:     Call HandleError(False, "GetDataset " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function

'Public Function GetSubstring(ByRef pComplete As String, _
'                              ByRef pStart As String, _
'                              ByRef pEnd As String) As String

'    On Error GoTo ErrorHandler
'    '
'    Dim pIndexStart As Long
'    Dim pIndexEnd As Long
'    '
'829:     pIndexStart = InStr(1, pComplete, pStart) + Len(pStart)
'830:     If pEnd = "" Then
'831:         pIndexEnd = Len(pComplete) + 1
'832:     Else
'833:         pIndexEnd = InStr(pIndexStart, pComplete, pEnd)
'834:     End If
'    '
'836:     GetSubstring = Mid(pComplete, pIndexStart, (pIndexEnd - pIndexStart))
'    '
'    Exit Function
'ErrorHandler:
'840:     Call HandleError(False, "GetSubstring " & MODULE_NAME & " (" & CStr(Erl) & ")", Err.Number, Err.Source, Err.Description)
'End Function



