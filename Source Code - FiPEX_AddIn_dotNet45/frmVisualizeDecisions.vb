Imports System.ComponentModel
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesOleDB
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports System.Text.RegularExpressions

Public Class frmVisualizeDecisions

    Private m_app As ESRI.ArcGIS.Framework.IApplication
    Private pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
    'Private m_pOutGxDatabase As IGxDatabase
    Private m_sCategory2 As String
    Private m_sCategory1 As String
    'Private m_pFWorkspace1 As IFeatureWorkspace
    Private m_pFWorkspace As IFeatureWorkspace
    Private m_sTableName1 As String
    Private m_sTableName2 As String
    ' Private progressThread As Threading.Thread
    Private m_sProgressString As String
    Private m_iProgressPercent As Integer


    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean


        Try
            m_app = m_application
            pMxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            m_FiPEx = FishPassageExtension.GetExtension

            m_UtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize Form")
            Return False
        End Try
    End Function



    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
    End Sub

    Private Sub cmdHighlight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHighlight.Click

        ' Me.CheckForIllegalCrossThreadCalls = False
        'progressThread = New Threading.Thread(AddressOf updateProgress)
        Call Highlightstuff()

    End Sub
    Private Sub Highlightstuff()
       
        ' Subroutine Highlight Stuff
        ' Created July 25, 2012
        ' highlights a list of features using Utility Network Analyst's highlight function

        ' 1. select non-spatial table list 
        ' 2. have user select the budget amount
        ' 3. read all rows that have a decision option number of 2 into a list object
        ' 4. use the list object to highlight features in the current network (check that network features exist). 


        ' =============== Section 1 ================
        '    1. select non-spatial table 
        '    
        If txtTable.Text = "" Then
            MsgBox("You must select Table 1 for input.")

            'exitsubroutine()
            Exit Sub
        End If

    
        Dim pTable As ITable
        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening table 1. Exiting." + ex.Message)

            Exit Sub
        End Try

        Dim iBarrierEIDField, iOptionField, iBudgetfield As Integer

        'check table 
        If pTable.FindField("BarrierEID") = -1 Then
            MsgBox("Cannot find field 'BarrierEID' in Table. Now Exiting.")

            'exitsubroutine()
            Exit Sub
        Else
            iBarrierEIDField = pTable.FindField("BarrierEID")
        End If

        If pTable.FindField("OptionNum") = -1 Then
            MsgBox("Cannot find field 'OptionNum' in Table. Now Exiting.")

            'exitsubroutine()
            Exit Sub
        Else
            iOptionField = pTable.FindField("OptionNum")
        End If

        If pTable.FindField("Budget") = -1 Then
            MsgBox("Cannot find field 'Budget' in Table. Now Exiting.")

            'exitsubroutine()
            Exit Sub
        Else
            iBudgetField = pTable.FindField("Budget")
        End If


        ' ============================= Section 2 ========================
        ' 2. Get the budget amount. 
        Dim iBudget, iTotalBudget As Integer
        If lstBudgets.SelectedItems.Count = 0 Then
            MsgBox("You must select a budget amount.")
            Exit Sub
        Else
            Try
                iBudget = Convert.ToInt32(lstBudgets.SelectedItem.ToString)
            Catch ex As Exception
                MsgBox("Trouble converting budget to integer value / variable.")
            End Try
        End If


        Dim iTabRows As Integer
        iTabRows = pTable.RowCount(Nothing)
        Dim pCursor As ICursor
        pCursor = pTable.Search(Nothing, False)

        Dim pRow As IRow

        ' ============================= Section 3 ========================
        ' 2. read all rows that have a decision option number of 2 into a list object
        ' read tables 1 and 2 into feature cursors
        ' read all table 2 into a list object

        ' Opentable and read the lines into an object
        ' Chose: read table 1 line by line and compare to table 2.  Read all of table 2 into a list object
        ' at the beginning of the sub.  
        ' custom object used: GLPKDecisionsObject

        Dim pGLPKDecisionsObjectTAB As New GLPKDecisionsObject(Nothing, Nothing, Nothing, Nothing)
        Dim lGLPKDecisionsObjectTAB As New List(Of GLPKDecisionsObject)

        Dim iBarrierEID As Integer
        Dim iOption As Integer
        Dim iBudgetTable As Integer


        Dim pField As IField
        pRow = pCursor.NextRow

        ' Loop through each row
        Do Until pRow Is Nothing

            pField = pRow.Fields.Field(iBarrierEIDField)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField.Type = esriFieldType.esriFieldTypeOID Then
                iBarrierEID = pRow.Value(iBarrierEIDField)
            Else
                MsgBox("Error reading Table. Could not convert the barrier EID to type 'integer'. Now Exiting.")
                'BackgroundWorker1.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField = pRow.Fields.Field(iBudgetfield)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField.Type = esriFieldType.esriFieldTypeDouble Then
                iBudgetTable = Convert.ToInt32(pRow.Value(iBudgetfield))
            Else
                MsgBox("Error reading Table 1. Could not convert the budget to type 'integer'. Now Exiting.")
                'BackgroundWorker1.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField = pRow.Fields.Field(iOptionField)
            If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
                iOption = pRow.Value(iOptionField)
            Else
                MsgBox("Error reading Table 1. Could not convert the optionnum to type 'integer'. Now Exiting.")
                'BackgroundWorker1.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            ' add to new object
            If iOption > 1 And iBudgetTable = iBudget Then
                pGLPKDecisionsObjectTAB = New GLPKDecisionsObject(Nothing, Nothing, iBarrierEID, iOption)
                lGLPKDecisionsObjectTAB.Add(pGLPKDecisionsObjectTAB)
                iTotalBudget = iTotalBudget + iBudget
            End If

            'pRow1 = Nothing
            pRow = pCursor.NextRow

        Loop

        If lGLPKDecisionsObjectTAB.Count = 0 Then
            MsgBox("No decisions found in table, only values of '1' or 'do nothing'. Now Exiting.")
            Exit Sub
        End If

        ' ====================== Section 4 ========================
        ' Get EID's for highlighting
 
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pUNAExt As IUtilityNetworkAnalysisExt

        Try
            pUNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
        Catch ex As Exception
            MsgBox("Trouble getting Utility Network Analyst Extension")
            Exit Sub
        End Try

        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        pNetworkAnalysisExt = CType(pUNAExt, INetworkAnalysisExt)

        If pNetworkAnalysisExt.NetworkCount = 0 Then
            MsgBox("No Networks Loaded. Now Exiting")
            Exit Sub
        End If

        ' iEnumNetEIDBuilderGEN
        '             As New EnumNetEIDArray
        ' iEnumNetEIDBuilder
        '             AS New EnumNetEIDArray
        ' supposedly these two are the same - but preferred to use GEN version
        ' http://edndoc.esri.com/arcobjects/9.2/ComponentHelp/esriGeoDatabase/IEnumNetEIDBuilder.htm
        ' iEnumNetEID
        '             CType from EnumNETEIDArray

        Dim pResultsEIDsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        pResultsEIDsGEN.Network = pNetworkAnalysisExt.CurrentNetwork.Network
        pResultsEIDsGEN.ElementType = esriElementType.esriETJunction

        For i = 0 To lGLPKDecisionsObjectTAB.Count - 1
            pResultsEIDsGEN.Add(lGLPKDecisionsObjectTAB(i).BarrierEID)
        Next

        Dim pResultsJunctions As IEnumNetEID
        pResultsJunctions = CType(pResultsEIDsGEN, IEnumNetEID)

        ' do edges just because
        ' edges were not returned -- create an empty enumeration
        pResultsEIDsGEN = New EnumNetEIDArray
        pResultsEIDsGEN.Network = pNetworkAnalysisExt.CurrentNetwork.Network
        pResultsEIDsGEN.ElementType = esriElementType.esriETEdge
        Dim pResultEdges As IEnumNetEID
        pResultEdges = CType(pResultsEIDsGEN, IEnumNetEID)


        Try

            pNetworkAnalysisExtResults = CType(pUNAExt, INetworkAnalysisExtResults)
            pNetworkAnalysisExtResults.ClearResults()
            pNetworkAnalysisExtResults.ResultsAsSelection = True
            'pNetworkAnalysisExtResults.SetResults(pResultsJunctions, Nothing)
            pNetworkAnalysisExtResults.CreateSelection(pResultsJunctions, pResultEdges)
        Catch ex As Exception
            MsgBox("Problem encountered setting results display. Exiting.")
            Exit Sub
        End Try

        lblHighlighted.Text = "Highlighted: " + Convert.ToString(pResultsJunctions.Count)

    End Sub
    Private Function GetWorkspace(ByVal sGDB As String) As IFeatureWorkspace
        ' This function gets a reference to the workspace specified by 
        ' the user in the options menu - used for DBF table output

        Dim pFWorkspace As IFeatureWorkspace
        Dim pWorkspace As IWorkspace

        ' Not sure if workspace name is actually necessary anymore.
        Dim pWorkspaceName As IWorkspaceName = New WorkspaceName
        pWorkspaceName.PathName = sGDB
        Dim pWSFactory As IWorkspaceFactory

        If InStr(sGDB, ".mdb") <> 0 Then
            pWSFactory = New AccessWorkspaceFactory
        ElseIf InStr(sGDB, ".gdb") <> 0 Then
            pWSFactory = New FileGDBWorkspaceFactory
        Else ' assume a folder / textfile workspace
            pWSFactory = New TextFileWorkspaceFactory
            'GetWorkspace = pFWorkspace
            'Exit Function
        End If

        Try
            pWorkspace = pWSFactory.OpenFromFile(sGDB, 0)
            pFWorkspace = CType(pWorkspace, IFeatureWorkspace)
            GetWorkspace = pFWorkspace
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("The geodatabase for DBF output was not found.  Please check the DST options and add an existing geodatabase." + ex.Message, "Geodatbase Not Found")
            GetWorkspace = Nothing
            Exit Function
        End Try

    End Function


    'API REGION FOR DRAG AND DROP TEXT BOX
    'API code that needs to be pasted in the general section of the form.
    ' From here: http://forums.esri.com/Thread.asp?c=93&f=993&t=153825

    Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long

    Private Declare Function CreateStreamOnHGlobal Lib "Ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Boolean, ByRef pStream As IStream) As Long
    Private Declare Function GetHGlobalFromStream Lib "Ole32" (ByVal pStm As IStream, ByRef hGlobal As Long) As Long

    'Textbox subs
    Private Sub Text1_OLEDragDrop(ByVal Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        Try

            Dim cfFormat As Long
        cfFormat = RegisterClipboardFormat("My Layer")
        OpenClipboard(0)

        Dim hGlobal As Long
        hGlobal = GetClipboardData(cfFormat)

        CloseClipboard()

        If hGlobal = 0 Then
            Exit Sub
        End If

        Dim pStm As IStream
        CreateStreamOnHGlobal(hGlobal, False, pStm)
        Dim pObjectStream As IObjectStream
        pObjectStream = New ObjectStream
        pObjectStream.Stream = pStm

        Dim pPropSet As IPropertySet
        pPropSet = New PropertySet

        Dim pPersistStream As IPersistStream
        pPersistStream = pPropSet
        pPersistStream.Load(pObjectStream)
        Dim pLayer As ILayer
        pLayer = pPropSet.GetProperty("Layer")

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = pLayer

        If pFeatLayer.DataSourceType = "Shapefile Feature Class" Then

            Dim pDS As IDataset
            pDS = pFeatLayer.FeatureClass
                txtTable.Text = pLayer.Name
        Else
            MsgBox("Please select a Shapefile.")
        End If

        Exit Sub
        Catch ex As Exception
            MsgBox("Text1_OLEDragDrop Issue")
            Exit Sub
        End Try

    End Sub

    Private Sub Text1_OLEDragOver(ByVal Data As DataObject, ByVal Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal State As Integer)
        Try

            Dim pMxDoc As IMxDocument
        Dim pLayer As ILayer
        Dim pApp As IApplication
        pApp = New AppRef
        pMxDoc = pApp.Document
        pLayer = pMxDoc.SelectedLayer
        If pLayer Is Nothing Then Exit Sub

        Dim cfFormat As Long
        cfFormat = RegisterClipboardFormat("My Layer")
        Dim pStm As IStream
        CreateStreamOnHGlobal(0, False, pStm)
        OpenClipboard(0)

        Dim pObjectStream As IObjectStream
        pObjectStream = New ObjectStream
        pObjectStream.Stream = pStm

        Dim pPropSet As IPropertySet
        pPropSet = New PropertySet
        pPropSet.SetProperty("Layer", pLayer)
        Dim pPersistStream As IPersistStream
        pPersistStream = pPropSet
        pPersistStream.Save(pObjectStream, 0)
        Dim hGlobal As Long
        GetHGlobalFromStream(pStm, hGlobal)
        SetClipboardData(cfFormat, hGlobal)
        CloseClipboard()

            Exit Sub

        Catch ex As Exception
            MsgBox("Error in Drag and Drop")
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
     Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterTables
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Table"
        pGxDialog.AllowMultiSelect = False

        pGxDialog.DoModalOpen(0, pEnumGx)

        If pEnumGx Is Nothing Then Exit Sub

        'Me.ZOrder(0) 'sendtoback?
        pEnumGx.Reset()

        Dim pGxObject As IGxObject = pEnumGx.Next
        'pGxName = CType(pGxObject, IGxDisplayName)

        'If pGxName Is Nothing Then
        '    Exit Sub
        'Else
        '    txtTable1.Text = pGxName.DisplayName
        'End If
        If pGxObject Is Nothing Then
            Exit Sub
        End If

        Dim sCategory As String
        sCategory = pGxObject.Category.ToString
        If sCategory <> "File Geodatabase Table" And sCategory <> "Text File" And sCategory <> "Personal Geodatabase Table" Then
            MsgBox("Input table must be a File Geodatabase Table, Personal Geodatabase Table, CSV, or Text File. Input cannot be of type " + sCategory)
            txtTable.Text = ""
            Exit Sub
        End If

        Dim sFullTableName As String
        sFullTableName = pGxObject.FullName
        Dim sTableDirectory As String
        Dim iLastBackSlash As Integer
        iLastBackSlash = sFullTableName.LastIndexOf("\")
        sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        m_sTableName1 = sFullTableName.Substring(iLastBackSlash + 1)
        m_pFWorkspace = GetWorkspace(sTableDirectory)

        If pGxObject Is Nothing Then
            MsgBox("You must select an input table")
            Exit Sub
        Else
            txtTable.Text = sFullTableName
        End If

        Dim pTable As ITable
        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening table 1. Exiting." + ex.Message)
            Exit Sub
        End Try
        Try
            Dim pFields As IFields
            pFields = pTable.Fields

            Dim sFieldname As String
            Dim pCursor As ICursor
            Dim pDataStats As IDataStatistics
            Dim vVar As Object


            ' clear the values listbox
            lstBudgets.Items.Clear()

            ' Get selected field
            sFieldname = "Budget"

            pCursor = pTable.Search(Nothing, False)

            ' Setup the datastatistics and get the unique values of the "Id" field
            pDataStats = New DataStatistics
            Dim pEnumerator As IEnumerator

            With pDataStats
                .Cursor = pCursor
                .Field = sFieldname
                pEnumerator = .UniqueValues
            End With

            pEnumerator.Reset()
            pEnumerator.MoveNext()

            vVar = pEnumerator.Current

            While Not IsNothing(vVar)
                lstBudgets.Items.Add(vVar.ToString)
                pEnumerator.MoveNext()
                vVar = pEnumerator.Current
            End While

        Catch ex As Exception
            MsgBox("Error trying to populate the budget list box.")
            Exit Sub
        End Try

    End Sub

End Class