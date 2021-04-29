Imports System.ComponentModel
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesOleDB
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Geoprocessing

Imports System.Text.RegularExpressions

Public Class frmVisualizeDecisionsAndNet
    Private m_app As ESRI.ArcGIS.Framework.IApplication
    Private pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
    'Private m_pOutGxDatabase As IGxDatabase
    Private m_sCategory2 As String
    Private m_sCategory1 As String
    'Private m_pFWorkspace1 As IFeatureWorkspace

    Private m_pFWorkspace As IFeatureWorkspace
    Private m_pFWorkspace2 As IFeatureWorkspace
    Private m_pFWorkspace3 As IFeatureWorkspace

    Private m_sTableName1 As String ' Results Table
    Private m_sTableName2 As String ' Decisions Table
    Private m_sTableName3 As String ' Connectivity Table

    ' Private progressThread As Threading.Thread
    Private m_sProgressString As String
    Private m_iProgressPercent As Integer

    'Private m_dHighestBudget As Double
    'Private m_dLowestBudget As Double
    Private m_lBudgets As List(Of Double)
    Private m_dBudget As Double = 0
    Private m_lEIDs As List(Of Integer)
    Private m_iBudgetCount As Integer
    Private m_sUniqueCode As String
    Private m_bUndirected As Boolean
    Private m_bGRB As Boolean
    Private m_lLineLayersFields As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
    Private m_lPolyLayersFields As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
    Private m_lSinks As List(Of Integer) = New List(Of Integer)
    Private m_lSelectAndUpdateFeaturesObject As List(Of SelectAndUpdateFeaturesObject) = New List(Of SelectAndUpdateFeaturesObject)
    Dim m_iTreatmentFieldIndex As Integer
    'Private m_iTreatmentFieldIndex As Integer

    Private Property bFoundEID As Boolean

    
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
    Private Sub frmVisualizeDecisionsAndNet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        txtDecisionsTab.Enabled = False
        txtResultsTab.Enabled = False
        cmdBrowseDecisionsTab.Enabled = False
        txtConnectivityTab.Enabled = False
        cmdBrowseConnectivityTab.Enabled = False
        grpLines.Enabled = False
        grpPolygons.Enabled = False
        lstBudgets.Enabled = False
        grpStep6.Enabled = False

        ' =============== Load Lines and Polygons from network =====================
        ' check current network 
        '                       - make sure one is there
        ' read FIPEX lines settings 
        '                       - make sure network is same as FIPEX by 
        '                         cross ref participating lines with those in FIPEX
        '                       - load network lines from FIPEX to checklist box
        ' read FIPEX polygons settings
        '                       - load polygon lines from FIPEX to checklist box
        ' 
        ' 
        ' 

        ' =============== Section 1: Check for a network =====================
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        If m_UtilityNetworkAnalysisExt Is Nothing Then
            MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
            Me.Close()
        Else

            pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
            If pNetworkAnalysisExt.NetworkCount = 0 Then
                MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
                Me.Close()
            End If
        End If
        Dim pGeometricNetwork As IGeometricNetwork
        Try
            pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Catch ex As Exception
            MsgBox("Trouble getting current geometric network.  Must have an active geometric network in ArcMap to use this toolset. Exiting.")
            Me.Close()
        End Try


        ' get all lines feature classes in network
        Dim pEnumNetSimpleEdges As IEnumFeatureClass
        Try
            pEnumNetSimpleEdges = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTSimpleEdge)
        Catch ex As Exception
            MsgBox("Trouble getting the network Simple Edge Feature Classes from the geometric network. Exiting.")
            Me.Close()
        End Try

        If pEnumNetSimpleEdges Is Nothing Then
            MsgBox("No simple Edge Lines feature classes were returned for this network.  FIPEX only uses Simple Edges for analyses.  Exiting.")
            Me.Close()
        End If

        Dim pFeatureClass As IFeatureClass
        Dim pFeatureLayer As IFeatureLayer
        pEnumNetSimpleEdges.Reset()
        pFeatureClass = pEnumNetSimpleEdges.Next
        If pFeatureClass Is Nothing Then
            MsgBox("No simple Edge Lines feature classes were returned for this network.  FIPEX only uses Simple Edges for analyses.  Exiting.")
            Me.Close()
        End If

        ' =============== Section 2: Load lines and polys from FIPEX =====================
        '  Right now Options saves layer 'names' from TOC, not FCIDs or anything more specific
        '  You can't retrieve map layer 'names' from network feature classes - only 'alias' names
        '  So to compare network lines with FIPEX lines there needs to be a retrieval of FCIDs for the layers
        '  using the line 'names' stored in the FIPEX extension.  
        '  Looping through ArcMap will be done to get the 'name' and then fCID of each line stored in extension. 
        '  In effect this means that ALL ACTIVE FIPEX LINES layers must be present in the map in order for this 
        '  analysis to see them.  

        Dim m_iLinesCount, m_iPolysCount As Integer
        Dim sLineLayer, sPolyLayer As String
        Dim LineLayerFCIDAndPermField As New LayersAndFCIDAndCumulativePassField(Nothing, Nothing, Nothing) ' layer to hold parameters
        Dim i, m, j As Integer
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim bLinesMatchFIPEXNetwork As Boolean
        Dim bLinesMatchMapFIPEX, bPolysMatchMapFIPEX As Boolean
        Dim iFIPEX_FCID As Integer
        Dim pLayer As ILayer

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Me.Close()
        End Try

        Try
            ' Get loaded settings
            If m_FiPEx.m_bLoaded = True Then
                m_iLinesCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numLines"))
                If m_iLinesCount > 0 Then
                    For j = 0 To m_iLinesCount - 1

                        Try
                            sLineLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncLine" + j.ToString)) ' get line layer
                        Catch ex As Exception
                            MsgBox("Could not get line layer from FIPEX: IncLine" + j.ToString + ". Now exiting.")
                            Me.Close()
                        End Try

                        bLinesMatchMapFIPEX = False
                        ' Loop through all map layers
                        For i = 0 To pMap.LayerCount - 1
                            If pMap.Layer(i).Valid = True Then
                                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                                    pLayer = pMap.Layer(i)
                                    pFeatureLayer = CType(pLayer, IFeatureLayer)
                                    If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge _
                                        Or pFeatureLayer.FeatureClass.ShapeType = esriShapeType.esriShapePolyline Then

                                        ' If the map layer name matches the FIPEX layer name
                                        If pFeatureLayer.Name = sLineLayer Then
                                            bLinesMatchMapFIPEX = True
                                            iFIPEX_FCID = pFeatureLayer.FeatureClass.FeatureClassID

                                            ' loop through current network layers and find a match
                                            bLinesMatchFIPEXNetwork = False
                                            pEnumNetSimpleEdges.Reset()
                                            pFeatureClass = pEnumNetSimpleEdges.Next
                                            Do Until pFeatureClass Is Nothing
                                                If pFeatureClass.FeatureClassID = iFIPEX_FCID Then
                                                    bLinesMatchFIPEXNetwork = True
                                                End If
                                                pFeatureClass = pEnumNetSimpleEdges.Next
                                            Loop
                                            If bLinesMatchFIPEXNetwork = True Then
                                                lstLineLayers.Items.Add(sLineLayer)
                                                LineLayerFCIDAndPermField = New LayersAndFCIDAndCumulativePassField(sLineLayer, iFIPEX_FCID, "NotSet")
                                                m_lLineLayersFields.Add(LineLayerFCIDAndPermField)
                                                Exit For
                                            Else
                                                MsgBox("The line layer " + sLineLayer + " loaded in FIPEX does not participate in the current active geometric network. " _
                                                       + "Please check settings in ArcMap and FIPEX Options to continue.")
                                                Me.Close()
                                            End If
                                        End If ' layer name match
                                    End If     ' simple edge
                                End If         ' LayerType Check
                            End If             ' layer is valid
                        Next                   ' Map Layer
                        If bLinesMatchMapFIPEX = False Then
                            MsgBox("The lines layer " + sLineLayer + "saved in FIPEX cannot be found in the active TOC. This layer will be omitted from analysis.")
                        End If

                    Next ' network line saved in FIPEX
                Else
                    MsgBox("There are currently no network lines participating in FIPEX analysis.  Please open FIPEX Options and check lines layer.")
                    Me.Close()
                End If ' there are network lines saved in FIPEX

                m_iPolysCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numPolys"))
                If m_iPolysCount > 0 Then
                    For j = 0 To m_iPolysCount - 1
                        Try
                            sPolyLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncPoly" + j.ToString)) ' get line layer
                        Catch ex As Exception
                            MsgBox("Could not get poly layer from FIPEX: IncPoly" + j.ToString + ". Now exiting.")
                            Me.Close()
                        End Try
                    Next
                    bPolysMatchMapFIPEX = False

                    ' Loop through all map layers
                    For i = 0 To pMap.LayerCount - 1
                        If pMap.Layer(i).Valid = True Then
                            If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                                pLayer = pMap.Layer(i)
                                pFeatureLayer = CType(pLayer, IFeatureLayer)
                                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                                    ' If the map layer name matches the FIPEX layer name
                                    If pFeatureLayer.Name = sPolyLayer Then
                                        bPolysMatchMapFIPEX = True
                                        iFIPEX_FCID = pFeatureLayer.FeatureClass.FeatureClassID
                                        lstPolyLayers.Items.Add(sPolyLayer)
                                        LineLayerFCIDAndPermField = New LayersAndFCIDAndCumulativePassField(sPolyLayer, iFIPEX_FCID, "NotSet")
                                        m_lPolyLayersFields.Add(LineLayerFCIDAndPermField)
                                        Exit For
                                    End If ' layer name match
                                End If     ' simple edge
                            End If         ' LayerType Check
                        End If             ' layer is valid
                    Next                   ' Map Layer
                    If bPolysMatchMapFIPEX = False Then
                        MsgBox("The polys layer " + sPolyLayer + "saved in FIPEX cannot be found in the active TOC. This layer will be omitted from analysis.")
                    End If
                End If ' there are polygons included in FIPEX options settings
            Else
                MsgBox("FIPEX settings have not been loaded.  Please open FIPEX, adjust settings, and save map document. Exiting.")
                Me.Close()
            End If ' FiPEX settings are loaded
        Catch ex As Exception
            MsgBox("Problem getting status of Fish Passage Extension. Lines and Polygons layers included in analysis are not loaded. " + ex.Message)
            Exit Sub
        End Try

        lblStep1Status.Text = ""
        lblStep2Status.Text = ""
        lblStep3Status.Text = ""
        lblStep4Status.Text = ""
        lblStep5Status.Text = ""
        lblStep6aStatus.Text = ""
        lblStep6bStatus.Text = ""
        lblStep1Warning.Text = ""
        lblStep2warning.Text = ""
        lblStep3Warning.Text = ""
        lblStep4Warning.Text = ""
        lblStep5Warning.Text = ""

    End Sub

    'Private Sub Highlightstuff()

    '    ' Subroutine Highlight Stuff
    '    ' Created July 25, 2012
    '    ' highlights a list of features using Utility Network Analyst's highlight function

    '    ' 1. select non-spatial table list 
    '    ' 2. have user select the budget amount
    '    ' 3. read all rows that have a decision option number of 2 into a list object
    '    ' 4. use the list object to highlight features in the current network (check that network features exist). 


    '    ' =============== Section 1 ================
    '    '    1. select non-spatial table 
    '    '    
    '    If txtTable.Text = "" Then
    '        MsgBox("You must select Table 1 for input.")

    '        'exitsubroutine()
    '        Exit Sub
    '    End If


    '    Dim pTable As ITable
    '    Try
    '        pTable = m_pFWorkspace.OpenTable(m_sTableName1)
    '    Catch ex As Exception
    '        MsgBox("Trouble opening table 1. Exiting." + ex.Message)

    '        Exit Sub
    '    End Try

    '    Dim iBarrierEIDField, iOptionField, iBudgetfield As Integer

    '    'check table 
    '    If pTable.FindField("BarrierEID") = -1 Then
    '        MsgBox("Cannot find field 'BarrierEID' in Table. Now Exiting.")

    '        'exitsubroutine()
    '        Exit Sub
    '    Else
    '        iBarrierEIDField = pTable.FindField("BarrierEID")
    '    End If

    '    If pTable.FindField("OptionNum") = -1 Then
    '        MsgBox("Cannot find field 'OptionNum' in Table. Now Exiting.")

    '        'exitsubroutine()
    '        Exit Sub
    '    Else
    '        iOptionField = pTable.FindField("OptionNum")
    '    End If

    '    If pTable.FindField("Budget") = -1 Then
    '        MsgBox("Cannot find field 'Budget' in Table. Now Exiting.")

    '        'exitsubroutine()
    '        Exit Sub
    '    Else
    '        iBudgetField = pTable.FindField("Budget")
    '    End If


    '    ' ============================= Section 2 ========================
    '    ' 2. Get the budget amount. 
    '    Dim iBudget, iTotalBudget As Integer
    '    If lstBudgets.SelectedItems.Count = 0 Then
    '        MsgBox("You must select a budget amount.")
    '        Exit Sub
    '    Else
    '        Try
    '            iBudget = Convert.ToInt32(lstBudgets.SelectedItem.ToString)
    '        Catch ex As Exception
    '            MsgBox("Trouble converting budget to integer value / variable.")
    '        End Try
    '    End If


    '    Dim iTabRows As Integer
    '    iTabRows = pTable.RowCount(Nothing)
    '    Dim pCursor As ICursor
    '    pCursor = pTable.Search(Nothing, False)

    '    Dim pRow As IRow

    '    ' ============================= Section 3 ========================
    '    ' 2. read all rows that have a decision option number of 2 into a list object
    '    ' read tables 1 and 2 into feature cursors
    '    ' read all table 2 into a list object

    '    ' Opentable and read the lines into an object
    '    ' Chose: read table 1 line by line and compare to table 2.  Read all of table 2 into a list object
    '    ' at the beginning of the sub.  
    '    ' custom object used: GLPKDecisionsObject

    '    Dim pGLPKDecisionsObjectTAB As New GLPKDecisionsObject(Nothing, Nothing, Nothing, Nothing)
    '    Dim lGLPKDecisionsObjectTAB As New List(Of GLPKDecisionsObject)

    '    Dim iBarrierEID As Integer
    '    Dim iOption As Integer
    '    Dim iBudgetTable As Integer


    '    Dim pField As IField
    '    pRow = pCursor.NextRow

    '    ' Loop through each row
    '    Do Until pRow Is Nothing

    '        pField = pRow.Fields.Field(iBarrierEIDField)
    '        If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
    '            Or pField.Type = esriFieldType.esriFieldTypeOID Then
    '            iBarrierEID = pRow.Value(iBarrierEIDField)
    '        Else
    '            MsgBox("Error reading Table. Could not convert the barrier EID to type 'integer'. Now Exiting.")
    '            'BackgroundWorker1.CancelAsync()
    '            'exitsubroutine()
    '            Exit Sub
    '        End If

    '        pField = pRow.Fields.Field(iBudgetfield)
    '        If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
    '            Or pField.Type = esriFieldType.esriFieldTypeDouble Then
    '            iBudgetTable = Convert.ToInt32(pRow.Value(iBudgetfield))
    '        Else
    '            MsgBox("Error reading Table 1. Could not convert the budget to type 'integer'. Now Exiting.")
    '            'BackgroundWorker1.CancelAsync()
    '            'exitsubroutine()
    '            Exit Sub
    '        End If

    '        pField = pRow.Fields.Field(iOptionField)
    '        If pField.Type = esriFieldType.esriFieldTypeInteger Or pField.Type = esriFieldType.esriFieldTypeSingle Or pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
    '            iOption = pRow.Value(iOptionField)
    '        Else
    '            MsgBox("Error reading Table 1. Could not convert the optionnum to type 'integer'. Now Exiting.")
    '            'BackgroundWorker1.CancelAsync()
    '            'exitsubroutine()
    '            Exit Sub
    '        End If

    '        ' add to new object
    '        If iOption > 1 And iBudgetTable = iBudget Then
    '            pGLPKDecisionsObjectTAB = New GLPKDecisionsObject(Nothing, Nothing, iBarrierEID, iOption)
    '            lGLPKDecisionsObjectTAB.Add(pGLPKDecisionsObjectTAB)
    '            iTotalBudget = iTotalBudget + iBudget
    '        End If

    '        'pRow1 = Nothing
    '        pRow = pCursor.NextRow

    '    Loop

    '    If lGLPKDecisionsObjectTAB.Count = 0 Then
    '        MsgBox("No decisions found in table, only values of '1' or 'do nothing'. Now Exiting.")
    '        Exit Sub
    '    End If

    '    ' ====================== Section 4 ========================
    '    ' Get EID's for highlighting

    '    Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
    '    Dim pUNAExt As IUtilityNetworkAnalysisExt

    '    Try
    '        pUNAExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
    '    Catch ex As Exception
    '        MsgBox("Trouble getting Utility Network Analyst Extension")
    '        Exit Sub
    '    End Try

    '    Dim pNetworkAnalysisExt As INetworkAnalysisExt
    '    pNetworkAnalysisExt = CType(pUNAExt, INetworkAnalysisExt)

    '    If pNetworkAnalysisExt.NetworkCount = 0 Then
    '        MsgBox("No Networks Loaded. Now Exiting")
    '        Exit Sub
    '    End If

    '    ' iEnumNetEIDBuilderGEN
    '    '             As New EnumNetEIDArray
    '    ' iEnumNetEIDBuilder
    '    '             AS New EnumNetEIDArray
    '    ' supposedly these two are the same - but preferred to use GEN version
    '    ' http://edndoc.esri.com/arcobjects/9.2/ComponentHelp/esriGeoDatabase/IEnumNetEIDBuilder.htm
    '    ' iEnumNetEID
    '    '             CType from EnumNETEIDArray

    '    Dim pResultsEIDsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
    '    pResultsEIDsGEN.Network = pNetworkAnalysisExt.CurrentNetwork.Network
    '    pResultsEIDsGEN.ElementType = esriElementType.esriETJunction

    '    For i = 0 To lGLPKDecisionsObjectTAB.Count - 1
    '        pResultsEIDsGEN.Add(lGLPKDecisionsObjectTAB(i).BarrierEID)
    '    Next

    '    Dim pResultsJunctions As IEnumNetEID
    '    pResultsJunctions = CType(pResultsEIDsGEN, IEnumNetEID)

    '    ' do edges just because
    '    ' edges were not returned -- create an empty enumeration
    '    pResultsEIDsGEN = New EnumNetEIDArray
    '    pResultsEIDsGEN.Network = pNetworkAnalysisExt.CurrentNetwork.Network
    '    pResultsEIDsGEN.ElementType = esriElementType.esriETEdge
    '    Dim pResultEdges As IEnumNetEID
    '    pResultEdges = CType(pResultsEIDsGEN, IEnumNetEID)


    '    Try

    '        pNetworkAnalysisExtResults = CType(pUNAExt, INetworkAnalysisExtResults)
    '        pNetworkAnalysisExtResults.ClearResults()
    '        pNetworkAnalysisExtResults.ResultsAsSelection = True
    '        'pNetworkAnalysisExtResults.SetResults(pResultsJunctions, Nothing)
    '        pNetworkAnalysisExtResults.CreateSelection(pResultsJunctions, pResultEdges)
    '    Catch ex As Exception
    '        MsgBox("Problem encountered setting results display. Exiting.")
    '        Exit Sub
    '    End Try

    '    lblHighlighted.Text = "Highlighted: " + Convert.ToString(pResultsJunctions.Count)

    'End Sub

    Private Sub cmdBrowseResultsTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseResultsTab.Click

        ' Sub created Aug. 17, 2012 by Greig Oldford
        ' Section 1: User select of File GDB, Personal GDB, or Text file
        ' Section 2: Get workspace of table and set to module variable
        ' Section 3: Check that table is correct format

        ' =============== Section X: Disable all subsequent steps ===
        ' (but keep track of their current state in case of 'Exit Sub'

        Dim bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled As Boolean
        If cmdBrowseDecisionsTab.Enabled = True Then
            bStep2Enabled = True
        Else
            bStep2Enabled = False
        End If
        If cmdBrowseConnectivityTab.Enabled = True Then
            bStep3Enabled = True
        Else
            bStep3Enabled = False
        End If
        If grpLines.Enabled = True Then
            bStep4Enabled = True
        Else
            bStep4Enabled = False
        End If
        If lstBudgets.Enabled = True Then
            bStep5Enabled = True
        Else
            bStep5Enabled = False
        End If
        If grpStep6.Enabled = True Then
            bStep6Enabled = True
        Else
            bStep6Enabled = False
        End If

        ' =============== Section 1: Get Table =====================
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
        pEnumGx.Reset()

        Dim pGxObject As IGxObject = pEnumGx.Next
        If pGxObject Is Nothing Then
            Exit Sub
        End If

        Dim sCategory As String
        sCategory = pGxObject.Category.ToString
        If sCategory <> "File Geodatabase Table" And sCategory <> "Text File" And sCategory <> "Personal Geodatabase Table" Then
            MsgBox("Input table must be a File Geodatabase Table, Personal Geodatabase Table, CSV, or Text File. Input cannot be of type " + sCategory)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)

            Exit Sub
        End If

        ' =============== Section 2: Get Workspace =====================
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
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End If

        Dim pTable As ITable
        Dim sNameArray(1) As String
        Dim sUniqueCode As String
        Dim bGRB As Boolean = False ' tracks whether analysis type is from Gurobi or GLPK
        Dim bUndirected As Boolean = False
        ' =============== Section 3: Check format of table =====================
        Try
            If m_sTableName1.Contains("Results_UNDIR") Or m_sTableName1.Contains("Results_DIR") Then
                ' get the unique 5 digit code from the table name
                sNameArray = m_sTableName1.Split(New String() {"DIR_"}, StringSplitOptions.RemoveEmptyEntries)
                If sNameArray.Length = 1 Then
                    MsgBox("No Unique five-digit code found in table name - need this for analysis. Exiting.")
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    Exit Sub
                End If
                sUniqueCode = sNameArray(1).Substring(0, 5)
                MsgBox("The Unique five-digit code of this analysis: " + sUniqueCode)
            Else
                MsgBox("Table name does not contain the string 'Results_UNDIR' or 'Results_DIR'. Now exiting.")
            End If
            If m_sTableName1.Contains("GLPK") Then
                bGRB = False
            ElseIf m_sTableName1.Contains("GRB") Then
                bGRB = True
            Else
                MsgBox("Need to confirm whether analysis comes from GLPK or Gurobi; no 'GRB' or 'GLPK' found in table name. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Trouble reading table 1 name. Exiting." + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening table 1. Exiting." + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try
        Try
            ' first check that the table has rows
            Dim iRowCount As Integer
            iRowCount = pTable.RowCount(Nothing)
            If iRowCount < 1 Then
                MsgBox("Row count of input table is zero.  Now exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            Dim pFields As IFields
            pFields = pTable.Fields

            ' Confirm the right fields are present
            Dim iBudgetFieldIndex, iTreatmentFieldIndex, iHabitatZMaxFieldIndex, iCentralBarrierEIDFieldIndex, iSinkEIDFieldIndex As Integer

            iBudgetFieldIndex = pFields.FindField("Budget")
            If iBudgetFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'Budget' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iTreatmentFieldIndex = pFields.FindField("Treatment")
            If iTreatmentFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'Treatment' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iSinkEIDFieldIndex = pFields.FindField("SinkEID")
            If iSinkEIDFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'SinkEID' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iHabitatZMaxFieldIndex = pFields.FindField("HabitatZMax")
            If iHabitatZMaxFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'HabitatZMax' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            ' If those fields are present continue and check to see if the central barrier EID
            ' field is present.  If so, this is an 'undirected' analysis.  
            iCentralBarrierEIDFieldIndex = pFields.FindField("CentralBarrierEID")
            If iCentralBarrierEIDFieldIndex <> -1 Then
                bUndirected = True
            End If

            ' Table Checks Out

        Catch ex As Exception
            MsgBox("Error trying to find required fields in input table.")
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        ' =============== Section 4: Get Budget Range =====================
        Dim sFieldname As String
        Dim pCursor As ICursor
        Dim pDataStats As IDataStatistics
        Dim vVar As Object
        Dim lBudgets As List(Of Double) = New List(Of Double)
        Dim dBudget As Double

        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  GLPKSolved (boolean/binary)
        '  Perc_Gap (double)
        '  MAxSolTime(integer)
        '  TimeUsed(double)
        '  HabitatZmax(double)
        '  CentralBarrierEID (integer)


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

        ' keep track of all budget amounts, lowest and highest, and count
        Dim dLowestBudget, dHighestBudget As Double
        ' default set really high or low
        dLowestBudget = 1000000000
        dHighestBudget = 0
        Dim iBudgetCount As Integer

        While Not IsNothing(vVar)
            Try
                dBudget = Convert.ToDouble(vVar)
                If dLowestBudget > dBudget Then
                    dLowestBudget = dBudget
                End If
                If dHighestBudget < dBudget Then
                    dHighestBudget = dBudget
                End If
                iBudgetCount += 1
            Catch ex As Exception
                MsgBox("Trouble Converting budget to type 'double.' Now Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End Try
            lBudgets.Add(dBudget)
            pEnumerator.MoveNext()
            vVar = pEnumerator.Current
        End While

        ' return variables for next step
        'm_dHighestBudget = dHighestBudget
        'm_dLowestBudget = dLowestBudget
        'm_lBudgets = lBudgets
        m_iBudgetCount = iBudgetCount
        lstBudgets.Items.Clear()
        ' add budgets to list
        For i = 0 To lBudgets.Count - 1
            lstBudgets.Items.Add(lBudgets(i).ToString)
        Next

        ' pass unique budgets to global variable for use by 
        ' temporary code for exporting decision count to 
        ' table (under 'highlight decisions' button currently)
        ' Dec. 2012
        m_lBudgets = lBudgets

        m_sUniqueCode = sUniqueCode
        m_bUndirected = bUndirected
        m_bGRB = bGRB

        ' going to loop through and get list of sinks
        Dim lSinkEIDs As New List(Of Integer)
        Dim iSinkEID As Integer
        Dim bFoundEID As Boolean
        sFieldname = "SinkEID"
        pCursor = pTable.Search(Nothing, False)

        ' Setup the datastatistics and get the unique values of the "Id" field
        pDataStats = New DataStatistics

        With pDataStats
            .Cursor = pCursor
            .Field = sFieldname
            pEnumerator = .UniqueValues
        End With

        ' get list of all unique sinks
        ' most of the time this should only be one sink
        pEnumerator.Reset()
        pEnumerator.MoveNext()
        vVar = pEnumerator.Current
        While Not IsNothing(vVar)
            Try
                iSinkEID = Convert.ToInt32(vVar)

            Catch ex As Exception
                MsgBox("Trouble Converting Sink EID to type 'integer' Now Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End Try

            bFoundEID = False
            For i = 0 To lSinkEIDs.Count - 1
                If lSinkEIDs(i) = iSinkEID Then
                    bFoundEID = True
                    Exit For
                End If
            Next
            If bFoundEID = False Then
                lSinkEIDs.Add(iSinkEID)
            End If
            pEnumerator.MoveNext()
            vVar = pEnumerator.Current
        End While

        m_lSinks = lSinkEIDs

        ' =============== Section 5: Enable Step 2 =====================
        lblStep1Status.Text = "Done!"

        txtResultsTab.Text = sFullTableName
        cmdBrowseDecisionsTab.Enabled = True

        ' disable the rest - in case this changes the other things' compatibility
        cmdBrowseConnectivityTab.Enabled = False
        grpLines.Enabled = False
        grpPolygons.Enabled = False
        lstBudgets.Enabled = False
        lstBudgets.Enabled = False
        grpStep6.Enabled = False

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

    Private Sub cmdHighlight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHighlight.Click
        '' Me.CheckForIllegalCrossThreadCalls = False
        ''progressThread = New Threading.Thread(AddressOf updateProgress)

        'Call Highlightstuff()


        ' THIS CODE WILL TEMPORARILY EXPORT TABLES CONTAINING 
        ' DECISION COUNTS (excluding 'do nothing') FOR ALL 
        ' BUDGETS IN THE LIST -- part of thesis 2012
        ' uses the DecisionCountObject created in 'shared objects' class

        ' ======================================
        ' 1. Get the parameters needed
        '    budget list = m_lBudgets
        '    Solver = m_bGRB
        '    Direction = m_bUndirected
        '    UniqueCode = m_sUniqueCode
        '    Treatment = 
        '    budget = 
        '    percgap = 
        '    decisioncount = 
        '    timeused = 

        ' LOOP THROUGH ALL RESULTS IN TABLE

        ' GET TREATMENT FROM 'RESULTS' TABLE
        Dim iTreatmentFieldIndex_Results, iTreatmentFieldIndex_Decisions As Integer
        Dim iBudgetFieldIndex_Results, iBudgetFieldIndex_Decisions As Integer
        Dim iPercentGapFieldIndex, iTimeUsedFieldIndex As Integer
        Dim iOptionNumIndex As Integer
        Dim pFields_Results, pFields_Decisions As IFields
        Dim pTable_Results, pTable_Decisions As ITable
        Dim iRowCount_Results, iRowCount_Decisions As Integer

        ' GET FIELDS FROM RESULTS TABLE
        Try
            ' Workspace set when opening table 1 (results table workspace)
            pTable_Results = m_pFWorkspace.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening decisions table 1. Exiting." + ex.Message)
            'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try
        Try
            ' first check that the table has rows
            iRowCount_Results = pTable_Results.RowCount(Nothing)
            If iRowCount_Results < 1 Then
                MsgBox("Row count of input results table is zero.  Now exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            pFields_Results = pTable_Results.Fields

            iBudgetFieldIndex_Results = pFields_Results.FindField("Budget")
            If iBudgetFieldIndex_Results = -1 Then
                MsgBox("Required field missing: the field 'Budget' could not be found in the input table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iPercentGapFieldIndex = pFields_Results.FindField("PercGap")
            If iPercentGapFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'PercGap' could not be found in the 'results' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iTimeUsedFieldIndex = pFields_Results.FindField("TimeUsed")
            If iTimeUsedFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'TimeUsed' could not be found in the 'results' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iTreatmentFieldIndex_Results = pFields_Results.FindField("Treatment")
            If iTreatmentFieldIndex_Results = -1 Then
                MsgBox("Required field missing: the field 'Treatment' could not be found in the 'results' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Trouble getting treatment or other fields from results table. " + ex.Message)
        End Try

        ' GET FIELDS FROM DECISIONS TABLE
        Try
            ' Workspace set when opening table 1 (results table workspace)
            pTable_Decisions = m_pFWorkspace.OpenTable(m_sTableName2)
        Catch ex As Exception
            MsgBox("Trouble opening decisions table 1. Exiting." + ex.Message)
            'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        Try
            ' first check that the table has rows
            iRowCount_Decisions = pTable_Decisions.RowCount(Nothing)
            If iRowCount_Decisions < 1 Then
                MsgBox("Row count of input decisions table is zero.  Now exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            pFields_Decisions = pTable_Decisions.Fields

            iBudgetFieldIndex_Decisions = pFields_Decisions.FindField("Budget")
            If iBudgetFieldIndex_Decisions = -1 Then
                MsgBox("Required field missing: the field 'Budget' could not be found in the 'decisions' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iTreatmentFieldIndex_Decisions = pFields_Decisions.FindField("Treatment")
            If iTreatmentFieldIndex_Decisions = -1 Then
                MsgBox("Required field missing: the field 'Treatment' could not be found in the 'decisions' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iOptionNumIndex = pFields_Decisions.FindField("OptionNum")
            If iOptionNumIndex = -1 Then
                MsgBox("Required field missing: the field 'OptionNum' could not be found in the 'decisions' table. Exiting.")
                'resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("Trouble getting treatment or other fields from results table. " + ex.Message)
        End Try

        ' For each Results Row
        ' Get Treatment, budget, percentgap, time used from table
        ' reset the decisionobject
        ' add the treatment, budget, percentgap, time used, GLPK/GRB, Direction, Analysis Code to decisionobject. 
        ' reset decision counter
        ' loop through decision table.  If treatment matches and budget matches and decision number is > 1 then increment counter.
        Dim pCursor_Results, pCursor_Decisions As ICursor
        Dim pRow_Results, pRow_Decisions As IRow

        pCursor_Results = pTable_Results.Search(Nothing, False)
        pRow_Results = pCursor_Results.NextRow

        Dim sTreatment_Results, sTreatment_Decisions As String
        Dim dBudget_Results, dBudget_Decisions As Double
        Dim dPercentGap, dTimeUsed As Double
        Dim pField_Results, pField_Decisions As IField
        Dim pDecisionCountObject As DecisionCountObject = New DecisionCountObject(Nothing, _
                                                                             Nothing, _
                                                                             Nothing, _
                                                                             Nothing, _
                                                                             Nothing, _
                                                                             Nothing, _
                                                                             Nothing, _
                                                                             Nothing)
        Dim lDecisionCountObject As List(Of DecisionCountObject) = New List(Of DecisionCountObject)

        Dim sSolver As String
        If m_bGRB = True Then
            sSolver = "GRB"
        Else
            sSolver = "GLPK"
        End If

        Dim sDirection As String
        If m_bUndirected = True Then
            sDirection = "UNDIR"
        Else
            sDirection = "DIR"
        End If

        Dim bUseDecision As Boolean = False
        Dim iDecisionCount As Integer = 0
        Dim iOptionNum As Integer = 0

        Do Until pRow_Results Is Nothing

            iDecisionCount = 0

            ' get results treatment string
            pField_Results = pRow_Results.Fields.Field(iTreatmentFieldIndex_Results)
            If pField_Results.Type = esriFieldType.esriFieldTypeString Then
                Try
                    sTreatment_Results = Convert.ToString(pRow_Results.Value(iTreatmentFieldIndex_Results))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table treatment to type 'string'. " + _
                           "Please check the 'results' table for treatment errors (might contain null values). " + _
                           ex.Message)
                    Exit Sub
                End Try
            Else
                MsgBox("The treatment field in the 'results' table cannot be of type " + pField_Results.Type.ToString)
                Exit Sub
            End If

            ' get results budget value
            pField_Results = pRow_Results.Fields.Field(iBudgetFieldIndex_Results)
            If pField_Results.Type = esriFieldType.esriFieldTypeInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSingle _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeDouble Then

                Try
                    dBudget_Results = Convert.ToDouble(pRow_Results.Value(iBudgetFieldIndex_Results))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table budget to type 'double'. " + _
                           "Please check the 'results' table for budget errors. " + _
                           ex.Message)
                    Exit Sub
                End Try
            Else
                MsgBox("The budget field in the 'results' table cannot be of type " + pField_Results.Type.ToString)
                Exit Sub
            End If

            ' get results percent gap from optimal value
            pField_Results = pRow_Results.Fields.Field(iPercentGapFieldIndex)
            If pField_Results.Type = esriFieldType.esriFieldTypeInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSingle _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeDouble Then

                Try
                    dPercentGap = Convert.ToDouble(pRow_Results.Value(iPercentGapFieldIndex))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table PercGap to type 'double'. " + _
                           "Please check the 'results' table for percent gap errors. " + _
                           ex.Message)
                    Exit Sub
                End Try
            Else
                MsgBox("The Percent Gap field in the 'results' table cannot be of type " + pField_Results.Type.ToString)
                Exit Sub
            End If

            ' get results time used value
            pField_Results = pRow_Results.Fields.Field(iTimeUsedFieldIndex)
            If pField_Results.Type = esriFieldType.esriFieldTypeInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSingle _
                Or pField_Results.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField_Results.Type = esriFieldType.esriFieldTypeDouble Then

                Try
                    dTimeUsed = Convert.ToDouble(pRow_Results.Value(iTimeUsedFieldIndex))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table TimeUsed to type 'double'. " + _
                           "Please check the 'results' table for time used errors. Error might be null values. " + _
                           ex.Message)
                    Exit Sub
                End Try
            Else
                MsgBox("The Time Used field in the 'results' table cannot be of type " + pField_Results.Type.ToString)
                Exit Sub
            End If

            pDecisionCountObject = New DecisionCountObject(sSolver, _
                                                           m_sUniqueCode, _
                                                           sDirection, _
                                                           sTreatment_Results, _
                                                           dBudget_Results, _
                                                           dPercentGap, _
                                                           Nothing, _
                                                           dTimeUsed)

            ' FOR EACH ROW IN DECISIONS TABLE
            ' STORE THE OPTION NUM IF
            ' it's > 1 (not 'do nothing)
            ' the budget matches
            ' the treatment matches
            pCursor_Decisions = pTable_Decisions.Search(Nothing, False)
            pRow_Decisions = pCursor_Decisions.NextRow
            Do Until pRow_Decisions Is Nothing

                bUseDecision = False

                ' get decisions table budget for this row
                pField_Decisions = pRow_Decisions.Fields.Field(iBudgetFieldIndex_Decisions)
                If pField_Decisions.Type = esriFieldType.esriFieldTypeInteger _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeSingle _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeSmallInteger _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeDouble Then

                    Try
                        dBudget_Decisions = Convert.ToDouble(pRow_Decisions.Value(iBudgetFieldIndex_Decisions))
                    Catch ex As Exception
                        MsgBox("Error. Trouble converting the table Budget to type 'double'. " + _
                               "Please check the 'decisions' table for Budget errors. Error might be null values. " + _
                               ex.Message)
                        Exit Sub
                    End Try
                Else
                    MsgBox("The budget field in the 'decisions' table cannot be of type " + pField_Decisions.Type.ToString)
                    Exit Sub
                End If

                ' get decisions Treatment for this row
                pField_Decisions = pRow_Decisions.Fields.Field(iTreatmentFieldIndex_Decisions)
                If pField_Decisions.Type = esriFieldType.esriFieldTypeString Then

                    Try
                        sTreatment_Decisions = Convert.ToString(pRow_Decisions.Value(iTreatmentFieldIndex_Decisions))
                    Catch ex As Exception
                        MsgBox("Error. Trouble converting the table treatment to type 'string'. " + _
                               "Please check the 'decisions' table for treatment errors. Error might be null values. " + _
                               ex.Message)
                        Exit Sub
                    End Try
                Else
                    MsgBox("The treatment field in the 'decisions' table cannot be of type " + pField_Decisions.Type.ToString)
                    Exit Sub
                End If

                ' get decisions OptionNum for this row
                pField_Decisions = pRow_Decisions.Fields.Field(iOptionNumIndex)
                If pField_Decisions.Type = esriFieldType.esriFieldTypeInteger _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeSingle _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeSmallInteger _
                    Or pField_Decisions.Type = esriFieldType.esriFieldTypeDouble Then

                    Try
                        iOptionNum = Convert.ToInt16(pRow_Decisions.Value(iOptionNumIndex))
                    Catch ex As Exception
                        MsgBox("Error. Trouble converting the table OptionNum to type 'Int16'. " + _
                               "Please check the 'decisions' table for OptionNum errors. Error might be null values. " + _
                               ex.Message)
                        Exit Sub
                    End Try
                Else
                    MsgBox("The OptionNum field in the 'decisions' table cannot be of type " + pField_Decisions.Type.ToString)
                    Exit Sub
                End If


                If sTreatment_Decisions = sTreatment_Results And _
                    dBudget_Decisions = dBudget_Results And _
                    iOptionNum > 1 Then

                    iDecisionCount += 1

                End If

                pRow_Decisions = pCursor_Decisions.NextRow()
            Loop

            ' For this Budget (row in the 'results' table)
            ' Add the count of the 'do something' decisions to the object
            pDecisionCountObject.iDecisionCount = iDecisionCount
            lDecisionCountObject.Add(pDecisionCountObject)

            pRow_Results = pCursor_Results.NextRow()
        Loop

        ' WRITE RESULTS TO TABLE SOMEWHERE
        ' TEMPORARY
        ' Solver = string (GLPK or GRB)
        ' UniqueCode = string
        ' Direction = string
        ' Treatment = string
        ' budget = double
        ' percgap = double
        ' decisioncount = integer
        ' timeused = double

        ' determine output workspace (use pop-up)
        ' =============== Section 1: Get Table =====================
        Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterFileGeodatabases
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Output File Geodatabase"
        pGxDialog.AllowMultiSelect = False

        pGxDialog.DoModalOpen(0, pEnumGx)

        If pEnumGx Is Nothing Then Exit Sub
        pEnumGx.Reset()

        Dim pGxObject As IGxObject = pEnumGx.Next
        If pGxObject Is Nothing Then
            Exit Sub
        End If

        Dim sCategory As String
        sCategory = pGxObject.Category.ToString
        If sCategory <> "File Geodatabase" Then
            MsgBox("Input table must be a File Geodatabase. Input cannot be of type " + sCategory)
            Exit Sub
        End If

        ' =============== Section 2: Get Workspace =====================
        Dim sFullGDBName As String
        sFullGDBName = pGxObject.FullName

        'Dim sTableDirectory As String
        'Dim iLastBackSlash As Integer
        'iLastBackSlash = sFullTableName.LastIndexOf("\")
        'sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        'm_sTableName1 = sFullTableName.Substring(iLastBackSlash + 1)

        ' have the GDB string -- will send it to get workspace
        Dim outFWorkspace As IFeatureWorkspace
        outFWorkspace = GetWorkspace(sFullGDBName)

        If pGxObject Is Nothing Then
            MsgBox("You must select an input geodatabase")
            Exit Sub
        End If

        Dim sTableOut As String
        Dim myTableName As String

        myTableName = InputBox("Choose a Table Name", "Table Name Select", "")
        sTableOut = TableName(myTableName, _
                              outFWorkspace, _
                              "Dec23")
        PrepDecisionCountTable(sTableOut, outFWorkspace)

        Dim pTable As ITable
        Dim pRowBuffer As IRowBuffer
        Dim pCursor As ICursor

        pTable = outFWorkspace.OpenTable(sTableOut)

        ' Solver = string (GLPK or GRB)
        ' UniqueCode = string
        ' Direction = string
        ' Treatment = string
        ' budget = double
        ' percgap = double
        ' decisioncount = integer
        ' timeused = double

        Dim j As Integer = 0
        For j = 0 To lDecisionCountObject.Count - 1
            Try
                pRowBuffer = pTable.CreateRowBuffer
            Catch ex As Exception
                MsgBox("Problem creating output table row buffer. " + ex.Message)
            End Try

            Try
                pRowBuffer.Value(1) = lDecisionCountObject(j).sSolver
                pRowBuffer.Value(2) = lDecisionCountObject(j).sUniqueCode
                pRowBuffer.Value(3) = lDecisionCountObject(j).sDirection
                pRowBuffer.Value(4) = lDecisionCountObject(j).sTreatment
                pRowBuffer.Value(5) = lDecisionCountObject(j).dBudget
                pRowBuffer.Value(6) = lDecisionCountObject(j).dPercGap
                pRowBuffer.Value(7) = lDecisionCountObject(j).iDecisionCount
                pRowBuffer.Value(8) = lDecisionCountObject(j).dTimeUsed
            Catch ex As Exception
                MsgBox("Problem with Row Buffer. " + ex.Message)

            End Try
            Try
                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
            Catch ex As Exception
                MsgBox("Problem inserting row into output table. " + _
                       ex.Message + _
                       " loop count: " + j.ToString)
            End Try

        Next

    End Sub
    Public Function GetCategory(ByVal pWorkspace As IWorkspace) As String
        Dim pClassID As New UID
        pClassID = pWorkspace.WorkspaceFactory.GetClassID
        Select Case pClassID.Value.ToString
            Case "{DD48C96A-D92A-11D1-AA81-00C04FA33A15}" ' pGDB
                'GetCategory = "Personal Geodatabase"
                Return "Personal Geodatabase"
            Case "{71FE75F0-EA0C-4406-873E-B7D53748AE7E}" ' fGDB
                'GetCategory = "File Geodatabase"
                Return "File Geodatabase"
            Case "{D9B4FA40-D6D9-11D1-AA81-00C04FA33A15}" ' GDB
                'GetCategory = "SDE Database"
                Return "SDE Database"
            Case "{A06ADB96-D95C-11D1-AA81-00C04FA33A15}" ' Shape
                'GetCategory = "Shapefile Workspace"
                Return "Shapefile Workspace"
            Case "{1D887452-D9F2-11D1-AA81-00C04FA33A15}" ' Coverage
                'GetCategory = "ArcInfo Coverage Workspace"
                Return "ArcInfo Coverage Workspace"
            Case Else
                'GetCategory = "Unknown Workspace Category"
                Return "Unknown Workspace Category"
        End Select
    End Function

    Private Function TableName(ByVal sKeyword As String, ByVal pFWorkspace As IFeatureWorkspace, ByRef sPrefix As String) As String
        ' Function:    Table Name
        ' Created By:  Greig Oldford
        ' Update Date: October 5, 2010
        '              December 16, 2012 (fix infinite loop on table prefix problem)
        ' Purpose:     1) To check for table name conflicts in output
        '                 workspace
        ' Keyword: "Habitat","Metrics","DCI","Connectivity"
        '

        Dim pWorkspace As IWorkspace
        Dim sFilepath As String
        Dim pWorkspaceFactory As IWorkspaceFactory
        Dim pWorkspace2 As IWorkspace2
        Dim message As String
        Dim title As String
        Dim defaultValue As String
        Dim myTableName As Object
        Dim reg As New Regex("^[A-Za-z0-9_]+$") ' checks if prefix contains allowed characters
        Dim reg2 As New Regex("^[0-9]")         ' for checking if string starts with numbers
        Dim sName As String

        pWorkspace = CType(pFWorkspace, IWorkspace)
        sFilepath = pWorkspace.PathName

        ' Need to connect to the personal or file geodatabase workspace
        ' to check to see if tablename already exists.
        Dim strCategory As String
        Dim pFGDBWFactory As New FileGDBWorkspaceFactory
        Dim pAccessWFactory As New AccessWorkspaceFactory

        ' Check what kind of workspace it is
        strCategory = GetCategory(pWorkspace)

        ' TODO: add code for if workspace is not one of these two types
        If strCategory = "File Geodatabase" Then
            pWorkspaceFactory = pFGDBWFactory
        ElseIf strCategory = "Personal Geodatabase" Then
            pWorkspaceFactory = pAccessWFactory
        End If

        ' Check if name prefix starts with numbers (not allowed)
        Dim bPrefixNumCheck As Boolean = False
        Dim bLengthCheck As Boolean = True
        Dim bPrefixCharCheck As Boolean = True
        Dim bValidate As Boolean = False

        Dim sNumMsg As String
        Dim sLengthMsg As String
        Dim sCharMsg As String

        ' Check for invalid characters
        ' Check for table names beginning with numbers
        ' Check that table name does not exist in workspace
        ' THIS SECTION HAS REPEATED CODE - INEFFICIENT AND COULD BE IMPROVED
        Do Until bValidate = True

            'Check that user hasn't entered invalid text for tablename
            bPrefixCharCheck = reg.IsMatch(sPrefix) 'For VB.NET
            ' Check if prefix starts with numbers (not allowed in ArcGIS for table names)
            bPrefixNumCheck = reg2.IsMatch(sPrefix)

            ' Check if table name is too long (>55)
            sName = sPrefix + "_" + sKeyword

            ' Check that length of table name isn't too long
            If Len(sName) > 55 Then
                bLengthCheck = False
            Else
                bLengthCheck = True
            End If

            If bPrefixCharCheck = False Or bPrefixNumCheck = True Or bLengthCheck = False Then
                title = "Problem with table name"

                If bPrefixCharCheck = False Or bPrefixNumCheck = True Then
                    If bPrefixCharCheck = False Then
                        sCharMsg = "Invalid character in table name. Table name may only include spaces, numbers, or _"
                    Else
                        sCharMsg = "OK"
                    End If
                    If bPrefixNumCheck = True Then
                        sNumMsg = "Table names cannot begin with numbers in ArcGIS."
                    Else
                        sNumMsg = "OK"
                    End If

                    message = "Invalid table name, please enter a new prefix." & Environment.NewLine & _
                    "Characters: " & sCharMsg & Environment.NewLine _
                    & "First Character: " & sNumMsg & Environment.NewLine

                    defaultValue = ""
                    myTableName = InputBox(message, title, defaultValue)

                    sPrefix = myTableName.ToString
                    If sPrefix = "" Then
                        Return "Cancel"
                    End If
                    bValidate = False

                End If

                If bLengthCheck = False Then
                    sLengthMsg = "Table name is too long.  Name cannot exceed 55 characters in ArcGIS"
                    message = "Invalid table name, please enter a new prefix." & Environment.NewLine & _
                        "Length: " & sLengthMsg
                    defaultValue = ""
                    myTableName = InputBox(message, title, defaultValue)
                    sPrefix = myTableName.ToString
                    If sPrefix = "" Then
                        Return "Cancel"
                    End If
                    bValidate = False
                End If

            Else
                bValidate = True
            End If

        Loop

        Dim bExists As Boolean = True
        Dim iUnderIndex As Integer
        Dim sTrimmedStrRight As String
        Dim sTrimmedStrLeft As String
        Dim iTrimmedSuffix As Integer

        Do Until bExists = False
            ' Open the database and check if name exists
            pWorkspace2 = CType(pWorkspaceFactory.OpenFromFile(sFilepath, 0), IWorkspace2)
            bExists = pWorkspace2.NameExists(esriDatasetType.esriDTTable, sName)

            ' If the table exists then rename it by either auto-incrementing the number
            ' on the end or tacking a _1 on the end.  
            If bExists = True Then
                iUnderIndex = sName.LastIndexOf("_")
                sTrimmedStrRight = sName.Substring(iUnderIndex + 1)
                If IsNumeric(sTrimmedStrRight) Then
                    iTrimmedSuffix = Convert.ToInt16(sTrimmedStrRight)
                    iTrimmedSuffix += 1
                    sTrimmedStrLeft = sName.Remove(iUnderIndex + 1, sTrimmedStrRight.Length)
                    sName = sTrimmedStrLeft + CStr(iTrimmedSuffix)
                Else
                    sName = sName + "_1"
                End If
            End If
        Loop

        ' make the sPrefix global so it persists in next loop
        ' to prevent userInput box from popping up all the time
        'm_sPrefix = sPrefix

        Return sName ' VB.NET

        'TableName = sName
    End Function

    Private Sub PrepDecisionCountTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with GLPK Model
        ' Created by: Greig Oldford
        ' For: Thesis

        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the table in the output workspace. 
        '
        ' Table fields:
        ' Solver = string (GLPK or GRB)
        ' UniqueCode = string
        ' Direction = string
        ' Treatment = string
        ' budget = double
        ' percgap = double
        ' decisioncount = integer
        ' timeused = double

        Dim pStTab As IStandaloneTable
        Dim pStTabColl As IStandaloneTableCollection
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit
        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        ' Create new Fields object
        pFields = New Fields
        pFieldsEdit = CType(pFields, IFieldsEdit)

        Dim iFields As Integer ' to keep track of number of fields

        pFieldsEdit.FieldCount_2 = 9
        iFields = 9

        ' ==========================================
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "ObjectID" 'VB.NET
        pFieldEdit.Name_2 = "ObID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        'pFieldEdit.AliasName = "ObjectID"
        'pFieldEdit.Name = "ObID"
        'pFieldEdit.Type = esriFieldType.esriFieldTypeOID

        pFieldsEdit.Field_2(0) = pField 'VB.NET

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Solver Field"
        pFieldEdit.Name_2 = "SOLVER"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 15
        pFieldsEdit.Field_2(1) = pField


        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Unique Code"
        pFieldEdit.Name_2 = "UNIQUE_CODE"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 15
        pFieldsEdit.Field_2(2) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Direction Type"
        pFieldEdit.Name_2 = "DIR_UNDIR"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 15
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "treatment"
        pFieldEdit.Name_2 = "TREATMENT"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 70
        pFieldsEdit.Field_2(4) = pField

        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(5) = pField

        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Percent Gap"
        pFieldEdit.Name_2 = "PERC_GAP"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(6) = pField

        ' ============ Eighth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Decision Count"
        pFieldEdit.Name_2 = "DEC_COUNT"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(7) = pField

        ' ============ Ninth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Time Used"
        pFieldEdit.Name_2 = "TIME_USED"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(8) = pField



        ' May be possible to add optional params for RDBMS behaviour
        pFWSpace.CreateTable(sTable, pFields, Nothing, Nothing, "")
        ' MsgBox "created table successfully"

        ' Add Table to Map Doc
        Dim pTable As ITable = pFWSpace.OpenTable(sTable)
        pStTabColl = CType(pMap, IStandaloneTableCollection)
        pStTab = New StandaloneTable
        pStTab.Table = pTable
        pStTabColl.AddStandaloneTable(pStTab)
        pMxDoc.UpdateContents()

    End Sub


    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
   
    Private Sub cmdBrowseDecisionsTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseDecisionsTab.Click

        ' Browse for Decisions Table
        ' Sub created Aug. 17, 2012 by Greig Oldford
        ' Section 1: User select of File GDB, Personal GDB, or Text file
        ' Section 2: Get workspace of table and set to module variable
        ' Section 3: Check that table is correct format
        ' Section 4: Check that table matches the 'Results table' - must use pairs

        ' =============== Section X: Disable all subsequent steps ===
        ' (but keep track of their current state in case of 'Exit Sub'

        Dim bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled As Boolean
        If cmdBrowseDecisionsTab.Enabled = True Then
            bStep2Enabled = True
        Else
            bStep2Enabled = False
        End If
        If cmdBrowseConnectivityTab.Enabled = True Then
            bStep3Enabled = True
        Else
            bStep3Enabled = False
        End If
        If grpLines.Enabled = True Then
            bStep4Enabled = True
        Else
            bStep4Enabled = False
        End If
        If lstBudgets.Enabled = True Then
            bStep5Enabled = True
        Else
            bStep5Enabled = False
        End If
        If grpStep6.Enabled = True Then
            bStep6Enabled = True
        Else
            bStep6Enabled = False
        End If

        ' =============== Section 1 =====================
        ' 1. get table
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
     
        If pGxObject Is Nothing Then
            Exit Sub
        End If

        Dim sCategory As String
        sCategory = pGxObject.Category.ToString
        If sCategory <> "File Geodatabase Table" And sCategory <> "Text File" And sCategory <> "Personal Geodatabase Table" Then
            MsgBox("Input table must be a File Geodatabase Table, Personal Geodatabase Table, CSV, or Text File. Input cannot be of type " + sCategory)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End If

        ' =============== Section 2 =====================
        ' 2. get table workspace
        Dim sFullTableName As String
        sFullTableName = pGxObject.FullName
        Dim sTableDirectory As String
        Dim iLastBackSlash As Integer
        iLastBackSlash = sFullTableName.LastIndexOf("\")
        sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        m_sTableName2 = sFullTableName.Substring(iLastBackSlash + 1)
        m_pFWorkspace2 = GetWorkspace(sTableDirectory)

        If pGxObject Is Nothing Then
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            MsgBox("You must select an input table")
            Exit Sub
        End If

        ' =============== Section 3: Check format of table =====================
        Dim sNameArray As String()
        Dim sUniqueCode As String
        Dim bGRB As Boolean = False

        Try
            If m_sTableName2.Contains("Decisions_UNDIR") Or m_sTableName2.Contains("Decisions_DIR") Then
                ' get the unique 5 digit code from the table name
                sNameArray = m_sTableName2.Split(New String() {"DIR_"}, StringSplitOptions.RemoveEmptyEntries)

                If sNameArray.Length = 1 Then
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    MsgBox("No Unique five-digit code found in table name - need this for analysis. Exiting.")
                    Exit Sub
                End If
                sUniqueCode = sNameArray(1).Substring(0, 5)

                If sUniqueCode <> m_sUniqueCode Then
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    MsgBox("The Unique code (" + sUniqueCode + ") from this table does not match the code in the results table (" + _
                           m_sUniqueCode + ").  These codes must match. Exiting.")
                    Exit Sub
                End If

            Else
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Table name does not contain the string 'Decisions_UNDIR' or 'Decisions_DIR'. Now exiting.")
            End If
            If m_sTableName2.Contains("GLPK") Then
                bGRB = False
            ElseIf m_sTableName2.Contains("GRB") Then
                bGRB = True
            Else
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Need to confirm whether analysis comes from GLPK or Gurobi; no 'GRB' or 'GLPK' found in table name. Exiting.")
                Exit Sub
            End If
            If bGRB <> m_bGRB Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("The analysis type doesn't match the 'Results' table;  must match GRB with GRB and GLPK with GLPK types.  Exiting")
                Exit Sub
            End If
        Catch ex As Exception
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            MsgBox("Trouble reading table 2 name. Exiting." + ex.Message)
            Exit Sub
        End Try

        ' Verify fields in table
        Dim pTable As ITable
        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName2)
        Catch ex As Exception
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            MsgBox("Trouble opening table 2. Exiting." + ex.Message)
            Exit Sub
        End Try

        Try
            ' first check that the table has rows
            Dim iRowCount As Integer
            iRowCount = pTable.RowCount(Nothing)
            If iRowCount < 1 Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Row count of input table is zero.  Now exiting.")
                'Exit Sub
            End If

            Dim pFields As IFields
            pFields = pTable.Fields

            '  SinkEID (integer)
            '  Treatment (string)
            '  Budget (double)
            '  BarrierEID (Integer)
            '  OptionNum (Integer)

            ' Confirm the right fields are present
            Dim iBudgetFieldIndex, iTreatmentFieldIndex, iOptionNumIndex, iBarrierEIDIndex As Integer

            iBudgetFieldIndex = pFields.FindField("Budget")
            If iBudgetFieldIndex = -1 Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Required field missing: the field 'Budget' could not be found in the input table. Exiting.")
                Exit Sub
            End If

            iTreatmentFieldIndex = pFields.FindField("Treatment")
            If iTreatmentFieldIndex = -1 Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Required field missing: the field 'Treatment' could not be found in the input table. Exiting.")
                Exit Sub
            End If

            iOptionNumIndex = pFields.FindField("OptionNum")
            If iOptionNumIndex = -1 Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Required field missing: the field 'OptionNum' could not be found in the input table. Exiting.")
                Exit Sub
            End If

            iBarrierEIDIndex = pFields.FindField("BarrierEID")
            If iBarrierEIDIndex = -1 Then
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Required field missing: the field 'BarrierEID' could not be found in the input table. Exiting.")
                Exit Sub
            End If

            ' Table Checks Out

        Catch ex As Exception
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            MsgBox("Error trying to check for required fields in input table. " + ex.Message)
            Exit Sub
        End Try

        ' =============== Section 4: Get Unique EID's =====================
        Dim lEIDs As List(Of Integer) = New List(Of Integer)
        Dim pCursor As ICursor
        Dim pDataStats As IDataStatistics
        Dim vVar As Object
        Dim iEID As Integer

        pCursor = pTable.Search(Nothing, False)

        ' Setup the datastatistics and get the unique values of the "Id" field
        pDataStats = New DataStatistics
        Dim pEnumerator As IEnumerator

        With pDataStats
            .Cursor = pCursor
            .Field = "BarrierEID"
            pEnumerator = .UniqueValues
        End With

        pEnumerator.Reset()
        pEnumerator.MoveNext()
        Try
            vVar = pEnumerator.Current
        Catch ex As Exception
            MsgBox("Format of input 'decisions' table matches, but no records were found. " + ex.Message)
        End Try

        While Not IsNothing(vVar)
            Try
                iEID = Convert.ToInt32(vVar)
            Catch ex As Exception
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                MsgBox("Trouble Converting BarrierEID to type 'integer.' Now Exiting.")
                Exit Sub
            End Try
            lEIDs.Add(iEID)
            pEnumerator.MoveNext()
            vVar = pEnumerator.Current
        End While

        m_lEIDs = lEIDs

        ' =============== Section 5: Enable Step 3 =====================
        lblStep2Status.Text = "Done!"
        cmdBrowseConnectivityTab.Enabled = True

        txtDecisionsTab.Text = sFullTableName

        ' disable the rest - in case this changes the other things' compatibility
        grpLines.Enabled = False
        grpPolygons.Enabled = False
        lstBudgets.Enabled = False
        lstBudgets.Enabled = False
        grpStep6.Enabled = False

    End Sub
#Region "DragDrop"
    ' THIS SECTION ISN't CURRENTLY BEING USED
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
                txtResultsTab.Text = pLayer.Name
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
#End Region

    Private Sub cmdBrowseConnectivityTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseConnectivityTab.Click

        ' check that it’s the correct table 
        ' by (a) all unique decision EID’s from ‘decisions’ table should be present
        ' Return all unique EIDs in a list
        ' Enable the listbox for 4


        ' =============== Section X: Disable all subsequent steps ===
        ' (but keep track of their current state in case of 'Exit Sub'

        Dim bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled As Boolean
        If cmdBrowseDecisionsTab.Enabled = True Then
            bStep2Enabled = True
        Else
            bStep2Enabled = False
        End If
        If cmdBrowseConnectivityTab.Enabled = True Then
            bStep3Enabled = True
        Else
            bStep3Enabled = False
        End If
        If grpLines.Enabled = True Then
            bStep4Enabled = True
        Else
            bStep4Enabled = False
        End If
        If lstBudgets.Enabled = True Then
            bStep5Enabled = True
        Else
            bStep5Enabled = False
        End If
        If grpStep6.Enabled = True Then
            bStep6Enabled = True
        Else
            bStep6Enabled = False
        End If

        ' =============== Section 1 =====================
        ' 1. get table
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
        If pGxObject Is Nothing Then
            Exit Sub
        End If

        Dim sCategory As String
        sCategory = pGxObject.Category.ToString
        If sCategory <> "File Geodatabase Table" And sCategory <> "Text File" And sCategory <> "Personal Geodatabase Table" Then
            MsgBox("Input table must be a File Geodatabase Table, Personal Geodatabase Table, CSV, or Text File. Input cannot be of type " + sCategory)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            txtConnectivityTab.Text = ""
            Exit Sub
        End If

        ' =============== Section 2 =====================
        ' 2. get table workspace
        Dim sFullTableName As String
        sFullTableName = pGxObject.FullName
        Dim sTableDirectory As String
        Dim iLastBackSlash As Integer
        iLastBackSlash = sFullTableName.LastIndexOf("\")
        sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        m_sTableName3 = sFullTableName.Substring(iLastBackSlash + 1)
        m_pFWorkspace3 = GetWorkspace(sTableDirectory)

        If pGxObject Is Nothing Then
            MsgBox("You must select an input table")
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        Else
            txtConnectivityTab.Text = sFullTableName
        End If

        ' =============== Section 3: Check format of table =====================
        Dim sNameArray As String()
        Dim sUniqueCode As String
        Dim bGRB As Boolean = False

        Try
            If m_sTableName3.Contains("connectivity_") Then
                ' get the unique 5 digit code from the table name
                sNameArray = m_sTableName3.Split(New String() {"connectivity_"}, StringSplitOptions.RemoveEmptyEntries)

                If sNameArray.Length = 1 Then
                    MsgBox("No Unique five-digit code found in table name - need this for analysis. Exiting.")
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    Exit Sub
                End If
                sUniqueCode = sNameArray(1).Substring(0, 5)

                If sUniqueCode <> m_sUniqueCode Then
                    MsgBox("The Unique code (" + sUniqueCode + ") from this table does not match the code in the results table (" + _
                           m_sUniqueCode + ").  These codes must match. Exiting.")
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    Exit Sub
                End If
            Else
                MsgBox("Table name does not contain the string 'connectivity_'. Now exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If
            
        Catch ex As Exception
            MsgBox("Trouble reading table 3 name. Exiting." + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        ' Verify fields in table
        Dim pTable As ITable
        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName3)
        Catch ex As Exception
            MsgBox("Trouble opening table 3. Exiting." + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        Try
            ' first check that the table has rows
            Dim iRowCount As Integer
            iRowCount = pTable.RowCount(Nothing)
            If iRowCount < 1 Then
                MsgBox("Row count of input table is zero.  Now exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            Dim pFields As IFields
            pFields = pTable.Fields

            '  BEID (integer)
            '  UpEID (integer)

            ' Confirm the right fields are present
            Dim iBEIDFieldIndex, iUpEIDFieldIndex As Integer

            iBEIDFieldIndex = pFields.FindField("BEID")
            If iBEIDFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'BEID' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            iUpEIDFieldIndex = pFields.FindField("UpEID")
            If iUpEIDFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'UpEID' could not be found in the input table. Exiting.")
                resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                Exit Sub
            End If

            ' Table Checks Out

        Catch ex As Exception
            MsgBox("Error trying to check for required fields in input table. " + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        ' =============== Section 4: Check EIDs =====================
        ' check all EIDs

        Dim lEIDs As List(Of Integer) = New List(Of Integer)
        Dim pCursor As ICursor
        Dim pDataStats As IDataStatistics
        Dim vVar As Object
        Dim iEID, i, j, k As Integer
        Dim bEIDFound As Boolean = False
        Dim bSinkEIDFound As Boolean = False

        Try
            ' double check all EIDs in the options / decisions table
            ' are present in the connectivity table. 
            pCursor = pTable.Search(Nothing, False)

            ' Setup the datastatistics and get the unique values of the "Id" field
            pDataStats = New DataStatistics
            Dim pEnumerator As IEnumerator

            With pDataStats
                .Cursor = pCursor
                .Field = "UpEID"
                pEnumerator = .UniqueValues
            End With

            pEnumerator.Reset()
            pEnumerator.MoveNext()
            vVar = pEnumerator.Current

            While Not IsNothing(vVar)
                Try
                    iEID = Convert.ToInt32(vVar)
                Catch ex As Exception
                    MsgBox("Trouble Converting BarrierEID to type 'integer.' Now Exiting.")
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    Exit Sub
                End Try
                lEIDs.Add(iEID)
                pEnumerator.MoveNext()
                vVar = pEnumerator.Current
            End While

            ' m_lEIDs were populated in the reading of the decisions table
            ' many EIDs will not be in m_lEID because sometimes 'do nothing' 
            ' is omitted from decisions table.  
            ' But check that all m_lEIDs from decisions table
            ' are present in the connectivity table (lEIDs)
            For i = 0 To m_lEIDs.Count - 1
                bEIDFound = False
                For j = 0 To lEIDs.Count - 1
                    If m_lEIDs(i) = lEIDs(j) Then
                        bEIDFound = True
                        Exit For
                    End If
                Next
                ' the the option of 'do nothing' is present in the decisions 
                ' table for the sink EID then it won't be in the lEIDs list because
                ' it is found in the BEID column, not the UpEID column
                If bEIDFound = False Then
                    ' there will be one EID not found in the 'UpEID' column - the sink
                    For k = 0 To m_lSinks.Count - 1
                        If m_lEIDs(i) = m_lSinks(k) Then
                            bEIDFound = True
                            lEIDs.Add(m_lEIDs(i))
                            Exit For
                        End If
                    Next
                End If
                If bEIDFound = False Then
                    MsgBox("The Barrier EID from the decisions table (" + m_lEIDs(i).ToString + ") cannot be found in the 'UpEID' field of the connectivity table.  Exiting.")
                    resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
                    Exit Sub
                End If
            Next

            ' if all EIDs are found then we'll switch up the lists because the
            ' connectivity list contains all the EIDs, whereas the Decisions table
            ' might not - 'Do nothing' options may be ommitted. 
            m_lEIDs = lEIDs

        Catch ex As Exception
            MsgBox("Error trying to get EIDs from Connectivity table. " + ex.Message)
            resetselectables(bStep2Enabled, bStep3Enabled, bStep4Enabled, bStep5Enabled, bStep6Enabled)
            Exit Sub
        End Try

        ' Table and EIDs check out. 

        ' =============== Section 5: Enable Step 3 =====================
        lblStep3Status.Text = "Done!"
        grpLines.Enabled = True
        grpPolygons.Enabled = True
        lstBudgets.Enabled = True


        ' disable the rest - in case this changes the other things' compatibility
        grpStep6.Enabled = False


    End Sub

    Private Sub lstLineLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstLineLayers.SelectedIndexChanged
        ' =================================================
        ' On selected Index Changed
        ' Clear the corresponding listbox containing fields of type: Double 
        ' Get the name of the selected lines layer
        ' find this layer in the map (use first encountered)
        ' get the fields for this layer
        ' populate the associate listbox with fields of type 'double'

        lstLineFieldsDouble.Items.Clear()
        Dim indexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstLineLayers.SelectedIndices
        Dim iSelectedIndex, iFieldCount As Integer
        Dim sLstLineName, sMapLineName, sFIPEXLineName As String
        Dim iMapFCID, i, j, k As Integer
        Dim pFields As IFields
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim pLayer As ILayer
        Dim pFeatureLayer As IFeatureLayer
        Dim bFoundLayerInMap As Boolean

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Me.Close()
        End Try


        ' check what item was selected
        ' sometimes things weren't selected, the box was just clicked. 
        If indexes.Count = 1 Then
            iSelectedIndex = lstLineLayers.SelectedIndex
            sLstLineName = lstLineLayers.SelectedItem.ToString
            For j = 0 To m_lLineLayersFields.Count - 1
                ' Find matching name and get fields for layer
                ' Loop through the TOC.  Could get lines through searching 
                ' the feature dataset of the active geometric network, 
                ' but want to leave things open in case we include lines that are not 
                ' part of the network (somewhat like polygons are currently included)
                If m_lLineLayersFields(j).Layer = sLstLineName Then
                    ' populate lstbox by getting fields 
                    ' for lines layer via the geometric network
                    ' Loop through all map layers
                    Try
                        bFoundLayerInMap = False
                        For i = 0 To pMap.LayerCount - 1
                            If pMap.Layer(i).Valid = True Then
                                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                                    pLayer = pMap.Layer(i)
                                    pFeatureLayer = CType(pLayer, IFeatureLayer)
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                                        Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
                                        Or pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge Then

                                        ' If the map layer name matches the FIPEX layer name
                                        If pFeatureLayer.Name = sLstLineName Then
                                            ' 
                                            pFields = pFeatureLayer.FeatureClass.Fields
                                            iFieldCount = pFields.FieldCount()

                                            For k = 0 To iFieldCount - 1
                                                If pFields.Field(k).Type() = esriFieldType.esriFieldTypeDouble Then
                                                    lstLineFieldsDouble.Items.Add(pFields.Field(k).Name)
                                                End If
                                            Next
                                            bFoundLayerInMap = True
                                            Exit For
                                        End If ' layer name match
                                    End If     ' simple edge
                                End If         ' LayerType Check
                            End If             ' layer is valid
                        Next                   ' Map Layer
                        If bFoundLayerInMap = False Then
                            MsgBox("Error. The layer selected in the listbox was not found in the active map.  Could not retrieve fields.")
                            Exit Sub
                        Else
                            ' check the list at this j for the field name
                            If m_lLineLayersFields(j).CumPermField = "NotSet" Then
                                lblLineFieldSet.Text = "None"
                            Else
                                lblLineFieldSet.Text = m_lLineLayersFields(j).CumPermField
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox("Error trying to retrieve the fields for the lines layer " + sLstLineName)
                        Exit Sub
                    End Try

                End If
            Next

        ElseIf indexes.Count > 1 Then
            MsgBox("Muliple layers selected in listbox.  This should be disabled... errrrror.")
            Exit Sub
        End If

    End Sub

    Private Sub lstPolyLayers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPolyLayers.SelectedIndexChanged
        ' =================================================
        ' On selected Index Changed
        ' Clear the corresponding listbox containing fields of type: Double 
        ' Get the name of the selected lines layer
        ' find this layer in the map (use first encountered)
        ' get the fields for this layer
        ' populate the associate listbox with fields of type 'double'

        lstPolyFieldsDouble.Items.Clear()
        Dim indexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstPolyLayers.SelectedIndices
        Dim iSelectedIndex, iFieldCount As Integer
        Dim sLstPolyName, sMapPolyName, sFIPEXPolyName As String
        Dim iMapFCID, i, j, k As Integer
        Dim pFields As IFields
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim pLayer As ILayer
        Dim pFeatureLayer As IFeatureLayer
        Dim bFoundLayerInMap As Boolean

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Me.Close()
        End Try

        ' check what item was selected
        ' sometimes things weren't selected, the box was just clicked. 
        If indexes.Count = 1 Then
            iSelectedIndex = lstPolyLayers.SelectedIndex
            sLstPolyName = lstPolyLayers.SelectedItem.ToString
            For j = 0 To m_lPolyLayersFields.Count - 1
                ' Find matching name and get fields for layer
                ' Loop through the TOC.  Could get lines through searching 
                ' the feature dataset of the active geometric network, 
                ' but want to leave things open in case we include lines that are not 
                ' part of the network (somewhat like polygons are currently included)
                If m_lPolyLayersFields(j).Layer = sLstPolyName Then
                    ' populate lstbox by getting fields 
                    ' for lines layer via the geometric network
                    ' Loop through all map layers
                    Try
                        bFoundLayerInMap = False
                        For i = 0 To pMap.LayerCount - 1
                            If pMap.Layer(i).Valid = True Then
                                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                                    pLayer = pMap.Layer(i)
                                    pFeatureLayer = CType(pLayer, IFeatureLayer)
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                                        ' If the map layer name matches the FIPEX layer name
                                        If pFeatureLayer.Name = sLstPolyName Then
                                            ' 
                                            pFields = pFeatureLayer.FeatureClass.Fields
                                            iFieldCount = pFields.FieldCount()

                                            For k = 0 To iFieldCount - 1
                                                If pFields.Field(k).Type() = esriFieldType.esriFieldTypeDouble Then
                                                    lstPolyFieldsDouble.Items.Add(pFields.Field(k).Name)
                                                End If
                                            Next
                                            bFoundLayerInMap = True
                                            Exit For
                                        End If ' layer name match
                                    End If     ' simple edge
                                End If         ' LayerType Check
                            End If             ' layer is valid
                        Next                   ' Map Layer
                        If bFoundLayerInMap = False Then
                            MsgBox("Error. The layer selected in the listbox was not found in the active map.  Could not retrieve fields.")
                            Exit Sub
                        Else
                            ' check the list at this j for the field name
                            If m_lPolyLayersFields(j).CumPermField = "NotSet" Then
                                lblPolyFieldSet.Text = "None"
                            Else
                                lblPolyFieldSet.Text = m_lPolyLayersFields(j).CumPermField
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox("Error trying to retrieve the fields for the lines layer " + sLstPolyName)
                        Exit Sub
                    End Try

                End If
            Next

        ElseIf indexes.Count > 1 Then
            MsgBox("Muliple layers selected in listbox.  This should be disabled... errrrror.")
            Exit Sub
        End If
    End Sub

    Private Sub cmdSetLineField_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetLineField.Click
        ' =================================================
        ' Get selected item from the fields listbox        
        ' Find the selected item in the layers listbox     
        ' crossref with master list of layers/fields       
        ' Set the 'field' in the master list.              
        ' =================================================

        ' ================= Section 1. Get Selected Item ================================
        Dim fieldsindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstLineFieldsDouble.SelectedIndices
        Dim iSelectedFieldIndex As Integer
        Dim sFieldName As String

        If fieldsindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the listbox.  Exiting.")
            Exit Sub
        ElseIf fieldsindexes.Count = 0 Then
            MsgBox("No fields in listbox are selected.  Please select a field.")
            Exit Sub
        ElseIf fieldsindexes.Count > 1 Then
            MsgBox("Multiple fields in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedFieldIndex = lstLineFieldsDouble.SelectedIndex
        End If

        Try
            sFieldName = lstLineFieldsDouble.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting fieldname from listbox.")
            Exit Sub
        End Try

        ' ================= Section 2. Get Selected Item from Layers Listbox ===============================
        Dim layersindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstLineLayers.SelectedIndices
        Dim iSelectedLayerIndex As Integer
        Dim sLayerName As String

        If layersindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the layers listbox.  Exiting.")
            Exit Sub
        ElseIf layersindexes.Count = 0 Then
            MsgBox("No layers in listbox are selected.  Please select a field.")
            Exit Sub
        ElseIf layersindexes.Count > 1 Then
            MsgBox("Multiple layers in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedLayerIndex = lstLineLayers.SelectedIndex
        End If

        Try
            sLayerName = lstLineLayers.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting layer name from listbox.")
        End Try

        ' ================= Section 3. Get corresponding master list layer ==========================
        Dim i, j As Integer
        Dim bFoundLayerMatch As Boolean = False

        For i = 0 To m_lLineLayersFields.Count - 1
            If m_lLineLayersFields(i).Layer = sLayerName Then
                bFoundLayerMatch = True
                Exit For
            End If
        Next
        If bFoundLayerMatch = False Then
            MsgBox("The layer for this field could not be found in the master list. Exiting.")
            Exit Sub
        End If

        m_lLineLayersFields(i).CumPermField = sFieldName
        MsgBox("Successfully set cumulative permeability field for layer " + sLayerName)
        lblLineFieldSet.Text = sFieldName

    End Sub

    Private Sub cmdSetPolyField_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetPolyField.Click
        ' =================================================
        ' Get selected item from the fields listbox        
        ' Find the selected item in the layers listbox     
        ' crossref with master list of layers/fields       
        ' Set the 'field' in the master list.              
        ' =================================================

        ' ================= Section 1. Get Selected Item ================================
        Dim fieldsindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstPolyFieldsDouble.SelectedIndices
        Dim iSelectedFieldIndex As Integer
        Dim sFieldName As String

        If fieldsindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the listbox.  Exiting.")
            Exit Sub
        ElseIf fieldsindexes.Count = 0 Then
            MsgBox("No fields in listbox are selected.  Please select a field.")
            Exit Sub
        ElseIf fieldsindexes.Count > 1 Then
            MsgBox("Multiple fields in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedFieldIndex = lstPolyFieldsDouble.SelectedIndex
        End If

        Try
            sFieldName = lstPolyFieldsDouble.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting fieldname from listbox.")
            Exit Sub
        End Try

        ' ================= Section 2. Get Selected Item from Layers Listbox ===============================
        Dim layersindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstPolyLayers.SelectedIndices
        Dim iSelectedLayerIndex As Integer
        Dim sLayerName As String

        If layersindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the listbox.  Exiting.")
            Exit Sub
        ElseIf layersindexes.Count = 0 Then
            MsgBox("No layers in listbox are selected.  Please select a field.")
            Exit Sub
        ElseIf layersindexes.Count > 1 Then
            MsgBox("Multiple layers in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedLayerIndex = lstPolyLayers.SelectedIndex
        End If

        Try
            sLayerName = lstPolyLayers.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting layer name from listbox.")
        End Try

        ' ================= Section 3. Get corresponding master list layer ==========================
        Dim i, j As Integer
        Dim bFoundLayerMatch As Boolean = False

        For i = 0 To m_lPolyLayersFields.Count - 1
            If m_lPolyLayersFields(i).Layer = sLayerName Then
                bFoundLayerMatch = True
                Exit For
            End If
        Next
        If bFoundLayerMatch = False Then
            MsgBox("The layer for this field could not be found in the master list. Exiting.")
            Exit Sub
        End If

        m_lPolyLayersFields(i).CumPermField = sFieldName
        MsgBox("Successfully set cumulative permeability field for layer " + sLayerName)
        lblPolyFieldSet.Text = sFieldName

    End Sub

    Private Sub cmdAddFieldLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddFieldLine.Click

        ' Get the field name from the textbox
        ' get the corresponding layer selected and layer in master list. 
        ' locate layer in Active TOC and get the fields
        ' confirm that the field doesn't already exist 
        ' attempt to add the field
        ' if successful, add the field to the listbox, change the setting in the master list, and update label


        ' ================= Section 1. Check that fieldname is okay ==========================
        Dim sNewFieldName As String
        Dim reg As New Regex("^[A-Za-z0-9_]+$") ' checks if prefix contains allowed characters
        Dim reg2 As New Regex("^[0-9]")         ' for checking if string starts with numbers
        Dim bCharCheck, bNumCheck As Boolean

        If txtNewLinefield.Text.Length = 0 Then
            MsgBox("Please enter a field name for the new field in the box above.")
            Exit Sub
        End If

        sNewFieldName = txtNewLinefield.Text

        'Check that user hasn't entered invalid text for tablename
        bCharCheck = False
        bCharCheck = reg.IsMatch(sNewFieldName) 'For VB.NET

        If bCharCheck = False Then
            MsgBox("Field contains an unacceptable character. Please use only characters a-z, A-Z, or 0-9 and rename the field.")
            Exit Sub
        End If

        ' Check if prefix starts with numbers 
        bNumCheck = False
        bNumCheck = reg2.IsMatch(sNewFieldName)

        If bNumCheck = True Then
            MsgBox("Field starts with a number. Please only use characters for first digit")
            Exit Sub
        End If

        ' Check if table name is too long (>55)
        ' Check that length of table name isn't too long
        If Len(sNewFieldName) > 55 Then
            MsgBox("The length of the field entered is too long.  Please keep the field below 55 characters")
            Exit Sub
        End If

        ' ================= Section 2. Get Selected Item from Layers Listbox ===============================
        Dim layersindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstLineLayers.SelectedIndices
        Dim iSelectedLayerIndex As Integer
        Dim sLayerName As String

        If layersindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the layers listbox. Please make sure a layer is selected. Exiting.")
            Exit Sub
        ElseIf layersindexes.Count = 0 Then
            MsgBox("No layers in listbox are selected. Please select a field.")
            Exit Sub
        ElseIf layersindexes.Count > 1 Then
            MsgBox("Multiple layers in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedLayerIndex = lstLineLayers.SelectedIndex
        End If

        Try
            sLayerName = lstLineLayers.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting layer name from listbox.")
        End Try

        ' ================= Section 3. Get corresponding master list layer ==========================
        Dim i, j As Integer
        Dim bFoundLayerMatch As Boolean = False

        For j = 0 To m_lLineLayersFields.Count - 1
            If m_lLineLayersFields(j).Layer = sLayerName Then
                bFoundLayerMatch = True
                Exit For
            End If
        Next
        If bFoundLayerMatch = False Then
            MsgBox("The layer selected in listbox could not be found in the master list. Exiting.")
            Exit Sub
        End If


        ' ================= Section 4. Get corresponding map layer and fields ==========================
        Dim bFoundLayerInMap As Boolean = False
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim pLayer As ILayer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFields As IFields
        Dim iFieldCount As Integer = 0

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Exit Sub
        End Try

        Try
            bFoundLayerInMap = False
            For i = 0 To pMap.LayerCount - 1
                If pMap.Layer(i).Valid = True Then
                    If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                        pLayer = pMap.Layer(i)
                        pFeatureLayer = CType(pLayer, IFeatureLayer)
                        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                            Or pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
                            Or pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge Then

                            ' If the map layer name matches the FIPEX layer name
                            If pFeatureLayer.Name = sLayerName Then
                                ' 
                                pFields = pFeatureLayer.FeatureClass.Fields
                                iFieldCount = pFields.FieldCount()

                                bFoundLayerInMap = True
                                Exit For
                            End If ' layer name match
                        End If     ' simple edge
                    End If         ' LayerType Check
                End If             ' layer is valid
            Next                   ' Map Layer
            If bFoundLayerInMap = False Then
                MsgBox("Error. The layer selected in the listbox was not found in the active map.  Could not retrieve fields.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error trying to retrieve the fields for the lines layer " + sLayerName)
            Exit Sub
        End Try

        ' ================= Section 5. Make sure layer name does not already exist ==========================
        ' Make sure field name does NOT already exist
        Dim bFieldAlready As Boolean = False
        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        If pFields Is Nothing Or iFieldCount = 0 Then
            MsgBox("Could not get fields for the selected layer. Exiting.")
            Exit Sub
        End If

        For i = 0 To iFieldCount - 1
            If pFields.Field(i).Name = sNewFieldName Then
                bFieldAlready = True
                Exit For
            End If
        Next

        If bFieldAlready = True Then
            MsgBox("Field Already exists in table. Please choose a different field name.")
            Exit Sub
        End If

        If pFeatureLayer Is Nothing Then
            MsgBox("Could not retrieve FeatureLayer object for the selected layer. Now exiting.")
            Exit Sub
        End If

        ' Add the field to hold the output to this layer
        Try
            pField = New Field
            pFieldEdit = CType(pField, IFieldEdit)
            pFieldEdit.AliasName_2 = sNewFieldName
            pFieldEdit.Name_2 = sNewFieldName
            pFieldEdit.Editable_2 = True
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pField = CType(pFieldEdit, IField)
            pFeatureLayer.FeatureClass.AddField(pField)

        Catch ex As Exception
            MsgBox("Failed adding new field to feature class. Please make sure layer is editable (close edit sessions, check user permissions, etc.). " & ex.Message)
            Exit Sub
        End Try

        m_lLineLayersFields(j).CumPermField = sNewFieldName
        MsgBox("Successfully set new cumulative permeability field for layer " + sLayerName)
        lblLineFieldSet.Text = sNewFieldName
        lstLineFieldsDouble.Items.Add(sNewFieldName)

    End Sub

    Private Sub cmdAddFieldPoly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddFieldPoly.Click
        ' ----------------------------------------------
        ' Get the field name from the textbox
        ' get the corresponding layer selected and layer in master list. 
        ' locate layer in Active TOC and get the fields
        ' confirm that the field doesn't already exist 
        ' attempt to add the field
        ' if successful, add the field to the listbox, change the setting in the master list, and update label

        ' ================= Section 1. Check that fieldname is okay ==========================
        Dim sNewFieldName As String
        Dim reg As New Regex("^[A-Za-z0-9_]+$") ' checks if prefix contains allowed characters
        Dim reg2 As New Regex("^[0-9]")         ' for checking if string starts with numbers
        Dim bCharCheck, bNumCheck As Boolean

        If txtNewPolyField.Text.Length = 0 Then
            MsgBox("Please enter a field name for the new field in the box above.")
            Exit Sub
        End If

        sNewFieldName = txtNewPolyField.Text

        'Check that user hasn't entered invalid text for tablename
        bCharCheck = False
        bCharCheck = reg.IsMatch(sNewFieldName) 'For VB.NET

        If bCharCheck = False Then
            MsgBox("Field contains an unacceptable character. Please use only characters a-z, A-Z, or 0-9 and rename the field.")
            Exit Sub
        End If

        ' Check if prefix starts with numbers 
        bNumCheck = False
        bNumCheck = reg2.IsMatch(sNewFieldName)

        If bNumCheck = True Then
            MsgBox("Field starts with a number. Please only use characters for first digit")
            Exit Sub
        End If

        ' Check if table name is too long (>55)
        ' Check that length of table name isn't too long
        If Len(sNewFieldName) > 55 Then
            MsgBox("The length of the field entered is too long.  Please keep the field below 55 characters")
            Exit Sub
        End If

        ' ================= Section 2. Get Selected Item from Layers Listbox ===============================
        Dim layersindexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstPolyLayers.SelectedIndices
        Dim iSelectedLayerIndex As Integer
        Dim sLayerName As String

        If layersindexes Is Nothing Then
            MsgBox("Trouble getting the selected index collection from the layers listbox. Please make sure a layer is selected. Exiting.")
            Exit Sub
        ElseIf layersindexes.Count = 0 Then
            MsgBox("No layers in listbox are selected. Please select a field.")
            Exit Sub
        ElseIf layersindexes.Count > 1 Then
            MsgBox("Multiple layers in listbox are selected - this should be disabled.... errrrrrror.")
            Exit Sub
        Else
            iSelectedLayerIndex = lstPolyLayers.SelectedIndex
        End If

        Try
            sLayerName = lstPolyLayers.SelectedItem.ToString
        Catch ex As Exception
            MsgBox("Trouble getting layer name from listbox.")
        End Try

        ' ================= Section 3. Get corresponding master list layer ==========================
        Dim i, j As Integer
        Dim bFoundLayerMatch As Boolean = False

        For j = 0 To m_lPolyLayersFields.Count - 1
            If m_lPolyLayersFields(j).Layer = sLayerName Then
                bFoundLayerMatch = True
                Exit For
            End If
        Next
        If bFoundLayerMatch = False Then
            MsgBox("The layer selected in listbox could not be found in the master list. Exiting.")
            Exit Sub
        End If

        ' ================= Section 4. Get corresponding map layer and fields ==========================
        Dim bFoundLayerInMap As Boolean = False
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim pLayer As ILayer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFields As IFields
        Dim iFieldCount As Integer = 0

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Exit Sub
        End Try

        Try
            bFoundLayerInMap = False
            For i = 0 To pMap.LayerCount - 1
                If pMap.Layer(i).Valid = True Then
                    If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                        pLayer = pMap.Layer(i)
                        pFeatureLayer = CType(pLayer, IFeatureLayer)
                        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                            ' If the map layer name matches the FIPEX layer name
                            If pFeatureLayer.Name = sLayerName Then
                                ' 
                                pFields = pFeatureLayer.FeatureClass.Fields
                                iFieldCount = pFields.FieldCount()

                                bFoundLayerInMap = True
                                Exit For
                            End If ' layer name match
                        End If     ' simple edge
                    End If         ' LayerType Check
                End If             ' layer is valid
            Next                   ' Map Layer
            If bFoundLayerInMap = False Then
                MsgBox("Error. The layer selected in the listbox was not found in the active map.  Could not retrieve fields.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error trying to retrieve the fields for the Polys layer " + sLayerName)
            Exit Sub
        End Try

        ' ================= Section 5. Make sure layer name does not already exist ==========================
        ' Make sure field name does NOT already exist
        Dim bFieldAlready As Boolean = False
        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        If pFields Is Nothing Or iFieldCount = 0 Then
            MsgBox("Could not get fields for the selected layer. Exiting.")
            Exit Sub
        End If

        For i = 0 To iFieldCount - 1
            If pFields.Field(i).Name = sNewFieldName Then
                bFieldAlready = True
                Exit For
            End If
        Next

        If bFieldAlready = True Then
            MsgBox("Field Already exists in table. Please choose a different field name.")
            Exit Sub
        End If

        If pFeatureLayer Is Nothing Then
            MsgBox("Could not retrieve FeatureLayer object for the selected layer. Now exiting.")
            Exit Sub
        End If

        ' Add the field to hold the output to this layer
        Try
            pField = New Field
            pFieldEdit = CType(pField, IFieldEdit)
            pFieldEdit.AliasName_2 = sNewFieldName
            pFieldEdit.Name_2 = sNewFieldName
            pFieldEdit.Editable_2 = True
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pField = CType(pFieldEdit, IField)
            pFeatureLayer.FeatureClass.AddField(pField)

        Catch ex As Exception
            MsgBox("Failed adding new field to feature class. Please make sure layer is editable (close edit sessions, check user permissions, etc.). " & ex.Message)
            Exit Sub
        End Try

        m_lPolyLayersFields(j).CumPermField = sNewFieldName
        MsgBox("Successfully set new cumulative permeability field for layer " + sLayerName)
        lblPolyFieldSet.Text = sNewFieldName
        lstPolyFieldsDouble.Items.Add(sNewFieldName)
        
    End Sub


    Private Sub lstBudgets_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstBudgets.SelectedIndexChanged
        Dim indexes As Windows.Forms.ListBox.SelectedIndexCollection = Me.lstBudgets.SelectedIndices

        If indexes Is Nothing Then
            MsgBox("Could not get reference to the budget list object. Exiting.")
            Exit Sub
        ElseIf indexes.Count = 0 Then
            MsgBox("Could not find selected budget in list.")
            Exit Sub
        ElseIf indexes.Count > 1 Then
            MsgBox("Muliple budgets selected - should only be one. errrrrrror.")
            Exit Sub
        ElseIf indexes.Count = -1 Then
            MsgBox("A count of negative one was found for the budgets. errrrrrror.")
            Exit Sub
        End If

        Dim sBudget As String
        sBudget = lstBudgets.SelectedItem.ToString
        Dim dBudget As Double

        If sBudget Is Nothing Then
            MsgBox("No budget retrieved from listbox.")
            Exit Sub
        ElseIf sBudget = "" Then
            MsgBox("No budget retrived from listbox.")
            Exit Sub
        End If

        Try
            dBudget = Convert.ToDouble(sBudget)
        Catch ex As Exception
            MsgBox("Failure converting value in budget listbox to type 'double'. Exiting. " + ex.Message)
            Exit Sub
        End Try

        m_dBudget = dBudget
        lblStep5Status.Text = "Done!"
        lblStep5Warning.Text = ""
        grpStep6.Enabled = True

    End Sub
    Private Sub resetselectables(ByVal bStep2Enabled As Boolean, _
                                 ByVal bStep3Enabled As Boolean, _
                                 ByVal bStep4Enabled As Boolean, _
                                 ByVal bStep5Enabled As Boolean, _
                                 ByVal bStep6Enabled As Boolean
                                  )
        If bStep2Enabled = True Then
            cmdBrowseDecisionsTab.Enabled = True
        Else
            cmdBrowseDecisionsTab.Enabled = False
        End If
        If bStep3Enabled = True Then
            cmdBrowseConnectivityTab.Enabled = True
        Else
            cmdBrowseConnectivityTab.Enabled = False
        End If
        If bStep4Enabled = True Then
            grpLines.Enabled = True
        Else
            grpLines.Enabled = False
        End If
        If bStep5Enabled = True Then
            lstBudgets.Enabled = True
        Else
            lstBudgets.Enabled = False
        End If
        If bStep6Enabled = True Then
            grpStep6.Enabled = True
        Else
            grpStep6.Enabled = False
        End If


    End Sub


    Private Sub cmdCalcCumPerm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCalcCumPerm.Click
        ' ===================================================================================
        ' 1. Check that there is a budget selected. 
        '
        ' 2. Check that there is at least one lines layer checked
        '   If not, and there are no polygons checked, then warn user and quit
        ' Check that all checked layers have a cumulative permeability field set
        ' Create a new list of all checked layers and their corresponding permeability field
        '
        ' 3. If this is an 'undirected' analysis get the central node from the results table
        '
        ' 4. Get the decision EIDs from the 'decisions table' (and option num if not '1' or 'do nothing)
        '   corresponding to this budget
        ' Check the feature class table fields for all EIDs and make sure the field corresponding to the 
        '  optionNum is present.  
        ' Check that EID for central node is a network junction, not an edge
        ' Check that all EIDs are network junctions, not edges.  
        ' 
        ' Begin analysis. 

        ' ================= Section 1: Check Budget =========================================
        If m_dBudget = 0 Then
            MsgBox("The budget selected is zero.  Please select a non-zero budget.")
            Exit Sub
        End If

        ' ================= Section 2: Get Lines and Poly Layers Checked ====================
        ' Check that there is at least one lines layer checked
        '   If not, and there are no polygons checked, then warn user and quit
        ' Check that all checked layers have a cumulative permeability field set
        ' Create a new list of all checked layers and their corresponding permeability field

        Dim i, m As Integer
        Dim checkedindexeslines As Windows.Forms.CheckedListBox.CheckedIndexCollection = Me.lstLineLayers.CheckedIndices
        Dim checkedindexespolys As Windows.Forms.CheckedListBox.CheckedIndexCollection = Me.lstPolyLayers.CheckedIndices

        If checkedindexeslines.Count = 0 Then
            If checkedindexespolys.Count = 0 Then
                MsgBox("There are no layers checked.  Please check at least one line or one polygon layer.")
                Exit Sub
            End If
        End If

        Dim iCheckedIndex As Integer
        Dim sListLayer As String
        Dim bFoundLayer As Boolean = False
        Dim lLineLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
        Dim lPolyLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField) = New List(Of LayersAndFCIDAndCumulativePassField)
        Dim pLayersFCIDAndField As LayersAndFCIDAndCumulativePassField = New LayersAndFCIDAndCumulativePassField(Nothing, Nothing, Nothing)

        For i = 0 To checkedindexeslines.Count - 1
            ' Get the layer name, from list of checked items
            Try
                iCheckedIndex = checkedindexeslines.Item(i)
                sListLayer = lstLineLayers.Items.Item(iCheckedIndex).ToString
            Catch ex As Exception
                MsgBox("Error trying to get checked item from the layer list box. Exiting")
                Exit Sub
            End Try

            ' Cross ref with the master list
            bFoundLayer = False
            For m = 0 To m_lLineLayersFields.Count - 1
                If m_lLineLayersFields.Item(m).Layer = sListLayer Then
                    bFoundLayer = True
                    Exit For
                End If
            Next

            If bFoundLayer = False Then
                MsgBox("Error 23.  Could not find checked item in layer list in the master list. Exiting.")
                Exit Sub
            End If

            If m_lLineLayersFields.Item(m).CumPermField = "NotSet" Then
                MsgBox("The cumulative permeability field for the layer " + m_lLineLayersFields.Item(m).Layer + _
                       " is not set. Please set this field to continue. Exiting.")
                Exit Sub
            End If
            pLayersFCIDAndField = New LayersAndFCIDAndCumulativePassField(m_lLineLayersFields.Item(m).Layer, _
                                                                          m_lLineLayersFields.Item(m).FCID, _
                                                                          m_lLineLayersFields.Item(m).CumPermField)
            lLineLayersFieldsCHECKED.Add(pLayersFCIDAndField)
        Next

        For i = 0 To checkedindexespolys.Count - 1
            ' Get the layer name, from list of checked items
            Try
                iCheckedIndex = checkedindexespolys.Item(i)
                sListLayer = lstPolyLayers.Items.Item(iCheckedIndex).ToString
            Catch ex As Exception
                MsgBox("Error trying to get checked item from the layer list box. Exiting")
                Exit Sub
            End Try

            ' Cross ref with the master list
            bFoundLayer = False
            For m = 0 To m_lPolyLayersFields.Count - 1
                If m_lPolyLayersFields.Item(m).Layer = sListLayer Then
                    bFoundLayer = True
                    Exit For
                End If
            Next

            If bFoundLayer = False Then
                MsgBox("Error 24.  Could not find checked item in layer list in the master list. Exiting.")
                Exit Sub
            End If

            If m_lPolyLayersFields.Item(m).CumPermField = "NotSet" Then
                MsgBox("The cumulative permeability field for the layer " + m_lPolyLayersFields.Item(m).Layer + _
                       " is not set. Please set this field to continue. Exiting.")
                Exit Sub
            End If
            pLayersFCIDAndField = New LayersAndFCIDAndCumulativePassField(m_lPolyLayersFields.Item(m).Layer, _
                                                                          m_lPolyLayersFields.Item(m).FCID, _
                                                                          m_lPolyLayersFields.Item(m).CumPermField)
            lPolyLayersFieldsCHECKED.Add(pLayersFCIDAndField)
        Next

        ' ================= Section 3: Get Central Node (if needed) =========================
        Dim pTable As ITable
        Dim pFields As IFields
        Dim iRowCount As Integer
        Dim pCursor As ICursor
        Dim pRow As IRow
        Dim pField As IField
        Dim iCentralBarrierEIDFieldIndex, iCentralBarrierEID, iBudgetFieldIndex, iSinkEID, iSinkEIDFieldIndex As Integer
        Dim dTableBudget As Double
        Dim bFoundBudgetAndCentralEID As Boolean = False
        Dim bFoundBudgetAndSinkEID As Boolean = False

        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening table 1. Exiting." + ex.Message)
            Exit Sub
        End Try
        Try
            ' first check that the table has rows
            iRowCount = pTable.RowCount(Nothing)
            If iRowCount < 1 Then
                MsgBox("Row count of input table 1 is zero.  Now exiting.")
                Exit Sub
            End If

            ' Confirm the right fields are present
            pFields = pTable.Fields
            If m_bUndirected = True Then
                iCentralBarrierEIDFieldIndex = pFields.FindField("CentralBarrierEID")
                If iCentralBarrierEIDFieldIndex = -1 Then
                    MsgBox("Could not find the 'CentralBarrierEID' field in table 1.  Column name must match exactly. Now exiting.")
                    Exit Sub
                End If
            End If
            

            iBudgetFieldIndex = pFields.FindField("Budget")
            If iBudgetFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'Budget' could not be found in the input table. Exiting.")
                Exit Sub
            End If

            ' Confirm the sinkEID is present
            ' need this no matter what type of analysis it is otherwise
            ' will try to query the sinks feature class for permeability value
            ' which are usually not set and always assumed 100% permeable
            iSinkEIDFieldIndex = pFields.FindField("SinkEID")
            If iSinkEIDFieldIndex = -1 Then
                MsgBox("Could not find the 'SinkEID' field in table 1.  Column name must match exactly. Now exiting.")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("Error trying to find required fields in input table.")
            Exit Sub
        End Try

        ' Get cursor and search for the central barrier EID for the given budget
        pCursor = pTable.Search(Nothing, False)
        pRow = pCursor.NextRow

            ' Loop through each row
        Do Until pRow Is Nothing
            ' get budget value
            pField = pRow.Fields.Field(iBudgetFieldIndex)
            If pField.Type = esriFieldType.esriFieldTypeInteger _
                Or pField.Type = esriFieldType.esriFieldTypeSingle _
                Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField.Type = esriFieldType.esriFieldTypeDouble Then

                Try
                    dTableBudget = Convert.ToDouble(pRow.Value(iBudgetFieldIndex))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table budget to type 'double'. Please check the 'results' table for budget errors.")
                    Exit Sub
                End Try
            Else
                MsgBox("The budget field in the 'Results' table cannot be of type " + pField.Type.ToString)
                Exit Sub
            End If

            ' if the budget corresponds then get the central barrier EID and sink
            If dTableBudget = m_dBudget Then

                ' Central node field EID (if undirected)
                If m_bUndirected = True Then
                    pField = pRow.Fields.Field(iCentralBarrierEIDFieldIndex)
                    If pField.Type = esriFieldType.esriFieldTypeInteger _
                        Or pField.Type = esriFieldType.esriFieldTypeSingle _
                        Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                        Or pField.Type = esriFieldType.esriFieldTypeOID Then
                        Try
                            iCentralBarrierEID = Convert.ToInt32(pRow.Value(iCentralBarrierEIDFieldIndex))
                        Catch ex As Exception
                            MsgBox("The central barrier EID could not be converted to type 'integer'. Please check the 'results' table.")
                            Exit Sub
                        End Try
                        bFoundBudgetAndCentralEID = True
                    End If
                End If

                ' SinkEID
                pField = pRow.Fields.Field(iSinkEIDFieldIndex)
                If pField.Type = esriFieldType.esriFieldTypeInteger _
                    Or pField.Type = esriFieldType.esriFieldTypeSingle _
                    Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                    Or pField.Type = esriFieldType.esriFieldTypeOID Then
                    Try
                        iSinkEID = Convert.ToInt32(pRow.Value(iSinkEIDFieldIndex))
                    Catch ex As Exception
                        MsgBox("The central sink EID could not be converted to type 'integer'. Please check the 'results' table.")
                        Exit Sub
                    End Try
                    bFoundBudgetAndSinkEID = True
                End If

            End If
            pRow = pCursor.NextRow
            If bFoundBudgetAndSinkEID = True And bFoundBudgetAndCentralEID = True Then
                Exit Do
            End If
        Loop

        If m_bUndirected = True And bFoundBudgetAndCentralEID = False Then
            MsgBox("Could not find central barrier EID corresponding to this budget. Please check 'results' table.")
            Exit Sub
        End If
        If bFoundBudgetAndSinkEID = False Then
            MsgBox("Could not find central sink EID corresponding to this budget. Please check 'results' table.")
            Exit Sub
        End If

        ' In the case where all decision options are listed, the decision for the sink EID would be an option (but always 1 or 'do nothing) ... therefore
        ' we know the sinkEID prior to selection of a budget, because this EID is the only one not present in the UpEID
        ' field of the connectivity table.  It was therefore also detected in the 'browse connectivity' stage and added
        ' the master list.  
        ' If the 'do nothing' options were not listed in the input table then we don't know the sink and it would not have been 
        ' added.  Need to check this now. 
        Dim bFoundEID As Boolean = False
        For i = 0 To m_lEIDs.Count - 1
            If m_lEIDs(i) = iSinkEID Then
                bFoundEID = True
                Exit For
            End If
        Next
        If bFoundEID = False Then
            m_lEIDs.Add(iSinkEID)
        End If

        ' ================= Section 4: Get Decision EIDs from Decisions Table and Permeabilities ===============
        ' for this budget get the list of decisions with EIDs
        Try
            pTable = m_pFWorkspace.OpenTable(m_sTableName2)
        Catch ex As Exception
            MsgBox("Trouble opening table 2. Exiting." + ex.Message)
            Exit Sub
        End Try

        Dim iOptionNumIndex, iBarrierEIDIndex, iTreatmentIndex As Integer
        Dim sTreatment As String
        Dim pEnumNetEIDBuilder As ESRI.ArcGIS.Geodatabase.IEnumNetEIDBuilder
        Dim pEnumNetEID As ESRI.ArcGIS.Geodatabase.IEnumNetEID
        Dim pEnumEidInfo As ESRI.ArcGIS.NetworkAnalysis.IEnumEIDInfo
        Dim pEidInfo As ESRI.ArcGIS.NetworkAnalysis.IEIDInfo
        Dim pFeature As ESRI.ArcGIS.Geodatabase.IFeature
        Dim pJuncFeat As ESRI.ArcGIS.Geodatabase.IJunctionFeature
        Dim pEidHelper As ESRI.ArcGIS.NetworkAnalysis.EIDHelper

        Try
            ' first check that the table has rows
            iRowCount = pTable.RowCount(Nothing)
            If iRowCount < 1 Then
                MsgBox("Row count of input table 2 is zero.  Now exiting.")
                'Exit Sub
            End If

            ' Confirm the right fields are present
            pFields = pTable.Fields

            iBudgetFieldIndex = pFields.FindField("Budget")
            If iBudgetFieldIndex = -1 Then
                MsgBox("Required field missing: the field 'Budget' could not be found in the input table #2. Exiting.")
                Exit Sub
            End If

            iOptionNumIndex = pFields.FindField("OptionNum")
            If iOptionNumIndex = -1 Then
                MsgBox("Required field missing: the field 'OptionNum' could not be found in the input table #2. Exiting.")
                Exit Sub
            End If

            iBarrierEIDIndex = pFields.FindField("BarrierEID")
            If iBarrierEIDIndex = -1 Then
                MsgBox("Required field missing: the field 'BarrierEID' could not be found in the input table #2. Exiting.")
                Exit Sub
            End If

            iTreatmentIndex = pFields.FindField("Treatment")
            If iTreatmentIndex = -1 Then
                MsgBox("Required field missing: the field 'treatment' could not be found in the input table #2. Exiting.")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("Error retrieving fields from table 2. Please check table.")
            Exit Sub
        End Try

        '------------------------------------------------------------------
        ' Get the Barrier ID, OptionNum, and permeability for given budget
        ' If optionnum is 1 (do nothing) then the permeability must be retrieved
        ' from the field set in FIPEX - a flexible field name.  
        '  
        ' Establish a list of barrier layers and permeability fields in this sub so that 
        ' if the feature class for an EID has been found in the TOC from a previous loop it 
        ' does get double-checked - waste.  
        ' Get the barrier ID Fields

        Dim iBarrierLayers As Integer
        Dim sBarrierIDLayer, sBarrierIDField, sBarrierPermField, sBarrierNaturalYNField As String
        Dim pBarrierIDObj As BarrierIDObj
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim bFoundLayerInMap As Boolean = False
        Dim pMap As IMap
        Dim pDoc As IDocument
        Dim pLayer As ILayer
        Dim pFeatureLayer As IFeatureLayer
        Dim sLayerPresentinTOC As String ' 'yes' or 'no'
        Dim j As Integer

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Exit Sub
        End Try

        If m_FiPEx.m_bLoaded = True Then
            iBarrierLayers = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierLayers > 0 Then
                For j = 0 To iBarrierLayers - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    sBarrierPermField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierPermField" + j.ToString))
                    
                    ' instead of using the NaturalYNField for what it's meant for, use it to keep track of whether the barrier
                    ' layer is found in the TOC.  This saves creating another custom object. 
                    'sBarrierNaturalYNField = Convert.ToString(m_FiPEx.pPropset.GetProperty("BarrierNaturalYNField" + j.ToString))
                    Try
                        bFoundLayerInMap = False
                        For i = 0 To pMap.LayerCount - 1
                            If pMap.Layer(i).Valid = True Then
                                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                                    pLayer = pMap.Layer(i)
                                    pFeatureLayer = CType(pLayer, IFeatureLayer)
                                    ' **************************************
                                    ' NO EDGE BARRIER SUPPORT YET
                                    ' **************************************
                                    If pFeatureLayer.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleJunction Then

                                        ' If the map layer name matches the FIPEX layer name
                                        If pFeatureLayer.Name = sBarrierIDLayer Then
                                            bFoundLayerInMap = True
                                            Exit For
                                        End If ' layer name match
                                    End If     ' simple junction
                                End If         ' LayerType Check
                            End If             ' layer is valid
                        Next                   ' Map Layer
                        If bFoundLayerInMap = False Then
                            sLayerPresentinTOC = "no"
                        Else
                            sLayerPresentinTOC = "yes"
                        End If
                    Catch ex As Exception
                        MsgBox("Error trying to retrieve the fields for the barrier layer " + sBarrierIDLayer)
                        Exit Sub
                    End Try

                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, _
                                                     sBarrierIDField, _
                                                     sBarrierPermField, _
                                                     sLayerPresentinTOC, _
                                                     Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            Else
                MsgBox("There are no barrier layers set in the FIPEX options.  This tool requires 'permeability' fields to be set for " _
                       + "barrier layers. Please check this setting and try again.")
                Exit Sub
            End If
        Else
            MsgBox("Error - It appears the FIPEX Options have not been set.  Please set these options (particularly the barrier 'permeability' field) to continue.")
            Exit Sub
        End If

        
        Dim pGLPKOptionsObject As GLPKOptionsObject = New GLPKOptionsObject(Nothing, _
                                                                            Nothing, _
                                                                            Nothing, _
                                                                            Nothing, _
                                                                            Nothing)
        Dim lGLPKOptionsObject As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject)
        Dim iBarrierEID, iOptionNum As Integer
        Dim dOptionPerm As Double
        Dim iPermFieldIndex As Integer
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pFeatureClass As IFeatureClass
        Dim pFeature2 As IFeature

        If m_UtilityNetworkAnalysisExt Is Nothing Then
            MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
            Me.Close()
        Else
            pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
            If pNetworkAnalysisExt.NetworkCount = 0 Then
                MsgBox("A Network must be loaded in ArcMap to use this tool.  Exiting.")
                Me.Close()
            End If
        End If
        Dim pGeometricNetwork As IGeometricNetwork
        Try
            pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Catch ex As Exception
            MsgBox("Trouble getting current geometric network.  Must have an active geometric network in ArcMap to use this toolset. Exiting.")
            Me.Close()
        End Try

        ' loop through the decisions table
        ' get any row that matches the budget amount
        ' get the EID for that decision
        ' query the feature and featureclass corresponding to EID
        ' do a re-query to get the feature field values for the permeability field 
        '     (this is due to an issue querying the first pFeature and thus pFeature2 is necessary)


        pCursor = pTable.Search(Nothing, False)
        pRow = pCursor.NextRow
        ' Loop through each row in the decisions table
        Do Until pRow Is Nothing

            ' get budget value
            pField = pRow.Fields.Field(iBudgetFieldIndex)
            If pField.Type = esriFieldType.esriFieldTypeInteger _
                Or pField.Type = esriFieldType.esriFieldTypeSingle _
                Or pField.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField.Type = esriFieldType.esriFieldTypeDouble Then

                Try
                    dTableBudget = Convert.ToDouble(pRow.Value(iBudgetFieldIndex))
                Catch ex As Exception
                    MsgBox("Error. Trouble converting the table budget to type 'double'. Please check the 'decisions' table for budget errors.")
                    Exit Sub
                End Try
            Else
                MsgBox("The budget field in the 'decisions' table cannot be of type " + pField.Type.ToString)
                Exit Sub
            End If

            ' if the row matches
            If dTableBudget = m_dBudget Then
                Try
                    iBarrierEID = Convert.ToInt32(pRow.Value(iBarrierEIDIndex))
                Catch ex As Exception
                    MsgBox("The barrier EID could not be converted to type 'integer'. Now exiting.")
                    Exit Sub
                End Try

                Try
                    sTreatment = Convert.ToString(pRow.Value(iTreatmentIndex))
                Catch ex As Exception
                    MsgBox("The treatment could not be converted to type 'treatment'. Now exiting.")
                    Exit Sub
                End Try

                Try
                    iOptionNum = Convert.ToInt32(pRow.Value(iOptionNumIndex))
                Catch ex As Exception
                    MsgBox("The treatment could not be converted to type 'OptionNum'. Now exiting.")
                    Exit Sub
                End Try

                ' need to get permeability gains for each decision from the feature class attribute table.
                ' unless it's the sink EID - never any decisions there, no perm field found in GDB (or set in FIPEX)
                If iBarrierEID <> iSinkEID Then

                    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=254373
                    ' this has better code which I will use. 
                    ' APPARENTLY must specify the ElementType. I would prefer not to because I would like to 
                    ' leave it open later to both junctions and edges, querying both edge/junction features.  
                    pEidHelper = New ESRI.ArcGIS.NetworkAnalysis.EIDHelper
                    pEidHelper.GeometricNetwork = pGeometricNetwork
                    pEidHelper.ReturnFeatures = True

                    pEnumNetEIDBuilder = New ESRI.ArcGIS.Geodatabase.EnumNetEIDArray
                    pEnumNetEIDBuilder.ElementType = ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction
                    pEnumNetEIDBuilder.Network = pGeometricNetwork.Network
                    pEnumNetEIDBuilder.Add(iBarrierEID)
                    pEnumNetEID = pEnumNetEIDBuilder

                    pEnumEidInfo = pEidHelper.CreateEnumEIDInfo(pEnumNetEID)
                    pEnumEidInfo.Reset()

                    pEidInfo = pEnumEidInfo.Next
                    pFeature = pEidInfo.Feature

                    ' the layer should also be present in the map TOC - otherwise problems may occur.  
                    ' lBarrierIDs (list of barriers and permeability fields)
                    '
                    bFoundLayer = False
                    For i = 0 To lBarrierIDs.Count - 1
                        If lBarrierIDs.Item(i).Layer = pFeature.Class.AliasName Then
                            bFoundLayer = True
                            Exit For
                        End If
                    Next
                    If bFoundLayer = False Then
                        MsgBox("The layer -" + pFeature.Class.AliasName + "- for this eid, " + iBarrierEID.ToString _
                               + ", could not be found in the FIPEX settings. Please add this layer to continue.")
                        Exit Sub
                    End If
                    If lBarrierIDs.Item(i).NaturalYNField = "no" Then
                        MsgBox("Cannot find the layer " + pFeature.Class.AliasName _
                               + " in the TOC.  Please add this barrier layer to map so analysis can continue.")
                        Exit Sub
                    End If

                    pFields = pFeature.Fields

                    ' find the option num field
                    ' if option is not 'do nothing' or NOT
                    If iOptionNum > 1 Then
                        ' get index of 'new' permeability field
                        iPermFieldIndex = pFields.FindField("option" + Convert.ToString(iOptionNum - 1) + "permafter")
                    Else
                        ' to get permeability field, need to retrieve the field from the one stored in the FIPEX settings 
                        ' for this EID / feature.  
                        ' have the EID / feature
                        ' have the list of barrier layers saved in FIPEX
                        ' the layer should be (a) saved in FIPEX and (b) field set for permeability. 
                        If lBarrierIDs.Item(i).PermField = "<None>" Then
                            MsgBox("The permeability field for layer " + lBarrierIDs.Item(i).Layer + " is not set" _
                                   + ". Please set the permeability field in the FIPEX Options Menu to continue.")
                            Exit Sub
                        Else
                            iPermFieldIndex = pFields.FindField(lBarrierIDs.Item(i).PermField)
                        End If
                    End If ' option num > 1

                    If iPermFieldIndex = -1 Then
                        If iOptionNum > 1 Then
                            MsgBox("Could not find the option number for field 'option" + Convert.ToString(iOptionNum - 1) + _
                            "permafter' in feature class " + pFeature.Class.AliasName + _
                            " for barrier EID " + iBarrierEID.ToString)
                        Else
                            MsgBox("Could not find the option number for field " + Convert.ToString(lBarrierIDs.Item(i).PermField) + _
                           " in feature class " + pFeature.Class.AliasName + _
                           " for barrier EID " + iBarrierEID.ToString)
                        End If
                        Exit Sub
                    Else
                        If iOptionNum > 1 Then
                            'MsgBox("Successfully found field 'option" + iOptionNum.ToString + _
                            '      "permafter' in feature class " + pFeature.Class.AliasName + _
                            '      "for barrier EID " + iBarrierEID.ToString)
                        Else
                            'MsgBox("Successfully found field " + lBarrierIDs.Item(i).PermField + _
                            '      " in feature class " + pFeature.Class.AliasName + _
                            '      "for barrier EID " + iBarrierEID.ToString)
                        End If
                    End If

                    ' get permeability value
                    Try
                        pField = pFields.Field(iPermFieldIndex)
                        If pField Is Nothing Then
                            MsgBox("Field could not be found in " + pFeature.Class.AliasName + " layer. ")
                            Exit Sub
                        End If
                    Catch ex As Exception
                        MsgBox("Field could not be found in " + pFeature.Class.AliasName + " layer. " + ex.Message)
                        Exit Sub
                    End Try

                    ' need to get value from pfeature - no columns were retrieved in the pfeature
                    Try
                        pFeatureClass = CType(pFeature.Class, IFeatureClass)
                        pFeature2 = pFeatureClass.GetFeature(pFeature.OID)
                    Catch ex As Exception
                        MsgBox("Ugh.")
                        Exit Sub
                    End Try

                    If pField.Type = esriFieldType.esriFieldTypeDouble Then
                        Try
                            dOptionPerm = Convert.ToDouble(pFeature2.Value(iPermFieldIndex))
                        Catch ex As Exception
                            MsgBox("Error. Trouble converting the permeability to type 'double'.")
                            Exit Sub
                        End Try
                    Else
                        MsgBox("Could not find the permeability field - not of type 'double'.  Please check all 'barriers' layers set in FIPEX Options. " + pField.Type.ToString)
                        Exit Sub
                    End If

                    pGLPKOptionsObject = New GLPKOptionsObject(iBarrierEID, iOptionNum, dOptionPerm, 0, Nothing)
                    lGLPKOptionsObject.Add(pGLPKOptionsObject)
                End If ' if this entry isn't for the network sink
            End If ' budgets match
            pRow = pCursor.NextRow
        Loop ' next decision option row

        ' ======================= Add 'Do Nothing' DECISIONS to LIST ===============================
        ' now have a list object of barrier EID and new Permeability for this budget amount
        ' (also could add in cost, but ignoring it)
        ' ISSUE here is that SOMETIMES (currently GLPK tables) include all decisions, even do nothing (decision option 1)
        ' and SOMETIMES (currently GRB tables) they do not.  The GRB tables omit decision option 1.  Therefore, the 
        ' list lGLPKOptionsObject may not include decision options for every EID.  And WON'T for the sink EID.  
        ' USE the master list of EIDs to cross-check and get the permeability if it is missing.  
        Dim bEIDFound As Boolean = False
        For i = 0 To m_lEIDs.Count - 1
            bEIDFound = False
            j = 0
            For j = 0 To lGLPKOptionsObject.Count - 1
                If lGLPKOptionsObject.Item(j).BarrierEID = m_lEIDs(i) Then
                    bEIDFound = True
                    Exit For
                End If
            Next
            ' if it isn't found and it's not the sink EID, 
            ' go through the same jazz as above
            ' (assume that because the EID is present in the master EID list table
            '   and it is not in the decisions table, that option number is 1)
            If bEIDFound = False Then
                If m_lEIDs(i) <> iSinkEID Then

                    ' http://forums.esri.com/Thread.asp?c=93&f=993&t=254373
                    ' this has better code which I will use. 
                    ' APPARENTLY must specify the ElementType. I would prefer not to because I would like to 
                    ' leave it open later to both junctions and edges, querying both edge/junction features.  
                    pEidHelper = New ESRI.ArcGIS.NetworkAnalysis.EIDHelper
                    pEidHelper.GeometricNetwork = pGeometricNetwork
                    pEidHelper.ReturnFeatures = True

                    pEnumNetEIDBuilder = New ESRI.ArcGIS.Geodatabase.EnumNetEIDArray
                    pEnumNetEIDBuilder.ElementType = ESRI.ArcGIS.Geodatabase.esriElementType.esriETJunction
                    pEnumNetEIDBuilder.Network = pGeometricNetwork.Network
                    pEnumNetEIDBuilder.Add(m_lEIDs(i))
                    pEnumNetEID = pEnumNetEIDBuilder

                    pEnumEidInfo = pEidHelper.CreateEnumEIDInfo(pEnumNetEID)
                    pEnumEidInfo.Reset()

                    pEidInfo = pEnumEidInfo.Next
                    pFeature = pEidInfo.Feature

                    ' cross ref and make sure layer is present in Map TOC
                    ' the layer should also be present in the map TOC - otherwise problems may occur.  
                    ' lBarrierIDs (list of barriers and permeability fields)
                    bFoundLayer = False
                    J = 0
                    For j = 0 To lBarrierIDs.Count - 1
                        If lBarrierIDs.Item(j).Layer = pFeature.Class.AliasName Then
                            bFoundLayer = True
                            Exit For
                        End If
                    Next
                    If bFoundLayer = False Then
                        MsgBox("The layer -" + pFeature.Class.AliasName + "- for this eid, " + iBarrierEID.ToString _
                               + ", could not be found in the FIPEX settings. Please add this layer to continue.")
                        Exit Sub
                    End If
                    If lBarrierIDs.Item(j).NaturalYNField = "no" Then
                        MsgBox("Cannot find the layer " + pFeature.Class.AliasName _
                               + " in the TOC.  Please add this barrier layer to map so analysis can continue.")
                        Exit Sub
                    End If

                    pFields = pFeature.Fields

                    ' to get permeability field, need to retrieve the field from the one stored in the FIPEX settings 
                    ' for this EID / feature.  
                    ' have the EID / feature
                    ' have the list of barrier layers saved in FIPEX
                    ' the layer should be (a) saved in FIPEX and (b) field set for permeability. 
                    If lBarrierIDs.Item(j).PermField = "<None>" Then
                        MsgBox("The permeability field for layer " + lBarrierIDs.Item(j).Layer + " is not set" _
                               + ". Please set the permeability field in the FIPEX Options Menu to continue.")
                        Exit Sub
                    Else
                        iPermFieldIndex = pFields.FindField(lBarrierIDs.Item(j).PermField)
                    End If

                    If iPermFieldIndex = -1 Then
                        MsgBox("Could not find the option number for field " + Convert.ToString(lBarrierIDs.Item(j).PermField) + _
                         " in feature class " + pFeature.Class.AliasName + _
                        " for barrier EID " + m_lEIDs(i).ToString)
                        Exit Sub
                    Else
                        'If m_lEIDs(i) = 5776 Then
                        '    MsgBox("Successfully got permeability field for Pockwock Lake Dam. That field is: " _
                        '           + lBarrierIDs.Item(j).PermField + " and it is number " + Convert.ToString(iPermFieldIndex) _
                        '           + ".")

                        'End If
                        'MsgBox("Successfully found field " + lBarrierIDs.Item(j).PermField + _
                        ' " in feature class " + pFeature.Class.AliasName + _
                        '" for barrier EID " + m_lEIDs(i).ToString)
                    End If

                    ' get permeability value
                    Try
                        pField = pFields.Field(iPermFieldIndex)
                        If pField Is Nothing Then
                            MsgBox("Field could not be found in " + pFeature.Class.AliasName + " layer. ")
                            Exit Sub
                        End If
                    Catch ex As Exception
                        MsgBox("Field could not be found in " + pFeature.Class.AliasName + " layer. " + ex.Message)
                        Exit Sub
                    End Try

                    ' need to get value from pfeature - no columns were retrieved in the pfeature
                    Try
                        pFeatureClass = CType(pFeature.Class, IFeatureClass)
                        pFeature2 = pFeatureClass.GetFeature(pFeature.OID)
                    Catch ex As Exception
                        MsgBox("Ugh.")
                        Exit Sub
                    End Try

                    If pField.Type = esriFieldType.esriFieldTypeDouble Then
                        Try
                            dOptionPerm = Convert.ToDouble(pFeature2.Value(iPermFieldIndex))
                        Catch ex As Exception
                            MsgBox("Error. Trouble converting the permeability to type 'double'.")
                            Exit Sub
                        End Try
                    Else
                        MsgBox("Could not find the permeability field - not of type 'double'.  Please check all 'barriers' layers set in FIPEX Options. " + pField.Type.ToString)
                        Exit Sub
                    End If

                    pGLPKOptionsObject = New GLPKOptionsObject(m_lEIDs(i), _
                                                               1, _
                                                               dOptionPerm, _
                                                               0, _
                                                               Nothing)
                Else
                    ' if it's the sink assume the permeability is 100% and optionnum = 1
                    pGLPKOptionsObject = New GLPKOptionsObject(iSinkEID, _
                                                               1, _
                                                               1, _
                                                               0, _
                                                               Nothing)
                End If

                lGLPKOptionsObject.Add(pGLPKOptionsObject)

            End If
        Next



        ' ================= Section 5: Run the analysis ===============
        ' save initial network settings

        ' --------- Be Nice and Save Original Selectable Layers -----------
        ' Also set all layers as selectable
        Dim pFLyrSlct As IFeatureLayer
        Dim iNotSelectables As List(Of Integer) = New List(Of Integer)
        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLyrSlct = CType(pMap.Layer(i), IFeatureLayer)
                    If pFLyrSlct.Selectable = False Then
                        iNotSelectables.Add(i)
                        pFLyrSlct.Selectable = True
                    End If
                End If
            End If
        Next
        i = 0

        ' ----------------- SAVE ORIGINAL GEONet SETTINGS -----------------
        ' DOES NOT CURRENTLY SAVE EDGE BARRIERS
        ' check if user has hit 'close/cancel'

        'If m_bCancel = True Then
        '    backgroundworker1.CancelAsync()
        '    backgroundworker1.Dispose()
        '    Exit Sub
        'End If
        'backgroundworker1.ReportProgress(8, "Saving Current Geometric Network Flags and Barriers")

        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        Dim pTraceTasks As ITraceTasks
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pOriginalBarriersListGEN, pOriginalEdgeFlagsListGEN, pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim bFlagDisplay As IFlagDisplay
        Dim bEID As Integer
        Dim pOriginalBarriersList, pOriginalEdgeFlagsList, pOriginaljuncFlagsList As IEnumNetEID
        Dim pEdgeFlagDisplay As IEdgeFlagDisplay
        Dim pFlagDisplay As IFlagDisplay
        Dim pSymbol As ISymbol
        Dim pFlagSymbol, pBarrierSymbol As ISimpleMarkerSymbol
        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        Dim bFCID, bFID, bSubID As Integer
        Dim iFCID, iFID, iSubID As Integer
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim pRgbColor As IRgbColor = New RgbColor
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)
        ' SET SYMBOLS
        ' For the Flag marker
        With pRgbColor
            .Red = 0
            .Green = 255
            .Blue = 0
        End With
        pSimpleMarkerSymbol = New SimpleMarkerSymbol
        With pSimpleMarkerSymbol
            .Color = pRgbColor
            .Style = esriSimpleMarkerStyle.esriSMSSquare
            .Outline = True
            .Size = 10
        End With

        ' Result is a global variable containing a flag marker
        pFlagSymbol = pSimpleMarkerSymbol
        pRgbColor = New RgbColor

        ' Set the barrier symbol color and parameters
        With pRgbColor
            .Red = 255
            .Green = 0
            .Blue = 0
        End With
        pSimpleMarkerSymbol = New SimpleMarkerSymbol
        With pSimpleMarkerSymbol
            .Color = pRgbColor
            .Style = esriSimpleMarkerStyle.esriSMSX
            .Outline = True
            .Size = 10
        End With

        ' Result is a global variable containing a barrier marker
        pBarrierSymbol = pSimpleMarkerSymbol
        ' END SET SYMBOLS 

        Try
            pDoc = m_app.Document
            ' hook into ArcMap
            pMxDoc = CType(pDoc, IMxDocument)
            pMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox("Trouble getting reference to map doc. Exiting")
            Exit Sub
        End Try

        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        ' QI for the ITraceTasks interface using IUtilityAnalysisExt
        pTraceTasks = CType(m_UtilityNetworkAnalysisExt, ITraceTasks)

        ' update extension with the results
        ' QI for the INetworkAnalysisExtResults interface using IUTilityNetworkAnalysisExt
        pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)

        ' clear any leftover results from previous calls to the cmd
        pNetworkAnalysisExtResults.ClearResults()

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)

        ' Was using ArrayList because of advantage of 'count' and 'add' properties
        ' but EnumNetEIDBuilderGEN addition to 9.2 has this functionality
        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        Dim iOriginalJunctionBarrierCount As Integer
        Try

            iOriginalJunctionBarrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        Catch ex As Exception
            MsgBox("Trouble getting the current barrier count from the loaded network. " + ex.Message)
            Exit Sub
        End Try

        ' Save the barriers
        For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            bFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

        ' Save the flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            bFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
            'pOriginaljuncFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)


        ' ******** NO EDGE FLAG SUPPORT YET *********
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            bFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETEdge)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)

        ' ******************************************
        pNetworkAnalysisExtFlags.ClearFlags()
        pNetworkAnalysisExtBarriers.ClearBarriers()

        ' ------------------------- END SAVE ORIGINAL GEONET SETTINGS -----------------

        ' ****************************************************************************
        ' ========================== BEGIN MAIN ANALYSIS BIT ========================
        Dim iEID As Integer
        Dim pEID_CPERM_AndDir As EIDCPermAndDir = New EIDCPermAndDir(Nothing, Nothing, Nothing)
        Dim lEID_CPERM_andDir As List(Of EIDCPermAndDir) = New List(Of EIDCPermAndDir) ' gets reset each order loop
        Dim pID As New UID
       
        Try
            If m_bUndirected = False Then
                pEID_CPERM_AndDir.Barrier = iSinkEID
                pEID_CPERM_AndDir.cPerm = 1
            Else
                pEID_CPERM_AndDir.Barrier = iCentralBarrierEID
                ' need to find perm for this barrier
                ' note: this barrier is *upstream* of the largest maximal subnetwork
                ' note: IF THE UNDIRECTED MODEL CHANGES AND RETURNS THE NODE BELOW THE LARGEST MAXIMAL
                '       SUBNETWORK THEN FROM HERE DOWN NEEDS TO CHANGE SLIGHLTLY. 
                '       SPECIFICALLY THE FIRST CPERM AT THE CENTRAL BARRIER EID MUST BE CONSIDERED 1
                ' *NOTE*: thank you, me!
                '       the code below has been changed to reflect the choice of changing the 'central' 
                '       barrier ID to be _below_ the single larges undirected subnetwork

                'For i = 0 To lGLPKOptionsObject.Count - 1
                '    If lGLPKOptionsObject.Item(i).BarrierEID = iCentralBarrierEID Then
                '        pEID_CPERM_AndDir.cPerm = lGLPKOptionsObject.Item(i).BarrierPerm
                '        Exit For
                '    End If
                'Next

                pEID_CPERM_AndDir.cPerm = 1
            End If
        Catch ex As Exception
            MsgBox("Problem getting central EID or sink EID for first upstream analysis. " + ex.Message)
            Exit Sub
        End Try

        lEID_CPERM_andDir.Add(pEID_CPERM_AndDir)

        'If pEID_CPERM_AndDir.Barrier = 11760 Then
        '    MsgBox("Impassable Channel enountered.")
        'End If

        pActiveView.Refresh() ' refresh the view
        DoIterativeUpstreamAnalysis(pMap, _
                                    lEID_CPERM_andDir, _
                                    lGLPKOptionsObject, _
                                    lLineLayersFieldsCHECKED, _
                                    lPolyLayersFieldsCHECKED, _
                                    pBarrierSymbol, _
                                    pFlagSymbol, _
                                    pActiveView)

        pActiveView.Refresh() ' refresh the view
        If m_bUndirected = True Then
        ' ==================== Perform DOWNSTREAM ANALYSIS ================
        ' 
        ' Logic:
        ' At barrier X get downstream barrier Y
        '   Set Flag on Y 
        '   Get upstream(edges)
        '   CPERM for edges = CPERMx
        '   Get flowstopping junctions upstream 
        '   Filter them (sources **OR X**)
        '   For each barrier upstream of Y = for each Y-up
        '   Get CPERMyup = CPERMx * PERMyup
        '   Perform upstream analysis for each (call DoIterativeUpstreamAnalysis sub_

        ' Now the permeability at the central node / sink 
            ' is whatever it is after optimal decision made at it
            For i = 0 To lGLPKOptionsObject.Count - 1
                If lGLPKOptionsObject.Item(i).BarrierEID = iCentralBarrierEID Then
                    pEID_CPERM_AndDir.cPerm = lGLPKOptionsObject.Item(i).BarrierPerm
                    Exit For
                End If
            Next
            'pEID_CPERM_AndDir.cPerm = 1
            pNetworkAnalysisExtBarriers.ClearBarriers()
            pNetworkAnalysisExtFlags.ClearFlags()

            'If pEID_CPERM_AndDir.Barrier = 11760 Then
            '    MsgBox("Impassable Channel enountered.")
            'End If

            pActiveView.Refresh() ' refresh the view
            DoIterativeDownstreamAnalysis(pMap, _
                                   pEID_CPERM_AndDir, _
                                   lGLPKOptionsObject, _
                                   lLineLayersFieldsCHECKED, _
                                   lPolyLayersFieldsCHECKED, _
                                   pBarrierSymbol, _
                                   pFlagSymbol, _
                                   pActiveView, _
                                   iSinkEID)


        End If
        ' ==================== END PERFORM DOWNSTREAM ANALYSIS ================
        ' 

        ' ================== UPDATE LINES LAYERS WITH CPERM ==================
        Dim pEditor As ESRI.ArcGIS.Editor.IEditor

        ' get reference to editor extension
        Try
            pID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"
            pEditor = m_app.FindExtensionByCLSID(pID)
            If pEditor Is Nothing Then
                MsgBox("Error getting reference to the Editor extension. Exiting. ")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Editor extension. Exiting. " + ex.Message)
            Exit Sub
        End Try

        ' this better fucking do it. 
        UpdateLineWithCPERM(pEditor)


        ' ========================== SELECT DECISION =========================
        ' have this object with EIDs and decisions
        ' if they are not 1, or do nothing, then select them on the map 
        ' can use the network to do this, using createselection from junction enumeration object
        ' this is the master list object: lGLPKOptionsObject


        Dim pResultsEIDsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        pResultsEIDsGEN.Network = pNetworkAnalysisExt.CurrentNetwork.Network
        pResultsEIDsGEN.ElementType = esriElementType.esriETJunction

        For i = 0 To lGLPKOptionsObject.Count - 1
            If lGLPKOptionsObject(i).OptionNum > 1 Then
                pResultsEIDsGEN.Add(lGLPKOptionsObject(i).BarrierEID)
            End If
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
            pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)
            pNetworkAnalysisExtResults.ClearResults()
            pNetworkAnalysisExtResults.ResultsAsSelection = True
            'pNetworkAnalysisExtResults.SetResults(pResultsJunctions, Nothing)
            pNetworkAnalysisExtResults.CreateSelection(pResultsJunctions, pResultEdges)
        Catch ex As Exception
            MsgBox("Problem encountered setting results display. Exiting.")
            Exit Sub
        End Try

        pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
        pActiveView.ScreenDisplay.UpdateWindow()

        ' ======================= END SELECT DECISIONS ======================


        ' ========================== END MAIN ANALYSIS BIT ========================
        ' ****************************************************************************


        ' 
        ' ----------------- RESET ORIGINAL GEONet SETTINGS ----------------- 
        '  Reset things the way the user had them

        ' check if user has hit 'close/cancel'
        'If m_bCancel = True Then
        '    backgroundworker1.CancelAsync()
        '    backgroundworker1.Dispose()
        '    Exit Sub
        'End If
        'backgroundworker1.ReportProgress(iProgress + 10, "Resetting Flags And Barriers")

        '  RESET BARRIERS
        m = 0
        pOriginalBarriersList.Reset()
        For m = 0 To pOriginalBarriersList.Count - 1
            bEID = pOriginalBarriersList.Next
            pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

            ' Display the barriers as a JunctionFlagDisplay type
            pFlagDisplay = New JunctionFlagDisplay
            pSymbol = CType(pBarrierSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = bFCID
                .FID = bFID
                .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
            pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
        Next
        '  END RESET BARRIERS 
        'MsgBox("Debug:62")
        ' Clear current flags
        pNetworkAnalysisExtFlags.ClearFlags()

        '  RESET FLAGS
        ' restore all EDGE flags
        m = 0
        pOriginalEdgeFlagsList.Reset()
        For m = 0 To pOriginalEdgeFlagsList.Count - 1

            iEID = pOriginalEdgeFlagsList.Next
            ' Query the corresponding user ID's to the element ID
            pNetElements.QueryIDs(iEID, esriElementType.esriETEdge, iFCID, iFID, iSubID)

            ' Display the flags as a JunfctionFlagDisplay type
            pFlagDisplay = New EdgeFlagDisplay
            pSymbol = CType(pFlagSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = iFCID
                .FID = iFID
                .Geometry = pGeometricNetwork.GeometryForEdgeEID(iEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)
            pNetworkAnalysisExtFlags.AddEdgeFlag(pEdgeFlagDisplay)
        Next

        ' restore all JUNCTION Flags
        m = 0
        pOriginaljuncFlagsList.Reset()
        For m = 0 To pOriginaljuncFlagsList.Count - 1

            iEID = pOriginaljuncFlagsList.Next
            ' Query the corresponding user ID's to the element ID
            pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

            ' Display the flags as a JunfctionFlagDisplay type
            pFlagDisplay = New JunctionFlagDisplay
            pSymbol = CType(pFlagSymbol, ISymbol)
            With pFlagDisplay
                .FeatureClassID = iFCID
                .FID = iFID
                .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEID)
                .Symbol = pSymbol
            End With

            ' Add the flags to the logical network
            pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
            pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)

        Next

        ' --------- Be Nice and Reset Original Selectable Layers -----------
        ' Also set all layers as selectable
        For i = 0 To iNotSelectables.Count - 1
            If pMap.Layer(iNotSelectables.Item(i)).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLyrSlct = CType(pMap.Layer(i), IFeatureLayer)
                    pFLyrSlct.Selectable = False
                End If
            End If
        Next
        i = 0

        pActiveView.Refresh() ' refresh the view
    End Sub

    Public Function TraceFlowSolverSetup() As ITraceFlowSolver
        ' Prepares the network for tracing

        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pGeometricNetwork As IGeometricNetwork
        Dim pUtilityNetwork As IUtilityNetwork
        Dim pNetSolver As INetSolver
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pEdgeElementBarriers As INetElementBarriers     ' Should these next two be INetElementBarriersGEN?
        Dim pJunctionElementBarriers As INetElementBarriers
        Dim pSelectionSetBarriers As ISelectionSetBarriers
        Dim pFeatureLayer As IFeatureLayer
        Dim pFlagDisplay As IFlagDisplay
        Dim pEdgeFlagDisplay As IEdgeFlagDisplay
        Dim pEdgeFlags() As IEdgeFlag
        Dim pJunctionFlags() As IJunctionFlag
        Dim pNetFlag As INetFlag
        Dim pEdgeFlag As IEdgeFlag
        'Dim pTraceFlowSolver As ITraceFlowSolver
        Dim pTraceTasks As ITraceTasks
        Dim pNetworkAnalysisExtWeightFilter As INetworkAnalysisExtWeightFilter
        Dim pNetSchema As INetSchema
        Dim pNetWeight As INetWeight
        Dim eWeightFilterType As esriWeightFilterType
        Dim pNetSolverWeights As INetSolverWeights
        Dim pTraceFlowSolver As ITraceFlowSolver

        ' Must change from Object type in VB.Net to Variant in VB6
        Dim lngFromValues() As Object
        Dim lngToValues() As Object

        Dim lngFeatureLayerCount As Integer
        Dim lngEdgeFlagCount As Integer
        Dim lngJunctionFlagCount As Integer
        Dim binFeatureLayerDisabled As Boolean
        Dim binApplyNotOperator As Boolean
        Dim lngFilterRangeCount As Integer
        Dim i As Integer

        If m_FiPEx Is Nothing Then
            m_FiPEx = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
        If m_UtilityNetworkAnalysisExt Is Nothing Then
            m_UtilityNetworkAnalysisExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetUNAExt
        End If
        'Dim FiPEx__1 As FishPassageExtension = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetExtension
        'Dim pUNAExt As IUtilityNetworkAnalysisExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
        'If m_pNetworkAnalysisExt Is Nothing Then
        '    m_pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        'End If

        ' Get reference to the current network through Utility Network interface
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)

        ' Assign a reference to the current geometric network via a local variable
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

        ' Assign the network to local IUtilityNetwork variable
        ' however, do it through GeometricNetwork which returns
        ' a QI for utilitynetwork using INetwork.... ?

        ' Not all members of IUtilityNetwork be directly calleable
        ' With VB.NET and it has been recommended to use IUtilityNetworkGEN
        ' instead...

        pUtilityNetwork = CType(pGeometricNetwork.Network, IUtilityNetwork)

        ' initialize the trace flow solver and set the network
        pNetSolver = New TraceFlowSolver
        pNetSolver.SourceNetwork = pUtilityNetwork

        ' Get barriers for the network
        ' QI for the interface using the IUtilityNetworkAnalysisExt
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        ' Get the element barriers
        pNetworkAnalysisExtBarriers.CreateElementBarriers(pJunctionElementBarriers, pEdgeElementBarriers)
        ' Get the selection set barriers
        pNetworkAnalysisExtBarriers.CreateSelectionBarriers(pSelectionSetBarriers)
        ' set the barriers for the network solver
        pNetSolver.ElementBarriers(esriElementType.esriETEdge) = pEdgeElementBarriers
        pNetSolver.ElementBarriers(esriElementType.esriETJunction) = pJunctionElementBarriers
        pNetSolver.SelectionSetBarriers = pSelectionSetBarriers


        ' set the barriers for the network solver
        ' for each feature determine if it is enabled.
        ' if it is disabled notify network solver
        ' determine the number of feature layers belonging to network
        lngFeatureLayerCount = pNetworkAnalysisExt.FeatureLayerCount
        For i = 0 To lngFeatureLayerCount - 1
            ' get the next feature layer and determine if it is disabled
            ' assign the feature layer to a local IFeatureLayer variable
            pFeatureLayer = pNetworkAnalysisExt.FeatureLayer(i)
            ' determine if the feature layer is disabled
            binFeatureLayerDisabled = pNetworkAnalysisExtBarriers.GetDisabledLayer(pFeatureLayer)
            ' if it is disabled then notify the network solver
            If binFeatureLayerDisabled Then
                pNetSolver.DisableElementClass(pFeatureLayer.FeatureClass.FeatureClassID)
            End If
        Next

        ' set up the weight filters for the network
        ' QI for the INetworkAnalysisEctWeightFilter interface using
        ' the UtilityNetworkAnalysisExt
        pNetworkAnalysisExtWeightFilter = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtWeightFilter)
        ' QI for the INetSolverWeights interace using the INetSolver
        pNetSolverWeights = CType(pNetSolver, INetSolverWeights)
        ' QI for the INetSchema interface using IUtilityNetwork
        pNetSchema = CType(pUtilityNetwork, INetSchema)

        With pNetworkAnalysisExtWeightFilter

            ' Create the junction weight filter
            lngFilterRangeCount = pNetworkAnalysisExtWeightFilter.FilterRangeCount(esriElementType.esriETJunction)
            ' If there are any weight filters
            If lngFilterRangeCount > 0 Then
                ' get a NetWeight object from the INetSchema interface
                pNetWeight = pNetSchema.WeightByName(.JunctionWeightFilterName)
                ' get the type and Not operator status from the InetworkAnalysisExtWeightFileter interface
                .GetFilterType(esriElementType.esriETJunction, eWeightFilterType, binApplyNotOperator)
                ' redimension the weight filter ranges arrays and get the ranges
                ReDim lngFromValues(0 To lngFilterRangeCount - 1)
                ReDim lngToValues(0 To lngFilterRangeCount - 1)
                For i = 0 To lngFilterRangeCount - 1
                    .GetFilterRange(esriElementType.esriETJunction, i, lngFromValues(i), lngToValues(i))
                Next
                ' add the filter ranges to the network solver
                pNetSolverWeights.SetFilterRanges(esriElementType.esriETJunction, lngFilterRangeCount, lngFromValues(0), lngToValues(0))

            End If


            ' create the edge weight filters
            lngFilterRangeCount = pNetworkAnalysisExtWeightFilter.FilterRangeCount(esriElementType.esriETEdge)
            If lngFilterRangeCount > 0 Then

                ' get the type and Not operator status from the INetworkAnalysisExtWeightFilter interface
                .GetFilterType(esriElementType.esriETEdge, eWeightFilterType, binApplyNotOperator)

                ' get a NetWeight object from the INetSchema interface
                pNetWeight = pNetSchema.WeightByName(.FromToEdgeWeightFilterName)
                ' add the weight filter to the network solver
                pNetSolverWeights.FromToEdgeFilterWeight = pNetWeight

                ' get a NetWeight object from the INetScema interface
                pNetWeight = pNetSchema.WeightByName(.ToFromEdgeWeightFilterName)
                ' add the weight filter to the network solver
                pNetSolverWeights.ToFromEdgeFilterWeight = pNetWeight

                'get the filter ranges and apply them to the network solver
                ReDim lngFromValues(0 To lngFilterRangeCount - 1)
                ReDim lngToValues(0 To lngFilterRangeCount - 1)
                For i = 0 To lngFilterRangeCount - 1
                    .GetFilterRange(esriElementType.esriETEdge, i, lngFromValues(i), lngToValues(i))
                Next

                pNetSolverWeights.SetFilterType(esriElementType.esriETEdge, eWeightFilterType, binApplyNotOperator)
                pNetSolverWeights.SetFilterRanges(esriElementType.esriETEdge, lngFilterRangeCount, lngFromValues(0), lngToValues(0))

            End If

        End With


        ' assign the flags to the Network Solver
        ' get the edge flags
        ' QI for the ITraceFlowSolver interface using INetSolver Interface
        pTraceFlowSolver = CType(pNetSolver, ITraceFlowSolver)
        ' QI for the INetworkAnalysisExtFlags interface using the IUtilitNetworkAnalysisExt
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        ' determine the number of edge flags on the current network
        lngEdgeFlagCount = pNetworkAnalysisExtFlags.EdgeFlagCount

        ' if there are edge flags then
        If Not lngEdgeFlagCount = 0 Then
            ' Redimension the array to hold the correct number of edge flags
            ReDim pEdgeFlags(0 To lngEdgeFlagCount - 1)
            For i = 0 To lngEdgeFlagCount - 1
                ' assign a local variable for IFlagDisplay and IEdgeFlagDisplay variables
                pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
                pEdgeFlagDisplay = CType(pFlagDisplay, IEdgeFlagDisplay)

                ' co-create(?)a new EdgeFlag Object
                pNetFlag = New EdgeFlag
                pEdgeFlag = CType(pNetFlag, IEdgeFlag)
                ' assign the properties of the EdgeFlagDisplay object to he EdgeFlag Object
                ' I think this is where you could determine the distance along the line
                ' an edge flag is. 
                pEdgeFlag.Position = Convert.ToSingle(pEdgeFlagDisplay.Percentage)
                pNetFlag.UserClassID = pFlagDisplay.FeatureClassID
                pNetFlag.UserID = pFlagDisplay.FID
                pNetFlag.UserSubID = pFlagDisplay.SubID
                ' add the new EdgeFlag object to the array
                pEdgeFlags(i) = CType(pNetFlag, IEdgeFlag)
            Next
            ' add the edge flags to the network solver
            'pTraceFlowSolver.PutEdgeOrigins(lngEdgeFlagCount, pEdgeFlags(0))
        End If


        ' Get the junction flags
        ' determine the number of junction flags on the network
        lngJunctionFlagCount = pNetworkAnalysisExtFlags.JunctionFlagCount
        ' only execute this if there are juntion flags
        If Not lngJunctionFlagCount = 0 Then
            ' redimension the array to hold the correct number of junction flags
            ReDim pJunctionFlags(0 To lngJunctionFlagCount - 1)
            For i = 0 To lngJunctionFlagCount - 1

                ' assign to a local IFlagDisplay variable
                pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
                ' co-create a new JunctionFlag object
                pNetFlag = New JunctionFlag
                ' assign the properties of the JunctionFlagDisplay object to the JunctionFlag object
                pNetFlag.UserClassID = pFlagDisplay.FeatureClassID
                pNetFlag.UserID = pFlagDisplay.FID
                pNetFlag.UserSubID = pFlagDisplay.SubID
                ' add the new junction flag to the array of junction flags
                pJunctionFlags(i) = CType(pNetFlag, IJunctionFlag)
            Next
            'add the junction flags to the network solver
            ' This function is not calleable from VB.NET:
            'pTraceFlowSolver.PutJunctionOrigins(lngJunctionFlagCount, pJunctionFlags(0))
            ' So we use this function
            Dim pTraceFlowSolverGEN As ITraceFlowSolverGEN
            pTraceFlowSolverGEN = CType(pNetSolver, ITraceFlowSolverGEN)
            pTraceFlowSolverGEN.PutJunctionOrigins(pJunctionFlags)

        End If

        ' set the option for tracing on indeterminate flow
        ' QI for the ITraceTasksinterface using IUtilityNetworkAnlysisExt
        pTraceTasks = CType(m_UtilityNetworkAnalysisExt, ITraceTasks)
        'pTraceFlowSolver.TraceIndeterminateFlow = pTraceTasks.TraceIndeterminateFlow
        pTraceFlowSolver.TraceIndeterminateFlow = True
        pTraceTasks.TraceIndeterminateFlow = True

        ' pass the traceFlowSolver object back to the network solver
        TraceFlowSolverSetup = pTraceFlowSolver

    End Function

    Public Sub IntersectFeatures(ByRef lPolyLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                 ByVal dCPERM As Double, _
                                 ByRef pEnumLayer As IEnumLayer)

        ' This intersect features sub was derived from the IntersectFeatures Sub in the Analaysis Class
        ' On August 24, 2012.  
        ' It was modified to 
        ' (a) not intersect features if they weren't included in lPolyLayersFieldsCHECKED or lLineLayersFieldsCHECKED
        ' i.e., if they weren't checked by the user in the listbox. 
        ' (b) not to read extension settings. 

        'Read Extension Settings
        ' ================== READ EXTENSION SETTINGS =================

        'Dim bDBF As Boolean = False         ' Include DBF output default 'no'
        'Dim pLLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
        'Dim pPLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
        'Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        'Dim iLinesCount As Integer = 0      ' number of lines layers currently using
        'Dim HabLayerObj As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property
        '' object to hold stats to add to list. 
        'Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim m As Integer = 0
        Dim k As Integer = 0
        Dim j As Integer = 0
        Dim i As Integer = 0

        'If m_FiPEx.m_bLoaded = True Then

        '    ' Populate a list of the layers using and habitat summary fields.
        '    ' match any of the polygon layers saved in stream to those in listboxes 
        '    iPolysCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numPolys"))
        '    If iPolysCount > 0 Then
        '        For k = 0 To iPolysCount - 1
        '            'sPolyLayer = m_FiPEx.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
        '            HabLayerObj = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
        '            With HabLayerObj
        '                .Layer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
        '                .ClsField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyClassField" + k.ToString))
        '                .QuanField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyQuanField" + k.ToString))
        '                .UnitField = Convert.ToString(m_FiPEx.pPropset.GetProperty("PolyUnitField" + k.ToString))
        '            End With

        '            ' Load that object into the list
        '            pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
        '        Next
        '    End If

        '    ' Need to be sure that quantity field has been assigned for each
        '    ' layer using. 
        '    Dim iCount1 As Integer = pPLayersFields.Count

        '    If iCount1 > 0 Then
        '        For m = 0 To iCount1 - 1
        '            If pPLayersFields.Item(m).QuanField = "Not set" Then
        '                System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
        '                Exit Sub
        '            End If
        '        Next
        '    End If

        '    iLinesCount = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numLines"))
        '    Dim HabLayerObj2 As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

        '    ' match any of the line layers saved in stream to those in listboxes
        '    If iLinesCount > 0 Then
        '        For j = 0 To iLinesCount - 1
        '            'sLineLayer = m_FiPEx.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
        '            HabLayerObj2 = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
        '            With HabLayerObj2
        '                '.Layer = sLineLayer
        '                .Layer = Convert.ToString(m_FiPEx.pPropset.GetProperty("IncLine" + j.ToString))
        '                .ClsField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineClassField" + j.ToString))
        '                .QuanField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineQuanField" + j.ToString))
        '                .UnitField = Convert.ToString(m_FiPEx.pPropset.GetProperty("LineUnitField" + j.ToString))
        '            End With
        '            ' add to the module level list
        '            pLLayersFields.Add(HabLayerObj2)
        '        Next
        '    End If

        '    ' Need to be sure that quantity field has been assigned for each
        '    ' layer using. 
        '    iCount1 = pLLayersFields.Count
        '    If iCount1 > 0 Then
        '        For m = 0 To iCount1 - 1
        '            If pLLayersFields.Item(m).QuanField = "Not set" Then
        '                System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for river layer. Please choose a field in the options menu.", "Parameter Missing")
        '                Exit Sub
        '            End If
        '        Next
        '    End If
        'Else
        '    System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
        '    Exit Sub
        'End If

        ' ====================================================================
        ' ////////////////////////////////////////////////////////////////////
        ' BEGIN NEW CODE
        ' created August 26, 2012 by Greig Oldford
        ' purpose: attempt to intersect using a 'spatial filter' query rather than 
        '          an intersect.  This should help avoid long waits for each GP session.
        ' Logic:
        ' > For each layer in the map, 
        '   - If it's in the polygon-checked list get the featurelayer object for the polygon
        '     Add the polygon feat. layer and cPerm field to a list to do spatial query with 
        '      (if it's not in list already due to duplicate layers in TOC) 
        '   - If it's a lines layer, and has a selection
        '   (we will anticipate intersection with non-network lines in the future)
        '     Add this pfeaturelayer and selectionset to a list to do spatial query with 
        '      (if it's not already in the list due to duplicate layers in TOC)
        ' > For each polygon layer in the featurelayer list        '     
        '   > For each lines layer in the list
        '       get a cursor for the lines selection set
        '       loop through the cursor and add the geometry of each feature to the spatial filter.
        '         create a selection in polygon layer by adding to existing selection using the spatial filter.
        '       If there's a selectionset for the polygon layer
        '         store the selectionset for layer
        '         Get the workspace for this layer
        '         Get the index of the permeability field
        '         Store the workspace, selectionset, featurelayer, CPERM
        '  Will be using code from here: 
        ' http://gis.stackexchange.com/questions/28266/geoprocessor-select-by-location-using-features-instead-of-layers
        '

        Dim pFeatureLayerPoly, pFeatureLayerLine, pFeatureLayerMap As IFeatureLayer

        Dim lFeatureLayerAndCPermFieldPOLY As List(Of FeatureLayerAndCPermField) = New List(Of FeatureLayerAndCPermField)
        Dim pFeatureLayerAndCPermField As FeatureLayerAndCPermField = New FeatureLayerAndCPermField(Nothing, Nothing)

        Dim lFeatureLayerAndSelSetLINE As List(Of FeatureLayerAndSelectionSet) = New List(Of FeatureLayerAndSelectionSet)
        Dim pFeatureLayerAndSelectionSet As FeatureLayerAndSelectionSet = New FeatureLayerAndSelectionSet(Nothing, Nothing)

        Dim pFeatureSelectionLine As IFeatureSelection
        Dim pSelectionSetLine As ISelectionSet
        Dim bFeatureLayerFound As Boolean

        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayerMap = CType(pEnumLayer.Next, IFeatureLayer)
        Do While Not pFeatureLayerMap Is Nothing ' these two lines must be separate
            If pFeatureLayerMap.Valid = True Then ' or there will be an empty object ref

                ' if it's a polygon cross-ref with list of polygons checked by user
                If pFeatureLayerMap.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                    ' loop through checked polys list and see if there's a match
                    ' if there is add it to list of poly FeatLayers (if not already there)
                    For i = 0 To lPolyLayersFieldsCHECKED.Count - 1
                        If lPolyLayersFieldsCHECKED(i).Layer = pFeatureLayerMap.Name Then
                            bFeatureLayerFound = False
                            For j = 0 To lFeatureLayerAndCPermFieldPOLY.Count - 1
                                If lFeatureLayerAndCPermFieldPOLY(j).pFeatureLayer Is pFeatureLayerMap Then
                                    bFeatureLayerFound = True
                                End If
                            Next
                            If bFeatureLayerFound = False Then
                                pFeatureLayerAndCPermField = New FeatureLayerAndCPermField(pFeatureLayerMap, lPolyLayersFieldsCHECKED(i).CumPermField)
                                lFeatureLayerAndCPermFieldPOLY.Add(pFeatureLayerAndCPermField)
                            End If
                        End If
                    Next

                    ' If it's  a lines layer check selection set and add to lines layer list
                ElseIf pFeatureLayerMap.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline _
                       Or pFeatureLayerMap.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
                       Or pFeatureLayerMap.FeatureClass.FeatureType = esriFeatureType.esriFTSimpleEdge Then

                    ' get the selection set
                    pFeatureSelectionLine = Nothing
                    pFeatureSelectionLine = CType(pFeatureLayerMap, IFeatureSelection)
                    pSelectionSetLine = Nothing ' like to reset these cuz sometimes whothefuckknowswhy it causes issues otherwise
                    pSelectionSetLine = pFeatureSelectionLine.SelectionSet

                    ' If there's a selection then add this layer to the list of unique lines feature layers
                    If pSelectionSetLine.Count > 0 Then
                        bFeatureLayerFound = False
                        For i = 0 To lFeatureLayerAndSelSetLINE.Count - 1
                            If lFeatureLayerAndSelSetLINE(i).pFeatureLayer Is pFeatureLayerMap Then
                                bFeatureLayerFound = True
                            End If
                        Next
                        If bFeatureLayerFound = False Then
                            pFeatureLayerAndSelectionSet = New FeatureLayerAndSelectionSet(pFeatureLayerMap, pSelectionSetLine)
                            lFeatureLayerAndSelSetLINE.Add(pFeatureLayerAndSelectionSet)
                        End If
                    End If

                End If ' feature type matches
            End If ' valid layer

            pFeatureLayerMap = CType(pEnumLayer.Next, IFeatureLayer)
        Loop

        ' if there are no line features selected then exit sub
        If lFeatureLayerAndSelSetLINE.Count = 0 Then
            Exit Sub
        End If

        Dim pWorkspacePoly As IWorkspace
        Dim iCPERMFieldIndex As Integer
        Dim pTablePoly As ITable
        Dim pLineCursor As IFeatureCursor
        Dim pLineFeature As IFeature
        Dim pSpatialFilter As ISpatialFilter
        Dim pFeatureSelectionPOLY As IFeatureSelection
        Dim pSelectionSetPOLY As ISelectionSet
        Dim pSelectAndUpdateFeaturesObject As SelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(Nothing, _
                                                                                                                             Nothing, _
                                                                                                                             Nothing, _
                                                                                                                             Nothing, _
                                                                                                                             Nothing)

        ' loop through polygons layer in list
        For i = 0 To lFeatureLayerAndCPermFieldPOLY.Count - 1

            ' get the polygon feature layer
            Try
                pFeatureLayerPoly = lFeatureLayerAndCPermFieldPOLY(i).pFeatureLayer
            Catch ex As Exception
                MsgBox("Trouble getting the polygon feature layer back from the object list in 'intersect features' sub. " + ex.Message)
                Exit Sub
            End Try

            ' get the field index for the CPERM field
            pTablePoly = CType(pFeatureLayerPoly.FeatureClass, ITable)
            iCPERMFieldIndex = pTablePoly.FindField(lFeatureLayerAndCPermFieldPOLY(i).sCPermField)
            If iCPERMFieldIndex = -1 Then
                MsgBox("Error in 'intersect features' sub. There must be a field named " _
                       & lFeatureLayerAndCPermFieldPOLY(i).sCPermField & " in layer " _
                       & pFeatureLayerPoly.Name)
                Exit Sub
            End If

            ' get the workspace for the polygon
            Try
                pWorkspacePoly = pFeatureLayerPoly.FeatureClass.FeatureDataset.Workspace
            Catch ex As Exception
                MsgBox("Trouble getting polygon workspace from the polygon feature layer in 'intersect features' sub. " + ex.Message)
                Exit Sub
            End Try

            ' set the featureselection object for this polygon
            Try
                pFeatureSelectionPOLY = CType(pFeatureLayerPoly, IFeatureSelection)
            Catch ex As Exception
                MsgBox("Trouble setting polygon feature selection object in 'intersect features' sub. " + ex.Message)
                Exit Sub
            End Try

            For j = 0 To lFeatureLayerAndSelSetLINE.Count - 1

                ' get the line feature layer
                Try
                    pFeatureLayerLine = lFeatureLayerAndSelSetLINE(j).pFeatureLayer
                Catch ex As Exception
                    MsgBox("Trouble getting the polygon feature layer back from the object list in 'intersect features' sub. " + ex.Message)
                    Exit Sub
                End Try

                ' get the line selection set from the list
                Try
                    pSelectionSetLine = lFeatureLayerAndSelSetLINE(j).pSelectionSet
                Catch ex As Exception
                    MsgBox("Trouble getting the line selection set back from the object list in 'intersect features' sub. " + ex.Message)
                    Exit Sub
                End Try

                ' get a cursor for the line selection set
                pLineCursor = Nothing 'reset in case
                Try
                    pSelectionSetLine.Search(Nothing, True, pLineCursor)
                Catch ex As Exception
                    MsgBox("Trouble getting featurecursor back from line feature layer selection set in 'intersect features' sub. " + ex.Message _
                           + ". This could be caused by an iCursor being returned instead of an iFeatureCursor.")
                    Exit Sub
                End Try

                If pLineCursor Is Nothing Then
                    MsgBox("Trouble getting featurecursor back from line feature layer selection set in 'intersect features' sub. " _
                           + "An empty cursor was returned.")
                    Exit Sub
                End If

                pSpatialFilter = New SpatialFilter
                pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                pLineFeature = pLineCursor.NextFeature
                Do Until pLineFeature Is Nothing
                    ' Define geometry of spatial filter
                    pSpatialFilter.Geometry = pLineFeature.Shape
                    Try
                        pFeatureSelectionPOLY.SelectFeatures(pSpatialFilter, esriSelectionResultEnum.esriSelectionResultAdd, False)
                    Catch ex As Exception
                        MsgBox("Error encountered executing spatial selection for polygon features. " + ex.Message)
                        Exit Sub
                    End Try
                    pLineFeature = pLineCursor.NextFeature
                Loop
            Next ' Lines layer with a selection set

            If pFeatureSelectionPOLY.SelectionSet.Count > 0 Then
                pSelectionSetPOLY = pFeatureSelectionPOLY.SelectionSet
                pSelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(pWorkspacePoly, _
                                                                                   pFeatureLayerPoly, _
                                                                                   iCPERMFieldIndex, _
                                                                                   pSelectionSetPOLY, _
                                                                                   dCPERM)
                m_lSelectAndUpdateFeaturesObject.Add(pSelectAndUpdateFeaturesObject)
            End If
        Next ' polygon layer 



        ' \\\\\\\\\\\\\\\\\\\\\\\\\\ OLD GP INTERSECT \\\\\\\\\\\\\\\\\\\\\\
        ' ======================= 1.0 INTERSECT ===========================
        ' This next section checks each of the polygon layers in the focusmap
        ' and intersects them with any layers in the focusmap that have a selection.

        ' PROCESS LOGIC:
        ' If there are polygon layers to use in this process then continue
        '   For each of the layers in the focusMap
        '     If it's a feature layer then
        '       If it's a polygon
        '         If it's on the 'include' list
        '           For each of the FocusMap layers 
        '             If it's a line layer (because we don't want to repeatedly 
        '             intersect a polygon layer with itself, do we? DO WEEE???)
        '               If there are any selected features 
        '                 If the parameter array is already populated then empty it
        '                 Populate parameter array for intersect process
        '                 Perform intersect - return results as selection

        'Dim pUID As New UID
        '' Get the pUID of the SelectByLayer command
        ''pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

        'Dim pGp As IGeoProcessor
        'pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
        'Dim pParameterArray As IVariantArray
        'Dim pMxDocument As IMxDocument
        'Dim pMap As IMap

        'Dim sFeatureFullPath As String
        'Dim lMaxLayerIndex As Integer
        'Dim pLayer2Intersect As IFeatureLayer
        ''Dim iFieldVal As Integer  ' The field index
        'Dim sTestPolygon As String
        'Dim bIncludePoly As Boolean 'For polygon inclusion
        'Dim pFeatureLayer As IFeatureLayer
        ''Dim pEnumLayer As IEnumLayer
        'Dim pFeatureSelection As IFeatureSelection
        'Dim pGPResults As IGeoProcessorResult
        'Dim pDoc As IDocument = My.ArcMap.Application.Document
        'pMxDocument = CType(pDoc, IMxDocument)
        'pMap = pMxDocument.FocusMap
        'lMaxLayerIndex = pMap.LayerCount - 1
        'i = 0
        'm = 0
        'Dim pFeatureSelection2 As IFeatureSelection ' for checking if any were returned

        ''pUID = New UID
        ''pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

        ''pEnumLayer = pMap.Layers(pUID, True)
        ''pEnumLayer.Reset()

        'If Not lPolyLayersFieldsCHECKED Is Nothing Then
        '    If lPolyLayersFieldsCHECKED.Count > 0 Then
        '        For i = 0 To lMaxLayerIndex
        '            If pMap.Layer(i).Valid = True Then
        '                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
        '                    pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
        '                    sTestPolygon = Convert.ToString(pFeatureLayer.FeatureClass.ShapeType)

        '                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
        '                        bIncludePoly = False ' Reset Variable
        '                        For j = 0 To lPolyLayersFieldsCHECKED.Count - 1
        '                            If pFeatureLayer.Name = lPolyLayersFieldsCHECKED(j).Layer Then
        '                                bIncludePoly = True
        '                            End If
        '                        Next

        '                        'sDataSourceType = pFeatureLayer.DataSourceType

        '                        If bIncludePoly = True Then
        '                            m = 0
        '                            For m = 0 To lMaxLayerIndex
        '                                If pMap.Layer(m).Valid = True Then
        '                                    If TypeOf pMap.Layer(m) Is IFeatureLayer Then
        '                                        pLayer2Intersect = CType(pMap.Layer(m), IFeatureLayer)

        '                                        If pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
        '                                        Or pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
        '                                            pFeatureSelection = CType(pLayer2Intersect, IFeatureSelection)

        '                                            If (pFeatureSelection.SelectionSet.Count <> 0) Then

        '                                                If pParameterArray IsNot Nothing Then
        '                                                    pParameterArray.RemoveAll()
        '                                                Else
        '                                                    'pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray 'VB.NET
        '                                                    pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray
        '                                                End If

        '                                                ' The GP doesn't need full path names to feature classes
        '                                                ' it only needs feature layer names as they appear in the TOC
        '                                                sFeatureFullPath = pFeatureLayer.Name
        '                                                pParameterArray.Add(sFeatureFullPath)
        '                                                pParameterArray.Add("INTERSECT")
        '                                                pParameterArray.Add(pLayer2Intersect.Name)
        '                                                pParameterArray.Add("#")
        '                                                pParameterArray.Add("ADD_TO_SELECTION")

        '                                                pGPResults = pGp.Execute("SelectLayerByLocation_management", pParameterArray, Nothing)

        '                                                'pFeatureSelection2 = CType(pFeatureLayer, IFeatureSelection)
        '                                                'MsgBox("The count of the features now selected in the polygon layer " + pFeatureLayer.Name + ": " + Convert.ToString(pFeatureSelection2.SelectionSet.Count))

        '                                            End If ' it has a feature selection
        '                                        End If ' It's a line
        '                                    End If ' It's a feature layer
        '                                End If
        '                            Next
        '                        End If
        '                    End If
        '                End If
        '            End If ' Layer is valid
        '        Next
        '    End If ' polygon list count is not zero
        'End If ' polygon list is not nothing

    End Sub

    Private Sub PreUpdateFeaturesWithCPERM(ByVal dCPERM As Double, _
                                        ByVal lLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                        ByRef pEnumLayer As IEnumLayer)

        Dim pFeatureSelection As IFeatureSelection
        Dim pSelectionSet As ISelectionSet
        Dim iFieldIndex As Integer
        Dim pTable As ITable
        Dim pFields As IFields
        Dim pFeatureLayer As IFeatureLayer
        Dim i, j, k, m As Integer
        Dim pWorkspaceA As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim pSelectionSetA As ISelectionSet
        Dim pSelectAndUpdateFeaturesObject As SelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(Nothing, _
                                                                                                                                     Nothing, _
                                                                                                                                     Nothing, _
                                                                                                                                     Nothing, _
                                                                                                                                     Nothing)
        ' Created by Greig Oldford
        ' August 25, 2012
        ' purpose - get selection set for the checked layers and store them in a master list for later.
        ' Later these selectionsets will be used to update the layer with CPERM.  This method groups all workspaces
        ' together, so multiple edit sessions per workspace aren't required, and significantly speeds up processing time. 

        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref
                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then

                    pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                    pSelectionSet = pFeatureSelection.SelectionSet

                    pFields = pFeatureLayer.FeatureClass.Fields
                    k = 0
                    For k = 0 To lLayersFieldsCHECKED.Count - 1
                        If lLayersFieldsCHECKED(k).Layer = pFeatureLayer.Name Then
                            'MsgBox("Updating CPERM Field: The selection set count for layer " + pFeatureLayer.Name _
                            '+ " is " + Convert.ToString(pFeatureSelection.SelectionSet.Count))
                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                ' pFeatureLayer and Workspace
                                ' List Object of Custom Object: pworkspace, pfeaturelayer, dPermField, pSelectionSet (for each workspace start editing)
                                ' Can Predicate Search to get refined list of workspaces later

                                pWorkspaceA = pFeatureLayer.FeatureClass.FeatureDataset.Workspace
                                pSelectionSetA = pFeatureSelection.SelectionSet
                                ' Get the field from the table to calculate
                                pTable = CType(pFeatureLayer.FeatureClass, ITable)
                                iFieldIndex = pTable.FindField(lLayersFieldsCHECKED(k).CumPermField)
                                If iFieldIndex = -1 Then
                                    MsgBox("There must be a field named " & lLayersFieldsCHECKED(k).CumPermField & " in layer " & pFeatureLayer.Name)
                                    Exit Sub
                                End If

                                pSelectAndUpdateFeaturesObject = New SelectAndUpdateFeaturesObject(pWorkspaceA, _
                                                                                                   pFeatureLayer, _
                                                                                                   iFieldIndex, _
                                                                                                   pSelectionSetA, _
                                                                                                   dCPERM)
                                m_lSelectAndUpdateFeaturesObject.Add(pSelectAndUpdateFeaturesObject)

                            End If ' bError is false
                        End If
                    Next
                End If
            End If
            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Loop
        ' =============== END UPDATE FEATURES WITH CUMULATIVE PERMEABILITY ============

    End Sub
    Private Sub ExcludeFeatures(ByVal pFeatureLayer As IFeatureLayer)


        ' Function:    ExcludeFeatures
        ' Created By:  Greig Oldford
        ' Date:        June 11, 2009
        ' Purpose:     To remove from selection any features which are to be
        '              excluded.
        ' Description: This function is called after the intersecting any polygons
        '              included for habitat statistics with returned trace features.
        '              It is meant to remove any selected features that are on the
        '              'excludes' list. It is also meant to remove selected features
        '              of feature classes that are not on the 'includes' list.
        ' Notes:
        '      June 30, 2009  --> This has been modified to read exclusions from extension
        '                         rather than retrieving from global variable. 
        '       Feb 29, 2008  --> The exludes list does not need to be divided into polygons
        '               and lines at all. However, it helps when removing excludes.
        '                  This function should also remove even initially returned
        '               line features since the user may opt only to output habitat
        '               statistics for polygons.


        ' =================== READ EXCLUSIONS FROM EXTENSION =================

        Dim iExclusions, j As Integer
        Dim lExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
        If m_FiPEx.m_bLoaded = True Then ' If there were any extension settings set

            iExclusions = Convert.ToInt32(m_FiPEx.pPropset.GetProperty("numExclusions"))
            Dim ExclusionsObj As New LayerToExclude(Nothing, Nothing, Nothing)

            ' match any of the line layers saved in stream to those in listboxes
            If iExclusions > 0 Then
                For j = 0 To iExclusions - 1
                    'sLineLayer = m_FIPEX.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    ExclusionsObj = New LayerToExclude(Nothing, Nothing, Nothing)
                    With ExclusionsObj
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldLayer" + j.ToString))
                        .Feature = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldFeature" + j.ToString))
                        .Value = Convert.ToString(m_FiPEx.pPropset.GetProperty("ExcldValue" + j.ToString))
                    End With

                    ' add to the module level list
                    lExclusions.Add(ExclusionsObj)
                Next
            End If
        Else
            MsgBox("The FIPEX Options haven't been loaded.  Exiting Exclusions Subroutine.")
            Exit Sub
        End If

        ' ==================== EXCLUDE FEATURES =======================
        ' PROCESS LOGIC: 
        ' 1.0 If there are selected features in the layer
        '   2.0 If there are exclusions
        '     3.0 For each exclusion 
        '       4.0 If exclusion layer matches the current layer
        '       4.1 Get a cursor to loop through selection set
        '         5.0 If the layer name matches current exclusion layer name
        '           6.0 If the exclusion field is found in the layer
        '             7.0 For each feature in the layer
        '             7.1 Get the value from the exclusion list
        '               8.0 If the value from feature matches exclusion value
        '               8.1 Add the object id of feature to a list of items to exclude
        '             7.2 If there are feature values that match found then
        '               9.0 For each value
        '               9.1 Remove that feature from the selection

        Dim pFeatureSelection As IFeatureSelection
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim iFieldVal As Integer
        Dim i As Integer
        Dim aOID As New List(Of Integer)
        Dim iCount As Integer
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pAV As IActiveView
        Dim vVal As Object
        Dim sTempVal As String
        Dim iCountTEMP As Integer
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMap = pMxDoc.FocusMap
        pAV = pMxDoc.ActiveView
        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
        j = 0

        'iCountTEMP = pFeatureSelection.SelectionSet.Count
        'MsgBox("The selected feature count before exclusion: " + iCountTEMP.ToString)

        If (pFeatureSelection.SelectionSet.Count <> 0) Then
            If iExclusions > 0 Then
                i = 0
                For i = 0 To iExclusions - 1
                    Dim pCursor As ICursor
                    pFeatureSelection.SelectionSet.Search(Nothing, True, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)


                    If pFeatureLayer.Name = lExclusions(i).Layer Then
                        ' Find the field
                        iFieldVal = pFeatureCursor.FindField(lExclusions(i).Feature)

                        If iFieldVal <> -1 Then
                            ' reset the excluded features count
                            iCount = 0
                            pFeature = pFeatureCursor.NextFeature

                            Do While Not pFeature Is Nothing
                                Try
                                    vVal = pFeature.Value(iFieldVal)
                                Catch ex As Exception
                                    MsgBox("Could not convert value found in Exclusions Sub to type 'variant/object'. " + _
                                           ". Found in Layer " + pFeatureLayer.Name.ToString + ex.Message)
                                    vVal = "nothing"
                                End Try
                                Try
                                    sTempVal = vVal.ToString
                                Catch ex As Exception
                                    MsgBox("Could not convert value found in Exclusions Sub to type 'string'. " + _
                                           ". Found in Layer " + pFeatureLayer.Name.ToString + ex.Message)
                                    sTempVal = "nothing"
                                End Try

                                'MsgBox(" m_aExcldVls(x): " + CStr(m_aExcldVls(x)))
                                'MsgBox("Is vVal Null? " + CStr(IsNull(vVal)))
                                'MsgBox("What kind of layer is feature from? " + CStr(pFeatureLayer.FeatureClass.ShapeType))
                                '

                                ' If the value matches then add the OID to an array
                                ' to subtract from the selection set.
                                If sTempVal IsNot Nothing Then
                                    If sTempVal = lExclusions(i).Value Then
                                        aOID.Add(pFeature.OID)
                                    End If
                                End If

                                pFeature = pFeatureCursor.NextFeature
                            Loop

                            iCount = aOID.Count
                            Dim iOID As Integer
                            If iCount <> 0 Then
                                For j = 0 To iCount - 1
                                    ' Remove excluded features from selection set by OID
                                    ' NOTE: remove list does not actually remove a list
                                    ' so we need a loop here; see http://forums.esri.com/Thread.asp?c=159&f=1707&t=224394
                                    iOID = aOID.Item(j)
                                    pFeatureSelection.SelectionSet.RemoveList(1, iOID)
                                    pFeatureSelection.SelectionChanged()
                                Next
                                pAV.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                                pAV.Refresh()

                            End If


                        End If ' the exclusion field is found
                    End If     ' if the layer is one with exclusions
                Next    ' exclusion


                'pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                'iCountTEMP = pFeatureSelection.SelectionSet.Count
                'MsgBox("The selected feature count after exclusion: " + iCountTEMP.ToString)

            End If      ' there are any exclusions
        End If          ' there is a selection set for this feature
    End Sub

    Private Sub DoIterativeUpstreamAnalysis(ByRef pMap As IMap, _
                                            ByVal lEID_CPERM_andDir As List(Of EIDCPermAndDir), _
                                            ByRef lGLPKOptionsObject As List(Of GLPKOptionsObject), _
                                            ByRef lLineLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                            ByRef lPolyLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                            ByVal pBarrierSymbol As ISimpleMarkerSymbol, _
                                            ByVal pFlagSymbol As ISimpleMarkerSymbol, _
                                            ByRef pActiveView As IActiveView)
        ' ================== PERFORM UPSTREAM TRACES =================
        ' Loop through barriers
        ' For each barrier encountered upstream (upstreambars; 'start' barrier in order loop zero)
        '   Return upstream lines and polygons
        '   Save them in the master list of lines and polygons
        '   Update lines and polygons with cumulative permeability (CPERM)
        '   Get downstream tracestopping elements
        '   Filter sources / ends out
        '   For each returned tracestopping barrier
        '     Get permeability (PERM) value for option
        '     CPERM = PERM x CPERM
        '     Add to ListForNextUpstream  
        '     Add them to master list of lines and polys
        '   Next tracestopping barrier
        ' Next upstreambars
        ' upstreambars = ListForNextupstream
        ' If upstreambars.count = 0 Then Exit For
        ' Next Orderloop
        '  

        ' ---------------- 1. Calculate Cumulative Permeability for Central Barrier / Sink and Upstream -----------------
        ' Set central barrier or sink as 'start'
        ' set list and object to hold EID, CPERM, and direction for each 
        ' trace EID encountered.  
        ' we have all EIDs non-cumulative permeability stored in the lGLPKOptionsObject
        ' now declare list to hold cumulative permeabilities

        Dim pNextTraceBarrierEIDGEN As IEnumNetEIDBuilderGEN
        pNextTraceBarrierEIDGEN = New EnumNetEIDArray
        Dim pNextTraceBarrierEIDs As IEnumNetEID
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim bFlagNode As Boolean = False
        Dim eFlowElements As esriFlowElements
        Dim pEnumNetEIDBuilder As ESRI.ArcGIS.Geodatabase.IEnumNetEIDBuilder
        Dim pFlowEndJunctionsPer As IEnumNetEID
        Dim pFlowEndEdgesPer As IEnumNetEID ' gets reset after each trace
        Dim pAllFlowEndBarriersGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pAllFlowEndBarriers As IEnumNetEID
        pAllFlowEndBarriersGEN = New EnumNetEIDArray
        pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID) ' doesn't get reset
        Dim pNoSourceFlowEndsGEN As IEnumNetEIDBuilderGEN ' gets reset each order loop
        Dim pNoSourceFlowEnds As IEnumNetEID
        pNoSourceFlowEndsGEN = New EnumNetEIDArray
        pNoSourceFlowEnds = CType(pNoSourceFlowEndsGEN, IEnumNetEID)
        Dim pResultEdges As IEnumNetEID
        Dim pResultJunctions As IEnumNetEID
        Dim lEID_CPERM_andDirNEXT As List(Of EIDCPermAndDir) = New List(Of EIDCPermAndDir) ' gets reset each order loop
        Dim bEID, bFCID, bFID, iEID, bSubID, iFCID, iFID, iSubID As Integer
        Dim dCPERMx As Double = 0
        Dim pNextFlagsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim bKeepEID As Boolean = False
        Dim iEndEID As Integer
        Dim i, j, k, m, p, q As Integer
        Dim dPERMy As Double = 0
        Dim iOrderUpMax As Integer = 10000 ' make this a ridiculous number
        Dim pEID_CPERM_AndDir As EIDCPermAndDir = New EIDCPermAndDir(Nothing, Nothing, Nothing)
        Dim pUID As New UID
        Dim pID As New UID
        Dim pEnumLayer As IEnumLayer
        Dim pFlagDisplay As IFlagDisplay
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim pSymbol As ISymbol
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pGeometricNetwork As IGeometricNetwork
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        Dim pFeatureLayer As IFeatureLayer
        Dim xEID As Integer ' barrier for any given loop
        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)

        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        '' get reference to editor extension
        'Try
        '    pID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"
        '    pEditor = m_app.FindExtensionByCLSID(pID)
        '    If pEditor Is Nothing Then
        '        MsgBox("Error getting reference to the Editor extension. Exiting. ")
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    MsgBox("Error getting reference to the Editor extension. Exiting. " + ex.Message)
        '    Exit Sub
        'End Try

        ' get reference to Map Layer Collection
        Try
            pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
            pEnumLayer = pMap.Layers(pUID, True)
            If pEnumLayer Is Nothing Then
                MsgBox("Error getting reference to the Map Layers Collection. Exiting. ")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Map Layers Collection. Exiting. " + ex.Message)
            Exit Sub
        End Try

        ' Until OrderMax is reached (set to arbitrarily large number because using 'exit for')
        '                           (perhaps should use a do while loop)
        '   For each upstream barrier 
        '      Get the Tracestopping barriers
        '      and get the edges between the barrier and the next upstream barriers
        '      intersect these edges with polygons checked by the user
        '      update the field for the selected lines / polygons with the cumulative permeability
        For i = 0 To iOrderUpMax - 1
            pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)
            lEID_CPERM_andDirNEXT = New List(Of EIDCPermAndDir)
            For j = 0 To lEID_CPERM_andDir.Count - 1

                xEID = lEID_CPERM_andDir.Item(j).Barrier
                dCPERMx = lEID_CPERM_andDir.Item(j).cPerm

                ' make dCPERM = 0 if < 0.05 to avoid
                ' infinite divisions and decimal places
                If dCPERMx < 0.05 Then
                    dCPERMx = 0
                End If

                ' if the dCPERMx is zero, we know all subsequent upstream barriers' permeability will be zero
                ' too. Therefore only set barriers if the cPERMx is not zero.  
                If dCPERMx > 0 Then
                ' create list of EIDs to set as barriers (omit the flag node)

                    pNextTraceBarrierEIDGEN = New EnumNetEIDArray
                    For k = 0 To m_lEIDs.Count - 1
                        ' If it's not the flag barrier
                        If m_lEIDs.Item(k) <> xEID Then
                            pNextTraceBarrierEIDGEN.Add(m_lEIDs.Item(k))
                        End If
                    Next
                    'QI to get 'next' and 'count'
                    pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)
                    ' ========================== SET BARRIERS  ===============================
                    m = 0
                    pNextTraceBarrierEIDs.Reset()
                    Try
                        For m = 0 To pNextTraceBarrierEIDs.Count - 1
                            bEID = pNextTraceBarrierEIDs.Next
                            pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

                            ' Display the barriers as a JunctionFlagDisplay type
                            pFlagDisplay = New JunctionFlagDisplay
                            pSymbol = CType(pBarrierSymbol, ISymbol)
                            With pFlagDisplay
                                .FeatureClassID = bFCID
                                .FID = bFID
                                .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                                .Symbol = pSymbol
                            End With

                            ' Add the flags to the logical network
                            pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                            pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
                        Next
                    Catch ex As Exception

                    End Try
                ' ========================== END SET BARRIERS ===========================
                End If ' dCPERMx is greater than zero

                ' ========================== SET FLAG ====================================
                Try
                    pNetElements.QueryIDs(xEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                    ' Display the flags as a JunctionFlagDisplay type
                    pFlagDisplay = New JunctionFlagDisplay
                    pSymbol = CType(pFlagSymbol, ISymbol)
                    With pFlagDisplay
                        .FeatureClassID = iFCID
                        .FID = iFID
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(xEID)
                        .Symbol = pSymbol
                    End With

                    ' Add the flags to the logical network
                    pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                    pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)
                Catch ex As Exception
                    MsgBox("Trouble setting the Flag " + Convert.ToString(xEID) + ex.Message)
                    Exit Sub
                End Try
                ' ========================== END SET FLAG ===========================

                'prepare the network solver
                Try
                    pTraceFlowSolver = TraceFlowSolverSetup()
                Catch ex As Exception
                    MsgBox("Failed to set up network. Exiting." + ex.Message)
                    Exit Sub
                End Try
                If pTraceFlowSolver Is Nothing Then
                    MsgBox("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver setup error.")
                    Exit Sub
                End If

                eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
                If eFlowElements = Nothing Then
                    MsgBox("Trouble getting the flow elements for the network. Exiting.")
                    Exit Sub
                End If

                ' =============== RUN TRACES TO SELECT UPSTREAM ELEMENTS ============
                pResultEdges = Nothing
                pResultJunctions = Nothing
                pMap.ClearSelection()
                Try
                    pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)
                Catch ex As Exception
                    MsgBox("Trouble finding flow elements upstream of Barrier " + Convert.ToString(xEID) + ex.Message)
                    Exit Sub
                End Try

                If pResultJunctions Is Nothing Then
                    ' junctions were not returned -- create an empty enumeration
                    pEnumNetEIDBuilder = New EnumNetEIDArray
                    pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                    pEnumNetEIDBuilder.ElementType = esriElementType.esriETJunction
                    pResultJunctions = CType(pEnumNetEIDBuilder, IEnumNetEID)
                End If

                If pResultEdges Is Nothing Then
                    ' edges were not returned -- create an empty enumeration
                    pEnumNetEIDBuilder = New EnumNetEIDArray
                    pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                    pEnumNetEIDBuilder.ElementType = esriElementType.esriETEdge
                    pResultEdges = CType(pEnumNetEIDBuilder, IEnumNetEID)
                End If

                Try
                    pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
                Catch ex As Exception
                    MsgBox("Error creating selection for returned results for barrier " + Convert.ToString(xEID) _
                           + ". " + ex.Message)
                    Exit Sub
                End Try
                pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                pActiveView.ScreenDisplay.UpdateWindow()
                'pActiveView.Refresh() ' refresh the view

                ' =============== END RUN TRACES TO SELECT UPSTREAM ELEMENTS ============
                

                ' =============== UPDATE FEATURES WITH CUMULATIVE PERMEABILITY ============
                ' Update selected features with Cumulative permeability
                ' Need
                ' CPERM - lEID_CPERM_andDir.Item(j).cperm
                ' layers with cperm fields - lLineLayersFieldsCHECKED and lPolyLayersFieldsCHECKED
                '                          (these are type LayersAndFCIDAndCumulativePassField)


                ' WHEN EDITS HAPPEN IN EDIT SESSION AS THEY DO IN UPDATEFEATURESWITHPERM
                ' THE MAP SELECTION IS CLEARED BY THE EDITOR WHEN THE SESSION IS STOPPED
                ' note: edit session must be started when editing geometric network features
                ' solution: first call polygons updater
                '           then re-created selection with returned features from trace (lines)
                '           then call lines updater
                'If lPolyLayersFieldsCHECKED.Count > 0 Then
                '    UpdateFeaturesWithCPERM(dCPERMx, lPolyLayersFieldsCHECKED, pEditor, pEnumLayer)
                '    pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
                'End If

                '' GOING TO USE A SINGLE SPATIAL SELECT RATHER THAN AN INTERSECT AT THE END
                '' OF THE ANALYSIS. CAN'T USE EXCLUSIONS BECAUSE OF THIS. 
                ' new intersect function below - adds to master list of selectionsets with workspace
                ' much like the preupdatefeatureswithCPERM below
                If lPolyLayersFieldsCHECKED.Count > 0 Then
                    Call IntersectFeatures(lPolyLayersFieldsCHECKED, dCPERMx, pEnumLayer)
                End If

                '' ---- EXCLUDE FEATURES -----
                'PreExcludeFEatures(pEnumLayer, lPolyLayersFieldsCHECKED, lLineLayersFieldsCHECKED)
                ' '' ---- END EXCLUDE FEATURES -----
                If lLineLayersFieldsCHECKED.Count > 0 Then
                    PreUpdateFeaturesWithCPERM(dCPERMx, lLineLayersFieldsCHECKED, pEnumLayer)
                End If

                ' =========== RUN TRACE UPSTREAM TO GET FLOW STOPPING ELEMENTS =======
                Try
                    pFlowEndJunctionsPer = Nothing
                    pFlowEndEdgesPer = Nothing
                Catch ex As Exception
                    MsgBox("Trouble setting FlowEndJunctions as 'nothing.' Weird. Exiting. " + ex.Message)
                    Exit Sub
                End Try

                'Return the features stopping the trace
                Try
                    pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMUpstream, eFlowElements, pFlowEndJunctionsPer, pFlowEndEdgesPer)
                Catch ex As Exception
                    MsgBox("Error. Problem encountered during 'find flow end elements' of analysis. Exiting. " + ex.Message)
                    Exit Sub
                End Try
                ' ================ END RUN TRACE FOR FLOW STOPPING ELEMENTS ==================

                ' =========== FILTER STOPPING ELEMENTS - NO EIDS NOT ON MASTER LIST =======
                ' eleminate flow stopping elements not on master list derived from connectivity table
                ' this is mainly meant to eliminate network sources / sinks
                ' to eliminate sources, keep a list of all Flow stopping features already encountered
                ' and cross reference this to avoid infinite loops (in the case where barriers have been placed on sources)
                If pFlowEndJunctionsPer.Count > 0 Then

                    pFlowEndJunctionsPer.Reset()
                    k = 0
                    For k = 0 To pFlowEndJunctionsPer.Count - 1

                        bKeepEID = False
                        iEndEID = pFlowEndJunctionsPer.Next
                        m = 0

                        ' loop through the master EID list
                        ' keep the EID only if 
                        '  a) they are on the master list
                        '  b) they have not been encountered before 
                        For m = 0 To m_lEIDs.Count - 1
                            If iEndEID = m_lEIDs(m) Then
                                ' if the junction isn't the same as the flag
                                If iEndEID <> xEID Then
                                    bKeepEID = True
                                End If
                                ' if it's already been returned from previous order loops then
                                ' eliminate it to avoid infinite loops - probably a source / sink
                                pAllFlowEndBarriers.Reset()
                                p = 0
                                For p = 0 To pAllFlowEndBarriers.Count - 1
                                    If iEndEID = pAllFlowEndBarriers.Next Then
                                        bKeepEID = False ' set false if already on master list
                                    End If
                                Next ' p
                            End If
                        Next ' in the master list of eids (m)

                        If bKeepEID = True Then
                            ' This variable does not get reset - used
                            ' to crosscheck in case of infinite loop problem
                            pAllFlowEndBarriersGEN.Add(iEndEID)
                            'This variable gets reset each order loop
                            pNoSourceFlowEndsGEN.Add(iEndEID)

                            ' ------------- GET CUMULATIVE PERMEABILITIES FOR FLOW STOPPING ELEMENTS -----------
                            pEID_CPERM_AndDir = New EIDCPermAndDir(Nothing, Nothing, Nothing)
                            pEID_CPERM_AndDir.Barrier = iEndEID

                            'If iEndEID = 5776 Then
                            '    MsgBox("Encountered Pockwock Lake Dam.")
                            'ElseIf iEndEID = 11760 Then
                            '    MsgBox("Encountered Impassable Channel.")
                            'End If

                            m = 0
                            For m = 0 To lGLPKOptionsObject.Count - 1
                                If lGLPKOptionsObject.Item(m).BarrierEID = iEndEID Then
                                    dPERMy = lGLPKOptionsObject.Item(m).BarrierPerm
                                    pEID_CPERM_AndDir.cPerm = dPERMy * dCPERMx
                                    Exit For
                                End If
                            Next
                            lEID_CPERM_andDirNEXT.Add(pEID_CPERM_AndDir)
                        End If
                    Next 'flowend element k
                    ' ================ END FILTER FLOW END ELEMENTS ==================

                End If ' there are barriers returned upstream stopping trace. 

                pActiveView.Refresh() ' refresh the view
                pNetworkAnalysisExtBarriers.ClearBarriers()
                pNetworkAnalysisExtFlags.ClearFlags()

            Next ' flags in set for this orderloop j in lEID_CPERM_andDir

            ' set the next lEID_CPERM_andDir as the filtered traceflow stopping barriers
            ' lEID_CPERM_andDir = pNoSourceFlowEnds
            lEID_CPERM_andDir = lEID_CPERM_andDirNEXT
            lEID_CPERM_andDirNEXT = New List(Of EIDCPermAndDir)

            'If the count of lEID_CPERM_andDir is zero then exit orderloop
            If lEID_CPERM_andDir.Count = 0 Then Exit For
        Next     ' orderloop (i)

    End Sub

    Private Sub DoIterativeDownstreamAnalysis(ByRef pMap As IMap, _
                                            ByVal pEID_CPERM_andDirDOWN As EIDCPermAndDir, _
                                            ByRef lGLPKOptionsObject As List(Of GLPKOptionsObject), _
                                            ByRef lLineLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                            ByRef lPolyLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                            ByVal pBarrierSymbol As ISimpleMarkerSymbol, _
                                            ByVal pFlagSymbol As ISimpleMarkerSymbol, _
                                            ByRef pActiveView As IActiveView, _
                                            ByVal iSinkEID As Integer)
        ' ==================== Perform DOWNSTREAM ANALYSIS ================
        ' 
        ' Logic:
        ' At barrier X get downstream barrier Y
        '   Set Flag on Y 
        '   Get upstream(edges)
        '   CPERM for edges = CPERMx
        '   Get flowstopping junctions upstream 
        '   Filter them (sources **OR X**)
        '   For each barrier upstream of Y = for each Y-up
        '   Get CPERMyup = CPERMx * PERMyup
        '   Perform upstream analysis for each (call DoIterativeUpstreamAnalysis sub_

        ' Now the permeability at the central node / sink 
        ' is considered 1

        Dim pNextTraceBarrierEIDGEN As IEnumNetEIDBuilderGEN
        pNextTraceBarrierEIDGEN = New EnumNetEIDArray
        Dim pNextTraceBarrierEIDs As IEnumNetEID
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim bFlagNode As Boolean = False
        Dim eFlowElements As esriFlowElements
        Dim pEnumNetEIDBuilder As ESRI.ArcGIS.Geodatabase.IEnumNetEIDBuilder

        Dim pFlowEndJunctionsUP As IEnumNetEID
        Dim pFlowEndEdgesUP As IEnumNetEID ' gets reset after each trace
        Dim pFlowEndJunctionsDOWN As IEnumNetEID
        Dim pFlowEndEdgesDOWN As IEnumNetEID ' gets reset after each trace
        Dim pAllFlowEndBarriersGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pAllFlowEndBarriers As IEnumNetEID
        pAllFlowEndBarriersGEN = New EnumNetEIDArray
        pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID) ' doesn't get reset
        Dim pNoSourceFlowEndsGEN As IEnumNetEIDBuilderGEN ' gets reset each order loop
        Dim pNoSourceFlowEnds As IEnumNetEID
        pNoSourceFlowEndsGEN = New EnumNetEIDArray
        pNoSourceFlowEnds = CType(pNoSourceFlowEndsGEN, IEnumNetEID)
        Dim pResultEdges As IEnumNetEID
        Dim pResultJunctions As IEnumNetEID
        Dim lEID_CPERM_andDirNEXT As List(Of EIDCPermAndDir) = New List(Of EIDCPermAndDir) ' gets reset each order loop
        Dim bEID, bFCID, bFID, iEID, bSubID, iFCID, iFID, iSubID As Integer
        Dim dCPERMx, dPERMy, dPERMyup As Double
        Dim pNextFlagsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim bContinue As Boolean = False
        Dim bKeepEID As Boolean = False
        Dim iEIDy, iEIDx, iEIDyup As Integer
        Dim i, j, k, m, p, q As Integer
        Dim dPerm As Double = 0
        Dim iOrderUpMax As Integer = 10000 ' make this a ridiculous number
        Dim pEID_CPERM_AndDir As EIDCPermAndDir = New EIDCPermAndDir(Nothing, Nothing, Nothing)
        Dim pUID As New UID
        Dim pEnumLayer As IEnumLayer
        Dim pFlagDisplay As IFlagDisplay
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim pSymbol As ISymbol
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pGeometricNetwork As IGeometricNetwork
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements
        Dim pEID_CPERM_andDirUP As EIDCPermAndDir = New EIDCPermAndDir(Nothing, Nothing, Nothing)
        Dim lEID_CPERM_andDirUP As List(Of EIDCPermAndDir) = New List(Of EIDCPermAndDir)
        Dim bContinueUp As Boolean = True
        Dim pFeatureLayer As IFeatureLayer

        Dim iOrderDownMax As Integer = 10000
        Dim iSinkTraceCounter As Integer = 0

        pNetworkAnalysisExtBarriers = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)
        pNetworkAnalysisExtFlags = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExt = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExt)
        pNetworkAnalysisExtResults = CType(m_UtilityNetworkAnalysisExt, INetworkAnalysisExtResults)

        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        ' get reference to Map Layer Collection
        Try
            pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
            pEnumLayer = pMap.Layers(pUID, True)
            If pEnumLayer Is Nothing Then
                MsgBox("Error getting reference to the Map Layers Collection. Exiting. ")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Error getting reference to the Map Layers Collection. Exiting. " + ex.Message)
            Exit Sub
        End Try

        For i = 0 To iOrderDownMax
            iEIDx = pEID_CPERM_andDirDOWN.Barrier
            dCPERMx = pEID_CPERM_andDirDOWN.cPerm
            pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)

            ' This prevents infinite decimal places by multiplying through
            ' chose 0.05 because >5% sensitivity / accuracy threshold unlikely 
            If dCPERMx < 0.05 Then
                dCPERMx = 0
            End If

            ' if the dCPERMx is zero, we know all subsequent downstream barriers' permeability will be zero
            ' too. Therefore only set barriers if the cPERMx is not zero. If cPERMx IS zero then only set
            ' We know the immediately downstream barrier is the sink.  Place flag on the sink, and a single barrier
            ' on X and finish analysis.  
            If dCPERMx > 0 Then

                ' for each downstream barrier (should always only be one)
                ' create list of EIDs to set as barriers (omit the flag node)
                pNextTraceBarrierEIDGEN = New EnumNetEIDArray
                For k = 0 To m_lEIDs.Count - 1
                    ' If it's not the flag barrier
                    If m_lEIDs.Item(k) <> iEIDx Then
                        pNextTraceBarrierEIDGEN.Add(m_lEIDs.Item(k))
                    End If
                Next

                'QI to get 'next' and 'count'
                pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)

                ' ========================== SET BARRIERS  ===============================
                m = 0
                pNextTraceBarrierEIDs.Reset()
                Try
                    For m = 0 To pNextTraceBarrierEIDs.Count - 1
                        bEID = pNextTraceBarrierEIDs.Next
                        pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

                        ' Display the barriers as a JunctionFlagDisplay type
                        pFlagDisplay = New JunctionFlagDisplay
                        pSymbol = CType(pBarrierSymbol, ISymbol)
                        With pFlagDisplay
                            .FeatureClassID = bFCID
                            .FID = bFID
                            .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                            .Symbol = pSymbol
                        End With

                        ' Add the flags to the logical network
                        pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                        pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
                    Next
                Catch ex As Exception

                End Try
                ' ========================== END SET BARRIERS ===========================


                ' ========================== SET FLAG ====================================
                Try
                    pNetElements.QueryIDs(iEIDx, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                    ' Display the flags as a JunctionFlagDisplay type
                    pFlagDisplay = New JunctionFlagDisplay
                    pSymbol = CType(pFlagSymbol, ISymbol)
                    With pFlagDisplay
                        .FeatureClassID = iFCID
                        .FID = iFID
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEIDx)
                        .Symbol = pSymbol
                    End With

                    ' Add the flags to the logical network
                    pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                    pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)
                Catch ex As Exception
                    MsgBox("Trouble setting the Flag " + Convert.ToString(iEIDx) + ex.Message)
                    Exit Sub
                End Try
                ' ========================== END SET FLAG ===========================

                'prepare the network solver
                Try
                    pTraceFlowSolver = TraceFlowSolverSetup()
                Catch ex As Exception
                    MsgBox("Failed to set up network. Exiting." + ex.Message)
                    Exit Sub
                End Try
                If pTraceFlowSolver Is Nothing Then
                    MsgBox("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver setup error.")
                    Exit Sub
                End If

                eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
                If eFlowElements = Nothing Then
                    MsgBox("Trouble getting the flow elements for the network. Exiting.")
                    Exit Sub
                End If


                ' =========== RUN TRACE DOWNSTREAM TO GET FLOW STOPPING DOWNSTREAM BARRIER =======
                Try
                    pFlowEndJunctionsDOWN = Nothing
                    pFlowEndEdgesDOWN = Nothing
                Catch ex As Exception
                    MsgBox("Trouble setting FlowEndJunctionsDOWN as 'nothing.' Weird. Exiting. " + ex.Message)
                    Exit Sub
                End Try

                'Return the features stopping the trace
                Try
                    pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMDownstream, eFlowElements, pFlowEndJunctionsDOWN, pFlowEndEdgesDOWN)
                Catch ex As Exception
                    MsgBox("Error. Problem encountered during 'find flow end elements downstream' of analysis. Exiting. " + ex.Message)
                    Exit Sub
                End Try
                ' ================ END RUN TRACE FOR FLOW STOPPING DOWNSTREAM ELEMENTS ==================

                ' ================ FILTER FLOW STOPPING DOWNSTREAM BARRIER ==============================
                ' make sure  NO EIDS NOT ON MASTER LIST and NOT THE SINK
                ' eliminate flow stopping elements not on master list derived from connectivity table
                ' this is mainly meant to eliminate network sources / sinks
                ' keep a list of all Flow stopping features already encountered
                ' and cross reference this to avoid infinite loops (in the case where barriers have been placed on sources)
                ' or otherwise weird situations that could cause this problem. 
                If pFlowEndJunctionsDOWN.Count > 0 Then

                    If pFlowEndJunctionsDOWN.Count > 1 Then
                        MsgBox("Warning: there are more than one immediately downstream barriers to barrier EID " + iEIDx.ToString _
                               + ". Please check this to continue analysis. Exiting.")
                        Exit Sub
                    End If

                    pFlowEndJunctionsDOWN.Reset()
                    k = 0
                    For k = 0 To pFlowEndJunctionsDOWN.Count - 1

                        bContinue = False
                        iEIDy = pFlowEndJunctionsDOWN.Next
                        m = 0

                        ' loop through the master EID list
                        ' keep the EID only if 
                        '  a) they are on the master list
                        '  b) they have not been encountered before 
                        For m = 0 To m_lEIDs.Count - 1
                            If iEIDy = m_lEIDs(m) Then
                                bContinue = True
                                ' if it's already been returned from previous order loops then
                                ' eliminate it to avoid infinite loops - probably a source / sink
                                pAllFlowEndBarriers.Reset()
                                p = 0
                                For p = 0 To pAllFlowEndBarriers.Count - 1
                                    If iEIDy = pAllFlowEndBarriers.Next Then
                                        bContinue = False ' set false if already on master list
                                    End If
                                Next ' p
                            End If
                        Next ' in the master list of eids (m)

                        If bContinue = True Then
                            ' This variable does not get reset - used
                            ' to crosscheck in case of infinite loop problem
                            pAllFlowEndBarriersGEN.Add(iEIDy)

                            ' ------------- GET CUMULATIVE PERMEABILITIES FOR FLOW STOPPING ELEMENTS -----------
                            m = 0
                            For m = 0 To lGLPKOptionsObject.Count - 1
                                If lGLPKOptionsObject.Item(m).BarrierEID = iEIDy Then
                                    dPERMy = lGLPKOptionsObject.Item(m).BarrierPerm
                                    Exit For
                                End If
                            Next
                        End If
                    Next 'flowend element k (should only be one)
                End If ' there are barriers returned upstream stopping trace. 

                If bContinue = False Then
                    Exit For
                End If
                ' ================ END FILTER FLOW END ELEMENTS ==================
            Else
                iEIDy = iSinkEID
            End If ' dCPERMx Is greater than zero

            pNetworkAnalysisExtBarriers.ClearBarriers()
            pNetworkAnalysisExtFlags.ClearFlags()

            'If iEIDy = 11760 Then
            '    MsgBox("Impassable Channel encountered.")
            'End If
            
            ' ================ SET BARRIERS ON ALL BUT Y ==============================
            
            pNextTraceBarrierEIDGEN = New EnumNetEIDArray
            ' If the permeability is not zero then set all barriers except for at Y
            ' create list of EIDs to set as barriers (omit the flag node)
            If dCPERMx > 0 Then
                For k = 0 To m_lEIDs.Count - 1
                    ' If it's not the flag barrier
                    If m_lEIDs.Item(k) <> iEIDy Then
                        pNextTraceBarrierEIDGEN.Add(m_lEIDs.Item(k))
                    End If
                Next
            Else
                ' if the permeabilty is zero the only barrier needed is at X
                pNextTraceBarrierEIDGEN.Add(iEIDx)
            End If

            'QI to get 'next' and 'count'
            pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)

            m = 0
            pNextTraceBarrierEIDs.Reset()
            Try
                For m = 0 To pNextTraceBarrierEIDs.Count - 1
                    bEID = pNextTraceBarrierEIDs.Next
                    pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

                    ' Display the barriers as a JunctionFlagDisplay type
                    pFlagDisplay = New JunctionFlagDisplay
                    pSymbol = CType(pBarrierSymbol, ISymbol)
                    With pFlagDisplay
                        .FeatureClassID = bFCID
                        .FID = bFID
                        .Geometry = pGeometricNetwork.GeometryForJunctionEID(bEID)
                        .Symbol = pSymbol
                    End With

                    ' Add the flags to the logical network
                    pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                    pNetworkAnalysisExtBarriers.AddJunctionBarrier(pJuncFlagDisplay)
                Next
            Catch ex As Exception
                MsgBox("Trouble setting barriers in the downstream analysis. " + ex.Message)
                Exit Sub
            End Try
            ' ========================== END SET BARRIERS ===========================

            ' ========================== SET FLAG ON Y ====================================
            Try
                pNetElements.QueryIDs(iEIDy, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                ' Display the flags as a JunctionFlagDisplay type
                pFlagDisplay = New JunctionFlagDisplay
                pSymbol = CType(pFlagSymbol, ISymbol)
                With pFlagDisplay
                    .FeatureClassID = iFCID
                    .FID = iFID
                    .Geometry = pGeometricNetwork.GeometryForJunctionEID(iEIDy)
                    .Symbol = pSymbol
                End With

                ' Add the flags to the logical network
                pJuncFlagDisplay = CType(pFlagDisplay, IJunctionFlagDisplay)
                pNetworkAnalysisExtFlags.AddJunctionFlag(pJuncFlagDisplay)
            Catch ex As Exception
                MsgBox("Trouble setting the Flag " + Convert.ToString(iEIDy) + ex.Message)
                Exit Sub
            End Try
            ' ========================== END SET FLAG ON Y===========================

            'prepare the network solver
            Try
                pTraceFlowSolver = TraceFlowSolverSetup()
            Catch ex As Exception
                MsgBox("Failed to set up network. Exiting." + ex.Message)
                Exit Sub
            End Try
            If pTraceFlowSolver Is Nothing Then
                MsgBox("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver setup error.")
                Exit Sub
            End If

            eFlowElements = esriFlowElements.esriFEJunctionsAndEdges
            If eFlowElements = Nothing Then
                MsgBox("Trouble getting the flow elements for the network. Exiting.")
                Exit Sub
            End If

            ' =============== RUN TRACES TO SELECT UPSTREAM ELEMENTS ============
            pResultEdges = Nothing
            pResultJunctions = Nothing
            pMap.ClearSelection()
            Try
                pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)
            Catch ex As Exception
                MsgBox("Trouble finding flow elements upstream of Barrier " + Convert.ToString(iEIDy) + ex.Message)
                Exit Sub
            End Try

            If pResultJunctions Is Nothing Then
                ' junctions were not returned -- create an empty enumeration
                pEnumNetEIDBuilder = New EnumNetEIDArray
                pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                pEnumNetEIDBuilder.ElementType = esriElementType.esriETJunction
                pResultJunctions = CType(pEnumNetEIDBuilder, IEnumNetEID)
            End If

            If pResultEdges Is Nothing Then
                ' edges were not returned -- create an empty enumeration
                pEnumNetEIDBuilder = New EnumNetEIDArray
                pEnumNetEIDBuilder.Network = pNetworkAnalysisExt.CurrentNetwork.Network
                pEnumNetEIDBuilder.ElementType = esriElementType.esriETEdge
                pResultEdges = CType(pEnumNetEIDBuilder, IEnumNetEID)
            End If

            Try
                pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
            Catch ex As Exception
                MsgBox("Error creating selection for returned results for barrier " + Convert.ToString(iEIDy) _
                       + ". " + ex.Message)
                Exit Sub
            End Try
            ' =============== END RUN TRACES TO SELECT UPSTREAM ELEMENTS ============

            ' =============== UPDATE FEATURES WITH CUMULATIVE PERMEABILITY ============
            ' Update selected features with Cumulative permeability
            ' Need
            ' CPERM - lEID_CPERM_andDir.Item(j).cperm
            ' layers with cperm fields - lLineLayersFieldsCHECKED and lPolyLayersFieldsCHECKED
            '                          (these are type LayersAndFCIDAndCumulativePassField)

            ' WHEN EDITS HAPPEN IN EDIT SESSION AS THEY DO IN UPDATEFEATURESWITHPERM
            ' THE MAP SELECTION IS CLEARED BY THE EDITOR WHEN THE SESSION IS STOPPED
            ' note: edit session must be started when editing geometric network features
            ' solution: first call polygons updater
            '           then re-created selection with returned features from trace (lines)
            '           then call lines updater
            'If lPolyLayersFieldsCHECKED.Count > 0 Then
            '    UpdateFeaturesWithCPERM(dCPERMx, lPolyLayersFieldsCHECKED, pEditor, pEnumLayer)
            '    pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
            'End If
            ' =============== INSERSECT AND EXCLUDE FEATURES  ============
            ' NOT GOING TO INTERSECT AFTER EACH TRACE - USE SPATIAL SELECT INSTEAD
            ' CANNOT USE EXCLUSIONS IN THIS CASE
            ' new intersect function below - adds to master list of selectionsets with workspace
            ' much like the preupdatefeatureswithCPERM below
            If lPolyLayersFieldsCHECKED.Count > 0 Then
                Call IntersectFeatures(lPolyLayersFieldsCHECKED, dCPERMx, pEnumLayer)
            End If

            '' ---- EXCLUDE FEATURES -----
            'PreExcludeFEatures(pEnumLayer, lPolyLayersFieldsCHECKED, lLineLayersFieldsCHECKED)
            ' '' ---- END EXCLUDE FEATURES -----

            ' =============== END INSERSECT AND EXCLUDE FEATURES  ============

            If lLineLayersFieldsCHECKED.Count > 0 Then
                PreUpdateFeaturesWithCPERM(dCPERMx, lLineLayersFieldsCHECKED, pEnumLayer)
            End If


            ' =========== RUN TRACE UPSTREAM TO GET FLOW STOPPING ELEMENTS FROM Y =======
            Try
                pFlowEndJunctionsUP = Nothing
                pFlowEndEdgesUP = Nothing
            Catch ex As Exception
                MsgBox("Trouble setting FlowEndJunctions as 'nothing.' Weird. Exiting. " + ex.Message)
                Exit Sub
            End Try

            'Return the features stopping the trace
            Try
                pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMUpstream, eFlowElements, pFlowEndJunctionsUP, pFlowEndEdgesUP)
            Catch ex As Exception
                MsgBox("Error. Problem encountered during 'find flow end elements up of Y' of analysis. Exiting. " + ex.Message)
                Exit Sub
            End Try
            ' ================ END RUN TRACE FOR FLOW STOPPING ELEMENTS FROM Y =================


            ' ================ FILTER FLOW STOPPING UPSTREAM BARRIERS FROM Y ==============================
            ' make sure  NO EIDS NOT ON MASTER LIST 
            ' eliminate flow stopping elements not on master list derived from connectivity table
            ' this is mainly meant to eliminate network sources 
            ' also eliminate X from this list
            lEID_CPERM_andDirUP = New List(Of EIDCPermAndDir)
            bContinueUp = True
            If pFlowEndJunctionsUP.Count > 0 Then

                pFlowEndJunctionsUP.Reset()
                k = 0
                For k = 0 To pFlowEndJunctionsUP.Count - 1

                    bKeepEID = False
                    iEIDyup = pFlowEndJunctionsUP.Next
                    m = 0

                    ' loop through the master EID list
                    ' keep the EID only if 
                    '  a) they are on the master list
                    '  b) it is not X
                    For m = 0 To m_lEIDs.Count - 1
                        If iEIDyup = m_lEIDs(m) Then
                            If iEIDyup <> iEIDx Then
                                bKeepEID = True
                            End If
                        End If
                    Next ' in the master list of eids (m)

                    If bKeepEID = True Then
                        ' This variable does not get reset - used
                        ' to crosscheck in case of infinite loop problem
                        pAllFlowEndBarriersGEN.Add(iEIDyup)
                        'If iEIDyup = 5776 Then
                        '    MsgBox("Pockwock Lake barrier encountered as yup")
                        'End If
                        ' ------------- GET CUMULATIVE PERMEABILITIES FOR FLOW STOPPING ELEMENTS -----------
                        m = 0
                        For m = 0 To lGLPKOptionsObject.Count - 1
                            If lGLPKOptionsObject.Item(m).BarrierEID = iEIDyup Then
                                dPERMyup = lGLPKOptionsObject.Item(m).BarrierPerm
                                Exit For
                            End If
                        Next

                        If iEIDyup <> iSinkEID Then
                            pEID_CPERM_andDirUP = New EIDCPermAndDir(iEIDyup, (dPERMyup * dCPERMx), "up")
                            lEID_CPERM_andDirUP.Add(pEID_CPERM_andDirUP)
                        End If
                    End If
                Next 'flowend element k upstream of Y
            Else
                bContinueUp = False
            End If ' there are barriers returned upstream stopping trace. 

            ' ================ END FILTER FLOW END ELEMENTS ==================

            pNetworkAnalysisExtBarriers.ClearBarriers()
            pNetworkAnalysisExtFlags.ClearFlags()

            ' ================ DO ITERATIVE UPSTREAM ANALYSIS ================== 
            If lEID_CPERM_andDirUP.Count = 0 Then
                bContinueUp = False
            End If

            ' this is strange but the sink seems to be returned as flow-stopping
            ' element when a flag is placed on the sink and a flow-stopping trace is done
            ' if this is the sink then exit
            If lEID_CPERM_andDirUP.Count = 1 Then
                If iEIDy = iSinkEID Then
                    Exit For 'orderloop - halt analysis
                End If
            End If

            ' seems to be a problem with repetitive tracing 
            ' upstream of the sink (the sink gets traced up twice, the 
            '  second time ignoring the barrier 'X' upstream and visiting
            '  the entire network)
            '  there should be only ONE trace upstream of the sink EID.  
            If iEIDy = iSinkEID Then
                iSinkTraceCounter += 1
                If iSinkTraceCounter > 1 Then
                    Exit For
                    bContinueUp = False ' just in case
                End If
            End If

            If bContinueUp = True Then

                'If iEIDx = 11760 Then
                '    MsgBox("Impassable Channel enountered.")
                'End If

                DoIterativeUpstreamAnalysis(pMap, _
                                            lEID_CPERM_andDirUP,
                                            lGLPKOptionsObject,
                                            lLineLayersFieldsCHECKED,
                                            lPolyLayersFieldsCHECKED,
                                            pBarrierSymbol,
                                            pFlagSymbol,
                                            pActiveView)
            End If
            ' ================ END DO ITERATIVE UPSTREAM ANALYSIS ================== 

            ' if this is the sink then exit
            If iEIDy = iSinkEID Then
                Exit For
            End If

            pEID_CPERM_andDirDOWN = New EIDCPermAndDir(iEIDy, (dPERMy * dCPERMx), "down")

        Next ' downstream order

    End Sub
    Private Sub PreExcludeFEatures(ByVal pEnumLayer As IEnumLayer, _
                                   ByRef lPolyLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField), _
                                   ByRef lLineLayersFieldsCHECKED As List(Of LayersAndFCIDAndCumulativePassField))


        Dim bAnalysisFeatureMatch As Boolean = False
        Dim pFeatureLayer As IFeatureLayer
        pEnumLayer.Reset()
        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then
                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                    bAnalysisFeatureMatch = False
                    For k = 0 To lPolyLayersFieldsCHECKED.Count - 1
                        If pFeatureLayer.Name = lPolyLayersFieldsCHECKED(k).Layer Then
                            bAnalysisFeatureMatch = True
                            Exit For
                        End If
                    Next
                    If bAnalysisFeatureMatch = False Then
                        For k = 0 To lLineLayersFieldsCHECKED.Count - 1
                            If pFeatureLayer.Name = lPolyLayersFieldsCHECKED(k).Layer Then
                                bAnalysisFeatureMatch = True
                                Exit For
                            End If
                        Next ' k
                    End If
                    If bAnalysisFeatureMatch = True Then
                        ExcludeFeatures(pFeatureLayer)
                    End If
                End If
            End If
            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
        Loop
    End Sub
    Private Sub UpdateLineWithCPERM(ByRef pEditor As ESRI.ArcGIS.Editor.IEditor)

        ' This is an attempt to use a model level list object to store
        ' the selected features from each trace.  This is so an update / edit session
        ' only needs to be opened once per workspace rather than once per featurelayer per trace
        ' Have: m_lSelectAndUpdateFeaturesObject containing the workspaces, layers, selection set, field indices, and cperms
        ' Need: list of the unique workspaces from the above object to loop through
        '       a refined list of the above object to loop through (for each unique workpace)

        Dim lWorkspaces As List(Of ESRI.ArcGIS.Geodatabase.IWorkspace) = New List(Of IWorkspace)
        ' can use findall to get the list of objects from master list containing each unique workspace
        Dim pWorkspace, pWorkspaceChecker As ESRI.ArcGIS.Geodatabase.IWorkspace
        Dim workspacecomparer As FindSelectionSetsByWorkspace 'predicate
        Dim lRefinedSelectAndUpdateFeaturesObject As List(Of SelectAndUpdateFeaturesObject)
        Dim bWorkspaceMatch As Boolean = False
        Dim pFeatureLayer As IFeatureLayer
        Dim iCPermFieldIndex As Integer
        Dim dCPerm As Double

        Dim pSelectionSet2 As ESRI.ArcGIS.Geodatabase.ISelectionSet2
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim bError As Boolean = False
        Dim k As Integer


        'barriercomparer = New FindBarriersBySinkEIDPredicate(iSinkEID)
        'refinedBarrierEIDList = lBarrierAndSinkEIDs.FindAll(AddressOf barriercomparer.CompareEID)

        ' It's a bit convoluted to get unique list of items from master list. 
        ' so here's the way I'll do it:
        ' for each object in the master list 
        ' get the workspace
        ' if the workspace isn't in the existing list of workspaces
        ' then get a refined list from the master list with just that workspace. 
        For i = 0 To m_lSelectAndUpdateFeaturesObject.Count - 1
            bWorkspaceMatch = False
            pWorkspace = m_lSelectAndUpdateFeaturesObject(i).pWorkspace
            For j = 0 To lWorkspaces.Count - 1
                If lWorkspaces(j) Is pWorkspace Then
                    bWorkspaceMatch = True
                End If
            Next

            ' if it's not in the master list yet the updating hasn't been done yet
            If bWorkspaceMatch = False Then

                ' add it to the list
                lWorkspaces.Add(pWorkspace)

                ' get a refined master list for this workspace
                workspacecomparer = New FindSelectionSetsByWorkspace(pWorkspace)
                lRefinedSelectAndUpdateFeaturesObject = New List(Of SelectAndUpdateFeaturesObject)
                lRefinedSelectAndUpdateFeaturesObject = m_lSelectAndUpdateFeaturesObject.FindAll(AddressOf workspacecomparer.CompareWorkspace)

                ' start one edit session for this refined list
                If pEditor.EditState = ESRI.ArcGIS.Editor.esriEditState.esriStateEditing Then
                    Try
                        pEditor.StopEditing(True)
                    Catch ex As Exception
                        MsgBox("Error trying to stop current edit session. Please save your edits and close this session and try again." + ex.Message)
                        Exit Sub
                    End Try
                End If
                Try
                    pEditor.StartEditing(pWorkspace)
                Catch ex As Exception
                    MsgBox("Error trying to start edit session for selected feature's feature class - layer " _
                           + ex.Message)
                End Try

                ' for all selection sets with the same workspace update their tables
                For j = 0 To lRefinedSelectAndUpdateFeaturesObject.Count - 1

                    ' just for safety check the workspaces
                    pWorkspaceChecker = lRefinedSelectAndUpdateFeaturesObject(j).pWorkspace
                    If Not pWorkspaceChecker Is pWorkspace Then
                        MsgBox("Error: the workspace in the refined master selection list isn't the one expected after predicate refining.")
                    End If

                    pEditor.StartOperation()
                    pFeatureCursor = Nothing

                    pSelectionSet2 = CType(lRefinedSelectAndUpdateFeaturesObject(j).pSelectionSet, ISelectionSet2)
                    pFeatureLayer = lRefinedSelectAndUpdateFeaturesObject(j).pFeatureLayer
                    iCPermFieldIndex = lRefinedSelectAndUpdateFeaturesObject(j).iCPermFieldIndex
                    dCPerm = lRefinedSelectAndUpdateFeaturesObject(j).dCPerm

                    Try
                        pSelectionSet2.Update(Nothing, True, pFeatureCursor)
                        pFeature = Nothing
                        pFeature = pFeatureCursor.NextFeature

                        Do Until pFeature Is Nothing
                            pFeature.Value(iCPermFieldIndex) = dCPerm
                            pFeatureCursor.UpdateFeature(pFeature)
                            pFeature = pFeatureCursor.NextFeature
                        Loop

                        pEditor.StopOperation("featureupdated")
                    Catch ex As Exception
                        MsgBox("Error trying to update selected features. " + ex.Message _
                               + ". For layer " + pFeatureLayer.Name)
                    End Try
                Next ' selection object in the refined list

                ' end the edit sessions
                pEditor.StopEditing(True)

            End If
        Next ' i - object in the master list, looking for unique workspaces

        ''If Not CanEditWOEditSession(pFeatureLayer.FeatureClass) Then
        ''If pEditor.EditState = ESRI.ArcGIS.Editor.esriEditState.esriStateNotEditing Then
        ''MsgBox("This sample requires that ArcMap is not in edit mode. Will now activate editor mode.")
        ''End If

        ''End If

        ' '' Also, check to see if the selected layer supports editing without
        ' '' an edit session
        ''If Not CanEditWOEditSession(pFeatureLayer.FeatureClass) Then
        ''    MsgBox("This layer (" & pFeatureLayer.Name & ") cannot be edited outside of an edit session")
        ''Else
        '' lEIDAndMetrics
        '' lFCIDandNameAndDCIMetricFieldObject
        '' iEID
        'If bError = False Then
        '    '' Get the field from the table to calculate
        '    'pTable = CType(pFeatureLayer.FeatureClass, ITable)
        '    'iFieldIndex = pTable.FindField(lLayersFieldsCHECKED(k).CumPermField)
        '    'If iFieldIndex = -1 Then
        '    '    MsgBox("There must be a field named " & lLayersFieldsCHECKED(k).CumPermField & " in layer " & pFeatureLayer.Name)
        '    '    Exit Sub
        '    'End If

        '    pEditor.StartOperation()
        '    pFeatureCursor = Nothing
        '    pSelectionSet2 = CType(pFeatureSelection.SelectionSet, ISelectionSet2)
        '    Try
        '        pSelectionSet2.Update(Nothing, True, pFeatureCursor)
        '        pFeature = Nothing
        '        pFeature = pFeatureCursor.NextFeature

        '        Do Until pFeature Is Nothing
        '            pFeature.Value(iFieldIndex) = dCPERM
        '            pFeatureCursor.UpdateFeature(pFeature)
        '            pFeature = pFeatureCursor.NextFeature
        '        Loop

        '        pEditor.StopOperation("featureupdated")
        '        pEditor.StopEditing(True)
        '    Catch ex As Exception
        '        MsgBox("Error trying to update selected features. " + ex.Message _
        '               + ". For layer " + pFeatureLayer.Name)
        '    End Try

        '    'pEditor.StartOperation()
        '    'Dim pEnumFeature As IEnumFeature
        '    'pEnumFeature = pEditor.EditSelection
        '    'pEnumFeature.Reset()
        '    'pFeature = pEnumFeature.Next

        '    'For m = 0 To pEditor.SelectionCount - 1

        '    '    pFeature.Val                                                ue(iMetricField) = dMetric

        '    '    pFeature = pEnumFeature.Next
        '    'Next
        'Else
        '    MsgBox("No selected features were found for layer " + pFeatureLayer.Name)
        'End If
    End Sub
    Private Class FindSelectionSetsByWorkspace
        ' predicate to retrieve matching workspace from a list of objects
        Private _pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace

        Public Sub New(ByVal pworkspace As ESRI.ArcGIS.Geodatabase.IWorkspace)
            Me._pWorkspace = pworkspace
        End Sub

        Public Function CompareWorkspace(ByVal obj As SelectAndUpdateFeaturesObject) As Boolean
            Return (_pWorkspace Is obj.pWorkspace)
        End Function
    End Class
End Class