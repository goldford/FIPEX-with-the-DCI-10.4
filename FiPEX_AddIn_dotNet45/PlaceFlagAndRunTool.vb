Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.SystemUI
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class PlaceFlagAndRunTool
    Inherits ESRI.ArcGIS.Desktop.AddIns.Tool
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    Private backgroundworker2 As New BackgroundWorker
    Private m_bCancel As Boolean
    Private m_iProgress As Integer
    Private ProgressForm As New frmAnalysisProgress
    Private bToolRunning As Boolean = False ' to pause onUpdate when tool is running

    Public Sub New()

        backgroundworker2.WorkerReportsProgress = True
        backgroundworker2.WorkerSupportsCancellation = True
        AddHandler backgroundworker2.DoWork, AddressOf backgroundworker2_DoWork
        AddHandler backgroundworker2.ProgressChanged, AddressOf backgroundworker2_ProgressChanged
        AddHandler backgroundworker2.RunWorkerCompleted, AddressOf backgroundworker2_RunWorkerCompleted
    End Sub
    Protected Overloads Overrides Sub OnUpdate()
        ' use the extension listener to avoid constant checks to the 
        ' map network.  The extension listener will only update the boolean
        ' check on network count if there's a map change
        ' upgrade at version 10
        ' protected override void OnUpdate()
        If bToolRunning = False Then
            If m_FiPEx__1 Is Nothing Then
                m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
            End If
            If m_UNAExt Is Nothing Then
                m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetUNAExt
            End If
            'Dim FiPEx__1 As FishPassageExtension = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetExtension
            'Dim pUNAExt As IUtilityNetworkAnalysisExt = FiPEX_AddIn_dotNet35_2.FishPassageExtension.GetUNAExt
            If m_pNetworkAnalysisExt Is Nothing Then
                m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
            End If

            If m_pNetworkAnalysisExt.NetworkCount > 0 Then
                Me.Enabled = True
            Else
                Me.Enabled = False
            End If
        End If

    End Sub

    Protected Overrides Sub OnMouseDown(ByVal arg As MouseEventArgs)

        ' Keeping things similar to Analysis command 
        ' so calling subroutine here which will invoke backgroundworker
        Call RunSimpleAnalysis(arg)


    End Sub
    Private Sub backgroundworker2_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        m_bCancel = False
        Dim iProgress As Integer

        If backgroundworker2.CancellationPending = True Then
            e.Cancel = True
        End If

        ProgressForm = New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmAnalysisProgress()
        If ProgressForm.Form_Initialize() Then
            ProgressForm.ShowDialog()
        End If

        m_bCancel = True
        ProgressForm = Nothing
        backgroundworker2.CancelAsync()



    End Sub

    Private Sub backgroundworker2_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        If Not ProgressForm Is Nothing Then
            If m_bCancel = False Then
                Try
                    If ProgressForm.IsDisposed = False Then
                        If ProgressForm.Visible = True Then
                            If ProgressForm.m_bCloseMe = False Then
                                ProgressForm.ChangeProgressBar(e.ProgressPercentage)
                                ProgressForm.ChangeLabel(e.UserState.ToString)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    ' there will be exceptions thrown at this crap
                    'MsgBox("Issue in ProgressChanged")
                    Exit Sub
                End Try
            End If
        End If
    End Sub
    Private Sub backgroundworker2_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)
        Try
            If e.Cancelled = True Then
                m_bCancel = True
                If Not ProgressForm Is Nothing Then
                    If ProgressForm.Visible = True And m_bCancel = False Then
                        ProgressForm.Close()
                    End If
                End If

            ElseIf e.Error IsNot Nothing Then
                If Not ProgressForm Is Nothing Then
                    If ProgressForm.Visible = True And m_bCancel = False Then
                        ProgressForm.lblProgress.Text = "Error: " & e.Error.Message
                    End If
                End If
            Else
                If Not ProgressForm Is Nothing Then
                    'ProgressForm.lblProgress.Text = "Done!"
                    'ProgressForm.cmdCancel.Text = "Close"
                    If ProgressForm.Visible = True And m_bCancel = False Then
                        ProgressForm.Close()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Error in runworkercompleted " + ex.Message)
        End Try

    End Sub
    Private Sub RunSimpleAnalysis(ByVal arg As MouseEventArgs)

        m_bCancel = False
        ProgressForm = Nothing

        ' create a new background worker
        If Not backgroundworker2.IsBusy = True Then
            backgroundworker2.RunWorkerAsync()
        End If


        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Getting FiPEx Option Settings")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(5, "Getting FiPEx Option Settings")

        Threading.Thread.Sleep(200)

        'Change the mouse cursor to hourglass
        Dim pMouseCursor As IMouseCursor
        pMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim BeginTime As DateTime = DateTime.Now
        Dim EndTime As DateTime

        ' Get current network
        ' Clear the results of previous trace tasks
        ' Clear the flags in the network
        If m_FiPEx__1 Is Nothing Then
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetUNAExt
        End If
        If m_pNetworkAnalysisExt Is Nothing Then
            m_pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        End If
        ' Get reference to the current network through Utility Network interface
        Dim pNetworkAnalysisExt As INetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)
        Dim pNetwork As INetwork = pGeometricNetwork.Network
        Dim pNetElements As INetElements = CType(pNetwork, INetElements)
        Dim pTraceTasks As ITraceTasks = CType(m_UNAExt, ITraceTasks)
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults = CType(m_UNAExt, INetworkAnalysisExtResults)

        pNetworkAnalysisExtResults.ClearResults()
        pMap.ClearSelection()


        Dim pFlagSymbol, pBarrierSymbol As ISimpleMarkerSymbol
        Dim pSimpleMarkerSymbol As ISimpleMarkerSymbol = New SimpleMarkerSymbol
        Dim pRgbColor As IRgbColor = New RgbColor

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


        ' Create the form
        ' Prepare the output results form
        'Dim pResultsForm As New FiPEx.frmResults
        'pResultsForm.Show()

        ' convert the EID to a feature class ID, feature ID, and sub ID
        Dim netElements As INetElements = CType(pGeometricNetwork.Network, INetElements)
        Dim FCID As Integer, FID As Integer, subID As Integer

        ' ===============GET CURRENT FLAGS AND BARRIERS ==================
        ' Before all current flags are cleared

        ''MsgBox("Debug:1")

        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Getting FiPEx Option Settings")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(5, "Saving Original FiPEx Settings")


        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN

        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pFlagDisplay As IFlagDisplay
        Dim bEID, i As Integer
        Dim iOrderNum As Integer = 0

        Dim pMetricsObject As New MetricsObject(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim lMetricsObject As New List(Of MetricsObject)
        Dim lAllFCIDs As New List(Of FCIDandNameObject) ' to hold FCIDs of included line layers
        Dim lBarrierAndSinkEIDs As List(Of BarrAndBarrEIDAndSinkEIDs) = New List(Of BarrAndBarrEIDAndSinkEIDs)
        Dim bBarrierPerm As Boolean = False ' Barrier perm field? 
        Dim bNaturalYN As Boolean = False   ' Natural Barrier y/n field?

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

        ' Was using ArrayList because of advantage of 'count' and 'add' properties
        ' but EnumNetEIDBuilderGEN addition to 9.2 has this functionality
        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        Dim iOriginalJunctionBarrierCount As Integer
        iOriginalJunctionBarrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        'MsgBox("Debug Mes: The Original junction barrier count: " + iOriginalJunctionBarrierCount.ToString)

        ' Save the barriers
        For i = 0 To iOriginalJunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            pFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
            'originalBarriersList(i) = bEID
        Next

        ' No edge flag support yet
        Dim pOriginalBarriersList As IEnumNetEID
        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

        ' Save the flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
            'pOriginaljuncFlagsList(i) = bEID
        Next

        Dim pOriginaljuncFlagsList As IEnumNetEID
        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ******** NO EDGE FLAG SUPPORT YET *********
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        Dim pOriginalEdgeFlagsList As IEnumNetEID
        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)
        ' ******************************************
        ' ============== UnLABEL FLAGS ===================
        Dim sFlagCheck As String
        Dim iFlagEID, m As Integer
        Dim flagBarrier As Boolean

        ' Check if flag is on a barrier
        pOriginaljuncFlagsList.Reset()
        For i = 0 To pOriginaljuncFlagsList.Count - 1

            iFlagEID = pOriginaljuncFlagsList.Next
            flagBarrier = False     ' assume flag is not on barrier
            m = 0

            pOriginalBarriersList.Reset()
            For m = 0 To pOriginalBarriersList.Count - 1
                If iFlagEID = pOriginalBarriersList.Next Then
                    flagBarrier = True
                End If
            Next

            ' unlabel if not over barrier
            If flagBarrier = False Then
                netElements.QueryIDs(iFlagEID, esriElementType.esriETJunction, FCID, FID, subID)
                UnLabelBarrier(FCID, FID)
            End If
        Next

        ' ========== GET THE JUNCTION =============

        ''MsgBox("Debug:2")
        Dim pPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        pPoint = pMxApp.Display.DisplayTransformation.ToMapPoint(arg.X, arg.Y)

        ' find the nearest junction element to this Point
        Dim EID As Integer
        Dim outPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
        Dim pointToEID As IPointToEID = New PointToEID
        With pointToEID
            .GeometricNetwork = pGeometricNetwork
            .SourceMap = pMap
            .SnapTolerance = 100     ' set a snap tolerance of 10 map units
            .GetNearestJunction(pPoint, EID, outPoint)
        End With

        If EID = 0 Then
            MsgBox("You have to click on a network junction element.")
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If

        ' convert the EID to a feature class ID, feature ID, and sub ID
        netElements.QueryIDs(EID, esriElementType.esriETJunction, FCID, FID, subID)

        ' =============== CLEAR FLAGS =================
        pNetworkAnalysisExtFlags.ClearFlags()

        ' create a new JunctionFlagDisplay object and populate it
        Dim junctionFlagDisplay As IJunctionFlagDisplay = New JunctionFlagDisplay
        Dim flagDisplay As IFlagDisplay = CType(junctionFlagDisplay, IFlagDisplay)
        With flagDisplay
            .FeatureClassID = FCID
            .FID = FID
            .SubID = subID
            .Geometry = CType(outPoint, IGeometry)
            .Symbol = CType(pFlagSymbol, ISymbol)
        End With

        ' add the JunctionFlagDisplay object to the Network Analysis extension
        Dim networkAnalysisExtFlags As INetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        networkAnalysisExtFlags.AddJunctionFlag(junctionFlagDisplay)

        ''MsgBox("Debug:3")
        ' =============== LABEL THE FLAG =================
        ' 1.2 Label the flag
        Call LabelFlag(FCID, FID)

        ''MsgBox("Debug:4")
        ' refresh the view
        pActiveView.Refresh()


        ' =============== GET THE FLAGS LIST AGAIN ==================
        ' Save the flags
        i = 0

        ' reset these variables. 
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray

        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
            'pOriginaljuncFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginaljuncFlagsList = Nothing
        pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ******** NO EDGE FLAG SUPPORT YET *********
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = Nothing
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)
        ' ******************************************

        ' =============== READ EXTENSION SETTINGS =========

        ''MsgBox("Debug:5")
        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(10, "Reading FiPEx Settings")

        ' 1.3 Get the extension settings needed 
        Dim j, k As Integer

        Dim sDirection As String = "up"     ' Analysis direction default to 'upstream'
        Dim pLLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
        Dim pPLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)
        Dim plExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
        Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        Dim iLinesCount As Integer = 0      ' number of lines layers currently using
        Dim iExclusions As Integer = 0

        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim bUpHab, bTotalUpHab, bDownHab, bTotalDownHab, bPathDownHab, bTotalPathDownHab As Boolean


        Dim lHabStatsList As List(Of StatisticsObject_2) = New List(Of StatisticsObject_2)
        Dim iBarrierIDs As Integer
        Dim sBarrierIDLayer, sBarrierIDField As String
        Dim sBarrierPermField, sBarrierNaturalYNField As String

        ' If settings have been set by the user then load them
        If m_FiPEx__1.m_bLoaded = True Then
            sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))
            iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
            bBarrierPerm = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("barrierperm"))
            bNaturalYN = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("NaturalYN"))
            bUpHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("UpHab"))
            bTotalUpHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalUpHab"))
            bDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("DownHab"))
            bTotalDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalDownHab"))
            bPathDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("PathDownHab"))
            bTotalPathDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalPathDownHab"))

            ' Populate a list of the layers using and habitat summary fields.
            ' match any of the polygon layers saved in stream to those in listboxes 
            Dim HabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

            If iPolysCount > 0 Then
                For k = 0 To iPolysCount - 1
                    'sPolyLayer = m_DFOExt.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
                    HabLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
                    With HabLayerObj
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
                        .HabUnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
                    End With

                    ' Load that object into the list
                    pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
                Next
            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            Dim iCount1 As Integer = pPLayersFields.Count

            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pPLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                Next
            End If

            iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
            Dim HabLayerObj2 As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

            ' match any of the line layers saved in stream to those in listboxes
            If iLinesCount > 0 Then
                For j = 0 To iLinesCount - 1
                    'sLineLayer = m_DFOExt.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    HabLayerObj2 = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    With HabLayerObj2
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))

                        .LengthField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthField" + j.ToString))
                        .LengthUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthUnits" + j.ToString))

                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabClassField" + j.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabQuanField" + j.ToString))
                        .HabUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabUnits" + j.ToString))
                    End With
                    ' add to the module level list
                    pLLayersFields.Add(HabLayerObj2)
                Next

            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            iCount1 = pLLayersFields.Count
            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pLLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for river layer. Please choose a field in the options menu.", "Parameter Missing")
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                Next
            End If

            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    sBarrierPermField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierPermField" + j.ToString))
                    sBarrierNaturalYNField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierNaturalYNField" + j.ToString))
                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, sBarrierPermField, sBarrierNaturalYNField, Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            End If

        Else
            ' TODO: Pop-up form with current settings
            ' Add a button on form that loads options
            ' upon close, if options form within for has been opened
            ' then keep looping
            Dim sMessage As String = "Please check options in menu and re-run tool"
            System.Windows.Forms.MessageBox.Show(sMessage, "Parameters Missing")
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        ' =========================== END READ EXTENSION SETTINGS ===================================


        ' ------------------------------------------------
        ' 1.4 Prepare traceflow solver

        ''MsgBox("Debug:6")
        '=====================
        ' Can use a 2 dimensional Array to populate statistics
        ' These are multiple recordsets with rows and columns.
        Dim statsArray3D(4, 1) As Double 'VB.NET

        Dim bFID, bFCID, bSubID, iEID As Integer
        Dim pFLyrSlct As IFeatureLayer  ' to set all layers as selectable

        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)

        ' =============== SAVE ORIGINAL GEONET SETTINGS ================
        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        ' =============== Save the junction barriers ===================
        ' Save the barriers
        Dim iOriginalJunctionBarrierCount2 As Integer
        iOriginalJunctionBarrierCount2 = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        'MsgBox("Debug Mes: The Original junction barrier count second check: " + iOriginalJunctionBarrierCount2.ToString)


        For i = 0 To iOriginalJunctionBarrierCount2 - 1

            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            pFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
            'originalBarriersList(i) = bEID
        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = Nothing
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

        Dim iOriginalJunctionBarrierCount3 As Integer
        iOriginalJunctionBarrierCount3 = pOriginalBarriersList.Count
        'MsgBox("Debug Mes: The Original junction barrier count third check: " + iOriginalJunctionBarrierCount3.ToString)

        ' =============== Save the Junction Flags ===================
        ' Save the flags
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.JunctionFlagCount - 1
            ' Use the bFlagDisplay to retrieve the EIDs of the junction flags
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.JunctionFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginaljuncFlagsListGEN.Add(bEID)
            'pOriginaljuncFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        'pOriginaljuncFlagsList = 

        pOriginaljuncFlagsList = Nothing
        pOriginaljuncFlagsList = CType(pOriginaljuncFlagsListGEN, IEnumNetEID)

        ' ============ Save the Edge Flags ======================
        ' ******** NO EDGE FLAG SUPPORT YET *********
        i = 0
        For i = 0 To pNetworkAnalysisExtFlags.EdgeFlagCount - 1

            ' Use the bFlagDisplay to retrieve EIDs of the Edge flags for later
            pFlagDisplay = CType(pNetworkAnalysisExtFlags.EdgeFlag(i), IFlagDisplay)
            bEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalEdgeFlagsListGEN.Add(bEID)
            'pOriginalEdgeFlagsList(i) = bEID
        Next

        ' QI to and get an array interface that has 'count' and 'next' methods
        pOriginalEdgeFlagsList = Nothing
        pOriginalEdgeFlagsList = CType(pOriginalEdgeFlagsListGEN, IEnumNetEID)
        ' ====================== END SAVE ORIGINAL GEONET SETTINGS ========================

        ' If there are no flags set exit sub
        If pNetworkAnalysisExtFlags.JunctionFlagCount = 0 Then
            'If pNetworkAnalysisExtFlags.EdgeFlagCount = 0 Then
            MsgBox("There are no flags set on junctions.  Please Set flags only on network junctions.")
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
            'End If
        End If

        ' =============== FLAG CONSISTENCY CHECK ====================
        ' Check for consistency that flags are all on barriers or all on non-barriers.

        sFlagCheck = flagcheck(pOriginalBarriersList, pOriginalEdgeFlagsList, pOriginaljuncFlagsList)
        '    MsgBox "FlagCheck may be null and crash..."
        '    MsgBox "sFlagcheck: " + sFlagCheck
        If sFlagCheck = "error" Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        ' ============= END FLAG CONSISTENCY CHECK ==================

        ' Make sure all the layers in the TOC are selectable
        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLyrSlct = CType(pMap.Layer(i), IFeatureLayer)
                    pFLyrSlct.Selectable = True
                End If
            End If
        Next

        Dim pFeatureClass As IFeatureClass
        Dim pFeatureLayer As IFeatureLayer
        Dim pLayer As ILayer
        Dim sFeatureLayerName As String
        Dim lFCIDs As New List(Of FCIDandNameObject) ' to hold FCIDs of included line layers
        Dim pFCIDandNameObject As New FCIDandNameObject(Nothing, Nothing)
        Dim iLineFCID, iPolyFCID As Integer
        Dim bLineFCIDFound, bPolyFCIDFound As Boolean


        ''MsgBox("Debug:7")

        ' Need to get the FC Id's of the line layers because
        ' names might change and there might be duplicate names, names
        ' that repeat but are different FC's, (weird stuff like that)
        ' NOTE: this should have been done in Options code and stored in Extension settings.
        i = 0
        j = 0
        m = 0
        Try
            For i = 0 To pLLayersFields.Count - 1
                For j = 0 To pMap.LayerCount - 1
                    If pMap.Layer(j).Valid = True And TypeOf pMap.Layer(j) Is IFeatureLayer Then
                        pLayer = pMap.Layer(j)
                        pFeatureLayer = CType(pLayer, IFeatureLayer)
                        sFeatureLayerName = pLLayersFields(i).Layer
                        If pFeatureLayer.Name = sFeatureLayerName Then
                            'Note: there are a number of options for FC ID's - this one seems most reasonable.
                            iLineFCID = pFeatureLayer.FeatureClass.FeatureClassID
                            If iLineFCID <> -1 Then ' check that this is not a shapefile
                                'Check that this FCID isn't already in the list
                                bLineFCIDFound = False
                                pFCIDandNameObject = New FCIDandNameObject(Nothing, Nothing)
                                If lFCIDs.Count > 0 Then
                                    For m = 0 To lFCIDs.Count - 1
                                        If lFCIDs(m).FCID = iLineFCID Then
                                            bLineFCIDFound = True
                                        End If
                                    Next
                                    If bLineFCIDFound = False Then
                                        pFCIDandNameObject.FCID = iLineFCID
                                        pFCIDandNameObject.Name = pFeatureLayer.Name
                                        lFCIDs.Add(pFCIDandNameObject)
                                        lAllFCIDs.Add(pFCIDandNameObject)
                                    End If
                                Else
                                    pFCIDandNameObject.FCID = iLineFCID
                                    pFCIDandNameObject.Name = pFeatureLayer.Name
                                    lFCIDs.Add(pFCIDandNameObject)
                                    lAllFCIDs.Add(pFCIDandNameObject)
                                End If
                            End If ' not a shapefile or coverage
                        End If
                    End If
                Next
            Next


        Catch ex As Exception
            MsgBox("Could not get the participating lines layers from table of contents. Now Exiting. " + ex.Message)
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End Try


        ' Add the polygon layers to the ALL FCIDs list\
        ' this is use later in the predicate search to get unique 
        ' feature classes to draw out from the master habitat object
        i = 0
        j = 0
        m = 0

        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(15, "Getting Participating Polygons Layers")
        Try
            For i = 0 To pPLayersFields.Count - 1
                For j = 0 To pMap.LayerCount - 1
                    If pMap.Layer(j).Valid = True And TypeOf pMap.Layer(j) Is IFeatureLayer Then
                        pLayer = pMap.Layer(j)
                        pFeatureLayer = CType(pLayer, IFeatureLayer)
                        If pFeatureLayer.Name = pPLayersFields(i).Layer Then
                            'Note: there are a number of options for FC ID's - this one seems most reasonable.
                            iPolyFCID = pFeatureLayer.FeatureClass.FeatureClassID
                            If iPolyFCID <> -1 Then ' check that this is not a shapefile
                                'Check that this FCID isn't already in the list
                                bPolyFCIDFound = False
                                pFCIDandNameObject = New FCIDandNameObject(Nothing, Nothing)
                                If lAllFCIDs.Count > 0 Then
                                    For m = 0 To lAllFCIDs.Count - 1
                                        If lAllFCIDs(m).FCID = iPolyFCID Then
                                            bPolyFCIDFound = True
                                        End If
                                    Next
                                    If bPolyFCIDFound = False Then
                                        pFCIDandNameObject.FCID = iPolyFCID
                                        pFCIDandNameObject.Name = pFeatureLayer.Name
                                        lAllFCIDs.Add(pFCIDandNameObject)
                                    End If
                                Else
                                    pFCIDandNameObject.FCID = iPolyFCID
                                    pFCIDandNameObject.Name = pFeatureLayer.Name
                                    lAllFCIDs.Add(pFCIDandNameObject)
                                End If
                            End If ' not a shapefile or coverage
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox("Could not get the participating polygons layers from table of contents. Now Exiting. " + ex.Message)
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End Try


        ' ========================== Begin Traces ====================

        ''MsgBox("Debug:8")
        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(20, "Beginning Network Traces")

        Dim barrierCount As Integer
        barrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        Dim pSymbol As ISymbol
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim sOutID As String 'ID of flag for output table
        Dim iFCID, iFID, iSubID As Integer
        Dim pNoSourceFlowEnds As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNoSourceFlowEndsTemp As IEnumNetEIDBuilderGEN

        Dim eFlowElements As esriFlowElements = esriFlowElements.esriFEJunctionsAndEdges
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim pResultEdges As IEnumNetEID
        Dim pResultJunctions As IEnumNetEID
        Dim pEnumNetEIDBuilder As IEnumNetEIDBuilder
        Dim pFlowEndJunctionsPer As IEnumNetEID
        Dim pFlowEndEdgesPer As IEnumNetEID
        Dim sHabTypeKeyword As String
        Dim pUID As UID

        pUID = New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        Dim pEnumLayer As IEnumLayer
        pEnumLayer = pMap.Layers(pUID, True)
        pEnumLayer.Reset()


        If sFlagCheck = "barriers" Then

            Dim pNextTraceBarrierEIDs As IEnumNetEID
            Dim pNextTraceBarrierEIDGEN As IEnumNetEIDBuilderGEN
            Dim flagOverBarrier As Boolean

            ' ============ FILTER BARRIERS (eliminate barriers where flags are) ===
            i = 0
            'reset / initialize OriginalBarriers list
            pOriginalBarriersList.Reset()
            pNextTraceBarrierEIDGEN = New EnumNetEIDArray

            ' Get the element ID of this flag
            pOriginaljuncFlagsList.Reset()
            iEID = pOriginaljuncFlagsList.Next()

            ' Query the corresponding user ID's to the element ID
            pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

            ' For each of the original barriers set
            ' Check if it overlaps with a flag
            ' if not, then add it to a list to use in 
            ' this trace
            For i = 0 To barrierCount - 1
                bEID = pOriginalBarriersList.Next

                flagOverBarrier = False
                If bEID = iEID Then
                    flagOverBarrier = True
                End If
                If Not flagOverBarrier Then
                    pNextTraceBarrierEIDGEN.Add(bEID)
                End If
            Next

            'QI to get 'next' and 'count'
            pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)
            ' ====================== END FILTER BARRIERS =============================

            ' clear current barriers
            pNetworkAnalysisExtBarriers.ClearBarriers()

            ' ========================== SET BARRIERS  ===============================
            m = 0
            pNextTraceBarrierEIDs.Reset()


            ' Set the barrier symbol color and parameters
            With pRgbColor
                .Red = 255
                .Green = 0
                .Blue = 0
            End With

            With pSimpleMarkerSymbol
                .Color = pRgbColor
                .Style = esriSimpleMarkerStyle.esriSMSX
                .Outline = True
                .Size = 10
            End With
            pSymbol = CType(pSimpleMarkerSymbol, ISymbol)

            For m = 0 To pNextTraceBarrierEIDs.Count - 1
                bEID = pNextTraceBarrierEIDs.Next
                pNetElements.QueryIDs(bEID, esriElementType.esriETJunction, bFCID, bFID, bSubID)

                ' Display the barriers as a JunctionFlagDisplay type
                pFlagDisplay = New JunctionFlagDisplay
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
            ' ========================== END SET BARRIERS ===========================
        End If ' flagcheck = barriers

        ' this variable holds trace flow ends that are not sources
        pNoSourceFlowEndsTemp = New EnumNetEIDArray  'Reset temp variable

        ' Get the element ID of this flag
        pOriginaljuncFlagsList.Reset()
        iEID = pOriginaljuncFlagsList.Next()

        ' Query the corresponding user ID's to the element ID
        pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

        ' ====================== RUN TRACE IN DIRECTION OF ANALYSIS ====================
        '                            TO GET FLOW END ELEMENTS

        ''MsgBox("Debug:9")
        'prepare the network solver
        pTraceFlowSolver = TraceFlowSolverSetup2()
        If pTraceFlowSolver Is Nothing Then
            System.Windows.Forms.MessageBox.Show("Could not set up the network. Check that there is a network loaded.", "TraceFlowSolver2 setup error.")
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If

        eFlowElements = esriFlowElements.esriFEJunctionsAndEdges


        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        backgroundworker2.ReportProgress(20, "Performing Network Traces" & ControlChars.NewLine & _
                                         "Trace: upstream immediate next barriers.")

        'Return the features stopping the trace
        pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMUpstream, eFlowElements, pFlowEndJunctionsPer, pFlowEndEdgesPer)

        ' ============================ END RUN TRACE  ===============================

        ' ================= GET BARRIER ID ===================
        Dim pIDAndType As New IDandType(Nothing, Nothing)
        Dim sType As String
        pIDAndType = New IDandType(Nothing, Nothing)
        pIDAndType = GetBarrierID(iFCID, iFID, lBarrierIDs)
        Dim bNaturalY As Boolean = False

        sOutID = pIDAndType.BarrID
        sType = pIDAndType.BarrIDType

        If sType <> "Sink" And sFlagCheck = "nonbarr" Then
            sType = "Flag - Node"
        ElseIf sType <> "Sink" And sFlagCheck = "barrier" Then
            sType = "Flag - barrier"
        End If

        Dim dBarrierPerm As Double
        Dim sNaturalYN As String

        If bBarrierPerm = True Then
            dBarrierPerm = GetBarrierPerm(iFCID, iFID, lBarrierIDs)
        Else
            dBarrierPerm = 0
        End If
        If bNaturalYN = True Then
            sNaturalYN = GetNaturalYN(iFCID, iFID, lBarrierIDs)
        Else
            sNaturalYN = "F"
        End If

        If sNaturalYN = "T" Then
            bNaturalY = True
        End If

        '' Will save this sOutID and sType for later use, if this is orderloop zero (flag)
        '' because will need to insert the DCI Metric at the end of this flag loop
        '' else if it's a barrier in a greater order loop we're visiting, then keep
        '' track of their id's for later use.  
        'If orderLoop = 0 Then
        '    f_sOutID = sOutID
        '    f_siOutEID = iEID
        '    f_sType = sType
        'Else
        '    pBarrierAndSinkEIDs = New BarrAndBarrEIDAndSinkEIDs(iEID, iEID, sOutID)
        '    lBarrierAndSinkEIDs.Add(pBarrierAndSinkEIDs)
        'End If

        pMetricsObject = New MetricsObject(sOutID, iEID, sOutID, iEID, sType, "Permeability", dBarrierPerm)
        lMetricsObject.Add(pMetricsObject)

        'sOutID = GetBarrierID(iFCID, iFID, lBarrierIDs)

        ' =========================== FILTER FLOW END ELEMENTS (no sources) =================
        ' Filter first flow end elements so that only barriers are included
        ' 
        ' 1) For each flow-stopping element 
        '   2) For each original barrier set by the user
        '     3) If there is a match 
        '       4) If they have not already been encountered before
        '         (need this check in case trace was on indeterminate flow - 
        '          to avoid infinite looping)
        '         5)Add them to a list to use as FLAGS later

        'pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)
        pFlowEndJunctionsPer.Reset()

        Dim p As Integer = 0
        Dim bKeepEID As Boolean
        Dim iEndEID As Integer

        For p = 0 To pFlowEndJunctionsPer.Count - 1

            bKeepEID = False 'initializw
            iEndEID = pFlowEndJunctionsPer.Next
            m = 0
            pOriginalBarriersList.Reset()
            For m = 0 To pOriginalBarriersList.Count - 1
                If iEndEID = pOriginalBarriersList.Next Then
                    bKeepEID = True ' set true if found
                    'pAllFlowEndBarriers.Reset()
                    'For k = 0 To pAllFlowEndBarriers.Count - 1
                    '    If endEID = pAllFlowEndBarriers.Next Then
                    '        keepEID = False ' set false if already on master list
                    '    End If
                    'Next
                End If
            Next

            If bKeepEID = True Then
                ' to crosscheck in case of infinite loop problem
                pNoSourceFlowEnds.Add(iEndEID) 'This variable gets reset each 
                ' order loop
            End If
        Next


        Dim sHabType, sHabDir As String
        Dim pTotalResultsJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pTotalResultsEdgesGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pSubtractUpJunctions, pSubtractUpEdges As IEnumNetEID
        Dim pConnectedJunctions, pConnectedEdges, pDownConnectedJunctions, pDownConnectedEdges As IEnumNetEID

        ' if the following are needed then have to do upstream trace
        ' downstream because if bDownhab then need to subtract upstream from connected trace
        If bUpHab = True Or bDownHab = True Then

            ''MsgBox("Debug:10")

            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(25, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: upstream immediate.")

            pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)

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

            If bUpHab = True Then
                ' ============ SAVE FOR HIGHLIGHTING ==================
                ' Get results to display as highlights at end of sub
                pResultJunctions.Reset()
                k = 0
                For k = 0 To pResultJunctions.Count - 1
                    pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
                Next
                pResultEdges.Reset()
                For k = 0 To pResultEdges.Count - 1
                    pTotalResultsEdgesGEN.Add(pResultEdges.Next)
                Next

                ' Get results as selection
                pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)


                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(30, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: upstream immediate." & ControlChars.NewLine & _
                                                 "Intersecting Features")

                ' ============== INTERSECT FEATURES ================

                ''MsgBox("Debug:11")
                Call IntersectFeatures()
                ''MsgBox("Debug:12")

                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(30, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: upstream immediate." & ControlChars.NewLine & _
                                                 "Excluding Features")
                ' ---- EXCLUDE FEATURES -----
                pEnumLayer.Reset()
                ' Look at the next layer in the list
                pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                    If pFeatureLayer.Valid = True Then
                        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                            ''MsgBox("Debug:13")
                            ExcludeFeatures(pFeatureLayer)
                            ''MsgBox("Debug:14")
                        End If
                    End If
                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                Loop

                ' ---- END EXCLUDE FEATURES -----
                ''MsgBox("Debug: 14a")
                ' ============ SAVE FOR HIGHLIGHTING ==================
                ' Get results to display as highlights at end of sub
                pResultJunctions.Reset()
                k = 0
                For k = 0 To pResultJunctions.Count - 1
                    pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
                Next
                pResultEdges.Reset()
                For k = 0 To pResultEdges.Count - 1
                    pTotalResultsEdgesGEN.Add(pResultEdges.Next)
                Next

                ' use results for Upstream habitat if needed
                sHabType = "Immediate"
                sHabDir = "upstream"

                Call calculateStatistics_2(lHabStatsList,
                                           sOutID,
                                           iEID,
                                           sType,
                                           sOutID,
                                           iEID,
                                           sHabType,
                                           sHabDir)
             
            End If

            ' need to do a "connected" trace and subtract upstream trace results from it
            If bDownHab = True Then

                ''MsgBox("Debug:13b")
                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(35, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: downstream immediate.")

                ' Store upstream elements from last trace to use to subtract from connected 
                pSubtractUpJunctions = pResultJunctions
                pSubtractUpEdges = pResultEdges

                pMap.ClearSelection() ' clear selection

                pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMConnected, eFlowElements, pResultJunctions, pResultEdges)

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

                pConnectedJunctions = pResultJunctions
                pConnectedEdges = pResultEdges

                ' Subtract the upstream edges and junctions from the connected edges and junctions list
                pDownConnectedJunctions = DownStreamConnected(pSubtractUpJunctions, pConnectedJunctions)
                pDownConnectedEdges = DownStreamConnected(pSubtractUpEdges, pConnectedEdges)

                ' ============ SAVE FOR HIGHLIGHTING ==================
                ' Get results to display as highlights at end of sub
                pDownConnectedJunctions.Reset()
                k = 0
                For k = 0 To pDownConnectedJunctions.Count - 1
                    pTotalResultsJunctionsGEN.Add(pDownConnectedJunctions.Next)
                Next
                pDownConnectedEdges.Reset()
                For k = 0 To pDownConnectedEdges.Count - 1
                    pTotalResultsEdgesGEN.Add(pDownConnectedEdges.Next)
                Next

                ' Get results as selection
                pNetworkAnalysisExtResults.CreateSelection(pDownConnectedJunctions, pDownConnectedEdges)

                ' Get results to display as highlights at end of sub
                pDownConnectedJunctions.Reset()
                k = 0
                For k = 0 To pDownConnectedJunctions.Count - 1
                    pTotalResultsJunctionsGEN.Add(pDownConnectedJunctions.Next)
                Next
                pDownConnectedEdges.Reset()
                For k = 0 To pDownConnectedEdges.Count - 1
                    pTotalResultsEdgesGEN.Add(pDownConnectedEdges.Next)
                Next

                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(35, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: downstream immediate." & ControlChars.NewLine & _
                                                 "Intersecting Features")
                ' ============== INTERSECT FEATURES ================
                ''MsgBox("Debug:14e")
                Call IntersectFeatures()
                ''MsgBox("Debug:15")
                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(35, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: downstream immediate." & ControlChars.NewLine & _
                                                 "Excluding Features")

                ' ---- EXCLUDE FEATURES -----
                pEnumLayer.Reset()
                ' Look at the next layer in the list
                pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                    If pFeatureLayer.Valid = True Then
                        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                            ExcludeFeatures(pFeatureLayer)
                        End If
                    End If
                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                Loop
                ' ---- END EXCLUDE FEATURES -----

                sHabType = "Immediate"
                sHabDir = "downstream"
                Call calculateStatistics_2(lHabStatsList, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
            End If
        End If  ' bUpHab = True Or bDCI = True Or bDownHab = True 


        ' If Downstream Path Habitat desired
        If bPathDownHab = True Then

            ''MsgBox("Debug:16")
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(40, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: path downstream.")

            pMap.ClearSelection() ' clear selection

            pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMDownstream, eFlowElements, pResultJunctions, pResultEdges)

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

            ' ============ SAVE FOR HIGHLIGHTING ==================
            ' Get results to display as highlights at end of sub
            pResultJunctions.Reset()
            k = 0
            For k = 0 To pResultJunctions.Count - 1
                pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
            Next
            pResultEdges.Reset()
            For k = 0 To pResultEdges.Count - 1
                pTotalResultsEdgesGEN.Add(pResultEdges.Next)
            Next

            ' Get results as selection
            pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
            ' =================== INTERSECT FEATURES =============
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(40, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: path downstream." & ControlChars.NewLine & _
                                                 "Intersecting Features")

            ''MsgBox("Debug:17")
            Call IntersectFeatures()

            ''MsgBox("Debug:18")
            ' ---- EXCLUDE FEATURES -----
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(40, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: path downstream." & ControlChars.NewLine & _
                                                 "Excluding Features")

            pEnumLayer.Reset()
            ' Look at the next layer in the list
            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                If pFeatureLayer.Valid = True Then
                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                        ExcludeFeatures(pFeatureLayer)
                    End If
                End If
                pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
            Loop
            ' ---- END EXCLUDE FEATURES -----

            pActiveView.Refresh() ' refresh the view
            sHabType = "Path"
            sHabDir = "downstream"
            Call calculateStatistics_2(lHabStatsList, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
        End If

        ' If any total tables desired clear all barriers and run traces. 
        If bTotalUpHab = True Or bTotalDownHab = True Then

            ''MsgBox("Debug:19")
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(45, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: total upstream/total downstream.")

            pNetworkAnalysisExtBarriers.ClearBarriers()
            pTraceFlowSolver = TraceFlowSolverSetup2()

            'prepare the network solver
            If pTraceFlowSolver Is Nothing Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If

            eFlowElements = esriFlowElements.esriFEJunctionsAndEdges

            If bTotalUpHab = True Or bTotalDownHab = True Then

                pMap.ClearSelection() ' clear selection

                pNetworkAnalysisExtBarriers.ClearBarriers()
                pTraceFlowSolver = TraceFlowSolverSetup2()

                ' perform UPSTREAM trace
                pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, eFlowElements, pResultJunctions, pResultEdges)

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

                If bTotalUpHab = True Then
                    ' ============ SAVE FOR HIGHLIGHTING ==================
                    ' Get results to display as highlights at end of sub
                    pResultJunctions.Reset()
                    k = 0
                    For k = 0 To pResultJunctions.Count - 1
                        pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
                    Next
                    pResultEdges.Reset()
                    For k = 0 To pResultEdges.Count - 1
                        pTotalResultsEdgesGEN.Add(pResultEdges.Next)
                    Next


                    ' Get results as selection
                    pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)

                    ' =================== INTERSECT FEATURES =============
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(45, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total upstream." & ControlChars.NewLine & _
                                                     "Intersecting Features.")

                    ''MsgBox("Debug:20")
                    Call IntersectFeatures()
                    ''MsgBox("Debug:21")
                    ' ---- EXCLUDE FEATURES -----
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(45, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total upstream." & ControlChars.NewLine & _
                                                     "Excluding Features.")

                    pEnumLayer.Reset()
                    ' Look at the next layer in the list
                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                    Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                        If pFeatureLayer.Valid = True Then
                            If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                            pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                            pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                ExcludeFeatures(pFeatureLayer)
                            End If
                        End If
                        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                    Loop
                    ' ---- END EXCLUDE FEATURES -----

                    pActiveView.Refresh() ' refresh the view
                    sHabType = "Total"
                    sHabDir = "upstream"
                    Call calculateStatistics_2(lHabStatsList, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                End If

                If bTotalDownHab = True Then

                    ''MsgBox("Debug:22")

                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(50, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total downstream.")

                    ' Store current results for cross-ref and subtraction.
                    pSubtractUpJunctions = pResultJunctions
                    pSubtractUpEdges = pResultEdges

                    pMap.ClearSelection() ' clear selection

                    ' find connected
                    pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMConnected, eFlowElements, pResultJunctions, pResultEdges)

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

                    ' ========= SUBTRACT UPSTREAM FROM ALL CONNECTED =======
                    pConnectedJunctions = pResultJunctions
                    pConnectedEdges = pResultEdges

                    ' Subtract the upstream edges and junctions from the connected edges and junctions list
                    pDownConnectedJunctions = DownStreamConnected(pSubtractUpJunctions, pConnectedJunctions)
                    pDownConnectedEdges = DownStreamConnected(pSubtractUpEdges, pConnectedEdges)

                    ' ============ SAVE FOR HIGHLIGHTING ==================
                    ' Get results to display as highlights at end of sub
                    pDownConnectedJunctions.Reset()
                    k = 0
                    For k = 0 To pDownConnectedJunctions.Count - 1
                        pTotalResultsJunctionsGEN.Add(pDownConnectedJunctions.Next)
                    Next
                    pDownConnectedEdges.Reset()
                    For k = 0 To pDownConnectedEdges.Count - 1
                        pTotalResultsEdgesGEN.Add(pDownConnectedEdges.Next)
                    Next

                    ' Get results as selection
                    pNetworkAnalysisExtResults.CreateSelection(pDownConnectedJunctions, pDownConnectedEdges)

                    ' =================== INTERSECT FEATURES =============
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(50, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total downstream." & ControlChars.NewLine & _
                                                     "Intersecting features.")

                    ''MsgBox("Debug:23")
                    Call IntersectFeatures()
                    ''MsgBox("Debug:24")
                    ' ---- EXCLUDE FEATURES -----
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(50, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total downstream." & ControlChars.NewLine & _
                                                     "Excluding features.")
                    pEnumLayer.Reset()
                    ' Look at the next layer in the list
                    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                    Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                        If pFeatureLayer.Valid = True Then
                            If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                            pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                            pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                ExcludeFeatures(pFeatureLayer)
                            End If
                        End If
                        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
                    Loop
                    ' ---- END EXCLUDE FEATURES -----
                    pActiveView.Refresh() ' refresh the view

                    sHabType = "Total"
                    sHabDir = "downstream"
                    Call calculateStatistics_2(lHabStatsList, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                End If
            End If ' bTotalUpHab = True Or bTotalDownHab = True
        End If ' bTotalUpHab = True Or bTotalDownHab = True Or bTotalPathDownHab = True

        If bTotalPathDownHab = True Then

            ''MsgBox("Debug:25")
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(55, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: total path downstream.")
            pMap.ClearSelection() ' clear selection

            pNetworkAnalysisExtBarriers.ClearBarriers()
            pTraceFlowSolver = TraceFlowSolverSetup2()

            ' perform UPSTREAM trace
            pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMDownstream, eFlowElements, pResultJunctions, pResultEdges)

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

            ' ============ SAVE FOR HIGHLIGHTING ==================
            ' Get results to display as highlights at end of sub
            pResultJunctions.Reset()
            k = 0
            For k = 0 To pResultJunctions.Count - 1
                pTotalResultsJunctionsGEN.Add(pResultJunctions.Next)
            Next
            pResultEdges.Reset()
            For k = 0 To pResultEdges.Count - 1
                pTotalResultsEdgesGEN.Add(pResultEdges.Next)
            Next

            ' Get results as selection
            pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
            ' =================== INTERSECT FEATURES =============
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(55, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: total path downstream." & ControlChars.NewLine & _
                                             "Intersecting features.")

            ''MsgBox("Debug:26")
            Call IntersectFeatures()
            ''MsgBox("Debug:27")
            ' ---- EXCLUDE FEATURES -----

            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(55, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: total path downstream." & ControlChars.NewLine & _
                                             "Excluding features.")
            pEnumLayer.Reset()
            ' Look at the next layer in the list
            pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                If pFeatureLayer.Valid = True Then
                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                        ExcludeFeatures(pFeatureLayer)
                    End If
                End If
                pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
            Loop
            ' ---- END EXCLUDE FEATURES -----

            pActiveView.Refresh() ' refresh the view
            sHabType = "Total Path"
            sHabDir = "downstream"
            Call calculateStatistics_2(lHabStatsList, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
        End If

        ' Create a result highlight of all areas traced
        Dim pTotalResultsEdges, pTotalResultsJunctions As IEnumNetEID
        pTotalResultsEdges = CType(pTotalResultsEdgesGEN, IEnumNetEID)
        pTotalResultsJunctions = CType(pTotalResultsJunctionsGEN, IEnumNetEID)
        pNetworkAnalysisExtResults.ResultsAsSelection = False
        pNetworkAnalysisExtResults.SetResults(pTotalResultsJunctions, pTotalResultsEdges)
        pNetworkAnalysisExtResults.ResultsAsSelection = True

        If sFlagCheck = "barriers" Or bTotalUpHab = True Or bTotalDownHab = True Or bTotalPathDownHab = True Then

            ''MsgBox("Debug:28")
            ' ========================= RESET BARRIERS =========================
            m = 0
            pOriginalBarriersList.Reset()
            pNetworkAnalysisExtBarriers.ClearBarriers()

            Dim iOriginalJunctionBarrierCount4 As Integer = pOriginalBarriersList.Count
            'MsgBox("Debug Mes: The Original junction barrier count fourth check: " + iOriginalJunctionBarrierCount4.ToString)

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
            ' ======================= END RESET BARRIERS ======================
        End If



        ' ======================== RESET FLAGS ================================

        ''MsgBox("Debug:29")
        Dim pEdgeFlagDisplay As IEdgeFlagDisplay
        ' Clear current flags
        pNetworkAnalysisExtFlags.ClearFlags()

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
        ' ================================= END RESET FLAGS ================================


        '' Update output form labels
        '' Bring results form to front
        'pResultsForm.BringToFront()
        'pResultsForm.txtRichResults.Select(0, 0)
        EndTime = DateTime.Now
        'pResultsForm.lblBeginTime.Text = "Begin Time: " & BeginTime
        'pResultsForm.lblEndTime.Text = "End Time: " & EndTime

        Dim TotalTime As TimeSpan
        TotalTime = EndTime - BeginTime
        'pResultsForm.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
        'pResultsForm.lblDirection.Text = "Analysis Direction: " + sDirection
        'pResultsForm.lblOrder.Text = "Order of Analysis: n/a (1)"
        'pResultsForm.lblNumBarriers.Text = "Number of Barriers Analysed: 1"


        ' ================== BEGIN WRITE TO OUTPUT FORM =================
        ''MsgBox("Debug:30")

        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        m_iProgress = 60
        backgroundworker2.ReportProgress(m_iProgress, "Preparing Output Form.")

        ' Output Form (will replace dockable window)
        Dim pResultsForm3 As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3
        pResultsForm3.Show()

        Dim pSinkAndDCIS As New SinkandDCIs(Nothing, Nothing, Nothing, Nothing)
        Dim lSinkAndDCIS As New List(Of SinkandDCIs)
        Dim pSinkIDAndTypes As New SinkandTypes(Nothing, Nothing, Nothing)
        Dim lSinkIDandTypes As New List(Of SinkandTypes)

        i = 0
        Dim bSinkThere, bDCIpMatch, bDCIdMatch, bEntered As Boolean
        Dim row As DataRow

        ' this bit gets a list of the unique sinks from the 
        ' metrics object list
        ' and it populates a datatable with unique sinks list
        For i = 0 To lMetricsObject.Count - 1
            j = 0
            bSinkThere = False
            For j = 0 To lSinkIDandTypes.Count - 1
                If lSinkIDandTypes(j).SinkEID = lMetricsObject(i).SinkEID Then
                    bSinkThere = True
                End If
            Next
            If bSinkThere = False Then
                pSinkIDAndTypes = New SinkandTypes(lMetricsObject(i).SinkEID, lMetricsObject(i).Sink, lMetricsObject(i).Type)
                lSinkIDandTypes.Add(pSinkIDAndTypes)
                'row = pSinkTable.NewRow()
                'row.ItemArray = New Object() {lMetricsObject(i).SinkEID, lMetricsObject(i).Type}
                'pSinkTable.Rows.Add(row)
            End If
        Next

        ' ============ BEGIN WRITE TO DATAGRID OUTPUT SUMMARY TABLE =============
        ' Set up the table - create columns if needed

        pResultsForm3.DataGridView1.Columns.Add("Flag", "Flag")           '0
        pResultsForm3.DataGridView1.Columns.Add("FlagEID", "FlagEID")       '1
        pResultsForm3.DataGridView1.Columns.Add("Type", "Type")           '2
        pResultsForm3.DataGridView1.Columns.Add("Metric", "Metric")       '3
        pResultsForm3.DataGridView1.Columns.Add("Value", "Value")         '4

        ' If there are habitat statistics then add the proper columns
        If lHabStatsList.Count > 0 Then
            pResultsForm3.DataGridView1.Columns.Add("Layer", "Layer")                 '5
            pResultsForm3.DataGridView1.Columns.Add("Direction", "Direction")         '6
            pResultsForm3.DataGridView1.Columns.Add("Type", "Type")                   '7
            pResultsForm3.DataGridView1.Columns.Add("HabitatClass", "Habitat_Class")  '8
            pResultsForm3.DataGridView1.Columns.Add("Quantity", "Quantity")           '9
            pResultsForm3.DataGridView1.Columns.Add("Unit", "Unit")                   '10
        End If

        i = 0
        For i = 0 To pResultsForm3.DataGridView1.Columns.Count - 1
            pResultsForm3.DataGridView1.Columns.Item(i).SortMode = DataGridViewColumnSortMode.Programmatic
        Next i

        i = 0
        Dim iMaxRowIndex, iSinkRowIndex, iSinkRowCount, iHabRowIndex, iHabRowcount, iBarrRowIndex, iBarrRowCount As Integer ' loop counters and grid indices
        Dim iMetricRowIndex, iMetricRowCount, iThisHabRowIndex, iThisHabRowCount As Integer
        Dim iSinkEID, iBarrEID As Integer
        Dim dTotalHab As Double 'running total of habitat for table
        Dim sLayer, sDirection2, sTraceType As String
        Dim bTrigger As Boolean = False
        Dim bTrigger2 As Boolean = False
        Dim bColorSwitcher = False
        Dim bSinkVisit As Boolean = True
        Dim t As Integer = 0
        i = 0
        iMaxRowIndex = 0
        j = 0

        Dim sinkBarrierLayerComparer As FindLayerAndBarrEIDAndSinkEIDPredicate ' for refining large stats object, reduce looping
        Dim sinkcomparer As FindBarriersBySinkEIDPredicate
        Dim barriercomparer As FindBarriersBySinkEIDPredicate ' used for refining habitat stats list 
        Dim barriermetriccomparer As FindBarrierMetricsBySinkEIDPredicate  ' used for refining barrier metrics stats list

        Dim refinedHabitatList As List(Of StatisticsObject_2)           ' for refining habitat stats list
        Dim refinedBarrierEIDList As List(Of BarrAndBarrEIDAndSinkEIDs) ' for refining barrier list
        Dim refinedBarrierMetricsList As List(Of MetricsObject)

        Dim HabStatsComparer As RefineHabStatsListPredicate


        Dim pDataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle
        ' For each sink
        ' 1. for each sink in the master sinks object list
        ' 2. for each barrier associated with the sink in the master barriers list.  
        ' 2a add the metrics.

        ' notes: -the maxrow index keeps track of which row we're at,
        '        it's needed if there are multiple sinks
        '        -the isinkrow count keeps track of which row for this
        '        sink we're at so that the first row can be found
        '        -results form is populated with sinkID, not EID
        For i = 0 To lSinkIDandTypes.Count - 1

            iSinkRowCount = 0
            iBarrRowCount = 0
            j = 0
            k = 0
            iSinkRowIndex = pResultsForm3.DataGridView1.Rows.Add()
            iMaxRowIndex = iSinkRowIndex ' the new maximum row count
            iSinkEID = lSinkIDandTypes(i).SinkEID

            ' Add the sink ID to the table
            '    record the row number
            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(0).Style
            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(1).Style
            'pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, pResultsForm3.DataGridView1.Font.Size, FontStyle.Bold)
            pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, 14, FontStyle.Bold)
            pDataGridViewCellStyle.ForeColor = Color.DarkGreen
            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(0).Style = pDataGridViewCellStyle

            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(0).Value = lSinkIDandTypes(i).SinkID
            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(1).Value = lSinkIDandTypes(i).SinkEID
            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Value = lSinkIDandTypes(i).Type

            ' post up the sink-specific metrics and stats. 
            '    For each of the records in the metrics list
            '     add the values associated with this sink
            '     keep track of the number of rows added
            '     and the max row count of the table.  

            For k = 0 To lMetricsObject.Count - 1
                ' matching the 'barrier' EID - which is redundant
                ' and includes 'sink' metrics, too.  
                If lMetricsObject(k).BarrEID = iSinkEID Then
                    If iSinkRowCount = 0 Then
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Value = lMetricsObject(k).MetricName
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Value = Math.Round(lMetricsObject(k).Metric, 2)
                    ElseIf iSinkRowCount > 0 Then
                        pResultsForm3.DataGridView1.Rows.Add()
                        ' keep track of the maximum number of rows in the table
                        iMaxRowIndex = iMaxRowIndex + 1 ' Row tracker
                        'pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Value = lMetricsObject(j).ID
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Value = lMetricsObject(k).MetricName
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Value = Math.Round(lMetricsObject(k).Metric, 2)

                    End If
                    iSinkRowCount += 1
                End If
            Next 'Metric Object

            t = 0
            iHabRowIndex = iSinkRowIndex
            iBarrRowIndex = iSinkRowIndex
            iHabRowcount = 0
            bSinkVisit = True ' to pass to Sub to tell it whether to increment the barrier loop counter
            For t = 0 To lAllFCIDs.Count - 1

                Dim iTemp As Integer
                'bUpHab, bTotalUpHab, bDownHab, bTotalDownHab, bPathDownHab, bTotalPathDownHab
                iTemp = lAllFCIDs(t).FCID

                If bUpHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "upstream", "Immediate")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
                If bTotalUpHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "upstream", "Total")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
                If bDownHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "downstream", "Immediate")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
                If bTotalDownHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "downstream", "Total")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
                If bPathDownHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "downstream", "Path")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
                If bTotalPathDownHab = True Then
                    HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, iSinkEID, lAllFCIDs(t).FCID, "downstream", "Total Path")
                    refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                    UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                End If
            Next ' layer included (t)

            If bColorSwitcher = True Then
                bColorSwitcher = False
            Else
                bColorSwitcher = True
            End If

            iHabRowcount = 0
            iBarrRowIndex = 0
            iBarrRowCount = 0
            bTrigger = False 'indicates if the there's been another row added beyond the sink (any barriers)
            dTotalHab = 0
            bColorSwitcher = True

            ' 1. a refined list of all barriers for the sink
            ' use a comparer to get all the records from the barriers and sinks list that match the sink
            barriercomparer = New FindBarriersBySinkEIDPredicate(iSinkEID)
            refinedBarrierEIDList = lBarrierAndSinkEIDs.FindAll(AddressOf barriercomparer.CompareEID)

            ' For each barrier
            '  1. a refined list of all habitat stats for this barrier 
            '     and sink and layer
            '  2. for each layer get a refined list of habitat metrics 
            '     associated with each layer, sink, barrier combo
            k = 0
            For k = 0 To refinedBarrierEIDList.Count - 1


                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                m_iProgress = m_iProgress + 1
                backgroundworker2.ReportProgress(m_iProgress, "Writing to Output Form." & ControlChars.NewLine & _
                                                 "Writing for Barrier " & (k + 1).ToString & " of " & (refinedBarrierEIDList.Count).ToString)

                iBarrRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                iMaxRowIndex = iBarrRowIndex
                iBarrRowCount = 0
                iSinkRowCount += 1
                iBarrRowCount += 1
                bTrigger = False

                ' attempt at a border
                ' border control not available as of 2005 and .net 2.0
                'Dim pPainter As Windows.Forms.DataGridViewRowPrePaintEventArgs
                'pPainter = pResultsForm3.DataGridView1..Rows(iMaxRowIndex).
                '' add barrier ID to the datagrid
                '    record the row number
                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Style
                If bColorSwitcher = False Then
                    pDataGridViewCellStyle.BackColor = Color.PowderBlue
                Else
                    pDataGridViewCellStyle.BackColor = Color.Lavender
                End If
                pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, pResultsForm3.DataGridView1.Font.Size, FontStyle.Bold)
                pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Value = refinedBarrierEIDList(k).BarrLabel
                pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Value = refinedBarrierEIDList(k).BarrEID

                ' get refined list of barrier/sink metrics
                barriermetriccomparer = New FindBarrierMetricsBySinkEIDPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID)
                refinedBarrierMetricsList = lMetricsObject.FindAll(AddressOf barriermetriccomparer.CompareEID)

                ' For each metric in the refined list
                ' insert a new row if necessary
                ' add the metric to the table
                t = 0
                iMetricRowCount = 0
                iMetricRowIndex = iMaxRowIndex
                For t = 0 To refinedBarrierMetricsList.Count - 1


                    If iMetricRowCount = 0 Then
                        'pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Value = lMetricsObject(j).ID
                        If bColorSwitcher = True Then
                            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Style
                            pDataGridViewCellStyle.BackColor = Color.Lavender
                        Else
                            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Style
                            pDataGridViewCellStyle.BackColor = Color.PowderBlue
                        End If
                        pResultsForm3.DataGridView1.Rows(iMetricRowIndex).Cells(5).Style = pDataGridViewCellStyle
                        pResultsForm3.DataGridView1.Rows(iMetricRowIndex).Cells(6).Style = pDataGridViewCellStyle

                        pResultsForm3.DataGridView1.Rows(iMetricRowIndex).Cells(5).Value = refinedBarrierMetricsList(t).MetricName
                        pResultsForm3.DataGridView1.Rows(iMetricRowIndex).Cells(6).Value = Math.Round(refinedBarrierMetricsList(t).Metric, 2)
                        iMetricRowCount += 1
                    ElseIf iMetricRowCount > 0 Then
                        If iMaxRowIndex <= (iBarrRowIndex + iMetricRowCount) Then
                            iMetricRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                            'iBarrRowIndex = iMetricRowIndex
                            iBarrRowCount += 1
                            iMetricRowCount += 1
                            iMaxRowIndex += 1 ' Row tracker
                            iSinkRowCount += 1
                        End If
                        If bColorSwitcher = True Then
                            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Style
                            pDataGridViewCellStyle.BackColor = Color.Lavender
                        Else
                            pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Style
                            pDataGridViewCellStyle.BackColor = Color.PowderBlue
                        End If
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Style = pDataGridViewCellStyle
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Style = pDataGridViewCellStyle
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Style = pDataGridViewCellStyle
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(6).Style = pDataGridViewCellStyle


                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Value = refinedBarrierMetricsList(t).MetricName
                        pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(6).Value = Math.Round(refinedBarrierMetricsList(t).Metric, 2)
                    End If
                Next ' refined metric (t)

                t = 0
                iHabRowIndex = iBarrRowIndex
                iHabRowcount = 0
                bSinkVisit = False
                For t = 0 To lAllFCIDs.Count - 1

                    Dim iTemp As Integer
                    'bUpHab, bTotalUpHab, bDownHab, bTotalDownHab, bPathDownHab, bTotalPathDownHab
                    iTemp = lAllFCIDs(t).FCID

                    If bUpHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "upstream", "Immediate")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                    If bTotalUpHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "upstream", "Total")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                    If bDownHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "downstream", "Immediate")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                    If bTotalDownHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "downstream", "Total")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                    If bPathDownHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "downstream", "Path")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                    If bTotalPathDownHab = True Then
                        HabStatsComparer = New RefineHabStatsListPredicate(iSinkEID, refinedBarrierEIDList(k).BarrEID, lAllFCIDs(t).FCID, "downstream", "Total Path")
                        refinedHabitatList = lHabStatsList.FindAll(AddressOf HabStatsComparer.CompareHabStuff)
                        UpdateSummaryTable(refinedHabitatList, iHabRowcount, pResultsForm3, iMaxRowIndex, iBarrRowIndex, bColorSwitcher, iSinkRowCount, iBarrRowCount, bSinkVisit)
                    End If
                Next ' layer included (t)

                If bColorSwitcher = True Then
                    bColorSwitcher = False
                Else
                    bColorSwitcher = True
                End If
            Next ' barrier for this sink (k)
        Next ' sink 

        pResultsForm3.DataGridView1.AutoResizeColumns()
        EndTime = DateTime.Now
        pResultsForm3.lblBeginTime.Text = "Begin Time: " & BeginTime
        pResultsForm3.lblEndtime.Text = "End Time: " & EndTime

        TotalTime = EndTime - BeginTime
        pResultsForm3.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
        pResultsForm3.lblDirection.Text = "Analysis Direction: " + sDirection
        If iOrderNum <> 999 Then
            pResultsForm3.lblOrder.Text = "Order of Analysis: " & CStr(iOrderNum)
        Else
            pResultsForm3.lblOrder.Text = "Order of Analysis: Max"
        End If

        'If Not pAllFlowEndBarriers Is Nothing Then
        '    If pAllFlowEndBarriers.Count <> 0 Then
        '        pResultsForm3.lblNumBarriers.Text = "Number of Barriers Analysed: " & CStr(pAllFlowEndBarriers.Count + pOriginaljuncFlagsList.Count)
        '    Else
        '        pResultsForm3.lblNumBarriers.Text = "Number of Barriers Analysed: 1"
        '    End If
        'End If

        ' ============== END WRITE TO OUTPUT SUMMARY TABLE =================

        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.Dispose()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If

        backgroundworker2.ReportProgress(100, "Completed!")
        m_iProgress = 100
        'If Not ProgressForm Is Nothing Then
        '    ProgressForm.Close()
        'End If

        backgroundworker2.Dispose()
        backgroundworker2.CancelAsync()
        ' refresh the view
        pActiveView.Refresh()

    End Sub
    Private Sub LabelFlag(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
        ' Purpose: Label barriers or flags if no label is present
        '          using user-set field from extension settings
        '   
        ' 1.0 Read the extension settings
        ' 2.0 label the flag if needed 
        ' 
        ' Bug Note: Since there are issues with finding visible label
        '           elements (see notes further down) the workaround
        '           shows labels as visible if they have been turned on
        '           and then off in ArcMap.  So if you have turned labels
        '           off they might still be in the annotation properties
        '           and return as 'visible.'

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)

        ' ------------------------------------
        ' 1.0 Read Extension Barrier label settings
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim j, i As Integer
        Dim iBarrierIDs As Integer
        Dim sBarrierIDLayer, sBarrierIDField As String

        ' If settings have been set by the user then load them
        If m_FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    ' Object to retrieve barrier label field (do not need other obj. params so set to 'nothing')
                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, Nothing, Nothing, Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            End If
        End If

        ' ------------------------------------
        ' 2.0 Label the flag if needed

        Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        'Dim pMapAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        Dim pAnnoLayerProps As IAnnotateLayerProperties
        Dim aLELayerProps As ILabelEngineLayerProperties
        Dim pNewLELayerProps As ILabelEngineLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        'Dim pElement As IElement
        'Dim pTextElement As ITextElement
        'Dim pAnnotateMapProps As IAnnotateMapProperties
        'Dim pAnnotateMap As IAnnotateMap
        Dim pEnumInVisibleElements As IElementCollection
        Dim sLabelField, sLabelValue As String
        Dim sSQL, strOIDName As String
        Dim sPreviousWhereClause As String = ""
        Dim pGFLayer As IGeoFeatureLayer
        Dim pFLayer As IFeatureLayer
        'Dim sText As String
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
        Dim pFieldType As esriFieldType
        Dim iFieldIndex As Integer
        Dim pFields As IFields
        Dim bClassMatch As Boolean = False
        Dim sSearchString As String
        Dim iClassNum As Integer
        Dim iStringSearchReturn As Integer

        Dim pFeature As IFeature
        Dim pFeatureClass As IFeatureClass
        'Dim iTemp As Integer

        ' For each layer in the map
        '   If it's a feature layer
        '     If it's the same layer as the new flag is being placed on
        '       For each layer in the BarrierID list
        '         If it matches this layer
        '           If the field is found in this layer
        '             Then use this field as a label        '
        '       Get the annotation properties of the layer
        '       For each of the properties
        '         Get the visible elements
        '         If any one label matches the ObjectID of the element
        '           Alert that there's a match (so labels won't be changed)
        '         If there was no match found then
        '           Create a new label class called "FlagID"

        ''MsgBox("Debug:L1")

        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(i), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        ''MsgBox("Debug:L2")

                        bLabelMatch = False 'reset match variable for label
                        pGFLayer = CType(pMap.Layer(i), IGeoFeatureLayer)

                        ' For each layer in the BarrierID list
                        '   If it matches this layer
                        '     If the field is found in this layer
                        '       Then use this field as a label
                        bUserLabel = False
                        If lBarrierIDs IsNot Nothing Then
                            If lBarrierIDs.Count <> 0 Then
                                For j = 0 To lBarrierIDs.Count - 1
                                    If lBarrierIDs.Item(j).Layer = pFLayer.Name Then

                                        ''MsgBox("Debug:L3")
                                        pFeatureClass = pFLayer.FeatureClass
                                        pFeature = pFeatureClass.GetFeature(iFID)
                                        pFields = pFeature.Fields

                                        If pFields.FindField(lBarrierIDs.Item(j).Field) <> -1 Then

                                            ''MsgBox("Debug:L4")
                                            sLabelField = lBarrierIDs.Item(j).Field
                                            Try
                                                sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))
                                            Catch ex As Exception
                                                ''MsgBox("Debug:Could not convert barrier ID label to string.")
                                                sLabelValue = "NoName"
                                            End Try

                                            ' If a value was returned then set the alert variable true
                                            If sLabelValue <> "" Then

                                                bUserLabel = True

                                                ' Get the field type of the field because if it is a string
                                                ' the sql requires quotation wrappers. 
                                                iFieldIndex = pFeatureClass.FindField(sLabelField)
                                                pFieldType = pFields.Field(iFieldIndex).Type

                                                If pFieldType = esriFieldType.esriFieldTypeString Then
                                                    bString = True
                                                End If

                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        ' Since pAnnotationLayerPropertiesCollection.QueryItem does not return the proper
                        ' visible elements a workaround found here 
                        ' http://forums.esri.com/Thread.asp?c=93&f=993&t=164358 
                        ' or here for details
                        ' http://edn.esri.com/index.cfm?fa=codeExch.sampleDetail&pg=/arcobjects/9.1/Samples/Cartography/Labeling_and_Annotation/LabelsToMapAnno.htm
                        ' was attempted - cloning the annotation properties did not work, though. 


                        'Dim propsIndex As Integer
                        'Dim pClone As IClone

                        '' Clone the map annotation properties collection
                        'For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)
                        '    pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps1, Nothing, Nothing)
                        '    If Not pAnnoLayerProps1 Is Nothing Then
                        '        'Clone the properties and add them to the new collection
                        '        pClone = CType(pAnnoLayerProps1, IClone)
                        '        pMapAnnoLayerPropsColl.Add(CType(pClone.Clone, IAnnotateLayerProperties))
                        '    End If
                        'Next

                        'pEnumVisibleElements = New ElementCollection
                        'pEnumInVisibleElements = New ElementCollection
                        'pMapAnnoLayerPropsColl.QueryItem(0, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                        ' See http://forums.esri.com/Thread.asp?c=93&f=992&t=69162#180195 for 
                        ' discussion of this code
                        ' Get the Annotation Collection
                        pAnnoLayerPropsColl = pGFLayer.AnnotationProperties

                        'MsgBox("Debug:L5")

                        Dim propsIndex As Integer
                        For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)

                            pEnumVisibleElements = New ElementCollection
                            pEnumInVisibleElements = New ElementCollection
                            pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                            ' If there is already a class called "BarrierOrFlagID" 
                            ' Then get the where clause and save it for later
                            If pAnnoLayerProps.Class.ToString = "BarrierOrFlagID" Then
                                sPreviousWhereClause = pAnnoLayerProps.WhereClause.ToString ' & " OR "
                                bClassMatch = True
                                iClassNum = propsIndex
                            End If

                            ' If the layer had no labels already then make sure the new 
                            ' class (default) is defaulted to 'off' - don't show labels
                            ' otherwise it will show labels for the 'default' class. 
                            If propsIndex = 0 And pAnnoLayerPropsColl.Count = 1 Then
                                If pGFLayer.DisplayAnnotation = False Then
                                    pAnnoLayerProps.DisplayAnnotation = False
                                End If
                            End If

                            'populate visible elements collection
                            'pAnnotateMap = pMap.AnnotationEngine
                            'pAnnotateMapProps = New AnnotateMapProperties
                            'pAnnotateMapProps.AnnotateLayerPropertiesCollection = pAnnoLayerPropsColl
                            'pAnnoLayerProps.FeatureLayer = pGFLayer
                            'pAnnotateMap.Label(pAnnotateMapProps, pMap, Nothing)

                            ' Query the element collection for the visible labels
                            ' And see if there is already a label that matches the 
                            ' one that would be placed.
                            ' If a label field has been found from the extension settings
                            ' then search for that in the visible labels.  If not search 
                            ' for the feature objectID. 
                            'For j = 0 To pEnumVisibleElements.Count - 1

                            '    pEnumVisibleElements.QueryItem(j, pElement)
                            '    pTextElement = CType(pElement, ITextElement)

                            '    If (Not pTextElement Is Nothing) Then
                            '        sText = pTextElement.Text
                            '        If bUserLabel = True Then
                            '            If sLabelValue = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        Else
                            '            If iFID.ToString = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        End If
                            '    End If
                            'Next j
                        Next ' PropsIndex

                        ''MsgBox("Debug:L6")
                        ' TEMP SOLUTION
                        ' In where clause from the class, if found, search for the current feature
                        '   Check if using user settings for label to create search string
                        If bClassMatch = True Then
                            bLabelMatch = False
                            If sPreviousWhereClause <> "" Then
                                If bUserLabel = True Then
                                    If bString = True Then
                                        sSearchString = sLabelField + " = '" + sLabelValue + "'"
                                    Else
                                        sSearchString = sLabelField + " = " + sLabelValue
                                    End If
                                Else
                                    sSearchString = pGFLayer.FeatureClass.OIDFieldName & " = " & iFID.ToString
                                End If
                                iStringSearchReturn = InStr(sPreviousWhereClause, sSearchString)
                                If iStringSearchReturn <> 0 Then
                                    bLabelMatch = True
                                End If
                            End If

                        End If

                        ''MsgBox("Debug:L7")
                        ' If there was no match between the visible labels and the feature's
                        ' value in the label field
                        If bLabelMatch = False Then

                            ' If we're using a label field specified by the user 
                            ' then use that field to label, otherwise use the FID field
                            If sPreviousWhereClause <> "" Then
                                If bUserLabel = True Then
                                    If bString = True Then
                                        sSQL = sPreviousWhereClause & " OR " & sLabelField & " = " & "'" & sLabelValue & "'"
                                    Else
                                        sSQL = sPreviousWhereClause & " OR " & sLabelField & " = " & sLabelValue
                                    End If
                                Else
                                    strOIDName = pGFLayer.FeatureClass.OIDFieldName
                                    sSQL = sPreviousWhereClause & " OR " & strOIDName & " = " & iFID
                                    sLabelField = strOIDName
                                End If
                            Else
                                If bUserLabel = True Then
                                    If bString = True Then
                                        sSQL = sLabelField & " = " & "'" & sLabelValue & "'"
                                    Else
                                        sSQL = sLabelField & " = " & sLabelValue
                                    End If
                                Else
                                    strOIDName = pGFLayer.FeatureClass.OIDFieldName
                                    sSQL = strOIDName & " = " & iFID
                                    sLabelField = strOIDName
                                End If
                            End If

                            ''MsgBox("Debug:L8")
                            ' If there was a class match found
                            '   Then get that annotation layer properties set
                            '    set the 'where' clause
                            ' If there was no match found
                            '    need to add a class so name it
                            '    set the 'where' clause
                            ' See here for explanation of new layerprops class
                            ' http://resources.esri.com/help/9.3/arcgisengine/dotnet/d3f93845-fedc-42f1-827b-912038c6271b.htm
                            Try
                                If bClassMatch = True Then
                                    pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                    pAnnoLayerProps.WhereClause = sSQL
                                    pAnnoLayerProps.DisplayAnnotation = True
                                Else
                                    pNewLELayerProps = New LabelEngineLayerPropertiesClass()
                                    pAnnoLayerProps = CType(pNewLELayerProps, IAnnotateLayerProperties)
                                    pAnnoLayerProps.Class = "BarrierOrFlagID"
                                    pAnnoLayerProps.WhereClause = sSQL
                                    pAnnoLayerProps.DisplayAnnotation = True
                                    pAnnoLayerPropsColl.Add(pAnnoLayerProps)
                                End If
                            Catch ex As Exception
                                MsgBox("A serious error was encountered trying to label this feature. " + ex.Message)
                            End Try


                            aLELayerProps = CType(pAnnoLayerProps, ILabelEngineLayerProperties)
                            aLELayerProps.Expression = "[" & sLabelField & "]"

                            ''MsgBox("Debug:L9")
                            ' Change symbol - add halo
                            Dim pBOLP As IBasicOverposterLayerProperties = New BasicOverposterLayerProperties
                            pBOLP.PointPlacementOnTop = False
                            pBOLP.PointPlacementMethod = esriOverposterPointPlacementMethod.esriAroundPoint
                            pBOLP.LabelWeight = esriBasicOverposterWeight.esriHighWeight

                            Dim pTxtSym As ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                            pTxtSym = aLELayerProps.Symbol

                            'mask color
                            Dim pRgbColor As IRgbColor = New RgbColor
                            ' Set the barrier symbol color and parameters
                            With pRgbColor
                                .Red = 255
                                .Green = 255
                                .Blue = 255
                            End With

                            Dim pFillSymbol As IFillSymbol = New SimpleFillSymbol
                            pFillSymbol.Color = pRgbColor

                            Dim pLineSymbol As ILineSymbol = New SimpleLineSymbol
                            pLineSymbol.Color = pRgbColor
                            pFillSymbol.Outline = pLineSymbol

                            Dim pMask As IMask
                            pMask = CType(pTxtSym, IMask)

                            pMask.MaskStyle = esriMaskStyle.esriMSHalo
                            pMask.MaskSize = 1.8
                            pMask.MaskSymbol = pFillSymbol

                            pTxtSym = CType(pMask, ITextSymbol)

                            aLELayerProps.Symbol = pTxtSym
                            aLELayerProps.BasicOverposterLayerProperties = pBOLP
                            pGFLayer.DisplayAnnotation = True

                            ''MsgBox("Debug:L10")
                        End If
                    End If
                End If
            End If
        Next
    End Sub
    Private Function flagcheck(ByRef pBarriersList As IEnumNetEID, ByRef pEdgeFlags As IEnumNetEID, ByRef pJuncFlags As IEnumNetEID) As String
        ' Function:     Flag Check
        ' Author:       Greig Oldford
        ' Date Created: March 26th, 2008
        ' Date Translated: May 11 - 19, 2009 
        ' Description:  This function will check to see if initially the flags are on barriers
        '               or not. It should stop the analysis if some flags are on barriers and some
        '               are not - it should be one or the other.
        '               It will return one of three strings - "barrier", "nonbarr", or "error"
        '
        ' Notes:   March 27, 2008 -> Not currently checking for edge flags since there is no support for
        '                      them yet.
        Dim i As Integer
        Dim m As Integer
        Dim pFlagsOnBarrGEN As IEnumNetEIDBuilderGEN  ' list holds flags on barriers
        Dim pFlagsNoBarrGEN As IEnumNetEIDBuilderGEN  ' list holds flag not on barriers
        pFlagsOnBarrGEN = New EnumNetEIDArray
        pFlagsNoBarrGEN = New EnumNetEIDArray

        Dim pFlagsOnBarr As IEnumNetEID
        Dim pFlagsNoBarr As IEnumNetEID
        Dim flagBarrier As Boolean
        Dim iEID As Integer

        pJuncFlags.Reset()
        i = 0

        ' For each flag
        For i = 0 To pJuncFlags.Count - 1

            flagBarrier = False     ' assume flag is not on barrier
            iEID = pJuncFlags.Next  ' get the EID of flag
            m = 0
            pBarriersList.Reset()

            ' For each barrier
            For m = 0 To pBarriersList.Count - 1
                'If endEID = pOriginalBarriersList(m) Then 'VB.NET
                If iEID = pBarriersList.Next Then
                    flagBarrier = True
                End If
            Next

            If flagBarrier = True Then  'put EID in flags on barrier list

                ' THIS LIST COULD BE USED IN FUTURE TO FILTER BAD FLAGS OUT
                ' I.E. - check which flags are on barriers and remove only those ones automatically.
                pFlagsOnBarrGEN.Add(iEID)

            Else   ' put EID in flags not on barrier list
                pFlagsNoBarrGEN.Add(iEID)
            End If
        Next

        ' QI to get "next" and "count"
        pFlagsOnBarr = CType(pFlagsOnBarrGEN, IEnumNetEID)
        pFlagsNoBarr = CType(pFlagsNoBarrGEN, IEnumNetEID)

        If pFlagsOnBarr.Count = pJuncFlags.Count Then
            flagcheck = "barriers"
            'return "barriers"? ' should be a return in VB.Net I think... but this works
        ElseIf pFlagsNoBarr.Count = pJuncFlags.Count Then
            flagcheck = "nonbarr"
        Else
            MsgBox("Inconsistent flag placement." + vbCrLf + _
            "Barrier flags: " & pFlagsOnBarr.Count & vbCrLf & _
            " Non-barrier flags: " & pFlagsNoBarr.Count)
            flagcheck = "error"
        End If
    End Function
    Private Sub UnLabelBarrier(ByVal iFCID As Integer, ByVal iFID As Integer)

        ' Created By: Greig Oldford
        ' Date: July 5, 2009
        ' Purpose: Label barriers or flags if no label is present
        '          using user-set field from extension settings
        '   
        ' 1.0 Read the extension settings
        ' 2.0 label the flag if needed 
        ' 
        ' Bug Note: Since there are issues with finding visible label
        '           elements (see notes further down) the workaround
        '           shows labels as visible if they have been turned on
        '           and then off in ArcMap.  So if you have turned labels
        '           off they might still be in the annotation properties
        '           and return as 'visible.'

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pMxApp As IMxApplication = CType(My.ArcMap.Application, IMxApplication)

        ' ------------------------------------
        ' 1.0 Read Extension Barrier label settings
        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim j, i As Integer
        Dim iBarrierIDs As Integer
        Dim sBarrierIDLayer, sBarrierIDField As String

        ' If settings have been set by the user then load them
        If m_FiPEx__1.m_bLoaded = True Then
            ' Get the barrier ID Fields
            iBarrierIDs = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numBarrierIDs"))
            If iBarrierIDs > 0 Then
                For j = 0 To iBarrierIDs - 1
                    sBarrierIDLayer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDLayer" + j.ToString))
                    sBarrierIDField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("BarrierIDField" + j.ToString))
                    ' Object to retrieve barrier label field (do not need other obj. params so set to 'nothing')
                    pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, Nothing, Nothing, Nothing)
                    lBarrierIDs.Add(pBarrierIDObj)
                Next
            End If
        End If

        ' ------------------------------------
        ' 2.0 Label the flag if needed

        Dim pAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        'Dim pMapAnnoLayerPropsColl As IAnnotateLayerPropertiesCollection
        Dim pAnnoLayerProps As IAnnotateLayerProperties
        Dim pNewLELayerProps As ILabelEngineLayerProperties
        Dim aLELayerProps As ILabelEngineLayerProperties
        Dim pEnumVisibleElements As IElementCollection
        'Dim pElement As IElement
        'Dim pTextElement As ITextElement
        'Dim pAnnotateMapProps As IAnnotateMapProperties
        'Dim pAnnotateMap As IAnnotateMap
        Dim pEnumInVisibleElements As IElementCollection
        Dim sLabelField, sLabelValue As String
        Dim sSQL, strOIDName As String
        Dim sPreviousWhereClause As String = ""
        Dim pGFLayer As IGeoFeatureLayer
        Dim pFLayer As IFeatureLayer
        'Dim sText As String
        Dim bLabelMatch As Boolean = False
        Dim bUserLabel As Boolean = False
        Dim bString As Boolean = False
        Dim pFieldType As esriFieldType
        Dim iFieldIndex As Integer
        Dim pFields As IFields
        Dim bClassMatch As Boolean
        Dim iClassNum As Integer
        Dim sSearchString As String
        Dim pFeature As IFeature
        Dim pFeatureClass As IFeatureClass
        Dim sORSearchString, sSearchStringOR As String
        Dim iStringPosit As Integer
        ' For each layer in the map
        '   If it's a feature layer
        '     If it's the same layer as the new flag is being placed on
        '       For each layer in the BarrierID list
        '         If it matches this layer
        '           If the field is found in this layer
        '             Then use this field as a label        '
        '       Get the annotation properties of the layer
        '       For each of the properties
        '         Get the visible elements
        '         If any one label matches the ObjectID of the element
        '           Alert that there's a match (so labels won't be changed)
        '         If there was no match found then
        '           Create a new label class called "FlagID"



        For i = 0 To pMap.LayerCount - 1
            If pMap.Layer(i).Valid = True Then
                If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(i), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        bLabelMatch = False 'reset match variable for label
                        pGFLayer = CType(pMap.Layer(i), IGeoFeatureLayer)

                        ' For each layer in the BarrierID list
                        '   If it matches this layer
                        '     If the field is found in this layer
                        '       Then use this field as a label
                        bUserLabel = False
                        If lBarrierIDs IsNot Nothing Then
                            If lBarrierIDs.Count <> 0 Then
                                For j = 0 To lBarrierIDs.Count - 1
                                    If lBarrierIDs.Item(j).Layer = pFLayer.Name Then

                                        pFeatureClass = pFLayer.FeatureClass
                                        pFeature = pFeatureClass.GetFeature(iFID)
                                        pFields = pFeature.Fields

                                        If pFields.FindField(lBarrierIDs.Item(j).Field) <> -1 Then

                                            sLabelField = lBarrierIDs.Item(j).Field
                                            sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))

                                            ' If a value was returned then set the alert variable true
                                            If sLabelValue <> "" Then

                                                bUserLabel = True

                                                ' Get the field type of the field because if it is a string
                                                ' the sql requires quotation wrappers. 
                                                iFieldIndex = pFeatureClass.FindField(sLabelField)
                                                pFieldType = pFields.Field(iFieldIndex).Type

                                                If pFieldType = esriFieldType.esriFieldTypeString Then
                                                    bString = True
                                                End If

                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        ' Since pAnnotationLayerPropertiesCollection.QueryItem does not return the proper
                        ' visible elements a workaround found here 
                        ' http://forums.esri.com/Thread.asp?c=93&f=993&t=164358 
                        ' or here for details
                        ' http://edn.esri.com/index.cfm?fa=codeExch.sampleDetail&pg=/arcobjects/9.1/Samples/Cartography/Labeling_and_Annotation/LabelsToMapAnno.htm
                        ' was attempted - cloning the annotation properties did not work, though. 


                        'Dim propsIndex As Integer
                        'Dim pClone As IClone

                        '' Clone the map annotation properties collection
                        'For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)
                        '    pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps1, Nothing, Nothing)
                        '    If Not pAnnoLayerProps1 Is Nothing Then
                        '        'Clone the properties and add them to the new collection
                        '        pClone = CType(pAnnoLayerProps1, IClone)
                        '        pMapAnnoLayerPropsColl.Add(CType(pClone.Clone, IAnnotateLayerProperties))
                        '    End If
                        'Next

                        'pEnumVisibleElements = New ElementCollection
                        'pEnumInVisibleElements = New ElementCollection
                        'pMapAnnoLayerPropsColl.QueryItem(0, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                        ' See http://forums.esri.com/Thread.asp?c=93&f=992&t=69162#180195 for 
                        ' discussion of this code
                        ' Get the Annotation Collection
                        pAnnoLayerPropsColl = pGFLayer.AnnotationProperties

                        Dim propsIndex As Integer
                        For propsIndex = 0 To (pAnnoLayerPropsColl.Count - 1)

                            pEnumVisibleElements = New ElementCollection
                            pEnumInVisibleElements = New ElementCollection
                            pAnnoLayerPropsColl.QueryItem(propsIndex, pAnnoLayerProps, pEnumVisibleElements, pEnumInVisibleElements)

                            ' If there is already a class called "BarrierOrFlagID" 
                            ' Then get the where clause and save it for later
                            If pAnnoLayerProps.Class.ToString = "BarrierOrFlagID" Then
                                sPreviousWhereClause = pAnnoLayerProps.WhereClause.ToString
                                bClassMatch = True
                                iClassNum = propsIndex
                            End If

                            '' If the layer had no labels already then make sure the new 
                            '' class (default) is defaulted to 'off' - don't show labels
                            '' otherwise it will show labels for the 'default' class. 
                            'If propsIndex = 0 And pAnnoLayerPropsColl.Count = 1 Then
                            '    If pGFLayer.DisplayAnnotation = False Then
                            '        pAnnoLayerProps.DisplayAnnotation = False
                            '    End If
                            'End If

                            'populate visible elements collection
                            'pAnnotateMap = pMap.AnnotationEngine
                            'pAnnotateMapProps = New AnnotateMapProperties
                            'pAnnotateMapProps.AnnotateLayerPropertiesCollection = pAnnoLayerPropsColl
                            'pAnnoLayerProps.FeatureLayer = pGFLayer
                            'pAnnotateMap.Label(pAnnotateMapProps, pMap, Nothing)

                            ' Query the element collection for the visible labels
                            ' And see if there is already a label that matches the 
                            ' one that would be placed.
                            ' If a label field has been found from the extension settings
                            ' then search for that in the visible labels.  If not search 
                            ' for the feature objectID. 
                            'For j = 0 To pEnumVisibleElements.Count - 1

                            '    pEnumVisibleElements.QueryItem(j, pElement)
                            '    pTextElement = CType(pElement, ITextElement)

                            '    If (Not pTextElement Is Nothing) Then
                            '        sText = pTextElement.Text
                            '        If bUserLabel = True Then
                            '            If sLabelValue = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        Else
                            '            If iFID.ToString = sText Then
                            '                bLabelMatch = True
                            '            End If
                            '        End If
                            '    End If
                            'Next j
                        Next ' PropsIndex

                        ' TEMP SOLUTION
                        ' In where clause from the class, if found, search for the current feature
                        '   Check if using user settings for label to create search string
                        If bClassMatch = True And sPreviousWhereClause <> "" And bUserLabel = True Then
                            If bString = True Then
                                sSearchString = sLabelField + " = '" + sLabelValue + "'"
                            Else
                                sSearchString = sLabelField + " = " + sLabelValue
                            End If
                        Else
                            sSearchString = pGFLayer.FeatureClass.OIDFieldName & " = " & iFID.ToString
                        End If
                        iStringPosit = InStr(sPreviousWhereClause, sSearchString)
                        If iStringPosit <> 0 Then
                            bLabelMatch = True

                            ' need to re-search to check if this clause is found in the middle of the 
                            ' where filter string or on the end (if middle then there is an "OR" to remove)
                            ' Three situations: 1) searchstring is the only one present (no "or")
                            '                   2) searchstring is at the end (preceded by "or")
                            '                   3) searchstring is at the beginning (followed by "or")
                            sORSearchString = "OR " & sSearchString
                            sSearchStringOR = sSearchString & " OR"

                            If InStr(sPreviousWhereClause, sSearchStringOR) <> 0 Then

                                sSearchString = sORSearchString
                                sSQL = sPreviousWhereClause.Remove(iStringPosit - 1, sSearchString.Length + 1)
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL

                            ElseIf InStr(sPreviousWhereClause, sORSearchString) <> 0 Then

                                sSearchString = "OR " + sSearchString
                                sSQL = sPreviousWhereClause.Remove(iStringPosit - 5, sSearchString.Length + 1)
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL

                                ' otherwise remove the preceding "OR" from this
                            Else
                                sSQL = ""
                                pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                                pAnnoLayerProps.WhereClause = sSQL
                                pAnnoLayerProps.DisplayAnnotation = False

                            End If

                        End If

                        '' If there was no match between the visible labels and the feature's
                        '' value in the label field
                        'If bLabelMatch = False Then

                        '    ' If we're using a label field specified by the user 
                        '    ' then use that field to label, otherwise use the FID field
                        '    If bUserLabel = True Then
                        '        If bString = True Then
                        '            sSQL = sPreviousWhereClause & sLabelField & " = " & "'" & sLabelValue & "'"
                        '        Else
                        '            sSQL = sPreviousWhereClause & sLabelField & " = " & sLabelValue
                        '        End If
                        '    Else
                        '        strOIDName = pGFLayer.FeatureClass.OIDFieldName
                        '        sSQL = sPreviousWhereClause & strOIDName & " = " & iFID
                        '        sLabelField = strOIDName
                        '    End If

                        '    ' If there was a class match found
                        '    '   Then get that annotation layer properties set
                        '    '    set the 'where' clause
                        '    ' If there was no match found
                        '    '    need to add a class so name it
                        '    '    set the 'where' clause
                        '    ' See here for explanation of new layerprops class
                        '    ' http://resources.esri.com/help/9.3/arcgisengine/dotnet/d3f93845-fedc-42f1-827b-912038c6271b.htm
                        '    If bClassMatch = True Then
                        '        pAnnoLayerPropsColl.QueryItem(iClassNum, pAnnoLayerProps, Nothing, Nothing)
                        '        pAnnoLayerProps.WhereClause = sSQL
                        '    Else
                        '        pNewLELayerProps = New LabelEngineLayerPropertiesClass()
                        '        pAnnoLayerProps = CType(pNewLELayerProps, IAnnotateLayerProperties)
                        '        pAnnoLayerProps.Class = "BarrierOrFlagID"
                        '        pAnnoLayerProps.WhereClause = sSQL
                        '        pAnnoLayerPropsColl.Add(pAnnoLayerProps)
                        '    End If

                        'aLELayerProps = CType(pAnnoLayerProps, ILabelEngineLayerProperties)
                        'aLELayerProps.Expression = "[" & sLabelField & "]"

                        '' Change symbol - add halo
                        'Dim pBOLP As IBasicOverposterLayerProperties = New BasicOverposterLayerProperties
                        'pBOLP.PointPlacementOnTop = False
                        'pBOLP.PointPlacementMethod = esriOverposterPointPlacementMethod.esriAroundPoint
                        'pBOLP.LabelWeight = esriBasicOverposterWeight.esriHighWeight

                        'Dim pTxtSym As ITextSymbol = New ESRI.ArcGIS.Display.TextSymbol
                        'pTxtSym = aLELayerProps.Symbol

                        ''mask color
                        'Dim pRgbColor As IRgbColor = New RgbColor
                        '' Set the barrier symbol color and parameters
                        'With pRgbColor
                        '    .Red = 255
                        '    .Green = 255
                        '    .Blue = 255
                        'End With

                        'Dim pFillSymbol As IFillSymbol = New SimpleFillSymbol
                        'pFillSymbol.Color = pRgbColor

                        'Dim pLineSymbol As ILineSymbol = New SimpleLineSymbol
                        'pLineSymbol.Color = pRgbColor
                        'pFillSymbol.Outline = pLineSymbol

                        'Dim pMask As IMask
                        'pMask = CType(pTxtSym, IMask)

                        'pMask.MaskStyle = esriMaskStyle.esriMSHalo
                        'pMask.MaskSize = 1.8
                        'pMask.MaskSymbol = pFillSymbol

                        'pTxtSym = CType(pMask, ITextSymbol)

                        'aLELayerProps.Symbol = pTxtSym
                        'aLELayerProps.BasicOverposterLayerProperties = pBOLP

                        pGFLayer.DisplayAnnotation = True
                    End If
                End If
            End If

        Next
    End Sub
    Private Function GetNaturalYN(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As String
        ' =============== FLAG ON POINT WITH Perm? =========================
        ' This section checks whether there is a BarrierPerm Field
        ' In which case it will use this field to identify the 
        ' barrier in the output.
        ' 1.0 For each layer in the table of contents
        '   2.0 If the layer is a FeatureLayer
        '     3.0 If the layer ClassID matches the layer of the CURRENT FLAG
        '     3.1 Get the field values for the feature
        '       4.0 For each of the Barrier IDs in the list
        '         5.0 If the layer in the BarrierID list matches the layer in the TOC
        '         5.1 Get the name of the field
        '           6.0 If the field has something in it
        '             7.0 Set the ID of the flag (dOutValue) equal to that value
        '
        Dim bCheck As Boolean
        Dim j, k As Integer

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeature As IFeature
        Dim pFields As IFields
        Dim sNaturalYNField As String
        Dim iBarrierIds As Integer
        Dim sOutNatural As String

        bCheck = False
        For j = 0 To pMap.LayerCount - 1
            If pMap.Layer(j).Valid = True Then
                If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(j), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        pFeatureClass = pFLayer.FeatureClass
                        pFeature = pFeatureClass.GetFeature(iFID)
                        pFields = pFeature.Fields
                        iBarrierIds = lBarrierIDs.Count

                        For k = 0 To iBarrierIds - 1
                            If lBarrierIDs.Item(k).Layer = pFLayer.Name Then
                                sNaturalYNField = lBarrierIDs.Item(k).NaturalYNField
                                If pFields.FindField(sNaturalYNField) <> -1 Then
                                    sOutNatural = Convert.ToString(pFeature.Value(pFields.FindField(sNaturalYNField)))
                                    bCheck = True
                                End If ' names match
                            End If
                        Next ' barrierIDField
                        If bCheck = False Then ' assume it's not natural
                            sOutNatural = "F"
                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        GetNaturalYN = sOutNatural

    End Function
    Private Function GetBarrierID(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As IDandType
        ' =============== FLAG ON POINT WITH BNUMBER? =========================
        ' This section checks whether there is a BNumber Field
        ' In which case it will use this field to identify the 
        ' barrier in the output.
        ' 1.0 For each layer in the table of contents
        '   2.0 If the layer is a FeatureLayer
        '     3.0 If the layer ClassID matches the layer of the CURRENT FLAG
        '     3.1 Get the field values for the feature
        '       4.0 For each of the Barrier IDs in the list
        '         5.0 If the layer in the BarrierID list matches the layer in the TOC
        '         5.1 Get the name of the field
        '           6.0 If the field has something in it
        '             7.0 Set the ID of the flag (sOID) equal to that value
        '
        Dim bCheck As Boolean
        Dim j, k As Integer

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeature As IFeature
        Dim pFields As IFields
        Dim sBarrierIDField As String
        Dim iBarrierIds, iOID As Integer
        Dim sOutID As String
        Dim q As Integer
        Dim sAncillaryRole As String

        bCheck = False
        For j = 0 To pMap.LayerCount - 1
            If pMap.Layer(j).Valid = True Then
                If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(j), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        pFeatureClass = pFLayer.FeatureClass
                        pFeature = pFeatureClass.GetFeature(iFID)
                        pFields = pFeature.Fields
                        iBarrierIds = lBarrierIDs.Count

                        For k = 0 To iBarrierIds - 1
                            If lBarrierIDs.Item(k).Layer = pFLayer.Name Then
                                sBarrierIDField = lBarrierIDs.Item(k).Field
                                If pFields.FindField(sBarrierIDField) <> -1 Then
                                    Try
                                        sOutID = Convert.ToString(pFeature.Value(pFields.FindField(sBarrierIDField)))
                                        bCheck = True
                                    Catch ex As Exception
                                        bCheck = False
                                    End Try
                                End If ' names match
                            End If
                        Next ' barrierIDField
                        If bCheck = False Then ' names don't match - set to default field
                            'sOutID = Convert.ToString(pFeature.Value(pFields.FindField("OBJECTID")))
                            ' If the layer has an ancillaryRole, check to see if the feature
                            q = pFields.FindField("AncillaryRole")
                            If Not q = -1 Then
                                sAncillaryRole = Convert.ToString(pFeature.Value(q))
                                If sAncillaryRole = "2" Then
                                    sOutID = "Sink"
                                    iOID = pFeature.OID.ToString ' Added Feb 18, 2012
                                End If
                            Else
                                sOutID = pFeature.OID.ToString
                            End If
                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        Dim pIDAndType As New IDandType(Nothing, Nothing)

        If sOutID = "Sink" Then
            pIDAndType.BarrID = iOID.ToString
            pIDAndType.BarrIDType = sOutID
        Else
            pIDAndType.BarrID = sOutID
            pIDAndType.BarrIDType = "Barrier"
        End If

        GetBarrierID = pIDAndType

    End Function
    Private Function GetBarrierPerm(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As Double
        ' =============== FLAG ON POINT WITH Perm? =========================
        ' This section checks whether there is a BarrierPerm Field
        ' In which case it will use this field to identify the 
        ' barrier in the output.
        ' 1.0 For each layer in the table of contents
        '   2.0 If the layer is a FeatureLayer
        '     3.0 If the layer ClassID matches the layer of the CURRENT FLAG
        '     3.1 Get the field values for the feature
        '       4.0 For each of the Barrier IDs in the list
        '         5.0 If the layer in the BarrierID list matches the layer in the TOC
        '         5.1 Get the name of the field
        '           6.0 If the field has something in it
        '             7.0 Set the ID of the flag (dOutValue) equal to that value
        '
        Dim bCheck As Boolean
        Dim j, k As Integer

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeature As IFeature
        Dim pFields As IFields
        Dim sBarrierPermField As String
        Dim iBarrierIds As Integer
        Dim dOutValue As Double
        Dim vPermValue As Object


        bCheck = False
        For j = 0 To pMap.LayerCount - 1
            If pMap.Layer(j).Valid = True Then
                If TypeOf pMap.Layer(j) Is IFeatureLayer Then
                    pFLayer = CType(pMap.Layer(j), IFeatureLayer)
                    If pFLayer.FeatureClass.FeatureClassID = iFCID Then

                        pFeatureClass = pFLayer.FeatureClass
                        pFeature = pFeatureClass.GetFeature(iFID)
                        pFields = pFeature.Fields
                        iBarrierIds = lBarrierIDs.Count

                        For k = 0 To iBarrierIds - 1
                            If lBarrierIDs.Item(k).Layer = pFLayer.Name Then
                                'MsgBox("The Layer in TOC: " + CStr(pFLayer.Name))
                                'MsgBox("The Layer in List: " + CStr(lBarrierIDs.Item(k).Layer))

                                sBarrierPermField = lBarrierIDs.Item(k).PermField
                                'MsgBox("Permeability field: " + CStr(sBarrierPermField))

                                If pFields.FindField(sBarrierPermField) <> -1 Then
                                    ''MsgBox("Debug:C12")
                                    vPermValue = pFeature.Value(pFields.FindField(sBarrierPermField))
                                    Try
                                        dOutValue = Convert.ToDouble(vPermValue)
                                    Catch ex As Exception
                                        MsgBox("The Permeability Value in the " & CStr(pFLayer.Name) & " was not convertible" _
                                        & " to type 'double'. Please check the attribute table. Assuming permeability = zero. " & ex.Message.ToString)
                                        ' If there's a null value or field value can't be converted, assume zero
                                        dOutValue = 0.0
                                    End Try
                                    bCheck = True
                                End If ' names match
                            End If
                        Next ' barrierIDField
                        If bCheck = False Then ' names don't match - assume it's a 
                            dOutValue = 0
                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        GetBarrierPerm = dOutValue

    End Function
    Public Sub IntersectFeatures()

        ' 2020 - note this intersects with polygon layers for habitat
        '        based on any _selected_ line layers 
        '        (i.e., network lines returned by trace)
        '        it does not intersect with more line layers 

        'Read Extension Settings
        ' ================== READ EXTENSION SETTINGS =================

        Dim bDBF As Boolean = False         ' Include DBF output default 'no'
        Dim pLLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
        Dim pPLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)
        Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        Dim iLinesCount As Integer = 0      ' number of lines layers currently using

        Dim HabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

        ' object to hold stats to add to list. 
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim m As Integer = 0
        Dim k As Integer = 0
        Dim j As Integer = 0
        Dim i As Integer = 0

        If m_FiPEx__1.m_bLoaded = True Then

            ' Populate a list of the layers using and habitat summary fields.
            ' match any of the polygon layers saved in stream to those in listboxes 
            iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
            If iPolysCount > 0 Then
                For k = 0 To iPolysCount - 1
                    'sPolyLayer = m_FiPEX__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
                    HabLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
                    With HabLayerObj
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
                        .HabUnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
                    End With

                    ' Load that object into the list
                    pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
                Next
            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            Dim iCount1 As Integer = pPLayersFields.Count

            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pPLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu. FIPEX Error code 576. ", "Parameter Missing")
                        Exit Sub
                    End If
                Next
            End If

            iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
            Dim HabLayerObj2 As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

            ' match any of the line layers saved in stream to those in listboxes
            If iLinesCount > 0 Then
                For j = 0 To iLinesCount - 1
                    'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    HabLayerObj2 = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    With HabLayerObj2
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))

                        .LengthField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthField" + j.ToString))
                        .LengthUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthUnits" + j.ToString))

                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabClassField" + j.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabQuanField" + j.ToString))
                        .HabUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabUnits" + j.ToString))
                    End With
                    ' add to the module level list
                    pLLayersFields.Add(HabLayerObj2)
                Next
            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            iCount1 = pLLayersFields.Count
            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pLLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for river layer. Please choose a field in the options menu.", "Parameter Missing")
                        Exit Sub
                    End If
                    If pLLayersFields.Item(m).LengthField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No length field is not set for river layer. Please choose a field in the options menu.", "Parameter Missing")
                        Exit Sub
                    End If
                Next

            End If
        Else
            System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
            Exit Sub
        End If

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
        '             intersect a polygon layer with itself)
        '               If there are any selected features 
        '                 If the parameter array is already populated then empty it
        '                 Populate parameter array for intersect process
        '                 Perform intersect - return results as selection


        Dim pUID As New UID
        ' Get the pUID of the SelectByLayer command
        'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

        Dim pGp As IGeoProcessor
        pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
        Dim pParameterArray As IVariantArray
        Dim pMxDocument As IMxDocument
        Dim pMap As IMap

        Dim sFeatureFullPath As String
        Dim lMaxLayerIndex As Integer
        Dim pLayer2Intersect As IFeatureLayer
        Dim iFieldVal As Integer  ' The field index
        Dim sTestPolygon As String
        Dim bIncludePoly As Boolean 'For polygon inclusion
        Dim pFeatureLayer As IFeatureLayer
        Dim pEnumLayer As IEnumLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pGPResults As IGeoProcessorResult
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDocument = CType(pDoc, IMxDocument)
        pMap = pMxDocument.FocusMap
        lMaxLayerIndex = pMap.LayerCount - 1
        i = 0
        m = 0

        pUID = New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

        pEnumLayer = pMap.Layers(pUID, True)
        pEnumLayer.Reset()

        If Not pPLayersFields Is Nothing Then
            If pPLayersFields.Count > 0 Then
                For i = 0 To lMaxLayerIndex
                    If pMap.Layer(i).Valid = True Then
                        If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                            pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
                            sTestPolygon = Convert.ToString(pFeatureLayer.FeatureClass.ShapeType)

                            If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
                                bIncludePoly = False ' Reset Variable
                                For j = 0 To pPLayersFields.Count - 1
                                    If pFeatureLayer.Name = pPLayersFields(j).Layer Then
                                        bIncludePoly = True
                                    End If
                                Next

                                'sDataSourceType = pFeatureLayer.DataSourceType

                                If bIncludePoly = True Then
                                    m = 0
                                    For m = 0 To lMaxLayerIndex
                                        If pMap.Layer(m).Valid = True Then
                                            If TypeOf pMap.Layer(m) Is IFeatureLayer Then
                                                pLayer2Intersect = CType(pMap.Layer(m), IFeatureLayer)

                                                If pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
                                                Or pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
                                                    pFeatureSelection = CType(pLayer2Intersect, IFeatureSelection)

                                                    If (pFeatureSelection.SelectionSet.Count <> 0) Then

                                                        If pParameterArray IsNot Nothing Then
                                                            pParameterArray.RemoveAll()
                                                        Else
                                                            'pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray 'VB.NET
                                                            pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray
                                                        End If

                                                        ' The GP doesn't need full path names to feature classes
                                                        ' it only needs feature layer names as they appear in the TOC
                                                        sFeatureFullPath = pFeatureLayer.Name
                                                        pParameterArray.Add(sFeatureFullPath)
                                                        pParameterArray.Add("INTERSECT")
                                                        pParameterArray.Add(pLayer2Intersect.Name)
                                                        pParameterArray.Add("#")
                                                        pParameterArray.Add("ADD_TO_SELECTION")

                                                        pGPResults = pGp.Execute("SelectLayerByLocation_management", pParameterArray, Nothing)
                                                    End If ' it has a feature selection
                                                End If ' It's a line
                                            End If ' It's a feature layer
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If ' Layer is valid
                Next
            End If ' polygon list count is not zero
        End If ' polygon list is not nothing

    End Sub


    'Public Sub IntersectFeatures()

    '    'Read Extension Settings
    '    ' ================== READ EXTENSION SETTINGS =================

    '    Dim bDBF As Boolean = False         ' Include DBF output default 'no'
    '    Dim pLLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
    '    Dim pPLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
    '    Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
    '    Dim iLinesCount As Integer = 0      ' number of lines layers currently using
    '    Dim HabLayerObj As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property
    '    ' object to hold stats to add to list. 
    '    Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    '    Dim m As Integer = 0
    '    Dim k As Integer = 0
    '    Dim j As Integer = 0
    '    Dim i As Integer = 0

    '    If m_FiPEx__1.m_bLoaded = True Then

    '        ' Populate a list of the layers using and habitat summary fields.
    '        ' match any of the polygon layers saved in stream to those in listboxes 
    '        iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
    '        If iPolysCount > 0 Then
    '            For k = 0 To iPolysCount - 1
    '                'sPolyLayer = m_DFOExt.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
    '                HabLayerObj = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
    '                With HabLayerObj
    '                    .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
    '                    .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
    '                    .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
    '                    .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
    '                End With

    '                ' Load that object into the list
    '                pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
    '            Next
    '        End If

    '        ' Need to be sure that quantity field has been assigned for each
    '        ' layer using. 
    '        Dim iCount1 As Integer = pPLayersFields.Count

    '        If iCount1 > 0 Then
    '            For m = 0 To iCount1 - 1
    '                If pPLayersFields.Item(m).QuanField = "Not set" Then
    '                    System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
    '                    backgroundworker2.CancelAsync()
    '                    Exit Sub
    '                End If
    '            Next
    '        End If

    '        iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
    '        Dim HabLayerObj2 As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

    '        ' match any of the line layers saved in stream to those in listboxes
    '        If iLinesCount > 0 Then
    '            For j = 0 To iLinesCount - 1
    '                'sLineLayer = m_DFOExt.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
    '                HabLayerObj2 = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
    '                With HabLayerObj2
    '                    '.Layer = sLineLayer
    '                    .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))
    '                    .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineClassField" + j.ToString))
    '                    .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineQuanField" + j.ToString))
    '                    .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineUnitField" + j.ToString))
    '                End With
    '                ' add to the module level list
    '                pLLayersFields.Add(HabLayerObj2)
    '            Next
    '        End If

    '        ' Need to be sure that quantity field has been assigned for each
    '        ' layer using. 
    '        iCount1 = pLLayersFields.Count
    '        If iCount1 > 0 Then
    '            For m = 0 To iCount1 - 1
    '                If pLLayersFields.Item(m).QuanField = "Not set" Then
    '                    System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for river layer. Please choose a field in the options menu.", "Parameter Missing")
    '                    backgroundworker2.CancelAsync()
    '                    Exit Sub
    '                End If
    '            Next
    '        End If
    '    Else
    '        System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
    '        backgroundworker2.CancelAsync()
    '        Exit Sub
    '    End If

    '    ' ======================= 1.0 INTERSECT ===========================
    '    ' This next section checks each of the polygon layers in the focusmap
    '    ' and intersects them with any layers in the focusmap that have a selection.

    '    ' PROCESS LOGIC:
    '    ' If there are polygon layers to use in this process then continue
    '    '   For each of the layers in the focusMap
    '    '     If it's a feature layer then
    '    '       If it's a polygon
    '    '         If it's on the 'include' list
    '    '           For each of the FocusMap layers 
    '    '             If it's a line layer (because we don't want to repeatedly 
    '    '             intersect a polygon layer with itself, do we? DO WEEE???)
    '    '               If there are any selected features 
    '    '                 If the parameter array is already populated then empty it
    '    '                 Populate parameter array for intersect process
    '    '                 Perform intersect - return results as selection


    '    Dim pUID As New UID
    '    ' Get the pUID of the SelectByLayer command
    '    'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

    '    Dim pGp As IGeoProcessor
    '    pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
    '    Dim pParameterArray As IVariantArray
    '    Dim pMxDocument As IMxDocument
    '    Dim pMap As IMap

    '    Dim sFeatureFullPath As String
    '    Dim lMaxLayerIndex As Integer
    '    Dim pLayer2Intersect As IFeatureLayer
    '    Dim iFieldVal As Integer  ' The field index
    '    Dim sTestPolygon As String
    '    Dim bIncludePoly As Boolean 'For polygon inclusion
    '    Dim pFeatureLayer As IFeatureLayer
    '    Dim pEnumLayer As IEnumLayer
    '    Dim pFeatureSelection As IFeatureSelection

    '    Dim pDoc As IDocument = My.ArcMap.Application.Document
    '    pMxDocument = CType(pDoc, IMxDocument)
    '    pMap = pMxDocument.FocusMap
    '    lMaxLayerIndex = pMap.LayerCount - 1
    '    i = 0
    '    m = 0

    '    pUID = New UID
    '    pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

    '    pEnumLayer = pMap.Layers(pUID, True)
    '    pEnumLayer.Reset()

    '    If Not pPLayersFields Is Nothing Then
    '        If pPLayersFields.Count > 0 Then
    '            For i = 0 To lMaxLayerIndex
    '                If pMap.Layer(i).Valid = True Then
    '                    If TypeOf pMap.Layer(i) Is IFeatureLayer Then
    '                        pFeatureLayer = CType(pMap.Layer(i), IFeatureLayer)
    '                        sTestPolygon = Convert.ToString(pFeatureLayer.FeatureClass.ShapeType)

    '                        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
    '                            bIncludePoly = False ' Reset Variable
    '                            For j = 0 To pPLayersFields.Count - 1
    '                                If pFeatureLayer.Name = pPLayersFields(j).Layer Then
    '                                    bIncludePoly = True
    '                                End If
    '                            Next

    '                            'sDataSourceType = pFeatureLayer.DataSourceType

    '                            If bIncludePoly = True Then
    '                                m = 0
    '                                For m = 0 To lMaxLayerIndex
    '                                    If pMap.Layer(m).Valid = True Then
    '                                        If TypeOf pMap.Layer(m) Is IFeatureLayer Then
    '                                            pLayer2Intersect = CType(pMap.Layer(m), IFeatureLayer)

    '                                            If pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine _
    '                                            Or pLayer2Intersect.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
    '                                                pFeatureSelection = CType(pLayer2Intersect, IFeatureSelection)

    '                                                If (pFeatureSelection.SelectionSet.Count <> 0) Then

    '                                                    If pParameterArray IsNot Nothing Then
    '                                                        pParameterArray.RemoveAll()
    '                                                    Else
    '                                                        'pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray 'VB.NET
    '                                                        pParameterArray = New ESRI.ArcGIS.esriSystem.VarArray
    '                                                    End If

    '                                                    ' The GP doesn't need full path names to feature classes
    '                                                    ' it only needs feature layer names as they appear in the TOC
    '                                                    sFeatureFullPath = pFeatureLayer.Name
    '                                                    pParameterArray.Add(sFeatureFullPath)
    '                                                    pParameterArray.Add("INTERSECT")
    '                                                    pParameterArray.Add(pLayer2Intersect.Name)
    '                                                    pParameterArray.Add("#")
    '                                                    pParameterArray.Add("ADD_TO_SELECTION")

    '                                                    pGp.Execute("SelectLayerByLocation_management", pParameterArray, Nothing)
    '                                                End If ' it has a feature selection
    '                                            End If ' It's a line
    '                                        End If ' It's a feature layer
    '                                    End If
    '                                Next
    '                            End If
    '                        End If
    '                    End If
    '                End If ' Layer is valid
    '            Next
    '        End If ' polygon list count is not zero
    '    End If ' polygon list is not nothing
    'End Sub
    Public Function TraceFlowSolverSetup2() As ITraceFlowSolver
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

        ' Get reference to the current network through Utility Network interface
        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)

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
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)
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
        pNetworkAnalysisExtWeightFilter = CType(m_UNAExt, INetworkAnalysisExtWeightFilter)
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
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
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
        pTraceTasks = CType(m_UNAExt, ITraceTasks)
        'pTraceFlowSolver.TraceIndeterminateFlow = pTraceTasks.TraceIndeterminateFlow
        'MsgBox "Before trace indeterminate flow"
        pTraceFlowSolver.TraceIndeterminateFlow = True
        pTraceTasks.TraceIndeterminateFlow = True
        'MsgBox "After Trace indeterminate flow"

        ' pass the traceFlowSolver object back to the network solver
        TraceFlowSolverSetup2 = pTraceFlowSolver
        'Return pTraceFlowSolver

    End Function

    Public Sub calculateStatistics_2(ByRef lHabStatsList As List(Of StatisticsObject_2), _
                                     ByRef ID As String, _
                                     ByRef iEID As Integer, _
                                     ByRef sType As String, _
                                     ByRef f_sOutID As String, _
                                     ByRef f_siOutEID As Integer,
                                     ByVal sHabTypeKeyword As String, _
                                     ByVal sDirection2 As String)

        '2020 note: sKeyword unused. Deleted. -GO
        ' **************************************************************************************
        ' Subroutine:  Calculate Statistics (2) 
        ' Created By:  Greig Oldford
        ' Update Date: October 5, 2010
        ' Purpose:     1) intersect other included layers with returned selection
        '                 from the trace.
        '              2) calculate habitat area and length using habitat classes 
        '                 and excluding unwanted features
        '              3) get array (matrix?) of statistics for each habitat class and
        '                 each layer included for habitat classification stats
        '              4) update statistics object and send back to onclick
        ' Keywords:    sHabTypeKeyword - "Total", "Immediate", or "Path"
        '              sKeyword - "barrier" or "nonbarr"
        '
        '
        ' Notes:
        ' 
        ' 
        '       Aug, 2020    --> Not sure why passing vars by ref other than lHabStatsList.
        '                        The function only changes Hab stats object
        '                        deleted 'sKeyword' arg (checks if flag on barr or nonbarr) - it was unused
        ' 
        ' 
        '       Oct 5, 2010  --> Changing this subroutine to a function so it can update the statistics 
        '                  object for habitat statistics (with classes) ONLY. i.e., there will be no 
        '                  other metrics included in this habitat statistics object.
        '                  Added another keyword to say whether this is TOTAL habitat or otherwise (sHabTypeKeyword). 
        '    
        '       Mar 3, 2008  --> only polygon feature layers are intersected.  The function
        '                  checks the config file for included polygons and will intersect any
        '                  network features returned by the trace with the polygons on the list.
        '                  There is probably no reason to have this explicitly for polygons, and
        '                  dividing the 'includes' list into line and polygon categories means that
        '                  the habitat classification also must be divided as such.  This would double
        '                  the number of variables for this process (polygon habitat class layer
        '                  variable, line hab class lyr var, polygon hab class case field var, etc.)
        '                  So since network feature layers are already being returned by the trace,
        '                  they don't need to be intersected.  If we have one 'includes' list that
        '                  contains both polygon and line layers then we need to find out which layers
        '                  in this list are not part of the geometric network, and only intersect these
        '                  features.
        '                  For each includes feature, For each current geometric feature, find match?  Next
        '                  If no match then continue intersection.


        Dim pMxDoc As IMxDocument
        Dim pEnumLayer As IEnumLayer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        ' Feb 29 --> There will be a variable number of "included" layers
        '            to use for the habitat classification summary tables.
        '            Each table corresponds to "pages" in the matrix.
        '            Matrix(pages, columns, rows)
        '            Only the farthest right element in a matrix can be
        '            redim "preserved" in VB6 meaning there must be a static
        '            number of columns and pages.  Pages isn't a problem.
        '            They will be the number of layers in the "includes" list
        '            Columns, however, will vary.  This is a problem.  They
        '            will vary between pages of the matrix too which means there
        '            will be empty columns on at least one page if the column count
        '            is different between pages.
        '            Answer to this problem is to avoid the matrix altogether and
        '            update the necessary tables within this function


        Dim dTotalArea As Double ' alternative to 2D Matrix when no classes are given
        Dim pUID As New UID
        ' Get the pUID of the SelectByLayer command
        'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

        Dim pGp As IGeoProcessor
        pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
        Dim pMxDocument As IMxDocument
        Dim pMap As IMap


        Dim i, j, k, m As Integer
        Dim iFieldVal As Integer  ' The field index

        Dim pFields As IFields
        Dim vVar As Object
        Dim pSelectionSet As ISelectionSet
        Dim sTemp As String
        Dim sUnit As String

        ' K REPRESENTS NUMBER OF POSSIBLE HABITAT CLASSES
        '  rows, columns.  ROWS SHOULD BE SET BY NUMBER OF SUMMARY FIELDS
        ' cannot be redimension preserved later

        'Dim mHabClassVals() As Object       'Matrix holds the unique classes and values
        'ReDim mHabClassVals(0 To 4, 0 To 400)         'temporary dimension statement
        Dim lHabStatsMatrix As New List(Of HabStatisticsObject)
        Dim pHabStatisticsObject As New HabStatisticsObject(Nothing, Nothing)

        'Dim pFeatureWkSp As IFeatureWorkspace
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor
        Dim vFeatHbClsVl As Object ' Feature Habitat Class Value (an object because classes can be numbers or string)
        Dim vTemp As Object
        Dim sFeatClassVal As String
        Dim sMatrixVal As String
        Dim dHabArea As Double

        Dim bClassFound As Boolean
        'For k = 1 To UBound(mHabClassVals, 2) vb6
        Dim classComparer As FindStatsClassPredicate
        Dim iStatsMatrixIndex As Integer ' for refining statistics list 
        Dim sClass As String
        Dim vHabTemp As Object

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMxDocument = CType(pDoc, IMxDocument)
        pMap = pMxDocument.FocusMap

        ' ================== READ EXTENSION SETTINGS =================
        'Dim sDirection, sDirection2 As String

        Dim bDBF As Boolean = False         ' Include DBF output default 'no'
        Dim pLLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
        Dim pPLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)
        Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        Dim iLinesCount As Integer = 0      ' number of lines layers currently using

        ' 2020 - change this two separate objects, lines polygons
        ' layer to hold parameters to send to property
        Dim PolyHabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
        Dim LineHabLayerObj As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)


        ' object to hold stats to add to list. 
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, _
                                                        Nothing, Nothing, Nothing, Nothing, Nothing, _
                                                        Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim sDirection As String

        If m_FiPEx__1.m_bLoaded = True Then

            sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))
            bDBF = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDBF"))

            ''Make the direction more readable for dockable window output
            'If sDirection = "up" Then
            '    sDirection2 = "upstream"
            'Else
            '    sDirection2 = "downstream"
            'End If

            ' Populate a list of the layers using and habitat summary fields.
            ' match any of the polygon layers saved in stream to those in listboxes 
            iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
            If iPolysCount > 0 Then
                For k = 0 To iPolysCount - 1
                    'sPolyLayer = m_FiPEX__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
                    PolyHabLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
                    With PolyHabLayerObj
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
                        .HabUnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
                    End With

                    ' Load that object into the list
                    pPLayersFields.Add(PolyHabLayerObj)  'what are the brackets about - this could be aproblem!!
                Next
            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            Dim iCount1 As Integer = pPLayersFields.Count

            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pPLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
                        Exit Sub
                    End If
                Next
            End If

            iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))

            ' match any of the line layers saved in stream to those in listboxes
            If iLinesCount > 0 Then
                For j = 0 To iLinesCount - 1
                    'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    LineHabLayerObj = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                    With LineHabLayerObj
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))
                        .LengthField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthField" + j.ToString))
                        .LengthUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthUnits" + j.ToString))

                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabClassField" + j.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabQuanField" + j.ToString))
                        .HabUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabUnits" + j.ToString))
                    End With
                    ' add to the module level list
                    pLLayersFields.Add(LineHabLayerObj)
                Next
            End If

            ' Need to be sure that quantity field has been assigned for each
            ' layer using. 
            iCount1 = pLLayersFields.Count
            If iCount1 > 0 Then
                For m = 0 To iCount1 - 1
                    If pLLayersFields.Item(m).HabQuanField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("Missing habitat quantity field for line layer. Please choose a field in the options menu.", "Parameter Missing")
                        Exit Sub
                    End If
                    If pLLayersFields.Item(m).LengthField = "Not set" Then
                        System.Windows.Forms.MessageBox.Show("Missing length field for line layer. Please choose a field in the options menu.", "Parameter Missing")
                        Exit Sub
                    End If
                Next
            End If
        Else
            System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
            Exit Sub
        End If

        ' ================ 2.0 Calculate Area and Length ======================
        ' This next section calculates the area or length of selected features
        ' in the TOC.
        '
        ' PROCESS LOGIC:
        '  1.0 For each Feature Layer in the map
        '  1.1 Filter out any excluded features
        '  1.2 Get a list of all fields in the layer
        '  1.3 Combine the polygon and line layers into one list
        '      2020 - changed this so they are no longer combined
        '  1.4 Prepare the dockable window 
        '    2.0 For each habitat layer in the new list (polygons and lines)
        '      3.0 If there's a match b/w the current layer and habitat layer in list
        '        4.0 then prepare Dockable Window and DBF tables if need be
        '        4.1 Search for the habitat class field in layer
        '        4.2a If the field is found
        '          5.0a If there is a selection set 
        '            6.0a Get the unique values in that field from the selection set
        '            6.1a Loop through unique values and add each to the left column
        '                of a two-column array/matrix to hold statistics
        '            6.2a For each selected feature in the layer
        '              7.0a Get the value in the habitat class field
        '              7.1a For each unique habitat class value in the statistics matrix
        '                8.0a If it matches the value of the class field found in the current feature
        '                  9.0a then add the value of the quantity field in that feature to the
        '                      quantity field for that row in the matrix
        '        4.2b Else if the habitat class field is not found
        '          5.0b If there is a selection set
        '            6.0b For each feature total up stats
        '          5.1b Send output to dockable window

        pUID = New UID
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

        pEnumLayer = pMap.Layers(pUID, True)
        pEnumLayer.Reset()
        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Dim iClassCheckTemp As Integer
        Dim iLoopCount As Integer = 0
        Dim dTempQuan As Double

        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                pSelectionSet = pFeatureSelection.SelectionSet

                ' get the fields from the featureclass
                pFields = pFeatureLayer.FeatureClass.Fields
                j = 0

                ' Need to combine both polygon and line lists into one
                ' This avoids the need for two loops
                Dim lLineLayersFields As New List(Of LineLayerToAdd)
                For j = 0 To pLLayersFields.Count - 1
                    lLineLayersFields.Add(pLLayersFields(j))
                Next
                Dim lPolyLayersFields As New List(Of PolyLayerToAdd)
                j = 0
                For j = 0 To pPLayersFields.Count - 1
                    lPolyLayersFields.Add(pPLayersFields(j))
                Next

                ' 2020 - separate loops for polys and lines introduced (not most efficient)
                j = 0
                For j = 0 To lLineLayersFields.Count - 1
                    If lLineLayersFields(j).Layer = pFeatureLayer.Name Then

                        ' Get the Units of measure, if any
                        sUnit = lLineLayersFields(j).HabUnits
                        If sUnit = "Metres" Then
                            sUnit = "m"
                        ElseIf sUnit = "Kilometres" Then
                            sUnit = "km"
                        ElseIf sUnit = "Square Metres" Then
                            sUnit = "m^2"
                        ElseIf sUnit = "Feet" Then
                            sUnit = "ft"
                        ElseIf sUnit = "Miles" Then
                            sUnit = "mi"
                        ElseIf sUnit = "Square Miles" Then
                            sUnit = "mi^2"
                        ElseIf sUnit = "Hectares" Then
                            sUnit = "ha"
                        ElseIf sUnit = "Acres" Then
                            sUnit = "ac"
                        ElseIf sUnit = "Hectometres" Then
                            sUnit = "hm"
                        ElseIf sUnit = "Dekametres" Then
                            sUnit = "dm"
                        ElseIf sUnit = "Square Kilometres" Then
                            sUnit = "km^2"
                        Else
                            sUnit = "n/a"
                        End If

                        iClassCheckTemp = pFields.FindField(lLineLayersFields(j).HabClsField)
                        'If pFields.FindField(lLayersFields(j).ClsField) <> -1 Then
                        If iClassCheckTemp <> -1 Then

                            ' Reset the stats objects
                            pDataStats = New DataStatistics
                            pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                            ' Clear the statsMatrix
                            lHabStatsMatrix = New List(Of HabStatisticsObject)

                            If pFeatureSelection.SelectionSet.Count <> 0 Then
                                pSelectionSet.Search(Nothing, False, pCursor)

                                pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                                With pHabStatisticsObject
                                    .UniqueHabClass = "Classes"
                                    .HabQuantity = Nothing '***
                                End With
                                lHabStatsMatrix.Add(pHabStatisticsObject)

                                pSelectionSet.Search(Nothing, False, pCursor) ' THIS LINE MAY BE REDUNDANT (SEE ABOVE)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)

                                pFeature = pFeatureCursor.NextFeature                            ' For each selected feature
                                Do While Not pFeature Is Nothing

                                    ' The habitat class field could be a number or a string
                                    ' so the variable used to hold it is an ambiguous object (variant)
                                    vFeatHbClsVl = pFeature.Value(pFields.FindField(lLineLayersFields(j).HabClsField))

                                    ' Loop through each unique habitat class again
                                    ' and check if it matches the class value of the feature
                                    k = 1
                                    bClassFound = False
                                    iStatsMatrixIndex = 0

                                    Try
                                        sClass = Convert.ToString(vFeatHbClsVl)
                                    Catch ex As Exception
                                        MsgBox("The Habitat Class found in the " & lLineLayersFields(j).Layer & " was not convertible" _
                                        & " to type 'string'.  " & ex.Message)
                                        sClass = "not set"
                                    End Try
                                    If sClass = "" Then
                                        sClass = "not set"
                                    End If

                                    vHabTemp = pFeature.Value(pFields.FindField(lLineLayersFields(j).HabQuanField))

                                    Try
                                        dHabArea = Convert.ToDouble(vHabTemp)
                                    Catch ex As Exception
                                        MsgBox("The Habitat Quantity found in the " & lLineLayersFields(j).Layer & " was not convertible" _
                                        & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                        dHabArea = 0
                                    End Try

                                    classComparer = New FindStatsClassPredicate(sClass)
                                    ' use the layer and sink ID to get a refined list of habitat stats for 
                                    ' this sink, layer combo
                                    iStatsMatrixIndex = lHabStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
                                    If iStatsMatrixIndex = -1 Then
                                        bClassFound = False
                                        pHabStatisticsObject = New HabStatisticsObject(sClass, dHabArea)
                                        lHabStatsMatrix.Add(pHabStatisticsObject)
                                    Else
                                        bClassFound = True
                                        lHabStatsMatrix(iStatsMatrixIndex).HabQuantity = lHabStatsMatrix(iStatsMatrixIndex).HabQuantity + dHabArea
                                    End If

                                    pFeature = pFeatureCursor.NextFeature

                                Loop     ' selected feature
                            End If ' There is a selection set


                            ' If there are items in the stats matrix
                            If lHabStatsMatrix.Count <> 0 Then
                                k = 1
                                ' For each unique value in the matrix
                                ' (always skip first row of matrix as it is the 'column headings')
                                For k = 1 To lHabStatsMatrix.Count - 1
                                    'If bDBF = True Then

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "habitat"
                                        .HabitatDimension = "length"
                                        .TotalImmedPath = sHabTypeKeyword
                                        .UniqueClass = CStr(lHabStatsMatrix(k).UniqueHabClass)
                                        .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                        .Quantity = lHabStatsMatrix(k).HabQuantity
                                        .Unit = sUnit
                                    End With
                                    lHabStatsList.Add(pHabStatsObject_2)

                                Next


                            Else ' If there are no statistics

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "length"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "none"
                                    .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                            End If ' There are items in the statsmatrix

                        Else   ' if the habitat class case field is not found

                            dTotalArea = 0

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)

                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature

                                ' Get the summary field and add the value to the
                                ' total for habitat area.
                                ' ** ==> Multiple fields could be added here in a 'for' loop.

                                iFieldVal = pFeatureCursor.FindField(lLineLayersFields(j).HabQuanField)

                                ' For each selected feature
                                m = 1
                                Do While Not pFeature Is Nothing
                                    Try
                                        vTemp = pFeature.Value(iFieldVal)
                                    Catch ex As Exception
                                        MsgBox("Could not convert quantity field found in " + lLineLayersFields(j).Layer.ToString + _
                                               " was not convertible to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(vTemp)
                                    Catch ex As Exception
                                        MsgBox("Could not convert the habitat quantity value found in the" + _
                                        lLineLayersFields(j).Layer.ToString + ". Was not convertable to type 'double'." + _
                                        ex.Message)
                                        dTempQuan = 0
                                    End Try
                                    ' Insert into the corresponding column of the second
                                    ' row the updated habitat area measurement.
                                    dTotalArea = dTotalArea + dTempQuan
                                    pFeature = pFeatureCursor.NextFeature
                                Loop     ' selected feature
                            End If ' there are selected features

                            ' If DBF tables are to be output
                            'If bDBF = True Then

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "habitat"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "none"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dTotalArea
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)


                        End If ' found habitat class field in layer

                        ' Insert loop for retrieving line length statistics
                        ' 



                        ' increment the loop counter for
                        iLoopCount = iLoopCount + 1

                    End If     ' feature layer matches hab class layer
                Next    ' line layer

                ' ##########################################################
                ' ##########################################################
                'Sep 2020 - separate loops for polys and lines.
                '       not efficient, but easiest and running short on time. 

                j = 0
                For j = 0 To lPolyLayersFields.Count - 1
                    If lPolyLayersFields(j).Layer = pFeatureLayer.Name Then

                        ' Get the Units of measure, if any
                        sUnit = lPolyLayersFields(j).HabUnitField
                        If sUnit = "Metres" Then
                            sUnit = "m"
                        ElseIf sUnit = "Kilometres" Then
                            sUnit = "km"
                        ElseIf sUnit = "Square Metres" Then
                            sUnit = "m^2"
                        ElseIf sUnit = "Feet" Then
                            sUnit = "ft"
                        ElseIf sUnit = "Miles" Then
                            sUnit = "mi"
                        ElseIf sUnit = "Square Miles" Then
                            sUnit = "mi^2"
                        ElseIf sUnit = "Hectares" Then
                            sUnit = "ha"
                        ElseIf sUnit = "Acres" Then
                            sUnit = "ac"
                        ElseIf sUnit = "Hectometres" Then
                            sUnit = "hm"
                        ElseIf sUnit = "Dekametres" Then
                            sUnit = "dm"
                        ElseIf sUnit = "Square Kilometres" Then
                            sUnit = "km^2"
                        Else
                            sUnit = "n/a"
                        End If

                        iClassCheckTemp = pFields.FindField(lPolyLayersFields(j).HabClsField)
                        'If pFields.FindField(lLayersFields(j).ClsField) <> -1 Then
                        If iClassCheckTemp <> -1 Then

                            ' Reset the stats objects
                            pDataStats = New DataStatistics
                            pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                            ' Clear the statsMatrix
                            lHabStatsMatrix = New List(Of HabStatisticsObject)

                            If pFeatureSelection.SelectionSet.Count <> 0 Then
                                pSelectionSet.Search(Nothing, False, pCursor)

                                pHabStatisticsObject = New HabStatisticsObject(Nothing, Nothing)
                                With pHabStatisticsObject
                                    .UniqueHabClass = "Classes"
                                    .HabQuantity = Nothing '***
                                End With
                                lHabStatsMatrix.Add(pHabStatisticsObject)

                                pSelectionSet.Search(Nothing, False, pCursor) ' THIS LINE MAY BE REDUNDANT (SEE ABOVE)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)

                                pFeature = pFeatureCursor.NextFeature                            ' For each selected feature
                                Do While Not pFeature Is Nothing
                                    'pFields = pFeature.Fields  '** removed because should be redundant

                                    ' The habitat class field could be a number or a string
                                    ' so the variable used to hold it is an ambiguous object (variant)
                                    vFeatHbClsVl = pFeature.Value(pFields.FindField(lPolyLayersFields(j).HabClsField))

                                    ' Loop through each unique habitat class again
                                    ' and check if it matches the class value of the feature
                                    k = 1
                                    bClassFound = False
                                    iStatsMatrixIndex = 0

                                    Try
                                        sClass = Convert.ToString(vFeatHbClsVl)
                                    Catch ex As Exception
                                        MsgBox("The Habitat Class found in the " & lPolyLayersFields(j).Layer & " was not convertible" _
                                        & " to type 'string'.  " & ex.Message)
                                        sClass = "not set"
                                    End Try
                                    If sClass = "" Then
                                        sClass = "not set"
                                    End If

                                    vHabTemp = pFeature.Value(pFields.FindField(lPolyLayersFields(j).HabQuanField))

                                    Try
                                        dHabArea = Convert.ToDouble(vHabTemp)
                                    Catch ex As Exception
                                        MsgBox("The Habitat Quantity found in the " & lPolyLayersFields(j).Layer & " was not convertible" _
                                        & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                        dHabArea = 0
                                    End Try

                                    classComparer = New FindStatsClassPredicate(sClass)
                                    ' use the layer and sink ID to get a refined list of habitat stats for 
                                    ' this sink, layer combo
                                    iStatsMatrixIndex = lHabStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
                                    If iStatsMatrixIndex = -1 Then
                                        bClassFound = False
                                        pHabStatisticsObject = New HabStatisticsObject(sClass, dHabArea)
                                        lHabStatsMatrix.Add(pHabStatisticsObject)
                                    Else
                                        bClassFound = True
                                        lHabStatsMatrix(iStatsMatrixIndex).HabQuantity = lHabStatsMatrix(iStatsMatrixIndex).HabQuantity + dHabArea
                                    End If

                                    pFeature = pFeatureCursor.NextFeature

                                Loop     ' selected feature
                            End If ' There is a selection set


                            ' If there are items in the stats matrix
                            If lHabStatsMatrix.Count <> 0 Then
                                k = 1
                                ' For each unique value in the matrix
                                ' (always skip first row of matrix as it is the 'column headings')
                                For k = 1 To lHabStatsMatrix.Count - 1
                                    'If bDBF = True Then

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing, Nothing, _
                                                                               Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "habitat"
                                        .HabitatDimension = "area"
                                        .TotalImmedPath = sHabTypeKeyword
                                        .UniqueClass = CStr(lHabStatsMatrix(k).UniqueHabClass)
                                        .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                        .Quantity = lHabStatsMatrix(k).HabQuantity
                                        .Unit = sUnit
                                    End With
                                    lHabStatsList.Add(pHabStatsObject_2)

                                    If sUnit = "n/a" Then
                                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                        'pResultsForm.txtRichResults.AppendText("         " & lStatsMatrix(k).UniqueClass & "    " & Format(lStatsMatrix(k).Quantity, "0.00") & Environment.NewLine)
                                    Else
                                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                        'pResultsForm.txtRichResults.AppendText("         " & lStatsMatrix(k).UniqueClass & "    " & Format(lStatsMatrix(k).Quantity, "0.00") & sUnit & Environment.NewLine)
                                    End If
                                Next


                            Else ' If there are no statistics

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing, Nothing, _
                                                                           Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "area"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "none"
                                    .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                                '                            End If

                                ' If no stats found then add zeros
                                If sUnit = "n/a" Then
                                    'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                    'pResultsForm.txtRichResults.AppendText("         none    0.00" & Environment.NewLine)
                                Else
                                    'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                    'pResultsForm.txtRichResults.AppendText("         none    0.00" & sUnit & Environment.NewLine)
                                End If
                            End If ' There are items in the statsmatrix

                        Else   ' if the habitat class case field is not found

                            dTotalArea = 0

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)

                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature

                                ' Get the summary field and add the value to the
                                ' total for habitat area.
                                ' ** ==> Multiple fields could be added here in a 'for' loop.

                                iFieldVal = pFeatureCursor.FindField(lPolyLayersFields(j).HabQuanField)

                                ' For each selected feature
                                m = 1
                                Do While Not pFeature Is Nothing
                                    Try
                                        vTemp = pFeature.Value(iFieldVal)
                                    Catch ex As Exception
                                        MsgBox("Could not convert quantity field found in " + lPolyLayersFields(j).Layer.ToString + _
                                               " was not convertible to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(vTemp)
                                    Catch ex As Exception
                                        MsgBox("Could not convert the habitat quantity value found in the" + _
                                        lPolyLayersFields(j).Layer.ToString + ". Was not convertable to type 'double'." + _
                                        ex.Message)
                                        dTempQuan = 0
                                    End Try
                                    ' Insert into the corresponding column of the second
                                    ' row the updated habitat area measurement.
                                    dTotalArea = dTotalArea + dTempQuan
                                    pFeature = pFeatureCursor.NextFeature
                                Loop     ' selected feature
                            End If ' there are selected features

                            ' If DBF tables are to be output
                            'If bDBF = True Then

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing, Nothing, _
                                                                       Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "habitat"
                                .HabitatDimension = "area"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "none"
                                .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                .Quantity = dTotalArea
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)


                            If sUnit = "n/a" Then
                                'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                'pResultsForm.txtRichResults.AppendText("        " & Format(dTotalArea, "0.00") & Environment.NewLine)
                            Else
                                'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
                                'pResultsForm.txtRichResults.AppendText("        " & Format(dTotalArea, "0.00") & sUnit & Environment.NewLine)
                            End If


                        End If ' found habitat class field in layer

                        ' increment the loop counter for
                        iLoopCount = iLoopCount + 1

                    End If     ' feature layer matches hab class layer
                Next    ' poly layer



            End If ' featurelayer is valid
            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Loop

    End Sub


    'Public Sub calculateStatistics_2(ByRef lHabStatsList As List(Of StatisticsObject_2), ByRef ID As String, _
    '      ByRef iEID As Integer, ByRef sType As String, ByRef f_sOutID As String, ByRef f_siOutEID As Integer, _
    '      ByVal sKeyword As String, ByVal sHabTypeKeyword As String, _
    '      ByVal sDirection2 As String)

    '    ' **************************************************************************************
    '    ' Subroutine:  Calculate Statistics (2) 
    '    ' Created By:  Greig Oldford
    '    ' Update Date: October 5, 2010
    '    ' Purpose:     1) To intersect other included layers with returned selection
    '    '                 from the trace.
    '    '              2) calculate habitat area and length using habitat classes 
    '    '                 and excluding unwanted features
    '    '              3) get array (matrix?) of statistics for each habitat class and
    '    '                 each layer included for habitat classification stats
    '    '              4) update statistics object and send back to onclick
    '    ' Keywords:    sHabTypeKeyword - "Total", "Immediate", or "Path"
    '    '              sKeyword - "barrier" or "nonbarr"
    '    '
    '    '
    '    ' Notes:
    '    '       Oct 5, 2010  --> Changing this subroutine to a function so it can update the statistics 
    '    '                  object for habitat statistics (with classes) ONLY. i.e., there will be no 
    '    '                  other metrics included in this habitat statistics object.
    '    '                  Added another keyword to say whether this is TOTAL habitat or otherwise (sHabTypeKeyword). 
    '    '    
    '    '       Mar 3, 2008  --> Currently, only polygon feature layers are intersected.  The function
    '    '                  checks the config file for included polygons and will intersect any
    '    '                  network features returned by the trace with the polygons on the list.
    '    '                  There is probably no reason to have this explicitly for polygons, and
    '    '                  dividing the 'includes' list into line and polygon categories means that
    '    '                  the habitat classification also must be divided as such.  This would double
    '    '                  the number of variables for this process (polygon habitat class layer
    '    '                  variable, line hab class lyr var, polygon hab class case field var, etc.)
    '    '                  So, since network feature layers are already being returned by the trace,
    '    '                  they don't need to be intersected.  If we have one 'includes' list that
    '    '                  contains both polygon and line layers then we need to find out which layers
    '    '                  in this list are not part of the geometric network, and only intersect these
    '    '                  features.
    '    '                  For each includes feature, For each current geometric feature, find match?  Next
    '    '                  If no match then continue intersection.


    '    Dim pMxDoc As IMxDocument
    '    Dim pEnumLayer As IEnumLayer
    '    Dim pFeatureLayer As IFeatureLayer
    '    Dim pFeatureSelection As IFeatureSelection
    '    Dim pFeatureCursor As IFeatureCursor
    '    Dim pFeature As IFeature
    '    ' Feb 29 --> There will be a variable number of "included" layers
    '    '            to use for the habitat classification summary tables.
    '    '            Each table corresponds to "pages" in the matrix.
    '    '            Matrix(pages, columns, rows)
    '    '            Only the farthest right element in a matrix can be
    '    '            redim "preserved" in VB6 meaning there must be a static
    '    '            number of columns and pages.  Pages isn't a problem.
    '    '            They will be the number of layers in the "includes" list
    '    '            Columns, however, will vary.  This is a problem.  They
    '    '            will vary between pages of the matrix too which means there
    '    '            will be empty columns on at least one page if the column count
    '    '            is different between pages.
    '    '            Answer to this problem is to avoid the matrix altogether and
    '    '            update the necessary tables within this function


    '    Dim dTotalArea As Double ' alternative to 2D Matrix when no classes are given
    '    Dim pUID As New UID
    '    ' Get the pUID of the SelectByLayer command
    '    'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

    '    Dim pGp As IGeoProcessor
    '    pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
    '    Dim pMxDocument As IMxDocument
    '    Dim pMap As IMap


    '    Dim i, j, k, m As Integer
    '    Dim iFieldVal As Integer  ' The field index

    '    Dim pFields As IFields
    '    Dim vVar As Object
    '    Dim pSelectionSet As ISelectionSet
    '    Dim sTemp As String
    '    Dim sUnit As String

    '    ' K REPRESENTS NUMBER OF POSSIBLE HABITAT CLASSES
    '    '  rows, columns.  ROWS SHOULD BE SET BY NUMBER OF SUMMARY FIELDS
    '    ' cannot be redimension preserved later

    '    'Dim mHabClassVals() As Object       'Matrix holds the unique classes and values
    '    'ReDim mHabClassVals(0 To 4, 0 To 400)         'temporary dimension statement
    '    Dim lStatsMatrix As New List(Of HabStatisticsObject)
    '    Dim pStatisticsObject As New HabStatisticsObject(Nothing, Nothing)

    '    'Dim pFeatureWkSp As IFeatureWorkspace
    '    Dim pDataStats As IDataStatistics
    '    Dim pCursor As ICursor
    '    Dim vFeatHbClsVl As Object ' Feature Habitat Class Value (an object because classes can be numbers or string)
    '    Dim vTemp As Object
    '    Dim sFeatClassVal As String
    '    Dim sMatrixVal As String
    '    Dim dHabArea As Double

    '    Dim bClassFound As Boolean
    '    'For k = 1 To UBound(mHabClassVals, 2) vb6
    '    Dim classComparer As FindStatsClassPredicate
    '    Dim iStatsMatrixIndex As Integer ' for refining statistics list 
    '    Dim sClass As String
    '    Dim vHabTemp As Object

    '    Dim pDoc As IDocument = My.ArcMap.Application.Document
    '    pMxDoc = CType(pDoc, IMxDocument)
    '    pMxDocument = CType(pDoc, IMxDocument)
    '    pMap = pMxDocument.FocusMap

    '    ' ================== READ EXTENSION SETTINGS =================
    '    ''MsgBox("Debug:C1")

    '    Dim bDBF As Boolean = False         ' Include DBF output default 'no'
    '    Dim pLLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
    '    Dim pPLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
    '    Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
    '    Dim iLinesCount As Integer = 0      ' number of lines layers currently using
    '    Dim HabLayerObj As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property
    '    ' object to hold stats to add to list. 
    '    Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    '    Dim sDirection As String

    '    If m_FiPEx__1.m_bLoaded = True Then

    '        sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))
    '        bDBF = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDBF"))

    '        ''Make the direction more readable for dockable window output
    '        'If sDirection = "up" Then
    '        '    sDirection2 = "upstream"
    '        'Else
    '        '    sDirection2 = "downstream"
    '        'End If

    '        ' Populate a list of the layers using and habitat summary fields.
    '        ' match any of the polygon layers saved in stream to those in listboxes 
    '        iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
    '        If iPolysCount > 0 Then
    '            For k = 0 To iPolysCount - 1
    '                'sPolyLayer = m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
    '                HabLayerObj = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
    '                With HabLayerObj
    '                    .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
    '                    .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
    '                    .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
    '                    .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
    '                End With

    '                ' Load that object into the list
    '                pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
    '            Next
    '        End If

    '        ' Need to be sure that quantity field has been assigned for each
    '        ' layer using. 
    '        Dim iCount1 As Integer = pPLayersFields.Count

    '        If iCount1 > 0 Then
    '            For m = 0 To iCount1 - 1
    '                If pPLayersFields.Item(m).QuanField = "Not set" Then
    '                    System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for lacustrine layer. Please choose a field in the options menu.", "Parameter Missing")
    '                    backgroundworker2.CancelAsync()
    '                    Exit Sub
    '                End If
    '            Next
    '        End If

    '        iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
    '        Dim HabLayerObj2 As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

    '        ' match any of the line layers saved in stream to those in listboxes
    '        If iLinesCount > 0 Then
    '            For j = 0 To iLinesCount - 1
    '                'sLineLayer = m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
    '                HabLayerObj2 = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
    '                With HabLayerObj2
    '                    '.Layer = sLineLayer
    '                    .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))
    '                    .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineClassField" + j.ToString))
    '                    .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineQuanField" + j.ToString))
    '                    .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineUnitField" + j.ToString))
    '                End With
    '                ' add to the module level list
    '                pLLayersFields.Add(HabLayerObj2)
    '            Next
    '        End If

    '        ' Need to be sure that quantity field has been assigned for each
    '        ' layer using. 
    '        iCount1 = pLLayersFields.Count
    '        If iCount1 > 0 Then
    '            For m = 0 To iCount1 - 1
    '                If pLLayersFields.Item(m).QuanField = "Not set" Then
    '                    System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for river layer. Please choose a field in the options menu.", "Parameter Missing")
    '                    backgroundworker2.CancelAsync()
    '                    Exit Sub
    '                End If
    '            Next
    '        End If
    '    Else
    '        System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
    '        backgroundworker2.CancelAsync()
    '        Exit Sub
    '    End If

    '    ' ======================== PREPARE DOCKABLE WINDOW FOR OUTPUT =================
    '    'Dim pDockWin As IDockableWindow
    '    'Dim pDSTDockWin As IDockableWindowDef
    '    'Dim pDockWinMgr As IDockableWindowManager
    '    'Dim containedBox As System.Windows.Forms.ListBox
    '    'pDockWinMgr = CType(m_application, IDockableWindowManager) 'QI

    '    'Dim u As New UID


    '    'If bDockWin = True Then
    '    '    u.Value = "{5904bd54-d8ec-4dd8-b1e1-9adcb6558d26}"

    '    '    pDockWin = pDockWinMgr.GetDockableWindow(u)
    '    '    pDockWin.Show(True)
    '    '    pDockWin.Dock(esriDockFlags.esriDockShow)

    '    '    ' If pDockwin works, try to link in and clear the listbox
    '    '    ' sample code from http://edndoc.esri.com/arcobjects/9.2/NET/ViewCodePages/e439cf8c-778b-4fdb-b4c4-fab51a546ac6ClearLoggingCommand.vb.htm
    '    '    If pDockWin IsNot Nothing Then
    '    '        containedBox = TryCast(pDockWin.UserData, System.Windows.Forms.ListBox)
    '    '    End If

    '    'End If
    '    ' ======================== END OLD DOCKWIN CODE =================================

    '    ' ================ 2.0 Calculate Area and Length ======================
    '    ' This next section calculates the area or length of selected features
    '    ' in the TOC.
    '    '
    '    ' PROCESS LOGIC:
    '    '  1.0 For each Feature Layer in the map
    '    '  1.1 Filter out any excluded features
    '    '  1.2 Get a list of all fields in the layer
    '    '  1.3 Combine the polygon and line layers into one list
    '    '  1.4 Prepare the dockable window 
    '    '    2.0 For each habitat layer in the new list (polygons and lines)
    '    '      3.0 If there's a match b/w the current layer and habitat layer in list
    '    '        4.0 then prepare Dockable Window and DBF tables if need be
    '    '        4.1 Search for the habitat class field in layer
    '    '        4.2a If the field is found
    '    '          5.0a If there is a selection set 
    '    '            6.0a Get the unique values in that field from the selection set
    '    '            6.1a Loop through unique values and add each to the left column
    '    '                of a two-column array/matrix to hold statistics
    '    '            6.2a For each selected feature in the layer
    '    '              7.0a Get the value in the habitat class field
    '    '              7.1a For each unique habitat class value in the statistics matrix
    '    '                8.0a If it matches the value of the class field found in the current feature
    '    '                  9.0a then add the value of the quantity field in that feature to the
    '    '                      quantity field for that row in the matrix
    '    '        4.2b Else if the habitat class field is not found
    '    '          5.0b If there is a selection set
    '    '            6.0b For each feature total up stats
    '    '          5.1b Send output to dockable window

    '    ''MsgBox("Debug:C2")

    '    pUID = New UID
    '    pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"

    '    pEnumLayer = pMap.Layers(pUID, True)
    '    pEnumLayer.Reset()
    '    pEnumLayer.Reset()

    '    ' Look at the next layer in the list
    '    pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
    '    Dim iClassCheckTemp As Integer
    '    Dim iLoopCount As Integer = 0
    '    Dim dTempQuan As Double

    '    ''MsgBox("Debug:C3")
    '    Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
    '        If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

    '            pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
    '            pSelectionSet = pFeatureSelection.SelectionSet

    '            ' get the fields from the featureclass
    '            pFields = pFeatureLayer.FeatureClass.Fields
    '            j = 0

    '            ' Need to combine both polygon and line lists into one
    '            ' This avoids the need for two loops
    '            Dim lLayersFields As New List(Of LayerToAdd)
    '            For j = 0 To pLLayersFields.Count - 1
    '                lLayersFields.Add(pLLayersFields(j))
    '            Next
    '            j = 0
    '            For j = 0 To pPLayersFields.Count - 1
    '                lLayersFields.Add(pPLayersFields(j))
    '            Next

    '            ''MsgBox("Debug:C4")
    '            j = 0
    '            For j = 0 To lLayersFields.Count - 1
    '                If lLayersFields(j).Layer = pFeatureLayer.Name Then

    '                    ' Get the Units of measure, if any
    '                    sUnit = lLayersFields(j).UnitField

    '                    If sUnit = "Metres" Then
    '                        sUnit = "m"
    '                    ElseIf sUnit = "Kilometres" Then
    '                        sUnit = "km"
    '                    ElseIf sUnit = "Square Metres" Then
    '                        sUnit = "m^2"
    '                    ElseIf sUnit = "Feet" Then
    '                        sUnit = "ft"
    '                    ElseIf sUnit = "Miles" Then
    '                        sUnit = "mi"
    '                    ElseIf sUnit = "Square Miles" Then
    '                        sUnit = "mi^2"
    '                    ElseIf sUnit = "Hectares" Then
    '                        sUnit = "ha"
    '                    ElseIf sUnit = "Acres" Then
    '                        sUnit = "ac"
    '                    ElseIf sUnit = "Hectometres" Then
    '                        sUnit = "hm"
    '                    ElseIf sUnit = "Dekametres" Then
    '                        sUnit = "dm"
    '                    ElseIf sUnit = "Square Kilometres" Then
    '                        sUnit = "km^2"
    '                    Else
    '                        sUnit = "n/a"
    '                    End If

    '                    ''MsgBox("Debug:C5")
    '                    iClassCheckTemp = pFields.FindField(lLayersFields(j).ClsField)
    '                    'If pFields.FindField(lLayersFields(j).ClsField) <> -1 Then
    '                    If iClassCheckTemp <> -1 Then

    '                        ''MsgBox("Debug:C6")
    '                        ' Reset the stats objects
    '                        pDataStats = New DataStatistics
    '                        pStatisticsObject = New StatisticsObject(Nothing, Nothing)
    '                        ' Clear the statsMatrix
    '                        lStatsMatrix = New List(Of StatisticsObject)

    '                        If pFeatureSelection.SelectionSet.Count <> 0 Then

    '                            ''MsgBox("Debug:C7")
    '                            pSelectionSet.Search(Nothing, False, pCursor)

    '                            '' Setup the datastatistics and get the unique values of the "Id" field
    '                            'Dim pEnum As IEnumerator
    '                            ''System.Windows.Forms.MessageBox.Show(pSelectionSet.Count.ToString)
    '                            'With pDataStats
    '                            '    .Cursor = pCursor
    '                            '    .Field = lLayersFields(j).ClsField
    '                            '    '.Field = "Strahler"
    '                            'End With

    '                            'pEnum = pDataStats.UniqueValues
    '                            ''Dim sTemp As String = pDataStats.Field
    '                            ''Try
    '                            ''    pEnumVar = CType(pEnum, IEnumVariantSimple)
    '                            ''Catch ex As Exception
    '                            ''    System.Windows.Forms.MessageBox.Show(ex.Message.ToString, "Initialize")
    '                            ''End Try
    '                            '' Add the top left corner label to matrix
    '                            ''mHabClassVals(0, 0) = "Classes"

    '                            ''MsgBox("Debug:C8")
    '                            pStatisticsObject = New StatisticsObject(Nothing, Nothing)
    '                            With pStatisticsObject
    '                                .UniqueClass = "Classes"
    '                                .Quantity = Nothing '***
    '                            End With
    '                            lStatsMatrix.Add(pStatisticsObject)


    '                            ''MsgBox("Debug:C9")

    '                            '' Loop through the unique classes and add them to the matrix
    '                            'pEnum.Reset()
    '                            'pEnum.MoveNext()
    '                            'vVar = pEnum.Current

    '                            'While vVar IsNot Nothing        'populate the column headings
    '                            '    'mHabClassVals(0, k) = vVar 'of the matrix with hab classes
    '                            '    pStatisticsObject = New StatisticsObject(vVar.ToString, Nothing)
    '                            '    ' Add a row to the list for each unique class
    '                            '    ' Leave quantities empty for now
    '                            '    'With pStatisticsObject
    '                            '    '    .UniqueClass = vVar.ToString
    '                            '    '    .Quantity = Nothing  '****
    '                            '    'End With
    '                            '    lStatsMatrix.Add(pStatisticsObject)
    '                            '    sTemp = pStatisticsObject.UniqueClass

    '                            '    pEnum.MoveNext()
    '                            '    vVar = pEnum.Current
    '                            'End While

    '                            ''For k = 0 To lStatsMatrix.Count - 1
    '                            ''    sTemp = lStatsMatrix.Item(k).UniqueClass
    '                            ''Next
    '                            ''ReDim Preserve mHabClassVals(4, k - 1) ' redimension columns of matrix

    '                            ''MsgBox("Debug:C10")
    '                            pSelectionSet.Search(Nothing, False, pCursor) ' THIS LINE MAY BE REDUNDANT (SEE ABOVE)
    '                            pFeatureCursor = CType(pCursor, IFeatureCursor)

    '                            ''MsgBox("Debug:C11")
    '                            pFeature = pFeatureCursor.NextFeature                            ' For each selected feature
    '                            Do While Not pFeature Is Nothing
    '                                'pFields = pFeature.Fields  ** removed because should be redundant

    '                                ' The habitat class field could be a number or a string
    '                                ' so the variable used to hold it is an ambiguous object (variant)
    '                                vFeatHbClsVl = pFeature.Value(pFields.FindField(lLayersFields(j).ClsField))
    '                                ''vFeatHbClsVl = pFeature.Value(pFields.FindField("Strahler"))
    '                                'Try
    '                                '    dHabArea = Convert.ToDouble(pFeature.Value(pFields.FindField(lLayersFields(j).QuanField)))
    '                                'Catch ex As Exception
    '                                '    MsgBox("No habitat quantity was found for a feature")
    '                                '    dHabArea = 0
    '                                'End Try

    '                                ' Loop through each unique habitat class again
    '                                ' and check if it matches the class value of the feature
    '                                k = 1
    '                                bClassFound = False
    '                                iStatsMatrixIndex = 0

    '                                ''MsgBox("Debug:C12")
    '                                Try
    '                                    sClass = Convert.ToString(vFeatHbClsVl)
    '                                Catch ex As Exception
    '                                    MsgBox("The Habitat Class found in the " & lLayersFields(j).Layer & " was not convertible" _
    '                                    & " to type 'string'.  " & ex.Message)
    '                                End Try
    '                                If sClass = "" Then
    '                                    sClass = "not set"
    '                                End If

    '                                ''MsgBox("Debug:C13")
    '                                Try
    '                                    vHabTemp = pFeature.Value(pFields.FindField(lLayersFields(j).QuanField))
    '                                Catch ex As Exception
    '                                    vHabTemp = 0
    '                                End Try

    '                                Try
    '                                    dHabArea = Convert.ToDouble(vHabTemp)
    '                                Catch ex As Exception
    '                                    MsgBox("The Habitat Quantity found in the " & lLayersFields(j).Layer & " was not convertible" _
    '                                    & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
    '                                End Try

    '                                ''MsgBox("Debug:C14")
    '                                classComparer = New FindStatsClassPredicate(sClass)
    '                                ' use the layer and sink ID to get a refined list of habitat stats for 
    '                                ' this sink, layer combo
    '                                iStatsMatrixIndex = lStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
    '                                If iStatsMatrixIndex = -1 Then
    '                                    bClassFound = False
    '                                    pStatisticsObject = New StatisticsObject(sClass, dHabArea)
    '                                    lStatsMatrix.Add(pStatisticsObject)
    '                                Else
    '                                    bClassFound = True
    '                                    lStatsMatrix(iStatsMatrixIndex).Quantity = lStatsMatrix(iStatsMatrixIndex).Quantity + dHabArea
    '                                End If

    '                                ''MsgBox("Debug:C15")
    '                                ' Check if the unique class exists in the object list yet
    '                                ' if it doesn't trigger flag or add it... 

    '                                '' k starts at 1 because the first item in the list will just be column
    '                                '' headings
    '                                'For k = 1 To lStatsMatrix.Count - 1
    '                                '    'vMatClmnVl = mHabClassVals(0, k) vb6
    '                                '    sMatrixVal = lStatsMatrix(k).UniqueClass
    '                                '    If Len(sFeatClassVal) <> 0 And Len(sMatrixVal) <> 0 Then
    '                                '        If sMatrixVal = sFeatClassVal Then

    '                                '            ' Get the summary field and add the value to the
    '                                '            ' total for habitat area.
    '                                '            ' ** ==> Multiple fields could be added here in a 'for' loop.
    '                                '            'lFieldVal = pFields.FindField(m_aHabSumFld(j))
    '                                '            ' Insert into the corresponding column of the second
    '                                '            ' row the updated habitat area measurement.
    '                                '            'mHabClassVals(1, k) = mHabClassVals(1, k) + lHabArea vb6
    '                                '            lStatsMatrix(k).Quantity = lStatsMatrix(k).Quantity + dHabArea
    '                                '            sTemp = lStatsMatrix(k).UniqueClass.ToString
    '                                '        End If
    '                                '        'ElseIf Len(sFeatClassVal) <>  Then
    '                                '    End If
    '                                'Next ' unique habitat class

    '                                pFeature = pFeatureCursor.NextFeature

    '                            Loop     ' selected feature

    '                            'k = 0
    '                            'Dim dTemp As Double
    '                            'For k = 0 To lStatsMatrix.Count - 1
    '                            '    Dim sTemp2 As String = lStatsMatrix(k).UniqueClass
    '                            '    dTemp = lStatsMatrix(k).Quantity
    '                            'Next

    '                        End If ' There is a selection set

    '                        '' Print Quantity Field
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                        'pResultsForm.txtRichResults.AppendText("        Quantity Field (" + sUnit + "): ")
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
    '                        'pResultsForm.txtRichResults.AppendText(lLayersFields(j).QuanField + Environment.NewLine)

    '                        '' Print Class Field
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                        'pResultsForm.txtRichResults.AppendText("        Class Field: ")
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
    '                        'pResultsForm.txtRichResults.AppendText(lLayersFields(j).ClsField + Environment.NewLine)

    '                        '' add 'columns' for each
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                        'pResultsForm.txtRichResults.AppendText("        <class> / <habitat>" + Environment.NewLine)

    '                        'If bDockWin = True Then
    '                        '    If containedBox IsNot Nothing Then
    '                        '        containedBox.Items.Add("   Quantity Field: " + lLayersFields(j).QuanField)
    '                        '        containedBox.Items.Add("      Class Field: " + lLayersFields(j).ClsField)
    '                        '    End If
    '                        'End If ' bDockWin True

    '                        ' If there are items in the stats matrix
    '                        If lStatsMatrix.Count <> 0 Then
    '                            k = 1
    '                            ' For each unique value in the matrix
    '                            ' (always skip first row of matrix as it is the 'column headings')
    '                            For k = 1 To lStatsMatrix.Count - 1
    '                                'If bDBF = True Then

    '                                ''MsgBox("Debug:C16")
    '                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    '                                With pHabStatsObject_2
    '                                    .Layer = pFeatureLayer.Name
    '                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
    '                                    .bID = ID
    '                                    .bEID = iEID
    '                                    .bType = sType
    '                                    .Sink = f_sOutID
    '                                    .SinkEID = f_siOutEID
    '                                    .Direction = sDirection2
    '                                    .TotalImmedPath = sHabTypeKeyword
    '                                    .UniqueClass = CStr(lStatsMatrix(k).UniqueClass)
    '                                    .ClassName = CStr(lLayersFields(j).ClsField)
    '                                    .Quantity = lStatsMatrix(k).Quantity
    '                                    .Unit = sUnit
    '                                End With
    '                                lHabStatsList.Add(pHabStatsObject_2)

    '                                ''MsgBox("Debug:C17")
    '                                'End If

    '                                If sUnit = "n/a" Then
    '                                    'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                                    'pResultsForm.txtRichResults.AppendText("         " & lStatsMatrix(k).UniqueClass & "    " & Format(lStatsMatrix(k).Quantity, "0.00") & Environment.NewLine)
    '                                Else
    '                                    'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                                    'pResultsForm.txtRichResults.AppendText("         " & lStatsMatrix(k).UniqueClass & "    " & Format(lStatsMatrix(k).Quantity, "0.00") & sUnit & Environment.NewLine)
    '                                End If
    '                            Next

    '                            ' Insert a line break
    '                            'pResultsForm.txtRichResults.AppendText(Environment.NewLine)

    '                        Else ' If there are no statistics
    '                            'If bDBF = True Then

    '                            ''MsgBox("Debug:C18")
    '                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    '                            With pHabStatsObject_2
    '                                .Layer = pFeatureLayer.Name
    '                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
    '                                .bID = ID
    '                                .bEID = iEID
    '                                .bType = sType
    '                                .Sink = f_sOutID
    '                                .SinkEID = f_siOutEID
    '                                .Direction = sDirection
    '                                .TotalImmedPath = sHabTypeKeyword
    '                                .UniqueClass = "none"
    '                                .ClassName = CStr(lLayersFields(j).ClsField)
    '                                .Quantity = 0.0
    '                                .Unit = sUnit
    '                            End With
    '                            lHabStatsList.Add(pHabStatsObject_2)

    '                            '                            End If

    '                            ''MsgBox("Debug:C19")
    '                            ' If no stats found then add zeros
    '                            If sUnit = "n/a" Then
    '                                'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                                'pResultsForm.txtRichResults.AppendText("         none    0.00" & Environment.NewLine)
    '                            Else
    '                                'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                                'pResultsForm.txtRichResults.AppendText("         none    0.00" & sUnit & Environment.NewLine)
    '                            End If


    '                            'If bDockWin = True Then
    '                            '    If containedBox IsNot Nothing Then
    '                            '        If sUnit = "n/a" Then
    '                            '            containedBox.Items.Add("            " & "0.00")
    '                            '        Else
    '                            '            containedBox.Items.Add("            " & "0.00" & sUnit)
    '                            '        End If
    '                            '    End If
    '                            'End If
    '                        End If ' There are items in the statsmatrix

    '                        '' Insert a line break
    '                        'pResultsForm.txtRichResults.AppendText(Environment.NewLine)

    '                    Else   ' if the habitat class case field is not found


    '                        '' Reset the stats objects
    '                        ''---------TEMP----------
    '                        'pDataStats = New DataStatistics
    '                        'pStatisticsObject = New StatisticsObject(Nothing, Nothing)
    '                        '' Clear the statsMatrix
    '                        'lStatsMatrix = New List(Of StatisticsObject)
    '                        '' -------- END TEMP -----------

    '                        ''MsgBox("Debug:C20")
    '                        dTotalArea = 0

    '                        If pFeatureSelection.SelectionSet.Count <> 0 Then

    '                            pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)

    '                            ''---------TEMP----------
    '                            'Dim pEnum As IEnumerator
    '                            'With pDataStats
    '                            '    .Cursor = pCursor
    '                            '    .Field = lLayersFields(j).ClsField
    '                            'End With
    '                            ''pEnum = pDataStats.UniqueValues

    '                            'With pStatisticsObject
    '                            '    .UniqueClass = "Classes"
    '                            '    .Quantity = Nothing '***
    '                            'End With
    '                            'lStatsMatrix.Add(pStatisticsObject)

    '                            'With pStatisticsObject
    '                            '    .UniqueClass = "1"
    '                            '    .Quantity = 40 '***
    '                            'End With
    '                            'lStatsMatrix.Add(pStatisticsObject)

    '                            'With pStatisticsObject
    '                            '    .UniqueClass = "2"
    '                            '    .Quantity = 100 '***
    '                            'End With
    '                            'lStatsMatrix.Add(pStatisticsObject)
    '                            'Dim n As Integer

    '                            'For n = 0 To lStatsMatrix.Count - 1

    '                            'Next
    '                            '' -------- END TEMP -----------

    '                            ''MsgBox("Debug:C21")
    '                            pFeatureCursor = CType(pCursor, IFeatureCursor)
    '                            pFeature = pFeatureCursor.NextFeature

    '                            ' Get the summary field and add the value to the
    '                            ' total for habitat area.
    '                            ' ** ==> Multiple fields could be added here in a 'for' loop.

    '                            ''MsgBox("Debug:C22")
    '                            iFieldVal = pFeatureCursor.FindField(lLayersFields(j).QuanField)

    '                            'Try
    '                            '    dHabArea = Convert.ToDouble(vHabTemp)
    '                            'Catch ex As Exception
    '                            '    MsgBox("The Habitat Quantity found in the " & lLayersFields(j).Layer & " was not convertible" _
    '                            '    & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
    '                            '    dHabArea = 0
    '                            'End Try

    '                            ' For each selected feature
    '                            m = 1
    '                            Do While Not pFeature Is Nothing
    '                                Try
    '                                    vTemp = pFeature.Value(iFieldVal)
    '                                Catch ex As Exception
    '                                    MsgBox("Could not retrieve the habitat quantity value found in the" + _
    '                                    lLayersFields(j).Layer.ToString + ". Was not convertable to type 'double'." + _
    '                                    ex.Message)
    '                                    vTemp = 0
    '                                End Try
    '                                Try
    '                                    dTempQuan = Convert.ToDouble(vTemp)
    '                                Catch ex As Exception
    '                                    'MsgBox("Could not convert the habitat quantity value found in the" + _
    '                                    'lLayersFields(j).Layer.ToString + ". Was not convertable to type 'double'." + _
    '                                    'ex.Message)
    '                                    dTempQuan = 0
    '                                End Try
    '                                ' Insert into the corresponding column of the second
    '                                ' row the updated habitat area measurement.
    '                                dTotalArea = dTotalArea + dTempQuan
    '                                pFeature = pFeatureCursor.NextFeature
    '                            Loop     ' selected feature

    '                            ''MsgBox("Debug:C24")
    '                        End If ' there are selected features

    '                        ' If DBF tables are to be output
    '                        'If bDBF = True Then

    '                        pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    '                        With pHabStatsObject_2
    '                            .Layer = pFeatureLayer.Name
    '                            .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
    '                            .bID = ID
    '                            .bEID = iEID
    '                            .bType = sType
    '                            .Sink = f_sOutID
    '                            .SinkEID = f_siOutEID
    '                            .Direction = sDirection2
    '                            .TotalImmedPath = sHabTypeKeyword
    '                            .UniqueClass = "none"
    '                            .ClassName = CStr(lLayersFields(j).ClsField)
    '                            .Quantity = dTotalArea
    '                            .Unit = sUnit
    '                        End With
    '                        lHabStatsList.Add(pHabStatsObject_2)

    '                        'End If ' DBF tables output

    '                        '' Print Quantity Field
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                        'pResultsForm.txtRichResults.AppendText("        Quantity Field (" + sUnit + "): ")
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
    '                        'pResultsForm.txtRichResults.AppendText(lLayersFields(j).QuanField + Environment.NewLine)

    '                        '' add 'columns' for each
    '                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                        'pResultsForm.txtRichResults.AppendText("        <habitat>" + Environment.NewLine)

    '                        If sUnit = "n/a" Then
    '                            'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                            'pResultsForm.txtRichResults.AppendText("        " & Format(dTotalArea, "0.00") & Environment.NewLine)
    '                        Else
    '                            'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Regular)
    '                            'pResultsForm.txtRichResults.AppendText("        " & Format(dTotalArea, "0.00") & sUnit & Environment.NewLine)
    '                        End If

    '                        'If bDockWin = True Then
    '                        '    If containedBox IsNot Nothing Then
    '                        '        containedBox.Items.Add("   Quantity Field: " + lLayersFields(j).QuanField)
    '                        '        If sUnit = "n/a" Then
    '                        '            containedBox.Items.Add("            " & Format(dTotalArea, "0.00"))
    '                        '        Else
    '                        '            containedBox.Items.Add("            " & Format(dTotalArea, "0.00") & sUnit)
    '                        '        End If
    '                        '    End If

    '                        'End If

    '                        '' Insert a line break
    '                        'pResultsForm.txtRichResults.AppendText(Environment.NewLine)

    '                    End If ' found habitat class field in layer

    '                    ' increment the loop counter for
    '                    iLoopCount = iLoopCount + 1

    '                End If     ' feature layer matches hab class layer
    '            Next           ' habitat layer
    '        End If ' featurelayer is valid
    '        pFeatureLayer = CType(pEnumLayer.Next, IFeatureLayer)
    '    Loop

    'End Sub
    Private Function DownStreamConnected(ByRef pUpConnected As IEnumNetEID, ByRef pConnected As IEnumNetEID) As IEnumNetEID

        ' =====================================================
        ' Function:     downstream connected
        ' Author:       Greig Oldford
        ' Date Created: October 11, 2010
        ' Description:  This function accepts two lists of edge or junction elements:
        '               (1) the upstream connected and (2) the connected.  It subtracts
        '               1 from 2 to give the downstream connected and returns this list. 
        ' Note:         This function uses the same logic as flagcheck, essentially. 
        ' =====================================================

        Dim i As Integer
        Dim m As Integer
        Dim pDownstreamConnectedGEN As IEnumNetEIDBuilderGEN  ' list holds flags on barriers
        pDownstreamConnectedGEN = New EnumNetEIDArray

        Dim pDownstreamConnected As IEnumNetEID
        Dim bCommonElement As Boolean
        Dim iEID As Integer

        pConnected.Reset()
        i = 0

        ' For each of the connected elements
        For i = 0 To pConnected.Count - 1

            bCommonElement = False  ' assume flag is not on barrier
            iEID = pConnected.Next  ' get the EID of flag
            m = 0
            pUpConnected.Reset()

            ' For each upstream element
            For m = 0 To pUpConnected.Count - 1
                'If endEID = pOriginalBarriersList(m) Then 'VB.NET
                If iEID = pUpConnected.Next Then
                    bCommonElement = True
                End If
            Next

            If bCommonElement = False Then  'put EID on downstream connected list
                ' THIS LIST COULD BE USED IN FUTURE TO FILTER BAD FLAGS OUT
                ' I.E. - check which flags are on barriers and remove only those ones automatically.
                pDownstreamConnectedGEN.Add(iEID)
            End If
        Next

        ' QI 
        pDownstreamConnected = CType(pDownstreamConnectedGEN, IEnumNetEID)
        DownStreamConnected = pDownstreamConnected

    End Function
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
        Dim plExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
        If m_FiPEx__1.m_bLoaded = True Then ' If there were any extension settings set

            iExclusions = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numExclusions"))
            Dim ExclusionsObj As New LayerToExclude(Nothing, Nothing, Nothing)

            ' match any of the line layers saved in stream to those in listboxes
            If iExclusions > 0 Then
                For j = 0 To iExclusions - 1
                    'sLineLayer = m_DFOExt.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    ExclusionsObj = New LayerToExclude(Nothing, Nothing, Nothing)
                    With ExclusionsObj
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("ExcldLayer" + j.ToString))
                        .Feature = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("ExcldFeature" + j.ToString))
                        .Value = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("ExcldValue" + j.ToString))
                    End With

                    ' add to the module level list
                    plExclusions.Add(ExclusionsObj)
                Next
            End If
        Else
            backgroundworker2.CancelAsync()
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
        Dim x As Integer
        Dim aOID As New List(Of Integer)
        Dim iCount As Integer
        Dim pMxDoc As IMxDocument
        Dim pMap As IMap
        Dim pAV As IActiveView
        Dim vVal As Object
        Dim sTempVal As String

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMap = pMxDoc.FocusMap
        pAV = pMxDoc.ActiveView
        pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)

        If (pFeatureSelection.SelectionSet.Count <> 0) Then
            If iExclusions > 0 Then
                x = 0
                For x = 0 To iExclusions - 1
                    Dim pCursor As ICursor
                    pFeatureSelection.SelectionSet.Search(Nothing, True, pCursor)
                    pFeatureCursor = CType(pCursor, IFeatureCursor)

                    ' FOR DEBUG - see how many feautres are selected in the layer
                    Dim iCountTEMP As Integer
                    iCountTEMP = pFeatureSelection.SelectionSet.Count

                    If pFeatureLayer.Name = plExclusions(x).Layer Then
                        ' Find the field
                        iFieldVal = pFeatureCursor.FindField(plExclusions(x).Feature)

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

                                ' If the value matches then add the OID to an array
                                ' to subtract from the selection set.
                                If sTempVal IsNot Nothing Then
                                    If sTempVal = plExclusions(x).Value Then
                                        aOID.Add(pFeature.OID)
                                    End If
                                End If

                                pFeature = pFeatureCursor.NextFeature
                            Loop

                            iCount = aOID.Count
                            Dim iOID As Integer
                            If iCount <> 0 Then
                                For i As Integer = 0 To iCount - 1
                                    ' Remove excluded features from selection set by OID
                                    ' NOTE: remove list does not actually remove a list
                                    ' so we need a loop here; see http://forums.esri.com/Thread.asp?c=159&f=1707&t=224394
                                    iOID = aOID.Item(i)
                                    pFeatureSelection.SelectionSet.RemoveList(1, iOID)
                                    pFeatureSelection.SelectionChanged()
                                Next
                                pAV.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
                                pAV.Refresh()

                            End If
                            ' If the feature value for the column name matches
                            ' Note that pFeature.Value() outputs type Variant

                            ' Set variable indicating to exclude feature

                            'pFeatureSelection.SelectionSet.RemoveList(
                            'MsgBox ("Set the exclusion var to: " + CStr(bExcludeFeat))
                        End If
                    End If
                Next
            End If
        End If

    End Sub
    Private Class FindStatsClassPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the layer matches 
        ' and the sink/barr EID matches as well.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _Class As String

        Public Sub New(ByVal class2 As String)
            Me._Class = class2
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareStatsClass(ByVal obj As HabStatisticsObject) As Boolean
            Return (_Class = obj.UniqueHabClass)
        End Function
    End Class
    Private Class RefineHabStatsListPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the 
        ' and the sink EID, barrier ID, and layer matches.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _SinkEID As Integer
        Private _BarrEID As Integer
        Private _LayerID As Integer
        Private _Direction As String
        Private _Type As String

        Public Sub New(ByVal sinkEID As Integer, ByVal barrEID As Integer, ByVal layerID As String, ByVal direction As String, ByVal type As String)
            Me._SinkEID = sinkEID
            Me._BarrEID = barrEID
            Me._LayerID = layerID
            Me._Direction = direction
            Me._Type = type
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareHabStuff(ByVal obj As StatisticsObject_2) As Boolean
            Return (_SinkEID = obj.SinkEID And _BarrEID = obj.bEID And _LayerID = obj.LayerID _
            And _Direction = obj.Direction And _Type = obj.TotalImmedPath)
        End Function
    End Class
    Private Sub UpdateSummaryTable(ByRef lRefinedHabStatsList As List(Of StatisticsObject_2), ByRef iHabRowCount As Integer, _
     ByRef pResultsForm3 As FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3, ByRef iMaxRowIndex As Integer, ByRef iBarrIndex As Integer, _
     ByRef bColorSwitcher As Boolean, ByRef iSinkRowCount As Integer, ByRef iBarrRowCount As Integer, ByVal bSinkVisit As Boolean)

        Dim m As Integer = 0
        Dim iThisRowIndex As Integer
        Dim pDataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle
        Dim bTrigger As Boolean = False
        Dim sLayer, sDirection, sTraceType, sClass, sUnit As String
        Dim dTotalHab, dQuantity As Double

        ' For each habitat stat in the refined list
        For m = 0 To lRefinedHabStatsList.Count - 1

            ' if it's the first habitat item in the list
            ' then it's going on the barrierindex line
            ' otherwise it's habrowindex + habrowcount
            ' if that number exceeds the Max row index
            ' then a new row will need to be inserted
            If (iBarrIndex + iHabRowCount) > iMaxRowIndex Then
                iMaxRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                iThisRowIndex = iMaxRowIndex
                iSinkRowCount += 1
                If bSinkVisit = False Then
                    iBarrRowCount += 1
                End If
                bTrigger = True
            Else
                iThisRowIndex = iBarrIndex + iHabRowCount
            End If

            If bColorSwitcher = True Then
                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style
                pDataGridViewCellStyle.BackColor = Color.Lavender
            Else
                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style
                pDataGridViewCellStyle.BackColor = Color.PowderBlue
            End If
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Style = pDataGridViewCellStyle

            ' if the row count of the habitat metrics exceeds the 
            ' statistics metrics then the colors of cells below the metrics
            ' will also have to be changed
            If bTrigger = True Then
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(1).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(2).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(4).Style = pDataGridViewCellStyle
            End If
            'End If

            If m = 0 Then
                sLayer = lRefinedHabStatsList(m).Layer
                sDirection = lRefinedHabStatsList(m).Direction
                sTraceType = lRefinedHabStatsList(m).TotalImmedPath
                sUnit = lRefinedHabStatsList(m).Unit

                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Value = sLayer
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Value = sDirection
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Value = sTraceType
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Value = sUnit
            End If

            sClass = lRefinedHabStatsList(m).UniqueClass
            dQuantity = Math.Round(lRefinedHabStatsList(m).Quantity, 2)
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Value = sClass
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Value = dQuantity

            dTotalHab = dTotalHab + lRefinedHabStatsList(m).Quantity
            iHabRowCount += 1

            ' Add the total field if necessary
            If m = lRefinedHabStatsList.Count - 1 Then

                If (iBarrIndex + iHabRowCount) > iMaxRowIndex Then
                    iMaxRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    iThisRowIndex = iMaxRowIndex
                    iSinkRowCount += 1
                    If bSinkVisit = False Then
                        iBarrRowCount += 1
                    End If
                    bTrigger = True
                Else
                    iThisRowIndex = iBarrIndex + iHabRowCount
                End If

                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Style
                pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, pResultsForm3.DataGridView1.Font.Size, FontStyle.Bold)
                pDataGridViewCellStyle.BackColor = Color.SlateGray
                'pDataGridViewCellStyle.ForeColor = Color.White
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style = pDataGridViewCellStyle

                ' If the habitat row exceeds maximum metric row then 
                ' color the cells below the metric row appropriately
                If bTrigger = True And bColorSwitcher = True Then
                    pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style
                    pDataGridViewCellStyle.BackColor = Color.Lavender
                ElseIf bTrigger = True And bColorSwitcher = False Then
                    pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style
                    pDataGridViewCellStyle.BackColor = Color.PowderBlue
                End If
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(1).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(2).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(4).Style = pDataGridViewCellStyle
                ' for border adjustment
                'dataGridViewAdvancedBorderStyleInput = pResultsForm3.DataGridView1.Rows(iHabRowIndex).DefaultCellStyle(8).
                'dataGridViewAdvancedBorderStylePlaceHolder = dataGridViewAdvancedBorderStyleInput
                'dataGridViewAdvancedBorderStylePlaceHolder.Top = System.Windows.Forms.DataGridViewAdvancedCellBorderStyle.InsetDouble

                'pResultsForm3.DataGridView1.Rows(iHabRowIndex).Cells(8).AdjustCellBorderStyle(dataGridViewAdvancedBorderStyleInput, dataGridViewAdvancedBorderStylePlaceHolder, False, False, False, False)
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Value = "Total"
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Value = Math.Round(dTotalHab, 2)

                iHabRowCount += 1

            End If
        Next

        '' Switch the color switcher
        'If bColorSwitcher = True Then
        '    bColorSwitcher = False
        'Else
        '    bColorSwitcher = True
        'End If


    End Sub
    Private Class FindBarrierMetricsBySinkEIDPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the layer matches 
        ' and the sink/barr EID matches as well.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _SinkEID As Integer
        Private _BarrierEID As Integer

        Public Sub New(ByVal sinkEID As Integer, ByVal barrierEID As Integer)
            Me._SinkEID = sinkEID
            Me._BarrierEID = barrierEID
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareEID(ByVal obj As MetricsObject) As Boolean
            Return (_SinkEID = obj.SinkEID And _BarrierEID = obj.BarrEID)
        End Function
    End Class
    Private Class FindLayerAndBarrEIDPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the layer matches 
        ' and the sink/barr EID matches as well.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _BarrEID As Integer
        Private _LayerID As Integer

        Public Sub New(ByVal barrEID As Integer, ByVal layerID As String)
            Me._BarrEID = barrEID
            Me._LayerID = layerID
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareEIDandLayer(ByVal obj As StatisticsObject_2) As Boolean
            Return (_BarrEID = obj.bEID And _LayerID = obj.LayerID)
        End Function
    End Class
    Private Class FindLayerAndBarrEIDAndSinkEIDPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the 
        ' and the sink EID, barrier ID, and layer matches.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _SinkEID As Integer
        Private _BarrEID As Integer
        Private _LayerID As Integer

        Public Sub New(ByVal sinkEID As Integer, ByVal barrEID As Integer, ByVal layerID As String)
            Me._SinkEID = sinkEID
            Me._BarrEID = barrEID
            Me._LayerID = layerID
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareEIDandLayer(ByVal obj As StatisticsObject_2) As Boolean
            Return (_SinkEID = obj.SinkEID And _BarrEID = obj.bEID And _LayerID = obj.LayerID)
        End Function
    End Class
    Private Class FindBarriersBySinkEIDPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the layer matches 
        ' and the sink/barr EID matches as well.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _SinkEID As Integer

        Public Sub New(ByVal sinkEID As Integer)
            Me._SinkEID = sinkEID
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareEID(ByVal obj As BarrAndBarrEIDAndSinkEIDs) As Boolean
            Return (_SinkEID = obj.SinkEID)
        End Function
    End Class
End Class
