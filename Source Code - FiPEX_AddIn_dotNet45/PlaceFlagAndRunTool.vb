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
                m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
            End If
            If m_UNAExt Is Nothing Then
                m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
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

        ProgressForm = New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmAnalysisProgress()
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
            m_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetExtension()
        End If
        If m_UNAExt Is Nothing Then
            m_UNAExt = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.FishPassageExtension.GetUNAExt
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
            MsgBox("No network junction found here. Please zoom in and click within the 100 map unit tolerance of a network junction.")
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
                ' =====================================================
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
                ' ==================================================
                ' ============== INTERSECT FEATURES ================

                Call IntersectFeatures()

                If m_bCancel = True Then
                    backgroundworker2.CancelAsync()
                    backgroundworker2.ReportProgress(100, "Closing")
                    Exit Sub
                End If
                backgroundworker2.ReportProgress(30, "Performing Network Traces" & ControlChars.NewLine & _
                                                 "Trace: upstream immediate." & ControlChars.NewLine & _
                                                 "Excluding Features")

                ' ==================================================
                ' ============== HABITAT STATS =====================
                ' use results for Upstream habitat if needed
                sHabType = "Immediate"
                sHabDir = "upstream"
                Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields,
                                                         sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                ' ============== END HABITAT STATS ==================
                ' ==================================================


                ' =================================================
                ' ============== EXCLUDE FEATURES ================
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

                ' =====================================================
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

                ' =====================================================
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

                ' ==================================================
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


                ' ==================================================
                ' ============== HABITAT STATS =====================
                sHabType = "Immediate"
                sHabDir = "downstream"

                Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields,
                                                         sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                ' ============ END HABITAT STATS ==================
                ' =================================================


                ' ==================================================
                ' ============== EXCLUDE FEATURES ================

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

            End If
        End If  ' bUpHab = True Or bDCI = True Or bDownHab = True 


        ' If Downstream Path Habitat desired
        If bPathDownHab = True Then

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

            ' =======================================================
            ' ============== SAVE FOR HIGHLIGHTING ==================
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

            ' ==================================================
            ' ================ INTERSECT FEATURES ==============
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(40, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: path downstream." & ControlChars.NewLine & _
                                                 "Intersecting Features")

            Call IntersectFeatures()

            ' ==================================================
            ' ================= HABITAT STATS ==================
            sHabType = "Path"
            sHabDir = "downstream"
            Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields,
                                                     sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
            ' =============== END HABITAT STATS ================
            ' ==================================================

            ' ==================================================
            ' ================= EXCLUDE FEATURES ===============
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

                    ' =================================================
                    ' ================ INTERSECT FEATURES =============
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(45, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total upstream." & ControlChars.NewLine & _
                                                     "Intersecting Features.")
                    Call IntersectFeatures()

                    ' ################################################
                    ' ############## HABITAT STATS ###################
                    sHabType = "Total"
                    sHabDir = "upstream"
                    Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields,
                                                             sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                    ' ############## END HABITAT STATS ###############
                    ' ################################################

                    ' ================================================
                    ' ================ EXCLUDE FEATURES ==============
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
                    ' ======================================================
                    ' =============== END EXCLUDE FEATURES =================

                    pActiveView.Refresh() ' refresh the view

                End If

                If bTotalDownHab = True Then

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

                    ' ====================================================
                    ' =================== INTERSECT FEATURES =============
                    If m_bCancel = True Then
                        backgroundworker2.CancelAsync()
                        backgroundworker2.ReportProgress(100, "Closing")
                        Exit Sub
                    End If
                    backgroundworker2.ReportProgress(50, "Performing Network Traces" & ControlChars.NewLine & _
                                                     "Trace: total downstream." & ControlChars.NewLine & _
                                                     "Intersecting features.")

                    Call IntersectFeatures()

                    ' ===================================================
                    ' ================= HABITAT STATS ===================
                    sHabType = "Total"
                    sHabDir = "downstream"
                    Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields,
                                                             sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
                    ' =========== === END HABITAT STATS ===================
                    ' =====================================================

                    ' ===================================================
                    ' =================== EXCLUDE FEATURES =============
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

                End If
            End If ' bTotalUpHab = True Or bTotalDownHab = True
        End If ' bTotalUpHab = True Or bTotalDownHab = True Or bTotalPathDownHab = True

        If bTotalPathDownHab = True Then

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
            ' =====================================================
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

            ' ======================================================
            ' =================== INTERSECT FEATURES ===============
            If m_bCancel = True Then
                backgroundworker2.CancelAsync()
                backgroundworker2.ReportProgress(100, "Closing")
                Exit Sub
            End If
            backgroundworker2.ReportProgress(55, "Performing Network Traces" & ControlChars.NewLine & _
                                             "Trace: total path downstream." & ControlChars.NewLine & _
                                             "Intersecting features.")
            Call IntersectFeatures()

            ' ===============================================
            ' =================== HAB STATS =================
            sHabType = "Total Path"
            sHabDir = "downstream"
            Call SharedSubs.calculateStatistics_2020(lHabStatsList, pLLayersFields, pPLayersFields, sOutID, iEID, sType, sOutID, iEID, sHabType, sHabDir)
            ' =================== HAB STATS =================
            ' ===============================================

            ' ======================================================
            ' =================== EXCLUDE FEATURES =================
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
            ' ================ END EXCLUDE FEATURES =================
            ' ======================================================

            pActiveView.Refresh() ' refresh the view

        End If

        ' Create a result highlight of all areas traced
        Dim pTotalResultsEdges, pTotalResultsJunctions As IEnumNetEID
        pTotalResultsEdges = CType(pTotalResultsEdgesGEN, IEnumNetEID)
        pTotalResultsJunctions = CType(pTotalResultsJunctionsGEN, IEnumNetEID)
        pNetworkAnalysisExtResults.ResultsAsSelection = False
        pNetworkAnalysisExtResults.SetResults(pTotalResultsJunctions, pTotalResultsEdges)
        pNetworkAnalysisExtResults.ResultsAsSelection = True

        If sFlagCheck = "barriers" Or bTotalUpHab = True Or bTotalDownHab = True Or bTotalPathDownHab = True Then

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
            ' ==================== END RESET BARRIERS =================
            ' =========================================================
        End If

        ' ===========================================================
        ' ================== RESET FLAGS ============================

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
        ' ============================ END RESET FLAGS ===========================
        ' ========================================================================

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

        ' ===============================================================
        ' ================== BEGIN WRITE TO OUTPUT FORM =================

        If m_bCancel = True Then
            backgroundworker2.CancelAsync()
            backgroundworker2.ReportProgress(100, "Closing")
            Exit Sub
        End If
        m_iProgress = 60
        backgroundworker2.ReportProgress(m_iProgress, "Preparing Output Form.")

        Dim pSinkAndDCIS As New SinkandDCIs(Nothing, Nothing, Nothing, Nothing)
        Dim lSinkAndDCIS As New List(Of SinkandDCIs)
        Dim pSinkIDAndTypes As New SinkandTypes(Nothing, Nothing, Nothing)
        Dim lSinkIDandTypes As New List(Of SinkandTypes)

        i = 0
        Dim bSinkThere, bDCIpMatch, bDCIdMatch, bEntered As Boolean
        Dim row As DataRow

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
            End If
        Next

        ' refresh the view
        pActiveView.Refresh()

        Dim pResultsForm3 As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2022.frmResults_3
        pResultsForm3.Show()
        Dim numbarrsnodes As String = CStr(pOriginaljuncFlagsList.Count)
        SharedSubs.ResultsForm2020(pResultsForm3, lSinkIDandTypes, lHabStatsList, lMetricsObject,
                                       BeginTime, numbarrsnodes, iOrderNum, sDirection)
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
                                            Try
                                                sLabelValue = Convert.ToString(pFeature.Value(pFields.FindField(sLabelField)))
                                            Catch ex As Exception
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

        'Dim pGp As IGeoProcessor
        'pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor

        Dim pGp2 As IGeoProcessor2
        pGp2 = New ESRI.ArcGIS.Geoprocessing.GeoProcessor

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

                                                        ' 2020 - turn off history otherwise results and MXD gets huge
                                                        pGp2.LogHistory = False
                                                        pGp2.AddToResults = False
                                                        pGPResults = pGp2.Execute("SelectLayerByLocation_management", pParameterArray, Nothing)
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
