Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Text.RegularExpressions
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesOleDB ' For DCI calculation
Imports ESRI.ArcGIS.GeoDatabaseUI    ' For use with DCI table conversion - IExportInterface
Imports System.IO                    ' For reading DCI output file
Imports System                       ' added based on MSDN instructions for writing test txt file (for permission check)
Imports System.ComponentModel

' MUST REMOVE THIS IMPORTS PROJECT REFERENCE TO GUROBI BEFORE
' REDISTRIBUTION.  GUROBI library is found in c:\gurobi500\win64\gurobi50.dll
' for re-add
'Imports Gurobi

Public Class Analysis
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    Private m_UNAExt As IUtilityNetworkAnalysisExt
    Private m_pNetworkAnalysisExt As INetworkAnalysisExt
    Private m_FiPEx__1 As FishPassageExtension
    Private backgroundworker1 As New BackgroundWorker
    Private ProgressForm As New frmAnalysisProgress
    Private m_bCancel As Boolean = False

    ' common decision object for thesis sensitivity analysis
    Private m_lSACommonDecisionsObject_DIR As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSACommonDecisionsObject_UNDIR As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    ' decisions for the initial 'best guess' treatment - thesis SA
    Private m_lSABestGuessDecisionsObject_DIR As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSABestGuessDecisionsObject_UNDIR As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSA_P_Results_DIR As List(Of DIR_OptResultsObject) = New List(Of DIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSA_A_Results_DIR As List(Of DIR_OptResultsObject) = New List(Of DIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSA_P_Results_UNDIR As List(Of UNDIR_OptResultsObject) = New List(Of UNDIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Private m_lSA_A_Results_UNDIR As List(Of UNDIR_OptResultsObject) = New List(Of UNDIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS
    Dim bCreateRandomPerm As Boolean



    Public Sub New()

        backgroundworker1.WorkerReportsProgress = True
        backgroundworker1.WorkerSupportsCancellation = True
        AddHandler backgroundworker1.DoWork, AddressOf backgroundWorker1_DoWork
        AddHandler backgroundworker1.ProgressChanged, AddressOf backgroundworker1_ProgressChanged
        AddHandler backgroundworker1.RunWorkerCompleted, AddressOf backgroundworker1_RunWorkerCompleted
    End Sub

    Private Property m_lSA_P_Results As Boolean



    Protected Overloads Overrides Sub OnUpdate()
        ' use the extension listener to avoid constant checks to the 
        ' map network.  The extension listener will only update the boolean
        ' check on network count if there's a map change
        ' upgrade at version 10
        ' protected override void OnUpdate()
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

            'Dim pNetworkAnalysisExt As INetworkAnalysisExt
            'Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
            'Dim barrNumber As Integer
            Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
            Dim iFlagNumber As Integer ' for junction flags
            'Dim eFlagNumber As Integer ' for edge flags

            'pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
            'pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)
            pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
            iFlagNumber = pNetworkAnalysisExtFlags.JunctionFlagCount

            If iFlagNumber > 0 Then
                Me.Enabled = True
            Else
                Me.Enabled = False
            End If
        Else
            Me.Enabled = False
        End If
    End Sub
    Protected Overrides Sub OnClick()

        ' Open Options in Modal
        Dim bKeepGoing As Boolean = False
        My.ArcMap.Application.CurrentTool = Nothing
        Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension
        Using MyForm As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmRunAdvancedAnalysis

            If MyForm.Form_Initialize(My.ArcMap.Application) Then
                MyForm.TabPage4.BringToFront()
                MyForm.ShowDialog()
                bKeepGoing = MyForm.m_bRun
            End If
        End Using

        ' Open Progress Form
        If bKeepGoing = True Then
            Call RunAnalysis()
        End If

    End Sub
    Private Sub backgroundworker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        m_bCancel = False

        If backgroundworker1.CancellationPending = True Then
            e.Cancel = True
        End If

        Call ShowProgressForm()


        '    Else
        '        ' Perform a time consuming operation and report progress.
        '        System.Threading.Thread.Sleep(500)
        '        backgroundworker1.ReportProgress(i * 10)
        '    End If
        'Next
    End Sub
    Private Sub backgroundworker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        If e.Cancelled = True Then
            m_bCancel = True
        ElseIf e.Error IsNot Nothing Then
            If Not ProgressForm Is Nothing Then
                Try
                    ProgressForm.lblProgress.Text = "Error: " & e.Error.Message
                Catch ex As Exception
                    MsgBox("Error encountered in RunWorkerCompleted of Analysis command. " + ex.Message)
                End Try
            End If
        Else
            If Not ProgressForm Is Nothing Then
                Try
                    ProgressForm.lblProgress.Text = "Done!"
                    ProgressForm.cmdCancel.Text = "Close"
                Catch ex As Exception
                    MsgBox("Error encountered in RunWorkerCompleted of Analysis command. " + ex.Message)
                End Try
            End If
        End If

    End Sub
    Private Sub backgroundworker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        If Not ProgressForm Is Nothing Then
            Try
                If ProgressForm.Visible = True Then
                    If ProgressForm.m_bCloseMe = False Then
                        ProgressForm.ChangeProgressBar(e.ProgressPercentage)
                        ProgressForm.ChangeLabel(e.UserState.ToString)
                    End If
                End If
            Catch ex As Exception
                'exception raised... donworryaboutit
                MsgBox("You can ignore this error. Exception raised in 'progress changed' subroutine of analysis command" + ex.Message)
            End Try
        End If
    End Sub
    Private Sub ShowProgressForm()
        ' User clicks 'Run' then continue to main analysis sub
        ProgressForm = New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmAnalysisProgress()
        If ProgressForm.Form_Initialize() Then
            ProgressForm.ShowDialog()
            If Not ProgressForm Is Nothing Then
                If ProgressForm.Visible = False Then
                    m_bCancel = True
                End If
            End If
        End If

        If m_bCancel = True Then
            ProgressForm = Nothing
            backgroundworker1.CancelAsync()
        End If

    End Sub
    Private Sub RunAnalysis()
        ' 2020 changed from 'Protected' sub due to debugging problem

        ' ==================================================================
        ' Command:       RunAnalysis
        ' Modified:      Sep 2020, G Oldford
        '
        ' Description:   read settings from the extension stream,
        '                perform a network trace, 
        '                intersect Trace results with desired polygon and non-network lines
        '                exclude selected features that are not wanted
        '                calculate habitat area sorted by habitat class,
        '                output results in desired format.
        '                An option to repeat this process iteratively up or
        '                downstream of a startpoint is possible.  It then returns
        '                the network to the way it was originally.
        '
        ' Old Notes:    
        '                This tool has some unused code to prepare for using edges
        '                as barriers or startpoints.  
        ' Notes 2020:    to do: should be cleaned with functions created to make it more compact
        '                Note the algorithm is 'breadth first search' (as opposed to 'depth first search')

        m_bCancel = False
        ProgressForm = Nothing

        ' create a new background worker
        If Not backgroundworker1.IsBusy = True Then
            backgroundworker1.RunWorkerAsync()
        End If

        Threading.Thread.Sleep(1000)

        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(0, "Beginning Analysis")
        Dim iProgress As Integer ' for use in loops reporting progress


        'Change the mouse cursor to hourglass
        Dim pMouseCursor As IMouseCursor
        pMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim BeginTime As DateTime = DateTime.Now
        Dim EndTime As DateTime

        Dim pNetworkAnalysisExtFlags As INetworkAnalysisExtFlags
        Dim pTraceFlowSolver As ITraceFlowSolver
        Dim pResultEdges As IEnumNetEID
        Dim pResultJunctions As IEnumNetEID

        ' variables to use to display all area traced after analysis
        Dim pTotalResultsJunctions As IEnumNetEID
        Dim pTotalResultsEdges As IEnumNetEID
        Dim pTotalResultsJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pTotalResultsEdgesGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray

        Dim pFlowEndJunctions As IEnumNetEID
        Dim pFlowEndJunctionsTemp As IEnumNetEID 'temporary enum array
        Dim pFlowEndEdges As IEnumNetEID

        Dim pFirstJunctionBarriers As IEnumNetEID
        Dim pFirstEdgeBarriers As IEnumNetEID

        Dim sTableName As String ' TableName for Stats output

        ' These are for flow end elements found when
        ' the analysis is being done per flag
        Dim pFlowEndJunctionsPer As IEnumNetEID
        Dim pFlowEndEdgesPer As IEnumNetEID

        Dim pTraceTasks As ITraceTasks
        Dim eFlowElements As esriFlowElements
        Dim pNetworkAnalysisExt As INetworkAnalysisExt
        Dim pEnumNetEIDBuilder As IEnumNetEIDBuilder
        Dim pNetworkAnalysisExtResults As INetworkAnalysisExtResults
        Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers

        Dim orderLoop As Integer
        Dim pFlagDisplay As IFlagDisplay
        Dim pFLyrSlct As IFeatureLayer  ' to set all layers as selectable

        ' Variant variable for looping through line layers,
        ' features, and values for exclusion
        Dim pNetwork As INetwork
        Dim pNetElements As INetElements

        ' these values are returned by INetElements.QueryIDs
        ' and need to be of type 'long'
        Dim iFCID, iFID, iSubID, iEID As Integer

        Dim dBarrierPerm As Double
        Dim sNaturalYN As String

        ' For table output
        'Dim pFeatureDataset As IFeatureDataset
        'Dim pWorkspace As IWorkspace
        'Dim sFilepath As String

        'Dim pWorkspace2 As IWorkspace2
        Dim pFWorkspace As IFeatureWorkspace
        Dim pCursor As ICursor
        Dim pRowBuffer As IRowBuffer
        'Dim pRow As IRow

        ' Get the list of barriers that were just used in the last trace
        Dim barrierCount As Integer
        Dim bFlagDisplay As IFlagDisplay
        Dim bFID, bFCID, bSubID, bEID, rEID As Integer

        Dim pNextTraceBarrierEIDs As IEnumNetEID
        Dim pNextTraceBarrierEIDGEN As IEnumNetEIDBuilderGEN

        Dim i, j, m, n, k, p As Integer ' counter variables

        Dim flagOverBarrier As Boolean
        Dim pTable As ITable

        ' Variables to store original user-entered
        ' barrier EIDs
        'Dim mFlagDisplay As IFlagDisplay
        'Dim mFID As Integer
        'Dim mFCID As Integer
        'Dim mSubID As Integer

        'Dim mEID As New List(Of Integer) 'VB.NET
        'Dim mEID() As Integer
        Dim sTemp As String

        ' A new EnumNetEIDArray object to hold EIDs of barriers at the analysis
        Dim pEnumNetEIDArray As New EnumNetEIDArray
        'Dim pEnumNetEIDBuilderGEN As IEnumNetEIDBuilderGEN

        ' Variables to hold downstream barrier
        Dim pDwnstrmFlowEndJuncs As IEnumNetEID
        Dim pDwnstrmFlowEndEdges As IEnumNetEID
        'Dim maxFlowEndsFound As Integer

        ' A variable to hold the original user-set barriers' EIDs
        ' and a variable to hold original user-set flags' EIDs
        ' VB.NET:
        'Dim originalBarriersList As New ArrayList
        'Dim pOriginalEdgeFlagsList As New ArrayList
        'Dim pOriginaljuncFlagsList As New ArrayList
        Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginalEdgeFlagsListGEN As IEnumNetEIDBuilderGEN
        Dim pOriginaljuncFlagsListGEN As IEnumNetEIDBuilderGEN

        Dim pOriginalBarriersList As IEnumNetEID
        Dim pOriginalEdgeFlagsList As IEnumNetEID
        Dim pOriginaljuncFlagsList As IEnumNetEID
        'Dim pTempEIDList As IEnumNetEIDBuilderGEN

        Dim keepEID As Boolean
        Dim endEID As Integer
        Dim pNoSourceFlowEnds As IEnumNetEIDBuilderGEN
        pNoSourceFlowEnds = New EnumNetEIDArray
        Dim pNoSourceFlowEndsTemp As IEnumNetEIDBuilderGEN

        Dim pMxDocument As IMxDocument = CType(My.ArcMap.Application.Document, IMxDocument)
        Dim pMap As IMap = pMxDocument.FocusMap
        Dim pActiveView As IActiveView = CType(pMap, IActiveView)

        ' to hold stats on the habitat - will be written to table at end of sub
        'Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        'Dim pDCIStatsObject As New DCIStatisticsObject(Nothing, Nothing, Nothing, Nothing)
        Dim lHabStatsList As List(Of StatisticsObject_2) = New List(Of StatisticsObject_2)
        Dim lDCIStatsList As List(Of DCIStatisticsObject) = New List(Of DCIStatisticsObject)

        '2020 new object below to eventually replace lDCIStatsList
        Dim lAdv_DCI_Data_Object As List(Of Adv_DCI_Data_Object) = New List(Of Adv_DCI_Data_Object)
        Dim pAdvDCIDataObj = New Adv_DCI_Data_Object(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, _
                                                                 Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim lGLPKStatsList As List(Of GLPKStatisticsObject) = New List(Of GLPKStatisticsObject) ' for habitat stats for GLPK
        Dim lGLPKOptionsList As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject) ' for options stats for GLPK
        Dim lGLPKOptionsListTEMP As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject) ' for TEMP options stats for GLPK

        Dim sHabType As String

        Dim f_sOutID, f_sType As String ' To hold the (possibly) user-set ID and the 'type' of each flag
        Dim f_siOutEID As Integer
        Dim lBarrierAndSinkEIDs As List(Of BarrAndBarrEIDAndSinkEIDs) = New List(Of BarrAndBarrEIDAndSinkEIDs)
        Dim pBarrierAndSinkEIDs As New BarrAndBarrEIDAndSinkEIDs(Nothing, Nothing, Nothing, Nothing)

        ' for pairing with DCI stats at the end of each flag's loop 

        ' ===================== GET EXTENSION SETTINGS ====================
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(5, "Getting FIPEX Option Settings")

        ' Declare settings variables and set defaults
        Dim sDirection As String = "up"     ' Analysis direction default to 'upstream'
        Dim iOrderNum As Integer = 1        ' the ordernum retrieved default to 1
        Dim bMaximum As Boolean = False     ' set maximum order yes/no default to no
        Dim bConnectTab As Boolean = False  ' Connectivity table default to none
        Dim bBarrierPerm As Boolean = False ' Barrier perm field? 
        Dim bNaturalYN As Boolean = False   ' Natural Barrier y/n field?
        Dim bDCI As Boolean = False         ' Calculate DCI?
        Dim bDCISectional As Boolean = False ' Calculate DCI Sectional?

        '2020
        Dim bUseHabLength, bUseHabArea As Boolean
        Dim bDistanceLim As Boolean = False
        Dim bDistanceDecay As Boolean = False
        Dim dMaxDist As Double = 0.0
        Dim sDDFunction As String = "none"
        Dim bAdvConnectTab As Boolean = False

        Dim bDBF As Boolean = False         ' Include DBF output default none
        Dim sGDB As String = ""             ' Output GDB for DBF output
        Dim sPrefix As String = ""          ' Prefix for output tables

        Dim bUpHab As Boolean = False
        Dim bTotalUpHab As Boolean = False
        Dim bDownHab As Boolean = False
        Dim bTotalDownHab As Boolean = False
        Dim bPathDownHab As Boolean = False
        Dim bTotalPathDownHab As Boolean = False


        Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        Dim iLinesCount As Integer = 0      ' number of lines layers currently using

        Dim pLLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
        Dim pPLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)

        '2020 ' replace hablayerobj above with two 
        Dim LineLayerObj As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property
        Dim PolyLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property


        Dim lBarrierIDs As List(Of BarrierIDObj) = New List(Of BarrierIDObj)
        Dim pBarrierIDObj As New BarrierIDObj(Nothing, Nothing, Nothing, Nothing, Nothing)

        Dim iBarrierIDs As Integer = 0
        Dim sBarrierIDLayer, sBarrierIDField, sBarrierNaturalYNField, sBarrierPermField As String
        Dim sBarrierType As String = "Barrier" ' field name found in layers (GLPK / 2020 adv connect)


        Dim bGLPKTables As Boolean = False
        Dim sGLPKModelDir As String = ""
        Dim sGnuWinDir As String = ""
        Dim uniqueBarrierEIDComparer As FindBarrierEIDPredicate

        'm_LLayersFields2.Clear() ' To be safe clear these
        'm_PLayersFields2.Clear() ' - issues during debugging found

        ' If settings have been set by the user then load them from the extension stream (stored in .mxd doc)
        If m_FiPEx__1.m_bLoaded = True Then

            ' to do 2020: 'loadproperties should be a 'shared function' 

            Try

                sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))
                iOrderNum = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("ordernum"))     'May have problems here - need to convert to integer
                bMaximum = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("maximum"))
                'If they want the maximum then set ordernum to very large number (haven't found a better way yet)
                If bMaximum = True Then
                    iOrderNum = 999
                End If
                bConnectTab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("connecttab"))
                bBarrierPerm = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("barrierperm"))
                bNaturalYN = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("NaturalYN"))
                bDCI = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("dciyn"))
                bDCISectional = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("dcisectionalyn"))

                bDBF = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDBF"))
                sGDB = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGDB"))
                sPrefix = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("TabPrefix"))

                '2020
                Try
                    bUseHabLength = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bUseHabLength"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property UseHabLength. Setting it to True.")
                    bUseHabLength = True
                End Try
                Try
                    bUseHabArea = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bUseHabArea"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property bUseHabArea. Setting it to False.")
                    bUseHabArea = False
                End Try
                Try
                    bDistanceDecay = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDistanceDecay"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property bDistanceDecay. Setting it to False.")
                    bDistanceDecay = False
                End Try

                Try
                    bDistanceLim = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDistanceLim"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property bDistanceLim. Setting it to False.")
                    bDistanceDecay = False
                End Try

                Try
                    sDDFunction = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sDDFunction"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property sDDFunction. Setting it to 'none'")
                    sDDFunction = "none"
                End Try


                Try
                    dMaxDist = Convert.ToDouble(m_FiPEx__1.pPropset.GetProperty("dMaxDist"))
                Catch ex As Exception
                    'MsgBox("Trouble loading FIPEX property dMaxDist. Setting it to 0.")
                    dMaxDist = 0.0
                End Try

                bAdvConnectTab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("advconnecttab"))

                bUpHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("UpHab"))
                bTotalUpHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalUpHab"))
                bDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("DownHab"))
                bTotalDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalDownHab"))
                bPathDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("PathDownHab"))
                bTotalPathDownHab = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("TotalPathDownHab"))

                iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))

                PolyLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)

                ' Populate a list of the layers to use and habitat summary fields.
                ' match any of the polygon layers saved in stream to those in listboxes 
                If iPolysCount > 0 Then
                    For k = 0 To iPolysCount - 1
                        'sPolyLayer = m_FiPEX__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
                        PolyLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
                        With PolyLayerObj
                            .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
                            .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
                            .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
                            .HabUnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
                        End With

                        ' Load that object into the list
                        pPLayersFields.Add(PolyLayerObj)  'what are the brackets about - this could be aproblem!!
                    Next
                End If

                ' Need to be sure that quantity field has been assigned for each
                ' layer using. 
                Dim iCount1 As Integer = pPLayersFields.Count

                If iCount1 > 0 Then
                    For m = 0 To iCount1 - 1
                        sTemp = pPLayersFields.Item(m).Layer
                        If pPLayersFields.Item(m).HabQuanField = "Not set" Then
                            System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
                            Exit Sub
                        End If
                    Next
                End If

                iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
                LineLayerObj = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

                ' match any of the line layers saved in stream to those in listboxes
                If iLinesCount > 0 Then
                    For j = 0 To iLinesCount - 1
                        'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                        LineLayerObj = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                        With LineLayerObj

                            .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))

                            .LengthField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthField" + j.ToString))
                            .LengthUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthUnits" + j.ToString))

                            .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabClassField" + j.ToString))
                            .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabQuanField" + j.ToString))
                            .HabUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabUnits" + j.ToString))

                        End With
                        pLLayersFields.Add(LineLayerObj)
                    Next
                End If

                ' Need to be sure that quantity field has been assigned for each
                ' layer using. 
                iCount1 = pLLayersFields.Count
                If iCount1 > 0 Then
                    For m = 0 To iCount1 - 1
                        If pLLayersFields.Item(m).HabQuanField = "Not set" Then
                            System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for line layer. Please choose a field in the options menu.", "Parameter Missing")
                            Exit Sub
                        End If
                        If pLLayersFields.Item(m).LengthField = "Not set" Then
                            System.Windows.Forms.MessageBox.Show("No length field set for line layer. Please choose a field in the options menu.", "Parameter Missing")
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
                        ' TEMP HARD CODE OF BARRIER TYPE - this contains a FIELD to look for
                        ' in each layer for barrier type ie. 'culvert' 'dam' -- thesis SA
                        pBarrierIDObj = New BarrierIDObj(sBarrierIDLayer, sBarrierIDField, sBarrierPermField, sBarrierNaturalYNField, "FIPEX_BarrierType")
                        lBarrierIDs.Add(pBarrierIDObj)
                    Next
                End If

                bGLPKTables = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bGLPKTables"))
                sGLPKModelDir = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGLPKModelDir"))
                sGnuWinDir = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGnuWinDir"))

                'TEMP HARDCODING
                bGLPKTables = False

                sGLPKModelDir = "c:\GunnsModel\"
                'sGnuWinDir = "c:\GnuWin32\"

            Catch ex As Exception
                MsgBox("Error when attempting to load stored FIPEX properties. This is normal if FIPEX has been updated. Try setting options again, saving document, closing document, and re-opening. " + ex.Message)
                Exit Sub
            End Try

        Else
            ' TODO: Pop-up form with current settings
            ' Add a button on form that loads options
            ' upon close, if options form within for has been opened
            ' then keep looping
            Dim sMessage As String = "Please set analysis options in menu and re-run tool"
            System.Windows.Forms.MessageBox.Show(sMessage, "Parameters Missing")
            Exit Sub
        End If


        ' =============== SAVE ORIGINAL GEONet SETTINGS =========================
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(8, "Saving Current Geometric Network Flags and Barriers")

        pNetworkAnalysisExt = CType(m_UNAExt, INetworkAnalysisExt)
        Dim pGeometricNetwork As IGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

        pNetwork = pGeometricNetwork.Network
        pNetElements = CType(pNetwork, INetElements)

        ' QI for the ITraceTasks interface using IUtilityAnalysisExt
        pTraceTasks = CType(m_UNAExt, ITraceTasks)

        ' update extension with the results
        ' QI for the INetworkAnalysisExtResults interface using IUTilityNetworkAnalysisExt
        pNetworkAnalysisExtResults = CType(m_UNAExt, INetworkAnalysisExtResults)

        ' clear any leftover results from previous calls to the cmd
        pNetworkAnalysisExtResults.ClearResults()

        ' QI the Flags and barriers
        pNetworkAnalysisExtFlags = CType(m_UNAExt, INetworkAnalysisExtFlags)
        pNetworkAnalysisExtBarriers = CType(m_UNAExt, INetworkAnalysisExtBarriers)

        pOriginalBarriersListGEN = New EnumNetEIDArray
        pOriginalEdgeFlagsListGEN = New EnumNetEIDArray
        pOriginaljuncFlagsListGEN = New EnumNetEIDArray

        ' 2020 note - these saved' objects are for the 'advanced connectivity' 
        '             'saved' is so if user selects advanced connectivity table 
        '              this object will not be modified and can be user to restor
        Dim pOriginalBarriersListSaved As IEnumNetEID
        Dim pOriginalBarriersListSavedGEN As IEnumNetEIDBuilderGEN
        pOriginalBarriersListSavedGEN = New EnumNetEIDArray

        Dim iOriginalJunctionBarrierCount As Integer
        iOriginalJunctionBarrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount

        ' Save the barriers
        For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
            ' Use bFlagDisplay to retrieve EIDs of the barriers for later
            bFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
            bEID = pNetElements.GetEID(bFlagDisplay.FeatureClassID, bFlagDisplay.FID, bFlagDisplay.SubID, esriElementType.esriETJunction)
            pOriginalBarriersListGEN.Add(bEID)
            pOriginalBarriersListSavedGEN.Add(bEID)
            'originalBarriersList(i) = bEID
        Next

        ' QI to and get an array object that has 'count' and 'next' methods
        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

        ' 2020: for branch connectivity feature - GO
        ' save the original list so it can be restored later
        pOriginalBarriersListSaved = CType(pOriginalBarriersListSavedGEN, IEnumNetEID)
        'MsgBox("Number of original barriers: " & CStr(pOriginalBarriersListSaved.Count))

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
        'pOriginaljuncFlagsList = 
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

        ' If there are no flags set exit sub
        If pNetworkAnalysisExtFlags.JunctionFlagCount = 0 Then
            'If pNetworkAnalysisExtFlags.EdgeFlagCount = 0 Then
            MsgBox("There are no flags set on junctions.  Please Set flags only on network junctions.")
            Exit Sub
            'End If
        End If

        ' =============== FLAG CONSISTENCY CHECK ====================
        ' Check for consistency that flags are all on barriers or all on non-barriers.
        Dim sFlagCheck As String
        sFlagCheck = flagcheck(pOriginalBarriersList, pOriginalEdgeFlagsList, pOriginaljuncFlagsList)
        '    MsgBox "FlagCheck may be null and crash..."
        '    MsgBox "sFlagcheck: " + sFlagCheck
        If sFlagCheck = "error" Then
            MsgBox("FIPEX Error 109: Please check that all flags are consistently on barriers or non-barriers.")
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
        i = 0

        ' =================== GET WORKSPACE ========================
        ' Prepare for creating output tables

        'Dim pWorkspaceName As IWorkspaceName = New WorkspaceName
        Dim sConnectTabName As String = ""
        Dim sAdvConnectTabName As String = ""

        ' =============== For DCI Table Output =====================

        ' obtain reference to current geometric network
        ' and get all participating point feature classes so only
        ' those are included in the barriers list. 
        pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

        ' Get all simple edges participating in the geometric network
        Dim pEnumSimpLineFCs As ESRI.ArcGIS.Geodatabase.IEnumFeatureClass = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTSimpleEdge)
        ' No complex edge support yet... 
        Dim pEnumCompLineFCs As ESRI.ArcGIS.Geodatabase.IEnumFeatureClass = pGeometricNetwork.ClassesByType(esriFeatureType.esriFTComplexEdge)

        Dim pFeatureClass As IFeatureClass
        Dim pFeatureLayer As IFeatureLayer
        Dim bEdgeMatch As Boolean ' load all layers in TOC to the lstbox
        Dim sHabTableName As String
        Dim lFCIDs As New List(Of FCIDandNameObject) ' to hold FCIDs of included line layers
        Dim lAllFCIDs As New List(Of FCIDandNameObject) ' to hold FCIDs of included line layers
        Dim pFCIDandNameObject As New FCIDandNameObject(Nothing, Nothing)
        Dim pMetricsObject As New MetricsObject(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim lMetricsObject As New List(Of MetricsObject)
        Dim iLineFCID, iPolyFCID As Integer
        Dim bLineFCIDFound, bPolyFCIDFound As Boolean
        Dim pLayer As ILayer

        ' 2020 - added a polygon object so DCI can use areas as well as lengths
        Dim lLLayersFieldsDCI As New List(Of LineLayerToAdd)
        Dim lPLayersFieldsDCI As New List(Of PolyLayerToAdd)

        ' for GLPK - 2020 - note the LayerToAdd object was split into two groups for lines and polys
        '                   and this issue would need to be fixed
        Dim lAllLayersFieldsGLPK As New List(Of LineLayerToAdd)

        Dim lGLPKUniqueEIDs As List(Of Integer)
        Dim sDCITableName, sMetricTableName As String
        Dim sFeatureLayerName As String
        Dim sGLPKHabitatTableName, sGLPKOptionsTableName, sGLPKConnectTabName As String
        ' Need to get the FC Id's of the line layers because
        ' names might change and there might be duplicate names, names
        ' that repeat but are different FC's, (weird stuff like that)
        ' NOTE: this should have been done in Options code and stored in Extension settings.
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(10, "Getting participating layers from the Table of Contents")

        'MsgBox("Debug:2")
        i = 0
        j = 0
        m = 0
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

        ' Add the polygon layers to the ALL FCIDs list\
        ' this is use later in the predicate search to get unique 
        ' feature classes to draw out from the master habitat object

        'MsgBox("Debug:3")
        i = 0
        j = 0
        m = 0
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

        ' ============== Begin DBF Table Preparation =======================
        ' 1 If tables are to be output in DBF format
        '   2 For each of the line layers included
        '     2.1 Get a table name (tablename function)
        '     2.2 Add that table name to a list for use... 
        '     2.3 Set up the fields for that table
        '     2.4 Create that table (PrepDBFOutTable Subroutine)
        '     2.5 Add the table to the TOC in ArcMap
        '     3.0a If the flag is not on a barrier
        '       3.1a Get a table name and add it to a list ("flagadjacent habitat")
        '       3.2a Create the table and add it to the TOC in ArcMap
        '       4.0a If the total impounded/impacted habitat is wanted
        '          4.1a Create fields and table for those statistics
        '       3.0b Else if the flag is on a barrier
        '       4.0b If the total impounded/impacted habitat is wanted
        '         4.1b Create fields and table for those statistics
        ' **** Repeat above process for each polygon layer ****

        'MsgBox("Debug:4")
        If bDBF = True Then

            ' Get the path to the database
            pFWorkspace = GetWorkspace(sGDB)
            If pFWorkspace Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Could not get reference to output workspace.  " _
                + "Please check output geodatabase in options menu", "Output Workspace Error")
                Exit Sub
            End If

            ' Change the active tab in the TOC to 'source'
            ' to show tables
            Dim pTOC As IContentsView
            pTOC = pMxDocument.ContentsView(1)
            pMxDocument.CurrentContentsView = pTOC

            If bDCI = True Then

                'MsgBox("Debug:5")
                ' Make sure that each of the included line layers are members of the 
                ' geometric network (and simple line feature classes). This is because
                ' lines are not currently intersected (only polygons).  

                'Initial set of feature class
                pEnumSimpLineFCs.Reset()
                pFeatureClass = pEnumSimpLineFCs.Next()
                i = 0
                j = 0
                If lFCIDs.Count > 0 Then
                    Do Until pFeatureClass Is Nothing
                        For i = 0 To lFCIDs.Count - 1
                            If lFCIDs(i).FCID = pFeatureClass.FeatureClassID Then
                                For j = 0 To pLLayersFields.Count - 1
                                    If pLLayersFields(j).Layer = lFCIDs(i).Name Then
                                        lLLayersFieldsDCI.Add(pLLayersFields(j))
                                    End If
                                Next 'j
                            End If
                        Next 'i
                        pFeatureClass = pEnumSimpLineFCs.Next()
                    Loop
                End If

                ' 2020 added a list of polygon layers and fields
                ' no need to do checks as above
                j = 0
                For j = 0 To pPLayersFields.Count - 1
                    lPLayersFieldsDCI.Add(pPLayersFields(j))
                Next

            End If
            ' ============== End for DCI Output ===============

            ' ###################################################
            ' ###################################################
            ' ############# FOR GLPK ############################
            ' do not delete

            ' For GLPK Filter out the lines that are not part of the network 
            ' as above, but keep all polygons included in the analysis
            lAllLayersFieldsGLPK = New List(Of LineLayerToAdd)
            ' 2020 - the above object needs to be fixed for GLPK to work
            '        because LayerToAdd object has been split into two objects
            '        one for line, one for poly

            'If bGLPKTables = True Then
            '    'Initial set of feature class
            '    pEnumSimpLineFCs.Reset()
            '    pFeatureClass = pEnumSimpLineFCs.Next()
            '    i = 0
            '    j = 0
            '    If lAllFCIDs.Count > 0 Then
            '        Do Until pFeatureClass Is Nothing
            '            For i = 0 To lAllFCIDs.Count - 1
            '                If lAllFCIDs(i).FCID = pFeatureClass.FeatureClassID Then
            '                    For j = 0 To pLLayersFields.Count - 1
            '                        If pLLayersFields(j).Layer = lAllFCIDs(i).Name Then
            '                            lAllLayersFieldsGLPK.Add(pLLayersFields(j))
            '                        End If
            '                    Next 'j                              
            '                End If
            '            Next 'i
            '            pFeatureClass = pEnumSimpLineFCs.Next()
            '        Loop
            '    End If
            '    j = 0
            '    For j = 0 To pPLayersFields.Count - 1
            '        lAllLayersFieldsGLPK.Add(pPLayersFields(j))
            '    Next 'j
            'End If
            ' ############################################################
            ' ######## end for GLPK ######################################


            ' 2020 unsure if below commented out section is for GLPK or not
            'Else ' If there is no DBF Table Output

            '' Still need the table names for output to dockable window
            '' Will use names to title each section of dockable window
            '' -------------------------------------------------------
            '' 1.0 For each line layer included
            '' 1.1 Get the name of the layer for TITLE for each output block
            '' 2.0a If the flag is not on a barrier
            ''   2.1a Add the word 'Accessible' before the title
            ''   2.2a Add this name to a list of titles for 'unimpacted/adjacent habitat'
            ''   3.0a If total impacted is needed
            ''     3.1a Add 'TtlAffctd' before the title and add it to a title list
            ''         of 'impacted/impounded habitat'
            '' 2.0b Else if the flag is on a barrier
            ''   3.0b If total impacted is needed 
            ''     3.1b Add 'TtlAffctd' before the title and add it to a title list
            ''          of 'impacted/impounded habitat'
            '' **** Repeat this for polygon layers included ****

            'For i = 0 To pLLayersFields.Count - 1
            '    lTableNames.Add(pLLayersFields(i).Layer)

            '    If sFlagCheck = "nonbarr" Then
            '        'sTableName = "FlagAdjacentHabitat" + pLLayersFields(i).Layer
            '        'lFlagTabNamesUImp.Add(sTableName)
            '        If bTotalHab = True Then
            '            sTableName = "TotalHabitat_" + pLLayersFields(i).Layer
            '            lFlagTabNamesImp.Add(sTableName)
            '        End If
            '    ElseIf sFlagCheck = "barriers" Then

            '        If bTotalHab = True Then
            '            sTableName = "TotalHabitat_" + pLLayersFields(i).Layer
            '            lBarTabNamesTtlImp.Add(sTableName)
            '        End If
            '    End If
            'Next

            '' Get names for all polygon layers using
            'For i = 0 To pPLayersFields.Count - 1
            '    lTableNames.Add(pPLayersFields(i).Layer)

            '    If sFlagCheck = "nonbarr" Then
            '        'sTableName = "FlagAdjacentHabitat" + pPLayersFields(i).Layer
            '        'lFlagTabNamesUImp.Add(sTableName)
            '        If bTotalHab = True Then
            '            sTableName = "TotalHabitat_" + pPLayersFields(i).Layer
            '            lFlagTabNamesImp.Add(sTableName)
            '        End If
            '    ElseIf sFlagCheck = "barriers" Then
            '        If bTotalHab = True Then
            '            sTableName = "TotalHabitat_" + pPLayersFields(i).Layer
            '            lBarTabNamesTtlImp.Add(sTableName)
            '        End If
            '    End If
            'Next ' polygon layer
        End If ' DBF layers are included


        ' ========================== Begin Traces ====================

        'MsgBox("Debug:7")
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(10, "Beginning Network Traces")

        barrierCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount
        Dim pSymbol As ISymbol
        Dim pJuncFlagDisplay As IJunctionFlagDisplay
        Dim sOutID As String 'ID of flag for output table
        Dim pAllFlowEndBarriersGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pAllFlowEndBarriers As IEnumNetEID
        Dim pNextOriginalJuncFlagGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pNextOriginalJuncFlag As IEnumNetEID
        Dim fEID As Integer ' to hold the flag ID for the next flag
        Dim pUID As New UID
        Dim pEnumLayer As IEnumLayer
        ' for connectivity info:
        Dim pBarrierAndDownstreamEID As New BarrierAndDownstreamID(Nothing, Nothing, Nothing, Nothing)
        Dim pGLPKBarrAndDownEID As New GLPKBarrierAndDownstreamEID(Nothing, Nothing)
        Dim lConnectivity As New List(Of BarrierAndDownstreamID)
        Dim lGLPKConnectivity As New List(Of GLPKBarrierAndDownstreamEID)
        Dim sFlowEndID, sHabDir, sTableType, sType As String
        Dim sLastUserSetID As String = ""
        Dim dDCIp, dDCId, dPerm As Double ' holds DCIp and DCId stats, 9999 if error
        Dim bNoPerm As Boolean ' holds trigger to check if a single barrier does have a zero permeability
        Dim bNaturalY As Boolean = False ' holds trigger to check if any barrier is labeled as 'Natural' - 
        ' determines how output DCI file is read

        Dim pDownConnectedJunctions, pDownConnectedEdges As IEnumNetEID
        Dim pSubtractUpJunctions, pSubtractUpEdges As IEnumNetEID
        Dim pConnectedJunctions, pConnectedEdges As IEnumNetEID
        Dim pIDAndType As New IDandType(Nothing, Nothing)
        Dim pGLPKOptionsObject As New GLPKOptionsObject(Nothing, _
                                                        Nothing, _
                                                        Nothing, _
                                                        Nothing, _
                                                        Nothing)
        Dim iMaxOptionNum As Integer = 0

        ' https://support.esri.com/en/technical-article/000008702
        ' {E156D7E5-22AF-11D3-9F99-00C04F6BC78E} IGeoFeatureLayer
        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        pEnumLayer = pMap.Layers(pUID, True)

        ' =========== GET SYMBOLS ============

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

        ' ============= END GET SYMBOLS ==============

        iProgress = 10
        Dim iProgressIncrementFactor As Integer
        Try
            iProgressIncrementFactor = Convert.ToInt16(30 / pOriginaljuncFlagsList.Count)
        Catch ex As Exception
            MsgBox("Error Converting ProgressIncrementFactor")
        End Try

        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################

        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################

        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################

        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################

        ' ######################### FIND ALL BRANCH JUNCTIONS ##############
        ' Created Aug 2020
        ' Purpose: testing, debugging of finding all junctions representing 
        '          a branch _in addition_ to the junctions representing 
        '          sinks or barriers. Create list of junctions that will 
        '          be treated like 'barriers' for purpose of connectivity 
        '          table generation; amend list of original junction barriers.
        '          the 'original' in object name should be changed since the
        '          actual originally set barriers are stored in pOriginalBarriersListSaved
        ' Logic: 
        ' If 'Advanced Connectivity is checked' 
        ' Clear All Flags and Barriers
        ' For Each Original Flag
        ' Set Flag 
        ' Trace Upstream and find neighbour junction
        ' If more than one upstream immediate junction 
        ' Save current flag junction EID
        ' Move flag to next 'order' and repeat above

        ' will need: pOriginaljuncFlagsList.Reset() see below
        'MsgBox("Debug2020: bAdvConnectTab is set " & Str(bAdvConnectTab))
        Dim pNetTopology As INetTopology
        Dim iEdges As Integer
        Dim iEdgeEID As Integer = 0
        Dim iFromEID As Integer = 0 ' referring to dig. direction
        Dim iToEID As Integer = 0
        Dim iThisEID As Integer = 0
        Dim iNextEID As Integer = 0
        Dim iEID_j As Integer = 0
        Dim iEID_p As Integer = 0
        Dim iEID_k As Integer = 0
        Dim iLastEID As Integer = 0
        Dim iActualFromEID As Integer = 0 ' for actual flow direction
        Dim iActualToEID As Integer = 0
        Dim iFlowDir As Integer = 0 ' esriFlowDirection 
        Dim bOrientation As Boolean = False
        Dim bUpstream As Boolean = False
        Dim bMatch As Boolean = False
        Dim bBranchJunction As Boolean = False
        Dim bSourceJunction As Boolean = False
        Dim iUpstreamEdges As Integer = 0
        Dim pAdjacentEdges As IEnumNetEID
        Dim pAdjacentEdgesGEN As IEnumNetEIDBuilderGEN
        Dim pBranchJunctions As IEnumNetEID
        Dim pBranchJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim pFilteredBranchJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFilteredBranchJunctionsList As IEnumNetEID
        Dim pSourceJunctions As IEnumNetEID
        Dim pSourceJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim pFilteredSourceJunctionsGEN As IEnumNetEIDBuilderGEN = New EnumNetEIDArray
        Dim pFilteredSourceJunctionsList As IEnumNetEID


        ' next junctions upstream direction
        Dim pNextJunctions As IEnumNetEID
        Dim pNextJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim pJunctions As IEnumNetEID
        Dim pJunctionsGEN As IEnumNetEIDBuilderGEN
        Dim iUpstreamJunctionCount As Integer = 0
        Dim pUtilityNetwork As IUtilityNetwork ' need to get flow direction relative to dig. direction
        Dim pForwardStar As IForwardStar
        Dim pForwardStarGEN As IForwardStarGEN
        Dim pNetEdge As INetworkEdge
        Dim iNetEdgeDirection As Integer ' the direction of flow relative to digitized direction


        If bAdvConnectTab = True Then

            '   If this is the first iteration
            'lHabStatsList = New List(Of StatisticsObject_2) ' set the habitat stats list

            pOriginaljuncFlagsList.Reset()
            pBranchJunctionsGEN = New EnumNetEIDArray
            pSourceJunctionsGEN = New EnumNetEIDArray

            ' for each flag the user has set (usually assuming the flag is on the furthest
            ' downstream nodes
            ' i.e., the outflow(s) or sink(s) of the network(s)
            For i = 0 To pOriginaljuncFlagsList.Count - 1


                ' check if user has hit 'close/cancel'
                If m_bCancel = True Then
                    backgroundworker1.CancelAsync()
                    backgroundworker1.Dispose()
                    Exit Sub
                End If
                backgroundworker1.ReportProgress(10, "Performing Network Traces for Advanced Connectivity Analysis Step 1 " & _
                                                 ControlChars.NewLine & _
                                                 "User Flag " & (i + 1).ToString & _
                                                 " of " & (pOriginaljuncFlagsList.Count).ToString)
                ' reset the counter

                ' ##### NEW CODE AUG 2020 #####
                ' https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#INetTopology.htm
                ' better / worse alternative? https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#IForwardStarGEN.htm

                ' IGeometricNetwork
                ' 
                'pGeometricNetwork already set
                'pNetwork already set,
                'pNetwork = pGeometricNetwork.Network

                pNetTopology = CType(pNetwork, INetTopology)
                fEID = pOriginaljuncFlagsList.Next()

                ' Use Do loop until no upstream neighbours are found
                pJunctionsGEN = New EnumNetEIDArray
                pJunctionsGEN.Add(fEID) ' initialize this object before first loop
                pJunctions = CType(pJunctionsGEN, IEnumNetEID)
                iUpstreamJunctionCount = pJunctions.Count ' initialize for first loop
                pUtilityNetwork = CType(pGeometricNetwork.Network, IUtilityNetwork)

                pForwardStarGEN = pUtilityNetwork.CreateForwardStar(True, Nothing, Nothing, Nothing, Nothing)
                pForwardStar = CType(pForwardStarGEN, IForwardStar)

                Do Until iUpstreamJunctionCount = 0

                    pJunctions.Reset()
                    pNextJunctionsGEN = New EnumNetEIDArray

                    For j = 0 To pJunctions.Count - 1
                        iThisEID = pJunctions.Next()
                        iUpstreamEdges = 0

                        'pForwardStar.FindAdjacent(0, iThisEID, iEdges)
                        iEdges = pNetTopology.GetAdjacentEdgeCount(iThisEID)
                        'MsgBox("Debug2020: # edge neighbours found for node EID " & Str(iThisEID) & "using pForwardStar: " & Str(iEdges))

                        If iEdges > 0 Then
                            'If pNetTopology.GetAdjacentEdgeCount(iThisEID) > 0 Then
                            'pEnumNetEIDBuilder.Add(lEID)
                            'MsgBox("Debug2020: neighbours found for flag " & Str(fEID))
                            ' if edge count is greater than 2 (one uo one down)
                            ' add current junction to list of branch junctions
                            ' for each edge find adjacent junctions
                            'iEdges = pNetTopology.GetAdjacentEdgeCount(iThisEID)
                            'MsgBox("Debug2020: # edge neighbours found for node EID " & Str(iThisEID) & "using pNetTopology: " & Str(iEdges))

                            For k = 0 To iEdges - 1

                                iEdgeEID = 0
                                iFromEID = 0
                                iToEID = 0

                                'MsgBox("Debug2020: # This junction  EID: " & Str(iThisEID))
                                'MsgBox("Debug2020: edgeEID initialized to: " & Str(iEdgeEID))
                                pNetTopology.GetAdjacentEdge(iThisEID, k, iEdgeEID, True)
                                'MsgBox("Debug2020: neighbour edge EID found using pNetTopology: " & Str(iEdgeEID))

                                'pForwardStar.QueryAdjacentEdge(k, iEdgeEID, bOrientation, Nothing)
                                ' MsgBox("Debug2020: neighbour edge EID found using pForwardStar: " & Str(iEdgeEID))
                                'MsgBox("Debug2020: is neighbour edge " & Str(iEdgeEID) & " found using pForwardStar upstream? " & Str(bOrientation))



                                ' Problem is that bOrientation (and other nettopology directions) are simply digitization direction
                                ' no way easily to determine the flow direction of the edges
                                ' need to:
                                ' get the adjacent edge
                                ' determine digitization direction relative to iThisEID
                                ' determine flow direction relative to digization direction
                                ' is it upstream?

                                ' IUtilityNetworkGEN.GetFlowDirection(EID) should get flow direction but it is relative to digitization
                                ' https://desktop.arcgis.com/en/arcobjects/latest/net/webframe.htm#esriFlowDirection.htm
                                ' 1 = with digitization direction, 2 = against digitization direction
                                'pUtilityNetwork.

                                'iToEID = iThisEID

                                'MsgBox("Debug2020: # The initialized next junction EID: " & Str(iNextJunctionEID))
                                pNetTopology.GetFromToJunctionEIDs(iEdgeEID, iFromEID, iToEID)
                                'MsgBox("Debug2020: # The returned 'to' node using PNetTopology for edge " & Str(iEdgeEID) & " :" & Str(iToEID))
                                'MsgBox("Debug2020: # The returned 'from' node using PNetTopology for edge " & Str(iEdgeEID) & " :" & Str(iFromEID))

                                'pForwardStar.QueryAdjacentJunction(k, iFromEID, Nothing)
                                'MsgBox("Debug2020: # The neighbour node using pForwardStar for edge " & Str(iEdgeEID) & " :" & Str(iFromEID))

                                iFlowDir = pUtilityNetwork.GetFlowDirection(iEdgeEID)
                                'MsgBox("Debug2020: # The esriflowdirection for edge " & Str(iEdgeEID) & " :" & Str(iFlowDir))

                                If iFromEID = iThisEID Then
                                    iNextEID = iToEID
                                    If iFlowDir = 1 Then
                                        'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is downstream of node " & Str(iThisEID))
                                        bUpstream = False
                                    ElseIf iFlowDir = 2 Then
                                        'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is upstream of node " & Str(iThisEID))
                                        bUpstream = True
                                    Else
                                        MsgBox("Debug2020: Edge flow direction is unitialized or indeterminate. For Distance-limited analysis all edges must have determinate flow direction. ")
                                        MsgBox("Debug2020: # The esriflowdirection for edge EID " & Str(iEdgeEID) & " :" & Str(iFlowDir))
                                        'pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, _
                                        '  iFCID, _
                                        '  iFID, _
                                        '  iSubID)
                                        bUpstream = False
                                    End If

                                ElseIf iToEID = iThisEID Then
                                    iNextEID = iFromEID
                                    If iFlowDir = 1 Then
                                        'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is upstream of node " & Str(iThisEID))
                                        bUpstream = True
                                    ElseIf iFlowDir = 2 Then
                                        'MsgBox("Debug2020: Edge " & Str(iEdgeEID) & " is downstream of node " & Str(iThisEID))
                                        bUpstream = False
                                    Else
                                        'MsgBox("Debug2020: Edge flow direction is unitialized or indeterminate")
                                        bUpstream = False
                                    End If
                                Else
                                    MsgBox("Debug2020: No match?? iThisEID: " & Str(iThisEID) & " iFromEID: " & Str(iFromEID) & "iToEID: " & Str(iToEID))
                                End If

                                ' add the upstream EID to a list of EIDs unless we are 
                                ' Note here would be the place to identify braids, if that is desired
                                '    Do this after this loop by checking if any duplicates in list of 
                                '    upstream EIDs
                                If bUpstream = True Then
                                    pNextJunctionsGEN.Add(iNextEID)
                                    iUpstreamEdges = iUpstreamEdges + 1
                                End If
                            Next
                            ' MsgBox("Debug2020: # of UPSTREAM neighbour edges found for Node " & Str(iThisEID) & ": " & Str(iUpstreamEdges))

                        End If ' upstream edges > 0
                        ' Add to list of branching junctions

                        If iUpstreamEdges > 1 Then
                            pBranchJunctionsGEN.Add(iThisEID)
                            'MsgBox("Debug2020: Found a branch junction at node: " & Str(iThisEID))
                            ' going to add this EID to list and remove duplicates later
                            'pOriginalBarriersListGEN.Add(iThisEID)
                        ElseIf iUpstreamEdges = 0 Then
                            'MsgBox("Debug2020 Source found at " & CStr(iThisEID))
                            pSourceJunctionsGEN.Add(iThisEID)
                        End If

                        'iLastEID = iThisEID ' track last loops EID 

                    Next 'junction in 'order
                    pNextJunctions = Nothing
                    pNextJunctions = CType(pNextJunctionsGEN, IEnumNetEID)
                    pNextJunctions.Reset()

                    'MsgBox("Debug2020: # of upstream junctions this 'order': " & Str(iUpstreamEdges))

                    ' not sure it's necessary to redeclare an empty object for junctions but will do
                    pJunctionsGEN = New EnumNetEIDArray
                    pJunctions = Nothing
                    pJunctions = CType(pJunctionsGEN, IEnumNetEID)
                    pJunctions = pNextJunctions
                    pNextJunctionsGEN = New EnumNetEIDArray

                    iUpstreamJunctionCount = pJunctions.Count()

                Loop

                ' MsgBox("Debug2020: # of branching junctions for network ': " & Str(pBranchJunctions.Count))

                ' merge the original barriers list and the branch junction list
                ' careful because the original barriers list likely has other objects
                ' linked to it that contain permeability, ID, etc etc
                ' remove duplicates

            Next ' original junction flag or sink

            ' For each pBranchJunctions add it to the original barrier list generator
            pBranchJunctions = CType(pBranchJunctionsGEN, IEnumNetEID)

            pBranchJunctions.Reset()
            For j = 0 To pBranchJunctions.Count - 1
                iEID_j = pBranchJunctions.Next()
                'MsgBox("Debug2020 iIEID_j: " + Str(iEID_j))

                ' make sure it's not a duplicate 
                '(if user has placed flag on branch junction, then omit the branch junction)
                pOriginalBarriersList.Reset()
                bMatch = False

                'MsgBox("Debug2020 pOriginalBarriersList.Count: " + Str(pOriginalBarriersList.Count))
                For k = 0 To pOriginalBarriersList.Count - 1
                    iEID_k = pOriginalBarriersList.Next()
                    'MsgBox("Debug2020 iIEID_k: " + Str(iEID_k))
                    If iEID_j = iEID_k Then
                        bMatch = True
                        'MsgBox("Debug2020 iIEID_k " + Str(iEID_k) + " = iEID_j " + Str(iEID_j))

                    End If
                Next
                If bMatch = False Then
                    ' add branch junction to user-set barrier list 
                    pOriginalBarriersListGEN.Add(iEID_j)
                    ' store filtered (no user-set barrier) branch junction for later
                    pFilteredBranchJunctionsGEN.Add(iEID_j)
                    'MsgBox("Debug2020 adding iIEID_j to list of branch junctions: " + Str(iEID_j))

                End If
            Next

            pFilteredBranchJunctionsList = Nothing
            pFilteredBranchJunctionsList = CType(pFilteredBranchJunctionsGEN, IEnumNetEID)
            pFilteredBranchJunctionsList.Reset()

            pSourceJunctions = CType(pSourceJunctionsGEN, IEnumNetEID)
            pSourceJunctions.Reset()
            'MsgBox("Debug2020: number of source junctions " + CStr(pSourceJunctions.Count))

            For j = 0 To pSourceJunctions.Count - 1
                iEID_j = pSourceJunctions.Next()
                'MsgBox("Debug2020 iIEID_j: " + Str(iEID_j))

                ' make sure it's not a duplicate 
                '(if user has placed flag on branch junction, then omit the branch junction)
                pOriginalBarriersList.Reset()
                bMatch = False

                'MsgBox("Debug2020 pOriginalBarriersList.Count: " + Str(pOriginalBarriersList.Count))
                For k = 0 To pOriginalBarriersList.Count - 1
                    iEID_k = pOriginalBarriersList.Next()
                    'MsgBox("Debug2020 iIEID_k: " + Str(iEID_k))
                    If iEID_j = iEID_k Then
                        bMatch = True
                        'MsgBox("Debug2020 iIEID_k " + Str(iEID_k) + " = iEID_j " + Str(iEID_j))

                    End If
                Next
                If bMatch = False Then
                    ' add branch junction to user-set barrier list 
                    pOriginalBarriersListGEN.Add(iEID_j)
                    ' store filtered (no user-set barrier) branch junction for later
                    pFilteredSourceJunctionsGEN.Add(iEID_j)
                    'MsgBox("Debug2020 adding iIEID_j to list of source junctions and to list of barriers: " + Str(iEID_j))

                End If
            Next

            pFilteredSourceJunctionsList = Nothing
            pFilteredSourceJunctionsList = CType(pFilteredSourceJunctionsGEN, IEnumNetEID)
            pFilteredSourceJunctionsList.Reset()

            pOriginalBarriersList = Nothing
            pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

        End If

        'MsgBox("Debug2020 Number of saved barriers after branch search: " & Str(pOriginalBarriersListSaved.Count))

        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################
        ' 2020_2020 ########################################################


        '###################### END ADVANCED CONNECTIVITY TAB ##############

        pNextOriginalJuncFlagGEN = New EnumNetEIDArray
        pNextOriginalJuncFlagGEN.Add(fEID)
        pNextOriginalJuncFlag = CType(pNextOriginalJuncFlagGEN, IEnumNetEID)

        ' For each order
        '   If this is the first iteration
        '     If this is a "Non Barrier" analysis
        '       For each original junction flag
        '       Clear all flags
        '       Get the element ID of the flag
        '       Get the user ID's of the flag
        '       Display the flag / Set as flag in geonet
        '       Run the TraceFlowSolverSetup Sub (setup network)
        '       Run the trace on all flags to get the set of barriers first
        '       encountered.
        '         If no results were returned
        '         Create empty objects for returned edges and junctions
        '       Make trace results a selection
        '       Get the features stopping the trace 
        '       Filter the features stopping the trace (no sinks/sources/non-barriers)
        '       If the current flag is on a point with field that matches

        lHabStatsList = New List(Of StatisticsObject_2) ' set the habitat stats list

        pOriginaljuncFlagsList.Reset()
        For i = 0 To pOriginaljuncFlagsList.Count - 1
            ' check if user has hit 'close/cancel'
            If m_bCancel = True Then
                backgroundworker1.CancelAsync()
                backgroundworker1.Dispose()
                Exit Sub
            End If
            backgroundworker1.ReportProgress(10, "Performing Network Traces " & _
                                             ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & _
                                             " of " & (pOriginaljuncFlagsList.Count).ToString)

            'MsgBox("Debug:8")
            ' reset the counter
            orderLoop = 0
            bNaturalY = False

            ' reset statistics lists
            lConnectivity = New List(Of BarrierAndDownstreamID)
            lDCIStatsList = New List(Of DCIStatisticsObject)

            ' reset the flow end elements tracker (barriers tracker)
            ' this variable prevents infinite looping but can be reset
            ' for each original flag.  
            pAllFlowEndBarriersGEN = New EnumNetEIDArray
            pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)

            fEID = pOriginaljuncFlagsList.Next()
            pNextOriginalJuncFlagGEN = New EnumNetEIDArray
            pNextOriginalJuncFlagGEN.Add(fEID)
            pNextOriginalJuncFlag = CType(pNextOriginalJuncFlagGEN, IEnumNetEID)

            For orderLoop = 0 To iOrderNum

                'MsgBox("Debug:9")
                If orderLoop = 0 Then

                    ' Change variable on first iteration for filtering next
                    pFlowEndJunctions = pNextOriginalJuncFlag
                ElseIf orderLoop > 0 Then
                    If pFlowEndJunctions.Count = 0 Or pFlowEndJunctions Is Nothing Then
                        Exit For
                    End If
                End If

                'MsgBox("Debug:10")
                pNoSourceFlowEnds = New EnumNetEIDArray

                ' ================ RUN TRACE ON ONE FLAG AT A TIME ======================
                ' check if user has hit 'close/cancel'
                If m_bCancel = True Then
                    backgroundworker1.CancelAsync()
                    backgroundworker1.Dispose()
                    Exit Sub
                End If
                If iProgress < 50 Then
                    iProgress = iProgress + 1
                End If
                backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & _
                                             Convert.ToString(pOriginaljuncFlagsList.Count) & _
                                             ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & _
                                         orderLoop.ToString & " of " & _
                                         iOrderNum.ToString & " (max)." & _
                                         ControlChars.NewLine & _
                                         "Direction: " & sDirection)

                'MsgBox("Debug:11")
                pFlowEndJunctions.Reset()


                For j = 0 To pFlowEndJunctions.Count - 1

                    ' this variable holds trace flow ends that are not sources
                    pNoSourceFlowEndsTemp = New EnumNetEIDArray

                    iEID = pFlowEndJunctions.Next()

                    ' ============ FILTER BARRIERS (eliminate barriers where flags will be) ===
                    ' For each of the original barriers
                    '   For each of the flow end junctions from the last trace
                    '     If the original barrier is NOT on a flow end junction 
                    '     Add it to a list to use in the upcoming trace.
                    ' DO NOT NEED TO DO THIS IN CASE THIS IS A 'NONBARR' ANALYSIS
                    ' AND ORDERLOOP = 0.  LOGIC THO?

                    p = 0
                    'reset / initialize OriginalBarriers list
                    pOriginalBarriersList.Reset()
                    ' clear the previous list
                    pNextTraceBarrierEIDGEN = New EnumNetEIDArray

                    For p = 0 To pOriginalBarriersList.Count - 1
                        bEID = pOriginalBarriersList.Next()
                        flagOverBarrier = False

                        ' If the EID of the end-of-flow junctions do not equal
                        ' that of the barriers set in the last trace then keep
                        ' that barrier for the next trace.
                        If bEID = iEID Then
                            flagOverBarrier = True
                        End If
                        'Next

                        ' For each barrier from the previous trace, add that barrier to a list
                        ' of barriers for the next trace if it doesn't overlap with one of the
                        ' flags for the next trace
                        If Not flagOverBarrier Then
                            ' add the barrier to a list of barriers to use in the next trace
                            pNextTraceBarrierEIDGEN.Add(bEID) 'VB.NET
                        End If
                    Next

                    'QI to get 'next' and 'count'
                    pNextTraceBarrierEIDs = CType(pNextTraceBarrierEIDGEN, IEnumNetEID)
                    ' ====================== END FILTER BARRIERS =============================

                    'MsgBox("Debug:12")
                    ' ========================== SET BARRIERS  ===============================
                    m = 0
                    pNextTraceBarrierEIDs.Reset()
                    pNetworkAnalysisExtBarriers.ClearBarriers()

                    For m = 0 To pNextTraceBarrierEIDs.Count - 1
                        bEID = pNextTraceBarrierEIDs.Next()
                        pNetElements.QueryIDs(bEID, _
                                              esriElementType.esriETJunction, _
                                              bFCID, _
                                              bFID, _
                                              bSubID)

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
                    ' ========================== END SET BARRIERS ===========================

                    ' ========================== SET FLAG ====================================

                    'MsgBox("Debug:13")
                    pNetworkAnalysisExtFlags.ClearFlags()
                    pNetElements.QueryIDs(iEID, esriElementType.esriETJunction, _
                                          iFCID, _
                                          iFID, _
                                          iSubID)

                    ' Display the flags as a JunctionFlagDisplay type
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


                    ' ====================== RUN TRACE IN DIRECTION OF ANALYSIS ====================
                    '                            TO GET FLOW END ELEMENTS

                    'MsgBox("Debug:14")
                    'prepare the network solver
                    pTraceFlowSolver = TraceFlowSolverSetup()
                    If pTraceFlowSolver Is Nothing Then
                        System.Windows.Forms.MessageBox.Show("Could not set up the network. Check that there is a network loaded.", _
                                                             "TraceFlowSolver setup error.")
                        Exit Sub
                    End If

                    eFlowElements = esriFlowElements.esriFEJunctionsAndEdges

                    'Return the features stopping the trace
                    If sDirection = "up" Then
                        pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMUpstream, _
                                                             eFlowElements, _
                                                             pFlowEndJunctionsPer, _
                                                             pFlowEndEdgesPer)
                    Else
                        pTraceFlowSolver.FindFlowEndElements(esriFlowMethod.esriFMDownstream, _
                                                             eFlowElements, _
                                                             pFlowEndJunctionsPer, _
                                                             pFlowEndEdgesPer)
                    End If
                    ' ============================ END RUN TRACE  ===============================

                    'MsgBox("Debug:15")
                    ' ================= GET BARRIER ID AND METRICS ===================
                    ' Gets the ID of the current flag?
                    pIDAndType = New IDandType(Nothing, Nothing)

                    pIDAndType = GetBarrierID(iFCID, _
                                              iFID, _
                                              lBarrierIDs)

                    sOutID = pIDAndType.BarrID

                    ' 2020 if the 'advanced connectivity' analysis (branching junctions respected)
                    '      then if element is a branch junction set permeability = 1
                    bBranchJunction = False
                    If bAdvConnectTab = True Then
                        pFilteredBranchJunctionsList.Reset()
                        For p = 0 To pFilteredBranchJunctionsList.Count - 1
                            iEID_p = pFilteredBranchJunctionsList.Next()
                            'MsgBox("Debug2020 The iEID_p: " + Str(iEID_p) + " And the iEID: " + Str(iEID))
                            If iEID = iEID_p Then
                                bBranchJunction = True
                                'MsgBox("Debug2020: Found junction will set perm to 1")
                                Exit For
                            End If
                        Next
                    End If

                    bSourceJunction = False
                    If bAdvConnectTab = True Then
                        pFilteredSourceJunctionsList.Reset()
                        For p = 0 To pFilteredSourceJunctionsList.Count - 1
                            iEID_p = pFilteredSourceJunctionsList.Next()
                            'MsgBox("Debug2020 The iEID_p: " + Str(iEID_p) + " And the iEID: " + Str(iEID))
                            If iEID = iEID_p Then
                                bSourceJunction = True
                                'MsgBox("Debug2020: Found source junction will set perm to 1")
                                'MsgBox("Debug2020 The source iEID_p: " + Str(iEID_p) + " And the iEID: " + Str(iEID))
                                Exit For
                            End If
                        Next
                    End If

                    ' track whehter it's a branch or source junction
                    If bBranchJunction = True Then
                        sType = "Branch Junction"
                    ElseIf bSourceJunction = True Then
                        sType = "Source Junction"
                    Else
                        sType = pIDAndType.BarrIDType
                    End If

                    If sType <> "Sink" And _
                        orderLoop = 0 And _
                        sFlagCheck = "nonbarr" Then

                        sType = "Flag - Node"

                        '2020 Note (Edge case fix): must add 'Flag - Node' type to 'barrier' list
                        ' during 'advanced analysis' with dd. otherwise the downstream path trace
                        ' from the next upstream nodes will pass the flag (if flag is not a 'sink').
                        ' the resulting connectivity table for R DCI will be wrong.
                        ' Fix here is thus to add the node to list of barriers temporarily to halt trace. 
                        ' This means that the 'path downstream' now does not proceed all the way to the 
                        ' network sink anymore. 
                        ' don't worry about modifying pOriginalBarriersList because the barriers are 
                        ' saved in pOriginalBarriersListSaved
                        pOriginalBarriersList.Reset()
                        pOriginalBarriersListGEN = New EnumNetEIDArray
                        For p = 0 To pOriginalBarriersList.Count - 1
                            iEID_p = pOriginalBarriersList.Next()
                            pOriginalBarriersListGEN.Add(iEID_p)
                        Next
                        pOriginalBarriersListGEN.Add(iEID)
                        pOriginalBarriersList = Nothing
                        pOriginalBarriersList = CType(pOriginalBarriersListGEN, IEnumNetEID)

                    ElseIf sType <> "Sink" And _
                        orderLoop = 0 And _
                        sFlagCheck = "barrier" Then

                        sType = "Flag - barrier"

                    End If

                    ' Get barrier Permeability 
                    'MsgBox("2020 message pFilteredBranchJunctionsList.Count : " + Str(pFilteredBranchJunctionsList.Count))
                    MsgBox("Debug2020: Flag / junction is a: " & sType)

                    ' Barrier permeability = 1 if branch junction
                    ' otherwise check the user-set attribute
                    If bBranchJunction = True Then
                        dBarrierPerm = 1
                    ElseIf bSourceJunction = True Then
                        dBarrierPerm = 1
                    Else
                        If bBarrierPerm = True Then
                            dBarrierPerm = GetBarrierPerm(iFCID, iFID, lBarrierIDs)
                        Else
                            dBarrierPerm = 0
                        End If
                    End If
                    

                    ' Get natural barrier
                    If bNaturalYN = True Then
                        sNaturalYN = GetNaturalYN(iFCID, _
                                                  iFID, _
                                                  lBarrierIDs)
                    Else
                        sNaturalYN = "F"
                    End If

                    If sNaturalYN = "T" Then
                        bNaturalY = True
                    End If

                    ' Barrier type - stores the 'culvert/dam/etc' type
                    ' of the layer... right now not in stored property set 
                    ' of the FIPEX extension - so temporary switch and function here
                    Dim bBarrierType As Boolean = True
                    '(- temp switch)
                    If bBranchJunction = True Then
                        sBarrierType = "Branch Junction"
                    ElseIf bSourceJunction = True Then
                        sBarrierType = "Source Junction"
                    Else
                        If bBarrierType = True Then
                            'Get BarrierType retrieves barrier type from user-set attribute
                            sBarrierType = GetBarrierType(iFCID, iFID, lBarrierIDs)
                        Else
                            sBarrierType = "Not Set"
                        End If

                    End If
                    If sBarrierType = "not found" Or sBarrierType = "Not Set" Then
                        sBarrierType = "Barrier"
                    End If
                   

                    ' populate tables for optimization analysis
                    If bGLPKTables = True Then

                        'MsgBox("Debug:16")
                        ' If this is a NEW FLAG / SINK, RESET THE OPTIONS LIST
                        If orderLoop = 0 Then
                            lGLPKOptionsList = New List(Of GLPKOptionsObject)  ' reset GLPK Connectivity table if this is a new flag
                            lGLPKConnectivity = New List(Of GLPKBarrierAndDownstreamEID)
                        End If

                        ' first add 'do nothing option', cost of zero
                        ' If orderloop is zero and the analysis type is 'nonbarr' then
                        ' assign a passability of 100% to the 'sink' (start-point)
                        If orderLoop = 0 And sFlagCheck = "nonbarr" Then
                            pGLPKOptionsObject = New GLPKOptionsObject(iEID, 1, 1, 0, "FLAG")
                            lGLPKOptionsList.Add(pGLPKOptionsObject)
                        Else
                            pGLPKOptionsObject = New GLPKOptionsObject(iEID, 1, dBarrierPerm, 0, sBarrierType)
                            lGLPKOptionsList.Add(pGLPKOptionsObject)
                        End If

                        lGLPKOptionsListTEMP = GetGLPKOptions(iFCID, _
                                                              iFID, _
                                                              lBarrierIDs, _
                                                              iEID, _
                                                              sBarrierType)

                        m = 0
                        For m = 0 To lGLPKOptionsListTEMP.Count - 1
                            lGLPKOptionsList.Add(lGLPKOptionsListTEMP(m))
                        Next

                    End If 'GLPK tables

                    ' Will save this sOutID and sType for later use, if this is orderloop zero (flag)
                    ' because will need to insert the DCI Metric at the end of this flag loop
                    ' else if it's a barrier in a greater order loop we're visiting, then keep
                    ' track of their id's for later use.  
                    If orderLoop = 0 Then
                        f_sOutID = sOutID
                        f_siOutEID = iEID
                        f_sType = sType
                    Else
                        pBarrierAndSinkEIDs = New BarrAndBarrEIDAndSinkEIDs(f_siOutEID, _
                                                                            iEID, _
                                                                            sOutID,
                                                                            sBarrierType)
                        lBarrierAndSinkEIDs.Add(pBarrierAndSinkEIDs)
                    End If

                    pMetricsObject = New MetricsObject(f_sOutID, _
                                                       f_siOutEID, _
                                                       sOutID, _
                                                       iEID, _
                                                       sType, _
                                                       "Permeability", _
                                                       dBarrierPerm)
                    lMetricsObject.Add(pMetricsObject)

                    ' ================ END GET IDS AND METRICS =============

                    ' Before the next stage, save the iFID variable because it may 
                    ' get reset in FILTER FLOW END ELEMENTS and it is used to label 
                    ' sinks in the form output
                    Dim sFID As String = iFID.ToString


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
                    'If this is orderloop zero
                    ' then it's okay to reset the 
                    pAllFlowEndBarriers = CType(pAllFlowEndBarriersGEN, IEnumNetEID)
                    pFlowEndJunctionsPer.Reset()
                    p = 0

                    'MsgBox("Debug:17")
                    For p = 0 To pFlowEndJunctionsPer.Count - 1

                        keepEID = False 'initialize
                        endEID = pFlowEndJunctionsPer.Next()
                        m = 0
                        pOriginalBarriersList.Reset()
                        For m = 0 To pOriginalBarriersList.Count - 1
                            If endEID = pOriginalBarriersList.Next() Then
                                keepEID = True ' set true if found
                                pAllFlowEndBarriers.Reset()
                                For k = 0 To pAllFlowEndBarriers.Count - 1
                                    If endEID = pAllFlowEndBarriers.Next() Then
                                        keepEID = False ' set false if already on master list
                                    End If
                                Next
                            End If
                        Next

                        If keepEID = True Then
                            pAllFlowEndBarriersGEN.Add(endEID) ' This variable does not get reset - used
                            ' to crosscheck in case of infinite loop problem
                            pNoSourceFlowEnds.Add(endEID) 'This variable gets reset each 
                            ' order loop

                            If bConnectTab = True Or bGLPKTables = True Then
                                ' this is for a running tally of connectivity:
                                ' first get the "barrierID" (set in Options by user)
                                '     using the getbarrierid sub which requires the 
                                '     fcid and fid from queryids method
                                pNetElements.QueryIDs(endEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)

                                Dim pIDAndType2 As New IDandType(Nothing, Nothing)
                                pIDAndType2 = GetBarrierID(iFCID, iFID, lBarrierIDs)
                                sLastUserSetID = sFlowEndID
                                sFlowEndID = pIDAndType2.BarrID

                                'Dim pIDAndType2 As New IDandType(Nothing, Nothing)

                                If bConnectTab = True Then
                                    ' TEMPORARY CODE UNTIL DCI MODEL IS UPDATED
                                    If orderLoop = 0 Then


                                        pBarrierAndDownstreamEID = New BarrierAndDownstreamID(endEID.ToString, "Sink", sFlowEndID, "Sink")
                                    Else
                                     
                                        pBarrierAndDownstreamEID = New BarrierAndDownstreamID(endEID.ToString, iEID.ToString, sFlowEndID, sOutID)

                                    End If

                                    lConnectivity.Add(pBarrierAndDownstreamEID)
                                End If

                                If bGLPKTables = True Then

                                    pGLPKBarrAndDownEID = New GLPKBarrierAndDownstreamEID(endEID, iEID)
                                    lGLPKConnectivity.Add(pGLPKBarrAndDownEID)
                                End If

                            End If

                            'NoSourceFlowEnds will have to
                            ' be changed to NoSourceJncFlowEnds in the future
                            ' and NoSourceEdgFlowEnds to differentiate
                            pNoSourceFlowEndsTemp.Add(endEID) ' This is a temporary variable.
                        End If
                    Next 'flowend element

                    ' if the following are needed then have to do upstream trace
                    If bUpHab = True Or _
                        bDCI = True Or _
                        bDownHab = True Or _
                        bGLPKTables = True Then

                        'MsgBox("Debug:18")
                        pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMUpstream, _
                                                          eFlowElements, _
                                                          pResultJunctions, _
                                                          pResultEdges)

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

                        If bUpHab = True _
                            Or bDCI = True _
                            Or bGLPKTables = True Then
                            ' Get results as selection
                            pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)

                            ' ============== INTERSECT FEATURES ================
                            ' check if user has hit 'close/cancel'
                            If m_bCancel = True Then
                                backgroundworker1.CancelAsync()
                                backgroundworker1.Dispose()
                                Exit Sub
                            End If
                            If iProgress < 50 Then
                                iProgress = iProgress + 1
                            End If
                            backgroundworker1.ReportProgress(iProgress, "Performing Network Traces (Intersecting Features) " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: " & sDirection)

                            'MsgBox("Debug:19")
                            Call IntersectFeatures()
                            'MsgBox("Debug:20")

                            ' ---- EXCLUDE FEATURES -----
                            pEnumLayer.Reset()
                            ' Look at the next layer in the list
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                        ExcludeFeatures(pFeatureLayer)
                                    End If
                                End If
                                pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Loop
                            ' ---- END EXCLUDE FEATURES -----

                            ' ============ SAVE FOR HIGHLIGHTING ==================
                            ' Get results to display as highlights at end of sub
                            pResultJunctions.Reset()
                            k = 0
                            For k = 0 To pResultJunctions.Count - 1
                                pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                            Next
                            pResultEdges.Reset()
                            For k = 0 To pResultEdges.Count - 1
                                pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                            Next

                            ' use results for DCI if needed
                            ' 2020 changed calculateDCIStatistics to use areas
                            ' as hab quantity or lengths. Also, differentiates lengths
                            ' from habitat (uses new settings from user)
                            If bDCI = True Then
                                Call calculateDCIStatistics(lDCIStatsList, _
                                                            lLLayersFieldsDCI, _
                                                            lPLayersFieldsDCI, _
                                                            iEID, _
                                                            dBarrierPerm, _
                                                            sNaturalYN, _
                                                            orderLoop)
                            End If

                            ' ############### FOR GLPK #####################
                            If bGLPKTables = True Then
                                ' If this is the first loop then empty the 
                                ' Statistics of results from other flags 
                                If orderLoop = 0 Then
                                    lGLPKStatsList = New List(Of GLPKStatisticsObject)
                                End If
                                Call calculateGLPKStatistics(lGLPKStatsList, _
                                                             lAllLayersFieldsGLPK, _
                                                             iEID, _
                                                             orderLoop)
                            End If ' ############### FOR GLPK #####################

                            ' use results for Upstream habitat if needed
                            If bUpHab = True Then
                                sHabType = "Immediate"
                                sHabDir = "upstream"
                                Call calculateStatistics_2(lHabStatsList, _
                                                           sOutID, _
                                                           iEID, _
                                                           sType, _
                                                           f_sOutID, _
                                                           f_siOutEID, _
                                                           sHabType, _
                                                           sHabDir)
                            End If

                        End If '  If bUpHab = True Or bDCI = True Or bGLPKTables = True

                        ' need to do a "connected" trace and subtract upstream trace results from it
                        If bDownHab = True Then
                            ' check if user has hit 'close/cancel'
                            If m_bCancel = True Then
                                backgroundworker1.CancelAsync()
                                backgroundworker1.Dispose()
                                Exit Sub
                            End If
                            If iProgress < 50 Then
                                iProgress = iProgress + 1
                            End If
                            backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Immediate Downstream ")

                            ' Store upstream elements from last trace to use to subtract from connected 
                            pSubtractUpJunctions = pResultJunctions
                            pSubtractUpEdges = pResultEdges

                            pMap.ClearSelection() ' clear selection

                            'MsgBox("Debug:21")
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

                            ' Get results as selection
                            'MsgBox("Debug:21")
                            pNetworkAnalysisExtResults.CreateSelection(pDownConnectedJunctions, pDownConnectedEdges)

                            ' Get results to display as highlights at end of sub
                            pDownConnectedJunctions.Reset()
                            k = 0
                            For k = 0 To pDownConnectedJunctions.Count - 1
                                pTotalResultsJunctionsGEN.Add(pDownConnectedJunctions.Next())
                            Next
                            pDownConnectedEdges.Reset()
                            For k = 0 To pDownConnectedEdges.Count - 1
                                pTotalResultsEdgesGEN.Add(pDownConnectedEdges.Next())
                            Next

                            ' ============== INTERSECT FEATURES ================
                            ' check if user has hit 'close/cancel'
                            If m_bCancel = True Then
                                backgroundworker1.CancelAsync()
                                backgroundworker1.Dispose()
                                Exit Sub
                            End If
                            If iProgress < 50 Then
                                iProgress = iProgress + 1
                            End If
                            backgroundworker1.ReportProgress(iProgress, "Performing Network Traces (Intersecting features) " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Immediate Downstream ")

                            'MsgBox("Debug:22")
                            Call IntersectFeatures()
                            'MsgBox("Debug:23")

                            ' ---- EXCLUDE FEATURES -----
                            pEnumLayer.Reset()
                            ' Look at the next layer in the list
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                        ExcludeFeatures(pFeatureLayer)
                                    End If
                                End If
                                pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Loop
                            ' ---- END EXCLUDE FEATURES -----

                            sHabType = "Immediate"
                            sHabDir = "downstream"

                            'MsgBox("Debug:24")
                            Call calculateStatistics_2(lHabStatsList, _
                                                       sOutID, _
                                                       iEID, _
                                                       sType, _
                                                       f_sOutID, _
                                                       f_siOutEID, _
                                                       sHabType, _
                                                       sHabDir)
                            'MsgBox("Debug:25")
                        End If
                    End If  ' bUpHab = True Or bDCI = True Or bDownHab = True 

                    ' If Downstream Path Habitat desired
                    ' 2020 - required if 'advanced connectivity' (distance decay)
                    If bPathDownHab = True Or bAdvConnectTab = True Then
                        ' check if user has hit 'close/cancel'
                        If m_bCancel = True Then
                            backgroundworker1.CancelAsync()
                            backgroundworker1.Dispose()
                            Exit Sub
                        End If
                        If iProgress < 50 Then
                            iProgress = iProgress + 1
                        End If
                        backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Path Downstream ")
                        'MsgBox("Debug:26")
                        pMap.ClearSelection() ' clear selection
                        ' perform downstream trace
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
                            pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                        Next
                        pResultEdges.Reset()
                        For k = 0 To pResultEdges.Count - 1
                            pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                        Next

                        ' Get results as selection
                        pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)

                        ' =================== INTERSECT FEATURES =============
                        ' check if user has hit 'close/cancel'
                        If m_bCancel = True Then
                            backgroundworker1.CancelAsync()
                            backgroundworker1.Dispose()
                            Exit Sub
                        End If
                        If iProgress < 50 Then
                            iProgress = iProgress + 1
                        End If
                        backgroundworker1.ReportProgress(iProgress, "Performing Network Traces (Intersecting features) " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Path Downstream ")
                        'MsgBox("Debug:27")
                        Call IntersectFeatures()
                        'MsgBox("Debug:28")
                        ' ---- EXCLUDE FEATURES -----
                        pEnumLayer.Reset()
                        ' Look at the next layer in the list
                        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                            If pFeatureLayer.Valid = True Then
                                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                    ExcludeFeatures(pFeatureLayer)
                                End If
                            End If
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                        Loop
                        ' ---- END EXCLUDE FEATURES -----

                        pActiveView.Refresh() ' refresh the view
                        sHabType = "Path" ' careful not to change - used later
                        sHabDir = "downstream" ' careful not to change - used later
                        'MsgBox("Debug:29")
                        Call calculateStatistics_2(lHabStatsList, _
                                                   sOutID, _
                                                   iEID, _
                                                   sType, _
                                                   f_sOutID, _
                                                   f_siOutEID, _
                                                   sHabType, _
                                                   sHabDir)
                        'MsgBox("Debug:30")
                    End If ' Downstream Path Habitat desired

                    ' If any total tables desired clear all barriers and run traces. 
                    If bTotalUpHab = True Or bTotalDownHab = True Then

                        ' check if user has hit 'close/cancel'
                        If m_bCancel = True Then
                            backgroundworker1.CancelAsync()
                            backgroundworker1.Dispose()
                            Exit Sub
                        End If
                        If iProgress < 50 Then
                            iProgress = iProgress + 1
                        End If
                        backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                         "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                     "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                     "Direction: Total Upstream (used to help get Total Downstream)")
                        'MsgBox("Debug:31")
                        pNetworkAnalysisExtBarriers.ClearBarriers()
                        pTraceFlowSolver = TraceFlowSolverSetup()

                        'prepare the network solver
                        If pTraceFlowSolver Is Nothing Then Exit Sub
                        eFlowElements = esriFlowElements.esriFEJunctionsAndEdges

                        pMap.ClearSelection() ' clear selection
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

                        ' ============ SAVE FOR HIGHLIGHTING ==================
                        ' Get results to display as highlights at end of sub
                        pResultJunctions.Reset()
                        k = 0
                        For k = 0 To pResultJunctions.Count - 1
                            pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                        Next
                        pResultEdges.Reset()
                        For k = 0 To pResultEdges.Count - 1
                            pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                        Next


                        If bTotalUpHab = True Then

                            ' Get results as selection
                            'MsgBox("Debug:31")
                            pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)

                            ' check if user has hit 'close/cancel'
                            If m_bCancel = True Then
                                backgroundworker1.CancelAsync()
                                backgroundworker1.Dispose()
                                Exit Sub
                            End If
                            If iProgress < 50 Then
                                iProgress = iProgress + 1
                            End If
                            backgroundworker1.ReportProgress(iProgress, "Performing Network Traces (intersecting features)" & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Total Upstream")

                            ' =================== INTERSECT FEATURES =============
                            Call IntersectFeatures()
                            'MsgBox("Debug:32")
                            ' ---- EXCLUDE FEATURES -----
                            pEnumLayer.Reset()
                            ' Look at the next layer in the list
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                        ExcludeFeatures(pFeatureLayer)
                                    End If
                                End If
                                pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Loop
                            ' ---- END EXCLUDE FEATURES -----

                            pActiveView.Refresh() ' refresh the view
                            sHabType = "Total"
                            sHabDir = "upstream"
                            'MsgBox("Debug:33")
                            Call calculateStatistics_2(lHabStatsList, _
                                                       sOutID, _
                                                       iEID, _
                                                       sType, _
                                                       f_sOutID, _
                                                       f_siOutEID, _
                                                       sHabType, _
                                                       sHabDir)
                            'MsgBox("Debug:34")
                        End If

                        If bTotalDownHab = True Then
                            ' check if user has hit 'close/cancel'
                            If m_bCancel = True Then
                                backgroundworker1.CancelAsync()
                                backgroundworker1.Dispose()
                                Exit Sub
                            End If
                            If iProgress < 50 Then
                                iProgress = iProgress + 1
                            End If
                            backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Total Downstream (subtracting upstream from 'connected' trace)")
                            'MsgBox("Debug:35")
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


                            ' ============ SAVE FOR HIGHLIGHTING ==================
                            ' Get results to display as highlights at end of sub
                            pResultJunctions.Reset()
                            k = 0
                            For k = 0 To pResultJunctions.Count - 1
                                pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                            Next
                            pResultEdges.Reset()
                            For k = 0 To pResultEdges.Count - 1
                                pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                            Next

                            ' ========= SUBTRACT UPSTREAM FROM ALL CONNECTED =======
                            pConnectedJunctions = pResultJunctions
                            pConnectedEdges = pResultEdges

                            ' Subtract the upstream edges and junctions from the connected edges and junctions list
                            pDownConnectedJunctions = DownStreamConnected(pSubtractUpJunctions, _
                                                                          pConnectedJunctions)
                            pDownConnectedEdges = DownStreamConnected(pSubtractUpEdges, _
                                                                      pConnectedEdges)

                            ' Get results as selection
                            pNetworkAnalysisExtResults.CreateSelection(pDownConnectedJunctions, pDownConnectedEdges)

                            ' =================== INTERSECT FEATURES =============
                            'MsgBox("Debug:36")
                            Call IntersectFeatures()
                            'MsgBox("Debug:37")
                            ' ---- EXCLUDE FEATURES -----
                            pEnumLayer.Reset()
                            ' Look at the next layer in the list
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                                If pFeatureLayer.Valid = True Then
                                    If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                    pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                        ExcludeFeatures(pFeatureLayer)
                                    End If
                                End If
                                pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                            Loop
                            ' ---- END EXCLUDE FEATURES -----
                            pActiveView.Refresh() ' refresh the view

                            sHabType = "Total"
                            sHabDir = "downstream"
                            'MsgBox("Debug:38")

                            ' returns a habitat list
                            Call calculateStatistics_2(lHabStatsList, _
                                                       sOutID, _
                                                       iEID, _
                                                       sType, _
                                                       f_sOutID, _
                                                       f_siOutEID, _
                                                       sHabType, _
                                                       sHabDir)
                            'MsgBox("Debug:39")
                        End If
                    End If ' bTotalUpHab = True Or bTotalDownHab = True 

                    If bTotalPathDownHab = True Then
                        ' check if user has hit 'close/cancel'
                        If m_bCancel = True Then
                            backgroundworker1.CancelAsync()
                            backgroundworker1.Dispose()
                            Exit Sub
                        End If
                        If iProgress < 50 Then
                            iProgress = iProgress + 1
                        End If
                        backgroundworker1.ReportProgress(iProgress, "Performing Network Traces " & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count) & ControlChars.NewLine & _
                                         "Current Barrier 'Order': " & orderLoop.ToString & " of " & iOrderNum.ToString & " (max)." & ControlChars.NewLine & _
                                         "Direction: Path Downstream")
                        'MsgBox("Debug:40")
                        pMap.ClearSelection() ' clear selection

                        pNetworkAnalysisExtBarriers.ClearBarriers()
                        pTraceFlowSolver = TraceFlowSolverSetup()

                        ' perform UPSTREAM trace
                        pTraceFlowSolver.FindFlowElements(esriFlowMethod.esriFMDownstream, _
                                                          eFlowElements, _
                                                          pResultJunctions, _
                                                          pResultEdges)

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
                            pTotalResultsJunctionsGEN.Add(pResultJunctions.Next())
                        Next
                        pResultEdges.Reset()
                        For k = 0 To pResultEdges.Count - 1
                            pTotalResultsEdgesGEN.Add(pResultEdges.Next())
                        Next

                        ' Get results as selection
                        pNetworkAnalysisExtResults.CreateSelection(pResultJunctions, pResultEdges)
                        ' =================== INTERSECT FEATURES =============
                        'MsgBox("Debug:41")
                        Call IntersectFeatures()
                        'MsgBox("Debug:42")
                        ' ---- EXCLUDE FEATURES -----
                        pEnumLayer.Reset()
                        ' Look at the next layer in the list
                        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
                            If pFeatureLayer.Valid = True Then
                                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Or _
                                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref
                                    ExcludeFeatures(pFeatureLayer)
                                End If
                            End If
                            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
                        Loop
                        ' ---- END EXCLUDE FEATURES -----

                        pActiveView.Refresh() ' refresh the view
                        sHabType = "Total Path"
                        sHabDir = "downstream"
                        'MsgBox("Debug:43")
                        Call calculateStatistics_2(lHabStatsList, _
                                                   sOutID, _
                                                   iEID, _
                                                   sType, _
                                                   f_sOutID, _
                                                   f_siOutEID, _
                                                   sHabType, _
                                                   sHabDir)
                        'MsgBox("Debug:44")
                    End If ' bTotalPathDownHab = true

                    'MsgBox("Debug2020: The iEID of this Barrier: " + CStr(iEID) + " and the sOutID of barrier: " + CStr(sOutID))


                    'End If ' calculate total impacted
                    pNetworkAnalysisExtFlags.ClearFlags()   ' clear the flags

                    ' clear selection
                    pMap.ClearSelection()

                    ' refresh the view
                    pActiveView.Refresh()

                Next 'flag
                ' ================= END RUN TRACE ON ONE FLAG AT A TIME =================

                ' change the flowEndJunctions to an array
                ' that no longer has sources, or non-barriers.
                pFlowEndJunctions = New EnumNetEIDArray
                pFlowEndJunctions = CType(pNoSourceFlowEnds, IEnumNetEID)

                ' Store the first flow end elements
                ' 2012- THESE VARS MAY NOT BE USED ANYMORE
                If orderLoop = 0 And sFlagCheck = "nonbarr" Then
                    pFirstJunctionBarriers = pFlowEndJunctions
                ElseIf orderLoop = 0 And sFlagCheck = "barriers" Then
                    pFirstJunctionBarriers = pOriginaljuncFlagsList
                End If

                pFirstEdgeBarriers = pFlowEndEdges

                ' Clear the barriers.
                pNetworkAnalysisExtBarriers.ClearBarriers()

                'End If ' orderloop = x
            Next ' orderloop

            iProgress = 50

            ' ======= CREATE and WRITE TO TABLES ========
            ' Use the stats object list to write 
            ' rows to the table. 
            If bDBF = True Then
                'MsgBox("Debug:45")
                ' check if user has hit 'close/cancel'
                If m_bCancel = True Then
                    backgroundworker1.CancelAsync()
                    backgroundworker1.Dispose()
                    Exit Sub
                End If
                If iProgress < 70 Then
                    iProgress = iProgress + 1
                End If
                backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count))

                ' ============== DCI, GLPK, Habitat and Metric Table Name ========
                ' First Generate a Random 5 digit code for the table name to keep track of multiple analyses
                Dim s As String = "abcdefghijklmnopqrstuvwxyz0123456789"
                Dim r As New Random
                Dim sb As New System.Text.StringBuilder
                For x As Integer = 1 To 5
                    Dim idx As Integer = r.Next(0, 35)
                    sb.Append(s.Substring(idx, 1))
                Next
                Dim sAnalysisCode As String

                sAnalysisCode = sb.ToString

                ' creates a table to be sent to R for DCI, but only needed
                ' if 'advanced' table isn't already created - GO, 2020
                If bDCI = True And bAdvConnectTab = False Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, _
                                                     "Creating Output Tables" & _
                                                     ControlChars.NewLine & _
                                                     "User Flag " & (i + 1).ToString & _
                                                     " of " & Convert.ToString(pOriginaljuncFlagsList.Count & _
                                                                               ControlChars.NewLine & _
                                                                               "Table: DCI Analysis"))
                    sDCITableName = TableName("DCI_" + sAnalysisCode, _
                                              pFWorkspace, _
                                              sPrefix)
                    'MsgBox("Debug:46")
                    PrepDCIOutTable(sDCITableName, pFWorkspace)

                End If

                If bConnectTab = True And bAdvConnectTab = False Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                                "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                            "Table: Connectivity"))
                    'MsgBox("Debug:47")
                    sConnectTabName = TableName("connectivity_" + sAnalysisCode, _
                                                pFWorkspace, _
                                                sPrefix)

                    If sConnectTabName = "Cancel" Then
                        MsgBox("Debug2020 - issue encountered naming output connectivity table")
                        Exit Sub
                    End If
                    PrepConnectivityTables(pFWorkspace, sConnectTabName)
                End If

                If bAdvConnectTab = True Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                                "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                            "Table: Connectivity"))
                    'MsgBox("Debug:47")
                    sAdvConnectTabName = TableName("ADV_connectivity_" + sAnalysisCode, _
                                                pFWorkspace, _
                                                sPrefix)

                    If sAdvConnectTabName = "Cancel" Then
                        MsgBox("Debug2020 - issue encountered naming output advanced connect table")
                        Exit Sub
                    End If
                    'PrepAdvanced2020DCIOutTable(sAdvConnectTabName, pFWorkspace)
                End If

                If bGLPKTables = True Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                               "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                           "Table: Habitat (For GLPK Optimization)"))

                    'MsgBox("Debug:48")
                    sGLPKHabitatTableName = TableName("GLPKHabitat_" + sAnalysisCode, _
                                                      pFWorkspace, _
                                                      sPrefix)
                    PrepGLPKHabTable(sGLPKHabitatTableName, pFWorkspace)

                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                               "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                           "Table: Options (For GLPK Optimization)"))

                    'MsgBox("Debug:49")
                    sGLPKOptionsTableName = TableName("GLPKOptions_" + sAnalysisCode, _
                                                      pFWorkspace, _
                                                      sPrefix)
                    PrepGLPKOptionsTable(sGLPKOptionsTableName, pFWorkspace)

                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                               "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                           "Table: Connectivity (For GLPK Optimization)"))
                    sGLPKConnectTabName = TableName("GLPKconnectivity_" + sAnalysisCode, _
                                                    pFWorkspace, _
                                                    sPrefix)
                    If sGLPKConnectTabName = "Cancel" Then
                        Exit Sub
                    End If
                    'MsgBox("Debug:50")
                    PrepGLPKConnectivityTable(pFWorkspace, sGLPKConnectTabName)
                End If

                ' If this is the first flag visited - don't need another table 
                ' for habitat and metrics
                ' for every flag visited - can append if it's the second flag or higher. 
                If i = 0 Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                              "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                          "Table: Habitat"))
                    'MsgBox("Debug:51")
                    sTableType = "Habitat_"
                    sHabTableName = TableName(sTableType + sAnalysisCode, _
                                              pFWorkspace, _
                                              sPrefix)
                    Call PrepHabitatTable(sHabTableName, pFWorkspace)


                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Creating Output Tables" & ControlChars.NewLine & _
                                                              "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                          "Table: Metrics"))
                    sTableType = "Metrics_"
                    sMetricTableName = TableName(sTableType + sAnalysisCode, _
                                                 pFWorkspace, _
                                                 sPrefix)
                    Call PrepMetricTable(sMetricTableName, _
                                         pFWorkspace)

                    ' ========== 2020 To Do =========
                    ' insert 'sMetricTableName' Advanced?

                    'MsgBox("Debug:52")

                End If
                ' ============== End New DCI, Hab and Metric Table Name ==========

                '2020 if advanced connectivity then no regular connectivity table
                If bConnectTab = True And bAdvConnectTab = False Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Writing To Output Tables" & ControlChars.NewLine & _
                                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                         "Table: Connectivity"))
                    'MsgBox("Debug:53")
                    pTable = pFWorkspace.OpenTable(sConnectTabName)
                    j = 0
                    For j = 0 To lConnectivity.Count - 1
                        pRowBuffer = pTable.CreateRowBuffer
                        pRowBuffer.Value(1) = lConnectivity(j).BarrID
                        pRowBuffer.Value(2) = lConnectivity(j).DownstreamBarrierID
                        pRowBuffer.Value(3) = lConnectivity(j).BarrLabel
                        pRowBuffer.Value(4) = lConnectivity(j).DownstreamBarrLabel
                        pCursor = pTable.Insert(True)
                        pCursor.InsertRow(pRowBuffer)
                        pCursor.Flush()
                    Next
                End If

                ' 2020 - 'allow for second connectivity containing downstream length and habitat 
                '      - match the path length stats for each node 

                If bAdvConnectTab = True And bDCI = True Then
                    ' to do: change lDCIStatsList to include Areas as well....
                    ' to do: above will fix issue when there are 'classes' used
                    '        and eliminate the need to use the lHabStatsObject

                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Writing To Output Tables" & ControlChars.NewLine & _
                                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                         "Table: Connectivity"))
                    'MsgBox("Debug:53")
                    'pTable = pFWorkspace.OpenTable(sAdvConnectTabName)
                    j = 0
                    For j = 0 To lConnectivity.Count - 1

                        'pRowBuffer = pTable.CreateRowBuffer

                        pAdvDCIDataObj = New Adv_DCI_Data_Object(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, _
                                                                 Nothing, Nothing, Nothing, Nothing, Nothing)

                        ' objectID
                        ' BarrierID
                        ' BarirerUserLabel
                        ' Habt Quan
                        ' Habit Quan Units
                        ' BarrPerm
                        ' BarrnaturalYN
                        ' DownstreamNeighbEID
                        ' DownstreamNeighbUserLabel
                        ' DownstreamNeighbDistance
                        ' DownstreamNeighbUnits

                        'pRowBuffer.Value(1) = lConnectivity(j).BarrID
                        'pRowBuffer.Value(2) = lConnectivity(j).BarrLabel

                        pAdvDCIDataObj.NodeEID = lConnectivity(j).BarrID
                        pAdvDCIDataObj.NodeLabel = lConnectivity(j).BarrLabel

                        For k = 0 To lDCIStatsList.Count - 1
                            If lDCIStatsList(k).Barrier = lConnectivity(j).BarrID Then
                                ' 2020 to-do - fix from here to pCursor
                                '            - DCIStatsList should have same stats as lhabstatslist need to add 
                                '              room for all params required for DCI
                                '              downstream habitat length, hab area, and distance to next barrier 
                                '              (distance often = length, but not always) 
                                'pRowBuffer.Value(5) = lDCIStatsList(k).BarrierPerm
                                'pRowBuffer.Value(6) = lDCIStatsList(k).BarrierYN

                                pAdvDCIDataObj.BarrierPerm = lDCIStatsList(k).BarrierPerm
                                pAdvDCIDataObj.NaturalTF = lDCIStatsList(k).BarrierYN

                            End If
                        Next

                        'pRowBuffer.Value(7) = lConnectivity(j).DownstreamBarrierID
                        'pRowBuffer.Value(8) = lConnectivity(j).DownstreamBarrLabel

                        pAdvDCIDataObj.DownstreamEID = lConnectivity(j).DownstreamBarrierID
                        pAdvDCIDataObj.DownstreamNodeLabel = lConnectivity(j).DownstreamBarrLabel

                        For k = 0 To lHabStatsList.Count - 1
                            If lHabStatsList(k).bEID = lConnectivity(j).BarrID Then
                                If lHabStatsList(k).TotalImmedPath = "Path" And lHabStatsList(k).Direction = "downstream" Then

                                    ' 2020 to do: currently this restricts habitat to length  - should fix
                                    '             and add user-options 

                                    'pRowBuffer.Value(3) = Math.Round(lHabStatsList(k).Quantity, 2)
                                    'pRowBuffer.Value(4) = lHabStatsList(k).Unit
                                    'pRowBuffer.Value(9) = Math.Round(lHabStatsList(k).Quantity, 2)
                                    'pRowBuffer.Value(10) = lHabStatsList(k).Unit

                                    pAdvDCIDataObj.NodeType = lHabStatsList(k).bType
                                    pAdvDCIDataObj.HabQuantity = Math.Round(lHabStatsList(k).Quantity, 2)
                                    pAdvDCIDataObj.HabQuanUnits = lHabStatsList(k).Unit
                                    pAdvDCIDataObj.DownstreamNeighDistance = Math.Round(lHabStatsList(k).Quantity, 2)
                                    pAdvDCIDataObj.DistanceUnits = lHabStatsList(k).Unit

                                    ' 2020 insert a check for path down distance / length




                                End If
                            End If
                        Next

                        'pCursor = pTable.Insert(True)
                        'pCursor.InsertRow(pRowBuffer)
                        'pCursor.Flush()

                        '2020
                        lAdv_DCI_Data_Object.Add(pAdvDCIDataObj)

                    Next

                End If

                ' DO NOT DELETE!
                'If bGLPKTables = True Then
                '    pTabl2e = pFWorkspace.OpenTable(sConnectTabName)
                '    j = 0
                '    For j = 0 To lConnectivity.Count - 1
                '        pRowBuffer = pTable.CreateRowBuffer
                '        pRowBuffer.Value(2) = lConnectivity(j).DownstreamBarrierID
                '        pRowBuffer.Value(1) = lConnectivity(j).BarrID

                '        pCursor = pTable.Insert(True)
                '        pCursor.InsertRow(pRowBuffer)
                '    Next
                'End If

                If bDCI = True Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Writing To Output Tables" & ControlChars.NewLine & _
                                                             "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count & ControlChars.NewLine & _
                                                                                                                         "Table: DCI Statistics (for output to 'R' Stats Software)"))
                    ' used to check there isn't an issue with table during regular DCI
                    ' (if zero rows then don't call R)
                    Dim iDCIRowCountTEMP As Integer = 1
                    If bAdvConnectTab = False Or bDistanceLim = False Then
                        ' ==== BEGIN WRITE TO DCI TABLE =====
                        ' convert to generic
                        'Dim lConnectTabName As New List(Of String)
                        'lConnectTabName.Add(sConnectTabName)
                        pTable = pFWorkspace.OpenTable(sDCITableName)
                        j = 0
                        bNoPerm = False ' NoPerm false means there aren't any barriers encountered in 
                        ' this list yet that have less than one permeability
                        For j = 0 To lDCIStatsList.Count - 1
                            pRowBuffer = pTable.CreateRowBuffer
                            pRowBuffer.Value(1) = lDCIStatsList(j).Barrier
                            pRowBuffer.Value(2) = Math.Round(lDCIStatsList(j).HabQuantity, 2)
                            pRowBuffer.Value(3) = lDCIStatsList(j).BarrierPerm
                            If lDCIStatsList(j).BarrierPerm < 1 Then
                                bNoPerm = True
                            End If
                            pRowBuffer.Value(4) = lDCIStatsList(j).BarrierYN

                            pCursor = pTable.Insert(True)
                            pCursor.InsertRow(pRowBuffer)
                            pCursor.Flush()
                        Next
                        ' ===== END WRITE TO DCI TABLE =======
                        iDCIRowCountTEMP = pTable.RowCount(Nothing)

                    End If
                    
                    '
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Calculating DCI" & ControlChars.NewLine & _
                                                            "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count))

                    '====== BEGIN DCI CALCULATION ======
                    ' TEMPORARY CHECK - if DCI Table has no rows then the 
                    ' DCI model will fail.  If there are no barriers then 
                    ' in UpdateResultsDCI function the DCIp and DCId will be
                    ' set to 100.   ' check if user has hit 'close/cancel'

                    ' run regular analysis if 'advanced connectivity' isn't checked
                    ' DCIRowCountTEmp checks that there is more than one segments in DCI table 
                    ' and if not skips R call
                 
                    If bAdvConnectTab = False Or bDistanceLim = False Then
                        If iDCIRowCountTEMP = 0 Or bNoPerm = False Then
                            UpdateResultsDCI(iDCIRowCountTEMP, dDCIp, dDCId, bNaturalY, "out.txt")
                        ElseIf iDCIRowCountTEMP > 0 Then
                            DCIShellCall(sDCITableName, sConnectTabName, pFWorkspace)
                            UpdateResultsDCI(iDCIRowCountTEMP, dDCIp, dDCId, bNaturalY, "out.txt")
                        Else
                            MsgBox("Debug2020: Error 306 in RunAnalysis. bAdvConnectTab must = True And bDistanceLim must = True to proceed with DCI. Exiting")
                            Exit Sub
                        End If
                    ElseIf bAdvConnectTab = True And bDistanceLim = True Then
                        If lAdv_DCI_Data_Object.Count > 1 Then
                            DCI_ADV2020_ShellCall(pFWorkspace, _
                                      lAdv_DCI_Data_Object, _
                                      bDCISectional, _
                                      bDistanceLim, _
                                      dMaxDist, _
                                      bDistanceDecay, sDDFunction)
                            UpdateResultsDCI(lAdv_DCI_Data_Object.Count, dDCIp, dDCId, bNaturalY, "out_dd.txt")
                        Else
                            MsgBox("Debug2020: Error 204 in RunAnalysis. There must be more than 1 row in the advanced summary table object. Exiting")
                            Exit Sub
                        End If
                    Else
                        MsgBox("Debug2020: Error 210 in RunAnalysis. bAdvConnectTab must = True And bDistanceLim must = True to proceed with DCI. Exiting")
                        Exit Sub

                    End If


                ' Insert a row into the Metrics Object List
                ' using pair of ID and Type for flag saved earlier
                pMetricsObject = New MetricsObject(f_sOutID, f_siOutEID, f_sOutID, f_siOutEID, f_sType, "DCIp", dDCIp)
                lMetricsObject.Add(pMetricsObject)
                pMetricsObject = New MetricsObject(f_sOutID, f_siOutEID, f_sOutID, f_siOutEID, f_sType, "DCId", dDCId)
                lMetricsObject.Add(pMetricsObject)

                ' Update the metrics object with the sectional DCI
                If bDCISectional = True Then
                    ' check if user has hit 'close/cancel'
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    If iProgress < 70 Then
                        iProgress = iProgress + 1
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Calculating DCI Sectional" & ControlChars.NewLine & _
                                                            "User Flag " & (i + 1).ToString & " of " & Convert.ToString(pOriginaljuncFlagsList.Count))
                    'MsgBox("Debug:55")
                    ' read output table if it exists
                    ' and add sectional DCIs to the metrics object
                    'UpdateResultsDCISectional()

                    ' Read the DCI Model Directory from the document properties
                    Dim sDCIModelDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sDCIModelDir"))
                    'Dim bDCISectional As Boolean = Convert.ToBoolean(m_FiPEX__1.pPropset.GetProperty("dcisectionalyn"))

                    ' Read the file and return two values - line one and line two from output
                    Dim ErrInfo As String

                    Dim objReader As StreamReader
                    Dim sLine As String
                    Dim iPosit As Integer
                    Dim dDCIs As Double
                    Dim iBarrierEID As Integer

                    Dim iLoopCount As Integer = 0

                        '2020
                        Dim DCIsFileName As String = "DCI_all_sections.csv"
                        Dim iDCIRowCount As Integer = iDCIRowCountTEMP

                        If bAdvConnectTab = True And bDistanceLim = True Then
                            DCIsFileName = "DCI_all_sections_dd.csv"
                            iDCIRowCount = lAdv_DCI_Data_Object.Count
                        End If


                    Try
                            If iDCIRowCount > 1 Then
                                objReader = New StreamReader(sDCIModelDir + "/" + DCIsFileName)
                                Do Until objReader.EndOfStream = True
                                    iLoopCount = iLoopCount + 1
                                    sLine = objReader.ReadLine()
                                    ' If the first line says ERROR then write this, otherwise write nothing
                                    ' because the first line of a successful DCI will say "Value" - skip.  
                                    ' Using FINDME to get to the position at top of results box.  
                                    If iLoopCount = 1 Then
                                        If sLine = "ERROR" Then
                                            Exit Do
                                        End If
                                    Else
                                        ' separate barrier from DCIs value
                                        Dim sLineArray(1) As String
                                        sLineArray = sLine.Split(",")

                                        ' separate barrier label 
                                        sLineArray(0) = sLineArray(0).Trim("""")
                                        If sLineArray(0) <> "sink" Then
                                            sLineArray(0) = sLineArray(0).TrimEnd("s")
                                            sLineArray(0) = sLineArray(0).TrimEnd("_")
                                            iBarrierEID = Convert.ToInt64(sLineArray(0))
                                        Else
                                            iBarrierEID = f_siOutEID
                                        End If
                                        dDCIs = Convert.ToDouble(sLineArray(1))

                                        'retrieve the barrier label, too
                                        pNetElements.QueryIDs(iBarrierEID, esriElementType.esriETJunction, iFCID, iFID, iSubID)
                                        pIDAndType = New IDandType(Nothing, Nothing)
                                        pIDAndType = GetBarrierID(iFCID, iFID, lBarrierIDs)
                                        sOutID = pIDAndType.BarrID

                                        ' ##############################################
                                        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_
                                        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_

                                        ' do not read the same DCI sectional for branch junctions
                                        ' they are the same as for the next downstream barrier
                                        '
                                        ' if EID = branch junction EID
                                        '  then flag and do not do next step
                                        bBranchJunction = False
                                        'If bAdvConnectTab = True Then
                                        '    'MsgBox("Debug2020: filtered branch juncs: " + Str(pFilteredBranchJunctionsList.Count))

                                        '    pFilteredBranchJunctionsList.Reset()
                                        '    For p = 0 To pFilteredBranchJunctionsList.Count - 1
                                        '        iEID_p = pFilteredBranchJunctionsList.Next()
                                        '        'MsgBox("Debug2020 The iEID_p: " + Str(iEID_p) + " And the iBarrierEID: " + Str(iBarrierEID))
                                        '        If iBarrierEID = iEID_p Then
                                        '            bBranchJunction = True
                                        '            'MsgBox("Debug2020: MAtch found")
                                        '            Exit For
                                        '        End If
                                        '    Next
                                        'End If

                                        bSourceJunction = False
                                        If bAdvConnectTab = True Then
                                            'MsgBox("Debug2020: filtered branch juncs: " + Str(pFilteredBranchJunctionsList.Count))

                                            pFilteredSourceJunctionsList.Reset()
                                            For p = 0 To pFilteredSourceJunctionsList.Count - 1
                                                iEID_p = pFilteredSourceJunctionsList.Next()
                                                'MsgBox("Debug2020 The iEID_p: " + Str(iEID_p) + " And the iBarrierEID: " + Str(iBarrierEID))
                                                If iBarrierEID = iEID_p Then
                                                    bSourceJunction = True
                                                    'MsgBox("Debug2020: MAtch found")
                                                    Exit For
                                                End If
                                            Next
                                        End If

                                        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_
                                        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_
                                        ' ###############################################

                                        If bBranchJunction = False And bSourceJunction = False Then
                                            ' Insert new metric into metrics object list
                                            pMetricsObject = New MetricsObject(f_sOutID, _
                                                                               f_siOutEID, _
                                                                               sOutID, _
                                                                               iBarrierEID, _
                                                                               f_sType, _
                                                                               "DCI Sectional", _
                                                                               Math.Round(dDCIs, 2))
                                            lMetricsObject.Add(pMetricsObject)
                                        End If
                                    End If
                                Loop
                                objReader.Close()
                            Else ' if there are no barriers encountered
                                dDCIs = 100
                            End If
                    Catch Ex As Exception
                            ErrInfo = Ex.Message
                            objReader.Close()
                        MsgBox("Error reading Sectional DCI output file. " + ErrInfo)
                    End Try

                End If ' output to DCI sectional = True
            End If 'output to DCI = True
            ' ======= END DCI CALCULATION =======

            ' Insert a randomization check here.  
            Dim bCostRandomization As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bCostRandomization"))
            Dim bPermRandomization As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bPermRandomization"))
            Dim iCostRandomCount As Integer '= Convert.ToInt16(m_FiPEx__1.pPropset.GetProperty("iCostRandomCount"))
            Dim iPermRandomCount As Integer '= Convert.ToInt16(m_FiPEx__1.pPropset.GetProperty("iPermRandomCount"))
            Dim bExcludeDams As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bExcludeDams"))
            ' THIS NEEDS TO BE ADDED TO DOC STREAM IN ALL READ/WRITE LOCATIONShttp://www.boingboing.net/features/cassini/cassini6.jpg
            ' MAY CORRUPT EXISTING DOCUMENTS
            Dim bUndirected As Boolean = False '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bUndirected"))
            Dim bDirected As Boolean = False '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDirected"))

            bPermRandomization = False
            bCostRandomization = False
            iCostRandomCount = 40

            Dim dRandomCost As Double
            Dim bUseRandomPerm As Boolean
            Dim bCreateRandomPerm As Boolean

            If bGLPKTables = True Then
                ' ===== BEGIN GLPK TABLES WRITE =====
                'MsgBox("Debug:56")
                pTable = pFWorkspace.OpenTable(sGLPKHabitatTableName)
                j = 0
                'bNoPe2wrm = False ' NoPerm false means there aren't any barriers encountered in 
                '' this list yet that have less than one permeability
                For j = 0 To lGLPKStatsList.Count - 1
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(0) = lGLPKStatsList(j).Barrier
                    'pRowBuffer.Value(1) = 1
                    Try
                        pRowBuffer.Value(1) = Math.Round(lGLPKStatsList(j).Quantity, 1)
                        'pRowBuffer.Value(1) = Convert.ToInt32(lGLPKStatsList(j).Quantity)

                    Catch ex As Exception
                        MsgBox("Error attempting to insert quantity into GLPKHabitatTable. Code4455. Continuing... " + ex.Message)
                        pRowBuffer.Value(1) = lGLPKStatsList(j).Quantity
                    End Try
                    pCursor = pTable.Insert(True)
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                Next


                pTable = pFWorkspace.OpenTable(sGLPKConnectTabName)
                j = 0
                For j = 0 To lGLPKConnectivity.Count - 1
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(0) = lGLPKConnectivity(j).DownstreamBarrierEID
                    pRowBuffer.Value(1) = lGLPKConnectivity(j).BarrEID
                    pRowBuffer.Value(2) = 1

                    pCursor = pTable.Insert(True)
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                Next

                pTable = pFWorkspace.OpenTable(sGLPKOptionsTableName)

                ' if there's no randomization then update options normally
                ' and run the GLPKShellCall once
                If bCostRandomization = False And bPermRandomization = False Then

                    j = 0
                    For j = 0 To lGLPKOptionsList.Count - 1

                        'If Not lGLPKUniqueEIDs.Contains(lGLPKOptionsList(j).BarrierEID) Then
                        '    lGLPKUniqueEIDs.Add(lGLPKOptionsList(j).BarrierEID)
                        'End If

                        ' TEMP - to eliminate all options except 'do nothing' - should make UNDIR work
                        'If lGLPKOptionsList(j).OptionNum < 2 Then

                        If lGLPKOptionsList(j).BarrierPerm <> 9999 And lGLPKOptionsList(j).OptionCost <> 9999 Then
                            pRowBuffer = pTable.CreateRowBuffer
                            pRowBuffer.Value(0) = lGLPKOptionsList(j).BarrierEID
                            pRowBuffer.Value(1) = lGLPKOptionsList(j).OptionNum
                            ' keep track of max options for use in input to GLPK Model
                            If iMaxOptionNum < lGLPKOptionsList(j).OptionNum Then
                                iMaxOptionNum = lGLPKOptionsList(j).OptionNum
                            End If
                            pRowBuffer.Value(2) = Math.Round(lGLPKOptionsList(j).BarrierPerm, 2)
                            'pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost, 2)
                            pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost)
                            pCursor = pTable.Insert(True)
                            pCursor.InsertRow(pRowBuffer)
                            pCursor.Flush()
                        End If
                        'End If
                    Next ' j

                    'j = 0
                    'For j = 0 To lGLPKUniqueEIDs.Count - 1
                    '    pRowBuffer = pTable.CreateRowBuffer
                    '    pRowBuffer.Value(0) = lGLPKOptionsList(j).BarrierEID
                    '    pRowBuffer.Value(1) = 1
                    '    pRowBuffer.Value(2) = Math.Round(lGLPKOptionsList(j).BarrierPerm, 2)
                    '    pRowBuffer.Value(3) = 0

                    '    pCursor = pTable.Insert(True)
                    '    pCursor.InsertRow(pRowBuffer)
                    'Next

                    ' ===== END GLPK TABLES WRITE =======

                    ' ===== BEGIN GLPK CALCULATION =====
                    ' check if user has hit 'close/cancel'
                    'MsgBox("Debug:57")

                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If

                    If iProgress < 70 Then
                        iProgress = iProgress + 1
                    End If

                    backgroundworker1.ReportProgress(iProgress, "Performing Optimization Analysis" & ControlChars.NewLine & _
                                                 "User Flag " & (i + 1).ToString & " of " & (pOriginaljuncFlagsList.Count).ToString)

                    If bDirected = True Then
                        GLPKShellCall(sGLPKHabitatTableName, _
                                      sGLPKOptionsTableName, _
                                      sGLPKConnectTabName, _
                                      iMaxOptionNum, _
                                      f_siOutEID, _
                                      pFWorkspace, _
                                      iProgress, _
                                      sAnalysisCode, _
                                      "DIR", _
                                      False)
                        ' run once to get 'initial state' of network
                        ' measured in permeability weighted accessible network
                        ' Need to eliminate all other options in 'options' table
                        'other(than) 'do nothing' in order to get the model to 
                        ' work, i.e., force a decision. 
                        'PrepareZeroBudgetOptionsTable(sGLPKOptionsTableName, _
                        '                              pFWorkspace, _
                        '                              lGLPKOptionsList)
                        'GLPKShellCall(sGLPKHabitatTableName, _
                        '           sGLPKOptionsTableName, _
                        '           sGLPKConnectTabName, _
                        '           1, _
                        '           f_siOutEID, _
                        '           pFWorkspace, _
                        '           iProgress, _
                        '           sAnalysisCode, _
                        '           "DIR", _
                        '           True)
                    End If

                    If bUndirected = True Then

                        GLPKShellCall(sGLPKHabitatTableName, _
                                      sGLPKOptionsTableName, _
                                      sGLPKConnectTabName, _
                                   iMaxOptionNum, _
                                    f_siOutEID, _
                                 pFWorkspace, _
                                 iProgress, _
                                sAnalysisCode, _
                               "UNDIR", _
                               False)

                        ' run once to get 'initial state' of network
                        ' measured in permeability weighted accessible network
                        ' Need to eliminate all other options in 'options' table
                        'other(than) 'do nothing' in order to get the model to 
                        ' work, i.e., force a decision. 
                        'PrepareZeroBudgetOptionsTable(sGLPKOptionsTableName, _
                        '                              pFWorkspace, _
                        '                              lGLPKOptionsList)
                        'GLPKShellCall(sGLPKHabitatTableName, _
                        '           sGLPKOptionsTableName, _
                        '           sGLPKConnectTabName, _
                        '           1, _
                        '           f_siOutEID, _
                        '           pFWorkspace, _
                        '           iProgress, _
                        '           sAnalysisCode, _
                        '           "UNDIR", _
                        '           True)
                    End If

                Else ' if Randomization is on then 
                    ' update Options table with random permeabilty and run 
                    ' for as many iterations as are chosen

                    ' 1. First run analysis as usual
                    '    But add a _A after the analysis code ('best guess' treatment)
                    ' 2. After optimization, delete all rows in the options table
                    ' 3. update options table again, with randomized variable (permeability or cost). 
                    ' 4. re-run, but add a _C1 (cost, run #1) or _P1 (perm, run #1) 
                    '    after the analysis code 
                    '  in the case of permeability randomization...
                    ' 5. re-run, but forcing the 'best guess' decisions under the new
                    ' permeability randomization to see the result (ZMAX)

                    j = 0
                    For j = 0 To lGLPKOptionsList.Count - 1

                        'If Not lGLPKUniqueEIDs.Contains(lGLPKOptionsList(j).BarrierEID) Then
                        '    lGLPKUniqueEIDs.Add(lGLPKOptionsList(j).BarrierEID)
                        'End If

                        If lGLPKOptionsList(j).BarrierPerm <> 9999 And lGLPKOptionsList(j).OptionCost <> 9999 Then
                            pRowBuffer = pTable.CreateRowBuffer
                            pRowBuffer.Value(0) = lGLPKOptionsList(j).BarrierEID
                            pRowBuffer.Value(1) = lGLPKOptionsList(j).OptionNum
                            ' keep track of max options for use in input to GLPK Model
                            If iMaxOptionNum < lGLPKOptionsList(j).OptionNum Then
                                iMaxOptionNum = lGLPKOptionsList(j).OptionNum
                            End If
                            pRowBuffer.Value(2) = Math.Round(lGLPKOptionsList(j).BarrierPerm, 2)
                            'pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost, 2)
                            pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost)
                            pCursor = pTable.Insert(True)
                            pCursor.InsertRow(pRowBuffer)
                            pCursor.Flush()
                        End If
                    Next

                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    If iProgress < 70 Then
                        iProgress = iProgress + 1
                    End If
                    backgroundworker1.ReportProgress(iProgress, "Performing Optimization Analysis" & ControlChars.NewLine & _
                                                 "User Flag " & (i + 1).ToString & " of " & (pOriginaljuncFlagsList.Count).ToString & _
                                                 ControlChars.NewLine & "Primary Run")

                    If bDirected = True Then
                        GLPKShellCall(sGLPKHabitatTableName, _
                                 sGLPKOptionsTableName, _
                                 sGLPKConnectTabName, _
                                 iMaxOptionNum, _
                                 f_siOutEID, _
                                 pFWorkspace, _
                                 iProgress, _
                                 sAnalysisCode + "_A", _
                                 "DIR", _
                                 False)
                        If bPermRandomization = True Then
                            runRandomization(sGLPKHabitatTableName, _
                                             sGLPKOptionsTableName, _
                                             sGLPKConnectTabName, _
                                             "DIR", _
                                             pTable, _
                                             lGLPKOptionsList, _
                                             iProgress, _
                                             iMaxOptionNum, _
                                             f_siOutEID, _
                                             pFWorkspace, _
                                             sAnalysisCode, _
                                             sPrefix)
                        End If
                        If bCostRandomization = True Then

                        End If
                    End If ' DIR

                    ' For SA, need to know the decisions reached -- added option to store in global variable in GLPKShell Call
                    ' m_lSABestGuessDecisionsObject
                    ' 
                    If bUndirected = True Then
                        GLPKShellCall(sGLPKHabitatTableName, _
                                 sGLPKOptionsTableName, _
                                 sGLPKConnectTabName, _
                                 iMaxOptionNum, _
                                 f_siOutEID, _
                                 pFWorkspace, _
                                 iProgress, _
                                 sAnalysisCode + "_A", _
                                 "UNDIR", _
                                 False)
                        If bPermRandomization = True Then
                            runRandomization(sGLPKHabitatTableName, _
                                           sGLPKOptionsTableName, _
                                           sGLPKConnectTabName, _
                                           "UNDIR", _
                                           pTable, _
                                           lGLPKOptionsList, _
                                           iProgress, _
                                           iMaxOptionNum, _
                                           f_siOutEID, _
                                           pFWorkspace, _
                                           sAnalysisCode, _
                                           sPrefix)
                        End If

                        If bCostRandomization = True Then

                        End If
                    End If ' UNDIR
                End If ' RANDOMIZATION - SA IS ON

                'MsgBox("Debug:58")
            End If ' GLPK is on
            ' ===== END GLPK CALCULATION =======


            End If ' output to dbf = true

            ' REset barriers?
            ' clear current barriers
            pNetworkAnalysisExtBarriers.ClearBarriers()

            'pEdgeFlagDisplay = New IEdgeFlagDisplay
            '               Reset things the way the user had them


            ' ========================= RESET BARRIERS ===========================
            'MsgBox("Debug:59")
            If bAdvConnectTab = True Then
                'MsgBox("Number of 'barriers' before: " & Str(pOriginalBarriersList.Count))
                pOriginalBarriersList = Nothing
                'MsgBox("Number of 'barriers' in pOriginalBarriersListSaved: " & Str(pOriginalBarriersListSaved.Count))
                pOriginalBarriersList = pOriginalBarriersListSaved
                'MsgBox("Number of 'barriers' in pOriginalBarriersListSaved: " & Str(pOriginalBarriersListSaved.Count))
                'MsgBox("Number of 'barriers' after: " & Str(pOriginalBarriersList.Count))
            End If

            m = 0
            pOriginalBarriersList.Reset()
            For m = 0 To pOriginalBarriersList.Count - 1
                bEID = pOriginalBarriersList.Next()
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
            ' ====================== END RESET BARRIERS ===========================

            ' =================================
            ' Update Habitat List Object with permeability weighted habitat
            ' if this option is selected.  
            ' Inputs: Habitat List Object (not unique to every flag)
            '         Metrics List Object (not unique to every flag)
            '         Connectivity List Object (unique to every flag)



        Next ' Flag /sink


        ' ============== 2020 DON'T WRITE REDUNDANT METRICS Objects And Correct ==========
        ' ################################################################################
        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020
        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020
        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020

        ' to do: reevaluate the metrics and habitat to remove reference to branch junctions (gets
        ' too long in table output! 10,000's of thousands of rows
        ' 

        ' for each row in habitat table
        ' do this until EID found is not junction 
        '   is EID branch junction? 
        '   if no find the next downstream barrier and repeat
        '   else exit do loop
        '
        ' if the EID and hab class is not already in the new hab list
        '   insert the EID and hab class and value
        'else update the new hab list value for the EID/class 

        ' will need a phabitatobject 

        j = 0
        For j = 0 To lConnectivity.Count - 1
            '       
            'lConnectivity(j).DownstreamBarrierID
            'lConnectivity(j).BarrID
        Next

        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020
        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020
        ' 2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020_2020
        ' ################################################################################




        ' if the 'advanced' connectivity table is output
        ' then duplicate metrics, habitat, connectivity tables
        ' then fix the original tables to remove reference to junctions
        ' (will need to do network search using these tables to fix by 
        '  aggregating habitat upstream of branch junctions to the nearest downstream barrier)
        ' ============== END 2020 Duplicate and Correct Tables ===========


        'MsgBox("Debug:60")
        If bDBF = True Then
            ' check if user has hit 'close/cancel'
            If m_bCancel = True Then
                backgroundworker1.CancelAsync()
                backgroundworker1.Dispose()
                Exit Sub
            End If
            backgroundworker1.ReportProgress(iProgress + 10, "Writing to DBF Tables")

            ' ================ BEGIN WRITE TO METRICS TABLE ===============
            ' Insert DCI values into the 'Metrics' table
            pTable = pFWorkspace.OpenTable(sMetricTableName)
            j = 0
            For j = 0 To lMetricsObject.Count - 1

                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(1) = lMetricsObject(j).Sink
                pRowBuffer.Value(2) = lMetricsObject(j).SinkEID
                pRowBuffer.Value(3) = lMetricsObject(j).ID
                pRowBuffer.Value(4) = lMetricsObject(j).BarrEID
                pRowBuffer.Value(5) = lMetricsObject(j).Type
                pRowBuffer.Value(6) = lMetricsObject(j).MetricName
                pRowBuffer.Value(7) = lMetricsObject(j).Metric

                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            Next
            ' ============ END INSERT DCI's to METRICS TABLE ==================
            'MsgBox("Debug:61")


            ' ============ BEGIN INSERT STATS INTO HABITAT TABLES =============
            pTable = pFWorkspace.OpenTable(sHabTableName)
            j = 0
            For j = 0 To lHabStatsList.Count - 1

                ' 2020 - do not print the line length to this file
                '       optionally can do this in the future
                If lHabStatsList(j).LengthOrHabitat = "Habitat" Then
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(1) = lHabStatsList(j).Sink
                    pRowBuffer.Value(2) = lHabStatsList(j).SinkEID
                    pRowBuffer.Value(3) = lHabStatsList(j).bID
                    pRowBuffer.Value(4) = lHabStatsList(j).bEID
                    pRowBuffer.Value(5) = lHabStatsList(j).bType
                    pRowBuffer.Value(6) = lHabStatsList(j).Layer
                    pRowBuffer.Value(7) = lHabStatsList(j).Direction
                    pRowBuffer.Value(8) = lHabStatsList(j).TotalImmedPath
                    pRowBuffer.Value(9) = lHabStatsList(j).UniqueClass
                    pRowBuffer.Value(10) = lHabStatsList(j).ClassName
                    pRowBuffer.Value(11) = lHabStatsList(j).Quantity
                    pRowBuffer.Value(12) = lHabStatsList(j).Unit

                    pCursor = pTable.Insert(True)
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                End If
            Next
        End If ' write to DBF is True

        ' ======== END WRITE TO TABLES ====

        ' clear current barriers
        pNetworkAnalysisExtBarriers.ClearBarriers()

        Dim pEdgeFlagDisplay As IEdgeFlagDisplay

        ' ========================= RESET BARRIERS ===========================
        '               Reset things the way the user had them
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(iProgress + 10, "Resetting Flags And Barriers")

        m = 0
        pOriginalBarriersList.Reset()
        For m = 0 To pOriginalBarriersList.Count - 1
            bEID = pOriginalBarriersList.Next()
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
        ' ====================== END RESET BARRIERS ===========================
        'MsgBox("Debug:62")
        ' Clear current flags
        pNetworkAnalysisExtFlags.ClearFlags()

        ' ======================== RESET FLAGS ================================
        ' restore all EDGE flags
        m = 0
        pOriginalEdgeFlagsList.Reset()
        For m = 0 To pOriginalEdgeFlagsList.Count - 1

            iEID = pOriginalEdgeFlagsList.Next()
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

            iEID = pOriginaljuncFlagsList.Next()
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

        'MsgBox("Debug:63")
        ' Create a result highlight of all areas traced
        pTotalResultsEdges = CType(pTotalResultsEdgesGEN, IEnumNetEID)
        pTotalResultsJunctions = CType(pTotalResultsJunctionsGEN, IEnumNetEID)
        pNetworkAnalysisExtResults.ResultsAsSelection = False
        pNetworkAnalysisExtResults.SetResults(pTotalResultsJunctions, pTotalResultsEdges)
        pNetworkAnalysisExtResults.ResultsAsSelection = True

        ' =========================== END RESET FLAGS =====================
        ' Bring results form to front
        'pResultsForm.BringToFront()
        'pResultsForm.txtRichResults.Select(0, 0)
        EndTime = DateTime.Now
        'pResultsForm.lblBeginTime.Text = "Begin Time: " & BeginTime
        'pResultsForm.lblEndTime.Text = "End Time: " & EndTime

        Dim TotalTime As TimeSpan
        TotalTime = EndTime - BeginTime

        'pResultsForm.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
        'pResultsForm.lblDirection.Text = "Analysis Direction: " + sDirection
        'If iOrderNum <> 999 Then
        '    'pResultsForm.lblOrder.Text = "Order of Analysis: " & CStr(iOrderNum)
        'Else
        '    'pResultsForm.lblOrder.Text = "Order of Analysis: Max"
        'End If

        'If Not pAllFlowEndBarriers Is Nothing Then
        '    If pAllFlowEndBarriers.Count <> 0 Then
        '        pResultsForm.lblNumBarriers.Text = "Number of Barriers Analysed: " & CStr(pAllFlowEndBarriers.Count + pOriginaljuncFlagsList.Count)
        '    Else
        '        pResultsForm.lblNumBarriers.Text = "Number of Barriers Analysed: 1"
        '    End If
        'End If

        ' ================== BEGIN WRITE TO OUTPUT FORM =================
        'MsgBox("Debug:64")
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(70, "Prepping Output Form")

        ' Output Form (will replace dockable window)
        Dim pResultsForm3 As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3
        pResultsForm3.Show()

        'Dim table As DataTable = New DataTable("Summary Results")
        'Dim column As DataColumn = New DataColumn("Sink ID", GetType(System.String))
        'Table.Columns.Add(column)

        'column = New DataColumn("bID", GetType(System.String))
        'table.Columns.Add(column)
        'column = New DataColumn("Type", GetType(System.String))
        'table.Columns.Add(column)

        ' the habitat stats list contains quantities for each class used
        ' if classes are used.  Otherwise they contain totals.  
        ' will need to re-read user settings in order to determine what CAN go into 
        ' the output.

        ' Logic - check our settings and what traces are being done. 
        ' give the user the available options
        ' given their choice determine how the summary table is made.  

        ' If ANY of the habitat trace types are checked, then the Habstatslist will have content. 
        ' Otherwise just the MetricsStatsObject will no matter what

        ' CASE 1: User wants classes, user wants totals.  
        ' So if the user wants to include classes in their output summary table
        ' then the classes can be taken from lHabStatsList and 
        ' IF the user is calculating DCI then 
        ' the total value can be grabbed from the lDCIStatList object. 
        ' However, in this case (user WANTS classes included in output) the 
        '  can be taken from the DCI hab list and probably SHOULD be because
        ' it's being looped through anyway - why loop through a second list? just
        ' grab a running total/tally from the first list as we move through.

        ' Hab stats will NOT be used if no traces are checked, but at least one always should be 

        ' Need to check IF there are HABITAT CLASSES set.
        '  YES
        '     LOOP through all columns
        '       check IF HAB CLASS is there
        '         NO
        '           ADD it as a COLUMN
        '       
        ' For now, only the total habitat will be included in the output
        ' for area and length. 
        '  If the user is using classes then they will have to be summed
        ' no matter what, it will have to be summed for each trace type done.  

        ' The easiest way to is to check the length of each object
        ' if it's not empty, then use it. 

        Dim pSinkAndDCIS As New SinkandDCIs(Nothing, Nothing, Nothing, Nothing)
        Dim lSinkAndDCIS As New List(Of SinkandDCIs)
        Dim pSinkIDAndTypes As New SinkandTypes(Nothing, Nothing, Nothing)
        Dim lSinkIDandTypes As New List(Of SinkandTypes)


        i = 0
        Dim bSinkThere, bDCIpMatch, bDCIdMatch, bEntered As Boolean
        Dim row As DataRow
        'Dim pSinkTable As DataTable = New DataTable("Sink Table")

        'column = New DataColumn("Sink", GetType(System.String))
        'pSinkTable.Columns.Add(column)
        ''pSinkTable.PrimaryKey = column
        'column = New DataColumn("Type", GetType(System.String))
        'pSinkTable.Columns.Add(column)

        Dim bPrintToResultsForm As Boolean = True

        If lMetricsObject.Count > 50 Then
            Dim sMetricsCount As String = Convert.ToString(lMetricsObject.Count)
            Dim result = Windows.Forms.MessageBox.Show("There are more than " & sMetricsCount & " barriers / junctions analyzed.  It may take some time to write to results form.  Continue?", "continue", Windows.Forms.MessageBoxButtons.YesNo)
            If result = Windows.Forms.DialogResult.No Then
                bPrintToResultsForm = False
            End If
        End If

        ' gets a list of the unique sinks from the 
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
        ' to do 2020 - this should be a separate sub wrapped in following 'if'
        If bPrintToResultsForm = True Then

            '  column one   - sink ID
            '  column two   - sink stat label
            '  column three - sink stat
            '  column four  - barrier ID
            '  column five  - trace direction (i.e. upstream:)  (Loop)
            '  column six   - trace type (i.e. total)
            '  column six   - habitat layer (i.e. Lines) (Loop) (OMIT)
            '  column seven - habitat class (i.e. river or stream)(Loop)
            '                 if there are classes insert a 'total' at the bottom 
            '                 of the list of classes (run totalling function?)
            '  column eight - the quantity
            '  column nine  - unit 

            ' For each Sink in the master Sinks List
            '   If it's the first iteration of that sink. 
            '     if DCI stats were calculated then 
            '     insert the DCIp stat label and value in column 2 and three
            ' 

            ' Do NOT bind the data table so I can program flexible
            ' rows and columns to handle the variety of input data necessary


            ' Set up the table - create columns 

            pResultsForm3.DataGridView1.Columns.Add("Sink", "Sink")           '0
            pResultsForm3.DataGridView1.Columns.Add("SinkID", "SinkID")       '1
            pResultsForm3.DataGridView1.Columns.Add("Type", "Type")           '2
            pResultsForm3.DataGridView1.Columns.Add("Node", "Node")           '3
            pResultsForm3.DataGridView1.Columns.Add("NodeID", "NodeID")       '4
            pResultsForm3.DataGridView1.Columns.Add("Metric", "Metric")       '5
            pResultsForm3.DataGridView1.Columns.Add("Value", "Value")         '6

            ' If there are habitat statistics then add the proper columns
            If lHabStatsList.Count > 0 Then
                pResultsForm3.DataGridView1.Columns.Add("Layer", "Layer")                 '7
                pResultsForm3.DataGridView1.Columns.Add("Direction", "Direction")         '8
                pResultsForm3.DataGridView1.Columns.Add("Type", "Type")                   '9
                pResultsForm3.DataGridView1.Columns.Add("HabitatClass", "Habitat_Class")  '10
                pResultsForm3.DataGridView1.Columns.Add("Quantity", "Quantity")           '11
                pResultsForm3.DataGridView1.Columns.Add("Unit", "Unit")                   '12
            End If

            i = 0
            For i = 0 To pResultsForm3.DataGridView1.Columns.Count - 1
                pResultsForm3.DataGridView1.Columns.Item(i).SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            Next i

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
            Dim refinedGLPKOptionsList As List(Of GLPKOptionsObject)

            Dim HabStatsComparer As RefineHabStatsListPredicate

            Dim pDataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle

            ' two tables joined manually (no SQL 'join' available):
            ' - the metrics table / list
            ' and 
            ' - the habitat statistics table / list
            '  the table added second should tend to be the larger, 
            '  and the table going second needs to know the  
            '  number of rows already inserted in the table so it 
            '  knows whether to insert another. 
            '  iterate through both tables for each unique sink


            ' For each sink
            ' 1. for each sink in the master sinks object list
            ' 2. for each barrier associated with the sink in the master barriers list.  
            ' 2a add the metrics.

            ' notes: -the maxrow index keeps track of which row we're at,
            '        it's needed if there are multiple sinks
            '        -the isinkrow count keeps track of which row for this
            '        sink we're at so that the first row can be found
            '        -results form is populated with sinkID, not EID

            'iProgressIncrementFactor = (40 / lBarrierAndSinkEIDs.Count)

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
                            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Value = lMetricsObject(k).MetricName
                            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(6).Value = Math.Round(lMetricsObject(k).Metric, 2)
                        ElseIf iSinkRowCount > 0 Then
                            pResultsForm3.DataGridView1.Rows.Add()
                            ' keep track of the maximum number of rows in the table
                            iMaxRowIndex = iMaxRowIndex + 1 ' Row tracker
                            'pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Value = lMetricsObject(j).ID
                            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(5).Value = lMetricsObject(k).MetricName
                            pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(6).Value = Math.Round(lMetricsObject(k).Metric, 2)

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
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    If iProgress < 90 Then
                        backgroundworker1.ReportProgress(iProgress + 1, "Writing to Output Form")
                    End If
                    'MsgBox("Debug:65")
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

                'MsgBox("Debug:66")

                ' For each barrier
                '  1. a refined list of all habitat stats for this barrier 
                '     and sink and layer
                '  2. for each layer get a refined list of habitat metrics 
                '     associated with each layer, sink, barrier combo
                k = 0
                For k = 0 To refinedBarrierEIDList.Count - 1

                    iBarrRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    iMaxRowIndex = iBarrRowIndex
                    iBarrRowCount = 0
                    iSinkRowCount += 1
                    iBarrRowCount += 1
                    bTrigger = False
                    If m_bCancel = True Then
                        backgroundworker1.CancelAsync()
                        backgroundworker1.Dispose()
                        Exit Sub
                    End If
                    If iProgress < 90 Then
                        backgroundworker1.ReportProgress(iProgress + 1, "Writing to Output Form" & ControlChars.NewLine & _
                                                         " Barrier / Node #: " & k.ToString)
                    End If
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
                    pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(3).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(4).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iMaxRowIndex).Cells(2).Value = refinedBarrierEIDList(k).NodeType
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

                    'MsgBox("Debug:67")
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

            'MsgBox("Debug:68")
            pResultsForm3.DataGridView1.AutoResizeColumns()
            EndTime = DateTime.Now
            pResultsForm3.lblBeginTime.Text = "Begin Time: " & BeginTime
            pResultsForm3.lblEndtime.Text = "End Time: " & EndTime

            TotalTime = EndTime - BeginTime
            pResultsForm3.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
            pResultsForm3.lblDirection.Text = "Analysis Direction: " + sDirection
            If iOrderNum <> 999 Then
                'MsgBox("Debug:69")
                pResultsForm3.lblOrder.Text = "Order of Analysis: " & CStr(iOrderNum)
                'MsgBox("Debug:70")
            Else
                pResultsForm3.lblOrder.Text = "Order of Analysis: Max"
            End If

            If Not pAllFlowEndBarriers Is Nothing Then
                If pAllFlowEndBarriers.Count <> 0 Then
                    'MsgBox("Debug:71")
                    pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: " & CStr(pAllFlowEndBarriers.Count + pOriginaljuncFlagsList.Count)
                    'MsgBox("Debug:72")
                Else
                    pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: 1"
                End If
            End If

            ' refresh the view
            pActiveView.Refresh()
            pResultsForm3.BringToFront()
            pResultsForm3.Activate()

        End If ' bPrintToResultsForm = True

        ' ============== END WRITE TO OUTPUT SUMMARY TABLE =================

        ' review of objects used in output:
        'lHabStatsList

        ' = lHabStatsList(j).Sink
        ' = lHabStatsList(j).bID
        ' = lHabStatsList(j).bType
        ' = lHabStatsList(j).Layer
        ' = lHabStatsList(j).Direction
        ' = lHabStatsList(j).TotalImmedPath
        ' = lHabStatsList(j).UniqueClass
        ' = lHabStatsList(j).ClassName
        ' = lHabStatsList(j).Quantity
        ' = lHabStatsList(j).Unit

        ' = lDCIStatsList(j).Barrier
        ' = lDCIStatsList(j).Quantity
        ' = lDCIStatsList(j).BarrierPerm
        ' = lDCIStatsList(j).BarrierYN

        ' = lMetricsObject(j).Sink
        ' = lMetricsObject(j).ID
        ' = lMetricsObject(j).Type
        ' = lMetricsObject(j).MetricName
        ' = lMetricsObject(j).Metric



        ' The output form contains a variable width table
        ' with a minimum set of data.  It will be up to the user to select
        ' whether they want this form or not.  This form could contain all
        ' the data presented in the two tables, Habitat and Metrics, and it 
        ' could contain very little of it.Can think about this table as a 1:M
        ' join - lots of wasted space and redundancy.  

        'MsgBox("Debug:73")
        ' ================== END WRITE TO OUTPUT FORM ===================
        ' check if user has hit 'close/cancel'
        If m_bCancel = True Then
            backgroundworker1.CancelAsync()
            backgroundworker1.Dispose()
            Exit Sub
        End If
        backgroundworker1.ReportProgress(100, "Completed!")

        backgroundworker1.Dispose()
        backgroundworker1.CancelAsync()


    End Sub
    Private Sub PrepareZeroBudgetOptionsTable(ByVal sGLPKOptionsTableName As String, _
                                              ByRef pFWorkspace As IFeatureWorkspace, _
                                              ByVal lGLPKOptionsList As List(Of GLPKOptionsObject))

        Dim iMaxOptionNum As Integer
        Dim pTable As ITable
        Dim iCountInsertOptionsTemp As Integer
        Dim bCreateRow As Boolean
        Dim pRowBuffer As IRowBuffer
        Dim pCursor As ICursor
        Dim lTempOptionsObject As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject)
        Dim pTempGLPKOptionsObject As GLPKOptionsObject
        Dim j As Integer = 0

        ' saves all options (including new ones for big barriers)
        lTempOptionsObject = New List(Of GLPKOptionsObject)
        Try

            j = 0
            For j = 0 To lGLPKOptionsList.Count - 1


                pTempGLPKOptionsObject = New GLPKOptionsObject(Nothing, _
                                                               Nothing, _
                                                               Nothing, _
                                                               Nothing, _
                                                               Nothing)
                ' keep list of random permeabilities used
                If lGLPKOptionsList(j).BarrierPerm <> 9999 And lGLPKOptionsList(j).OptionCost <> 9999 Then

                    pTempGLPKOptionsObject.BarrierEID = lGLPKOptionsList(j).BarrierEID
                    pTempGLPKOptionsObject.OptionNum = lGLPKOptionsList(j).OptionNum
                    pTempGLPKOptionsObject.BarrierPerm = lGLPKOptionsList(j).BarrierPerm
                    pTempGLPKOptionsObject.OptionCost = Math.Round(lGLPKOptionsList(j).OptionCost)
                    pTempGLPKOptionsObject.BarrierType = lGLPKOptionsList(j).BarrierType

                    ' new permeabilities are stored in this list (for 'do nothing' run next)
                    lTempOptionsObject.Add(pTempGLPKOptionsObject)

                End If
            Next 'j 

        Catch ex As Exception
            MsgBox("Error code 4232. " + ex.Message)
        End Try



        iMaxOptionNum = 0
        pTable = pFWorkspace.OpenTable(sGLPKOptionsTableName)
        Try
            ' Delete all rows in the options table
            pTable.DeleteSearchedRows(Nothing)
            iCountInsertOptionsTemp = 0
            j = 0
            For j = 0 To lTempOptionsObject.Count - 1

                bCreateRow = False
                If lTempOptionsObject(j).OptionNum = 1 Then
                    bCreateRow = True
                End If

                If lTempOptionsObject(j).BarrierPerm <> 9999 And lTempOptionsObject(j).OptionCost <> 9999 Then
                    If bCreateRow = True Then
                        pRowBuffer = pTable.CreateRowBuffer
                        pRowBuffer.Value(0) = lTempOptionsObject(j).BarrierEID
                        pRowBuffer.Value(1) = lTempOptionsObject(j).OptionNum
                        ' keep track of max options for use in input to GLPK Model
                        If iMaxOptionNum < lTempOptionsObject(j).OptionNum Then
                            iMaxOptionNum = lTempOptionsObject(j).OptionNum
                        End If
                        pRowBuffer.Value(2) = lTempOptionsObject(j).BarrierPerm
                        pRowBuffer.Value(3) = Math.Round(lTempOptionsObject(j).OptionCost)
                        pCursor = pTable.Insert(True)
                        pCursor.InsertRow(pRowBuffer)
                        pCursor.Flush()
                        ' this counter keeps track of 'budget' for 'do nothing'
                        'iCountInsertOptionsTemp += 1
                    End If
                End If
            Next ' j option
        Catch ex As Exception
            MsgBox("Trouble creating 'do nothing' options table for the zero budget run in SA. " + _
                   "Code 40. RunRandomization routine. " + ex.Message)
        End Try

    End Sub
    Private Sub runRandomization(ByVal sGLPKHabitatTableName As String, _
                                 ByVal sGLPKOptionsTableName As String, _
                                 ByVal sGLPKConnectTabName As String, _
                                 ByVal sAnalysisType As String, _
                                 ByRef pTable As ITable, _
                                 ByVal lGLPKOptionsList As List(Of GLPKOptionsObject), _
                                 ByVal iProgress As Integer, _
                                 ByVal iMaxOptionNum As Integer, _
                                 ByVal f_siOutEID As Integer, _
                                 ByRef pFWorkspace As IFeatureWorkspace, _
                                 ByVal sAnalysisCode As String, _
                                 ByVal sPrefix As String)

        ' Sub Description
        ' Deletes all rows in options table
        ' run through, forcing the decisions to the 'best guess' decisions
        ' (not needed for cost randomization SA)
        ' (sAnalysisType parameter keeps track if this is a DIR or UNDIR analysis)
        ' run the optimization analysis (_A[X] suffix where X = randomization count)
        '  (passes 'analysis type' to it to tell GLPKShellCall that it's only DIR or UNDIR)
        ' Update all rows in options table again -- this time need to do a randomization for 
        ' those barriers that were forced.  (In previous run the permeability would not have 
        ' been randomized for these barriers, since they were forced and option 1 was not there)
        ' run the optimization analysis (_P[X]) suffix where X = randomization count)


        'Number of randomized permeabilities to use
        Dim iPermRandomCount As Integer = 30
        Dim bExcludeDams As Boolean = True
        Dim dRandomPerm As Double

        Dim liRandomPerm As List(Of Double) = New List(Of Double)
        Dim bCreateRow As Boolean
        Dim bUseRandomPerm As Boolean
        Dim bEIDMatch As Boolean
        Dim j As Integer = 0
        Dim pRowBuffer As IRowBuffer
        Dim pCursor As ICursor
        Dim iTracker As Integer

        Dim lSABestGuessDecisionsObject As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject)
        If sAnalysisType = "DIR" Then
            lSABestGuessDecisionsObject = m_lSABestGuessDecisionsObject_DIR
        ElseIf sAnalysisType = "UNDIR" Then
            lSABestGuessDecisionsObject = m_lSABestGuessDecisionsObject_UNDIR
        End If

        If lSABestGuessDecisionsObject.Count = 0 Then
            MsgBox("Warning: the 'best guess' decision set for this sensitivity analysis (" + sAnalysisType + _
                   ") is empty. Warning is from runRandomization subroutine.")
        End If

        Dim lTempOptionsObject As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject)
        Dim pTempGLPKOptionsObject As GLPKOptionsObject
        Dim iCountInsertOptionsTemp As Integer = 0

        ' loop for number of randomizations
        For k = 0 To iPermRandomCount - 1


            Try

                ' Delete all rows in the options table
                pTable.DeleteSearchedRows(Nothing)
                Randomize() 'initializes random number gen
                liRandomPerm = New List(Of Double) ' reset rand perm list

                ' run through, forcing the decisions to the ones made earlier, 
                ' stored by GLPKShellCall in m_lSABestGuessDecisionsObject
                j = 0
                For j = 0 To lGLPKOptionsList.Count - 1

                    ' For each barrier / option combo in the lGLPKOptionsList
                    bCreateRow = False
                    bUseRandomPerm = False
                    bEIDMatch = False

                    ' keep list of random permeabilities used
                    If lGLPKOptionsList(j).BarrierPerm <> 9999 And lGLPKOptionsList(j).OptionCost <> 9999 Then

                        ' do not create row if this is a barrier from the decision set (m_lSABestGuessDecisionsObject)
                        ' and if the option here is 1
                        ' also do not create a row if this is not a barrier from the decision set and the 
                        ' option is greater than 1. (in this case use a randomized variable)
                        ' NOTE: WOULD NEED A CHECK HERE IF UNDIR IS RUN, AND A SEPARATE "J" LOOP
                        '       THIS ASSUMES GRB IS BEING USED, ASSUMES ONLY OPTIONS > 1 ARE IN THE OUTPUT
                        '       FROM GLPKSHELLCALL (which is then transfered to the global variable)
                        For m = 0 To lSABestGuessDecisionsObject.Count - 1
                            If lGLPKOptionsList(j).BarrierEID = lSABestGuessDecisionsObject(m).BarrierEID Then
                                bEIDMatch = True
                                If lGLPKOptionsList(j).OptionNum = lSABestGuessDecisionsObject(m).DecisionOption Then
                                    bCreateRow = True
                                End If
                            End If
                        Next

                        ' if this isn't a barrier in the 'decision' object list
                        ' then if the option is '1', 
                        ' then create the row 
                        ' if the barrier is a "CULVERT" then randomize the perm
                        If bEIDMatch = False Then

                            ' don't randomize the flag... 
                            ' HARD CODED --> IF FLAG IS ON BARRIER THIS WON'T GET RANDOMIZED
                            If lGLPKOptionsList(j).OptionNum = 1 Then
                                bCreateRow = True

                                If lGLPKOptionsList(j).BarrierType <> "FLAG" Then
                                    If bExcludeDams = True Then
                                        If lGLPKOptionsList(j).BarrierType <> "DAM" Then
                                            bUseRandomPerm = True
                                        End If
                                    Else
                                        bUseRandomPerm = True
                                    End If
                                End If
                            End If
                        End If

                        If bCreateRow = True Then

                            pRowBuffer = pTable.CreateRowBuffer
                            pRowBuffer.Value(0) = lGLPKOptionsList(j).BarrierEID
                            pRowBuffer.Value(1) = lGLPKOptionsList(j).OptionNum
                            ' keep track of max options for use in input to GLPK Model
                            If iMaxOptionNum < lGLPKOptionsList(j).OptionNum Then
                                iMaxOptionNum = lGLPKOptionsList(j).OptionNum
                            End If

                            ' don't update permeability with random permeability unless 
                            ' it's option 1 ('do nothing')
                            If bUseRandomPerm = True Then

                                dRandomPerm = Rnd()
                                pRowBuffer.Value(2) = Math.Round(dRandomPerm, 2)

                                ' add the random value to a list of re-use later. 
                                ' (later, if the barrier is a 'fix' barrier, another random num
                                '  will be generated, because it wouldn't have been here)
                                liRandomPerm.Add(Math.Round(dRandomPerm, 2))
                            Else
                                pRowBuffer.Value(2) = lGLPKOptionsList(j).BarrierPerm
                            End If

                            'pRowBuffer.Value(2) = Math.Round(lGLPKOptionsList(j).BarrierPerm, 2)
                            'pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost, 2)
                            pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost)
                            pCursor = pTable.Insert(True)
                            pCursor.InsertRow(pRowBuffer)
                            pCursor.Flush()
                        End If ' creatrow is true
                    End If
                Next 'j

            Catch ex As Exception
                MsgBox("Error round 1. " + ex.Message)
            End Try

            If m_bCancel = True Then
                backgroundworker1.CancelAsync()
                backgroundworker1.Dispose()
                Exit Sub
            End If
            If iProgress < 70 Then
                iProgress = iProgress + 1
            End If

            backgroundworker1.ReportProgress(iProgress, "Performing Sensitivitiy Analysis" & ControlChars.NewLine & _
                                 ControlChars.NewLine & "Randomization Run #" + k.ToString)

            ' DO SA FOR 
            Try

                GLPKShellCall(sGLPKHabitatTableName, _
                              sGLPKOptionsTableName, _
                              sGLPKConnectTabName, _
                              iMaxOptionNum, _
                              f_siOutEID, _
                              pFWorkspace, _
                              iProgress, _
                              sAnalysisCode + "_A" + k.ToString, _
                              sAnalysisType, _
                              False)

            Catch ex As Exception
                MsgBox("Error - round 1ba. " + ex.Message)
            End Try

            ' now recreate the options table a second time, but do not force the model to choose the
            ' decision variables.  This will not work if the global 'decisions' list m_lSABestGuessDecisionsObject_DIR
            ' created in GLPKShellCall is created wrongly -- that is, it should be created only if the suffix
            ' of the sAnalysisCode ends in _A, not _P or _A1. 
            ' Will check each barrier EID for a match (unfortunately) to the decisions in 'best guess' object
            ' This is so the same randomized permeabilities can be assigned to the same barriers.  
            ' AND if it is a barrier match, we will need to generate a new permeability, since it was skipped 
            ' in the last round

            Try

                ' Delete all rows in the options table
                pTable.DeleteSearchedRows(Nothing)
                iTracker = 0 ' helps tracking list index for random perms

                ' saves all options (including new ones for big barriers)
                lTempOptionsObject = New List(Of GLPKOptionsObject)

                j = 0
                For j = 0 To lGLPKOptionsList.Count - 1

                    bCreateRandomPerm = False
                    bUseRandomPerm = False

                    pTempGLPKOptionsObject = New GLPKOptionsObject(Nothing, _
                                                                   Nothing, _
                                                                   Nothing, _
                                                                   Nothing, _
                                                                   Nothing)
                    ' keep list of random permeabilities used
                    If lGLPKOptionsList(j).BarrierPerm <> 9999 And lGLPKOptionsList(j).OptionCost <> 9999 Then

                        ' DETERMINE IF RANDOM PERM IS NEEDED
                        ' NOTE HERE IS ONLY USING DECISIONS FROM THE 'DIRECTED' MODEL!
                        ' Only needed if optionnum is 1 (do nothing)
                        If lGLPKOptionsList(j).OptionNum = 1 And lGLPKOptionsList(j).BarrierType <> "FLAG" Then

                            If bExcludeDams = True And lGLPKOptionsList(j).BarrierType <> "DAM" Then
                                bUseRandomPerm = True
                            ElseIf bExcludeDams = False Then
                                bUseRandomPerm = True
                            End If

                            ' each of the decision barriers
                            ' if this barrier is in the decision list
                            ' a randomized permeability was not created in the last run. 
                            ' so need to here
                            If bUseRandomPerm = True Then
                                For m = 0 To lSABestGuessDecisionsObject.Count - 1
                                    If lGLPKOptionsList(j).BarrierEID = lSABestGuessDecisionsObject(m).BarrierEID Then

                                        ' EXCLUDE DAMS IF NEEDED
                                        If bExcludeDams = True Then
                                            If lGLPKOptionsList(j).BarrierType = "DAM" Then
                                                Exit For
                                            Else
                                                bCreateRandomPerm = True
                                                Exit For
                                            End If
                                        Else
                                            bCreateRandomPerm = True
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        pRowBuffer = pTable.CreateRowBuffer
                        pRowBuffer.Value(0) = lGLPKOptionsList(j).BarrierEID
                        pTempGLPKOptionsObject.BarrierEID = lGLPKOptionsList(j).BarrierEID

                        pRowBuffer.Value(1) = lGLPKOptionsList(j).OptionNum
                        pTempGLPKOptionsObject.OptionNum = lGLPKOptionsList(j).OptionNum

                        ' keep track of max options for use in input to GLPK Model
                        If iMaxOptionNum < lGLPKOptionsList(j).OptionNum Then
                            iMaxOptionNum = lGLPKOptionsList(j).OptionNum
                        End If

                        ' insert a random value from the stored ran perm list 
                        ' if the option num is 1 and the EID is not a decision EID
                        ' insert a new generated random number if the option num is 1
                        ' and the EID is a decision EID
                        ' if the option num > 1 then do not insert a random number, use
                        ' the actual estimated gain (for thesis usually 1)
                        ' iTracker keeps track of how many inserts were done that would not 
                        ' have been done before
                        If bUseRandomPerm = False Then
                            pRowBuffer.Value(2) = lGLPKOptionsList(j).BarrierPerm
                            pTempGLPKOptionsObject.BarrierPerm = lGLPKOptionsList(j).BarrierPerm
                        Else
                            If bCreateRandomPerm = True Then
                                dRandomPerm = Rnd()
                                pRowBuffer.Value(2) = Math.Round(dRandomPerm, 2)
                                pTempGLPKOptionsObject.BarrierPerm = Math.Round(dRandomPerm, 2)
                            Else
                                pRowBuffer.Value(2) = liRandomPerm(iTracker)
                                pTempGLPKOptionsObject.BarrierPerm = liRandomPerm(iTracker)
                                iTracker += 1
                            End If
                        End If

                        'pRowBuffer.Value(2) = Math.Round(lGLPKOptionsList(j).BarrierPerm, 2)
                        'pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost, 2)
                        pRowBuffer.Value(3) = Math.Round(lGLPKOptionsList(j).OptionCost)
                        pTempGLPKOptionsObject.OptionCost = Math.Round(lGLPKOptionsList(j).OptionCost)
                        pTempGLPKOptionsObject.BarrierType = lGLPKOptionsList(j).BarrierType

                        pCursor = pTable.Insert(True)
                        pCursor.InsertRow(pRowBuffer)
                        pCursor.Flush()
                        ' new permeabilities are stored in this list (for 'do nothing' run next)
                        lTempOptionsObject.Add(pTempGLPKOptionsObject)

                    End If
                Next 'j 

            Catch ex As Exception
                MsgBox("Error round 2. " + ex.Message)
            End Try

            If m_bCancel = True Then
                backgroundworker1.CancelAsync()
                backgroundworker1.Dispose()
                Exit Sub
            End If
            If iProgress < 70 Then
                iProgress = iProgress + 1
            End If

            backgroundworker1.ReportProgress(iProgress, "Performing Sensitivity Analysis" & ControlChars.NewLine & _
                                 ControlChars.NewLine & "Randomization Run #" + k.ToString)

            Try
                GLPKShellCall(sGLPKHabitatTableName, _
                              sGLPKOptionsTableName, _
                              sGLPKConnectTabName, _
                              iMaxOptionNum, _
                              f_siOutEID, _
                              pFWorkspace, _
                              iProgress, _
                              sAnalysisCode + "_P" + k.ToString, _
                              sAnalysisType, _
                              False)
            Catch ex As Exception
                MsgBox("Error round 2b. " + ex.Message)
            End Try


            ' update decisions - force only 'do nothing' 
            ' by removing all other options.  
            ' these options and permeabilities (randomized)
            ' are stored with pTempGLPKOptionsObject
            ' (only need to do this with UNDIR)

            If sAnalysisType = "UNDIR" Then
                iMaxOptionNum = 0
                Try
                    ' Delete all rows in the options table
                    pTable.DeleteSearchedRows(Nothing)
                    iCountInsertOptionsTemp = 0
                    j = 0
                    For j = 0 To lTempOptionsObject.Count - 1

                        bCreateRow = False
                        If lTempOptionsObject(j).OptionNum = 1 Then
                            bCreateRow = True
                        End If

                        If lTempOptionsObject(j).BarrierPerm <> 9999 And lTempOptionsObject(j).OptionCost <> 9999 Then
                            If bCreateRow = True Then
                                pRowBuffer = pTable.CreateRowBuffer
                                pRowBuffer.Value(0) = lTempOptionsObject(j).BarrierEID
                                pRowBuffer.Value(1) = lTempOptionsObject(j).OptionNum
                                ' keep track of max options for use in input to GLPK Model
                                If iMaxOptionNum < lTempOptionsObject(j).OptionNum Then
                                    iMaxOptionNum = lTempOptionsObject(j).OptionNum
                                End If
                                pRowBuffer.Value(2) = lTempOptionsObject(j).BarrierPerm
                                pRowBuffer.Value(3) = Math.Round(lTempOptionsObject(j).OptionCost)
                                pCursor = pTable.Insert(True)
                                pCursor.InsertRow(pRowBuffer)
                                pCursor.Flush()
                                ' this counter keeps track of 'budget' for 'do nothing'
                                'iCountInsertOptionsTemp += 1
                            End If
                        End If
                    Next ' j option
                Catch ex As Exception
                    MsgBox("Trouble creating 'do nothing' options table for the zero budget run in SA. " + _
                           "Code 40. RunRandomization routine. " + ex.Message)
                End Try
            End If

            '' RUN A CONTROL - SEND BUDGET OVERRIDE 
            '' GLPKShellCall now calculates for 'zero' budget
            '' suffix is identified using (P0P, P1P, P2P etc.)
            'Try
            '    GLPKShellCall(sGLPKHabitatTableName, _
            '                  sGLPKOptionsTableName, _
            '                  sGLPKConnectTabName, _
            '                  iMaxOptionNum, _
            '                  f_siOutEID, _
            '                  pFWorkspace, _
            '                  iProgress, _
            '                  sAnalysisCode + "_P" + k.ToString + "P", _
            '                  sAnalysisType, _
            '                  True)
            'Catch ex As Exception
            '    MsgBox("Error round 2b. " + ex.Message)
            'End Try

        Next 'k random treatment

        ' Export the common decisions b/w treatments for SA
        ExportCommonDecisions_RandomPerm(pFWorkspace, _
                                         sAnalysisCode, _
                                         lSABestGuessDecisionsObject, _
                                         sAnalysisType)

        ' Export the 'results' tables for SA 
        If sAnalysisType = "DIR" Then
            If m_lSA_A_Results_DIR.Count > 0 Then
                Export_SA_Results_DIR(pFWorkspace, _
                                      sAnalysisCode, _
                                      sPrefix, _
                                      "A")
            End If
            If m_lSA_P_Results_DIR.Count > 0 Then
                Export_SA_Results_DIR(pFWorkspace, _
                                      sAnalysisCode, _
                                      sPrefix, _
                                      "P")
            End If

        ElseIf sAnalysisType = "UNDIR" Then
            If m_lSA_A_Results_UNDIR.Count > 0 Then
                Export_SA_Results_UNDIR(pFWorkspace, _
                                      sAnalysisCode, _
                                      sPrefix, _
                                      "A")
            End If
            If m_lSA_P_Results_UNDIR.Count > 0 Then
                Export_SA_Results_UNDIR(pFWorkspace, _
                                      sAnalysisCode, _
                                      sPrefix, _
                                      "P")
            End If
        End If
      

    End Sub
    Private Sub ExportCommonDecisions_RandomPerm(ByRef pFWorkspace As IFeatureWorkspace, _
                                                 ByRef sAnalysisCode As String, _
                                                 ByVal lSABestGuessDecisionsObject As List(Of GLPKDecisionsObject), _
                                                 ByVal sAnalysisType As String)

        ' called from the runRandomization subroutine
        ' exports a decisions table for randomized permeability
        ' common decisions amongst all
        ' First prep table
        ' then populate table
        ' (lSABestGuessDecisionsObject was set in runRandomization based on analysis type - UNDIR or DIR)

        Dim sTableName As String

        ' Create the table
        Try
            sTableName = TableName(sAnalysisCode, pFWorkspace, "SA_RandomOverlap_" + sAnalysisType)

        Catch ex As Exception
            MsgBox("Problem getting Table name for the SA random perm. " + ex.Message)
        End Try

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

        pFieldsEdit.FieldCount_2 = 3
        iFields = 3

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= Second Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrEID"
        pFieldEdit.Name_2 = "BarrEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(1) = pField

        ' ============= Third Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "DecisionNum"
        pFieldEdit.Name_2 = "DecisionNum"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(2) = pField

        Try
            pFWorkspace.CreateTable(sTableName, pFields, Nothing, Nothing, "")
        Catch ex As Exception
            MsgBox("Problem creating the SA Random Perm export table. " + ex.Message)
        End Try
        ' MsgBox "created table successfully"

        Dim pTable As ITable = pFWorkspace.OpenTable(sTableName)
        Try

            ' Add Table to Map Doc
            pStTabColl = CType(pMap, IStandaloneTableCollection)
            pStTab = New StandaloneTable
            pStTab.Table = pTable
            pStTabColl.AddStandaloneTable(pStTab)
            pMxDoc.UpdateContents()

        Catch ex As Exception
            MsgBox("Probelm adding new SA table to map doc. " + ex.Message)
        End Try


        Dim pRowBuffer As IRowBuffer
        Dim pCursor As ICursor


        Try
            ' UPDATE THE TABLE
            pRowBuffer = pTable.CreateRowBuffer

            For n = 0 To lSABestGuessDecisionsObject.Count - 1
                '  SinkEID (integer)
                '  Treatment (string)
                '  Budget (double)
                '  GLPKSolved (boolean/binary)
                '  Perc_Gap (double)
                '  MAxSolTime(integer)
                '  TimeUsed(double)
                '  HabitatZmax(double)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(1) = lSABestGuessDecisionsObject(n).BarrierEID
                pRowBuffer.Value(2) = lSABestGuessDecisionsObject(n).DecisionOption

                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            Next
        Catch ex As Exception
            MsgBox("Error updating SA Table with common random decisions. " + ex.Message)
        End Try


    End Sub
    Private Sub Export_SA_Results_DIR(ByRef pFWorkspace As IFeatureWorkspace, _
                                      ByRef sAnalysisCode As String, _
                                      ByVal sPrefix As String, _
                                      ByVal sTreatmentType As String)

        ' exports table from stored global variable 
        ' containing 'results' table rows
        ' created in GLPKShellCall for thesis Sensitivity analysis
        ' -- directed model

        Dim pRowBuffer As IRowBuffer
        Dim pTable As ITable
        Dim pCursor As ICursor
        Dim sSAResultsTableName As String
        Dim n As Integer

        Dim lSA_Results_DIR As List(Of DIR_OptResultsObject) = New List(Of DIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS


        If sTreatmentType = "A" Then
            lSA_Results_DIR = m_lSA_A_Results_DIR
        ElseIf sTreatmentType = "P" Then
            lSA_Results_DIR = m_lSA_P_Results_DIR
        End If



        ' ---------------
        ' Create and write to 'results' table for SA Analysis
        ' results table
        Try
            sSAResultsTableName = TableName("SAResults_DIR_" + sAnalysisCode + "_" + sTreatmentType, pFWorkspace, sPrefix)
        Catch ex As Exception
            MsgBox("Error trying to get SA Results table name. " + ex.Message)
            Exit Sub
        End Try

        Try
            ' create table
            PrepDIRResultsOutTable(sSAResultsTableName, pFWorkspace)
        Catch ex As Exception
            MsgBox("Error preparing DIR SA results table. " + ex.Message)
            Exit Sub
        End Try

        ' populate table
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  GLPKSolved (boolean/binary)
        '  Perc_Gap (double)
        '  MAxSolTime(integer)
        '  TimeUsed(double)
        '  HabitatZmax(double)
        Try
            pTable = pFWorkspace.OpenTable(sSAResultsTableName)
            n = 0

            For n = 0 To lSA_Results_DIR.Count - 1
                '  SinkEID (integer)
                '  Treatment (string)
                '  Budget (double)
                '  GLPKSolved (boolean/binary)
                '  Perc_Gap (double)
                '  MAxSolTime(integer)
                '  TimeUsed(double)
                '  HabitatZmax(double)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(1) = lSA_Results_DIR(n).SinkEID

                If sTreatmentType = "A" Then
                    If n = 0 Then
                        pRowBuffer.Value(2) = lSA_Results_DIR(n).Treatment_Name + "_A"
                    Else
                        pRowBuffer.Value(2) = lSA_Results_DIR(n).Treatment_Name + "_A" + Convert.ToString(n - 1)
                    End If
                ElseIf sTreatmentType = "P" Then
                    pRowBuffer.Value(2) = lSA_Results_DIR(n).Treatment_Name + "_P" + Convert.ToString(n)
                Else
                    pRowBuffer.Value(2) = lSA_Results_DIR(n).Treatment_Name
                End If

                pRowBuffer.Value(3) = lSA_Results_DIR(n).Budget
                pRowBuffer.Value(4) = Convert.ToInt16(lSA_Results_DIR(n).GLPK_Solved)
                pRowBuffer.Value(5) = lSA_Results_DIR(n).Perc_Gap
                pRowBuffer.Value(6) = lSA_Results_DIR(n).MaxSolTime
                pRowBuffer.Value(7) = lSA_Results_DIR(n).TimeUsed
                pRowBuffer.Value(8) = lSA_Results_DIR(n).Habitat_ZMax

                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            Next

        Catch ex As Exception
            MsgBox("Error writing to GRB Results out table. " + ex.Message)
            Exit Sub
        End Try
    End Sub
    Private Sub Export_SA_Results_UNDIR(ByRef pFWorkspace As IFeatureWorkspace, _
                                      ByRef sAnalysisCode As String, _
                                      ByVal sPrefix As String, _
                                      ByVal sTreatmentType As String)

        ' exports table from stored global variable 
        ' containing 'results' table rows
        ' created in GLPKShellCall for thesis Sensitivity analysis
        ' -- directed model

        Dim pRowBuffer As IRowBuffer
        Dim pTable As ITable
        Dim pCursor As ICursor
        Dim sSAResultsTableName As String
        Dim n As Integer

        Dim lSA_Results_UNDIR As List(Of UNDIR_OptResultsObject) = New List(Of UNDIR_OptResultsObject) ' FOR THESIS SENSITIVTY ANALYSIS


        If sTreatmentType = "A" Then
            lSA_Results_UNDIR = m_lSA_A_Results_UNDIR
        ElseIf sTreatmentType = "P" Then
            lSA_Results_UNDIR = m_lSA_P_Results_UNDIR
        End If

        ' ---------------
        ' Create and write to 'results' table for SA Analysis
        ' results table
        Try
            sSAResultsTableName = TableName("SAResults_UNDIR_" + sAnalysisCode + "_" + sTreatmentType, pFWorkspace, sPrefix)
        Catch ex As Exception
            MsgBox("Error trying to get UNDIR SA Results table name. " + ex.Message)
            Exit Sub
        End Try

        Try
            ' create table
            PrepUNDIRResultsOutTable(sSAResultsTableName, pFWorkspace)
        Catch ex As Exception
            MsgBox("Error preparing UNDIR SA results table. " + ex.Message)
            Exit Sub
        End Try

        ' populate table
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  GLPKSolved (boolean/binary)
        '  Perc_Gap (double)
        '  MAxSolTime(integer)
        '  TimeUsed(double)
        '  HabitatZmax(double)
        Try
            pTable = pFWorkspace.OpenTable(sSAResultsTableName)
            n = 0

            For n = 0 To lSA_Results_UNDIR.Count - 1
                '  SinkEID (integer)
                '  Treatment (string)
                '  Budget (double)
                '  GLPKSolved (boolean/binary)
                '  Perc_Gap (double)
                '  MAxSolTime(integer)
                '  TimeUsed(double)
                '  HabitatZmax(double)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(1) = lSA_Results_UNDIR(n).SinkEID

                If sTreatmentType = "A" Then
                    If n = 0 Then
                        pRowBuffer.Value(2) = lSA_Results_UNDIR(n).Treatment_Name + "_A"
                    Else
                        pRowBuffer.Value(2) = lSA_Results_UNDIR(n).Treatment_Name + "_A" + Convert.ToString(n - 1)
                    End If
                ElseIf sTreatmentType = "P" Then
                    pRowBuffer.Value(2) = lSA_Results_UNDIR(n).Treatment_Name + "_P" + Convert.ToString(n)
                Else
                    pRowBuffer.Value(2) = lSA_Results_UNDIR(n).Treatment_Name
                End If

                pRowBuffer.Value(3) = lSA_Results_UNDIR(n).Budget
                pRowBuffer.Value(4) = Convert.ToInt16(lSA_Results_UNDIR(n).GLPK_Solved)
                pRowBuffer.Value(5) = lSA_Results_UNDIR(n).Perc_Gap
                pRowBuffer.Value(6) = lSA_Results_UNDIR(n).MaxSolTime
                pRowBuffer.Value(7) = lSA_Results_UNDIR(n).TimeUsed
                pRowBuffer.Value(8) = lSA_Results_UNDIR(n).Habitat_ZMax
                pRowBuffer.Value(9) = lSA_Results_UNDIR(n).CentralBarrierEID

                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            Next

        Catch ex As Exception
            MsgBox("Error writing to SA UNDIR Results out table. " + ex.Message)
            Exit Sub
        End Try
    End Sub
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
    Private Class FindBarrierEIDPredicate
        ' this class should help return a double-check 
        ' list object of Statistics where the 
        ' and the sink EID, barrier ID, and layer matches.  
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        Private _bEID As Integer
        Private _passability As Double

        Public Sub New(ByVal bEID As Integer, ByVal passability As Double)
            Me._bEID = bEID
            Me._passability = passability
        End Sub

        ' this is a double-check and should 
        ' return true if there's a match between
        ' BOTH the layer and EID of the sink/barrier
        Public Function CompareEIDandLayer(ByVal obj As StatisticsObject_2) As Boolean
            Return (_bEID = obj.bEID)
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
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
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

                        .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineClassField" + j.ToString))
                        .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineQuanField" + j.ToString))
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
    Public Sub calculateGLPKStatistics(ByRef lGLPKStatisticsList As List(Of GLPKStatisticsObject), _
                                       ByVal lAllLayersFieldsGLPK As List(Of LineLayerToAdd), _
                                       ByRef ID As Integer, _
                                       ByVal iOrderLoop As Integer)

        ' **************************************************************************************
        ' Subroutine:  Calculate DCI Statistics 
        ' Created By:  Greig Oldford
        ' Update Date: October 5, 2010
        ' Purpose:     
        '              1) calculate habitat area and length ignoring habitat classes 
        '              2) combine multiple network-participating line layers into one
        '                 habitat measure (intent to introduce weighting options and 
        '                 inclusion of non-network lines which requires additions to intersect sub)
        '              4) update statistics object and send back to onclick

        ' ##################### 2020 Commented out the code below ########################
        '                       do not delete

        'Dim pMxDoc As IMxDocument
        'Dim pEnumLayer As IEnumLayer
        'Dim pFeatureLayer As IFeatureLayer
        'Dim pFeatureSelection As IFeatureSelection
        'Dim pFeatureCursor As IFeatureCursor
        'Dim pFeature As IFeature
        'Dim dTotalArea As Double ' alternative to 2D Matrix when no classes are given
        'Dim pUID As New UID
        '' Get the pUID of the SelectByLayer command
        ''pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"

        'Dim pMxDocument As IMxDocument
        'Dim pMap As IMap

        'Dim j, k, m As Integer
        'Dim iFieldVal As Integer  ' The field index

        'Dim pFields As IFields
        'Dim pSelectionSet As ISelectionSet

        '' K REPRESENTS NUMBER OF POSSIBLE HABITAT CLASSES
        ''  rows, columns.  ROWS SHOULD BE SET BY NUMBER OF SUMMARY FIELDS
        '' cannot be redimension preserved later

        'Dim lStatsMatrix As New List(Of StatisticsObject)
        'Dim pStatisticsObject As New StatisticsObject(Nothing, Nothing)

        'Dim pCursor As ICursor
        'Dim vTemp As Object

        'Dim pDoc As IDocument = My.ArcMap.Application.Document
        'pMxDoc = CType(pDoc, IMxDocument)
        'pMxDocument = CType(pDoc, IMxDocument)
        'pMap = pMxDocument.FocusMap

        '' ================== READ EXTENSION SETTINGS =================
        'Dim sDirection, sDirection2 As String

        'Dim bDBF As Boolean = False         ' Include DBF output default 'no'
        'Dim pLLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
        'Dim pPLayersFields As List(Of LayerToAdd) = New List(Of LayerToAdd)
        'Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        'Dim iLinesCount As Integer = 0      ' number of lines layers currently using
        'Dim HabLayerObj As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property
        '' object to hold stats to add to list. 
        'Dim pGLPKStatisticsObject As New GLPKStatisticsObject(Nothing, Nothing)

        'If m_FiPEx__1.m_bLoaded = True Then

        '    sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))

        '    'Make the direction more readable for dockable window output
        '    If sDirection = "up" Then
        '        sDirection2 = "upstream"
        '    Else
        '        sDirection2 = "downstream"
        '    End If

        '    ' Populate a list of the layers using and habitat summary fields.
        '    ' match any of the polygon layers saved in stream to those in listboxes 
        '    iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
        '    If iPolysCount > 0 Then
        '        For k = 0 To iPolysCount - 1
        '            'sPolyLayer = m_FiPEX__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
        '            HabLayerObj = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
        '            With HabLayerObj
        '                .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer
        '                .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
        '                .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
        '                .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
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
        '                System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for lacustrine layer. Please choose a field in the options menu.", "Parameter Missing")
        '                Exit Sub
        '            End If
        '        Next
        '    End If

        '    iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
        '    Dim HabLayerObj2 As New LayerToAdd(Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

        '    ' match any of the line layers saved in stream to those in listboxes
        '    If iLinesCount > 0 Then
        '        For j = 0 To iLinesCount - 1
        '            'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
        '            HabLayerObj2 = New LayerToAdd(Nothing, Nothing, Nothing, Nothing)
        '            With HabLayerObj2
        '                '.Layer = sLineLayer
        '                .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))
        '                .ClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineClassField" + j.ToString))
        '                .QuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineQuanField" + j.ToString))
        '                .UnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineUnitField" + j.ToString))
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

        '' ========== End Read Extension settings ===============

        'pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        'pEnumLayer = pMap.Layers(pUID, True)
        'pEnumLayer.Reset()

        '' Look at the next layer in the list
        'pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)

        'Dim iLoopCount As Integer = 0
        'Dim dTempQuan As Double
        'dTotalArea = 0

        'Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
        '    If pFeatureLayer.Valid = True Then ' or there will be an empty object ref
        '        If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Or _
        '        pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolygon Then ' or there will be an empty object ref

        '            pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
        '            pSelectionSet = pFeatureSelection.SelectionSet

        '            ' get the fields from the featureclass
        '            pFields = pFeatureLayer.FeatureClass.Fields
        '            j = 0

        '            For j = 0 To lAllLayersFieldsGLPK.Count - 1
        '                If lAllLayersFieldsGLPK(j).Layer = pFeatureLayer.Name Then

        '                    If pFeatureSelection.SelectionSet.Count <> 0 Then

        '                        pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)
        '                        pFeatureCursor = CType(pCursor, IFeatureCursor)
        '                        pFeature = pFeatureCursor.NextFeature

        '                        ' Get the summary field and add the value to the
        '                        ' total for habitat area.
        '                        ' ** ==> Multiple fields could be added here in a 'for' loop.

        '                        iFieldVal = pFeatureCursor.FindField(lAllLayersFieldsGLPK(j).QuanField)

        '                        ' For each selected feature
        '                        m = 1
        '                        Do While Not pFeature Is Nothing
        '                            Try
        '                                vTemp = pFeature.Value(iFieldVal)
        '                            Catch ex As Exception
        '                                vTemp = 0
        '                            End Try

        '                            Try
        '                                dTempQuan = Convert.ToDouble(pFeature.Value(iFieldVal))
        '                            Catch ex As Exception
        '                                dTempQuan = 0
        '                                MsgBox("The Habitat Quantity found in the " & pFeatureLayer.Name & " was not convertible" _
        '                                    & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
        '                            End Try
        '                            ' Insert into the corresponding column of the second
        '                            ' row the updated habitat area measurement.
        '                            dTotalArea = dTotalArea + dTempQuan
        '                            pFeature = pFeatureCursor.NextFeature
        '                        Loop     ' selected feature
        '                    End If ' there are selected features

        '                    ' increment the loop counter for
        '                    iLoopCount = iLoopCount + 1

        '                End If     ' feature layer matches hab class layer
        '            Next           ' habitat layer
        '        End If ' featurelayer is valid
        '    End If
        '    pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        'Loop ' next map layer

        'With pGLPKStatisticsObject
        '    ' If iOrderLoop = 0 Then
        '    ' .Barrier = "Sink"
        '    ' Else
        '    .Barrier = ID
        '    'End If
        '    '.BarrierPerm = dBarrierPerm
        '    '.BarrierYN = sNaturalYN
        '    .Quantity = dTotalArea
        'End With

        'lGLPKStatisticsList.Add(pGLPKStatisticsObject)
    End Sub
    Public Sub calculateDCIStatistics(ByRef lDCIStatsList As List(Of DCIStatisticsObject), _
                                      ByVal lLLayersFieldsDCI As List(Of LineLayerToAdd), _
                                      ByVal lPLayersFieldsDCI As List(Of PolyLayerToAdd), _
                                      ByRef ID As String, _
                                      ByVal dBarrierPerm As Double, _
                                      ByVal sNaturalYN As String, _
                                      ByVal iOrderLoop As Integer)

        ' **************************************************************************************
        ' Subroutine:  Calculate DCI Statistics 
        ' Created By:  Greig Oldford
        ' Update Date: October 5, 2010
        ' Purpose:     
        '              1) calculate habitat area and length ignoring habitat classes 
        '              2) combine multiple network-participating line layers into one
        '                 habitat measure (intent to introduce weighting options and 
        '                 inclusion of non-network lines which requires additions to intersect sub)
        '              4) update statistics object and send back to onclick
        '
        ' 2020 - note this doesn't currently make use of area and polygons, just lines

        Dim pMxDoc As IMxDocument
        Dim pEnumLayer As IEnumLayer
        Dim pFeatureLayer As IFeatureLayer
        Dim pFeatureSelection As IFeatureSelection
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeature As IFeature
        Dim pUID As New UID

        Dim pMxDocument As IMxDocument
        Dim pMap As IMap

        Dim j, k, m As Integer
        Dim iHabField, iLengthField As Integer  ' The field index

        Dim pFields As IFields
        Dim pSelectionSet As ISelectionSet

        Dim lStatsMatrix As New List(Of HabStatisticsObject)
        Dim pStatisticsObject As New HabStatisticsObject(Nothing, Nothing)

        Dim pCursor As ICursor
        Dim vTemp As Object

        ' object to hold stats to add to list. 
        Dim pDCIStatsObject As New DCIStatisticsObject(Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim dTempQuan, dTempLength As Double

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMxDocument = CType(pDoc, IMxDocument)
        pMap = pMxDocument.FocusMap

        ' 2020 - removed read of extension settings - unnecessary
        ' ================== READ EXTENSION SETTINGS =================

        ' Dim sDirection, sDirection2 As String

        'Dim bDBF As Boolean = False         ' Include DBF output default 'no'
        'Dim pLLayersFields As List(Of LineLayerToAdd) = New List(Of LineLayerToAdd)
        'Dim pPLayersFields As List(Of PolyLayerToAdd) = New List(Of PolyLayerToAdd)
        'Dim iPolysCount As Integer = 0      ' number of polygon layers currently using
        'Dim iLinesCount As Integer = 0      ' number of lines layers currently using

        'If m_FiPEx__1.m_bLoaded = True Then

        '    sDirection = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("direction"))
        '    bDBF = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDBF"))

        'Make the direction more readable for dockable window output
        'If sDirection = "up" Then
        '    sDirection2 = "upstream"
        'Else
        '    sDirection2 = "downstream"
        'End If

        ' Populate a list of the layers using and habitat summary fields.
        ' match any of the polygon layers saved in stream to those in listboxes 
        'iPolysCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numPolys"))
        'Dim HabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
        'If iPolysCount > 0 Then
        '    For k = 0 To iPolysCount - 1
        '        'sPolyLayer = m_FiPEX__1.pPropset.GetProperty("IncPoly" + k.ToString) ' get poly layer
        '        HabLayerObj = New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
        '        With HabLayerObj
        '            .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncPoly" + k.ToString)) ' get poly layer

        '            .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyClassField" + k.ToString))
        '            .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyQuanField" + k.ToString))
        '            .HabUnitField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("PolyUnitField" + k.ToString))
        '        End With

        '        ' Load that object into the list
        '        pPLayersFields.Add(HabLayerObj)  'what are the brackets about - this could be aproblem!!
        '    Next
        'End If

        '' Need to be sure that quantity field has been assigned for each
        '' layer using. 
        'Dim iCount1 As Integer = pPLayersFields.Count

        'If iCount1 > 0 Then
        '    For m = 0 To iCount1 - 1
        '        If pPLayersFields.Item(m).HabQuanField = "Not set" Then
        '            System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for polygon layer. Please choose a field in the options menu.", "Parameter Missing")
        '            Exit Sub
        '        End If
        '    Next
        'End If

        'iLinesCount = Convert.ToInt32(m_FiPEx__1.pPropset.GetProperty("numLines"))
        'Dim HabLayerObj2 As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) ' layer to hold parameters to send to property

        '' match any of the line layers saved in stream to those in listboxes
        'If iLinesCount > 0 Then
        '    For j = 0 To iLinesCount - 1
        '        'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
        '        HabLayerObj2 = New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        '        With HabLayerObj2
        '            '.Layer = sLineLayer
        '            .Layer = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("IncLine" + j.ToString))

        '            .LengthField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthField" + j.ToString))
        '            .LengthUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineLengthUnits" + j.ToString))

        '            .HabClsField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabClassField" + j.ToString))
        '            .HabQuanField = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabQuanField" + j.ToString))
        '            .HabUnits = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("LineHabUnits" + j.ToString))

        '        End With
        '        ' add to the module level list
        '        pLLayersFields.Add(HabLayerObj2)
        '    Next
        'End If

        '' Need to be sure that quantity field has been assigned for each
        '' layer using. 
        'iCount1 = pLLayersFields.Count
        'If iCount1 > 0 Then
        '    For m = 0 To iCount1 - 1
        '        If pLLayersFields.Item(m).HabQuanField = "Not set" Then
        '            System.Windows.Forms.MessageBox.Show("No habitat quantity parameter set for line layer. Please choose a field in the options menu.", "Parameter Missing")
        '            Exit Sub
        '        End If
        '        If pLLayersFields.Item(m).LengthField = "Not set" Then
        '            System.Windows.Forms.MessageBox.Show("No length field set for line layer. Please choose a field in the options menu.", "Parameter Missing")
        '            Exit Sub
        '        End If
        '    Next
        'End If
        'Else
        'System.Windows.Forms.MessageBox.Show("Cannot read extension settings.", "Calculate Stats Error")
        'Exit Sub
        'End If

        ' ========== End Read Extension settings ===============

        pUID.Value = "{E156D7E5-22AF-11D3-9F99-00C04F6BC78E}"
        pEnumLayer = pMap.Layers(pUID, True)
        pEnumLayer.Reset()

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)

        Dim iLoopCount As Integer = 0
        Dim dTotalHab As Double = 0
        Dim dTotalLength As Double = 0

        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                ' 2020 

                If pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryLine Or _
                pFeatureLayer.FeatureClass.ShapeType = esriGeometryType.esriGeometryPolyline Then ' or there will be an empty object ref

                    pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                    pSelectionSet = pFeatureSelection.SelectionSet

                    ' get the fields from the featureclass
                    pFields = pFeatureLayer.FeatureClass.Fields
                    j = 0

                    For j = 0 To lLLayersFieldsDCI.Count - 1
                        If lLLayersFieldsDCI(j).Layer = pFeatureLayer.Name Then

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature


                                ' 2020 - only get habitat from line if the user has chosen line
                                'IF...

                                ' For each selected feature
                                ' Get the habitat field and add the value to the
                                ' total for habitat quantity.

                                Try
                                    iHabField = pFeatureCursor.FindField(lLLayersFieldsDCI(j).HabQuanField)
                                Catch ex As Exception
                                    MsgBox("Error finding the field in line layer for habitat. FIPEX code 986")
                                    Exit Sub

                                End Try

                                Try
                                    iLengthField = pFeatureCursor.FindField(lLLayersFieldsDCI(j).LengthField)
                                Catch ex As Exception
                                    MsgBox("Error finding the field in line layer for length. FIPEX code 987")
                                    Exit Sub

                                End Try

                                m = 1
                                Do While Not pFeature Is Nothing
                                    Try
                                        vTemp = pFeature.Value(iHabField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from habitat attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 984.")
                                        vTemp = 0
                                    End Try

                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iHabField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The habitat quantity found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                    End Try
                                    dTotalHab = dTotalHab + dTempQuan


                                    ' 2020 get distance / length field and quantity
                                    Try
                                        vTemp = pFeature.Value(iLengthField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from length attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 984.")
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iLengthField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The length found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                    End Try
                                    dTotalLength = dTotalLength + dTempQuan

                                    pFeature = pFeatureCursor.NextFeature
                                Loop     ' selected feature
                            End If ' there are selected features

                            ' increment the loop counter for
                            iLoopCount = iLoopCount + 1

                        End If     ' feature layer matches hab class layer
                    Next           ' habitat layer

                    ' 2020 - if user has selected poly for hab then....



                End If ' featurelayer is valid
            End If
            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Loop ' next map layer

        With pDCIStatsObject
            If iOrderLoop = 0 Then
                .Barrier = "Sink"
            Else
                .Barrier = ID
            End If
            .BarrierPerm = dBarrierPerm
            .BarrierYN = sNaturalYN
            .Length = dTotalLength
            .HabQuantity = dTotalHab
        End With

        lDCIStatsList.Add(pDCIStatsObject)
    End Sub
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
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
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

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "Habitat"
                                        .TotalImmedPath = sHabTypeKeyword
                                        .UniqueClass = CStr(lHabStatsMatrix(k).UniqueHabClass)
                                        .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                        .Quantity = lHabStatsMatrix(k).HabQuantity
                                        .Unit = sUnit
                                    End With
                                    lHabStatsList.Add(pHabStatsObject_2)

                                Next


                            Else ' If there are no statistics

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection
                                    .LengthOrHabitat = "Habitat"
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

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "Habitat"
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

                                    pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                                    With pHabStatsObject_2
                                        .Layer = pFeatureLayer.Name
                                        .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                        .bID = ID
                                        .bEID = iEID
                                        .bType = sType
                                        .Sink = f_sOutID
                                        .SinkEID = f_siOutEID
                                        .Direction = sDirection2
                                        .LengthOrHabitat = "Habitat"
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

                                pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                                With pHabStatsObject_2
                                    .Layer = pFeatureLayer.Name
                                    .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                    .bID = ID
                                    .bEID = iEID
                                    .bType = sType
                                    .Sink = f_sOutID
                                    .SinkEID = f_siOutEID
                                    .Direction = sDirection
                                    .LengthOrHabitat = "Habitat"
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

                            pHabStatsObject_2 = New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
                            With pHabStatsObject_2
                                .Layer = pFeatureLayer.Name
                                .LayerID = pFeatureLayer.FeatureClass.FeatureClassID
                                .bID = ID
                                .bEID = iEID
                                .bType = sType
                                .Sink = f_sOutID
                                .SinkEID = f_siOutEID
                                .Direction = sDirection2
                                .LengthOrHabitat = "Habitat"
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
    Private Sub UpdateSummaryTable(ByRef lRefinedHabStatsList As List(Of StatisticsObject_2),
                                   ByRef iHabRowCount As Integer,
                                   ByRef pResultsForm3 As FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3,
                                   ByRef iMaxRowIndex As Integer,
                                   ByRef iBarrIndex As Integer,
                                   ByRef bColorSwitcher As Boolean,
                                   ByRef iSinkRowCount As Integer,
                                   ByRef iBarrRowCount As Integer,
                                   ByVal bSinkVisit As Boolean)

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
                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style
                pDataGridViewCellStyle.BackColor = Color.Lavender
            Else
                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style
                pDataGridViewCellStyle.BackColor = Color.PowderBlue
            End If
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(11).Style = pDataGridViewCellStyle
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(12).Style = pDataGridViewCellStyle

            ' if the row count of the habitat metrics exceeds the 
            ' statistics metrics then the colors of cells below the metrics
            ' will also have to be changed
            If bTrigger = True Then
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(4).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Style = pDataGridViewCellStyle
            End If
            'End If

            If m = 0 Then
                sLayer = lRefinedHabStatsList(m).Layer
                sDirection = lRefinedHabStatsList(m).Direction
                sTraceType = lRefinedHabStatsList(m).TotalImmedPath
                sUnit = lRefinedHabStatsList(m).Unit

                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Value = sLayer
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Value = sDirection
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Value = sTraceType
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(12).Value = sUnit
            End If

            sClass = lRefinedHabStatsList(m).UniqueClass
            dQuantity = Math.Round(lRefinedHabStatsList(m).Quantity, 2)
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Value = sClass
            pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(11).Value = dQuantity

            dTotalHab = dTotalHab + lRefinedHabStatsList(m).Quantity
            iHabRowCount += 1

            ' Add the total field if necessary
            If m = lRefinedHabStatsList.Count - 1 And
                lRefinedHabStatsList.Count <> 1 Then

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

                pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(11).Style
                pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily,
                                                       pResultsForm3.DataGridView1.Font.Size,
                                                       FontStyle.Bold)
                'pDataGridViewCellStyle.BackColor = Color.SlateGray
                If bColorSwitcher = True Then
                    pDataGridViewCellStyle.BackColor = Color.Lavender
                Else
                    pDataGridViewCellStyle.BackColor = Color.PowderBlue
                End If
                'pDataGridViewCellStyle.ForeColor = Color.White
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(12).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(11).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(9).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(8).Style = pDataGridViewCellStyle
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(7).Style = pDataGridViewCellStyle

                ' If the habitat row exceeds maximum metric row then 
                ' color the cells below the metric row appropriately
                If bTrigger = True And bColorSwitcher = True Then
                    pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style
                    pDataGridViewCellStyle.BackColor = Color.Lavender
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(4).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Style = pDataGridViewCellStyle
                ElseIf bTrigger = True And bColorSwitcher = False Then
                    pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style
                    pDataGridViewCellStyle.BackColor = Color.PowderBlue
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(3).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(4).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(5).Style = pDataGridViewCellStyle
                    pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(6).Style = pDataGridViewCellStyle
                End If
                ' for border adjustment
                'dataGridViewAdvancedBorderStyleInput = pResultsForm3.DataGridView1.Rows(iHabRowIndex).DefaultCellStyle(8).
                'dataGridViewAdvancedBorderStylePlaceHolder = dataGridViewAdvancedBorderStyleInput
                'dataGridViewAdvancedBorderStylePlaceHolder.Top = System.Windows.Forms.DataGridViewAdvancedCellBorderStyle.InsetDouble

                'pResultsForm3.DataGridView1.Rows(iHabRowIndex).Cells(8).AdjustCellBorderStyle(dataGridViewAdvancedBorderStyleInput, dataGridViewAdvancedBorderStylePlaceHolder, False, False, False, False)
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(10).Value = "(Total)"
                pResultsForm3.DataGridView1.Rows(iThisRowIndex).Cells(11).Value = Math.Round(dTotalHab, 2)

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

    Public Function FileInUse(ByVal sFile As String) As Boolean
        If System.IO.File.Exists(sFile) Then
            Try
                Dim F As Short = FreeFile()
                FileOpen(F, sFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(F)
            Catch
                Return True
            End Try
        End If
    End Function
    Private Sub GLPKShellCall(ByVal sGLPKHabitatTableName As String, _
                              ByVal sGLPKOptionsTableName As String, _
                                      ByVal sGLPKConnectTabName As String, _
                                     ByVal iMaxOptions As Integer, _
                                      ByVal iSinkEID As Integer, _
                                      ByRef pFWorkspace As IFeatureWorkspace, _
                                      ByRef iProgress1 As Integer, _
                                      ByRef sAnalysisCode As String, _
                                      ByVal sAnalysisType As String, _
                                      ByVal bZeroBudgetOverride As Boolean)
        '    ' ====================================================================
        '    ' SubRoutine:    GLPK GNUwin Shell Call
        '    ' Author:        Greig Oldford
        '    ' Created For:   Thesis
        '    '
        '    ' Description:
        '    '             Runs the GLPK Optimization Model on the network and other CSV files 
        '    '             First overwrites a model data file containing parameters 
        '    '             Then overwrites model file with correct model paths (customizable by user)
        '    '             Runs iteratively for a series of budget amounts
        '    '             Optionally runs the 'Directed' or 'Undirected' model
        '    '             Optionally uses the free GLPK optimisation solver (called from Shell)
        '    '             Or uses the proprietary Gurobi Optimisation software (called from Gurobi API)
        '    '             BUDGET     - budget for this analysis 
        '    '             TIME LIMIT - second time limit to GLPK run
        '    '             MAX OPTIONS - max number of options at any node 
        '    '             FIRST NODE - network sink
        '    '
        '    ' Logic:
        '    '       exports already created FC GDB Tables to CSV's in the GLPK Model folder
        '    '       iteratively runs the GLPK model for incremental budget amounts.  
        '    '       writes results to output tables to Geodatabase set by user in FiPEX Options



        '    ' Read FIPEX Options
        '    Dim bGLPKTables As Boolean = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bGLPKTables"))
        '    Dim sGLPKModelDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGLPKModelDir"))
        '    Dim sGnuWinDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGnuWinDir"))
        '    Dim sGLPKTreatment As String '= Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sGLPKTreatment")) 'to name the treatment (optional)
        '    Dim dBudgetIncrement As Double '= Convert.ToDouble(m_FiPEx__1.pPropset.GetProperty("dGLPKBudgetIncrement"))
        '    Dim dMinBudget As Double '= Convert.ToDouble(m_FiPEx__1.pPropset.GetProperty("dGLPKMinBudget"))
        '    Dim dMaxBudget As Double '= Convert.ToDouble(m_FiPEx__1.pPropset.GetProperty("dGLPKMaxBudget"))
        '    Dim iGLPKTimeLimit As Integer '= Convert.ToInt16(m_FiPEx__1.pPropset.GetProperty("iGLPKTimeLimit"))

        '    Dim sPrefix As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("TabPrefix"))

        '    ' THIS NEEDS TO BE ADDED TO DOC STREAM IN ALL READ/WRITE LOCATIONS
        '    ' MAY CORRUPT EXISTING DOCUMENTS
        '    Dim bUndirected As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bUndirected"))
        '    Dim bDirected As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bDirected"))

        '    If sAnalysisType = "DIR" Then
        '        bDirected = True
        '        bUndirected = False
        '    ElseIf sAnalysisType = "UNDIR" Then
        '        bDirected = False
        '        bUndirected = True
        '    Else
        '        MsgBox("Error getting analysis type -- DIR or UNDIR!  Exiting GLPKShellCall.")
        '        Exit Sub
        '    End If

        '    Dim bGurobiPickup As Boolean '= Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("bGurobiPickup"))
        '    Dim iGRBTimeLimit As Integer '= Convert.ToInt16(m_FiPEx__1.pPropset.GetProperty("iGRBTimeLimit"))
        '    Dim bGRBTimeLimit As Boolean = True ' this is not FULLY INTEGRATED YET (false = no time limit)

        '    'TEMP HARDCODE!
        '    bGLPKTables = False
        '    sGLPKModelDir = "c:\GunnsModel"
        '    sGnuWinDir = "C:\gnuwin32"
        '    'sGLPKTreatment = "Feb4_DeleteMe_" + sAnalysisType + "_test"
        '    sGLPKTreatment = "Apr28_MER_ANoSpn_" + sAnalysisType
        '    dBudgetIncrement = 240
        '    dMinBudget = 15
        '    dMaxBudget = 10000

        '    ' Specifies the acceptable gap tolerance for gurobi
        '    Dim dMIPGap As Double = 0.02
        '    Dim bMIPGap As Boolean = True

        '    ' Budget amounts are overridden based on bZeroBudgetOverride
        '    ' - this is used by runRandomization (SA) to force a zero budget
        '    ' run to get a 'do nothing' ZMAX
        '    If bZeroBudgetOverride = True Then
        '        ' the budget for 'do nothing' run must be one
        '        ' dollar per option... zero doesn't seem to work
        '        dMinBudget = 10
        '        dMaxBudget = 10
        '    End If

        '    iGLPKTimeLimit = 20
        '    bGurobiPickup = True
        '    iGRBTimeLimit = 100

        '    Dim bSA As Boolean = False ' sensitivity analysis?  
        '    Dim sTreatmentType As String = "N"

        '    ' Keep track of treatment type if this 
        '    ' is the sensitivity analysis -- determines
        '    ' which global variable the 'results get thrown into
        '    If bSA = True Then
        '        If (sAnalysisCode.Substring(sAnalysisCode.Length - 2)).StartsWith("_A") = True _
        '                   Or (sAnalysisCode.Substring(sAnalysisCode.Length - 3)).StartsWith("_A") = True _
        '                   Or (sAnalysisCode.Substring(sAnalysisCode.Length - 4)).StartsWith("_A") = True _
        '                   Then
        '            sTreatmentType = "A"

        '        ElseIf (sAnalysisCode.Substring(sAnalysisCode.Length - 2)).StartsWith("_P") = True _
        '            Or (sAnalysisCode.Substring(sAnalysisCode.Length - 3)).StartsWith("_P") = True _
        '            Or (sAnalysisCode.Substring(sAnalysisCode.Length - 4)).StartsWith("_P") = True _
        '            Or (sAnalysisCode.Substring(sAnalysisCode.Length - 5)).StartsWith("_P") = True _
        '            Then

        '            sTreatmentType = "P"

        '        End If
        '    End If


        '    Dim dMArea As Double ' maximum area possible (bounds the model), for undirected param. file only.  
        '    ' equal to total hab area above sink

        '    Dim dBudget As Double
        '    Dim bGLPKSolved As Boolean = False
        '    Dim bGRBSolved As Boolean = False
        '    Dim dPercGap As Double
        '    Dim dTimeUsed As Double
        '    Dim dGRBTimeUsed As Double
        '    Dim iSolutionFoundRow As Integer = 0
        '    Dim dMaxHabitat As Double = 0

        '    Dim sSplitTimeUsed As String
        '    Dim sSplitTimeUsedFirstHalf As String
        '    Dim sSplitTimeUsedFirstHalf2 As String
        '    Dim pTable As ITable
        '    'Dim pCursor As ICursor
        '    Dim i As Integer
        '    Dim pTxtFactory As IWorkspaceFactory = New TextFileWorkspaceFactory

        '    Dim pFWorkspaceOut As IFeatureWorkspace
        '    Dim pWorkspaceOut As IWorkspace
        '    Dim sw As StreamWriter
        '    ' Get output workspace
        '    pWorkspaceOut = pTxtFactory.OpenFromFile(sGLPKModelDir, My.ArcMap.Application.hWnd)
        '    pFWorkspaceOut = CType(pWorkspaceOut, IFeatureWorkspace)

        '    Dim sOutputFile(0) As String
        '    Dim sModelFile(0) As String
        '    Dim sPreviousOutputLine As String
        '    Dim sPreviousOutputLineSplit(1) As String
        '    Dim pDIR_OptResultsObject As New DIR_OptResultsObject(Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing, _
        '                                                          Nothing)
        '    Dim pUNDIR_ResultsObject As New UNDIR_OptResultsObject(Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing, _
        '                                                           Nothing)
        '    Dim n As Integer
        '    Dim pGLPKDecisionsObject As New GLPKDecisionsObject(Nothing, _
        '                                                        Nothing, _
        '                                                        Nothing, _
        '                                                        Nothing)
        '    Dim lGLPKDecisionsObject As New List(Of GLPKDecisionsObject) 'for directed model
        '    Dim lGLPKDecisionsObject2 As New List(Of GLPKDecisionsObject) ' for undirected model
        '    Dim lGRBDecisionsObject As New List(Of GLPKDecisionsObject) 'for directed model
        '    Dim lGRBDecisionsObject2 As New List(Of GLPKDecisionsObject) ' for undirected model
        '    Dim sDecisionOutputSplit(2) As String
        '    Dim iDecisionEID, iDecisionNum As Integer
        '    ' Check that user has write permissions in output workspace
        '    'Dim pc As New CheckPerm
        '    'pc.Permission = "Modify"
        '    'If Not pc.CheckPerm(sDCIModelDir) Then
        '    '    MsgBox("You do not have the necessary permission to create, delete, or modify files in the DCI Model installation directory." + _
        '    '    " Please change directory, attain necessary permissions, or uncheck DCI Output checkbox in 'Advanced Tab' of Options Menu.")
        '    '    Exit Sub
        '    'End If

        '    If m_bCancel = True Then
        '        backgroundworker1.CancelAsync()
        '        backgroundworker1.Dispose()
        '        Exit Sub
        '    End If
        '    If iProgress1 < 70 Then
        '        iProgress1 = iProgress1 + 1
        '    End If
        '    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                                 "Checking User permissions in output directory.")


        '    ' Check that the user currently has file permissions to write to 
        '    ' this directory
        '    Dim bPermissionCheck
        '    bPermissionCheck = FileWriteDeleteCheck(sGLPKModelDir)
        '    If bPermissionCheck = False Then
        '        MsgBox("File / folder permission check: " & Str(bPermissionCheck))
        '        MsgBox("It appears you do not have write permission to the GLPK Model Directory.  Write permission to this directory is needed in order to run DCI Analysis.")
        '        Exit Sub
        '    End If

        '    Dim pExportOp As IExportOperation = New ExportOperation
        '    Dim pDataSetIn As IDataset
        '    Dim pDataSetOut As IDataset
        '    Dim pDSNameIn As IDatasetName
        '    Dim pDSNameOut As IDatasetName
        '    Dim FileToDelete As String

        '    ' ========================
        '    ' start logfile
        '    ' Delete the table if it already exists
        '    Dim log As StreamWriter
        '    FileToDelete = sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt"
        '    If System.IO.File.Exists(FileToDelete) = True Then
        '        System.IO.File.Delete(FileToDelete)
        '    End If
        '    log = File.CreateText(sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt")
        '    log.WriteLine("Begin new analysis")
        '    log.Close()
        '    ' ========================


        '    ' Get the output dataset name ready.
        '    pDataSetOut = CType(pWorkspaceOut, IDataset)

        '    ' Take DBF habitat tables in geodatabase workspace
        '    ' and print it to the new text/csv file

        '    ' export three tables to output directory:
        '    ' 1 options
        '    ' 2 connectivity
        '    ' 3 habitat
        '    If m_bCancel = True Then
        '        backgroundworker1.CancelAsync()
        '        backgroundworker1.Dispose()
        '        Exit Sub
        '    End If
        '    If iProgress1 < 70 Then
        '        iProgress1 = iProgress1 + 1
        '    End If
        '    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                                 "Exporting connectivity table to GLPK Model Directory.")

        '    ' ============ EXPORT CONNECTIVITY TABLE TO CSV ==============
        '    ' Delete the table if it already exists
        '    FileToDelete = sGLPKModelDir + "/FIPEX_GLPKConnectivity.csv"
        '    If System.IO.File.Exists(FileToDelete) = True Then
        '        System.IO.File.Delete(FileToDelete)

        '        ' need to reset the workspace so the file list will refresh and
        '        ' arcgis will know the file doesn't exist now.
        '        pWorkspaceOut = pTxtFactory.OpenFromFile(sGLPKModelDir, 0)
        '        pFWorkspaceOut = CType(pWorkspaceOut, IFeatureWorkspace)
        '        pDataSetOut = CType(pWorkspaceOut, IDataset)
        '    End If

        '    pExportOp = New ExportOperation

        '    ' Get input table
        '    pTable = pFWorkspace.OpenTable(sGLPKConnectTabName)

        '    ' Get the dataset name for the input table
        '    pDataSetIn = CType(pTable, IDataset)
        '    pDSNameIn = CType(pDataSetIn.FullName, IDatasetName)

        '    ' Get dataset for output table
        '    pDSNameOut = New TableName
        '    pDSNameOut.Name = "FIPEX_GLPKConnectivity.csv"
        '    pDSNameOut.WorkspaceName = CType(pDataSetOut.FullName, IWorkspaceName)

        '    Try
        '        pExportOp.ExportTable(pDSNameIn, Nothing, Nothing, pDSNameOut, My.ArcMap.Application.hWnd)
        '    Catch ex As Exception
        '        MsgBox("Error trying to export DBF table to GLPK Directory. " & ex.Message)
        '    End Try

        '    ' ======== EXPORT HABITAT TABLE TO CSV ==========
        '    ' Convert the habitat table
        '    If m_bCancel = True Then
        '        backgroundworker1.CancelAsync()
        '        backgroundworker1.Dispose()
        '        Exit Sub
        '    End If
        '    If iProgress1 < 70 Then
        '        iProgress1 = iProgress1 + 1
        '    End If
        '    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                     "Exporting habitat table to GLPK Model Directory.")

        '    pExportOp = New ExportOperation

        '    ' Input table
        '    pTable = pFWorkspace.OpenTable(sGLPKHabitatTableName)

        '    ' if undirected model is used, want to get maximum area
        '    ' for the parameters file
        '    Dim pCursor As ICursor
        '    Dim iNetworkQuanField As Integer

        '    If bUndirected = True Then
        '        iNetworkQuanField = pTable.FindField("HABITAT")
        '        If iNetworkQuanField <> -1 Then
        '            pCursor = pTable.Search(Nothing, False)
        '            Dim pField As IField
        '            Dim pRow As IRow
        '            pRow = pCursor.NextRow
        '            ' Loop through each row
        '            Do Until pRow Is Nothing

        '                pField = pRow.Fields.Field(iNetworkQuanField)
        '                If pField.Type = esriFieldType.esriFieldTypeDouble Or pField.Type = esriFieldType.esriFieldTypeInteger Or _
        '                    pField.Type = esriFieldType.esriFieldTypeSmallInteger Then
        '                    dMArea = dMArea + pRow.Value(iNetworkQuanField)
        '                Else
        '                    MsgBox("Error reading habtiat Table. Could not read the network quan field. Now Exiting.")
        '                    'BackgroundWorker1.CancelAsync()
        '                    'exitsubroutine()
        '                    Exit Sub
        '                End If
        '                pRow = pCursor.NextRow
        '            Loop
        '        Else
        '            MsgBox("Error reading habitat Table. Could not find field 'habitat'. Now Exiting.")
        '            Exit Sub
        '        End If
        '    End If ' bUndDirected is true

        '    ' Get the dataset name for the input table
        '    pDataSetIn = CType(pTable, IDataset)
        '    pDSNameIn = CType(pDataSetIn.FullName, IDatasetName)

        '    ' Delete the table if it already exists
        '    FileToDelete = sGLPKModelDir + "/FIPEX_GLPKHabitat3.csv"
        '    If System.IO.File.Exists(FileToDelete) = True Then
        '        System.IO.File.Delete(FileToDelete)

        '        ' need to reset the workspace so the file list will refresh and
        '        ' arcgis will know the file doesn't exist now.
        '        pWorkspaceOut = pTxtFactory.OpenFromFile(sGLPKModelDir, 0)
        '        pFWorkspaceOut = CType(pWorkspaceOut, IFeatureWorkspace)
        '        pDataSetOut = CType(pWorkspaceOut, IDataset)

        '    End If

        '    ' Get dataset for output table and export
        '    pDSNameOut = New TableName
        '    pDSNameOut.Name = "FIPEX_GLPKHabitat3.csv"
        '    pDSNameOut.WorkspaceName = CType(pDataSetOut.FullName, IWorkspaceName)

        '    Try
        '        pExportOp.ExportTable(pDSNameIn, _
        '                              Nothing, _
        '                              Nothing, _
        '                              pDSNameOut, _
        '                              My.ArcMap.Application.hWnd)
        '    Catch ex As Exception
        '        MsgBox("Error exporting GLPK habitat table to GLPK Directory. " & ex.Message)
        '    End Try

        '    ' ======== EXPORT OPTIONS TABLE TO CSV ==========
        '    ' Convert the options table

        '    If m_bCancel = True Then
        '        backgroundworker1.CancelAsync()
        '        backgroundworker1.Dispose()
        '        Exit Sub
        '    End If
        '    If iProgress1 < 70 Then
        '        iProgress1 = iProgress1 + 1
        '    End If
        '    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                     "Exporting options table to GLPK Model Directory.")

        '    pExportOp = New ExportOperation

        '    ' Input table
        '    pTable = pFWorkspace.OpenTable(sGLPKOptionsTableName)

        '    ' Get the dataset name for the input table
        '    pDataSetIn = CType(pTable, IDataset)
        '    pDSNameIn = CType(pDataSetIn.FullName, IDatasetName)

        '    ' Delete the table if it already exists
        '    FileToDelete = sGLPKModelDir + "/FIPEX_GLPKOptions.csv"
        '    If System.IO.File.Exists(FileToDelete) = True Then
        '        System.IO.File.Delete(FileToDelete)
        '    End If

        '    ' Get dataset for output table and export
        '    pDSNameOut = New TableName
        '    pDSNameOut.Name = "FIPEX_GLPKOptions.csv"
        '    pDSNameOut.WorkspaceName = CType(pDataSetOut.FullName, IWorkspaceName)

        '    Try
        '        pExportOp.ExportTable(pDSNameIn, _
        '                              Nothing, _
        '                              Nothing, _
        '                              pDSNameOut, _
        '                              My.ArcMap.Application.hWnd)
        '    Catch ex As Exception
        '        MsgBox("Error exporting connectivity table to GLPK Directory. " & ex.Message)
        '    End Try

        '    ' ==== PRIOR TO ANALYSIS SEARCH AND REPLACE PATH TXT IN MODEL FILE ===
        '    '  search and replace 
        '    ' replace c:\GunnsModel\ 
        '    ' with path to Model DIR (sGLPKModelDir)

        '    If m_bCancel = True Then
        '        backgroundworker1.CancelAsync()
        '        backgroundworker1.Dispose()
        '        Exit Sub
        '    End If
        '    If iProgress1 < 70 Then
        '        iProgress1 = iProgress1 + 1
        '    End If
        '    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                     "Rewriting Model file with user output directory.")

        '    If bDirected = True Then
        '        Try
        '            If System.IO.File.Exists(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod") = True Then
        '                sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod")
        '            End If

        '            n = 0
        '            For n = 0 To sModelFile.GetUpperBound(0)
        '                If sModelFile(n).Contains("C:\GunnsModel_REPLACE") = True Then
        '                    sModelFile(n) = sModelFile(n).Replace("C:\GunnsModel_REPLACE", sGLPKModelDir)
        '                End If
        '            Next

        '            ' overwrite file
        '            System.IO.File.WriteAllLines(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod", sModelFile)
        '        Catch ex As Exception
        '            MsgBox("Error trying to do 'search and replace' in the directed model file. " + ex.Message)
        '            Exit Sub
        '        End Try
        '    End If
        '    If bUndirected = True Then
        '        Try

        '            If System.IO.File.Exists(sGLPKModelDir + "/GreigUndirected.mod") = True Then
        '                sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/GreigUndirected.mod")
        '            End If

        '            n = 0
        '            For n = 0 To sModelFile.GetUpperBound(0)
        '                If sModelFile(n).Contains("C:\GunnsModel_REPLACE") = True Then
        '                    sModelFile(n) = sModelFile(n).Replace("C:\GunnsModel_REPLACE", sGLPKModelDir)
        '                End If
        '            Next

        '            ' overwrite file
        '            System.IO.File.WriteAllLines(sGLPKModelDir + "/GreigUndirected.mod", sModelFile)
        '        Catch ex As Exception
        '            MsgBox("Error trying to read or write to the undirected model file during 'replace'. " + ex.Message)
        '            Exit Sub
        '        End Try

        '    End If


        '    Dim lGLPKResultsObject As New List(Of DIR_OptResultsObject) ' for directed model
        '    Dim lGLPKResultsObject2 As New List(Of UNDIR_OptResultsObject)  ' for undirected model
        '    Dim lGRBResultsObject As New List(Of DIR_OptResultsObject) ' for directed model
        '    Dim lGRBResultsObject2 As New List(Of UNDIR_OptResultsObject)  ' for undirected model

        '    Dim iTotalBudgetIncrements As Integer
        '    iTotalBudgetIncrements = Convert.ToInt16((dMaxBudget - dMinBudget) / dBudgetIncrement)

        '    Dim iAnalysisCounter As Integer = 0
        '    Dim iEstimatedMaxTimeRemaining As Integer
        '    Dim GRBEnv1 As Gurobi.GRBEnv
        '    Dim sUndirectedModelPath1, sUndirectedModelPath2 As String
        '    Dim sDirectedModelPath1, sDirectedModelPath2 As String
        '    Dim iGRBLogFileLength As Integer
        '    Dim bGRBOptimal As Boolean = False
        '    Dim iGRBStatusCode As Integer
        '    Dim dGRBZMax As Double
        '    Dim dObjBound, dObjVal, dGRBPercentGap As Double
        '    Dim bBeginBounds As Boolean
        '    Dim sX, sIAMX As String
        '    Dim pGRBVar As Gurobi.GRBVar()
        '    Dim pSingleVar As Gurobi.GRBVar
        '    Dim sDecisionString As String
        '    Dim sOptionStringArray() As String
        '    Dim sOptionString As String
        '    Dim sBarrierIDString As String
        '    Dim GRBModel1 As Gurobi.GRBModel
        '    Dim pRowBuffer As IRowBuffer
        '    Dim dMinutesRemaining As Double
        '    Dim iMinutesRemaining As Integer
        '    Dim iEstimatedSecondsRemaining As Integer
        '    Dim iCentralBarrierEID As Integer
        '    Dim sCentralBarrierEID As String
        '    Dim dDecisionString As Double
        '    Dim bFileLocked As Boolean

        '    ' ========== FOR EACH BUDGET AMOUNT ===============
        '    ' Do until Max budget is exceeded

        '    dBudget = dMinBudget
        '    Do Until dBudget > dMaxBudget
        '        iAnalysisCounter += 1
        '        log = New System.IO.StreamWriter(sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt", True)
        '        log.WriteLine("Begin new analysis")
        '        log.Close()

        '        If bDirected = True Then
        '            ' ========== WRITE / OVERWRITE .DAT MODEL PARAMETERS FILE ===========
        '            ' Delete the table if it already exists
        '            FileToDelete = sGLPKModelDir + "/modelparameters.dat"
        '            If System.IO.File.Exists(FileToDelete) = True Then
        '                System.IO.File.Delete(FileToDelete)
        '            End If

        '            Try
        '                sw = File.CreateText(sGLPKModelDir + "/modelparameters.dat")
        '                sw.WriteLine("data;")
        '                sw.WriteLine("param FirstNod:=" + Convert.ToString(iSinkEID) + ";")
        '                sw.WriteLine("param mOptions:=" + Convert.ToString(iMaxOptions) + ";")
        '                sw.WriteLine("param Budget:=" + Convert.ToString(dBudget) + ";")
        '                sw.Close()

        '            Catch e As Exception
        '                MsgBox("The modelparameters.dat file in the GLPK Model directory could not be written to.  The following exception was found: " & e.Message)

        '            End Try

        '            ' ===== CALL THE MODEL AND WAIT TILL COMPLETED ========\
        '            ' will need to call the model as a batch executable. 
        '            ' the model executable will need to be rewritten each time
        '            ' to adapt for path name variations
        '            Try
        '                FileToDelete = sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat"
        '                If System.IO.File.Exists(FileToDelete) = True Then
        '                    System.IO.File.Delete(FileToDelete)
        '                End If

        '            Catch ex As Exception
        '                MsgBox("Error deleting the batch command file for GLPK Shell for directed model. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                sw = File.CreateText(sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat")
        '                sw.WriteLine("@ECHO ON")
        '                sw.WriteLine(":BEGIN")
        '                If iGLPKTimeLimit = 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --wcpxlp " + _
        '                                 sGLPKModelDir + "\forgurobi.lp >" + sGLPKModelDir + _
        '                                 "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit = 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat >" + _
        '                                 sGLPKModelDir + "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " >" + sGLPKModelDir + "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " --wcpxlp " + sGLPKModelDir + "\forgurobi.lp >" + sGLPKModelDir + "\outputMay251.txt")
        '                End If

        '                sw.WriteLine(":END")
        '                sw.Close()

        '            Catch e As Exception
        '                ' MsgBox("The FiPEX_GLPSOLCommand.bat file could not be written to.  The following exception was found: " & e.Message)
        '                Threading.Thread.Sleep(5000)
        '                sw.WriteLine("@ECHO ON")
        '                sw.WriteLine(":BEGIN")
        '                If iGLPKTimeLimit = 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --wcpxlp " + _
        '                                 sGLPKModelDir + "\forgurobi.lp >" + sGLPKModelDir + _
        '                                 "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit = 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat >" + _
        '                                 sGLPKModelDir + "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " >" + sGLPKModelDir + "\outputMay251.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + _
        '                                 "\May25_EldonsModel_FiPEx.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " --wcpxlp " + sGLPKModelDir + "\forgurobi.lp >" + sGLPKModelDir + "\outputMay251.txt")
        '                End If

        '                sw.WriteLine(":END")
        '                sw.Close()
        '            End Try

        '            ' notify user of time remaining
        '            iEstimatedMaxTimeRemaining = iGLPKTimeLimit * (iTotalBudgetIncrements - iAnalysisCounter)
        '            If bGurobiPickup = True Then
        '                iEstimatedMaxTimeRemaining = iGRBTimeLimit * (iTotalBudgetIncrements - iAnalysisCounter) + iEstimatedMaxTimeRemaining
        '            End If
        '            If bUndirected = True Then
        '                ' have to double the time 
        '                iEstimatedMaxTimeRemaining = iEstimatedMaxTimeRemaining * 2
        '            End If
        '            If m_bCancel = True Then
        '                backgroundworker1.CancelAsync()
        '                backgroundworker1.Dispose()
        '                Exit Sub
        '            End If
        '            If iProgress1 < 70 Then
        '                iProgress1 = iProgress1 + 1
        '            End If
        '            If iAnalysisCounter > iTotalBudgetIncrements Then
        '                iTotalBudgetIncrements = iAnalysisCounter
        '            End If
        '            Try
        '                If iEstimatedMaxTimeRemaining > 100 Then

        '                    dMinutesRemaining = iEstimatedMaxTimeRemaining / 60
        '                    iMinutesRemaining = Convert.ToInt32(Math.Round(dMinutesRemaining))
        '                    iEstimatedSecondsRemaining = Convert.ToInt32(Math.IEEERemainder(Convert.ToDouble(iEstimatedMaxTimeRemaining), 60))
        '                    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                                iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                "Estimated time remaining for optimization (max): " & iMinutesRemaining.ToString & "min " & iEstimatedSecondsRemaining.ToString _
        '                                                & "sec")
        '                Else
        '                    backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis" & ControlChars.NewLine & _
        '                                                 iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                 "Estimated time remaining for optimization (max): " & iEstimatedMaxTimeRemaining.ToString & "sec.")

        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error trying to update time remaining for user. " + ex.Message)
        '                Exit Sub

        '            End Try


        '            ' ***************************************
        '            Try
        '                Shell(sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat", AppWinStyle.MinimizedNoFocus, True)
        '            Catch ex As Exception
        '                MsgBox("Error calling the directed GLPK model from the shell. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            ' ***************************************


        '            ' ==== READ OUTPUT DECISION CSV AND UPDATE DECISION OBJECT ========
        '            pDIR_OptResultsObject = New DIR_OptResultsObject(Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing, _
        '                                                             Nothing)

        '            ' ------------------------------
        '            ' GET SINKEID iSinkEID
        '            pDIR_OptResultsObject.SinkEID = iSinkEID

        '            ' ------------------------------
        '            ' GET BUDGET dBudget
        '            pDIR_OptResultsObject.Budget = dBudget

        '            ' ------------------------------
        '            ' GET TREATMENT sGLPKTreatment
        '            pDIR_OptResultsObject.Treatment_Name = sGLPKTreatment

        '            ' ------------------------------
        '            ' GET GLPK_SOLVED bGLPKSolved
        '            Try
        '                If System.IO.File.Exists(sGLPKModelDir + "/outputMay251.txt") = True Then
        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/outputMay251.txt")
        '                End If

        '                n = 0
        '                bGLPKSolved = False
        '                For n = 0 To sOutputFile.GetUpperBound(0)
        '                    If sOutputFile(n).Contains("INTEGER OPTIMAL SOLUTION FOUND") = True Then
        '                        bGLPKSolved = True
        '                        iSolutionFoundRow = n
        '                        Exit For
        '                    End If
        '                Next
        '            Catch ex As Exception
        '                MsgBox("Error trying to read and scan the output file from the GLPK model. " + ex.Message)
        '                Exit Sub
        '            End Try


        '            pDIR_OptResultsObject.GLPK_Solved = bGLPKSolved
        '            Try
        '                ' ------------------------------
        '                ' GET PERCENT GAP dPercGap
        '                If bGLPKSolved = True Then
        '                    dPercGap = 0.0
        '                Else
        '                    n = 0
        '                    For n = 0 To sOutputFile.GetUpperBound(0)
        '                        If sOutputFile(n).Contains("TIME LIMIT EXCEEDED; SEARCH TERMINATED") = False And _
        '                           sOutputFile(n).Contains("PROBLEM HAS NO INTEGER FEASIBLE SOLUTION") = False Then
        '                            sPreviousOutputLine = sOutputFile(n)
        '                        Else
        '                            ' contingency code (should always contain %)
        '                            If sPreviousOutputLine.Contains("%") = False Then
        '                                dPercGap = 100
        '                                Exit For
        '                            End If
        '                            If sPreviousOutputLine.Contains("not yet found") Or _
        '                                sPreviousOutputLine.Contains("tree is empty") Then
        '                                dPercGap = 100
        '                            Else
        '                                ' found the last line before the 'time limit exceeded' warning
        '                                ' search for the '%' and keep the three digits prior as the Perc gap
        '                                sPreviousOutputLineSplit = sPreviousOutputLine.Split("%")
        '                                sPreviousOutputLine = sPreviousOutputLineSplit(0)
        '                                ' now trim all but the last four characters
        '                                dPercGap = Convert.ToDouble(sPreviousOutputLine.Remove(0, sPreviousOutputLine.Length - 4))
        '                            End If
        '                            Exit For
        '                        End If
        '                    Next

        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error getting percentage gap from the directed model output file (GLPK). Will set it to 100%. " + ex.Message)
        '                dPercGap = 100
        '            End Try


        '            pDIR_OptResultsObject.Perc_Gap = dPercGap

        '            ' ---------------------------
        '            ' GET TIME USED iGLPKTimeLimit
        '            pDIR_OptResultsObject.MaxSolTime = iGLPKTimeLimit

        '            ' ---------------------------
        '            ' GET TIME USED dTimeUsed
        '            Try
        '                If bGLPKSolved = True Then
        '                    ' split output file 'time used:   0.1 secs"
        '                    sSplitTimeUsed = sOutputFile(iSolutionFoundRow + 1)
        '                    sSplitTimeUsedFirstHalf = sSplitTimeUsed.Remove(0, 10)
        '                    sSplitTimeUsedFirstHalf2 = sSplitTimeUsedFirstHalf.TrimEnd("s", "e", "c")
        '                    dTimeUsed = Convert.ToDouble(sSplitTimeUsedFirstHalf2)
        '                Else
        '                    dTimeUsed = iGLPKTimeLimit
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Problem getting the solve time from the GLPK directed model output.  Will set it to 100%. " + ex.Message)
        '                dTimeUsed = iGLPKTimeLimit
        '            End Try


        '            pDIR_OptResultsObject.TimeUsed = dTimeUsed

        '            Try
        '                ' --------------------------
        '                ' GET HABITAT ZMAX
        '                If System.IO.File.Exists(sGLPKModelDir + "/ZMaxOutput.txt") = True Then
        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/ZMaxOutput.txt")
        '                    If sOutputFile(1).Length <> 0 Then
        '                        dMaxHabitat = Convert.ToDouble(sOutputFile(1))
        '                    Else
        '                        dMaxHabitat = 0
        '                    End If
        '                Else
        '                    dMaxHabitat = 0
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Problem getting the maximal network area from directed model output (GLPK).  Will set to zero. " + ex.Message)
        '                dMaxHabitat = 0
        '            End Try


        '            pDIR_OptResultsObject.Habitat_ZMax = dMaxHabitat
        '            lGLPKResultsObject.Add(pDIR_OptResultsObject)

        '            ' ===== READ THE DECISIONS INTO AN OBJECT ===========

        '            Try
        '                ' open output res2.csv (could be res1) and read each line
        '                If System.IO.File.Exists(sGLPKModelDir + "/Res1.csv") = True Then

        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/Res1.csv")
        '                    n = 1
        '                    ' split each line up, write to master list for this treatment
        '                    For n = 1 To sOutputFile.GetUpperBound(0)
        '                        sDecisionOutputSplit = sOutputFile(n).Split(",")
        '                        iDecisionEID = Convert.ToInt32(sDecisionOutputSplit(0))
        '                        iDecisionNum = Convert.ToInt16(sDecisionOutputSplit(1))
        '                        pGLPKDecisionsObject = New GLPKDecisionsObject(dBudget, _
        '                                                                       sGLPKTreatment, _
        '                                                                       iDecisionEID, _
        '                                                                       iDecisionNum)
        '                        lGLPKDecisionsObject.Add(pGLPKDecisionsObject)
        '                    Next
        '                Else
        '                    ' contingency code on fail.  should be a goto error 
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error trying to read output from directed model (GLPK). Could not find the decisions output in Res1 CSV file. " + ex.Message)
        '            End Try

        '            If bGurobiPickup = True Then
        '                Try
        '                    ' gurobi help examples ( http://www.gurobi.com/documentation/5.0/example-tour/node106)
        '                    GRBEnv1 = New Gurobi.GRBEnv(sGLPKModelDir + "\gurobilogfile.log")

        '                    ' delete any previous decisions output file
        '                    ' note: this file is not used anymore anyway
        '                    FileToDelete = sGLPKModelDir + "/GRBDecisions.txt"
        '                    If System.IO.File.Exists(FileToDelete) = True Then
        '                        System.IO.File.Delete(FileToDelete)
        '                    End If

        '                    Try
        '                        sw = File.CreateText(sGLPKModelDir + "/GRBDecisions.txt")
        '                        sw.WriteLine("Gurobi Decision Output File")
        '                        sw.WriteLine("For Budget Amount: " & dBudget.ToString)
        '                    Catch ex As Exception
        '                        MsgBox("Trouble Creating gurobi decision output file. now exiting. " & ex.Message)
        '                        Exit Sub
        '                    End Try

        '                    ' create the model object, importing model file from glpk output

        '                    bFileLocked = FileInUse(sGLPKModelDir + "\forgurobi.lp")
        '                    If System.IO.File.Exists(sGLPKModelDir + "\forgurobi.lp") = False Then
        '                        MsgBox("Error 234a: forgurobi.lp file does not exist." + _
        '                               "Likely error during GLPK analysis preventing file creation.")
        '                        Exit Try
        '                    End If
        '                    'MsgBox("File is locked before GRB optimization? " & bFileLocked)

        '                    sDirectedModelPath1 = sGLPKModelDir + "\forgurobi.lp"
        '                    sDirectedModelPath2 = sDirectedModelPath1.Replace("\", "/")

        '                    GRBModel1 = New Gurobi.GRBModel(GRBEnv1, sDirectedModelPath2)

        '                    ' set the time limit
        '                    If bGRBTimeLimit = True Then
        '                        GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.TimeLimit, iGRBTimeLimit)
        '                    End If
        '                    If bMIPGap = True Then
        '                        GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.MIPGap, dMIPGap)
        '                    End If

        '                    Try
        '                        GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.NodefileStart, 0.1)
        '                    Catch ex As Exception
        '                        MsgBox("Error 324511d. " + ex.Message)
        '                    End Try

        '                    'Try
        '                    '    GRBModel1.GetEnv().Set(Gurobi.GRB.IntParam.Threads, 2)
        '                    'Catch ex As Exception
        '                    '    MsgBox("Errort rhe3. " + ex.Message)
        '                    'End Try

        '                    GRBModel1.Optimize()

        '                    ' ========= Write GRB Results Summary to List Object ==================
        '                    pDIR_OptResultsObject = New DIR_OptResultsObject(Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing, _
        '                                                                     Nothing)

        '                    ' get solution status by code and get the current objective value, ZMax (AMaxMax)
        '                    ' see here for list of possible codes http://www.gurobi.com/documentation/5.0/reference-manual/node740
        '                    iGRBStatusCode = GRBModel1.Get(Gurobi.GRB.IntAttr.Status)
        '                    If iGRBStatusCode = Gurobi.GRB.Status.OPTIMAL Then
        '                        pDIR_OptResultsObject.GLPK_Solved = True
        '                        pDIR_OptResultsObject.Habitat_ZMax = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                    Else
        '                        pDIR_OptResultsObject.GLPK_Solved = False
        '                        'try to get objective value, otherwise set to zero
        '                        Try
        '                            dGRBZMax = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                            pDIR_OptResultsObject.Habitat_ZMax = dGRBZMax
        '                        Catch ex As Exception
        '                            pDIR_OptResultsObject.Habitat_ZMax = 0
        '                        End Try
        '                    End If

        '                    dGRBTimeUsed = GRBModel1.Get(Gurobi.GRB.DoubleAttr.Runtime)  ' time used
        '                    pDIR_OptResultsObject.TimeUsed = dGRBTimeUsed
        '                    pDIR_OptResultsObject.Budget = dBudget
        '                    pDIR_OptResultsObject.MaxSolTime = iGRBTimeLimit

        '                    ' apparently there's no way to get the gap easily, though it's printed to file
        '                    ' compute it manually
        '                    dObjBound = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjBound)
        '                    dObjVal = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                    Try
        '                        If dObjVal <> 0 Then
        '                            dGRBPercentGap = Math.Abs((dObjBound / dObjVal - 1) * 100)
        '                        Else
        '                            dGRBPercentGap = 100
        '                        End If
        '                    Catch ex As Exception
        '                        MsgBox("Problem calculating gap of solution from Gurobi. " + ex.Message)
        '                        dGRBPercentGap = 99
        '                    End Try
        '                    pDIR_OptResultsObject.Perc_Gap = dGRBPercentGap
        '                    pDIR_OptResultsObject.SinkEID = iSinkEID
        '                    pDIR_OptResultsObject.Treatment_Name = sGLPKTreatment

        '                    lGRBResultsObject.Add(pDIR_OptResultsObject)

        '                    ' ========== Write GRB Decision Results to List Object =============
        '                    ' get the decision result options and write them to a CSV (x[i,k] for all k)
        '                    ' x(node, decision) for all nodes
        '                    ' could do this by reading all individual decisions from the 'bounds' section of forgurobi.lp
        '                    ' read forgurobi.lp 'bounds' section, line by line, and get all x variables that do not have 1 as option
        '                    ' i.e., x(112354, 2) or x(1122354, 3)
        '                    ' print these to file - if their value is 1. i.e., [112334, 2]

        '                    bBeginBounds = False
        '                    pGRBVar = GRBModel1.GetVars()

        '                    ' If the LP file exists
        '                    If System.IO.File.Exists(sGLPKModelDir + "\forgurobi.lp") = True Then
        '                        ' Read it
        '                        sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "\forgurobi.lp")

        '                        bFileLocked = FileInUse(sGLPKModelDir + "\forgurobi.lp")
        '                        'MsgBox("File is locked after readalllines? " & bFileLocked)

        '                        n = 0
        '                        ' loop through it
        '                        For n = 0 To sModelFile.GetUpperBound(0)
        '                            ' if a flag is raised that we're in the 'Generals' (i.e., var declaration) section
        '                            If sModelFile(n).Length > 0 Then
        '                                If bBeginBounds = True Then
        '                                    ' if the variable is a decision variable in the second position
        '                                    If sModelFile(n).Chars(1).ToString = "x" Then


        '                                        ' get the variable name starting in the second position
        '                                        ' while trimming any leading or trailing spaces
        '                                        sX = sModelFile(n).Trim

        '                                        If sX = "End" Then
        '                                            MsgBox("Reached end of file of GRB output. Please check (shouldn't happen). Now exiting this round. ")
        '                                            Exit For
        '                                        End If

        '                                        ' get the variable from the solved GRB Model
        '                                        pSingleVar = GRBModel1.GetVarByName(sX)
        '                                        Try
        '                                            If iGRBStatusCode = Gurobi.GRB.Status.OPTIMAL Then
        '                                                sDecisionString = pSingleVar.Get(GRB.DoubleAttr.X)
        '                                            Else
        '                                                ' there should be more than one solution - looking for the next
        '                                                ' best solution (suboptimal)
        '                                                ' see here http://www.gurobi.com/documentation/5.0/example-tour/node114
        '                                                If GRBModel1.Get(GRB.IntAttr.SolCount) > 0 Then
        '                                                    ' this sets solution number to 1 -- the second best solution found
        '                                                    GRBModel1.GetEnv().Set(GRB.IntParam.SolutionNumber, 1)
        '                                                    ' Xn means the optimal value for the nth solution number
        '                                                    sDecisionString = pSingleVar.Get(GRB.DoubleAttr.Xn)
        '                                                End If
        '                                            End If
        '                                        Catch ex As Exception
        '                                            MsgBox("Exception raised trying to get the decision string for variable from the GRB variable using GRB API (directed model). " + ex.Message)
        '                                            Exit Sub
        '                                        End Try
        '                                        ' If this variable has a decision at it (and every x will have a decision, 
        '                                        '     even if it is 'do nothing' (i.e. x(i,1)) 

        '                                        ' need to convert the decision string to an double and round it to the nearest integer
        '                                        ' this is because of an error where the decision string
        '                                        ' is sometimes a weird number like 0.999999920249585
        '                                        ' this problem was encountered Dec. 2, 2012
        '                                        Try
        '                                            dDecisionString = Convert.ToDouble(sDecisionString)
        '                                        Catch ex As Exception
        '                                            MsgBox("Could not convert decision string to type 'double' in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                            dDecisionString = 0
        '                                        End Try
        '                                        Try
        '                                            dDecisionString = Math.Round(dDecisionString, 0, MidpointRounding.AwayFromZero)
        '                                        Catch ex As Exception
        '                                            MsgBox("trouble rounding decision string to single digit integer in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                            dDecisionString = 0
        '                                        End Try


        '                                        If dDecisionString > 0.49 Then
        '                                            ' find out if it is 'do nothing' 
        '                                            sOptionStringArray = sX.Split(",")
        '                                            sOptionString = sOptionStringArray(1).TrimEnd(")")
        '                                            ' if it's not do nothing then get print the barrier ID and option chosen to file
        '                                            If sOptionString.Trim <> "1" Then
        '                                                sBarrierIDString = sOptionStringArray(0).TrimStart("(", "x")
        '                                                sBarrierIDString = sBarrierIDString.Trim
        '                                                sw.WriteLine(sBarrierIDString + "," + sOptionString)
        '                                                Try
        '                                                    iDecisionEID = Convert.ToInt32(sBarrierIDString)
        '                                                Catch ex As Exception
        '                                                    MsgBox("Trouble converting Node ID to type Integer in Gurobi Decision Output (undirected model). Converting ID to '999999'.")
        '                                                    iDecisionEID = 999999
        '                                                End Try
        '                                                Try
        '                                                    iDecisionNum = Convert.ToInt32(sOptionString)
        '                                                Catch ex As Exception
        '                                                    MsgBox("Trouble converting OptionNum to type Integer in Gurobi Decision Output (undirected model). Converting optionnum to '99'.")
        '                                                    iDecisionNum = 99
        '                                                End Try
        '                                                ' using an object named here prior to GUROBI  - same fields/attribute/parameters though
        '                                                pGLPKDecisionsObject = New GLPKDecisionsObject(dBudget, _
        '                                                                                               sGLPKTreatment + "_GRB", _
        '                                                                                               iDecisionEID, _
        '                                                                                               iDecisionNum)
        '                                                lGRBDecisionsObject.Add(pGLPKDecisionsObject)
        '                                            End If
        '                                        End If
        '                                    End If
        '                                End If

        '                            End If
        '                            If sModelFile(n).Contains("Generals") = True Then
        '                                bBeginBounds = True
        '                            End If
        '                        Next ' row in the LP model input file
        '                    End If
        '                    sw.Close()
        '                    sDirectedModelPath1 = sGLPKModelDir + "\dummy.lp"
        '                    sDirectedModelPath2 = sDirectedModelPath1.Replace("\", "/")
        '                    GRBModel1 = New Gurobi.GRBModel(GRBEnv1, sDirectedModelPath2)
        '                    GRBModel1.Dispose()
        '                    GRBEnv1.Dispose()
        '                Catch ex As GRBException
        '                    MsgBox("Problem encountered running gurobi analysis after GLPK directed analysis." + ex.Message)
        '                End Try
        '            End If ' GUROBI is being used for tough problems

        '        End If ' bdirected is true


        '        bFileLocked = FileInUse(sGLPKModelDir + "\forgurobi.lp")
        '        'MsgBox("File is locked After GRB optimization? " & bFileLocked)


        '        Dim sTemp As String
        '        Dim bCentralBarrierEIDFound As Boolean

        '        If bUndirected = True Then

        '            ' ========== WRITE / OVERWRITE .DAT MODEL PARAMETERS FILE ===========
        '            Try
        '                ' Delete the table if it already exists
        '                FileToDelete = sGLPKModelDir + "/modelparameters_undirected.dat"
        '                If System.IO.File.Exists(FileToDelete) = True Then
        '                    System.IO.File.Delete(FileToDelete)
        '                End If

        '            Catch ex As Exception
        '                MsgBox("Error trying to delete old model parameters file for undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                sw = File.CreateText(sGLPKModelDir + "/modelparameters_undirected.dat")
        '                sw.WriteLine("data;")
        '                sw.WriteLine("param FirstNod:=" + Convert.ToString(iSinkEID) + ";")
        '                sw.WriteLine("param mOptions:=" + Convert.ToString(iMaxOptions) + ";")
        '                sw.WriteLine("param Budget:=" + Convert.ToString(dBudget) + ";")
        '                sw.WriteLine("param MArea:=" + Convert.ToString(dMArea) + ";")
        '                sw.Close()

        '            Catch e As Exception
        '                MsgBox("The modelparameters_undirected.dat file in the GLPK Model directory could not be written to.  The following exception was found: " & e.Message)

        '            End Try
        '            Try
        '                FileToDelete = sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat"
        '                If System.IO.File.Exists(FileToDelete) = True Then
        '                    System.IO.File.Delete(FileToDelete)
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Problem Deleting the model batch file. Error message: " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                sw = File.CreateText(sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat")
        '                sw.WriteLine("@ECHO ON")
        '                sw.WriteLine(":BEGIN")
        '                If iGLPKTimeLimit = 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                ElseIf iGLPKTimeLimit = 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat --wcpxlp " + _
        '                                 sGLPKModelDir + "\forgurobiundir.lp >" + sGLPKModelDir + "\outputGreigUndir.txt")

        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " --wcpxlp " + sGLPKModelDir + "\forgurobiundir.lp >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters_undirected.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + " >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                End If
        '                sw.WriteLine(":END")
        '                sw.Close()

        '            Catch e As Exception
        '                'MsgBox("The FiPEX_GLPSOLCommand.bat file could not be written to.  The following exception was found: " & e.Message)
        '                Threading.Thread.Sleep(5000)
        '                sw = File.CreateText(sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat")
        '                sw.WriteLine("@ECHO ON")
        '                sw.WriteLine(":BEGIN")
        '                If iGLPKTimeLimit = 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                ElseIf iGLPKTimeLimit = 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat --wcpxlp " + _
        '                                 sGLPKModelDir + "\forgurobiundir.lp >" + sGLPKModelDir + "\outputGreigUndir.txt")

        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = True Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + _
        '                                 sGLPKModelDir + "\modelparameters_undirected.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + _
        '                                 " --wcpxlp " + sGLPKModelDir + "\forgurobiundir.lp >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                ElseIf iGLPKTimeLimit <> 0 And bGurobiPickup = False Then
        '                    sw.WriteLine(sGnuWinDir + "\bin\glpsol -m " + sGLPKModelDir + "\GreigUndirected.mod -d " + sGLPKModelDir + _
        '                                 "\modelparameters_undirected.dat --tmlim " + Convert.ToString(iGLPKTimeLimit) + " >" + sGLPKModelDir + "\outputGreigUndir.txt")
        '                End If
        '                sw.WriteLine(":END")
        '                sw.Close()
        '            End Try

        '            ' notify user of time remaining
        '            iEstimatedMaxTimeRemaining = iGLPKTimeLimit * (iTotalBudgetIncrements - iAnalysisCounter)
        '            If bGurobiPickup = True Then
        '                iEstimatedMaxTimeRemaining = iGRBTimeLimit * (iTotalBudgetIncrements - iAnalysisCounter) + iEstimatedMaxTimeRemaining
        '            End If
        '            If bDirected = True Then
        '                ' have to double the time 
        '                iEstimatedMaxTimeRemaining = iEstimatedMaxTimeRemaining * 2
        '                ' directed model would have completed its round at this budget by now
        '                iEstimatedMaxTimeRemaining = iEstimatedMaxTimeRemaining - iGLPKTimeLimit
        '                If bGurobiPickup = True Then
        '                    iEstimatedMaxTimeRemaining = iEstimatedMaxTimeRemaining - iGRBTimeLimit
        '                End If
        '            End If
        '            If m_bCancel = True Then
        '                backgroundworker1.CancelAsync()
        '                backgroundworker1.Dispose()
        '                Exit Sub
        '            End If
        '            If iProgress1 < 70 Then
        '                iProgress1 = iProgress1 + 1
        '            End If
        '            If iAnalysisCounter > iTotalBudgetIncrements Then
        '                iTotalBudgetIncrements = iAnalysisCounter
        '            End If
        '            Try
        '                If iEstimatedMaxTimeRemaining > 100 Then
        '                    dMinutesRemaining = iEstimatedMaxTimeRemaining / 60
        '                    iMinutesRemaining = Convert.ToInt32(Math.Round(dMinutesRemaining))
        '                    iEstimatedSecondsRemaining = Convert.ToInt32(Math.IEEERemainder(Convert.ToDouble(iEstimatedMaxTimeRemaining), 60))
        '                    backgroundworker1.ReportProgress(iProgress1, "Performing GLPK Optimisation Analysis ('undirected' model)" & ControlChars.NewLine & _
        '                                                iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                "Estimated time remaining for optimization (max): " & iMinutesRemaining.ToString & "min " & iEstimatedSecondsRemaining.ToString _
        '                                                & "sec")
        '                Else
        '                    backgroundworker1.ReportProgress(iProgress1, "Performing GLPK Optimisation Analysis ('undirected' model)" & ControlChars.NewLine & _
        '                                                 iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                 "Estimated time remaining for optimization (max): " & iEstimatedMaxTimeRemaining.ToString & "sec.")
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error trying to update user with remaining time. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' ---------------------------------------------
        '            ' delete content from previous GLPK file
        '            ' so if the model fails then it won't give an answer anyway. 
        '            Try
        '                ' --------------------------
        '                ' GET HABITAT ZMAX
        '                If System.IO.File.Exists(sGLPKModelDir + "/UNDIROutput.txt") = True Then
        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/UNDIROutput.txt")
        '                    If sOutputFile(1).Length <> 0 Then
        '                        dMaxHabitat = Convert.ToDouble(sOutputFile(1))
        '                    Else
        '                        dMaxHabitat = 0
        '                    End If
        '                    If sOutputFile(3).Length <> 0 Then
        '                        iCentralBarrierEID = Convert.ToInt32(sOutputFile(3))
        '                    Else
        '                        iCentralBarrierEID = 0
        '                    End If
        '                Else
        '                    dMaxHabitat = 0
        '                    iCentralBarrierEID = 0
        '                End If
        '            Catch ex As Exception
        '                'MsgBox("Trouble getting the maximum network habitat area and central barrier. Will set them to zero. " + ex.Message)
        '                dMaxHabitat = 0
        '                iCentralBarrierEID = 0
        '            End Try




        '            ' **********************************************
        '            Try
        '                Shell(sGLPKModelDir + "/FiPEX_GLPSOLCommand.bat", AppWinStyle.MinimizedNoFocus, True)
        '            Catch ex As Exception
        '                MsgBox("Error running GLPK model shell command, calling the batch file. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            ' **********************************************

        '            ' ==== READ OUTPUT DECISION CSV AND UPDATE DECISION OBJECT ========
        '            pUNDIR_ResultsObject = New UNDIR_OptResultsObject(Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing, _
        '                                                              Nothing)

        '            ' ------------------------------
        '            ' GET SINKEID iSinkEID
        '            pUNDIR_ResultsObject.SinkEID = iSinkEID
        '            ' ------------------------------
        '            ' GET BUDGET dBudget
        '            pUNDIR_ResultsObject.Budget = dBudget
        '            ' ------------------------------
        '            ' GET TREATMENT sGLPKTreatment
        '            pUNDIR_ResultsObject.Treatment_Name = sGLPKTreatment
        '            ' ------------------------------
        '            ' GET GLPK_SOLVED bGLPKSolved
        '            Try

        '                If System.IO.File.Exists(sGLPKModelDir + "/outputGreigUndir.txt") = True Then
        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/outputGreigUndir.txt")
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Problem Deleting the model batch file. Error message: " + ex.Message)
        '                Exit Sub
        '            End Try

        '            n = 0
        '            bGLPKSolved = False
        '            For n = 0 To sOutputFile.GetUpperBound(0)
        '                If sOutputFile(n).Contains("INTEGER OPTIMAL SOLUTION FOUND") = True Then
        '                    bGLPKSolved = True
        '                    iSolutionFoundRow = n
        '                    Exit For
        '                End If
        '            Next

        '            pUNDIR_ResultsObject.GLPK_Solved = bGLPKSolved
        '            Try
        '                ' ------------------------------
        '                ' GET PERCENT GAP dPercGap
        '                If bGLPKSolved = True Then
        '                    dPercGap = 0.0
        '                Else
        '                    n = 0
        '                    For n = 0 To sOutputFile.GetUpperBound(0)
        '                        If sOutputFile(n).Contains("TIME LIMIT EXCEEDED; SEARCH TERMINATED") = False And _
        '                            sOutputFile(n).Contains("PROBLEM HAS NO INTEGER FEASIBLE SOLUTION") = False Then
        '                            sPreviousOutputLine = sOutputFile(n)
        '                        Else
        '                            ' contingency code (should always contain %)
        '                            If sPreviousOutputLine.Contains("%") = False Then
        '                                dPercGap = 100
        '                                Exit For
        '                            End If
        '                            If sPreviousOutputLine.Contains("not yet found") Or _
        '                                sPreviousOutputLine.Contains("tree is empty") Then
        '                                dPercGap = 100
        '                            Else
        '                                ' found the last line before the 'time limit exceeded' warning
        '                                ' search for the '%' and keep the three digits prior as the Perc gap
        '                                sPreviousOutputLineSplit = sPreviousOutputLine.Split("%")
        '                                sPreviousOutputLine = sPreviousOutputLineSplit(0)
        '                                ' now trim all but the last four characters
        '                                dPercGap = Convert.ToDouble(sPreviousOutputLine.Remove(0, sPreviousOutputLine.Length - 4))
        '                            End If
        '                            Exit For
        '                        End If
        '                    Next

        '                End If

        '            Catch ex As Exception
        '                MsgBox("Problem encountered getting perc. gap from undirected model output (GLPK). will set it to 100%. " + ex.Message)
        '                dPercGap = 100
        '            End Try

        '            pUNDIR_ResultsObject.Perc_Gap = dPercGap

        '            ' ---------------------------
        '            ' GET TIME USED iGLPKTimeLimit
        '            pUNDIR_ResultsObject.MaxSolTime = iGLPKTimeLimit

        '            Try
        '                ' ---------------------------
        '                ' GET TIME USED dTimeUsed
        '                If bGLPKSolved = True Then
        '                    ' split output file 'time used:   0.1 secs"
        '                    sSplitTimeUsed = sOutputFile(iSolutionFoundRow + 1)
        '                    sSplitTimeUsedFirstHalf = sSplitTimeUsed.Remove(0, 10)
        '                    sSplitTimeUsedFirstHalf2 = sSplitTimeUsedFirstHalf.TrimEnd("s", "e", "c")
        '                    dTimeUsed = Convert.ToDouble(sSplitTimeUsedFirstHalf2)
        '                Else
        '                    dTimeUsed = iGLPKTimeLimit
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Problem getting solution time used from undirected model output (GLPK).  Will set it to 100%. " + ex.Message)
        '                dTimeUsed = iGLPKTimeLimit
        '            End Try


        '            pUNDIR_ResultsObject.TimeUsed = dTimeUsed

        '            Try
        '                ' --------------------------
        '                ' GET HABITAT ZMAX
        '                If System.IO.File.Exists(sGLPKModelDir + "/UNDIROutput.txt") = True Then
        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/UNDIROutput.txt")
        '                    If sOutputFile(1).Length <> 0 Then
        '                        dMaxHabitat = Convert.ToDouble(sOutputFile(1))
        '                    Else
        '                        dMaxHabitat = 0
        '                    End If
        '                    If sOutputFile(3).Length <> 0 Then
        '                        iCentralBarrierEID = Convert.ToInt32(sOutputFile(3))
        '                    Else
        '                        iCentralBarrierEID = 0
        '                    End If
        '                Else
        '                    dMaxHabitat = 0
        '                    iCentralBarrierEID = 0
        '                End If
        '            Catch ex As Exception
        '                'MsgBox("Trouble getting the maximum network habitat area and central barrier. Will set them to zero. " + ex.Message)
        '                dMaxHabitat = 0
        '                iCentralBarrierEID = 0
        '            End Try

        '            pUNDIR_ResultsObject.Habitat_ZMax = dMaxHabitat
        '            pUNDIR_ResultsObject.CentralBarrierEID = iCentralBarrierEID
        '            lGLPKResultsObject2.Add(pUNDIR_ResultsObject)

        '            Try
        '                ' ===== READ THE DECISIONS INTO AN OBJECT ===========
        '                ' open output res2.csv (could be res1) and read each line
        '                If System.IO.File.Exists(sGLPKModelDir + "/Res1_undirected.csv") = True Then

        '                    sOutputFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/Res1_undirected.csv")
        '                    n = 1
        '                    ' split each line up, write to master list for this treatment
        '                    For n = 1 To sOutputFile.GetUpperBound(0)
        '                        sDecisionOutputSplit = sOutputFile(n).Split(",")
        '                        iDecisionEID = Convert.ToInt32(sDecisionOutputSplit(0))
        '                        iDecisionNum = Convert.ToInt16(sDecisionOutputSplit(1))
        '                        pGLPKDecisionsObject = New GLPKDecisionsObject(dBudget, _
        '                                                                       sGLPKTreatment, _
        '                                                                       iDecisionEID, _
        '                                                                       iDecisionNum)
        '                        lGLPKDecisionsObject2.Add(pGLPKDecisionsObject)
        '                    Next
        '                Else
        '                    ' contingency code on fail.  should be a goto error 
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Trouble finding the res1_undirected CSV output file to retrieve decisions from undirected model output. " _
        '                       + "There will be no decisions output for this budget amount (GLPK model), budget: " + Convert.ToString(dBudget) _
        '                       + " " + ex.Message)
        '            End Try




        '            If bGurobiPickup = True Then
        '                ' gurobi help examples ( http://www.gurobi.com/documentation/5.0/example-tour/node106)
        '                Try

        '                    GRBEnv1 = New Gurobi.GRBEnv(sGLPKModelDir + "\gurobilogfile.log")
        '                Catch ex As Exception
        '                    MsgBox("Trouble getting gurobi environment. Now Exiting. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                Try
        '                    ' delete any previous decisions output file
        '                    FileToDelete = sGLPKModelDir + "/GRBDecisions.txt"
        '                    If System.IO.File.Exists(FileToDelete) = True Then
        '                        System.IO.File.Delete(FileToDelete)
        '                    End If
        '                Catch ex As Exception
        '                    MsgBox("Trouble deleting GRBDecisions file. " + ex.Message)
        '                    Exit Sub
        '                End Try


        '                Try
        '                    sw = File.CreateText(sGLPKModelDir + "/GRBDecisions.txt")
        '                    sw.WriteLine("Gurobi Decision Output File")
        '                    sw.WriteLine("For Budget Amount: " & dBudget.ToString)
        '                Catch ex As Exception
        '                    MsgBox("Trouble Creating gurobi decision output file. now exiting. " & ex.Message)
        '                    Exit Sub
        '                End Try

        '                ' create the model object, importing model file from glpk output

        '                sUndirectedModelPath1 = sGLPKModelDir + "\forgurobiundir.lp"
        '                sUndirectedModelPath2 = sUndirectedModelPath1.Replace("\", "/")

        '                GRBModel1 = New Gurobi.GRBModel(GRBEnv1, sUndirectedModelPath2)

        '                ' set the time limit
        '                If bGRBTimeLimit = True Then
        '                    GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.TimeLimit, iGRBTimeLimit)
        '                End If
        '                If bMIPGap = True Then
        '                    GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.MIPGap, dMIPGap)
        '                End If

        '                Try
        '                    GRBModel1.GetEnv().Set(Gurobi.GRB.DoubleParam.NodefileStart, 0.1)
        '                Catch ex As Exception
        '                    MsgBox("Error 324511d. " + ex.Message)
        '                End Try

        '                'Try
        '                '    GRBModel1.GetEnv().Set(Gurobi.GRB.IntParam.Threads, 2)
        '                'Catch ex As Exception
        '                '    MsgBox("Errort rhe3. " + ex.Message)
        '                'End Try

        '                Try

        '                    If iEstimatedMaxTimeRemaining > 100 Then
        '                        dMinutesRemaining = iEstimatedMaxTimeRemaining / 60
        '                        iMinutesRemaining = Convert.ToInt32(Math.Round(dMinutesRemaining))
        '                        iEstimatedSecondsRemaining = Convert.ToInt32(Math.IEEERemainder(Convert.ToDouble(iEstimatedMaxTimeRemaining), 60))
        '                        backgroundworker1.ReportProgress(iProgress1, "Performing Gurobi Optimization Analysis ('undirected' model)" & ControlChars.NewLine & _
        '                                                    iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                    "Estimated time remaining for optimization (max): " & iMinutesRemaining.ToString & "min " & iEstimatedSecondsRemaining.ToString _
        '                                                    & "sec")
        '                    Else
        '                        backgroundworker1.ReportProgress(iProgress1, "Performing Optimization Analysis ('undirected' model)" & ControlChars.NewLine & _
        '                                                     iAnalysisCounter.ToString & " of ~" & iTotalBudgetIncrements.ToString & " rounds" & ControlChars.NewLine & _
        '                                                     "Estimated time remaining for optimization (max): " & iEstimatedMaxTimeRemaining.ToString & "sec.")
        '                    End If

        '                    GRBModel1.Optimize()
        '                Catch ex As Exception
        '                    MsgBox("Error encountered during optimisation. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                ' ========= Write GRB Results Summary to List Object ==================
        '                ' Note, just using same variable here as GLPK, rather than creating another object for GRB case
        '                ' . List object is different though. 
        '                pUNDIR_ResultsObject = New UNDIR_OptResultsObject(Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing, _
        '                                                                  Nothing)

        '                ' get solution status by code and get the current objective value, ZMax (AMaxMax)
        '                ' see here for list of possible codes http://www.gurobi.com/documentation/5.0/reference-manual/node740
        '                iGRBStatusCode = GRBModel1.Get(Gurobi.GRB.IntAttr.Status)
        '                If iGRBStatusCode = Gurobi.GRB.Status.OPTIMAL Then
        '                    pUNDIR_ResultsObject.GLPK_Solved = True
        '                    pUNDIR_ResultsObject.Habitat_ZMax = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                Else
        '                    pUNDIR_ResultsObject.GLPK_Solved = False
        '                    'try to get objective value, otherwise set to zero
        '                    Try
        '                        dGRBZMax = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                        pUNDIR_ResultsObject.Habitat_ZMax = dGRBZMax
        '                    Catch ex As Exception
        '                        pUNDIR_ResultsObject.Habitat_ZMax = 0
        '                    End Try
        '                End If

        '                dGRBTimeUsed = GRBModel1.Get(Gurobi.GRB.DoubleAttr.Runtime)  ' time used
        '                pUNDIR_ResultsObject.TimeUsed = dGRBTimeUsed
        '                pUNDIR_ResultsObject.Budget = dBudget
        '                pUNDIR_ResultsObject.MaxSolTime = iGRBTimeLimit

        '                ' apparently there's no way to get the gap easily, though it's printed to file
        '                ' compute it manually
        '                Try

        '                    dObjBound = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjBound)
        '                Catch ex As Exception
        '                    'MsgBox("Could not get Objective Bound from undirected model Gurobi output. ")
        '                    dObjBound = 0
        '                End Try
        '                If dObjBound <> 0 Then
        '                    dObjVal = GRBModel1.Get(Gurobi.GRB.DoubleAttr.ObjVal)
        '                    Try
        '                        If dObjVal <> 0 Then
        '                            dGRBPercentGap = Math.Abs((dObjBound / dObjVal - 1) * 100)
        '                        Else
        '                            dGRBPercentGap = 100
        '                        End If
        '                    Catch ex As Exception
        '                        MsgBox("Problem calculating gap of solution from Gurobi. " + ex.Message)
        '                        dGRBPercentGap = 99
        '                    End Try
        '                Else
        '                    dGRBPercentGap = 0
        '                End If
        '                pUNDIR_ResultsObject.Perc_Gap = dGRBPercentGap
        '                pUNDIR_ResultsObject.SinkEID = iSinkEID
        '                pUNDIR_ResultsObject.Treatment_Name = sGLPKTreatment

        '                ' Gotta look at all the possible nodes, and then find the maximal one (iamx).
        '                ' easiest way to get loop through these vars is to read them each out of the model file
        '                ' and then see if their value is 'one'
        '                ' If the LP file exists
        '                bBeginBounds = False
        '                bCentralBarrierEIDFound = False
        '                If System.IO.File.Exists(sGLPKModelDir + "\forgurobiundir.lp") = True Then
        '                    ' Read it
        '                    sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "\forgurobiundir.lp")
        '                    n = 0
        '                    ' loop through it
        '                    For n = 0 To sModelFile.GetUpperBound(0)
        '                        ' if a flag is raised that we're in the 'Generals' (i.e., var declaration) section
        '                        If sModelFile(n).Length > 0 Then
        '                            If bBeginBounds = True Then
        '                                ' if the variable is a decision variable in the second position
        '                                Try
        '                                    sTemp = sModelFile(n).Substring(1)
        '                                Catch ex As Exception
        '                                    MsgBox("Error in GRB retrieval of IAMX")
        '                                    Exit Sub
        '                                End Try
        '                                Try
        '                                    sTemp = sModelFile(n).Substring(0, 3)
        '                                Catch ex As Exception
        '                                    MsgBox("Error in checking whether reached end of file.")
        '                                    Exit Sub
        '                                End Try
        '                                If sTemp = "End" Then
        '                                    Exit For
        '                                End If
        '                                Try
        '                                    sTemp = sModelFile(n).Substring(1, 4)
        '                                Catch ex As Exception
        '                                    MsgBox("Error retrieving first four characters of file in GRB retrieval of IAMX. ")
        '                                    Exit Sub
        '                                End Try
        '                                If sTemp = "iamx" Then
        '                                    ' get the variable name starting in the second position
        '                                    ' while trimming any leading or trailing spaces
        '                                    sIAMX = sModelFile(n).Trim
        '                                    ' get the variable from the solved GRB Model
        '                                    pSingleVar = GRBModel1.GetVarByName(sIAMX)
        '                                    If pSingleVar Is Nothing Then
        '                                        MsgBox("Could not retrieve variable by name in GRB API. variable: " + sIAMX)
        '                                        Exit Sub
        '                                    End If
        '                                    Try
        '                                        If iGRBStatusCode = Gurobi.GRB.Status.OPTIMAL Then
        '                                            sDecisionString = pSingleVar.Get(GRB.DoubleAttr.X)
        '                                        Else
        '                                            ' there should be more than one solution - looking for the next
        '                                            ' best solution (suboptimal)
        '                                            ' see here http://www.gurobi.com/documentation/5.0/example-tour/node114
        '                                            If GRBModel1.Get(GRB.IntAttr.SolCount) > 0 Then
        '                                                ' this sets solution number to 1 -- the second best solution found
        '                                                GRBModel1.GetEnv().Set(GRB.IntParam.SolutionNumber, 1)
        '                                                ' Xn means the optimal value for the nth solution number
        '                                                sDecisionString = pSingleVar.Get(GRB.DoubleAttr.Xn)
        '                                            End If
        '                                        End If
        '                                    Catch ex As Exception
        '                                        MsgBox("Exception raised trying to get the decision string for variable from the GRB variable using GRB API. " + ex.Message)
        '                                        Exit Sub
        '                                    End Try


        '                                    ' need to convert the decision string to an double and round it to the nearest integer
        '                                    ' this is because of an error where the decision string
        '                                    ' is sometimes a weird number like 0.999999920249585
        '                                    ' this problem was encountered Dec. 2, 2012
        '                                    Try
        '                                        dDecisionString = Convert.ToDouble(sDecisionString)
        '                                    Catch ex As Exception
        '                                        MsgBox("Could not convert decision string to type 'double' in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                        dDecisionString = 0
        '                                    End Try
        '                                    Try
        '                                        dDecisionString = Math.Round(dDecisionString, 0, MidpointRounding.AwayFromZero)
        '                                    Catch ex As Exception
        '                                        MsgBox("trouble rounding decision string to single digit integer in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                        dDecisionString = 0
        '                                    End Try

        '                                    ' If this variable has a decision at it (and every x will have a decision, 
        '                                    '     even if it is 'do nothing' (i.e. x(i,1)) 
        '                                    If dDecisionString > 0.49 Then
        '                                        ' find out if it is 'do nothing' 
        '                                        sCentralBarrierEID = sIAMX.TrimEnd(")")
        '                                        sCentralBarrierEID = sCentralBarrierEID.TrimStart("i", "a", "m", "x", "(")

        '                                        Try
        '                                            iCentralBarrierEID = Convert.ToInt32(sCentralBarrierEID)
        '                                        Catch ex As Exception
        '                                            MsgBox("Trouble converting Central Node ID to type Integer in Gurobi Decision Output (undirected model). Converting ID to '999999'.")
        '                                            iCentralBarrierEID = 999999
        '                                        End Try
        '                                        bCentralBarrierEIDFound = True
        '                                        pUNDIR_ResultsObject.CentralBarrierEID = iCentralBarrierEID
        '                                        Exit For
        '                                    End If
        '                                End If

        '                            End If
        '                        End If
        '                        If sModelFile(n).Contains("Generals") = True Then
        '                            bBeginBounds = True
        '                        End If
        '                    Next ' row in the LP model input file
        '                    If bCentralBarrierEIDFound = False Then
        '                        MsgBox("No central barrier EID could be retrieved from the Gurobi API and output. ID is set to 999999")
        '                        pUNDIR_ResultsObject.CentralBarrierEID = 999999
        '                    End If
        '                End If

        '                ' get the node/barrier which is immediately downstream of the central network feature
        '                lGRBResultsObject2.Add(pUNDIR_ResultsObject)

        '                ' ========== Write GRB Decision Results to List Object =============
        '                ' get the decision result options and write them to a CSV (x[i,k] for all k)
        '                ' x(node, decision) for all nodes
        '                ' could do this by reading all individual decisions from the 'bounds' section of forgurobi.lp
        '                ' read forgurobi.lp 'bounds' section, line by line, and get all x variables that do not have 1 as option
        '                ' i.e., x(112354, 2) or x(1122354, 3)
        '                ' print these to file - if their value is 1. i.e., [112334, 2]

        '                bBeginBounds = False
        '                pGRBVar = GRBModel1.GetVars()

        '                ' If the LP file exists
        '                If System.IO.File.Exists(sGLPKModelDir + "\forgurobiundir.lp") = True Then
        '                    ' Read it
        '                    sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "\forgurobiundir.lp")
        '                    n = 0
        '                    ' loop through it
        '                    For n = 0 To sModelFile.GetUpperBound(0)
        '                        ' if a flag is raised that we're in the 'Generals' (i.e., var declaration) section
        '                        If sModelFile(n).Length > 0 Then
        '                            If bBeginBounds = True Then
        '                                ' if the variable is a decision variable in the second position
        '                                If sModelFile(n).Chars(1).ToString = "x" Then
        '                                    ' get the variable name starting in the second position
        '                                    ' while trimming any leading or trailing spaces
        '                                    sX = sModelFile(n).Trim
        '                                    ' get the variable from the solved GRB Model
        '                                    pSingleVar = GRBModel1.GetVarByName(sX)
        '                                    Try
        '                                        If iGRBStatusCode = Gurobi.GRB.Status.OPTIMAL Then
        '                                            sDecisionString = pSingleVar.Get(GRB.DoubleAttr.X)
        '                                        Else
        '                                            ' there should be more than one solution - looking for the next
        '                                            ' best solution (suboptimal)
        '                                            ' see here http://www.gurobi.com/documentation/5.0/example-tour/node114
        '                                            If GRBModel1.Get(GRB.IntAttr.SolCount) > 0 Then
        '                                                ' this sets solution number to 1 -- the second best solution found
        '                                                GRBModel1.GetEnv().Set(GRB.IntParam.SolutionNumber, 1)
        '                                                ' Xn means the optimal value for the nth solution number
        '                                                sDecisionString = pSingleVar.Get(GRB.DoubleAttr.Xn)
        '                                            End If
        '                                        End If
        '                                    Catch ex As Exception
        '                                        MsgBox("Exception raised trying to get the decision string for variable from the GRB variable using GRB API. " + ex.Message)
        '                                        Exit Sub
        '                                    End Try

        '                                    ' need to convert the decision string to an double and round it to the nearest integer
        '                                    ' this is because of an error where the decision string
        '                                    ' is sometimes a weird number like 0.999999920249585
        '                                    ' this problem was encountered Dec. 2, 2012
        '                                    Try
        '                                        dDecisionString = Convert.ToDouble(sDecisionString)
        '                                    Catch ex As Exception
        '                                        MsgBox("Could not convert decision string to type 'double' in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                        dDecisionString = 0
        '                                    End Try
        '                                    Try
        '                                        dDecisionString = Math.Round(dDecisionString, 0, MidpointRounding.AwayFromZero)
        '                                    Catch ex As Exception
        '                                        MsgBox("trouble rounding decision string to single digit integer in 'directed' analysis in GLPKShellCall. " + ex.Message)
        '                                        dDecisionString = 0
        '                                    End Try

        '                                    ' If this variable has a decision at it (and every x will have a decision, 
        '                                    '     even if it is 'do nothing' (i.e. x(i,1)) 
        '                                    If dDecisionString > 0.49 Then
        '                                        ' find out if it is 'do nothing' 
        '                                        sOptionStringArray = sX.Split(",")
        '                                        sOptionString = sOptionStringArray(1).TrimEnd(")")
        '                                        ' if it's not do nothing then get print the barrier ID and option chosen to file
        '                                        If sOptionString.Trim <> "1" Then
        '                                            sBarrierIDString = sOptionStringArray(0).TrimStart("(", "x")
        '                                            sBarrierIDString = sBarrierIDString.Trim
        '                                            sw.WriteLine(sBarrierIDString + "," + sOptionString)
        '                                            Try
        '                                                iDecisionEID = Convert.ToInt32(sBarrierIDString)
        '                                            Catch ex As Exception
        '                                                MsgBox("Trouble converting Node ID to type Integer in Gurobi Decision Output (undirected model). Converting ID to '999999'.")
        '                                                iDecisionEID = 999999
        '                                            End Try
        '                                            Try
        '                                                iDecisionNum = Convert.ToInt32(sOptionString)
        '                                            Catch ex As Exception
        '                                                MsgBox("Trouble converting OptionNum to type Integer in Gurobi Decision Output (undirected model). Converting optionnum to '99'.")
        '                                                iDecisionNum = 99
        '                                            End Try
        '                                            ' using an object named here prior to GUROBI  - same fields/attribute/parameters though
        '                                            pGLPKDecisionsObject = New GLPKDecisionsObject(dBudget, _
        '                                                                                           sGLPKTreatment + "_GRB", _
        '                                                                                           iDecisionEID, _
        '                                                                                           iDecisionNum)
        '                                            lGRBDecisionsObject2.Add(pGLPKDecisionsObject)
        '                                        End If
        '                                    End If
        '                                End If
        '                            End If
        '                        End If
        '                        If sModelFile(n).Contains("Generals") = True Then
        '                            bBeginBounds = True
        '                        End If
        '                    Next ' row in the LP model input file

        '                End If
        '                sw.Close()
        '                GRBModel1.Dispose()
        '                GRBEnv1.Dispose()

        '            End If ' GUROBI is being used for tough problems

        '        End If ' bundirected is true

        '        dBudget = dBudget + dBudgetIncrement
        '    Loop

        '    Dim sGRBResultsTableName1 As String
        '    Dim sGRBResultsTableName2 As String

        '    Dim sGRBResultsDecisionsTableName1 As String
        '    Dim sGRBResultsDecisionsTableName2 As String

        '    Dim sGLPKResultsTableName As String
        '    Dim sGLPKResultsTableName2 As String

        '    Dim sGLPKResultsDecisionsTableName As String
        '    Dim sGLPKResultsDecisionsTableName2 As String

        '    ' ===== PREP OUTPUT RESULTS TABLES =======
        '    ' ----------------
        '    ' Table 1
        '    ' Get Table1 Name
        '    If bDirected = True Then

        '        ' (only export GLPK tables if this is NOT an Sensitivity Analysis)
        '        If bSA = False Then

        '            log = New System.IO.StreamWriter(sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt", True)
        '            log.WriteLine("Writing to GLPK directed output tables.")
        '            log.Close()

        '            Try
        '                sGLPKResultsTableName = TableName("GLPKResults_" + sAnalysisCode, _
        '                                                  pFWorkspace, _
        '                                                  sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error getting GLPK Table name." + ex.Message)
        '                Exit Sub
        '            End Try
        '            Try

        '                PrepDIRResultsOutTable(sGLPKResultsTableName, _
        '                                       pFWorkspace)
        '            Catch ex As Exception
        '                MsgBox("Error creating GLPK Results Output table. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            ' Populate Table
        '            Try

        '                pTable = pFWorkspace.OpenTable(sGLPKResultsTableName)
        '                n = 0
        '                For n = 0 To lGLPKResultsObject.Count - 1
        '                    '  SinkEID (integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  GLPKSolved (boolean/binary)
        '                    '  Perc_Gap (double)
        '                    '  MAxSolTime(integer)
        '                    '  TimeUsed(double)
        '                    '  HabitatZmax(double)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = lGLPKResultsObject(n).SinkEID
        '                    pRowBuffer.Value(2) = lGLPKResultsObject(n).Treatment_Name
        '                    pRowBuffer.Value(3) = lGLPKResultsObject(n).Budget
        '                    pRowBuffer.Value(4) = Convert.ToInt16(lGLPKResultsObject(n).GLPK_Solved)
        '                    pRowBuffer.Value(5) = lGLPKResultsObject(n).Perc_Gap
        '                    pRowBuffer.Value(6) = lGLPKResultsObject(n).MaxSolTime
        '                    pRowBuffer.Value(7) = lGLPKResultsObject(n).TimeUsed
        '                    pRowBuffer.Value(8) = lGLPKResultsObject(n).Habitat_ZMax

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next

        '            Catch ex As Exception
        '                MsgBox("Problem writing to GLPK Results Table. Error Message: " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' ---------------
        '            ' Table 2
        '            ' get table name
        '            Try
        '                sGLPKResultsDecisionsTableName = TableName("GLPKResultsDecisions_DIR_" + sAnalysisCode, _
        '                                                           pFWorkspace, _
        '                                                           sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error getting GLPK results output table name. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            Try
        '                ' create table
        '                PrepGLPKResultsDecisionsOutTable(sGLPKResultsDecisionsTableName, _
        '                                                 pFWorkspace)
        '            Catch ex As Exception
        '                MsgBox("Error preparing GLPK Decisions Output table. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' populate table
        '            '  SinkEID(Of Integer)()
        '            '  Treatment (string)
        '            '  Budget (double)
        '            '  BarrierEID (Integer)
        '            '  OptionNum (Integer)
        '            Try
        '                n = 0
        '                pTable = pFWorkspace.OpenTable(sGLPKResultsDecisionsTableName)
        '                For n = 0 To lGLPKDecisionsObject.Count - 1
        '                    '  SinkEID(Of Integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  BarrierEID (Integer)
        '                    '  OptionNum (Integer)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = iSinkEID
        '                    pRowBuffer.Value(2) = lGLPKDecisionsObject(n).Treatment
        '                    pRowBuffer.Value(3) = lGLPKDecisionsObject(n).Budget
        '                    pRowBuffer.Value(4) = lGLPKDecisionsObject(n).BarrierEID
        '                    pRowBuffer.Value(5) = lGLPKDecisionsObject(n).DecisionOption

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next
        '            Catch ex As Exception
        '                MsgBox("Error writing to GLPK decisions out table. " + ex.Message)
        '                Exit Sub
        '            End Try
        '        End If ' this is a sensitivity analysis


        '        ' ======================GUROBI Table ============================
        '        ' If GUROBI analysis is being used, 
        '        ' Create and populate a Gurobi Table
        '        If bGurobiPickup = True Then


        '            ' Don't export to results tables if this is a
        '            ' sensitivity analysis (will do this only once). 
        '            ' (do export to 'decision' tables though). 
        '            If bSA = False Then

        '                ' ---------------
        '                ' Create and write to table 1
        '                ' results table
        '                Try
        '                    sGRBResultsTableName1 = TableName("GRBResults_DIR_" + sAnalysisCode, _
        '                                                      pFWorkspace, _
        '                                                      sPrefix)
        '                Catch ex As Exception
        '                    MsgBox("Error trying to get GRB Results table name. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                Try
        '                    ' create table
        '                    PrepDIRResultsOutTable(sGRBResultsTableName1, pFWorkspace)
        '                Catch ex As Exception
        '                    MsgBox("Error preparing DIR GRB results table. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                ' populate table
        '                '  SinkEID (integer)
        '                '  Treatment (string)
        '                '  Budget (double)
        '                '  GLPKSolved (boolean/binary)
        '                '  Perc_Gap (double)
        '                '  MAxSolTime(integer)
        '                '  TimeUsed(double)
        '                '  HabitatZmax(double)
        '                Try
        '                    pTable = pFWorkspace.OpenTable(sGRBResultsTableName1)
        '                    n = 0
        '                    For n = 0 To lGRBResultsObject.Count - 1
        '                        '  SinkEID (integer)
        '                        '  Treatment (string)
        '                        '  Budget (double)
        '                        '  GLPKSolved (boolean/binary)
        '                        '  Perc_Gap (double)
        '                        '  MAxSolTime(integer)
        '                        '  TimeUsed(double)
        '                        '  HabitatZmax(double)
        '                        pRowBuffer = pTable.CreateRowBuffer
        '                        pRowBuffer.Value(1) = lGRBResultsObject(n).SinkEID
        '                        pRowBuffer.Value(2) = lGRBResultsObject(n).Treatment_Name
        '                        pRowBuffer.Value(3) = lGRBResultsObject(n).Budget
        '                        pRowBuffer.Value(4) = Convert.ToInt16(lGRBResultsObject(n).GLPK_Solved)
        '                        pRowBuffer.Value(5) = lGRBResultsObject(n).Perc_Gap
        '                        pRowBuffer.Value(6) = lGRBResultsObject(n).MaxSolTime
        '                        pRowBuffer.Value(7) = lGRBResultsObject(n).TimeUsed
        '                        pRowBuffer.Value(8) = lGRBResultsObject(n).Habitat_ZMax

        '                        pCursor = pTable.Insert(True)
        '                        pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                    Next

        '                Catch ex As Exception
        '                    MsgBox("Error writing to GRB Results out table. " + ex.Message)
        '                    Exit Sub
        '                End Try
        '            Else
        '                ' if this is an SA
        '                ' loop through results (if this is for a range of budgets)
        '                ' and add them to a global variable to save for later

        '                Try
        '                    n = 0
        '                    If sTreatmentType = "A" Then
        '                        For n = 0 To lGRBResultsObject.Count - 1
        '                            m_lSA_A_Results_DIR.Add(lGRBResultsObject(n))
        '                        Next
        '                    ElseIf sTreatmentType = "P" Then
        '                        For n = 0 To lGRBResultsObject.Count - 1
        '                            m_lSA_P_Results_DIR.Add(lGRBResultsObject(n))
        '                        Next
        '                    End If

        '                Catch ex As Exception
        '                    MsgBox("trouble code wdfs22. " + ex.Message)
        '                End Try


        '            End If ' this is a sensitivity analysis

        '            ' ---------------
        '            ' Table 2
        '            ' get table name
        '            Try
        '                sGRBResultsDecisionsTableName1 = TableName("GRBResultsDecisions_DIR_" + sAnalysisCode, _
        '                                                           pFWorkspace, _
        '                                                           sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error getting the GRB Decisions table name. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                ' create table
        '                PrepGRBResultsDecisionsOutTable(sGRBResultsDecisionsTableName1, _
        '                                                pFWorkspace)
        '            Catch ex As Exception
        '                MsgBox("Error preparing the GRB Results Decisions Output table. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' populate table
        '            '  SinkEID(Of Integer)()
        '            '  Treatment (string)
        '            '  Budget (double)
        '            '  BarrierEID (Integer)
        '            '  OptionNum (Integer)
        '            Try
        '                n = 0
        '                pTable = pFWorkspace.OpenTable(sGRBResultsDecisionsTableName1)
        '                For n = 0 To lGRBDecisionsObject.Count - 1
        '                    '  SinkEID(Of Integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  BarrierEID (Integer)
        '                    '  OptionNum (Integer)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = iSinkEID
        '                    pRowBuffer.Value(2) = lGRBDecisionsObject(n).Treatment
        '                    pRowBuffer.Value(3) = lGRBDecisionsObject(n).Budget
        '                    pRowBuffer.Value(4) = lGRBDecisionsObject(n).BarrierEID
        '                    pRowBuffer.Value(5) = lGRBDecisionsObject(n).DecisionOption

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next
        '            Catch ex As Exception
        '                MsgBox("Error writing to the GRB Decisions output table for the directed model. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                If bSA = True And bZeroBudgetOverride = False Then

        '                    ' If there's a Sensitivity Analysis - keep track of common decisions... 
        '                    SA_Analysis(lGRBDecisionsObject, "Directed")

        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error calling the SA_Analysis function in the undirected section of the GLPKShellCall function. " + ex.Message)
        '            End Try

        '            ' Check if the analysis code ends in _A
        '            ' If so, save the options in a global object for use in the 
        '            ' main runanalysis sub.  
        '            ' (Sensitivity Analysis code)
        '            Try
        '                If bSA = True Then
        '                    If sAnalysisCode.Substring(sAnalysisCode.Length - 2) = "_A" Then
        '                        ' MsgBox("the substring from the analysis code: " + sAnalysisCode.Substring(sAnalysisCode.Length - 2))
        '                        m_lSABestGuessDecisionsObject_DIR = lGRBDecisionsObject
        '                        m_lSACommonDecisionsObject_DIR = lGRBDecisionsObject
        '                    End If
        '                End If
        '            Catch ex As Exception
        '                MsgBox("Error reading the Analysis code substrin in GLPKShellCall subroutine. " + ex.Message)
        '                Exit Sub
        '            End Try

        '        End If ' gurobi is used

        '        ' ==== AFTER ANALYSIS SEARCH AND REPLACE PATH TXT IN MODEL FILE ===
        '        '  search and replace 
        '        ' replac epath to Model DIR (sGLPKModelDir)  
        '        ' with c:\GunnsModel\
        '        Try
        '            If System.IO.File.Exists(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod") = True Then
        '                sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod")
        '            End If
        '        Catch ex As Exception
        '            MsgBox("Error reading the GLPk directed model file. " + ex.Message)
        '            Exit Sub
        '        End Try

        '        Try
        '            n = 0
        '            For n = 0 To sModelFile.GetUpperBound(0)
        '                If sModelFile(n).Contains(sGLPKModelDir) = True Then
        '                    sModelFile(n) = sModelFile(n).Replace(sGLPKModelDir, "C:\GunnsModel_REPLACE")
        '                End If
        '            Next
        '        Catch ex As Exception
        '            MsgBox("Error replacing lines in the directed model file. " + ex.Message)
        '            Exit Sub
        '        End Try

        '        Try
        '            ' overwrite file
        '            System.IO.File.WriteAllLines(sGLPKModelDir + "/May25_EldonsModel_FiPEx.mod", sModelFile)
        '        Catch ex As Exception
        '            MsgBox("Error trying to write to the directed model file. " + ex.Message)
        '            Exit Sub
        '        End Try
        '    End If ' bDirected model is true


        '    If bUndirected = True Then

        '        ' Only export to GLPK if this is NOT a sensitivity analysis
        '        If bSA = False Then

        '            log = New System.IO.StreamWriter(sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt", True)
        '            log.WriteLine("Writing to undirected output tables.")
        '            log.Close()
        '            Try
        '                sGLPKResultsTableName2 = TableName("GLPKResults_UNDIR_" + sAnalysisCode, _
        '                                                   pFWorkspace, _
        '                                                   sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error trying to get GLPK Results table name for undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            Try
        '                PrepUNDIRResultsOutTable(sGLPKResultsTableName2, _
        '                                         pFWorkspace)
        '                ' Populate Table
        '            Catch ex As Exception
        '                MsgBox("Error preparing undirected model results out table. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            Try
        '                pTable = pFWorkspace.OpenTable(sGLPKResultsTableName2)
        '                n = 0
        '                For n = 0 To lGLPKResultsObject2.Count - 1
        '                    '  SinkEID (integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  GLPKSolved (boolean/binary)
        '                    '  Perc_Gap (double)
        '                    '  MAxSolTime(integer)
        '                    '  TimeUsed(double)
        '                    '  HabitatZmax(double)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = lGLPKResultsObject2(n).SinkEID
        '                    pRowBuffer.Value(2) = lGLPKResultsObject2(n).Treatment_Name
        '                    pRowBuffer.Value(3) = lGLPKResultsObject2(n).Budget
        '                    pRowBuffer.Value(4) = Convert.ToInt16(lGLPKResultsObject2(n).GLPK_Solved)
        '                    pRowBuffer.Value(5) = lGLPKResultsObject2(n).Perc_Gap
        '                    pRowBuffer.Value(6) = lGLPKResultsObject2(n).MaxSolTime
        '                    pRowBuffer.Value(7) = lGLPKResultsObject2(n).TimeUsed
        '                    pRowBuffer.Value(8) = lGLPKResultsObject2(n).Habitat_ZMax
        '                    pRowBuffer.Value(9) = lGLPKResultsObject2(n).CentralBarrierEID

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next

        '            Catch ex As Exception
        '                MsgBox("Error trying to write results to GLPK results table for the undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try


        '            ' ---------------
        '            ' Table 2
        '            ' get table name
        '            Try
        '                sGLPKResultsDecisionsTableName2 = TableName("GLPKResultsDecisions_UNDIR_" + sAnalysisCode, _
        '                                                            pFWorkspace, _
        '                                                            sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error getting thable name for GLPK decisions table for the undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' create table
        '            Try
        '                PrepGLPKResultsDecisionsOutTable(sGLPKResultsDecisionsTableName2, _
        '                                                 pFWorkspace)
        '            Catch ex As Exception
        '                MsgBox("Error preparing GLPK Decisions output table for the undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            ' populate table
        '            '  SinkEID(Of Integer)()
        '            '  Treatment (string)
        '            '  Budget (double)
        '            '  BarrierEID (Integer)
        '            '  OptionNum (Integer)
        '            Try
        '                n = 0
        '                pTable = pFWorkspace.OpenTable(sGLPKResultsDecisionsTableName2)
        '                For n = 0 To lGLPKDecisionsObject2.Count - 1
        '                    '  SinkEID(Of Integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  BarrierEID (Integer)
        '                    '  OptionNum (Integer)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = iSinkEID
        '                    pRowBuffer.Value(2) = lGLPKDecisionsObject2(n).Treatment
        '                    pRowBuffer.Value(3) = lGLPKDecisionsObject2(n).Budget
        '                    pRowBuffer.Value(4) = lGLPKDecisionsObject2(n).BarrierEID
        '                    pRowBuffer.Value(5) = lGLPKDecisionsObject2(n).DecisionOption

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next

        '            Catch ex As Exception
        '                MsgBox("Error trying to write to the GLPK decisions output table for the undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try

        '        End If ' this is a sensitivity analysis


        '        ' ======================GUROBI Table ============================
        '        ' If GUROBI analysis is being used, 
        '        ' Create and populate a Gurobi Table
        '        If bGurobiPickup = True Then

        '            ' only do 'results' tables if this is not a sensitivity analysis
        '            ' (if so they are done once in a custom function). 
        '            If bSA = False Then

        '                ' ---------------
        '                ' Create and write to table 1
        '                ' results table
        '                Try
        '                    sGRBResultsTableName2 = TableName("GRBResults_UNDIR_" + sAnalysisCode, _
        '                                                      pFWorkspace, _
        '                                                      sPrefix)

        '                Catch ex As Exception
        '                    MsgBox("Error trying to get tablename for the undirected model and GRB run. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                ' create table
        '                Try
        '                    PrepUNDIRResultsOutTable(sGRBResultsTableName2, _
        '                                             pFWorkspace)
        '                Catch ex As Exception
        '                    MsgBox("Error trying to prepare GRB results output table for the undirected model.")
        '                    Exit Sub
        '                End Try

        '                ' populate table
        '                '  SinkEID (integer)
        '                '  Treatment (string)
        '                '  Budget (double)
        '                '  GLPKSolved (boolean/binary)
        '                '  Perc_Gap (double)
        '                '  MAxSolTime(integer)
        '                '  TimeUsed(double)
        '                '  HabitatZmax(double)
        '                Try
        '                    pTable = pFWorkspace.OpenTable(sGRBResultsTableName2)
        '                    n = 0
        '                    For n = 0 To lGRBResultsObject2.Count - 1
        '                        '  SinkEID (integer)
        '                        '  Treatment (string)
        '                        '  Budget (double)
        '                        '  GLPKSolved (boolean/binary)
        '                        '  Perc_Gap (double)
        '                        '  MAxSolTime(integer)
        '                        '  TimeUsed(double)
        '                        '  HabitatZmax(double)
        '                        pRowBuffer = pTable.CreateRowBuffer
        '                        pRowBuffer.Value(1) = lGRBResultsObject2(n).SinkEID
        '                        pRowBuffer.Value(2) = lGRBResultsObject2(n).Treatment_Name
        '                        pRowBuffer.Value(3) = lGRBResultsObject2(n).Budget
        '                        pRowBuffer.Value(4) = Convert.ToInt16(lGRBResultsObject2(n).GLPK_Solved)
        '                        pRowBuffer.Value(5) = lGRBResultsObject2(n).Perc_Gap
        '                        pRowBuffer.Value(6) = lGRBResultsObject2(n).MaxSolTime
        '                        pRowBuffer.Value(7) = lGRBResultsObject2(n).TimeUsed
        '                        pRowBuffer.Value(8) = lGRBResultsObject2(n).Habitat_ZMax
        '                        pRowBuffer.Value(9) = lGRBResultsObject2(n).CentralBarrierEID

        '                        pCursor = pTable.Insert(True)
        '                        pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                    Next

        '                Catch ex As Exception
        '                    MsgBox("Trouble writing to the GRB results table for the undirected model. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '            Else

        '                Try
        '                    n = 0
        '                    If sTreatmentType = "A" Then
        '                        For n = 0 To lGRBResultsObject2.Count - 1
        '                            m_lSA_A_Results_UNDIR.Add(lGRBResultsObject2(n))
        '                        Next
        '                    ElseIf sTreatmentType = "P" Then
        '                        For n = 0 To lGRBResultsObject2.Count - 1
        '                            m_lSA_P_Results_UNDIR.Add(lGRBResultsObject2(n))
        '                        Next
        '                    End If

        '                Catch ex As Exception
        '                    MsgBox("trouble code wdfs42. " + ex.Message)
        '                End Try


        '            End If ' if this is a sensitivity analysis

        '            ' ---------------
        '            ' Table 2
        '            ' get table name
        '            Try
        '                sGRBResultsDecisionsTableName2 = TableName("GRBResultsDecisions_UNDIR_" + sAnalysisCode, _
        '                                                           pFWorkspace, _
        '                                                           sPrefix)
        '            Catch ex As Exception
        '                MsgBox("Error getting table name for the decisions output table for GRB undirected model. " + ex.Message)
        '                Exit Sub
        '            End Try
        '            Try
        '                ' create table
        '                PrepGRBResultsDecisionsOutTable(sGRBResultsDecisionsTableName2, _
        '                                                pFWorkspace)
        '            Catch ex As Exception
        '                MsgBox("Error preparing decision results for the GRB undirected model output. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' populate table
        '            '  SinkEID(Of Integer)()
        '            '  Treatment (string)
        '            '  Budget (double)
        '            '  BarrierEID (Integer)
        '            '  OptionNum (Integer)
        '            Try
        '                n = 0
        '                pTable = pFWorkspace.OpenTable(sGRBResultsDecisionsTableName2)
        '                For n = 0 To lGRBDecisionsObject2.Count - 1
        '                    '  SinkEID(Of Integer)
        '                    '  Treatment (string)
        '                    '  Budget (double)
        '                    '  BarrierEID (Integer)
        '                    '  OptionNum (Integer)
        '                    pRowBuffer = pTable.CreateRowBuffer
        '                    pRowBuffer.Value(1) = iSinkEID
        '                    pRowBuffer.Value(2) = lGRBDecisionsObject2(n).Treatment
        '                    pRowBuffer.Value(3) = lGRBDecisionsObject2(n).Budget
        '                    pRowBuffer.Value(4) = lGRBDecisionsObject2(n).BarrierEID
        '                    pRowBuffer.Value(5) = lGRBDecisionsObject2(n).DecisionOption

        '                    pCursor = pTable.Insert(True)
        '                    pCursor.InsertRow(pRowBuffer)
        ' pCursor.Flush()
        '                Next

        '            Catch ex As Exception
        '                MsgBox("Error writing to the GRB decisions output table for the undirected mode. " + ex.Message)
        '                Exit Sub
        '            End Try

        '            ' Check if the analysis code ends in _A
        '            ' If so, save the options in a global object for use in the 
        '            ' main runanalysis sub.  
        '            ' (Sensitivity Analysis code)
        '            If bSA = True Then
        '                Try
        '                    If sAnalysisCode.Substring(sAnalysisCode.Length - 2) = "_A" Then
        '                        'MsgBox("the substring from the analysis code: " + sAnalysisCode.Substring(sAnalysisCode.Length - 2))
        '                        m_lSABestGuessDecisionsObject_UNDIR = lGRBDecisionsObject2
        '                        m_lSACommonDecisionsObject_UNDIR = lGRBDecisionsObject2
        '                    End If
        '                Catch ex As Exception
        '                    MsgBox("Error reading the Analysis code substrin in GLPKShellCall subroutine. " + ex.Message)
        '                    Exit Sub
        '                End Try

        '                ' using GRB only
        '                ' If Sensitivity Analysis is happening
        '                ' then Loop through both GRBDecisions Object and get common decisions
        '                If bZeroBudgetOverride = False Then
        '                    Try
        '                        SA_Analysis(lGRBDecisionsObject2, "Undirected")
        '                    Catch ex As Exception
        '                        MsgBox("Error calling the SA_Analysis function in the undirected section of the GLPKShellCall function. " + ex.Message)
        '                    End Try
        '                End If

        '            End If
        '        End If


        '        ' ===== PRINT RESULTS OBJECTS TO GDB TABLES =====

        '        ' ==== AFTER ANALYSIS SEARCH AND REPLACE PATH TXT IN MODEL FILE ===
        '        '  search and replace 
        '        ' replac epath to Model DIR (sGLPKModelDir)  
        '        ' with c:\GunnsModel\
        '        Try
        '            If System.IO.File.Exists(sGLPKModelDir + "/GreigUndirected.mod") = True Then
        '                sModelFile = System.IO.File.ReadAllLines(sGLPKModelDir + "/GreigUndirected.mod")
        '            End If
        '        Catch ex As Exception
        '            MsgBox("Error reading the undirected model after analyses are complete. " + ex.Message)
        '            Exit Sub
        '        End Try

        '        Try
        '            n = 0
        '            For n = 0 To sModelFile.GetUpperBound(0)
        '                If sModelFile(n).Contains(sGLPKModelDir) = True Then
        '                    sModelFile(n) = sModelFile(n).Replace(sGLPKModelDir, "C:\GunnsModel_REPLACE")
        '                End If
        '            Next
        '        Catch ex As Exception
        '            MsgBox("Error reading and replacing the lines in the undirected model after analysis is complete. " + ex.Message)
        '        End Try

        '        Try
        '            ' overwrite file
        '            System.IO.File.WriteAllLines(sGLPKModelDir + "/GreigUndirected.mod", sModelFile)
        '        Catch ex As Exception
        '            MsgBox("Error overwriting the undirected model after analyses are complete. " + ex.Message)
        '            Exit Sub
        '        End Try
        '    End If ' bUnDirected model is true
        '    log = New System.IO.StreamWriter(sGLPKModelDir + "/FIPEX_logfile_Optimisation.txt", True)
        '    log.WriteLine("Successfully ran optimisation with FIPEX GLPKShellCall sub.")
        '    log.Close()
    End Sub
    Private Sub SA_Analysis(ByRef lDecisionObject As List(Of GLPKDecisionsObject), _
                            ByVal sConnectivityType As String)

        ' Function purpose: 
        '                  cross reference input list
        '                  with global level list
        '                  keep in global list only matching / common decisions

        Dim TEMPlDecisionObject As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject)
        Dim lSACommonDecisionObject As List(Of GLPKDecisionsObject) = New List(Of GLPKDecisionsObject)

        If sConnectivityType = "Directed" Then
            lSACommonDecisionObject = m_lSACommonDecisionsObject_DIR

        ElseIf sConnectivityType = "Undirected" Then
            lSACommonDecisionObject = m_lSACommonDecisionsObject_UNDIR
        End If

        For i = 0 To lSACommonDecisionObject.Count - 1
            If lDecisionObject.Count > 0 Then
                For j = 0 To lDecisionObject.Count - 1
                    If lSACommonDecisionObject(i).BarrierEID = lDecisionObject(j).BarrierEID Then
                        If lSACommonDecisionObject(i).DecisionOption = lDecisionObject(j).DecisionOption Then

                            TEMPlDecisionObject.Add(lSACommonDecisionObject(i))

                        End If
                    End If
                Next
            End If
        Next

        If sConnectivityType = "Directed" Then
            m_lSACommonDecisionsObject_DIR = New List(Of GLPKDecisionsObject)
            m_lSACommonDecisionsObject_DIR = TEMPlDecisionObject
        ElseIf sConnectivityType = "Undirected" Then
            m_lSACommonDecisionsObject_UNDIR = New List(Of GLPKDecisionsObject)
            m_lSACommonDecisionsObject_UNDIR = TEMPlDecisionObject
        End If


    End Sub
    Private Sub DCIShellCall(ByVal sDCITableName As String, ByVal sConnectTabName As String, ByRef pFWorkspace As IFeatureWorkspace)
        ' ====================================================================
        ' SubRoutine:    DCI R Shell Call
        ' Author:        Greig Oldford
        ' Last Edited: August, 2020
        '
        ' Description: 
        '       Re-exports DBF as CSV
        '       calls on R from the DOS command line 
        '       DCI R models prints to a TXT file.  


        Dim sDCIModelDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sDCIModelDir"))    ' DCI model install directory
        Dim sRInstallDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sRInstallDir"))     ' R Program install directory
        Dim bDCISectional As Boolean = Convert.ToBoolean(m_FiPEx__1.pPropset.GetProperty("dcisectionalyn")) ' DCI sectional on/off?

        Dim pTable As ITable
        'Dim pCursor As ICursor
        Dim i As Integer
        Dim pTxtFactory As IWorkspaceFactory = New TextFileWorkspaceFactory

        Dim pFWorkspaceOut As IFeatureWorkspace
        Dim pWorkspaceOut As IWorkspace

        ' Get output workspace
        pWorkspaceOut = pTxtFactory.OpenFromFile(sDCIModelDir, My.ArcMap.Application.hWnd)
        pFWorkspaceOut = CType(pWorkspaceOut, IFeatureWorkspace)

        ' Check that user has write permissions in output workspace
        'Dim pc As New CheckPerm
        'pc.Permission = "Modify"
        'If Not pc.CheckPerm(sDCIModelDir) Then
        '    MsgBox("You do not have the necessary permission to create, delete, or modify files in the DCI Model installation directory." + _
        '    " Please change directory, attain necessary permissions, or uncheck DCI Output checkbox in 'Advanced Tab' of Options Menu.")
        '    Exit Sub
        'End If

        ' Check that the user currently has file permissions to write to 
        ' this directory
        Dim bPermissionCheck
        bPermissionCheck = FileWriteDeleteCheck(sDCIModelDir)
        If bPermissionCheck = False Then
            MsgBox("File / folder permission check: " & CStr(bPermissionCheck))
            MsgBox("It appears you do not have write permission to the DCI Model Directory.  Write permission to this directory is needed in order to run DCI Analysis.")
            Exit Sub
        End If

        Dim pExportOp As IExportOperation = New ExportOperation
        Dim pDataSetIn As IDataset
        Dim pDataSetOut As IDataset
        Dim pDSNameIn As IDatasetName
        Dim pDSNameOut As IDatasetName
        Dim FileToDelete As String

        ' Get the output dataset name ready.
        pDataSetOut = CType(pWorkspaceOut, IDataset)

        ' Take DBF habitat tables in geodatabase workspace
        ' and print it to the new text/csv file
        ' If there is more than one then merge it 
        ' into one table.  


        ' If it is the first table then delete other tables 
        ' and create the output table. 
        If i = 0 Then

            ' Delete the table if it already exists
            FileToDelete = sDCIModelDir + "/FIPEX_BarrierHabitatLine.csv"
            If System.IO.File.Exists(FileToDelete) = True Then
                System.IO.File.Delete(FileToDelete)

                ' need to reset the workspace so the file list will refresh and
                ' arcgis will know the file doesn't exist now.
                pWorkspaceOut = pTxtFactory.OpenFromFile(sDCIModelDir, 0)
                pFWorkspaceOut = CType(pWorkspaceOut, IFeatureWorkspace)
                pDataSetOut = CType(pWorkspaceOut, IDataset)
            End If

            pExportOp = New ExportOperation

            ' Get input table
            pTable = pFWorkspace.OpenTable(sDCITableName)

            ' Get the dataset name for the input table
            pDataSetIn = CType(pTable, IDataset)
            pDSNameIn = CType(pDataSetIn.FullName, IDatasetName)

            ' Get dataset for output table
            pDSNameOut = New TableName
            pDSNameOut.Name = "FIPEX_BarrierHabitatLine.csv"
            pDSNameOut.WorkspaceName = CType(pDataSetOut.FullName, IWorkspaceName)

            Try
                pExportOp.ExportTable(pDSNameIn, Nothing, Nothing, pDSNameOut, My.ArcMap.Application.hWnd)
            Catch ex As Exception
                MsgBox("Error trying to export DBF table to DCI Directory. " & ex.Message)
            End Try

            'Else ' if this is the second loop just ADD data from the input table to output table

        End If

        ' Convert the connectivity table
        pExportOp = New ExportOperation

        ' Input table
        pTable = pFWorkspace.OpenTable(sConnectTabName)

        ' Get the dataset name for the input table
        pDataSetIn = CType(pTable, IDataset)
        pDSNameIn = CType(pDataSetIn.FullName, IDatasetName)

        ' Delete the table if it already exists
        FileToDelete = sDCIModelDir + "/FIPEX_connectivity.csv"
        If System.IO.File.Exists(FileToDelete) = True Then
            System.IO.File.Delete(FileToDelete)
        End If

        ' Get dataset for output table and export
        pDSNameOut = New TableName
        pDSNameOut.Name = "FIPEX_connectivity.csv"
        pDSNameOut.WorkspaceName = CType(pDataSetOut.FullName, IWorkspaceName)

        Try
            pExportOp.ExportTable(pDSNameIn, Nothing, Nothing, pDSNameOut, My.ArcMap.Application.hWnd)
        Catch ex As Exception
            MsgBox("Error exporting connectivity table to DCI Directory. " & ex.Message)
        End Try

        ' navigate to the model directory and run the R program.  Pause until completed.
        ChDir(sDCIModelDir)
        If bDCISectional = True Then
            Shell(sRInstallDir + "/bin/r" + " CMD BATCH FIPEX_run_DCI_Sectional.r", AppWinStyle.Hide, True)
        Else
            Shell(sRInstallDir + "/bin/r" + " CMD BATCH FIPEX_run_DCI.r", AppWinStyle.Hide, True)
        End If



    End Sub
    Private Sub DCI_ADV2020_ShellCall(ByRef pFWorkspace As IFeatureWorkspace, _
                                      ByRef lAdv_DCI_Data_Object As List(Of Adv_DCI_Data_Object), _
                                      ByVal bDCISectional As Boolean, _
                                      ByVal bDistanceLim As Boolean, _
                                      ByVal dMaxDist As Double, _
                                      ByVal bDistanceDecay As Boolean, _
                                      ByVal sDDFunction As String)

        ' ====================================================================
        ' Exports 'advanced' connectivity table for distance decay calcs to CSV 
        ' Calls DCI via Shell

        Dim sDCIModelDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sDCIModelDir"))    ' DCI model install directory
        Dim sRInstallDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sRInstallDir"))     ' R Program install directory

        'Dim pTable As ITable
        Dim i As Integer

        Dim bPermissionCheck
        bPermissionCheck = FileWriteDeleteCheck(sDCIModelDir)
        If bPermissionCheck = False Then
            MsgBox("File / folder permission check: " & CStr(bPermissionCheck))
            MsgBox("It appears you do not have write permission to the DCI Model Directory.  Write permission to this directory is needed in order to run DCI Analysis.")
            Exit Sub
        End If


        ' ##############################################################
        ' 2020 Write FIPEX FIPEX_Advanced_DD_2020
        ' objectID
        ' Node Type
        ' BarrierID
        ' BarirerUserLabel
        ' Habt Quan
        ' Habit Quan Units
        ' BarrPerm
        ' BarrnaturalYN
        ' DownstreamNeighbEID
        ' DownstreamNeighbUserLabel
        ' DownstreamNeighbDistance
        ' DownstreamNeighbUnits
        'ObID,NodeEID,NodeLabel,HabQuantity,HabUnits,BarrierPerm,NaturalTF,DownstreamEID,DownstreamNodeLabel,DownstreamNeighDistance,DistanceUnits

        Try
            If File.Exists(sDCIModelDir & "\FIPEX_Advanced_DD_2020.csv") Then
                File.Delete(sDCIModelDir & "\FIPEX_Advanced_DD_2020.csv")
            End If
        Catch ex As Exception
            MsgBox("Issue reading / writing / deleting to FIPEX FIPEX_Advanced_DD_2020 file. Error code: d101" & ex.Message)
        End Try


        Dim sw1 As New System.IO.StreamWriter(sDCIModelDir & "\FIPEX_Advanced_DD_2020.csv")

        Try
            sw1.Write("NodeType,NodeEID,NodeLabel,HabQuantity,HabUnits,BarrierPerm,NaturalTF,DownstreamEID,DownstreamNodeLabel,DownstreamNeighDistance,DistanceUnits")
            sw1.Write(Environment.NewLine)
        Catch ex As Exception
            MsgBox("Issue  writing to FIPEX_Advanced_DD_2020 file. Error code: d102 " & ex.Message)
        End Try


        Dim sNodeType, sNodeEID, sNodeLabel, sHabQuantity, sHabQuanUnits, sBarrierPerm, sNaturalTF As String
        Dim sDownstreamEID, sDownstreamNodeLabel, sDownstreamNeighDistance, sDistanceUnits As String

        Try
            For i = 0 To lAdv_DCI_Data_Object.Count - 1
                'MsgBox(lAdv_DCI_Data_Object(i).NodeEID)
                sNodeType = Convert.ToString(lAdv_DCI_Data_Object(i).NodeType)

                sNodeEID = Convert.ToString(lAdv_DCI_Data_Object(i).NodeEID)

                'If Char.TryParse(lAdv_DCI_Data_Object(i).NodeLabel, sNodeLabel) = True Then
                sNodeLabel = Convert.ToString(lAdv_DCI_Data_Object(i).NodeLabel)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).HabQuantity, sHabQuantity) = True Then
                sHabQuantity = Convert.ToString(lAdv_DCI_Data_Object(i).HabQuantity)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).HabQuanUnits, sHabQuanUnits) = True Then
                sHabQuanUnits = Convert.ToString(lAdv_DCI_Data_Object(i).HabQuanUnits)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).BarrierPerm, sBarrierPerm) = True Then
                sBarrierPerm = Convert.ToString(lAdv_DCI_Data_Object(i).BarrierPerm)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).NaturalTF, sNaturalTF) = True Then
                sNaturalTF = Convert.ToString(lAdv_DCI_Data_Object(i).NaturalTF)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).DownstreamEID, sDownstreamEID) = True Then
                sDownstreamEID = lAdv_DCI_Data_Object(i).DownstreamEID
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).DownstreamNodeLabel, sDownstreamNodeLabel) = True Then
                sDownstreamNodeLabel = Convert.ToString(lAdv_DCI_Data_Object(i).DownstreamNodeLabel)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).DownstreamNeighDistance, sDownstreamNeighDistance) = True Then
                sDownstreamNeighDistance = Convert.ToString(lAdv_DCI_Data_Object(i).DownstreamNeighDistance)
                'End If
                'If Char.TryParse(lAdv_DCI_Data_Object(i).DistanceUnits, sDistanceUnits) Then
                sDistanceUnits = Convert.ToString(lAdv_DCI_Data_Object(i).DistanceUnits)
                'End If

                sw1.Write(sNodeType & "," & sNodeEID & "," & sNodeLabel & "," & _
                          sHabQuantity & "," & sHabQuanUnits & "," & _
                          sBarrierPerm & "," & sNaturalTF & "," & _
                          sDownstreamEID & "," & sDownstreamNodeLabel & "," & _
                          sDownstreamNeighDistance & "," & sDistanceUnits)
                sw1.Write(Environment.NewLine)
            Next
        Catch ex As Exception
            MsgBox("Issue  writing to FIPEX_Advanced_DD_2020 file. Error code: d103 " & ex.Message)
        End Try

        sw1.Close()


        ' ##############################################################
        ' 2020 Write param file to CSV for R to read for distance decay
        ' bDCISectional, bDistanceLim, dMaxDist, bDistanceDecay
        Try
            If File.Exists(sDCIModelDir & "\FIPEX_2020_Params.csv") Then
                File.Delete(sDCIModelDir & "\FIPEX_2020_Params.csv")
            End If
        Catch ex As Exception
            MsgBox("Issue reading / writing / deleting to FIPEX param file. Error code: d100 " & ex.Message)
        End Try

        Try
            Dim sw2 As New System.IO.StreamWriter(sDCIModelDir & "\FIPEX_2020_Params.csv")

            Dim sDCISectional As String = Convert.ToString(bDCISectional)
            Dim sDistanceLim As String = Convert.ToString(bDistanceLim)
            Dim sMaxDist As String = Convert.ToString(dMaxDist)
            Dim sDistanceDecay As String = Convert.ToString(bDistanceDecay)

            sw2.Write("bDCISectional,bDistanceLim,dMaxDist,bDistanceDecay, sDDFunction")
            sw2.Write(Environment.NewLine)
            sw2.Write(sDCISectional & "," _
                      & sDistanceLim & "," _
                      & sMaxDist & "," _
                      & sDistanceDecay & "," _
                      & sDDFunction
                      )
            sw2.Write(Environment.NewLine)
            sw2.Close()

        Catch ex As Exception
            MsgBox("Issue writing to FIPEX param file. Error code: d101 " & ex.Message)
        End Try


        ChDir(sDCIModelDir)
        Shell(sRInstallDir + "/bin/r" + " CMD BATCH FIPEX_run_DCI_DD_2020.r", AppWinStyle.Hide, True)


    End Sub
    Private Sub UpdateResultsDCI(ByRef iBarrierCount As Integer, ByRef dDCIp As Double, _
                                 ByRef dDCId As Double, ByRef bNaturalY As Boolean, _
                                 ByVal sOutFileName As String)

        ' ====================================================================
        ' SubRoutine:    Update Results Form With DCI
        ' Description:   read the output file from DCI Model
        ' ====================================================================

        ' Read the DCI Model Directory from the document properties
        Dim sDCIModelDir As String = Convert.ToString(m_FiPEx__1.pPropset.GetProperty("sDCIModelDir"))
        'Dim bDCISectional As Boolean = Convert.ToBoolean(m_FiPEX__1.pPropset.GetProperty("dcisectionalyn"))

        ' Read the file and return two values - line one and line two from output
        Dim ErrInfo As String

        Dim objReader As StreamReader
        Dim sLine As String
        Dim iPosit As Integer

        ''Select the DCI Output title and set it to bold.  
        'iPosit = pResultsForm.txtRichResults.Find("DCI OUTPUT")
        'pResultsForm.txtRichResults.Select(iPosit, 10)
        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Bold)

        Dim iLoopCount As Integer = 0

        Try
            If iBarrierCount > 1 Then ' TEMP 'if' here in case no barriers - otherwise model fails (always one barrier - the flag)
                objReader = New StreamReader(sDCIModelDir + "/" + sOutFileName)
                Do Until objReader.EndOfStream = True
                    iLoopCount = iLoopCount + 1
                    sLine = objReader.ReadLine()
                    ' If the first line says ERROR then write this, otherwise write nothing
                    ' because the first line of a successful DCI will say "Value" - skip.  
                    ' Using FINDME to get to the position at top of results box.  
                    If iLoopCount = 1 Then

                        If sLine = "ERROR" Then
                            'pResultsForm.txtRichResults.SelectedText = "    DCI Stats" + Environment.NewLine + "    FINDME"
                            'iPosit = pResultsForm.txtRichResults.Find("FINDME")
                            'pResultsForm.txtRichResults.Select(iPosit, 6)
                            'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
                            'pResultsForm.txtRichResults.SelectedText = sLine + Environment.NewLine + "    R DCI model encountered a problem with FIPEX tables" + _
                            'Environment.NewLine + "    FINDME"

                            dDCIp = 9999
                            dDCId = 9999

                            Exit Do
                        End If
                    ElseIf iLoopCount = 2 Then

                        '' Take selected bold DCI OUTPUT, change it to two lines including a findme string
                        'pResultsForm.txtRichResults.SelectedText = "    DCI Stats" + Environment.NewLine + "    FINDME"
                        'iPosit = pResultsForm.txtRichResults.Find("FINDME")
                        'pResultsForm.txtRichResults.Select(iPosit, 6)
                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
                        'pResultsForm.txtRichResults.SelectedText = sLine + Environment.NewLine + "    FINDME"

                        If bNaturalY = True Then
                            sLine = sLine.Remove(0, 15)
                        Else
                            sLine = sLine.Remove(0, 6)
                        End If
                        dDCIp = Convert.ToDouble(sLine)

                    ElseIf iLoopCount = 3 Then
                        'iPosit = pResultsForm.txtRichResults.Find("FINDME")
                        'pResultsForm.txtRichResults.Select(iPosit, 6)
                        'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
                        'pResultsForm.txtRichResults.SelectedText = sLine + Environment.NewLine + "    FINDME"
                        If bNaturalY = True Then
                            sLine = sLine.Remove(0, 15)
                        Else
                            sLine = sLine.Remove(0, 6)
                        End If
                        dDCId = Convert.ToDouble(sLine)

                    End If
                Loop

                objReader.Close()

                '' Clean up - get rid of final FINDME
                'iPosit = pResultsForm.txtRichResults.Find("FINDME")
                'pResultsForm.txtRichResults.Select(iPosit, 6)
                'pResultsForm.txtRichResults.SelectedText = "    "

            Else ' if there are no barriers encountered

                'pResultsForm.txtRichResults.SelectedText = "    DCI Stats" + Environment.NewLine + "    FINDME"
                'iPosit = pResultsForm.txtRichResults.Find("FINDME")
                'pResultsForm.txtRichResults.Select(iPosit, 6)
                'pResultsForm.txtRichResults.SelectionFont = New Font("Tahoma", 10, FontStyle.Italic)
                'pResultsForm.txtRichResults.SelectedText = "    DCIp = 100" + Environment.NewLine + "       DCId = 100"

                dDCId = 100
                dDCIp = 100

            End If

        Catch Ex As Exception
            ErrInfo = Ex.Message
            MsgBox("Error reading DCI output file. " + ErrInfo)

        End Try

    End Sub
    Private Function PrepConnectivityTables(ByVal pFWorkspace As IFeatureWorkspace, ByVal sConnectTabName As String) As String
        ' ==============================================
        ' Function: Prepare Connectivity Table(s)
        ' Author:   Greig(Oldford)
        ' Purpose: To prepare connectivity tables in 
        '          output workspace
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 1.1 If Connectivity tables are desired
        '   2.0 Check for table name conflicts and get name
        '   2.1 Create three fields
        '   2.2 Set first as objectID
        '   2.3 Set second as barrierID (string)
        '   2.4 Set third as downstream_barrierID (should be string)
        '
        ' # 2020 - added the user-set labels back as well as the EIDs
        '
        '  Note on field types:  Since there can be multiple barriers layers
        '  there could be a variety of ID types that will get added to the one connectivity
        '  table.  Since anything can be inserted to a text field then both fields should
        '  be string. 

        Dim pStTab As IStandaloneTable
        Dim pStTabColl As IStandaloneTableCollection
        'Dim sConnectTabName As String

        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit
        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        pStTabColl = CType(pMap, IStandaloneTableCollection)
        Dim pTable As ITable


        ' Create new Fields object
        pFields = New Fields
        pFieldsEdit = CType(pFields, IFieldsEdit)
        pFieldsEdit.FieldCount_2 = 5

        ' Create  Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "ObjectID"
        pFieldEdit.Name_2 = "ObID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID

        ' Set first field
        pFieldsEdit.Field_2(0) = pField

        ' Create Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrierFlagID assigned"
        pFieldEdit.Name_2 = "BarrierOrFlagID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55

        pFieldsEdit.Field_2(1) = pField

        'Create Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Downstream Barrier/Sink"
        pFieldEdit.Name_2 = "Downstream_Barrier"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55

        pFieldsEdit.Field_2(2) = pField

        'Create Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrierFlagID Label"
        pFieldEdit.Name_2 = "BarrierOrFlagLabel"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55

        pFieldsEdit.Field_2(3) = pField

        'Create Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Downstream Barrier/Sink User Label"
        pFieldEdit.Name_2 = "Downstream_BarrierLabel"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55

        pFieldsEdit.Field_2(4) = pField

        ' May be possible to add optional params for RDBMS behaviour
        If pFWorkspace IsNot Nothing Then
            Try
                pFWorkspace.CreateTable(sConnectTabName, pFields, Nothing, Nothing, "")
            Catch ex As Exception
                MsgBox("Problem creating connectivity table. " & ex.Message)
            End Try

            ' Add Table to Map Doc
            pTable = pFWorkspace.OpenTable(sConnectTabName)
            pStTab = New StandaloneTable
            pStTab.Table = pTable
            pStTabColl.AddStandaloneTable(pStTab)
            pMxDoc.UpdateContents()
        Else
            System.Windows.Forms.MessageBox.Show("Output workspace for DBF table must be within a Personal or File GeoDatabase", "Output workspace conflict")
        End If

        PrepConnectivityTables = sConnectTabName
    End Function
    Private Function PrepGLPKConnectivityTable(ByVal pFWorkspace As IFeatureWorkspace, ByVal sConnectTabName As String) As String
        ' ==============================================
        ' Function: Prepare GLPK Connectivity Table(s)
        '   Author: Greig(Oldford)
        '      For: Thesis   
        '  Purpose: To prepare GLPK connectivity tables in 
        '          output workspace
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 1.1 If Connectivity tables are desired
        '   2.0 Check for table name conflicts and get name
        '   2.1 Create three fields
        '   2.2 Set one as downstream_barrierID (should be string)
        '   2.3 Set second as barrierID (string)
        '   2.4 Set third as DUMMY
        '

        Dim pStTab As IStandaloneTable
        Dim pStTabColl As IStandaloneTableCollection
        'Dim sConnectTabName As String

        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit
        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        pStTabColl = CType(pMap, IStandaloneTableCollection)
        Dim pTable As ITable


        ' Create new Fields object
        pFields = New Fields
        pFieldsEdit = CType(pFields, IFieldsEdit)
        pFieldsEdit.FieldCount_2 = 3

        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BEID"
        pFieldEdit.Name_2 = "BEID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger

        ' Set first field
        pFieldsEdit.Field_2(0) = pField

        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "UpEID"
        pFieldEdit.Name_2 = "UpEID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger

        pFieldsEdit.Field_2(1) = pField

        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        ' THIS COLUMN IS SET AS AN INTEGER
        ' MAY CAUSE PROBLEMS IN THE FUTURE IF BARRIER
        ' IDs ARE STRINGS!!
        pFieldEdit.AliasName_2 = "DUMMY"
        pFieldEdit.Name_2 = "DUMMY"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger

        pFieldsEdit.Field_2(2) = pField

        ' May be possible to add optional params for RDBMS behaviour
        If pFWorkspace IsNot Nothing Then
            Try
                pFWorkspace.CreateTable(sConnectTabName, pFields, Nothing, Nothing, "")
            Catch ex As Exception
                MsgBox("Problem creating connectivity table. " & ex.Message)
            End Try

            ' Add Table to Map Doc
            pTable = pFWorkspace.OpenTable(sConnectTabName)
            pStTab = New StandaloneTable
            pStTab.Table = pTable
            pStTabColl.AddStandaloneTable(pStTab)
            pMxDoc.UpdateContents()
        Else
            System.Windows.Forms.MessageBox.Show("Output workspace for DBF table must be within a Personal or File GeoDatabase", "Output workspace conflict")
        End If

        PrepGLPKConnectivityTable = sConnectTabName
    End Function
    Private Sub PrepGLPKHabTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

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
        '   BarrierID (integer)
        '   Habitat (double)
        '   dummy

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

        pFieldsEdit.FieldCount_2 = 2
        iFields = 2

        '' ============ First Field ============
        '' Create ObjectID Field
        'pField = New Field
        'pFieldEdit = CType(pField, IFieldEdit)

        'pFieldEdit.AliasName_2 = "ObjectID" 'VB.NET
        'pFieldEdit.Name_2 = "ObID"
        'pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        ''pFieldEdit.AliasName = "ObjectID"
        ''pFieldEdit.Name = "ObID"
        ''pFieldEdit.Type = esriFieldType.esriFieldTypeOID

        'pFieldsEdit.Field_2(0) = pField 'VB.NET
        ''pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BARRIER"
        pFieldEdit.Name_2 = "BARRIER"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(0) = pField

        '' ============ Second Field ============
        'pField = New Field
        'pFieldEdit = CType(pField, IFieldEdit)

        'pFieldEdit.AliasName_2 = "DUMMY4"
        'pFieldEdit.Name_2 = "DUMMY4"

        'pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        'pFieldEdit.Scale_2 = 2
        'pFieldsEdit.Field_2(1) = pField


        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Habitat Quantity Field"
        pFieldEdit.Name_2 = "HABITAT"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(1) = pField

        '' ============ Second Field ============
        'pField = New Field
        'pFieldEdit = CType(pField, IFieldEdit)

        'pFieldEdit.AliasName_2 = "Habitat Quantity Field"
        'pFieldEdit.Name_2 = "HABITAT"

        'pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        'pFieldsEdit.Field_2(1) = pField

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
    Private Sub PrepGLPKOptionsTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with GLPK Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '   BARRIER (Integer)
        '   OPTION (double)
        '   PERM (double)
        '   COST (double)

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

        pFieldsEdit.FieldCount_2 = 4
        iFields = 4

        '' ============ First Field ============
        '' Create ObjectID Field
        'pField = New Field
        'pFieldEdit = CType(pField, IFieldEdit)

        'pFieldEdit.AliasName_2 = "ObjectID" 'VB.NET
        'pFieldEdit.Name_2 = "ObID"
        'pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        ''pFieldEdit.AliasName = "ObjectID"
        ''pFieldEdit.Name = "ObID"
        ''pFieldEdit.Type = esriFieldType.esriFieldTypeOID

        'pFieldsEdit.Field_2(0) = pField 'VB.NET
        ''pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BARRIER"
        pFieldEdit.Name_2 = "BARRIER"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(0) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "OPTION1"
        pFieldEdit.Name_2 = "OPTION1"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(1) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "PERM"
        pFieldEdit.Name_2 = "PERM"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(2) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "COST"
        pFieldEdit.Name_2 = "COST"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(3) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepDIRResultsOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with DCI Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  GLPKSolved (boolean/binary)
        '  Perc_Gap (double)
        '  MAxSolTime(integer)
        '  TimeUsed(double)
        '  HabitatZmax(double)


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

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "SinkEID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(1) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment"
        pFieldEdit.Name_2 = "Treatment"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldsEdit.Field_2(2) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "GLPKSolved"
        pFieldEdit.Name_2 = "GLPKSolved"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
        pFieldsEdit.Field_2(4) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "PercGap"
        pFieldEdit.Name_2 = "PercGap"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(5) = pField

        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "MaxSolTime"
        pFieldEdit.Name_2 = "MaxSolTime"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(6) = pField

        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "TimeUsed"
        pFieldEdit.Name_2 = "TimeUsed"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(7) = pField

        ' ============ Eighth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "HabitatZMax"
        pFieldEdit.Name_2 = "HabitatZMax"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(8) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepUNDIRResultsOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with DCI Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  GLPKSolved (boolean/binary)
        '  Perc_Gap (double)
        '  MAxSolTime(integer)
        '  TimeUsed(double)
        '  HabitatZmax(double)
        '  centralBarrierEID (integer)


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

        pFieldsEdit.FieldCount_2 = 10
        iFields = 10

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "SinkEID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(1) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment"
        pFieldEdit.Name_2 = "Treatment"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldsEdit.Field_2(2) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "GLPKSolved"
        pFieldEdit.Name_2 = "GLPKSolved"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
        pFieldsEdit.Field_2(4) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "PercGap"
        pFieldEdit.Name_2 = "PercGap"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(5) = pField

        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "MaxSolTime"
        pFieldEdit.Name_2 = "MaxSolTime"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(6) = pField

        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "TimeUsed"
        pFieldEdit.Name_2 = "TimeUsed"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(7) = pField

        ' ============ Eighth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "HabitatZMax"
        pFieldEdit.Name_2 = "HabitatZMax"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(8) = pField

        ' ============= Ninth Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "CentralBarrierEID"
        pFieldEdit.Name_2 = "CentralBarrierEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(9) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepGLPKResultsDecisionsOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with DCI Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  BarrierEID (Integer)
        '  OptionNum (Integer)


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

        pFieldsEdit.FieldCount_2 = 6
        iFields = 6

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "SinkEID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(1) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment"
        pFieldEdit.Name_2 = "Treatment"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldsEdit.Field_2(2) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrierEID"
        pFieldEdit.Name_2 = "BarrierEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(4) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "OptionNum"
        pFieldEdit.Name_2 = "OptionNum"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(5) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepGRBResultsDecisionsOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with DCI Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '  SinkEID (integer)
        '  Treatment (string)
        '  Budget (double)
        '  BarrierEID (Integer)
        '  OptionNum (Integer)


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

        pFieldsEdit.FieldCount_2 = 6
        iFields = 6

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "SinkEID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(1) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment"
        pFieldEdit.Name_2 = "Treatment"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldsEdit.Field_2(2) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldEdit.Scale_2 = 2
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrierEID"
        pFieldEdit.Name_2 = "BarrierEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(4) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "OptionNum"
        pFieldEdit.Name_2 = "OptionNum"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldsEdit.Field_2(5) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepDCIOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace for use with DCI Model
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        ' Table fields:
        '   ObjectID (autoinc)
        '   BarrierID (string 55)
        '   Quantity (double)
        '   BarrierPerm (double)
        '   BarrierYN (string 55)

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

        pFieldsEdit.FieldCount_2 = 5
        iFields = 5

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= Second Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "BarrierID"
        pFieldEdit.Name_2 = "BarrierID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(1) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Habitat Quantity Field"
        pFieldEdit.Name_2 = "Quantity"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(2) = pField

        ' =========== Fourth Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Barrier Permeability Field"
        pFieldEdit.Name_2 = "BarrierPerm"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(3) = pField

        ' =========== Fifth Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Barrier Natural TF Field"
        pFieldEdit.Name_2 = "NaturalYN"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55

        pFieldsEdit.Field_2(4) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepAdvanced2020DCIOutTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' Note I eliminated this table, which becomes huge, from DBF, and just skipped
        ' to printing this table as CSV only to R Model folder
        ' ==============================================
        ' 2020 - DCI Output Table that includes columns for downstream neighbour 
        ' and distance to downstream neighbout


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

        pFieldsEdit.FieldCount_2 = 11
        iFields = 11

        ' objectID
        ' BarrierID
        ' BarirerUserLabel
        ' Habt Quan
        ' Habit Quan Units
        ' BarrPerm
        ' BarrnaturalYN
        ' DownstreamNeighbEID
        ' DownstreamNeighbUserLabel
        ' DownstreamNeighbDistance
        ' DownstreamNeighbUnits

        ' ============ Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "ObjectID"
        pFieldEdit.Name_2 = "ObID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        pFieldsEdit.Field_2(0) = pField

        ' ============= Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Node EID"
        pFieldEdit.Name_2 = "NodeEID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(1) = pField

        ' ============= Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Node User Label"
        pFieldEdit.Name_2 = "NodeLabel"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(2) = pField

        ' ============Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Habitat Quantity"
        pFieldEdit.Name_2 = "Quantity"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(3) = pField

        ' ============  Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Habitat Units"
        pFieldEdit.Name_2 = "HabUnits"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(4) = pField

        ' =========== Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Barrier Permeability"
        pFieldEdit.Name_2 = "BarrierPerm"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(5) = pField

        ' =========== Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Barrier Natural TF"
        pFieldEdit.Name_2 = "NaturalTF"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(6) = pField

        ' =========== Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Downstream Neighbour EID"
        pFieldEdit.Name_2 = "DownstreamEID"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(7) = pField

        ' ============= Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Downstream Neighbour User Label"
        pFieldEdit.Name_2 = "DownstreamNodeLabel"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(8) = pField

        ' =========== Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Downstream Neighbour Distance"
        pFieldEdit.Name_2 = "DownstreamNeighDistance"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(9) = pField

        ' =========== Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Distance Units"
        pFieldEdit.Name_2 = "DistanceUnits"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(10) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepHabitatTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        '
        ' Creates the following fields:
        '       ObjectID
        '       Sink      (string - 55)
        '       SinkEID   (integer - long)
        '       Label   (string - 55)   2020 Barrier   (string - 55)
        '       ElementID(integer - long)  2020 BarrierEID(integer - long)
        '       Layer     (string - 255)
        '       Direction (string - 12)
        '       Trace Type(string - 20)
        '       Class     (string - 85)
        '       Quantity  (double)
        '       Measure   (string - 30)

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

        pFieldsEdit.FieldCount_2 = 13
        iFields = 13

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= Second Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Sink ID"
        pFieldEdit.Name_2 = "SinkID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(1) = pField

        ' ============= Third Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Sink EID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(2) = pField

        ' ============= Fourth Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Label"
        pFieldEdit.Name_2 = "Label"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(3) = pField

        ' ============= Fifth Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Element ID"
        pFieldEdit.Name_2 = "ElementID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(4) = pField

        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Node Type"
        pFieldEdit.Name_2 = "nType"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 30
        pFieldsEdit.Field_2(5) = pField

        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Habitat Layer"
        pFieldEdit.Name_2 = "Layer"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 255
        pFieldsEdit.Field_2(6) = pField

        ' ============ Eigth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Direction"
        pFieldEdit.Name_2 = "Direction"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 12
        pFieldsEdit.Field_2(7) = pField

        ' ============ Ninth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Trace Type"
        pFieldEdit.Name_2 = "TraceType"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 20
        pFieldsEdit.Field_2(8) = pField

        ' ============ tenth Field ============
        ' THIS COLUMN IS SET AS A STRING IN CASE HAB CLASSES
        ' ARE DIVIDED INTO PRIM, SECOND, TERT etc.
        ' This may cause problems when populating from
        ' a numeric field.

        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Habitat Class"
        pFieldEdit.Name_2 = "HabClass"
        'End If
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 85
        pFieldsEdit.Field_2(9) = pField

        ' ============ Eleventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)
        pFieldEdit.AliasName_2 = "Habitat Class"
        pFieldEdit.Name_2 = "HabClassField"
        'End If
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 85
        pFieldsEdit.Field_2(10) = pField

        ' =========== Twelfth Field ===========
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Habitat Quantity"
        pFieldEdit.Name_2 = "Quantity"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(11) = pField

        ' =========== Thirteenth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "UnitMeasure"
        pFieldEdit.Name_2 = "Units"
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 30
        pFieldsEdit.Field_2(12) = pField


        ' MsgBox "Going to create table in workspace named - & sTableName
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
    Private Sub PrepMetricTable(ByVal sTable As String, ByVal pFWSpace As IFeatureWorkspace)

        ' ==============================================
        ' Purpose: To prepare dbf tables in output workspace
        '
        ' PROCESS LOGIC:
        ' 1.0 Set the table collection to the map document
        ' 2.0 Create a fields collection with name and type of each field
        ' 3.0 Create the tables in the output workspace. 
        '
        '
        ' Creates the following fields:
        '       ObjectID
        '       SinkID    (string - 55)
        '       SinkEID   (integer - long)
        '       Label      (string -55) 2020 
        '       ElementID (integer - long) 2020
        '       MetricName (string - 55)
        '       Metric (double)

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

        pFieldsEdit.FieldCount_2 = 8
        iFields = 8

        ' ============ First Field ============
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
        'pFieldsEdit.Field(0) = pField

        ' ============= Second Field ============
        ' Create Barrier ID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Sink ID"
        pFieldEdit.Name_2 = "SinkID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(1) = pField

        ' ============= Third Field ============
        ' Create Barrier ID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Sink EID"
        pFieldEdit.Name_2 = "SinkEID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(2) = pField

        ' ============= Fourth Field ============
        ' Create Barrier ID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Label"
        pFieldEdit.Name_2 = "Label"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(3) = pField

        ' ============= Fifth Field ============
        ' Create Barrier ID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Element ID"
        pFieldEdit.Name_2 = "ElementID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        pFieldEdit.Length_2 = 5
        pFieldsEdit.Field_2(4) = pField

        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Node Type"
        pFieldEdit.Name_2 = "NodeType"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 30
        pFieldsEdit.Field_2(5) = pField

        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Metric Name"
        pFieldEdit.Name_2 = "MetricName"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        pFieldEdit.Length_2 = 55
        pFieldsEdit.Field_2(6) = pField

        ' ============ Eigth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Metric"
        pFieldEdit.Name_2 = "Metric"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        pFieldsEdit.Field_2(7) = pField

        ' MsgBox "Going to create table in workspace named - & sTableName
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
                    'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
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

                                'MsgBox(" m_aExcldVls(x): " + CStr(m_aExcldVls(x)))
                                'MsgBox("Is vVal Null? " + CStr(IsNull(vVal)))
                                'MsgBox("What kind of layer is feature from? " + CStr(pFeatureLayer.FeatureClass.ShapeType))
                                '

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
    'Public Function GetFileContents(ByVal FullPath As String, _
    '   Optional ByRef ErrInfo As String = "") As String

    '    Dim sFirstLine, sSecondLine As String
    '    Dim objReader As StreamReader
    '    Try

    '        objReader = New StreamReader(FullPath)
    '        sFirstLine = objReader.ReadLine()
    '        sSecondLine = objReader.ReadLine

    '        objReader.Close()
    '        'GetFileContents.m_LineOne = sFirstLine
    '        'GetFileContents.m_LineTwo = sSecondLine
    '    Catch Ex As Exception
    '        ErrInfo = Ex.Message
    '        MsgBox("Error reading DCI output file. " + ErrInfo)

    '    End Try
    'End Function

    'Public Structure DCIFileStrings
    '    'Structure suggested here: http://www.codewidgets.com/product.aspx?key=55
    '    Public m_LineOne As String
    '    Public m_LineTwo As String
    'End Structure
    Public Function FileWriteDeleteCheck(ByVal sDCIOutputDir As String) As Boolean

        ' Checks user permissions at runtime to read/write/delete to output directory. 
        ' Need to take the previous out.txt and put a default error string in there in case
        ' the DCI model fails.  

        Dim FILE_NAME As String = "out.txt"
        Try
            If File.Exists(sDCIOutputDir + "\" + FILE_NAME) Then
                'MsgBox("tempmsg: this is the file name tested: " + sDCIOutputDir + FILE_NAME)
                'MsgBox("Output directory already contains the write test file. It will be deleted")
                File.Delete(sDCIOutputDir + "\" + FILE_NAME)
            End If
        Catch e As Exception
            MsgBox("You require read/write/delete permissions on the DCI Model Directory. " & e.Message)
        End Try

        Try
            Dim path As String = sDCIOutputDir + "\" + FILE_NAME
            Dim sw As StreamWriter = File.CreateText(path)
            sw.Write("ERROR")
            sw.Close()

            ' Ensure that the target does not exist.
            'File.Delete(path)

            Return True

        Catch e As Exception
            MsgBox("The out.txt file in the DCI Model directory could not be written to.  The following exception was found: " & e.Message)
            Return False
        End Try

    End Function
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
        pTraceFlowSolver.TraceIndeterminateFlow = True
        pTraceTasks.TraceIndeterminateFlow = True

        ' pass the traceFlowSolver object back to the network solver
        TraceFlowSolverSetup = pTraceFlowSolver

    End Function
    Public Function GetSourcesBranchJunctions()

    End Function

    Private Function GetNaturalYN(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As String
        'returns whether the barrier is natural or not based on user-declared field in attribute table
        'if none found, assumes NOT natural (returns "F")

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
    Private Function GetBarrierPerm(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As Double
        'returns permeability drawn from user-declared field in attribute table
        'if not it returns a zero permeability
        '
        ' If there's a barrierperm field provided, 
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
                                sBarrierPermField = lBarrierIDs.Item(k).PermField
                                If pFields.FindField(sBarrierPermField) <> -1 Then
                                    Try
                                        dOutValue = Convert.ToDouble(pFeature.Value(pFields.FindField(sBarrierPermField)))
                                        bCheck = True
                                    Catch ex As Exception
                                        MsgBox("The Permeability Value in the " & CStr(pFLayer.Name) & " was not convertible" _
                                        & " to type 'double'. Please check the attribute table. Assuming permeability = zero. " & ex.Message.ToString)
                                        ' If there's a null value or field value can't be converted, assume zero
                                        dOutValue = 0.0
                                    End Try

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
    Private Function GetBarrierType(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As String
        ' returns a type category drawn from attribute table such as "dam" or "culvert"
        ' This section checks whether there is a BarrierType Field
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
        Dim sBarrierTypeField As String
        Dim iBarrierIds As Integer
        Dim sOutValue As String

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
                                sBarrierTypeField = lBarrierIDs.Item(k).LayerType
                                If pFields.FindField(sBarrierTypeField) <> -1 Then
                                    Try
                                        sOutValue = Convert.ToString(pFeature.Value(pFields.FindField(sBarrierTypeField)))
                                        bCheck = True
                                    Catch ex As Exception
                                        sOutValue = "not found"
                                    End Try

                                End If ' names match
                            End If
                        Next ' barrierIDField
                        If bCheck = False Then ' names don't match - assume it's a 
                            sOutValue = "not found"
                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        GetBarrierType = sOutValue

    End Function
    Private Function GetGLPKOptions(ByVal iFCID As Integer, _
                                    ByVal iFID As Integer, _
                                    ByVal lBarrierIDs As List(Of BarrierIDObj), _
                                    ByVal iEID As Integer, _
                                    ByVal sBarrierType As String) As List(Of GLPKOptionsObject)

        ' =============== FLAG ON POINT WITH Perm? =========================
        ' This function returns a list of the GLPK Options for this barrier
        ' which will be added to the running list in the main sub.
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
        Dim j, k, m As Integer

        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer
        Dim pFeatureClass As IFeatureClass
        Dim pFeature As IFeature
        Dim pFields As IFields
        Dim iBarrierIds As Integer
        Dim dOptionPermValue, dOptionCostValue As Double
        Dim lGLPKOptionsListTEMP As List(Of GLPKOptionsObject) = New List(Of GLPKOptionsObject) ' for TEMP options stats for GLPK
        Dim pGLPKOptionsObject As GLPKOptionsObject = New GLPKOptionsObject(Nothing, _
                                                                            Nothing, _
                                                                            Nothing, _
                                                                            Nothing, _
                                                                            Nothing)
        Dim sBarrierOptionPermField, sBarrierOptionCostField As String

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

                        'For k = 0 To iBarrierIds - 1
                        '    If lBarrierIDs.Item(k).Layer = pFLayer.Name Then
                        '        sBarrierPermField = lBarrierIDs.Item(k).PermField
                        '        If pFields.FindField(sBarrierPermField) <> -1 Then
                        '            dOutValue = Convert.ToDouble(pFeature.Value(pFields.FindField(sBarrierPermField)))
                        '            bCheck = True
                        '        End If ' names match
                        '    End If
                        'Next ' barrierIDField

                        ' TEMP SEARCH FOR SPECIFIC FIELDS REPLACES ABOVE
                        For k = 0 To iBarrierIds - 1
                            If lBarrierIDs.Item(k).Layer = pFLayer.Name Then

                                ' for each of the four maximum options to consider
                                ' get the 
                                m = 1
                                For m = 1 To 4
                                    sBarrierOptionPermField = "option" + Convert.ToString(m) + "permafter"
                                    sBarrierOptionCostField = "option" + Convert.ToString(m) + "cost"
                                    If pFields.FindField(sBarrierOptionPermField) <> -1 Then
                                        Try
                                            dOptionPermValue = Convert.ToDouble(pFeature.Value(pFields.FindField(sBarrierOptionPermField)))

                                        Catch ex As Exception
                                            'System.Windows.Forms.MessageBox.Show("The permeability value in the Barriers Layer is Null. Please provide a value. Default value will be used otherwise (0).  " + ex.Message, "")
                                            dOptionPermValue = 9999
                                        End Try

                                        If pFields.FindField(sBarrierOptionCostField) <> -1 Then
                                            Try
                                                dOptionCostValue = Convert.ToDouble(pFeature.Value(pFields.FindField(sBarrierOptionCostField)))

                                            Catch ex As Exception
                                                ' System.Windows.Forms.MessageBox.Show("The cost value in the Barriers Layer is Null. Please provide a value. " + ex.Message, "")
                                                dOptionCostValue = 9999
                                            End Try
                                            bCheck = True
                                            pGLPKOptionsObject = New GLPKOptionsObject(iEID, _
                                                                                       m + 1, _
                                                                                       dOptionPermValue, _
                                                                                       dOptionCostValue, _
                                                                                       sBarrierType)
                                            lGLPKOptionsListTEMP.Add(pGLPKOptionsObject)
                                        Else
                                            ' MsgBox("The Option Cost Value was not found for the " + _
                                            'Convert.ToString(m) + " Option field, but the Permeability field was found")
                                        End If
                                    End If
                                    'm += 1
                                Next 'm (option)

                            End If
                        Next ' barrierIDField

                        If bCheck = False Then ' names don't match - assume it's a 
                            'MsgBox("No options or permeability fields were found in the barriers layer.  No projects will be considered for this barrier/layer except 'do nothing.'")

                        End If
                    End If ' feature class in map equals feature class of flag
                End If
            End If
        Next

        GetGLPKOptions = lGLPKOptionsListTEMP

    End Function
    Private Function GetBarrierID(ByVal iFCID As Integer, ByVal iFID As Integer, ByVal lBarrierIDs As List(Of BarrierIDObj)) As IDandType
        ' =============== FLAG ON POINT WITH BNUMBER? =========================
        'returns user-set ID or Object ID
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
                                    Catch ex As Exception
                                        MsgBox("Could not convert barrier ID to string.")
                                        sOutID = "Unknown"
                                    End Try
                                    bCheck = True
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
        Else
            pFWorkspace = Nothing
            GetWorkspace = pFWorkspace
            Exit Function
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
            iEID = pJuncFlags.Next()  ' get the EID of flag
            m = 0
            pBarriersList.Reset()

            ' For each barrier
            For m = 0 To pBarriersList.Count - 1
                'If endEID = pOriginalBarriersList(m) Then 'VB.NET
                If iEID = pBarriersList.Next() Then
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
            iEID = pConnected.Next() ' get the EID of flag
            m = 0
            pUpConnected.Reset()

            ' For each upstream element
            For m = 0 To pUpConnected.Count - 1
                'If endEID = pOriginalBarriersList(m) Then 'VB.NET
                If iEID = pUpConnected.Next() Then
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

    ' This function returns the kind of workspace the geometric network is in
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





End Class
