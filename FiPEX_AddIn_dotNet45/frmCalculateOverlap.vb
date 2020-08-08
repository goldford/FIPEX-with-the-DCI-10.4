Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesOleDB
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class frmCalculateOverlap
    Private m_app As ESRI.ArcGIS.Framework.IApplication
    Private pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Private m_FiPEx As FishPassageExtension
    Private m_UtilityNetworkAnalysisExt As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt
    Private m_pOutGxDatabase As IGxDatabase
    Private m_sCategory2 As String
    Private m_sCategory1 As String
    Private m_pFWorkspace1 As IFeatureWorkspace
    Private m_pFWorkspace2 As IFeatureWorkspace
    Private m_sTableName1 As String
    Private m_sTableName2 As String
    ' Private progressThread As Threading.Thread
    Private m_sProgressString As String
    Private m_iProgressPercent As Integer

    Public Function Form_Initialize(ByVal m_application As ESRI.ArcGIS.Framework.IApplication) As Boolean

        'Dim pMap As ESRI.ArcGIS.Carto.IMap
        'Dim pLayer As ESRI.ArcGIS.Carto.ILayer
        'Dim pFLayer As ESRI.ArcGIS.Carto.IFeatureLayer2
        'Dim pGeoDataset As ESRI.ArcGIS.Geodatabase.IGeoDataset

        Try
            m_app = m_application
            pMxDoc = CType(m_app.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
            m_FiPEx = FishPassageExtension.GetExtension

            ' Obtain a reference to Utility Network Analysis Ext in the current Doc
            'Dim pUID As New ESRI.ArcGIS.esriSystem.UID
            'pUID.Value = "{98528F9B-B971-11D2-BABD-00C04FA33C20}"

            'Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension = m_application.FindExtensionByCLSID(pUID)
            'm_UtilityNetworkAnalysisExt = CType(pExtension, IUtilityNetworkAnalysisExt)
            m_UtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt


            ' Obtain a reference to the DFOBarriersAnalysis Extension
            'Dim pUID2 As New ESRI.ArcGIS.esriSystem.UID
            'pUID2.Value = "FiPEx.DFOBarriersAnalysisExtension"

            'pExtension = m_application.FindExtensionByCLSID(pUID2)

            'm_DFOExt = CType(pExtension, FiPEx.DFOBarriersAnalysisExtension)

            Return True

        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString, "Initialize Form")
            Return False
        End Try
    End Function

    Private Sub backgroundworker3_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker3.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
        lblProgress.Text = e.UserState.ToString
    End Sub

    Private Sub backgroundworker3_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted

    End Sub


    Private Sub BackgroundWorker3_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Do Until e.Cancel = True
            If BackgroundWorker3.CancellationPending = True Then
                e.Cancel = True
            End If
            System.Threading.Thread.Sleep(200)
        Loop

    End Sub
  
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        BackgroundWorker3.Dispose()
        Me.Close()
    End Sub

    Private Sub exitsubroutine()
        'backgroundworker3.CancelAsync()
        BackgroundWorker3.ReportProgress(0)
        lblProgress.Text = m_sProgressString
        ' BackgroundWorker3.Dispose() 'read that dispose is a relic and does nothing
        BackgroundWorker3.CancelAsync()

    End Sub
    Private Sub cmdCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCalculate.Click

        ' Me.CheckForIllegalCrossThreadCalls = False
        'progressThread = New Threading.Thread(AddressOf updateProgress)
        Call CalculateOverlap()

        BackgroundWorker3.CancelAsync()

    End Sub

    Private Sub cmdBrowseTab2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseTab2.Click
        Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterTables
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Table 1"
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
            txtTable2.Text = ""
            Exit Sub
        End If

        'm_sCategory = sCategory


        Dim sFullTableName As String
        sFullTableName = pGxObject.FullName
        Dim sTableDirectory As String
        Dim iLastBackSlash As Integer
        iLastBackSlash = sFullTableName.LastIndexOf("\")
        sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        m_sTableName2 = sFullTableName.Substring(iLastBackSlash + 1)

        'Dim pStandaloneTable As IStandaloneTable
        'Dim pTable As ITable

        'Dim pTxtFactory As IWorkspaceFactory = New TextFileWorkspaceFactory
        'Dim pFWorkspace As IFeatureWorkspace
        'Dim pWorkspace As IWorkspace
        '' Get output workspace
        'Try
        '    pWorkspace = pTxtFactory.OpenFromFile(sTableDirectory, My.ArcMap.Application.hWnd)
        'Catch ex As Exception
        '    MsgBox("Failed getting workspace for selected table. " & ex.Message)
        'End Try

        'pFWorkspace = CType(pWorkspace, IFeatureWorkspace)

        'm_pFWorkspace2 = pWorkspace

        m_pFWorkspace2 = GetWorkspace(sTableDirectory)

        If pGxObject Is Nothing Then
            MsgBox("You must select an input table")
            Exit Sub
        Else
            txtTable2.Text = sFullTableName
        End If

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

    Private Function TableName(ByVal pFWorkspace As IFeatureWorkspace, ByRef sPrefix As String) As String
        ' Function:    Table Name
        ' Created By:  Greig Oldford
        ' Update Date: June 18, 2012
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
        Dim bNumCheck As Boolean = False
        Dim bLengthCheck As Boolean = True
        Dim bCharCheck As Boolean = True
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
            bCharCheck = reg.IsMatch(sPrefix) 'For VB.NET
            ' Check if prefix starts with numbers (not allowed in ArcGIS for table names)
            bNumCheck = reg2.IsMatch(sPrefix)
            ' Check if table name is too long (>55)
            sName = sPrefix
            ' Check that length of table name isn't too long
            If Len(sName) > 55 Then
                bLengthCheck = False
            Else
                bLengthCheck = True
            End If

            If bCharCheck = False Or bNumCheck = True Or bLengthCheck = False Then

                title = "Problem with table name"
                If bCharCheck = False Then
                    sCharMsg = "Invalid character in table name. Table name may only include spaces, numbers, or _"
                Else
                    sCharMsg = "OK"
                End If
                If bNumCheck = True Then
                    sNumMsg = "Table names cannot begin with numbers in ArcGIS."
                Else
                    sNumMsg = "OK"
                End If
                If bLengthCheck = False Then
                    sLengthMsg = "Table name is too long.  Name cannot exceed 55 characters in ArcGIS"
                Else
                    sLengthMsg = "OK"
                End If

                message = "Invalid table name, please enter a new prefix." & Environment.NewLine & _
                "Characters: " & sCharMsg & Environment.NewLine _
                & "First Character: " & sNumMsg & Environment.NewLine _
                & "Length: " & sLengthMsg

                defaultValue = ""
                myTableName = InputBox(message, title, defaultValue)

                sName = myTableName.ToString
                If sName = "" Then
                    Return "Cancel"
                End If

                bValidate = False
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
    Private Sub cmdBrowseTab1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseTab1.Click
        Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterTables
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Table 1"
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
            txtTable1.Text = ""
            Exit Sub
        End If

        Dim sFullTableName As String
        sFullTableName = pGxObject.FullName
        Dim sTableDirectory As String
        Dim iLastBackSlash As Integer
        iLastBackSlash = sFullTableName.LastIndexOf("\")
        sTableDirectory = sFullTableName.Remove(iLastBackSlash)
        m_sTableName1 = sFullTableName.Substring(iLastBackSlash + 1)

        'Dim pStandaloneTable As IStandaloneTable
        'Dim pTable As ITable

        'Dim pTxtFactory As IWorkspaceFactory = New TextFileWorkspaceFactory
        'Dim pFWorkspace As IFeatureWorkspace
        'Dim pWorkspace As IWorkspace
        '' Get output workspace
        'Try
        '    pWorkspace = pTxtFactory.OpenFromFile(sTableDirectory, My.ArcMap.Application.hWnd)
        'Catch ex As Exception
        '    MsgBox("Failed getting workspace for selected table. " & ex.Message)
        'End Try

        'pFWorkspace = CType(pWorkspace, IFeatureWorkspace)

        'm_pFWorkspace1 = pWorkspace

        m_pFWorkspace1 = GetWorkspace(sTableDirectory)

        If pGxObject Is Nothing Then
            MsgBox("You must select an input table")
            Exit Sub
        Else
            txtTable1.Text = pGxObject.FullName
        End If


    End Sub


    Private Sub cmdGDBOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGDBOut.Click
        Dim pGxDialog As IGxDialog
        Dim pDstFilter As IGxObjectFilter
        Dim pEnumGx As IEnumGxObject

        pDstFilter = New GxFilterFileGeodatabases
        pGxDialog = New GxDialog

        Dim pFilterCol As IGxObjectFilterCollection
        pFilterCol = CType(pGxDialog, IGxObjectFilterCollection)

        pFilterCol.AddFilter(pDstFilter, True)

        pGxDialog.Title = "Browse for Geodatabase"
        pGxDialog.AllowMultiSelect = False

        pGxDialog.DoModalOpen(0, pEnumGx)

        If pEnumGx Is Nothing Then Exit Sub

        'Me.ZOrder(0) 'sendtoback?
        pEnumGx.Reset()

        Dim pGxDatabase As IGxDatabase
        Dim pGxObject As IGxObject = pEnumGx.Next
        pGxDatabase = CType(pGxObject, IGxDatabase)

        If pGxDatabase Is Nothing Then
            MsgBox("You must select a geodatabase for output.")
            txtGDBOut.Text = ""
            Exit Sub
        Else
            txtGDBOut.Text = pGxDatabase.Workspace.PathName
        End If

        m_pOutGxDatabase = pGxDatabase

    End Sub

    Private Class GLPKDecisionsObjectPredicateBudgets
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        'Private _BarrierEID As Integer
        Private _Budget As Double

        'Public Sub New(ByVal barriereid As Integer, ByVal budget As Double)
        Public Sub New(ByVal budget As Double)

            'Me._BarrierEID = barriereid
            Me._Budget = budget
        End Sub

        Public Function CompareBudgets(ByVal obj As GLPKDecisionsObject) As Boolean
            'Return (_BarrierEID = obj.BarrierEID And _Budget = obj.Budget)
            Return (_Budget = obj.Budget)
        End Function

    End Class
    Private Class GLPKDecisionsObjectPredicateBudgetsANDOptionNums
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        'Private _BarrierEID As Integer
        Private _Budget As Double
        Private _OptionNum As Integer

        'Public Sub New(ByVal barriereid As Integer, ByVal budget As Double)
        Public Sub New(ByVal budget As Double, ByVal optionnum As Integer)

            'Me._BarrierEID = barriereid
            Me._Budget = budget
            Me._OptionNum = optionnum
        End Sub

        Public Function CompareBudgetsANDOptions(ByVal obj As GLPKDecisionsObject) As Boolean
            'Return (_BarrierEID = obj.BarrierEID And _Budget = obj.Budget)
            Return (_OptionNum < obj.DecisionOption And _Budget = obj.Budget)
        End Function

    End Class



    Private Class DecisionsOverlapObjectPredicateBudgets
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        'Private _BarrierEID As Integer
        Private _Budget As Double

        'Public Sub New(ByVal barriereid As Integer, ByVal budget As Double)
        Public Sub New(ByVal budget As Double)

            'Me._BarrierEID = barriereid
            Me._Budget = budget
        End Sub

        Public Function CompareBudgets2(ByVal obj As DecisionsOverlapObject) As Boolean
            'Return (_BarrierEID = obj.BarrierEID And _Budget = obj.Budget)
            Return (_Budget = obj._Budget)
        End Function
    End Class
    Private Class GLPKDecisionsObjectPredicateBarriers
        ' tutorial here: http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/bad5193a-3bf1-4675-8a05-625f46f9c158/
        'Private _BarrierEID As Integer
        Private _Barriers As Integer

        'Public Sub New(ByVal barriereid As Integer, ByVal budget As Double)
        Public Sub New(ByVal barriers As Double)

            'Me._BarrierEID = barriereid
            Me._Barriers = barriers
        End Sub

        Public Function CompareBarriers(ByVal obj As GLPKDecisionsObject) As Boolean
            'Return (_BarrierEID = obj.BarrierEID And _Budget = obj.Budget)
            Return (_Barriers = obj.BarrierEID)
        End Function
    End Class

    Private Sub CalculateOverlap()

        ' Created by: Greig Oldford
        ' for: Thesis
        ' Description:  calculates 'overlap' between optimisation treatments
        ' Logic: 
        '  check tables have been selected
        '  check fields in tables, make sure they are correct
        '  compare row counts of input tables
        '  create table for output
        '  read all rows in input tables and compare decisions
        '  prep statistics and write to output tables

        Me.CheckForIllegalCrossThreadCalls = False
        If Not BackgroundWorker3.IsBusy Then
            BackgroundWorker3.RunWorkerAsync()
        End If

        m_iProgressPercent = 0
        m_sProgressString = ""

        BackgroundWorker3.ReportProgress(0, "do")

        If m_pOutGxDatabase Is Nothing Then
            MsgBox("You must select an output Geodatabase")
            'exitsubroutine()
            Exit Sub
        End If
        If txtTable1.Text = "" Then
            MsgBox("You must select Table 1 for input.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End If
        If txtTable2.Text = "" Then
            MsgBox("You must select Table 2 for input.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End If
        If txtTable1.Text = txtTable2.Text Then
            MsgBox("Tables 1 and 2 must be different. Exception for debugging made... continuing")
            'BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            ' Exit Sub
        End If
        If m_pFWorkspace1 Is Nothing Or m_pFWorkspace2 Is Nothing Then
            MsgBox("Trouble getting path to one of the input tables")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End If

        Dim pTable1, pTable2 As ITable

        Try
            pTable1 = m_pFWorkspace1.OpenTable(m_sTableName1)
        Catch ex As Exception
            MsgBox("Trouble opening table 1. Exiting." + ex.Message)
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End Try
        Try
            pTable2 = m_pFWorkspace2.OpenTable(m_sTableName2)
        Catch ex As Exception
            MsgBox("Trouble opening table 2. Exiting. " + ex.Message)
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End Try
        m_iProgressPercent = 10
        m_sProgressString = "Checking Table Fields"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        ' =================================================================
        ' check that each table has the right fields
        ' field names must correspond to those created during the GLPK process
        ' fields:
        '   sinkEID
        '   Treatment
        '   Budget
        '   BarrierEID
        Dim iBarrierEIDField1, iTreatmentField1, iBudgetField1, iOptionField1 As Integer
        Dim iBarrierEIDField2, iTreatmentField2, iBudgetField2, iOptionField2 As Integer

        'check table 2
        If pTable2.FindField("BarrierEID") = -1 Then
            MsgBox("Cannot find field 'BarrierEID' in Table 2. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iBarrierEIDField2 = pTable2.FindField("BarrierEID")
        End If

        If pTable2.FindField("Budget") = -1 Then
            MsgBox("Cannot find field 'Budget' in Table 2. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iBudgetField2 = pTable2.FindField("Budget")
        End If

        If pTable2.FindField("Treatment") = -1 Then
            MsgBox("Cannot find field 'Treatment' in Table 2. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iTreatmentField2 = pTable2.FindField("Treatment")
        End If

        If pTable1.FindField("OptionNum") = -1 Then
            MsgBox("Cannot find field 'OptionNum' in Table 2. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iOptionField2 = pTable1.FindField("OptionNum")
        End If

        ' check table 1
        If pTable1.FindField("BarrierEID") = -1 Then
            MsgBox("Cannot find field 'BarrierEID' in Table 1. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iBarrierEIDField1 = pTable1.FindField("BarrierEID")
        End If

        If pTable1.FindField("Budget") = -1 Then
            MsgBox("Cannot find field 'Budget' in Table 1. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iBudgetField1 = pTable1.FindField("Budget")
        End If

        If pTable1.FindField("Treatment") = -1 Then
            MsgBox("Cannot find field 'Treatment' in Table 1. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iTreatmentField1 = pTable1.FindField("Treatment")
        End If

        If pTable1.FindField("OptionNum") = -1 Then
            MsgBox("Cannot find field 'OptionNum' in Table 1. Now Exiting.")
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        Else
            iOptionField1 = pTable1.FindField("OptionNum")
        End If

        ' =================================================================
        ' compare row count of tables
        Dim iTab1Rows, iTab2Rows As Integer
        iTab1Rows = pTable1.RowCount(Nothing)
        iTab2Rows = pTable2.RowCount(Nothing)

        'If iTab2Rows <> iTab1Rows Then
        '    MsgBox("The number of rows in the two input tables do not match.  Currently, to compare tables the number of barriers, the budget increments, and all other settings must match. Now exiting.")
        '    Exit Sub
        'End If

        ' =================================================================
        ' try to create output table (test user permissions to create table)
        ' get workspace of output 

        m_iProgressPercent = 20
        m_sProgressString = "Creating Output Table"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        Dim pWorkspaceOUT As IWorkspace
        Dim pFWorkspaceOUT As IFeatureWorkspace
        Try
            pWorkspaceOUT = m_pOutGxDatabase.Workspace
            pFWorkspaceOUT = CType(pWorkspaceOUT, IFeatureWorkspace)
        Catch ex As Exception
            MsgBox("Error trying to get feature workspace of output geodatabase. Now Exiting. " + ex.Message)
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End Try

        ' create table
        ' fields:
        '  barrierEID - integer
        '  budget     - double
        '  overlap    - integer (1)
        '  treatment1 - string (140)
        '  treatment2 - string (140)
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

        ' ============= First Field ============
        ' Create ObjectID Field
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "ObjectID" 'VB.NET
        pFieldEdit.Name_2 = "ObID"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeOID
        pFieldsEdit.Field_2(0) = pField

        ' ============ Second Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Budget"
        pFieldEdit.Name_2 = "Budget"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(1) = pField

        ' ============ Third Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "OverlapCount"
        pFieldEdit.Name_2 = "OverlapCount"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(2) = pField

        ' ============ Fourth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Overlap_Statistic"
        pFieldEdit.Name_2 = "Overlap_Statistic"

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(3) = pField

        ' ============ Fifth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment1"
        pFieldEdit.Name_2 = "Treatment1"
        pFieldEdit.Length_2 = 140

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(4) = pField


        ' ============ Sixth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "T1_DecisionCount"
        pFieldEdit.Name_2 = "Treatment1_DecisionsCount"
        'pFieldEdit.Length_2 = 140

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(5) = pField


        ' ============ Seventh Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "Treatment2"
        pFieldEdit.Name_2 = "Treatment2"
        pFieldEdit.Length_2 = 140

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(6) = pField

        ' ============ Eigth Field ============
        pField = New Field
        pFieldEdit = CType(pField, IFieldEdit)

        pFieldEdit.AliasName_2 = "T2_DecisionCount"
        pFieldEdit.Name_2 = "Treatment2_DecisionsCount"
        'pFieldEdit.Length_2 = 140

        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        'pFieldEdit.Scale_2 = 1
        pFieldsEdit.Field_2(7) = pField



        'get output tablename
        Dim sTablePrefix As String
        sTablePrefix = txtTableName.Text
        Dim sTableOut As String
        sTableOut = TableName(pFWorkspaceOUT, sTablePrefix)

        ' create the table
        Try
            pFWorkspaceOUT.CreateTable(sTableOut, pFields, Nothing, Nothing, "")
        Catch ex As Exception
            MsgBox("Problem creating output table in provided geodatabase. Now Exiting. " + ex.Message)
            BackgroundWorker3.CancelAsync()
            'exitsubroutine()
            Exit Sub
        End Try

        m_iProgressPercent = 30
        m_sProgressString = "Reading Input Table 1"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        ' update contents with table
        Dim pStTab As IStandaloneTable
        Dim pStTabColl As IStandaloneTableCollection
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        Dim pMxDoc As IMxDocument = CType(pDoc, IMxDocument)
        Dim pMap As IMap = pMxDoc.FocusMap

        ' Add Table to Map Doc
        Dim pTable As ITable = pFWorkspaceOUT.OpenTable(sTableOut)
        pStTabColl = CType(pMap, IStandaloneTableCollection)
        pStTab = New StandaloneTable
        pStTab.Table = pTable
        pStTabColl.AddStandaloneTable(pStTab)
        pMxDoc.UpdateContents()

        ' ============================================================
        ' read tables 1 and 2 into feature cursors
        ' read all table 2 into a list object

        Dim pCursor1, pCursor2 As ICursor
        pCursor1 = pTable1.Search(Nothing, False)
        pCursor2 = pTable2.Search(Nothing, False)

        Dim pRow1, pRow2 As IRow

        ' Open each table and read the lines into an object? Or read each line one at a time? 
        ' Chose: read table 1 line by line and compare to table 2.  Read all of table 2 into a list object
        ' at the beginning of the sub.  
        ' custom object used: GLPKDecisionsObject

        Dim pGLPKDecisionsObjectTAB1 As New GLPKDecisionsObject(Nothing, Nothing, Nothing, Nothing)
        Dim pGLPKDecisionsObjectTAB2 As New GLPKDecisionsObject(Nothing, Nothing, Nothing, Nothing)
        Dim lGLPKDecisionsObjectTAB1 As New List(Of GLPKDecisionsObject)
        Dim lGLPKDecisionsObjectTAB2 As New List(Of GLPKDecisionsObject)

        Dim iBarrierEID As Integer
        Dim dBudget As Double
        Dim sTreatment As String
        Dim iOption As Integer

        m_iProgressPercent = 40
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)
        ' variant / objects for each to be safe
        Dim pField2 As IField

        ' read each row into the object
        pRow1 = pCursor1.NextRow

        ' Loop through each row
        Do Until pRow1 Is Nothing

            pField2 = pRow1.Fields.Field(iBarrierEIDField1)
            If pField2.Type = esriFieldType.esriFieldTypeInteger Or pField2.Type = esriFieldType.esriFieldTypeSingle Or pField2.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField2.Type = esriFieldType.esriFieldTypeOID Then
                iBarrierEID = pRow1.Value(iBarrierEIDField1)
            Else
                MsgBox("Error reading Table 1. Could not convert the barrier EID to type 'integer'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow1.Fields.Field(iBudgetField1)
            If pField2.Type = esriFieldType.esriFieldTypeDouble Then
                dBudget = pRow1.Value(iBudgetField1)
            Else
                MsgBox("Error reading Table 1. Could not convert the budget to type 'double'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow1.Fields.Field(iTreatmentField1)
            If pField2.Type = esriFieldType.esriFieldTypeString Then
                sTreatment = pRow1.Value(iTreatmentField1)
            Else
                MsgBox("Error reading Table 1. Could not convert the treatment to type 'string'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow1.Fields.Field(iOptionField1)
            If pField2.Type = esriFieldType.esriFieldTypeInteger Or pField2.Type = esriFieldType.esriFieldTypeSingle Or pField2.Type = esriFieldType.esriFieldTypeSmallInteger Then
                iOption = pRow1.Value(iOptionField1)
            Else
                MsgBox("Error reading Table 1. Could not convert the optionnum to type 'integer'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            ' add to new object
            pGLPKDecisionsObjectTAB1 = New GLPKDecisionsObject(dBudget, sTreatment, iBarrierEID, iOption)
            lGLPKDecisionsObjectTAB1.Add(pGLPKDecisionsObjectTAB1)

            'pRow1 = Nothing
            pRow1 = pCursor1.NextRow
        Loop

        m_iProgressPercent = 50
        m_sProgressString = "Reading Input Table 2"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        ' read each row into the object
        pRow2 = pCursor2.NextRow

        ' Loop through each row
        Do Until pRow2 Is Nothing

            pField2 = pRow2.Fields.Field(iBarrierEIDField2)
            If pField2.Type = esriFieldType.esriFieldTypeInteger Or pField2.Type = esriFieldType.esriFieldTypeSingle Or pField2.Type = esriFieldType.esriFieldTypeSmallInteger _
                Or pField2.Type = esriFieldType.esriFieldTypeOID Then
                iBarrierEID = pRow2.Value(iBarrierEIDField2)
            Else
                MsgBox("Error reading table 2. Could not convert the barrier EID to type 'integer'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow2.Fields.Field(iBudgetField2)
            If pField2.Type = esriFieldType.esriFieldTypeDouble Then
                dBudget = pRow2.Value(iBudgetField2)
            Else
                MsgBox("Error reading table 2. Could not convert the budget to type 'double'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow2.Fields.Field(iTreatmentField2)
            If pField2.Type = esriFieldType.esriFieldTypeString Then
                sTreatment = pRow2.Value(iTreatmentField2)
            Else
                MsgBox("Error reading table 2. Could not convert the treatment to type 'string'. Now Exiting.")
                BackgroundWorker3.CancelAsync()
                'exitsubroutine()
                Exit Sub
            End If

            pField2 = pRow2.Fields.Field(iOptionField2)
            If pField2.Type = esriFieldType.esriFieldTypeInteger Or pField2.Type = esriFieldType.esriFieldTypeSingle Or pField2.Type = esriFieldType.esriFieldTypeSmallInteger Then
                iOption = pRow2.Value(iOptionField2)
            Else
                MsgBox("Error reading table 2. Could not convert the optionnum to type 'integer'. Now Exiting.")
                'exitsubroutine()
                Exit Sub
            End If

            ' add to new object
            pGLPKDecisionsObjectTAB2 = New GLPKDecisionsObject(dBudget, sTreatment, iBarrierEID, iOption)
            lGLPKDecisionsObjectTAB2.Add(pGLPKDecisionsObjectTAB2)

            'pRow1 = Nothing
            pRow2 = pCursor2.NextRow
        Loop

        m_iProgressPercent = 60
        m_sProgressString = "Comparing Tables"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        ' get unique budget amounts from each table
        Dim lUniqueBudgetsTAB1 As New List(Of Double)
        Dim lUniqueBudgetsTAB2 As New List(Of Double)
        Dim i As Integer
        ' for now the treatments are all assumed the same
        ' and the same as the first row
        Dim sTreatment1Temp As String
        Dim sTreatment2Temp As String
        Try
            sTreatment1Temp = lGLPKDecisionsObjectTAB1(0).Treatment
            sTreatment2Temp = lGLPKDecisionsObjectTAB2(0).Treatment
        Catch ex As Exception
            MsgBox("Error 15. " + ex.Message)
            sTreatment1Temp = "not sure"
            sTreatment2Temp = "not sure"
        End Try

        For i = 0 To lGLPKDecisionsObjectTAB1.Count - 1
            If lUniqueBudgetsTAB1.Contains(lGLPKDecisionsObjectTAB1(i).Budget) = False Then
                lUniqueBudgetsTAB1.Add(lGLPKDecisionsObjectTAB1(i).Budget)
            End If
        Next

        i = 0
        For i = 0 To lGLPKDecisionsObjectTAB2.Count - 1
            If lUniqueBudgetsTAB2.Contains(lGLPKDecisionsObjectTAB2(i).Budget) = False Then
                lUniqueBudgetsTAB2.Add(lGLPKDecisionsObjectTAB2(i).Budget)
            End If
        Next

        ' compare two - notify if descrepancy
        If lUniqueBudgetsTAB2.Count <> lUniqueBudgetsTAB1.Count Then
            MsgBox("The number of unique budget classes in table 1 does not equal the number in table 2. Table 1 count: " + lUniqueBudgetsTAB1.Count.ToString + " Table 2 count: " _
                   + lUniqueBudgetsTAB2.Count.ToString + " Exception allowed... continuing.")
            'backgroundworker3.CancelAsync()
            ''exitsubroutine()
            'Exit Sub
        End If

        ' get unique barriers from each table
        ' get unique budget amounts from each table
        Dim lUniqueBarriersTAB1 As New List(Of Integer)
        Dim lUniqueBarriersTAB2 As New List(Of Integer)
        

        For i = 0 To lGLPKDecisionsObjectTAB1.Count - 1
            If lUniqueBarriersTAB1.Contains(lGLPKDecisionsObjectTAB1(i).BarrierEID) = False Then
                lUniqueBarriersTAB1.Add(lGLPKDecisionsObjectTAB1(i).BarrierEID)
            End If
        Next

        For i = 0 To lGLPKDecisionsObjectTAB2.Count - 1
            If lUniqueBarriersTAB2.Contains(lGLPKDecisionsObjectTAB2(i).BarrierEID) = False Then
                lUniqueBarriersTAB2.Add(lGLPKDecisionsObjectTAB2(i).BarrierEID)
            End If
        Next

        ' compare two - notify if descrepancy
        If lUniqueBarriersTAB1.Count <> lUniqueBarriersTAB2.Count Then
            MsgBox("The number of unique barrier IDs in table 1 does not equal the number in table 2. Table 1 count: " + lUniqueBarriersTAB1.Count.ToString + " Table 2 count: " _
                               + lUniqueBarriersTAB2.Count.ToString + " Exception allowed... continuing")
            'backgroundworker3.CancelAsync()
            'backgroundworker3.Dispose()
            'Exit Sub
        End If


        m_iProgressPercent = 70
        m_sProgressString = "Comparing Decisions"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)

        ' *** basically, the two above lists should be the same ***
        ' so it doesn't matter which one is looped through.  

        ' need list of overlap results object
        '   parameters:
        '     barrierEID
        '     budget
        '     perc_overlap
        '     treatment1
        '     treatment2
        Dim pDecisionsOverlapObject As New DecisionsOverlapObject(Nothing, _
                                                                  Nothing, _
                                                                  Nothing, _
                                                                  Nothing, _
                                                                  Nothing, _
                                                                  Nothing, _
                                                                  Nothing, _
                                                                  Nothing)
        Dim lDecisionsOverlapObject As New List(Of DecisionsOverlapObject)
        Dim j As Integer = 0
        i = 0

        'Dim refinedGLPKOptionsList As List(Of GLPKOptionsObject)
        'Dim HabStatsComparer As RefineHabStatsListPredicate
        Dim DecisionBudgetComparer As GLPKDecisionsObjectPredicateBudgets
        Dim lRefinedGLPKDecisionsObjectTAB1 As New List(Of GLPKDecisionsObject)
        Dim lRefinedGLPKDecisionsObjectTAB2 As New List(Of GLPKDecisionsObject)
        Dim DecisionBarrierComparer As GLPKDecisionsObjectPredicateBarriers
        Dim lRefined_AGAIN_GLPKDecisionsObjectTAB1 As New List(Of GLPKDecisionsObject)
        Dim lRefined_AGAIN_GLPKDecisionsObjectTAB2 As New List(Of GLPKDecisionsObject)
        Dim bOverlap, bOverlap_DoSomething As Boolean
        Dim iOverlap_DoSomething As Integer
        Dim bWarning1, bWarning2 As Boolean ' warning for zero length or >1 length lists of unique barrier/budget combo
        Dim lRefinedGLPKDecisionObjectTAB1_DoSomething As New List(Of GLPKDecisionsObject)
        Dim lRefinedGLPKDecisionObjectTAB2_DoSomething As New List(Of GLPKDecisionsObject)
        Dim DecisionBudgetANDOptionComparer As GLPKDecisionsObjectPredicateBudgetsANDOptionNums

        Dim lCombined_GLPKDecisionsObject As New List(Of GLPKDecisionsObject)
        Dim lTEMPlist_GLPKDecisionsObject As New List(Of GLPKDecisionsObject)
        Dim lOverlap_Decisions As New List(Of GLPKDecisionsObject)
        Dim pRowBuffer As IRowBuffer
        Dim pCursor As ICursor
        Dim iTreatment1DecisionCount As Integer = 0
        Dim iTreatment2DecisionCount As Integer = 0
        Dim iOverlapCount As Integer = 0

        ' get a refined list of for each budget amount (not each barrier and budget amount
        ' for each unique budget
        ' get a list of barriers and decisions for this budget for each table
        ' compare decisions for barriers between tables using unique barriers list for table 1
        ' then using unique barriers list for table 2
        For i = 0 To lUniqueBudgetsTAB1.Count - 1

            lOverlap_Decisions = New List(Of GLPKDecisionsObject)
            lRefinedGLPKDecisionsObjectTAB1 = New List(Of GLPKDecisionsObject)
            lRefinedGLPKDecisionsObjectTAB2 = New List(Of GLPKDecisionsObject)
            lCombined_GLPKDecisionsObject = New List(Of GLPKDecisionsObject)

            ' get refined list of Table 1
            DecisionBudgetComparer = New GLPKDecisionsObjectPredicateBudgets(lUniqueBudgetsTAB1(i))
            lRefinedGLPKDecisionsObjectTAB1 = lGLPKDecisionsObjectTAB1.FindAll(AddressOf DecisionBudgetComparer.CompareBudgets)
            ' get refined list of table 2
            lRefinedGLPKDecisionsObjectTAB2 = lGLPKDecisionsObjectTAB2.FindAll(AddressOf DecisionBudgetComparer.CompareBudgets)

            ' warn if count of decisions at a budget don't match
            'If lRefinedGLPKDecisionsObjectTAB1.Count <> lRefinedGLPKDecisionsObjectTAB2.Count Then
            '    MsgBox("Warning: decisions at budget amount " + lUniqueBudgetsTAB1(i).ToString + " don't match. Overlap will be calculated at 0% for this budget amount.")
            'End If

            If lRefinedGLPKDecisionsObjectTAB1.Count = 0 Then
                MsgBox("Warning: decisions at budget amount " + lUniqueBudgetsTAB1(i).ToString + " for Table 1 weren't found. Overlap will be calculated at 0% for this budget amount.")
            End If

            If lRefinedGLPKDecisionsObjectTAB2.Count = 0 Then
                MsgBox("Warning: decisions at budget amount " + lUniqueBudgetsTAB2(i).ToString + " for Table 2 weren't found. Overlap will be calculated at 0% for this budget amount.")
            End If

            ' herein alternative to commented out below
            ' (because of changing statistic slightly)
            ' 1. Get all barriers with decisions for this budget for each input table
            ' 2. Loop through both lists created in previous step and combine into a master list
            '    but only if decision option number is > 1 (ie not do nothing)
            ' 3. Search this list for duplicates - those are overlaps.  
            '     Overlap stat = count of overlaps / (total length combined list - count of overlaps)
            Try
                For j = 0 To lRefinedGLPKDecisionsObjectTAB1.Count - 1
                    If lRefinedGLPKDecisionsObjectTAB1(j).DecisionOption > 1 Then
                        lCombined_GLPKDecisionsObject.Add(lRefinedGLPKDecisionsObjectTAB1(j))
                        iTreatment1DecisionCount += 1
                    End If
                Next
            Catch ex As Exception
                MsgBox("Error 23. " + ex.Message)
            End Try

            Try
                For j = 0 To lRefinedGLPKDecisionsObjectTAB2.Count - 1
                    If lRefinedGLPKDecisionsObjectTAB2(j).DecisionOption > 1 Then
                        lCombined_GLPKDecisionsObject.Add(lRefinedGLPKDecisionsObjectTAB2(j))
                        iTreatment2DecisionCount += 1
                    End If
                Next
            Catch ex As Exception
                MsgBox("Error 24. " + ex.Message)
            End Try

            ' check for pairs of barriers between treatments / input tables
            ' if pair is found then if it isn't already in the list
            ' add that pair
            Try
                For j = 0 To lCombined_GLPKDecisionsObject.Count - 1

                    DecisionBarrierComparer = New GLPKDecisionsObjectPredicateBarriers(lCombined_GLPKDecisionsObject(j).BarrierEID)
                    lTEMPlist_GLPKDecisionsObject = New List(Of GLPKDecisionsObject)
                    lTEMPlist_GLPKDecisionsObject = lCombined_GLPKDecisionsObject.FindAll(AddressOf DecisionBarrierComparer.CompareBarriers)

                    ' If there's only one (or less) then there is no overlap
                    If lTEMPlist_GLPKDecisionsObject.Count > 1 Then

                        ' If there's greater than two decisions for this budget and barrier then something is wrong
                        ' because there's only one decision possible at each barrier for each budget amount
                        If lTEMPlist_GLPKDecisionsObject.Count > 2 Then
                            MsgBox("Error. There should be only one occurence of each barrier at a given budget")
                        Else
                            ' if this barrier is not already in the overlap list and the decision matches between treatments then 
                            ' add it to the list. 
                            If Not lOverlap_Decisions.Exists(AddressOf DecisionBarrierComparer.CompareBarriers) Then
                                If lTEMPlist_GLPKDecisionsObject(0).DecisionOption = lTEMPlist_GLPKDecisionsObject(1).DecisionOption Then
                                    lOverlap_Decisions.Add(lTEMPlist_GLPKDecisionsObject(0))
                                    lOverlap_Decisions.Add(lTEMPlist_GLPKDecisionsObject(1))
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox("Error 25. " + ex.Message)
            End Try

            ' add the overlap percent and budget to the list
            Try
                pDecisionsOverlapObject = New DecisionsOverlapObject(Nothing, _
                                                                     Nothing, _
                                                                     Nothing, _
                                                                     Nothing, _
                                                                     Nothing,
                                                                     Nothing, _
                                                                     Nothing, _
                                                                     Nothing)
                If lOverlap_Decisions.Count > 0 And lCombined_GLPKDecisionsObject.Count > 0 Then
                    ' find the overlap percentage but
                    ' multiply by 1000 because the object only accepts integers
                    ' (using a relic object)
                    ' (will divide by 10 later)
                    iOverlapCount = lOverlap_Decisions.Count / 2
                    With pDecisionsOverlapObject
                        ._Budget = lUniqueBudgetsTAB1(i)
                        ._OverlapCount = iOverlapCount
                        ._OverlapStatistic = Convert.ToInt16(((lOverlap_Decisions.Count / lCombined_GLPKDecisionsObject.Count) * 1000))
                        ' going to borrow two unused integer object fields - barriereid and overlap -
                        ' for the decision count for each treatment
                        ._T1DecisionCount = iTreatment1DecisionCount
                        ._T2DecisionCount = iTreatment2DecisionCount
                    End With
                Else
                    With pDecisionsOverlapObject
                        ._Budget = lUniqueBudgetsTAB1(i)
                        ._OverlapCount = 0
                        ._OverlapStatistic = 0
                        ' going to borrow two unused integer object fields - barriereid and overlap -
                        ' for the decision count for each treatment
                        ._T1DecisionCount = iTreatment1DecisionCount
                        ._T2DecisionCount = iTreatment2DecisionCount
                    End With
                End If

                lDecisionsOverlapObject.Add(pDecisionsOverlapObject)
            Catch ex As Exception
                MsgBox("Error 26. " + ex.Message)
            End Try

            iTreatment1DecisionCount = 0
            iTreatment2DecisionCount = 0
            iOverlapCount = 0

        Next ' unique budget

        Dim dBudgetTemp As Double
        ' insert into the table
        Try
            For i = 0 To lDecisionsOverlapObject.Count - 1
                pRowBuffer = pTable.CreateRowBuffer
                dBudgetTemp = lDecisionsOverlapObject(i)._Budget

                pRowBuffer.Value(1) = lDecisionsOverlapObject(i)._Budget
                pRowBuffer.Value(2) = lDecisionsOverlapObject(i)._OverlapCount

                If lDecisionsOverlapObject(i)._OverlapStatistic > 0 Then
                    pRowBuffer.Value(3) = (lDecisionsOverlapObject(i)._OverlapStatistic / 10)
                End If

                pRowBuffer.Value(4) = sTreatment1Temp
                ' THIS IS NOT BARRIER EID, this is decision count
                pRowBuffer.Value(5) = lDecisionsOverlapObject(i)._T1DecisionCount
                pRowBuffer.Value(6) = sTreatment2Temp
                ' THIS IS NOT OVERLAP, this is decision count
                pRowBuffer.Value(7) = lDecisionsOverlapObject(i)._T2DecisionCount

                pCursor = pTable.Insert(True)
                pCursor.InsertRow(pRowBuffer)
            Next

        Catch ex As Exception
            MsgBox("Problem encountered inserting row in output table. " + ex.Message)
        End Try





        '    ' compare the decision at each barrier between each table
        '    For j = 0 To lUniqueBarriersTAB1.Count - 1

        '        bOverlap = False  ' overlap between decisions (do nothing included)
        '        iOverlap_DoSomething = 0 ' overlap between decisions (do nothing excluded)
        '        bWarning1 = False ' notifies if barrier is missing at this budget amount
        '        bWarning2 = False ' notifies if barier is missing at this budget amount

        '        ' re-refine the list to get what should be ONE item for a given budget and barrier (from each list)
        '        ' i.e., get the BARRIER for this budget amount 
        '        DecisionBarrierComparer = New GLPKDecisionsObjectPredicateBarriers(lUniqueBarriersTAB1(j))

        '        ' try to catch errors if no results were returned
        '        If lRefinedGLPKDecisionsObjectTAB1.Count <> 0 Then
        '            lRefined_AGAIN_GLPKDecisionsObjectTAB1 = lRefinedGLPKDecisionsObjectTAB1.FindAll(AddressOf DecisionBarrierComparer.CompareBarriers)
        '        Else
        '            lRefined_AGAIN_GLPKDecisionsObjectTAB1 = New List(Of GLPKDecisionsObject)
        '        End If

        '        If lRefinedGLPKDecisionsObjectTAB2.Count <> 0 Then
        '            lRefined_AGAIN_GLPKDecisionsObjectTAB2 = lRefinedGLPKDecisionsObjectTAB2.FindAll(AddressOf DecisionBarrierComparer.CompareBarriers)
        '        Else
        '            lRefined_AGAIN_GLPKDecisionsObjectTAB2 = New List(Of GLPKDecisionsObject)
        '        End If

        '        ' check that the count is 1 for each re-refined list (unique combo key of barrier and budget)
        '        If lRefined_AGAIN_GLPKDecisionsObjectTAB1.Count <> 1 Then
        '            'MsgBox("Warning. For budget amount " + lUniqueBudgetsTAB1(i).ToString + " and barrier ID " + lUniqueBarriersTAB1(j).ToString + " there were " _
        '            '+ lRefined_AGAIN_GLPKDecisionsObjectTAB1.Count.ToString + " records found. Overlap at this budget and barrier will be set to zero.")
        '            bWarning1 = True
        '        End If

        '        If lRefined_AGAIN_GLPKDecisionsObjectTAB2.Count <> 1 Then
        '            'MsgBox("Warning. For budget amount " + lUniqueBudgetsTAB1(i).ToString + " and barrier ID " + lUniqueBarriersTAB1(j).ToString + " there were " _
        '            '+ lRefined_AGAIN_GLPKDecisionsObjectTAB2.Count.ToString + " records found. Overlap at this budget and barrier will be set to zero.")
        '            bWarning2 = True
        '        End If

        '        ' compare the two
        '        ' if both barriers were found
        '        If bWarning1 = False And bWarning2 = False Then

        '            ' get overlap including do nothing case
        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption Then
        '                bOverlap = True
        '            End If

        '            ' get overlap excluding the 'do nothing' decisions
        '            ' both 'do nothing'             = 0  *default*
        '            ' one barrier is 'do something' = 1
        '            ' both barriers same decision   = 2
        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption > 1 Or lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption > 1 Then
        '                If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption Then
        '                    iOverlap_DoSomething = 2
        '                Else
        '                    iOverlap_DoSomething = 1
        '                End If
        '            End If
        '        ElseIf bWarning1 = False And bWarning2 = True Then
        '            bOverlap = False
        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption > 1 Then
        '                iOverlap_DoSomething = 1
        '            End If

        '        ElseIf bWarning2 = False And bWarning1 = True Then
        '            bOverlap = False
        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption > 1 Then
        '                iOverlap_DoSomething = 1
        '            End If

        '        End If

        '        pDecisionsOverlapObject = New DecisionsOverlapObject(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

        '        With pDecisionsOverlapObject
        '            ._BarrierEID = lUniqueBarriersTAB1(j)
        '            ._Budget = lUniqueBudgetsTAB1(i)
        '            If bWarning1 = False Then
        '                ._Treatment1 = lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).Treatment
        '            Else
        '                ._Treatment1 = "not sure"
        '            End If
        '            If bWarning2 = False Then
        '                ._Treatment2 = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).Treatment
        '            Else
        '                ._Treatment2 = "not sure"
        '            End If
        '            ._Overlap_DoSomething = iOverlap_DoSomething
        '        End With

        '        If bOverlap = False Then
        '            pDecisionsOverlapObject._Overlap = 0
        '        Else
        '            pDecisionsOverlapObject._Overlap = 1
        '        End If

        '        lDecisionsOverlapObject.Add(pDecisionsOverlapObject)
        '    Next ' unique barrier

        '    ' have to do the reverse now, to check overlap the other way.  
        '    ' for all barriers in the second list, 
        '    ' if they can't be found in the decisions overlap for this budget amount

        '    For j = 0 To lUniqueBarriersTAB2.Count - 1
        '        If lUniqueBarriersTAB1.Contains(lUniqueBarriersTAB2(j)) = False Then

        '            bOverlap = False  ' overlap between decisions (do nothing included)
        '            iOverlap_DoSomething = 0 ' overlap between decisions (do nothing excluded)
        '            bWarning1 = False ' notifies if barrier is missing at this budget amount
        '            bWarning2 = False ' notifies if barier is missing at this budget amount

        '            ' re-refine the list to get what should be ONE item for a given budget and barrier (from each list)
        '            ' i.e., get the BARRIER for this budget amount 
        '            DecisionBarrierComparer = New GLPKDecisionsObjectPredicateBarriers(lUniqueBarriersTAB2(j))

        '            ' try to catch errors if no results were returned
        '            If lRefinedGLPKDecisionsObjectTAB1.Count <> 0 Then
        '                lRefined_AGAIN_GLPKDecisionsObjectTAB1 = lRefinedGLPKDecisionsObjectTAB1.FindAll(AddressOf DecisionBarrierComparer.CompareBarriers)
        '            Else
        '                lRefined_AGAIN_GLPKDecisionsObjectTAB1 = New List(Of GLPKDecisionsObject)
        '            End If

        '            If lRefinedGLPKDecisionsObjectTAB2.Count <> 0 Then
        '                lRefined_AGAIN_GLPKDecisionsObjectTAB2 = lRefinedGLPKDecisionsObjectTAB2.FindAll(AddressOf DecisionBarrierComparer.CompareBarriers)
        '            Else
        '                lRefined_AGAIN_GLPKDecisionsObjectTAB2 = New List(Of GLPKDecisionsObject)
        '            End If

        '            ' check that the count is 1 for each re-refined list (unique combo key of barrier and budget)
        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB1.Count <> 1 Then
        '                'MsgBox("Warning. For budget amount " + lUniqueBudgetsTAB1(i).ToString + " and barrier ID " + lUniqueBarriersTAB1(j).ToString + " there were " _
        '                '+ lRefined_AGAIN_GLPKDecisionsObjectTAB1.Count.ToString + " records found. Overlap at this budget and barrier will be set to zero.")
        '                bWarning1 = True
        '            End If

        '            If lRefined_AGAIN_GLPKDecisionsObjectTAB2.Count <> 1 Then
        '                'MsgBox("Warning. For budget amount " + lUniqueBudgetsTAB1(i).ToString + " and barrier ID " + lUniqueBarriersTAB1(j).ToString + " there were " _
        '                '+ lRefined_AGAIN_GLPKDecisionsObjectTAB2.Count.ToString + " records found. Overlap at this budget and barrier will be set to zero.")
        '                bWarning2 = True
        '            End If

        '            ' compare the two
        '            ' if both barriers were found
        '            If bWarning1 = False And bWarning2 = False Then

        '                ' get overlap including do nothing case
        '                If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption Then
        '                    bOverlap = True
        '                End If

        '                ' get overlap excluding the 'do nothing' decisions
        '                ' both 'do nothing'             = 0  *default*
        '                ' one barrier is 'do something' = 1
        '                ' both barriers same decision   = 2
        '                If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption > 1 Or lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption > 1 Then
        '                    If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption Then
        '                        iOverlap_DoSomething = 2
        '                    Else
        '                        iOverlap_DoSomething = 1
        '                    End If
        '                End If
        '            ElseIf bWarning1 = False And bWarning2 = True Then
        '                bOverlap = False
        '                If lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).DecisionOption > 1 Then
        '                    iOverlap_DoSomething = 1
        '                End If

        '            ElseIf bWarning2 = False And bWarning1 = True Then
        '                bOverlap = False
        '                If lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).DecisionOption > 1 Then
        '                    iOverlap_DoSomething = 1
        '                End If

        '            End If

        '            pDecisionsOverlapObject = New DecisionsOverlapObject(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

        '            With pDecisionsOverlapObject
        '                ._BarrierEID = lUniqueBarriersTAB2(j)
        '                ._Budget = lUniqueBudgetsTAB1(i)
        '                If bWarning1 = False Then
        '                    ._Treatment1 = lRefined_AGAIN_GLPKDecisionsObjectTAB1(0).Treatment
        '                Else
        '                    ._Treatment1 = "not sure"
        '                End If
        '                If bWarning2 = False Then
        '                    ._Treatment2 = lRefined_AGAIN_GLPKDecisionsObjectTAB2(0).Treatment
        '                Else
        '                    ._Treatment2 = "not sure"
        '                End If
        '                ._Overlap_DoSomething = iOverlap_DoSomething
        '            End With

        '            If bOverlap = False Then
        '                pDecisionsOverlapObject._Overlap = 0
        '            Else
        '                pDecisionsOverlapObject._Overlap = 1
        '            End If

        '            lDecisionsOverlapObject.Add(pDecisionsOverlapObject)

        '        End If
        '    Next ' unique barrier from second table

        '    '----------------- second stat ------------------------------
        '    ' for second stat - only overlap between decisions made - 
        '    ' get refined list for each budget AND where decision options are > 1 (not nothing)
        '    DecisionBudgetANDOptionComparer = New GLPKDecisionsObjectPredicateBudgetsANDOptionNums(lUniqueBudgetsTAB1(i), 1)
        '    lRefinedGLPKDecisionObjectTAB1_DoSomething = lGLPKDecisionsObjectTAB1.FindAll(AddressOf DecisionBudgetANDOptionComparer.CompareBudgetsANDOptions)
        '    lRefinedGLPKDecisionObjectTAB2_DoSomething = lGLPKDecisionsObjectTAB1.FindAll(AddressOf DecisionBudgetANDOptionComparer.CompareBudgetsANDOptions)

        '    ' for each list loop through and compare at this barrier
        '    ' whether the other list contains this barrier
        '    ' if so, is the decision option the same?


        ''Next ' unique budget

        'Dim iTEMPCOUNT As Integer
        'iTEMPCOUNT = lDecisionsOverlapObject.Count

        '' =================================================================
        '' will need a custom predicate to refine list for comparisons

        '' for each budget amount
        '' get refined list
        '' calculate overlap statistic
        'Dim pDecisionOverlapObjectDinal As New DecisionsOverlapObjectFinal(Nothing, Nothing, Nothing, Nothing, Nothing)
        'Dim lDecisionOverlapObjectFinal As New List(Of DecisionsOverlapObjectFinal)
        'Dim DecisionBudgetComparer2 As DecisionsOverlapObjectPredicateBudgets ' have to create a second because of slight difference in objects 

        'Dim lRefinedDecisionsOverlapObject As New List(Of DecisionsOverlapObject)

        'Dim pRowBuffer As IRowBuffer
        'Dim pCursor As ICursor
        'Dim dOverlapPercent, dOverlapPercent_DoSomething As Double
        'Dim iTallyOverlap, iTallyOverlap_DoSomething_Same, iTallyOverlap_DoSomething As Integer

        'pTable = pFWorkspaceOUT.OpenTable(sTableOut)
        'i = 0
        'j = 0
        'For i = 0 To lUniqueBudgetsTAB1.Count - 1

        '    DecisionBudgetComparer2 = New DecisionsOverlapObjectPredicateBudgets(lUniqueBudgetsTAB1(i))
        '    lRefinedDecisionsOverlapObject = New List(Of DecisionsOverlapObject)
        '    lRefinedDecisionsOverlapObject = lDecisionsOverlapObject.FindAll(AddressOf DecisionBudgetComparer2.CompareBudgets2)
        '    dOverlapPercent = 0
        '    dOverlapPercent_DoSomething = 0
        '    iTallyOverlap = 0
        '    iTallyOverlap_DoSomething_Same = 0
        '    iTallyOverlap_DoSomething = 0

        '    For j = 0 To lRefinedDecisionsOverlapObject.Count - 1

        '        ' get overlap percent including the 'do nothing' cases
        '        If lRefinedDecisionsOverlapObject(j)._Overlap = 1 Then
        '            iTallyOverlap = iTallyOverlap + 1
        '        End If

        '        If lRefinedDecisionsOverlapObject(j)._Overlap_DoSomething > 0 Then
        '            iTallyOverlap_DoSomething = iTallyOverlap_DoSomething + 1
        '            If lRefinedDecisionsOverlapObject(j)._Overlap_DoSomething > 1 Then
        '                iTallyOverlap_DoSomething_Same = iTallyOverlap_DoSomething_Same + 1
        '            End If
        '        End If

        '    Next
        '    Try
        '        dOverlapPercent = (iTallyOverlap / lRefinedDecisionsOverlapObject.Count) * 100
        '    Catch ex As Exception
        '        MsgBox("Problem calculating overlap percent. May be a divide by zero prob. " + ex.Message)
        '        dOverlapPercent = 0
        '    End Try
        '    Try
        '        dOverlapPercent_DoSomething = (iTallyOverlap_DoSomething_Same / iTallyOverlap_DoSomething) * 100
        '    Catch ex As Exception
        '        MsgBox("Problem calculating overlap percent (do somethings). May be a divide by zero prob. " + ex.Message)
        '        dOverlapPercent_DoSomething = 0
        '    End Try

        '    ' get overlap percent excluding the 'do nothing' cases (both are do nothing)
        '    Try
        '        pRowBuffer = pTable.CreateRowBuffer
        '        dBudgetTemp = lRefinedDecisionsOverlapObject(i)._Budget

        '        pRowBuffer.Value(1) = lRefinedDecisionsOverlapObject(i)._Budget
        '        If dOverlapPercent > 0 Then
        '            pRowBuffer.Value(2) = Math.Round(dOverlapPercent, 2)
        '        End If
        '        If dOverlapPercent_DoSomething > 0 Then
        '            pRowBuffer.Value(3) = Math.Round(dOverlapPercent_DoSomething, 2)
        '        End If

        '        sTreatment1Temp = lRefinedDecisionsOverlapObject(i)._Treatment1
        '        sTreatment2Temp = lRefinedDecisionsOverlapObject(i)._Treatment2
        '        pRowBuffer.Value(4) = lRefinedDecisionsOverlapObject(i)._Treatment1
        '        pRowBuffer.Value(5) = lRefinedDecisionsOverlapObject(i)._Treatment2

        '        pCursor = pTable.Insert(True)
        '        pCursor.InsertRow(pRowBuffer)
        '    Catch ex As Exception
        '        MsgBox("Problem encountered inserting row in output table. " + ex.Message)
        '    End Try

        'Next


        m_iProgressPercent = 100
        m_sProgressString = "Complete!"
        BackgroundWorker3.ReportProgress(m_iProgressPercent, m_sProgressString)
        pCursor = Nothing
        pTable = Nothing
        pRowBuffer = Nothing
        'progressThread.Abort()
    End Sub
    'Private Sub frmCalculateOverlap_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    '    progressThread.Abort()
    'End Sub


End Class
