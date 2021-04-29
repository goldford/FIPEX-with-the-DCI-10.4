
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.esriSystem
Imports System.Runtime.InteropServices
'Imports System.Windows.Data.Binding
Imports System.IO
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Desktop.AddIns
Imports System.Windows.Forms


'Namespace PersistExtensionAddIn
Public Class FishPassageExtension
    Inherits ESRI.ArcGIS.Desktop.AddIns.Extension

    ' =================
    '==================
    'Imports system.windows.forms

    'Implements IExtension
    'Implements IExtensionConfig
    'Implements IPersistVariant

    ' self-reference for inter component communication 
    ' from http://help.arcgis.com/en/sdk/10.0/arcobjects_net/conceptualhelp/index.html#/Add_in_coding_patterns/0001000000zz000000/
    Private Shared s_extension As FishPassageExtension
    Private Shared m_UNAextension As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt

    Private m_bHasNetworks As Boolean = False
    Private m_application As IApplication
    Public m_enableState As ExtensionState

    Public pPropset As ESRI.ArcGIS.esriSystem.IPropertySet = New PropertySet
    'Public _sDictionary As New Dictionary(Of String, String)
    Public m_bLoaded As Boolean = False ' variable to check whether this is a new document or a loaded document.
    Public m_bFlagsLoaded As Boolean = False
    Public m_bBarriersLoaded As Boolean = False

    ' Below is for ArcGIS 10.0
    ' HOWEVER I don't think it is used again anywhere.  
    'Private Const RequiredProductCode As esriLicenseProductCode = esriLicenseProductCode.esriLicenseProductCodeArcInfo
   
    ' having issues with handler removal (removing unadded handlers)
    ' these keep track of successful handle adds
    Public m_bItemAdded_HandlerLoaded As Boolean = False
    Public m_bItemDeleted_HandlerLoaded As Boolean = False
    Public m_bContentsChanged_HandlerLoaded As Boolean = False
    Public m_bNewDocument_HandlerLoaded As Boolean = False
    Public m_bOpenDocument_HandlerLoaded As Boolean = False

    Public Sub New()
        s_extension = Me
    End Sub

    Public Property ToolsEnabledProp() As Boolean
        ' used to tell tools whether they're enabled or disabled.  
        ' tied to the 
        Get

        End Get
        Set(ByVal value As Boolean)

        End Set
    End Property
    Public Property HasNetworkProp() As Boolean
        Get
            Return m_bHasNetworks
        End Get
        Set(ByVal value As Boolean)
            m_bHasNetworks = value
        End Set
    End Property


    Friend Shared Function GetExtension() As FishPassageExtension
        ' Extension loads just in time, call FindExtension to load it.
        ' this gets a reference to the instance of this extension occuring in the '
        ' map application
        If s_extension Is Nothing Then
            Dim extID As UID = New UIDClass()
            extID.Value = "FIPEX_AddIn_dotNet45_2020_FishPassageExtension"
            My.ArcMap.Application.FindExtensionByCLSID(extID)
        End If
        Return s_extension
    End Function

    Friend Shared Function GetUNAExt() As ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt

        If m_UNAextension Is Nothing Then

            ' sets the guid of the utility network analyst extension
            Dim pUID As New ESRI.ArcGIS.esriSystem.UID
            pUID.Value = "{98528F9B-B971-11D2-BABD-00C04FA33C20}"

            Dim pExtension As ESRI.ArcGIS.esriSystem.IExtension
            If Not My.ArcMap.Application Is Nothing Then
                Try
                    pExtension = My.ArcMap.Application.FindExtensionByCLSID(pUID)

                    If pExtension Is Nothing Then
                        MsgBox("FIPEX could not retrieve Utility Network Analyst extension.")
                        m_UNAextension = Nothing
                        Return m_UNAextension
                    End If

                    m_UNAextension = TryCast(pExtension, ESRI.ArcGIS.EditorExt.IUtilityNetworkAnalysisExt)

                Catch ex As Exception
                    MsgBox("Error trying to open the Utility Network Analyst Extension during document load. Code 6. " + _
                           "Documents saved using FIPEX should be opened after opening an empty ArcMap session " + _
                           "(rather than opening an MXD file directly). The UNA extension has not been loaded. ")
                End Try
            Else
                MsgBox("The Utility Network Analyst Extension could not be retrieved by FIPEX during document load. Code 7." + _
                       "This is a bug with this version of ArcMap. Document will not load properly.  Please try opening ArcMap.exe first " + _
                       "rather than opening an MXD directly.  ")
            End If
           

        End If
        Return m_UNAextension
    End Function

    Friend Sub DoWork()
        System.Windows.Forms.MessageBox.Show("Do work")
    End Sub


    Protected Overrides Sub OnStartup()
        '
        ' TODO: Uncomment to start listening to document events
        '

        ' 4MsgBox("removeme Startup FIPEX")

        Try
            WireDocumentEvents()
        Catch ex As Exception
            MsgBox("Error on Startup loading FIPEX. " + ex.Message)
            Exit Sub
        End Try


        ' 4MsgBox("removeme before s_extension")

        s_extension = Me


        ' 5MsgBox("removeme Startup Before GETUNA")

        Try
            GetUNAExt()
        Catch ex As Exception
            MsgBox("Error on Startup loading FIPEX.  Code 2. " + ex.Message)
            Exit Sub
        End Try

        'If Utility Network Analyst is not loaded, do not initialize
        If Not m_UNAextension Is Nothing Then
            Try
                Initialize()
            Catch ex As Exception
                MsgBox("Error on Startup loading FIPEX. Code 3. " + ex.Message)
                Uninitialize()
                s_extension.State = ExtensionState.Disabled
            End Try
        End If


        'MsgBox("Continuing.")

        'Dim pNetworkAnalysisExt As INetworkAnalysisExt
        'pNetworkAnalysisExt = CType(m_UNAextension, INetworkAnalysisExt)
        'Dim iNetworkCountSource As Integer = pNetworkAnalysisExt.NetworkCount

        'Dim myBinding As New Binding("MyDataProperty")

    End Sub

    Private Sub Initialize()
        '5 MsgBox("removeme Initializing FIPEX")
        Try
            If s_extension Is Nothing Then
                Return
            ElseIf Me.State <> ExtensionState.Enabled Then
                Return
            End If
        Catch ex As Exception
            MsgBox("Error trying to check FIPEX extension state. Code 10. " + _
                   "Try opening a new ArcMap session (ArcMap.exe) rather than opening a " + _
                   " MXD document directly. " + ex.Message)
            Return
        End Try

        Try
            If m_bLoaded = False Then
                With pPropset
                    ' Load defaults
                    .SetProperty("direction", "up")
                    .SetProperty("ordernum", Convert.ToInt32(999))
                    .SetProperty("maximum", Convert.ToBoolean("True"))
                    .SetProperty("connecttab", Convert.ToBoolean("False"))
                    .SetProperty("advconnecttab", Convert.ToBoolean("False"))
                    .SetProperty("barrierperm", Convert.ToBoolean("False"))
                    .SetProperty("naturalyn", Convert.ToBoolean("False"))
                    .SetProperty("dciyn", Convert.ToBoolean("False"))
                    .SetProperty("dcisectionalyn", Convert.ToBoolean("False"))
                    .SetProperty("sDCIModelDir", Convert.ToString("not set"))
                    .SetProperty("sRInstallDir", Convert.ToString("not set"))
                    .SetProperty("bDBF", Convert.ToBoolean("False"))
                    .SetProperty("sGDB", Convert.ToString("not set"))
                    .SetProperty("TabPrefix", Convert.ToString("not set"))

                    .SetProperty("UpHab", Convert.ToBoolean("True"))
                    .SetProperty("TotalUpHab", Convert.ToBoolean("True"))
                    .SetProperty("DownHab", Convert.ToBoolean("False"))
                    .SetProperty("TotalDownHab", Convert.ToBoolean("False"))
                    .SetProperty("PathDownHab", Convert.ToBoolean("False"))
                    .SetProperty("TotalPathDownHab", Convert.ToBoolean("False"))

                    .SetProperty("numPolys", Convert.ToInt32(0))
                    .SetProperty("numLines", Convert.ToInt32(0))
                    .SetProperty("numExclusions", Convert.ToInt32(0))
                    .SetProperty("numBarrierIDs", Convert.ToInt32(0))
                    .SetProperty("barrierEIDsLoadedyn", Convert.ToBoolean("False"))
                    .SetProperty("numBarrierEIDs", Convert.ToInt32(0))
                    .SetProperty("flagEIDsLoadedyn", Convert.ToBoolean("False"))
                    .SetProperty("numFlagEIDs", Convert.ToInt32(0))

                    .SetProperty("bGLPKTables", Convert.ToBoolean("False"))
                    .SetProperty("sGLPKModelDir", Convert.ToString("not set"))
                    .SetProperty("sGnuWinDir", Convert.ToString("not set"))

                    m_bLoaded = True
                    m_bFlagsLoaded = False
                    m_bBarriersLoaded = False
                End With
            End If
        Catch ex As Exception
            MsgBox("Error during Initialization of FIPEX. Code 15. " + _
                   "Trouble setting property set. " + ex.Message)
        End Try
       

        ' Reset event handlers
        If My.ArcMap.Document Is Nothing Then
            MsgBox("FIPEX could not get reference to ArcMap Document during initialization. Code 11. " + _
                   "This is often due to opening an MXD document directly rather than an ArcMap session first, " + _
                   "then opening the document. ")
            Exit Sub
        End If
        Dim avEvent As IActiveViewEvents_Event

        If My.ArcMap.Document.FocusMap Is Nothing Then
            MsgBox("FIPEX could not attain reference to ArcMap 'focusMap'. Code 14. " + _
                   "This is often due to opening an MXD document directly rather than an ArcMap session first, " + _
                   "then opening the document. ")
            'Me.State = ExtensionState.Disabled
            Exit Sub
        End If


        ' MsgBox("removeme Initialize after focus map retrieved")

        Try
            avEvent = TryCast(My.ArcMap.Document.FocusMap, IActiveViewEvents_Event)
        Catch ex As Exception
            MsgBox("Issue encountered retrieving map document reference during FIPEX initialization. Code 12. " + _
                   "This is often due to opening an MXD document directly rather than an ArcMap session first, " + _
                   "then opening the document. " + ex.Message)

        End Try

        AddHandler avEvent.ItemAdded, AddressOf AvEvent_ItemAdded
        m_bItemAdded_HandlerLoaded = True
        AddHandler avEvent.ItemDeleted, AddressOf AvEvent_ItemAdded
        m_bItemDeleted_HandlerLoaded = True
        AddHandler avEvent.ContentsChanged, AddressOf MapEvents_ContentsChanged
        m_bContentsChanged_HandlerLoaded = True

        ' Wire up events
        AddHandler My.ArcMap.Events.NewDocument, AddressOf ArcMap_NewOpenDocument
        m_bNewDocument_HandlerLoaded = True
        AddHandler My.ArcMap.Events.OpenDocument, AddressOf ArcMap_NewOpenDocument
        m_bOpenDocument_HandlerLoaded = True
        ' Update the UI
        'm_map = ArcMap.Document.FocusMap
        Try
            m_bHasNetworks = CheckNetworkCount()

        Catch ex As Exception
            MsgBox("Problem retrieving network count from Utility Network Analyst during FIPEX load. Code 13." + _
            "This is often due to opening an MXD document directly rather than an ArcMap session first, " + _
                   "then opening the document. " + ex.Message)
        End Try

        'If m_bLoaded = False Then
        '    m_bLoaded = True
        'End If
      

    End Sub
    Private Sub Uninitialize()
        If s_extension Is Nothing Then
            Return
        End If

        ' Detach event handlers
        Dim avEvent As IActiveViewEvents_Event
        Try
            avEvent = TryCast(My.Document.FocusMap, IActiveViewEvents_Event)
        Catch ex As Exception
            MsgBox("Trouble getting reference to Focus Map during FIPEX unitialization. Code 21. " + _
                   ex.Message)
            Exit Sub
        End Try
        Try

            'm_bItemAdded_HandlerLoaded = False
            'm_bItemDeleted_HandlerLoaded = False
            'm_bContentsChanged_HandlerLoaded = False
            ' m_bNewDocument_HandlerLoaded = False
            'm_bOpenDocument_HandlerLoaded = False

            If m_bItemAdded_HandlerLoaded = True Then
                RemoveHandler avEvent.ItemAdded, AddressOf AvEvent_ItemAdded
            End If
            If m_bItemDeleted_HandlerLoaded = True Then
                RemoveHandler avEvent.ItemDeleted, AddressOf AvEvent_ItemAdded
            End If
            If m_bContentsChanged_HandlerLoaded = True Then
                RemoveHandler avEvent.ContentsChanged, AddressOf MapEvents_ContentsChanged
            End If

            avEvent = Nothing
        Catch ex As Exception
            MsgBox("Error during 'remove handler' process of FIPEX uninitialization. " + ex.Message)
            Exit Sub
        End Try

        ' Update UI
        ' set all the buttons disabled.
        '.SetEnabled(False)
    End Sub
    Protected Overrides Function OnGetState() As ExtensionState
        'MsgBox("removeme Getting State of FIPEX")
        m_enableState = Me.State
        Return Me.State
    End Function
    Protected Overrides Function OnSetState(ByVal state As ExtensionState) As Boolean
        ' Optionally check for a license here
        'MsgBox("removeme Setting State of FIPEX")

        Me.State = state
        m_enableState = state

        If state = ExtensionState.Enabled Then
            Try
                Initialize()
            Catch ex As Exception
                MsgBox("Error during FIPEX initialization process during setstate routine. Code 22. " + _
                       ex.Message)
            End Try
        Else
            Try
                Uninitialize()
            Catch ex As Exception
                MsgBox("Error during FIPEX unitialization process during setstate rountine. Code 23. " + _
                       ex.Message)
            End Try
        End If

        Return MyBase.OnSetState(state)
    End Function
    Private Sub AvEvent_ItemAdded(ByVal Item As Object)
        'm_map = ArcMap.Document.FocusMap
        'FillComboBox()
        'UpdateSelCountDockWin()
        'MsgBox("removeme AvEvent_ItemAdded FIPEX triggered")
        Try
            m_bHasNetworks = CheckNetworkCount()

        Catch ex As Exception
            MsgBox("Error trying to get network count during FIPEX load. Code 23. " + _
                   ex.Message)
        End Try
    End Sub

    Protected Overrides Sub OnShutdown()

    End Sub

    Private Sub ArcMap_NewOpenDocument()

        ' 5MsgBox("removeme AvEvent_NewOpenDocument FIPEX triggered")
        Dim pageLayoutEvent As IActiveViewEvents_Event

        Try
            pageLayoutEvent = TryCast(My.ArcMap.Document.PageLayout, IActiveViewEvents_Event)
        Catch ex As Exception
            MsgBox("Error during 'new open document' routine of FIPEX Extension. Code 17. " + ex.Message)
        End Try

        Try
            AddHandler pageLayoutEvent.FocusMapChanged, AddressOf AVEvents_FocusMapChanged
        Catch ex As Exception
            MsgBox("Error trying to 'add focusmap changed' handler to FIPEX. Code 18. " + ex.Message)
        End Try

        
    End Sub

    Private Sub AVEvents_FocusMapChanged()

        ' 2MsgBox("removeme AVEvents_FocusMapChanged FIPEX triggered")
        Try
            Initialize()
        Catch ex As Exception
            MsgBox("Error during map initialization of FIPEX. Code 19. " + _
                   "Try opening this document after an empty ArcMap session is started " + _
                   "(open ArcMap.exe then the document). " + ex.Message)
        End Try
    End Sub
    Private Sub WireDocumentEvents()
        '
        ' TODO: Sample document event wiring code. Change as needed.
        '
        'AddHandler My.ArcMap.Events.NewDocument, AddressOf ArcMapNewDocument

    End Sub

    Private Sub ArcMapNewDocument()
        ' TODO: Add code to handle new document event
    End Sub

    ''' <summary>
    ''' Determine extension state
    ''' </summary>
    'Private Function StateCheck(ByVal requestEnable As Boolean) As esriExtensionState
    '    'Turn on or off extension directly 
    '    If requestEnable Then
    '        'Check if the correct product is licensed
    '        'Dim aoInitTestProduct As IAoInitialize = New AoInitializeClass()
    '        'Dim prodCode As esriLicenseProductCode = aoInitTestProduct.InitializedProduct()
    '        'If prodCode = RequiredProductCode Then _
    '        Return esriExtensionState.esriESEnabled

    '        'Return esriExtensionState.esriESUnavailable
    '    Else
    '        Return esriExtensionState.esriESDisabled
    '    End If
    'End Function

    '#Region "IExtension Members"
    '    ''' <summary>
    '    ''' Name of extension. Do not exceed 31 characters
    '    ''' </summary>
    '    Public ReadOnly Property Name() As String Implements ESRI.ArcGIS.esriSystem.IExtension.Name
    '        Get
    '            Return "DFOBarriersAnalysisExtension"
    '        End Get
    '    End Property

    '    Public Sub Shutdown() Implements ESRI.ArcGIS.esriSystem.IExtension.Shutdown
    '        m_application = Nothing
    '    End Sub

    '    Public Sub Startup(ByRef initializationData As Object) Implements ESRI.ArcGIS.esriSystem.IExtension.Startup
    '        m_application = CType(initializationData, IApplication)
    '        If m_application Is Nothing Then Return

    '    End Sub
    '#End Region

    '#Region "IExtensionConfig Members"
    '    Public ReadOnly Property Description() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.Description
    '        Get
    '            Return "DFO Maritimes Fish Passage Extension"
    '        End Get
    '    End Property

    '    ''' <summary>
    '    ''' Friendly name shown in the Extensions dialog
    '    ''' </summary>
    '    Public ReadOnly Property ProductName() As String Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.ProductName
    '        Get
    '            Return "DFO Maritimes Fish Passage Extension"
    '        End Get
    '    End Property

    '    Public Property State() As ESRI.ArcGIS.esriSystem.esriExtensionState Implements ESRI.ArcGIS.esriSystem.IExtensionConfig.State
    '        Get
    '            Return m_enableState
    '        End Get
    '        Set(ByVal value As ESRI.ArcGIS.esriSystem.esriExtensionState)
    '            If m_enableState <> 0 And value = m_enableState Then Exit Property

    '            'Check if ok to enable or disable extension
    '            Dim requestState As esriExtensionState = value
    '            If requestState = esriExtensionState.esriESEnabled Then
    '                'Cannot enable if it's already in unavailable state
    '                If m_enableState = esriExtensionState.esriESUnavailable Then
    '                    Throw New COMException("Not running the appropriate product license.")
    '                End If

    '                'Determine if state can be changed
    '                Dim checkState As esriExtensionState = StateCheck(True)
    '                m_enableState = checkState

    '                If m_enableState = esriExtensionState.esriESUnavailable Then
    '                    Throw New COMException("Not running the appropriate product license.")
    '                End If

    '            ElseIf requestState = esriExtensionState.esriESDisabled Then
    '                'Determine if state can be changed
    '                Dim checkState As esriExtensionState = StateCheck(False)
    '                If (m_enableState <> checkState) Then m_enableState = checkState
    '            End If
    '        End Set
    '    End Property
    '#End Region

    '#Region "Persistence Members"

    '    Public ReadOnly Property ID() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.esriSystem.IPersistVariant.ID
    '        Get
    '            Dim extUID As New UIDClass()
    '            extUID.Value = Me.GetType().GUID.ToString("B") '"b" is a format - not sure what it is exactly from object def
    '            Return extUID
    '        End Get
    '    End Property
    ' this sub is run at the time the mxd document is loaded

    Protected Overloads Overrides Sub OnLoad(ByVal inStrm As Stream)
        'Public Sub Load(ByVal Stream As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Load

        ' ================================================
        ' Purpose: Load Extension settings to property set
        '
        ' 1) Check the first item from the stream
        '   2) If it says "check" then
        '   3) Populate Propertyset with stream properties
        '   4) Set the PUBLIC variable bLoaded = True
        '
        ' NOTE: Be careful when adding new properties then 
        '       accessing a document that has properties saved
        '       from old stream - will not match!

        ' 2MsgBox("removeme OnLoad FIPEX triggered")

        Dim iPolysCount As Integer = 0   ' number of polygon layers currently using
        Dim i As Integer = 0             ' loop counter
        Dim iLinesCount As Integer = 0   ' number of lines layers currently using
        Dim j As Integer = 0             ' loop counter
        Dim iExclusions As Integer = 0
        Dim iBarrierIDs As Integer = 0
        Dim iBarrierEIDCount As Integer = 0
        Dim lBarrierEIDs As List(Of Integer) = New List(Of Integer)
        Dim iFlagEIDCount As Integer = 0

        'Dim sDictionary As New Dictionary(Of String, String)
        Dim sDictionary As New Dictionary(Of String, String)
        Dim bf = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()

        Try
            sDictionary = CType(bf.Deserialize(inStrm), Dictionary(Of String, String))
        Catch ex As Exception
            MsgBox("Error during OnLoad of FIPEX. Document dictionary loading error. " + _
                   "Try opening ArcMap.exe rather than double-clicking MXD document. " + _
                   ex.Message)
        End Try


        ' '' Load the stream of properties into a property set
        ''Dim propcheck As String = Convert.ToString(Stream.Read())
        ''If propcheck = "yes" Then
        'Dim streamCheck As String ' to check if extension settings have been saved in mxd
        '    streamCheck = Convert.ToString(inStrm.)

        If sDictionary.Item("check") = "check" Then
            Try
                With pPropset
                    .SetProperty("direction", sDictionary.Item("direction"))
                    .SetProperty("ordernum", Convert.ToInt32(sDictionary.Item("ordernum")))
                    .SetProperty("maximum", Convert.ToBoolean(sDictionary.Item("maximum")))
                    .SetProperty("connecttab", Convert.ToBoolean(sDictionary.Item("connecttab")))
                    .SetProperty("barrierperm", Convert.ToBoolean(sDictionary.Item("barrierperm")))
                    .SetProperty("naturalyn", Convert.ToBoolean(sDictionary.Item("naturalyn")))
                    .SetProperty("dciyn", Convert.ToBoolean(sDictionary.Item("dciyn")))
                    .SetProperty("dcisectionalyn", Convert.ToBoolean(sDictionary.Item("dcisectionalyn")))
                    .SetProperty("sDCIModelDir", Convert.ToString(sDictionary.Item("sDCIModelDir")))
                    .SetProperty("sRInstallDir", Convert.ToString(sDictionary.Item("sRInstallDir")))

                    '2020
                    .SetProperty("advconnecttab", Convert.ToBoolean(sDictionary.Item("advconnecttab")))
                    .SetProperty("bUseHabLength", Convert.ToBoolean(sDictionary.Item("bUseHabLength")))
                    .SetProperty("bUseHabArea", Convert.ToBoolean(sDictionary.Item("bUseHabArea")))
                    .SetProperty("bDistanceDecay", Convert.ToBoolean(sDictionary.Item("bDistanceDecay")))
                    .SetProperty("bDistanceLim", Convert.ToBoolean(sDictionary.Item("bDistanceLim")))
                    .SetProperty("dMaxDist", Convert.ToDouble(sDictionary.Item("dMaxDist")))
                    .SetProperty("sDDFunction", Convert.ToString(sDictionary.Item("sDDFunction")))

                    .SetProperty("bDBF", Convert.ToBoolean(sDictionary.Item("bDBF")))
                    .SetProperty("sGDB", Convert.ToString(sDictionary.Item("sGDB")))
                    .SetProperty("TabPrefix", Convert.ToString(sDictionary.Item("TabPrefix")))

                    .SetProperty("UpHab", Convert.ToBoolean(sDictionary.Item("UpHab")))
                    .SetProperty("TotalUpHab", Convert.ToBoolean(sDictionary.Item("TotalUpHab")))
                    .SetProperty("DownHab", Convert.ToBoolean(sDictionary.Item("DownHab")))
                    .SetProperty("TotalDownHab", Convert.ToBoolean(sDictionary.Item("TotalDownHab")))
                    .SetProperty("PathDownHab", Convert.ToBoolean(sDictionary.Item("PathDownHab")))
                    .SetProperty("TotalPathDownHab", Convert.ToBoolean(sDictionary.Item("TotalPathDownHab")))

                    .SetProperty("numPolys", Convert.ToInt32(sDictionary.Item("numPolys")))

                    ' Add each polygon and respective hab quan and qual fields as a property
                    ' Add each habitat class and quantity field as property
                    iPolysCount = Convert.ToInt32(pPropset.GetProperty("numPolys"))
                    If iPolysCount > 0 Then
                        i = 0
                        For i = 0 To iPolysCount - 1
                            .SetProperty("IncPoly" + i.ToString, Convert.ToString(sDictionary.Item("IncPoly" + i.ToString)))
                            .SetProperty("PolyClassField" + i.ToString, Convert.ToString(sDictionary.Item("PolyClassField" + i.ToString)))
                            .SetProperty("PolyQuanField" + i.ToString, Convert.ToString(sDictionary.Item("PolyQuanField" + i.ToString)))
                            .SetProperty("PolyUnitField" + i.ToString, Convert.ToString(sDictionary.Item("PolyUnitField" + i.ToString)))
                        Next
                    End If

                    .SetProperty("numLines", Convert.ToInt32(sDictionary.Item("numLines")))

                    ' Add each line included as a property
                    ' Add each habitat class and quantity field as property
                    iLinesCount = Convert.ToInt32(pPropset.GetProperty("numLines"))
                    If iLinesCount > 0 Then
                        j = 0
                        For j = 0 To iLinesCount - 1
                            .SetProperty("IncLine" + j.ToString, Convert.ToString(sDictionary.Item("IncLine" + j.ToString)))

                            .SetProperty("LineLengthField" + j.ToString, Convert.ToString(sDictionary.Item("LineLengthField" + j.ToString)))
                            .SetProperty("LineLengthUnits" + j.ToString, Convert.ToString(sDictionary.Item("LineLengthUnits" + j.ToString)))

                            .SetProperty("LineHabClassField" + j.ToString, Convert.ToString(sDictionary.Item("LineHabClassField" + j.ToString)))
                            .SetProperty("LineHabQuanField" + j.ToString, Convert.ToString(sDictionary.Item("LineHabQuanField" + j.ToString)))
                            .SetProperty("LineHabUnits" + j.ToString, Convert.ToString(sDictionary.Item("LineHabUnits" + j.ToString)))
                        Next
                    End If

                    .SetProperty("numExclusions", Convert.ToInt32(sDictionary.Item("numExclusions")))

                    ' Add each line included as a property
                    ' Add each habitat class and quantity field as property
                    iExclusions = Convert.ToInt32(pPropset.GetProperty("numExclusions"))
                    If iExclusions > 0 Then
                        i = 0
                        For i = 0 To iExclusions - 1
                            .SetProperty("ExcldLayer" + i.ToString, Convert.ToString(sDictionary.Item("ExcldLayer" + i.ToString)))
                            .SetProperty("ExcldFeature" + i.ToString, Convert.ToString(sDictionary.Item("ExcldFeature" + i.ToString)))
                            .SetProperty("ExcldValue" + i.ToString, Convert.ToString(sDictionary.Item("ExcldValue" + i.ToString)))
                        Next
                    End If

                    .SetProperty("numBarrierIDs", Convert.ToInt32(sDictionary.Item("numBarrierIDs")))

                    ' Add each line included as a property
                    ' Add each habitat class and quantity field as property
                    iBarrierIDs = Convert.ToInt32(pPropset.GetProperty("numBarrierIDs"))
                    If iBarrierIDs > 0 Then
                        i = 0
                        For i = 0 To iBarrierIDs - 1
                            .SetProperty("BarrierIDLayer" + i.ToString, Convert.ToString(sDictionary.Item("BarrierIDLayer" + i.ToString)))
                            .SetProperty("BarrierIDField" + i.ToString, Convert.ToString(sDictionary.Item("BarrierIDField" + i.ToString)))
                            .SetProperty("BarrierPermField" + i.ToString, Convert.ToString(sDictionary.Item("BarrierPermField" + i.ToString)))
                            .SetProperty("BarrierNaturalYNField" + i.ToString, Convert.ToString(sDictionary.Item("BarrierNaturalYNField" + i.ToString)))
                        Next
                    End If

                    '.SetProperty("barrierEIDsLoadedyn", Convert.ToBoolean(sDictionary.Item("barrierEIDsLoadedyn")))
                    .SetProperty("barrierEIDsLoadedyn", Convert.ToBoolean("false"))
                    .SetProperty("numBarrierEIDs", Convert.ToInt32(sDictionary.Item("numBarrierEIDs")))

                    iBarrierEIDCount = Convert.ToInt32(pPropset.GetProperty("numBarrierEIDs"))
                    If iBarrierEIDCount > 0 Then
                        i = 0
                        For i = 0 To iBarrierEIDCount - 1
                            .SetProperty("BarrierEID" + i.ToString, Convert.ToInt32(sDictionary.Item("BarrierEID" + i.ToString)))
                        Next
                    End If

                    '.SetProperty("flagEIDsLoadedyn", Convert.ToBoolean(sDictionary.Item("flagEIDsLoadedyn")))
                    .SetProperty("flagEIDsLoadedyn", Convert.ToBoolean("false"))
                    .SetProperty("numFlagEIDs", Convert.ToInt32(sDictionary.Item("numFlagEIDs")))

                    iFlagEIDCount = Convert.ToInt32(pPropset.GetProperty("numFlagEIDs"))
                    If iFlagEIDCount > 0 Then
                        i = 0
                        For i = 0 To iFlagEIDCount - 1
                            .SetProperty("FlagEID" + i.ToString, Convert.ToInt32(sDictionary.Item("FlagEID" + i.ToString)))
                        Next
                    End If

                    Dim bGLPKTables As Boolean

                    bGLPKTables = sDictionary.Item("bGLPKTables")
                    .SetProperty("bGLPKTables", Convert.ToBoolean(sDictionary.Item("bGLPKTables")))
                    .SetProperty("sGLPKModelDir", Convert.ToString(sDictionary.Item("sGLPKModelDir")))
                    .SetProperty("sGnuWinDir", Convert.ToString(sDictionary.Item("sGnuWinDir")))

                End With

            Catch ex As Exception
                MsgBox("Error during Property Set retrieval of FIPEX. Document dictionary loading error. " + _
               "Try opening ArcMap.exe rather than double-clicking MXD document. " + _
               ex.Message)
            End Try

            m_bLoaded = True
        Else
            Try

           
            With pPropset
                ' Load defaults
                .SetProperty("direction", "up")
                .SetProperty("ordernum", Convert.ToInt32(999))
                .SetProperty("maximum", Convert.ToBoolean("True"))
                .SetProperty("connecttab", Convert.ToBoolean("False"))
                .SetProperty("advconnecttab", Convert.ToBoolean("False"))
                .SetProperty("barrierperm", Convert.ToBoolean("False"))
                .SetProperty("naturalyn", Convert.ToBoolean("False"))
                .SetProperty("dciyn", Convert.ToBoolean("False"))
                .SetProperty("dcisectionalyn", Convert.ToBoolean("False"))
                .SetProperty("sDCIModelDir", Convert.ToString("not set"))
                .SetProperty("sRInstallDir", Convert.ToString("not set"))

                    '2020
                    .SetProperty("bUseHabLength", Convert.ToBoolean("True"))
                    .SetProperty("bUseHabArea", Convert.ToBoolean("False"))
                    .SetProperty("bDistanceDecay", Convert.ToBoolean("False"))
                    .SetProperty("bDistanceLim", Convert.ToBoolean("False"))
                    .SetProperty("dMaxDist", Convert.ToDouble(0))
                    .SetProperty("sDDFunction", Convert.ToString("none"))

                .SetProperty("bDBF", Convert.ToBoolean("False"))
                .SetProperty("sGDB", Convert.ToString("not set"))
                .SetProperty("TabPrefix", Convert.ToString("not set"))

                .SetProperty("UpHab", Convert.ToBoolean("True"))
                .SetProperty("TotalUpHab", Convert.ToBoolean("True"))
                .SetProperty("DownHab", Convert.ToBoolean("False"))
                .SetProperty("TotalDownHab", Convert.ToBoolean("False"))
                .SetProperty("PathDownHab", Convert.ToBoolean("False"))
                .SetProperty("TotalPathDownHab", Convert.ToBoolean("False"))

                .SetProperty("numPolys", Convert.ToInt32(0))
                .SetProperty("numLines", Convert.ToInt32(0))
                .SetProperty("numExclusions", Convert.ToInt32(0))
                .SetProperty("numBarrierIDs", Convert.ToInt32(0))
                .SetProperty("barrierEIDsLoadedyn", Convert.ToBoolean("False"))
                .SetProperty("numBarrierEIDs", Convert.ToInt32(0))
                .SetProperty("flagEIDsLoadedyn", Convert.ToBoolean("False"))
                .SetProperty("numFlagEIDs", Convert.ToInt32(0))

                .SetProperty("bGLPKTables", Convert.ToBoolean("False"))
                .SetProperty("sGLPKModelDir", Convert.ToString("not set"))
                .SetProperty("sGnuWinDir", Convert.ToString("not set"))

            End With

            Catch ex As Exception
                MsgBox("Error during Property Set retrieval of FIPEX. Document dictionary loading error. " + _
             "Try opening ArcMap.exe rather than double-clicking MXD document. Code 4. " + _
             ex.Message)
            End Try
            m_bLoaded = True
        End If

    End Sub
    ' this sub is run at the time the document is saved
    Protected Overloads Overrides Sub OnSave(ByVal outStrm As Stream)
        'As ESRI.ArcGIS.esriSystem.IVariantStream) Implements ESRI.ArcGIS.esriSystem.IPersistVariant.Save

        Dim iPolysCount As Integer = 0   ' number of polygon layers currently using
        Dim i As Integer = 0             ' loop counter
        Dim iLinesCount As Integer = 0   ' number of lines layers currently using
        Dim j As Integer = 0             ' loop counter
        Dim iExclusions As Integer = 0
        Dim iBarrierIDs As Integer = 0

        Dim sDictionary As New Dictionary(Of String, String)
        Dim iEID As Integer = 0
        Dim iBarrierEIDCount As Integer = 0
        Dim lBarrierEIDs As List(Of Integer) = New List(Of Integer)
        Dim iFlagEIDCount As Integer = 0
        Dim lFlagEIDs As List(Of Integer) = New List(Of Integer)

        ' =========  BEGIN GRAB OF BARRIERS IN CURRENT NETWORK =========
        ' this section sets the property set with barriers in current network. 

        ' THIS NEEDS CHECKING ON
        If Me.State = ExtensionState.Enabled Then

            Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension

            Dim pUtilityNetworkAnalysisExt As IUtilityNetworkAnalysisExt
            pUtilityNetworkAnalysisExt = FishPassageExtension.GetUNAExt

            Dim pGeometricNetwork As IGeometricNetwork

            ' Get the active network
            ' The network shouldn't be empty because the extension wouldn't be enabled otherwise
            Dim pNetworkAnalysisExt As INetworkAnalysisExt
            pNetworkAnalysisExt = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExt)

            ' If there is a network, and only one network then 
            ' save the barriers of the current network
            ' save the flags of the current network
            If pNetworkAnalysisExt.NetworkCount <> 0 Then

                Try
                    pGeometricNetwork = pNetworkAnalysisExt.CurrentNetwork

                    ' Get the network elements for EID retrieval
                    Dim pNetwork As INetwork
                    Dim pNetElements As INetElements
                    pNetwork = pGeometricNetwork.Network
                    pNetElements = CType(pNetwork, INetElements)

                    'Dim pOriginalBarriersListGEN As IEnumNetEIDBuilderGEN
                    'pOriginalBarriersListGEN = New EnumNetEIDArray
                    Dim pNetworkAnalysisExtBarriers As INetworkAnalysisExtBarriers
                    Dim pFlagDisplay As IFlagDisplay

                    'pNetworkAnalysisExtFlags = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
                    pNetworkAnalysisExtBarriers = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtBarriers)

                    '' coclass of geomtricnetwork
                    'Dim pGeometricNetworkName As IGeometricNetworkName
                    'pGeometricNetworkName = CType(pGeometricNetwork, IGeometricNetworkName)

                    ' Get barriers from network (junction only)
                    i = 0
                    iBarrierEIDCount = pNetworkAnalysisExtBarriers.JunctionBarrierCount
                    For i = 0 To pNetworkAnalysisExtBarriers.JunctionBarrierCount - 1
                        ' Use pFlagDisplay to retrieve EIDs of the barriers for later
                        pFlagDisplay = CType(pNetworkAnalysisExtBarriers.JunctionBarrier(i), IFlagDisplay)
                        iEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
                        lBarrierEIDs.Add(iEID)
                        'pOriginalBarriersListGEN.Add(iEID)
                        'originalBarriersList(i) = bEID
                    Next


                    'Dim pOriginalflagsListGEN As IEnumNetEIDBuilderGEN
                    'pOriginalflagsListGEN = New EnumNetEIDArray
                    Dim pNetworkAnalysisExtflags As INetworkAnalysisExtFlags

                    'pNetworkAnalysisExtFlags = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)
                    pNetworkAnalysisExtflags = CType(pUtilityNetworkAnalysisExt, INetworkAnalysisExtFlags)

                    '' coclass of geomtricnetwork
                    'Dim pGeometricNetworkName As IGeometricNetworkName
                    'pGeometricNetworkName = CType(pGeometricNetwork, IGeometricNetworkName)

                    ' Get flags from network (junction only)
                    i = 0
                    iFlagEIDCount = pNetworkAnalysisExtflags.JunctionFlagCount
                    For i = 0 To pNetworkAnalysisExtflags.JunctionFlagCount - 1
                        ' Use pFlagDisplay to retrieve EIDs of the flags for later
                        pFlagDisplay = CType(pNetworkAnalysisExtflags.JunctionFlag(i), IFlagDisplay)
                        iEID = pNetElements.GetEID(pFlagDisplay.FeatureClassID, pFlagDisplay.FID, pFlagDisplay.SubID, esriElementType.esriETJunction)
                        lFlagEIDs.Add(iEID)
                        'pOriginalBarriersListGEN.Add(iEID)
                        'originalBarriersList(i) = bEID
                    Next
                Catch ex As Exception
                    MsgBox("Debug2020 - Error saving FIPEX options 1/3")
                    Exit Sub
                End Try
               

            End If


            ' =========  END GRAB OF BARRIERS IN CURRENT NETWORK ===========

            ' Write to pPropset now for immediate retrieval
            ' seems redundant but for saving of barriers to document stream it's correct

            Try
                ' Save the barrier count and iEIDs if any
                With pPropset
                    .SetProperty("numBarrierEIDs", Convert.ToInt32(iBarrierEIDCount))
                    ' Add each EID to the propertyset if there are any
                    If iBarrierEIDCount > 0 Then
                        i = 0
                        For i = 0 To lBarrierEIDs.Count - 1
                            .SetProperty("BarrierEID" + i.ToString, lBarrierEIDs(i))
                        Next
                    End If

                    .SetProperty("numFlagEIDs", iFlagEIDCount)
                    ' Add each EID to the propertyset if there are any
                    If iFlagEIDCount > 0 Then
                        i = 0
                        For i = 0 To lFlagEIDs.Count - 1
                            .SetProperty("flagEID" + i.ToString, lFlagEIDs(i))
                        Next
                    End If

                End With
            Catch ex As Exception
                MsgBox("Debug2020 - Error saving FIPEX options 2/3")
                Exit Sub
            End Try

            ' ========== BEGIN ADD PROPERTIES TO DICTIONARY OBJECT =============
            'If there are any properties to write then save them to the stream
            If m_bLoaded = True Then
                Try
                    'MsgBox("Debug2020 - Testing Save of Document 3/3")

                    ' load the properties into a data dictionary for serialization to the stream
                    ' (there may be a better way to do this, but for add-ins the system.io.stream
                    ' seems to limit us to a binary format for write/reads
                    ' all of these are converted to string format to keep it simple
                    ' some of them will be strings already but ohhh wellll

                    sDictionary.Add("check", "check")
                    sDictionary.Add("direction", Convert.ToString(pPropset.GetProperty("direction")))
                    sDictionary.Add("ordernum", Convert.ToString(pPropset.GetProperty("ordernum")))
                    sDictionary.Add("maximum", Convert.ToString(pPropset.GetProperty("maximum")))
                    sDictionary.Add("connecttab", Convert.ToString(pPropset.GetProperty("connecttab")))
                    sDictionary.Add("barrierperm", Convert.ToString(pPropset.GetProperty("barrierperm")))
                    sDictionary.Add("naturalyn", Convert.ToString(pPropset.GetProperty("naturalyn")))
                    sDictionary.Add("dciyn", Convert.ToString(pPropset.GetProperty("dciyn")))
                    sDictionary.Add("dcisectionalyn", Convert.ToString(pPropset.GetProperty("dcisectionalyn")))
                    sDictionary.Add("sDCIModelDir", Convert.ToString(pPropset.GetProperty("sDCIModelDir")))

                    '2020
                    sDictionary.Add("advconnecttab", Convert.ToString(pPropset.GetProperty("advconnecttab")))

                    sDictionary.Add("bUseHabLength", Convert.ToString(pPropset.GetProperty("bUseHabLength")))
                    sDictionary.Add("bUseHabArea", Convert.ToString(pPropset.GetProperty("bUseHabArea")))

                    sDictionary.Add("bDistanceDecay", Convert.ToString(pPropset.GetProperty("bDistanceDecay")))
                    sDictionary.Add("bDistanceLim", Convert.ToString(pPropset.GetProperty("bDistanceLim")))
                    sDictionary.Add("dMaxDist", Convert.ToString(pPropset.GetProperty("dMaxDist")))
                    sDictionary.Add("sDDFunction", Convert.ToString(pPropset.GetProperty("sDDFunction")))

                    sDictionary.Add("sRInstallDir", Convert.ToString(pPropset.GetProperty("sRInstallDir")))
                    sDictionary.Add("bDBF", Convert.ToString(pPropset.GetProperty("bDBF")))
                    sDictionary.Add("sGDB", Convert.ToString(pPropset.GetProperty("sGDB")))
                    sDictionary.Add("TabPrefix", Convert.ToString(pPropset.GetProperty("TabPrefix")))
                    sDictionary.Add("UpHab", Convert.ToString(pPropset.GetProperty("UpHab")))
                    sDictionary.Add("TotalUpHab", Convert.ToString(pPropset.GetProperty("TotalUpHab")))
                    sDictionary.Add("DownHab", Convert.ToString(pPropset.GetProperty("DownHab")))
                    sDictionary.Add("TotalDownHab", Convert.ToString(pPropset.GetProperty("TotalDownHab")))
                    sDictionary.Add("PathDownHab", Convert.ToString(pPropset.GetProperty("PathDownHab")))
                    sDictionary.Add("TotalPathDownHab", Convert.ToString(pPropset.GetProperty("TotalPathDownHab")))
                    sDictionary.Add("numPolys", Convert.ToString(pPropset.GetProperty("numPolys")))

                    iPolysCount = Convert.ToInt32(pPropset.GetProperty("numPolys"))
                    If iPolysCount > 0 Then
                        i = 0
                        For i = 0 To iPolysCount - 1
                            sDictionary.Add("IncPoly" + i.ToString, Convert.ToString(pPropset.GetProperty("IncPoly" + i.ToString)))
                            sDictionary.Add("PolyClassField" + i.ToString, Convert.ToString(pPropset.GetProperty("PolyClassField" + i.ToString)))
                            sDictionary.Add("PolyQuanField" + i.ToString, Convert.ToString(pPropset.GetProperty("PolyQuanField" + i.ToString)))
                            sDictionary.Add("PolyUnitField" + i.ToString, Convert.ToString(pPropset.GetProperty("PolyUnitField" + i.ToString)))
                        Next
                    End If

                    sDictionary.Add("numLines", Convert.ToString(pPropset.GetProperty("numLines")))

                    iLinesCount = Convert.ToInt32(pPropset.GetProperty("numLines"))
                    If iLinesCount > 0 Then
                        j = 0
                        For j = 0 To iLinesCount - 1
                            sDictionary.Add("IncLine" + j.ToString, Convert.ToString(pPropset.GetProperty("IncLine" + j.ToString)))
                            sDictionary.Add("LineLengthField" + j.ToString, Convert.ToString(pPropset.GetProperty("LineLengthField" + j.ToString)))
                            sDictionary.Add("LineLengthUnits" + j.ToString, Convert.ToString(pPropset.GetProperty("LineLengthUnits" + j.ToString)))
                            sDictionary.Add("LineHabClassField" + j.ToString, Convert.ToString(pPropset.GetProperty("LineHabClassField" + j.ToString)))
                            sDictionary.Add("LineHabQuanField" + j.ToString, Convert.ToString(pPropset.GetProperty("LineHabQuanField" + j.ToString)))
                            sDictionary.Add("LineHabUnits" + j.ToString, Convert.ToString(pPropset.GetProperty("LineHabUnits" + j.ToString)))
                        Next
                    End If

                    sDictionary.Add("numExclusions", Convert.ToString(pPropset.GetProperty("numExclusions")))

                    iExclusions = Convert.ToInt32(pPropset.GetProperty("numExclusions"))
                    If iExclusions > 0 Then
                        i = 0
                        For i = 0 To iExclusions - 1

                            sDictionary.Add("ExcldLayer" + i.ToString, Convert.ToString(pPropset.GetProperty("ExcldLayer" + i.ToString)))
                            sDictionary.Add("ExcldFeature" + i.ToString, Convert.ToString(pPropset.GetProperty("ExcldFeature" + i.ToString)))
                            sDictionary.Add("ExcldValue" + i.ToString, Convert.ToString(pPropset.GetProperty("ExcldValue" + i.ToString)))
                        Next
                    End If

                    sDictionary.Add("numBarrierIDs", Convert.ToString(pPropset.GetProperty("numBarrierIDs")))
                    iBarrierIDs = Convert.ToInt32(pPropset.GetProperty("numBarrierIDs"))
                    If iBarrierIDs > 0 Then
                        i = 0
                        For i = 0 To iBarrierIDs - 1

                            sDictionary.Add("BarrierIDLayer" + i.ToString, Convert.ToString(pPropset.GetProperty("BarrierIDLayer" + i.ToString)))
                            sDictionary.Add("BarrierIDField" + i.ToString, Convert.ToString(pPropset.GetProperty("BarrierIDField" + i.ToString)))
                            sDictionary.Add("BarrierPermField" + i.ToString, Convert.ToString(pPropset.GetProperty("BarrierPermField" + i.ToString)))
                            sDictionary.Add("BarrierNaturalYNField" + i.ToString, Convert.ToString(pPropset.GetProperty("BarrierNaturalYNField" + i.ToString)))
                        Next
                    End If

                    sDictionary.Add("barrierEIDsLoadedyn", "false")
                    sDictionary.Add("numBarrierEIDs", Convert.ToString(pPropset.GetProperty("numBarrierEIDs")))

                    iBarrierEIDCount = Convert.ToInt32(pPropset.GetProperty("numBarrierEIDs"))

                    If iBarrierEIDCount > 0 Then
                        i = 0
                        For i = 0 To iBarrierEIDCount - 1
                            sDictionary.Add("BarrierEID" + i.ToString, Convert.ToString(pPropset.GetProperty("BarrierEID" + i.ToString)))
                        Next
                    End If

                    sDictionary.Add("flagEIDsLoadedyn", "false")
                    sDictionary.Add("numFlagEIDs", Convert.ToString(pPropset.GetProperty("numFlagEIDs")))

                    iFlagEIDCount = Convert.ToInt32(pPropset.GetProperty("numFlagEIDs"))

                    If iFlagEIDCount > 0 Then
                        i = 0
                        For i = 0 To iFlagEIDCount - 1
                            sDictionary.Add("FlagEID" + i.ToString, Convert.ToString(pPropset.GetProperty("FlagEID" + i.ToString)))
                        Next
                    End If

                    sDictionary.Add("bGLPKTables", Convert.ToString(pPropset.GetProperty("bGLPKTables")))
                    sDictionary.Add("sGLPKModelDir", Convert.ToString(pPropset.GetProperty("sGLPKModelDir")))
                    sDictionary.Add("sGnuWinDir", Convert.ToString(pPropset.GetProperty("sGnuWinDir")))

                Catch ex As Exception
                    ' 2020 added all these tries because when the above code fails
                    ' some properties have already been set, so adding a second time triggers
                    ' an unhandled failure
                    Try
                        sDictionary.Add("check", "check")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("direction", "up")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("ordernum", Convert.ToString(999))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("maximum", Convert.ToString("True"))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("connecttab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("advconnecttab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("barrierperm", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("naturalyn", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("dciyn", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("dcisectionalyn", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sDCIModelDir", "not set")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sRInstallDir", "not set")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bUseHabLength", Convert.ToString(True))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bUseHabArea", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bDistanceDecay", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bDistanceLim", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("dMaxDist", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sDDFunction", "none")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("advconnecttab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bDBF", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sGDB", "not set")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("TabPrefix", "not set")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("TabPrefix", "not set")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("UpHab", Convert.ToString(True))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("TotalUpHab", Convert.ToString(True))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("DownHab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("TotalDownHab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("PathDownHab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("TotalPathDownHab", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numPolys", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numLines", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numExclusions", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numBarrierIDs", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("barrierEIDsLoadedyn", "false")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numBarrierEIDs", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("flagEIDsLoadedyn", "false")
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("numFlagEIDs", Convert.ToString(0))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("bGLPKTables", Convert.ToString(False))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sGLPKModelDir", Convert.ToString("not set"))
                    Catch ex1 As Exception

                    End Try
                    Try
                        sDictionary.Add("sGnuWinDir", Convert.ToString("not set"))
                    Catch ex1 As Exception

                    End Try
                    
                End Try
            Else ' if (for some weird reason) the extension hasn't been loaded
                '  (the bloaded var is set as false) then save defaults to stream

                sDictionary.Add("check", "check")
                sDictionary.Add("direction", "up")
                sDictionary.Add("ordernum", Convert.ToString(999))
                sDictionary.Add("maximum", Convert.ToString("True"))
                sDictionary.Add("connecttab", Convert.ToString(False))
                sDictionary.Add("advconnecttab", Convert.ToString(False))
                sDictionary.Add("barrierperm", Convert.ToString(False))
                sDictionary.Add("naturalyn", Convert.ToString(False))
                sDictionary.Add("dciyn", Convert.ToString(False))
                sDictionary.Add("dcisectionalyn", Convert.ToString(False))
                sDictionary.Add("sDCIModelDir", "not set")
                sDictionary.Add("sRInstallDir", "not set")

                '2020
                sDictionary.Add("bUseHabLength", Convert.ToString(True))
                sDictionary.Add("bUseHabArea", Convert.ToString(False))
                sDictionary.Add("bDistanceDecay", Convert.ToString(False))
                sDictionary.Add("bDistanceLim", Convert.ToString(False))
                sDictionary.Add("dMaxDist", Convert.ToString(0))
                sDictionary.Add("sDDFunction", "none")
                sDictionary.Add("advconnecttab", Convert.ToString(False))


                sDictionary.Add("bDBF", Convert.ToString(False))
                sDictionary.Add("sGDB", "not set")
                sDictionary.Add("TabPrefix", "not set")
                sDictionary.Add("UpHab", Convert.ToString(True))
                sDictionary.Add("TotalUpHab", Convert.ToString(True))
                sDictionary.Add("DownHab", Convert.ToString(False))
                sDictionary.Add("TotalDownHab", Convert.ToString(False))
                sDictionary.Add("PathDownHab", Convert.ToString(False))
                sDictionary.Add("TotalPathDownHab", Convert.ToString(False))
                sDictionary.Add("numPolys", Convert.ToString(0))
                sDictionary.Add("numLines", Convert.ToString(0))
                sDictionary.Add("numExclusions", Convert.ToString(0))
                sDictionary.Add("numBarrierIDs", Convert.ToString(0))
                sDictionary.Add("barrierEIDsLoadedyn", "false")
                sDictionary.Add("numBarrierEIDs", Convert.ToString(0))
                sDictionary.Add("flagEIDsLoadedyn", "false")
                sDictionary.Add("numFlagEIDs", Convert.ToString(0))

                sDictionary.Add("bGLPKTables", Convert.ToString(False))
                sDictionary.Add("sGLPKModelDir", Convert.ToString("not set"))
                sDictionary.Add("sGnuWinDir", Convert.ToString("not set"))

            End If
            ' ========== END ADD PROPERTIES TO DICTIONARY OBJECT =============

            Dim bf = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
            bf.Serialize(outStrm, sDictionary)

        End If ' extenson is on

    End Sub


    Friend Function CheckNetworkCount() As Boolean
        ' 4MsgBox("removeme CheckNetworkCount FIPEX triggered")
        'Dim FiPEx__1 As FishPassageExtension = FishPassageExtension.GetExtension

        Dim pNetworkAnalysisExt As INetworkAnalysisExt

        If m_UNAextension Is Nothing Then
            Return False
        End If

        Try
            pNetworkAnalysisExt = CType(m_UNAextension, INetworkAnalysisExt)

            If pNetworkAnalysisExt.NetworkCount > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("Error during network count check (FIPEX). " + ex.Message)
        End Try
        

    End Function

    '#End Region

    Private Sub MapEvents_ContentsChanged()

        ' 4MsgBox("removeme MapEvents_ContentsChanged FIPEX triggered")
        m_bHasNetworks = CheckNetworkCount()
    End Sub


    Friend Function HasNetwork() As Boolean

        ' 4MsgBox("removeme HasNetwork FIPEX triggered")
        Return m_bHasNetworks
    End Function

    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
        Catch ex As Exception
            MsgBox("Error during FIPEX 'finalize' process. Code 20. " + ex.Message)
        End Try
    End Sub
End Class

'End Namespace


