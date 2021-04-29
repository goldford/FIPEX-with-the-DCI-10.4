Imports System.Drawing
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.EditorExt
Imports ESRI.ArcGIS.NetworkAnalysis

Public Class SharedSubs
    Public Shared Sub ResultsForm2020(ByRef pResultsForm3 As FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3,
                                      ByRef lSinkIDandTypes As List(Of SinkandTypes),
                                      ByRef lHabStatsList As List(Of StatisticsObject_2),
                                      ByRef lMetricsObject As List(Of MetricsObject),
                                      ByRef BeginTime As DateTime, ByRef numbarrsnodes As String,
                                      ByRef iOrderNum As Integer, sDirection As String)
        ' col 0 - sink ID
        ' col 1 - sink EID
        ' col 2 - sink node type (barrier / junction )
        ' col 3 - barrier ID
        ' col 4 - barrier EID
        ' col 5 - stat (e.g., perm, DCI)
        ' col 6 - trace type (e.g. upstream)
        ' col 7 - trace subtype (e.g., immediate)
        ' col 8 - class
        ' col 9 - value
        ' col 10 - units
        ' col 11 - hab_dimension (can be length, area)
        ' col 10 - units

        ' Output Form (will replace dockable window)
        'Dim pResultsForm3 As New FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.frmResults_3
        Dim iSinkRowIndex As Integer


        ' Set up the table - create columns 
        pResultsForm3.DataGridView1.Columns.Add("SinkID", "Sink User ID")             '0
        pResultsForm3.DataGridView1.Columns.Add("SinkEID", "Sink Net EID")           '1
        pResultsForm3.DataGridView1.Columns.Add("SinkNodeType", "Sink Node Type") '2
        pResultsForm3.DataGridView1.Columns.Add("BarrierID", "Barrier User ID")       '3
        pResultsForm3.DataGridView1.Columns.Add("BarrierEID", "Barrier Net EID")     '4
        pResultsForm3.DataGridView1.Columns.Add("Stat", "Statistic")                 '5
        pResultsForm3.DataGridView1.Columns.Add("TraceType", "Trace Type")       '6
        pResultsForm3.DataGridView1.Columns.Add("TraceSubtype", "Trace Subtype") '7
        pResultsForm3.DataGridView1.Columns.Add("class", "class")               '8
        pResultsForm3.DataGridView1.Columns.Add("value", "value")               '9
        pResultsForm3.DataGridView1.Columns.Add("units", "units")               '10
        pResultsForm3.DataGridView1.Columns.Add("dimension", "dimension")       '11
        pResultsForm3.DataGridView1.Columns.Add("layer", "layer")               '12

        pResultsForm3.DataGridView1.Columns(0).Width = 46
        pResultsForm3.DataGridView1.Columns(1).Width = 46
        pResultsForm3.DataGridView1.Columns(2).Width = 46
        pResultsForm3.DataGridView1.Columns(3).Width = 46
        pResultsForm3.DataGridView1.Columns(4).Width = 46
        pResultsForm3.DataGridView1.Columns(5).Width = 72
        pResultsForm3.DataGridView1.Columns(6).Width = 65
        pResultsForm3.DataGridView1.Columns(7).Width = 65
        pResultsForm3.DataGridView1.Columns(8).Width = 90
        pResultsForm3.DataGridView1.Columns(9).Width = 46
        pResultsForm3.DataGridView1.Columns(10).Width = 46
        pResultsForm3.DataGridView1.Columns(11).Width = 65
        pResultsForm3.DataGridView1.Columns(12).Width = 75


        For i = 0 To lSinkIDandTypes.Count - 1

            'pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Style
            'pDataGridViewCellStyle = pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Style
            'pDataGridViewCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, 14, FontStyle.Bold)
            pResultsForm3.DataGridView1.AllowUserToResizeColumns = True
            pResultsForm3.DataGridView1.AllowUserToResizeRows = True
            pResultsForm3.DataGridView1.ColumnHeadersDefaultCellStyle.Font = New Font(pResultsForm3.DataGridView1.Font.FontFamily, 8, FontStyle.Bold)

            For j = 0 To lMetricsObject.Count - 1
                If lMetricsObject(j).SinkEID = lSinkIDandTypes(i).SinkEID Then
                    iSinkRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Value = lSinkIDandTypes(i).SinkID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Value = lSinkIDandTypes(i).SinkEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(2).Value = lSinkIDandTypes(i).Type
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(3).Value = lMetricsObject(j).ID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(4).Value = lMetricsObject(j).BarrEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(5).Value = lMetricsObject(j).MetricName
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(6).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(7).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(8).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(9).Value = Math.Round(lMetricsObject(j).Metric, 2)
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(10).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(11).Value = "-"
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(12).Value = "-"
                End If
            Next

            ' for each habitat object item add the stats (hab by class, etc)
            For j = 0 To lHabStatsList.Count - 1
                If lHabStatsList(j).SinkEID = lSinkIDandTypes(i).SinkEID Then
                    iSinkRowIndex = pResultsForm3.DataGridView1.Rows.Add()
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(0).Value = lSinkIDandTypes(i).SinkID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(1).Value = lSinkIDandTypes(i).SinkEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(2).Value = lSinkIDandTypes(i).Type
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(3).Value = lHabStatsList(j).bID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(4).Value = lHabStatsList(j).bEID
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(5).Value = lHabStatsList(j).LengthOrHabitat
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(6).Value = lHabStatsList(j).Direction
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(7).Value = lHabStatsList(j).TotalImmedPath
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(8).Value = lHabStatsList(j).UniqueClass
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(9).Value = Math.Round(lHabStatsList(j).Quantity, 2)
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(10).Value = lHabStatsList(j).Unit
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(11).Value = lHabStatsList(j).HabitatDimension
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(12).Value = lHabStatsList(j).Layer

                End If
            Next
        Next

        Dim EndTime As DateTime
        EndTime = DateTime.Now

        Dim TotalTime As TimeSpan
        TotalTime = EndTime - BeginTime
        pResultsForm3.lblBeginTime.Text = "Begin Time: " & BeginTime
        pResultsForm3.lblEndtime.Text = "End Time: " & EndTime
        pResultsForm3.lblTotalTime.Text = "Total Time: " & TotalTime.Hours & "hrs " & TotalTime.Minutes & "minutes " & TotalTime.Seconds & "seconds"
        pResultsForm3.lblDirection.Text = "Analysis Direction: " + sDirection
        If iOrderNum <> 99999 Then
            pResultsForm3.lblOrder.Text = "Order of Analysis: " & CStr(iOrderNum)
        Else
            pResultsForm3.lblOrder.Text = "Order of Analysis: Max (all nodes in analysis direction)"
        End If

        If Not numbarrsnodes Is Nothing Then
            pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: " & numbarrsnodes
        Else
            pResultsForm3.lblNumBarriers.Text = "Number of Barriers / Nodes Analysed: 1"
        End If
        pResultsForm3.BringToFront()

    End Sub
    Public Shared Sub exclusions2020(ByRef bExclude As Boolean, ByRef pFeature As IFeature, ByRef pFeatureLayer As IFeatureLayer)

        ' =============================================
        ' ============== EXCLUSIONS 2020 ==============
        Dim e_FiPEx__1 As FishPassageExtension
        If e_FiPEx__1 Is Nothing Then
            e_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If

        Dim iExclusions, j As Integer
        Dim plExclusions As List(Of LayerToExclude) = New List(Of LayerToExclude)
        If e_FiPEx__1.m_bLoaded = True Then ' If there were any extension settings set

            iExclusions = Convert.ToInt32(e_FiPEx__1.pPropset.GetProperty("numExclusions"))
            Dim ExclusionsObj As New LayerToExclude(Nothing, Nothing, Nothing)

            ' match any of the line layers saved in stream to those in listboxes
            If iExclusions > 0 Then
                For j = 0 To iExclusions - 1
                    'sLineLayer = m_FiPEX__1.pPropset.GetProperty("IncLine" + j.ToString) ' get line layer
                    ExclusionsObj = New LayerToExclude(Nothing, Nothing, Nothing)
                    With ExclusionsObj
                        '.Layer = sLineLayer
                        .Layer = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldLayer" + j.ToString))
                        .Feature = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldFeature" + j.ToString))
                        .Value = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("ExcldValue" + j.ToString))
                    End With

                    ' add to the module level list
                    plExclusions.Add(ExclusionsObj)
                Next
            End If
        Else
            MsgBox("The FIPEX Options haven't been loaded.  Exiting Exclusions Subroutine. FIPEX code 2642. ")
            Exit Sub
        End If

        Dim x As Integer = 0
        Dim iFieldVal As Integer
        Dim vVal As Object
        Dim sTempVal As String
        For x = 0 To iExclusions - 1
            If pFeatureLayer.Name = plExclusions(x).Layer Then
                ' try to find the field
                iFieldVal = pFeature.Fields.FindField(plExclusions(x).Feature)
                If iFieldVal <> -1 Then
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

                    If sTempVal IsNot Nothing Then
                        If sTempVal = plExclusions(x).Value Then
                            bExclude = True
                        End If
                    End If
                End If
            End If
        Next

    End Sub
    Public Shared Sub calculateStatistics_2020(ByRef lHabStatsList As List(Of StatisticsObject_2), _
                                        ByRef lLineLayersFields As List(Of LineLayerToAdd), _
                                        ByRef lPolyLayersFields As List(Of PolyLayerToAdd), _
                                        ByRef ID As String, _
                                        ByRef iEID As Integer, _
                                        ByRef sType As String, _
                                        ByRef f_sOutID As String, _
                                        ByRef f_siOutEID As Integer,
                                        ByVal sHabTypeKeyword As String, _
                                        ByVal sDirection2 As String)

        ' **************************************************************************************
        ' Subroutine:  Calculate Statistics (2) 
        ' Author:       G Oldford
        ' Purpose:     1) intersect other included layers with returned selection
        '                 from the trace.
        '              2) calculate habitat area and length using habitat classes 
        '                 and excluding unwanted features
        '              3) get array (matrix?) of statistics for each habitat class and
        '                 each layer included for habitat classification stats
        '              4) update statistics object and send back to onclick
        ' Keywords:    sHabTypeKeyword - "Total", "Immediate", or "Path"
        '
        '
        ' Notes:
        ' 
        '       Aug, 2020    --> Not sure why passing vars by ref other than lHabStatsList.
        '                        Deleted 'sKeyword' arg (checks if flag on barr or nonbarr) - it was unused. 
        '                        The object returned must differentiate between 'length' and 'habitat' for DD,  
        '                        so changed object by adding a 'LengthOrHabitat' param. 
        '                        Since object returned is used in output tables (not only for DCI) then I must 
        '                        keep all habitat returned (not just either line or poly habitat depending on user
        '                        choice). introduced object to tag the habitat as 'length' or 
        '                        'area' based on whether it is drawn from polygon or line layer. It's probably 
        '                        a better option long-term to use the units as the basis for differentiating 
        '                        whether a quantity returned from a feature represents habitat area or habitat length. 
        '                        In other words, right now the TOC layer type determines whether the habitat extracted from
        '                        the TOC layer is 'area' or 'line'
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
        Dim e_FiPEx__1 As FishPassageExtension
        If e_FiPEx__1 Is Nothing Then
            e_FiPEx__1 = FiPEX_ArcMap_10p4_up_AddIn_dotNet45_2020.FishPassageExtension.GetExtension()
        End If

        Dim pUID As New UID
        ' Get the pUID of the SelectByLayer command
        'pUID.Value = "{82B9951B-DD63-11D1-AA7F-00C04FA37860}"
        'Dim pGp As IGeoProcessor
        'pGp = New ESRI.ArcGIS.Geoprocessing.GeoProcessor
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
        Dim lHabStatsMatrix As New List(Of HabStatisticsObject)
        Dim pHabStatisticsObject As New HabStatisticsObject(Nothing, Nothing)

        'Dim pFeatureWkSp As IFeatureWorkspace
        Dim pDataStats As IDataStatistics
        Dim pCursor As ICursor
        Dim vFeatHbClsVl As Object ' Feature Habitat Class Value (an object because classes can be numbers or string)
        Dim vTemp As Object
        Dim sFeatClassVal As String
        Dim sMatrixVal As String
        Dim dHabArea, dHabLength As Double
        Dim bClassFound As Boolean
        'For k = 1 To UBound(mHabClassVals, 2) vb6
        Dim classComparer As FindStatsClassPredicate2020
        Dim iStatsMatrixIndex As Integer ' for refining statistics list 
        Dim sClass As String
        Dim vHabTemp As Object
        Dim pDoc As IDocument = My.ArcMap.Application.Document
        pMxDoc = CType(pDoc, IMxDocument)
        pMxDocument = CType(pDoc, IMxDocument)
        pMap = pMxDocument.FocusMap

        ' 2020 - change this two separate objects, lines polygons
        ' layer to hold parameters to send to property
        Dim PolyHabLayerObj As New PolyLayerToAdd(Nothing, Nothing, Nothing, Nothing)
        Dim LineHabLayerObj As New LineLayerToAdd(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)

        ' object to hold stats to add to list. 
        Dim pHabStatsObject_2 As New StatisticsObject_2(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        Dim sDirection As String
        sDirection = Convert.ToString(e_FiPEx__1.pPropset.GetProperty("direction"))

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

        ' Look at the next layer in the list
        pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Dim iClassCheckTemp, iLengthField As Integer
        Dim iLoopCount As Integer = 0
        Dim dTempQuan As Double = 0
        Dim dTotalLength As Double = 0
        Dim bExclude As Boolean = False

        Do While Not pFeatureLayer Is Nothing ' these two lines must be separate
            If pFeatureLayer.Valid = True Then ' or there will be an empty object ref

                pFeatureSelection = CType(pFeatureLayer, IFeatureSelection)
                pSelectionSet = pFeatureSelection.SelectionSet

                ' get the fields from the featureclass
                pFields = pFeatureLayer.FeatureClass.Fields
                j = 0

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
                        ElseIf sUnit = "None" Then
                            sUnit = "none"
                        Else
                            sUnit = "n/a"
                        End If

                        'MsgBox("Debug 2020: Check for the habitat class field for line layer. Is it <none> or 'not set'?: " & lLineLayersFields(j).HabClsField)

                        Try
                            iLengthField = pFields.FindField(lLineLayersFields(j).LengthField)
                        Catch ex As Exception
                            MsgBox("Error finding the field in line layer for length. FIPEX code 971")
                            Exit Sub
                        End Try
                        ' 

                        ' if we find class field being used then use an intermediate object pHabStatisticsObject
                        ' then use a list of these objects (lHabStatsMatrix) to keep track of total habitat by class
                        iClassCheckTemp = pFields.FindField(lLineLayersFields(j).HabClsField)
                        If iClassCheckTemp <> -1 And lLineLayersFields(j).HabClsField <> "<None>" _
                            And lLineLayersFields(j).HabClsField <> "Not set" And lLineLayersFields(j).HabClsField <> "<none>" Then

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

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then
                                        ' =========================================
                                        ' ============ HABITAT STATS ==============
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
                                            dHabLength = Convert.ToDouble(vHabTemp)
                                        Catch ex As Exception
                                            MsgBox("The Habitat Quantity found in the " & lLineLayersFields(j).Layer & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. " & ex.Message)
                                            dHabLength = 0
                                        End Try

                                        classComparer = New FindStatsClassPredicate2020(sClass)
                                        ' use the layer and sink ID to get a refined list of habitat stats for 
                                        ' this sink, layer combo
                                        iStatsMatrixIndex = lHabStatsMatrix.FindIndex(AddressOf classComparer.CompareStatsClass)
                                        If iStatsMatrixIndex = -1 Then
                                            bClassFound = False
                                            pHabStatisticsObject = New HabStatisticsObject(sClass, dHabLength)
                                            lHabStatsMatrix.Add(pHabStatisticsObject)
                                        Else
                                            bClassFound = True
                                            lHabStatsMatrix(iStatsMatrixIndex).HabQuantity = lHabStatsMatrix(iStatsMatrixIndex).HabQuantity + dHabLength
                                        End If
                                        ' ============ END HABITAT STATS ==============
                                        ' =========================================

                                    End If

                                    ' ====================================================
                                    ' ============ LENGTH / DISTANCE STATS ===============
                                    ' 2020 exclusions don't apply to line length fields
                                    ' 2020 get distance / length field and quantity separetely from habitat (no classes for lengths)
                                    Try
                                        vTemp = pFeature.Value(iLengthField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from length attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 972.")
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iLengthField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The length found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. FIPEX Code 973." & ex.Message)
                                    End Try
                                    dTotalLength = dTotalLength + dTempQuan
                                    ' ========== END LENGTH / DISTANCE STATS =============
                                    ' ====================================================

                                    pFeature = pFeatureCursor.NextFeature

                                Loop     ' next selected feature
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
                                    .Direction = sDirection2
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "length"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "not set"
                                    .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                            End If ' There are items in the statsmatrix

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
                                .LengthOrHabitat = "length"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dTotalLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)


                        Else   ' if the habitat class case field is not found

                            dHabLength = 0

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

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then
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
                                            MsgBox("Could not convert the habitat quantity value found in the " + _
                                            lLineLayersFields(j).Layer.ToString + ". The values in the " + _
                                            lLineLayersFields(j).HabQuanField.ToString + " was not convertable to type 'double'." + _
                                            ex.Message)
                                            dTempQuan = 0
                                        End Try

                                        dHabLength = dHabLength + dTempQuan
                                    End If

                                    ' 2020 get distance / length field and quantity separetely from habitat (no classes for lengths)
                                    Try
                                        vTemp = pFeature.Value(iLengthField)
                                    Catch ex As Exception
                                        MsgBox("Problem converting value from length attribute in line layer to double / decimal. Please ensure values in attribute are of type double. FIPEX code 972.")
                                        vTemp = 0
                                    End Try
                                    Try
                                        dTempQuan = Convert.ToDouble(pFeature.Value(iLengthField))
                                    Catch ex As Exception
                                        dTempQuan = 0
                                        MsgBox("DCI table creation error. The length found in field in " & pFeatureLayer.Name & " was not convertible" _
                                            & " to type 'double'.  Null values in field may be responsible. FIPEX Code 973." & ex.Message)
                                    End Try
                                    dTotalLength = dTotalLength + dTempQuan

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
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dHabLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

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
                                .LengthOrHabitat = "length"
                                .HabitatDimension = "length"
                                .TotalImmedPath = sHabTypeKeyword
                                .UniqueClass = "not set"
                                .ClassName = CStr(lLineLayersFields(j).HabClsField)
                                .Quantity = dTotalLength
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

                        End If ' found habitat class field in line layer

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
                        ElseIf sUnit = "None" Then
                            sUnit = "none"
                        Else
                            sUnit = "n/a"
                        End If

                        iClassCheckTemp = pFields.FindField(lPolyLayersFields(j).HabClsField)
                        'If pFields.FindField(lLayersFields(j).ClsField) <> -1 Then
                        If iClassCheckTemp <> -1 And lPolyLayersFields(j).HabClsField <> "<None>" _
                            And lPolyLayersFields(j).HabClsField <> "Not set" And lPolyLayersFields(j).HabClsField <> "<none>" Then

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

                                pFeature = pFeatureCursor.NextFeature          ' For each selected feature
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    'pFields = pFeature.Fields  '** removed because should be redundant

                                    If bExclude = False Then

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

                                        classComparer = New FindStatsClassPredicate2020(sClass)
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
                                    .Direction = sDirection2
                                    .LengthOrHabitat = "habitat"
                                    .HabitatDimension = "area"
                                    .TotalImmedPath = sHabTypeKeyword
                                    .UniqueClass = "not set"
                                    .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                    .Quantity = 0.0
                                    .Unit = sUnit
                                End With
                                lHabStatsList.Add(pHabStatsObject_2)

                            End If ' There are items in the statsmatrix
                        Else   ' if the habitat class case field is not found

                            dHabArea = 0

                            If pFeatureSelection.SelectionSet.Count <> 0 Then

                                pFeatureSelection.SelectionSet.Search(Nothing, False, pCursor)
                                pFeatureCursor = CType(pCursor, IFeatureCursor)
                                pFeature = pFeatureCursor.NextFeature

                                ' Get the summary field and add the value to the total for habitat area.
                                ' ** ==> Multiple fields could be added here in a 'for' loop.
                                iFieldVal = pFeatureCursor.FindField(lPolyLayersFields(j).HabQuanField)

                                ' For each selected feature
                                m = 1
                                Do While Not pFeature Is Nothing

                                    ' 2020 - do not add if feature is to be excluded
                                    bExclude = False
                                    SharedSubs.exclusions2020(bExclude, pFeature, pFeatureLayer)

                                    If bExclude = False Then

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
                                            MsgBox("Could not convert the habitat quantity value found in the " + _
                                            lPolyLayersFields(j).Layer.ToString + "layer. The " + _
                                            +lPolyLayersFields(j).HabQuanField.ToString + " field was not convertable to type 'double'." + _
                                            ex.Message)

                                            dTempQuan = 0
                                        End Try
                                        ' Insert into the corresponding column of the second
                                        ' row the updated habitat area measurement.
                                        dHabArea = dHabArea + dTempQuan
                                    End If

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
                                .UniqueClass = "not set"
                                .ClassName = CStr(lPolyLayersFields(j).HabClsField)
                                .Quantity = dHabArea
                                .Unit = sUnit
                            End With
                            lHabStatsList.Add(pHabStatsObject_2)

                        End If ' found habitat class field in layer

                        ' increment the loop counter for
                        iLoopCount = iLoopCount + 1

                    End If  ' feature layer matches hab class layer
                Next    ' poly layer
            End If ' featurelayer is valid
            pFeatureLayer = CType(pEnumLayer.Next(), IFeatureLayer)
        Loop

    End Sub

    Public Class FindStatsClassPredicate2020
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

    Public Shared Function flagcheck2021(ByRef pBarriersList As IEnumNetEID, ByRef pEdgeFlags As IEnumNetEID, ByRef pJuncFlags As IEnumNetEID) As String
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
        '          Apr 28, 2021 -> Copied this function here from Analysis.vb, made public
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
            flagcheck2021 = "barriers"
            'return "barriers"? ' should be a return in VB.Net I think... but this works
        ElseIf pFlagsNoBarr.Count = pJuncFlags.Count Then
            flagcheck2021 = "nonbarr"
        Else
            MsgBox("Inconsistent flag placement." + vbCrLf + _
            "Barrier flags: " & pFlagsOnBarr.Count & vbCrLf & _
            " Non-barrier flags: " & pFlagsNoBarr.Count)
            flagcheck2021 = "error"
        End If
    End Function

End Class

