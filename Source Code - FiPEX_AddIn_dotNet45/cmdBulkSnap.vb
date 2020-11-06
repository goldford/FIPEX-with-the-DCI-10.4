
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem


Public Class cmdBulkSnap
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Private m_editor As IEditor
    Private m_isSnappedToItSelf As Boolean
    Friend snapForm As frmBulkSnap

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()
        Try
            'Check if any SnapAgent is checked. 
            'The FeatureSnapAgent is turned off if the HitType is set to 
            'esriGeometryPartNone. Other SnapAgents are checked if they 
            'are registered in the snap environment that means that if 
            'the number of snapagents is greater then the number of 
            'featureSnapAgents a normal SnapAgent is checked.
            Dim snapEnvironment As ISnapEnvironment = CType(m_editor, ISnapEnvironment)
            Dim snapAgentCount As Integer
            Dim featSnapAgents As Integer = 0
            Dim checkedfSnapAgents As Integer = 0
            Dim snapAgents As Integer = snapEnvironment.SnapAgentCount
            'Loop through the feature snap agents in the snap environment.
            For snapAgentCount = 0 To snapEnvironment.SnapAgentCount - 1
                Dim currentSnapAgent As ISnapAgent = snapEnvironment.SnapAgent(snapAgentCount)
                If TypeOf currentSnapAgent Is IFeatureSnapAgent Then
                    Dim fSnapAgent As IFeatureSnapAgent = CType(currentSnapAgent, IFeatureSnapAgent)
                    featSnapAgents = featSnapAgents + 1
                    If Not fSnapAgent.HitType = esriGeometryHitPartType.esriGeometryPartNone Then
                        checkedfSnapAgents = checkedfSnapAgents + 1
                    End If
                End If
            Next snapAgentCount
            If checkedfSnapAgents = 0 And featSnapAgents = featSnapAgents Then
                System.Windows.Forms.MessageBox.Show("You need to turn on at least one snap agent." + vbCrLf + "In ArcGIS 10 you may need to turn on 'Classic Snapping' in the General Editing Options." + vbCrLf + "Followed by checking the lines you want to snap to in the 'Snapping Window' available in the dropdown of the Editor toolbar.")
                Exit Sub
            End If
            Dim pEditorProperties4 As IEditProperties4 = CType(m_editor, IEditProperties4)
            If pEditorProperties4.ClassicSnapping = False Then
                System.Windows.Forms.MessageBox.Show("You need to turn on at least one snap agent." + vbCrLf + "In ArcGIS 10 you may need to turn on 'Classic Snapping' in the General Editing Options." + vbCrLf + "Followed by checking the lines you want to snap to in the 'Snapping Window' available in the dropdown of the Editor toolbar.")
                Exit Sub
            End If

            snapForm = New frmBulkSnap
            'Retrieve settings back from the form.
            snapForm.ShowDialog()

            If Not snapForm.m_bCancel Then
                'Loop through the selected features and snap the ones you can
                Dim selectedFeatures As IEnumFeature = m_editor.EditSelection
                selectedFeatures.Reset()
                Dim numberOfSelectedFeatures As Integer
                numberOfSelectedFeatures = m_editor.SelectionCount

                'Start edit operation, enabling undo/redo.
                m_editor.StartOperation()
                Dim successfulSnappedFeatures As Integer = 0
                Dim selfSnappedFeatures As Integer = 0
                Dim extent As IEnvelope = New Envelope
                Dim featureCount As Integer
                For featureCount = 0 To numberOfSelectedFeatures - 1
                    Dim currentFeature As IFeature = selectedFeatures.Next()
                    'Reset self snapped Features. It may happen, that features is snapped to itself, if e.g. feature's FeaturesClass and the FeatureClass which is associated with the SnapAgent are the same;
                    m_isSnappedToItSelf = False
                    Dim clone As IClone = CType(currentFeature.Shape, IClone)
                    'Clone the geometry
                    Dim originalGeom As IGeometry = CType(clone.Clone(), IGeometry)
                    'snap Features and get the altered Geometry
                    Dim updatedGeometry As IGeometry = snapFeature(currentFeature, snapEnvironment)
                    If m_isSnappedToItSelf Then
                        selfSnappedFeatures = selfSnappedFeatures + 1
                    End If
                    If Not updatedGeometry Is Nothing Then
                        'update extent to zoom to snapped features
                        'update the feature`s geometry and store it
                        currentFeature.Shape = updatedGeometry
                        currentFeature.Store()
                        extent.Union(currentFeature.Extent)
                        successfulSnappedFeatures = successfulSnappedFeatures + 1
                    Else
                        'Geometry should not be updated, so the feature should maintain it's original geometry.
                        currentFeature.Shape = originalGeom
                        currentFeature.Store()
                    End If
                Next featureCount

                'Abort operation if no features were altered.
                If successfulSnappedFeatures = 0 Then
                    m_editor.AbortOperation()
                Else
                    m_editor.StopOperation("Snap features")
                    If extent.Height = 0 And extent.Width = 0 Then
                        Dim active As IActiveView = CType(m_editor.Map, IActiveView)
                        m_editor.Display.Invalidate(active.Extent, True, -2)
                    Else
                        extent.Expand(1.1, 1.1, True)
                        m_editor.Display.Invalidate(extent, True, -2)
                    End If
                End If

                'Display results in text box. 
                If selfSnappedFeatures > 0 Then
                    System.Windows.Forms.MessageBox.Show(successfulSnappedFeatures.ToString() & " Feature Geometries updated" & ControlChars.NewLine _
                    & selfSnappedFeatures.ToString() & " Feature(s) snapped to itself" & ControlChars.NewLine & _
                    "---------------------------------------------------------------" & ControlChars.NewLine & _
                    "Total: " & (successfulSnappedFeatures + selfSnappedFeatures).ToString() & " Feature(s) snapped.")
                Else
                    System.Windows.Forms.MessageBox.Show(successfulSnappedFeatures & " Feature(s) snapped.")
                End If
            Else
                snapForm.Close()
                Exit Sub
            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' This method stores all the features points in a collection 
    ''' and attempts to snap each point.
    ''' </summary>
    ''' <param name="featureToSnap">feature to snap</param>
    ''' <param name="snapEnvironment">Editors snap environment</param>
    ''' <returns>The resulting geometry from snapping the feature.</returns>
    Public Function snapFeature(ByVal featureToSnap As IFeature, ByVal snapEnvironment As ISnapEnvironment) As IGeometry
        Dim snapGeom As IGeometry = featureToSnap.Shape
        Dim pointCollection As IPointCollection
        Dim currentPoint As IPoint = New PointClass()
        Dim alteredPoint As IPoint = New PointClass()

        'Attempt to snap the feature.
        Select Case snapGeom.GeometryType
            Case esriGeometryType.esriGeometryPoint
                currentPoint = CType(snapGeom, ESRI.ArcGIS.Geometry.IPoint)
                alteredPoint = SnapPoint(currentPoint, snapEnvironment)
                'Use a cast to assign the geometry of the snapped feature to the altered geometry.
                Return CType(alteredPoint, IGeometry)

            Case esriGeometryType.esriGeometryPolyline
                pointCollection = CType(snapGeom, IPointCollection)
                ' Loop through the point collection, look at all points.
                ' If the user selected Line Snapping End Points Only then
                ' snap only the start and endpoints if the user selected 
                ' Otherwise snap all verticies.
                Dim i As Integer = 0
                Dim linePointSelfSnapped As Integer = 0
                For i = 0 To pointCollection.PointCount - 1
                    If i = 0 Or i = pointCollection.PointCount - 1 Then 'Or Not snapForm.m_isLineSnappingEndPointsOnly Then

                        currentPoint = pointCollection.Point(i)
                        alteredPoint = SnapPoint(currentPoint, snapEnvironment)
                        If Not alteredPoint Is Nothing Then
                            pointCollection.UpdatePoint(i, alteredPoint)
                        Else
                            'Check to see if the feature snapped to itself.
                            If m_isSnappedToItSelf Then
                                linePointSelfSnapped = linePointSelfSnapped + 1
                            Else
                                'Feature could not be snapped.
                                Return Nothing
                            End If
                        End If
                    End If
                Next
                'Determine how many points have to be snapped and compare that              'with the number of selfSnappedPoints.
                'Dim toStoreOrNot As Boolean

                'If snapForm.m_isLineSnappingEndPointsOnly Then
                If linePointSelfSnapped = 2 Then
                    m_isSnappedToItSelf = True
                    Return Nothing
                Else
                    m_isSnappedToItSelf = False
                    'Check the point collection geometry then return the geometry.
                    snapFeature = checkGeometry(CType(pointCollection, IGeometry))
                    Return snapFeature
                End If
                'Else
                '    'User selected to snap ends and vertices, so check all vertices.
                '    If linePointSelfSnapped = pointCollection.PointCount Then
                '        m_isSnappedToItSelf = True
                '        Return Nothing
                '    Else
                '        m_isSnappedToItSelf = False
                '        snapFeature = checkGeometry(CType(pointCollection, IGeometry))
                '        Return snapFeature
                '    End If
                'End If

                Return Nothing

            Case esriGeometryType.esriGeometryPolygon
                pointCollection = CType(snapGeom, IPointCollection)
                Dim polygonPointsSnappedToItself As Integer = 0

                'If snapForm.m_isPolygonSnapping Then
                Dim i As Integer = 0
                For i = 0 To pointCollection.PointCount - 1
                    alteredPoint = SnapPoint(pointCollection.Point(i), snapEnvironment)

                    If Not alteredPoint Is Nothing Then
                        pointCollection.UpdatePoint(i, alteredPoint)
                    Else
                        If m_isSnappedToItSelf Then
                            polygonPointsSnappedToItself = polygonPointsSnappedToItself + 1
                        Else
                            Return Nothing
                        End If
                    End If
                Next
                'End If

                'Determine how many points have to be snapped and compare 
                'that with the number of selfSnappedPoints all points have to be            'snapped.
                If polygonPointsSnappedToItself = pointCollection.PointCount Then
                    m_isSnappedToItSelf = True
                    Return Nothing
                Else
                    m_isSnappedToItSelf = False
                    snapFeature = checkGeometry(CType(pointCollection, IGeometry))
                    Return snapFeature
                End If
                Return Nothing
            Case Else
                m_editor.AbortOperation()
                Return Nothing
        End Select
    End Function
    ''' <summary>
    ''' This method checks the geometry to make sure that it is not empty and is correct. the
    ''' </summary>
    ''' <param name="snapGeometry">Geometry from the point collection that was snapped.</param>
    ''' <returns>The simplified geometry</returns>
    Private Function checkGeometry(ByVal snapGeometry As ESRI.ArcGIS.Geometry.IGeometry) As IGeometry
        Dim topoOp2 As ESRI.ArcGIS.Geometry.ITopologicalOperator2
        topoOp2 = CType(snapGeometry, ITopologicalOperator2)
        'Make sure that the geometry is not empty, then simply it again.
        topoOp2.IsKnownSimple_2 = False
        topoOp2.Simplify()
        If snapGeometry.IsEmpty = True Then
            Return Nothing
        Else
            Return snapGeometry
        End If
    End Function
    ''' <summary>
    ''' This method snaps a point by passing this point to 
    ''' ISnapEnvironment:SnapPoint where every selected SnapAgent
    ''' is asked to perform the snapjob.
    ''' </summary>
    ''' <param name="currentPoint">Point to snap.</param>
    ''' <param name="snapEnvironment">Which snap agents are turned on.</param>
    ''' <returns>The snapped point.</returns>
    Public Function SnapPoint(ByVal currentPoint As ESRI.ArcGIS.Geometry.IPoint, ByVal snapEnvironment As ISnapEnvironment) As IPoint

        Dim x As Double = currentPoint.X
        Dim y As Double = currentPoint.Y
        Dim successfullySnapped As Boolean = snapEnvironment.SnapPoint(currentPoint)
        Dim deltaX As Double = x - currentPoint.X
        Dim deltaY As Double = y - currentPoint.Y

        If Not successfullySnapped Then
            Return Nothing
        Else
            'If successful check to see if the location of the point changed
            If successfullySnapped AndAlso deltaX = 0.0 AndAlso deltaY = 0.0 Then
                'If the location has not changed flag it as snapped to itself.
                m_isSnappedToItSelf = True
                Return Nothing
            Else
                Return currentPoint
            End If
        End If
        Return Nothing
    End Function

    Protected Overrides Sub OnUpdate()
        ' get the application 
        ' get the editor
        ' check if there's an edit session
        ' check if there's a selection

        If My.ArcMap.Application IsNot Nothing Then
            Dim uID As New UID
            uID.Value = "esriEditor.Editor"

            If m_editor Is Nothing Then
                m_editor = CType(My.ArcMap.Application.FindExtensionByCLSID(uID), IEditor)
            End If

            If m_editor Is Nothing Then
                Me.Enabled = False
                Exit Sub
            End If

            If m_editor.SelectionCount > 0 And m_editor.EditState = esriEditState.esriStateEditing Then
                Me.Enabled = True
            Else
                Me.Enabled = False
            End If

        End If

    End Sub
End Class
