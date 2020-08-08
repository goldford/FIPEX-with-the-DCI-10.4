Public Class SharedObjectClasses

End Class
' Defines the object that can be added to Exclude List
Public Class StatisticsObject
    Public UniqueClass As String
    Public Quantity As Double

    Public Sub New(ByVal uniqueclass As String, ByVal quantity As Double)
        Me.UniqueClass = uniqueclass
        Me.Quantity = quantity
    End Sub

End Class
Public Class StatisticsObject_2
    Public Layer As String
    Public LayerID As Integer
    Public Sink As String
    Public SinkEID As Integer
    Public bID As String
    Public bEID As Integer
    Public bType As String
    Public Direction As String
    Public TotalImmedPath As String
    Public UniqueClass As String
    Public ClassName As String
    Public Quantity As Double
    Public Unit As String

    Public Sub New(ByVal layer As String, ByVal layerID As Integer, ByVal sink As String, ByVal sinkEID As Integer, ByVal bid As String, ByVal barrEID As Integer, ByVal btype As String, ByVal direction As String, _
    ByVal totalimmedpath As String, ByVal uniqueclass As String, ByVal classname As String, ByVal quantity As Double, ByVal unit As String)
        Me.Layer = layer
        Me.LayerID = layerID
        Me.Sink = sink
        Me.SinkEID = sinkEID
        Me.bID = bid
        Me.bEID = barrEID
        Me.bType = btype
        Me.Direction = direction
        Me.TotalImmedPath = totalimmedpath
        Me.UniqueClass = uniqueclass
        Me.ClassName = classname
        Me.Quantity = quantity
        Me.Unit = unit
    End Sub

End Class
Public Class DCIStatisticsObject

    '   ObjectID (autoinc)
    '   BarrierID (string 55)
    '   Quantity (double)
    '   BarrierPerm (double)
    '   BarrierYN (string 55)
    Public Barrier As String
    Public Quantity As Double
    Public BarrierPerm As Double
    Public BarrierYN As String


    Public Sub New(ByVal barrier As String, ByVal quantity As Double, ByVal barrierperm As Double, _
    ByVal barrieryn As String)
        Me.Barrier = barrier
        Me.Quantity = quantity
        Me.BarrierPerm = barrierperm
        Me.BarrierYN = barrieryn
    End Sub

End Class
Public Class GLPKStatisticsObject

    '   BarrierID (Integer)
    '   Quantity (double)
    Public Barrier As Integer
    Public Quantity As Double


    Public Sub New(ByVal barrier As Integer, ByVal quantity As Double)
        Me.Barrier = barrier
        Me.Quantity = quantity
    End Sub
End Class
Public Class GLPKOptionsObject

    '   BarrierID (Integer)
    '   OPtion# (integer)
    '   BarrierPerm (double)
    '   OptionCost(Double)
    Public BarrierEID As Integer
    Public OptionNum As Integer
    Public BarrierPerm As Double
    Public OptionCost As Double
    Public BarrierType As String

    Public Sub New(ByVal barrier As Integer, _
                   ByVal optionnum As Integer, _
                   ByVal barrierperm As Double, _
                   ByVal optioncost As Double, _
                   ByVal _barriertype As String)
        Me.BarrierEID = barrier
        Me.OptionNum = optionnum
        Me.BarrierPerm = barrierperm
        Me.OptionCost = optioncost
        Me.BarrierType = _barriertype
    End Sub
End Class

Public Class MetricsObject
    ' ------------------------
    '   For populating metric table
    '   ObjectID (autoinc)
    '   User-set Sink ID
    '   The Sink Element ID (integer)
    '   BarrierID (string 55)
    '   The barrier element ID (Integer)
    '   metricname (string 55)
    '   metric (double)
    ' 
    '   This object is used to populate the 
    '   the DBF table and the output results form
    ' ------------------------

    Public Sink As String
    Public SinkEID As Integer
    Public ID As String
    Public BarrEID As Integer
    Public Type As String
    Public MetricName As String
    Public Metric As Double


    Public Sub New(ByVal sink As String, ByVal iSinkEID As Integer, ByVal sID As String, ByVal iBarrEID As Integer, ByVal sType As String, ByVal metricname As String, ByVal metric As Double)
        Me.Sink = sink
        Me.SinkEID = iSinkEID
        Me.ID = sID
        Me.BarrEID = iBarrEID
        Me.Type = sType
        Me.MetricName = metricname
        Me.Metric = metric
    End Sub

End Class
Public Class FCIDandNameObject
    ' this object holds the FCID and name
    ' of layers being used in DCI habitat
    ' table output.  It's used to eliminate
    ' layers from inclusion that are not 
    ' part of the network. 
    Public FCID As Integer
    Public Name As String


    Public Sub New(ByVal fcid As Integer, ByVal name As String)
        Me.FCID = fcid
        Me.Name = name
    End Sub

End Class
Public Class BarrierAndDownstreamID
    ' this object is used to 
    Public BarrID As String
    Public DownstreamBarrierID As String


    Public Sub New(ByVal barrid As String, ByVal downstreambarrierid As String)
        Me.BarrID = barrid
        Me.DownstreamBarrierID = downstreambarrierid
    End Sub

End Class
Public Class GLPKBarrierAndDownstreamEID
    ' this object is used to 
    Public BarrEID As Integer
    Public DownstreamBarrierEID As Integer


    Public Sub New(ByVal barreid As Integer, ByVal downstreambarriereid As Integer)
        Me.BarrEID = barreid
        Me.DownstreamBarrierEID = downstreambarriereid
    End Sub

End Class
Public Class IDandType
    ' this object is used to get the label and type of barrier
    ' from the feature class
    Public BarrID As String
    Public BarrIDType As String


    Public Sub New(ByVal barrid As String, ByVal barridtype As String)
        Me.BarrID = barrid
        Me.BarrIDType = barridtype
    End Sub

End Class
Public Class SinkandDCIs
    ' this object is used to]
    Public SinkEID As Integer
    Public Type As String
    Public DCIp As Double
    Public DCId As Double


    Public Sub New(ByVal sinkeid As Integer, ByVal Type As String, ByVal dcip As Double, ByVal dcid As Double)
        Me.SinkEID = sinkeid
        Me.Type = Type
        Me.DCIp = dcip
        Me.DCId = dcid
    End Sub

End Class
Public Class SinkandTypes
    ' this object is used to 
    Public SinkEID As Integer
    Public SinkID As String
    Public Type As String


    Public Sub New(ByVal sinkeid As Integer, ByVal sinkid As String, ByVal type As String)
        Me.SinkEID = sinkeid
        Me.SinkID = sinkid
        Me.Type = type
    End Sub

End Class
Public Class BarrierAndSinkEIDs
    ' this object is used to 
    Public SinkEID As Integer
    Public BarrEID As Integer


    Public Sub New(ByVal sinkeid As Integer, ByVal barreid As Integer)
        Me.SinkEID = sinkeid
        Me.BarrEID = barreid
    End Sub

End Class
Public Class BarrAndBarrEIDAndSinkEIDs
    ' this object is used to 
    Public SinkEID As Integer
    Public BarrEID As Integer
    Public BarrLabel As String


    Public Sub New(ByVal sinkeid As Integer, ByVal barreid As Integer, ByVal barrlabel As String)
        Me.SinkEID = sinkeid
        Me.BarrEID = barreid
        Me.BarrLabel = barrlabel
    End Sub
End Class

Public Class LayersAndTypes
    ' this object is used to keep track 
    ' of unique layers and type (polygon, line)
    ' encountered during traces.  
    Public LayerName As String
    Public Type As String


    Public Sub New(ByVal layername As String, ByVal type As String)
        Me.LayerName = layername
        Me.Type = type
    End Sub
End Class
Public Class DIR_OptResultsObject
    ' this object is used to 
    Public SinkEID As Integer
    Public Budget As Double
    Public GLPK_Solved As Boolean
    Public Perc_Gap As Double
    Public MaxSolTime As Integer
    Public TimeUsed As Double
    Public Habitat_ZMax As Double
    Public Treatment_Name As String


    Public Sub New(ByVal sinkeid As Integer, ByVal treatment_name As String, ByVal budget As Double, ByVal glpk_solved As Boolean, ByVal perc_gap As Double, ByVal maxsoltime As Integer, _
                   ByVal timeused As Double, ByVal habitat_zmax As Double)
        Me.SinkEID = sinkeid
        Me.Treatment_Name = treatment_name
        Me.Budget = budget
        Me.GLPK_Solved = glpk_solved
        Me.Perc_Gap = perc_gap
        Me.MaxSolTime = maxsoltime
        Me.TimeUsed = timeused
        Me.Habitat_ZMax = habitat_zmax
    End Sub
End Class
Public Class UNDIR_OptResultsObject
    ' this object is used to 
    Public SinkEID As Integer
    Public Budget As Double
    Public GLPK_Solved As Boolean
    Public Perc_Gap As Double
    Public MaxSolTime As Integer
    Public TimeUsed As Double
    Public Habitat_ZMax As Double
    Public Treatment_Name As String
    Public CentralBarrierEID As Integer


    Public Sub New(ByVal sinkeid As Integer, ByVal treatment_name As String, ByVal budget As Double, ByVal glpk_solved As Boolean, ByVal perc_gap As Double, ByVal maxsoltime As Integer, _
                   ByVal timeused As Double, ByVal habitat_zmax As Double, ByVal centralbarriereid As Integer)
        Me.SinkEID = sinkeid
        Me.Treatment_Name = treatment_name
        Me.Budget = budget
        Me.GLPK_Solved = glpk_solved
        Me.Perc_Gap = perc_gap
        Me.MaxSolTime = maxsoltime
        Me.TimeUsed = timeused
        Me.Habitat_ZMax = habitat_zmax
        Me.CentralBarrierEID = centralbarriereid
    End Sub
End Class
Public Class GLPKDecisionsObject
    ' this object is used to 
    Public Budget As Double
    Public Treatment As String
    Public BarrierEID As Integer
    'Public BarrierLabel As String
    'Public BarrierOBJID As Integer
    'Public BarrierLayer As String
    'Public BarrierLayerID As Integer
    Public DecisionOption As Integer


    Public Sub New(ByVal budget As Double, ByVal treatment As String, ByVal barriereid As Integer, ByVal decisionoption As Integer)
        Me.Budget = budget
        Me.Treatment = treatment
        Me.BarrierEID = barriereid
        'Me.BarrierLabel = barrierlabel
        'Me.BarrierOBJID = barrierobjid
        'Me.BarrierLayer = barrierlayer
        'Me.BarrierLayerID = barrierlayerid
        Me.DecisionOption = decisionoption
    End Sub
End Class
Public Class LayersAndFCIDs
    ' this object is used to keep track 
    ' of map layer names and unique FCIDs
    ' and field that will be updated with results
    Public LayerName As String
    Public FCID As Integer



    Public Sub New(ByVal layername As String, ByVal fcid As Integer)
        Me.LayerName = layername
        Me.FCID = fcid
    End Sub
End Class
Public Class LayersAndFCIDsAndFields
    Public Layer As String
    Public QuanField As String
    Public ClsField As String
    Public UnitField As String

    Public Sub New(ByVal layer As String, ByVal QuanField As String, ByVal ClsField As String, ByVal UnitField As String)
        Me.Layer = layer
        Me.QuanField = QuanField
        Me.ClsField = ClsField
        Me.UnitField = UnitField

    End Sub

End Class

' Defines the object that can be added to a habParamList
Public Class LayerToAdd
    Public Layer As String
    Public QuanField As String
    Public ClsField As String
    Public UnitField As String

    Public Sub New(ByVal layer As String, ByVal QuanField As String, ByVal ClsField As String, ByVal UnitField As String)
        Me.Layer = layer
        Me.QuanField = QuanField
        Me.ClsField = ClsField
        Me.UnitField = UnitField

    End Sub

End Class
' Defines the object that can be added to Exclude List
Public Class LayerToExclude
    Public Layer As String
    Public Feature As String
    Public Value As String

    Public Sub New(ByVal layer As String, ByVal feature As String, ByVal value As String)
        Me.Layer = layer
        Me.Feature = feature
        Me.Value = value
    End Sub

End Class
' Defines the object used to add barrier IDs
Public Class BarrierIDObj
    Public Layer As String
    Public Field As String
    Public PermField As String
    Public NaturalYNField As String
    Public LayerType As String ' holds culvert / barrier type -- added for thesis SA
    Public Sub New(ByVal layer As String, ByVal field As String, ByVal permfield As String, ByVal naturalynfield As String, ByVal _layertype As String)
        Me.Layer = layer
        Me.Field = field
        Me.PermField = permfield
        Me.NaturalYNField = naturalynfield
        Me.LayerType = _layertype
    End Sub
End Class
Public Class LayerFCIDAndQuanField
    Public _Layer As String
    Public _FCID As Integer
    Public _Quanfield As String
    Public Sub New(ByVal layer As String, ByVal fcid As Integer, ByVal quanfield As String)
        Me._Layer = layer
        Me._FCID = fcid
        Me._Quanfield = quanfield
    End Sub
End Class
Public Class LayerFCIDAndHabQuan
    Public _Layer As String
    Public _FCID As Integer
    Public _Quan As Double
    Public Sub New(ByVal layer As String, ByVal fcid As Integer, ByVal quan As Double)
        Me._Layer = layer
        Me._FCID = fcid
        Me._Quan = quan
    End Sub
End Class
Public Class ResultsIDsObject
    ' this object is used to 
    Public _EID As Integer
    Public _FCID As Integer
    Public _FID As Integer
    Public _SubID As Integer


    Public Sub New(ByVal eid As Integer, ByVal fcid As Integer, ByVal fid As Integer, ByVal subid As Integer)
        Me._EID = eid
        Me._FCID = fcid
        Me._FID = fid
        Me._SubID = subid
    End Sub
End Class

Public Class FCID_FID_dTotalObject
    ' this object is used to 
    Public _FCID As Integer
    Public _FID As Integer
    Public _dTotal As Double

    Public Sub New(ByVal fcid As Integer, ByVal fid As Integer, ByVal dTotal As Double)
        Me._FCID = fcid
        Me._FID = fid
        Me._dTotal = dTotal
    End Sub
End Class
Public Class DecisionsOverlapObject
    ' this object is used in frmCalculateOverlap
    Public _BarrierEID As Integer
    Public _Budget As Double
    Public _OverlapCount As Integer
    Public _OverlapStatistic As Integer
    Public _Treatment1 As String
    Public _T1DecisionCount As Integer
    Public _Treatment2 As String
    Public _T2DecisionCount As Integer

    Public Sub New(ByVal barrierEID As Integer, _
                   ByVal budget As Double, _
                   ByVal overlapcount As Integer, _
                   ByVal overlapstatistic As Integer, _
                   ByVal treatment1 As String, _
                   ByVal t1decisioncount As Integer, _
                   ByVal treatment2 As String, _
                   ByVal t2decisioncount As Integer)
        Me._BarrierEID = barrierEID
        Me._Budget = budget
        Me._OverlapCount = overlapcount
        Me._OverlapStatistic = overlapstatistic
        Me._Treatment1 = treatment1
        Me._T1DecisionCount = t1decisioncount
        Me._Treatment2 = treatment2
        Me._T2DecisionCount = t2decisioncount
    End Sub
End Class
Public Class DecisionsOverlapObjectFinal
    ' this object is used in frmCalculateOverlap
    ' it has a type double for 'overlap' 
    Public _BarrierEID As Integer
    Public _Budget As Double
    Public _Overlap As Double
    Public _Treatment1 As String
    Public _Treatment2 As String

    Public Sub New(ByVal barrierEID As Integer, ByVal budget As Double, ByVal overlap As Double, ByVal treatment1 As String, ByVal treatment2 As String)
        Me._BarrierEID = barrierEID
        Me._Budget = budget
        Me._Overlap = overlap
        Me._Treatment1 = treatment1
        Me._Treatment2 = treatment2
    End Sub
End Class
Public Class LayersAndFCIDAndCumulativePassField
    Public Layer As String
    Public FCID As Integer
    Public CumPermField As String

    Public Sub New(ByVal layer As String, ByVal fcid As Integer, ByVal cumpermfield As String)
        Me.Layer = layer
        Me.FCID = fcid
        Me.CumPermField = cumpermfield
    End Sub

End Class
Public Class EIDCPermAndDir
    ' for use with visualize decisions and net tool
    '   BarrierID (Integer)
    '   cPerm      (Double)
    '   Dir       (String)
    Public Barrier As Integer
    Public cPerm As Double
    Public Dir As String

    Public Sub New(ByVal barrier As Integer, ByVal cperm As Double, ByVal dir As String)
        Me.Barrier = barrier
        Me.cPerm = cperm
        Me.Dir = dir
    End Sub
End Class

Public Class SelectAndUpdateFeaturesObject
    'Custom Object: pworkspace, pfeaturelayer, iCPermFieldIndex, pSelectionSet, dCPerm
    ' For use in Visualize Decisions and Network algorithm
    Public pWorkspace As ESRI.ArcGIS.Geodatabase.IWorkspace
    Public pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Public iCPermFieldIndex As Integer
    Public pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet
    Public dCPerm As Double

    Public Sub New(ByVal pworkspace As ESRI.ArcGIS.Geodatabase.IWorkspace, _
                   ByVal pfeaturelayer As ESRI.ArcGIS.Carto.IFeatureLayer, _
                   ByVal icpermfieldindex As Integer, _
                   ByVal pselectionset As ESRI.ArcGIS.Geodatabase.ISelectionSet, _
                   ByVal dcperm As Double)
        Me.pWorkspace = pworkspace
        Me.pFeatureLayer = pfeaturelayer
        Me.iCPermFieldIndex = icpermfieldindex
        Me.pSelectionSet = pselectionset
        Me.dCPerm = dcperm
    End Sub

End Class
Public Class FeatureLayerAndSelectionSet
    ' Custom Object: pfeaturelayer, pSelectionSet
    ' For use in Visualize Decisions and Network algorithm
    Public pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Public pSelectionSet As ESRI.ArcGIS.Geodatabase.ISelectionSet

    Public Sub New(ByVal pfeaturelayer As ESRI.ArcGIS.Carto.IFeatureLayer, _
                ByVal pselectionset As ESRI.ArcGIS.Geodatabase.ISelectionSet)
        Me.pFeatureLayer = pfeaturelayer
        Me.pSelectionSet = pselectionset
    End Sub

End Class
Public Class FeatureLayerAndCPermField
    ' Custom Object: pfeaturelayer, pSelectionSet
    ' For use in Visualize Decisions and Network algorithm
    Public pFeatureLayer As ESRI.ArcGIS.Carto.IFeatureLayer
    Public sCPermField As String

    Public Sub New(ByVal pfeaturelayer As ESRI.ArcGIS.Carto.IFeatureLayer, _
                ByVal scpermfield As String)
        Me.pFeatureLayer = pfeaturelayer
        Me.sCPermField = scpermfield
    End Sub

End Class
Public Class DecisionCountObject
    ' Custom Object: for counting decisions at each budget amount
    '                of a given 'results' table
    ' For use in Visualize Decisions and Network algorithm
    ' Added as a separate subroutine to help in thesis. 
    ' The 'decision count' should be added to 'results' tables
    ' from GLPK but at this point had not been. 
    ' 
    ' Solver = string (GLPK or GRB)
    ' UniqueCode = string
    ' Direction = string
    ' Treatment = string
    ' budget = double
    ' percgap = double
    ' decisioncount = integer
    ' timeused = double

    Public sSolver As String
    Public sUniqueCode As String
    Public sDirection As String
    Public sTreatment As String
    Public dBudget As Double
    Public dPercGap As Double
    Public iDecisionCount As Integer
    Public dTimeUsed As Double

    Public Sub New(ByVal _sSolver As String, _
                   ByVal _sUniqueCode As String, _
                   ByVal _sDirection As String, _
                   ByVal _sTreatment As String, _
                   ByVal _dBudget As Double, _
                   ByVal _dPercGap As Double, _
                   ByVal _iDecisionCount As Integer, _
                   ByVal _dTimeUsed As Double)
        Me.sSolver = _sSolver
        Me.sUniqueCode = _sUniqueCode
        Me.sDirection = _sDirection
        Me.sTreatment = _sTreatment
        Me.dBudget = _dBudget
        Me.dPercGap = _dPercGap
        Me.iDecisionCount = _iDecisionCount
        Me.dTimeUsed = _dTimeUsed
    End Sub

End Class