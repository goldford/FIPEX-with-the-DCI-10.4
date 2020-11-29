Imports System.Drawing
Imports ESRI.ArcGIS.Geodatabase

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
                    pResultsForm3.DataGridView1.Rows(iSinkRowIndex).Cells(9).Value = Math.Round(lMetricsObject(j).Metric, 2)
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
End Class

