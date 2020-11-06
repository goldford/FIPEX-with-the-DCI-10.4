Public Class Combo1
  Inherits ESRI.ArcGIS.Desktop.AddIns.ComboBox

  Public Sub New()

  End Sub

  Protected Overrides Sub OnUpdate()
    Enabled = My.ArcMap.Application IsNot Nothing
  End Sub
End Class
