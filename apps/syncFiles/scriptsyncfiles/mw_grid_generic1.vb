
#Para DebugMode Internal

Sub AutoStart()
  Dim gridView = ZZ.IDEHelper.Addins("siaGridView")
  Dim form = gridView.CreateGenericGrid("mw_grid_generic1.testGrid")
  Dim ig As IGrid = form.Grid
  
  form.setDefaultColumnHeaders
  
  ig.Rows.Count = 20
  For i as integer = 0 To 19
    ig.Cells(i,i).Value = "####"
  Next
End Sub





