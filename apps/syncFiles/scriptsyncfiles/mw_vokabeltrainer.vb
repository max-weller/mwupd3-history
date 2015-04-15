#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

#Imports System.Text.RegularExpressions

#Para DebugMode External
#Para AssemblyType winexe

#WindowData frm_main
			550	500	Form		Width=550|Height=500|Text="Vokabeltrainer"
	10	230	500	200	IGridEx	igData	Dock=5
	10	440	200	30	Label	lab_Statusbar	Text="Hallo Welt"|Anchor=6

//DockStyle.Top=1							
	0	0	0	40	Panel	pnl_TopBar	Dock=1|BackColor=ColorTranslator.FromHtml("#dcdfef")
1	15	4	90	30	Button	btn_read	Text="Laden"
1	106	4	90	30	Button	btn_save	Text="Speichern"
1	270	10	147	20	TextBox	txt_Filename	Text=""
1	204	12	65	15	Label	Label_0	Text="Dateiname"
1	438	7	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9
#EndData
'Anchor=13/15

Dim MAIN As frm_main

Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MAIN.Close
End Sub

Sub btn_read_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  
  
  
End Sub

Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  
  
End Sub




Sub Onfrm_main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MAIN.igData.Cols.Add("x1", 20)
  MAIN.igData.Cols.Add("x2", 20)
  MAIN.igData.Cols.Add("Sprache1", 180)
  MAIN.igData.Cols.Add("Sprache2", 180)
  
  MAIN.igData.Rows.count = 40
End Sub

Sub Autostart()
  MAIN = New frm_main(Me)
  MAIN.Show()
  #If IDE = False Then
  Application.Run(MAIN)
  #End If
End Sub









'-->
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByVal para1 As String)
Public Event updateUserFeedback2(ByVal para1 As String)

Sub iGrid_TextBoxKeyDown(sender As Object, e As iGTextBoxKeyDownEventArgs) Handles myBase.TextBoxKeyDown
    'msgbox("hallo")
  If e.KeyValue  = Keys.Enter Then
    Me.CommitEditCurCell
    If Me.CurCell.ColIndex = 2 Then
      Me.setCurCell(Me.CurCell.RowIndex, 3)
    Else
      Me.setCurCell(Me.CurCell.RowIndex+1, 2)
    End If
  End If
End Sub

Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  'printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  'printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
End Sub


Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  printLine  4, "myIgrid_KeyDn", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
  
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDn", e.keycode 
  'e.handled =true
End Sub


Private Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

Private Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
  dim curRow as integer
  curRow=me.curRow.index
  trace "curRow", curRow
  RaiseEvent updateUserFeedback2(curRow.ToString)

End Sub

:Sub checkSelectionInRowMode(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then anz = anz +1
  Next
  RaiseEvent updateUserFeedback1(anz.ToString)

End Sub

:Sub resetAllSelections(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    if row.Selected=true then row.Selected=false
  Next
End Sub

End Class



'-->
'--> --> IgridHelper2

  
Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
  Dim max = line.Cells.Count - 1
  Dim out(max) As String
  :For i As Integer = 0 To max
    :out(i) = line.Cells(i).Value.ToString
  :Next
  :Return Join(out, Delimiter)
End Function

Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  Dim max = line.Cells.Count - 1
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  :For i As Integer = 0 To max
  :  line.Cells(i).Value = out(i)
  :Next
End Sub

Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim max = Grid.Rows.Count - 1
  Dim out(max) As String
  For i As Integer = 0 To max
    out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
  Next
  FileContent = Join(out, LineDelim)
End Sub
  
Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim out() = Split(FileContent, LineDelim)
  Grid.Rows.Clear()
  Grid.Rows.Count = out.Length
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub
  


