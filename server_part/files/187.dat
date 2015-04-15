


#Para DebugMode intern
#Para SilentMode true

#reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic
#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib

public WithEvents FormRef As Form
public WithEvents ref As Object
shared winId as string ="bookmark.vb"
public myWin2 as object
public winCaption as string = "Snippet-Manager"



public WithEvents IGrid1 As Igrid

'''#####################################################
'''#####################################################

''http://snap.teamwiki.net/100211-203909-es-snippet-manager02.png




'-->
sub test1()
      Dim myDate As Date
      myDate = Now()
  msgBox(  myDate.ToString("yyyy-MM-dd__HH:mm"))

  ''myDate.ToString("yyyy-MM-dd__HH:mm", ci))
  ''msgBox("OK - 1")
end sub

sub test2()
  msgBox("OK - 2")
end sub

sub test3()
  msgBox("OK - 3")
end sub

'-->



'''#####################################################
'''#####################################################


'''#####################################################
'''#####################################################
Sub GetFormRef()

  'If ref IsNot Nothing Then Exit Sub
  ref = ZZ.IDEHelper.CreateDockWindow(winID, 4)
  formRef = ref.form
  formRef.text=winCaption
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
  
  ''DAS MACHT EINE NORMALE FORM
  ''formRef = New frmTB_scriptWin 'ZZ.IDEHelper.CreateForm(winID)
  ''ref = CType(FormRef, Object).PNL
  ''FormRef.Text = "Snippet-Manager"
End Sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls
End Sub

Sub onResizeControls()
  Igrid1.Width = formRef.Width - 10
  Igrid1.Height = formRef.Height - 160
End Sub


'-->
'''#####################################################
Sub AutoStart()
  myWin2 = ZZ.scriptClass("es_snippetEditor")
  myWin2.setParent(me)
  zz.traceClear()
  zz.printLineReset()
  
  GetFormRef()
  With ref
     dim el as object
    .resetControls ()
    .insertX = 11 :.insertY = 12
    .addIcon("appPicture", "http://es.teamwiki.net/docs/icons/insert-object64.png" )
     
.BR  '--------------------------------------------------
     .insertX = 75 ::.insertY = 12
     .addButton ("insertItem"     , "Einfügen"  , "#1DD910")
.BR  (24)'--------------------------------------------------
     .insertX = 75 :
    .addButton ("showHideZoom"    , "Bearbeiten"      , "#ccc")

 .BR (24) '--------------------------------------------------
    .insertX = 75 
    .addButton ("addLine"        , "N E U"  , "#DF7500")

.BR  '--------------------------------------------------
   .insertX = -3 :.insertY = 90
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("outMonitor" ,  -2         , "Roh-"+vbNewLine+"Daten"    , 0  , "#afa", , , "multiline=50")
    ref.element("txt_outMonitor").WordWrap=false

.BR (55) '--------------------------------------------------
    .insertX = 5 
    el = .addButton ("read"            , "Lesen"      , "#ccc")
    el.width= 70
    .insertX = 80
    el = .addButton ("save"            , "Speichern"  , "#ccc")
    el.width=70


End With
  IGrid1 = New IGrid
  IGrid1.Top = 180
  'leere Zelle als "" statt nothing
  
  iGrid1.DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
  iGrid1.DefaultCol.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
  iGrid1.DefaultCol.ColHdrStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None
  
  IGrid1.Cols.Add("id").width=5
  IGrid1.Cols.Add("gruppe").width=18
  IGrid1.Cols.Add("nickname").width=150
  IGrid1.Cols.Add("caption").width=3
  IGrid1.Cols.Add("body").width=3
  IGrid1.Cols.Add("tags").width=3
  IGrid1.GroupBox.Visible = True
  Igrid1.rows.count = 100
  iGrid1.ReadOnly = false
  iGrid1.AllowDrop = True
  iGrid1.HotTracking = False
  iGrid1.SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
  iGrid1.SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText
  iGrid1.rowMode=true
  
  ref.Controls.Add(IGrid1)
  onResizeControls()
  'FormRef.Show()
  FormRef.Visible=true
  

  readSnippetFile()
  iGrid1.rows(0).selected=true
  
  '...ersten Eintrag markieren fehlt noch
  
  'ZZ.scriptClass("es_snippetEditor").createEditWindow(0)
  
End Sub




'-->
'--> EVENTS

Sub insertItem_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
End Sub

Sub appPicture_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
End Sub

Sub showHideZoom_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  'ZZ.scriptClass("es_snippetEditor").hideSnippetEditor()
  zoomInSnippetEditor()
End Sub

Sub addLine_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  zoomInSnippetEditor()
End Sub

sub zoomInSnippetEditor()
  dim ig as IGrid = IGrid1
  dim curRow as iGRow
  curRow=ig.curRow
  dim curRowIndex as integer = curRow.index
  ZZ.scriptClass("es_snippetEditor").createEditWindow(curRowIndex)
end sub

Sub read_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  readSnippetFile()
End Sub

Sub save_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  saveSnippetFile() 
End Sub




'-->

sub readSnippetFile()
  trace "OK","OK"
  dim FileName as string =""
  Dim Content as string =""
  Content = readWebFile(FileName)
  '' Content = ZZ.fileGetContents("C:\yPara\scriptIDE\codeBookmarks.txt")
  '' msgBox(content)

  Igrid_put (IGrid1, Content)
End Sub

sub saveSnippetFile()
  Dim Content As String
  Igrid_get (IGrid1, Content)

  Dim myDate As Date = Now()
  dim path as string="C:\yPara\scriptIDE\codeBookmarks\"
  dim fileName=myDate.ToString("yyyy-MM-dd__HH_mm__")+"codeBookmarks.txt"
  printLine 2, filename, filename
  ZZ.filePutContents(path+fileName, Content)

  saveWebFile("", Content)
end sub


sub refreshZoom(curRow as iGRow)
  'dim rowIndex as integer = curRow.index
  'trace "refreshZoom", rowIndex
  'trace "refreshZoom", curRow.cells(2).Value.ToString
  dim code as string = curRow.cells(2).Value.ToString
  printLine 2, "code", code
end sub

function getDragDropData(curRow as iGRow) as string
  trace "getDragDropData", curRow.cells(2).Value.ToString
  dim code as string = curRow.cells(4).Value.ToString
  printLine 2, "code", code
  code=replace(code,"||LF||",vbNewLine)
  return code + vbNewLine
end function

sub onZoomDataChanged()

end sub



sub switchToCodeEditor()
  dim tab as object=zz.getActiveTab()
  'Dim scin As ScintillaNet.Scintilla = tab.RTF
  Dim scin As object = tab.RTF
  
  scin.focus()
  'Dim codeText As String = scin.Text

end sub


'-->

Sub iGrid1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles iGrid1.MouseDown
  'umstellen auf startDragCell
  'printLine 9, "iGrid1_MouseDown", 'OK'
  Dim mousePos As Integer = e.Y
  Dim groupBoxHeight As Integer = iGrid1.GroupBox.Height
  If mousePos < (groupBoxHeight + 15) Then Exit Sub

  dim ig as IGrid = sender
  dim curRow as iGRow
  curRow=sender.curRow
  
  refreshZoom(curRow)
  dim dragDropData as string =getDragDropData(curRow)
  if e.Button = Windows.Forms.MouseButtons.Right then
    iGrid1.ReadOnly = false 
  else
    'iGrid1.ReadOnly = true
    Dim dat As New DataObject
    dat.SetText(dragDropData)
    iGrid1.DoDragDrop(dat, DragDropEffects.All)
    '... umstellen auf mouseMove
    switchToCodeEditor()
  end if
End Sub


Sub iGrid1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles iGrid1.MouseUp
  '--> !!! ... der kommt hier gar nicht vorbei
  printLine 9, "iGrid1_MouseDown", "OK"
  iGrid1.ReadOnly = false
End Sub

Private Sub iGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles iGrid1.DoubleClick
  iGrid1.ReadOnly = false
End Sub

Private Sub iGrid1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles iGrid1.DragDrop
  printLine 9, "igGlobsearch_DragDrop", "OK"
End Sub

Private Sub iGrid1_CellMouseUp(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseUpEventArgs) Handles iGrid1.CellMouseUp
  printLine 9, "iGrid1_CellMouseUp", "OK"
End Sub


'-->
'--> --> IgridHelper2

  
:Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
  on error resume next
  Dim max = line.Cells.Count - 1
  Dim out(max) As String
  For i As Integer = 0 To max
    out(i) = line.Cells(i).Value.ToString
  Next
  Return Join(out, Delimiter)
End Function

:Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  Dim max = line.Cells.Count - 1
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  For i As Integer = 0 To max
    line.Cells(i).Value = out(i)
  Next
End Sub

:Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim max = Grid.Rows.Count - 1
  Dim out(max) As String
  For i As Integer = 0 To max
    out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
  Next
  FileContent = Join(out, LineDelim)
End Sub
  
:Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim out() = Split(FileContent, LineDelim)
  Grid.Rows.Clear()
  Grid.Rows.Count = out.Length
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub
  
Sub iGrid1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles iGrid1.DragEnter
    'Debug.Print("iGrid1_DragEnter")
    e.Effect = DragDropEffects.All
End Sub


Sub iGrid1_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles iGrid1.DragOver
    e.Effect = DragDropEffects.All
End Sub



'-->

Function readWebFile(strFileName as string) 'lesen
  printLine 11, "readWebFile", strFileName
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  strFileName="es-code-snippets.txt"
  dim strContent as string
  strContent=ZZ.TwReadFile(strAppName, strFileName) 
  return strContent
  
  'ref.element("txt_outMonitor").text=strContent
end Function


sub saveWebFile(strFileName as string, strContent as string) 'speichern
  printLine 11, "readWebFile", strFileName
  'dim strAppName: strAppName="scriptIDE_test_es"

  dim strAppName as string
  strAppName="projektmanager"
  ''strAppName="ERR_kfldaskfldskflö"
  strFileName="es-code-snippets.txt"

  dim errMes as string
  '' Igrid_get(IGrid1, strContent, vbNewLine, "|")
  errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  if Not errMes.startsWith("OK:") then msgBox(errMes)
  ref.element("txt_outMonitor").text=errMes
  
end sub







'--> EOF

