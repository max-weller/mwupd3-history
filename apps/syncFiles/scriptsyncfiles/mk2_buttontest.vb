



'-->
'-->START



#Para DebugMode extern
#Para IconFile c:\icon.ico

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports Tentec.Windows.iGridLib
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper

Dim PnlRef As ScriptedPanel

Dim withevents tim as new Timer




'-->
Sub AutoStart()
'--> Scripted Panel/Formeigenschaften
  PnlRef = ZZ.CreateWindow("mk2_buttontest.Hello", DWndFlags.DockWindow)
  with PnlRef
    'DockWindow Position merken oder nicht
    CType(.Form, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = True
    'Formdesign
    .Form.Size = New Size(300,200)
    .Form.Text = "scriptIDE Starter Edition"
    .ResetControls(50, 50)
    .addIcon("myIcon", "http://mw.teamwiki.net/docs/img/win-icons/spcplui_138-48.png", , , -2)
    '.addButton("btn1", "ffffff")
    .addTextbox("Birthdate", 50)
    .br
    .addButton("btnCalcAge", "alter berechnen")
    .br
    .addLabel("Age", "ssdfsd")    
  end With
End Sub


'-->
'--> FORM1
Sub btnCalcAge_MouseClick(e)
    Dim year as string = My.Computer.Clock.LocalTime.Year
    Dim birth as integer = PnlRef("txt_Birthdate").Text
    If birth > year then
      msgbox("Wie willst du bitteschön in der Zukunft geboren sein?")
      exit sub
    end if
    PnlRef("lab_Age").Text = (Val(year) - Val(PnlRef("txt_Birthdate").Text)).ToString
End Sub

Sub tim_Tick(sender As Object, e As EventArgs) Handles tim.Tick
  tim.stop
  msgbox("test")
End Sub

Sub btn1_MouseClick(e)
  tim.Interval = 5000
  tim.start
End Sub



'-->
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
'Public Event updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
'Public Event updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
Public Event updateUserFeedback1(ByRef iGrid As IgridEx, ByVal para1 As String)
Public Event updateUserFeedback2(ByRef iGrid As IgridEx, ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback1
'sub IGrid1_updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback2

private g_fileSpec as string
: Property fileSpec () As String
:   Get 
:     Return g_fileSpec
:   End Get
:   Set (ByVal value As String)
:     g_fileSpec=value
:   End Set
: End Property

private g_name as string
: Property name () As String
:   Get 
:     Return g_name
:   End Get
:   Set (ByVal value As String)
:     g_name=value
:   End Set
: End Property


Sub myIgrid_ColHdrClick(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGColHdrClickEventArgs) Handles myBase.ColHdrClick
  '=false ... dann sortiert er nicht
  e.DoDefault = true
End Sub

Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  'printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  'printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
End Sub


Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  dim keyCode=e.keycode
  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  if e.control=true then
    'trace "CONTROL",e.keycode 
    'myScriptClass.monitorAdd("CONTROL: "+keyCode.toString)
    if keycode=88 then onClipboardCutSelected(me)   'X
    if keycode=67 then onClipboardCopySelected(me)  'C
    if keycode=86 then onClipboardPasteBeforeSelectedLine(me) 'V
    e.handled =true
  end if
End Sub


sub onClipboardCutSelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardCut()")
  onClipboardCopySelected(iGrid,true)
end sub


sub onClipboardCopySelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid, optional remove as boolean=false)
  myScriptClass.monitorAdd("CONTROL: onClipboardCopy()")
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  dim max =iGrid.Rows.count
  dim OUT(max) as string
  dim DEL(max) as integer
  dim counter as integer=0
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then 
      'myScriptClass.monitorAdd(row.index.toString)
      OUT(counter)=myScriptClass.JoinIGridLine(row)
      DEL(counter)=row.index
      counter=counter+1
      anz = anz +1
    end if
  Next
  redim preserve OUT(counter-1)
  redim preserve DEL(counter-1)
  if remove=true then
    dim i as integer
    for i=DEL.length-1 to 0 step -1
      iGrid.Rows.RemoveAt(DEL(i))
    next
  end if
  dim RESULT as string=join(OUT, vbNewLine)
  zz.setClipboardText(RESULT)

  myScriptClass.monitorAdd(RESULT)
  myScriptClass.monitorAdd("==========================================")
  myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  myScriptClass.monitorAdd("")
end sub


sub onClipboardPasteBeforeSelectedLine(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardPasteBeforeSelectedLine()")
  dim selIndex as integer = -1
  
  '...die erste selektierte zeile ermitteln
  Dim row As TenTec.Windows.iGridLib.iGRow
  For Each row In iGrid.Rows
    if row.Selected=true then
      if selIndex=-1 then selIndex=row.index
      row.Selected=false
    end if
  Next

  '... noch nicht entschieden
  if selIndex=-1 then
    '... es ist nichts Selektiert
    '... in Zeile 1 einfügen / ... oder ans ende ???
    'ist grid überhaupt initialisiert ???
  end if
  
  dim newData as string =zz.getClipboardText()
  myScriptClass.monitorAdd(newData)
  dim DATA=split(newData,vbNewLine)
  dim i as integer=0
  dim skip as boolean =true
  for i=DATA.length-1 to 0 step -1
    dim item=DATA(i)
    iGrid.Rows.Insert(selIndex)
    myScriptClass.SplitToIGridLine(iGrid.Rows(selIndex), item)
    '... alle außer der ersten einfügung
    if skip = true  then 
      skip = false
    else
      iGrid.Rows(selIndex+1).selected=true
    end if
  next
  '...den ersten auch
  iGrid.Rows(selIndex).selected=true
  '' myScriptClass.monitorAdd(RESULT)
  '' myScriptClass.monitorAdd("==========================================")
  '' myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  '' myScriptClass.monitorAdd("")
end sub



Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
  dim curRow as integer
  curRow=me.curRow.index
  trace "curRow", curRow
  RaiseEvent updateUserFeedback2(me, curRow.ToString)
End Sub


:Sub checkSelectionInRowMode(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then anz = anz +1
  Next
  RaiseEvent updateUserFeedback1(me, anz.ToString)

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


Sub Igrid_put2(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim out() = Split(FileContent, LineDelim)
  Grid.Cols.Clear
  Grid.Rows.Count = out.Length
  dim header()=Split(out(0), ColDelim)
  for j as integer=0 to header.length-1
    grid.Cols.Add(header(j),66)
  next
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub


sub igrid_setDefaultPara(iGrid)
  with iGrid
    .DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
    .DefaultCol.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
    .DefaultCol.CellStyle.EmptyStringAs = TenTec.Windows.iGridLib.iGEmptyStringAs.EmptyString 'zack!
    .DefaultCol.ColHdrStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None
    .SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
    .SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText
  end with
end sub





'-->
'--> outMONITOR

 sub clearAll()
   on error resume next
   monitorClear()
   statustext_reset()
   zz.traceClear()
   zz.printLineReset()
   err.number=0
end sub



sub setMonitorRef()
  'on error resume next
  globMonitorRef = zz.scriptClass("es_testLabor")
  trace "globMonitorRef", globMonitorRef.name
end sub

: function getMonitorRef()
  on error resume next
  if globMonitorRef is nothing then
    globMonitorRef = zz.scriptClass("es_testLabor")
  end if
  return globMonitorRef
end function

: sub monitorClear()
   on error resume next
   globMonitorRef.clear()
   err.number=0
end sub


: sub monitorAdd(para1 as string, optional para2 as string="")
   on error resume next
   globMonitorRef.add(para1+": "+para2+"<--")
   err.number=0
end sub

: sub statustext(para1 as string, optional para2 as string="")
   on error resume next
   'globMonitorRef.statustext(para1+": "+para2+"<--")
   globMonitorRef.statustext(para1+"")
   err.number=0
end sub

: sub statustext_reset()
   on error resume next
   globMonitorRef.statustext_reset()
   err.number=0
end sub


: sub errorText(para as string)
   on error resume next
   globMonitorRef.errorText(para)
   err.number=0
end sub


