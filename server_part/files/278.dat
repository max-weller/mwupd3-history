
'es_internSuche.vb

'-->
'--> G L O B A L   -   P A R A 

#Para DebugMode intern
#Para SilentMode true


'NEU...
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



'' #reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic

#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib

public WithEvents FormRef As Form
public WithEvents ref As Object
public shared toolBar As Object
public shared statBar As Object
public shared container1 As Object
public toolBarColor as string

shared winId as string ="es_internSuche.mainWin"
public winCaption as string = "internSucher: "
dim WithEvents myTextArea as textbox
dim WithEvents myPicture as pictureBox      
public withEvents myWin2 as object
dim myImg as object
'public WithEvents IGrid1 As Igrid
Dim WithEvents IGrid1 As IgridEx

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

dim globFileSpec as string ="C:\yPara\scriptIDE\scriptPara\"+"myBookmarks01.txt"


'-->
'--> T E S T 

sub test1()
  printLine 3, "test1","test1"
  monitorAdd("--- test1 ------------------------------------")
end sub

sub test2()
  printLine 3, "test2","test2"
  monitorAdd("--- test2 ------------------------------------")
end sub

sub test3()
  printLine 3, "test3","test3"
  monitorAdd("--- test3 ------------------------------------")
end sub






'-->
'--> I N I 

Sub GetFormRef()
  'If ref IsNot Nothing Then Exit Sub
  ''ref = ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  formRef = ref.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
End Sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls
End Sub

Sub onResizeControls()
   return
   trace "onResizeControls", formRef.Height
  'Igrid1.Width = formRef.Width - 10
  myTextArea.Height = container1.Height - 90 '42
End Sub



'-->
'--> A U T O S T A R T 
Sub AutoStart()
  'zz.traceClear()
  'zz.printLineReset()

  GetFormRef()
  'zz.doevents
  'dim id as string="es_internSuche.mainWin"
  'dim id as string="ToolBar|##|tbScriptWin|##|es_internsuche.mainwin"
  'ToggleDockWindowFromFormRef(formRef, 0)
  
'exit sub
  
  
  dim el as object
  toolBarColor="#BFBFBF"
  toolBarColor="#2A4100"
  
  With ref
    .resetControls ()
    '.activateEvents = "|IconMouseDown|   |TextboxKeyDown|"

    ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
    container1  =.addPanel("container1"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)
    container1.top=112
    toolBar.resetControls  (10,5,1)
    statBar.resetControls  (10,5,1)
    container1.resetControls  (10,5,1)
   
    toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    container1.BackColor = ColorTranslator.FromHtml("#fff")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
  end with

'--> toolBar    
  With toolbar
    toolbar.visible=true
   .height=48
     .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
 
    '.BR  '--------------------------------------------------
    ''.insertY=5: .insertX=5
    ''.addIcon("expand_right"  , "http://es.teamwiki.net/docs/icons/arrow_right.png" ,7 ,7 ,16,16  )
    ''.addIcon("expand_left"   , "http://es.teamwiki.net/docs/icons/arrow_left.png"  ,7 ,7 ,16,16  )


    .backColor=ColorTranslator.FromHtml("#243E56")
    .backColor=ColorTranslator.FromHtml("#2F3950")
    .backColor=ColorTranslator.FromHtml("#47493F")

    .OffsetX = 77 :. insertY = 11
    '' el=.addButton ("cmdExit"     , "Fenster schließen"  , "#E0E0E0"): el.width=113
                     'el=.addButton ("outMonitor_toClipboard"   , "--> Clipboard"   , "#E0E0E0"): el.width=77
     '' .insertY = 4  
     '' .insertX = 120: el=.addButton ("hideMe"    , "- X -"   , "#43759F"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     '' .insertX = 165: el=.addButton ("test1"    , "test 1"  , "#43759F"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     '' .insertX = 210: el=.addButton ("test2"    , "test 2"  , "#43759F"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     '' .insertX = 255: el=.addButton ("test3"    , "test 3"  , "#43759F"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     '' 
     ''   
     '' .insertY = 35
     '' .insertX = 143:.OffsetX = 143
     '' .insertX = 120 :  el=.addButton ("outMonitor_Clear"         , "Clear"           , "#FB7929"): el.width=66
     '' .insertX = 190: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66
     ''  el=.addButton ("toggleWeb"    , "W=E=B"    , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     ''  
    '...bild
    .insertX = 5  '70 '35 
    .insertY = 2 ' 11 ' 28 '-5
    '' myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/logMonitor.png" )
     myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/gnome-searchtool.png" )
 

    '--> ...suchen
    '--> ...topbar ========================================================
     '.insertX = 220: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66

     'errLine
    .insertX = 60:.insertY = 16
     el= .addTextbox ("such",   -2 )
     el= ref.element("txt_such")
     el.backColor=ColorTranslator.FromHtml("#CFD4AF")
     el.foreColor=ColorTranslator.FromHtml("#484A40")
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0
     el.text= "msgbox"
  end with


'--> ...container    

 with container1

 

    '' '--> ...textarea
    '' .insertX = 5 :.insertY = 90
    '' .addTextbox ("outMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
    ''    myTextArea=ref.element("txt_outMonitor")
    ''    myTextArea.WordWrap=false
    ''    myTextArea.scrollbars=System.Windows.Forms.ScrollBars.Vertical
    ''    myTextArea.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''    myTextArea.bringToFront()
    ''    '--> myTextArea.text=getMiniHelpText()
    '' '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
    ''   myTextArea.Top = 5
    ''   myTextArea.Left = 5
    ''   myTextArea.Width = container1.Width-10
    ''   myTextArea.Height = container1.Height - myTextArea.Top -4: 
    ''   myTextArea.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    '' 
'--> ...iGrid
    IGrid1 = New IgridEx
    IGrid1.DefaultCol.CellStyle.EmptyStringAs = TenTec.Windows.iGridLib.iGEmptyStringAs.EmptyString 'zack!
    'IGrid1.Top = 80 : IGrid1.Width = pLeft.Width : IGrid1.Height = .Height - 90: 
    IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    IGrid1.ReadOnly = True
    'leere Zelle als "" statt nothing
    IGrid1.Cols.Add("Nr",33)
    IGrid1.Cols.Add("r",20)
    IGrid1.Cols.Add("ZeilenInhalt",135)
    'IGrid1.Cols.Add("Interpret")
    'IGrid1.Cols.Add("Tags")
    
    'Igrid1.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
    
    'IGrid1.GroupBox.Visible = True
    container1.Controls.Add(IGrid1)
    
     IGrid1.Top = 1
     IGrid1.Left = 0
     IGrid1.Width = container1.Width-IGrid1.Left -0
     IGrid1.Height = container1.Height - IGrid1.Top -0: 
     IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    
   ref.visible=true
   ref.parent.visible=true
   ref.parent.parent.visible=true
   ref.show
  end with
  
  
'--> ...statBar   
  with statbar
      .visible=false
      .backColor=ColorTranslator.FromHtml("#939393")
      '.backColor=ColorTranslator.FromHtml("#243E56")
    
     .insertX = 20: .insertY= 10: 
     el= .addTextbox ("status",   160 )
     el= ref.element("txt_status")
     el.backColor=ColorTranslator.FromHtml("#777")
     el.foreColor=ColorTranslator.FromHtml("#bbb")
     el.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0

     .insertX = 205: 
     el= .addTextbox ("status2",   -2 )
     el= ref.element("txt_status2")
     el.backColor=ColorTranslator.FromHtml("#D7D7D7")
     el.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0
  End With

  'onResizeControls()
end sub



'-->
'--> E V E N T S

Sub onButtonEvent(e)
  '' ref.ResumeLayout(False)
  '' ref.PerformLayout()
  monitorClear
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim tag as string = e.sender.tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  dim action as string =name(1)
  monitorAdd("==============================================")
  monitorAdd("Sender.Name: " + e.Sender.Name+"<--")
  monitorAdd("      action: " +action+" <==")
  monitorAdd(" MouseButton: " +e.MouseButton+" <<<")
  'monitorAdd("   ClassName: " +e.ClassName+" <<<")
  monitorAdd(" ControlType: " +e.ControlType+" <<<")
  'monitorAdd("      MouseX: " +e.MouseX.toString+" <<<")
  monitorAdd("==============================================")
  Select Case action.toLower
    case "apppicture"      : searchCurDoc         (e, tagDATA)
    '' Case "caption"       '... nichts weiter machen
    Case else         
      dim errMes as string = "ERR: onButtonEvent(e): '"+name(1)+"' ...Typ nicht erkannt: "
      errorText(errMes)
      monitorAdd("! === ! === ! === ! === ! === ! === ! === ! === ! === ")
      monitorAdd(errMes)
      monitorAdd("! === ! === ! === ! === ! === ! === ! === ! === ! === ")
  End Select
  
  'ref auf toolStripPanel über parent...
  'dim id2 ="btn_expand_01"
  'dim el =  ref.element(id2)
  dim el =  e.sender
  'onTimerEvent()
  

  ref.ResumeLayout(False)
  ref.PerformLayout()
End Sub


sub txt_such_KeyDown(e)
  dim el as object = ref.element("txt_such")
  printLine 1, "e.keystring", e.keystring
  'trace  "keyDown(e)",      e.Keystring
  if e.keyString="RETURN" then
    searchCurDoc(nothing, nothing)
    el.focus
  end if
  if e.keyString="DOWN" then
    selectIGridLine(+1)
    el.focus
  end if
  if e.keyString="UP" then
    selectIGridLine(-1)
    el.focus
  end if
End Sub

sub selectIGridLine(deltaIndex as integer)
dim ig = iGrid1
  if ig.rows is nothing then 
    ig.rows.count=1
    exit sub
  end if
  if ig.curRow is nothing then
    printLine 2, "curRow", "nothing"
    ig.rows.count=1
    exit sub
  end if
  
  dim ir =ig.curRow
  dim curRowIndex = ig.curRow.index
  dim newIndex as integer =curRowIndex+deltaIndex
  ig.SetCurRow(newIndex)
  ''ig.row(newIndex).selected=true
  '' dim lineNumber as integer=cInt(ir.cells(0).value)
  '' monitorAdd("lineNumber: "+lineNumber.toString)
  '' 
  '' navigateCodeLink(lineNumber)
  '' 'ig.focus()
end sub


sub txt_such_TextChanged(e)
End Sub


sub txt_such_GotFocus(e)
End Sub


sub txt_such_LostFocus(e)
End Sub


sub txt_such_MouseUp(e)
End Sub


sub IGrid1_updateUserFeedback2(ByVal para1 As String) Handles iGrid1.updateUserFeedback2
  monitorAdd("para1: "+para1)
 'printLine 1, "insertLine", "START"
  dim ig = iGrid1
  if ig.rows is nothing then 
    ig.rows.count=1
    exit sub
  end if
  if ig.curRow is nothing then
    printLine 2, "curRow", "nothing"
    ig.rows.count=1
    exit sub
  end if
  
  dim ir =ig.curRow
  dim curRowIndex = ig.curRow.index
  dim lineNumber as integer=cInt(ir.cells(0).value): err.number=0
  monitorAdd("lineNumber: "+lineNumber.toString)
  
  navigateCodeLink(lineNumber)
  'ig.focus()

End Sub

sub navigateCodeLink(lineNumber as integer)
  'dim lineNumber as integer = cInt(tagDATA(4))
  monitorAdd("navigiere zu: " +lineNumber.toString)
  dim codeEditor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEditor.Lines.Current.text
  'codeEditor.Lines.Current.number=lineNumber
  codeEditor.goTo.line(lineNumber+50)
  codeEditor.goTo.line(lineNumber-10)
  codeEditor.goTo.line(lineNumber-1)
  'codeEditor.selection.start=100
  'codeEditor.selection.length=22
  'codeEditor.goTo.line(lineNumber).SelectionStartPosition()
  'codeEditor.select.line(lineNumber-1)
  codeEditor.focus()
end sub

:sub searchCurDoc(e, tagDATA)
  'msgbox("searchCurDoc")
  dim such2 as string = ref.element("txt_such").text
  dim such=such2.toLower
  dim tab as object=zz.getActiveRTF()
  'Dim scin As ScintillaNet.Scintilla = tab.RTF
  Dim scin As object = tab.RTF
  scin.focus()
  Dim codeText As String = scin.Text
  dim LINES() as string=split(codeText, vbNewLine)
  dim SCANN() as string=split(codeText.toLower, vbNewLine)
  dim treffer as string
  dim rank as string
  dim display as string
  dim isDefLine as boolean
  dim OUT() as string
  redim OUT(333)
  dim counter as integer=0
  dim trenn as string="<³³>"
  dim temp as string
  dim i as integer=0
  dim max as integer =LINES.length
  monitorAdd("____Suchbegriff: "+such2)
  monitorAdd("__Anzahl Zeilen: "+max.toString)
  monitorAdd("------------------------------------")
  dim item as string
  for i = 0 to max-1
    isDefLine=false
    item=trim(SCANN(i))
    if item.contains(such) then
      rank = "5"
      'monitorAdd(i.toString)
      treffer=trim(item)
      if treffer.contains("s"+"ub") then isDefLine=true
      if treffer.contains("f"+"unction") then isDefLine=true
      if treffer.contains("p"+"roperty") then isDefLine=true
      if isDefLine=true and not treffer.contains("exit")then
        dim findPos as integer = instr(item,"(")
        if findPos>0 then 
          item=mid(item,1,findPos-1)
          monitorAdd("-->"+item+"<--")
        end if
        rank="1"
      end if
      if treffer.startsWith("'") then 
         rank="9"
         temp=replace(LINES(i),"'","")
         LINES(i)="''"+trim(temp)
      end if
      'monitorAdd(i.toString+": "+rank+" --- "+trim(LINES(i)))
      OUT(counter)=(i+1).toString+trenn+rank+trenn+trim(LINES(i))
      counter=counter+1
    end if
  next
  redim preserve OUT(counter)
  dim RESULT as string=join(OUT, vbNewLine)
  monitorAdd(RESULT)
  Igrid_put(iGrid1,RESULT, vbNewLine, trenn)

  '--> ...sort
  iGrid1.GroupObject.Clear()
  iGrid1.SortObject.Clear()
  iGrid1.SortObject.Add(1, TenTec.Windows.iGridLib.iGSortOrder.Ascending)
  'iGrid.SortObject.Add("prio", TenTec.Windows.iGridLib.iGSortOrder.Descending)
  iGrid1.Sort()
  iGrid1.Group()

end sub



'-->
'-->  --> E V E N T S - 2 


function ToggleDockWindowFromId(id as string, optional forceState as integer=-1)as boolean
  dim toolWindow=ZZ.getDockPanelRef(id)
  
  '' exit function
  if toolWindow is nothing then exit function
  dim curState as boolean  = toolWindow.visible
  
  ''msgBox(curState)
  dim newState as boolean= not curState
  if forceState=0 then newState=false
  if forceState=1 then newState=true
  dim id2 as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  ''monitorAdd(id)
  ''monitorAdd(id2)
  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    if id=id2 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    if id=id2 then  toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function


function ToggleDockWindowFromFormRef(toolWindow as form, optional forceState as integer=-1)as boolean
  'dim toolWindow=ZZ.getDockPanelRef(id)
  
  '' exit function
  if toolWindow is nothing then exit function
  dim curState as boolean  = toolWindow.visible
  
  ''msgBox(curState)
  dim newState as boolean= not curState
  if forceState=0 then newState=false
  if forceState=1 then newState=true
  dim id2 as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  ''monitorAdd(id)
  ''monitorAdd(id2)
  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    'if id=id2 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    toolWindow.parent.visible =false
    'if id=id2 then  toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function



'-->
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByVal para1 As String)
Public Event updateUserFeedback2(ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByVal para1 As String) Handles iGrid1.updateUserFeedback1
'sub IGrid1_updateUserFeedback2(ByVal para1 As String) Handles iGrid1.updateUserFeedback2





Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  'printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  'printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
End Sub


Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  e.handled =true
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
'--> I G R I D - H E L P E R 

  
  
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
  



'-->
'--> outMONITOR

sub monitorAdd(para1 as string, optional para2 as string="")
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.add(para1+": "+para2+"<--")
   : end if
end sub

sub monitorClear()
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :    scriptClass.clear()
   : end if
end sub


sub statustext_reset()
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.statustext_reset
   : end if
end sub


sub errorText(para as string)
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.errorText(para)
   : end if
end sub

sub statustext(para as string)
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.statustext(para)
   : end if
end sub




'-->
'--> E O F


























