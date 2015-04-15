
' es_snippetManager.vb

#Para DebugMode intern
#Para SilentMode true

'-->
'-->N E U 
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


#reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic
#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib

public WithEvents FormRef As Form
public WithEvents ref As Object
shared winId as string ="es_snippetManager.mainWin"
public myWin2 as object
public winCaption as string = "Snippet-Manager (2)"

''public WithEvents IGrid1 As Igrid


'--> ini

Dim WithEvents IGrid1 As IgridEx
dim igFett = New TenTec.Windows.iGridLib.iGCellStyleDesign
dim  IgDefaultCellStyle = New TenTec.Windows.iGridLib.iGCellStyle(True)

shared myScriptClass



'''#####################################################
'''#####################################################

''http://snap.teamwiki.net/100211-203909-es-snippet-manager02.png


'--> FORM-ini
' private parent as Object

Dim WithEvents FormRef2 As Form
Dim ref2 As Object
shared winId2 as string ="es_snippetManager.zoomEditor"
dim curIGridRow as integer=-1
'' shared winId2 as string ="bookmark.vb"






'' public sub setParent(myParent)
''   parent=myParent
'' end sub
'' 
public sub hideSnippetEditor()
  printLine 1, "hideSnippetEditor", "...hide"
  '' with formRef2.parent.parent
  ''   .hide()
  '' end with
  formRef2.hide()
end sub



Sub GetFormRef2()
  'If ref IsNot Nothing Then Exit Sub
  'ref2 = ZZ.IDEHelper.CreateDockWindow(winID2,1)
  ref2 = ZZ.CreateWindow(winID2, DWndFlags.DockWindow, 1)'  : err.number=0

  formRef2=ref2.form
  formRef2=getOuterWindowRef(ref2)
  FormRef2.Text = "SnippetBearbeiten (2)"
  CType(FormRef2, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = false
End Sub


function getOuterWindowRef(panelRef as object) as object
  if panelRef.form.parent.parent is Nothing then
    '...falls es kein DockPanel-Fenster ist:
    :return panelRef.form
  else
    '...DockPanel-Fenster haben noch weitere Schichten dazwischen:
    :return panelRef.form.parent.parent
  end if
End function


'-->
'--> F O R M - 2 

sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


sub createEditWindow(index as integer)
  ''msgBox ("hallo")
  GetFormRef2()
  dim nicName as string=""
  dim code as string =""

  
  curIGridRow=index
  if index > -1 then
    dim ig as IGrid = IGrid1
    dim ir as iGRow = ig.rows(index)
    nicName =ir.cells(2).value.toString
    code=ir.cells(4).value.toString
    code=replace(code, "||LF||",vbNewLine)  
  end if
  
  With ref2
    dim el as object
    .activateEvents = "|TextboxKeyDown|  |TextboxKeyDown|  |TextboxTextChanged|"

    .resetControls ()
    .insertX = 11 :.insertY = 12
    .addIcon("xxx_appPicture", "http://es.teamwiki.net/docs/icons/insert-object64.png" )
     
    .BR  '--------------------------------------------------
    .insertX = 80  :.insertY = 12
     dim label as string=nicName
     el=.addLabel  ("lab01", label ,  ,"#ddd",,,-2,35) : el.backColor=ColorTranslator.FromHtml("#777"): makeJumboLabel(el)
     
    .BR  '--------------------------------------------------
   .insertX = 5 :.insertY = 70
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("nicName" ,  -2         , " nicName"    , 70  ,"#bbb" , , , )
    ref2.element("nicName").text=nicName
    ref2.element("nicName").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
 
   .insertX = 5 :.insertY = 100
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("codeTextarea" ,  -2         , "  Code"+vbNewLine+""    , 70 ,"#bbb" , , , "multiline=250")
    ref2.element("txt_codeTextarea").WordWrap=false
    ref2.element("txt_codeTextarea").text=code
    ref2.element("txt_codeTextarea").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    

     .BR  '--------------------------------------------------
   .insertX = 80 :.insertY = 360
   .addButton ("cmdOK"         , " O K "  , "#1DD910")
     .insertX = 160:
    .addButton ("cmdEscape"    , "Abbrechen"      , "#ccc")

  end with
  '' with formRef2.parent.parent
  ''   '.left=333
  ''   '.top=0
  ''   .show()
  '' end with
  '' dim getOuterWin=getOuterWindowRef(formRef)
  '' dim width=getOuterWin.width
  '' dim left=getOuterWin.left
  '' formRef2.left=left+width 'formRef.left+formRef.width-1
  '' formRef2.top=getOuterWin.top ' formRef.top
  '' formRef2.height=getOuterWin.height 'formRef.height
  '' formRef2.width=600
  formRef2.show()
  formRef2.bringToFront
end sub



'-->
'--> EVENTS -form 2

sub nicName_keyDown(e)
  ''....bringtNichts
  printLine 1, "keyDown(e)",  "keyDown(e)"
  ''printLine 2, "keyDown(e)",  "keyDown(e)"
  ''printLine 3, "keyDown(e)",  "keyDown(e)"
  trace 1, "keyDown(e)", e.Keystring
End Sub

sub nicName_TextChanged(e)
  printLine 2, "nicName_TextChanged",  e.Sender.text
  ''printLine 11, "keyDown(e)",  "keyDown(e)"
  ''printLine 22, "keyDown(e)",  "keyDown(e)"
  ''printLine 23, "keyDown(e)",  "keyDown(e)"
  'trace 1, "nicName_TextChanged", e.keyCode
  dim newText as string=e.Sender.text
  setNewTextToIgridCell(2, newText)
  ref2.element("nicName").text=newText
  ref2.element("lab_lab01").text=newText
End Sub


sub codeTextarea_TextChanged(e)
  'printLine 3, "codeTextarea_TextChanged", e.keyCode
  'trace  "codeTextarea_TextChanged", e.keyCode
  dim newText as string=e.Sender.text
  setNewTextToIgridCell(4, newText)
  ref2.element("txt_codeTextarea").text=newText
End Sub


sub setNewTextToIgridCell(colIndex as integer, newText as string)
  dim index as integer=curIGridRow
  if index > -1 then
    dim ig as IGrid = IGrid1
    dim ir as iGRow = ig.rows(index)
    newText=replace(newText, vbNewLine, "||LF||") 
    ir.cells(colIndex).value = newText
   end if 
end sub



sub cmdEscape_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  hideSnippetEditor()
End Sub

sub cmdOK_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  hideSnippetEditor()
  saveSnippetFile() 
End Sub




'-->
'--> T E S T 


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

''sub test3()
''  msgBox("OK - 3")
''end sub

'-->


'-->
'--> I N I - T E R M I N A T E

Sub __GetFormRef()

  'If ref IsNot Nothing Then Exit Sub
  'ref = ZZ.IDEHelper.CreateDockWindow(winID, 1)
  'ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0 
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  'formRef = ref.form
  formRef=getOuterWindowRef(ref)

  formRef.text=winCaption
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
  
  ''DAS MACHT EINE NORMALE FORM
  ''formRef = New frmTB_scriptWin 'ZZ.IDEHelper.CreateForm(winID)
  ''ref = CType(FormRef, Object).PNL
  ''FormRef.Text = "Snippet-Manager"
End Sub

Sub GetFormRef()
  If ref IsNot Nothing Then Exit Sub
  '' PanelRef = ZZ.IDEHelper.CreateDockWindow(winID)
  '' FormRef = PanelRef.Form
  '' ref = ZZ.IDEHelper.CreateDockWindow(winID, 1)
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)
  formRef=ref.form
  ''formRef.hide
  formRef.Text = winCaption
End Sub




Sub globAddHandler()
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf onTimerEvent
  AddHandler formRef.Resize, AddressOf Form_Resize
  'AddHandler scNet.MouseDown, AddressOf scNet_MouseDown
  setMonitorRef()
  monitorAdd("globAddHandler")
end sub


Sub globRemoveHandler()
  'trace "globRemoveHandler","try..."
  setMonitorRef()
  monitorAdd("!!! globRemoveHandler")
  
  'if formRef is nothing then exit sub
  'RemoveHandler Timer1.Tick, AddressOf onTimerEvent
  RemoveHandler formRef.Resize, AddressOf Form_Resize
  'RemoveHandler scNet.MouseDown, AddressOf scNet_MouseDown
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'trace "globRemoveHandler","DONE"
end sub




Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  'glob.saveFormPos(FormRef)
  'glob.saveParaFile()
  trace "formClosing" , "xxx"
  'globRemoveHandler()
End Sub


Sub onTerminate()
  'trace "!!! onTerminate", "!!! es_codeList"
  globRemoveHandler()
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
End Sub



'''#####################################################
'''#####################################################


'''#####################################################
'''#####################################################


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls
End Sub

Sub onResizeControls()
  Igrid1.Width = ref.Width -10
  Igrid1.Height =ref.Height - 96
End Sub

'-->
'--> A U T O S T A R T
'''#####################################################
Sub AutoStart()

  myScriptClass = me

  ''myWin2 = ZZ.scriptClass("es_snippetEditor")
  ''myWin2.setParent(me)
  zz.traceClear()
  zz.printLineReset()
  
  GetFormRef()
  With ref
     .BackColor = ColorTranslator.FromHtml("#17A90C")

     dim el as object
    .resetControls ()
    .insertX = 11 :.insertY = 12
    .addIcon("appPicture", "http://es.teamwiki.net/docs/icons/insert-object64.png" )

   .insertX = 75 :.insertY = 3
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("outMonitor" ,  -2         , "Roh-"+vbNewLine+"Daten"    , 0  , "#afa", , , "multiline=50")
    ref.element("txt_outMonitor").WordWrap=false



.BR  '--------------------------------------------------
     .insertX = 5 ::.insertY = 60
     .addButton ("insertItem"     , "Einfügen"  , "#1DD910")
'.BR  (24)'--------------------------------------------------
     '.insertX = 75 :
    .addButton ("showHideZoom"    , "Bearbeiten"      , "#ccc")

 '.BR (24) '--------------------------------------------------
    '.insertX = 150
    .addButton ("addLine"        , "N E U"  , "#DF7500")

'.BR  '--------------------------------------------------

'.BR (55) '--------------------------------------------------
'    .insertX = 5 
    el = .addButton ("read"            , "Lesen"      , "#ccc")
    el.width= 70
'    .insertX = 80
    el = .addButton ("save"            , "Speichern"  , "#ccc")
    el.width=70


End With
  IGrid1 = New IgridEx
  IGrid1.Top = 90
  iGrid1.left=5
  'leere Zelle als "" statt nothing
  
  iGrid1.DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
  iGrid1.DefaultCol.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
  iGrid1.DefaultCol.ColHdrStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None
  
  IGrid1.Cols.Add("id").width=5
  IGrid1.Cols.Add("gruppe").width=18
  IGrid1.Cols.Add("nickname").width=250
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
  
  '    '''ZZ.scriptClass("es_snippetEditor").createEditWindow(0)
  'createEditWindow(0)
  ref.show
  ref.visible=true
  ref.parent.show
  ref.parent.visible=true
  ref.parent.parent.show
  ref.parent.parent.visible=true
  
  globAddHandler()
  
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
  'msgBox("zoomInSnippetEditor")
  dim ig as IGrid = IGrid1
  dim curRow as iGRow
  curRow=ig.curRow
  dim curRowIndex as integer = curRow.index
  ''...waren mal getrennte Fenster
  ''ZZ.scriptClass("es_snippetEditor").createEditWindow(curRowIndex)
  createEditWindow(curRowIndex)
end sub

Sub read_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  readSnippetFile()
End Sub

Sub save_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  saveSnippetFile() 
End Sub





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
'--> i G R I D 

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
    'switchToCodeEditor()
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
'--> R E A D / S A V E




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



'-->
'--> ----------------------------------- L I B s (auslagern) 



'-->
'-->  I G R I D - E X 

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
Public Event updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback1 '@@-
'sub IGrid1_updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback2 '@@-

private g_fileSpec as string
: Property fileSpec () As String
:   Get 
:     Return g_fileSpec
:   End Get
:   Set (ByVal value As String)
:     g_fileSpec=value
:   End Set
: End Property

  '' Property ItemText(ByVal idx) As String
  ''   Get
  ''     Return Me.Items(idx)
  ''   End Get
  ''   Set(ByVal value As String)
  ''     Me.Items(idx) = value
  ''   End Set
  '' End Property

Private Sub myIgrid_ColHdrClick(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGColHdrClickEventArgs) Handles myBase.ColHdrClick
  '... dann sortiert er nicht
  e.DoDefault = False
End Sub

: Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
  if e.Button= Windows.Forms.MouseButtons.Right then
    printLine  8, "IGrid1_CellMouseDown - col", "rClick"
  else
    printLine  8, "IGrid1_CellMouseDown - col", "NORMAL"
  end if
End Sub


: Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  dim keyCode=e.keycode
  myScriptClass.setMonitorRef()
  myScriptClass.monitorAdd("myIgrid_KeyDown")

  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  if e.control=true then
    'trace "CONTROL",e.keycode 
    'myScriptClass.monitorAdd("CONTROL: "+keyCode.toString)
    myScriptClass.monitorAdd("e.control=true")
    if keycode=88 then onClipboardCutSelected(me)   'X
    if keycode=67 then onClipboardCopySelected(me)  'C
    if keycode=86 then onClipboardPasteBeforeSelectedLine(me) 'V
    e.handled =true
  end if
End Sub


: sub onClipboardCutSelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardCut()")
  onClipboardCopySelected(iGrid,true)
end sub


: sub onClipboardCopySelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid, optional remove as boolean=false)
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

: sub onClipboardPasteBeforeSelectedLine(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
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



: Private Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

: Private Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
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

  
: Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
  Dim max = line.Cells.Count - 1
  Dim out(max) As String
  :For i As Integer = 0 To max
    :out(i) = line.Cells(i).Value.ToString
  :Next
  :Return Join(out, Delimiter)
End Function

: Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  Dim max = line.Cells.Count - 1
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  :For i As Integer = 0 To max
  :  line.Cells(i).Value = out(i)
  :Next
End Sub

: Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim max = Grid.Rows.Count - 1
  Dim out(max) As String
  For i As Integer = 0 To max
    out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
  Next
  FileContent = Join(out, LineDelim)
End Sub
  
  
: Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim out() = Split(FileContent, LineDelim)
  Grid.Rows.Clear()
  Grid.Rows.Count = out.Length
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub

: Sub Igrid_put2(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
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


: sub igrid_setDefaultPara(iGrid)
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

public globMonitorRef as object

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
  'trace "globMonitorRef", globMonitorRef.name
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





