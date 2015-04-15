'' ...zwischengelagert

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


'' es_popUpMonitor                          

#Para DebugMode intern
#Para SilentMode true

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

shared winId as string ="es_popUpMonitor.mainWin"
public winCaption as string = "Test-Labor"
dim WithEvents myTextArea as textbox
dim WithEvents myPicture as pictureBox      
public withEvents myWin2 as object
dim myImg as object
public WithEvents IGrid1 As Igrid

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

dim globFileSpec as string ="C:\yPara\scriptIDE\scriptPara\"+"myBookmarks01.txt"


'-->
'--> T E S T 

sub test1()
  printLine 3, "aaa","bbbbb"
  add("---------------------------------------")
  '' ZZ.IDEHelper.Exec("Window.Reflection", "zz","")
  ZZ.IDEHelper.Exec("File.Save", "","")
  ZZ.IDEHelper.Exec("Debug.Run", "","")
  
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
      
  : scriptClass.test2()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if

''zz.IDEHelper.window.reflection("zz")
  'ZZ.IIDEIndexList.RestorePos(30,20)

  'ZZ.IDEHelper.indexList.RestorePos(30,20)
   add(ZZ.getActiveRTF.RTF.Lines.Current.Text)
end sub

sub test2()
  msgBox("bbbbbbbbbbbb")
  dim mainWin as object: mainWin =ZZ.IDEHelper.MainWindow
  trace mainWin.ActiveMdiChild,toString
  add(ZZ.getActiveRTF.RTF.toString)
  add(ZZ.getActiveRTF.RTF.Lines.Current.number)
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim lineNumber as integer = codeEitor.Lines.Current.number
  dim lineContent as string = codeEitor.Lines.Current.text
  add(lineNumber.toString+" --> "+lineContent+"<--")
  add(lineNumber.toString+" --> "+trim(lineContent)+"<--")
 end sub
 
sub test3()
  ZZ.IDEHelper.getToolBarList()
  
   'msgBox(  ZZ.IDEHelper.getToolBarList)
exit sub

   msgBox(  ZZ.IDEHelper.glob.Para("userToolbar3-paraTest"))
   ZZ.IDEHelper.glob.saveParaFile
exit sub


  ''Sub refreshMainTitle()
  dim MAIN=ZZ.IDEHelper.getMainFormRef
    On Error Resume Next
    Dim prefix = ""
    If MAIN.chkSticky.Checked Then prefix = "* "
    Dim filename = ""
    Dim win As Object = MAIN.ActiveMdiChild
    
    dim labelText1 as string
    dim labelText2 as string

   If win IsNot Nothing Then
      'MAIN.PictureBox2.Image = win.getIcon()(1).ToBitmap
      filename = ": " + win.getViewFilename()
      'MAIN.labWinTitle.Text = win.getViewFilename()
       labelText1 =filename
      labelText2 =win.getFileTag()
      'setLabelText(MAIN.lblGlobAktFileSpec, win.getFileTag())
    Else
      'MAIN.PictureBox2.Image = MAIN.Icon.ToBitmap
      'MAIN.labWinTitle.Text = "scriptIDE"
      'setLabelText(MAIN.lblGlobAktFileSpec, "")
     labelText1 ="scriptIDE"
      labelText2 ="???" 'MAIN.lblGlobAktFileSpec

    End If
    add(labelText1 )
    add(labelText2 )
    add(MAIN.chkSticky.Checked)
    ''MAIN.Text = prefix + "scriptIDE2" + filename

 '' End Sub

exit sub





  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
add(activeTabType)
add(ActiveTabFilespec)
add(activeTab.toString)
add(activeTab.parent.toString)
add(activeTab.parent.toString)
add(activeTab.parent.parent.toString)
exit sub
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim selection as string =codeEitor.selection.text
  add(selection)
end sub

:function callScriptClassTestProc(filespec as string)
  '' printLine 11, "...", "..." 
  '' printLine 1, "splitFilespecData", fileSpec
  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  '' :printLine 2, "splitFilespecData", name
  :statustext(name)
  :scriptClass=zz.scriptClass(name)
  :zz.doEvents
  :zz.doEvents
  :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     statustext("ERR(scriptClass): "+filespec)
     return nothing
  end if
  return scriptClass
End function



'--> ...
Sub GetFormRef()
  'If ref IsNot Nothing Then Exit Sub
  ref = ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0
  formRef = ref.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
End Sub

Sub GetFormRef__xxx()
  ''DAS MACHT EINE NORMALE FORM
  formRef = New frmTB_scriptWin 
  'ref = ZZ.IDEHelper.CreateForm(winID)
  ref = CType(FormRef, Object).PNL
  FormRef.Text = winCaption
  formRef.topmost=true
  formRef.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
End Sub

sub show()
  ''msgBox ("show(es_popUpMonitor.mainWin)")
  FormRef.show()
  : FormRef.visible=true
  :FormRef.parent.visible=true
  :FormRef.parent.parent.visible=true
end sub
  
sub hide()
  'FormRef.hide
  'FormRef.parent.hide
  '' FormRef.parent.parent.hide
   '  FormRef.parent.visible=false
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_popUpMonitor.mainWin")
  toolWindow.hide()
end sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls
End Sub

Sub onResizeControls()
   trace "onResizeControls", formRef.Height
  'Igrid1.Width = formRef.Width - 10
  myTextArea.Height = container1.Height - 95 '42
End Sub


'''#####################################################









'-->
'--> A U T O S T A R T 
Sub AutoStart()
  'zz.traceClear()
  'zz.printLineReset()
  GetFormRef()
  dim el as object
  toolBarColor="#BFBFBF"
  toolBarColor="#2A4100"
With ref
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"

    ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
    container1  =.addPanel("container1"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)
    toolbar.visible=false
    container1.top=112
    toolBar.resetControls  (10,5,1)
    statBar.resetControls  (10,5,1)
    container1.resetControls  (10,5,1)
   
    toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    container1.BackColor = ColorTranslator.FromHtml("#fff")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
  end with

'--> DATA: toolBar    
  With toolbar
    '.height=1
     .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
 
    '.BR  '--------------------------------------------------
    ''.insertY=5: .insertX=5
    .addIcon("expand_right"  , "http://es.teamwiki.net/docs/icons/arrow_right.png" ,7 ,7 ,16,16  )
    .addIcon("expand_left"   , "http://es.teamwiki.net/docs/icons/arrow_left.png"  ,7 ,7 ,16,16  )
  end with



'--> ...container    

 with container1

    .backColor=ColorTranslator.FromHtml("#BFBFBF")
    .backColor=ColorTranslator.FromHtml("#222")
    .backColor=ColorTranslator.FromHtml("#243E56")
    .backColor=ColorTranslator.FromHtml("#253833")
     .OffsetX = 77 :. insertY = 16
    '' el=.addButton ("cmdExit"     , "Fenster schließen"  , "#E0E0E0"): el.width=113
                     'el=.addButton ("outMonitor_toClipboard"   , "--> Clipboard"   , "#E0E0E0"): el.width=77
     .insertY = 8  
     .insertX = 100: el=.addButton ("hideMe"    , "- X -"          , "#345F52"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     .insertX = 145: el=.addButton ("test1"    , "test 1"          , "#345F52"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     .insertX = 190: el=.addButton ("test2"    , "test 2"          , "#345F52"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     .insertX = 235: el=.addButton ("test3"    , "test 3"          , "#345F52"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")

  
     .insertY = 35
     .insertX = 143 : .insertY = 40
 '               el=.addButton ("saveMe"    , "save"     , "#222"): el.width=45:  el.foreColor=ColorTranslator.FromHtml("#eee")
     .insertX = 100 :  el=.addButton ("outMonitor_Clear"         , "Clear"           , "#FB7929"): el.width=66
     .insertX = 170: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66
     '' '.BR(24):  
     ''  el=.addButton ("readMe"    , "read"     , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     ''  '.BR(24):  
     ''  el=.addButton ("testy2"    , "r- C -"   , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     ''  '.BR(24):  
     ''  el=.addButton ("testy4"    , "- D -"    , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
     ''  '.BR(24):  
      el=.addButton ("toggleWeb"    , "W=E=B"    , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
 
    '...bild
    .insertX = 15  '70 '35 
    .insertY = 0 ' 11 ' 28 '-5
     ''myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/logMonitor.png" )
     'myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/gnome-monitor.png" )
     myPicture=.addIcon("appPicture", "http://icons3.iconfinder.netdna-cdn.com/data/icons/humano2/72x72/apps/gnome-monitor.png" )
  
    '--> ...topbar ========================================================
     '.insertX = 220: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66

     'errLine
    .insertX = 10:.insertY = 73 
     el= .addTextbox ("errMes",   -2 )
     el= ref.element("txt_errMes")
     el.backColor=ColorTranslator.FromHtml("#CEF2E9")
     el.foreColor=ColorTranslator.FromHtml("#b00")
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0
     el.text= "lastErr"


    '--> ...textarea
    .insertX = 5 :.insertY = 95
    .addTextbox ("outMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
       myTextArea=ref.element("txt_outMonitor")
       myTextArea.WordWrap=false
       myTextArea.scrollbars=System.Windows.Forms.ScrollBars.Vertical
       myTextArea.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
       myTextArea.bringToFront()
       '--> myTextArea.text=getMiniHelpText()
    '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown

    
   ref.visible=true
   ref.show
   show
  end with
  
  
'--> ...statBar   
  with statbar
      .backColor=ColorTranslator.FromHtml("#939393")
      .backColor=ColorTranslator.FromHtml("#253833")
    .insertX = 20: .insertY= 10: 
     el= .addTextbox ("status",   160 )
     el= ref.element("txt_status")
     el.backColor=ColorTranslator.FromHtml("#D7D7D7")
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

  onResizeControls()
  readBookmarkData()
end sub

function getMiniHelpText()
#Data myMiniHelp as String
=======================================
Test 1 (,2,3) ruft die sub test1 (,2,3) 
im gerade aktiven Fenster auf.


Die Felder 1-18 merken sich die 
aktuelle CursorPosition und den 
zugehörigen Dateinamen 
(...setzen über rClick).

Bei normalem Click wird zur 
jeweiligen Datei gesprungen
und der Cursor wieder an die 
gleiche Stelle gesetzt.

======================================
   Anregen sind herzlich willkommen
======================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=myMiniHelp 
  return RESULT
end function


Sub myTextArea_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)Handles myTextArea.KeyDown
  trace "textarea_KeyDown",e.keyCode
  ref.element("txt_status").text=e.keyCode
end sub

Sub myPicture_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles myPicture.DoubleClick
  trace "myImage_DoubleClick","..........."
End Sub

 'Sub myPicture_MouseDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles myPicture.MouseDown
Sub myPicture_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles myPicture.MouseDown
  'trace "PictureBox1_MouseDown","..........."
  'formRef.parent.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
  if e.Button=Windows.Forms.MouseButtons.Right then
    hide()
    exit Sub
  end if
  
  dim hwnd as integer
  Call ReleaseCapture()
  ''Call SendMessage(myPicture.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  
  Call SendMessage(FormRef.parent.parent.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  ''Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub


'-->
'--> E V E N T S

sub onLabelEvent(e)
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  Add("==============================================")
  Add("Sender.Name ===>" + e.MouseButton)
  ''Add("keystring ===>" + e.keyString)
  '' dim tag as string = e.sender.tag.toString
  '' dim tagDATA()= Split(tag, "<³³>")
  Add(name(1))
  '' if name(1) = "navNr" then
  ''   setNewCodeLink(e)
  '' end if
  if name(1) = "navLab" then
    dim highlighter=ref.element("lab_highlighter_0")
    highlighter.top=e.sender.top-1
    highlighter.left=e.sender.left-1
    highlighter.height=e.sender.height+2
    highlighter.width=e.sender.width+2
    highlighter.visible=true
  
    if e.MouseButton = "Right" then
      setNewCodeLink(e)
      saveCodeLink()
    else
      navCodeLink(e)
    end if
  end if
end sub

sub saveCodeLink()
  dim OUT(100) as string
  dim tag as string=""
  dim counter as integer = 0
  dim el as object
  dim i as integer=0
  dim max as integer = container1.controls.count
  add(max.toString)
  for i = 0 to max-1
    el=container1.controls(i)
    '' add(i.toString()+" <-> "+el.name.toString())
    if not el.name is nothing and el.name.startsWith("lab_navLab")then
      tag=el.tag.toString()
      if tag = "System.Object[]" then tag=""
      out(counter)=tag
      counter =counter+1
    end if
  next
  redim preserve OUT(counter)
  dim RESULT as string = join(OUT, vbNewLine)
  add(RESULT+"<--(endOfArray)")
  
  '...forTest
  restoreCodeLink(RESULT)
end sub

sub restoreCodeLink(codeLinkDATA)
  dim LINES() as string=split(codeLinkDATA, vbNewLine)
  dim lineData as string
  dim FIELDS() as string
  dim id as string
  dim labelText as string=""
  dim counter as integer = 0
  dim el as object
  dim i as integer=0
  dim max as integer = LINES.length
  '' add(max.toString)
  for i = 0 to max-1
    lineDATA=LINES(i)
    if trim(lineDATA)="" then continue for
    FIELDS =split(lineDATA,"<³³>")
    labelText=FIELDS(3)
    id="lab_navLab_"+(i+1).toString()
    el=container1.controls(id)
    '' add(i.toString()+" <-> "+el.name.toString())
    if el is nothing then continue for
    el.tag=lineDATA
    el.text=labelText
  next
end sub

sub toggleWeb_MouseClick(e)
  ''msgBox("toggleWeb")
  dim id = ("ToolBar|##|tbScriptWin|##|es_webbrowser3.tab")
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit sub
  dim curState as boolean  = toolWindow.visible
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
   else
    toolWindow.hide()
    toolWindow.visible =false
  end if
end sub


'-->
 
sub setNewCodeLink(e)
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim lineContent as string = trim(codeEitor.Lines.Current.text)
  lineContent=replace(lineContent,vbNewLine,"")
  lineContent=replace(lineContent,"sub","S:",,,1)
  lineContent=replace(lineContent,"function","F:",,,1)
  lineContent=replace(lineContent,"private","",,,1)
  lineContent=replace(lineContent,"public","",,,1)
  lineContent=trim(lineContent)
  lineContent=replace(lineContent," ","_")
  dim lineNumber as integer = codeEitor.Lines.Current.number
  
  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  
  dim tagDATA(10) as string
  tagDATA(0)="index"
  tagDATA(1)=ActiveTabFilespec
  tagDATA(2)="subFunc"
  tagDATA(3)=lineContent
  tagDATA(4)=lineNumber.toString
  tagDATA(5)="reserve1"
  dim newTag as string=join(tagDATA,"<³³>")
  'add (newTag)
  add(lineContent)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim index as string=name(2)
  dim id ="lab_navLab_"+index+"<--"
  add("id: "+id)
  dim label=ref.element("lab_navLab_"+index)
  label.text=lineContent
  label.tag=newTag
  add(e.sender.text)
end sub

sub navCodeLink(e)
  ' 
  dim tag as string = e.sender.tag.toString
  add("tag: "+tag+"<--")
  
  if tag = "System.Object[]" then exit sub
  
  dim tagDATA()= Split(tag, "<³³>")
  '' tagData(0)="index"
  '' tagData(1)="filespec"
  '' tagData(2)="subFunc"
  '' tagData(3)=lineContent
  '' tagData(4)=lineNumber.toString
  '' tagData(5)="reserve1"
  
  dim fileSpec as string=tagDATA(1)
  
  '...provisorisch
  hideWebBrowser()
  
  dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
  
  dim lineNumber as integer = cInt(tagDATA(4))
  add("navigiere zu: " +lineNumber.toString)
  dim codeEitor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEitor.Lines.Current.text
  'codeEitor.Lines.Current.number=lineNumber
  codeEitor.goTo.line(lineNumber+50)
  codeEitor.goTo.line(lineNumber-10)
  codeEitor.goTo.line(lineNumber)
  codeEitor.focus()
  
end sub

sub hideWebBrowser()

  ''msgBox("toggleWeb")
  dim id = ("ToolBar|##|tbScriptWin|##|es_webbrowser3.tab")
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit sub
  toolWindow.hide()
  toolWindow.visible =false
end sub


sub saveMe_MouseClick(e)
  ''add("saveMe__MouseClick")
  saveBookmarkData()
end sub


sub readMe_MouseClick(e)
  'add("readMeMe_MouseClick")
  readBookmarkData()
end sub


sub saveBookmarkData()
  dim preFix as string ="lab_navLab_"
  dim el as control
  dim id as string
  dim content as string
  dim OUT(19) as string
  dim i as integer
  for i =1 to 18
    id = preFix+i.toString
    el = ref.element(id)
    content=el.tag.toString
    if content = "System.Object[]" then content=""
    OUT(i)=content
  next
  dim RESULT as string=join(OUT, vbNewLine)
  'add(Result)
  add(globFileSpec)
  ''dim fileSpec as string ="C:\yPara\scriptIDE\scriptPara\"+"myBookmarks01.txt"
  ZZ.filePutContents(globFileSpec, RESULT)
end sub






sub readBookmarkData()
  dim preFix as string ="lab_navLab_"
  dim el as control
  dim id as string
  dim content as string
  dim DATA(19) as string
  dim OUT(19) as string
 
  content=ZZ.fileGetContents(globFileSpec)
  'add(content)
  DATA=split(content, vbNewLine)
  dim ELEMENTS() as string
  dim item as string
  dim label as string
  dim i as integer
  dim max as integer 
  max=DATA.length
  for i =1 to 18
    id = preFix+i.toString
    el = ref.element(id)
    item=DATA(i)
    if trim(item)="" then continue for
    
    ELEMENTS=split(item,"<³³>")
    el.tag=item
    label=ELEMENTS(3)
    el.text=label
  next
end sub



sub test1_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test1()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub

sub test2_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test2()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test2()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub

sub test3_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = callScriptClassTestProc(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test3()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test3()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub

sub outMonitor_Clear_MouseClick(e)
  myTextArea.text=""
end sub

sub outMonitor_selectAll_MouseClick(e)
  dim lng as integer=myTextArea.text.length
  dim txt = container1.controls("txt_outMonitor")
  trace "myTextArea.text.length", lng
  ' myTextArea.Select(1,3)
  txt.selectAll()
  txt.focus()
  ' txt.text="aaaaaaaaaaa"
end sub

sub txt_status2_TextChanged(e)
  '...nix zu tun
end sub

sub txt_outMonitor_TextChanged(e)
  '...nix zu tun
end sub

sub txt_outMonitor_LostFocus(e)
  '...nix zu tun
end sub

sub txt_outMonitor_MouseUp(e)
  '...nix zu tun
end sub

sub txt_outMonitor_GotFocus(e)
  '...nix zu tun
end sub
sub txt_outMonitor_KeyDown(e)
  printLine 11, "", e.keyString
  if e.keyString="F3" then togglePopupMonitor()
  '...nix zu tun
end sub

sub togglePopupMonitor()
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_popUpMonitor.mainWin" )
  dim curState as boolean  = toolWindow.visible
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
   else
    toolWindow.hide()
  end if
end sub

Sub enumDocPanels_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  enumDocPanels()
End Sub

Sub appPicture_MouseClick(e)
  printLine 11, "MouseClick(e)", e.Sender.Name
  ' hide()
  'myImg.enable=false
  ''ref.element("pic_appPicture").selectet=true
End Sub

Sub appPicture_MouseDown(e)
  trace "appPicture_MouseDown", e.Sender.Name
  ' hide()
  'myImg.enable=false
  ''ref.element("pic_appPicture").selectet=true
End Sub

sub enumDocPanels()
   msgBox ("enumDocPanels")
   dim curDocPanel=zz.getActiveTab()
   ''msgBox (curDocPanel.parent.count)
end sub








'-->
'--> M O N I T O R
sub add(para as string)
  : dim txt as textbox = ref.element("txt_outMonitor")
  : if txt is nothing then exit sub
  : err.number=0
  
  dim oldContent=txt.text
  dim newContent=oldContent+para+vbNewLine
  txt.text=newContent
  dim lg=newContent.length
  txt.selectionStart=lg
  txt.ScrollToCaret()
end sub

sub clear()
  : dim txt as textbox = ref.element("txt_outMonitor")
  : if txt is nothing then exit sub
  : err.number=0
  txt.text=""
end sub

sub statustext_reset()
  : dim txt as textbox = ref.element("txt_status2")
  : if txt is nothing then exit sub
  : err.number=0
  txt.backColor=ColorTranslator.FromHtml("#D4D0C8")
  txt.text=""
end sub


sub statusText(para as string)
  : dim txt as textbox = ref.element("txt_status2")
  : if txt is nothing then exit sub
  if para.startsWith("ERR")then 
    txt.backColor=ColorTranslator.FromHtml("#f77")
  end if
  txt.text=para
end sub

sub errorText(para as string)
  : dim txt as textbox = ref.element("txt_errMes")
  : if txt is nothing then exit sub
  if para.startsWith("ERR")then 
    txt.backColor=ColorTranslator.FromHtml("#000")
    txt.foreColor=ColorTranslator.FromHtml("#DF7500")
  else
    txt.backColor=ColorTranslator.FromHtml("#D7D7D7")
    txt.foreColor=ColorTranslator.FromHtml("#FF0000")
  end if
  txt.text=para
end sub


sub setData(para as string)
  dim txt = ref.element("txt_outMonitor")
  txt.text=para
end sub

function getData() as string 
  dim txt = ref.element("txt_outMonitor")
  return txt.text
end function

