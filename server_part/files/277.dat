'' ...zwischengelagert



'--> es_bookmarks2                          

#Para DebugMode intern
#Para SilentMode true

'' #reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic

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

#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib

public WithEvents FormRef As Form
public WithEvents ref As Object


public shared toolBar As Object
public shared statBar As Object
public shared container1 As Object
public toolBarColor as string

shared winId as string ="es_bookmarks2.mainWin"
public winCaption as string = "myBookmarks2"

dim WithEvents myTextArea as textbox
dim WithEvents myPicture as pictureBox      
public withEvents myWin2 as object
dim myImg as object
public WithEvents IGrid1 As Igrid

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

public globMonitorRef as object
dim globFileSpec as string ="C:\yPara\scriptIDE\scriptPara\"+"myBookmarks01.txt"


'-->
'--> T E S T 

sub test1()
  printLine 3, "aaa","bbbbb"
  monitorAdd("---------------------------------------")
  ZZ.IDEHelper.Exec("Window.Reflection", "zz","")
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
   monitorAdd(ZZ.getActiveRTF.RTF.Lines.Current.Text)
end sub

sub test2()
  msgBox("bbbbbbbbbbbb")
  dim mainWin as object: mainWin =ZZ.IDEHelper.MainWindow
  trace mainWin.ActiveMdiChild,toString
  monitorAdd(ZZ.getActiveRTF.RTF.toString)
  monitorAdd(ZZ.getActiveRTF.RTF.Lines.Current.number)
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim lineNumber as integer = codeEitor.Lines.Current.number
  dim lineContent as string = codeEitor.Lines.Current.text
  monitorAdd(lineNumber.toString+" --> "+lineContent+"<--")
  monitorAdd(lineNumber.toString+" --> "+trim(lineContent)+"<--")
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
    monitorAdd(labelText1 )
    monitorAdd(labelText2 )
    monitorAdd(MAIN.chkSticky.Checked)
    ''MAIN.Text = prefix + "scriptIDE2" + filename

 '' End Sub

exit sub





  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
monitorAdd(activeTabType)
monitorAdd(ActiveTabFilespec)
monitorAdd(activeTab.toString)
monitorAdd(activeTab.parent.toString)
monitorAdd(activeTab.parent.toString)
monitorAdd(activeTab.parent.parent.toString)
exit sub
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim selection as string =codeEitor.selection.text
  monitorAdd(selection)
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
  'ref = ZZ.IDEHelper.CreateDockWindow(winID, 1)
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  formRef = ref.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = false
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


Sub Form_Resize(sender as Object, e as EventArgs)
  onResizeControls
End Sub

Sub onResizeControls()
   trace "onResizeControls", formRef.Height
  'Igrid1.Width = formRef.Width - 10
  container1.height=ref.height
  container1.width=ref.width
  monitorAdd("container1.controls.count")
  monitorAdd(container1.controls.count)
  monitorAdd("container1.controls.count")
  'myTextArea.Height = container1.Height - 60 '42
  dim el as object
  for each el in container1.controls
    if el.name.startsWith("lab_navLab") then
       el.width=container1.width - 60 
    end if
  next
End Sub


sub onInitialize
   '
end sub


sub onTerminate
  trace "onTerminate","OK ??? "
  globRemoveHandler()
end sub

'''#####################################################


sub globAddHandler()
  addHandler FormRef.Resize, AddressOf  Form_Resize
end sub

sub globRemoveHandler()
  removeHandler FormRef.Resize, AddressOf  Form_Resize
  trace " ==> globRemoveHandler","OK ?"
end sub







'-->
'--> A U T O S T A R T 
Sub AutoStart()
  'onTerminate() 
  'zz.traceClear()
  'zz.printLineReset()
  GetFormRef()
  setMonitorRef()
  
  globAddHandler()

  dim el as object
  toolBarColor="#BFBFBF"
With ref
    .BackColor = ColorTranslator.FromHtml("#222")

    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"

    ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
    container1  =.addPanel("container1"   , DockStyle.Fill)
    container1.visible=false
    toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)
 
    toolBar.resetControls  (10,5,1)
    statBar.resetControls  (10,5,1)
    container1.resetControls  (10,5,1)
   
    container1.BackColor = ColorTranslator.FromHtml("#fff")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
  end with

'--> DATA: toolBar    
  With toolbar
    toolbar.visible=true
    .height=22
    toolbar.BackColor = ColorTranslator.FromHtml("#333")
     .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
 
    '.BR  '--------------------------------------------------
    ''.insertY=5: .insertX=5
    .addIcon("expand_right"  , "http://es.teamwiki.net/docs/icons/arrow_right.png" ,7 ,7 ,16,16  )
    .addIcon("expand_left"   , "http://es.teamwiki.net/docs/icons/arrow_left.png"  ,7 ,7 ,16,16  )
  end with



'--> ...container    

 with container1
     ''container1.top=112
    .backColor=ColorTranslator.FromHtml("#BFBFBF")
    .backColor=ColorTranslator.FromHtml("#222")
     '' .OffsetX = 45 :. insertY = 11
''     '' el=.addButton ("cmdExit"     , "Fenster schließen"  , "#E0E0E0"): el.width=113
''                      'el=.addButton ("outMonitor_toClipboard"   , "--> Clipboard"   , "#E0E0E0"): el.width=77
''      .insertY = 4  
''      .insertX = 010: el=.addButton ("hideMe"    , "- X -"          , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''      .insertX = 055: el=.addButton ("test1"    , "test 1"          , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''      .insertX = 100: el=.addButton ("test2"    , "test 2"          , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''      .insertX = 145: el=.addButton ("test3"    , "test 3"          , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
'' 
''      .insertX = 205 :  el=.addButton ("outMonitor_Clear"         , "Clear"           , "#FB7929"): el.width=66
''      .insertX = 271: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66
''  
''      .insertY = 35
''      .insertX = 143:.OffsetX = 143
''                 el=.addButton ("saveMe"    , "save"     , "#222"): el.width=45:  el.foreColor=ColorTranslator.FromHtml("#eee")
''       .BR(24):  el=.addButton ("readMe"    , "read"     , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''       .BR(24):  el=.addButton ("testy2"    , "r- C -"   , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''       .BR(24):  el=.addButton ("testy4"    , "- D -"    , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''       .BR(24):  el=.addButton ("toggleWeb"    , "W=E=B"    , "#222"): el.width=45: el.foreColor=ColorTranslator.FromHtml("#eee")
''  
''     '...bild
''     .insertX = 15  '70 '35 
''     .insertY = 35 ' 11 ' 28 '-5
''      myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/logMonitor.png" )
''    .BR(130) '--------------------------------------------------
    ''.OffsetX = 5 : .insertY = 5
    
    '--> ...Labels
    '.AutoScroll = False
    createLabels(container1,50)
    .AutoScroll = true   
     container1.visible=true

     el=.addLabel ("highlighter_0"     , ""  , "#0f0", "#ddd") :el.autoSize=false:  el.width=180: el.height=17: el.visible=false


'' .BR (25)'--------------------------------------------------
    '' el=.addButton ("test8"     , "Suchen Ruckwärts"  , "#E0E0E0"): el.width=175
    '' .BR (33)'--------------------------------------------------
    '' el=.addButton ("enumDocPanels"     , "enumDocPanels"  , "#E0E0E0"): el.width=175
    '' .BR (33)'--------------------------------------------------
    '' 
    '' el =.addButton ("test5"     , "--> Clipboard"  , "#E0E0E0",,,,,): el.width=175
    '' .BR(23) '--------------------------------------------------
    '' el=.addButton ("test6"     , "Save As..."  , "#E0E0E0"): el.width=175
    '' .BR(30) '--------------------------------------------------
    '' .OffsetX = 10 '::.insertY = 160
    ''  el = .addTextbox ("funcName" ,  173        , "  Code"+vbNewLine+""    , 0,"#bbb" , , , )
    '' 'ref.element("txt_codeTextarea").WordWrap=false
    '' 'ref.element("txt_codeTextarea").text=code
    '' 
    '' 'ref.element("txt_funcName").Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    '' ref.element("txt_funcName").backColor=ColorTranslator.FromHtml("#fff")
    '' ref.element("txt_funcName").foreColor=ColorTranslator.FromHtml("#000")
    '' ref.element("txt_funcName").borderStyle=0
    '' 
  
    '--> ...topbar ========================================================
     '.insertX = 220: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66

     'errLine
    .insertX = 205:.insertY = 40 
     el= .addTextbox ("errMes",   -2 )
     el= ref.element("txt_errMes")
     el.backColor=ColorTranslator.FromHtml("#D7D7D7")
     el.foreColor=ColorTranslator.FromHtml("#b00")
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0
     el.text= "lastErr"
     el.visible=false

    '--> ...textarea
    '' .insertX = 200 :.insertY = 60
    '' .addTextbox ("outMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=260")
    ''    myTextArea=ref.element("txt_outMonitor")
    ''    myTextArea.WordWrap=false
    ''    myTextArea.scrollbars=System.Windows.Forms.ScrollBars.Vertical
    ''    myTextArea.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    '' 
    ''    '--> myTextArea.text=getMiniHelpText()
    '' '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
    '' 
    '' 
   ref.visible=true
   ref.show
   show
  end with
  
  
'--> ...statBar   
  with statbar
   .visible=false
      .backColor=ColorTranslator.FromHtml("#939393")
      .backColor=ColorTranslator.FromHtml("#222")
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


:sub createLabels(container as object,count as integer)
  on error resume next
  ref.SuspendLayout()
  dim deltaY as integer = 18
  dim x as integer=5
  dim y as integer=5 - deltaY
  dim i as integer
  
  for i=1 to count 
    y=y+deltaY
    createLabel(container, x, y, i.toString)
  next

end sub

:sub createLabel(container as object,x as integer,y as integer, index as integer)
  on error resume next
  dim index2 as string=index.toString
  dim bgColor1="#222" '"#007"    
  dim bgColor2="#0a0"    
  dim bgColor3="#333"

  dim fColor1="#aaa"
  dim fColor2="#ccc"
  dim fColor3="#ccc"
  with container
    if index=1 or index mod 10 = 0 then
      bgColor3="#AD9C01"
      fColor3="#000"
    end if
    
    dim el as object
 ''   el=.addLabel ("test444a"     , "0  ERR"      ,"#696E73", ,5   ,24 ,42   ,16): el.textAlign=System.Drawing.ContentAlignment.MiddleCenter: el.foreColor=ColorTranslator.FromHtml("#bbb")
    el=.addLabel ("navNr_"+index2+" "     , index  ,bgColor1, fColor1, x ,y ,22  ,16) 
    el.textAlign=System.Drawing.ContentAlignment.MiddleRight
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

    el=.addLabel ("navColor_"+index2+" "     , ""  ,bgColor2, fColor2, x+25 ,y ,4  ,16) 
    el.textAlign=System.Drawing.ContentAlignment.MiddleRight
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))



    el=.addLabel ("navLab_"+index2     , ""   ,bgColor3, fColor3, x+35 ,y ,150  ,16) 
    el.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    '.BR (19)'--------------------------------------------------
   end with
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
  monitorAdd("==============================================")
  monitorAdd("Sender.Name ===>" + e.MouseButton)
  ''Add("keystring ===>" + e.keyString)
  '' dim tag as string = e.sender.tag.toString
  '' dim tagDATA()= Split(tag, "<³³>")
  monitorAdd(name(1))
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
  monitorAdd(max.toString)
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
  monitorAdd(RESULT+"<--(endOfArray)")
  
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


'--> ...
 
sub setNewCodeLink(e)
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim lineContent as string = trim(codeEitor.Lines.Current.text)
  lineContent=replace(lineContent,vbNewLine,"")
  lineContent=replace(lineContent,"sub","",,,1)
  lineContent=replace(lineContent,"function","",,,1)
  lineContent=replace(lineContent,"private","",,,1)
  lineContent=replace(lineContent,"public","",,,1)
  lineContent=trim(lineContent)
  'lineContent=replace(lineContent," ","_")
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
  monitorAdd(lineContent)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim index as string=name(2)
  dim id ="lab_navLab_"+index+"<--"
  monitorAdd("id: "+id)
  dim label=ref.element("lab_navLab_"+index)
  label.text=lineContent
  label.tag=newTag
  monitorAdd(e.sender.text)
  saveBookmarkData()
end sub

sub navCodeLink(e)
  ' 
  dim tag as string = e.sender.tag.toString
  monitorAdd("tag: "+tag+"<--")
  
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
  monitorAdd("navigiere zu: " +lineNumber.toString)
  dim codeEitor = ZZ.getActiveRTF.RTF
  'dim lineContent as string = codeEitor.Lines.Current.text
  'codeEitor.Lines.Current.number=lineNumber
  codeEitor.goTo.line(lineNumber+50)
  codeEitor.goTo.line(lineNumber-10)
  codeEitor.goTo.line(lineNumber)
  codeEitor.focus()
  
end sub



sub saveBookmarkData()
  dim preFix as string ="lab_navLab_"
  dim el as control
  dim id as string
  dim content as string
  dim OUT(51) as string
  dim i as integer
  for i =1 to 50
    id = preFix+i.toString
    el = ref.element(id)
    content=el.tag.toString
    if content = "System.Object[]" then content=""
    OUT(i)=content
  next
  dim RESULT as string=join(OUT, vbNewLine)
  'add(Result)
  monitorAdd(globFileSpec)
  ''dim fileSpec as string ="C:\yPara\scriptIDE\scriptPara\"+"myBookmarks01.txt"
  ZZ.filePutContents(globFileSpec, RESULT)
end sub






sub readBookmarkData()
  dim preFix as string ="lab_navLab_"
  dim el as control
  dim id as string
  dim content as string
  dim DATA() as string
  'dim OUT(19) as string
 
  content=ZZ.fileGetContents(globFileSpec)
  'add(content)
  DATA=split(content, vbNewLine)
  dim ELEMENTS() as string
  dim item as string
  dim label as string
  dim i as integer
  dim max as integer 
  max=DATA.length
  for i =1 to 50
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


'--> ... 
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


'-->
'--> outMONITOR

 sub clearAll()
   on error resume next
   monitorClear()
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
   globMonitorRef.statustext(para1+"<--")
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




