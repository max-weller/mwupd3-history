dim appIcon as string ="http://icons3.iconfinder.netdna-cdn.com/data/icons/humano2/72x72/apps/gnome-monitor.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icon24/emblem-noread.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icons/128player_playlist.png"
'dim appIcon as string ="http://es.teamwiki.net/docs/icons/Folder-Downloads.png"
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


'' es_testLabor.vb                          

#Para DebugMode intern
#Para SilentMode true

'-->
'--> G L O B

'' #reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic

#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib

shared winId as string ="es_testLabor.mainWin"
public winCaption as string = "Test-Labor"

public WithEvents FormRef As Form
public WithEvents ref As Object
dim WithEvents myTextArea as textbox
dim WithEvents myPicture as pictureBox      
public withEvents myWin2 as object
public WithEvents IGrid1 As Igrid
public WithEvents Timer1 As System.Windows.Forms.Timer

public WithEvents check1 as System.Windows.Forms.CheckBox
public WithEvents check2 as System.Windows.Forms.CheckBox
public WithEvents check3 as System.Windows.Forms.CheckBox

shared myClassRef as object
dim myImg as object

'' public shared toolBar As Object
public shared toolBar1 As Object
public shared toolBar2 As Object
public shared statBar As Object
public shared container1 As Object
public toolBarColor as string
private globStatusText1 as object
private globStatusText2 as object
private globLabFileName as object

dim globErrorCounter as integer = 0

public gStatus1 as object
public gStatus2 as object
public gSubNotFound as object
public gSubNotFoundClose as object
public gSpeed as object
public gLoopCounter as object
public gTimerLab as object
public gFileName as object
public gErrCounter as object



'--> W I N - A P I

Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short

Function isControl() As Boolean
  isControl = False
  If isKeyPressed(Keys.ControlKey) Then
    isControl = True
  End If
End Function


Function isShiftControl() As Boolean
  isShiftControl = False
  If isKeyPressed(Keys.ShiftKey) And isKeyPressed(Keys.ControlKey) Then
    isShiftControl = True
  End If
End Function


Function isKeyPressed(ByVal key As Keys) As Boolean
  isKeyPressed = False
  Dim stat As Short
  GetAsyncKeyState(key) 'puffer leeren
  stat = GetAsyncKeyState(key)
  'Debug.Print(key.ToString + vbTab + stat.ToString)
  If stat <> 0 Then
    isKeyPressed = True
  End If
End Function



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
  
  exit sub
  
  
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

dim id="ToolBar|##|tbScriptWin|##|es_webbrowser.tab"
dim winRef=zz.IDEHelper.GetTabByPersistString(id)
winRef.show

exit sub
  msgBox(   join(ZZ.IDEHelper.getToolBarList(),vbnewline))
  msgBox(   join(ZZ.IDEHelper.getTabList(),vbnewline))
  
  exit sub

  dim id2="ToolBar|##|tbScriptWin|##|es_dialogEditor.editorWin"
  dim toolWindow=ZZ.getDockPanelRef(id2)
  toolWindow.visible= not toolwindow.visible
  ' ToggleDockWindowFromId(id)


exit sub

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
  'ref = ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
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
  ''msgBox ("show(es_testLabor.mainWin)")
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
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_testLabor.mainWin")
  toolWindow.hide()
end sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  'ACHTUNG: UMSTELLEN auf removehandler fehlt noch
  onResizeControls
End Sub

: Sub onResizeControls()
  '... der kommt hier bis zu 30 mal vorbaei
   'trace "onResizeControls", formRef.Height
  'Igrid1.Width = formRef.Width - 10
  myTextArea.Height = container1.Height -0' - 110 '42
  myTextArea.width = container1.width -20' - 110 '42
End Sub


'''#####################################################



sub onInitialize()
    AddHandler TT.TraceWrite, AddressOf OnTrace
    AddHandler TT.PrintLineChanged, AddressOf OnPrintLine
end sub


sub onTerminate()
   stopTimer()
   RemoveHandler TT.TraceWrite, AddressOf OnTrace
   RemoveHandler TT.PrintLineChanged, AddressOf OnPrintLine

end sub


'-->
'--> A U T O S T A R T 
Sub AutoStart()
  myClassRef=me
  'onTerminate()
  'zz.traceClear()
  'zz.printLineReset()
  GetFormRef()
  dim el as object
  toolBarColor="#BFBFBF"
  toolBarColor="#65CE84"
With ref
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    container1  =.addPanel("container1"   , DockStyle.Fill)
    toolBar1     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)

    toolbar1.visible=true
    toolbar1.height=44

    container1.top=112
    toolBar1.resetControls  (10,5,1)
    statBar.resetControls  (10,5,1)
    container1.resetControls  (10,5,1)
   
    toolbar1.BackColor = ColorTranslator.FromHtml(toolBarColor)
    container1.BackColor = ColorTranslator.FromHtml("#fff")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.height=22
  end with





'--> ... toolBar1  
  With toolbar1
    .height=95' 88
    .backColor=ColorTranslator.FromHtml("#BFBFBF")
    .backColor=ColorTranslator.FromHtml("#222")
    .backColor=ColorTranslator.FromHtml("#243E56")
    .backColor=ColorTranslator.FromHtml("#aaa")
    .backColor=ColorTranslator.FromHtml("#253833")
     '' .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"

 
    '--> ...error-reset

    .insertX = 10: .insertY= 3: 
    dim icon as string="http://es.teamwiki.net/docs/icon24/emblem-noread.png"
    'el=.addButton("close" ,"",bgColor,Left+deltaLeft,Top+deltaTop,24 ,24 ,icon) ' ,flags,handler)
    el=.addButton("close" ,"","#253833",             ,             ,26 ,26 ,icon) ' ,flags,handler)
    el.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    el.FlatAppearance.BorderSize = 0
    el.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#FFD350")  '#4A7CDB
    el.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#5A8CeB")
     

    ''.insertX = 220: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66

    '' .insertX = 85: .insertY= 5: 
''     '.addLabel(ID, Text,                       bgColor, fgColor ,Left,Top,Width,Height
''      el=.addLabel ("errlabel"     , "ERR:"   , "#253833", "#bbb"        ,   , ,70   ,22):  el.textAlign=System.Drawing.ContentAlignment.MiddleRight' : el.foreColor=ColorTranslator.FromHtml("#bbb")
''      
 
 
 '--> ...error-counter
    .insertX = 38: .insertY= 7: 
    el=.addLabel ("ErrCounter"     , "speed"   , "#3B5951", "#6DFF50"        ,   , ,45,19   ): 
    gErrCounter=el 
    '' el= .addTextbox ("ErrCounter",   50,"Err:"    , 0,"" , , , )
    '' el= ref.element("lab_ErrCounter")
    'el.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
    'el.backColor=ColorTranslator.FromHtml("#f66")
    el.textAlign=System.Drawing.ContentAlignment.MiddleCenter
    el.backColor=ColorTranslator.FromHtml("#3B5951")
    el.foreColor=ColorTranslator.FromHtml("#bbb")
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderstyle=0
    el.autosize=false
    el.height=20
    el.text="OK"
    
    '...bild
    .insertX = 15  '70 '35 
    .insertY = 25 ' 11 ' 28 '-5
    ''myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/logMonitor.png" )
    ''myPicture=.addIcon("appPicture", "http://es.teamwiki.net/docs/icons/gnome-monitor.png" )
    myPicture=.addIcon("appPicture", "http://icons3.iconfinder.netdna-cdn.com/data/icons/humano2/72x72/apps/gnome-monitor.png" )
  

    check1 = New System.Windows.Forms.CheckBox
    check1.Location = New System.Drawing.Point(97, 7)
    check1.Size = New System.Drawing.Size(19, 19)
    check1.Text = ""
    'check1.UseVisualStyleBackColor = True
    check1.BackColor = ColorTranslator.FromHtml("#253833")
    check1.foreColor = ColorTranslator.FromHtml("#bbb")
    check1.visible=true
    toolbar1.Controls.Add(check1)

    
    'filename
    .insertX = 118: .insertY= 5: 
    el=.addLabel ("fileName"     , "...fileName"   , "#3B5951", "#6DFF50"        ,   , ,-2   ,20):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    globLabFileName=el
    gFileName = el
    el.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

 
    .insertX = 120: .insertY= 48: 
    el=.addLabel ("speed"     , "speed"   , "#3B5951", "#6DFF50"        ,   , ,75,17   ):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    'globLabFileName=el
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    'el.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.text="1211"

    .insertX = 200: .insertY= 48: 
    el=.addLabel ("loop"     , "speed"   , "#3B5951", "#6DFF50"        ,   , ,75,17   ):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    'globLabFileName=el
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    'el.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.text="10.000"

    .insertX = 280: .insertY= 48: 
    el=.addLabel ("time"     , "speed"   , "#3B5951", "#6DFF50"        ,   , ,75,17   ):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    'globLabFileName=el
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    'el.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.text="5:987"


 
    '--> ...fileNotFound
 
     .insertX = 95: .insertY= 28: 
    el=.addLabel ("gSubNotFoundClose"     , "?"   , "#FF7777", "#000"        ,   , ,15,17   ):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    gSubNotFoundClose=el
    el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    'el.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.text="x"
    el.visible=false

    .insertX = 120:.insertY = 28
    'el= .addTextbox ("status2",   -2 )
    el=.addLabel ("gSubNotFound"     , "FileNotFound"   , "#FF7777", "#000"        ,   , ,-2,17   ):  el.textAlign=System.Drawing.ContentAlignment.MiddleCenter' : el.foreColor=ColorTranslator.FromHtml("#bbb")
    gSubNotFound= el
    el.backColor=ColorTranslator.FromHtml("#D7D7D7")
    el.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderstyle=0
    el.autosize=false
    el.height=18
    el.bringToFront()
    el.text="y"
    el.visible=false



    check2 = New System.Windows.Forms.CheckBox
    check2.Location = New System.Drawing.Point(98, 70)
    check2.Size = New System.Drawing.Size(77, 19)
    check2.Text = " loopMode"
    'check2.UseVisualStyleBackColor = True
    check2.BackColor = ColorTranslator.FromHtml("#253833")
    check2.foreColor = ColorTranslator.FromHtml("#bbb")
    check2.visible=true
    toolbar1.Controls.Add(check2)

  check3 = New System.Windows.Forms.CheckBox
    check3.Location = New System.Drawing.Point(180, 70)
    check3.Size = New System.Drawing.Size(177, 19)
    check3.Text = "dump / autoClear / opt?"
    'check3.UseVisualStyleBackColor = True
    check3.BackColor = ColorTranslator.FromHtml("#253833")
    check3.foreColor = ColorTranslator.FromHtml("#bbb")
    check3.visible=true
    toolbar1.Controls.Add(check3)










    
    '--> ...statustexterrLine
     .insertX = 120: .insertY= 28: 
    el= .addTextbox ("errMes",   -2 )
     el= ref.element("txt_errMes")
     '' el.backColor=ColorTranslator.FromHtml("#CEF2E9")
     '' el.foreColor=ColorTranslator.FromHtml("#b00")
     el.backColor=ColorTranslator.FromHtml("#253833")
     el.foreColor=ColorTranslator.FromHtml("#bbb")
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.borderstyle=0
     el.text= "... der kommt wieder weg"


end with



'--> ...container    

 with container1
    .backColor=ColorTranslator.FromHtml("#BFBFBF")
    .backColor=ColorTranslator.FromHtml("#222")
    .backColor=ColorTranslator.FromHtml("#243E56")
    .backColor=ColorTranslator.FromHtml("#253833")

'--> ...textarea
    .insertX = 5 :.insertY = 0' 110
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

    .insertX = 10: .insertY= 5: 
     el= .addTextbox ("status",   -2)
     globStatusText1= ref.element("txt_status")
     globStatusText1.backColor=ColorTranslator.FromHtml("#D7D7D7")
     globStatusText1.backColor=ColorTranslator.FromHtml("#3B5951")
     globStatusText1.foreColor=ColorTranslator.FromHtml("#bbb")
     globStatusText1.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     globStatusText1.borderstyle=0

End With

  

    '--> 
    '--> ... müll
    '--> ... ...topbar 
     '.insertX = 220: el=.addButton ("outMonitor_selectAll"     , "SelectAll"       , "#B1B13B"): el.width=66

  onResizeControls()
  ''readBookmarkData()

  Timer1 = New System.Windows.Forms.Timer()
  startTimer()
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
  
  trace "=== ACHTUNG === ", "...testFehler"
  dim dummy as object
  dummy.makeError


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


sub onKey_ESCAPE()
  '' add("TREFFER: onKey_ESCAPE")
  dim sc=zz.scriptclass("es_codeList")
  if sc is nothing then exit sub
  with sc
    .historyNavigateDeltaIndex(-1)
  end with
end sub



sub onKey_F1()
  ''add("TREFFER: onKey_F1")
  dim sc=zz.scriptclass("es_codeList")
  if sc is nothing then exit sub
  with sc
    .historyNavigateDeltaIndex(+1)
  end with
end sub

sub onKey_F2()
  add("TREFFER: onKey_F2")
  '' dim state as boolean
  '' dim i as integer
  '' for i=1 to 33
  ''   state=isKeyPressed(27)
  ''   add(state.toString)
  ''   zz.idle(33)
  '' next
  '' 
  toggleDockPanel(8)'left

end sub

sub onKey_F3()
  add("TREFFER: onKey_F3")
  '' dim id as string ="ToolBar|##|tbScriptWin|##|es_internSuche.mainWin"
  '' ToggleDockWindowFromId(id)

  dim id as string ="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  
  ''ToggleDockWindowFromId(id)
  
  toggleTab(nothing)
end sub


Sub toggleTab(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  static state as boolean =false
  state= not state
  dim fileSpec="loc:/C:/yPara/scriptIDE4/scriptClass/es_kommentar.txt"
    dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
    if state=true then
      tab.hide
    else
      tab.show
    end if
End Sub


sub onKey_F4()
    add("TREFFER: onKey_F4")
    toggleTraceWindow()
end sub

sub onKey_RETURN()
    'add("TREFFER: onKey_RETURN")
end sub


sub onKey_AltGrShift
  static lastState as integer=1
  if lastState=1 then
    lastState=0
  else
    lastState=1
  end if
  add("TREFFER: onKey_AltGrShift")
  toggleDockPanel(8,lastState)'left
  toggleDockPanel(9,lastState)'left
  toggleDockPanel(10,lastState)'left
  toggleDockPanel(7,lastState)'left
end sub


Function toggleDockPanel(intPanel as integer, optional forceState as integer=-1)
  static state as boolean = true: 'state = not state
  state = not state
  if forceState=0 then state=false
  if forceState=1 then state=true
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  ' mainWin.xxxxxxxxxx    '''FEHLER-TEST
  dim Panel as object = mainWin.DockPanel1
  dim p as object: dim i as integer
  For Each p In Panel.Panes
     For i = 0 To p.Contents.Count - 1
       if p.DockState = intPanel  then 
         'p.parent.visible = not p.parent.visible
         p.parent.visible = state
         exit Function
       end if
     Next
  Next
End Function



'-->

public sub toggleTraceWindow()
  dim id as string ="ToolBar|##|tbScriptWin|##|es_iconBar_B.mainWin"
  ToggleDockWindowFromId(id)
end sub

function ToggleDockWindowFromId(id as string, optional forceState as integer=-1)as boolean
  dim toolWindow=ZZ.getDockPanelRef(id)
  
  '' exit function
  if toolWindow is nothing then exit function
  dim curState as boolean  = toolWindow.visible
  
  ''msgBox(curState)
  dim newState as boolean= not curState
  if forceState=0 then newState=false
  if forceState=1 then curState=true
  dim id2 as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  dim id3 as string="ToolBar|##|tbScriptWin|##|es_iconBar_B.mainWin"
  Add(id)
  Add(id2)
  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    if id=id2 or id=id3 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    if id=id2 or id=id3 then toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function




sub txt_ErrCounter_MouseUp(e)
  errorCounterReset()
End Sub



Sub Check1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Check1.CheckedChanged
  msgBox("CheckBox1")
End Sub
  
sub onButtonEvent(e)
  errorCounterReset()
  zz.traceClear()
  zz.printLineReset()
end sub


sub onLabelEvent(e)
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  Add("======================== onLabelEvent")
  Add("____SenderName: " + name(1))
  Add("___MouseButton: " + e.MouseButton)
  ''Add("keystring ===>" + e.keyString)
  '' dim tag as string = e.sender.tag.toString
  '' dim tagDATA()= Split(tag, "<³³>")
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
    if el is nothing then continue for
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
  'myTextArea.text=""
   clear()
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
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_testLabor.mainWin" )
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
'--> startTimer
sub startTimer()
  Timer1.Interval = 555
  Timer1.Enabled = True
  statustext("timer gestartet")
end sub


:sub stopTimer()
  on error resume next
  Timer1.Enabled = false
  statustext("timer ist ausgeschaltet")
  Add("Timer gestopt...")
  err.number = 0
end sub


Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  ': stopTimer()
  : onTimerEvent()
End Sub


'--> onTimerEvent
sub onTimerEvent()
  'add("onTimerEvent")
  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  dim name as string=My.Computer.FileSystem.GetName(ActiveTabFilespec)
  'statustext(name)
  ref.element("lab_fileName").text=name
  
end sub






'-->
'--> T R A C E - C L A S S


  Sub OnTrace(ByVal para1 As String, Optional ByVal para2 As String = "", Optional ByVal typ As String = "info", Optional ByVal pcodeLink As String = "")
  'msgBox(typ)
  static skipMe as boolean=false
  if skipMe=true then exit sub
  skipMe=true
  if typ.toLower.startsWith("err") then
      'msgBox(para2)
      errorCounterAdd()
      zz.doevents
  end if
  skipMe=false
  exit sub

  Dim highlightLine As Boolean = False
    '' If typ.EndsWith(",highlight") Then
    ''   highlightLine = True
    ''   typ = Replace(typ, ",highlight", "")
    '' End If

    '' If isIDEMode And ScriptHost.Instance.InformationWindowVisible Then
    ''   Dim ir = MAIN.IGrid2.Rows.Add
    ''   ir.Cells(0).Value = typ
    ''   If MAIN.iml_TraceTypes.Images.ContainsKey(typ) Then ir.Cells(0).ImageIndex = MAIN.iml_TraceTypes.Images.IndexOfKey(typ)
    ''   ir.Cells(1).Value = para1
    ''   ir.Cells(2).Value = para2
    ''   ir.Tag = pcodeLink
    '' 
    ''   MAIN.autoScrollFlag = True
    ''   'If MAIN.checkTraceAutoscroll.Checked Then MAIN.IGrid2.SetCurRow(ir.Index)
    ''   MAIN.autoScrollFlag = True
    ''   'If MAIN.checkTraceAutoscroll.Checked Then MAIN.IGrid2.SetCurRow(ir.Index)
    ''   If highlightLine Then
    ''     MAIN.TextBox2.BackColor = Color.Gold
    ''     MAIN.timer_fadeOutErrBox.Start()
    ''   End If
    '' End If
  End Sub

  Sub OnPrintLine(ByVal index As Integer, ByVal title As String, ByVal data As String, Optional ByVal pcodeLink As String = "")
    '' With ScriptHost.Instance
    ''   If isIDEMode And .PrintLineWndVisible Then
    ''     .getPrintlineWndRef().setPrintLine(index, data, title)
    ''   End If
    '' End With
  End Sub


'-->
'--> M O N I T O R
: sub errorCounterAdd(optional item as string="")
  on error resume next
  globErrorCounter=globErrorCounter+1
  dim lab as object=gErrCounter
  lab.text=globErrorCounter.toString
  lab.BackColor=ColorTranslator.FromHtml("#f66")
  lab.foreColor=ColorTranslator.FromHtml("#000")
  err.number=0
end sub


: sub errorCounterReset()
   on error resume next
   globErrorCounter=0
   dim lab as object=gErrCounter
   lab.text="OK"
   lab.backColor=ColorTranslator.FromHtml("#3B5951")
   lab.foreColor=ColorTranslator.FromHtml("#bbb")
   err.number=0
end sub


: sub cmd(para as string)
  on error resume next
  trace "cmd(para as string)", para
  ''add("CMD: "+para)
  if para.endsWith("_F1") then onKey_F1()
  if para.endsWith("_F2") then onKey_F2()
  if para.endsWith("_F3") then onKey_F3()
  if para.endsWith("_F4") then onKey_F4()
  if para.endsWith("_ESCAPE") then onKey_ESCAPE()
  if para.endsWith("_RETURN") then onKey_RETURN()
  if para.endsWith("_AltGrShift") then onKey_AltGrShift()
end sub


: sub add(para as string)
  on error resume next
  dim txt as textbox = ref.element("txt_outMonitor")
  if txt is nothing then exit sub
  err.number=0
  
  dim oldContent=txt.text
  dim newContent=oldContent+para+vbNewLine
  txt.text=newContent
  dim lg=newContent.length
  txt.selectionStart=lg
  txt.ScrollToCaret()
end sub

: sub clear()
  on error resume next
  ' msgbox(ref.toString)
  dim txt as textbox = ref.element("txt_outMonitor")
  add(": sub clear() !!!")
  if txt is nothing then exit sub
  txt.text=""
  err.number=0
end sub

: sub statustext_reset()
  on error resume next
  gSubNotFound.backColor=ColorTranslator.FromHtml("#3B5951")
  gSubNotFound.text=""

  gSubNotFound.visible=false 
  gSubNotFoundClose.visible=false  
  err.number=0

end sub


: sub statusText(para as string)
  on error resume next
  if para.startsWith("ERR")then 
    gSubNotFound.backColor=ColorTranslator.FromHtml("#f77")
    gSubNotFound.text=para
    gSubNotFound.visible=true 
    gSubNotFoundClose.visible=true  
  end if
  'gSubNotFound.text=para
  globLabFileName.text=para


  zz.doevents
  err.number=0
end sub

: sub errorText(para as string)
  on error resume next
  dim txt as textbox = ref.element("txt_errMes")
  if txt is nothing then exit sub
  if para.startsWith("ERR")then 
    'txt.backColor=ColorTranslator.FromHtml("#000")
    txt.foreColor=ColorTranslator.FromHtml("#DF7500")
  else
    'txt.backColor=ColorTranslator.FromHtml("#D7D7D7")
    txt.foreColor=ColorTranslator.FromHtml("#FF0000")
  end if
  txt.text=para
  err.number=0
end sub


: sub setData(para as string)
  on error resume next
  dim txt = ref.element("txt_outMonitor")
  txt.text=para
  err.number=0
end sub

: function getData() as string 
  on error resume next
  dim txt = ref.element("txt_outMonitor")
  err.number=0
  return txt.text
end function

