
' es_iconBar_B

'--> 
'--> C O N F I G - G L O B A L
#Para DebugMode intern
#Para SilentMode true

shared winId as string ="es_iconBar_B.mainWin"
public winTitle="IconBar-B"

shared winId2 as string ="es_iconBar_B.dialogWin"
public winTitle2="IconBar-Dialog"


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

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer

  '' <DllImport("user32.dll")> _
  '' Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
  '' End Function
  

public WithEvents FormRef As Form
public WithEvents ref As Object

public WithEvents FormRef2 As Form
public WithEvents ref2 As Object


public withEvents myWin2 as object
dim myImg as object
public WithEvents IGrid1 As Igrid

public shared toolBar As Object
public shared statBar As Object
public shared container1 As Object
public toolBarColor as string

public iconList() as string
public globMonitorRef as object
'''#####################################################


'-->
'--> TEST

: sub test1()

  showDialog()
  
exit sub

   on error resume next
  'msgBox("OK - 1")
  static counter as integer
  counter=counter+1
  'statustext(counter.toString)
  'monitorAdd("es_iconBar.mainWin","...ich bin NACHRICHT")
  'printLine 11, "test1()", counter
  'trace  "test1()", counter
  'zz.printLineReset()
  'listDocumentPane()
end sub



:sub test2()
  static counter as integer
  counter=counter+1
  'trace "test2:",  counter
  for i as integer=1 to 50
    printLine i,"test"+i.tostring("000"), i.tostring("000") & " " & counter
    'statustext(counter.toString)
  next
exit sub

  dim toolWindow=ZZ.getDockPanelRef("Addin|##|siacodecompiler|##|SHPrintline" )
  'dim curState as boolean  = toolWindow.visible
  'if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.left=1000
    toolWindow.parent.top=0
    toolWindow.parent.parent.left=1000
    toolWindow.parent.parent.top=0
   'else
    toolWindow.hide()
  'end if
exit sub


  monitorAdd("111111")
  monitorAdd("222222222222")
  monitorAdd("3333333333333333333333")
  'listDocumentPane(7, true)
  'sendMesTest2()
end sub





sub test3()
  listDocumentPane(7, false)
  listDocumentPane(8, false)
  listDocumentPane(9, false)
  listDocumentPane(10, false)

  'sendMesTest3()
end sub


'-->

Sub GetFormRef()
  onTerminate()
  'If ref IsNot Nothing Then Exit Sub
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)
  formRef = ref.form
  formRef.text=winTitle
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
End Sub

Sub GetFormRef2()
  ''onTerminate()
  'If ref IsNot Nothing Then Exit Sub
  ref2 = ZZ.CreateWindow(winID2, DWndFlags.DockWindow, 1)
  formRef2 = ref2.form
  formRef2.text=winTitle2
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
End Sub


sub onTerminate()
  'globMonitorRef = nothing
end sub



'-->
'--> A U T O S T A R T

 Sub AutoStart()
  GetFormRef()
  setMonitorRef()
  'myWin2 = ZZ.scriptClass(xxxx)
  'myWin2.setParent(me)
  'zz.traceClear()
  zz.printLineReset()
  
  with ref
    ''  ref.SuspendLayout()
    .resetControls ()
    '.activateEvents = "|IconMouseDown|    |TextboxKeyDown|"

    ''toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=75)
    container1  =.addPanel("container1"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)
    toolbar.visible=true
    'container1.top=112 
    toolBar.resetControls  (10,5,1)
    toolbar.height=32
    ''toolbar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    toolbar.BackColor = ColorTranslator.FromHtml("#fff")
    toolbar.visible=false
 
    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml("#3B5951")
    statBar.visible=false

    container1.resetControls  (10,5,1)
    'container1.BackColor = ColorTranslator.FromHtml("#4579D7")
    container1.BackColor = ColorTranslator.FromHtml("#3B5951")
  end with

  'With ref
  with container1


  end with
  
   
  ''container1.SuspendLayout()

  iniIconList()
  ''insertIconsFromIconList(toolBar)

  insertIconsFromIconList(container1)
  container1.ResumeLayout(False)
  container1.PerformLayout()
  
  ref.ResumeLayout(False)
  ref.PerformLayout()
  

end sub


Sub AutoStart2()
  GetFormRef2()
  setMonitorRef()
  'myWin2 = ZZ.scriptClass(xxxx)
  'myWin2.setParent(me)
  'zz.traceClear()
  zz.printLineReset()
  
  with ref2
    ' ref.SuspendLayout()
    .resetControls ()
  end with
end Sub


:Function listDocumentPane(Pane, showState) As object 'DockPane
  dim mainWin as object
  mainWin = ZZ.IDEHelper.MainWindow
  dim i as integer
  dim isMdiChild as boolean
  dim tag as object
  'dim DockState as object
  dim DockContent as WeifenLuo.WinFormsUI.Docking.DockContent

  dim panel1 as object = mainWin.DockPanel1
  dim p as object
  dim dockHandler as object
  For Each p In panel1.Panes
    For i = 0 To p.Contents.Count - 1
      isMdiChild=false
      dockHandler = p.Contents(i).DockHandler
      if p.Contents(i) Is mainWin.ActiveMdiChild Then isMdiChild=true
      ''if p.DockState = Pane or p.DockState = 7 or p.DockState = 8 or p.DockState = 9 or p.DockState = 10 then 
     if p.DockState = Pane  then 
        'p.show()
        'p.visible=true
       
        'p.parent.visible= not showState
        p.parent.visible= not p.parent.visible
        exit function
      end if
      
      ''trace p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText
      ''monitorAdd(p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText)
     
     
     ''tag=DirectCast(p.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      ''trace p.DockState,  isMdiChild.toString+"  "+p.Contents(i).DockHandler.TabText
      'if dockHandler.isFloat = true  then 
      'end if
      
      'tbOpenedFiles.IGrid1.Cells(i, 0).Value = doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'If doc.Contents(i) Is MAIN.ActiveMdiChild Then tbOpenedFiles.IGrid1.SetCurRow(i)
    Next
    '' If p.DockState = 7 'DockState.Document Then
    ''    Return p
    ''  End If
  Next
End Function





'-->
'--> E V E N T S
sub txt_dlg_TextChanged(e)
  monitorAdd("================txt_dlg_TextChanged")
end sub





Sub onTextBoxEvent(e)
  setMonitorRef()
  statustext("OK")
  errorText("")

  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action as string =name(1)
  dim index as string =name(2)
  dim value as string = e.Sender.text
  printLine 11, "Inhalt"        ,e.Sender.text+"<--"
  printLine 3, "onTextBoxEvent"  ,e.Sender.Name
  printLine 4, "eventName"       ,e.eventName
  printLine 5, "index"           ,index
  printLine 6, "action"          ,action
  ' MonitorAdd("eventType: " +   e.eventName)
  '' if action="dlg" and e.eventName = "KeyDown" then
  ''   syncIGrid(e, index, value)
  '' end if
 
  if action="dlg" and e.eventName = "GotFocus" then
    refreshHelpLabels(e, index, value)
  end if

  '' MonitorAdd("======================== onLabelEvent")
  ''MonitorAdd("SenderFullName: " + e.Sender.Name)
  '' MonitorAdd("___MouseButton: " + e.MouseButton)
  '' MonitorAdd("________Action: " + action)
end sub




sub lab_editIconBar_MouseClick(e)
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  monitorAdd("================lab_editIconBar_MouseClick")
  monitorAdd("____SenderName: " + name(1))
  monitorAdd("___MouseButton: " + e.MouseButton)
  showDialog()
end sub

Sub onButtonEvent(e)
  setMonitorRef()
  monitorClear
  trace "onButtonEvent", "onButtonEvent111"
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim tag as string = e.sender.tag
  dim tagDATA()= Split(tag, "<³³>")
  dim action as string =name(1)
  monitorAdd("==============================================")
  monitorAdd("Sender.Name: " + e.Sender.Name+"")
  monitorAdd("      action: " +action+"")
  monitorAdd(" MouseButton: " +e.MouseButton+"")
  'monitorAdd("   ClassName: " +e.ClassName+"")
  'monitorAdd(" ControlType: " +e.ControlType+"")
  'monitorAdd("      MouseX: " +e.MouseX.toString+"")
  monitorAdd("==============================================")

 
  Select Case action.toLower
    '' case "togglepane"      : onTogglePane         (e, tagDATA)
    '' case "url"             : onNavigateUrl        (e, tagDATA)
    '' case "cmdoptions"      : onCmdOptions         (e, tagDATA)
    '' case "cmdopenexplorer" : onCmdOpenExplorer    (e, tagDATA)
    '' 
    Case "toggle"             : onToggleDockWindow   (e, tagDATA)
    Case "saverun"            : onSaveRun            (e, tagDATA)
    Case "starttest"          : startTest            (e, tagDATA)
    Case "exec"               : onExecute            (e, tagDATA)
    Case "exe"                : onExe                (e, tagDATA)
    Case "callbyname"         : cmdCallByName        (e, tagDATA)
    '' Case "cmdtoggleweb"    : oncmdToggleWeb       (e, tagDATA)
    '' Case "exe"             : onExe                 (e, tagDATA)
    '' case "mnuexe"          : onExe                 (e, tagDATA)
    '' Case "nav"             : onNavigateUrl         (e, tagDATA)
    '' Case "toggleimg"       : onToggleImg           (e, tagDATA)
    '' Case "checkbox"        : onToggleImg           (e, tagDATA)
    '' case "createIconList"  : onCreateIconList      (e, tagDATA)
    '' 'Case "expand"         : onToggleDockWindow    (e, tagDATA)
    '' 'Case "expand"         : toogleToolbarItems    (e, tagDATA)
    '' 'Case "switch"         : onSwitch              (e, tagDATA)
    '' 
    '' Case "cmdtoolbarman" :onCmdToolbarMan          (e, tagDATA)
    '' Case "cmdiconman"    :onCmdIconMan             (e, tagDATA)
    '' Case "item"            '... nichts weiter machen
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
  'dim el =  e.sender
  'onTimerEvent()
  'checkDisplayState(el)
End Sub



sub onToggleDockWindow(e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  monitorAdd("tagPara",tagPara)
  if tagPara = "" then exit sub

  dim id = (trim(tagPara))
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
  monitorAdd(id)
  monitorAdd(id2)
  
  'newState=true
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    'if id=id2 then toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    'if id=id2 then  toolWindow.parent.parent.visible =false
  end if
  
  return newState
end function


sub onExe(e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.shellExecute(tagPara) : Err.Clear
end sub



sub onExecute(e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  : ZZ.IDEHelper.Exec(tagPara, "",""): Err.Clear
end sub




sub onSaveRun (e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  'monitorClear()
  
  monitorAdd("onSaveRun")
  monitorAdd("------- SAVING")
  ZZ.IDEHelper.Exec("File.Save", "","")
  : monitorAdd("------- SAVED"): err.number=0
  zz.doevents
  : ZZ.IDEHelper.Exec("Debug.Run", "","") : err.number=0
  monitorAdd("----------- READY")
  zz.doevents
  ''ZZ.shellExecute(tagPara) : Err.Clear
  statustext("OK")
  'ref.ResumeLayout(False)
  'ref.PerformLayout()
end sub


:sub startTest (e, tagDATA)
  on error resume next
  dim timerStart1 = GetTime()
  dim tagPara as string = trim(tagDATA(1))
  'monitorClear()
  statustext_reset()
  monitorAdd("startTest")
  
  monitorAdd(tagPara)
  monitorAdd("------- START")
  'ZZ.IDEHelper.Exec("File.Save", "","")
  monitorAdd("------- SAVED")
  zz.doevents

  ZZ.IDEHelper.Exec("Debug.Run", "","")
  monitorAdd("----------- COMPILED")
  zz.doevents
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  
  'scriptClass.autostart
  scriptClass.setMonitorRef()  '... falls autoStart noch nicht aufgerufen wurde
  
  
  ' ### warum der hier einen Fehler macht ist mir nicht klar !!!
  : dim funcName as string=tagPara : err.number=0
  trace "funcName", funcName
  dim i as integer
  dim timerStart2 = GetTime()
  dim deltaTime as integer
  
  'err.number=1111
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
  Dim myType As Type = scriptClass.GetType()
  :Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
  if myMethod.toString = "" then
    dim errMes="ERR: Sub '"+funcName+"' ...nicht gefunden" 
    statustext(errMes)
    exit sub
  end if
  
  monitorAdd("procName: "+myMethod.toString)
  monitorAdd("AAA: "+err.number.tostring)
  dim args() 
  '' : myMethod.Invoke(scriptClass, Args)
  '' monitorAdd("BBB: "+err.number.tostring)
  '' 'myMethod.Invoke(scriptClass)
  '' 
  '' 'where "strFnName" is your method name and "Args" is an array of parameter
  '' 'values (objects) for the method's arguments in the same order.
  '' 
  '' 
  
  monitorAdd("????")
  monitorAdd(err.description)
  monitorAdd("????")
  'exit sub
  
  
  if ERR.number<>0 then
    ERR.number=0: dim errMes="ERR: Sub '"+funcName+"' ...nicht gefunden" 
    statustext(errMes)
    'printLine 11, "ERR", errMes
    'trace         "ERR", errMes
    'ref.ResumeLayout(False)
    'ref.PerformLayout()
    exit sub
  End if
  
  
  
  for i = 1 to 1000000
     'scriptClass.test2() 
     'msgbox(GetAsyncKeyState(19).toString)
     if GetAsyncKeyState(19) <> 0 then      '...PAUSE = abbrechen
       'msgbox("break")
       errorText("...Abgebrochen")
       exit for
     end if

     ' ACHTUNG: 'myMethod.Invoke' ist rund 7 mal so schnell(fast 10 mal schneller) !!!
     ' ACHTUNG: ... und ich kann heausfinden, ob es die proc überhaupt gibt (siehe weiter oben)
     '          ... damit komme ich auf circa 350 000 aufrufe einer leeren funktion
     '          ... bzw. rund 370 000 ohne checkAsyncKeyState
       '' '' :Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
       '' '' if myMethod.toString = "" then ...

     myMethod.Invoke(scriptClass, Args)
     'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
     'scriptClass.test2()
     ' statustext(i.tostring)
     '' : if ERR.number<>0 then
     '' :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
     ''  statustext(errMes)
     ''  'printLine 11, "ERR", errMes
     ''  'trace         "ERR", errMes
     ''  'ref.ResumeLayout(False)
     ''  'ref.PerformLayout()
     ''  exit sub
''     else
''        'statustext("OK: Test1 ausgeführt")
''     End if
     If i Mod 22177 = 0 Then
       deltaTime=GetTime()-timerStart2
       statustext("sek: "+deltaTime.toString("00,000")+ "  anz: "+i.toString("0000"))
       'zz.doEvents
     end if
   next    
  deltaTime=GetTime()-timerStart2
  monitorAdd("-------------- DONE")
  monitorAdd("anz schleifen je sek: "+(i / deltaTime * 1000).toString("0000"))
  monitorAdd("")
  statustext("sek: "+deltaTime.toString("00,000")+ "  anz: "+i.toString("0000"))
  'ref.ResumeLayout(False)
  'ref.PerformLayout()
end sub


sub cmdCallByName (e, tagDATA)
  dim tagPara as string = trim(tagDATA(1))
  statustext_reset()
  monitorAdd("cmdCallByName: "+tagPara)
  
  'dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  'dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  dim scriptClass = me
  if scriptClass is nothing then exit sub
  dim funcName as string=tagPara
  dim i as integer
  for i = 0 to 0
    CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
    statustext(i.tostring)
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
      exit sub
    else
      statustext("OK: '"+funcName + "' ... ausgeführt")
    End if
  next    
end sub



sub aufZu_MouseClick(e)
  'msgBox("aufZu_MouseClick")
  dim parentWin = ref.parent.parent.parent
  static state as boolean =false
  state=not state
  '' monitorAdd(ref.parent.Proportion.toString)
  '' monitorAdd(ref.parent.parent.Proportion.toString)
  '' monitorAdd(ref.parent.parent.parent.Proportion.toString)
  'exit sub
  dim el as object
  monitorAdd(parentWin.controls.count)
  dim i as integer
  '' for i = 0 to parentWin.controls.count-1
  ''   dim element as string=parentWin.controls(i).toString
  ''   if element.contains("SplitterControl") then
  ''     parentWin.controls(i).top=33
  ''     monitorAdd(parentWin.controls(i).controls(0).toString)
  ''     monitorAdd(parentWin.controls(i).controls(1).toString)
  ''     monitorAdd(parentWin.controls(i).controls(2).toString)
  ''   end if
  ''   monitorAdd(parentWin.controls(i).toString)
  '' next
  
  '' for each el in parentWin.controls.count
  ''   monitorAdd(el.name)
  '' next
  
  
  parentWin.SuspendLayout()
  if state =true then
   parentWin.height=70
  else
   parentWin.height=200
  end if
  parentWin.ResumeLayout(False)
  parentWin.PerformLayout()

end sub


sub test1_MouseClick(e)
  'printLine 3, "aaa","bbbbb"
  'monitorClear()
  monitorAdd("=== START")
  statustext_reset()
  zz.doevents
  '' ZZ.IDEHelper.Exec("Window.Reflection", "zz","")
  
  ZZ.IDEHelper.Exec("File.Save", "","")
  'monitorAdd("------- SAVED")
  'zz.doevents
  
  ZZ.IDEHelper.Exec("Debug.Run", "","")
  monitorAdd("----------- COMPILED")
  zz.doevents
  
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
      
  ''': scriptClass.test1()
   'CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)

   'CallByName(scriptClass, "test999", Microsoft.VisualBasic.CallType.Method)

   
   
   
   dim i as integer
   for i = 0 to 0
     CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)
   next




   'CallByName(scriptClass, "test1", Microsoft.VisualBasic.CallType.Method)
  
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    else
       statustext("OK: Test1 ausgeführt")
    End if
  monitorAdd("-------------- DONE")
  monitorAdd("")

''zz.IDEHelper.window.reflection("zz")
  'ZZ.IIDEIndexList.RestorePos(30,20)

  'ZZ.IDEHelper.indexList.RestorePos(30,20)
   'add(ZZ.getActiveRTF.RTF.Lines.Current.Text)
exit sub



  '' statustext_reset()
  '' dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  '' : dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  '' if scriptClass is nothing then exit sub
  '' : scriptClass.test1()
  '' : if ERR.number<>0 then
  '' :   ERR.number=0: dim errMes="ERR: Sub 'test1()' ...nicht gefunden" 
  ''     statustext(errMes)
  ''     'printLine 11, "ERR", errMes
  ''     'trace         "ERR", errMes
  ''   End if
End Sub

sub test2_MouseClick(e)
  statustext_reset()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  : dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
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
  : dim scriptClass = getScriptClassFromFileSpec(ActiveTabFilespec)
  if scriptClass is nothing then exit sub
  : scriptClass.test3()
  : if ERR.number<>0 then
  :   ERR.number=0: dim errMes="ERR: Sub 'test3()' ...nicht gefunden" 
      statustext(errMes)
      'printLine 11, "ERR", errMes
      'trace         "ERR", errMes
    End if
End Sub



:function getScriptClassFromFileSpec(filespec as string)
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

sub myTest(optional para1="", optional para2="",)
  msgbox("xxx myTest: "+para1)
  msgbox("xxx myTest: "+para2)
end sub

sub myTest2(optional para1="", optional para2="",)
  msgbox("xxx myTest xxxxx: "+para1)
  msgbox("xxx myTest xxxxxxx: "+para2)
end sub
 



'-->

'--> ICON-LIST


: Function iniIconList() as string
dim curEditorLine as integer=zLN
#Data myIconList as String
''============================================================================================================================================
callByName   |clearAll                                           |http://es.teamwiki.net/docs/icon24/emblem-noread.png
'' saveRun      |Debug.Run                                          |http://es.teamwiki.net/docs/icon24/Good-mark.png 
'' startTest    |test1                                              |http://es.teamwiki.net/docs/icon24/white01.png 
'' startTest    |test2                                              |http://es.teamwiki.net/docs/icon24/white02.png 
'' startTest    |test3                                              |http://es.teamwiki.net/docs/icon24/white03.png 



toggle       |ToolBar|##|tbScriptWin|##|es_testLabor.mainWin     |http://es.teamwiki.net/docs/icons/gnome-monitor3.png
toggle       |Addin|##|siacodecompiler|##|SHPrintline            |http://es.teamwiki.net/docs/icon24/gnome-dev-printer.png
toggle       |Addin|##|siaCodeCompiler|##|SHDebug                |http://es.teamwiki.net/docs/icon24/debug.png
toggle       |ToolBar|##|tbScriptWin|##|es_internSuche.mainWin   |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
toggle       |ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin      |http://es.teamwiki.net/docs/icon24/insert-object.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_colorPicker.dll|##|sw_colorPicker.sg_colorPicker        |http://icons3.iconfinder.netdna-cdn.com/data/icons/nuvola2/32x32/apps/kwikdisk.png


'' toggle       |ToolBar|##|tbScriptWin|##|es_bookmarks2.mainWin    |http://es.teamwiki.net/docs/icon24/FavouritesGelb.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo        |http://es.teamwiki.net/docs/icon24/ModifyEdit.png
toggle       |ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock      |http://es.teamwiki.net/docs/icon24/Clock.png

exe          |C:\yEXE\traceMonitor.exe                         |http://es.teamwiki.net/docs/icon24/agentsvr_113-32.png
exe          |C:\yEXE\ide3\IprocViewer.exe                     |http://es.teamwiki.net/docs/icon24/checkered_flag.png
exe          |C:\yEXE\extractAssociatedIcon.exe                |http://es.teamwiki.net/docs/icon24/stock_draw-selection.png 
exe          |C:\yEXE\mwSidebar.exe                            |http://es.teamwiki.net/docs/icon24/SysTrayDemo2_1-32.png 

startTest    |AutoStart        |http://es.teamwiki.net/docs/icon24/pifmgr_14-32.png 

startTest    |test4            |http://es.teamwiki.net/docs/icon24/darkblue01.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/gtk-edit.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_effects-preview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/redhat-home.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/Textpreview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_effects-preview.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/stock_calc-cancel.png
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
startTest    |test4            |http://es.teamwiki.net/docs/icon24/ViewSearch.png 
 '' 
 '' 
 '' 010 | http://es.teamwiki.net/docs/icon24/checkered_flag.png 
 '' 011 | http://es.teamwiki.net/docs/icon24/SearchFernglas.png
 '' 012 | http://es.teamwiki.net/docs/icons/navigation-180-white.png
 '' 013 | http://es.teamwiki.net/docs/icons/navigation-270-white.png 
 '' 014 | xxx 
 '' 015 | http://es.teamwiki.net/docs/icons/navigation-090-button.png 
 '' 016 | http://es.teamwiki.net/docs/icons/navigation-180-button.png 
 '' 017 | http://es.teamwiki.net/docs/icons/navigation-270-button.png 
 '' 017 | http://es.teamwiki.net/docs/icon24/Good-mark.png
 '' 018 | xxx http://es.teamwiki.net/docs/icon24/stock_test-mode.png
 '' 019 | xxx 
 
''============================================================================================================================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=curEditorLine & vbNewLine & myIconList
  redim iconList(111)
  dim DATA() as string=split(RESULT,vbNewLine)
  dim LINE() as string
  dim temp as string
  dim index as integer
  dim counter as integer =0
  'dim icon,type, para as string
  dim max as integer=DATA.length
  dim i as integer
  for i =1 to 25'   max-1
    temp=trim(DATA(i))
    if temp="" then continue for
    if temp.startsWith("'") then continue for
    temp = replace(temp, "|##|","_##_")+"||"             '...scriptedPanelId maskieren
    temp = replace(temp, "|","<³³>")                     '...scriptedPanelId maskieren
    temp = replace(temp, "_##_", "|##|")           '...scriptedPanelId maskieren
    LINE = split(DATA(i), "<³³>")
    index = val(LINE(0))
    iconList(counter)=temp
    counter=counter+1
  next
  return RESULT
end function 


: sub insertIconsFromIconList(PnlRef)
  on error resume next
  dim index, left, top, deltaTop, deltaLeft as integer
  dim text, icon, item  as string
  dim DATA() as string
  dim el as object
  ''dim bgColor="#696E73"
  dim bgColor="#4A7CDB"  'E0DDD7

  with pnlRef
      .insertX = 0: .insertY= 0: 
    el=.addLabel ("editIconBar"     , "e"   , "#3B5951", "#6DFF50"        ,   , ,45,19   ): 
    'gErrCounter=el 
    '' el= .addTextbox ("ErrCounter",   50,"Err:"    , 0,"" , , , )
    '' el= ref.element("lab_ErrCounter")
    'el.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
    'el.backColor=ColorTranslator.FromHtml("#f66")
    'el.textAlign=System.Drawing.ContentAlignment.MiddleCenter
    el.textAlign=System.Drawing.ContentAlignment.topCenter
    el.backColor=ColorTranslator.FromHtml("#4579D7")
    el.foreColor=ColorTranslator.FromHtml("#bbb")
    el.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.borderstyle=0
    el.autosize=false
    el.height=32
    el.width=15
    el.text="e"
  end with

  deltaTop=0
  deltaLeft=0
  for index =0 to 11
    text="" ' index.toString
    item=iconList(index)
    if trim(item)="" then continue for


    DATA=split(item,"<³³>")
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    icon=trim(DATA(2))
    if icon="xxx" then 
      icon=""
      text="" '"<"+index.toString+">"
    end if
    left=(index mod 10)*30
    top=(index - (index mod 10)) *4
    left=15+(index )*30
    top=(index - (index mod 10)) *4
    top=0
    dim id as string=trim(DATA(0))+"_"+index.toString
    ''monitorAdd(top.toString, left.tostring)
   
     el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,32 ,32 ,icon) ' ,flags,handler)
     el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
    'el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    'el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay 'ImageAboveText
    el.foreColor=ColorTranslator.FromHtml("#fff")
    'el.UseVisualStyleBackColor = True
    el.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    el.FlatAppearance.BorderSize = 1
    el.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#FFD350")  '#4A7CDB
    el.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#5A8CeB")
    el.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.tag=item
    el.AccessibleRole = System.Windows.Forms.AccessibleRole.CheckButton
    'el.checked=true
    'el.Checked = True
    'el.CheckState = System.Windows.Forms.CheckState.Checked

  next
end sub


:sub insertIconsFromIndex(ref as object,index as integer, left as integer, top as integer)
  on error resume next
  dim bgColor as string = "#fff"
  dim txt as string=val(index)
  left=(index mod 10)*30
  top=(index - (index mod 10)) *40
  
  
  'ref.addButton(id,txt,bgColor,Left,Top,30 ,40 ,iconList(index)) ' ,flags,handler)

  dim el as object
  with ref
    .insertX = 5:.insertY = 5:
    el=.addButton  ("imgList01"      , "33"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(40))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    
    .insertX = 40:.insertY = 150:
    el=.addButton  ("imgList02"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter

   .insertX = 10:.insertY = 190:
    el=.addButton  ("imgList03"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
  end with

end sub 

: sub xxx_insertIconsFromIndex(ref as object,index as integer, left as integer, top as integer)
  
  dim bgColor as string = "#fff"
  dim txt as string=val(index)
  left=(index mod 10)*30
  top=(index - (index mod 10)) *40
  
  
  'ref.addButton(id,txt,bgColor,Left,Top,30 ,40 ,iconList(index)) ' ,flags,handler)

  dim el as object
  with ref
    .insertX = 5:.insertY = 5:
    el=.addButton  ("imgList01"      , "33"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(40))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    
    .insertX = 40:.insertY = 150:
    el=.addButton  ("imgList02"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter

   .insertX = 10:.insertY = 190:
    el=.addButton  ("imgList03"      , "45"    , "#ccc")
    el.tag="...tag": el.height=40:el.width=30
    el.image=ZZ.GetImageCached(iconList(41))
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
  end with

end sub 



: sub xxx_insertIconsFromIconList(PnlRef)
  PnlRef.SuspendLayout()
  dim index, left, top, deltaTop, deltaLeft as integer
  dim text, icon, item  as string
  dim DATA() as string
  dim el as object
  ''dim bgColor="#696E73"
  dim bgColor="#3B5951" '"#2F394E"  '"#4A7CDB"  'E0DDD7
  deltaTop=0
  deltaLeft=0
  
  for index =0 to 25
    text="" ' index.toString
    item=iconList(index)
    if trim(item)="" then continue for

    DATA=split(item,"<³³>")
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    icon=trim(DATA(2))
    if icon="xxx" then 
      icon=""
      text="" '"<"+index.toString+">"
    end if
    left=(index mod 10)*30
    top=(index - (index mod 10)) *4
    
    top=index*29'34

     top=2
    'left=2
     dim id as string=trim(DATA(0))+"_"+index.toString
    ''monitorAdd(top.toString, left.tostring)
   
     el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,45 ,29 ,icon) ' ,flags,handler)
     'el=PnlRef.addButton(id ,text,bgColor,Left+deltaLeft,Top+deltaTop,32 ,32) ' ,flags,handler)
    
     'el.BackgroundImage = ZZ.GetImageCached(icon) ''CType(resources.GetObject("Button1.BackgroundImage"), System.Drawing.Image)
     'el.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom

     el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
    'el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    'el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay 'ImageAboveText
    el.foreColor=ColorTranslator.FromHtml("#fff")
    'el.UseVisualStyleBackColor = True
    el.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    el.FlatAppearance.BorderSize = 0
    el.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#FFD350")  '#4A7CDB
    el.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#5A8CeB")
    el.Padding = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
    el.tag=item
    el.AccessibleRole = System.Windows.Forms.AccessibleRole.CheckButton
    'el.checked=true
    'el.Checked = True
    'el.CheckState = System.Windows.Forms.CheckState.Checked

  next

  PnlRef.ResumeLayout(False)
  PnlRef.PerformLayout()

end sub



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

  Sub createOpenedTabList(DockPanel1)
      trace "--------------", "------------------------------ ???"
    Dim doc as object
    'doc = listDocumentPane(DockPanel1)
    
    If doc Is Nothing Then Exit Sub
    'Dim scrollPos = tbOpenedFiles.IGrid1.VScrollBar.Value
    'tbOpenedFiles.IGrid1.Rows.Clear()
    'tbOpenedFiles.IGrid1.Rows.Count = doc.Contents.Count
    dim i 
    For i = 0 To doc.Contents.Count - 1
      trace "xxx",  doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Cells(i, 0).Value = doc.Contents(i).DockHandler.TabText
      'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()
      'If doc.Contents(i) Is MAIN.ActiveMdiChild Then tbOpenedFiles.IGrid1.SetCurRow(i)
    Next
    'tbOpenedFiles.IGrid1.VScrollBar.Value = scrollPos
  End Sub




sub sendMesTest2
   monitorAdd("schnellTester-3","...ich bin die nachricht")
exit sub

  'msgBox("OK - 3")
   monitorAdd("schnellTester-3","...ich bin die nachricht")
  dim funcName =ref.element("funcName").text
  ''msgbox(funcName)
  dim para1 = "paaaaaaaaaara"
  ''printLine 1, "MouseClick(e)", funcName
    ''msgbox("myTest: "+para1)

  if funcName<>"" then
    callMySub(funcName, para1)
    '' CallByName(Me, funcName, Microsoft.VisualBasic.CallType.Method, para1)
    '' exit sub
  else
    callMySub("test3",nothing)
  end if
  'msgBox("....OK - 3")
end sub



sub sendMesTest3
   monitorAdd("schnellTester-2","...ich bin die nachricht")
exit sub

  'msgBox("OK - 2")
   monitorAdd("schnellTester-2","...ich bin die nachricht")
     printLine 11, "MouseClick(e)", "test2"
    dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
 
  trace "ActiveTabFilespec", ActiveTabFilespec
  trace "activeTabType", activeTabType
  trace "", 
  trace "", 
 
  dim filespec as string
  dim path, name, ext as Object
  filespec=ActiveTabFilespec
  zz.splitFilespecData(filespec, path, name, ext)
  trace "path",  path
  'printLine 2, "splitFilespecData", name
  dim scriptClass as object
  scriptClass=zz.scriptClass(name)
  if not scriptClass is nothing then
     : scriptClass.test1()
     : if err.number <>0 then
     :   printLine 11, "ERR", "test1 nicht gefunden"
     :   err.number =0
     : end if
  end if
  ''msgBox ("fileAction 333")
  
  '' msgBox(name)

 ' msgBox("....OK - 2")
end sub



'--> ACHTUNG: besser gelöst in 'popupMonitor'
sub callMySub(funcName, para)
  dim fileSpec as string="C:\yPara\scriptIDE\scriptClass\es_schnellTester.vb"
  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  :printLine 2, "splitFilespecData", name
  :scriptClass=zz.scriptClass(name)
  :zz.doEvents
  :zz.doEvents
  :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     exit sub
  end if
  
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method, para)
  : CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)
  : if ERR.number<>0 then
  :   ERR.number=0
      printLine 11, "ERR", "ERR: Sub 'test2()' ...nicht gefunden" 
      trace         "ERR", "ERR: Sub 'test2()' ...nicht gefunden" 
      exit sub
  end if
end sub

















'-->
'--> D I A L O G

public auto as integer =-2
public globTempHeaderData(200) as string
public globTempHeaderCounter as integer =0

sub showDialog()
  'msgBox("???")
  AutoStart2
  dim dialogData as string = getDialogData()
  'msgBox(dialogData)
  createDialogfromData(ref2, "", dialogData)
end sub


function getDialogData()                  
#Data myDialogData as String              
'======================================START-DIALOG
TYPE	id / cmd	nicName	fullName	Titel	DIALOG	edit/view/hidden	labBgColor	labForeColor	textBgColor	textForeColor	GRID	defaulWidth	gridBgColor	textForeColor	HELP	toolTip	text1	text2	icon	
1											xxx			1995-01-01	1999-01-02	SONDER(es)				
dialog			feld	titel																
dialog			vorname	hinz und kunz ?																
dialog			name	huber oder mayer ?								1	2	3	4	5	6	7	8	
dialog			PLZ	...5-stellig !								a	b	c	d	e	f	g	h	
1																				
1																				
1																				
1																				
dialog			schuhgröße	38-46																
dialog			schuhgröße	38-46																
dialog			schuhgröße	38-46																
																				
																				
'======================================END-DIALOG
#EndData                                  
  '' Trick 17: Zeilennummer dazupacken    
  dim RESULT as string=myDialogData       
  return RESULT                           
end function                              


Sub createDialogfromData(container As ScriptedPanel, fileSpec as string, optional optData as string="")
  '' dim DATA as string="LEER"
  ''msgBox("createDialogfromData")
  dim dialogData as string
  if optData = "" then
    dialogData = IO.File.ReadAllText(fileSpec)
  else
    dialogData = optData
  end if
  
  Dim DATA() as string=split(dialogData,vbNewLine)
  dim i as integer
  dim max as integer=DATA.length
  dim item as string
  dim FIELDS() as string
  
  dim index as integer = 0
  dim fullName as String
  
  container.resetControls (5,5)

  for i=0 to max-1
    item=DATA(i)
    FIELDS=split(item, vbTab)
    if FIELDS(0).toUpper.startsWith("DIALOG") then
      monitorAdd(item)
      fullName=FIELDS(3)
      createDialogElement(container, index, fullName , index.toString, item)
      index=index+1
    end if
  next
end sub

function createDialogElement(container as object, index as integer, labText as string, defaultValue as string, tagData as object)
    ' ### ??? stripHeaderData(labText)
    ''zNN=1
    'dim labBgColor="#A1A1DB"
    'dim labBgColor="#629FFF"
    dim labBgColor="#3477D7"
    dim el as object
    dim labWidth as integer=120
    with container 
      dim id="dlg_"+index.toString
      .addTextbox(id, Auto, labText+"  ", labWidth, labBgColor)
      el = .element("txt_"+id)
      el.text=defaultValue.toString
      el.AutoSize = false
      el.height=18
      'el.backColor=ColorTranslator.FromHtml("#E7E7FF")
      el.backColor=ColorTranslator.FromHtml("#fff")
      el.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      el.BorderStyle = System.Windows.Forms.BorderStyle.None      
      el.tag=tagData
      
      'label...
      el= .element("txtDesc_"+id)
      el.textAlign=System.Drawing.ContentAlignment.MiddleRight
      el.height=18
      el.foreColor=ColorTranslator.FromHtml("#fff")
      el.tag=tagData
      .br(20)
    end with
end Function


sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    'el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


sub refreshHelpLabels(e as object, index as string, value as string )
  dim tagData as string=e.sender.tag.toString
  monitorAdd("tagData: "+tagData)
  dim FIELDS() as string = split(tagData, vbTab)
  dim titel=FIELDS(4)
  monitorAdd(titel)
  ''dim el as object=panelRef.element("lab_labA_01")
  dim el as object=ref2.element("lab_labA_01")
  ''msgbox(el.toString)
  el.text=titel
end sub


'' sub stripHeaderData(labelText)
''   globTempHeaderData(globTempHeaderCounter)=labelText
''   globTempHeaderCounter=globTempHeaderCounter+1
'' end sub
'' 

'' Function getHeaderData() as string
''   redim preserve globTempHeaderData(globTempHeaderCounter)
''   globTempHeaderCounter=0
''   dim RESULT = join(globTempHeaderData,vbNewLine)
''   redim globTempHeaderData(200)
''   return RESULT
'' end function


'' sub XXX_setGridHeaderData(grid as TenTec.Windows.iGridLib.iGrid, labelTextAsNewLineString as string)
''   dim DATA() as string = split(labelTextAsNewLineString, vbNewLine)
''   dim max as integer=DATA.length
''   grid.Cols.Count=max
''  
''   dim i as integer
''   for i = 0 to max-1
''     grid.cols(i).text=DATA(i)
''     grid.cells(0,i).value=DATA(i)
''   next
''    
''   'dim grid as TenTec.Windows.iGridLib.iGrid = iGrid1
''   'Dim row As TenTec.Windows.iGridLib.iGRow
''   'dim curRow As TenTec.Windows.iGridLib.iGRow
''   'dim curRowIndex as integer
''   
''   '' curRow=grid.curRow
''   '' curRowIndex=grid.curRow.index
''   '' trace "curRow", curRow
''   '' curRow.cells(cInt(index)).value=value
''   '' 
''   '' dim lineData as string = JoinIGridLine(grid.curRow)
''   '' monitorAdd(lineData)
''   '' dbUpdateFromLineData(lineData)
''   '' with grid
''   ''   grid.cols(0).text="000"
''   ''   .cols(1).text="111"
''   ''   .cols(2).text="222"
''   ''   
''   '' end with
''   '' 
'' 
'' end sub








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




