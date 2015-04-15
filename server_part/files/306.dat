
''es_temp2.vb

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





'-->
'--> G L O B 
#Para DebugMode intern
#Para SilentMode true



'--> P A R A M E T E R -------------------------
Const  winID = "es_direktFenster.mainWin"
const  globHideOnClose as boolean = true           'bei true merkt er sich die Position
public winCaption as string="Direkt-Fenster 00000"
public glob_defaultPath as string = "C:\yPara\scriptIDE\direktFenster\"
public glob_scriptPath as string = "C:\yPara\scriptIDE\scriptClass\"

dim toolBarColor as string="#3B5951"


public WithEvents FormRef As Form
public PanelRef As ScriptedPanel

public shared toolBar As ScriptedPanel
public shared statBar As ScriptedPanel
public shared panelMain As ScriptedPanel

dim myTextArea as textbox

Public Const Auto = -2
shared myScriptClass


'-->
sub test1()
  'trace ("test1()[es_temp2....xxxxxxxxxxxxxx.....]")
  monitorAdd ("test1()[.]")
  'exit sub
  monitorAdd ("test1()[..]")
  monitorAdd ("test1()[...]")
exit sub

end sub




sub test2()
  'msgBox("OK - 2")
  dim script=zz.scriptClass("es_temp3")
  script.test1()

exit sub

  'demoCodeInsert()
end sub



sub test3()
  'msgBox("OK - 3")
  executeDirektFenster()
end sub







'-->
'--> FORM-HELPER/...Factory ________________________


:function getOuterWindowRef(panelRef as object) as object
  on error resume next
  if panelRef.form.parent.parent is Nothing then : err.number=0
    '...falls es kein DockPanel-Fenster ist
    :return panelRef.form
  else : err.number=0
    '...DockPanel-Fenster haben noch weitere Schichten dazwischen:
    :return panelRef.form.parent.parent
  end if
End function



Sub globAddHandler()
   AddHandler  myTextArea.KeyDown, AddressOf onMyTextArea_KeyDown
   AddHandler  myTextArea.KeyPress, AddressOf onMyTextArea_KeyPress
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf Timer1_Tick
  'AddHandler formRef.Resize, AddressOf Form_Resize
  'AddHandler TT.TraceWrite, AddressOf OnTrace
  'AddHandler TT.PrintLineChanged, AddressOf OnPrintLine
  'AddHandler  myTextArea.KeyDown, AddressOf OnPrintLine
end sub


Sub globRemoveHandler()
  trace "globRemoveHandler","try..."
  'if formRef is nothing then exit sub
  RemoveHandler  myTextArea.KeyDown, AddressOf onMyTextArea_KeyDown
   RemoveHandler  myTextArea.KeyPress, AddressOf onMyTextArea_KeyPress
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'RemoveHandler Timer1.Tick, AddressOf Timer1_Tick
  'RemoveHandler formRef.Resize, AddressOf Form_Resize
  'RemoveHandler TT.TraceWrite, AddressOf OnTrace
  'RemoveHandler TT.PrintLineChanged, AddressOf OnPrintLine
  trace "globRemoveHandler","DONE"
end sub


sub onInitialize()
  globAddHandler
end sub


sub onTerminate()
  globRemoveHandler
  'stopTimer()
end sub


Function GetPanelRef()
  onTerminate() 'aufruf, um das alte leben ein bischen anzuhalten 
  ' If PanelRef IsNot Nothing Then Exit Sub

  '--> ... --- dockWindow / normales Fenster
  'VERALTET: PanelRef = ZZ.IDEHelper.CreateDockWindow(winID, 1) '1=popupWin
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.StdWindow, 1)'  : err.number=0

  'formRef = getOuterWindowRef(panelRef)
  formRef = panelRef.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose : err.number=0
  return panelRef
End Function



'-->
'--> A U T O S T A R T ___________________________

Sub AutoStart()

  monitorClear()
  GetPanelRef()
  dim el as object
  
  '--> ... containerMain
  With PanelRef
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"
    panelMain  =.addPanel("containerMain"   , DockStyle.Fill)
    toolBar     =.addPanel("toolBar"        , DockStyle.Top, intHeight:=33)
    statBar     =.addPanel("statBar"        , DockStyle.Bottom, intHeight:=33)
  end with
  
  with toolBar
    .resetControls  (10,3)
    .visible=true
    .height=30
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    '.BackColor = ColorTranslator.FromHtml("#f88")
  end with
    
  with panelMain 
    .resetControls  (0,0)
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    .BackColor = ColorTranslator.FromHtml("#E3E0DB")
  end with
  
  with statbar
    .resetControls  (10,5)
    .visible=true
    .height=30
    .BackColor = ColorTranslator.FromHtml(toolBarColor)
    '.BackColor = ColorTranslator.FromHtml("#88f")
    statBar.height=30
  end with
   
  With panelMain
  '--> ... --- textarea
    '.insertX = 5 :.insertY = 0' 110
    .addTextbox ("myTextArea" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
     el=panelRef.element("txt_myTextArea")
     myTextArea=el
     el.Dock = DockStyle.Fill
     el.WordWrap=false
     el.scrollbars = System.Windows.Forms.ScrollBars.Both
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.bringToFront()
     
     dim miniHelp as string="" 
     miniHelp=miniHelp+vbNewLine+"'...hier Code eingeben und dann contr.ENTER drücken"
     miniHelp=miniHelp+vbNewLine+"'--------------------------------------------------"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"'z.B."
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"  msgBox(""hello World"")"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+"'...happy coding ;-)"
     miniHelp=miniHelp+vbNewLine+""
     miniHelp=miniHelp+vbNewLine+""
     myTextArea.text=miniHelp
  end with

  globAddHandler()

  myTextArea.SelectionStart=myTextArea.text.length
  myTextArea.focus()

end sub


'-->
'--> E V E N T S ___________________________

dim skipNextKeyPress as boolean=false

sub txt_myTextArea_KeyDown(e)


  '' ??? wie kann man contr.ENTER bequem abgreifen, ohne daß er eine in der Textbox einfügt ???

  '' printLine 1, "direktFenster", e.keyString
  '' if e.keyString="CTRL-RETURN" then
  ''   executeDirektFenster()
  ''   e.cancel=true
  '' end if
end sub


sub onMyTextArea_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
  'monitorAdd(e.keycode.toString)
  if e.control=true
    if e.keyCode=13 then
       skipNextKeyPress=true
      'monitorAdd("KeyDown-TREFFER")
    end if
  end if
end sub



sub onMyTextArea_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
  'monitorAdd(e.KeyChar)
  if skipNextKeyPress =true then
    skipNextKeyPress=false
    e.handled=true
    monitorAdd("onMyTextArea_KeyPress: CONTROL-ENTER")
    executeDirektFenster()
  end if
end sub


'--> 

sub executeDirektFenster()
  'msgBox("OK - 3")
  dim template as string=getCodeTemplate()
  dim insert as string =myTextArea.text
  dim content as string=replace(template, "[[INSERT-TEST-CODE]]", insert)

  dim filespec as string =glob_scriptPath+"es_temp3.vb"
  IO.Directory.CreateDirectory(glob_defaultPath)
  zz.filePutContents(filespec, content) ', [append As System.Object = False]) As System.Object
  
  'dim script=zz.ScriptClass("es_temp3")
  
  dim script=ScriptIDE.ScriptHost.ScriptHost.Instance.ScriptClass("es_temp3", True)'recompileScriptclass=True
  script.test1()
end sub 


'-->
'--> ......................... D A T A - B L O C K 
function getCodeTemplate()
#Data myData as String

'#####################################################

'NEU-2...
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Reference System.Windows.Forms.dll
#Imports System.Windows.Forms

#Reference System.Drawing.dll
#Imports System.Drawing

#Reference System.Data.dll
#Imports System.Data

'' #Reference TenTec.Windows.iGridLib.iGrid.dll
'' #Imports Tentec.Windows.iGridLib
'' 
'' #Reference WeifenLuo.WinFormsUI.Docking.dll
'' 



'--> G L O B 
#Para DebugMode intern
#Para SilentMode true



sub test1

  'msgBox("ich bin die test1 von ES_TEMP3")
  [[INSERT-TEST-CODE]]
end sub
'###################################################

#EndData
  dim RESULT as string=myData 
  return RESULT
end function


'-->
'--> ........................................LIBs
'--> outMONITOR

sub monitorAdd(para1 as string, optional para2 as string="")
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

