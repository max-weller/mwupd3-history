
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
Const  winID = "{ScriptClass}.mainWin"
const  globHideOnClose as boolean = false           'bei true merkt er sich die Position
public winCaption as string="Direkt-Fenster-xxx"
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
  monitorAdd ("test1()[.]")
  monitorAdd ("test1()[..]")
  monitorAdd ("test1()[...]")
  monitorAdd ("test1()[....]")
  monitorAdd ("test1()[.....]")
exit sub

end sub




sub test2()
  'msgBox("OK - 2")
  dim script=zz.scriptClass("es_temp3")
  script.test1()

exit sub

  demoCodeInsert()
end sub



sub test3()
  msgBox("OK - 3")
  IO.Directory.CreateDirectory(glob_defaultPath)
  dim filespec as string =glob_scriptPath+"es_temp3.vb"
  dim template as string=getCodeTemplate()
  dim insert as string =myTextArea.text
  '' msgbox(insert)
  '' myTextArea.text="msgBox(123)"
  '' insert=myTextArea.text
  msgbox(insert)
  dim content as string=replace(template, "[[INSERT-TEST-CODE]]", insert)
  msgBox(content)
  'exit sub
  zz.filePutContents(filespec, content) ', [append As System.Object = False]) As System.Object
  dim script=zz.newScriptClass("es_temp3")
  script.test1()

end sub
'-->

'--> spielwiese für shortcuts
  '##fn

  '##fn



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



'-->
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








sub demoCodeInsert()
  dim codeEitor = ZZ.getActiveRTF.RTF
  dim lineContent as string = trim(codeEitor.Lines.Current.text)
  lineContent=replace(lineContent,vbnewLine,"")
  ''msgBox(lineContent)
  '##fn
  if lineContent="'##fn" then
    ''monitorAdd("TREFFER")
    dim codeBlock=getDataBlock()
    
    Dim i as integer
    dim DATA()=split(codeBlock,vbNewLine)
    dim item as string
    dim max as integer=DATA.length
    dim counter as integer = -1
    dim OUT() as String
    redim OUT(1000)

    for i = 0 to max -1
      item = trim(DATA(i))
      if item = "" then continue for
      if item.startsWith("'' ") then continue for
      '' ...
      counter=counter+1
      OUT(counter)=item
    next
    redim preserve OUT(counter)
    dim RESULT as String = join(OUT,vbNewLine)
    monitorAdd(RESULT)
    dim forInsert as string=RESULT
    forInsert=replace(forInsert, vbNewLine, "")
    forInsert=replace(forInsert, "<br>", vbNewLine)
    ''monitorAdd(forInsert)
    
    codeEitor.Lines.Current.text=forInsert
    
  end if
  
exit sub


  '' dim tab
  '' If Left(tab.URL,5) <> "loc:/" Then Exit Sub
  '' Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  '' :dim scriptClass = callScriptClassTestProc(filespec)
  '' :scriptClass.test1()
  '' :if ERR.number<>0 then
  '' :   ERR.number=0
  ''     printLine 11, "ERR", "ERR: Sub 'test1()' ...nicht gefunden" 
  ''     trace         "ERR", "ERR: Sub 'test1()' ...nicht gefunden" 
  ''     exit sub
  '' end if
  '' 

end sub




function getDataBlock()
#Data myData as String
'' ############################################################
||NEXT-ITEM||
fn__||__
<br>  '------------------------------------------------------- '...fn
<br>  '[fn = for next] ...versuch einer allgemeingültigen Schleife
<br>
<br>  Dim i as integer
<br>  dim DATA()
<br>  dim item as string
<br>  dim max as integer=DATA.length
<br>  dim counter as integer = -1
<br>  dim OUT() as String
<br>  redim OUT(1000)
<br>
<br>  for i = 0 to max -1
<br>    item = trim(DATA(i))
<br>    if item = "" then continue for
<br>    if item.startsWith("'") then continue for
<br>    '' ...
<br>    counter=counter+1
<br>    OUT(counter)=item
<br>  next
<br>  redim preserve OUT(counter)
<br>  dim RESULT as String = join(OUT,vbNewLine)
<br>  monitorAdd(RESULT)
<br>
<br>  '-----------------------------------------------------
||NEXT-ITEM||
test__||__
<br>  '-----------------------------------------------------
<br>'-->
<br>'-->T E S T 
<br>sub test1()
<br>  msgBox("OK - 1")
<br>end sub

<br>sub test2()
<br>  msgBox("OK - 2")
<br>end sub

<br>sub test3()
<br>  msgBox("OK - 3")
<br>end sub
<br>'-->
<br>  '-----------------------------------------------------
||NEXT-ITEM||
xxx__||__
'' ##########################################################
#EndData
  dim RESULT as string=myData 
  return RESULT
end function




'-->
Function GetPanelRef()
  onTerminate() 'aufruf, um das alte leben ein bischen anzuhalten 
  ' If PanelRef IsNot Nothing Then Exit Sub
  'PanelRef = ZZ.IDEHelper.CreateDockWindow(winID, 1) '1=popupWin
  '--> dockWindow / normales Fenster
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.StdWindow, 1)'  : err.number=0

  'formRef = getOuterWindowRef(panelRef)
  formRef = panelRef.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose : err.number=0
  return panelRef
End Function

Sub GetFormRef()

  'If ref IsNot Nothing Then Exit Sub
  'ref = ZZ.IDEHelper.CreateDockWindow("es_temp2.mainWin", 1)
          'ZZ.IDEHelper.CreateDockWindow(winID, 4): err.number=0 

  panelRef = ZZ.CreateWindow("es_temp2.mainWin", DWndFlags.DockWindow, 1)'  : err.number=0
  'formRef = panelRef.form
  ''formRef=getOuterWindowRef(panelRef)
  formRef = panelRef.form
  ''formRef.text=winCaption

  formRef.text=winCaption
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = true
  
  ''DAS MACHT EINE NORMALE FORM
  ''formRef = New frmTB_scriptWin 'ZZ.IDEHelper.CreateForm(winID)
  ''panelRef = CType(FormRef, Object).PNL
  ''FormRef.Text = "Snippet-Manager"
End Sub


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
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf Timer1_Tick
  'AddHandler formRef.Resize, AddressOf Form_Resize
  'AddHandler TT.TraceWrite, AddressOf OnTrace
  'AddHandler TT.PrintLineChanged, AddressOf OnPrintLine
end sub


Sub globRemoveHandler()
  trace "globRemoveHandler","try..."
  if formRef is nothing then exit sub
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








'-->
'--> A U T O S T A R T

Sub AutoStart()
    monitorClear()
   '' 'msgbox("halli hallo")
   ''  monitorAdd ("........ AUTOSTART .........")
   ''  monitorAdd ("........ ......... .........")
   ''  monitorAdd ("........ AUTOSTART .........")
   ''  
   
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
  '--> ...textarea
    '.insertX = 5 :.insertY = 0' 110
    .addTextbox ("myTextArea" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
     el=panelRef.element("txt_myTextArea")
     myTextArea=el
     el.Dock = DockStyle.Fill
     el.WordWrap=false
     el.scrollbars = System.Windows.Forms.ScrollBars.Both
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.bringToFront()
     myTextArea.text="'abc"
     '--> ... ... el.text=getMiniHelpText()
     '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
  end with

  With panelMain
    '.activateEvents = "|IconMouseDown|   |TextboxKeyDown|"
    '' '...Listbox erzeugen
    '' Dim typName
    '' typName = "System.Windows.Forms.ListBox, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089, processorArchitecture=MSIL"
    '' .addControl ("ctrl_lv",typName)
    '' .Element("ctrl_lv").Top=20
    '' .Element("ctrl_lv").Width=400 : .Element("ctrl_lv").Height=400
    '' With .Element("ctrl_lv").Items
    ''   .add ("test 1")
    ''   .add ("test 22 ")
    ''   .add ("test 333")
    ''   .add ("test 4444")
    '' End With
  end with

end sub


  

'-->
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

