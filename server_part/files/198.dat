
'' es_UserToolbar3


'' 
'' '...get codeLink
'' dim codeEitor = ZZ.IDEHelper.getActiveTab.RTF
'' dim lineContent as string = trim(codeEitor.Lines.Current.text)
'' 
'' '...navigate codeLink
'' dim tab = ZZ.IDEHelper.NavigateFile(fileSpec)
'' dim lineNumber as integer = cInt(tagDATA(4))
'' add("navigiere zu: " +lineNumber.toString)
'' dim codeEitor = ZZ.IDEHelper.getActiveTab.RTF
'' 'dim lineContent as string = codeEitor.Lines.Current.text
'' 'codeEitor.Lines.Current.number=lineNumber
'' codeEitor.goTo.line(lineNumber+50)
'' codeEitor.goTo.line(lineNumber-10)
'' codeEitor.goTo.line(lineNumber)
'' codeEitor.focus()
'' 
'' !!! er kennt RTF nicht mehr
'' ??? wie heißt der neue Befehl ???
'' 
'' --------------------------------------------------------
'' 
'' ' trace schaltet sich immer bei Fehler auf visible 
'' sound / blinky, counter ???
'' trace: getDeserializedDockContent ...ausblenden ?
'' 
'' --------------------------------------------------------
'' 
'' cpuMonitor geht jetzt
'' colorPicker, analogClock und mediaPlayer gehen noch nicht
'' 
'' so schwalben nester
'' di - 10.April 60.peter mayer
'' 
'' hansjörg
'' woche nach ostern 6.00
'' 10 / 11 
'' 
'' 

#Para DebugMode internal
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



#Imports System.Diagnostics
#Imports System.Collections.Generic

#Reference WeifenLuo.WinFormsUI.Docking.dll
#Imports WeifenLuo.WinFormsUI.Docking.DockPanel

'-->
'--> C O N F I G 
public defaultBrowserId as string = "ToolBar|##|tbScriptWin|##|es_webbrowser3.tab" 

'public defaultBrowserScriptId as string = "es_webbrowser3"
public defaultBrowserScriptId as string = "ToolBar|##|tbScriptWin|##|es_webbrowser.tab"
' ??? externe Browser einbinden;: zusatzPara ??? - shell, fileName, umschaltbar machen
''... ZZ.shellExecute("C:\yPara\scriptIDE\compiledScripts\es_webbrowser3\es_webbrowser3.exe")

'iconList
Dim PnlRef As ScriptedPanel
Dim PnlRef2 As ScriptedPanel
dim appId as string="es_UserToolbar3"
' ende IconList

Public Const ToolbarID1 = "es_UserToolbar3.tools"
Public Const ToolbarID2 = "es_UserToolbar3.toolbar2"

shared ref as ScriptedToolstrip'object
shared ref2 as ScriptedToolstrip'object
shared skipAllEvents as boolean =false
shared topPanel(20) as string
shared rightPanel(20) as string
shared popups(20) as string

private privateRefCounter as integer
public publicRefCounter as integer
shared sharedRefCounter as integer
shared sharedEventCounter as integer = 0

Const toggleLB_off = "http://www.iconfinder.net/data/icons/iconic/raster/16/arrow_right.png"
Const toggleLB_on = "http://www.iconfinder.net/data/icons/iconic/raster/16/arrow_left.png"

Friend WithEvents Timer1 As System.Windows.Forms.Timer

Public Const DockTop = 7
Public Const DockLeft = 8
Public Const DockBottom =9 
Public Const DockRight = 10

public iconList() as string
Public cName() as string = _
    {"00 ________ cIndex", _  
     "01 _____ cMainType", _  
     "02 __ cCaptionTiny", _  
     "03 ______ cCaption", _  
     "04 ______ cToolTip", _  
     "05 ______ cSubType", _ 
     "06 __ cGroupMaster", _ 
     "07 ________ cGroup", _  
     "08 ___ cHidenOnIni", _  
     "09 ______ cTagPara", _ 
     "10 ______ cIconURL", _ 
     "11 _____ cIconURL2", _ 
     "12 ________ cFlags", _ 
     "13 ________ cFrei1", _ 
     "14 _____ cDataPlus", _ 
     "15 ___________ XXX", _ 
     "16 ___________ XXX", _ 
     "17 ___________ XXX", _ 
     "17 ___________ XXX", _ 
     "18 ___________ XXX", _ 
     "19 ___________ XXX", _ 
     "20 ___________ XXX", _ 
     "21 ========== LAST" }


Public Const cIndex as integer       = 0 
Public Const cMainType as integer    = 1
Public Const cCaptionTiny as integer = 2
Public Const cCaption as integer     = 3
Public Const cToolTip as integer     = 4
Public Const cSubType as integer     = 5
Public Const cGroupMaster as integer = 6
Public Const cGroup as integer       = 7
Public Const cHidenOnIni as integer  = 8
Public Const cTagPara as integer     = 9
Public Const cIconURL as integer    = 10
Public Const cIconURL2 as integer   = 11
Public Const cFlags as integer      = 12
Public Const cFrei1 as integer      = 13
Public Const cDataPlus as integer   = 14

public globTagData1() as string
public globTagData2() as string




  

'' Public Const cTagPara as integer = 4
'' Public Const cIconURL as integer = 5
'' Public Const cFlags as integer = 6
'' Public Const cFrei1  as integer = 7
'' Public Const cFrei2  as integer = 8
'' Public Const cDataPlus as integer = 9
''     

'-->
'--> TEST
sub test1()
  ref.SuspendLayout()
  static newState as boolean =true
  newState= not newState
  toggleToolBarLabels2(newState)
  'toogleToolbarItems(nothing, nothing)
  ref.ResumeLayout(False)
  ref.PerformLayout()
End Sub




sub test2()
  createIconList()
  
exit sub


  ' msgBox("OK - 2")
  
     dim mainWin as object: mainWin =ZZ.IDEHelper.getMainFormRef
     'dim title as string=ZZ.IDEHelper.MainWindow.ActiveMdiChild.text
     dim panel1 as object = mainWin.DockPanel1
     trace mainWin.text, "xxx"
     trace mainWin.ActiveMdiChild().text, "yyy"
  
  'toogleToolbarItems(nothing, nothing)
  
End Sub


sub test3()
  ZZ.IDEHelper.glob.Para("userToolbar3-paraTest")="ABC...XYZ"
exit sub

  'msgBox("OK - 3")
   trace "-------->>>",  "p.Contents(i).DockHandler.TabText"
   'dim activeTab         = ZZ.getActiveTab()
   'dim activeTabType     = ZZ.getActiveTabType()
   'dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
   '... bringt nicht viel; ausgabe über trace
   dim mainWin as object: mainWin =ZZ.IDEHelper.MainWindow
   listDocumentPane(ZZ.IDEHelper.MainWindow)
End Sub



'-->
'--> T O O L B A R - F A C T O R Y 
   

:Function toolbarFactory_preproc (dataBlock as string)as Array
  on error resume next
  'trace "datenBlock", dataBlock
  dim editorLineNumner as integer
  dim DATA() as string
  dim i as integer
  dim max
  dim out(1000) as string
  dim line as string
  dim counter as integer=0
  
  
  dim u
  
  DATA=split(dataBlock, vbNewLine)
  editorLineNumner=DATA(0)
  for i =1 to DATA.length -1
    line = trim(DATA(i))
    if line="" then continue for
    if line.startsWith("'") then continue for
    if line.startsWith("==========") then continue for
    
    if out(counter)="" then out(counter)=(editorLineNumner+i+1).toString & "|"  '...wichtig für zeilenNummern/mainType
    if line.endswith(" _") then
      out(counter)=out(counter)+Mid(line,1,line.length-2)
    else
      out(counter)=out(counter)+line
      counter =counter +1
    end if
   ''msgBox(line)
  next
  redim preserve out(counter-1)
  return out
End Function


'--> splitTypeInfo
:sub splitTypeInfo(byVal field as string, byRef typeInfo as string, byRef para as string)
  dim findPos as integer
  findPos=instr(field,":")
  ''trace findPos, field
  if findPos>0 then
    typeInfo=mid(field,1,findPos)
    para=mid(field,findPos+1)
  else
    typeInfo=""
    para=field
  end if
End Sub





'--> toolbarFactory()
:sub toolbarFactory(container as object, toolbarData as string, byRef rGlobTagData as string)    
  on error resume next



  'msgBox ("toolbarFactory")
  'monitorClear
  monitorAdd("toolbarFactory: ")
  dim exclusiveId As New Dictionary(Of String, integer)

  dim dataBlock as string = toolbarData
  '' msgBox(dataBlock)
  dim DATA() as string =toolbarFactory_preproc(dataBlock)
  dim max as integer=DATA.length
  dim OUT(max)
  dim tagData
  redim tagData(max)

  dim isMenuStart as boolean=false 
  dim isMenuMode as boolean=false
  dim isExpanderStartPlus, isExpanderStart, isHidenState as boolean
  dim i, ii, counter, tagIndex as integer
  dim fields() as string
  dim line, mainType, mainTypeUpper, Caption, CaptionTiny, toolTip, subType as string
  
  dim groupMaster, group, hidenOnIni, tagPara, iconURL, iconUrl2, flags, unknown, id as string
  dim rTypeInfo, rPara as string
  dim groupId as string = "" 
  dim groupePreFix as string = "GROUPE_"

  isHidenState=false
  for i = 0  to max-1
    'tagIndex=i
    iconURL = "": iconURL2 = "": tagPara = "": toolTip = "": flags = "": unknown = ""  '...resetOutVar
    SubType = "": Flags = "":   
    ' ...syntaxCheck, parameter zerlegen, ...dann dialogItem erzeugen 
    line=DATA(i)
    if trim(line)="" then continue for
    line = replace(line, "|##|","_##_")+"|||||||"                       '...scriptedPanelId maskieren
    fields=line.split("|")
    mainType=trim(fields(1))
    

    mainTypeUpper=mainType.toUpper
    if trim(mainType)="" then
       dim errMes as string= "FEHLER: keine Type-Information gefunden "
       errMes=errMes+vbNewLine+"... oder Zeilenende/Zeilenfortsetzung fehlerhaft"
       errMes=errMes+vbNewLine+"========================================================================"
       errMes=errMes+vbNewLine
       msgBox (errMes+ "Zeile: "+DATA(i))
       continue for
    end if
    'trace mainType+"<--", "???"
    if mainTypeUpper="#MENUE-START" or mainTypeUpper.startsWith("+++") then
          ''msgBox("menü-start ???")
          'monitorAdd("MENUE-START: "+fields(0))
          isMenuStart=true
          isMenuMode=true

          continue for
    end if
    if mainTypeUpper="#MENUE-START" or mainTypeUpper.startsWith("###") then
          'monitorAdd("MENUE-START: "+fields(0))
          ''msgBox("menü-start ???")
          isMenuStart=true
          isMenuMode=true
          continue for
    end if
    
    if mainTypeUpper="#MENUE-END" or mainTypeUpper.startsWith("XXXXX") then
          ''msgBox("menü-start ???")
          'monitorAdd("MENUE-START: "+fields(0))
          isMenuStart=false
          isMenuMode=false
          ref.ActiveMenu = Nothing
          continue for
    end if
    
    if mainTypeUpper.startsWith("---+") then
          ''msgBox("menü-start ???")
          'monitorAdd("EXPANDER-START-OPEN: "+fields(0))
          isExpanderStart=true
          isHidenState=false
          groupId=groupePreFix+tagIndex.toString
          continue for
    end if
    if mainTypeUpper.startsWith("---") then
          'monitorAdd("EXPANDER-START-CLOSE: "+fields(0))
          ''msgBox("menü-start ???")
          isExpanderStart=true
          isHidenState=true
          groupId=groupePreFix+tagIndex.toString
          continue for
    end if
    if mainTypeUpper.startsWith("===") then
          'monitorAdd("EXPANDER-ENDE: "+fields(0))
         ''msgBox("menü -end")
          isExpanderStart=false        '...macht die Schleife bereits
          isExpanderStartPlus=false    '...macht die Schleife bereits
          isHidenState=false           'WICHTIG
          groupId="NONE"
          continue for
    end if

   'checkbox
    if mainTypeUpper="CHECKBOX" then
      if iconUrl="" then
        iconUrl=iconList(0)
      end if
      if iconUrl2="" then
        iconUrl2=iconList(1)
      end if
    end if
 
 
   'make Id exclusive
    dim idCounter as integer =0
    dim mainTypeLower as string=mainType.toLower
    if exclusiveId.TryGetValue(mainTypeLower, idCounter) = true then
      idCounter=idCounter + 1
      exclusiveId.Item(mainTypeLower)=idCounter
      'increment
    else
      exclusiveId.Add(mainTypeLower, 0)
    end if 
    
    
    '' monitorAdd("exclusiveId.: "+mainType+"_"+idCounter.toString)
    ''if idCounter>1 then
      mainType  = mainType+"_"+idCounter.toString
    ''end if
    
    
    captionTiny=trim(fields(2))
    caption=trim(fields(3))
    if trim(caption)="" then caption=captionTiny
 
    for ii = 5  to fields.length-1
      if fields(ii)="" then continue for
      if fields(ii)="..." then continue for
      splitTypeInfo (trim(fields(ii)),rTypeInfo,rPara)
      
      'trace rTypeInfo.toLower, rPara
      Select Case rTypeInfo.toLower
        'case "#menue-start"
        'case "#menue-end"
        '  isMenuStart=false
        '  ref.ActiveMenu = Nothing
        '  continue for
        case "ico:"
          iconUrl=rPara
        case "ico2:"
          iconUrl2=rPara
        case "tt:"
          toolTip=rPara
        case "t:"
          subType=rPara
        case "p:"
          tagPara=rPara
        case "flags:"
          flags=rPara
        case else
          unknown=unknown+"<²²>"+rTypeInfo+rPara
      end select
    next
    
   '' ...groupManager

   ''monitorAdd("groupId: "+groupId)
   if groupId <> "NONE" then
      group=groupId
    else 
      group=""
    end if

   ''monitorAdd("isExpanderStart: "+isExpanderStart.toString)
    if isExpanderStart = true then
      groupMaster=group
      isExpanderStart = false
    else
      groupMaster=""
    end if

   if isHidenState=true then
      if hidenOnIni="" then 
        hidenOnIni="NEXT"    'erst beim nächsten durchgang zuweisen
      else
        hidenOnIni="HIDDEN"
      end if
    else
      hidenOnIni=""
    end if  


    '' outPara zuweissen
    tagIndex=counter
    
    dim btnPara(20) as String
    btnPara(cIndex) = tagIndex
    btnPara(cMainType) = mainType
    btnPara(cCaption) = caption
    btnPara(cCaptionTiny) = captionTiny
    btnPara(cToolTip) = toolTip
    btnPara(cSubType) = subType
    btnPara(cGroupMaster) = groupMaster
    btnPara(cGroup) = group
    btnPara(cHidenOnIni) = hidenOnIni
    btnPara(cTagPara) = tagPara
    btnPara(cIconURL) = iconURL
    btnPara(cIconURL2) = iconURL2
    btnPara(cFlags) = flags
    btnPara(cFrei1) = "...frei"
    btnPara(cDataPlus) = "PLUS-DATA: "+unknown

   'monitorAdd(btnPara(cIndex)+"---"+groupId+"+++"+btnPara(cGroupMaster)+"==="+btnPara(cGroup))

  
    ' Auf Los gehts los: ...buttons erzeugen 
    
    tagData(counter)=join(btnPara, "<³³>") 
    createToolbarElements(container,btnPara,isMenuStart,isMenuMode)
    if isMenuStart = true then
      isMenuMode=true
      isMenuStart=false
    end if
    
    '...provisorische Ausgabe
    OUT(counter)=join(btnPara, "<--"+vbNewLine) 
    counter=counter+1
  next
  
  redim preserve OUT(counter)
  '...provisorische Ausgabe
  ''msgBox(join(OUT, vbNewLine+"--------------------------------------------------------------"+vbNewLine))
  rGlobTagData=join(tagData,vbNewLine)
  
End Sub




'--> createToolbarElements
: sub createToolbarElements(container as object, btnPara() as string, isMenuStart as boolean, isMenuMode as boolean)
  on error resume next
  dim ref as object=container
  
  dim mainType     as string = btnPara(cMainType)
  dim caption      as string = btnPara(cCaption)
  dim captionTiny  as string = btnPara(cCaptionTiny)
  dim toolTip      as string = btnPara(cToolTip)
  dim subType      as string = btnPara(cSubType)
  dim tagPara      as string = trim(btnPara(cTagPara))
  dim iconURL      as string = btnPara(cIconURL)
  dim iconURL2     as string = btnPara(cIconURL2)
  dim flags        as string = btnPara(cFlags)
  dim frei1        as string = btnPara(cFrei1)
  dim unknown      as string = btnPara(cDataPlus )
  dim hidenOnIni   as string = btnPara(cHidenOnIni)
  dim index as integer
  index=val(iconURL)
  if index>0 then iconUrl = iconList(index) 



  dim el as object
  dim btnParaAsString=join(btnPara, "<³³>")
  btnParaAsString=replace(btnParaAsString,"_##_","|##|")   '...scriptedPanelIds deMaskieren
  if mainType.startsWith("###") or mainType.startsWith("...") then
    el=ref.addLabel  ("", "-")
    el.tag=btnParaAsString
    el.name="btn_"+mainType
    if hidenOnIni="HIDDEN" then el.visible=false
    'monitorAdd("SEPERATOR"+btnParaAsString)
    '' el.checked=true
  else
    if isMenuStart =true then
      el = ref.AddButton(mainType, caption,IconURL:=iconURL, Flags:=ButtonFlags.StartMenu)
      ''el.tag=subType+"<³³>"+tagPara+"<³³>"+unknown
      ''el.tag=subType+"<³³>"+btnParaAsString 
      el.tag=btnParaAsString
      if hidenOnIni="HIDDEN" then el.visible=false
      
      if tagPara<>"" and tagPara.startsWith("callback:")then
        dim forCallBack=trim(mid(tagPara,10))
        '' msgbox(forCallBack)
        ''CallByName(Me, tagPara, Microsoft.VisualBasic.CallType.Method, ref)
        CallByName(Me, forCallBack, Microsoft.VisualBasic.CallType.Method)
      end if
    
    else
      el = ref.AddButton(mainType, caption,IconURL:=iconURL)
      ''el.tag=subType+"<³³>"+btnParaAsString 
      el.tag=btnParaAsString
      el.ToolTipText=caption
      if hidenOnIni="HIDDEN" and isMenuMode=false then el.visible=false
    end if
  end if
End Sub


 
'--> toggleToolBarLabels1()
sub toggleToolBarLabels1(newState)
  on error resume next
  dim toolbarHeight
  dim toolbar =Ref
  if toolbar is nothing then exit sub
  
  dim hideLabels as boolean
  toolbarHeight = toolbar.height

  if toolbarHeight<55 then hideLabels=true
  static toggle as boolean =false
  toggle = not toggle

  toggle = newState
  hideLabels=toggle

  dim max as integer=globTagData1.length
  dim i as integer
  dim item
  dim LINE() as string
  dim caption as string
  dim groupMaster as string
  dim captionTiny as string
  dim elementId  as string
  dim el
  
  for i=0 to max-1
    item =globTagData1(i)
    if item="" then continue for
    
    ''monitorAdd(item)
    LINE = split(item, "<³³>")
    caption=LINE(cCaption)
    
    '' trace caption, item 
    captionTiny = LINE(cCaptionTiny)
    groupMaster=LINE(cGroupMaster)
    'captionTiny=LINE(cCaption2)
    elementID="btn_"+LINE(cMainType)
    'monitorAdd("elementID: ", elementID)
    el=ref.element(elementID)
    if el is nothing then                     'sonderfallMenü
      '' elementID="mnu_"+LINE(cMainType)
      '' el=ref.element(elementID)
      continue for
    end if 
    if hideLabels = true then
      el.text=captionTiny+""
      el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
    else
      el.text="      "+caption
      el.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
      el.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay
      el.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
      el.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)

    end if
    if groupMaster<>"" then
      el.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
      el.text=""+caption
      el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
      el.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
      el.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)
    end if
   next
end sub


 
'--> toggleToolBarLabels2()
  sub toggleToolBarLabels2(newState)
  on error resume next
  dim toolbarHeight
  dim toolbar =Ref
  if toolbar is nothing then exit sub
  
  dim hideLabels as boolean
  toolbarHeight = toolbar.height

  if toolbarHeight<55 then hideLabels=true
  static toggle as boolean =false
  toggle = not toggle

  toggle = newState
  hideLabels=toggle

  dim max as integer=globTagData2.length
  dim i as integer
  dim item
  dim LINE() as string
  dim caption as string
  dim groupMaster as string
  dim captionTiny as string
  dim elementId  as string
  dim el
  
  for i=0 to max-1
    item =globTagData2(i)
    if item="" then continue for
    
    ''monitorAdd(item)
    LINE = split(item, "<³³>")
    caption=LINE(cCaption)
    
    '' trace caption, item 
    captionTiny = LINE(cCaptionTiny)
    groupMaster=LINE(cGroupMaster)
    'captionTiny=LINE(cCaption2)
    elementID="btn_"+LINE(cMainType)
    'monitorAdd("elementID: ", elementID)
    el=ref2.element(elementID)
    if el is nothing then                     'sonderfallMenü
      '' elementID="mnu_"+LINE(cMainType)
      '' el=ref.element(elementID)
      continue for
    end if 
    if hideLabels = true then
      el.text=caption+""
      el.ImageAlign = System.Drawing.ContentAlignment.MiddleCenter
      el.TextDirection = System.Windows.Forms.ToolStripTextDirection.Vertical270
      el.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage

    else
      el.text="      "+caption
      el.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
      el.TextAlign = System.Drawing.ContentAlignment.MiddleRight
      el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay
      el.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
      el.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)

    end if
    if groupMaster<>"" then
      el.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
      el.text=""+caption
      el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
      el.Margin = New System.Windows.Forms.Padding(0, 0, 0, 0)
      el.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)
    end if
   next
end sub



'-->
sub onInitialize()
  '...
End Sub


sub onTerminate(optional intern as string="")
   : Timer1.Enabled = false
   : dim extern as string =""
   : if intern="" then
   :    extern ="=================EXTERN===: "+name
   :    intern=""
   :  else
   :  intern=intern+": "+name
   : end if
   : trace "onTerminate: "+intern,"ON-TERMINATE "+extern
   : dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
   : dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
   : removeHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged
  'ZZ.IDEHelper.RemoveToolbar("toolbar.tb1")
End Sub


'--> 
'--> ICON-LIST
:sub createIconList()
  on error resume next

'###################################

  '...Statposition festlegen
  dim mainForm as object = ZZ.IDEHelper.MainWindow
  dim startPosX as integer
  dim startPosY as integer
    if mainForm is nothing then
     startPosY=0
     startPosX=500
    else
      startPosX = mainForm.left+ mainForm.width +600' -800
     startPosX = mainForm.left+ mainForm.width -800
      startPosY = mainForm.Top+ 00
    end if
  '...1. Fenster erzeugen
  PnlRef = ZZ.IDEHelper.CreateDockWindow(appId+".Icon-List", 1)
  dim el as object
  with PnlRef
    .ResetControls(25, 20)
    with getOuterWindowRef(PnlRef)': err.number=0
      .top=startPosY
      .left=startPosX
      .height=600
      .width=450
    end with
    '...ACHTUNG: die Ausgabe (z.B. fenster-Hintergrundfarbe) erfolg nicht auf der Form                   
    '...sondern auf einem panel, das die gesammte Form ausfüllt.                                         
    '... deshalb darf z.B. für die Fenster-Hintergrund-Farbe NICHT getOuterWindowRef verwendet werden    
    .backColor=ColorTranslator.FromHtml("#394242") 
    .addIcon("myIcon", "http://es.teamwiki.net/docs/icons/graphic_design2.png", , , -2)
    .insertX = 160:.insertY = 15:
    el=.addLabel("myLabel", " Die Icon-Liste")
    el.Font = New System.Drawing.Font("Courier New", 20.0!, System.Drawing.FontStyle.bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#CFD4D4")
    
     .insertX = 160:.insertY = 50:
    dim mes as string=" ...soll die Auswahl des richtigen Icons für die userToolbar vereinfachen."
    mes =mes+vbNewLine+vbNewLine+"Bei Click auf ein Icon wird die Icon-Nummer in das Clipboard kopiert."
    el=.addLabel("myLabel2", mes)
    
    '??? el.multiline=true
    el.Font = New System.Drawing.Font("Courier New", 11.0!, System.Drawing.FontStyle.regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#CFD4D4")
    el.autosize=false
    el.width=260
    el.height=140
    
    insertIconsFromIconList(PnlRef)
    PnlRef.show()
    PnlRef.parent.parent.show()
  end With
  
  
''   
''   '... 2.Fenster erzeugen
''   PnlRef2 = ZZ.IDEHelper.CreateDockWindow(appId+".Team", 1)
''   with PnlRef2
''     .ResetControls(25, 20)
''     with getOuterWindowRef(PnlRef2)': err.number=0
''       .left=startPosX-25
''       .top=startPosY+200
''       .height=185
''       .width=500
''     end with
''     .backColor=ColorTranslator.FromHtml("#394242")
''     .addIcon("myIcon", "http://icons3.iconfinder.netdna-cdn.com/data/icons/Siena/128/globe%20yellow.png", , , -2)
''     .insertX = 160:.insertY = 120:
''     el=.addLabel("myLabel", "hello, world 2")
''     el.Font = New System.Drawing.Font("Courier New", 20.0!, System.Drawing.FontStyle.bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
''     el.foreColor=ColorTranslator.FromHtml("#FEC408")
''   end With
end sub


sub insertIconsFromIconList(PnlRef)
  dim index, left, top, deltaTop, deltaLeft as integer
  dim text, icon  as string
  dim el as object
  dim bgColor="#DEECFF"
  deltaTop=200
  deltaLeft=20
  for index =0 to 50
    text=index.toString
    icon=iconList(index)
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    if icon="xxx" then 
      icon=""
      text="<"+index.toString+">"
    end if
    left=(index mod 10)*40
    top=(index - (index mod 10)) *4
    ''monitorAdd(top.toString, left.tostring)
    el=PnlRef.addButton("iconList_"+text,text,bgColor,Left+deltaLeft,Top+deltaTop,40 ,40 ,icon) ' ,flags,handler)
    el.ImageAlign = System.Drawing.ContentAlignment.TopCenter
    el.TextAlign = System.Drawing.ContentAlignment.BottomCenter
    el.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay 'ImageAboveText
    'el.foreColor=ColorTranslator.FromHtml("#999")
    el.ForeColor = System.Drawing.Color.Transparent
  next
end sub


sub insertIconsFromIndex(ref as object,index as integer, left as integer, top as integer)
  
  dim bgColor as string = ""
  dim txt as string=val(index)
  left=(index mod 10)*30
  top=(index - (index mod 10)) *40
  
  
  'ref.addButton(id,txt,bgColor,Left,Top,30 ,40 ,iconList(index)) ' ,flags,handler)

  dim el as object
  with PnlRef
    .insertX = 10:.insertY = 150:
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
    el.ForeColor = System.Drawing.Color.Transparent
  end with

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
'--> AUTOSTART + DATA-BLOCK

'http://leftlogic.com/lounge/articles/entity-lookup/

    '' AddHandler TT.TraceWrite, AddressOf OnTrace
    '' AddHandler TT.PrintLineChanged, AddressOf OnPrintLine
    '' 
'' getDockPanelRef("")    '... schaltet nicht mehr in vordergrund
'' createDockPanelRef("")




Sub AutoStart() '--------------------------------------------
  'msgBox("AutoStart()")
  monitorAdd("AutoStart: 1 --------------------------------------")
      
  'zz.traceClear
  onTerminate("-----------------------------INTERN---"): Err.Number=0
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  mainWin.SuspendLayout()

  Ref = ZZ.IDEHelper.CreateToolbar(ToolbarID1)
  ref.SuspendLayout()

  Ref2 = ZZ.IDEHelper.CreateToolbar(ToolbarID2)
  ref2.SuspendLayout()

  dim el as object
  iniIconList()

  '' With Ref
  ''   dim tagData1 as string
  ''   .ResetControls()
  ''   dim dataBlock1 as string = getToolbarData1()
  ''   toolbarFactory(ref,dataBlock1,tagData1) '...DatenBlock holen und verwursteln
  ''   'msgbox(tagData1.length)
  ''   globTagData1=split(tagData1, vbNewLine)
  ''   toggleToolBarLabels1(true)
  '' end with
  
  With Ref2
    dim tagData2 as string
    .ResetControls()
    dim dataBlock2 as string = getToolbarData2()
    toolbarFactory(ref2,dataBlock2,tagData2) '...DatenBlock holen und verwursteln
    'msgbox(tagData2.length)
    globTagData2=split(tagData2, vbNewLine)
    toggleToolBarLabels2(true)
  end with
  
 'spezialEvent aktivieren, mit timer entkoppelt 
 'ACHTUNG: zur zeit noch etwas empfindlich auf Fehler
  dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
  dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
  'addHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged

  'Timer1 = New System.Windows.Forms.Timer()
  'Timer1.Enabled = true
  'Timer1.Interval = 111
  dim menuItem=ref.element("btn_xxxDynaMenu_0")
  

  
  '' ref.ResumeLayout(False)
  '' ref.PerformLayout()
  ref2.ResumeLayout(False)
  ref2.PerformLayout()
  mainWin.ResumeLayout(False)
  mainWin.PerformLayout()
  
End Sub


Function getToolbarData2() as string
dim curEditorLine as integer=zLN
''winSpy: C:\yEXE\WinSpy.exe

#Data myMenu as String
============================================================================================================================================
'TYPE  |CAPTION|LEER/kom  |tt: = toolTipText     |p: = PARA                                  |ico: = IconURL |xyz: Beliebig _  '...Zeilenfortsetzung
'                                                                                            |ico???:IconURL _    
' p=panel, e=exe, u=url, o=optionen, x=sonder                                                                                           |ico2:IconURL _
' andere typen als btn: #lab, #img, #chk, #txt                                                                                           |flags:ABC   ||||  

'--> ...Ansicht
'##########################################################################################################################
togglePane      | |ein-/aus-blenden           ||tip:         |ico:034       |p:10              
toggle          | |Sync              |tip:          |ico: 25 |t:pnl |p:Addin|##|siaScriptSyncMini|##|ScriptSyncMini          
toggle          | |localeFiles       |tip:                   |t:pnl |p:Addin|##|siaCommonProtocols|##|LocalFileBrowser        |ico:http://es.teamwiki.net/docs/icons/folder_open.png    
toggle          | |Ftp               |tip:                   |t:pnl |p:Addin|##|siaCommonProtocols|##|FTPBrowser              |ico:http://mw.teamwiki.net/docs/img/win-icons/Microsoft_883-16.png    
toggle          | |rtf               |tip:                   |t:pnl |p:Addin|##|siaRTF|##|TwAjaxBrowser                       |ico:http://mw.teamwiki.net/docs/img/win-icons/SHELL32_2-16.png   
toggle          | |Solution          |tip:                   |t:pnl |p:Addin|##|siaSolution|##|SolutionExplorer               |ico:http://es.teamwiki.net/docs/icons/navigator-insert-as-link.png     
cmdOpenExplorer | |explorer          |tip:                   |t:xxx |p:                         |ico:http://es.teamwiki.net/docs/icons/folder_open.png 
'' 
''http://es.teamwiki.net/docs/icons/16-clock.png
============================================================================================================================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=curEditorLine & vbNewLine & myMenu  
  return RESULT
end function



Function getToolbarData1() as string
dim curEditorLine as integer=zLN
''winSpy: C:\yEXE\WinSpy.exe

#Data myMenu as String
============================================================================================================================================
'TYPE  |CAPTION|LEER/kom  |tt: = toolTipText     |p: = PARA                                  |ico: = IconURL |xyz: Beliebig _  '...Zeilenfortsetzung
'                                                                                            |ico???:IconURL _    
' p=panel, e=exe, u=url, o=optionen, x=sonder                                                                                           |ico2:IconURL _
' andere typen als btn: #lab, #img, #chk, #txt                                                                                           |flags:ABC   ||||  

'--> ...Ansicht
'##########################################################################################################################
''caption       |Ansicht|          ||tip:         |xxx ico:001   |p:               

cmdToggleTestLab| testLab       ||tip:         |xxx ico:http://es.teamwiki.net/docs/icons/ledlightblue.png       |p:ToolBar|##|tbScriptWin|##|es_testLabor.mainWin                                                           |xxx ico:
cmdToggleWeb    | = WEB =       ||tip:         |xxx ico:http://es.teamwiki.net/docs/icons/ledlightblue.png       |p:ToolBar|##|tbScriptWin|##|es_webbrowser3.tab                                                           |xxx ico:
.....
-----
togglePane      |fullScreen|                ||tip:         |xxx ico:032       |p:-1               
.....
togglePane      | |                ||tip:         |ico:033       |p:8               
togglePane      | |                ||tip:         |ico:034       |p:10              
.....
togglePane      | |                ||tip:         |ico:035       |p:7                
togglePane      | |                ||tip:         |ico:36        |p:9             
=====  
.....
-----
caption          |Links            ||tip:         |xxx ico:      |xxx p:ToolBar|##|tbScriptWin|##|es_webbrowser3.tab                                                           |xxx ico:
''caption        |Web              ||tip:         |xxx ico:      |p:                                                          
url              |...ZURÜCK        ||tip:         |xxx ico:      |p:http://es.teamwiki.net/script-ide-01.html?view=3           
url              |google           ||tip:         |xxx ico:      |p:http://www.google.de/                                      
url              |iconfinder       ||tip:         |xxx ico:      |p:http://www.iconfinder.net/icondetails/34206/16/            
url              |codeProj.        ||tip:         |xxx ico:      |p:http://www.codeproject.com/info/search.aspx                
url              |Max              ||tip:         |xxx ico:      |p:http://mw.teamwiki.net/                                    
url              |Manuel           ||tip:         |xxx ico:      |p:http://mk2.teamwiki.net/                                   
url              |Elmar            ||tip:         |xxx ico:      |p:http://es.teamwiki.net/                                    
=====
.....
-----
caption          |IDE              ||tip:         |ico:040       |p:   |ico: http://es.teamwiki.net/docs/icons/folder_open.png                        
cmdOptions       |Fenster          ||tip:         |ico:041       |p:fensterverwaltung          
cmdOptions       |addins           ||tip:         |ico:042       |p:addins                     


'http://es.teamwiki.net/docs/icons/music_note.png



+++++++++++++++++ ' MENÜ-START
mnu             |menüTest|             |tip:          |p:        | ico:http://es.teamwiki.net/docs/icons/starBlue16.gif
checkbox        |hier kann man dann längere Texte eingeben |      |p:ToolBar|##|tbconsole   
checkbox        |chkBox |              |tip:           |p:ToolBar|##|tbconsole 
mnu             |musik |               |tip:        |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe    |ico:http://es.teamwiki.net/docs/icons/music_note.png
mnu             |Video |               |tip:        |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe   
mnu             |. . . |               |tip:        |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe  
mnu             |. . . |               |tip:        |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             
xxxxxxxxxxxxxxxxx 'MENÜ-ENDE

cmdOptions      |Skins|                ||tip:         |p:siaskinnable.ctl_options   |xxx ico:http://es.teamwiki.net/docs/icons/color-swatch.png  
cmdToolbarMan   |toolBars|             ||tip:         |p:                           |xxx ico:http://es.teamwiki.net/docs/icons/color-swatch.png  
cmdIconMan      |icons|                ||tip:         |p:                           |xxx ico:http://es.teamwiki.net/docs/icons/color-swatch.png  
checkBox        |links ? |             ||tip:         |p:ToolBar|##|tbconsole 
checkBox        |oben |                |tip:         |p:ToolBar|##|tbconsole  
checkBox        ||Check-1            ||tip:         |p:ToolBar|##|tbconsole  
checkBox        ||Check-2            ||tip:         |p:ToolBar|##|tbconsole  
=====
.....
'--> ...files
'' ---
'' caption         |Files               ||tip:                     |xxx ico:http://es.teamwiki.net/docs/icons/folder_open.png 
'' toggle          | |Sync              |tip:          |ico: 25 |t:pnl |p:Addin|##|siaScriptSyncMini|##|ScriptSyncMini          
'' toggle          | |Solution          |tip:                   |t:pnl |p:Addin|##|siaSolution|##|SolutionExplorer               |ico:http://es.teamwiki.net/docs/icons/navigator-insert-as-link.png     
'' toggle          | |localeFiles       |tip:                   |t:pnl |p:Addin|##|siaCommonProtocols|##|LocalFileBrowser        |ico:http://es.teamwiki.net/docs/icons/folder_open.png    
'' toggle          | |Ftp               |tip:                   |t:pnl |p:Addin|##|siaCommonProtocols|##|FTPBrowser              |ico:http://mw.teamwiki.net/docs/img/win-icons/Microsoft_883-16.png    
'' toggle          | |rtf               |tip:                   |t:pnl |p:Addin|##|siaRTF|##|TwAjaxBrowser                       |ico:http://mw.teamwiki.net/docs/img/win-icons/SHELL32_2-16.png   
'' cmdOpenExplorer | |explorer          |tip:                   |t:xxx |p:                         |ico:http://es.teamwiki.net/docs/icons/folder_open.png 
'' exe             | |EXE               |tip:          |t:url |p:xxx                      |ico:http://es.teamwiki.net/docs/icons/16-circle-blue-add.png
'' =====
'' .....
---+ ' GROUP-START
toggle         |DEBUG              ||tip:         |t:exe |                            |xxx ico:http://es.teamwiki.net/docs/icons/starGreen16.gif
'exec          |R U N|             |tip:          |t:opt |p:Debug.Run                   |ico:http://es.teamwiki.net/docs/icons/tick-button.png  
'.....
saveRun           | = RUN =|        |tip:          |t:opt |p:Debug.Run                   |ico:http://es.teamwiki.net/docs/icons/tick-button.png 
checkbox       |clear|           | |tip:          |p:        |xxx ico:  
checkbox       |dump|            | |tip:          |p:        |xxx ico:
..... 
toggle         | |trace            |tip:          |t:opt |p:Addin|##|siaCodeCompiler|##|SHDebug                    |ico:http://es.teamwiki.net/docs/icons/stock_macro-insert-breakpoint.png  
''toggle       | |Debug            |tip:          |t:exe |p:ToolBar|##|tbdebug                                    |ico:http://es.teamwiki.net/docs/icons/alert_gelb.png
toggle         | |Err              |tip:          |t:url |p:Addin|##|siaCodeCompiler|##|SHCompilerErrors           |ico:http://es.teamwiki.net/docs/icons/alert_gelb.png 
toggle         | |Cmd              |tip:          |t:url |p:Addin|##|ScriptIDE.Main|##|Console                   |ico:http://es.teamwiki.net/docs/icons/text-x-script.png
exe            | |Trace            |tip:          |t:exe |p:C:\yEXE\traceMonitor.exe                |ico:http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png
exe            |interProc|         |tip:           |t:exe |p:C:\yEXE\ide3\IprocViewer.exe                |ico:http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png
===== ' GROUP-END
.....
--- ' GROUP-START(expanded)
caption        |DEV                 ||tip:          |t:pnl |p:    |xxx ico:http://es.teamwiki.net/docs/icons/tools_blau.gif       
toggle         | |Snippets           |tip:          |t:pnl |p:ToolBar|##|tbScriptWin|##|es_snippetManager.mainWin                |ico:http://es.teamwiki.net/docs/icons/script_(start)_16x16.gif
toggle         | |Reflection         |tip:          |t:pnl |p:Addin|##|siaCodeCompiler|##|ReflectionInfo                 |ico:http://es.teamwiki.net/docs/icons/stock_navigator-references.png    
toggle         | |Suche              |tip:          |t:pnl |p:Addin|##|ScriptIDE.Main|##|GlobSearch                              |ico:http://es.teamwiki.net/docs/icons/Find_fernglas.png
toggle         | |testLabor          |tip:          |t:pnl |p:ToolBar|##|tbScriptWin|##|es_testLabor.mainWin             |ico:http://es.teamwiki.net/docs/StartupWizard.png        
toggle         | |regExpress         |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe                    |ico:http://es.teamwiki.net/docs/icons/page_white_code_red.png
exec           | |Konverter          |tip:          |t:exe |p:Tools.ConvertFile                    |ico:http://es.teamwiki.net/docs/icons/arrowCcurved.png
toggle         | |color              |tip:          |t:exe |p:ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_colorPicker.dll|##|sw_colorPicker.sg_colorPicker   |ico:http://es.teamwiki.net/docs/icons/rainbow.png  
===== ' GROUP-END
.....
----- ' GROUP-START
toggle         |TOOLs         ||tip:      |t:exe   |                                                      |xxx ico:http://es.teamwiki.net/docs/icons/starGreen16.gif
toggle         |memo|         |tip:          |t:exe |p:ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo                   |xxx ico:http://es.teamwiki.net/docs/icons/arrowCcurved.png
toggle         |radio|        |tip:          |t:exe |p:ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_mediaPlayere.dll|##|DOM_Player_1._0.sw_mediaPlayer |xxx ico:http://es.teamwiki.net/docs/icons/arrowCcurved.png
toggle         |clock|        |tip:          |t:exe |p:ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock                 |xxxico:http://es.teamwiki.net/docs/icons/arrowCcurved.png
exe            |clipView-1   ||          |tip:    |p:C:\Windows\system\ZZ_UGB\vbDragTheClipboard.exe     |xxx ico:http://es.teamwiki.net/docs/icons/stock_macro-insert-breakpoint.png  
exe            |clipView-2   ||          |tip:    |p:C:\Windows\system\ZZ_UGB\vbClipViewer.exe           |xxx ico:http://es.teamwiki.net/docs/icons/alert_gelb.png
exe            |taskList     ||          |tip:    |p:C:\Windows\system\ZZ_UGB\taskList.exe               |xxx ico:http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png
exe            |fileInfo     ||          |tip:    |p:C:\Windows\system\ZZ_UGB\FileVersionInfo.exe        |xxx ico:http://es.teamwiki.net/docs/icons/alert_gelb.png 
exe            |ajax         ||          |tip:    |p:C:\yEXE\TwAjaxDebug.exe                             |xxx ico:http://es.teamwiki.net/docs/icons/text-x-script.png
===== ' GROUP-END





''weitere tools:
     ''c:\yEXE\TwAjaxDebug.exe
     ''C:\Windows\system\ZZ_UGB\BlueScreen.exe
     ''C:\Windows\system\ZZ_UGB\vbDragTheClipboard.exe
     ''C:\Windows\system\ZZ_UGB\vbClipViewer.exe
     'C:\Windows\system\ZZ_UGB\taskList.exe
     ''???C:\Windows\system\ZZ_UGB\FileVersionInfo.exe
     ''diverse colorPickers
     ''ajaxDebug
     ''update
     ''folderSync

'--> ...news
.....

-----
caption          |NEWS          ||tip:            |t:opt |p:2              |xxx ico:http://es.teamwiki.net/docs/icons/recherche.gif
url              | google        ||tip:            |p:http://news.google.de/news?pz=0&hl=de&ned=de                                       |xxx ico:
url              | heise         ||tip:            |p:http://www.heise.de/newsticker/                  |xxx ico:
url              | ARD           ||tip:            |p:http://www.tagesschau.de/          |xxx ico:
url              | Wetter        ||tip:            |p:http://www.donnerwetter.de/region/suchort.mv?x=0&y=0&search=gie%DFen           |xxx ico:
=====
.....
+++++++++++++++++ ' MENÜ-START
cmdAddinConf2    |PAUSE |         |tip:          |t:opt |p:             |xxx ico:http://es.teamwiki.net/docs/icons/starBlue16.gif
toggle           |musik |               |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/music_note.png
toggle           |Video |               |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/music_note.png
toggle           |. . . |               |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/music_note.png
toggle           |. . . |               |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/music_note.png
xxxxxxxxxxxxxxxxx 'MENÜ-ENDE
'' 
.....
+++++++++++++++++ ' MENÜ-START
DynaMenu     |exeFiles ||            |tip:          |t:opt |p:callback:  addExeFilesToMenu              |xxx ico:http://es.teamwiki.net/docs/icons/starBlue16.gif
xxxxxxxxxxxxxxxxx 'MENÜ-ENDE
.....
---+ ' GROUP-START
toggle         |toolBar           ||tip:              |xxx ico:http://es.teamwiki.net/docs/icons/starGreen16.gif
test1          |design(1)|      | |tip:    |p:        |xxx ico:
createIconList |iconList|       | |tip:    |p:        |xxx ico:
checkbox       |oben|           | |tip:    |p:        |xxx ico:  
checkbox       |links|          | |tip:    |p:        |xxx ico:
===== ' GROUP-END


 

'' 
'' toggle           | lupe       |         |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/zoom_in16.png
'' toggle           | Clock      |         |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/16-clock.png
'' toggle           | Memo       |         |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe             |ico:http://es.teamwiki.net/docs/icons/edit_16x16.gif
'' 
'' ''|ico:http://mw.teamwiki.net/docs/img/win-icons/CSCDLL_143-16.png      
'' 
'' nav             |Links|                |tip:          |t:url |p:http://devnet.teamwiki.net/xymemo.html?doc=es-pb-mw-script-ide02     |ico:http://mw.teamwiki.net/docs/img/win-icons/iexplore_32542-16.png
'' .....
'' 

'' exe           |lupe |                |tip:          |t:exe |p:C:\Windows\system\ZZ_UGB\BlueScreen.exe   _
''                                                                   |ico:ico:http://es.teamwiki.net/docs/icons/block.png _
''                                                                   |ico2:http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png _
''                                                                   |ico3:http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png _
''                                                               |...|SPLASH:http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png
'' toggle          ||...leer              |tip:          |t:url |p:xxx            |
'' toggle          |44|...leer            |tip:          |t:url |p:xxx            |ico:http://es.teamwiki.net/docs/icons/star.png 
'' 
'' toggle          |c|...leer             |tip:toolTipText   |t:url |p:xxx     
'' toggle          |a|...leer             |tip:toolTipText   |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/16-tool-a.png
'' toggle          |11|...leer            |tip:toolTipText   |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/chevron-expand.png  
'' toggle          ||...leer              |tip:toolTipText     |t:url |p:xxx   |
'' toggle          ||...leer              |tip:toolTipText     |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/dummy.png 
'' toggle          ||...leer              |tip:toolTipText     |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/stock_macro-insert-breakpoint.png  
'' .....
'' toggle          ||...leer             |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/settings2_16x16.gif 
'' toggle          ||...leer             |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/tools_blau.gif 
'' toggle          |88|...leer           |tip:toolTipText  |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/script16x16.gif
'' toggle          ||...leer             |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/text-x-script.png
'' toggle          ||...leer             |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/stock_navigator-references.png
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/minus16.png
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/lifebuoy.png 
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:ico:http://es.teamwiki.net/docs/icons/toggle.png
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/Wizard111.png
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/star_black.png 
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:http://es.teamwiki.net/docs/icons/starRed16.gif   
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:
'' toggle          |x|...leer            |tip:toolTipText    |t:url |p:xxx   |ico:
''  
'' 
''http://es.teamwiki.net/docs/icons/16-clock.png
============================================================================================================================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=curEditorLine & vbNewLine & myMenu  
  return RESULT
end function


Function iniIconList() as string
dim curEditorLine as integer=zLN
#Data myIconList as String
''============================================================================================================================================
 000 | http://es.teamwiki.net/docs/icons/checkbox_yes.png 
 001 | http://es.teamwiki.net/docs/icons/checkbox_no.png 
'' 002 | http://es.teamwiki.net/docs/icons/checkbox_no.png
 003 | http://es.teamwiki.net/docs/icons/toggle.png 
 004 | http://es.teamwiki.net/docs/icons/toggle_collapse.png
 005 | http://es.teamwiki.net/docs/icons/chevron-expand.png
 006 | http://es.teamwiki.net/docs/icons/chevron-collapse.png  
 007 | http://es.teamwiki.net/docs/icons/cross-white.png
 007 | http://es.teamwiki.net/docs/icons/plus-white.png
 008 | http://es.teamwiki.net/docs/icons/zoom_in16.png 
 009 | xxx 


 010 | http://es.teamwiki.net/docs/icons/navigation-000-white.png 
 011 | http://es.teamwiki.net/docs/icons/navigation-090-white.png 
 012 | http://es.teamwiki.net/docs/icons/navigation-180-white.png
 013 | http://es.teamwiki.net/docs/icons/navigation-270-white.png 
 014 | xxx 
 015 | http://es.teamwiki.net/docs/icons/navigation-090-button.png 
 016 | http://es.teamwiki.net/docs/icons/navigation-180-button.png 
 017 | http://es.teamwiki.net/docs/icons/navigation-270-button.png 
 017 | xxx 
 018 | xxx 
 019 | xxx 

 020 | http://es.teamwiki.net/docs/icons/arrow-retweet.png
 021 | http://es.teamwiki.net/docs/icons/arrow-resize.png
 022 | http://es.teamwiki.net/docs/icons/arrow-circle.png
 023 | http://es.teamwiki.net/docs/icons/arrow-move.png 
 024 | http://es.teamwiki.net/docs/icons/arrow_two_head_2.png 
 025 | http://es.teamwiki.net/docs/icons/arrowRightLeft2Red.png 
 026 | xxx 
 027 | xxx 
 027 | xxx 
 028 | xxx 
 029 | xxx 

 030 | xxx
 031 | http://es.teamwiki.net/docs/icons/gnome-dev-computer.png  
 032 | http://es.teamwiki.net/docs/icons/layer-resize.png
 033 | http://es.teamwiki.net/docs/icons/arrow-180.png 
 034 | http://es.teamwiki.net/docs/icons/arrow_blau_rechts.png
 035 | http://es.teamwiki.net/docs/icons/arrow-090.png 
 036 | http://es.teamwiki.net/docs/icons/arrow-270.png
 037 | xxx 
 037 | xxx 
 038 | xxx 
 039 | xxx 

 040 | http://es.teamwiki.net/docs/icons/starGreen16.gif
 041 | http://es.teamwiki.net/docs/icons/hammer-screwdriver.png
 042 | http://mw.teamwiki.net/docs/img/win-icons/vsmacros_216-32.png
 043 | xxx 
 044 | xxx 
 045 | xxx 
 046 | xxx 
 047 | xxx 
 047 | xxx 
 048 | xxx 
 049 | xxx 


 050 | xxx 
 051 | xxx 
 052 | xxx 
 053 | xxx 
 054 | xxx 
 055 | xxx 
 056 | xxx 
 057 | xxx 
 057 | xxx 
 058 | xxx 
 059 | xxx 

 060 | xxx 
 061 | xxx 
 062 | xxx 
 063 | xxx 
 064 | xxx 
 065 | xxx 
 066 | xxx 
 067 | xxx 
 067 | xxx 
 068 | xxx 
 069 | xxx 

 '' 0 | xxx 
 '' 1 | xxx 
 '' 2 | xxx 
 '' 3 | xxx 
 '' 4 | xxx 
 '' 5 | xxx 
 '' 6 | xxx 
 '' 7 | xxx 
 '' 7 | xxx 
 '' 8 | xxx 
 '' 9 | xxx 
 '' 


''============================================================================================================================================
#EndData
  '' Trick 17: Zeilennummer dazupacken
  dim RESULT as string=curEditorLine & vbNewLine & myIconList
  redim iconList(111)
  dim DATA() as string=split(RESULT,vbNewLine)
  dim LINE() as string
  dim temp as string
  dim index as integer
  dim icon as string
  dim max as integer=DATA.length
  dim i as integer
  for i =1 to max-1
    temp=trim(DATA(i))
    if temp="" then continue for
    if temp.startsWith("'") then continue for
    LINE = split(DATA(i), "|")
    index = val(LINE(0))
    icon=trim(LINE(1))
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    'if icon="xxx" then icon="http://es.teamwiki.net/docs/icons/edit-number.png"
    '...check auf mehrfache Belegung fehlt noch
    iconList(index)=icon
  next
  return RESULT
end function


'... der ist noch nicht aktiviert
sub addExeFilesToMenu '' (ref as object)
  ''msgBox("ich bin die proc: addExeFilesToMenu")
  dim root as string="C:\yPara\scriptIDE\compiledScripts\"
  Dim finder = ZZ.OpenFileFinder(root, "*.*")
  Dim RESULT() = Split(finder.FindRecursive(), vbNewLine)
  dim el as object
  dim tagDATA(20) as string
  For Each Line As String In RESULT
    If Line.toString="" Then Continue For
    Dim Parts() as string = Split(Line, vbTab)
    dim fileName as string=Parts(1)
    If fileName.ToLower().EndsWith(".exe") Then 
    dim subFolder=IO.Path.GetFilenameWithoutExtension(fileName)+"\"
    'monitorAdd(line)
    'ref.AddButton("mnuexe_"+fileName, fileName)
    el=ref.AddButton("mnuexe_"+fileName, fileName)
    tagDATA(cTagPara)=root +subFolder+ fileName
    el.tag=join(tagDATA,"<³³>")
    end if
  Next
end sub






'-->
'--> E V E N T S

Sub onButtonEvent(e)
  ref.ResumeLayout(False)
  ref.PerformLayout()
  monitorClear
  '#######################################
  ''msgBox (e.keyString)
  '' MAX: ich brauch die MouseKnöpfe
  '#######################################
  errorText("")
  printLine 11, "" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim tag as string = e.sender.tag
  dim tagDATA()= Split(tag, "<³³>")
  dim action as string =name(1)
  monitorAdd("==============================================")
  monitorAdd("Sender.Name: " + e.Sender.Name+"<--")
  monitorAdd("")
  monitorAdd("      action: " +action+" <==")
  monitorAdd("     tagPara: " +tagDATA(cTagPara)+" <<<")
  monitorAdd(" MouseButton: " +e.MouseButton+" <<<")
  monitorAdd("   ClassName: " +e.ClassName+" <<<")
  monitorAdd(" ControlType: " +e.ControlType+" <<<")
  monitorAdd("      MouseX: " +e.MouseX.toString+" <<<")
  monitorAdd("==============================================")
  dim i as integer=0
  for i=0 to tagDATA.length-1
    'monitorAdd(cName(i),tagDATA(i))
  next
  monitorAdd("==============================================")

'monitorAdd(join(tagDATA,vbNewLine))
  'printLine 11, "" , tagDATA(0)

  '....SonderFall
  if action.toLower="togglepane" then
    onTogglePane          (e, tagDATA)
    exit sub
  end if

  checkForExpanderState(e, tagDATA)

  Select Case action.toLower
    case "togglepane"       : onTogglePane          (e, tagDATA)
    case "url"              : onNavigateUrl         (e, tagDATA)
    case "cmdoptions"       : onCmdOptions          (e, tagDATA)
    case "cmdopenexplorer"  : onCmdOpenExplorer     (e, tagDATA)

    Case "toggle"           : onToggleDockWindow    (e, tagDATA)
    Case "cmdtoggleweb"     : oncmdToggleWeb        (e, tagDATA)
    Case "cmdtoggletestlab" : onCmdToggleTestLab    (e, tagDATA)
    Case "saverun"          : onSaveRun             (e, tagDATA)
    Case "test1"            : ontest1               (e, tagDATA)
    Case "exec"             : onExecute             (e, tagDATA)
    Case "exe"              : onExe                 (e, tagDATA)
    case "mnuexe"           : onExe                 (e, tagDATA)
    Case "nav"              : onNavigateUrl         (e, tagDATA)
    Case "toggleimg"        : onToggleImg           (e, tagDATA)
    Case "checkbox"         : onToggleImg           (e, tagDATA)
    case "createIconList"   : onCreateIconList      (e, tagDATA)
    'Case "expand"          : onToggleDockWindow    (e, tagDATA)
    'Case "expand"          : toogleToolbarItems    (e, tagDATA)
    'Case "switch"          : onSwitch              (e, tagDATA)
    
    Case "cmdtoolbarman" :onCmdToolbarMan          (e, tagDATA)
    Case "cmdiconman"    :onCmdIconMan             (e, tagDATA)
    Case "item"            '... nichts weiter machen
    Case "caption"       '... nichts weiter machen
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
  
  checkDisplayState(el)
  ref.ResumeLayout(False)
  ref.PerformLayout()
End Sub


:sub checkForExpanderState(e,tagDATA)
  on error resume next
  ref.SuspendLayout()
  monitorAdd("checkForExpanderState ...")
  ''monitorAdd("xxxxxxxxxxxx tag: "+tagDATA(cGroup))
  if e.sender.tag is nothing then exit sub

  :dim tagIndex as integer = cInt(tagDATA(cIndex)):err.number=0  'z.b. wenn leer
  'monitorAdd("xxxxxxx tagIndex: "+tagIndex.toString)
  'monitorAdd("xxx cGroupMaster: "+tagDATA(cGroupMaster))
  'monitorAdd("xxxxxxxxx cGroup: "+tagDATA(cGroup))
  if tagDATA(cGroupMaster) = "" then exit sub
  
  monitorAdd("TREFFER: ...checkForExpanderState")
  dim i as integer
  dim LINE() as string
  dim groupId as string
  dim elementId as string
  dim el as object
  dim newState as boolean
  dim item as string
  
  '...prüfen, ob das erste Element sichtbar ist
  item =globTagData1(tagIndex+1) 'Array hat(te ???) zur Zeit einen offset von 1 
  LINE = split(item, "<³³>")
  elementID="btn_"+LINE(cMainType)
  el=ref.element(elementID)
  newState = NOT el.visible
  groupId = LINE(cGroup )
  
  dim title as string=trim(e.sender.text)
  title = replace(title, "< -","")
  title = replace(title, "...","")

  if newState=true then
    'el2.checked = true
    e.sender.Image = nothing '' ZZ.GetImageCached("http://icons3.iconfinder.netdna-cdn.com/data/icons/softwaredemo/PNG/16x16/Box_Red.png")
    e.sender.text=" "+title
  else
    'el2.checked = false
    e.sender.Image  = nothing '' = ZZ.GetImageCached(tagDATA(cIconURL))
    'e.sender.text=title+" ..."
    e.sender.text=title+""
  end if
 
  'monitorAdd("!!! TREFFER: "+tagIndex.toString+"<-->"+tagDATA(cGroupMaster)+"<-->"+tagDATA(cCaption))
  'ab tagIndex+1 toggel sichtbar/nicht sichtbar
   trace ref.controls.count
   '' for i=1 to ref.Items.count-1
   ''   trace ref.items(i).name
   '' next
   '' 
  for i=tagIndex+1 to globTagData1.length-1 '... ein ungefährer wert ;-)
    item =globTagData1(i)
    if trim(item)="" then continue for
     'monitorAdd("zzzzzzzzz tag: "+item)
    LINE = split(item, "<³³>")
    if LINE(cGroup) <> groupId  then exit for
    if LINE(cGroup) = "" then exit for
      
    elementID="btn_"+LINE(cMainType)
    'monitorAdd("elementID: ", elementID)
    el=ref.element(elementID)
    if el is nothing then
       elementID="mnu_"+LINE(cMainType)
      el=ref.element(elementID)
    end if 
    
    :el.visible=newState:err.Number=0 '...bei geschachtelten Menü hab ich keine referenz
  next
  '' else
  ''   monitorAdd("nix groupMode")
  '' end if
  '' 
  
  ref.ResumeLayout(False)
  ref.PerformLayout()
end sub




sub checkDisplayState(sender)
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls(3).name
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls("userToolbar_" + ToolbarID1).childs.count
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls(1).key

  ref.SuspendLayout()

  dim el as object
  dim item as object
  '' GEHT NICHT MEHR: dim toolbar=ZZ.IDEHelper.getMainFormRef.flpToolbar.controls("userToolbar_" + ToolbarID1)
  'dim toolbar =ZZ.IDEHelper.CreateToolbar(ToolbarID1)

  if sender is nothing then exit sub
  dim toolbar =sender.GetCurrentParent
  if toolbar is nothing then exit sub

  trace toolbar.name
  dim FIELDS() as string
  dim id
  for each el in toolbar.items
    if not el.Tag is nothing then
       'trace el.tag
       'monitorAdd( el.name, el.tag)
       checkDisplayState2(el, el.tag)
       '' FIELDS=split(el.tag,"<³³>") 
       '' dim id = DATA(1)
       '' monitorAdd("id",id)
       '' dim toolWindow=ZZ.getDockPanelRef("id" )
       '' dim curState as boolean  = toolWindow.visible
       '' monitorAdd(id,curState)
     end if
   next
   
  ref.ResumeLayout(False)
  ref.PerformLayout()
end sub



'--> checkDisplayState2
:sub checkDisplayState2(el, tag)
  '' monitorAdd(tag)  
  dim FIELDS() as string=split(tag,"<³³>")
  dim id = FIELDS(cTagPara)
  id=replace(id,"_##_","|##|")
  ' monitorAdd("id",id)
  dim toolWindow=ZZ.getDockPanelRef(id )
  if not toolWindow is Nothing then
    dim curState as boolean  = toolWindow.visible
    'monitorAdd(curState.toString , id)
    if curState=true then
      '''monitorAdd(curState.toString , id)
      : el.checked=true: err.number=0 'MenuItems haben das nicht
    else
      : el.checked=false: err.number=0 'MenuItems haben das nicht
    end if
  else
    ''monitorAdd("notFound", id)
  end if
end sub


'-->


sub onCmdToggleTestLab(e, tagDATA)
  dim id as string="ToolBar|##|tbScriptWin|##|es_testLabor.mainWin"
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit sub
  dim curState as boolean  = toolWindow.visible
  dim newState as boolean= not curState
  monitorAdd(id)
  if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    toolWindow.parent.parent.visible =true   
  else
    toolWindow.hide()
    toolWindow.visible =false
    toolWindow.parent.parent.visible =false
  end if
end sub




sub onToggleDockWindow(e, tagDATA)
  dim tagPara as string = tagDATA(cTagPara)
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

'newState=true
 if newState = true then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
   else
    toolWindow.hide()
    toolWindow.visible =false
  end if
  
  return newState
end function


sub onSaveRun (e, tagDATA)
  dim para=tagDATA(cTagPara)
  monitorClear()
  
  monitorAdd("onSaveRun")
  monitorAdd("------- SAVING")
  ZZ.IDEHelper.Exec("File.Save", "","")
  monitorAdd("------- SAVED")
  zz.doevents
  ZZ.IDEHelper.Exec("Debug.Run", "","")
  monitorAdd("----------- READY")
  zz.doevents
  '' ZZ.shellExecute(para) : Err.Clear
  
  ref.ResumeLayout(False)
  ref.PerformLayout()

end sub


sub onCreateIconList (e, tagDATA)
  createIconList()
end sub

sub onTest1(e, tagDATA)
  test1()
end sub


sub onCmdToggleWeb(e, tagDATA)
  dim id as string = defaultBrowserId 
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit sub
  dim curState as boolean  = toolWindow.visible
  dim newState as boolean= not curState
  showHideNavigateWebBrowser("",newState)
end sub


sub showHideNavigateWebBrowser(url as string, newState as boolean)
  dim id as string = defaultBrowserId 
  dim toolWindow=ZZ.getDockPanelRef(id)
  if toolWindow is nothing then exit sub
  
  dim curState as boolean  = toolWindow.visible
  if newState <> curState then
    if newState = true then
      toggleDockPanel(8, 0)  'sideBarLeft verstecken
      toolWindow.visible =true
      toolWindow.parent.visible =true
      toolWindow.show()
     else
      toggleDockPanel(8, 1)  'sideBarLeft anzeigen
      toolWindow.hide()
      toolWindow.visible =false
    end if
  end if
  
  if url<>"" then
    dim scriptClassName=defaultBrowserScriptId   '..."es_webbrowser3"
    dim webBrowser =zz.scriptClass(scriptClassName)
    if webBrowser is nothing then exit sub
    webBrowser.navigateWebbrowser(url)
  end if
end sub




sub onNavigateUrl(e, tagDATA)
  '... WAS NOCH ALLES FEHLT:
  '... derzeitiges panel merken
  '... schaltfläche hin / zurück ???
  '... ??? alle pupups sichern und verstecken ???
  '... panorama aktivieren ???

  dim url=tagDATA(cTagPara)
  showHideNavigateWebBrowser(url,true)
  exit sub
  
  dim id as string = defaultBrowserId 
  ToggleDockWindowFromId(id, 1) 
  'ZZ.getDockPanelRef(id).show()                ' eine defaultMethode NAVIGATE fehlt noch --> Interprog ???
  dim scriptClassName=defaultBrowserScriptId   '..."es_webbrowser3"
  dim webBrowser =zz.scriptClass(scriptClassName)
  webBrowser.navigateWebbrowser(url)
  
  toggleDockPanel(8, 0)   'sideBarLeft verstecken
end sub





sub onToggleImg(e, tagDATA)
  dim img1=tagDATA(cIconURL)
  dim img2=tagDATA(cIconURL2)
  dim tag =e.sender.tag
  monitorAdd(img1)
  monitorAdd(img2)
  dim newImage as string=""
  if tag.startsWith(" ")then
  'if e.sender.Image is ZZ.GetImageCached(img1) then
     'newImage=img1
     newImage=iconList(0)
     e.sender.tag=trim(tag)
  else
    'newImage = img2
     newImage=iconList(1)
    e.sender.tag=" "+tag
  end if
  e.sender.Image = ZZ.GetImageCached(newImage)
 
end sub

sub onExe(e, tagDATA)
  dim para=tagDATA(cTagPara)
  : ZZ.shellExecute(para) : Err.Clear
end sub

sub onExecute(e, tagDATA)
  dim para=tagDATA(cTagPara)
  : ZZ.IDEHelper.Exec(para, "",""): Err.Clear

end sub


sub onCmdOptions(e, tagDATA)
  dim para=tagDATA(cTagPara)
  :  ZZ.IDEHelper.ShowOptionsDialog(para) : Err.Clear
end sub


sub onCmdOpenExplorer(e, tagDATA)
  Dim tab As IDockContentForm = ZZ.IDEHelper.getActiveTab()
  Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  Process.start("explorer.exe", "/select," + fileSpec)
end sub


sub onCmdToolbarMan(e, tagDATA)
  msgBox("TOLLBAR-MANAGER: coming soon")
end sub

sub onCmdIconMan(e, tagDATA)
  msgBox("ICON-MANAGER: coming soon")
end sub




'-->

Function toggleFullScreen()
  static state as boolean = true: state = not state
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  dim Panel as object = mainWin.DockPanel1
  dim p as object: dim i as integer
  mainWin.SuspendLayout()
  For Each p In Panel.Panes
    For i = 0 To p.Contents.Count - 1
      if p.DockState = 7 or p.DockState = 8 or p.DockState = 9 or p.DockState = 10 then 
        p.parent.visible = state
      end if
     Next
  Next
  mainWin.ResumeLayout(False)
  mainWin.PerformLayout()
End Function


Function onTogglePane(e, tagDATA)

  if e.mouseButton ="Right" then
    checkForExpanderState(e, tagDATA)
    exit Function
  end if
  
  dim tagPara as string=trim(tagDATA(cTagPara))
  if tagPara="" then exit function ' ...nichts zu tun
  
  dim intPanel as integer =cInt(tagPara)
  monitorAdd("PARA: "+intPanel.toString)
  if intPanel=-1 then
    toggleFullScreen()
  else
    toggleDockPanel(intPanel)
  end if
end function


Function togglePanel(panelRef)
  dim curState as boolean  = panelRef.visible
  if curState =false then
    panelRef.show()
    panelRef.visible =true
   else
    panelRef.hide()
  end if
  ZZ.IDEHelper.getMainFormRef.focus()
  Return True '=Cancel
end function


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
'--> T I M E R

'--> DockPanel1_ActivePaneChanged
:Sub DockPanel1_ActivePaneChanged(ByVal sender As Object, ByVal e As System.EventArgs)
  startTimer()
  monitorAdd("ActivePaneChanged")
  dim activePane = sender.ActivePane()
  
  '... der macht fehler: 
  'trace "activePane.name", activePane.name
  
  '...der geht auch nicht
  'dim caption as string=sender.ActivePane.captionText
  
   dim traceMode as boolean =true
   dim tabName as string
   : if not activePane is nothing then
   :   tabName=activePane.CaptionText
   :   if traceMode=true then
   :      statustext(tabName)
   :   end if 
   : else
   :   if traceMode=true then
   :     trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+"activePane ...gibts nicht"
   :     statustext("activeTabName: Fehleanzeige !!!")
   :   end if
   : end if

  exit sub

  'trace "caption", caption
  
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.bringToFront()
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.left=0' 200
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.top=33
  ''  : sender.top=65' 26
  ''  : sender.height=ZZ.IDEHelper.MainWindow.height-70
  ''  : ZZ.IDEHelper.MainWindow.StatusStrip1.visible=false
  ''  : dim tabName as string
  ''  : if not activePane is nothing then
  ''  :   tabName=activePane.CaptionText
  ''  :   if traceMode=true then
  ''  :      trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+tabName
  ''  :   end if 
  ''  :   'printLine 11,  "! <<  ActivePanelChanged",tabName
  ''  :   dim scriptClass as object = zz.scriptClass("es_testLabor")
  ''  :   if not scriptClass is nothing then
  ''  :      scriptClass.add("activeTab(utb2): "+tabName)
  ''  :   end if
  ''  : else
  ''  :   if traceMode=true then
  ''  :     trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+"activePane ...gibts nicht"
  ''  :   end if
  ''  : end if

end sub


'--> startTimer
:sub startTimer()
  'Timer1.Interval = 111
  'Timer1.Enabled = True
  statustext("timer gestartet")
end sub


sub stopTimer()
  'Timer1.Enabled = false
  statustext("timer ist ausgeschaltet")
  monitorAdd("Timer gestopt...")
end sub


Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  ': stopTimer()
  : onTimerEvent()
End Sub


'--> onTimerEvent
:sub onTimerEvent()
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls(3).name
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls("userToolbar_" + ToolbarID1).childs.count
  'trace "elName", ZZ.IDEHelper.getMainFormRef.flpToolbar.controls(1).key
  dim el as object
  dim item as object
  '' GEHT NICHT MEHR: dim toolbar=ZZ.IDEHelper.getMainFormRef.flpToolbar.controls("userToolbar_" + ToolbarID1)
  dim toolbar=ZZ.IDEHelper.CreateToolbar(ToolbarID1)
  dim FIELDS() as string
  dim id
  for each el in toolbar.items
    if not el.Tag is nothing then
       trace el.name
       'monitorAdd( el.name, el.tag)
       checkDisplayState2(el, el.tag)
       '' FIELDS=split(el.tag,"<³³>") 
       '' dim id = DATA(1)
       '' monitorAdd("id",id)
       '' dim toolWindow=ZZ.getDockPanelRef("id" )
       '' dim curState as boolean  = toolWindow.visible
       '' monitorAdd(id,curState)
     end if
   next
end sub












'-->
'-->    
'--> P A R K P L A T Z:



Sub DockPanel1_DocumentTabActivated(ByVal Tab As Object, ByVal Key As String)
  '... der kommt nicht
  'msgbox("ActivePaneChanged ????")
  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  zz.doevents()
   trace "! DockPanel1_DocumentTabActivated", key
   'printLine 11, "!!!! ActivePanelChanged", ActiveTabFilespec
   'trace "!!!! ActivePanelChanged", name
end sub



sub ichWerdeNichtMehrGebraucht

  '... war mal bei DockPanel1_ActivePaneChanged drinnen

  '' 'msgbox("ActivePaneChanged ????")
  ''  sharedEventCounter=sharedEventCounter+1
  ''  dim traceMode as boolean=false
  ''  'traceMode=true
  ''  'dim activePane = sender.ActivePane()
  ''  ' trace activePane.CaptionText, activePane.name
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.bringToFront()
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.left=0' 200
  ''  : ZZ.IDEHelper.MainWindow.flpToolbar.top=33
  ''  : sender.top=65' 26
  ''  : sender.height=ZZ.IDEHelper.MainWindow.height-70
  ''  : ZZ.IDEHelper.MainWindow.StatusStrip1.visible=false
  ''  : dim tabName as string
  ''  : if not activePane is nothing then
  ''  :   tabName=activePane.CaptionText
  ''  :   if traceMode=true then
  ''  :      trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+tabName
  ''  :   end if 
  ''  :   'printLine 11,  "! <<  ActivePanelChanged",tabName
  ''  :   dim scriptClass as object = zz.scriptClass("es_testLabor")
  ''  :   if not scriptClass is nothing then
  ''  :      scriptClass.add("activeTab(utb2): "+tabName)
  ''  :   end if
  ''  : else
  ''  :   if traceMode=true then
  ''  :     trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+"activePane ...gibts nicht"
  ''  :   end if
  ''  : end if
  '' 
end sub

'-->

sub xxx_btn_togglePane_1_mouseClick(e) '2
   monitorAdd("========>>> btn_togglePane_1_mouseClick")
end sub

sub xxx_togglePane_1_mouseClick(e)  '1
   monitorAdd("========>>> togglePane_1_mouseClick")
end sub


sub XXX_toogleToolbarItems(e, tagDATA)
dim id as string
static state as boolean =true
state = not state
  id="btn_togglePane_1" : ref.element(id).visible=state
  id="btn_togglePane_2" : ref.element(id).visible=state
  id="btn_togglePane_3" : ref.element(id).visible=state
  id="btn_togglePane_4" : ref.element(id).visible=state
  id="btn_togglePane_5" : ref.element(id).visible=state
  'id="btn_expand_01"    : ref.element(id).checked= state and true
  dim id2 ="btn_expand_01"
  dim el2 =  ref.element(id2)
  dim title = el2.text
  title = replace(title, "< -","")
  title = replace(title, "- >","")
  if state=true then
    'el2.checked = true
    el2.Image = ZZ.GetImageCached("http://es.teamwiki.net/docs/icons/minus16.png")
    el2.text="< - "+"Ansicht"
  else
    'el2.checked = false
    el2.Image = ZZ.GetImageCached("http://es.teamwiki.net/docs/icons/gnome-dev-computer.png")
    el2.text="Ansicht"+" - >"
  end if
  
end sub



Sub createOpenedTabList_xxx(DockPanel1)
      trace "--------------", "------------------------------ ???"
    Dim doc = listDocumentPane(DockPanel1)
    
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


 Function listDocumentPane(mainWin) As object 'DockPane
   dim panel1 as object = mainWin.DockPanel1
   dim p as object
   dim i
   dim isMdiChild as boolean
   dim dockHandler as object
   dim tag
   dim DockState as object
   dim DockContent as WeifenLuo.WinFormsUI.Docking.DockContent
    For Each p In panel1.Panes
      For i = 0 To p.Contents.Count - 1
        isMdiChild=false
        dockHandler = p.Contents(i).DockHandler
        'if p.Contents(i) Is mainWin.ActiveMdiChild Then isMdiChild=true
        if dockHandler Is mainWin.ActiveMdiChild Then isMdiChild=true
        ''tag=DirectCast(p.Contents(i).DockHandler.Form, DockContent).GetPersistString()
        'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()

      ''trace p.DockState,  isMdiChild.toString+"  "+p.Contents(i).DockHandler.TabText
      
      'if dockHandler.isFloat = true  then 
        
        'trace p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText
        trace  mainWin.DockPanel1.ActiveMdiChild.title, "  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText
      
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
'-->
'--> outMONITOR

sub monitorAdd(para1 as string, optional para2 as string="")
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.add(para1+": "+para2+"<--")
   : end if
   : err.number=0
end sub

sub monitorClear()
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :    scriptClass.clear()
   : end if
   : err.number=0
end sub


sub statustext_reset()
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.statustext_reset
   : end if
   : err.number=0
end sub


sub errorText(para as string)
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.errorText(para)
   : end if
   : err.number=0
end sub

sub statustext(para as string)
   : on error resume next
   : dim scriptClass as object = zz.scriptClass("es_testLabor")
   : if not scriptClass is nothing then
   :   scriptClass.statustext(para)
   : end if
   : err.number=0
end sub





'' Sub onMenuEvent(e)
''   If e.EventName = "ItemClicked" Then
''     Dim btnName = e.Sender.Name.Substring(4)
''     
'' :   CallByName(Me, btnName, Microsoft.VisualBasic.CallType.Method, e)
''     If Err.Number Then MsgBox("Funktion nicht gefunden"+vbnewline+"..."+btnName):Err.Clear
''     
''   End If
'' End Sub
'' 
'' 


'-->




