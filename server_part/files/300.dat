
' es_dialogEditor

'biegelmaier 06441 747 84


'#KOM: hier ist Platz für Kommentar
'#KOM: oder ein eigenes Feld dafür vorsehen ???
'#KOM: 
'#KOM: 
'#KOM: 


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

''//  tpl_layout_splitterbar.vb
'-->
'--> G L O B 
#Para DebugMode intern
#Para SilentMode true
Public Const Auto = -2
shared myScriptClass

'--> P A R A M E T E R -------------------------
Const  winID = "{ScriptClass}.editorWin"
const  globHideOnClose as boolean = false           'bei true merkt er sich die Position

public winCaption as string="Parameter-Dialoge: ... bequem erstellen und verwalten"
public glob_defaultGridPath as string = "C:\yPara\scriptIDE\gridTest\"

dim toolBarColor as string="#E4E1DC"

Dim WithEvents FormRef As Form
Dim PanelRef As ScriptedPanel

' Dim toolbar As ScriptedToolstrip

public shared toolBar1 As ScriptedPanel
public shared toolBar2 As ScriptedPanel
public shared statBar As ScriptedPanel
public shared containerMain As ScriptedPanel
public shared container1 As ScriptedPanel
public shared container2 As ScriptedPanel
public shared container3 As ScriptedPanel
public shared container4 As ScriptedPanel
public shared container5 As ScriptedPanel
public shared container6 As ScriptedPanel

Dim WithEvents IGrid1 As IgridEx
dim globGridHeaderText1=getHeaderData()


Dim WithEvents IGrid2 As IgridEx
Dim WithEvents IGrid3 As IgridEx


public shared splitContainer1  As Object
Dim pLeft1, pRight1 As ScriptedPanel

public shared splitContainer2  As Object
Dim pLeft2, pRight2 As ScriptedPanel

public shared splitContainer2sub  As Object
Dim pLeft22, pRight22 As ScriptedPanel

public shared splitContainer3  As Object
Dim pLeft3, pRight3 As ScriptedPanel

public shared splitContainer33  As Object
Dim pLeft33, pRight33 As ScriptedPanel

public shared splitContainer4  As Object
Dim pLeft4, pRight4 As ScriptedPanel

public shared splitContainer44  As Object
Dim pLeft44, pRight44 As ScriptedPanel

dim myTextArea as textbox
dim myTextArea2 as textbox
   
 dim jumboBgColor as string="#2E74D5"
 dim jumboForeColor as string="#fff"


Dim WithEvents tmr1 As Timer
Dim glob As cls_globPara

Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer





'-->
sub test1(dummy as object)
  printLine 3, "aaa","bbbbb"
  monitorAdd("---------------------------------------")
  switchContainer(5)  
  dumpDialogData()
  
exit sub

  testAdr()
  
end sub


sub test2()
  
  startSearch()
end sub




sub test3()
end sub



Sub startSearch() '... fileFinder
  dim root as string
  root="C:/yPara/"
  root="C:\_sync\"
  root="C:\_sync\vbNet\_AppVorlage"
  root="C:\_sync\vbNet"
  root="C:\yPara"
  dim fileFinder as object
  dim erg as string
  fileFinder = ZZ.OpenFileFinder(root) 
  fileFinder.SortCriteria=1
  ' 'erg= fileFinder.findFiles()
  erg = fileFinder.FindRecursive()
  'ZZ.setOutMonitor (erg)
  monitorAdd(erg)
  fileFinder = nothing

  'msgBox ("...aktuelle Baustelle")
end sub


'-->

Sub onTerminate()
  trace "ugb_emailMischer", "onTerminate"
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  'stop timer  : fehlen noch
  'stop resizer 
End Sub



Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  saveGlobPara
End Sub


sub saveGlobPara()
  trace "saveGlobPara","try..."
  if glob is nothing then exit sub
  if formRef is nothing then exit sub
  glob.saveFormPos(FormRef)
  glob.saveTuttiFrutti(FormRef)
  glob.para("myMes")="ich bin eine Nachricht"
  glob.saveParaFile()
  trace "saveGlobPara","DONE"
end sub


Sub globAddHandler()
  'AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'AddHandler Timer1.Tick, AddressOf Timer1_Tick
  AddHandler formRef.Resize, AddressOf Form_Resize
end sub


Sub globRemoveHandler()
  trace "globRemoveHandler","try..."
  if formRef is nothing then exit sub
  'RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
  'RemoveHandler Timer1.Tick, AddressOf Timer1_Tick
   RemoveHandler formRef.Resize, AddressOf Form_Resize
  trace "globRemoveHandler","DONE"
end sub



Sub Form_Resize(sender as Object, e as EventArgs) 
  onResizeControls
End Sub


Sub onResizeControls()
   trace " Resize(height): ", formRef.Height.toString+" ... "+winID
   'myTextArea.Height = container1.Height - 60 '42
End Sub



Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr1.Tick
  '...z.B. aktuelle Urzeit ausgeben
End Sub



'-->

Function GetPanelRef()
  onTerminate() 'aufruf, um das alte leben ein bischen anzuhalten 
  ' If PanelRef IsNot Nothing Then Exit Sub
  'PanelRef = ZZ.IDEHelper.CreateDockWindow(winID, 1) '1=popupWin
  '--> dockWindow / normales Fenster
  'PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.StdWindow, 1)'  : err.number=0
  formRef = getOuterWindowRef(panelRef)
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose : err.number=0
  return panelRef
End Function



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



'-->
'--> DIALOG  ____________________________

function getDialogData()                  
#Data myDialogData as String              
'======================================START-DIALOG
TYPE	id / cmd	nicName	fullName	Titel	DIALOG	edit/view/hidden	labBgColor	labForeColor	textBgColor	textForeColor	GRID	defaulWidth	gridBgColor	textForeColor	HELP	toolTip	text1	text2	icon	
1											xxx			1995-01-01	1999-01-02	SONDER(es)				
dialog			feld	titel																
dialog			vorname	hinz und kunz ?																
dialog			name	huber oder mayer ?								1	2	3	4	5	6	7	8	
dialog			PLZ	...5-stellig !								a	b	c	d	e	f	g	h	
dialog			schuhgröße	38-46																
dialog			schuhgröße	38-46																
dialog			schuhgröße	38-46																
																				
																				
'======================================END-DIALOG
#EndData                                  
  '' Trick 17: Zeilennummer dazupacken    
  dim RESULT as string=myDialogData       
  return RESULT                           
end function                              



sub dumpDialogData()
  dim content as string
  Igrid_get(iGrid1,content)
  monitorAdd(content)
  monitorAdd("---------------------------------------")
  
  dim x as string="": dim y as string=""
  
  x=x+vbNewLine
  x=x+vbNewLine+"function getDialogData()                  "
  x=x+vbNewLine+"#Data myDialogData as String              "
  x=x+vbNewLine+"'======================================START-DIALOG" + vbNewLine
  y=y+vbNewLine  '[[CONTENT]]                              "  '!!! kein Plus-Zeichen
  y=y          +"'======================================END-DIALOG"
  y=y+vbNewLine+"#EndData                                  "
  y=y+vbNewLine+"  '' Trick 17: Zeilennummer dazupacken    "
  y=y+vbNewLine+"  dim RESULT as string=myDialogData       "
  y=y+vbNewLine+"  return RESULT                           "
  y=y+vbNewLine+"end function                              "
  y=y+vbNewLine
  
  myTextarea.text=x+content+y
 
  
exit sub
  
  Dim DATA() as string=split(content,vbNewLine)
  dim i as integer
  dim max as integer=DATA.length
  dim item as string
  dim FIELDS() as string
  
  dim index as integer = 0
  dim fullName as String
  
  '' container.resetControls (5,5)
  '' 
  '' for i=0 to max-1
  ''   item=DATA(i)
  ''   FIELDS=split(item, vbTab)
  ''   if FIELDS(0).toUpper.startsWith("DIALOG") then
  ''     monitorAdd(item)
  ''     fullName=FIELDS(3)
  ''     createDialogElement(container, index, fullName , index.toString, item)
  ''     index=index+1
  ''   end if
  '' next
  '' 
end sub

sub grabDialogData()
  dim content as string
  content=myTextarea2.text
  
  dim isEmpty as boolean =true
  if content.length>0 then isEmpty=false
  if isEmpty=true then exit sub
  
  dim DATA() as string=split(content,vbNewLine)
  
  dim item as string
  dim OUT(1000) as string
  dim counter as integer=0
  dim isDataBlock as boolean =False
  dim i as integer
  dim max as integer=DATA.length
  for i=0 to max-1
    item=DATA(i)
    if item.contains("===START-DIALOG") then isDataBlock=true 
    if item.contains("===END-DIALOG") then isDataBlock=false
    if isDataBlock = true then
      if trim(item).startsWith("'==") then continue for
      out(counter)=item
      counter=counter+1
    end if
  next
  redim preserve OUT(counter)
  dim RESULT=join(OUT,vbNewLine)
  Igrid_put(iGrid1,RESULT)
  
  'automatischer zurückschalter, wenn daten eingefügt wurden
  myTextarea2.text=""
  switchContainer(2) 
  
end sub

Sub cmdDumpDialogData(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdDumpDialogData': ...in Arbeit")
  'msgBox("'cmdDumpDialogData': " + "...Coming soon ")
  '...
  'msgbox(index)
  switchContainer(index)  
  dumpDialogData()

End Sub


Sub cmdGrabDialogData(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdGrabDialogData': ...in Arbeit")
  'msgBox("'cmdGrabDialogData': " + "...Coming soon ")
  '...
  switchContainer(index) 
  grabDialogData()
End Sub





'-->
'--> A U T O S T A R T _____________________

'__________ DIASLOG-EDITOR _____________
 
  ' ... lesen byCode und/oder copyPaste

  '' speichern + update grün
  '' splitterbar auf fix
  '' DUMP und ausblenden
  
   ' ... DesignParameter: eigene Seite ???

   ' OK: dumpModus vorbereiten
   ' OK: dialogDefinition in DATA-BLOCK speichern
   ' ... lesen byCode und/oder copyPaste


  'OK: autoscroll
  'OK: grid: no soart
  'OK: trenn(provisorisch)
  'OK: bgColor(iGrid)

  'OK: index umstellen
  'OK: !!! SYNC dialog --> grid
  'OK: !!! SYNC grid --> dialog  ... fast fertig
  'OK: createGridHaederText herausziehen
  
  
'________ DIALOG-MANAGER __________________
  '' + toggle sidebar
  '' statt xxx-3: update
  '' NEU: OK / Übernehmen
  '' splitterbar oben auf fix
  
  
  
  ' OK: Form leer raümen
  ' ... OK-Knopf, grün
  ' ... Abbrechen, rot
  ' ... toolBar für seitenTitel verwenden
  ' ... beim speichern vorschau aktualisieren (outomatisch)
  ' ... DesignParameter: ...auswerten
  ' - icon
  ' - bgColor
  ' - Hauptüberschrift - FensterTitel
  
'______ BETA-TEST: welche Anwendung: ____________ 
'  - eMail   ???
'  - iconBar ??? 

'______ ...VORGEMERKT: _______________________
  ' onEnter: next field
  ' gridDesign speichern in Para
  
  'jumboLabel
  '
  ' ???  gridColNr <-> dialogPosition
  ' ???  defaultWerte überspringen ??? checkbox 
  
 
Sub AutoStart()
   myScriptClass = me
  'onTerminate/formClosing geht bei mir nicht
  saveGlobPara()
  globRemoveHandler()
  
  PanelRef = GetPanelRef()    '... panelRef ist als globale var definiert
  glob = ZZ.newGlobPara()

  dim el as object
  'toolBarColor="#BFBFBF"
  'toolBarColor="#65CE84"
  dim cMain      As ScriptedPanel
  dim  cToolBar As ScriptedPanel
  dim  cStatBar  As ScriptedPanel

  '--> ... containerMain
  With PanelRef
    .resetControls ()
    .activateEvents = "|IconMouseDown|   |TextboxKeyDown|"


    containerMain  =.addPanel("containerMain"    , DockStyle.Fill)
    toolBar1     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)

    toolBar1.resetControls  (10,3,1)
    toolbar1.visible=true
    toolbar1.height=24
    toolbar1.height=30
    toolbar1.BackColor = ColorTranslator.FromHtml(toolBarColor)
 
    containerMain.resetControls  (10,5,1)
    containerMain.resetControls  (10,5,1)
    containerMain.BackColor = ColorTranslator.FromHtml("#fff")
    containerMain.BackColor = ColorTranslator.FromHtml("#f00")


    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml("#8f8")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.height=30
  end with
  
  '--> ... subContainer
  with containerMain
    container1  =.addPanel("container1"    , DockStyle.Fill)
    container1.visible=false
    container1.BackColor = ColorTranslator.FromHtml("#f8f")
    container1.resetControls  (10,5,1)
 
    container2  =.addPanel("container2"    , DockStyle.Fill)
    container2.visible=false
    container2.BackColor = ColorTranslator.FromHtml("#82C084")
    container2.resetControls  (10,5,1)
    
    container3  =.addPanel("container3"    , DockStyle.Fill)
    container3.visible=false
    container3.BackColor = ColorTranslator.FromHtml("#ff8")
    container3.resetControls  (10,5,1)
    
    container4  =.addPanel("container4"    , DockStyle.Fill)
    container4.visible=false
    container4.BackColor = ColorTranslator.FromHtml("#AD9C01")
    container4.resetControls  (10,5,1)
 
    container5  =.addPanel("container5"    , DockStyle.Fill)
    container5.visible=false
    container5.BackColor = ColorTranslator.FromHtml("#8f8")
    container5.resetControls  (10,5,1)

    container6  =.addPanel("container6"    , DockStyle.Fill)
    container6.visible=false
    container6.BackColor = ColorTranslator.FromHtml("#8f8")
    container6.resetControls  (10,5,1)
 
    '--> ... setDefaultVisible = 2
    container2.visible=true

 
  end with

  '--> ... toolbar           :GLOB
  with toolbar1
   .addButton ("cmdSwitchContainer_1"  , "Start"        ) ' , "#1DD910")
   .addButton ("cmdSwitchContainer_2"  , "Felder"       ) ' , "#1DD910")
   .addButton ("cmdSwitchContainer_3"  , "Design"       ) ' , "#1DD910")
   '.addButton ("cmdSwitchContainer_5" , "Dump"        ) ' , "#1DD910")
   '.addButton ("cmdSwitchContainer_4" , "xxx"         ) ' , "#1DD910")
   .addButton ("cmdDumpDialogData_5"   , "Dump(Code)"  ) ' , "#1DD910")
   .addButton ("cmdGrabDialogData_6"   , "Grab(Code)"  ) ' , "#1DD910")
   .addButton ("cmdXXXXXXXXX_7"   , "xxxx"  ) ' , "#1DD910")
   '.addButton ("cmdEscape"            , "xxx-3"      , "#f88")
  end with
  
  '--> ... statbar           :GLOB
  with statBar
    .resetControls (5,3)
    '.addButton ("cmdEscape"             , "Aktualisieren"      , "#8AC130")
    .addButton ("cmdPreview"             , "Vorschau"      , "#8AC130")
  end with
 




  '-->
  '--> ... container_1   :START ____________
  with  container1
    .resetControls (5,5)
    .activateEvents = "|ButtonMouseClick| |TEXTBOXTEXTCHANGED|"
    
    '--> ... --- splitContainer - 1)
    el=.addSplitcontainer("splMain", "left", pLeft1, "right", pRight1, DockStyle.Fill)
    splitContainer1=el
    el.backColor=ColorTranslator.FromHtml("#88f")
    el.splitterDistance=240
  end with 
    
     
  '--> ... --- pLeft1
  with pLeft1
    .backColor=ColorTranslator.FromHtml("#4488E9")
    .resetControls (5,5)
    .insertX = 11 :.insertY = 12
    .addIcon("appPicture11", "http://es.teamwiki.net/docs/icons/Help-File.png" )
  end with
 
 
 
 
  '--> ... --- pRight1
  with pRight1
    .resetControls (5,5)
    .backColor=ColorTranslator.FromHtml("#fff")
    .BR  '--------------------------------------------------
    .insertX = 10  :.insertY = 10
    dim label as string="DialogEditor"
    el=.addLabel  ("lab001", label ,  ,"",,,-2,35) :
    el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
    makeJumboLabel(el)

    label ="??? ein RTF integrieren ??? "+vbNewLine+vbNewLine+"... für Hyperlinks"+vbNewLine+"... und Überschriften "
    '.insertX = 15  :.insertY = 210

    .insertX = 10  :.insertY = 60
    el=.addLabel  ("lab002", label ,  ,"",,,-2,35) :
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.height=1111
    el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
    el.textAlign=System.Drawing.ContentAlignment.topLeft
    ' ??? el.padding=10
  end with
 
 
 

 '-->  
 '--> ... container_2  :FELDER_____________
  with container2
    .backColor = ColorTranslator.FromHtml("#B1D8B2")
    cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                          'cMain.backColor = ColorTranslator.FromHtml("#faa")
    cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                          cToolBar.backColor = ColorTranslator.FromHtml("#BDBBB7")
                          cToolBar.backColor = ColorTranslator.FromHtml("#fff")
    cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                          cStatBar.backColor = ColorTranslator.FromHtml("#BDBBB7")
  end with
  
  with cToolBar
    .visible=false
    .addButton ("cmdXXX_10"     , "xxx-10" )' , "#1DD910")
    .addButton ("cmdXXX_11"     , "xxx-11" )' , "#1DD910")
  end with
  
  with cMain
    el=.addSplitcontainer("splitContainer3", "left", pLeft3, "right", pRight3, DockStyle.Fill)
    el.Orientation = System.Windows.Forms.Orientation.Horizontal
    splitContainer3=el
    el.backColor=ColorTranslator.FromHtml("#BDBBB7")
    el.splitterDistance=150
  end with
  
  
  '--> ...  --- pLeft3            :...LOGO-ZOOM
  with pLeft3
    .backColor=ColorTranslator.FromHtml("#fff")
    .backColor=ColorTranslator.FromHtml("#4488E9")
    .insertX = 11 :.insertY = 12
    .addIcon("appPicture33", "http://es.teamwiki.net/docs/icons/Help-File.png " )

    .insertX = 230 :.insertY = 15
    dim label as string="T I T E L "
    el=.addLabel  ("lab001", label ,  ,"",,,-2,35) :
    el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
    makeJumboLabel(el)

    label ="  Überschrift..."+vbNewLine+vbNewLine+"... 2. Zeile"+vbNewLine+"... 3. Zeile"
    .insertX = 230  :.insertY = 60
    el=.addLabel  ("lab002", label ,  ,"",,,-2,35) :
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

    el.height=77
    el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
    el.textAlign=System.Drawing.ContentAlignment.topLeft
    ' ??? el.padding=10
  end with
   
         
   
  
  '--> ...  --- pRight3         :...UNTEN
  with pRight3 
    el=.addSplitcontainer("splitContainer33", "left", pLeft33, "right", pRight33, DockStyle.Fill)
    ''el.Orientation = System.Windows.Forms.Orientation.Horizontal
    splitContainer33=el
    el.backColor=ColorTranslator.FromHtml("#4488E9")
    el.splitterDistance=230
  end with


  '--> ...  --- pLeft33         :...UNTEN-LINKS
  with pLeft33 
   .resetControls (5,5)


  '--> ... ... --- iGrid-1
    IGrid1 = New IgridEx
    IGrid1.name="IGrid1"
    .Controls.Add(IGrid1)
    IGrid1.Dock = DockStyle.Fill
    'IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid1)
    with IGrid1

      .fileSpec=glob_defaultGridPath+"esGrid1.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")

      '.AllowDrop = True
      '.ReadOnly = True
      .Cols.Add("Nr",111)
      .Cols.Add("r",222)
      .Cols.Add("ZeilenInhalt",111)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
      .rows.count = 100
    end with
  end with



  '--> ...  --- pRight33         :...UNTEN-Rechts
  with pRight33 
    .AutoScroll = False
    .backColor=ColorTranslator.FromHtml("#C3DAFF")
    .resetControls (5,5)
    dim container As ScriptedPanel=pRight33    'ACHTUNG:...nicht vergessen
    dim labBgColor as string="#00f"
    dim tagData as object=nothing
    dim index as integer=-1
    index=index+1: createDialogElement(container, index, "TYPE"              , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "id / cmd"          , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "nicName"           , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "fullName"          , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "Titel"             , index.toString, nothing)
    .br(5)
    index=index+1: createDialogElement(container, index, "GRID"              , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "defaulWidth"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "gridBgColor"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "textForeColor"     , index.toString, nothing)
    .br(5)
    index=index+1: createDialogElement(container, index, "DIALOG"            , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "edit/view/hidden"  , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labBgColor"        , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labForeColor"      , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "textBgColor"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "textForeColor"     , index.toString, nothing)
    .br(5)
    index=index+1: createDialogElement(container, index, "HELP"              , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "toolTip"           , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "text1"             , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "text2"             , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "icon"              , index.toString, nothing)
    'index=index+1: createDialogElement(container, index, "..."               , index.toString, nothing)
    'index=index+1: createDialogElement(container, index, "..."               , index.toString, nothing)
    .br(1)
     el=.addLabel  ("dummyForAutoScroll", "xxxx" ,  ,"",,,33,5) : el.left=-999  '...fügt noch etwas luft ein
    .AutoScroll = True
  end with
  
  globGridHeaderText1=getHeaderData()
  monitorAdd(globGridHeaderText1)
  ' msgbox(globGridHeaderText1)


  '--> ... --- cStatbar 
   with cStatbar
    .addButton ("cmdCopyIGrid_1"         , "Kopieren 1" )' , "#1DD910")
    .addButton ("cmdInsertIGrid_1"       , "Einfügen 1" )' , "#1DD910")
    .addButton ("cmdReadDefaultIGrid_1"  , "Lesen 1" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_1"  , "Speichern 1" )' , "#1DD910")
  end with
 

 '-->  
 '--> ... container_3   : DESIGN ______________________
   with container3
     cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                           cMain.backColor = ColorTranslator.FromHtml("#ccc")
     cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                           cToolBar.backColor = ColorTranslator.FromHtml("#FFFF88")
     cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                           cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
   end with


  with cToolBar
    .visible=false
    .addButton ("cmdXXX_10"     , "xxx-10" )' , "#1DD910")
    .addButton ("cmdXXX_11"     , "xxx-11" )' , "#1DD910")
  end with
  
  with cMain
    el=.addSplitcontainer("splitContainer4", "left", pLeft4, "right", pRight4, DockStyle.Fill)
    el.Orientation = System.Windows.Forms.Orientation.Horizontal
    splitContainer4=el
    el.backColor=ColorTranslator.FromHtml("#BDBBB7")
    el.splitterDistance=150
  end with
 
'#################################
  
  '--> ...  pLeft4            :...LOGO-ZOOM
  with pLeft4
    .backColor=ColorTranslator.FromHtml("#fff")
    .backColor=ColorTranslator.FromHtml("#4488E9")
    .insertX = 11 :.insertY = 12
    .addIcon("appPicture33", "http://es.teamwiki.net/docs/icons/graphic_design2.png" )

    '.insertX = 15  :.insertY = 160
    .insertX = 230 :.insertY = 15
    dim label as string="T I T E L "
    el=.addLabel  ("lab101", label ,  ,"",,,-2,35) :
    el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
    makeJumboLabel(el)

    label ="  Überschrift..."+vbNewLine+vbNewLine+"... 2. Zeile"+vbNewLine+"... 3. Zeile"
    '.insertX = 15  :.insertY = 210
    .insertX = 230  :.insertY = 60
    el=.addLabel  ("lab102", label ,  ,"",,,-2,35) :
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))

    el.height=77
   el.backColor=ColorTranslator.FromHtml(jumboBgColor): el.foreColor=ColorTranslator.FromHtml(jumboForeColor): 
       el.textAlign=System.Drawing.ContentAlignment.topLeft
    ' ??? el.padding=10
  end with
   
         
   
  
  '--> ...  pRight4         :...UNTEN
 
  with pRight4 
    el=.addSplitcontainer("splitContainer44", "left", pLeft44, "right", pRight44, DockStyle.Fill)
    ''el.Orientation = System.Windows.Forms.Orientation.Horizontal
    splitContainer44=el
    el.backColor=ColorTranslator.FromHtml("#4488E9")
    el.splitterDistance=230
  end with


  '--> ...  pLeft44         :...UNTEN-LINKS
  with pLeft44 
   .resetControls (5,5)
    '--> ... --- iGrid-2  
    IGrid2 = New IgridEx
    IGrid2.name="IGrid2"
    .Controls.Add(IGrid2)
    IGrid2.Dock = DockStyle.Fill
    'IGrid2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid2)
    with IGrid2
      .fileSpec=glob_defaultGridPath+"esGrid2.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")
   
      '.AllowDrop = True
      '.ReadOnly = True
   
      .Cols.Add("Nr",111)
      .Cols.Add("r",222)
      .Cols.Add("ZeilenInhalt",111)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check

      .rows.count = 100
    end with
  end with

  with pRight44 
    .AutoScroll = False
    .backColor=ColorTranslator.FromHtml("#C3DAFF")
    .resetControls (5,5)
    dim container As ScriptedPanel=pRight44    'ACHTUNG:...nicht vergessen
    dim labBgColor as string="#00f"
    dim tagData as object=nothing
    dim index as integer=99
    index=index+1: createDialogElement(container, index, "A L L G E M E I N"  , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "DialogIcon-1"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "DialogIcon-1"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "Fenster-Titel"      , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "jumboLabelText"     , index.toString, nothing)
    .br(5)
    index=index+1: createDialogElement(container, index, "C O L O R S"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "defaulWidth"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "gridBgColor"       , index.toString, nothing)
    index=index+1: createDialogElement(container, index, "textForeColor"     , index.toString, nothing)
    .br(5)
    '' index=index+1: createDialogElement(container, index, "DIALOG"            , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "edit/view/hidden"  , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "labBgColor"        , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "labForeColor"      , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "textBgColor"       , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "textForeColor"     , index.toString, nothing)
    '' .br(5)
    '' index=index+1: createDialogElement(container, index, "HELP"              , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "toolTip"           , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "text1"             , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "text2"             , index.toString, nothing)
    '' index=index+1: createDialogElement(container, index, "icon"              , index.toString, nothing)
    '' 'index=index+1: createDialogElement(container, index, "..."               , index.toString, nothing)
    '' 'index=index+1: createDialogElement(container, index, "..."               , index.toString, nothing)
    '' .br(1)
    ''  el=.addLabel  ("dummyForAutoScroll", "xxxx" ,  ,"",,,33,5) : el.left=-999  '...fügt noch etwas luft ein
    .AutoScroll = True
  end with

  
  with cStatbar
    .addButton ("cmdCopyIGrid_2"  , "Kopieren 2" )' , "#1DD910")
    .addButton ("cmdInsertIGrid_2"  , "Einfügen 2" )' , "#1DD910")
    .addButton ("cmdReadDefaultIGrid_2"  , "Lesen 2" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_2"  , "Speichern 2" )' , "#1DD910")
  end with


 '--> 
 '--> ... container_4    :z.Zeit ausgeblendet
  with container4
     cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                           cMain.backColor = ColorTranslator.FromHtml("#00f")
     cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                           cToolBar.backColor = ColorTranslator.FromHtml("#aaa")
     cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                           cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
  end with

  with  cToolBar
    .addButton ("cmdCheckIsInGrid_1"       , "Check1-IsInGrid" )' , "#1DD910")
    .addButton ("cmdCheckGridFields_1"     , "CheckGridFields" )' , "#1DD910")
  end with

  '--> ... ---splitContainer - 2
  el=cMain.addSplitcontainer("splMain2", "left", pLeft2, "right", pRight2, DockStyle.Fill)
  splitContainer2=el
  el.backColor=ColorTranslator.FromHtml("#88f")
  el.splitterDistance=55
  el.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
  'el.IsSplitterFixed = True



'--> ... --- pLeft2
  with pLeft2
    .resetControls (5,5)
    .backColor=ColorTranslator.FromHtml("#D2D2FF")
  '--> ... --- --- iGrid-3
    IGrid3 = New IgridEx
    .Controls.Add(IGrid3)
    IGrid3.Dock = DockStyle.Fill
    'IGrid3.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid3)
    with IGrid3
      .fileSpec=glob_defaultGridPath+"esGrid3.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      .GroupBox.Visible = True
      '.AllowDrop = True
      '.ReadOnly = True

      .Cols.Add("Nr",44)
      .Cols.Add("r",100)
      .Cols.Add("ZeilenInhalt",44)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check

      .rows.count = 95
    end with
  End With

 '--> ... --- pRight2
  with pRight2
    el=pRight2.addSplitcontainer("splMain22", "left", pLeft22, "right", pRight22, DockStyle.Fill)
    splitContainer2sub=el
    el.backColor=ColorTranslator.FromHtml("#ccc")
    el.splitterDistance=250
  end with
  
  '--> ... --- pLeft22
  with pLeft22
    .resetControls (5,5)
    .AutoScroll = True
    .backColor=ColorTranslator.FromHtml("#f55")
    
    dim container As ScriptedPanel=pLeft22 'ACHTUNG:...nicht vergessen
    dim index as integer=50
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)

  end with
  
  
  '--> ... --- pRight22
  with pRight22
    .resetControls (5,5)
    .AutoScroll = True
    .backColor=ColorTranslator.FromHtml("#9f9")

    dim container As ScriptedPanel=pRight22 'ACHTUNG:...nicht vergessen
    dim index as integer=70
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)
    index=index+1: createDialogElement(container, index, "labText_"+index.toString, index.toString, nothing)

    '' .addTextbox("xxx-test", Auto, "labText", 100, "#aaa" )
    '' el= .element("txtDesc_xxx-test")
    '' el.textAlign=System.Drawing.ContentAlignment.MiddleRight
    '' 
    '' '' el= .element("txt_xxx-test")
    '' '' el.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
  end with
  
  with cStatbar
    .addButton ("cmdcloneIGrid_1"       , "Clone 1" )' , "#1DD910")
    .addButton ("cmdcloneIGrid_2"       , "Clone 2" )' , "#1DD910")
    .addButton ("cmdReadDefaultIGrid_3"  , "Lesen 3" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_3"  , "Speichern 3" )' , "#1DD910")
   end with

  
 '-->  
 '--> ... container_5   : DUMP - CODE _____________
  with container5
    cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                         'cMain.backColor = ColorTranslator.FromHtml("#aaa")
    cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                         cToolBar.backColor = ColorTranslator.FromHtml("#aaa")
    cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                         cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
  end with

  with cMain
  '--> ...textarea
    .insertX = 5 :.insertY = 0' 110
    .addTextbox ("dumpMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
     el=panelRef.element("txt_dumpMonitor")
     myTextArea=el
     el.Dock = DockStyle.Fill
     el.WordWrap=false
     el.scrollbars = System.Windows.Forms.ScrollBars.Both
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.bringToFront()
     
     
     '--> ... ... el.text=getMiniHelpText()
     '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
  end with

  with  cToolBar
    '.addButton ("cmdConnectServer"  , "Verbinden" )' , "#1DD910")
    '.addButton ("cmdSendTestMail"  ,         "test1: (senden)"  )', "#1DD910")
   end with




 '-->  
 '--> ... container_6   : Grab - CODE _____________
  with container6
    cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                         'cMain.backColor = ColorTranslator.FromHtml("#aaa")
    cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                         cToolBar.backColor = ColorTranslator.FromHtml("#aaa")
    cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                         cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
  end with

  with cMain
  '--> ...textarea
    .insertX = 5 :.insertY = 0' 110
    .addTextbox ("grabMonitor" ,  -2         , "inBox"    , 0  , "#FFFF99", , , "multiline=240")
     el=panelRef.element("txt_grabMonitor")
     myTextArea2=el
     el.Dock = DockStyle.Fill
     el.WordWrap=false
     el.scrollbars = System.Windows.Forms.ScrollBars.Both
     el.Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
     el.bringToFront()
     
     
     '--> ... ... el.text=getMiniHelpText()
     '' ??? addHandler textarea.KeyDown, AddressOf  textarea_KeyDown
  end with

  with  cToolBar
    '.addButton ("cmdConnectServer"  , "Verbinden" )' , "#1DD910")
    '.addButton ("cmdSendTestMail"  ,         "test1: (senden)"  )', "#1DD910")
   end with
  
  
  tmr1 = New Timer
  tmr1.Interval = 1000
  tmr1.Enabled = false

  glob.readFormPos(FormRef)
  'glob.readTuttiFrutti(FormRef)

  readDefaultIGridFromIndex(1)
  readDefaultIGridFromIndex(2)
  readDefaultIGridFromIndex(3)
  
  with IGrid1
    .Cols(0).width=50
    .Cols(1).width=50
    .Cols(2).width=0
    .Cols(3).width=105
  end with  
  
  'onResizeControls()
  FormRef.show
  formRef.bringToFront
  formRef.topmost=true
End Sub



'-->
'--> P A N E L - H E L P E R


dim globTempHeaderData(200) as string
dim globTempHeaderCounter as integer =0

function createDialogElement(container as object, index as integer, labText as string, defaultValue as string, tagData as object)
    stripHeaderData(labText)
    ''zNN=1
    'dim labBgColor="#A1A1DB"
    'dim labBgColor="#629FFF"
    dim labBgColor="#3477D7"
    dim el as object
    dim labWidth as integer=120
    with container 
    dim id="dlg_"+index.toString
      .addTextbox(id, Auto, labText+"  ", labWidth, labBgColor)
      'trace "index",index
      'trace "zNN",zNN
      'trace "zBB(zLN)",zBB(zLN)
      el = .element("txt_"+id)
      el.text=defaultValue.toString
      el.AutoSize = false
      el.height=18
      'el.backColor=ColorTranslator.FromHtml("#E7E7FF")
      el.backColor=ColorTranslator.FromHtml("#fff")
      ':el.TextAlign = System.Drawing.ContentAlignment.MiddleRight: err.number=0
      el.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      el.BorderStyle = System.Windows.Forms.BorderStyle.None      
      'label...
      el= .element("txtDesc_"+id)
      el.textAlign=System.Drawing.ContentAlignment.MiddleRight
      el.height=18
      el.foreColor=ColorTranslator.FromHtml("#fff")
      .br(20)
    end with
end Function


sub stripHeaderData(labelText)
  globTempHeaderData(globTempHeaderCounter)=labelText
  globTempHeaderCounter=globTempHeaderCounter+1
end sub


Function getHeaderData() as string
  redim preserve globTempHeaderData(globTempHeaderCounter)
  globTempHeaderCounter=0
  dim RESULT = join(globTempHeaderData,vbNewLine)
  redim globTempHeaderData(200)
  return RESULT
end function


sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    'el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


sub setGridHeaderData(grid as TenTec.Windows.iGridLib.iGrid, labelTextAsNewLineString as string)
  dim DATA() as string = split(labelTextAsNewLineString, vbNewLine)
  dim max as integer=DATA.length
  grid.Cols.Count=max
 
  dim i as integer
  for i = 0 to max-1
    grid.cols(i).text=DATA(i)
    grid.cells(0,i).value=DATA(i)
  next
   
  'dim grid as TenTec.Windows.iGridLib.iGrid = iGrid1
  'Dim row As TenTec.Windows.iGridLib.iGRow
  'dim curRow As TenTec.Windows.iGridLib.iGRow
  'dim curRowIndex as integer
  
  '' curRow=grid.curRow
  '' curRowIndex=grid.curRow.index
  '' trace "curRow", curRow
  '' curRow.cells(cInt(index)).value=value
  '' 
  '' dim lineData as string = JoinIGridLine(grid.curRow)
  '' monitorAdd(lineData)
  '' dbUpdateFromLineData(lineData)
  '' with grid
  ''   grid.cols(0).text="000"
  ''   .cols(1).text="111"
  ''   .cols(2).text="222"
  ''   
  '' end with
  '' 

end sub




'--> 
'--> E V E N T S  1 -------------------------


Sub onButtonEvent(e)
  setMonitorRef()
  clearAll
  statustext("OK")
  errorText("")
  printLine 11, "e.Sender.Name" , e.Sender.Name
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  MonitorAdd("======================== onLabelEvent")
  MonitorAdd("SenderFullName: " + e.Sender.Name)
  MonitorAdd("___MouseButton: " + e.MouseButton)
  MonitorAdd("________Action: " + action)
  
  '' if action.contains("cmdConnectServer") then connectToServer()
  '' if action.contains("cmdSendTestMail") then sendTestMail()
 
  callCmdByName(e)

end Sub


sub cmdSwitchContainer(e)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer =val(name(2))
  switchContainer(index)
end sub


sub switchContainer(index as integer)
 container1.visible=false
  container2.visible=false
  container3.visible=false
  container4.visible=false
  container5.visible=false
  container6.visible=false
  if index=1 then container1.visible=true
  if index=2 then container2.visible=true
  if index=3 then container3.visible=true
  if index=4 then container4.visible=true
  if index=5 then container5.visible=true
  if index=6 then container6.visible=true
end sub



sub callCmdByName(e)
  on error resume next
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim funcName as string=name(1)
  trace "funcName", funcName
  dim i as integer=1
  dim timerStart = GetTime()
  dim deltaTime as integer
  
  dim scriptClass= Me
  Dim myType As Type = scriptClass.GetType()
  Dim myMethod As System.Reflection.MethodInfo = myType.GetMethod(funcName)
  if myMethod.toString = "" then
    dim errMes="ERR: Sub '"+funcName+"' nicht gefunden" 
    statustext(errMes)
    dumpUnknownEventHandler(funcName)
    exit sub
  end if
  
  monitorAdd("______procName: "+myMethod.toString)
  monitorAdd("AAA: "+err.number.tostring)
  monitorAdd("BBB: "+err.description)

  dim args(0)' 
  args(0)=e
  myMethod.Invoke(scriptClass, args)
  'CallByName(scriptClass, funcName, Microsoft.VisualBasic.CallType.Method)

  deltaTime=GetTime()-timerStart
  monitorAdd("--------------- DONE")
  monitorAdd("anz schleifen je sek: "+(i / deltaTime * 1000).toString("0000"))
  monitorAdd("")
end sub




sub dumpUnknownEventHandler(funcName)
  dim tpl=getTemplateUnknownEventHandler()
  tpl=replace(tpl,"||FUNC-NAME||",funcName)
  MonitorAdd(tpl)
end sub


function getTemplateUnknownEventHandler()
#Data myData as String

ERR: EventHandler für '||FUNC-NAME||' nicht gefunden:
§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§


Sub ||FUNC-NAME||(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! '||FUNC-NAME||': ...in Arbeit")
  'msgBox("'||FUNC-NAME||': " + "...Coming soon ")
  '...
End Sub



§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
#EndData
  dim RESULT as string=myData 
  return RESULT
end function






'--> 
'--> E V E N T S  2 -------------------------
'' 
'' sub txt_dlg_TextChanged(e)                                 ' ...der tut leider nicht
''   MonitorAdd("SenderFullName: " + e.Sender.Name)
'' end sub



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
  if action="dlg" and e.eventName = "KeyDown" then

    syncIGrid(e, index, value)
  end if

  '' MonitorAdd("======================== onLabelEvent")
  '' MonitorAdd("SenderFullName: " + e.Sender.Name)
  '' MonitorAdd("___MouseButton: " + e.MouseButton)
  '' MonitorAdd("________Action: " + action)
end sub


'--> handles igrid1 und iGrid2
'sub IGrid1_updateUserFeedback2(ByRef grid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback2, iGrid2.updateUserFeedback2
sub IGrid1_updateUserFeedback2(ByRef grid As iGridEx, ByVal para1 As String) Handles iGrid1.updateUserFeedback2, iGrid2.updateUserFeedback2
  setMonitorRef()
  MonitorAdd("Feedback2: " + para1)
  PanelRef.element("txt_dlg_2").text="--- "+para1+" ---"

  dim curRow As TenTec.Windows.iGridLib.iGRow
  dim curRowIndex as integer
  curRow=grid.curRow
  curRowIndex=grid.curRow.index
  trace "curRow", curRow.index
  dim max as integer=grid.cols.count
  dim i as integer
  dim item as string
  
  'msgBox(grid.name)
  dim deltaIndex as integer=0
  if grid.name="IGrid2" then deltaIndex=100
  
  : for i = 0 to max-1
  :   item=curRow.cells(i).value
  :   'monitorAdd(i.toString+"="+item)
  :   PanelRef.element("txt_dlg_"+(i+deltaIndex).toString).text=item
  : next 

end sub


sub syncIGrid(e as object, index as string, value as string )
  dim grid as TenTec.Windows.iGridLib.iGrid
  monitorAdd(index)

dim deltaIndex as integer=0
  if index>50 then
     grid=iGrid2
     deltaIndex=-100
  else
     grid=iGrid1
  end if
  
 '... der geht hier natürlich nicht: ...grid=getIGridFromIndex(index)
  'Dim row As TenTec.Windows.iGridLib.iGRow
  dim curRow As TenTec.Windows.iGridLib.iGRow
  dim curRowIndex as integer
  curRow=grid.curRow

  curRowIndex=grid.curRow.index
  trace "curRow", curRow.index
  curRow.cells(cInt(index+deltaIndex)).value=value
  
  '' dim lineData as string = JoinIGridLine(grid.curRow)
  '' monitorAdd(lineData)
  '' dbUpdateFromLineData(lineData)
end sub







'-->
'--> C M D - Form / iGrid



Function getIGridFromIndex(index as integer) As IgridEx 
  if index=1 then  return  iGrid1
  if index=2 then  return  iGrid2
  if index=3 then  return  iGrid3
end function






Sub cmdReadDefaultIGrid(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  readDefaultIGridFromIndex(index)
End Sub

Sub readDefaultIGridFromIndex(index as integer)
  dim grid As IgridEx = getIGridFromIndex(index)
  dim fileSpec=grid.fileSpec  
  Dim fileContent = IO.File.ReadAllText(fileSpec)
  'Dim out() = Split(fileContent, vbNewLine)
  IGrid_put2(grid, fileContent)
end sub

Sub cmdSaveDefaultIGrid(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '' statustext("!!! 'cmdSaveDefaultIGrid': ...in Arbeit")
  '' 'msgBox("'cmdSaveDefaultIGrid': " + "...Coming soon ")
  dim grid As IgridEx = getIGridFromIndex(index)
  dim fileSpec=grid.fileSpec  
  Dim data As String
  IGrid_get (grid, data)
  dim lines() as string = split(data,vbNewLine)
  IO.File.WriteAllLines(fileSpec, lines)
End Sub


Sub cmdCopyIGrid(e)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer =val(name(2))
  dim data as string
  
  dim grid  As IgridEx  = getIGridFromIndex(index)
  Igrid_get(grid, data)
  zz.setClipboardText(data)
End Sub


Sub cmdInsertIGrid(e) 
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  
  dim data as string=zz.getClipboardText()
  dim grid As IgridEx = getIGridFromIndex(index)
  Igrid_put2(grid, data)
End Sub



Sub cmdEscape(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer =val(name(2))
  '...
  statustext("!!! 'cmdEscape(e)': ...in Arbeit")
  'msgBox("'cmdEscape(e)': " +    "...Coming soon ")
  '...
End Sub










'-->
'--> C M D - Datenbank


Sub cmdUpdateTest(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdUpdateTest': ...in Arbeit")
  'msgBox("'cmdUpdateTest': " + "...Coming soon ")

  dim grid as TenTec.Windows.iGridLib.iGrid = iGrid1
  'Dim row As TenTec.Windows.iGridLib.iGRow
  dim curRow as integer
  curRow=grid.curRow.index
  trace "curRow", curRow
  
  dim lineData as string = JoinIGridLine(grid.curRow)
  monitorAdd(lineData)
  '' dbUpdateFromLineData(lineData)

End Sub



Sub cmdUpdateAll(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdUpdateAll': ...in Arbeit")
  'msgBox("'cmdUpdateAll': " + "...Coming soon ")

  dim grid as TenTec.Windows.iGridLib.iGrid = iGrid1
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim lastSelRow As TenTec.Windows.iGridLib.iGRow = nothing
 
  '' dim curRow as integer
  '' curRow=grid.curRow.index
  '' trace "curRow", curRow
  '' 
  
  dim lineData as string
  For Each row In grid.Rows
    if lastSelRow isNot nothing then lastSelRow.selected=false
    lastSelRow=row
    row.Selected=true
    row.EnsureVisible
    lineData = JoinIGridLine(row)
    dbUpdateFromLineData(lineData)
    
    'monitorAdd(lineData)
    'trace "lineData", lineData
   Next

End Sub




sub dbUpdateFromLineData(lineData as string)
 dim DATA() as string = split(lineData, vbTab)
  dim item as string
  dim max as integer=DATA.length
  dim i as integer
  dim iNummer as string
  dim iNummerID as integer
  iNummer=DATA(0)
  
  if iNummer="ID" then exit sub
  if trim(iNummer)="" then exit sub
  
  
      'dbAdr muß offen sein
    '=======================================
    trace "check, ob dbADR offen ist" ,"dbADR.name"   'wie testen ohne "nothing"
    
    dim myPath as string=dbFileSpec
    monitorAdd("myPath: "+myPath)
    monitorAdd("iNummer: "+iNummer)
    if dbADR is nothing then
    
    :   dbADR = CreateObject("ADODB.connection")
    :   dbADR.Open("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & myPath)
    ':   If Err.Number<>0 Then Trace "ErrDbOpen: ", Err.Description
    :   Err.Clear()
    end if


    dim RS = CreateObject("ADODB.recordset")
    With RS
      '...forUpdate ???
      Dim cursorType     = 2
      Dim lockType       = 3
      Dim cursorLocation = 1
      .cursorType = cursorType      '
      .lockType = lockType            '
      .cursorLocation = cursorLocation  '??? Client
      
      dim SQL as string
      '' dSQL = "SELECT * FROM [" + tableName + "] Where 1=2"
      '' Trace "!!! SQL", SQL
      '' 
      iNummer = Replace(iNummer, "i-", "")
      iNummerID = (iNummer)
      Dim ss
      ss = ""
      ss = ss & "SELECT *"
      ss = ss & " FROM [UGB-ADR-TAB]"
      ss = ss & " WHERE"
      ss = ss & " ((([UGB-ADR-TAB].ID) =" & iNummerID & "))"
      'ss = ss & " ORDER BY [UGB-ADR-TAB].NAME, [UGB-ADR-TAB].VORNAME "
      ss = ss & " ;"
      SQL = ss
      Trace "SQL:" + SQL

      .Open(SQL, dbADR)   ', dbRead, cursorType ,lockType , CursorLocation   ' adCmdText
    end with
    
    Dim maxFields as integer
    With RS
      maxFields = .Fields.Count
      
      For i = 1 To maxFields - 1
        item=DATA(i)
        Trace i.toString + "   "+.Fields(i).Name ,  item
        .Fields(i).value=item
        'logfile CStr(i) + "  " + .fields(i).name
      Next
       ''.Fields(149).value="XXXX"
       .update
      Trace "maxFields", maxFields
    End With


  
  
  
  
  
  
  '' 'monitorAdd("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$4")
  '' monitorAdd("i-nummer: "+iNummer)
  '' for i=1 to max-1
  ''   item=DATA(i)
  ''   'monitorAdd(i.toString+"   "+item)
  '' next
  'monitorAdd("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$4")
end sub






Sub cmdCheckGridFields(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdCheckGridFields': ...in Arbeit")
  'msgBox("'cmdCheckGridFields': " + "...Coming soon ")
  '...
  dim id as string=iGrid3.curRow.cells(0).value
  monitorAdd("curId: "+id)
  
  dim idSuch=id
  dim lineData1 as string=""
  dim found1 as boolean
  dim DATA1() as string
  
  found1= ___CheckIsInGrid2(iGrid1 , idSuch,"",lineData1)
  if found1 = true then
    DATA1=split(lineData1, vbTab)
  end if
  dim lineData2 as string=""
  dim found2 as boolean
  dim DATA2() as string
  found2= ___CheckIsInGrid2(igrid2 , idSuch,"",lineData2)
  if found2 = true then
    DATA2=split(lineData2, vbTab)
  end if
  
  monitorAdd("AAA: "+lineData1)
  monitorAdd("BBB: "+lineData2)
  
  pLeft22.resetControls (5,5)
  pRight22.resetControls (5,5)
  
  if found1 = true and found2 = true then
    if lineData1 <> lineData2 then
      dim i as integer
      dim max as integer=DATA1.length
      for i=0 to max-1
        dim item1 as string=DATA1(i)
        dim item2 as string=DATA2(i)
        if item1 <> item2 then
          monitorAdd(i.toString+": "+item1)
          monitorAdd(i.toString+": "+item2)
          createDialogElement(pLeft22,  i , i.toString, item1, nothing)
          createDialogElement(pRight22, i , i.toString, item2, nothing)
        end if
      next
    end if
  end if
  '' dim curRow as integer
  '' curRow=me.curRow.index
  '' trace "curRow", curRow
  '' 

End Sub
  

Sub cmdCheckIsInGrid(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))

  ''statustext("!!! 'cmdCheckIsInGrid': ...in Arbeit")

  dim grid1 As IgridEx = iGrid1
  dim grid2 As IgridEx = iGrid2
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim isInGrid2 as boolean
  dim lineData as string
  dim id as string=""
  dim i as integer
  dim max as integer=grid1.rows.count
  dim OUT(max) as string
  dim counter as integer=0
  for i=0 to max-1
    row=grid1.rows(i)
    id=row.cells(0).value
    lineData=JoinIGridLine(row)
    isInGrid2 = ___CheckIsInGrid2(iGrid2,id, lineData)
    if isInGrid2 = true then
      OUT(counter)=id
      counter=counter+1
      monitorAdd(id)
    end if
  next
End Sub


Function ___CheckIsInGrid2(grid as IgridEx, idSuch as string, lineData as string, optional  byRef getLineData as string="") as boolean ' ...
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim id as string=""
  dim i as integer
  dim myLineData as string
  dim max as integer=grid.rows.count
  for i=0 to max-1
    row=grid.rows(i)
    id=row.cells(0).value
    if id=idSuch then 
      myLineData=JoinIGridLine(row)
      getLineData=myLineData
      if myLineData <> lineData then
        monitorAdd("grid-1: "+lineData)
        monitorAdd("grid-2: "+myLineData)
        return true
      else
        return false
      end if
    end if
  next
  return false
end Function





Sub cmdNavigateINummer(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdNavigateINummer': ...in Arbeit")
  'msgBox("'cmdNavigateINummer': " + "...Coming soon ")
  '...
End Sub




Sub cmdcloneIGrid(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer =val(name(2))
  dim data as string
  dim grid  As IgridEx  = getIGridFromIndex(index)
  Igrid_get(grid, data)
  IGrid_put2(IGrid3, data)
End Sub



'-->

'-->
'--> C M D - eMail-Massenversand


Sub cmdConnectServer(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdConnectServer': ...in Arbeit")
  'msgBox("'cmdConnectServer': " + "...Coming soon ")
  '...
End Sub



Sub cmdSendTestMail(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  '...
  statustext("!!! 'cmdSendTestMail': ...in Arbeit")
  'msgBox("'cmdSendTestMail': " + "...Coming soon ")
  '...
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
  '... dann sortiert er nicht
  e.DoDefault = False
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
'--> --- DB ---------------------------------------------
'--> ... .... tut noch nicht richtig - alter Mist

  Private dbADR As Object = nothing
  Private rsADR As Object 'RecordSet
  Private Const tableName = "UGB-ADR-TAB" 'ohne Klammer!!!

  Private Function dbFileSpec() As String
    'dbFileSpec = "U:\" + "SVS\DB\es-adr-test2.mdb"
    dbFileSpec = "C:\" + "_Alias\u\SVS\DB\es-adr-test2.mdb"
    dbFileSpec = "C:\" + "backupDB\test\es-adr-test2.mdb"
    dbFileSpec = "U:\" + "SVS\DB\es-adr-test2.mdb"
  End Function


Function zoomAdr(ByVal iNummer As String) As Object 'RecordSet
    On Error Resume Next
    Dim Modus
    Dim ctlName As String
    Dim index As Integer
    Dim dbFeldName As String
    Modus = "doppeltGeblockt"

    navigateADR(rsADR, iNummer, Modus)

    zoomAdr = rsADR 'geht das ???
  End Function
  
  
  
  Sub testAdr()
    On Error Resume Next
    Dim iNummer
    Dim Modus
    Modus = "doppeltGeblockt"
    iNummer = "1111"
    ''navigateADR(rsADR, iNummer, Modus)
    navigateADR(rsADR, "1110", Modus)
    ''navigateADR(rsADR, 1100, Modus)
    Trace rsADR.Fields.Count
    'closeDbAdr
  End Sub
  
  
  Sub closeDbAdr()
    dbClose(dbADR, rsADR)
  End Sub


  Sub navigateADR(ByRef RS, ByVal iNummer, ByVal Modus)
    On Error Resume Next
    'Stop
    Dim SQL
    Dim iNummerID
    If iNummer = "" Then Exit Sub
    If Len(iNummer) > 8 Then Exit Sub
    If InStr(iNummer, ".") > 0 Then Exit Sub
    '...oder generell ausschalten abhängig von Typ ???


    iNummer = Replace(iNummer, "i-", "")
    iNummerID = (iNummer)
    Dim ss
    ss = ""
    ss = ss & "SELECT *"
    ss = ss & " FROM [UGB-ADR-TAB]"
    ss = ss & " WHERE"
    ss = ss & " ((([UGB-ADR-TAB].ID) =" & iNummerID & "))"
    ss = ss & " ORDER BY [UGB-ADR-TAB].NAME, [UGB-ADR-TAB].VORNAME "
    ss = ss & " ;"
    SQL = ss
    Trace "SQL:" + SQL

    Dim myPath : myPath = dbFileSpec + ""


    '=======================================
    If dbADR Is Nothing Then
       trace "!!! If dbADR Is Nothing Then", "dbOpenForUpdate..."
      'DBopen(dbADR, dbFileSpec, tableName)
      'dbOpenForUpdate(dbADR, dbFileSpec,  tableName)
      : dbADR = CreateObject("ADODB.connection")
      : dbADR.Open("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & myPath)
      : If Err.Number<>0 Then 
      :   Trace "ErrDbOpen: ", Err.Description
      : end if
      : Err.Clear()
    End If
    '=======================================


    ''Modus ="UPDATE"
    RS = CreateObject("ADODB.recordset")
    With RS
      '...forUpdate ???
      Dim cursorType = 2
      Dim lockType = 3
      Dim cursorLocation = 1
      .cursorType = cursorType      '
      .lockType = lockType            '
      .cursorLocation = 1  '??? Client
      .Open(SQL, dbADR) '!!! 
      '''', dbRead, cursorType ,lockType , CursorLocation   ' adCmdText
    end with
 
    Trace "navigateADR..."
    With RS
      Dim i, maxFields
      maxFields = .Fields.Count
      'For i = 0 To maxFields - 1
      For i = 0 To 13
        Trace i.toString + "   "+.Fields(i).Name , .Fields(i).value
        'logfile CStr(i) + "  " + .fields(i).name
      Next
       .Fields(13).value="xxxxxxxxxxxxxxxxxxxxxxx"
       .update
      Trace "maxFields", maxFields
    End With

  End Sub



  '#########################################################
  '#########################################################
  '#########################################################


  ''Sub navigateSQL(ByRef db As Object, ByRef RS As Object, ByVal SQL As String, ByVal Modus As String)
  Sub navigateSQL(ByVal db As Object, ByVal RS As Object, ByVal SQL As String, ByVal Modus As String)
    'Stop
    'On Error Resume Next
    'db muß bereits offen sein
    : RS.Close()    : Err.Clear()
    trace "111 navigateSQL", db.toString
    trace "222 navigateSQL", rs.toString
    trace "333 navigateSQL", sql
    trace "444 navigateSQL", modus
    Try
      RS = CreateObject("ADODB.recordset")
    Catch ex As Exception
      MsgBox("(es) --> navigateSQL: ..." + Err.Description)
      Err.Clear()
      Exit Sub
    End Try
    With RS
      Dim cursorType
      Dim lockType
      Dim cursorLocation

      Modus = "READ"
      If Modus = "READ" Then
        cursorType = 3 'adOpenStatic
        lockType = 1   'adLockReadOnly
      Else
        cursorType = 3 'adOpenStatic
        lockType = 3   'adLockOptimistic
      End If
      'cursorType =  0 'adOpenForwardOnly
      'cursorType =  1 'adOpenKeyset
      'cursorType =  2 'adOpenDynamic
      'cursorType =  3 'adOpenStatic

      'lockType    = 1 'adLockReadOnly
      'lockType    = 2 'adLockPessimistic
      'lockType    = 3 'adLockOptimistic
      'lockType    = 4 'adLockBatchOptimistic
      .cursorType = cursorType      '
      .lockType = lockType            '
      .cursorLocation = 1  '??? Client
      cursorLocation = 1
      ''''.source=db
      ''''.Open ss, dbZE,3 ,3 , 1   ' adCmdText

      ''.Open  SQL, db, cursorType ,lockType , CursorLocation   ' adCmdText
      Try
        .Open(SQL, db)   ', dbRead, cursorType ,lockType , CursorLocation   ' adCmdText
      Catch ex As Exception
        Trace "navigateSQL: ", Err.Description
        'Stop
        ' errlog "SQL: ", SQL
        ' errlog "", ""
      End Try
      'Stop
    End With
  End Sub

  ''==========================
  ''_ with rsZE
  ''_     .CursorType = 3  '
  ''_                     '0 adOpenForwardOnly
  ''_                     '1 adOpenKeyset
  ''_                     '2 adOpenDynamic
  ''_                     '3 adOpenStatic
  ''_
  ''_   .LockType =  3   '
  ''_                     '1 adLockReadOnly
  ''_                     '2 adLockPessimistic
  ''_                     '3 adLockOptimistic
  ''_                     '4 adLockBatchOptimistic
  ''_   .CursorLocation=1 '??? Client
  ''_   .source=dbZE
  ''_     .Open ss, dbZE,3 ,3 , 1   ' adCmdText
  ''_
  ''==========================


  '#########################################################
  '#########################################################
  '#########################################################



sub dbOpenForUpdate(ByRef db As Object, ByVal dbFileSpec As String, ByVal tableName As String)
    'dbAdr muß offen sein
    '=======================================
    trace "check, ob dbADR offen ist" ,"dbADR.name"   'wie testen ohne "nothing"
    
    dim myPath as string=dbFileSpec

:   db = CreateObject("ADODB.connection")
:   db.Open("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & myPath)
':   If Err.Number<>0 Then Trace "ErrDbOpen: ", Err.Description
:   Err.Clear()



    dim RS = CreateObject("ADODB.recordset")
    With RS
      Dim cursorType
      Dim lockType
      Dim cursorLocation

      '...forUpdate ???
      cursorLocation = 1
      cursorType=2
      lockType=3
      .cursorType = cursorType      '
      .lockType = lockType            '
      .cursorLocation = 1  '??? Client
      
      dim SQL as string = "SELECT * FROM [" + tableName + "] Where 1=2"
      Trace "!!! SQL", SQL

      .Open(SQL, db)   ', dbRead, cursorType ,lockType , CursorLocation   ' adCmdText

    end with
    With RS
      Dim i, maxFields
      maxFields = .Fields.Count
      For i = 0 To maxFields - 1
        Trace i.toString + "   "+.Fields(i).Name ,  "" '.Fields(i).value
        'logfile CStr(i) + "  " + .fields(i).name
      Next
       ''.Fields(149).value="XXXX"
       ''.update
      Trace "maxFields", maxFields
    End With

end sub


Sub DBopen(ByRef db As Object, ByVal dbFileSpec As String, ByVal tableName As String)
    'On Error Resume Next
    'Stop
    Dim myPath : myPath = dbFileSpec + ""
    Trace "myPath???", "file:" + myPath
    Dim SQL
    'Dim RS as object = Nothing
    dim RS as object = CreateObject("ADODB.recordset")


    SQL = "SELECT * FROM [" + tableName + "] Where 1=2"
    Trace "SQL", SQL
    Trace "START DB_open..  ", Now()

    
:   db = CreateObject("ADODB.connection")
:   db.Open("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & myPath)
':   If Err.Number<>0 Then Trace "ErrDbOpen: ", Err.Description
:   Err.Clear()
    
    trace "!!! db", db.toString
    trace "!!! rs", rs.toString
    trace "!!! sql", sql
    
    navigateSQL(db, RS, SQL, "READ")

    '___Felder anzeigen lassen ...
    With RS
      Dim i, maxFields
      maxFields = .Fields.Count
      For i = 0 To maxFields - 1
        Trace i, .Fields(i).Name
        'logfile CStr(i) + "  " + .fields(i).name
      Next
      Trace "maxFields", maxFields
    End With
    '  Stop
  End Sub


  '#########################################################
  '#########################################################
  '#########################################################


  Sub dbClose(ByRef db As Object, ByRef RS As Object)
    On Error Resume Next
    : RS.Close() : Err.Clear()
    : RS = Nothing : Err.Clear()
    : db = Nothing : Err.Clear()
  End Sub












'-->
'--> S E N D - M A I L -------------------------

'#Imports TenTec.Windows.iGridLib
#Imports System.Net.Mail


' Dim client As New SmtpClient
Dim client As SmtpClient
Dim senderData As String

dim glob_senderMail as String = "dev@ugb.de"
dim glob_senderName as string = "scriptHost"



sub sendTestMail()
  msgBox("sendTestMail")
  'exit sub
  
  'connectToServer()
  
  Dim subj = "betreff_test"
  Dim body = "abc"+vbNewLine+"bbb"+vbNewLine+"ccc"
  Dim tAdr As String = "Elmar.Schropp@ugb.de"
  Dim tName As String = "an Elmar Schropp"
  '' 
  '' ''sendMailSyntax1(ig, subj, body, tAdr, tName)
  sendMailSyntax2(nothing, subj, body, tAdr, tName)
  '' 
end sub


sub connectToServer()
  msgBox("connectToServer")
  'exit sub
  
  dim host as string ="212.223.100.233"
  dim user as string ="323789@mxrelay1.nbg.internet1.de"
  dim pass as string ="5iHKEC1Dv@"
  connectServer(host, user, pass)
end sub



Function MyMailClient() As SmtpClient
  if client is nothing then client= new SmtpClient
  Return client
End Function



Sub sendMailForIGridLine(ByVal ig As iGRow)
  '' Dim subj = Form1.txtBetreff.Text
  '' Dim body = Form1.txtBody.Text
  '' Dim tAdr As String = ig.Cells("mail").Value
  '' Dim tName As String = ig.Cells("targetname").Value
  '' 
  '' ''sendMailSyntax1(ig, subj, body, tAdr, tName)
  '' sendMailSyntax2(ig, subj, body, tAdr, tName)
  '' 
End Sub

'Sub sendMailSyntax2(ByVal ig As iGRow, ByVal subj As String, ByVal body As String, ByVal mailAdr As String, ByVal name As String)
Sub sendMailSyntax2(ByVal ig As object, ByVal subj As String, ByVal body As String, ByVal mailAdr As String, ByVal name As String)

  If Trim(mailAdr) = "" Then
    monitorAdd("LEER: <--")
    Exit Sub
  End If

  Dim target As String = mailAdr + " <" + name + ">"
  monitorAdd("target: " + target)
  ''monitorAdd("tName: " + tName)
  ''monitorAdd("target: " + target)
  monitorAdd("<--")
  
    ''Dim strFrom As String = "bla@bla.com"
    ''Dim strSubj As String = "bla bla: " & strDatum
    ''Dim strBody As String = "bla bla" & strNaam & strVnaam

    Dim strFrom As String = mailAdr
    Dim strSubj As String = subj '' "bla bla: " & strDatum
    Dim strBody As String = body '' "bla bla" & strNaam & strVnaam

    'Create instance of MailMessage
:   Dim insMail As New MailMessage()
:
:    'Create from address item
:    Dim senderAddress As New MailAddress(glob_senderMail, glob_senderName)
:    'create to address item
:    Dim sendTo As New MailAddress(mailAdr, name)
:
:
:    'Build the email
:    With insMail
:      .From = senderAddress
:      .To.Add(sendTo)
:      .Subject = strSubj
:      .Body = strBody
:    End With
:
:    ' ''Create instance of the SMTP Client
:   ''Dim smtp As SmtpClient = New SmtpClient("mysmtp")
:    ''smtp.Credentials = New Net.NetworkCredential("myuser", "mypass")
:
:    'Send the email
:    '' smtp.Send(insMail)
:    client.Send(insMail)
:
:    'Clean up
:    '' insMail.Dispose() '''??? macht der auch den client zu ??? 
:
:    '' Return True
:    'msgBox("OK: " + target)
:   monitorAdd("DONE ??? :"+err.description)
  if err.number<>0 then
  'Catch ex As Exception
    'Display the error message and return false
    'MessageBox.Show("Error Sending Mail:" & vbNewLine & ex.Message)
    Dim msg = err.description
    monitorAdd("")
    monitorAdd(target)
    monitorAdd("ERR: " + Mid(msg, 1, 55))
    'ig.Cells("err").Value = msg.Replace(vbNewLine, "|ZS|")
    
    'Form1.txtErrMes.Text = msg
    showErrMes(msg)
    
    '' Return False
  end if

End Sub



:Sub sendMailSyntax1(ByVal ig As iGRow, ByVal subj As String, ByVal body As String, ByVal mailAdr As String, ByVal name As String)

  If Trim(mailAdr) = "" Then
    monitorAdd("LEER: <--")
    Exit Sub
  End If

  Dim target As String = mailAdr + " <" + name + ">"
  monitorAdd("target: " + target)
  ''monitorAdd("tName: " + tName)
  ''monitorAdd("target: " + target)
  monitorAdd("<--")
  Try
    sendEmail(subj, body, target)
    '' sendEmail(subj, body, target2)     ''falls NAME sonderzeichen enthällt 
    monitorAdd("OK: " + target)
  Catch ex As Exception
    Dim msg = ex.ToString
    monitorAdd("")
    monitorAdd(target)
    monitorAdd("ERR: " + Mid(msg, 1, 55))
    ig.Cells("err").Value = msg.Replace(vbNewLine, "|ZS|")

   'Form1.txtErrMes.Text = msg
    showErrMes(msg)
  End Try

End Sub

Sub setSenderData(ByVal senderName As String, ByVal senderAddress As String)
  glob_senderMail = senderAddress
  glob_senderName = senderName
  senderData = senderName + " <" + senderAddress + ">"
End Sub


Sub connectServer(ByVal host As String, ByVal user As String, ByVal pass As String)
  monitorAdd("host: " + host)
  monitorAdd("user: " + user)
  monitorAdd("pass: " + pass)
 
  '...provisorisch
  if client is nothing then client= new SmtpClient
  
  client.Host = host
  client.Credentials = New System.Net.NetworkCredential(user, pass)
  'client.EnableSsl = True

End Sub

Sub sendEmail(ByVal subject As String, ByVal body As String, ByVal targetMail As String)

  client.Send(senderData, targetMail, subject, body)

End Sub


sub showErrMes(msg)
  msgbox(msg)
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




