
' ugb_emailMischer

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

'--> P A R A M E T E R
Const  winID = "{ScriptClass}.demoWin"
const  globHideOnClose as boolean = true           'bei true merkt er sich die Position
public winCaption as string="UGB-WerbeMailer: eMail und/oder Schneckenpost"
public glob_defaultGridPath as string = "C:\yPara\scriptIDE\gridTest\"

dim toolBarColor as string="#E4E1DC"

Dim WithEvents FormRef As Form
Dim PanelRef As ScriptedPanel

public globMonitorRef as object
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

Dim WithEvents IGrid1 As IgridEx
Dim WithEvents IGrid2 As IgridEx
Dim WithEvents IGrid3 As IgridEx


public shared splitContainer1  As Object
public shared splitContainer2  As Object
public shared splitContainer2sub  As Object



Dim pLeft1, pRight1 As ScriptedPanel
Dim pLeft2, pRight2 As ScriptedPanel
Dim pLeft22, pRight22 As ScriptedPanel

Dim WithEvents tmr1 As Timer
Dim glob As cls_globPara

Public Declare Function GetTime Lib "winmm.dll" Alias "timeGetTime" () As Integer





'-->
sub test1()
  printLine 3, "aaa","bbbbb"
  monitorAdd("---------------------------------------")
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
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)'  : err.number=0
  formRef = getOuterWindowRef(panelRef)
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose : err.number=0
  return panelRef
End Function



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
'--> A U T O S T A R T
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
    'toolbar1.BackColor = ColorTranslator.FromHtml("#8ff")

    'container1.top=112
    containerMain.resetControls  (10,5,1)
    containerMain.resetControls  (10,5,1)
    containerMain.BackColor = ColorTranslator.FromHtml("#fff")
    containerMain.BackColor = ColorTranslator.FromHtml("#f00")


    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml("#8f8")
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.height=22
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
 
    'DEFAULT-VISBLE
    container4.visible=true

 
  end with

  '--> ... toolbar
  with toolbar1
   .addButton ("cmdSwitchContainer_1"  , "Auswahl" )' , "#1DD910")
   .addButton ("cmdSwitchContainer_2"  , "Daten"  )', "#1DD910")
   .addButton ("cmdSwitchContainer_3"  , "Briefe" )' , "#1DD910")
   .addButton ("cmdSwitchContainer_4"  , "DB-TOOL"  )', "#1DD910")
   .addButton ("cmdSwitchContainer_5"  , "email"  )', "#1DD910")
   .addButton ("cmdEscape"   , "Beenden"  , "#f88")

  end with
  
  'With PanelRef
  '--> ... container_1 
  with  container1
    .resetControls (5,5)
    'dim el as object
    .activateEvents = "|ButtonMouseClick| |TEXTBOXTEXTCHANGED|"
    
    '--> ... --- splitContainer - 1)
    el=.addSplitcontainer("splMain", "left", pLeft1, "right", pRight1, DockStyle.Fill)
    splitContainer1=el
    el.backColor=ColorTranslator.FromHtml("#88f")
    el.splitterDistance=150
  end with 
    
     
  '--> ... --- Left Pane - 1
  with pLeft1
    .resetControls (5,5)
    Dim typName
    typName = "System.Windows.Forms.ListBox, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089, processorArchitecture=MSIL"
    .addControl ("ctrl_lv",typName)
    .Element("ctrl_lv").Top=5
    .Element("ctrl_lv").Width=200 : .Element("ctrl_lv").Height=400
    With .Element("ctrl_lv").Items
      .add ("test")
      .add ("test")
      .add ("test")
      .add ("test")
    End With
  end with
 
 
 
 
  '--> ... --- Right Pane - 2
  with pRight1
    .resetControls (5,5)
    .backColor=ColorTranslator.FromHtml("#bbf")
    .addIcon("icoSound", "http://www.iconfinder.net/data/icons/nuvola2/22x22/apps/kmix.png")
    .addTextbox("url", Auto, X:=40)
  end with
 
 
 '--> ... container_2
   with container2
     .backColor = ColorTranslator.FromHtml("#B1D8B2")
     cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                           'cMain.backColor = ColorTranslator.FromHtml("#faa")
     cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                           cToolBar.backColor = ColorTranslator.FromHtml("#BDBBB7")
     cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                           cStatBar.backColor = ColorTranslator.FromHtml("#BDBBB7")
   end with
   
   with cToolBar
    .addButton ("cmdUpdateTest_1"     , "DB-rück-test" )' , "#1DD910")
    .addButton ("cmdUpdateAll_1"     , "DB-rück-all" )' , "#1DD910")

   end with
   
   with cMain
   '--> ... --- iGrid-1
    IGrid1 = New IgridEx
    .Controls.Add(IGrid1)
    IGrid1.Dock = DockStyle.Fill
    'IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid1)
    with IGrid1
      .fileSpec=glob_defaultGridPath+"esGrid1.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      .GroupBox.Visible = True
      '.AllowDrop = True
      '.ReadOnly = True

      .Cols.Add("Nr",111)
      .Cols.Add("r",222)
      .Cols.Add("ZeilenInhalt",111)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check

      .rows.count = 100
    end with
  end with
  
  with cStatbar
    .addButton ("cmdCopyIGrid_1"         , "Kopieren" )' , "#1DD910")
    .addButton ("cmdInsertIGrid_1"       , "Einfügen" )' , "#1DD910")
    .addButton ("cmdReadDefaultIGrid_1"  , "Lesen" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_1"  , "Speichern" )' , "#1DD910")
  end with
 

 
 '--> ... container_3
   with container3
     cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                           cMain.backColor = ColorTranslator.FromHtml("#ccc")
     cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                           cToolBar.backColor = ColorTranslator.FromHtml("#FFFF88")
     cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                           cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
   end with

   with cMain
   '--> ... --- iGrid-2
    IGrid2 = New IgridEx
    .Controls.Add(IGrid2)
    IGrid2.Dock = DockStyle.Fill
    'IGrid2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid2)
    with IGrid2
      .fileSpec=glob_defaultGridPath+"esGrid2.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      .GroupBox.Visible = True
      '.AllowDrop = True
      '.ReadOnly = True

      .Cols.Add("Nr",111)
      .Cols.Add("r",222)
      .Cols.Add("ZeilenInhalt",111)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check

      .rows.count = 100
    end with
  end with
  
  with cStatbar
    .addButton ("cmdCopyIGrid_2"  , "Kopieren" )' , "#1DD910")
    .addButton ("cmdInsertIGrid_2"  , "Einfügen" )' , "#1DD910")
    .addButton ("cmdReadDefaultIGrid_2"  , "Lesen" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_2"  , "Speichern" )' , "#1DD910")
  end with


 
 '--> ... container_4
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



'--> ... --- Left Pane - 2
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

 '--> ... --- Right Pane - 2
  with pRight2
    el=pRight2.addSplitcontainer("splMain22", "left", pLeft22, "right", pRight22, DockStyle.Fill)
    splitContainer2sub=el
    el.backColor=ColorTranslator.FromHtml("#ccc")
    el.splitterDistance=250
  end with
  
  '--> ... --- Left Pane - 2-2
  with pLeft22
    .resetControls (5,5)
    .AutoScroll = True
    .backColor=ColorTranslator.FromHtml("#f55")
    
    dim container As ScriptedPanel=pLeft22
    dim index as integer=0
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
  
  
  '--> ... --- Right Pane - 2-2
  with pRight22
    .resetControls (5,5)
    .AutoScroll = True
    .backColor=ColorTranslator.FromHtml("#9f9")

    dim container As ScriptedPanel=pRight22
    dim index as integer=20
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
    .addButton ("cmdReadDefaultIGrid_3"  , "Lesen" )'    , "#1DD910")
    .addButton ("cmdSaveDefaultIGrid_3"  , "Speichern" )' , "#1DD910")
   end with

  
 
 '--> ... container_5
  with container5
    cMain     = .addPanel("containerMain"    , DockStyle.Fill)
                         'cMain.backColor = ColorTranslator.FromHtml("#aaa")
    cToolBar  = .addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
                         cToolBar.backColor = ColorTranslator.FromHtml("#aaa")
    cStatBar  = .addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
                         cStatBar.backColor = ColorTranslator.FromHtml("#aaa")
  end with

  with  cToolBar
    .addButton ("cmdConnectServer"  , "Verbinden" )' , "#1DD910")
    .addButton ("cmdSendTestMail"  ,         "test1: (senden)"  )', "#1DD910")
   end with
  
  tmr1 = New Timer
  tmr1.Interval = 1000
  tmr1.Enabled = false

  glob.readFormPos(FormRef)
  glob.readTuttiFrutti(FormRef)

  ___readDefaultIGridFromIndex(1)
  ___readDefaultIGridFromIndex(2)
  ___readDefaultIGridFromIndex(3)
  'onResizeControls()
  
End Sub


function createDialogElement(container as object, index as integer, labText as string, defaultValue as string, tagData as object)
    ''zNN=1
    dim labBgColor="#A1A1DB"
    dim el as object
    dim labWidth as integer=100
    with container 
    dim id="dlg_"+index.toString
      .addTextbox(id, Auto, labText, labWidth, labBgColor )
      'trace "index",index
      'trace "zNN",zNN
      'trace "zBB(zLN)",zBB(zLN)
      el = .element("txt_"+id)
      el.text=defaultValue.toString
      el.AutoSize = false
      el.height=20
      el.backColor=ColorTranslator.FromHtml("#E7E7FF")
      :el.TextAlign = System.Drawing.ContentAlignment.MiddleRight: err.number=0
      el.BorderStyle = System.Windows.Forms.BorderStyle.None      
      'label...
      el= .element("txtDesc_"+id)
      el.textAlign=System.Drawing.ContentAlignment.MiddleRight
      .br(21)
    end with
end Function











'-->
'--> --- DB ---------------------------------------------

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
'--> E V E N T S

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
  '' if action.contains("xxx") then connectToServer()
  '' if action.contains("xxx") then connectToServer()
  '' if action.contains("xxx") then connectToServer()
  '' 
  callCmdByName(e)

end Sub



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















'-->
'--> C M D


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
  dbUpdateFromLineData(lineData)

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



sub cmdSwitchContainer(e)
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer =val(name(2))
  container1.visible=false
  container2.visible=false
  container3.visible=false
  container4.visible=false
  container5.visible=false
  if index=1 then container1.visible=true
  if index=2 then container2.visible=true
  if index=3 then container3.visible=true
  if index=4 then container4.visible=true
  if index=5 then container5.visible=true
end sub


Sub cmdReadDefaultIGrid(e)  ' ...
  Dim name() = Split(e.Sender.Name+"_", "_")
  dim action=name(1)
  dim index as integer = val(name(2))
  ___readDefaultIGridFromIndex(index)
End Sub

Sub ___readDefaultIGridFromIndex(index as integer)
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
'--> C M D - H E L P E R


Function getIGridFromIndex(index as integer) As IgridEx 
  if index=1 then  return  iGrid1
  if index=2 then  return  iGrid2
  if index=3 then  return  iGrid3
end function



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
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByVal para1 As String)
Public Event updateUserFeedback2(ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByVal para1 As String) Handles iGrid1.updateUserFeedback1
'sub IGrid1_updateUserFeedback2(ByVal para1 As String) Handles iGrid1.updateUserFeedback2

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
'--> S E N D - M A I L 

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




