
'' es_UserToolbar2

#Para DebugMode internal
#Para SilentMode true

#Imports System.Diagnostics

#Reference WeifenLuo.WinFormsUI.Docking.dll
#Imports WeifenLuo.WinFormsUI.Docking.DockPanel

Public Const ToolbarID1 = "es_UserToolbar2.tools"
shared ref as ScriptedToolstrip'object
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

'-->
sub test1()
   trace "-------->>>",  "p.Contents(i).DockHandler.TabText"
   dim mainWin as object
   mainWin =ZZ.IDEHelper.MainWindow
   'dim activeTab         = ZZ.getActiveTab()
   'dim activeTabType     = ZZ.getActiveTabType()
   'dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  dim docPanelRef()
   listDocumentPane(mainWin)
end sub

sub test2()
  msgBox("OK - 2")
end sub

sub test3()
  msgBox("OK - 3")
end sub


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
        if p.Contents(i) Is mainWin.ActiveMdiChild Then isMdiChild=true
        ''tag=DirectCast(p.Contents(i).DockHandler.Form, DockContent).GetPersistString()
        'tbOpenedFiles.IGrid1.Rows(i).Tag = DirectCast(doc.Contents(i).DockHandler.Form, DockContent).GetPersistString()

      ''trace p.DockState,  isMdiChild.toString+"  "+p.Contents(i).DockHandler.TabText
      
      'if dockHandler.isFloat = true  then 
        trace p.DockState,  isMdiChild.toString+"  "+dockHandler.isHidden.toString+"  "+ dockHandler.TabText
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
     


sub onInitialize()
end sub

sub onTerminate(optional intern as string="")
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
end sub
   
'--> 
  
Sub DockPanel1_ActivePaneChanged(ByVal sender As Object, ByVal e As System.EventArgs)
  exit sub
  
  'msgbox("ActivePaneChanged ????")
   sharedEventCounter=sharedEventCounter+1
   dim traceMode as boolean=false
   'traceMode=true
   dim activePane = sender.ActivePane()
   : ZZ.IDEHelper.MainWindow.flpToolbar.bringToFront()
   : ZZ.IDEHelper.MainWindow.flpToolbar.left=0' 200
   : ZZ.IDEHelper.MainWindow.flpToolbar.top=33
   : sender.top=65' 26
   : sender.height=ZZ.IDEHelper.MainWindow.height-70
   : ZZ.IDEHelper.MainWindow.StatusStrip1.visible=false
   : dim tabName as string
   : if not activePane is nothing then
   :   tabName=activePane.CaptionText
   :   if traceMode=true then
   :      trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+tabName
   :   end if 
   :   'printLine 11,  "! <<  ActivePanelChanged",tabName
   :   dim scriptClass as object = zz.scriptClass("es_popUpMonitor")
   :   if not scriptClass is nothing then
   :      scriptClass.add("activeTab(utb2): "+tabName)
   :   end if
   : else
   :   if traceMode=true then
   :     trace "... refCounter",sharedEventCounter.toString + "  "+privateRefCounter.toString + "  "+publicRefCounter.toString + "  "+sharedRefCounter.toString+ "  "+"activePane ...gibts nicht"
   :   end if
   : end if
end sub


  Sub createOpenedTabList(DockPanel1)
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

Sub DockPanel1_DocumentTabActivated(ByVal Tab As Object, ByVal Key As String)
   'msgbox("ActivePaneChanged ????")
  dim activeTab         = ZZ.getActiveTab()
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()
  zz.doevents()
   trace "! DockPanel1_DocumentTabActivated", key
   'printLine 11, "!!!! ActivePanelChanged", ActiveTabFilespec
   'trace "!!!! ActivePanelChanged", name
end sub





   





'-->
Sub AutoStart() '--------------------------------------------
   
   ' zz.traceClear
   onTerminate("-----------------------------INTERN---")
   
   privateRefCounter = privateRefCounter+1
   publicRefCounter = publicRefCounter+1
   sharedRefCounter = sharedRefCounter+1

''msgbox("ini ????")
   dim id, caption, iconUrl, tag, tooltip, flags as string

   dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
   dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
   addHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged

dim i as integer 
  i = 0
  i+=1:topPanel(i)="ToolBar|##|tbdebug"                                           '1
  i+=1:topPanel(i)="ToolBar|##|tbglobsearch"                                      '2
  i+=1:topPanel(i)="ToolBar|##|tbconsole"                                         '3
  'i+=1:topPanel(i)=""
  
  i = 0
  i+=1:rightPanel(i)="Addin|##|siaCommonProtocols.frmTB_fileExplorer"         '1
  i+=1:rightPanel(i)="Addin|##|siaSolution.frmTB_solutionExplorer"            '2
  i+=1:rightPanel(i)="Addin|##|siaScriptSyncMini.frmTB_scriptSync"            '3
  i+=1:rightPanel(i)="Addin|##|siaCommonProtocols.frmTB_ftpExplorer"          '4
  i+=1:rightPanel(i)="Addin|##|siaRTF.frmTB_twAjaxExplorer"                   '5 
  i+=1:rightPanel(i)=""                                                       '6
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|es_fileFinder.eingabe"        '7
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist"  '8
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|es_fileFinder02.eingabe"      '9
  i+=1:rightPanel(i)="// Toolbar|##|tbScriptWin|##|webbrowser1.tab"           '10

  i = 0
  i+=1:popups(i)="ToolBar|##|tbScriptWin|##|bookmark.vb"                   '1
  i+=1:popups(i)="Addin|##|siaCodeCompiler.frmTB_infoTips"                   '2
  i+=1:popups(i)=""                                                          '3
  i+=1:popups(i)="ToolBar|##|tbScriptWin|##|es_popUpMonitor.mainWin"         '4



  Ref = ZZ.IDEHelper.CreateToolbar(ToolbarID1)
  With Ref
    dim el as object
    .ResetControls()
      
    Dim btn=.addIcon ("cmdToggleLeftBar", toggleLB_on) : onRefreshLeftBarToggleBtn(btn)
   .addButton ("cmdRightBar2"       , "projekte"      ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/VCProject_7-16.png") :
    .addButton ("cmdRightBar3"       , "Sync"     ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/CSCDLL_143-16.png") :
    .addButton ("cmdPopups1"         , "snip"           ,IconURL:="http://es.teamwiki.net/docs/icons/arrow_down16.png" ) :
    .addButton ("cmdPopups4"        , "Mon"            ,IconURL:="http://es.teamwiki.net/docs/img/package_settings16.png" ) :
 
    .addLabel  ("", "-")              '-----------------------------------------------
    .AddButton("cmdAddinConf"        , "F"    ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/vsmacros_216-32.png")
    .AddButton ("cmdAddinConf2"      , "A"       , IconURL:="http://es.teamwiki.net/docs/icons/add_milestone.png")
    .AddButton ("cmdopenExplorer"    , "nav"       , IconURL:="http://es.teamwiki.net/docs/icons/folder_horizontal_open.png")
     .addButton ("cmdWeb"            ,"WEB" ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/iexplore_32542-16.png") :
    .addLabel  ("", "-")            '-----------------------------------------------
    .addLabel  ("", "-")            '-----------------------------------------------
     .addButton ("cmdPopups2"        , "ref"            ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/vcpkgui_6087-16.png" ) :
    .addLabel  ("", "-")              '-----------------------------------------------


 'http://es.teamwiki.net/docs/icons/configure_shortcuts.png
 
 
    .addIcon   ("cmdCloseTopBar"    , "http://es.teamwiki.net/docs/icons/16-square-red.png") :
     .addButton ("cmdTopBar1"       , "Debug"    ,"#85A7CB" ,  , , 50,22) :
    .addButton ("cmdTopBar2"        , "Suche"    ,"#D5C851" ,  , , 50,22) :
    .addButton ("cmdTopBar3"        , "Cmd  "  ,"#85A7CB" ,  , , 55,22) :
 
          id = "menuTools02" : caption = "T"
         tag = "C:\yEXE\traceMonitor.exe"  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png"  
     toolTip = ""
       ''flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
           el=.AddButton (id, caption, IconURL:=iconUrl): el.tag=tag

    .addLabel  ("", "-")            '-----------------------------------------------
    '.AddButton ("cmdAddinConf"      ,"+ "        , IconURL:="http://es.teamwiki.net/docs/icons/16-tool-a.png")
    
    .addButton ("cmdEsTest1"        , "M"     ,IconURL:="http://es.teamwiki.net/docs/icons/edit_16x16.gif") :
    .addButton ("cmdEsTest2"        , "clock" ,"" ,  , , 50,22) :
    .addButton ("cmdEsTest3"        , "Color" ,"" ,  , , 50,22) :

    
'-->  ... menü: Tools     
          id = "mnutest1" : caption = "Tools"
         tag = ""  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/syncui_120-16.png"  
     toolTip = ""
       flags = ButtonFlags.StartMenu : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------

    .AddButton ("cmdEditMe"         , "editMe...", IconURL:="http://es.teamwiki.net/docs/icons/edit_16x16.gif")
 
               '------------------------------------------------------------------------------------------------
          id = "menuTools01" : caption = "Bluescreen"
         tag = "C:\Windows\system\ZZ_UGB\BlueScreen.exe"  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag

    .addButton ("cmdEsTest4"        , "Musik  " ,"" ,  , , 50,22) :

               '------------------------------------------------------------------------------------------------
          id = "menuTools03" : caption = "InterProg"
         tag = "c:\yEXE\IprocViewer.exe"  
     iconUrl = "http://es.teamwiki.net/docs/icons/information.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag

               '------------------------------------------------------------------------------------------------
          id = "menuTools04" : caption = "WinSpy"
         tag = "c:\yEXE\WinSpy.exe"  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag

               '------------------------------------------------------------------------------------------------
          id = "menuTools05" : caption = "DragTheClipboard"
         tag = "C:\Windows\system\ZZ_UGB\vbDragTheClipboard.exe"  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "menuTools06" : caption = "ClipboardViewer"
         tag = "C:\Windows\system\ZZ_UGB\vbClipViewer.exe"  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "menuTools07" : caption = "taskList"
         tag = "C:\Windows\system\ZZ_UGB\taskList.exe"  
     iconUrl = ""  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "menuTools08" : caption = "FileInfo"
         tag = "C:\Windows\system\ZZ_UGB\FileVersionInfo.exe"  
     iconUrl = ""  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "conv_vb" : caption = "Zu VB.Net konvertieren"
         tag = ""  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/msvbprj_4510-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "conv_cs" : caption = "Zu C# konvertieren"
         tag = ""  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/csproj_101-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "cmdAutorunMan" : caption = "Autostarts ändern"
         tag = ""  
     iconUrl = ""  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "cmdShowTb2" : caption = "TB2 anzeigen"
         tag = ""  
     iconUrl = "http://mw.teamwiki.net/docs/img/win-icons/wmploc_602-16.png"  
     toolTip = ""
       flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
          id = "" : caption = ""
         tag = ""  
     iconUrl = ""  
     toolTip = ""
      ''flags = "" : el=.AddButton (id, caption, IconURL:=iconUrl, Flags:=flags): el.tag=tag
               '------------------------------------------------------------------------------------------------
    .ActiveMenu = Nothing


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

    
    .addLabel  ("", "-")            '-----------------------------------------------
   
    '.addIcon   ("cmdHideSidebarLeft"  ,"http://es.teamwiki.net/docs/icons/16-square-red.png") :
    '.addLabel  ("", " ")            '-----------------------------------------------
    .addButton ("cmdHideSidebarLeft"  ," "     ,IconURL:="http://es.teamwiki.net/docs/icons/16-square-red.png") :
    .addButton ("cmdRightBar1"        ,"Loc "     ,IconURL:="http://es.teamwiki.net/docs/icons/folder_open.png") :
    .addButton ("cmdRightBar4"        ,"FTP "      ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/Microsoft_883-16.png" ) :
    .addButton ("cmdRightBar5"        ,"rtf "     ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/SHELL32_2-16.png") :
    '.addButton ("test1"        ,"test1"     ,""  ,  , , 50,22) :
   
    .AddButton("mnuCompiled"        , "EXE"                 , IconURL:="http://mw.teamwiki.net/docs/img/win-icons/NETSHELL_1610-16.png", Flags:=ButtonFlags.StartMenu)
     addExeFilesToMenu(ref)
    .ActiveMenu = Nothing
 
    ''.addTextBox("aaaa","bbbb")
    
    '.addButton ("cmdClose"      ,"- X - "   ,"#c44"     ,  , , 44,22) :
    '.addButton ("cmdShowEsFileFinder"      ,"Webbrowser"  ,""     ,  , , 50,22) :
    '.addButton ("cmdShowPrintline"        ,"print"      ,""  ,  , , 50,22) :
    '.br(22)
    '.offsetX = 333
   ' .addIcon ("cmdClose2" ,"http://mw.teamwiki.net/docs/img/win-icons/imageres_170-16.png") :
    '.addButton ("cmdClose2"       ,"- X - "   ,"#c44"    ,  , , 44,22) :
 
    '.AddIcon("cmdCloseTbByName","http://mw.teamwiki.net/docs/img/win-icons/pifmgr_39-32.png")
  End With
End Sub

sub test1_MouseClick(e)
  trace "---------","test1_MouseClick"
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
  trace "", 
  trace "", 
  trace "", 
  'printLine 2, "splitFilespecData", name
  dim scriptClass as object
  scriptClass=zz.scriptClass(name)
  msgBox ("fileAction 333")
  
  msgBox(name)
  ''scriptClass.test1()
 
end sub


sub addExeFilesToMenu(ref as object)
  Dim finder = ZZ.OpenFileFinder("C:\yPara\scriptIDE\compiledScripts\", "*.*")
  Dim RESULT() = Split(finder.FindRecursive(), vbNewLine)
  For Each Line As String In RESULT
    If Line="" Then Continue For
    Dim Parts() = Split(Line, vbTab)
    If Parts(1).ToLower().EndsWith(".exe") Then _
    ref.AddButton("mnuexe_"+Parts(1), Parts(1))
  Next
end sub

'-->



Sub onRefreshLeftBarToggleBtn(btn)
  Dim stat = ZZ.getDockPanelRef("ToolBar|##|tbIndexList").Visible
  btn.Checked = stat
  If stat Then
    btn.Image = ZZ.GetImageCached(toggleLB_off)
  Else
    btn.Image = ZZ.GetImageCached(toggleLB_on)
  End If
End Sub

Sub cmdToggleLeftBar_MouseClick(e)
  setDockpanelVisible("ToolBar|##|tbIndexList", Not e.sender.Checked)
  setDockpanelVisible("ToolBar|##|tbOpenedFiles", Not e.sender.Checked)
  setDockpanelVisible("ToolBar|##|tbScriptWin|##|es_schnellTester.testWin", Not e.sender.Checked)
  setDockpanelVisible("ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_cpuMonitor2.dll|##|sw_cpuMonitor2.sw_cpuMonitor2", Not e.sender.Checked)
  onRefreshLeftBarToggleBtn(e.sender)
  
End Sub



Sub setDockpanelVisible(pnlID As String, stat As Boolean)
  If stat Then ZZ.getDockPanelRef(pnlID).show() Else ZZ.getDockPanelRef(pnlID).hide()
End Sub


'-->
sub cmdWeb_MouseClick(e)
  ZZ.shellExecute("C:\yPara\scriptIDE\compiledScripts\es_webbrowser3\es_webbrowser3.exe")
  exit sub
  
  trace "---------","MouseClick(e)"
  dim id3="ToolBar|##|tbScriptWin|##|es_webbrowser3.tab"
  dim panel2=ZZ.getDockPanelRef(id3)
  'dim activeTab         = ZZ.getActiveTab.CaptionText
  dim activeTabType     = ZZ.getActiveTabType()
  dim ActiveTabFilespec = ZZ.getActiveTabFilespec()

  trace "getActiveTab",  ActiveTabFilespec
  'trace "IsActiveDocumentPane",panel2.IsActiveDocumentPane
  'trace "IsAutoHide()",panel2.IsAutoHide()
  trace "IsFloat",panel2.IsFloat
  trace "IsHidden",panel2.IsHidden
  'trace "IsActiveDocumentPane()",panel2.IsActiveDocumentPane()
  panel2.activate()
  exit sub
  if panel2.visible=true then
     panel2.hide()
  else
     panel2.show()
     panel2.ReloadWidget()
  end if
end sub
sub cmdTopBar1_MouseClick(e)
  trace "---------","MouseClick(e)"
  toggleTopBar(e, 1)
end sub
sub cmdTopBar2_MouseClick(e)
  toggleTopBar(e, 2)
end sub
sub cmdTopBar3_MouseClick(e)
  toggleTopBar(e, 3)
end sub



sub cmdRightBar1_MouseClick(e)
  trace "---------","MouseClick(e)"
  toggleRightBar(e, 1)
end sub
sub cmdRightBar2_MouseClick(e)
  toggleRightBar(e, 2)
end sub
sub cmdRightBar3_MouseClick(e)
  toggleRightBar(e, 3)
end sub
sub cmdRightBar4_MouseClick(e)
  toggleRightBar(e, 4)
end sub
sub cmdRightBar5_MouseClick(e)
  toggleRightBar(e, 5)
end sub

sub cmdPopups1_MouseClick(e)
  ''ZZ.getDockPanelRef(rightPanel(5)).hide()
  ''ref.element("btn_cmdRightBar6").checked=false
  togglePopups(e, 1)
end sub

sub cmdPopups2_MouseClick(e)
  ''ZZ.getDockPanelRef(rightPanel(5)).hide()
  ''ref.element("btn_cmdRightBar6").checked=false
  togglePopups(e, 2)
end sub

sub cmdPopups4_MouseClick(e)
  togglePopups(e, 4)
end sub

'-->


sub cmdopenExplorer_MouseClick(e)
  Dim tab As IDockContentForm = ZZ.IDEHelper.getActiveTab()
  'Dim intUrl = tab.URL
  ' ### Dim fileSpec = ZZ.IDEHelper.ProtocolManager.MapToLocalFilename(tab.URL)
  ' ### Process.start("explorer.exe", "/select," + fileSpec)
end sub

sub cmdCloseTopBar_MouseClick(e)
  resetTopBar()
end sub

sub cmdHideSidebarLeft_MouseClick(e)
  resetRightBar()
  exit sub
 
  'printLine  2, "cmdHideSidebarLeft", e.Sender.Name
  dim id
  id="Addin|##|siaCommonProtocols.frmTB_ftpExplorer"         :  ZZ.getDockPanelRef(id).hide()
  id="Addin|##|siaScriptSyncMini.frmTB_scriptSync"           :  ZZ.getDockPanelRef(id).hide()
  id="Addin|##|siaRTF.frmTB_twAjaxExplorer"                  :  ZZ.getDockPanelRef(id).hide()
  id="Addin|##|siaCommonProtocols.frmTB_fileExplorer"        :  ZZ.getDockPanelRef(id).hide()
  id="Addin|##|siaSolution.frmTB_solutionExplorer"           :  ZZ.getDockPanelRef(id).hide()
  id="ToolBar|##|tbScriptWin|##|es_fileFinder.eingabe"       :  ZZ.getDockPanelRef(id).hide()
  id="ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist" :  ZZ.getDockPanelRef(id).hide()
  id="ToolBar|##|tbScriptWin|##|es_fileFinder02.eingabe"     :  ZZ.getDockPanelRef(id).hide()
  id="Toolbar|##|tbScriptWin|##|webbrowser1.tab"             :  ZZ.getDockPanelRef(id).hide()
end sub

sub toggleTopBar(e as object, index as integer)
  if skipAllEvents=true then exit sub
  dim toolWindow=ZZ.getDockPanelRef(topPanel(index))
  dim curState as boolean  = toolWindow.visible
  resetTopBar()
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    e.sender.checked=true
   else
    toolWindow.hide()
    e.sender.checked=false
  end if
end sub

sub toggleRightBar(e as object, index as integer)
  if skipAllEvents=true then exit sub
  if rightPanel(index)="" then exit sub
  dim toolWindow=ZZ.getDockPanelRef(rightPanel(index))
  dim curState as boolean  = toolWindow.visible
  resetRightBar()
  if curState =false then
    toolWindow.show()
    ''toolWindow.visible =true
    e.sender.checked=true
   else
    toolWindow.hide()
    e.sender.checked=false
  end if
end sub

sub togglePopups(e as object, index as integer)
  if skipAllEvents=true then exit sub
  dim toolWindow=ZZ.getDockPanelRef(popups(index))
  dim curState as boolean  = toolWindow.visible
  ''resetPopups()
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    toolWindow.parent.visible =true
    e.sender.checked=true
   else
    toolWindow.hide()
    toolWindow.visible =false
    e.sender.checked=false
  end if
end sub

sub resetTopBar()
  skipAllEvents=true
  ZZ.getDockPanelRef(topPanel(1)).hide()
  ZZ.getDockPanelRef(topPanel(2)).hide()
  ZZ.getDockPanelRef(topPanel(3)).hide()
  ref.element("btn_cmdTopBar1").checked=false
  ref.element("btn_cmdTopBar2").checked=false
  ref.element("btn_cmdTopBar3").checked=false
  skipAllEvents=false
end sub

sub resetRightBar()
  skipAllEvents=true
  : ZZ.getDockPanelRef(rightPanel(1)).hide()
  : ZZ.getDockPanelRef(rightPanel(2)).hide()
  : ZZ.getDockPanelRef(rightPanel(3)).hide()
  : ZZ.getDockPanelRef(rightPanel(4)).hide()
  : if rightPanel(5) > "" then ZZ.getDockPanelRef(rightPanel(5)).hide()
  : 'ZZ.getDockPanelRef(rightPanel(6)).hide()
  : ref.element("btn_cmdRightBar1").checked=false
  : ref.element("btn_cmdRightBar2").checked=false
  : ref.element("btn_cmdRightBar3").checked=false
  : ref.element("btn_cmdRightBar4").checked=false
  : ref.element("btn_cmdRightBar5").checked=false
  : 'ref.element("btn_cmdRightBar6").checked=false
  err.number=0
  skipAllEvents=false
end sub

sub resetPopups()
  skipAllEvents=true
  ZZ.getDockPanelRef(popups(1)).hide()
  ZZ.getDockPanelRef(popups(2)).hide()
  ref.element("btn_cmdPopups1").checked=false
  ref.element("btn_cmdPopups2").checked=false
  skipAllEvents=false
end sub


'-->

Sub menuTools01_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools02_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools03_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools04_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools05_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools06_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools07_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools08_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub
Sub menuTools09_MouseClick(e)
  printLine 1, e.sender.name, e.sender.tag
  ZZ.shellExecute(e.sender.tag)
end sub

sub cmdEsTest4_MouseClick(e)
  dim id3="ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_mediaPlayere.dll|##|DOM_Player_1._0.sw_mediaPlayer"
  dim panel2=ZZ.getDockPanelRef(id3)
  if panel2.visible=true then
     panel2.hide()
  else
     panel2.show()
     panel2.ReloadWidget()
  end if
end sub
sub cmdEsTest3_MouseClick(e)
  dim id3="ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sw_colorPicker.dll|##|sw_colorPicker.sg_colorPicker"
  dim panel2=ZZ.getDockPanelRef(id3)
  if panel2.visible=true then
     panel2.hide()
  else
     panel2.show()
     panel2.ReloadWidget()
  end if
end sub


sub cmdEsTest2_MouseClick(e)
  dim id3="ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock"
  dim panel2=ZZ.getDockPanelRef(id3)
  if panel2.visible=true then
     panel2.hide()
  else
     panel2.show()
     panel2.ReloadWidget()
  end if
end sub


sub cmdEsTest1_MouseClick(e)
  dim id3="ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\sg_memo.dll|##|root_sg_memo.sg_memo"
  dim panel2=ZZ.getDockPanelRef(id3)
  if panel2.visible=true then
     panel2.hide()
  else
     panel2.show()
     panel2.ReloadWidget()
  end if
  exit sub


  'msgBox("xxxx")
  e.sender.checked=true
  dim id="ToolBar|##|tbindexlist"
  dim indexList=ZZ.getDockPanelRef(id)
  printLine 3, "indexList", indexList.visible
  if indexList.visible =true then
    'indexList.hide()
  else
    'indexList.show()
    'indexList.visible =true
  end if
  
  '' indexList.parent.parent.parent.DockLeftPortion=0.22
  
  dim id2="ToolBar|##|tbLegacyWidget|##|C:\yPara\mwSidebar\widgets\mwTimer.dll|##|mwTimer.sw_analogClock"
  dim panel=ZZ.getDockPanelRef(id2)
  if panel.parent.parent.height < 333 then
    panel.parent.parent.top = 133
    panel.parent.parent.height = 1000
    panel.parent.parent.width = 800
  else
    panel.parent.parent.top = 22
    panel.parent.parent.height = 22
    panel.parent.parent.width = 111
  end if
  panel.text="Beim Gongschlag ist es genau ..."
  panel.parent.text="11111"
  panel.parent.parent.text="expand..."
  ''panel.parent.parent.parent.text="33333"  
  '' panel.parent.parent.parent.parent.text="44444444"  

End Sub



Sub CreateTB2()
  Dim Ref = ZZ.IDEHelper.CreateToolbar("mw_UserToolbar1.tb2")
  Ref.ResetControls()
  Ref.AddButton("test","Test3")
  Ref.AddButton("test2","Test4")
  Ref.AddIcon("cmdCloseTb2","http://mw.teamwiki.net/docs/img/win-icons/wiashext_131-32.png")
End Sub



'-->
'-->


Sub mnuCompiled_ItemClicked(e)
  Dim fileSpec = "C:\yPara\scriptIDE\compiledScripts\"+IO.Path.GetFilenameWithoutExtension(e.Sender.Text)+"\"+e.Sender.Text
  printLine 2,"fileSpec",fileSpec
  Process.Start(fileSpec)
End Sub

Sub cmdEditMe_MouseClick(e)
  ZZ.IDEHelper.NavigateFile("loc:/C:/yPara/scriptIDE/scriptClass/es_UserToolbar2.vb")
End Sub

Sub cmdCloseTb2_MouseClick(e)
  ZZ.IDEHelper.RemoveToolbar("mw_UserToolbar1.tb2")
End Sub

Sub cmdShowTb2_MouseClick(e)
  CreateTB2
End Sub

Sub cmdCloseTbByName_MouseClick(e)
  Dim x=Inputbox("Name der Toolbar")
  If x<>"" Then _
    ZZ.IDEHelper.RemoveToolbar(x)
  
End Sub





'-->

Sub cmdAutorunMan_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog(2) : Err.Clear
End Sub


'' Treenode "xxxxxx" nicht gefunden. Verfügbare Knoten:
'' home
'' allgemein
''    fensterverwaltung
''    winswitcher
''    toolbars
''    siaskinnable.ctl_options
'' erweiterungen
''    addins
''    scriptklassen
'' protokolle
''    siacommonprotocols.ctl_options
'' tools
''    siasolution.ctl_options
'' 






Sub cmdAddinConf_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog("fensterverwaltung") : Err.Clear
End Sub

Sub cmdAddinConf2_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog("addins") : Err.Clear
End Sub

sub cmdShowPrintline_MouseClick(e)
    dim id: cmdHideSidebarLeft_MouseClick(e)
    ''id="ToolBar|##|tbScriptWin|##|webbrowser1.tab" : ZZ.getDockPanelRef(id).show()
end sub


sub cmdShowEsFileFinder_MouseClick(e)
    dim id: cmdHideSidebarLeft_MouseClick(e)
    id="ToolBar|##|tbScriptWin|##|es_webbrowser3.tab" : ZZ.getDockPanelRef(id).show()
end sub

'-->


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
'' Sub onButtonEvent(e)
''   'trace "aaa"
''   
''   Dim btnName = e.Sender.Name.Substring(4)
''   printLine 2,"buttonEvent",btnName
''   
'' : CallByName(Me, btnName, Microsoft.VisualBasic.CallType.Method, e)
''   If Err.Number Then MsgBox("Funktion nicht gefunden"+vbnewline+"..."+btnName):Err.Clear
''   
''   
''   
'' End Sub







