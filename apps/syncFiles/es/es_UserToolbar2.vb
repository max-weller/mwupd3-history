


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




sub onInitialize()
 end sub


sub onTerminate()
   trace "es ??? onTerminate","es ??? onTerminate"
   dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
   dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
   removeHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged
  'ZZ.IDEHelper.RemoveToolbar("toolbar.tb1")
end sub


Sub DockPanel1_ActivePaneChanged(ByVal sender As Object, ByVal e As System.EventArgs)
   'msgbox("ActivePaneChanged ????")
   trace "!!!! ActivePanelChanged", "!!! ??????? H A L L O ???????"
end sub


'-->
Sub AutoStart() '--------------------------------------------

  ''msgbox("ini ????")
   dim dockPanel as WeifenLuo.WinFormsUI.Docking.DockPanel
   dockPanel = cType(zz.ideHelper.mainwindow.DockPanel1, WeifenLuo.WinFormsUI.Docking.DockPanel)
   addHandler dockPanel.ActivePaneChanged, AddressOf  DockPanel1_ActivePaneChanged


dim i as integer 
  i = 0
  i+=1:topPanel(i)="ToolBar|##|tbdebug"               '0
  i+=1:topPanel(i)="ToolBar|##|tbglobsearch"          '1
  i+=1:topPanel(i)="ToolBar|##|tbconsole"             '2
  'i+=1:topPanel(i)=""
  
  i = 0
  i+=1:rightPanel(i)="Addin|##|siaCommonProtocols.frmTB_fileExplorer"         '1
  i+=1:rightPanel(i)="Addin|##|siaSolution.frmTB_solutionExplorer"            '2
  i+=1:rightPanel(i)="Addin|##|siaScriptSyncMini.frmTB_scriptSync"            '3
  i+=1:rightPanel(i)="Addin|##|siaCommonProtocols.frmTB_ftpExplorer"          '4
  i+=1:rightPanel(i)="Addin|##|siaRTF.frmTB_twAjaxExplorer"                   '5
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|es_fileFinder.eingabe"        '6
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist"  '7
  i+=1:rightPanel(i)="ToolBar|##|tbScriptWin|##|es_fileFinder02.eingabe"      '8
  i+=1:rightPanel(i)="// Toolbar|##|tbScriptWin|##|webbrowser1.tab"           '9
 
  Ref = ZZ.IDEHelper.CreateToolbar(ToolbarID1)
  With Ref
    .ResetControls()
      
    .addButton ("cmdTraceMonitor"   ," " ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png")

    
    .AddButton ("cmdAddinConf"      ," ", IconURL:="http://es.teamwiki.net/docs/icons/16-tool-a.png")
    .addLabel  ("", "-")            '-----------------------------------------------
    .AddButton ("cmdopenExplorer"   , "", IconURL:="http://es.teamwiki.net/docs/icons/folder_horizontal_open.png")
    .addLabel  ("", " ")            '-----------------------------------------------
    
    .addIcon   ("cmdCloseTopBar"    ,"http://es.teamwiki.net/docs/icons/16-square-red.png") :
    .addLabel  ("", "-")            '-----------------------------------------------
    .addButton ("cmdTopBar1"        ,"Debug"    ,"#85A7CB" ,  , , 50,22) :
    .addButton ("cmdTopBar2"        ,"Suche"    ,"#D5C851" ,  , , 50,22) :
    .addButton ("cmdTopBar3"        ,"Cmd  "  ,"#85A7CB" ,  , , 55,22) :
    .addLabel  ("", "-")            '-----------------------------------------------
    
    .AddButton ("cmdEditMe"         , "", IconURL:="http://es.teamwiki.net/docs/icons/edit_16x16.gif")
    .addButton ("cmdEsTest1"        , " M E M O "     ,"" ,  , , 50,22) :
    .addButton ("cmdEsTest2"        , "Clock" ,"" ,  , , 50,22) :
    .addButton ("cmdEsTest3"        , "Color" ,"" ,  , , 50,22) :
    .addButton ("cmdEsTest4"        , "Musik" ,"" ,  , , 50,22) :
    .addLabel  ("", "-")            '-----------------------------------------------
    
    .AddButton ("mnutest1"          , "Menü..."             , IconURL:="http://mw.teamwiki.net/docs/img/win-icons/syncui_120-16.png", Flags:=ButtonFlags.StartMenu)
    .AddButton ("test2"             , "Test2"               ,,,,,,"http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png")
    .AddButton ("conv_vb"           , "Zu VB.Net konvertieren",,,,,,"http://mw.teamwiki.net/docs/img/win-icons/msvbprj_4510-16.png")
    .AddButton ("conv_cs"           , "Zu C# konvertieren"  ,,,,,,"http://mw.teamwiki.net/docs/img/win-icons/csproj_101-16.png")
    .AddButton ("cmdAutorunMan"     , "Autostarts ändern")
    .AddButton ("cmdShowTb2"        , "TB2 anzeigen"        , IconURL:="http://mw.teamwiki.net/docs/img/win-icons/wmploc_602-16.png")
    .ActiveMenu = Nothing
    
    .AddButton("mnuCompiled"        , "EXE"                 , IconURL:="http://mw.teamwiki.net/docs/img/win-icons/syncui_120-16.png", Flags:=ButtonFlags.StartMenu)
     addExeFilesToMenu(ref)
    .ActiveMenu = Nothing
    
    .addLabel  ("", "-")              '-----------------------------------------------
    .addLabel  ("", "-")              '-----------------------------------------------
    .addLabel  ("", "-")              '-----------------------------------------------
    .addIcon   ("cmdHideSidebarLeft"  ,"http://es.teamwiki.net/docs/icons/16-square-red.png") :
    .addLabel  ("", "-")              '-----------------------------------------------
    .addButton ("cmdRightBar1"        ,"Loc"     ,IconURL:="http://es.teamwiki.net/docs/icons/folder_open.png") :
    .addLabel  ("", "-")              '-----------------------------------------------
    .addButton ("cmdRightBar2"        ,"Pro"     ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/VCProject_7-16.png") :
    .addLabel  ("", "-")              '-----------------------------------------------
    .addButton ("cmdRightBar3"        ,"Sync"     ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/CSCDLL_143-16.png") :
    .addLabel  ("", "-")              '-----------------------------------------------
    .addButton ("cmdRightBar4"        ,"FTP"      ,""  ,  , , 50,22) :
    .addLabel  ("", "-")              '-----------------------------------------------
    .addButton ("cmdRightBar5"        ,"rtf"     ,""  ,  , , 50,22) :
    
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
sub cmdTopBar1_MouseClick(e)
  toggleTopBar(e, 1)
end sub
sub cmdTopBar2_MouseClick(e)
  toggleTopBar(e, 2)
end sub
sub cmdTopBar3_MouseClick(e)
  toggleTopBar(e, 3)
end sub

sub cmdRightBar1_MouseClick(e)
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

'-->

sub cmdopenExplorer_MouseClick(e)
  Dim tab As IDockContentForm = ZZ.IDEHelper.getActiveTab()
  'Dim intUrl = tab.URL
  Dim fileSpec = ZZ.IDEHelper.ProtocolManager.MapToLocalFilename(tab.URL)
  Process.start("explorer.exe", "/select," + fileSpec)
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
    'toolWindow.hide()
    'e.sender.checked=false
  end if
end sub

sub toggleRightBar(e as object, index as integer)
  if skipAllEvents=true then exit sub
  dim toolWindow=ZZ.getDockPanelRef(rightPanel(index))
  dim curState as boolean  = toolWindow.visible
  resetRightBar()
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    e.sender.checked=true
   else
    'toolWindow.hide()
    'e.sender.checked=false
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
  ZZ.getDockPanelRef(rightPanel(1)).hide()
  ZZ.getDockPanelRef(rightPanel(2)).hide()
  ZZ.getDockPanelRef(rightPanel(3)).hide()
  ZZ.getDockPanelRef(rightPanel(4)).hide()
  ZZ.getDockPanelRef(rightPanel(5)).hide()
  ref.element("btn_cmdRightBar1").checked=false
  ref.element("btn_cmdRightBar2").checked=false
  ref.element("btn_cmdRightBar3").checked=false
  ref.element("btn_cmdRightBar4").checked=false
  ref.element("btn_cmdRightBar5").checked=false
  skipAllEvents=false
end sub


'-->

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

Sub cmdTraceMonitor_MouseClick(e)
  ZZ.shellExecute("C:\yEXE\traceMonitor.exe")
  
End Sub





'-->

Sub cmdAutorunMan_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog(2) : Err.Clear
End Sub
Sub cmdAddinConf_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog() : Err.Clear
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







