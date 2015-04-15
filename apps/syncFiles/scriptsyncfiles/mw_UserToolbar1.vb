
#Para SilentMode True
#Para DebugMode internal
#Imports System.Diagnostics
#Imports System.Collections.Generic

Public Const ToolbarID1 = "mw_UserToolbar1.tools"
Public Const ToolbarID2 = "mw_UserToolbar1.tools2"
Const toggleLB_off = "http://www.iconfinder.net/data/icons/iconic/raster/16/arrow_right.png"
Const toggleLB_on = "http://www.iconfinder.net/data/icons/iconic/raster/16/arrow_left.png"
shared topPanel As New List(Of String)
shared rightPanel As New List(Of String)
Dim Ref As ScriptedToolstrip

Sub AutoStart()
  
  topPanel.Add("ToolBar|##|tbdebug")               '0
  topPanel.Add("ToolBar|##|tbglobsearch")          '1
  topPanel.Add("ToolBar|##|tbconsole")             '2
  
  rightPanel.Add("Addin|##|siaSolution.frmTB_solutionExplorer")            '0
  rightPanel.Add("Addin|##|siaCommonProtocols.frmTB_fileExplorer")         '1
  rightPanel.Add("Addin|##|siaCommonProtocols.frmTB_ftpExplorer")          '2
  rightPanel.Add("Addin|##|siaScriptSyncMini.frmTB_scriptSync")            '3
  rightPanel.Add("Toolbar|##|tbScriptWin|##|webbrowser1.tab")              '4
  rightPanel.Add("Addin|##|siaCodeCompiler.frmTB_infoTips")                '5
  rightPanel.Add("Addin|##|siaRTF.frmTB_twAjaxExplorer")                   '6
  rightPanel.Add("ToolBar|##|tbScriptWin|##|bookmark.vb")                  '7
  ' rightPanel.Add("ToolBar|##|tbScriptWin|##|es_fileFinder.eingabe")      '8
  ' rightPanel.Add("ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist")'9
  ' rightPanel.Add("ToolBar|##|tbScriptWin|##|es_fileFinder02.eingabe")    '10
  
  Dim Ref2 = ZZ.IDEHelper.CreateToolbar(ToolbarID2)
  Ref = ZZ.IDEHelper.CreateToolbar(ToolbarID1)
  With Ref2
    .ResetControls()
    
    Dim btn=.addIcon ("cmdToggleLeftBar", toggleLB_on) : onRefreshLeftBarToggleBtn(btn)
  End With
  With Ref
    .ResetControls()
    
    
    '' .AddButton("mnutest1", "Menü...", IconURL:="http://mw.teamwiki.net/docs/img/win-icons/syncui_120-16.png", Flags:=ButtonFlags.StartMenu)
    '' .AddButton("test2", "Test2",,,,,,"http://mw.teamwiki.net/docs/img/win-icons/WININET_137-16.png")
    '' .AddButton("conv_vb", "Zu VB.Net konvertieren",,,,,,"http://mw.teamwiki.net/docs/img/win-icons/msvbprj_4510-16.png")
    '' .AddButton("conv_cs", "Zu C# konvertieren",,,,,,"http://mw.teamwiki.net/docs/img/win-icons/csproj_101-16.png")
    '' .AddButton("cmdAutorunMan","Autostarts ändern")
    '' .AddButton("cmdShowTb2", "TB2 anzeigen", IconURL:="http://mw.teamwiki.net/docs/img/win-icons/wmploc_602-16.png")
    '' .ActiveMenu = Nothing
    
    .AddButton("cmdEditMe", "",  IconURL:="http://es.teamwiki.net/docs/icons/edit_16x16.gif")
    .addLabel  ("", "-") '-----------------------------------------------
    .addButton ("cmdTraceMonitor", "traceMon" ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/agentsvr_113-16.png")
    
    .AddButton("mnuCompiled", "EXE", IconURL:="http://mw.teamwiki.net/docs/img/win-icons/NETSHELL_1610-16.png", Flags:=ButtonFlags.StartMenu)
    Dim finder = ZZ.OpenFileFinder("C:\yPara\scriptIDE\compiledScripts\", "*.*")
    Dim RESULT() = Split(finder.FindRecursive(), vbNewLine)
    For Each Line As String In RESULT
      If Line="" Then Continue For
      Dim Parts() = Split(Line, vbTab)
      If Parts(1).ToLower().EndsWith(".exe") Then _
        .AddButton("mnuexe_"+Parts(1), Parts(1))
    Next
    .ActiveMenu = Nothing
    
    
    .AddButton("cmdAddinConf","" ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/vsmacros_216-32.png")
    .AddButton("cmdSkins","" ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/themeui_701-16.png")
    
    .addLabel  ("", "  ") '-----------------------------------------------
    .addIcon ("cmdResetBar_TopBar" ,"http://mw.teamwiki.net/docs/img/win-icons/imageres_170-16.png") :
    .addButton ("cmdToggle_TopBar_1"     ,"Suche"    ,"#D5C851" , , , 50,22) :
    .addButton ("cmdToggle_TopBar_0"     ,"Debug"    ,"#85A7CB" , , , 50,22) :
    .addButton ("cmdToggle_TopBar_2"     ,"Konsole"  ,"#85A7CB" , , , 55,22) :
    .addLabel  ("", "  ") '-----------------------------------------------
    '.br(22)
    .addIcon ("cmdResetBar_RightBar" ,"http://mw.teamwiki.net/docs/img/win-icons/imageres_170-16.png") :
    '.addButton ("cmdClose"      ,"- X - "   ,"#c44"     ,  , , 44,22) :
    .addButton ("cmdToggle_RightBar_0"    ,"Project"    ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/VCProject_7-16.png") :
    .addButton ("cmdToggle_RightBar_1"    ,"Locale"     ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/explorer_252-16.png") :
    .addButton ("cmdToggle_RightBar_2"    ,"FTP"        ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/Microsoft_883-16.png") :
    .addButton ("cmdToggle_RightBar_3"    ,"Sync"       ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/imageres_1_175-16.png") :
    .addButton ("cmdToggle_RightBar_4"    ,"Webbrowser" ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/iexplore_32542-16.png") :
    .addButton ("cmdToggle_RightBar_5"    ,"reflect"    ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/vcpkgui_6087-16.png") :
    .addButton ("cmdToggle_RightBar_6"    ,"rtf"        ,IconURL:="http://mw.teamwiki.net/docs/img/win-icons/SHELL32_2-16.png") :
    .addButton ("cmdToggle_RightBar_7"    ,"snip"       ,IconURL:="http://es.teamwiki.net/docs/icons/arrow_down16.png") :
    '.addButton ("cmdShowPrintline"        ,"print"      ,""  ,  , , 50,22) :
    '' .addLabel  ("", "-") '-----------------------------------------------
    '' .br(22)
    '' .offsetX = 333
    '' .addIcon ("cmdClose2" ,"http://mw.teamwiki.net/docs/img/win-icons/imageres_170-16.png") :
    '' '.addButton ("cmdClose2"      ,"- X - "   ,"#c44"    ,  , , 44,22) :
    '' .addButton ("estest2"        , "ugb"     ,"" ,  , , 50,22) :
    '' .addButton ("estest1"        , "sidebar" ,"" ,  , , 50,22) :
    '' .addLabel  ("", "-") '-----------------------------------------------
    '' .addLabel  ("", "-") '-----------------------------------------------
    .addLabel  ("", "   ") '-----------------------------------------------
    .AddIcon("cmdCloseTbByName","http://mw.teamwiki.net/docs/img/win-icons/pifmgr_39-32.png")
  End With
End Sub

Sub CreateTB2()
  
  Dim Ref = ZZ.IDEHelper.CreateToolbar("mw_UserToolbar1.tb2")
  
  Ref.ResetControls()
  Ref.AddButton("test","Test3")
  Ref.AddButton("test2","Test4")
  Ref.AddIcon("cmdCloseTb2","http://mw.teamwiki.net/docs/img/win-icons/wiashext_131-32.png")
  
End Sub


Sub mnuCompiled_ItemClicked(e)
  Dim fileSpec = "C:\yPara\scriptIDE\compiledScripts\"+IO.Path.GetFilenameWithoutExtension(e.Sender.Text)+"\"+e.Sender.Text
  printLine 2,"fileSpec",fileSpec
  Process.Start(fileSpec)
End Sub

Sub cmdEditMe_MouseClick(e)
  ZZ.IDEHelper.NavigateFile("loc:/C:/yPara/scriptIDE/scriptClass/mw_UserToolbar1.vb")
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
Function getBarList(bar As String) As List(Of String)
  If bar = "TopBar" Then Return topPanel
  If bar = "RightBar" Then Return rightPanel
End Function

sub toggleBar(bar As String, sender as object, index as integer, Optional hideOthers As Boolean = True)
  'if skipAllEvents=true then exit sub
  Dim barList = getBarList(bar)
  dim toolWindow=ZZ.getDockPanelRef(barList(index))
  dim curState as boolean  = toolWindow.visible
  If hideOthers Then resetBar(bar)
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
    sender.Checked=true
  elseIf Not hideOthers Then
    toolWindow.hide()
    sender.checked=false
  end if
end sub

sub resetBar(bar As String)
  Dim barList As List(Of String) = getBarList(bar)
  'skipAllEvents=true
  For i As Integer = 0 To barList.Count - 1
    ZZ.getDockPanelRef(barList(i)).hide()
    Dim el = ref.element("btn_cmdToggle_" & bar & "_" & (i))
    If el IsNot Nothing Then el.checked=false
  Next
  'skipAllEvents=false
end sub


Sub onButtonEvent(e)
  Dim name() = Split(e.Sender.Name, "_")
  
  Select Case name(1)
    Case "cmdToggle"   : toggleBar(name(2), e.Sender, name(3), e.MouseButton = "Left")
    Case "cmdResetBar" : resetBar(name(2))
  End Select
End Sub

'-->

Sub cmdAddinConf_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog("addins") : Err.Clear
End Sub
Sub cmdSkins_MouseClick(e)
  'komische Fehlermeldung
  '5 - Im Modul wurde ein Assemblymanifest erwartet. (Ausnahme von HRESULT: 0x80131018)
:  ZZ.IDEHelper.ShowOptionsDialog("siaskinnable.ctl_options") : Err.Clear
End Sub

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
  onRefreshLeftBarToggleBtn(e.sender)
End Sub

Sub setDockpanelVisible(pnlID As String, stat As Boolean)
  If stat Then ZZ.getDockPanelRef(pnlID).show() Else ZZ.getDockPanelRef(pnlID).hide()
End Sub

'' '-->
'' sub cmdCloseTopBar_MouseClick(e)
''   resetBar("TopBar", topBar)
''     '' dim id
''     '' id="ToolBar|##|tbglobsearch"       :  ZZ.getDockPanelRef(id).hide()
''     '' id="ToolBar|##|tbdebug"            :  ZZ.getDockPanelRef(id).hide()
''     '' id="ToolBar|##|tbconsole"          :  ZZ.getDockPanelRef(id).hide()
'' end sub
'' 
'' sub cmdShowSearch_MouseClick(e) 
''     dim id :cmdCloseTopBar_MouseClick(e)
''     id="ToolBar|##|tbglobsearch"       :  ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowDebug_MouseClick(e)
''     dim id :cmdCloseTopBar_MouseClick(e)
''     id="ToolBar|##|tbdebug"            :  ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowKonsole_MouseClick(e)
''     dim id :cmdCloseTopBar_MouseClick(e)
''     id="ToolBar|##|tbconsole"          :  ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' '-->
'' sub cmdHideSidebarLeft_MouseClick(e)
''     'printLine  2, "cmdHideSidebarLeft", e.Sender.Name
''     dim id
''     id="Addin|##|siaCommonProtocols.frmTB_ftpExplorer"         :  ZZ.getDockPanelRef(id).hide()
''     id="Addin|##|siaScriptSyncMini.frmTB_scriptSync"           :  ZZ.getDockPanelRef(id).hide()
''     id="Addin|##|siaRTF.frmTB_twAjaxExplorer"                  :  ZZ.getDockPanelRef(id).hide()
''     id="Addin|##|siaCommonProtocols.frmTB_fileExplorer"        :  ZZ.getDockPanelRef(id).hide()
''     id="Addin|##|siaSolution.frmTB_solutionExplorer"           :  ZZ.getDockPanelRef(id).hide()
''     id="ToolBar|##|tbScriptWin|##|es_fileFinder.eingabe"       :  ZZ.getDockPanelRef(id).hide()
''     id="ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist" :  ZZ.getDockPanelRef(id).hide()
''     id="ToolBar|##|tbScriptWin|##|es_fileFinder02.eingabe"     :  ZZ.getDockPanelRef(id).hide()
''     id="Toolbar|##|tbScriptWin|##|webbrowser1.tab"             :  ZZ.getDockPanelRef(id).hide()
'' end sub
'' 
'' sub cmdShowFtpExplorer_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="Addin|##|siaCommonProtocols.frmTB_ftpExplorer"         :  ZZ.getDockPanelRef(id).show()
''     id="ToolBar|##|tbScriptWin|##|mw_serverBackupHttp.dirlist" :  ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowSolutions_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="Addin|##|siaSolution.frmTB_solutionExplorer" : ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowPrintline_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="ToolBar|##|tbScriptWin|##|webbrowser1.tab" : ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowRTFFileList_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="Addin|##|siaRTF.frmTB_twAjaxExplorer" :   ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowEsFileFinder_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="ToolBar|##|tbScriptWin|##|webbrowser1.tab" : ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowScriptSync_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="Addin|##|siaScriptSyncMini.frmTB_scriptSync" : ZZ.getDockPanelRef(id).show()
'' end sub
'' 
'' sub cmdShowLocaleFolder_MouseClick(e)
''     dim id: cmdHideSidebarLeft_MouseClick(e)
''     id="Addin|##|siaCommonProtocols.frmTB_fileExplorer" 
''     dim panel: panel=ZZ.getDockPanelRef(id)
''    ' panel.width=200
''    ' panel.parent.width=200
''     panel.show()
'' end sub

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

Sub onTerminate()
  'ZZ.IDEHelper.RemoveToolbar("toolbar.tb1")
  
  
  
End Sub

