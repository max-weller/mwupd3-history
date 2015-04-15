

#Para DebugMode Internal
#Para SilentMode true

Public Const DockTop = 7
Public Const DockLeft = 8
Public Const DockBottom =9 
Public Const DockRight = 10
  

Function keyMan_13(keyStr As String) As Boolean
  'Msgbox("Taste 13: "+keyStr,,"KeyHandler")
  if keyStr = "SHIFT-RETURN" then
    'Msgbox("Taste 13: "+keyStr,,"KeyHandler")
    zz.scriptClass("es_popUpMonitor").show()
    Return True '=Cancel
end if  
  'Return True '=Cancel
End Function



Function keyMan_16(keyStr As String) As Boolean
  'Msgbox("Taste 13: "+keyStr,,"KeyHandler")
  if keyStr = "CTRL-ALT-SHIFT-SHIFTKEY" then
    'Msgbox("Taste 13: CTRL-ALT-SHIFT-SHIFTKEY", ,"KeyHandler")
    toggleFullScreen()
    'Return True '=Cancel
  end if  
  'Return True '=Cancel
End Function

Function keyMan_27(keyStr As String) As Boolean  'ESCAPE
  'Msgbox("Taste 112: "+keyStr,,"KeyHandler")
  toggleDockPanel(DockLeft)
  'toggleFullScreen()
  Return True '=Cancel
End Function

Function keyMan_112(keyStr As String) As Boolean  'F1
  'Msgbox("Taste 112: "+keyStr,,"KeyHandler")
  toggleDebugWindow()
  'toggleDockPanel(DockTop)
  'toggleFullScreen()
  Return True '=Cancel
End Function

Function keyMan_114(keyStr As String) As Boolean 'F3
  'toggleDockPanel(DockTop)
  toggleMonitorWindow()
End Function 

Function keyMan_115(keyStr As String) As Boolean 'F4
  ''toggleDockPanel(DockLeft)
  dim id as string="ToolBar|##|tbScriptWin|##|bookmark.vb"
  
  togglePopups(id)
  Return True '=Cancel
End Function 


Function keyMan_117(keyStr As String) As Boolean 'F6
  toggleDebugWindow()
  Return True '=Cancel
End Function 

Function keyMan_118(keyStr As String) As Boolean 'F7
  toggleMonitorWindow()
  Return True '=Cancel
End Function 

Function keyMan_122(keyStr As String) As Boolean 'F11
  'toggleFullScreen()
  'Return True '=Cancel
End Function 

Function keyMan_123(keyStr As String) As Boolean 'F12
  toggleFullScreen()
  'toggleMonitorWindow()
  'Return True '=Cancel
End Function 



sub togglePopups(id)
  dim toolWindow=ZZ.getDockPanelRef(id)
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


'-->
Function toggleMonitorWindow()
  togglePanel(ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_popUpMonitor.mainWin"))
End Function 


Function toggleDebugWindow()
  togglePanel(ZZ.getDockPanelRef("ToolBar|##|tbdebug"))
End Function



Function toggleDockPanel(intPanel)
  static state as boolean = true: 'state = not state
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  ' mainWin.xxxxxxxxxx    '''FEHLER-TEST
  dim Panel as object = mainWin.DockPanel1
    '' trace Panel.Panes.count, "--xxx--"
    '' trace Panel.Panes(0).isHidden, "--0--"
    '' trace Panel.Panes(1).isHidden, "--1--"
    '' trace Panel.Panes(2).isHidden, "--2--"
    '' trace Panel.Panes(3).isHidden, "--3--"
  dim p as object: dim i as integer
  For Each p In Panel.Panes
    '' trace "Panel.Panes...1", p.captionText
    '' trace "Panel.Panes...2", p.isHidden
    '' trace "Panel.Panes...3", p.isActivated
     For i = 0 To p.Contents.Count - 1
       if p.DockState = intPanel  then 
         p.parent.visible = not p.parent.visible
         exit Function
         '' if state =true then
         ''   'p.show()
         ''   p.parent.visible=true
         ''   'p.visible=true
         ''   p.parent.show()
         '' else
         ''   'p.hide
         ''   p.parent.hide()
         '' end if
       end if
     Next
  Next
End Function


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



