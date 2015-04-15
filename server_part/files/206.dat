

#Para DebugMode Internal

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

Function keyMan_112(keyStr As String) As Boolean  'F1
  'Msgbox("Taste 112: "+keyStr,,"KeyHandler")
  toggleDebugWindow()
  Return True '=Cancel
End Function


Function keyMan_114(keyStr As String) As Boolean 'F3
  toggleMonitorWindow()
End Function 

Function keyMan_115(keyStr As String) As Boolean 'F4
  toggleDockPanel(DockLeft)
  Return True '=Cancel
End Function 


Function toggleMonitorWindow()
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbScriptWin|##|es_popUpMonitor.mainWin" )
  dim curState as boolean  = toolWindow.visible
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
   else
    toolWindow.hide()
  end if
  ZZ.IDEHelper.getMainFormRef.focus()
End Function 


Function toggleDebugWindow()
  dim toolWindow=ZZ.getDockPanelRef("ToolBar|##|tbdebug" )
  dim curState as boolean  = toolWindow.visible
  if curState =false then
    toolWindow.show()
    toolWindow.visible =true
   else
    toolWindow.hide()
  end if
  ZZ.IDEHelper.getMainFormRef.focus()
  Return True '=Cancel
End Function


:Function toggleDockPanel(intPanel)
  static state as boolean = true: 'state = not state
  dim mainWin as object = ZZ.IDEHelper.MainWindow
  dim Panel as object = mainWin.DockPanel1
  dim p as object: dim i as integer
  For Each p In Panel.Panes
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




:Function toggleFullScreen()
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



