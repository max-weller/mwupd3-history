#Para DebugMode internal


Sub syntaxHighlight(QQ)
  QQ.scRef.ConfigurationManager.Language = "vbscript"
  QQ.setSCStyle (5, "#000000", "#D4E4F0")  'preproc 
  
End Sub

Function createIntelliPanel(QQ)
  QQ.resetPanel
  QQ.addLabel ("VB.NET  ", "#FFF600", "12")
  QQ.addButton ("save", "Save", "#AAAAAA")
  'QQ.addButton ("test", "test", "#AAAAAA")
  
  'sonderbehandlung durch onIntelliButtonMouseUp:
  QQ.addButton ("scriptclass", "Compile", "#AAAAAA")
  QQ.addButton ("autostart", "AutoStart", "#AAAAAA")
  
  createIntelliPanel = True
End Function

Sub onKeyDown(keyString)
  Select Case keyString
    Case "ALT-F5"
      'onBtnClick_autostart ""
  End Select
End Sub

Sub onBtnClick_save(button)
  'On Error Resume Next
  ZZ.getActiveTab().onSave()
End Sub
Sub onBtnClick_test(button)
  'On Error Resume Next
  msgbox ("test'111")
  
  msgbox ("test'222") 'test "strip" comment
  
  'ZZ.getActiveRTF().onSave()
End Sub

'' Sub onBtnClick_scriptclass(button)
''   'wird nicht mehr aufgerufen, sonderbehandlung durch onIntelliButtonMouseUp
''   stop
''   ZZ.getActiveRTF().onSave()
''   getActScriptClass()
'' End Sub

'' Function getActScriptClass()
''   'On Error Resume Next
''   
''   Dim filename As String
''   filename = ZZ.getActiveTabFilespec()
''   If filename = "" Then Exit Function
''   Dim path,file,ext
''   ZZ.splitFilespecData(filename,path,file,ext)
''   
''   Set getActScriptClass = ZZ.scriptClass(file)
'' End Function
'' 
'' Sub onBtnClick_autostart(button)
''   'wird nicht mehr aufgerufen, sonderbehandlung durch onIntelliButtonMouseUp
''   stop
''   
''   Dim cls
''   Set cls = getActScriptClass()
''   If cls is nothing then exit sub
''   :cls.autoStart
''   :If Err.Number <>0 Then Msgbox ("Fehler beim aufrufen von Autostart: "+Err.Description)
'' End Sub
