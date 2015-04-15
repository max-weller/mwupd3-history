

#Para DebugMode Internal
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



#Reference ScintillaNet.dll
'#Imports ScintillaNet

Sub AutoStart()
  ZZ.IDEHelper.RegisterFileactionHandler(Me)
  trace "RegisterFileactionHandler(Me)", "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   O K "
End Sub

'... onEnter...
<FileAction("*.vb", "onEnter", "http://mw.teamwiki.net/docs/img/icons/build.png", HandlesKeyString:="CTRL-RETURN")> _
Public Sub onEnter(ByVal tab As Object)
  Dim scin As ScintillaNet.Scintilla = tab.RTF
  Dim codeText As String = scin.Text
  msgBox("onEnter")
  'Clipboard.Clear
  'Clipboard.SetText(codeText)
End Sub
  '' 
'... erzeugt auch gleich den knopf dazu
<FileAction("*.txt", "Kopieren", "http://mw.teamwiki.net/docs/img/icons/build.png", HandlesKeyString:="CTRL-ALT-C")> _
Public Sub DateiAktion1(ByVal tab As Object)
  Dim scin As ScintillaNet.Scintilla = tab.RTF
  Dim codeText As String = scin.Text
  Clipboard.Clear
  Clipboard.SetText(codeText)
End Sub
  '' 
<FileAction("*.*", "Speichern unter...", "http://www.iconfinder.net/data/icons/discovery/16x16/actions/document-save-as.png")> _
Public Sub DateiAktion2(ByVal tab As IDockContentForm)
  If Left(tab.URL,5) <> "loc:/" Then Exit Sub
  Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  Using sfd As New SaveFileDialog
    sfd.Filename = fileSpec
    If sfd.ShowDialog() = DialogResult.OK Then
      Dim scin As ScintillaNet.Scintilla = tab.RTF
      ZZ.filePutContents(sfd.FileName, scin.Text)
    End If
  End Using
  
End Sub

<FileAction("*.vb", "Test 1", "http://es.teamwiki.net/docs/icons/draw-square-blue.png")> _
Public Sub DateiAktion3(ByVal tab As IDockContentForm)
  If Left(tab.URL,5) <> "loc:/" Then Exit Sub
  Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  :dim scriptClass = callScriptClassTestProc(filespec)
  :scriptClass.test1()
  :if ERR.number<>0 then
  :   ERR.number=0
      printLine 11, "ERR", "ERR: Sub 'test1()' ...nicht gefunden" 
      trace         "ERR", "ERR: Sub 'test1()' ...nicht gefunden" 
      exit sub
  end if
End Sub

<FileAction("*.vb", "Test 2", "http://es.teamwiki.net/docs/icons/draw-square-blue.png")> _
Public Sub DateiAktion4(ByVal tab As IDockContentForm)
  If Left(tab.URL,5) <> "loc:/" Then Exit Sub
  Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  callScriptClassTestProc(filespec)
  :dim scriptClass = callScriptClassTestProc(filespec)
  :scriptClass.test2()
  :if ERR.number<>0 then
  :   ERR.number=0
      printLine 11, "ERR", "ERR: Sub 'test2()' ...nicht gefunden" 
      trace         "ERR", "ERR: Sub 'test2()' ...nicht gefunden" 
      exit sub
  end if
End Sub

<FileAction("*.vb", "Test 3", "http://es.teamwiki.net/docs/icons/draw-square-blue.png")> _
Public Sub DateiAktion5(ByVal tab As IDockContentForm)
  If Left(tab.URL,5) <> "loc:/" Then Exit Sub
  Dim fileSpec = ProtocolService.MapToLocalFilename(tab.URL)
  :dim scriptClass = callScriptClassTestProc(filespec)
  :scriptClass.test3()
  :if ERR.number<>0 then
  :   ERR.number=0
      printLine 11, "ERR", "ERR: Sub 'test3()' ...nicht gefunden" 
      trace         "ERR", "ERR: Sub 'test3()' ...nicht gefunden" 
      exit sub
  end if
End Sub



:function callScriptClassTestProc(filespec as string)
  printLine 11, "...", "..." 
  printLine 1, "splitFilespecData", fileSpec

  dim path, name, ext as Object
  dim scriptClass as object
  :zz.splitFilespecData(filespec, path, name, ext)
  :printLine 2, "splitFilespecData", name
  :scriptClass=zz.scriptClass(name)
  :zz.doEvents
  :zz.doEvents
  :zz.doEvents
  :if scriptClass is Nothing then
     :ERR.number=0
     printLine 11, "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     trace         "ERR", "ERR: scriptKlasse '"+name+"' ...nicht gefunden" 
     return nothing
  end if
  return scriptClass
End function









