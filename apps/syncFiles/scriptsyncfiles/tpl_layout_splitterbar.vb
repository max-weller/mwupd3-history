
''//  tpl_layout_splitterbar.vb
'-->
'--> G L O B 
#Para DebugMode intern
#Para SilentMode true
Public Const Auto = -2


'--> P A R A M E T E R
Const  winID = "{ScriptClass}.demoWin"
const  globHideOnClose as boolean = true          'bei true merkt er sich die Position
public winCaption as string="...ich bin ein Template für splitterbarContainer"


Dim WithEvents FormRef As Form
Dim PanelRef As ScriptedPanel
Dim toolbar As ScriptedToolstrip

Dim WithEvents tmr1 As Timer
Dim glob As cls_globPara


'-->
Function GetPanelRef()
  onTerminate() 'aufruf, um das alte leben ein bischen anzuhalten 
  ' If PanelRef IsNot Nothing Then Exit Sub
  PanelRef = ZZ.IDEHelper.CreateDockWindow(winID, 1) '1=popupWin
  formRef = getOuterWindowRef(panelRef)
  formRef.text=winCaption
  CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = globHideOnClose
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


Sub onTerminate()
  trace ("onTerminate")
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  'stop timer  : fehlen noch
  'stop resizer 
End Sub


Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  glob.saveFormPos(FormRef)
  glob.para("myMes")="ich bin eine Nachricht"
  glob.saveParaFile()
End Sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
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
'--> A U T O S T A R T
Sub AutoStart()
  PanelRef = GetPanelRef()    '... panelRef ist als globale var definiert
  glob = ZZ.newGlobPara()
  glob.readFormPos(FormRef)
  'msgBox(glob.para("myMes"))
  ''FormRef.Text = "Music Box"
  
  Dim pLeft, pRight As ScriptedPanel
  
  With PanelRef
    .resetControls (5,5)
    dim el as object
    .activateEvents = "|ButtonMouseClick| |TEXTBOXTEXTCHANGED|"
    
    '--> ...splitContainer
    el=.addSplitcontainer("splMain", "left", pLeft, "right", pRight, DockStyle.Fill)
    el.backColor=ColorTranslator.FromHtml("#88f")
    el.splitterDistance=200
    
     
    '--> ... Left Pane
    with pLeft
      .resetControls (5,5)
      .backColor=ColorTranslator.FromHtml("#bbf")
      toolbar = pLeft.addToolstrip("main", 5, 5, 50, 20)
      toolbar.Dock=0
      toolbar.ImageScalingSize = New Size(16,16)
      'toolbar.addButton("cmdMenu", "", flags:=ButtonFlags.StartMenu, iconURL:="http://www.iconfinder.net/data/icons/nuvola2/16x16/apps/kmenuedit.png")
      toolbar.addButton("cmdPlay", "Play", iconURL:="http://www.iconfinder.net/data/icons/humano2/16x16/actions/gtk-media-play-ltr.png")
      toolbar.addButton("cmdStop", "Stop", iconURL:="http://www.iconfinder.net/data/icons/humano2/16x16/actions/gtk-media-pause.png")
      toolbar.addButton("cmdIndexDirectory", "Liste aktualisieren", iconURL:="http://www.iconfinder.net/data/icons/icojoy/shadow/standart/gif/24x24/001_38.gif")
      toolbar.addButton("cmdFolderlist", "Ordner")
      'toolbar.ActiveMenu = Nothing
      
      .BR(30)
      .addIcon("cmdSuchIcon", "http://www.iconfinder.net/data/icons/icojoy/shadow/standart/gif/24x24/001_38.gif")
      .addTextbox("suchBox", auto) : 
      .BR
    end with 
 
 
    '--> ... Right Pane
    with pRight
      .resetControls (5,5)
      .backColor=ColorTranslator.FromHtml("#bbf")
      .addIcon("icoSound", "http://www.iconfinder.net/data/icons/nuvola2/22x22/apps/kmix.png")
      .addTextbox("url", Auto, X:=40)
    end with
 
    tmr1 = New Timer
    tmr1.Interval = 1000 : tmr1.Enabled = True
  End With
  
  'onResizeControls()
End Sub


'--> 
'--> E V E N T S

Sub onButtonEvent(e)
  '#######################################
  ''msgBox (e.keyString)
  '' MAX: ich brauch die MouseKnöpfe
  '#######################################
  Dim nameParts() = Split(e.Sender.Name+"_", "_")
  dim name=nameParts(1)

  printLine 11, "e.Sender.Name" , e.Sender.Name
  printLine 12, "name" , name
  ''errorText("")
  
  dim tag as string = e.sender.tag
  printLine 12, "tag" , tag.toString
  dim tagDATA()= Split(tag, "<³³>")
  
  '' monitorAdd("==============================================")
  '' monitorAdd("Sender.Name ===>" + e.Sender.Name)
  '' monitorAdd("=>" +name(1)+"<=")
  
  Select Case name(1)
    case "ABC"        : cmd_ABC (e, tagDATA)
    case "XYZ"        : cmd_XYZ  (e, tagDATA)
    Case else         
      trace "ERR: onButtonEvent(e): '"+name(1)+"' ...Typ nicht erkannt: "
  End Select
end Sub


sub cmd_ABC (e, tagDATA)
  'msgBox("cmd_ABC")
end sub

sub cmd_XYZ (e, tagDATA)
  'msgBox("cmd_ABC")
end sub














