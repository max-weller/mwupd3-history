

'#Para DebugMode intern
'#Para SilentMode false

'--> G L O B


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


Public ref

Sub AutoStart()
  msgbox("autostaaaart ???")
  Set ref = zz.getDockPanelRef("Toolbar|##|tbScriptWin|##|tb_testWindow.Adressen")
  With Ref
    .setBackColor "#DFDFFF"
    .setVisible True
    ' ref.Text = "aaabbbccc"
    
    .resetControls 5,5
    .activateEvents = "|ButtonMouseclick|   |TextboxGotFocus|TextBoxLostFocus|"
    .addLabel "header1", "     Adress-Verwaltung", "#333", "#fff",,,Auto,20 : .BR
    .BR
    .addIcon "myicon1", "http://http.cdnlayer.com/dreamincode/home/images/peopleicon.png"
    .addMenu "icon1", "pic_myicon1", "Left", Array("111","222","333")
    .addLabel "icon1_label", "...ich bin ein Label", , ,5,160
    .addButton "read","Lesen","#aaf"
    .addButton "openLine","Öffnen","#aaf"
    .addButton "newLine","Neu","#aaf"
    '' .addButton "test333","autoSize? ich bin ein langer Text","#aaf"
    .BR
    .AddControl "lst_names", "scriptIDE.LstBox", 140,60,Auto,150
    .BR 60 'zeilenhöhe zurücksetzen
    ZZ.handleEvent Ref.Element("lst_names"), "SelectedIndexChanged"
    'ZZ.handleEvent Ref.Element("lst_names"), "MouseDown"
    'Err.Raise 33333, "tb_testWindow.Autostart", "Stopcode ..."
    
    'ref.offsetX=5 : ref.offsetY=40
    'ref.direction=2
    '           ID      , XX  , LabelText       ,lblXX,lblColor
    .addLabel "header2", "     Anschrift", "#333", "#fff",,,Auto,20 : .BR
    .addTextbox "txtID2", 290 , "Vorname"       , 80  , "#ccc" : .BR
    .addTextbox "txtID1", 290, "Name"          , 80  , "#ccc" : .BR
    .addTextbox "txtID3", 225 , "Straße/Hausnr.", 80  , "#ccc"
    .addTextbox "txtID4", 60                                   : .BR
    .addTextbox "txtID5", 60  , "PLZ/Ort"       , 80  , "#ccc"
    .addTextbox "txtID6", 225                                 : .BR
    
    .addTextbox "txtID7", 120, "Telefon", 80, "#ccc"
    .addTextbox "txtID8", 120, "Handy", 40, "#ccc"
    'ref.addTextbox "txtID9", 140, "zzz", 60, "#ccc" 
    .br 40 'absatz
    
    .addLabel "header3", "     Details ...", "#333", "#fff",,,Auto,20 : .BR
    .addControl "rtf_test", "scriptIDE.RtfBox",,,Auto : .BR
    
    
    'stop
    
    .addButton "applyLine","     Übernehmen     ", "#00000000"
    '.addButton "doOut","     Ausgabe anzeigen ...    ","#aaf" : .BR
    
    '.addTextbox "out", 333, , , ,,,"multiline=400"
    .show
    
    readListbox
  End With
  'stop
End Sub

sub onUnknownEvent(sender,e)
  
End Sub

Sub onMenuEvent(e)
  trace "onMenuEvent"+e.EventName,typename(e)
  If e.EventName = "BeforeOpen" Then
    If e.MouseX>40 Then e.Cancel=True
  End If
  If e.EventName = "ItemClicked" Then
    trace "itemClick"
    Select Case e.keystring
      Case "111"
        Msgbox "einseinseins geklickt"
      Case Else
        Msgbox "Ein menüpunkt wurde angeklickt:"+vbnewline+e.keystring
    End Select
  End If
End Sub

Sub onTextboxEvent(e)
  If IsEmpty(ref) Then Set ref=zz.getDockPanelRef("Toolbar|##|tbScriptWin|##|tb_testWindow.Adressen")
  
  
  Dim descName: descName="txtDesc_"+Mid(e.Sender.Name,5)
  'e.Sender.Text=descname
  trace "Event:"+e.EventName,"Ctrl:"+e.Sender.Name
  trace "descLabel:",descname&"   "&typename(ref.Element(descName))
  Dim label
  'trace "form-ref",typename(ref)
  
  'Msgbox descName
  'Exit Sub
  if e.EventName = "GotFocus" Then
    ZZ.setBgColor e.Sender,"#DCE8F7"
    
    If IsEmpty(ref.Element(descName)) = False Then
      'Set label = 
      'trace "label",typename(label)
      ZZ.setBgColor ref.Element(descName), "#4170BE"
      ZZ.setFgColor ref.Element(descName), "#fff"
    End If
    'e.sender.SelectAll
    'e.Sender.Text = "test"
  End If
  if e.EventName = "LostFocus" Then
    ZZ.setBgColor e.Sender,"#fff"
    
    If IsEmpty(ref.Element(descName)) = False Then
      'Set lab = ref.Element(descName)
      ZZ.setBgColor ref.Element(descName), "#ccc"
      ZZ.setFgColor ref.Element(descName), "#000"
    End If
    
    'e.Sender.Text = "test"
  End If
End Sub

Sub onButtonEvent(e)
  If IsEmpty(ref) Then Set ref = ZZ.getDockPanelRef("Toolbar|##|tbScriptWin|##|tb_testWindow.Adressen")
  dim el,i, parts,out
  'e.sender.text=now
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_read" Then
    readListbox
  End If
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_newLine" Then
    Ref.Element("lst_names").itemAdd "neu, " + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + "|##|" + vbTab + "neu" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + "" + vbTab + ""
    'Ref.Element("lst_names").itemAdd "Mustermann, Max" + vbTab + vbTab + vbTab + vbTab + "|##|" + vbTab + "Mustermann" + vbTab + "Max" + vbTab + "Musterstraße" + vbTab + "11" + vbTab + "12345" + vbTab + "Musterstadt" + vbTab + "012345" + vbTab + "98765432" + vbTab + "RTF"
    
  End If
  
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_openLine" Then
    Dim line: line=Ref.Element("lst_names").SelectedItem
    'msgbox  Instr(line,"|##|")
    line=Right(line,Len(line)- Instr(line,"|##|"))
    'msgbox line
    parts = Split(line, vbTab)
    For i=1 to 8
      ref.Element("txt_txtID" & i).Text = parts(i)
    Next
    ref.Element("rtf_test").Text=""
    ref.Element("rtf_test").RTF = Replace(parts(9), "³zs²", vbNewLine) :err.clear
  End If
  
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_applyLine" Then
    out = ref.Element("txt_txtID1").Text & ", " & ref.Element("txt_txtID2").Text & vbTab & vbTab & vbTab & vbTab &  vbTab & vbTab & "|##|"
    For i=1 to 8
      out = out & vbTab & ref.Element("txt_txtID" & i).Text
    Next
    out = out & vbTab & Replace(ref.Element("rtf_test").RTF,vbNewLine, "³zs²")
    Ref.Element("lst_names").ItemText(Ref.Element("lst_names").SelectedIndex)=out
    
    saveListbox
  End If
  
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_doOut" Then
    Dim max: max=8
    'Dim out(): Redim out(max)
    
    Dim label
    out="----------------------------"
    For i=1 to max
      'out(i) = ""
      If ref.HasElement("txtDesc_txtID" & i) Then out = out+vbnewline+ref.Element("txtDesc_txtID" & i).Text + ":"
      out = out & vbTab & ref.Element("txt_txtID" & i).Text
    Next
    out = out & vbNewLine & "----------------------------" & vbNewLine
    out = out & ref.Element("rtf_test").Text
    out = out & vbNewLine & "----------------------------" & vbNewLine
    
    ref.Element("out").Text=out 'Join(out,vbNewline)
  End If
  If e.EventName = "MouseClick" And e.Sender.Name = "btn_test333" Then
    Msgbox "Heute ist der " & Date & " und es ist " & Time,64,"Datum und Uhrzeit"
  End If
  
End Sub

Sub saveListbox
  Dim i,out()
  Redim out(Ref.Element("lst_names").ItemCount)
  For i = 0 To Ref.Element("lst_names").ItemCount-1
    out(i) = Ref.Element("lst_names").ItemText(i)
  Next
  Dim RESULT: RESULT = Join(out, vbNewLine)
  ZZ.filePutContents "C:\_test\adressen.txt", RESULT
End Sub

Sub readListbox
  Dim i,LINES
  LINES = Split(ZZ.fileGetContents("C:\_test\adressen.txt"), vbNewLine)
  
  With Ref.Element("lst_names")
    .itemClear
    
    For i = 0 To UBound(LINES)
      ' parts = Split(LINES(i), vbTab)
      If LINES(i)<>"" Then _
      .itemAdd LINES(i)'parts(0) + ", "+ parts(1) + vbTab + vbTab + vbTab + vbTab + "|##|" + vbTab + LINES(i)
    Next
    
    
  End With
  'msgbox typename(el)
  'el.Focus
End Sub
