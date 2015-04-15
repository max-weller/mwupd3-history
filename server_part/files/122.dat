#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic

#Imports System.Windows.Forms.Form

#Reference C:\yEXE\wb_test_es1.dll
#Reference C:\yEXE\Skybound.Gecko.dll
#Imports wb_test_es1.geckoWB

'#Imports Skybound.Gecko

Dim WithEvents FormRef As Form

dim WithEvents WB2 as Skybound.Gecko.GeckoWebBrowser
dim panelRef as object


Sub AutoStart()
  Skybound.Gecko.Xpcom.Initialize("C:\yEXE\xulrunner\")
  
  printLine 1, ".", ""
  printLine 2, "..", ""
  printLine 3, "...", ""
  printLine 4, "....", ""
  printLine 1, "autostart", now.tostring
  printLine 5, ".....", ""
  printLine 6, "......", ""
  printLine 7, ".......", ""
  printLine 8, "........", ""
  '--> Formreferenz holen
  panelRef= ZZ.IDEHelper.CreateDockWindow("{ScriptClass}.tab")
  ''FormRef = ZZ.getDockPanelRef("Toolbar|##|tbScriptWin|##|webbrowser1.tab")'AsSystem.Windows.Forms,System.Windows.Forms.Form
  FormRef=panelRef.form
  FormRef.Text = "W=E=B"
  
  
  '--> Toolbar erzeugen
  ''With DirectCast(FormRef, Object)
 With panelRef
    .resetControls  (10,5,1)
    .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
                 'id         txtXX   lab             labXX labBG  X     Y   noBR
    .addTextbox ("url",      400   , "URL:"         , 50  , "#aaa")
    .addButton  ("nav",      ">>>" , "#8ca")
    .addTextbox ("phpsuch",  200   , "PHP-Referenz:", 80  , "#aaa")
    .addButton  ("phpsuch",  ">>>" , "#8ca")
    .addTextbox ("status",   350   ,                ,     ,       , 0   , 0 )
  End With
             
  '--> WebBrowser erzeugen
  'WB = New WebBrowser
  
  
  'WB.Top = 0
  'WB.Top = 30'/WB.Top
  'ref.Controls.Add(WB)
  'WB.Navigate("http://www.google.de/ie")
  
  trace "autostart", "11111111111"
   WB2 = New Skybound.Gecko.GeckoWebBrowser
  trace "autostart", "22222222"
   WB2.Top = 30'/WB.Top
  trace "autostart", "3333333333"
  'msgbox(wb2.name)
  trace "autostart", "444444444"
   panelRef.Controls.Add(WB2)
  ''''''WB2.Navigate("http://www.google.de/ie")
   WB2.visible=true
  trace "autostart", "5555555555555"
  
   
    onResizeControls()
  
  'WB.Top = 30
End Sub

public sub navigateWebbrowser(url)
  WB2.Navigate(url)
end sub


  Sub onButtonEvent(e As Object)
    If e.Sender.Name = "btn_nav" Then
      WB2.Navigate(panelRef.Controls("txt_url").Text)
    End If
    If e.Sender.Name = "btn_phpsuch" Then
      WB2.Navigate("http://www.php.net/"+panelRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  Sub onTextboxEvent(e As Object)
    TRACE e.Keystring, e.Sender.Name
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_url" Then
      WB2.Navigate(panelRef.Controls("txt_url").Text)
    End If
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_phpsuch" Then
      WB2.Navigate("http://www.php.net/"+panelRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  
  Private Sub WB_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WB2.DocumentCompleted
    panelRef.Controls("txt_url").Text = WB2.URL.ToString()
  End Sub

  Private Sub WB_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB2.DocumentTitleChanged
    panelRef.Text="miniBrowser" 'WB.DocumentTitle
  End Sub

  Private Sub WB_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB2.StatusTextChanged
    panelRef.Controls("txt_status").Text= WB2.StatusText
  End Sub
  
  Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
    
    onResizeControls()
  End Sub

  Sub onResizeControls()
    'WB.Width = ref.Width - 10
    'WB.Height = ref.Height - 55
    'ref.Controls("txt_status").Top = ref.Height - 20
     'ref.Controls("txt_status").Width = ref.Width - 20
   
     WB2.Width = panelRef.Width - 10
     WB2.Height = panelRef.Height - 55
    panelRef.Controls("txt_status").Top = panelRef.Height - 20
    panelRef.Controls("txt_status").Width = panelRef.Width - 20
  End Sub

