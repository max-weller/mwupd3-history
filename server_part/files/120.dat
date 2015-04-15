#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic

#Imports System.Windows.Forms.Form

#Reference wb_test_es1.dll
#Imports wb_test_es1.geckoWB

'#Imports Skybound.Gecko

Dim WithEvents FormRef As Form
Dim WithEvents WB As WebBrowser

dim WB2 as wb_test_es1.geckoWB
dim ref as object


Sub AutoStart()
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
  ref= ZZ.IDEHelper.CreateDockWindow("es_webbrowser1.tab")
  ''FormRef = ZZ.getDockPanelRef("Toolbar|##|tbScriptWin|##|webbrowser1.tab")'AsSystem.Windows.Forms,System.Windows.Forms.Form
  FormRef=ref.form
  FormRef.Text = "W=E=B"
  
  
  '--> Toolbar erzeugen
  ''With DirectCast(FormRef, Object)
 With ref
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
   WB2 = New wb_test_es1.geckoWB
  trace "autostart", "22222222"
   WB2.Top = 30'/WB.Top
  trace "autostart", "3333333333"
  'msgbox(wb2.name)
  trace "autostart", "444444444"
   ref.Controls.Add(WB2)
  ''''''WB2.Navigate("http://www.google.de/ie")
   WB2.visible=true
  trace "autostart", "5555555555555"
  
   
    onResizeControls()
  
  'WB.Top = 30
End Sub

public sub navigateWebbrowser(url)
  WB.Navigate(url)
end sub


  Sub onButtonEvent(e As Object)
    If e.Sender.Name = "btn_nav" Then
      WB.Navigate(ref.Controls("txt_url").Text)
    End If
    If e.Sender.Name = "btn_phpsuch" Then
      WB.Navigate("http://www.php.net/"+FormRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  Sub onTextboxEvent(e As Object)
    TRACE e.Keystring, e.Sender.Name
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_url" Then
      WB.Navigate(ref.Controls("txt_url").Text)
    End If
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_phpsuch" Then
      WB.Navigate("http://www.php.net/"+FormRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  
  Private Sub WB_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WB.DocumentCompleted
    FormRef.Controls("txt_url").Text=e.Url.Tostring
  End Sub

  Private Sub WB_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB.DocumentTitleChanged
    FormRef.Text="miniBrowser" 'WB.DocumentTitle
  End Sub

  Private Sub WB_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB.StatusTextChanged
    FormRef.Controls("txt_status").Text= WB.StatusText
  End Sub
  
  Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
    
    onResizeControls()
  End Sub

  Sub onResizeControls()
    'WB.Width = ref.Width - 10
    'WB.Height = ref.Height - 55
    'ref.Controls("txt_status").Top = ref.Height - 20
     'ref.Controls("txt_status").Width = ref.Width - 20
   
     WB2.Width = ref.Width - 10
     WB2.Height = ref.Height - 55
    ref.Controls("txt_status").Top = ref.Height - 20
    ref.Controls("txt_status").Width = ref.Width - 20
  End Sub

