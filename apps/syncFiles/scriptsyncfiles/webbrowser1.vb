#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic

Dim WithEvents FormRef As Form
Dim WithEvents PanelRef As Object
Dim WithEvents WB As WebBrowser



Sub navigateWebbrowser(url As String)
  If PanelRef Is Nothing Then PanelRef = ZZ.IDEHelper.CreateDockWindow("webbrowser1.tab")'AsSystem.Windows.Forms,System.Windows.Forms.Form
  PanelRef.Element("wb1").Navigate(url)
End Sub

Sub AutoStart()
  'BuildMainToolbar
  trace "Hallo welt"
  
  '--> Formreferenz holen
  PanelRef = ZZ.IDEHelper.CreateDockWindow("webbrowser1.tab")'AsSystem.Windows.Forms,System.Windows.Forms.Form
  FormRef = PanelRef.Form
  FormRef.Text = "..."
  FormRef.Show()
  trace "Hallo welt2"
  
  '--> Toolbar erzeugen
  With PanelRef
    .resetControls ()
    .activateEvents = "|ButtonMouseClick|   |TextboxKeyDown|"
                 'id         txtXX   lab             labXX labBG  X     Y   
    .addTextbox ("url",      400   , "URL:"         , 50  , "#aaa")
    .addButton  ("nav",      ">>>" , "#8ca")
    .addTextbox ("phpsuch",  200   , "PHP-Referenz:", 80  , "#aaa")
    .addButton  ("phpsuch",  ">>>" , "#8ca")
    .addTextbox ("status",   350   ,                ,     ,       , 0   , 0 )
    
    
    '--> WebBrowser erzeugen
    WB = New WebBrowser
    'WB.Top = 0
    WB.Top = 30'/WB.Top
    WB.name = "wb1"
    .Controls.Add(WB)
    WB.Navigate("http://www.google.de/ie")
  End With
  onResizeControls()
  
  'WB.Top = 30
End Sub

  Sub onButtonEvent(e As Object)
    printLine 1,"buttonEvent", e.ID
    If e.Sender.Name = "btn_nav" Then
      WB.Navigate(PanelRef.Controls("txt_url").Text)
    End If
    If e.Sender.Name = "btn_phpsuch" Then
      WB.Navigate("http://www.php.net/"+PanelRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  Sub onTextboxEvent(e As Object)
    TRACE e.Keystring, e.Sender.Name
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_url" Then
      WB.Navigate(PanelRef.Controls("txt_url").Text)
    End If
    If e.Keystring = "RETURN" And e.Sender.Name = "txt_phpsuch" Then
      WB.Navigate("http://www.php.net/"+PanelRef.Controls("txt_phpsuch").Text)
    End If
  End Sub
  
  Private Sub WB_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WB.DocumentCompleted
    PanelRef.Element("txt_url").Text=e.Url.Tostring
    printLine 1, "URL", e.URL.Tostring
  End Sub

  Private Sub WB_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB.DocumentTitleChanged
    FormRef.Text=WB.DocumentTitle
  End Sub

  Private Sub WB_StatusTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles WB.StatusTextChanged
    PanelRef.Element("txt_status").Text= WB.StatusText
  End Sub
  
  Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
    onResizeControls
  End Sub
  
  Sub onResizeControls()
    WB.Width = panelRef.Width - 10
    WB.Height = panelRef.Height - 55
    PanelRef.Element("txt_status").Top = panelRef.Height - 20
    PanelRef.Element("txt_status").Width = panelRef.Width - 20
  End Sub

