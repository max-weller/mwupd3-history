
#Para DebugMode External
#Para AssemblyType winexe

#WindowData MyForm
			400	280	Form		Width=400
	10	10	55	30	Button	btn_Test_1	Text="Knopf1"
	70	10	55	30	Button	btn_Test_2	Text="Knopf2"
	130	10	55	30	Button	btn_Test_3	Text="Knopf3" |BackColor=ColorTranslator.FromHtml("#fea")
	275	10	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
	223	49	181	184	TextBox	txtDetails	Text="...dieser Text wird später überschrieben"|MultiLine=True|Anchor=15|ScrollBars=2
	10	60	48	48	Picturebox	farbeAendern_1	Tag="#2B95B6"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/VCProject_7-48.png")
	10	120	48	48	Picturebox	farbeAendern_2	Tag="#FBABAB"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/AcroRd32_2-48.png")
	10	180	48	48	Picturebox	farbeAendern_3	Tag="#FFE680"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/SHELL32_308-48.png")
	9	251	370	20	Label	lab_Statusbar	Text="Hallo Welt"|Anchor=6
	72	247	105	15	CheckBox	CheckBox_1	Text="CheckBox_1"
	193	247	101	18	CheckBox	CheckBox_2	Text="CheckBox_2"
	72	45	138	187	GroupBox	GroupBox_0	Text="GroupBox_0"
1	12	27	103	30	CheckBox	CheckBox_3	Text="CheckBox_3"
1	13	65	102	31	Label	Label_0	Text="Label_0"
1	10	107	118	69	ListBox	ListBox_0	Text="ListBox_0"
	199	12	55	27	CheckBox	CheckBox_0	Text="CheckBox_0"
#EndData

			'' 400	280	Form		Width=400
'' 	10	10	55	30	Button	btn_Test_1	Text="Knopf1"
'' 	70	10	55	30	Button	btn_Test_2	Text="Knopf2"
'' 	130	10	55	30	Button	btn_Test_3	Text="Knopf3" |BackColor=ColorTranslator.FromHtml("#fea")
'' 	275	10	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
'' 	65	50	300	190	TextBox	txtDetails	Text="...dieser Text wird später überschrieben"|MultiLine=True|Anchor=15|ScrollBars=2
'' 	10	60	48	48	Picturebox	farbeAendern_1	Tag="#2B95B6"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/VCProject_7-48.png")
'' 	10	120	48	48	Picturebox	farbeAendern_2	Tag="#FBABAB"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/AcroRd32_2-48.png")
'' 	10	180	48	48	Picturebox	farbeAendern_3	Tag="#FFE680"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/SHELL32_308-48.png")
'' 	10	250	200	30	Label	lab_Statusbar	Text="Hallo Welt"|Anchor=6
'' 
Dim MyForm1 As MyForm

Sub txtDetails_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
  '' printLine 11,"","EVENT: txtDetails_TextChanged"
  '' printLine 1,"sender.Name",sender.Name
  MyForm1.lab_Statusbar.Text = MyForm1.txtDetails.TextLength & " Zeichen ..."
End Sub

Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MyForm1.Close
End Sub

' Dieser Event-Handler wird für alle Controls aufgerufen, deren Name
' mit btn_test_ anfängt, und die keinen passenderen Handler haben
Sub btn_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  EventLog("EVENT: btn_TEST_Click", "sender.Name=" + sender.Name)
End Sub

' Dieser Handler gilt für alle Controls, deren Name mit farbeAendern_ beginnt
Sub farbeAendern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  EventLog("EVENT: farbeAendern_Click", "sender.Name=" + sender.Name)
  Dim cBack = ColorTranslator.FromHtml(sender.Tag)
  MyForm1.txtDetails.BackColor = cBack
End Sub

Sub EventLog(ParamArray LogText() As String)
  MyForm1.txtDetails.AppendText(Join(LogText, vbTab) + vbNewLine)
End Sub

' In dieser Funktion sind schon alle Controls verfügbar
Sub OnMyForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
  #Data testMe As String
===========================
Hallo Welt !
Ich bin ein Datenblock
"**" mitten im Script "**"
' Sonderzeichen? '
kein Problem ;:_!"§$%&/()=?"
===========================
  #EndData
  MyForm1.txtDetails.Text = testMe
End Sub

Sub Autostart()
  MyForm1 = New MyForm(Me)
  MyForm1.Show()
  #If IDE = False Then
  Application.Run(MyForm1)
  #End If
End Sub
