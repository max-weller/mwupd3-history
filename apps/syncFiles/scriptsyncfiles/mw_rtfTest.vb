'--> ... mw_formTest1

#Para DebugMode extern
#runtime EXE

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Text.RegularExpressions
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

#Para CompilerOptions /debug+ /debug:pdbonly
#Para AssemblyType winexe

'-->
'--> Window Data

#WindowData MyForm
//	x	y	w	h
			430	480	Form		Width=430|Height=480
	10	10	85	30	Button	btn_Test_1	Text="Code ausgeben"
	100	10	55	30	Button	btn_Test_2	Text="Bold"
	160	10	55	30	Button	btn_Test_3	Text="Knopf3" |BackColor=ColorTranslator.FromHtml("#fea")
	275	10	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
	20	50	400	185	RichTextBox	rtf1	Text="...dieser Text wird später überschrieben"|MultiLine=True|Anchor=13|ScrollBars=2
	20	260	400	185	TextBox	txtDetails	Text="...dieser Text wird später überschrieben"|MultiLine=True|Anchor=15|ScrollBars=2

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

'-->

Dim MyForm1 As MyForm


Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MyForm1.Close
End Sub


Sub btn_Test_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MyForm1.txtDetails.text = MyForm1.rtf1.Rtf
End Sub

Sub btn_Test_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  
  
End Sub


'-->

' In dieser Funktion sind schon alle Controls verfügbar
Sub OnMyForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
  #Data testMe As String
{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}}\viewkind4\uc1\pard
\f0\fs17 ...xxxxxxxxxxxxxxxxdieser Text wird sp\'e4ter \'fcberschrieben\par
}

  #EndData
  MyForm1.rtf1.RTF = testMe
End Sub

Sub Autostart()
  MyForm1 = New MyForm(Me)
  MyForm1.Show()
  Application.Run(MyForm1)
End Sub
