
#Runtime EXE

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

#Imports System.Text.RegularExpressions

#Para DebugMode External
#Para AssemblyType winexe
'<font(\s*([a-z])+="([^"]*)")*>
'Lorem ipsum <font color="#fea" size="!">dolor sit amet, consectetuer adipiscing elit.
#WindowData frm_main
			550	500	Form		Width=550|Height=500|Text="RegexTester"
//DockStyle.Fill=5
	10	50	500	100	SplitContainer	pnl_split1	Dock=5|Orientation=0
1.Panel1	10	50	500	100	SplitContainer	pnl_split2	Dock=5
2.Panel1	10	50	500	100	TextBox	txtInput	Text="Lorem ipsum dolor sit amet, consectetuer adipiscing elit."|MultiLine=True|Dock=5|ScrollBars=2
2.Panel2	10	155	500	70	TextBox	txtPattern	Text="(\\w+)"|MultiLine=True|Dock=5|ScrollBars=2
1.Panel2	10	230	500	200	Tentec.Windows.iGridLib.IGrid	igOutput	Dock=5|Cols.Count=2
//	10	120	48	48	Picturebox	farbeAendern_2	Tag="#FBABAB"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/AcroRd32_2-48.png")
//	10	180	48	48	Picturebox	farbeAendern_3	Tag="#FFE680"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/SHELL32_308-48.png")
	10	440	200	30	Label	lab_Statusbar	Text="Hallo Welt"|Anchor=6

//DockStyle.Top=1
	0	0	0	40	Panel	pnl_TopBar	Dock=1|BackColor=ColorTranslator.FromHtml("#dcdfef")
1	10	5	90	30	Button	btn_FindMatches	Text="Find Matches"
//	70	10	55	30	Button	btn_Test_2	Text="Knopf2"
//	130	10	55	30	Button	btn_Test_3	Text="Knopf3" |BackColor=ColorTranslator.FromHtml("#fea")
1	420	5	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9
#EndData
'Anchor=13/15

Dim MAIN As frm_main

Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MAIN.Close
End Sub

Sub btn_FindMatches_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  Dim input = MAIN.txtInput.Text
  Dim pattern = MAIN.txtPattern.Text
  
  Dim matches = Regex.Matches(input, pattern)
  
  MAIN.igOutput.Rows.Clear()
  If matches.Count = 0 Then
    MAIN.lab_Statusbar.Text = "No matches"
    Exit Sub
  End If
  MAIN.igOutput.Cols.Count = matches(0).Groups.Count+1
  Dim i As Integer = 0
  For Each mat As Match In matches
    i += 1
    Dim ir = MAIN.igOutput.Rows.Add()
    ir.Cells(0).Value = i
    ir.Cells(1).Value = mat.Value
    For j As Integer = 1 To mat.Groups.Count-1
      ir.Cells(j + 1).Value = ""
      for each capt as capture in mat.Groups(j).Captures
        ir.Cells(j + 1).Value += """"+capt.Value+""" / "
      next
    Next
  Next
  
  MAIN.lab_Statusbar.Text = "Match Count: " & i
End Sub

Sub OnMyForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
  
End Sub

Sub Autostart()
  MAIN = New frm_main(Me)
  MAIN.Show()
  Application.Run(MAIN)
End Sub



