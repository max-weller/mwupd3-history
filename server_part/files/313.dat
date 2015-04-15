#Para DebugMode extern
'=== Script/Programm-Icon === wird in .exe kompiliert
'#Para IconFile c:\icon.ico

'=== Standard Programmbibliotheken ===
#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports Tentec.Windows.iGridLib
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper


'=== Scripted-Panel ===
Dim PnlRef As ScriptedPanel

Sub AutoStart()
  PnlRef = ZZ.CreateWindow("mk2_buttontest.Hello", DWndFlags.DockWindow)
  With PnlRef
    '=== Formeigenschaften ===
    .Form.Size = New Size(300,200) 'Startgroesse
    .Form.Text = "Template Script" 'Fenstertitel
    'DockWindow Position merken oder nicht
    CType(.Form, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = True
    '=== Steuerelemente erstellen ===
    .ResetControls(50, 50)                                  'Startposition aller Controls (x, y)
    .addButton("Button1", "Button-Beschriftung")            'Button (Name as String, Caption as String)
    .br                                                     'neue Zeile
    .addTextbox("Textbox1", 50, "Textbox-Beschriftung", 70) 'Textbox (Name as String, Width as Integer, txtBoxLabText as String, txtBoxLabWidth as Integer)
    .br
    .addLabel("Label1", "Label-Beschriftung")               'Label (Name as String, Caption as String)
    .addIcon("Picturebox1", "http://mw.teamwiki.net/docs/img/win-icons/spcplui_138-48.png", , , )'Picturebox (Name as String, FilePath as String, xPos as Integer, yPos as Integer, Size as Integer)
  End With
End Sub

Sub Button1_MouseClick(e)
    PnlRef("lab_Label1").Text = "Hello World"
    MsgBox("Labelbeschriftung geändert.", 64)
End Sub