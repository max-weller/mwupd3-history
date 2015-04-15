#Para DebugMode intern

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

Dim PnlRef As ScriptedPanel

Dim withevents tim as new Timer

Sub AutoStart()
  PnlRef = ZZ.CreateWindow("mk2_buttontest.Hello", DWndFlags.DockWindow)
  with PnlRef
    'Definieren ob
    CType(.Form, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = True
    .Form.Size = New Size(300,200)
    .ResetControls(20, 30)
    .addIcon("myIcon", "http://mw.teamwiki.net/docs/img/win-icons/spcplui_138-48.png", , , -2)
    .br
    .addButton("btn1", "test")
    .br
    .addLabel("myLabel", "hello, world")
  end With
End Sub

Sub tim_Tick(sender As Object, e As EventArgs) Handles tim.Tick
  tim.stop
  msgbox("test")
End Sub
  
Sub btn1_MouseClick(e)
  tim.Interval = 5000
  tim.start
End Sub