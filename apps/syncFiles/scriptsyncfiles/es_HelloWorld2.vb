'' helloWorld

#Para DebugMode internal
#Para SilentMode true

Dim PnlRef As ScriptedPanel
Dim PnlRef2 As ScriptedPanel
dim appId as string="HelloWorld2"

Sub AutoStart()
  '...Statposition festlegen
  dim mainForm as object = ZZ.IDEHelper.MainWindow
  dim startPosX as integer
  dim startPosY as integer
    if mainForm is nothing then
     startPosY=0
     startPosX=500
    else
      startPosX = mainForm.left+ mainForm.width -700
      startPosY = mainForm.Top+ 100
    end if
  '...1. Fenster erzeugen
  PnlRef = ZZ.IDEHelper.CreateDockWindow(appId+".Hello", 1)
  dim el as object
  with PnlRef
    .ResetControls(25, 20)
    with getOuterWindowRef(PnlRef)': err.number=0
      .top=startPosY
      .left=startPosX
      .height=185
      .width=450
    end with
    '...ACHTUNG: die Ausgabe (z.B. fenster-Hintergrundfarbe) erfolg nicht auf der Form                   
    '...sondern auf einem panel, das die gesammte Form ausfüllt.                                         
    '... deshalb darf z.B. für die Fenster-Hintergrund-Farbe NICHT getOuterWindowRef verwendet werden    
    .backColor=ColorTranslator.FromHtml("#394242") 
    .addIcon("myIcon", "http://icons3.iconfinder.netdna-cdn.com/data/icons/bnw/128x128/apps/package_network.png", , , -2)
    .insertX = 160:.insertY = 120:
    el=.addLabel("myLabel", "hello, world 1")
    el.Font = New System.Drawing.Font("Courier New", 20.0!, System.Drawing.FontStyle.bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#CFD4D4")
  end With
  
  '... 2.Fenster erzeugen
  PnlRef2 = ZZ.IDEHelper.CreateDockWindow(appId+".Team", 1)
  with PnlRef2
    .ResetControls(25, 20)
    with getOuterWindowRef(PnlRef2)': err.number=0
      .left=startPosX-25
      .top=startPosY+200
      .height=185
      .width=500
    end with
    .backColor=ColorTranslator.FromHtml("#394242")
    .addIcon("myIcon", "http://icons3.iconfinder.netdna-cdn.com/data/icons/Siena/128/globe%20yellow.png", , , -2)
    .insertX = 160:.insertY = 120:
    el=.addLabel("myLabel", "hello, world 2")
    el.Font = New System.Drawing.Font("Courier New", 20.0!, System.Drawing.FontStyle.bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    el.foreColor=ColorTranslator.FromHtml("#FEC408")
  end With
end sub


function getOuterWindowRef(panelRef as object) as object
  if panelRef.form.parent.parent is Nothing then
    '...falls es kein DockPanel-Fenster ist:
    :return panelRef.form
  else
    '...DockPanel-Fenster haben noch weitere Schichten dazwischen:
    :return panelRef.form.parent.parent
  end if
End function

