




'--> 
'--> mw_fraktal

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference siaIDEMain.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports System.Drawing.Drawing2D
#Imports System.Drawing.Imaging
#Imports System.Runtime.InteropServices
#Imports ScriptIDE.Core
#Imports ScriptIDE.Main
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

'' es_UserToolbar3

#runtime exe
#Para DebugMode internal
'#Para SilentMode true

Public Declare Function calcMandelbrot Lib "apfelLib.dll" Alias "#2" _
  (real As Single, imag As Single, maxIterations As Integer) As Integer
  
'' APFELLIB_API int calcMandelbrot(float real, float imag, int maxIterations) {
''  float Znr=0.0, Zni=0.0, Zr=0.0;
''  
''  for(int i = 1; i<=maxIterations; i++) {
''   Zr = Znr;
''   Znr = (Zr*Zr) - (Zni*Zni);
''   Zni = Zr*Zni*2;
''   Znr = Znr + real;
''   Zni = Zni + imag;
''   if(Znr*Znr + Zni*Zni > 4) return i;
''  }
''  return 0;
'' }

Structure ComplexNr
  Dim r,i As double
End Structure


#WindowData MyForm
//	x	y	w	h
			430	320	Form		Width=680|Height=320|Text="Mandelbrot"
	10	10	55	30	Button	btn_Test_1	Text="Fraktal"
	70	10	55	30	Button	btn_Cancel	Text="Abbruch"|Enabled=false
	140	10	55	20	Label	lab_1	Text="xPos"
	140	30	55	30	TextBox	txt_x	Text="-2,5"
	200	10	55	20	Label	lab_2	Text="yPos"
	200	30	55	30	TextBox	txt_y	Text="-1"
	260	10	60	20	Label	lab_3	Text="Zoomstufe"
	260	30	55	30	TextBox	txt_zoom	Text="200"
	320	10	60	20	Label	lab_4	Text="Genauigkeit"
	320	30	55	30	TextBox	txt_maxit	Text="1000"
	400	20	160	40	Label	labTime	Text="...Willkommen!"|Font=New Font("Microsoft Sans Serif", 12, FontStyle.Regular, GraphicsUnit.Point)
	580	10	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
	10	60	650	220	Picturebox	pic1	Anchor=15|Borderstyle=1
#EndData
'' 
'' <Form id="MyForm" Width="680" Height="320" Text="Mandelbrot">
''   <Panel id="pnl_top" Dock="DockStyle.Top" height="60">
''     
''     <Set Position="10;10" MarginX="5" Size="55;30" />
''     <Button id="btn_Test_1" Text="Fraktal" />
''     <Button id="btn_Cancel" Text="Abbruch" Enabled="false" />
''     
''     <Set Position="140;10" MarginX="5" Size="55;20" />
''     <Label id="lab_1" Text="xPos" />
''     <Label id="lab_2" Text="yPos" />
''     <Label id="lab_3" Text="Zoomstufe" />
''     <Label id="lab_4" Text="Genauigkeit" />
''     
''     <Set Position="140;30" MarginX="5" Size="55;30" />
''     <Textbox id="txt_x" Text="-2,5" />
''     <Textbox id="txt_y" Text="-1" />
''     <Textbox id="txt_zoom" Text="200" />
''     <Textbox id="txt_maxIt" Text="1000" />
''     
''     <Set Position="400;20" />
''     <Label id="labTime" Size="160;40" Text="...Willkommen!" Font="Microsoft Sans Serif; 12; Regular" />
''     
''     <Set Position="580;10" />
''     <Button id="btn_Ende" Size="90;30" Text="Beenden" Anchor="9" Font="Segoe UI; 12; Regular" />
''     
''   </Panel>
''   
''   <Picturebox id="pic1" Bounds="10;60;650;220" Anchor="15" BorderStyle="1" />
'' </Form>
'' 
'' 

'' 
'' <Form id="MyForm" Width="680" Height="320" Text="Mandelbrot">
''   <Panel id="pnl_top" Dock="DockStyle.Top" height="60">
''     
''     <Set Position="10;10" MarginX="5" Size="55;30" />
''     <Button id="btn_Test_1" Text="Fraktal" />
''     <Button id="btn_Cancel" Text="Abbruch" Enabled="false" />
''     
''     <Set Position="140;10" MarginX="5" Label.Size="55;20" TextBox.Size="55;30" />
''     <cv><Label id="lab_1" Text="xPos" /><Textbox id="txt_x" Text="-2,5" /></cv>
''     
''     <Div Orientation="Vertical">
''       <Label id="lab_2" Text="yPos" />
''       <Textbox id="txt_y" Text="-1" />
''     </Div>
''     
''     <Div Orientation="Vertical">
''       <Label id="lab_3" Text="Zoomstufe" />
''       <Textbox id="txt_zoom" Text="200" />
''     </Div>
''     
''     <Div Orientation="Vertical">
''       <Label id="lab_4" Text="Genauigkeit" />
''       <Textbox id="txt_maxIt" Text="1000" />
''     </Div>
''     
''     
''     
''     <Set Position="400;20" />
''     <Label id="labTime" Size="160;40" Text="...Willkommen!" Font="Microsoft Sans Serif; 12; Regular" />
''     
''     <Set Position="580;10" />
''     <Button id="btn_Ende" Size="90;30" Text="Beenden" Anchor="9" Font="Segoe UI; 12; Regular" />
''     
''   </Panel>
''   
''   <Picturebox id="pic1" Bounds="10;60;650;220" Anchor="15" BorderStyle="1" />
'' </Form>
'' 
'' 


Dim MAIN As MyForm
Dim colors(1001) as color
Dim colorAlloc(1001) As integer'color
dim colorCount as integer=1000
Dim cancelFlag As Boolean

Dim glob As cls_globPara

Dim teiler,maxIterations As Integer
Dim xVon,xBis,yVon,yBis As Integer

:Sub btn_Test_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  ZZ.TimerStart()
  MAIN.btn_Cancel.Enabled = True
  
  teiler = MAIN.txt_zoom.Text
  xVon = MAIN.txt_x.Text * teiler
  xBis = xVon+MAIN.pic1.Width
  yVon = MAIN.txt_y.Text * teiler
  yBis = yVon+MAIN.pic1.Height
  maxIterations = MAIN.txt_maxit.Text
  'calculateMandelbrot
  Dim th As New Threading.Thread(AddressOf calculateMandelbrot)
  th.Start
End Sub

Dim bmp As Bitmap

delegate Sub delonPaintFinished()
Sub onPaintFinished()
  MAIN.pic1.image = bmp
  MAIN.btn_Cancel.Enabled = False
  MAIN.labTime.Text = "Zeit: " & (ZZ.TimerGetElapsed()/1000).ToString()
End Sub

:Sub calculateMandelbrot()
  cancelFlag = False
  Dim col As Integer
  
  bmp = New Bitmap(xBis-xVon+1, yBis-yVon+1, PixelFormat.Format32bppArgb)
  Dim ImageBytes() As Int32
  Dim bounds As Rectangle = New Rectangle(0, 0, bmp.Width, bmp.Height)
  Dim m_BitmapData As BitmapData = bmp.LockBits(bounds, _
    Imaging.ImageLockMode.ReadWrite, _
    Imaging.PixelFormat.Format32bppRgb)
  Dim RowSizeBytes As Integer = m_BitmapData.Stride
  Dim total_size As Integer = m_BitmapData.Stride * m_BitmapData.Height
  ReDim ImageBytes(total_size \ 4)
  Dim pix As Integer=0
  
  'dim multi As single = 1000/maxIterations
  'TRACE "multi",multi
  TRACE colorCount,maxIterations
  
  For y as Integer = yVon To yBis '-2*teiler To 1*teiler
    For x as Integer = xVon To xBis '-1*teiler To 1*teiler
      'v = calcZ(New ComplexNumber(x,y))
      'col = calcZ(x/teiler,y/teiler)
      col = calcMandelbrot(x/teiler, y/teiler, maxIterations)
      'If v then col=color.black else col=color.white
      'bmp.SetPixel(x-xVon, y-yVon, colorAlloc(col*multi))
      ImageBytes(pix) = colorAlloc(col*1000\maxIterations)
      pix+=1
    Next
    If (y mod 4 )=0 Then
      'MAIN.pic1.image = bmp
      If cancelFlag Then  Exit Sub
    End If
  Next
  
  Marshal.Copy(ImageBytes, 0, m_BitmapData.Scan0, total_size\4)
  bmp.UnlockBits(m_BitmapData)
  'ImageBytes = Nothing
  
  MAIN.Invoke(New delonPaintFinished(AddressOf onPaintFinished))
End Sub

Sub btn_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  cancelFlag=True
  MAIN.btn_Cancel.Enabled=False
End Sub

Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  main.close()
End Sub

Sub pic1_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) ''... Handles Control.MouseDown
  If e.Button = MouseButtons.Right Then
    Dim para As String = ((e.X + xVon) / teiler).ToString + "|" + ((e.Y + yVon) / teiler).ToString
    ZZ.InterProc.SendCommand("fraktal02", "Goto", para)
  Else
    MAIN.txt_x.Text = ((e.X + xVon) / teiler).ToString
    MAIN.txt_y.Text = ((e.Y + yVon) / teiler).ToString
  End If
End Sub



Sub OnMyForm_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs)
  glob.saveFormPos(MAIN)
  glob.saveTuttiFrutti(MAIN)
  glob.saveParaFile()
End Sub

Sub OnMyForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MAIN.CheckForIllegalCrossThreadCalls= False
  glob = New cls_globPara()
  glob.readFormPos(MAIN)
  glob.readTuttiFrutti(MAIN)
  colors(0) = Color.Black
  dim i as integer
  'set up a palette
  For i = 1 To 39
    colors(i) = Color.FromArgb(0, i, 5*i )
  Next
  For i = 40 To 127
    colors(i) = Color.FromArgb(0, 0, 2*i )
  Next
  For i = 128 To 191
    colors(i) = Color.FromArgb(0, (i-128) * 4, (191 - i) * 4)
  Next
  For i = 192 To 255
    colors(i) = Color.FromArgb((i - 192) * 4, (255 - i) * 4, 0)
  Next
  For i = 256 To 319
    colors(i) = Color.FromArgb(255, 0, (i - 256) * 4)
  Next
  For i = 320 To 383
    colors(i) = Color.FromArgb((383 - i) * 4, 0, 255)
  Next
  colorCount=383
  '' For i =1 to 1000
  ''   dim b=math.log(i)*36
  ''   colors(i)=color.fromargb(255,b,b,b)
  '' next
  '' colorCount=1000
  colorAlloc(0)=&HFFFFFFFF
  Dim max As Single=math.Log(1000)
  For i =1 to 1000
    dim b=math.log(i)/max*colorcount'i*colorCount\1000
    'colorAlloc(i)= colors(b)
    colorAlloc(i)= (&HFF<<24) OR (cint(colors(b).R)<<16) OR (cint(colors(b).G)<<8) OR (cint(colors(b).B))
    'TRACE colortranslator.ToHtml( colors(b)), colorValue.tostring("x25")
  Next
  Dim bmp As New Bitmap(1001, 100)
  For i =0 to 1000
    for y as integer=0 to 40
      bmp.setpixel(i,y,colors(i))
    next
    for y as integer=41 to 80
      'bmp.setpixel(i,y,colorAlloc(i))
    next
  next
  
  MAIN.pic1.image=bmp
End Sub

Sub Autostart()
  MAIN = New MyForm(Me)
  MAIN.Show()
  Application.Run(MAIN)
End Sub



'--> 
'--> altlasten ...

':Function calcZ(ByVal c As ComplexNumber) As Boolean
:Function calcZ(ByVal cr As Double, ByVal ci as double) As color
  Dim i As Integer
  'Dim Z As ComplexNr,Zn As ComplexNr
  Dim Znr,Zni,Zr As Double
  'Zn.r = Zn_r : Zn.i = Zn_i
  For i=1 To 1000
    'hoch 2
    Znr = Znr ^2 - Zni'^2
    Zni = 2*Znr*Zni
    'Z=Zn
    'Zn.r = (Z.r*Z.r) - (Z.i*Z.i)
    'Zn.i = (Z.r * Z.i) *2
    
    '' Zr=Znr
    '' Znr = (Zr*Zr) - (Zni*Zni)
    '' Zni = (Zr*Zni) * 2
    
    'a.Real * b.Imaginary + a.Imaginary * b.Real
    'plus c
    Znr = Znr + cr
    Zni = Zni + ci
    
    
    If Znr^2 + Zni^2 > 4 Then Return colors(i)
    'If Math.sqrt(Zn.r^2 + Zn.i^2) > 2 Then Return colors(i)'color.fromargb(255,i/4,i/4,255-i/4)
  Next
  
  Return color.black
End Function





Function BerechneMandelbrot(C As ComplexNumber) As Integer
  Dim Z As ComplexNumber = 0
  For n = 1 To 1000
    Z = Z^2 + c
    If Z.GetAbs() > 2 Then Return n
  Next
  Return -1
End Function













Structure ComplexNumber
    Public Real As Double
    Public Imaginary As Double

    Public Sub New(ByVal realPart As Double, ByVal imaginaryPart As Double)
        Me.Real = realPart
        Me.Imaginary = imaginaryPart
    End Sub

    Public Sub New(ByVal sourceNumber As ComplexNumber)
        Me.Real = sourceNumber.Real
        Me.Imaginary = sourceNumber.Imaginary
    End Sub
    
    Public Function GetAbs() As Double
        Return Math.Sqrt(Me.Real^2 + me.Imaginary^2)
    End Function
    
    Public Overrides Function ToString() As String
        Return Real & "+" & Imaginary & "i"
    End Function

    Public Shared Operator +(ByVal a As ComplexNumber, _
            ByVal b As ComplexNumber) As ComplexNumber
        Return New ComplexNumber(a.Real + b.Real, a.Imaginary + b.Imaginary)
    End Operator

    Public Shared Operator -(ByVal a As ComplexNumber, _
            ByVal b As ComplexNumber) As ComplexNumber
        Return New ComplexNumber(a.Real - b.Real, a.Imaginary - b.Imaginary)
    End Operator

    Public Shared Operator *(ByVal a As ComplexNumber, _
            ByVal b As ComplexNumber) As ComplexNumber
        Return New ComplexNumber(a.Real * b.Real - a.Imaginary * b.Imaginary, _
            a.Real * b.Imaginary + a.Imaginary * b.Real)
    End Operator

    Public Shared Operator /(ByVal a As ComplexNumber, _
            ByVal b As ComplexNumber) As ComplexNumber
        Return a * Reciprocal(b)
    End Operator
    
    Public Shared Function Reciprocal(ByVal a As ComplexNumber) As ComplexNumber
        Dim divisor As Double

        divisor = a.Real * a.Real + a.Imaginary * a.Imaginary
        If (divisor = 0.0#) Then Throw New DivideByZeroException

        Return New ComplexNumber(a.Real / divisor, -a.Imaginary / divisor)
    End Function
End Structure




