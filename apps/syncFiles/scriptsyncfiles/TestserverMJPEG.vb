#Imports System.Net
#Imports System.Text
#Imports System.Text.RegularExpressions
#Reference System.Management.dll
#Imports System.Management
#Imports System.Threading
#Imports System.Net.Sockets

Dim sock As Sockets.Socket

Dim remoteEP As IPEndPoint
Dim requestFilespec As String
Dim queryString As String
Dim recvData As String, recvLen As Integer

Dim stopListening As Boolean = False
Dim FormRef As Object

Dim serverName As String = "Server: screenshot-live-server/1.0"
Dim wwwRoot As String = "E:/testserver-www/"
Dim bindToIP As String = getIPAddress()'"192.168.178.27"
Dim bindToPort As Integer = 8080

Dim session As New System.Collections.Generic.Dictionary(Of String,String)

Const Auto=-2

Sub AutoStart()
  '' FormRef = ZZ.getDialogFormRef("TestserverMJPEG.debug")
  '' With FormRef
  ''   .checkForIllegalCrossThreadCalls=False
  ''   .resetControls (5,5)
  ''   .AddButton ("close", "   Free Socket   ")  :  .Br()
  ''   '.AddControl ("lst_debug", "scriptIDE.LstBox", ,,Auto,500)  :  .Br()
  ''   '.AddControl ("lst_ips", "scriptIDE.LstBox", ,,Auto,500)
  '' End With
  trace "Herzlich willkommen zum MJPEG-Testserver ..."
  sock = New Sockets.Socket(Sockets.AddressFamily.InterNetwork, Sockets.SocketType.Stream, Sockets.ProtocolType.IP)
  
  Dim th As New Threading.Thread(AddressOf listenerThread)
  th.Start()
End Sub

Sub onButtonEvent(e)
  If e.Sender.Name="btn_close" Then
    stopListening = True
  End If
End Sub

Sub listenerThread()
  'If FormRef Is Nothing Then FormRef = ZZ.getDockPanelRef("TestserverWWW.debug")
  Dim ep = New IPEndPoint(IPAddress.Parse(bindToIP), bindToPort)
  trace "start..."
  sock.Bind (ep)
  trace "bound to IP "+bindToIP+" at port "+bindToPort.Tostring+" ..."
  sock.Listen(20)
  trace "listening..."
  Dim acceptSock As Sockets.Socket
  Dim buf(2048) As Byte
  Dim headerBytes() As Byte
  Do
    acceptSock = sock.Accept()
    remoteEP = acceptSock.RemoteEndPoint
    Dim ip=remoteEP.Address.ToString
    trace "Accepted - "+ip+" - "+Now.ToString("yy-MM-dd HH:mm:ss")
    'If FormRef.Element("lst_ips").Items.Contains(ip) = False Then FormRef.Element("lst_ips").Items.Add(ip)
    recvData = ""
    requestFilespec = "" : queryString = ""
    Redim headerBytes(0)
    Do
      recvLen = acceptSock.Receive(buf, 2048, Sockets.SocketFlags.Partial)
      trace "recvLen",recvLen.ToString
      'If recvLen=0 Then Exit Do
      recvData += Encoding.ASCII.GetString(buf,0,recvlen)
    Loop Until acceptSock.Available=0 Or recvLen=0
    Dim lines()=Split(recvData,vbcrlf)
    trace "HTTP: ",lines(0)
    'trace lines(0)
    If lines(0).Startswith("GET ") OR lines(0).Startswith("POST ") Then
      Dim cmd()=Split(lines(0)," ")
      If cmd(1).Contains("?") Then queryString=cmd(1).Substring(cmd(1).IndexOf("?")):cmd(1)=cmd(1).Substring(0,cmd(1).Indexof("?"))
      If cmd(1)="/" Then cmd(1)="/index.html"
      requestFilespec = ZZ.FP(wwwRoot,cmd(1),True)
      
      If cmd(1)="/ScreenShot.jpg" Then
        Dim contentType="Content-Type: multipart/x-mixed-replace; boundary=--wg45asd98dfa034ga4"
        trace "200 OK - FileSpec: "+requestFilespec
        trace contentType
        Dim header = "HTTP/1.0 200 OK" + vbCrLf + contentType + vbCrLf + _
                     serverName + vbCrLf + vbCrLf
        
        trace "Screenshot ..."
        trace "Head:"+header
        headerBytes = Encoding.ASCII.GetBytes(header)
        acceptSock.Send(headerBytes)
        
        Dim del As New ParameterizedThreadStart(AddressOf motionStreamThread)
        Dim th As New Thread(del)
        th.Start(acceptSock)
        
        Continue Do 'nicht socket schließen!
      ElseIf cmd(1)="/index.html" Then
        Dim header = "HTTP/1.0 200 OK" + vbCrLf + serverName + vbCrLf + vbCrLf + _
                     "<html><body>" + _
                     "<br /><h2>mw Screenshot Server</h2>" + _
                     "<br />Herzlich willkommen!" + _
                     "<p><img src=""/ScreenShot.jpg""></p><br><br>...ich bin ein Screenshot</body></html>"
        
        headerBytes = Encoding.ASCII.GetBytes(header)
        acceptSock.Send(headerBytes)
      Else
        trace "404 Not Found"
        Dim header = "HTTP/1.0 404 Not Found" + vbCrLf + serverName + vbCrLf + vbCrLf + _
                     "<html><body>" + _
                     "<br /><h2>mw Screenshot Server</h2>" + _
                     "<p>Die Seite kann nicht angezeigt werden.</p></body></html>"
        
        headerBytes = Encoding.ASCII.GetBytes(header)
        acceptSock.Send(headerBytes)
      End If
    End If
    acceptSock.Close()
  Loop While stopListening = False
  trace "terminated"
  sock.Close
  
End Sub


Sub motionStreamThread(obj)
  Dim acceptSock As Socket=obj
:  While acceptSock.Connected And stopListening = False
:    trace "motionStream"
:    Dim arr() As Byte = getScreenshotAsByteArray()
:    acceptSock.Send(Encoding.ASCII.GetBytes(vbCrLf+"--wg45asd98dfa034ga4"+vbCrLf+ _
:                    "Content-Type: image/jpeg"+vbCrLf+"Content-Length: "+arr.Length.ToString+vbCrlf+vbCrlf))
:    acceptSock.Send(arr)
:    acceptSock.Send(Encoding.ASCII.GetBytes(vbCrLf+"--wg45asd98dfa034ga4"))
:    Threading.Thread.Sleep(500)
:  End While
End Sub

Sub onTerminate()
  trace "onTerminate"
  stopListening = True
  sock.Close
End Sub

Function getScreenshotAsByteArray() As Byte()
  Dim bounds As Rectangle
  Dim screenshot As System.Drawing.Bitmap
  Dim graph As Graphics
  bounds = Screen.PrimaryScreen.Bounds 
  'screenshot = New System.Drawing.Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb) 
  screenshot = New System.Drawing.Bitmap(400, 400, System.Drawing.Imaging.PixelFormat.Format32bppArgb) 
  graph = Graphics.FromImage(screenshot)
  Dim xPos = Math.max(0,Cursor.Position.X-200)
  Dim yPos = Math.max(0,Cursor.Position.Y-200)
  
  'graph.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy) 
  graph.CopyFromScreen(xPos, yPos, 0, 0, New Size(400,400), CopyPixelOperation.SourceCopy) 
  
  Dim ms As New System.Io.MemoryStream()
  screenshot.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
  
  getScreenshotAsByteArray = ms.GetBuffer()
  ms.Close() : graph.Dispose() : screenshot.Dispose()
End Function

'' Sub xtrace(item As String)
''   With FormRef.Element("lst_debug")
''     .items.Add(item)
''     Io.File.AppendAllText("E:\testserver-www\log.txt", item + vbNewLine)
''     .SelectedIndex = .Items.Count-1
''   End With
'' End Sub


Function getIPAddress() As String
  Dim mc As ManagementClass
  Dim mo As ManagementObject
  dim out as string=""
  mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
  Dim moc As ManagementObjectCollection = mc.GetInstances()
  For Each mo In moc
    If mo.Item("IPEnabled") = True Then
      Return mo.Item("IPAddress")(0)
      
    End If
  Next
End Function

