'--> CompilerOptions

#Para AssemblyType exe
#Para SilentMode True
#Para DebugMode External

#Imports System.Web
#Imports System.Net
#Imports System.Net.Sockets
#Imports System.Text
#Imports System.Text.RegularExpressions
#Reference System.Management.dll
#Reference System.Web.dll
#Imports System.Management
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Imports System.Collections.Specialized
#Imports System.Collections.Generic


'--> globVars

Public Shared serverName As String = "Server: testserver-www/1.0"
Public Shared wwwRoot As String = "E:/testserver-www/"
Public Shared bindToIP As String = getIPAddress()'"192.168.178.27"
Public Shared bindToPort As Integer = 8080

Public Shared sessionData As New Dictionary(Of String,SessionInfo)


Dim sock As Socket

Dim stopListening As Boolean = False
Dim FormRef As Object

Const Auto=-2

'-->

Sub AutoStart()
  '' FormRef = ZZ.getDialogFormRef("TestserverWWW.debug")
  '' With FormRef
  ''   .checkForIllegalCrossThreadCalls=False
  ''   .resetControls (5,5)
  ''   .AddButton ("close", "   Free Socket   ")
  ''   .Br()
  ''   .AddControl ("lst_debug", "scriptIDE.LstBox", ,,Auto,500)
  ''   .Br()
  ''   .AddControl ("lst_ips", "scriptIDE.LstBox", ,,Auto,500)
  '' End With
  
  sock = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP)
  
  Dim th As New Threading.Thread(AddressOf listenerThread)
  th.Start()
End Sub

Sub onButtonEvent(e)
  If e.Sender.Name="btn_close" Then
    stopListening = True
  End If
End Sub

Sub onTerminate()
  trace "onTerminate"
  stopListening = True
  sock.Close
End Sub

Sub listenerThread()
  If FormRef Is Nothing Then FormRef = ZZ.getDockPanelRef("TestserverWWW.debug")
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
    Dim resp As New ResponseGen(acceptSock)
  Loop While stopListening = False
  trace "terminated"
  sock.Close
End Sub


'-->
'--> Class SessionInfo
Class SessionInfo

Public ID As String
Public session As New System.Collections.Generic.Dictionary(Of String,String)

Sub New()
  Me.ID = Guid.NewGuid().ToString
End Sub

Function GetCookieHeader() As String
  Return "Set-Cookie: SessionID=" + Me.ID + "; expires="+Date.Now.AddYears(30).ToString("R")+"; path=/; domain=" + bindToIP
End Function

End Class

'-->
'--> Class ResponseGen
Class ResponseGen

Dim remoteEP As IPEndPoint
Dim requestFilespec As String
Dim queryString As String
Dim recvData As String, recvLen As Integer
Dim actSessionInfo As SessionInfo
Dim method As String
Dim actHeaders As New System.Collections.Generic.Dictionary(Of String,String)
Dim actCookies As New System.Collections.Generic.Dictionary(Of String,String)
Dim actInput As New System.Collections.Generic.Dictionary(Of String,String)
Dim actGET As NameValueCollection
Dim acceptSock As Socket

Sub New(Sock As Socket)
  acceptSock = Sock
  Dim th As New Threading.Thread(AddressOf ReadRequest)
  th.Start()
End Sub

Sub ReadRequest()
  remoteEP = acceptSock.RemoteEndPoint
  Dim ip=remoteEP.Address.ToString
  xtrace( "Accepted - "+ip+" - "+Now.ToString("yy-MM-dd HH:mm:ss"))
  
  recvData = ""
  requestFilespec = "" : queryString = ""
  Dim headerBytes(0) As Byte, buf(2048) As Byte
  Do
    recvLen = acceptSock.Receive(buf, 2048, Sockets.SocketFlags.Partial)
    trace "recvLen",recvLen.ToString
    recvData += Encoding.ASCII.GetString(buf,0,recvlen)
  Loop Until acceptSock.Available=0 Or recvLen=0
  Dim lines()=Split(recvData,vbcrlf)
  trace "HTTP: ",lines(0)
  xtrace(lines(0))
  ParseRequestHeaders(lines)
  
  If method <> "GET" And method <> "POST" Then sendErrorResult(405) : Exit Sub
  
  
  If IO.File.Exists(requestFilespec) Then
    Dim contentType=GetMimeTyp(requestFilespec)
    xtrace("200 OK - FileSpec: "+requestFilespec)
    xtrace(contentType)
    If contentType = "Content-Type: text/html" Then
      Dim fileCont=IO.File.ReadAllText(requestFilespec)
      replaceVars (fileCont)
      Dim header = "HTTP/1.0 200 OK" + vbCrLf + contentType + vbCrLf + _
                 "Content-Length: " + fileCont.Length.ToString + vbCrLf + serverName + vbCrLf + _
                 actSessionInfo.GetCookieHeader() + vbCrLf + vbCrLf
      headerBytes = Encoding.ASCII.GetBytes(header+fileCont)
      acceptSock.Send(headerBytes)
    Else
      Dim header = "HTTP/1.0 200 OK" + vbCrLf + contentType + vbCrLf + _
                 "Content-Length: " + FileLen(requestFilespec).Tostring + vbCrLf + serverName + vbCrLf + _
                 actSessionInfo.GetCookieHeader() + vbCrLf + vbCrLf
      headerBytes = Encoding.ASCII.GetBytes(header)
      acceptSock.SendFile(requestFilespec, headerBytes, New Byte(){}, 0)
    End If
  Else
    sendErrorResult(404)
  End If
  
  acceptSock.Close()
End Sub

Sub ParseRequestHeaders(lines() As String)
  Dim cmd()=Split(lines(0)," ")
  Redim Preserve cmd(3)
  method = cmd(0).ToUpper
  
  If cmd(1).Contains("?") Then queryString=cmd(1).Substring(cmd(1).IndexOf("?")):cmd(1)=cmd(1).Substring(0,cmd(1).Indexof("?"))
  If cmd(1)="/" Then cmd(1)="/index.html"
  requestFilespec = ZZ.FP(wwwRoot,cmd(1),True)
  
  actHeaders.Clear()
  For Each line As String In lines
    Dim header() = Split(line, ":", 2)
    If header.Length <> 2 Then Continue For
    actHeaders.Add(header(0).Trim.ToLower, header(1).Trim)
  Next
  
  actCookies.Clear()
  If actHeaders.ContainsKey("cookie") Then
    ParseCookies(actHeaders("cookie"))
  End If
  
  actInput.Clear()
  If queryString <> "" Then
    actGET = HttpUtility.ParseQueryString(querystring)
  End If
  
  actSessionInfo = GetSession()
End Sub

Sub ParseCookies(line As String)
  Dim cookies() = Split(line, ";")
  For Each cookie As String In cookies
    Dim parts() = Split(cookie, "=", 2)
    If parts.Length<>2 Then Continue For
    'trace "Cookie "+parts(0),parts(1)
    actCookies.Add(parts(0).Trim.ToLower, parts(1).Trim)
  Next
End Sub

Function GetSession() As SessionInfo
  If actCookies.ContainsKey("SessionID") AndAlso sessionData.ContainsKey(actCookies("SessionID")) Then
    Return sessionData(actCookies("SessionID"))
  Else
    Dim sess As New SessionInfo
    sessionData.Add(sess.ID, sess)
    Return sess
  End If
End Function




Sub xtrace(item As String)
  trace "X",item
End Sub


Sub replaceVars(ByRef str As String)
  str = Regex.Replace(str, "\<\#\#(.*?)\#\#\>", AddressOf replaceVarsMatchEvaluator)
  'ZZ.setOutMonitor(str)
  str = Regex.Replace(str, "222START_COMMENT333(.*?)222END_COMMENT333", "", RegexOptions.Singleline)
  str = Replace(str, "222END_COMMENT333", "")
End Sub


Function GetMimeTyp(fileSpec As String) As String
  Dim ext = IO.Path.GetExtension(fileSpec).toUpper()
  If ext = ".JPG" Then Return "Content-Type: image/jpeg"
  If ext = ".GIF" Then Return "Content-Type: image/gif"
  If ext = ".PNG" Then Return "Content-Type: image/png"
  If ext = ".HTML" OR ext = ".XHTML" Then Return "Content-Type: text/html"
  If ext = ".CSS" Then Return "Content-Type: text/css"
  If ext = ".TXT" OR ext = ".CSV" Then Return "Content-Type: text/plain"
  If ext = ".EXE" OR ext = ".DLL" Then Return "Content-Type: application/octet-stream"
  If ext = ".JS" Then Return "Content-Type: application/javascript"
  
  Return "Content-Type: text/plain" '"Content-Type: application/octet-stream"
End Function


Function getInternalServerVar(varName As String) As String
  If VarName = Nothing Then Return Nothing
  Select Case VarName
    Case "SERVER_ADDR" : Return bindToIP
    Case "SERVER_PORT" : Return bindToPort.Tostring
    Case "REMOTE_ADDR" : Return remoteEP.Address.ToString
    Case "REMOTE_PORT" : Return remoteEP.Port.ToString
    
    Case "DATE" : Return Now.ToLongDateString()
    Case "TIME" : Return Now.ToLongTimeString()
    Case "DAT_INVERS" : Return Now.ToString("yyyy-MM-dd-HH-mm-ss")
    
    Case "FILESPEC" : Return requestFilespec
    Case "REQUEST_URI" : Return "http://" + bindToIP + ":" + bindToPort.ToString + "/" + requestFilespec.Replace(wwwRoot, "")
    Case "QUERY_STRING" : Return queryString
    
    Case "RECEIVED_DATA" : Return recvData
    Case "RECEIVED_LENGTH" : Return recvLen.Tostring
  End Select
  
  Return Nothing
End Function


Function getInputVar(varName As String) As String
  If actCookies.ContainsKey(varName.ToLower) Then
    Return actCookies(varName.ToLower)
  End If
  Return actGET(varName.ToLower)
End Function


Function getServerVar(varName As String) As String
  If VarName = Nothing Then Return Nothing
  
  If VarName.StartsWith("#") Then Return getInternalServerVar(VarName.Substring(1))
  If VarName.StartsWith("%[") Then Return getInputVar(VarName.Substring(2, VarName.Length-3))
  
  If VarName.StartsWith("$") Then
    Dim key=varName.Substring(1).Trim
    If actSessionInfo.session.ContainsKey(key) Then Return actSessionInfo.session(key)
  End If
  Return Nothing
End Function

Function parseStringMatchEvaluator(ByVal match As System.Text.RegularExpressions.Match) As String
  Return getServerVar(match.Groups(1).Value)
End Function

Function parseString(str As String) As String
  str = Regex.Replace(str, "([\$#][_a-zA-Z0-9]+)", AddressOf parseStringMatchEvaluator)
  str = Regex.Replace(str, "{([\$#][^}]+)}", AddressOf parseStringMatchEvaluator)
  str = Regex.Replace(str, "(%\[[^\]]+\])", AddressOf parseStringMatchEvaluator)
  
  Return str
End Function

Function replaceVarsMatchEvaluator(ByVal match As System.Text.RegularExpressions.Match) As String
  'xtrace("ServerVar: "+match.Value)
  Dim VarName = match.Value.ToUpper
  'Dim VarCont=getServerVar(VarName)
  'If VarCont <> Nothing Then Return VarCont
  If VarName = "<## START_COMMENT ##>" Then Return "222START_COMMENT333"
  If VarName = "<## END_COMMENT ##>" Or VarName = "<## END_IF ##>" Then Return "222END_COMMENT333"
  
  If VarName.StartsWith("<##=""") And VarName.EndsWith("""##>") Then
    Dim Cont As String = match.Value.Substring(5, match.Value.Length-9)
    Return parseString(Cont)
  End If
  If VarName.StartsWith("<## IF ") Then
    Dim m=Regex.Match(VarName, """([^""]*)""[\s]*([a-zA-Z]{2,4})[\s]*""([^""]*)""[\s]*(SET[\s]*""([^""]*)""[\s]*""([^""]*)"")?")
    Dim varName2 As String = m.Groups(1).Value
    Dim strOperator As String = m.Groups(2).Value.ToUpper
    Dim expectedValue As String = m.Groups(3).Value.ToUpper
    Dim varCont2 = parseString(varName2)
    If varCont2 = Nothing Then varCont2=""
    varCont2=varCont2.ToUpper
    Dim bResult As Boolean
    Select Case strOperator
      Case "CNT" : bResult = varCont2.Contains(expectedValue)
      Case "NCNT" : bResult = Not varCont2.Contains(expectedValue)
      Case "SW" : bResult = varCont2.StartsWith(expectedValue)
      Case "NSW" : bResult = Not varCont2.StartsWith(expectedValue)
      Case "EW" : bResult = varCont2.EndsWith(expectedValue)
      Case "NEW" : bResult = Not varCont2.EndsWith(expectedValue)
      Case "EQ" : bResult = varCont2=expectedValue
      Case "NEQ" : bResult = Not varCont2=expectedValue
      Case Else : Return "Invalid If Operator"
    End Select
    
    Dim setKey=m.Groups(5).Value.ToUpper
    If setKey<>"" Then
      If bResult = False Then Return "<!-- didn't change  session key "+setKey+" -->"
      Dim newSetValue=parseString(m.Groups(6).Value)
      'If newSetValue.StartsWith("#") Then newSetValue = getServerVar("<## "+newSetValue.Substring(1).ToUpper+" ##>")
      If actSessionInfo.session.ContainsKey(setKey) Then
        actSessionInfo.session(setKey)=newSetValue
      Else
        actSessionInfo.session.Add(setKey,newSetValue)
      End If
      Return "<!-- set session key "+setKey+" to """+newSetValue+""" -->"
      'ENDE
    End If
    
    'xtrace("<## "+varName2+" ##>="+varCont2 + "~~~["+strOperator+"]~~~"+expectedValue+"<    bResult=" +bResult.Tostring+"   gCount="+m.Groups.Count.TOstring+"   "+m.Groups(5).Value)
    
    
    
    If bResult = True Then
      Return ""
    Else
      Return "222START_COMMENT333"
    End If
  End If
  If VarName.StartsWith("<## INCLUDE """) Then
    Dim abPos=VarName.IndexOf("""")
    Dim bisPos=VarName.LastIndexOf("""")
    If abPos=-1 OR abPos=bisPos OR bisPos=-1 Then Return "Illegal Include Statement"
    Dim includeFile=parseString(VarName.Substring(abPos+1, bisPos-abPos-1))
    includeFile=IO.Path.Combine(IO.Path.GetDirectoryName(requestFilespec), includeFile)
    'xtrace("Include File: "+includeFile)
    If IO.File.Exists(includeFile)=False Then Return "<b>Warning:</b> Include File not found: "+includeFile+"<br>"
    Dim incText= IO.File.ReadAllText(includeFile)
    replaceVars(incText)
    return incText
  End If
  Return ""
End Function

Sub sendErrorResult(errCode As Integer)
  Select Case errCode
    Case 404
      xtrace("404 Not Found")
      Dim header = "HTTP/1.0 404 Not Found" + vbCrLf + serverName + vbCrLf + actSessionInfo.GetCookieHeader() + vbCrLf + vbCrLf + _
                   "<html><body><img src='/warning.png' alt='Fehler' style='float: left; padding: 4px 20px 5px 100px;' width='128' height='128' />" + _
                   "<br /><h2>404: Not found</h2><p>The requested URL <i>"+requestFilespec+"</i> was not found on this server. </p>" + _
                   "<p><a href=""/"">Go back to the Home Page.</A></p></body></html>"
      
      acceptSock.Send(Encoding.ASCII.GetBytes(header))
    Case 405
      xtrace("405 Method Not Allowed")
      Dim header = "HTTP/1.0 405 Method Not Allowed" + vbCrLf + serverName + vbCrLf + vbCrLf + _
                   "<html><body><img src='/warning.png' alt='Fehler' style='float: left; padding: 4px 20px 5px 100px;' width='128' height='128' />" + _
                   "<br /><h2>405 Method Not Allowed</h2><p></p></body></html>"
      
      acceptSock.Send(Encoding.ASCII.GetBytes(header))
  End Select
End Sub

End Class

Public Shared Function getIPAddress() As String
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

