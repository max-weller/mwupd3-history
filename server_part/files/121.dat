'-->

#Para CompilerOptions /debug+ /debug:pdbonly
#Para IconFile C:\yPara\globIcons\wmploc_612.ico
#Para SilentMode True
#Para DebugMode external

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll
#Reference siaIDEMain.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib
#Imports System.Collections.Generic
#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic
#Reference TenTec.Windows.iGridLib.iGrid.dll
#Imports Tentec.Windows.iGridLib
#Reference c:\yexe\AxInterop.AXVLC.dll
#Reference c:\yexe\Interop.AXVLC.dll
#Imports System.Runtime.InteropServices
#Imports System.Text.RegularExpressions
#Reference WeifenLuo.WinFormsUI.Docking.dll

#runtime EXE

'--> Window "downloadStatus"
#WindowData frm_downloadStatus
			350	240	Form		Width=350|Height=240|Text="Download ..."|TopMost=True|FormBorderStyle=3'Dialog|MaximizeBox=False|MinimizeBox=False|ControlBox=False
	70	15	250	30	Label	labInfo	Text="Video wird heruntergeladen, bitte warten ..."
	10	10	48	48	Picturebox	picIcon	Tag="#FBABAB"|Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/oobefldr_104-48.png")
	70	45	250	90	TextBox	txtURL	Text="URL"|MultiLine=True
	70	145	250	30	ProgressBar	pb1	Minimum=0|Maximum=100|Style=1'continous
	210	180	110	25	Button	btnCancel	Text="Abbrechen"|Enabled=false
#EndData

Dim WithEvents FormRef As Form
Dim PanelRef As ScriptedPanel
Dim toolbar As ScriptedToolstrip
Dim TBRef As Object
Dim WithEvents IGrid1 As Igrid
Dim WithEvents pbVolume As ProgressBar
Dim WithEvents pbPos As ProgressBar
Dim WithEvents VLC1 As AxAXVLC.AxVLCPlugin2
Dim WithEvents tmr1 As Timer
Dim glob As cls_globPara

Public Const winID = "mw_mp3player.PlayList"
Public Const Auto = -2

Sub GetFormRef()
  If PanelRef IsNot Nothing Then Exit Sub
  'trace winID
  PanelRef = ZZ.CreateWindow(winID, DWndFlags.StdWindow Or DWndFlags.NoAutoShow)
  FormRef = PanelRef.Form
End Sub

Sub PlayList_Load(e)
 ' msgbox("load")
  FormRef = e.Sender
  FormRef.Text = "Music Box"
  glob.readFormPos(FormRef)
End Sub

Sub AutoStart()
  Application.EnableVisualStyles()
  glob = ZZ.newGlobPara()
  GetFormRef()
  
  Dim pTop, pLeft, pRight As ScriptedPanel
  Dim el As Control
  With PanelRef
    .resetControls (5,5)
    .activateEvents = "|ButtonMouseClick| |TEXTBOXTEXTCHANGED|"
    
    Dim spl=.addSplitcontainer("splMain", "left", pLeft, "right", pRight, DockStyle.Fill)
    spl.FixedPanel = 1
    
    pTop = .addPanel("topBar", DockStyle.Top, intHeight:=35)
    pTop.resetControls (5,5) : ZZ.setBgColor(pTop, "#fea")
    toolbar = pTop.addToolstrip("main", 5, 5, 50, 20)
    toolbar.className="{ScriptClass}"
    toolbar.Dock=0
    toolbar.ImageScalingSize = New Size(16,16)
    'toolbar.addButton("cmdMenu", "", flags:=ButtonFlags.StartMenu, iconURL:="http://www.iconfinder.net/data/icons/nuvola2/16x16/apps/kmenuedit.png")
    toolbar.addButton("cmdPlay", "Play", iconURL:="http://www.iconfinder.net/data/icons/humano2/16x16/actions/gtk-media-play-ltr.png")
    toolbar.addButton("cmdStop", "Stop", iconURL:="http://www.iconfinder.net/data/icons/humano2/16x16/actions/gtk-media-pause.png")
    'toolbar.addButton("cmdRead", "Read")
    toolbar.addButton("cmdIndexDirectory", "", iconURL:="http://mw.teamwiki.net/docs/img/win-icons/wmploc_719-16.png")
    toolbar.addButton("cmdFolderlist", "", iconURL:="http://mw.teamwiki.net/docs/img/win-icons/shell32_173-16.png")
    toolbar.addButton("cmdDownloadMulti", "", iconURL:="http://mw.teamwiki.net/docs/img/win-icons/fdm_128-16.png")
    
    'toolbar.ActiveMenu = Nothing
    pTop.addIcon("icoSound", "http://www.iconfinder.net/data/icons/nuvola2/22x22/apps/kmix.png", intLeft:=200)
    
    pTop.insertY=10
    pTop.addIcon("cmdDownloadVideo", "http://mw.teamwiki.net/docs/img/win-icons/ieframe_1_113-16.png", intLeft:=375)
    pTop.addIcon("cmdNav", "http://mw.teamwiki.net/docs/img/win-icons/ECMangen_189-17.png", intLeft:=400)
    pTop.addTextbox("url", Auto, X:=420)
    
    pbVolume = New ProgressBar : pbVolume.Maximum = 200 : pbVolume.Value = 50
    pbVolume.Top = 5 : pbVolume.left = 230 : pbVolume.Width = 100 ': pbVolume.Height = 240
    pbVolume.Style=1
    pTop.Controls.Add(pbVolume)
    
    pLeft.resetControls (5,5)
    pRight.resetControls (5,5)
    
    '--> ... Left Pane
    
    pLeft.addIcon("cmdSuchIcon", "http://www.iconfinder.net/data/icons/icojoy/shadow/standart/gif/24x24/001_38.gif")
    pLeft.addTextbox("suchBox", 150) : pLeft.BR
    
    IGrid1 = New IGrid
    IGrid1.DefaultCol.CellStyle.EmptyStringAs = TenTec.Windows.iGridLib.iGEmptyStringAs.EmptyString 'zack!
    IGrid1.Top = 40 : IGrid1.Width = pLeft.Width : IGrid1.Height = .Height - 90: 
    IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    IGrid1.ReadOnly = True
    'leere Zelle als "" statt nothing
    IGrid1.Cols.Add("fileSpec",0)
    IGrid1.Cols.Add("Playlist",20)
    IGrid1.Cols.Add("Titel",180)
    IGrid1.Cols.Add("Interpret")
    IGrid1.Cols.Add("Tags")
    IGrid1.Cols.Add("LastMod")
    
    Igrid1.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
    
    'IGrid1.GroupBox.Visible = True
    pLeft.Controls.Add(IGrid1)
    
    
    '--> ... Right Pane
    
    VLC1 = New AxAXVLC.AxVLCPlugin2
    VLC1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    pRight.Controls.Add(VLC1)
    VLC1.Top = 5 : VLC1.Height = pRight.Height-45 : VLC1.Width = pRight.Width - 5
    VLC1.volume = 50
    
    pbPos = New ProgressBar
    pbPos.Top = pRight.Height-35 : pbPos.Width = pRight.Width - 10 ': pbPos.Height = 240
    pbPos.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
    pRight.Controls.Add(pbPos)
    pbPos.Maximum = 1000
    
    tmr1 = New Timer
    tmr1.Interval = 300 : tmr1.Enabled = True
  End With
  
  FormRef.Show()
  
  cmdRead_MouseClick(nothing)
  
  'onResizeControls()
End Sub


'-->
'--> Command Buttons

Sub cmdDownloadVideo_MouseClick(e)
  Dim url = PanelRef.Element("url").Text
  Dim rLocFile As String
  DownloadVideo(url, "D:\Videos\mwVideoDownloader", rLocFile)
  If Not String.IsNullOrEmpty(rLocFile) Then
    PanelRef.Element("url").Text = rLocFile
    playURL(rLocFile)
  End If
End Sub

Sub cmdDownloadMulti_MouseClick(e)
  showMultiDownloader()
End Sub

Sub cmdNav_MouseClick(e)
  Dim url = PanelRef.Element("url").Text
    If e.MouseButton = "Left" Then
    If url.Contains("youtube.com/") Then
      url = GetYTVideoURL(url)
    End If
    playURL(url)
  Else
    VLC1.MRL = url
    
  End If
End Sub


Sub cmdIndexDirectory_MouseClick(e)
  '' Using fbd As New FolderBrowserDialog()
  ''   'fbd.SelectedPath="D:\videos"
  ''   If fbd.ShowDialog() = DialogResult.OK Then
  ''     startDriveIndex(fbd.SelectedPath)
  ''     
  ''   End If
  '' End Using
  Igrid1.Rows.Clear()
  
  Dim fileCont = ZZ.fileGetContents("C:\yPara\scriptIDE\scriptPara\musicBoxIndexDirs.txt")
  Dim Folders() = Split(fileCont, vbNewLine)
  For Each folder As String In Folders
    If IO.Directory.Exists(folder) = False Then Continue For
    startDriveIndex(folder)
    
  Next
End Sub

Sub cmdSuchIcon_MouseClick(e)
  PanelRef.Element("suchBox").Focus()
  PanelRef.Element("suchBox").SelectAll()
  
End Sub

Sub cmdFolderlist_MouseClick(e)
  Dim fileSpec As String = "C:\yPara\scriptIDE\scriptPara\musicBoxIndexDirs.txt"
  If IO.File.Exists(fileSpec) = False Then ZZ.filePutContents(fileSpec, "")
  
  Diagnostics.Process.Start(fileSpec)
  
End Sub

Sub cmdRead_MouseClick(e)
  Dim Content = ZZ.fileGetContents("C:\yPara\scriptIDE\mp3_idx.txt")
  Igrid_put (IGrid1, Content)
End Sub


Sub cmdPause_MouseClick(e)
  VLC1.playlist.togglePause()
End Sub


Sub cmdStop_MouseClick(e)
  If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  Application.DoEvents() : Threading.Thread.Sleep(22) : Application.DoEvents()
  If VLC1.playlist.items.count>0 Then VLC1.playlist.items.clear()
  Application.DoEvents()
  
End Sub

Sub cmdPlay_MouseClick(e)
  If IGrid1.CurRow Is Nothing Then Exit Sub
  
  Dim fileSpec As String = IGrid1.CurRow.Cells(0).Value
  PanelRef.Element("url").Text = fileSpec
  playURL(fileSpec)
End Sub

Sub playURL(fileSpec As String)
  printLine 2,"fileSpec",fileSpec
  'mms://ndr-fs-nds-hi-bkp-wmv.wm.llnwd.net/ndr_fs_nds_hi_bkp_wmv?MSWMExt=.asf
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  Application.DoEvents()
  'If VLC1.playlist.items.count>0 Then VLC1.playlist.items.clear()
  Application.DoEvents()
  Dim itemID = VLC1.playlist.add(fileSpec, "")
  'VLC1.MRL = "D:\Videos\mwVideoDownloader\BassHunter_-_I_can_walk_on_water__I_can_fly.flv"
  VLC1.playlist.playItem(itemID)
End Sub


Sub cmdMenu(e)
End Sub 'dummyFunc

'-->

'--> Window "downloadMulti"
Dim dmForm As frm_downloadMulti

#WindowData frm_downloadMulti
			450	360	Form		Width=450|Height=360|Text="Multi-Downloader"|TopMost=True|FormBorderStyle=3'Dialog|MaximizeBox=False
	14	20	48	48	Picturebox	dm_picIcon	Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/wuweb_IDI_MUICON-48.png")
	70	15	350	220	CheckedListBox	dm_list	Text="URL"
	69	242	66	20	Label	dm_lab01	Text="Zielordner:"
	140	240	280	20	TextBox	dm_txtTargetFolder	Text="D:\\Videos\\mwVideoDownloader\\"
	70	270	350	20	ProgressBar	dm_pb	Minimum=0|Maximum=100|Style=1'continous
	70	300	110	25	Button	dm_btnInsert	Text="Einfügen"
	200	300	110	25	Button	dm_btnStart	Text="Start"
	310	300	110	25	Button	bm_btnCancel	Text="Abbrechen"
#EndData

Sub showMultiDownloader()
  If dmForm Is Nothing OrElse dmForm.IsDisposed Then
    dmForm = New frm_downloadMulti(me)
  End If
  dmForm.Show() : dmForm.Activate()
End Sub

Sub dm_btnInsert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
  Dim clip = Clipboard.GetText()
  Dim lines() = Split(clip, vbNewLine)
  For Each line As String in lines
    dmForm.dm_list.Items.Add(line, line.StartsWith("http://"))
  Next
End Sub

Sub dm_btnStart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
  While dmForm.dm_list.CheckedIndices.Count > 0
    Dim idx As Integer = dmForm.dm_list.CheckedIndices(0)
    Dim url As String = dmForm.dm_list.Items(idx)
    dmForm.dm_list.SelectedIndex = idx
    dmForm.Text = "[" & idx & "] Multi-Downloader"
    If url.StartsWith("http://") Then
      'download ...
      Dim rLocFile As String
      DownloadVideo(url, dmForm.dm_txtTargetFolder.Text, rLocFile)
      If String.IsNullOrEmpty(rLocFile) Then
        dmForm.dm_list.Items.Insert(idx+1, "...ERR: Fehler beim Downloaden!")
        Beep()
      End If
      
    End If
    Application.DoEvents()
    dmForm.dm_list.SetItemChecked(idx, False)
    
    Application.DoEvents()
  End While
  
End Sub

Sub bm_btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
  dmForm.Close()
  dmForm = Nothing
End Sub




'-->
'--> Events

:Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr1.Tick
  On Error Resume Next
  If VLC1.playlist.isPlaying Then
'    pbPos.Maximum = VLC1.input.Length
    pbPos.Value = VLC1.input.position*1000
  End If
End Sub

Sub pbPos_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbPos.MouseDown
  VLC1.input.position = e.X / pbPos.Width
  
End Sub

Sub pbVolume_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pbVolume.MouseDown
  pbVolume.Value = e.X / pbVolume.Width * pbVolume.Maximum
  VLC1.volume = pbVolume.Value
  
End Sub

Sub IGrid1_CellMouseUp(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseUpEventArgs) Handles IGrid1.CellMouseUp
  If e.RowIndex < 0 Then Exit Sub
  If e.Button = MouseButtons.Right Then
    IGrid1.SetCurRow(e.RowIndex)
    cmdPlay_MouseClick(nothing)
  End If
End Sub

Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  glob.saveFormPos(FormRef)
  glob.saveParaFile()
  
End Sub

Sub onTerminate()
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
  
End Sub

'' Sub onMenuEvent(e)
''   If e.EventName = "ItemClicked" Then
''     GetFormRef()
''     Dim btnName = "cmd"+e.Sender.Text
''     trace btnName,"menu-clicked"
''     CallByName(Me, btnName, Microsoft.VisualBasic.CallType.Method, e)
''     If Err.Number Then MsgBox("Funktion nicht gefunden"+vbnewline+"..."+btnName):Err.Clear
''     
''   End If
'' End Sub
'' 
'' Sub onButtonEvent(e)
''   GetFormRef()
''   Dim btnName = e.Sender.Name.Substring(4)
''   
'' : CallByName(Me, btnName, Microsoft.VisualBasic.CallType.Method, e)
''   If Err.Number Then MsgBox("Funktion nicht gefunden"+vbnewline+"..."+btnName):Err.Clear
'' End Sub

Sub suchBox_TextChanged(e)
  GetFormRef()
  Dim suchWort = e.Sender.Text.Tolower
: For i As Integer = 0 To IGrid1.Rows.Count - 1
:   Igrid1.Rows(i).Visible = Igrid1.Cells(i,0).Value.ToLower.Contains(suchWort)
: Next
End Sub



'-->


Sub onFileFinderStatusEvent(FolderCount, FileCount, ActFolder, ActFileSpec, ByRef Cancel)
  PanelRef.Element("url").Text = ActFolder
End Sub


Sub startDriveIndex(startDir)
  
  PanelRef.Element("url").Text = "Suche läuft  " + startDir
  
  ZZ.TimerStart
  Dim finder As Object = ZZ.OpenFileFinder(startDir, "*.*")
  finder.CallbackScriptClassFunc = "onFileFinderStatusEvent"
  
  Dim RESULT: RESULT=finder.FindRecursive()
  
  PanelRef.Element("url").Text = Finder.FoundDirectoriesCount & " / " & Finder.FoundFilesCount & "   Benötigte Zeit: " & Finder.ElapsedTime & " ms" & ZZ.TimerGetElapsed() & " ms"
  
  Dim allowedExt = "|.MP3|.MP4|.FLV|"
  Dim Lines()=Split(RESULT,vbNewLine)
  Dim Parts() As String
  For i As Integer = 0 To Lines.Length - 1
    If Lines(i)="" Then Continue For
    Parts = Split(Lines(i), vbTab)
    If allowedExt.Contains("|"+Parts(2).ToUpper+"|") = False Then Continue For
    
    Dim iR = Igrid1.Rows.Add()
    iR.Cells(0).Value = Parts(7) 'fileSpec
    iR.Cells(4).Value = Parts(4) 'lastMod
    
    Dim fileName = Replace(Parts(1), "_", " ")
    If fileName.Contains("-") Then
      Dim abPos = fileName.IndexOf("-")
      iR.Cells(2).Value = fileName.Substring(abPos+1).Trim
      iR.Cells(3).Value = fileName.Substring(0, abPos).Trim
      
    Else
      iR.Cells(2).Value = fileName
    End If
    
    
    
  Next
  
  Dim Content As String = ""
  printline 1,"ig",igrid1.tostring
  Igrid_get(IGrid1, Content)
  
  ZZ.filePutContents("C:\yPara\scriptIDE\mp3_idx.txt", Content)
  PanelRef.Element("url").Text = ""
End Sub


'' Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
''   onResizeControls()
'' End Sub
'' 
'' Sub onResizeControls()
''   Igrid1.Width = formRef.Width - 10
''   Igrid1.Height = formRef.Height - 100
''   
'' End Sub
'' 



'-->
'--> Video Download

Public Const ifoVideoID = 0
Public Const ifoVideoTitle = 1
Public Const ifoT = 2
Dim frmDlStatus As frm_downloadStatus


Private Sub downloader_OnDownloadProgress(file As String, ReceivedBytes As Long, TotalBytes As Long)
  frmDlStatus.pb1.Value = ReceivedBytes / TotalBytes * 100
  'FormRef.Text="Downloading...   " & cint(ReceivedBytes / TotalBytes * 100) & "%     (" & file & ")"
  :Application.DoEvents : Err.Clear()
End Sub

Sub DownloadVideo(url As String, targetFolder As String, byRef resultFile As String)
  frmDlStatus = new frm_downloadStatus(Me)
  frmDlStatus.Show()
  frmDlStatus.txtURL.Text = "Quelle:"+url
  Dim ifoVideoID, ifoVideoTitle, ifoVideoT As String
  GetVideoInfo(url, ifoVideoID, ifoVideoTitle, ifoVideoT)
  
  frmDlStatus.txtURL.Text = ifoVideoTitle+vbNewLine+"Quelle:"+url+vbNewLine
  trace "ifoVideoID", ifoVideoID
  trace "ifoVideoTitle", ifoVideoTitle
  trace "ifoT", ifoVideoT
  
  Dim dlUrl = "http://www.youtube.com/get_video?video_id=" + ifoVideoID + "&t=" + ifoVideoT
  Dim locFile = ZZ.FP(targetFolder) + Makefilename(ifoVideoTitle) + "." + ifoVideoID + ".flv"
  
  frmDlStatus.txtURL.Text = ifoVideoTitle+vbNewLine+"Quelle:"+dlUrl+vbNewLine+"-->"+locFile
  
  trace "downloadURL", dlUrl
  trace "localFile", locFile
  
  
  Using downloader As New Downloader()
    downloader.SourcePath = dlUrl
    downloader.LocalPath = locFile
    AddHandler downloader.OnDownloadProgress, AddressOf downloader_OnDownloadProgress
    
    downloader.Download()
  End Using
  If IO.File.Exists(locFile) Then resultFile = locFile
  frmDlStatus.Close
  frmDlStatus=Nothing
End Sub

Function GetYTVideoURL(url As String) As String
  Dim ifoVideoID, ifoVideoTitle, ifoVideoT As String
  GetVideoInfo(url, ifoVideoID, ifoVideoTitle, ifoVideoT)
  
  Return "http://www.youtube.com/get_video?video_id=" + ifoVideoID + "&t=" + ifoVideoT
  
End Function

Function GetVideoInfo(url As String, ByRef VideoID As String, ByRef VideoTitle As String, ByRef VideoT As String) As String()
  Dim info(10) As String
  
  :Dim urlCont = ZZ.http_get(url) : Err.Clear()
  trace "CONT",urlCont
  
  Dim res = Regex.Matches(urlCont, "'([A-Z_]+)': (.*)$", RegexOptions.Multiline)
  For Each m As Match in res
    Dim key = m.Groups(1).value
    If key = "VIDEO_ID" Then VideoID = Trimpara(m.groups(2).value)
    If key = "VIDEO_TITLE" Then VideoTitle = Trimpara(m.groups(2).value)
    If key = "SWF_ARGS" Then 
      VideoT = GetValueFromSwfArgs(m.groups(2).value, "t")
    End If
  Next
  
  Return info
End Function

Function GetValueFromSwfArgs(args As String, keyName As String) As String
  Dim rx = """" + keyName + """: ""([^""]*)"""
  trace "regex",rx
  Dim mat = Regex.Match(args, rx)
  If mat Is Nothing Then Return ""
  Return mat.Groups(1).Value
End Function

Function Trimpara(args As String) As String
  If String.IsNullOrEmpty(args) Then Return ""
  Return args.Trim(" "c, ","c, "'"c, """"c)
End Function

Function Makefilename(title As String) As String
  Return Regex.Replace(title, "[^a-zA-Z0-9_-]", "_")
End Function

'-->

  Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
    Dim max = line.Cells.Count - 1
    Dim out(max) As String
    For i As Integer = 0 To max
      If line.Cells(i).Value IsNot Nothing Then _
        out(i) = line.Cells(i).Value.ToString
    Next
    Return Join(out, Delimiter)
  End Function

  Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
:    Dim max = line.Cells.Count - 1
:    Dim out() = Split(input, Delimiter)
:    ReDim Preserve out(max)
:    For i As Integer = 0 To max
:      line.Cells(i).Value = out(i)
:    Next
  End Sub

  Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
    Dim max = Grid.Rows.Count - 1
    Dim out(max) As String
    For i As Integer = 0 To max
      out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
    Next
    FileContent = Join(out, LineDelim)
  End Sub
  
  Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
    Dim out() = Split(FileContent, LineDelim)
    Grid.Rows.Clear()
    Grid.Rows.Count = out.Length
:    For i As Integer = 0 To out.Length - 1
:      SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
:    Next
  End Sub
  


'-->
'--> URLDownloadToFile
:BEGIN

#Imports System.Runtime.InteropServices
#Imports System.Runtime.CompilerServices
#Imports System.IO
#Imports System.ComponentModel
'#Imports System.Threading

	'--> Interfaces


	<ComImport, Guid("79EAC9C1-BAF9-11CE-8C82-00AA004BA90B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
	Public Interface IBindStatusCallback
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnStartBinding(<[In]> dwReserved As UInteger, <[In], MarshalAs(UnmanagedType.[Interface])> pib As IBinding)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub GetPriority(ByRef pnPriority As Integer)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnLowResource(<[In]> reserved As UInteger)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnProgress(<[In]> ulProgress As UInteger, <[In]> ulProgressMax As UInteger, <[In]> ulStatusCode As BINDSTATUS, <[In], MarshalAs(UnmanagedType.LPWStr)> szStatusText As String)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnStopBinding(<[In], MarshalAs(UnmanagedType.[Error])> hresult As UInteger, <[In], MarshalAs(UnmanagedType.LPWStr)> szError As String)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub GetBindInfo(ByRef grfBINDF As BINDF, <[In], Out> ByRef pbindinfo As BINDINFO)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnDataAvailable(<[In]> grfBSCF As UInteger, <[In]> dwSize As UInteger, <[In]> ByRef pFormatetc As FORMATETC, <[In]> ByRef pStgmed As STGMEDIUM)

		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub OnObjectAvailable(<[In]> ByRef riid As Guid, <[In], MarshalAs(UnmanagedType.IUnknown)> punk As Object)
	End Interface

    <ComImport, Guid("79EAC9C0-BAF9-11CE-8C82-00AA004BA90B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
	Public Interface IBinding
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub Abort()
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub Suspend()
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub [Resume]()
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub SetPriority(<[In]> nPriority As Integer)
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub GetPriority(ByRef pnPriority As Integer)
		<MethodImpl(MethodImplOptions.InternalCall, MethodCodeType := MethodCodeType.Runtime)> _
		Sub GetBindResult(ByRef pclsidProtocol As Guid, ByRef pdwResult As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByRef pszResult As String, <[In], Out> ByRef dwReserved As UInteger)
	End Interface



	'--> User-Defined Types:

	<Flags> _
	Public Enum BINDVERB As UInteger
		BINDVERB_GET = &H0
		' default action
		BINDVERB_POST = &H1
		' post verb
		BINDVERB_PUT = &H2
		' put verb
		BINDVERB_CUSTOM = &H3
		' custom verb
	End Enum

	' flags that describe the type of transaction that caller wants
	<Flags> _
	Public Enum BINDF As UInteger
		BINDF_DEFAULT = &H0
		BINDF_ASYNCHRONOUS = &H1
		BINDF_ASYNCSTORAGE = &H2
		BINDF_NOPROGRESSIVERENDERING = &H4
		BINDF_OFFLINEOPERATION = &H8
		BINDF_GETNEWESTVERSION = &H10
		BINDF_NOWRITECACHE = &H20
		BINDF_NEEDFILE = &H40
		BINDF_PULLDATA = &H80
		BINDF_IGNORESECURITYPROBLEM = &H100
		BINDF_RESYNCHRONIZE = &H200
		BINDF_HYPERLINK = &H400
		BINDF_NO_UI = &H800
		BINDF_SILENTOPERATION = &H1000
		BINDF_PRAGMA_NO_CACHE = &H2000

		BINDF_GETCLASSOBJECT = &H4000
		BINDF_RESERVED_1 = &H8000

		' bindstatus callback from client is free threaded
		BINDF_FREE_THREADED = &H10000
		' client does not need to know excat size of data available
		' hence the read goes directly to e.g. socket
		BINDF_DIRECT_READ = &H20000
		' is the transaction a forms submit.
		BINDF_FORMS_SUBMIT = &H40000
		BINDF_GETFROMCACHE_IF_NET_FAIL = &H80000
		' binding is from UrlMoniker
		BINDF_FROMURLMON = &H100000
		BINDF_FWD_BACK = &H200000

		BINDF_PREFERDEFAULTHANDLER = &H400000
		BINDF_ENFORCERESTRICTED = &H800000

		' Note:
		' the highest byte 0x??000000 is used internally
		' see other documentation
	End Enum

	<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto, Pack := 4)> _
	Public Structure SECURITY_ATTRIBUTES
		Public nLength As UInteger
		Public lpSecurityDescriptor As UInteger
		Public bInheritHandle As Integer
	End Structure

	<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto, Pack := 4)> _
	Public Structure BINDINFO
		Public cbSize As UInteger
		<MarshalAs(UnmanagedType.LPWStr)> _
		Public szExtraInfo As String
		Public stgmedData As STGMEDIUM
		Public grfBindInfoF As UInteger
		Public dwBindVerb As BINDVERB
		<MarshalAs(UnmanagedType.LPWStr)> _
		Public szCustomVerb As String
		Public cbstgmedData As UInteger
		Public dwOptions As UInteger
		Public dwOptionsFlags As UInteger
		Public dwCodePage As UInteger
		Public securityAttributes As SECURITY_ATTRIBUTES
		Public iid As Guid
		<MarshalAs(UnmanagedType.IUnknown)> _
		Public punk As Object
		Public dwReserved As UInteger
	End Structure

	<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto, Pack := 4), ComConversionLoss> _
	Public Structure FORMATETC
		Public cfFormat As UInteger
		<ComConversionLoss> _
		Public ptd As IntPtr
		Public dwAspect As UInteger
		Public lindex As Integer
		Public tymed As UInteger
	End Structure

	<StructLayout(LayoutKind.Sequential, CharSet := CharSet.Auto, Pack := 4), ComConversionLoss> _
	Public Structure STGMEDIUM
		Public tymed As UInteger
		<ComConversionLoss> _
		Public data As IntPtr
		<MarshalAs(UnmanagedType.IUnknown)> _
		Public pUnkForRelease As Object
	End Structure

	Public Enum BINDSTATUS As UInteger
		BINDSTATUS_FINDINGRESOURCE = 1
		BINDSTATUS_CONNECTING
		BINDSTATUS_REDIRECTING
		BINDSTATUS_BEGINDOWNLOADDATA
		BINDSTATUS_DOWNLOADINGDATA
		BINDSTATUS_ENDDOWNLOADDATA
		BINDSTATUS_BEGINDOWNLOADCOMPONENTS
		BINDSTATUS_INSTALLINGCOMPONENTS
		BINDSTATUS_ENDDOWNLOADCOMPONENTS
		BINDSTATUS_USINGCACHEDCOPY
		BINDSTATUS_SENDINGREQUEST
		BINDSTATUS_CLASSIDAVAILABLE
		BINDSTATUS_MIMETYPEAVAILABLE
		BINDSTATUS_CACHEFILENAMEAVAILABLE
		BINDSTATUS_BEGINSYNCOPERATION
		BINDSTATUS_ENDSYNCOPERATION
		BINDSTATUS_BEGINUPLOADDATA
		BINDSTATUS_UPLOADINGDATA
		BINDSTATUS_ENDUPLOADDATA
		BINDSTATUS_PROTOCOLCLASSID
		BINDSTATUS_ENCODING
		BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE
		BINDSTATUS_CLASSINSTALLLOCATION
		BINDSTATUS_DECODING
		BINDSTATUS_LOADINGMIMEHANDLER
		BINDSTATUS_CONTENTDISPOSITIONATTACH
		BINDSTATUS_FILTERREPORTMIMETYPE
		BINDSTATUS_CLSIDCANINSTANTIATE
		BINDSTATUS_IUNKNOWNAVAILABLE
		BINDSTATUS_DIRECTBIND
		BINDSTATUS_RAWMIMETYPE
		BINDSTATUS_PROXYDETECTING
		BINDSTATUS_ACCEPTRANGES
		BINDSTATUS_COOKIE_SENT
		BINDSTATUS_COMPACT_POLICY_RECEIVED
		BINDSTATUS_COOKIE_SUPPRESSED
		BINDSTATUS_COOKIE_STATE_UNKNOWN
		BINDSTATUS_COOKIE_STATE_ACCEPT
		BINDSTATUS_COOKIE_STATE_REJECT
		BINDSTATUS_COOKIE_STATE_PROMPT
		BINDSTATUS_COOKIE_STATE_LEASH
		BINDSTATUS_COOKIE_STATE_DOWNGRADE
		BINDSTATUS_POLICY_HREF
		BINDSTATUS_P3P_HEADER
		BINDSTATUS_SESSION_COOKIE_RECEIVED
		BINDSTATUS_PERSISTENT_COOKIE_RECEIVED
		BINDSTATUS_SESSION_COOKIES_ALLOWED
	End Enum

	Public Enum HRESULTS As Long
		S_OK = 0
		S_FALSE = 1

		E_NOTIMPL = &H80004001
		E_OUTOFMEMORY = &H8007000e
		E_INVALIDARG = &H80070057
		E_NOINTERFACE = &H80004002
		E_POINTER = &H80004003
		E_HANDLE = &H80070006
		E_ABORT = &H80004004
		E_FAIL = &H80004005
		E_ACCESSDENIED = &H80070005

		' IConnectionPoint errors

		CONNECT_E_FIRST = &H80040200
		CONNECT_E_NOCONNECTION
		' there is no connection for this connection id
		CONNECT_E_ADVISELIMIT
		' this implementation's limit for advisory connections has been reached
		CONNECT_E_CANNOTCONNECT
		' connection attempt failed
		CONNECT_E_OVERRIDDEN
		' must use a derived interface to connect
		' DllRegisterServer/DllUnregisterServer errors
		SELFREG_E_TYPELIB = &H80040200
		' failed to register/unregister type library
		SELFREG_E_CLASS
		' failed to register/unregister class
		' INET errors

		INET_E_INVALID_URL = &H800c0002
		INET_E_NO_SESSION = &H800c0003
		INET_E_CANNOT_CONNECT = &H800c0004
		INET_E_RESOURCE_NOT_FOUND = &H800c0005
		INET_E_OBJECT_NOT_FOUND = &H800c0006
		INET_E_DATA_NOT_AVAILABLE = &H800c0007
		INET_E_DOWNLOAD_FAILURE = &H800c0008
		INET_E_AUTHENTICATION_REQUIRED = &H800c0009
		INET_E_NO_VALID_MEDIA = &H800c000a
		INET_E_CONNECTION_TIMEOUT = &H800c000b
		INET_E_INVALID_REQUEST = &H800c000c
		INET_E_UNKNOWN_PROTOCOL = &H800c000d
		INET_E_SECURITY_PROBLEM = &H800c000e
		INET_E_CANNOT_LOAD_DATA = &H800c000f
		INET_E_CANNOT_INSTANTIATE_OBJECT = &H800c0010
		INET_E_USE_DEFAULT_PROTOCOLHANDLER = &H800c0011
		INET_E_DEFAULT_ACTION = &H800c0011
		INET_E_USE_DEFAULT_SETTING = &H800c0012
		INET_E_QUERYOPTION_UNKNOWN = &H800c0013
		INET_E_REDIRECT_FAILED = &H800c0014
		'INET_E_REDIRECTING
		INET_E_REDIRECT_TO_DIR = &H800c0015
		INET_E_CANNOT_LOCK_REQUEST = &H800c0016
		INET_E_USE_EXTEND_BINDING = &H800c0017
		INET_E_ERROR_FIRST = &H800c0002
		INET_E_ERROR_LAST = &H800c0017
		INET_E_CODE_DOWNLOAD_DECLINED = &H800c0100
		INET_E_RESULT_DISPATCHED = &H800c0200
		INET_E_CANNOT_REPLACE_SFP_FILE = &H800c0300

	End Enum
    
	'--> file downloader class
	<ClassInterface(ClassInterfaceType.None)> _
	Public Class Downloader
		Implements IBindStatusCallback
		Implements IDisposable

				'
				' TODO: Add constructor logic here
				'
		Public Sub New()
		End Sub
		Public Delegate Sub DownloadEvent(file As String)
		Public Delegate Sub DownloadAborted(file As String, Errorcode As Long, Message As String)
		Public Delegate Sub DownloadProgress(file As String, ReceivedBytes As Long, TotalBytes As Long)
		Public Event OnDownloadStarted As DownloadEvent
		Public Event OnDownloadComplete As DownloadEvent
		Public Event OnDownloadAborted As DownloadAborted
		Public Event OnDownloadProgress As DownloadProgress

		Private mobjBinding As IBinding
		Private IsAborted As Boolean = False
		Public LocalPath As String
		Public SourcePath As String
		Protected Overrides Sub Finalize()
          Try

            'finalizer
            Dispose(False)
          Finally
            MyBase.Finalize()
          End Try
		End Sub
		'disposer
		Public Sub Dispose() Implements IDisposable.Dispose
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

		'dispose
		Private disposed As Boolean = False
		''' <summary>
		''' Dispose and free used resources
		''' </summary>
		Protected Overridable Sub Dispose(disposing As Boolean)
          If disposing Then
            If Not disposed Then
              StopDownload()
              If mobjBinding IsNot Nothing Then
                Marshal.ReleaseComObject(mobjBinding)
              End If
              mobjBinding = Nothing
              disposed = True
            End If
          End If
		End Sub

		''' <summary>
		''' The URLMON library contains this function, URLDownloadToFile, which is a way
		''' to download files without user prompts.
		''' </summary>
		''' <param name="pCaller">Pointer to caller object (AX).</param>
		''' <param name="szURL">String of the URL.</param>
		''' <param name="szFileName">String of the destination filename/path.</param>
		''' <param name="dwReserved">[reserved].</param>
		''' <param name="lpfnCB">A callback function to monitor progress or abort.</param>
		''' <returns>throws exception if not success</returns>
		<DllImport("urlmon.dll", CharSet := CharSet.Auto, SetLastError := True, PreserveSig := False)> _
		Public Shared Sub URLDownloadToFile(<MarshalAs(UnmanagedType.IUnknown)> pAxCaller As Object, <MarshalAs(UnmanagedType.LPWStr)> szURL As String, <MarshalAs(UnmanagedType.LPWStr)> szFileName As String, <MarshalAs(UnmanagedType.U4)> dwReserved As UInteger, <MarshalAs(UnmanagedType.[Interface])> lpfnCB As IBindStatusCallback)
		End Sub
        
        
        '--> Class Members
		Public Sub Download()
          Try
            IsAborted = False
            If mobjBinding Is Nothing Then
              Dim destdir As String = Path.GetDirectoryName(Me.LocalPath)
              If Not Directory.Exists(destdir) Then
                  Directory.CreateDirectory(destdir)
              End If
                  'use 0x10 for new download
              URLDownloadToFile(IntPtr.Zero, SourcePath, LocalPath, 0, DirectCast(Me, IBindStatusCallback))
            End If
          Catch e As Exception
            Dim we As New Win32Exception(Marshal.GetLastWin32Error(), e.Message)
            System.Windows.Forms.MessageBox.Show("Download error: \r\n"+we.Message)
            '
          End Try
		End Sub

		Public Sub StopDownload()
          Try
            If mobjBinding IsNot Nothing Then
              SyncLock mobjBinding
                If mobjBinding IsNot Nothing Then
                  mobjBinding.Abort()
                End If
                IsAborted = True
              End SyncLock
            End If
          Catch ex As Exception

            System.Windows.Forms.MessageBox.Show("Download error: " & vbCr & vbLf & ex.Message)
          End Try
		End Sub

		Public Sub GetBindInfo(ByRef grfBINDF As BINDF, ByRef pbindinfo As BINDINFO) Implements IBindStatusCallback.GetBindInfo
          grfBINDF = BINDF.BINDF_IGNORESECURITYPROBLEM
          Try
            Dim cbSize As UInteger = pbindinfo.cbSize
            ' remember incoming cbSize
            pbindinfo = New BINDINFO()
            'reset
            pbindinfo.cbSize = cbSize
            ' restore cbSize
            pbindinfo.dwBindVerb = BINDVERB.BINDVERB_GET
            ' set verb
            pbindinfo.stgmedData.tymed = 0
            pbindinfo.cbstgmedData = CUInt(Marshal.SizeOf(pbindinfo.stgmedData))
            pbindinfo.dwOptions = 0
            pbindinfo.dwOptionsFlags = 0
            pbindinfo.dwReserved = 0
          Catch ex As Exception

            System.Windows.Forms.MessageBox.Show("Download error: " & vbCr & vbLf & ex.Message)
          End Try
		End Sub
		Public Sub GetPriority(ByRef pnPriority As Integer) Implements IBindStatusCallback.GetPriority
          pnPriority = 0
          'THREAD_PRIORITY_NORMAL=0,THREAD_PRIORITY_BELOW_NORMAL=-1
		End Sub
		Public Sub OnDataAvailable(grfBSCF As UInteger, dwSize As UInteger, ByRef pformatetc As FORMATETC, ByRef pstgmed As STGMEDIUM) Implements IBindStatusCallback.OnDataAvailable
		End Sub
		Public Sub OnLowResource(reserved As UInteger) Implements IBindStatusCallback.OnLowResource
		End Sub
		Public Sub OnObjectAvailable(ByRef riid As Guid, punk As Object) Implements IBindStatusCallback.OnObjectAvailable
		End Sub
		Public Sub OnProgress(ulProgress As UInteger, ulProgressMax As UInteger, ulStatusCode As BINDSTATUS, szStatusText As String) Implements IBindStatusCallback.OnProgress
          If ulProgressMax > 0 Then
            RaiseEvent OnDownloadProgress(SourcePath, ulProgress, ulProgressMax)
          End If
          If mobjBinding IsNot Nothing AndAlso IsAborted Then
            mobjBinding.Abort()
          End If
		End Sub
		Public Sub OnStartBinding(dwReserved As UInteger, pib As IBinding) Implements IBindStatusCallback.OnStartBinding
          IsAborted = False
          mobjBinding = pib
          RaiseEvent OnDownloadStarted(SourcePath)
		End Sub
		Public Sub OnStopBinding(hresult As UInteger, szError As String) Implements IBindStatusCallback.OnStopBinding
          Try
            If hresult = 1 Then
              RaiseEvent OnDownloadComplete(SourcePath)
            Else
              RaiseEvent OnDownloadAborted(SourcePath, hresult, ErrorDescription(hresult))
            End If
          Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("Download error: " & vbCr & vbLf & ex.Message)
          End Try
          mobjBinding = Nothing
		End Sub

		Private Function ErrorDescription(ErrNum As Long) As String
          Dim Description As String = ""

          Select Case CType(ErrNum, HRESULTS)
            Case HRESULTS.INET_E_AUTHENTICATION_REQUIRED
            Description = "Authentication Failure."
            Case HRESULTS.INET_E_CANNOT_CONNECT
            Description = "Cannot Connect"
            Case HRESULTS.INET_E_CANNOT_INSTANTIATE_OBJECT
            Description = "Cannot Instantiate Object."
            Case HRESULTS.INET_E_CANNOT_LOAD_DATA
            Description = "Cannot Load Data."
            Case HRESULTS.INET_E_CANNOT_LOCK_REQUEST
            Description = "Cannot Lock Request."
            Case HRESULTS.INET_E_CANNOT_REPLACE_SFP_FILE
            Description = "Cannot Replace SFP File."
            Case HRESULTS.INET_E_CODE_DOWNLOAD_DECLINED
            Description = "Code Download Declined."
            Case HRESULTS.INET_E_CONNECTION_TIMEOUT
            Description = "Connection Timeout."
            Case HRESULTS.INET_E_DATA_NOT_AVAILABLE
            Description = "Data Not Available."
            Case HRESULTS.INET_E_DEFAULT_ACTION
            Description = "Default Action."
            Case HRESULTS.INET_E_DOWNLOAD_FAILURE
            Description = "Download Failure."
            Case HRESULTS.INET_E_INVALID_REQUEST
            Description = "Invalid Request."
            Case HRESULTS.INET_E_INVALID_URL
            Description = "Invalid URL."
            Case HRESULTS.INET_E_NO_SESSION
            Description = "No Session."
            Case HRESULTS.INET_E_NO_VALID_MEDIA
            Description = "No Valid Media."
            Case HRESULTS.INET_E_OBJECT_NOT_FOUND
            Description = "File Not Found."
            Case HRESULTS.INET_E_QUERYOPTION_UNKNOWN
            Description = "QueryOption Unknown."
            Case HRESULTS.INET_E_REDIRECT_FAILED
            Description = "Redirect Failed."
            Case HRESULTS.INET_E_REDIRECT_TO_DIR
            Description = "Redirect To Dir."
            Case HRESULTS.INET_E_RESOURCE_NOT_FOUND
            Description = "Resource Not Found."
            Case HRESULTS.INET_E_RESULT_DISPATCHED
            Description = "Result Dispatched."
            Case HRESULTS.INET_E_SECURITY_PROBLEM
            Description = "Security Problem."
            Case HRESULTS.INET_E_UNKNOWN_PROTOCOL
            Description = "Unknown Protocol."
            Case Else
            Description = "Unknown Error."
          End Select
          If IsAborted Then
            Description = "Download aborted."
          End If
          Return Description
		End Function

	End Class

:END


