
'##################################################
'-->
'--> C O N F I G - G L O B A L

#Para DebugMode intern
#Para SilentMode true




'##################################################



#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic

'NEU...
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

Public Const Auto = -2
shared myScriptClass

shared winId as string ="es_FileFinder02.mainWin"
public winCaption as string = "mein neuer FileFinder 02"
dim toolBarColor as string="#ddd"

public ref As ScriptedPanel

public shared toolBar1 As ScriptedPanel
public shared toolBar2 As ScriptedPanel
public shared statBar As ScriptedPanel
public shared containerMain As ScriptedPanel

Dim cancelSearch
Dim WithEvents FormRef As Form
dim WithEvents myPicture as pictureBox      

Dim WithEvents IGrid1 As IgridEx
' dim globGridHeaderText1=getHeaderData()

public glob_defaultGridPath as string = "C:\yPara\scriptIDE\gridTest2\"
public globMonitorRef as object



'' Dim WithEvents fileFinder As cls_fileSearcher

'' Dim fso
'' Dim fileCount,folderCount
'' 


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
Public Declare Function ReleaseCapture Lib "user32" () As Integer
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2




'-->
'--> I N I

sub onInitialize()
  '...
End Sub


sub onTerminate()
   ''stopTimer()
   globRemoveHandler()
end sub


Sub globAddHandler()
    AddHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
    'AddHandler Timer1.Tick, AddressOf Timer1_Tick
end sub


Sub globRemoveHandler()
    RemoveHandler myPicture.MouseDown, AddressOf myPicture_MouseDown
    'RemoveHandler Timer1.Tick, AddressOf Timer1_Tick
end sub


Function GetFormRef()
  'If ref IsNot Nothing Then Exit Sub
  'ref = ZZ.IDEHelper.CreateDockWindow(winID, 1)
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 1)
  formRef = ref.form
  formRef.text=winCaption
  : CType(FormRef, WeifenLuo.WinFormsUI.Docking.DockContent).HideOnClose = false
  return ref
End function




'-->
'--> A U T O S T A R T

Sub AutoStart()
  dim el as object
  myScriptClass = me
  ref=GetFormRef()
  
  '--> ... iniPanels
  with ref
    .resetControls  (10,5,1)
    
    containerMain  =.addPanel("containerMain"    , DockStyle.Fill)
    toolBar1     =.addPanel("toolBar"      , DockStyle.Top, intHeight:=33)
    'statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=25)
    statBar     =.addPanel("statBar"      , DockStyle.Bottom, intHeight:=28)

    toolBar1.resetControls  (10,3,1)
    toolbar1.visible=true
    toolbar1.height=24
    toolbar1.height=110
    toolbar1.BackColor = ColorTranslator.FromHtml(toolBarColor)
    toolbar1.BackColor = ColorTranslator.FromHtml("#8484DE")

    containerMain.resetControls  (10,5,1)
    containerMain.resetControls  (10,5,1)
    containerMain.BackColor = ColorTranslator.FromHtml("#fff")
    containerMain.BackColor = ColorTranslator.FromHtml("#bbf")

    statBar.resetControls  (10,5,1)
    statBar.BackColor = ColorTranslator.FromHtml(toolBarColor)
    statBar.BackColor = ColorTranslator.FromHtml("#8484DE")
    statBar.height=60
  end with
 





  '--> ... toolbar1
  with toolbar1
     .resetControls (15,8,1)
    .activateEvents = "|ButtonMouseClick|"
    el=.addIcon ("myicon1", "http://es.teamwiki.net/docs/icons/filefind128.png",,,99,99)
    myPicture=el

    .offsetX = 130:  .insertY = 15
    .addTextbox ("startDir", Auto, "Start-Dir:", 70, "#99f") : .BR 
    .offsetX = 130:      .insertY = 40
    .addTextbox ("filter", Auto, "Dateifilter:", 70, "#99f" ) : .BR
    '.addButton ("start",  "   Suche (FSO)   " )
    .addButton ("start2", "  S T A R T  " )
    '.addButton ("start3", "   Suche (int)    " ) 
    .addButton ("cancel", "   Abbruch    ") 

     '--> ... !!! umstellen auf globPara
    .Element("startDir").Text="C:\yPara"'forTest
    .Element("filter").Text="*.*"'forTest
   end with





  
  '--> ... containerMain
  '' With containerMain
  ''   .resetControls (10,5,1)
  ''   .addTextbox  ("out", auto )
  ''   with .Element("out")
  ''     .MultiLine=True
  ''     .Wordwrap=false
  ''     '.Height=300
  ''     .Scrollbars=3
  ''     .MaxLength=0
  ''     .Dock = DockStyle.Fill
  ''   end with
    
  with containerMain
    .resetControls (10,5,1)
      '--> ... ... --- iGrid-1
    IGrid1 = New IgridEx
    IGrid1.name="IGrid1"
    .Controls.Add(IGrid1)
    IGrid1.Dock = DockStyle.Fill
    'IGrid1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
    igrid_setDefaultPara(IGrid1)
    with IGrid1

      .fileSpec=glob_defaultGridPath+"esGrid1.txt"
      .SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
      .rowMode=true
      .HotTracking = False
      '.GroupBox.Visible = True
      .GroupBox.Visible = false
      .SelCellsBackColorNoFocus = ColorTranslator.FromHtml("#4488E9")

      '.AllowDrop = True
      '.ReadOnly = True
      .Cols.Add("Sort",33)
      .Cols.Add("Filename",222)
      .Cols.Add("Ext",77)
      .Cols.Add("Size",77)
      .Cols.Add("Datum",111)
      .Cols.Add("xxx1",000)
      .Cols.Add("xxx2",000)
      .Cols.Add("Pfad",555)
      .Cols.Add("xxx3",555)
      .Cols.Add("xxx4",555)
      '.Cols(1).CellStyle.Type = TenTec.Windows.iGridLib.iGCellType.Check
      .rows.count = 100
    end with
  end with

    'alternativ als Listbox...
    '.addControl "lb1", "System.Windows.Forms.ListBox, System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089, processorArchitecture=MSIL"
    '.addControl "lst_out", "scriptIDE.LstBox",,,Auto,333  : .BR
    
    '.addControl "lst_out", "scriptIDE.TreeVw",,,Auto,333  : .BR
    'ZZ.setBgColor .Element("lst_out"), "#A9FAAD"
    '.br (310)

    '-->      --> VORSCHLAG: 
    ' ... als iGrid: (fehlt noch)
    ' ... und das ganze dynamisch umschaltbar machen ???




'--> ... statbar
  with statbar
    .resetControls (10,5,1)
    .addLabel  ("lab01", "    Folder(s):",  "", ""        ,,,66,22 )
    .addLabel  ("lab02", "             0",  "#333", "#fff",,,66,22 )
    .addLabel  ("lab03", "       File(s)",  "", ""        ,,,66,22 ) 
    .addLabel  ("lab04", "       File(s)",  "#333", "#fff",,,66,22 )
    .br
    .addTextbox ("status2", Auto) 
  end with

  '--> ... add Handler
  globAddHandler()

  formRef.show
  formRef.visible=true

End Sub











'-->
'--> E V E N T S
Sub onButtonEvent(e)
  '' If e.Sender.Name="btn_start" Then
  ''   'startSearch
  '' End If
  
  If e.Sender.Name="btn_start2" Then
    startSearch2
  End If
  
  If e.Sender.Name="btn_start3" Then
    '...
  End If
  
  If e.Sender.Name="btn_cancel" Then
    cancelSearch = True
  End If
  
  '' If e.Sender.Name="btn_indexall" Then
  ''   Dim drive
  ''   For Each drive In ZZ.FSO.Drives
  ''     If drive.DriveType = 2 Then 'Festplatte
  ''       trace drive.DriveLetter,drive.DriveType
  ''       startDriveIndex (drive.DriveLetter+":\")
  ''     End If
  ''   Next
  '' End If

End Sub


Sub onTextboxEvent(e As Object)
  '' TRACE e.Keystring, e.Sender.Name
  '' If e.Keystring = "RETURN" And e.Sender.Name = "txt_url" Then
  ''   'WB.Navigate(FormRef.Controls("txt_url").Text)
  '' End If
  '' If e.Keystring = "RETURN" And e.Sender.Name = "txt_phpsuch" Then
  ''   'WB.Navigate("http://www.php.net/"+FormRef.Controls("txt_phpsuch").Text)
  '' End If
End Sub



Sub myPicture_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
  'trace "PictureBox1_MouseDown","..........."
  'msgBox("????????????????")
  trace "=== ACHTUNG === ", "xxx...testFehler"
  dim dummy as object
  dummy.makeError
  trace "=== ACHTUNG === ", err.number


  if e.Button=Windows.Forms.MouseButtons.Right then
    ' hideMe()
    exit Sub
  end if
  
  dim hwnd as integer
  Call ReleaseCapture()
  ''Call SendMessage(myPicture.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  
  Call SendMessage(FormRef.parent.parent.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  ''Call SendMessage(FormRef.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub







'-->
'--> M A L O C H E 



sub startSearch2

'Datum: der macht nur eingabeDatum: sinnvoll wäre ein UND Änddat
'zwischen Datum und Uhrzeit 2 zeichen Lücke
'Zeit mit Doppelpunkten trennen



  cancelSearch=false
  dim root as string
  root="C:\_sync\"
  root="C:\_sync\vbNet\_AppVorlage"
  root="C:\windows"
  root="C:\_alias"
  root="V:\00_Redaktion"
  root="V:\00_Akademie"
  root="C:\"
  root="C:\_alias"
  root="C:\yExe"
  root="C:\yPara\scriptIDE\scriptClass"
  ref.Element("out").Text= vbNewLine+"   ... SUCHE" + vbNewLine+vbNewLine+"    "+root
  startDriveIndex(root)
end sub



Sub startDriveIndex(startDir)
  'msgbox("startDriveIndex start ????")
  '' ??? if IsEmpty(FormRef) Then Set FormRef=ZZ.getDockPanelRef("Toolbar|##|tbScriptWin|##|fileFinderWin32.eingabe")
  
  ' ref.Element("lst_out").nodeAdd ("",startdir & "0", "Suche läuft  " + startDir)
  ref.Element("status").Text= "Suche läuft  " + startDir
  
  ZZ.TimerStart
  Dim finder as object
  finder= ZZ.OpenFileFinder(startDir, ref.Element("filter").Text)
  finder.SortCriteria=1 'Name
  finder.SortCriteria=4 'Datum 
  
  finder.CallbackScriptClassFunc = "onFileFinderStatusEvent"
  
  Dim RESULT: RESULT=finder.FindRecursive()
  
  ref.Element("lst_out").nodeAdd ("",startdir & "1",Finder.FoundDirectoriesCount & " / " & Finder.FoundFilesCount & "   Benötigte Zeit: " & Finder.ElapsedTime & " ms" & ZZ.TimerGetElapsed() & " ms")
  zz.DoEvents()
  
  setMonitorRef()
  monitorAdd(RESULT)
  
  IGrid_put(iGrid1, RESULT)
  

  with iGrid1
  
  end with
  
  
  
  '...war für die textBox
  'ref.Element("out").Text= RESULT
  
  '...clipboard
  'ZZ.setClipboardText ( RESULT)
  
  '...Datei
  'ZZ.filePutContents (ZZ.FP(startDir,"scannDir2.txt"), RESULT)
  
  
  
End Sub


'-->
'--> C A L L B A C K
Sub onFileFinderStatusEvent(FolderCount, FileCount, ActFolder, ActFileSpec, ByRef Cancel)
  ref.Element("status2").text=FolderCount.toString+"  "+FileCount.toString+"  "+ActFolder
  
  if cancelSearch = true then
    cancelSearch=false
    cancel=true
    ref.Element("out").Text= vbNewLine+"         Suche abgebrochen"
  end if
End Sub




'-->
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
'Public Event updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
'Public Event updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String)
Public Event updateUserFeedback1(ByRef iGrid As IgridEx, ByVal para1 As String)
Public Event updateUserFeedback2(ByRef iGrid As IgridEx, ByVal para1 As String)

'sub IGrid1_updateUserFeedback1(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback1
'sub IGrid1_updateUserFeedback2(ByRef iGrid As TenTec.Windows.iGridLib.iGrid, ByVal para1 As String) Handles iGrid1.updateUserFeedback2

private g_fileSpec as string
: Property fileSpec () As String
:   Get 
:     Return g_fileSpec
:   End Get
:   Set (ByVal value As String)
:     g_fileSpec=value
:   End Set
: End Property

private g_name as string
: Property name () As String
:   Get 
:     Return g_name
:   End Get
:   Set (ByVal value As String)
:     g_name=value
:   End Set
: End Property


Sub myIgrid_ColHdrClick(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGColHdrClickEventArgs) Handles myBase.ColHdrClick
  '=false ... dann sortiert er nicht
  e.DoDefault = true
End Sub

Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  'printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  'printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
End Sub


Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  dim keyCode=e.keycode
  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  if e.control=true then
    'trace "CONTROL",e.keycode 
    'myScriptClass.monitorAdd("CONTROL: "+keyCode.toString)
    if keycode=88 then onClipboardCutSelected(me)   'X
    if keycode=67 then onClipboardCopySelected(me)  'C
    if keycode=86 then onClipboardPasteBeforeSelectedLine(me) 'V
    e.handled =true
  end if
End Sub


sub onClipboardCutSelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardCut()")
  onClipboardCopySelected(iGrid,true)
end sub


sub onClipboardCopySelected(ByVal iGrid As TenTec.Windows.iGridLib.iGrid, optional remove as boolean=false)
  myScriptClass.monitorAdd("CONTROL: onClipboardCopy()")
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  dim max =iGrid.Rows.count
  dim OUT(max) as string
  dim DEL(max) as integer
  dim counter as integer=0
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then 
      'myScriptClass.monitorAdd(row.index.toString)
      OUT(counter)=myScriptClass.JoinIGridLine(row)
      DEL(counter)=row.index
      counter=counter+1
      anz = anz +1
    end if
  Next
  redim preserve OUT(counter-1)
  redim preserve DEL(counter-1)
  if remove=true then
    dim i as integer
    for i=DEL.length-1 to 0 step -1
      iGrid.Rows.RemoveAt(DEL(i))
    next
  end if
  dim RESULT as string=join(OUT, vbNewLine)
  zz.setClipboardText(RESULT)

  myScriptClass.monitorAdd(RESULT)
  myScriptClass.monitorAdd("==========================================")
  myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  myScriptClass.monitorAdd("")
end sub


sub onClipboardPasteBeforeSelectedLine(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  myScriptClass.monitorAdd("CONTROL: onClipboardPasteBeforeSelectedLine()")
  dim selIndex as integer = -1
  
  '...die erste selektierte zeile ermitteln
  Dim row As TenTec.Windows.iGridLib.iGRow
  For Each row In iGrid.Rows
    if row.Selected=true then
      if selIndex=-1 then selIndex=row.index
      row.Selected=false
    end if
  Next

  '... noch nicht entschieden
  if selIndex=-1 then
    '... es ist nichts Selektiert
    '... in Zeile 1 einfügen / ... oder ans ende ???
    'ist grid überhaupt initialisiert ???
  end if
  
  dim newData as string =zz.getClipboardText()
  myScriptClass.monitorAdd(newData)
  dim DATA=split(newData,vbNewLine)
  dim i as integer=0
  dim skip as boolean =true
  for i=DATA.length-1 to 0 step -1
    dim item=DATA(i)
    iGrid.Rows.Insert(selIndex)
    myScriptClass.SplitToIGridLine(iGrid.Rows(selIndex), item)
    '... alle außer der ersten einfügung
    if skip = true  then 
      skip = false
    else
      iGrid.Rows(selIndex+1).selected=true
    end if
  next
  '...den ersten auch
  iGrid.Rows(selIndex).selected=true
  '' myScriptClass.monitorAdd(RESULT)
  '' myScriptClass.monitorAdd("==========================================")
  '' myScriptClass.monitorAdd("Anz. selected: "+anz.toString)
  '' myScriptClass.monitorAdd("")
end sub



Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
  dim curRow as integer
  curRow=me.curRow.index
  trace "curRow", curRow
  RaiseEvent updateUserFeedback2(me, curRow.ToString)
End Sub


:Sub checkSelectionInRowMode(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then anz = anz +1
  Next
  RaiseEvent updateUserFeedback1(me, anz.ToString)

End Sub

:Sub resetAllSelections(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    if row.Selected=true then row.Selected=false
  Next
End Sub

End Class
















'-->
'--> --> IgridHelper2

  
Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
  Dim max = line.Cells.Count - 1
  Dim out(max) As String
  :For i As Integer = 0 To max
    :out(i) = line.Cells(i).Value.ToString
  :Next
  :Return Join(out, Delimiter)
End Function


Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
  Dim max = line.Cells.Count - 1
  Dim out() = Split(input, Delimiter)
  ReDim Preserve out(max)
  :For i As Integer = 0 To max
  :  line.Cells(i).Value = out(i)
  :Next
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
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub


Sub Igrid_put2(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
  Dim out() = Split(FileContent, LineDelim)
  Grid.Cols.Clear
  Grid.Rows.Count = out.Length
  dim header()=Split(out(0), ColDelim)
  for j as integer=0 to header.length-1
    grid.Cols.Add(header(j),66)
  next
  For i As Integer = 0 To out.Length - 1
    SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
  Next
End Sub


sub igrid_setDefaultPara(iGrid)
  with iGrid
    .DefaultCol.CellStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None'=IGCellStyleDesign2
    .DefaultCol.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
    .DefaultCol.CellStyle.EmptyStringAs = TenTec.Windows.iGridLib.iGEmptyStringAs.EmptyString 'zack!
    .DefaultCol.ColHdrStyle.TextTrimming = TenTec.Windows.iGridLib.iGStringTrimming.None
    .SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
    .SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText
  end with
end sub





'-->
'--> outMONITOR

 sub clearAll()
   on error resume next
   monitorClear()
   statustext_reset()
   zz.traceClear()
   zz.printLineReset()
   err.number=0
end sub



sub setMonitorRef()
  'on error resume next
  globMonitorRef = zz.scriptClass("es_testLabor")
  trace "globMonitorRef", globMonitorRef.name
end sub

: function getMonitorRef()
  on error resume next
  if globMonitorRef is nothing then
    globMonitorRef = zz.scriptClass("es_testLabor")
  end if
  return globMonitorRef
end function

: sub monitorClear()
   on error resume next
   globMonitorRef.clear()
   err.number=0
end sub


: sub monitorAdd(para1 as string, optional para2 as string="")
   on error resume next
   globMonitorRef.add(para1+": "+para2+"<--")
   err.number=0
end sub

: sub statustext(para1 as string, optional para2 as string="")
   on error resume next
   'globMonitorRef.statustext(para1+": "+para2+"<--")
   globMonitorRef.statustext(para1+"")
   err.number=0
end sub

: sub statustext_reset()
   on error resume next
   globMonitorRef.statustext_reset()
   err.number=0
end sub


: sub errorText(para as string)
   on error resume next
   globMonitorRef.errorText(para)
   err.number=0
end sub





'-->
'--> M Ü L L


Sub startSearch()
''   'msgbox("startSearch start ????")
''   dim root as string
''   root="C:/yPara/"
''   root="C:\_sync\"
''   root="C:\_sync\vbNet\_AppVorlage"
''   root="C:\yExe"
''   dim fileFinder as object
''   dim erg as string
''   fileFinder = ZZ.OpenFileFinder(root)  
''   ''erg= fileFinder.findFiles()
''   erg = fileFinder.FindRecursive()
''   ZZ.setOutMonitor (erg)
''   fileFinder = nothing
'' 
''   'msgBox ("...aktuelle Baustelle")
'' exit sub
'' 
  '' 'falls die scriptClass zwischendurch zerstört wurde
  '' '' ??? if IsEmpty(FormRef) Then Set FormRef=zz.getDockPanelRef("Toolbar|##|tbScriptWin|##|fileFinderWin32.eingabe")
  '' 
  '' fso=ZZ.CreateObject("Scripting.FileSystemObject")
  '' 
  '' Dim startDir
  '' startDir = ref.Element("startDir").Text
  '' Dim out()
  '' Redim out(0)
  '' fileCount=0 : folderCount=0
  '' cancelSearch=False
  '' 
  '' ref.Element("lst_out").nodeClear
  '' ZZ.TimerStart
  '' 
  '' recursiveFindFile (startDir,out)
  '' out(0)="Anzahl Ordner: " & folderCount & "       Anzahl Dateien: " & fileCount & "   Benötigte Zeit: " & zz.TimerGetElapsed() & " ms"
  '' ref.Element("status").Text= "FERTIG"
  '' ref.Element("status").Text= out(0)
  '' Dim RESULT : RESULT=Join(out,vbNewline)
  '' 'formref.Element("out").Text= RESULT
  '' 
  '' ZZ.filePutContents (startDir+"scannDir.txt",RESULT)
  '' 
  '' 'FormRef.Element("lab02").Text= "OK-1"
  '' 'FormRef.Element("lab04").Text= "OK-2"
End Sub




Sub recursiveFindFile(path,out)
''   Dim folder,files,element,fileSpec
''   If cancelSearch=True Then Exit Sub
''   
''   folder = fso.GetFolder(path)
''   Err.Clear
'' :  Dim x:x=folder.Files.Count
'' ' ??? :  If IsEmpty(folder) Or Err.Number<>0 Then
'' ' ???     Err.Clear 'nachher auch löschen, um Anhalten zu vermeiden
'' ' ???     Msgbox "Pfad nicht gefunden"+vbnewline+path
'' ' ???     Exit Sub
'' ' ???   End If
''   ref.Element("status").Text= folder.Path
''   
''   addToOut (out, ">>>"+folder.path)
''   'FormRef.Element("lst_out").nodeAdd folder.Parent.Path,folder.Path,folder.Name
''   trace "Dateien im Ordner ",path+"   "+typeName(folder)
''   
''   For Each element In folder.SubFolders
''     folderCount=folderCount+1
''     fileSpec = element.Path
''     ref.Element("lst_out").nodeAdd (folder.Path, element.Path,element.Name)
''     recursiveFindFile (fileSpec, out)
''   Next
'' 
''   For Each element In folder.Files
''     fileCount=fileCount+1
''     fileSpec = element.Name
''     ref.Element("lst_out").nodeAdd (folder.Path,element.Path,element.Name)
''     'trace "FI'LE", fileSpec
''     addToOut (out, fileSpec)
''   Next
End Sub


Sub addToOut(out,var)
''   Redim Preserve out(UBound(out)+1)
''   out(UBound(out))=var
''   'FormRef.Element("lst_out").itemAdd var
End Sub
