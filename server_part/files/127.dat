
#Para SilentMode false
#Para DebugMode extern

#Reference TenTec.Windows.iGridLib.iGrid.dll
#Reference System.Web.dll
#Imports System.Web

#Reference System.Xml.dll
#Imports System.Xml

'#imports Tentec.Windows.Igridlib

Dim WithEvents FormRef As Form
Dim WithEvents ref As Object
'Dim PanelRef As ScriptedPanel
Dim WithEvents IGrid1 As IgridEx
dim skipResizeEvent as boolean = true
dim skipIGridEvent as boolean = true
Dim WithEvents tmr1 As Timer
dim WithEvents check1 as System.Windows.Forms.CheckBox
dim WithEvents check2 as System.Windows.Forms.CheckBox
dim WithEvents check3 as System.Windows.Forms.CheckBox


''Public Const winID = "{ScriptClass}.PlayList"
Public Const winID = "es_webReader.vb"
Public Const Auto = -2
Public twLoginuser, twLoginPass, twLoginFullname, twSessID As String

Dim glob As cls_globPara
dim LABELS1(70) as string
dim LABELS2(70) as string


'-->
Sub GetFormRef()
  If ref IsNot Nothing Then Exit Sub
  '' PanelRef = ZZ.IDEHelper.CreateDockWindow(winID)
  '' FormRef = PanelRef.Form
  ref = ZZ.IDEHelper.CreateDockWindow(winID, 4)
  formRef=ref.form
  ''formRef.hide
  formRef.Text = "Auswertung   xxx UGB - Onlinefragebogen    (c) mw, es"
End Sub

sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


Sub createIGrid(ref As Object)
  IGrid1 = New IgridEx
  dim ig as IgridEx = IGrid1
  ig.visible=false
  ig.RowMode = True
  ig.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
  ig.SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
  ig.SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText

  ig.Top = 70
  ig.left = 10
  ig.rows.count = 1
  
  dim i as integer

  :ig.LayoutObject.Text = glob.para("frm_kundenkonto__iGrid1_Layout")
  Err.number=0

 
  ig.cols.count = 55

'' 'leere Zelle als "" statt nothing ??? 
  '' ig.Cols.Add("id")
  '' ig.Cols.Add("nickname")
  '' ig.Cols.Add("gruppe")
  '' ig.Cols.Add("codelink")
  '' ig.GroupBox.Visible = True
  '' ig.Cols(1).width=300
 
  ref.Controls.Add(ig)
  ig.visible=false
End Sub



Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  '' glob.para("frm_webReader__iGrid1_Layout") = IGrid1.LayoutObject.Text
  ''glob.saveTuttiFrutti(FormRef)
  glob.saveFormPos(FormRef)
  glob.saveParaFile()
  trace "formClosing" , "xxx"
End Sub
  




'-->
Sub AutoStart()
  glob = ZZ.newGlobPara()
  glob.readFormPos(FormRef)

  '' ZZ.traceClear()
  '' ZZ.printLineReset
  trace "AutoStart web-READER.vb", "1111111111"
  glob = ZZ.newGlobPara()
  GetFormRef()
  glob.readFormPos(FormRef)
  dim el as object
  With ref
    .resetControls (10,8)

    .insertX = 10 :.insertY = 7
    .addIcon("appPicture", "http://es.teamwiki.net/docs/img/chart_64.png" )

   .insertX = 73 :.insertY = 4
   ''Func+++tion addLabel(strID,  strText,   bgColor fgColor  intLeft = -1,  intTop = -1, intWidth = -1, intHeight = -1)
   ''                  el=.addLabel  ("lab01", "Auswertung   UGB - Onlinefragebogen" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
                     el=.addLabel  ("lab01", "" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
    .insertX = 460 : el=.addLabel  ("curRow",         "zeile" ,  ,"#ffffff" , , ,55 ,31 ) :makeJumboLabel(el)
    .insertX = 520 : el=.addLabel  ("selectionCount", "sel" ,  ,"#ffffff" , , ,88 ,31 ) :makeJumboLabel(el)
 
    'el.width=200
    'el.height=50
    

.BR  '--------------------------------------------------
    .insertX = 70 :.insertY = 40
     el=.addLabel  ("label10"      , "Umfrage-Nr:" ,    ,"#555" )
    .addTextbox ("umfrageNr"       ,  50              , "..."  , 0  , "#aaa")
    .addButton  ("readData"        , "Umfrage lesen" )


.BR  '--------------------------------------------------
.BR  '--------------------------------------------------
    check1 = New System.Windows.Forms.CheckBox
    ''check1.Location = New System.Drawing.Point(360, 10)
    check1.Location = New System.Drawing.Point(12, 350)
    check1.Size = New System.Drawing.Size(100, 19)
    check1.Text = "Spalten-Modus"
    check1.UseVisualStyleBackColor = True
    ref.Controls.Add(check1)
    check1.visible=true
    check1.checked=true

.BR  '--------------------------------------------------
    .insertX = 130 :.insertY = 350
    .addButton  ("copyToClipbord" , "ins Clipboard kopieren "     , "#6f6")
    .addButton  ("saveGridLayout"  , "Spalten-Design speichern" )
    .addButton  ("login"           , "logIn" )
    .addButton  ("logout"          , "logOut (geht noch nicht)" )


.BR  '--------------------------------------------------
    .insertX = 11 :.insertY = 410
    .addButton  ("test12"      , "Alles Markieren"    , "#ccc")
    .addButton  ("test13"      , "Kopieren"           , "#ccc")
    .addButton  ("test14"      , "Ausschneiden"       , "#f66")
    .addButton  ("test15"      , "Einfügen"           , "#6f6")
    .addButton  ("insertLine"  , "Zeile Einfügen"     , "#ccc")
   
.BR  '--------------------------------------------------TEST...
   .addButton  ("readCaptions"   , "Überschriften lesen"    , "#ccc")




.BR  '--------------------------------------------------
   .insertX = 11 :.insertY = 470
   .addTextbox ("outMonitor" ,  -2         , "Roh-"+vbNewLine+"Daten"    , 50  , "#afa", , , "multiline=80")
       ref.element("txt_outMonitor").WordWrap=false
    .br  '--------------------------------------------------
    .addButton  ("clear"          , "clear"     , "#f66")
    .addButton  ("clearGrid"      , "clearGrid" )
    .addButton  ("readXml"        , "XML lesen" )
    '.addTextbox ("inummer"       ,  200     , "iNummer:"    , 50  , "#aaa")

.br  '--------------------------------------------------
    .addButton  ("test1"      , "Lesen"          , "#ccc")
    .addButton  ("test3"      , "testDaten"      , "#ccc")
    .addButton  ("test2"      , "Speichern"      , "#6f6")
 
.br  '--------------------------------------------------
  
     ''trace "appPath", application.executablePath
     dim exePath as string
     exePath=application.executablePath
     if exePath.toLower.endsWith("rtftabmdi.exe") then
       formRef.Text = "--> SCRIPT: ------- web-READER.vb"
     else
       '' formRef.Text = "web-READER.vb   (c)es, mw"
     end if
     'startTimer()
     trace "AutoStart", "FERTIG"
  End With

  createIGrid(ref)

  skipResizeEvent = False
  skipIGridEvent = false
  onResizeControls()
  formRef.show()
 
  umfrageDatenLesen()
  readCaptions_mouseClick(nothing)
  ref.element("txt_outMonitor").focus()

End Sub


sub startTimer
  tmr1 = New Timer
  tmr1.Interval = 300 : tmr1.Enabled = True
end sub


Sub onResizeControls()
  if skipResizeEvent = true then exit sub
  Igrid1.Width = ref.Width - 15
  'Igrid1.Height = ref.Height - 380
  Igrid1.Height = 280
  'ref.element("test2").top=ref.Height - 28
End Sub



'-->
'--> E V E N T S - 1

sub saveGridLayout_mouseClick(e)
  glob.para("frm_kundenkonto__iGrid1_Layout") = IGrid1.LayoutObject.Text
  glob.saveParaFile()
End Sub


sub readData_mouseClick(e)
  umfrageDatenLesen()
End Sub


sub umfrageDatenLesen
  logIn()
  dim RESULT as string
  RESULT=getWebData()
  ref.element("txt_outMonitor").text=RESULT

  '' iGird füllen...
  Dim out() = Split(RESULT, vbNewLine)
  dim ig as IgridEx = IGrid1
  ig.Rows.Clear()
  :ig.LayoutObject.Text = glob.para("frm_kundenkonto__iGrid1_Layout")
  Err.number=0
  ig.cols.count=55
  dim j as integer
  : for j=0 to ig.cols.count-1
  :   ig.Cols(j).text=(j+1).toString
  : next
  ig.Cols(1).CellStyle.BackColor=ColorTranslator.FromHtml("#ccffcc")
  ig.visible=true

  
  '' ig.cols.count=70
  ig.Rows.Count = out.Length
  :For i As Integer = 0 To out.Length - 1
  :  SplitToIGridLine(ig.Rows(i), out(i), vbTab)
  :Next

end sub


function getWebData()
  dim url as string
  dim cookies as string
  dim RESULT as string
  url="http://secure.teamwiki.net/q/profileservice.php"
  url=url+"?m=mydocs_getfile"
  url=url+"&type=out"
  url=url+"&u=ugb"
  url=url+"&f=shared_docs/umfrage_ergebnisse/"
  url=url+"&n=results_001.txt"
  url=url+"&dontcache=" & Now.Ticks.Tostring
  cookies="twnetSID=" + twSessID
  
  RESULT=ZZ.http_get(url, cookies)
  RESULT= replace(RESULT,"%0D%0A","||ZS||")
  RESULT = HttpUtility.UrlDecode(RESULT)
  return RESULT
end function



:sub readCaptions_mouseClick(e)
  dim fileSpec as string = "C:\yPara\scriptIDE\scriptClass\umfrage001-header.txt"
  dim RESULT as string = zz.fileGetContents(filespec)
  dim DATA() as string = split(RESULT,vbNewLine)
  dim i as integer
  dim max as integer=DATA.length
  trace "readCaptions_mouseClick", max
  dim item as string
  dim lastCaption as string
  dim counter as integer =0
  counter=0: LABELS1(counter)="Reserve-1": LABELS2(counter)="Reserve1"
  counter=1: LABELS1(counter)="Reserve-2": LABELS2(counter)="Reserve2"
  counter=2: LABELS1(counter)="Reserve-3": LABELS2(counter)="Reserve3"

  
  for i = 0 to max-1
    item=trim(DATA(i))
    if trim(item)="" then continue for
    if item.startsWith("|||") then             'Hauptüberschriften zuweisen
      lastCaption=trim(mid(item, 4))
      'trace counter, "!!! " + lastCaption
      continue for
    end if
    counter=counter+1
    LABELS1(counter)=lastCaption
    LABELS2(counter)=item
    'trace counter, LABELS2(counter)
  next
End Sub


Sub IGrid1_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles IGrid1.CellMouseDown
  onCollMouseDown(sender, e)
End Sub

sub onCollMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs)
  printLine  6, "AA-IGrid1_CellMouseDown - row", e.RowIndex
  printLine  7, "AA-IGrid1_CellMouseDown - col", e.ColIndex
  ref.element("lab_lab01").text=LABELS2(e.ColIndex)
end sub

'-->
'-->  --> L O G - I N 


sub login_mouseClick(e) 
  logIn()
End Sub


sub logIn()
  dim userName as string
  dim passWord as string
  dim RESULT as boolean
  userName="es"
  passWord="sandus"
  
  RESULT = doLogin( userName, passWord)
  trace "logIn", RESULT
end sub  

Function doLogin(ByVal userName As String, ByVal passWord As String) As Boolean
  Dim postData As String = "u=" + userName + "&p=" + passWord
  Dim RESULT = TwAjax.postUrl("http://teamwiki.net/php/vb/app_login2.php?", postData)
  Dim lines() = Split(RESULT, vbNewLine)
  ReDim Preserve lines(4)
   If lines(0) = "Login OK" Then
    twLoginuser = userName : twLoginPass = passWord : twLoginFullname = lines(3)
    '' glob.para("twLoginData") = glob.RC4StringEncrypt(userName + ":" + passWord, kData)
    twSessID = lines(2)
    Return True
  Else
    MsgBox("Ungültige Login-Daten!", MsgBoxStyle.Exclamation)
    Return False
  End If
End Function



'-->
'--> E V E N T S - 2

sub copyToClipbord_mouseClick(e)
  dim RESULT as string = ref.element("txt_outMonitor").text
  zz.setClipboardText(RESULT)
end sub


sub insertLine_MouseClick(e) 'lesen
  printLine 1, "insertLine", "START"
  dim ig = iGrid1
  if ig.rows is nothing then 
    ig.rows.count=1
    exit sub
  end if
  if ig.curRow is nothing then
    printLine 2, "curRow", "nothing"
    ig.rows.count=1
    exit sub
  end if
  
  dim curRowIndex = ig.curRow.index
  ig.Rows.Insert(curRowIndex+1)
  ig.curRow.index=curRowIndex
  ig.focus()
end sub


sub readXml_mouseClick(e) 'lesen
  printLine  1, "readXml_mouseClick" , "xxxx"
  dim RESULT as string= zz.http_get("http://umfragen.teamwiki.net/survey_001.xml")
  ref.element("txt_outMonitor").text=RESULT
  
  ' Import System.Xml namespace

Dim ws, pi, dc, cc, ac, et, el, xd As Integer
' Read a document
Dim textReader As XmlTextReader = New XmlTextReader(RESULT, XmlNodeType.Document, Nothing)
' Read until end of file
dim processed as boolean
dim out(500) as string
dim i as integer =0

While textReader.Read()
  processed=false
  Dim nType As XmlNodeType = textReader.NodeType
  ' if node type is white space
  If nType = XmlNodeType.Whitespace Then
    '' trace "WhiteSpace:" + textReader.Name.ToString()
    ws = ws + 1
    continue while
  End If
  
  dim curItem as string = textReader.ReadString
  if trim(curItem)<>"" then
    out(i)=curItem
    i = i+1
  end if
  
' If node type us a declaration
  If nType = XmlNodeType.XmlDeclaration Then
    trace "Declaration:" + curItem
    xd = xd + 1
    processed=true
  End If
  ' if node type is a comment
  If nType = XmlNodeType.Comment Then
    trace "Comment:" + curItem
    cc = cc + 1
    processed=true
  End If
  ' if node type us an attribute
  If nType = XmlNodeType.Attribute Then
    trace "Attribute:" + curItem
    ac = ac + 1
    processed=true
  End If
  ' if node type is an element
  If nType = XmlNodeType.Element Then
    trace "Element:" + textReader.Name.ToString() ,curItem
    el = el + 1
    processed=true
  End If
  ' if node type is an entity
  If nType = XmlNodeType.Entity Then
    trace "Entity:" + textReader.Name.ToString(), curItem
    et = et + 1
    processed=true
  End If
  ' if node type is a Process Instruction
  If nType = XmlNodeType.Entity Then
    trace "Entity:" + textReader.Name.ToString(), curItem
    pi = pi + 1
    processed=true
  End If
  ' if node type a document
  If nType = XmlNodeType.DocumentType Then
    trace "Document:" + textReader.Name.ToString(), curItem
    dc = dc + 1
    processed=true
  End If

  if processed=false then
    trace "!!!: "+nType.toString+ " --- "+textReader.Name.ToString() ,curItem
  end if


End While
trace "Total Comments:" + cc.ToString()
trace "Total Attributes:" + ac.ToString()
trace "Total Elements:" + el.ToString()
trace "Total Entity:" + et.ToString()
trace "Total Process Instructions:" + pi.ToString()
trace "Total Declaration:" + xd.ToString()
trace "Total DocumentType:" + dc.ToString()
trace "Total WhiteSpaces:" + ws.ToString()

dim out2=join(out,vbnewline)
'msgBox (out2)
zz.setOutMonitor(out2)

  
  
  
  
End Sub


Private Sub IGrid1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles IGrid1.MouseDown
  printLine  1, "mausKnopf" , e.button
  if e.Button = Windows.Forms.MouseButtons.Right then
    printLine  2, "mausKnopf" , "rClick"
    'switchColMode
    'selectAllInCol(colIndex)
  else
    printLine  2, "mausKnopf" , "NORMAL"
  end if
End Sub



Sub check1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check1.CheckedChanged
  if skipIGridEvent = true then exit sub 
  trace "check1_CheckedChanged", check1.checked
  dim ig = iGrid1
  dim col = IGrid1.CurCell.ColIndex
  dim row = IGrid1.CurCell.RowIndex
  ig.rowmode=not sender.checked
  ig.resetAllSelections(ig)
  if ig.rowmode=true then
    ig.rows(row).selected=true
  else
    ig.cells(row,col).selected=true
  end if
  ig.focus()
End Sub


Sub check2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check2.CheckedChanged
  trace "check22_CheckedChanged", check1.checked
End Sub


Sub check3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles check3.CheckedChanged
  trace "check333_CheckedChanged", check1.checked
End Sub


Sub iGrid1_updateUserFeedback1(ByVal para1 As String) Handles iGrid1.updateUserFeedback1
  'trace "iGrid1_updateUserFeedback: ", para1
  ref.element("lab_selectionCount").text=para1
End Sub

Sub iGrid1_updateUserFeedback2(ByVal para1 As String) Handles iGrid1.updateUserFeedback2
  ref.element("lab_curRow").text=para1
End Sub











sub test1_mouseClick(e) 'lesen
  printLine 11, "xxx", "12345"
  '' '...testfehler
  '' dim y as object
  '' y.kjdfkdjskfj

  ''msgBox ("xxx-test1")
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string
  strContent=ZZ.TwReadFile(strAppName, strFileName) 
  ref.element("txt_outMonitor").text=strContent
  IGrid1.cols.Count = 25

  Igrid_put(IGrid1, strContent, vbNewLine, "|")
  IGrid1.rows.count=20  'provisorisch
end sub


sub test2_mouseClick(e) 'speichern
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string: strContent=""
  dim errMes as string
  Igrid_get(IGrid1, strContent, vbNewLine, "|")
  errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  ref.element("txt_outMonitor").text=errMes
end sub


sub test3_mouseClick(e) 'Testdaten anzeigen
  'dim strAppName: strAppName="scriptIDE_test_es"
  dim strAppName: strAppName="projektmanager"
  dim strFileName: strFileName="gridData.grid"
  dim strContent as string: strContent=""
  dim errMes as string
  strContent="1111111111||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  strContent=strContent+vbNewLine+"1111111112||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  strContent=strContent+vbNewLine+"1111111113||1252173651038-es||DEFAULT-ITEM||145|131|324|22||||0px  solid #fff|||#DDD|#000|| [_] ||||||"
  ''IgridHelper.Igrid_get(IGrid1, strContent, vbNewLine, "|")
  '' errMes=ZZ.TwSaveFile(strAppName, strFileName, strContent)
  Igrid_put(IGrid1, strContent, vbNewLine, "|")
end sub


sub clearMonitor_mouseClick(e)
  ref.element("txt_outMonitor").text="gelöscht"
  ref.element("txt_outMonitor").WordWrap=false
end sub

sub clear_mouseClick(e)
  ref.element("txt_outMonitor").text="gelöscht"
  ref.element("txt_outMonitor").WordWrap=false
end sub


sub clearGrid_mouseClick(e)
  IGrid1.Rows.Clear()
end sub


'--> --> ...obolet
Sub onButtonEvent(e As Object)
  dim sName as string: sName=e.Sender.Name
  printLine  1, "ToolbarButtonClick", sName
  printLine  2, sName , "..........."
  
  If sName = "btn_read" Then
    Dim Content = ZZ.fileGetContents("C:\yPara\scriptIDE\codeBookmarks.txt")
    Igrid_put (IGrid1, Content)
    printLine  2, sName , "OK"
  End If
  If sName = "btn_save" Then
    Dim Content As String
    Igrid_get (IGrid1, Content)
    ZZ.filePutContents("C:\yPara\scriptIDE\codeBookmarks.txt", Content)
    printLine  2, sName , "OK"
  End If
End Sub

'-->


Sub onTimerEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr1.Tick
'On Error Resume Next
'  If VLC1.playlist.isPlaying Then
'    pbPos.Maximum = VLC1.input.Length
'    pbPos.Value = VLC1.input.position*1000
'  End If
End Sub


Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls()
End Sub



Sub onTerminate()
  'If VLC1.playlist.isPlaying Then VLC1.playlist.stop()
End Sub


'-->
'-->  --> CLASS: IgridEx

Class IgridEx

Inherits Tentec.Windows.Igridlib.IGrid
Public Event updateUserFeedback1(ByVal para1 As String)
Public Event updateUserFeedback2(ByVal para1 As String)


Sub myIgrid_CellMouseDown(ByVal sender As Object, ByVal e As TenTec.Windows.iGridLib.iGCellMouseDownEventArgs) Handles myBase.CellMouseDown
  'printLine  6, "IGrid1_CellMouseDown - row", e.RowIndex
  'printLine  7, "IGrid1_CellMouseDown - col", e.ColIndex
End Sub


Sub myIgrid_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles myBase.KeyDown
  printLine  4, "myIgrid_KeyDown", e.keycode 
  '' 'test: fehler
  '' dim y as object
  '' y.kjdfkdjskfj
    
  '' msgBox ("aaaaaaaaaaaaaaa")
  printLine  5, "myIgrid_KeyDown", e.keycode 
  e.handled =true
End Sub


Private Sub IgridEx_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.SelectionChanged
  trace "IGrid1_SelectionChanged..."
  checkSelectionInRowMode(me)
  trace "---------------------------------" 
End Sub

Private Sub IgridEx_CurRowChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles myBase.CurRowChanged
  dim curRow as integer
  curRow=me.curRow.index
  trace "curRow", curRow
  RaiseEvent updateUserFeedback2(curRow.ToString)

End Sub

:Sub checkSelectionInRowMode(ByVal iGrid As TenTec.Windows.iGridLib.iGrid)
  Dim row As TenTec.Windows.iGridLib.iGRow
  dim anz as integer
  For Each row In iGrid.Rows
    'Debug.Print(row.Selected)
    if row.Selected=true then anz = anz +1
  Next
  RaiseEvent updateUserFeedback1(anz.ToString)

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
  




'-->
'--> --> IMPORT


