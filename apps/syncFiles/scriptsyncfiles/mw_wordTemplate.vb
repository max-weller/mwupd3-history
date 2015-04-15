'-->
'--> --- GLOB -----------------------------------------

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
'#Reference siaIDEMain.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

'#Reference siaGridView.dll
#Para DebugMode int
'' 
'' #WindowData frm_eingabe
'' 			500	400	Form		Text="Rechnung"|Size = New Size(330,245)|FormBorderStyle=FormBorderStyle.SizableToolWindow
'' 	20	125	420	180	IGrid	ig_Pos	Text=Nothing|Anchor=15
'' 	15	8	48	48	Picturebox	pic_Icon	Image=ZZ.getImageCached("http://mw.teamwiki.net/docs/img/win-icons/SHELL32_308-48.png")
'' //	90	8	225	50	Label	lab_Prompt	Text=""|Anchor=13|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
'' 	135	350	80	25	Button	btn_OK	Text="Ausgabe"|DialogResult=DialogResult.OK|Anchor=10
'' 	220	350	80	25	Button	btn_Cancel	Text="Schließen"|DialogResult=DialogResult.Cancel|Anchor=10
'' #EndData


'U:\zPara\script4\i-ST-EtikettenVorlage.txt

'' Dim docFileName As String :docFileName = "V:\00_Pressearbeit\Aktion_Probeabo\baustein_Fastenklinik.doc"
'' Dim outFile As String     :outFile = "V:\00_Pressearbeit\Aktion_Probeabo\AUSGABE-1.doc" 'monitor_ausgabeFile.Text
'' 
'' Dim docFileName As String    = "V:\00_Pressearbeit\Aktion_Probeabo\baustein_Wellnesshotels.doc"
'' Dim outFile As String        = "V:\00_Pressearbeit\Aktion_Probeabo\AUSGABE-2.doc" 'monitor_ausgabeFile.Text
'' 
Dim docFileName As String    = "D:\EigeneDateien\Dokumente\Sekreteriat\Impressionen\Rechnungen\reVorlage1.doc"
Dim reDataFolder As String   = "D:\EigeneDateien\Dokumente\Sekreteriat\Impressionen\Rechnungen\data\"
Dim outRootFolder As String  = "D:\EigeneDateien\Dokumente\Sekreteriat\Impressionen\Rechnungen\" 'monitor_ausgabeFile.Text

Dim WithEvents FormRef As Form
Dim WithEvents ref As Object
'Dim PanelRef As ScriptedPanel
Dim WithEvents IGrid1 As IgridEx
Dim WithEvents IGrid2 As IgridEx
dim skipResizeEvent as boolean = true
dim skipIGridEvent as boolean = true
Dim WithEvents tmr1 As Timer
dim WithEvents list1 as ListBox
dim WithEvents check2 as System.Windows.Forms.CheckBox
dim WithEvents check3 as System.Windows.Forms.CheckBox


''Public Const winID = "{ScriptClass}.PlayList"
Public Const winID = "{ScriptClass}.eingabeForm"
Public Const Auto = -2
Public twLoginuser, twLoginPass, twLoginFullname, twSessID As String

Dim glob As cls_globPara
dim LABELS1(70) as string
dim LABELS2(70) as string

Dim reContentArea As String
Dim reIsDuplikat As Boolean

'-->
Sub GetFormRef()
  If ref IsNot Nothing Then Exit Sub
  '' PanelRef = ZZ.IDEHelper.CreateDockWindow(winID)
  '' FormRef = PanelRef.Form
  ref = ZZ.CreateWindow(winID, DWndFlags.DockWindow, 4) 'ZZ.IDEHelper.CreateDockWindow(winID, 4)
  formRef=ref.form
  ''formRef.hide
  formRef.Text = "Rechnungsgenerator"
End Sub

sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


Sub createIGrid(ref As Object, ByRef ig As IgridEx, name As string, xPos As Integer, yPos As Integer, rowCount As Integer, ParamArray cols() As String)
  'IGrid1 = New IgridEx
  ig  = New IgridEx
  ig.Name = name
  ig.visible=false
  ig.RowMode = True
  ig.SelectionMode = TenTec.Windows.iGridLib.iGSelectionMode.MultiExtended
  ig.SelCellsBackColorNoFocus = System.Drawing.SystemColors.Highlight
  ig.SelCellsForeColorNoFocus = System.Drawing.SystemColors.HighlightText

  ig.left = xPos
  ig.Top = yPos
  ig.rows.count = rowCount
  
  dim i as integer

  If (cols.Length Mod 2) <> 0 Then msgbox("cols.length must be even")
  
  For i=0 To UBound(cols) Step 2
    Dim col=ig.Cols.Add(cols(i), cols(i+1))
    col.CellStyle.ValueType = GetType(String)
    col.CellStyle.EmptyStringAs = iGEmptyStringAs.EmptyString
  Next
  
  :ig.LayoutObject.Text = glob.para("wordTemplate__IGrid_"+name+"_Layout")
  Err.number=0
  
'' 'leere Zelle als "" statt nothing ??? 
  '' ig.Cols.Add("id")
  '' ig.Cols.Add("nickname")
  '' ig.Cols.Add("gruppe")
  '' ig.Cols.Add("codelink")
  '' ig.GroupBox.Visible = True
  '' ig.Cols(1).width=300
 
  ref.Controls.Add(ig)
  ig.visible=true
End Sub



Sub Frm_FormClosing(Sender As Object,e As System.Windows.Forms.FormClosingEventArgs) Handles FormRef.FormClosing
  '' glob.para("frm_webReader__iGrid1_Layout") = IGrid1.LayoutObject.Text
  ''glob.saveTuttiFrutti(FormRef)
  glob.saveFormPos(FormRef)
  glob.saveParaFile()
  trace "formClosing" , "xxx"
End Sub
  

Sub Frm_Resize(Sender As Object,e As System.EventArgs) Handles FormRef.Resize
  onResizeControls
End Sub
  




'-->

Sub onTerminate()
  closeDbConnection()
End Sub

Sub AutoStart()
  glob = ZZ.newGlobPara()
  glob.readFormPos(FormRef)
  
  '' ZZ.traceClear()
  '' ZZ.printLineReset
  trace "AutoStart word-TEMPLATE", "1111111111"
  glob = ZZ.newGlobPara()
  GetFormRef()
  glob.readFormPos(FormRef)
  dim el as object
  Dim pnlBottom1 as object
  '--> ... Main
  With ref
    .resetControls (10,8)

    .insertX = 10 :.insertY = 7
    .addIcon("appPicture", "http://es.teamwiki.net/docs/img/chart_64.png" )

   .insertX = 73 :.insertY = 4
   ''Func+++tion addLabel(strID,  strText,   bgColor fgColor  intLeft = -1,  intTop = -1, intWidth = -1, intHeight = -1)
   ''                  el=.addLabel  ("lab01", "Auswertung   UGB - Onlinefragebogen" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
                     el=.addLabel  ("lab01", "Rechnungs-Generator" ,  ,"#ffffff",,,380,31) :makeJumboLabel(el)
    .insertX = 460 : el=.addLabel  ("curRow",         "zeile" ,  ,"#ffffff" , , ,55 ,31 ) :makeJumboLabel(el)
    .insertX = 520 : el=.addLabel  ("selectionCount", "sel" ,  ,"#ffffff" , , ,88 ,31 ) :makeJumboLabel(el)
 
    'el.width=200
    'el.height=50
    

.BR  '--------------------------------------------------
    .insertX = 70 :.insertY = 40
     el=.addLabel  ("label10"      , "Re-Nr:" ,    ,"#555" )
    .addTextbox ("reNr"    ,  50   ,          , 0)
    .addButton  ("readData"        , "Rechnung laden" )
    .addButton  ("nextNr"          , " N E U " )
    .addButton  ("saveData"        , "Speichern" )


.BR  '--------------------------------------------------
   .insertX = 300 :.insertY = 70
   .addButton("newAddr"            , "Neue Adresse")
   .addButton("saveAddr"            , "Übernehmen")
.BR  '--------------------------------------------------
   .insertX = 300
   .addTextbox ("reAddr" ,  300         , "Rechnungs-"+vbNewLine+"Adresse"    , 80  , "#afa", , , "multiline=80")
   
    createIGrid(ref, IGrid2, "igAdr", 11, 70, 0, _
                "ID"        , "ID", _
                "Nachname"  , "Nachname", _
                "Vorname"   , "Vorname", _
                "Firma"     , "Firma")
    IGrid2.Size = New Size(280, 110)
    
.BR  '--------------------------------------------------
    createIGrid(ref, IGrid1, "igPosition", .insertX, .insertY, 20, _
                "desc"   , "Beschreibung", _
                "anz"    , "Anzahl", _
                "ez"     , "Einzelpreis", _
                "ges"    , "Gesamtpreis")
   
.BR  '--------------------------------------------------
    '' check1 = New System.Windows.Forms.CheckBox
    '' ''check1.Location = New System.Drawing.Point(360, 10)
    '' check1.Location = New System.Drawing.Point(12, 350)
    '' check1.Size = New System.Drawing.Size(100, 19)
    '' check1.Text = "Spalten-Modus"
    '' check1.UseVisualStyleBackColor = True
    '' ref.Controls.Add(check1)
    '' check1.visible=true
    '' check1.checked=true

    pnlBottom1 = .addPanel("pnl_Bottom1", DockStyle.Bottom, intHeight := 220)
   ' pnlBottom2 = .addPanel("pnl_Bottom2", DockStyle.Bottom)
    
  End With
  
  
  '--> ... Panel
  
  With pnlBottom1
    .resetControls(11,11)
    
    
.BR  '--------------------------------------------------
    .insertX = 350 :.insertY = 0
    
    .addLabel  ("label11"      , "Nettobetrag:"    , "#bbb" ,"#333" , , , 80, 15)
    .addLabel  ("label10"      , "MwSt:"           , "#bbb" ,"#333" , , , 80, 15)
    .addLabel  ("label12"      , "Rechnungsbetrag" , "#bbb" ,"#333" , , , 100, 15)
    
    
.BR  '--------------------------------------------------
    .insertX = 350 :.insertY = 16
    
    .addLabel  ("netto"      , "0,00" , "#999" ,"#eee" , , , 80, 22)
    .addLabel  ("mwstSum"    , "0,00" , "#999" ,"#eee" , , , 80, 22)
    .addLabel  ("brutto"     , "0,00" , "#999" ,"#eee" , , , 100, 22)
    
    
.BR  '--------------------------------------------------
    .insertX = 11 :.insertY = 12
    .addButton  ("createDoc" , "    OK - Dokument ausgeben    "     , "#6f6")
    
    
    
.BR  '--------------------------------------------------

    .addTextbox ("reTitle"    ,  180      , "Titel(1)"    , 80  , "#afa")
    .addTextbox ("reSubtitle" ,  300      , "Titel(2)"    , 80  , "#afa")
.BR  '--------------------------------------------------
    .addTextbox ("reDatum"    ,  180      , "Datum"       , 80  , "#afa")
    .addTextbox ("reMwst"     ,  100      , "MwSt-Satz"   , 80  , "#afa")
.BR  '--------------------------------------------------
    .addTextbox ("targetFile" ,  380      , "Dateiname"   , 80  , "#afa")
    .addButton  ("buildTargetFilename"    , "make" )
   
.BR  '--------------------------------------------------
    '.insertX = 11 :.insertY = 510
    .addButton  ("saveGridLayout"  , "Spalten-Design speichern" )
    .addButton  ("test12"      , "Alles Markieren"    , "#ccc")
    .addButton  ("test13"      , "Kopieren"           , "#ccc")
    .addButton  ("test14"      , "Ausschneiden"       , "#f66")
    .addButton  ("test15"      , "Einfügen"           , "#6f6")
    .addButton  ("insertLine"  , "Zeile Einfügen"     , "#ccc")
   
    '.addButton  ("readCaptions"   , "Überschriften lesen"    , "#ccc")



  
  End With
  
  ref.Element("txt_reMwst").Text = "19"
  ref.Element("txt_reTitle").Text = "Rechnung"
  ref.Element("txt_reDatum").Text = Today.ToShortDateString
  
  
  '--> ... ini
  
  openDBConnection("D:\EigeneDateien\Dokumente\Sekreteriat\Impressionen\adb\adr-impressionen.mdb")
  readAdrList
  
  skipResizeEvent = False
  skipIGridEvent = false
  onResizeControls()
  formRef.show()
 
  '' umfrageDatenLesen()
  '' readCaptions_mouseClick(nothing)
  'ref.element("txt_outMonitor").focus()

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

sub saveGridLayout_mouseClick(e)
Trace "saveGridLAYOUT"
  glob.para("wordTemplate__IGrid_"+IGrid1.Name+"_Layout") = IGrid1.LayoutObject.Text
  glob.para("wordTemplate__IGrid_"+IGrid2.Name+"_Layout") = IGrid2.LayoutObject.Text
  glob.saveParaFile()
End Sub

'-->

Sub readAdrList()
  IGrid2.ReadOnly = True
  FillDbGrid(IGrid2, "Select Nachname,Vorname,Firma,ID From Kunden")
  
End Sub

Sub iGrid_CurRowChanged(sender As Object, e As EventArgs) Handles iGrid2.CurRowChanged
  If IGrid2.Currow Is Nothing Then Exit Sub
  
  Dim id As String = dbn2es(Igrid2.Currow.Cells("ID").Value)
  Dim read = dbGetReader("SELECT * FROM Kunden WHERE ID = "+id) : read.Read()
  
  Dim out As New System.Text.StringBuilder
  For i As Integer = 2 To 8
    out.AppendLine(dbn2es(read.Item("A" & i)))
  Next
  
  ref.Element("txt_reAddr").Text = out.TOstring
End Sub

Function getAdrLine(idx As Integer) As String
  Dim tb As TextBox = ref.Element("txt_reAddr")
  Return tb.Lines(idx)
End Function

'-->

Function getReNrFormatted() As String
  Return ref.Element("reNr").Text.PadLeft(5,"0")
End Function

Sub buildTargetFilename_MouseClick(e)
  Dim reNr As String = getReNrFormatted()
  
  Dim dat = ref.Element("reDatum").Text
  Dim datInvers As String = Mid(dat,7,4)+"-"+Mid(dat,4,2)+"-"+Mid(dat, 1,2)
  'ref.Element("reSubtitle").Text = out(7)
  
  Dim name = Today.Year & "\" & "RE-" & reNr & "-" & datInvers & "_" & ".doc"
  ref.Element("targetFile").Text = name
End Sub

Sub readData_MouseClick(e)
  Dim fileSpec = reDataFolder + ref.Element("reNr").Text + ".txt"
  Dim fileContent = IO.File.ReadAllText(fileSpec)
  
  Dim out() = Split(fileContent, vbNewLine, 16)
  IGrid_put(IGrid1, out(15))
  
  
  ref.Element("targetFile").Text = out(1)
  
  ref.Element("reDatum").Text = out(3)
  ref.Element("reMwst").Text = out(4)
  
  ref.Element("reTitle").Text = out(6)
  ref.Element("reSubtitle").Text = out(7)
  
  ref.Element("reAddr").Text = Replace(out(9), "|ZS|", vbNewLine)
  
End Sub


Sub saveData_MouseClick(e)
  Dim fileSpec = reDataFolder + ref.Element("reNr").Text + ".txt"
  
  Dim out(16) As String
  IGrid_get (IGrid1, out(15))
  
  out(1) = ref.Element("targetFile").Text
  
  out(3) = ref.Element("reDatum").Text
  out(4) = ref.Element("reMwst").Text
  
  out(6) = ref.Element("reTitle").Text
  out(7) = ref.Element("reSubtitle").Text
  
  out(9) = Replace(ref.Element("reAddr").Text, vbNewLine, "|ZS|")
  
  IO.File.WriteAllLines(fileSpec, out)
  
End Sub


Sub createDoc_MouseClick(e)
  createRechnung()
End Sub

Sub createRechnung()
    Trace "<<<<<< listHeader >>>>>>>>>>"
    
    
    IGrid_get (IGrid1, reContentArea)
    
    If wApp Is Nothing Then
      createWordApp()
    End If
    
    ''wApp.Visible = False
    
    
    Dim targetFile As String = outRootFolder+ref.Element("targetFile").Text
    
    curDoc = getWordDoc(targetFile) ''("briefeSchreibenTest.doc")
    
    'curDoc.ActiveWindow.Visible = False
    curDoc.ActiveWindow.Selection.WholeStory()
    curDoc.ActiveWindow.Selection.Delete(wdCharacter, 1)
    
    
    '.picUpTab ("now     "+cStr(now))
    '.picUpTab ("Anzahl: ... "+cStr(listCount))
    
    'Trace "ListBody", .bbb & " " & .ccc
    
    'actInummer = iNummer
    'Debug.Print(i & " = iNummer " & grid.CellValue(i, 1))
    
    reIsDuplikat = False
    insertDocument(curDoc, docFileName, False)
    replacePlaceholder(curDoc)
    
    reIsDuplikat = True
    insertDocument(curDoc, docFileName, True)
    replacePlaceholder(curDoc)

    curDoc.ActiveWindow.Visible = True
    curDoc.Activate()
    
End Sub

'' 
'' 
''   out(1) = ref.Element("targetFile").Text
''   
''   out(3) = ref.Element("reDatum").Text
''   out(4) = ref.Element("reMwst").Text
''   
''   out(6) = ref.Element("reTitle").Text
''   out(7) = ref.Element("reSubtitle").Text
''   
''   out(9) = Replace(ref.Element("reAddr").Text, vbNewLine, "|ZS|")
''   
'-->
'--> --- getPlaceHolder -------------------------------------------------------------------

  Public actInummer As String
  Public globRS As Object

  Function getContentForPlaceholder(ByVal key As String) As String
    On Error Resume Next
    
    Select Case key
      Case "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8"
        Dim idx As Integer
        If Int32.TryParse(key.SubString(1), idx) Then Return getAdrLine(idx-1)
        
      Case "DUPLIKAT"
        If reIsDuplikat Then
          Return "- Duplikat -"
        Else
          Return ""
        End If
        
      Case "HEADING1"
        Return ref.Element("reTitle").Text
        
      Case "HEADING2"
        Return ref.Element("reSubtitle").Text
        
      Case "DATUM"
        Return ref.Element("reMwst").Text
        
      Case "MWSTP"
        Return ref.Element("reDatum").Text
        
      Case "RENR"
        Return getReNrFormatted()
        
      Case "CONTENTAREA"
        Return reContentArea
        
        
    End Select
''     
''     If key.Startswith("#") Then
''       Dim field = key.SubString(1)
''       Dim fieldIndex As Integer
''       If Not Integer.TryParse(field, fieldIndex) Then
''         fieldIndex = Asc(UCase(field)) - 65
''         
''       End If
''       'Return gridInfo.Col(fieldIndex)
''     End If
''     
''     Static nr As Integer = 0
''     Static lastInummer As String
''     Trace lastInummer
''     Trace actInummer
''     ''lastInummer = ""
''     If lastInummer <> actInummer Then
''       navigateINummer(actInummer)
''       lastInummer = actInummer
''     End If
'' 
''     Dim RESULT As String = ""
''     ''Debug.Write(actInummer) : Debug.Print(key)
''     Dim keyValue As String = "EMPTY"
''     keyValue = globRS.Fields(key).value()
''     '' Debug.Print(keyValue)
''     If keyValue <> "EMPTY" Then
''       RESULT = keyValue
''     Else
''       RESULT = "x" & actInummer & "x" & key & "x" & nr
''       nr = nr + 1
''     End If
'' 
''     Return RESULT

  End Function

  Function navigateINummer(ByVal iNr As String)
    Dim iNummer As String
    ''Dim RS As Object
    Static globCounter As Integer
    globCounter = globCounter + 1
    iNummer = iNr
    iNummer = Replace(iNummer, "i-", "")
    If iNummer = "" Then Exit Function
    globRS = zoomAdr(iNummer)
    If globRS.eof = False Then
      Trace "ich bin das recordSet: ", globRS.Fields("name").value
      'DbFieldText("kurzanschrift") = RS.Fields("kurzanschrift").value
      'globCounter.ToString + ": " + 
      'frm_main.monitor_curAdr2.Text = globCounter.ToString + ": " + globRS.Fields("kurzanschrift").value
      'frm_main.monitor_email.Text = globRS.Fields("email").value
    Else
      ''DbFieldText("iNummer") = "FEHLER: Adresse nicht Gefunden"
      Trace "FEHLER: Adresse nicht Gefunden"
    End If
    Application.DoEvents()
    Return (globRS)
  End Function





'-->
'--> --- WORD -----------------------------------------------------------------------------

  Public wApp As Object 'Word.Application
  Public curDoc As Object 'Word.Document

  Public Const wdCollapseEnd = 0
  'Public Const wdPageBreak = 7
  Public Const wdSeekCurrentPageHeader = 9
  Public Const wdCharacter = 1
  Public Const msoTextOrientationHorizontal = 1
  Public Const wdAlignParagraphRight = 2
  Public Const wdScreen = 7
  Public Const wdStory = 6
  'Public Const wdFindContinue = 1
  Public Const wdReplaceAll = 2

  Public Const wdGoToPage = 1
  Public Const wdGoToAbsolute = 1
  Public Const wdGoToNext = 2
  Public Const wdPageBreak = 7
  Public Const wdFindContinue = 1




  Function getWordDoc(ByVal fileSpec As String) As Object 'Word.Document
    ' Stop 
    On Error Resume Next
    Trace "Anzahl Docs: ", wApp.Documents.Count
    For Each doc As Object In wApp.Documents
      Trace " -->  " & doc.Name, doc.FullName
      ''If doc.FullName.ToLower = fileSpec Then

      ''--> !!! filespec hat falsches format
      If doc.FullName.ToLower = fileSpec.ToLower Then
        doc.ActiveWindow.Visible = True
        Return doc
      End If
    Next
    Dim newDoc As Object 'Word.Document
    Dim filePath As String = IO.Path.GetDirectoryName(fileSpec)
    Dim fileName As String = IO.Path.GetFileName(fileSpec)
    Trace "filePath", filePath
    Trace "fileName", fileName
    If IO.File.Exists(fileSpec) Then
      wApp.ChangeFileOpenDirectory(filePath)

      newDoc = wApp.Documents.Open(fileName)
      Trace "err", Err.Description


      newDoc.ActiveWindow.Visible = True
    Else
      newDoc = wApp.Documents.Add()
      newDoc.SaveAs(fileSpec)
    End If

    newDoc.ActiveWindow.Visible = True
    Return newDoc

  End Function

  Sub createWordApp()

    If wApp IsNot Nothing Then Exit Sub

:   wApp = GetObject(, "Word.Application")
:   If Err.Number<>0 Then wApp = CreateObject("Word.Application")
:   Err.Clear

':   wApp.Visible = False : Err.Clear

  End Sub

  Sub insertDocument(ByVal wDoc As Object, ByVal docFilename As String, ByVal insertNewPageBefore As Boolean)

    'Dim doc As Word.Document
    'doc = getWordDoc(frm_main.monitor_ausgabeFile.Text)

    'doc.Activate()
    'Dim vorlage = frm_main.getTextBaustein

    With wDoc.ActiveWindow
      ' .ActiveWindow.Selection.EndKey(6)

      .Selection.EndKey(6) 'Word.WdUnits.wdStory

      Dim pageNr As Integer = .Selection.Information(3) 'Word.WdInformation.wdActiveEndPageNumber
      If insertNewPageBefore = True Then .Selection.InsertBreak(Type:=wdPageBreak)
      .Selection.InsertFile(FileName:=docFileName, Range:="", _
          ConfirmConversions:=False, Link:=False, Attachment:=False)
      .Selection.GoTo(wdGoToPage, wdGoToNext, , CStr(pageNr))

    End With
  End Sub

  Sub replacePlaceholder(ByVal wDoc As Object)

    With wDoc.ActiveWindow
      Do
        .Selection.Find.ClearFormatting()
        With .Selection.Find
          .Text = "[["
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
        End With
        If .Selection.Find.Execute() = False Then Exit Do

        Dim pos1 As Long
        pos1 = .Selection.Start

        With .Selection.Find
          .Text = "]]"
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
        End With
        If .Selection.Find.Execute() = False Then Exit Do

        Dim pos2 As Long
        pos2 = .Selection.Start

        .Selection.SetRange(pos1 + 2, pos2)

        Dim keyName As String = .Selection.Text
        Dim content As String = getContentForPlaceholder(keyName)

        .Selection.SetRange(pos1, pos2 + 2)

        .Selection.Text = content
      Loop
    End With
  End Sub



'-->
'--> --- DB(neu) --------------------------------------------------------------------------


'' Module sys_navigateDB
  Public conADR As OleDb.OleDbConnection

Function dbGetSingle(ByVal sql As String) As Object
    'Dim con = getDBConnection()

    Dim cm2 As New System.Data.OleDb.OleDbCommand
    cm2.Connection = conADR

    trace sql 
    cm2.CommandText = sql

    ''Try
      dbGetSingle = cm2.ExecuteScalar()
    ''Catch ex As Exception
      '' ??? MsgBox("Fehler bei der Abfrage:" + vbNewLine + sql + vbNewLine + vbNewLine + ex.Message)
    ''End Try
  End Function


  Function readLineAsTabArray(ByVal dataReader As OleDb.OleDbDataReader) As String
    If dataReader.IsClosed Or dataReader.HasRows = False Then Return ""
    Dim array(dataReader.FieldCount - 1) As String
    For i As Integer = 0 To dataReader.FieldCount - 1
      array(i) = dataReader.GetString(i)
    Next
    Return Join(array, vbTab)
  End Function
  
  Function readLineAsArray(ByVal dataReader As OleDb.OleDbDataReader) As String()
    If dataReader.IsClosed Or dataReader.HasRows = False Then Return New String() {}
    Dim array(dataReader.FieldCount - 1) As String
    For i As Integer = 0 To dataReader.FieldCount - 1
      array(i) = dataReader.GetString(i)
    Next
    Return array
  End Function

  Function dbGetReader(ByVal sql As String) As OleDb.OleDbDataReader
    Dim cm2 As New System.Data.OleDb.OleDbCommand
    cm2.Connection = conADR

    trace sql
    cm2.CommandText = sql

    '' Try
      dbGetReader = cm2.ExecuteReader(CommandBehavior.Default)
    '' Catch ex As Exception
       '' ??? MsgBox("Fehler bei der Abfrage:" + vbNewLine + sql + vbNewLine + vbNewLine + ex.Message)
    '' End Try
  End Function

  Function dbExecuteSQL(ByVal sql As String) As Integer
    Dim cm2 As New System.Data.OleDb.OleDbCommand
    cm2.Connection = conADR

    trace sql 
    cm2.CommandText = sql

    ''Try
      Return cm2.ExecuteNonQuery()
    ''Catch ex As Exception
    '' ???  MsgBox("Fehler bei der Abfrage:" + vbNewLine + sql + vbNewLine + vbNewLine + ex.Message)
    '' End Try
  End Function

  Sub openDBConnection(dbFilespec As String)
    If dbFilespec = "" Then MsgBox("bitte Pfad zur Datenbankdatei angeben") : Exit Sub
    Dim connectionString As String
    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbFilespec + ";User ID=Admin;Password=;"
    
    conADR = New OleDb.OleDbConnection(connectionString)
    conADR.Open()
  End Sub
  
  
  Sub closeDBConnection()
    If conADR Is Nothing Then Exit Sub
    conADR.Close()
  End Sub

  Sub FillDbGrid(ByVal grid As TenTec.Windows.iGridLib.iGrid, ByVal sql As String)
    'Dim con = getDBConnection()
    
    dim myGrid 
    myGrid=DirectCast(grid, Object)
    Dim cm2 As New System.Data.OleDb.OleDbCommand
    cm2.Connection = conADR

    trace sql
    cm2.CommandText = sql

    '' Try
      myGrid.FillWithData(cm2, False)
    '' Catch ex As Exception
      ''MsgBox("Fehler bei der Abfrage:" + vbNewLine + sql + vbNewLine + vbNewLine + ex.Message)
    '' End Try


    ''grid.Cols(0).Width = 110
    ''grid.Cols(1).Width = 60
    ''grid.Cols(2).Width = 220
    ''grid.Cols(3).Width = 220

    ''...warum geht der nicht mehr ???
    'For Each r As TenTec.Windows.iGridLib.iGRow In grid.Rows
    '  r.Cells(1).BackColor = Color.Lavender
    'Next
    'If grid.Rows.Count > 0 Then
    '  grid.SetCurRow(0)
    'End If
  End Sub
  
Public Shared Function NotNull(Of T)(ByVal Value As T, ByVal DefaultValue As T) As T
        If Value Is Nothing OrElse IsDBNull(Value) Then
                Return DefaultValue
        Else
                Return Value
        End If
End Function

'DBNull2EmptyString
Public Shared Function dbn2es(ByVal Value As Object) As String
        If Value Is Nothing OrElse IsDBNull(Value) Then
                Return ""
        Else
                Return Value
        End If
End Function







'-->
'--> --- DB -------------------------------------------------------------------------------

  Private dbADR As Object
  Private rsADR As Object 'RecordSet
  Private Const tableName = "UGB-ADR-TAB" 'ohne Klammer!!!

  Private Function dbFileSpec() As String
    dbFileSpec = "U:\" + "SVS\DB\es-adr-test2.mdb"
  End Function
  Function zoomAdr(ByVal iNummer As String) As Object 'RecordSet
    On Error Resume Next
    Dim Modus
    Dim ctlName As String
    Dim index As Integer
    Dim dbFeldName As String
    Modus = "doppeltGeblockt"

    navigateADR(rsADR, iNummer, Modus)

    zoomAdr = rsADR 'geht das ???
  End Function
  Sub testAdr()
    On Error Resume Next
    Dim iNummer
    Dim Modus
    Modus = "doppeltGeblockt"
    iNummer = "1111"
    navigateADR(rsADR, iNummer, Modus)
    navigateADR(rsADR, 1110, Modus)
    navigateADR(rsADR, 1100, Modus)
    Trace rsADR.Fields.Count
    'closeDbAdr
  End Sub
  Sub closeDbAdr()
    dbClose(dbADR, rsADR)
  End Sub


  Sub navigateADR(ByRef RS, ByVal iNummer, ByVal Modus)
    On Error Resume Next
    'Stop
    Dim SQL
    Dim iNummerID
    If iNummer = "" Then Exit Sub
    If Len(iNummer) > 8 Then Exit Sub
    If InStr(iNummer, ".") > 0 Then Exit Sub
    '...oder generell ausschalten abhängig von Typ ???


    iNummer = Replace(iNummer, "i-", "")
    iNummerID = (iNummer)
    Dim ss
    ss = ""
    ss = ss & "SELECT *"
    ss = ss & " FROM [UGB-ADR-TAB]"
    ss = ss & " WHERE"
    ss = ss & " ((([UGB-ADR-TAB].ID) =" & iNummerID & "))"
    ss = ss & " ORDER BY [UGB-ADR-TAB].NAME, [UGB-ADR-TAB].VORNAME "
    ss = ss & " ;"
    SQL = ss
    Trace "SQL:" + SQL


    'dbAdr muß offen sein
    '=======================================
    'trace "check, ob dbADR offen ist" dbADR.name  'wie testen ohne "nothing"

    If dbADR Is Nothing Then
      DBopen(dbADR, dbFileSpec, tableName)
    End If
    '=======================================



    'Stop
    navigateSQL(dbADR, RS, SQL, Modus)

    Trace "navigateADR..."
    'Debug.Print(RS.Fields(1))
    'TRACE "xxxxxxxxxxxxxx --- wo bin ich ???"
    'Debug.Print(dbADR.ToString)
  End Sub



  '#########################################################
  '#########################################################
  '#########################################################


  Sub navigateSQL(ByRef db As Object, ByRef RS As Object, ByVal SQL As String, ByVal Modus As String)
    'Stop
    'On Error Resume Next
    'db muß bereits offen sein
    : RS.Close()    : Err.Clear()
    Try
      RS = CreateObject("ADODB.recordset")
    Catch ex As Exception
      MsgBox("(es) --> navigateSQL: ..." + Err.Description)
      Err.Clear()
      Exit Sub
    End Try
    With RS
      Dim cursorType
      Dim lockType
      Dim cursorLocation

      Modus = "READ"
      If Modus = "READ" Then
        cursorType = 3 'adOpenStatic
        lockType = 1   'adLockReadOnly
      Else
        cursorType = 3 'adOpenStatic
        lockType = 3   'adLockOptimistic
      End If
      'cursorType =  0 'adOpenForwardOnly
      'cursorType =  1 'adOpenKeyset
      'cursorType =  2 'adOpenDynamic
      'cursorType =  3 'adOpenStatic

      'lockType    = 1 'adLockReadOnly
      'lockType    = 2 'adLockPessimistic
      'lockType    = 3 'adLockOptimistic
      'lockType    = 4 'adLockBatchOptimistic
      .cursorType = cursorType      '
      .lockType = lockType            '
      .cursorLocation = 1  '??? Client
      cursorLocation = 1
      ''''.source=db
      ''''.Open ss, dbZE,3 ,3 , 1   ' adCmdText

      ''.Open  SQL, db, cursorType ,lockType , CursorLocation   ' adCmdText
      Try
        .Open(SQL, db)   ', dbRead, cursorType ,lockType , CursorLocation   ' adCmdText
      Catch ex As Exception
        Trace "navigateSQL: ", Err.Description
        'Stop
        ' errlog "SQL: ", SQL
        ' errlog "", ""
      End Try
      'Stop
    End With
  End Sub

  ''==========================
  ''_ with rsZE
  ''_     .CursorType = 3  '
  ''_                     '0 adOpenForwardOnly
  ''_                     '1 adOpenKeyset
  ''_                     '2 adOpenDynamic
  ''_                     '3 adOpenStatic
  ''_
  ''_   .LockType =  3   '
  ''_                     '1 adLockReadOnly
  ''_                     '2 adLockPessimistic
  ''_                     '3 adLockOptimistic
  ''_                     '4 adLockBatchOptimistic
  ''_   .CursorLocation=1 '??? Client
  ''_   .source=dbZE
  ''_     .Open ss, dbZE,3 ,3 , 1   ' adCmdText
  ''_
  ''==========================


  '#########################################################
  '#########################################################
  '#########################################################


  Sub DBopen(ByRef db As Object, ByVal dbFileSpec As String, ByVal tableName As String)
    'On Error Resume Next
    'Stop
    Dim myPath : myPath = dbFileSpec + ""
    Trace "myPath???", "file:" + myPath
    Dim SQL
    Dim RS = Nothing

    SQL = "SELECT * FROM [" + tableName + "] Where 1=2"
    Trace "SQL", SQL
    Trace "START DB_open..  ", Now()

    
:   db = CreateObject("ADODB.connection")
:   db.Open("Driver={Microsoft Access Driver (*.mdb)};DBQ=" & myPath)
':   If Err.Number<>0 Then Trace "ErrDbOpen: ", Err.Description
:   Err.Clear()
    
    navigateSQL(db, RS, SQL, "READ")

    '___Felder anzeigen lassen ...
    With RS
      Dim i, maxFields
      maxFields = .Fields.Count
      For i = 0 To maxFields - 1
        Trace i, .Fields(i).Name
        'logfile CStr(i) + "  " + .fields(i).name
      Next
      Trace "maxFields", maxFields
    End With
    '  Stop
  End Sub


  '#########################################################
  '#########################################################
  '#########################################################


  Sub dbClose(ByRef db As Object, ByRef RS As Object)
    On Error Resume Next
    : RS.Close() : Err.Clear()
    : RS = Nothing : Err.Clear()
    : db = Nothing : Err.Clear()
  End Sub


'-->
'--> IgridHelper


  Shared Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
    Dim max = line.Cells.Count - 1
    Dim out(max) As String
    For i As Integer = 0 To max
      out(i) = CStr(line.Cells(i).Value)
    Next
    Return Join(out, Delimiter)
  End Function

  Shared Sub SplitToIGridLine(ByVal line As iGRow, ByVal input As String, Optional ByVal Delimiter As String = vbTab)
    Dim max = line.Cells.Count - 1
    Dim out() = Split(input, Delimiter)
    ReDim Preserve out(max)
    For i As Integer = 0 To max
      line.Cells(i).Value = out(i)
    Next
  End Sub

  Shared Sub Igrid_get(ByVal Grid As iGrid, ByRef FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
    Dim max = Grid.Rows.Count - 1
    Dim out(max) As String
    For i As Integer = 0 To max
      out(i) = JoinIGridLine(Grid.Rows(i), ColDelim)
    Next
    FileContent = Join(out, LineDelim)
  End Sub
  
  Shared Sub Igrid_put(ByVal Grid As iGrid, ByVal FileContent As String, Optional ByVal LineDelim As String = vbNewLine, Optional ByVal ColDelim As String = vbTab)
    Dim out() = Split(FileContent, LineDelim)
    Grid.Rows.Clear()
    Grid.Rows.Count = out.Length
    For i As Integer = 0 To out.Length - 1
      SplitToIGridLine(Grid.Rows(i), out(i), ColDelim)
    Next
  End Sub
  


