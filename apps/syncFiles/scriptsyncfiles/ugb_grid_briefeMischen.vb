'-->
'--> --- GLOB -----------------------------------------

#Reference siaGridView.dll
#Para DebugMode Internal

'U:\zPara\script4\i-ST-EtikettenVorlage.txt

'' Dim docFileName As String :docFileName = "V:\00_Pressearbeit\Aktion_Probeabo\baustein_Fastenklinik.doc"
'' Dim outFile As String     :outFile = "V:\00_Pressearbeit\Aktion_Probeabo\AUSGABE-1.doc" 'monitor_ausgabeFile.Text
'' 
'' Dim docFileName As String    = "V:\00_Pressearbeit\Aktion_Probeabo\baustein_Wellnesshotels.doc"
'' Dim outFile As String        = "V:\00_Pressearbeit\Aktion_Probeabo\AUSGABE-2.doc" 'monitor_ausgabeFile.Text
'' 
Dim docFileName As String    = "V:\00_Pressearbeit\Aktion_Probeabo\baustein_fastenärzte.doc"
Dim outFile As String        = "V:\00_Pressearbeit\Aktion_Probeabo\AUSGABE-3.doc" 'monitor_ausgabeFile.Text




'-->
'--> FOR-EACH-GRID-LINE...

Sub ListHeader(simSalaBim, headerData, dsData, listCount, dsFileTime)
  With simSalaBim
    Trace "<<<<<< listHeader >>>>>>>>>>"
    
    If wApp Is Nothing Then
      createWordApp()
    End If
    
    ''wApp.Visible = False
    
    
    
    
    curDoc = getWordDoc(outFile) ''("briefeSchreibenTest.doc")
    
    'curDoc.ActiveWindow.Visible = False
    curDoc.ActiveWindow.Selection.WholeStory()
    curDoc.ActiveWindow.Selection.Delete(wdCharacter, 1)
    
    
    '.picUpTab ("now     "+cStr(now))
    '.picUpTab ("Anzahl: ... "+cStr(listCount))
    
    
  End With
End Sub



Sub ListBodyLoop(simSalaBim, lineData, dsData, index, Cancel)
  With simSalaBim
    Trace "ListBody", .bbb & " " & .ccc
    
    'actInummer = iNummer
    'Debug.Print(i & " = iNummer " & grid.CellValue(i, 1))
    insertDocument(curDoc, docFileName, False)
    replacePlaceholder(curDoc, simSalaBim)

  End With
End Sub





Sub ListFooter(simSalaBim, headerData, dsData, listCount, dsFileTime)
  With simSalaBim
    Trace "<<<<<< listFooter >>>>>>>>>>"
    
    curDoc.ActiveWindow.Visible = True
    curDoc.Activate()
    
  End With
End Sub


Sub ListOut(simSalaBim, OutData, dsData, UserAction, UserFileSpec, cancel)
  With simSalaBim
    'ZZ.setClipboardText(OutData)
    
    '' Trace "<<<<<< listOut >>>>>>>>>>"
    '' 'setClipboardText outData
    '' 
    '' 
    '' forSave =outData
    '' fileSpec=UserFileSpec 
    '' fileSpec = "U:\zPara\SQL\ds-Dateien\llxxx.txt"
    '' saveFile fileSpec, forSave, errMes
    '' trace forSave
    '' trace fileSpec
    '' trace errmes
    '' 
    '' call sendToLabelPrinter(fileSpec)
    '' 
    '' cancel = True

    '' .sendToLabelPrinter
    '' .sendToClipboard
    'cancel = false
    ''-----------------------------
    '' .showWithNotePad
    '' .showWithExcel
    '' .showAsGrid
    '' .sendToLabelPrinter
  End With
End Sub



'-->
'--> --- getPlaceHolder ----------------------

  Public actInummer As String
  Public globRS As Object

  Function getContentForPlaceholder(ByVal key As String, gridInfo As Object) As String
    On Error Resume Next
    If key.Startswith("#") Then
      Dim field = key.SubString(1)
      Dim fieldIndex As Integer
      If Not Integer.TryParse(field, fieldIndex) Then
        fieldIndex = Asc(UCase(field)) - 65
        
      End If
      Return gridInfo.Col(fieldIndex)
    End If
    
    Static nr As Integer = 0
    Static lastInummer As String
    Trace lastInummer
    Trace actInummer
    ''lastInummer = ""
    If lastInummer <> actInummer Then
      navigateINummer(actInummer)
      lastInummer = actInummer
    End If

    Dim RESULT As String = ""
    ''Debug.Write(actInummer) : Debug.Print(key)
    Dim keyValue As String = "EMPTY"
    keyValue = globRS.Fields(key).value()
    '' Debug.Print(keyValue)
    If keyValue <> "EMPTY" Then
      RESULT = keyValue
    Else
      RESULT = "x" & actInummer & "x" & key & "x" & nr
      nr = nr + 1
    End If

    Return RESULT

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

'' 
''   Function getIgridRef() As Object
''     If igridRef Is Nothing Then
''       Dim sql = CreateObject("sqlMonitor.sqlScriptServer")
''       igridRef = sql.gridRef
''     End If
'' 
''     Return igridRef
''   End Function
'' '' 
''   Sub readFromSqlMonitor()
''     On Error Resume Next
''     Dim grid = getIgridRef()
''     '' Dim wDoc = getWordDoc("briefeSchreibenTest.doc")
''     Dim outFile As String
''     outFile = frm_main.monitor_ausgabeFile.Text
''     Dim wDoc = getWordDoc(outFile) ''("briefeSchreibenTest.doc")
'' 
''     Debug.Print(grid.RowCount)
'' 
''     Dim iNummer As String
''     wDoc.Visible = False
''     Dim insertPageBreakBefore As Boolean = False
''     For i As Integer = 1 To grid.RowCount
''       iNummer = grid.CellValue(i, 1)
''       grid.CurRow = i
'' 
''       actInummer = iNummer
''       Debug.Print(i & " = iNummer " & grid.CellValue(i, 1))
'' 
''       insertDocument(wDoc, frm_main.monitor_textbaustein.Text, insertPageBreakBefore)
''       insertPageBreakBefore = True
''       replacePlaceholder(wDoc)
'' 
'' 
''     Next
''     wApp.Visible = True
'' 
''     'Stop
''     'For i=
'' 
'' 
''   End Sub
'' 




'-->
'--> --- WORD -------------------------------------

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

:   wApp.Visible = False : Err.Clear

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

  Sub replacePlaceholder(ByVal wDoc As Object, gridInfo As Object)

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
        Dim content As String = getContentForPlaceholder(keyName, gridInfo)

        .Selection.SetRange(pos1, pos2 + 2)

        .Selection.Text = content
      Loop
    End With
  End Sub


'-->
'--> --- DB ---------------------------------------------

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





