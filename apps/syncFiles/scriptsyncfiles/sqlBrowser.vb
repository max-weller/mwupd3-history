
#Para SilentMode false
#Para DebugMode intern

'#Reference ScriptIDE.AddinIfaces.dll
'#Imports ScriptIDE.AddinIfaces

#reference Microsoft.visualbasic.dll
' #imports microsoft.visualbasic
#reference TenTec.Windows.iGridLib.iGrid.dll
'#imports Tentec.Windows.Igridlib


#imports System.Xml

Dim WithEvents FormRef As Form
Dim WithEvents PanelRef As Object
dim frm as object
Dim WithEvents IGrid1 As Igrid
'' dim dbFileSpec_ADR as string = "U:\SVS\DB\es-adr-test2.mdb"
'' dim dbFileSpec_OPOS as string = "U:\Buchhaltung\opos\opos01.mdb"
dim dbFileSpec_ADR as string =  "C:\_Alias\u\SVS\DB\es-adr-test2.mdb"
dim dbFileSpec_OPOS as string = "C:\_Alias\u\Buchhaltung\opos\opos01.mdb"

'##############################

'##############################

'-->
Sub AutoStart()

  PanelRef = ZZ.IDEHelper.CreateDockWindow("sqlBrowser.app")
  formRef=panelRef.form
  FormRef.Text = "es_sqlBrowser"
  
  ''frm = DirectCast(FormRef, Object).PNL
  frm=panelRef
  With frm
    .resetControls (10,5)
    .addButton  ("openDb"  , "DB öffnen"     , "#ccc")
    .addTextbox ("inummer" ,  200            , "iNummer:"    , 50  , "#aaa")
    .addButton  ("sql"     , "Abfrage"       , "#ccc")
    .addButton  ("closeDb" , "DB schließen"  , "#ccc")
    .br
    .activateEvents = "|ButtonMouseClick|"
    .addButton ("read"    , "Lesen"      , "#ccc")
    .addButton ("save"    , "Speichern"  , "#ccc")
    .activateEvents = "|ButtonMouseClick| |ICONMOUSECLICK|"

 End With
  
  Dim x As integer=0
'  x=5/x   '...testfehler
  
  IGrid1 = New IGrid
  IGrid1.Top = 70
  'leere Zelle als "" statt nothing
  IGrid1.Cols.Add("id")
  IGrid1.Cols.Add("nickname")
  IGrid1.Cols.Add("gruppe")
  IGrid1.Cols.Add("codelink")
  IGrid1.GroupBox.Visible = True
  Igrid1.rows.count = 100
  frm.Controls.Add(IGrid1)
  onResizeControls()
End Sub

'-->
Sub onButtonEvent(e As Object)
  dim sName as string: sName=e.Sender.Name
  printLine  1, "ToolbarButtonClick", sName
  printLine  2, sName , "..........."
  If sName = "btn_read" Then
    Dim Content = ZZ.fileGetContents("C:\yPara\scriptIDE\codeBookmarks.txt")
    IgridHelper.Igrid_put (IGrid1, Content)
    printLine  2, sName , "OK"
  End If
  If sName = "btn_save" Then
    Dim Content As String
    IgridHelper.Igrid_get (IGrid1, Content)
    ZZ.filePutContents("C:\yPara\scriptIDE\codeBookmarks.txt", Content)
    printLine  2, sName , "OK"
  End If
  If sName = "btn_openDb" Then
    openDBConnection()
    printLine  2, sName , "OK"
    printLine  3, "DBConnection", "open"
    'msgBox(sName)
  End If
  If sName = "btn_sql" Then
    printLine  4, "Abfrage", "Start...."
    '' msgBox(sName)
    executeOposAbfrage()
    printLine  4, "Abfrage", "OK"
    printLine  2, sName , "OK"
    ''msgBox(sName)
    End If
  If sName = "btn_closeDb" Then
    closeDBConnection()
    ''msgBox(sName)
    printLine  2, sName , "OK"
  End If
End Sub

Sub Form_Resize(sender as Object, e as EventArgs) Handles formRef.Resize
  onResizeControls
End Sub

Sub onResizeControls()
  Igrid1.Width = formRef.Width - 10
  Igrid1.Height = formRef.Height - 100
  
End Sub


sub executeOposAbfrageOld()
   dim sql as string
   dim iNummer as string
   iNummer=frm.Controls("txt_inummer").Text
   ''msgBox (iNummer)
   sql=getSqlStatement1(iNummer)
   ''msgBox(sql)
   FillDbGrid(IGrid1, sql)
end sub

sub executeOposAbfrage()
   dim sql as string
   dim iNummer as string
   iNummer=frm.Controls("txt_inummer").Text
   ''msgBox (iNummer)
   sql=getSqlStatement1(iNummer)
   ''msgBox(sql)
   FillDbGrid(IGrid1, sql)
   
   '5 und 8
   with IGrid1
     dim max as integer
     dim i as integer
     dim betrag1
     dim betrag2
     max=IGrid1.rows.count
     printLine 3, "Anzahl Treffer" ,max
     for i = 0 to max-1
       betrag1=.cells(i,6).value
       betrag2=.cells(i,8).value
      if typeOf(betrag1) is DBNull then betrag1=""
      if typeOf(betrag2) is DBNull then betrag2=""
      if betrag1.trim() <>betrag2.trim() then 
         .cells(i,6).BackColor=ColorTranslator.FromHtml("#f66")
         'msgbox (betrag1.toString)
         'msgbox (betrag2.toString)
       end if
     next
     'msgbox(max)
   end with
end sub




Function getSqlStatement1(iNummer as string) as string
  dim q as string: q=""
  dim LF as string: LF=vbNewLine

  q=q+LF+""
  q=q+LF+"SELECT "
  q=q+LF+"  tab_opos.iNummer"
  q=q+LF+" , tbZE.iNummer"
  q=q+LF+" , tab_opos.ID"
  q=q+LF+" , refZE"
  q=q+LF+""
  q=q+LF+" , kurzanschrift"
  q=q+LF+" , reDatum"
  q=q+LF+" , BetragEuro"
  q=q+LF+" , neuerBetrag"

  q=q+LF+" , BetragZE"
  q=q+LF+" , BetragRest"
  q=q+LF+" , ZEKom1"
  q=q+LF+" , frei1 "
  q=q+LF+" , tbZE.kontoName"

  q=q+LF+" , tab_opos.KontoName "
  q=q+LF+" , tabKtr.zusatzNummer "
  q=q+LF+" , tabKtr.kontoNummer"
  q=q+LF+" , kommentar"
  q=q+LF+" , bezahlArt"
  q=q+LF+" , bankeinzug"
  q=q+LF+" , errinerung"
  q=q+LF+" , mahnung"
  q=q+LF+" , storno"
  q=q+LF+""
  q=q+LF+""
  q=q+LF+" FROM ([tab_opos]"
  q=q+LF+" LEFT JOIN tabKtr"
  q=q+LF+"   ON tab_opos.KontoName = tabKtr.KontoName)"
  q=q+LF+"LEFT JOIN tbZE"
  q=q+LF+"   ON tab_opos.reNummer = tbZE.reNummer"
  q=q+LF+""
  q=q+LF+"WHERE"
  q=q+LF+"( tab_opos.iNummer = '||INUMMER||')" '60358
  q=q+LF+"OR"
  q=q+LF+"(tbZE.iNummer = '||INUMMER||')"
  q=q+LF+""
  q=q+LF+"ORDER BY "
  q=q+LF+"  tab_opos.reDatum"
  q=q+LF+""
  q=replace(q, "||INUMMER||",iNummer)
  return q
End Function



Class IgridHelper

  Shared Function JoinIGridLine(ByVal line As iGRow, Optional ByVal Delimiter As String = vbTab) As String
    Dim max = line.Cells.Count - 1
    Dim out(max) As String
    For i As Integer = 0 To max
      out(i) = line.Cells(i).Value.ToString
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
  
End Class




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

  Sub openDBConnection()
    If dbFileSpec_ADR = "" Then MsgBox("bitte Pfad zur Datenbankdatei angeben") : Exit Sub
    Dim connectionString As String
    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbFileSpec_ADR + ";User ID=Admin;Password=;"
    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbFileSpec_OPOS + ";User ID=Admin;Password=;"

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



'' End Module
