private parent as Object


#Para DebugMode intern
#Para SilentMode true

#reference Microsoft.visualbasic.dll
'' #imports microsoft.visualbasic
#reference TenTec.Windows.iGridLib.iGrid.dll
'' #imports Tentec.Windows.Igridlib



Dim WithEvents FormRef2 As Form
Dim ref2 As Object
shared winId2 as string ="es_snippetEditor.vb"
dim curIGridRow as integer=-1
'' shared winId2 as string ="bookmark.vb"


public sub setParent(myParent)
  parent=myParent
end sub

public sub hideSnippetEditor()
  printLine 1, "hideSnippetEditor", "...hide"
  with formRef2.parent.parent
    .hide()
  end with
end sub



Sub GetFormRef2()
  'If ref IsNot Nothing Then Exit Sub
  ref2 = ZZ.IDEHelper.CreateDockWindow(winID2,4)
  formRef2=ref2.form
  FormRef2.Text = "SnippetBearbeiten"
End Sub

'-->
'--> F O R M - 2 

Sub AutoStart()
  '' createEditWindow(-1)
end sub


sub makeJumboLabel(el)
    el.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    ''el.Size = New System.Drawing.Size(117, 37)
    el.backColor=ColorTranslator.FromHtml("#777")
    el.AutoSize = false
    el.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
end sub


public sub createEditWindow(index as integer)
  ''msgBox ("hallo")
  GetFormRef2()
  dim nicName as string=""
  dim code as string =""

  
  curIGridRow=index
  if index > -1 then
    dim ig as IGrid = parent.IGrid1
    dim ir as iGRow = ig.rows(index)
    nicName =ir.cells(2).value.toString
    code=ir.cells(4).value.toString
    code=replace(code, "||LF||",vbNewLine)  
  end if
  
  With ref2
    dim el as object
    .activateEvents = "|TextboxKeyDown|  |TextboxKeyDown|  |TextboxTextChanged|"

    .resetControls ()
    .insertX = 11 :.insertY = 12
    .addIcon("xxx_appPicture", "http://es.teamwiki.net/docs/icons/insert-object64.png" )
     
    .BR  '--------------------------------------------------
    .insertX = 80  :.insertY = 12
     dim label as string=nicName
     el=.addLabel  ("lab01", label ,  ,"#ddd",,,-2,35) : el.backColor=ColorTranslator.FromHtml("#777"): makeJumboLabel(el)
     
    .BR  '--------------------------------------------------
   .insertX = 5 :.insertY = 70
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("nicName" ,  -2         , " nicName"    , 70  ,"#bbb" , , , )
    ref2.element("nicName").text=nicName
    ref2.element("nicName").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
 
   .insertX = 5 :.insertY = 100
    .activateEvents = "|ButtonMouseClick|"
    el = .addTextbox ("codeTextarea" ,  -2         , "  Code"+vbNewLine+""    , 70 ,"#bbb" , , , "multiline=250")
    ref2.element("txt_codeTextarea").WordWrap=false
    ref2.element("txt_codeTextarea").text=code
    ref2.element("txt_codeTextarea").Font = New System.Drawing.Font("Courier New", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    

     .BR  '--------------------------------------------------
   .insertX = 80 :.insertY = 360
   .addButton ("cmdOK"         , " O K "  , "#1DD910")
     .insertX = 160:
    .addButton ("cmdEscape"    , "Abbrechen"      , "#ccc")

  end with
  with formRef2.parent.parent
    '.left=333
    '.top=0
    .show()
  end with

end sub

'-->
'--> EVENTS

sub nicName_keyDown(e)
  ''....bringtNichts
  printLine 1, "keyDown(e)",  "keyDown(e)"
  ''printLine 2, "keyDown(e)",  "keyDown(e)"
  ''printLine 3, "keyDown(e)",  "keyDown(e)"
  trace 1, "keyDown(e)", e.Keystring
End Sub

sub nicName_TextChanged(e)
  printLine 2, "nicName_TextChanged",  e.Sender.text
  ''printLine 11, "keyDown(e)",  "keyDown(e)"
  ''printLine 22, "keyDown(e)",  "keyDown(e)"
  ''printLine 23, "keyDown(e)",  "keyDown(e)"
  'trace 1, "nicName_TextChanged", e.keyCode
  dim newText as string=e.Sender.text
  setNewTextToIgridCell(2, newText)
  ref2.element("nicName").text=newText
  ref2.element("lab_lab01").text=newText
End Sub


sub codeTextarea_TextChanged(e)
  'printLine 3, "codeTextarea_TextChanged", e.keyCode
  'trace  "codeTextarea_TextChanged", e.keyCode
  dim newText as string=e.Sender.text
  setNewTextToIgridCell(4, newText)
  ref2.element("txt_codeTextarea").text=newText
End Sub


sub setNewTextToIgridCell(colIndex as integer, newText as string)
  dim index as integer=curIGridRow
  if index > -1 then
    dim ig as IGrid = parent.IGrid1
    dim ir as iGRow = ig.rows(index)
    newText=replace(newText, vbNewLine, "||LF||") 
    ir.cells(colIndex).value = newText
   end if 
end sub



sub cmdEscape_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  hideSnippetEditor()
End Sub

sub cmdOK_MouseClick(e)
  printLine 1, "MouseClick(e)", e.Sender.Name
  hideSnippetEditor()
  parent.saveSnippetFile() 
End Sub




