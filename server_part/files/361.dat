

'''-->
'--> NEU: Imports...

'#Reference System.Windows.Forms.dll
'#Reference System.Data.dll
'#Reference System.Drawing.dll
'#Reference WeifenLuo.WinFormsUI.Docking.dll
'#Reference TenTec.Windows.iGridLib.iGrid.dll

'#Imports System.Windows.Forms
'#Imports System.Data
'#Imports System.Drawing
#Imports ScriptIDE.Core
#Imports ScriptIDE.ScriptHost
'#Imports ScriptIDE.ScriptWindowHelper
'#Imports Tentec.Windows.iGridLib

'#imports System.IO
'#imports Microsoft.VisualBasic.Strings
'#imports System.Collections.Generic


'_________ logFile(FSO)

'' dim c_TestCounter
'' Property Let TestCounter(pTestCounter)
'' 	c_TestCounter = pTestCounter
'' End Property
'' '- - - - - - - - - - - 
'' Property Get TestCounter()
'' 	TestCounter = c_TestCounter
'' End Property
'' '- - - - - - - - - - - - - - - - - - - - - - -
'' 
sub autostart
  msgBox("hello")
  logFileOpen
  logFile("aaaaaaaaaa")
  logFileAdd("bbbbbbbbbbb")
  logFile("ccccccccccccc")
  logFileShow
  'showAsTextFile
  'showAsGrid("C:\es-logFile.txt")
end sub


'--> ----------------------------------------------


private g_logFileFileSpec as string = "C:\es-logFile.txt"
: Property logFileFileSpec () As String
:   Get 
:     Return g_logFileFileSpec
:   End Get
:   Set (ByVal value As String)
:     g_logFileFileSpec=value
:   End Set
: End Property


sub logFileOpen
  fWrite(logFileFileSpec, "=== "+cStr(now)+" ==="+vbCrLf)
end sub


sub logFileAdd(addData)
  fAppend (logFileFileSpec, addData)
end sub


sub logFile(addData)
  fAppend (logFileFileSpec, addData)
end sub


function logFileClose
  logFileClose= fRead(logFileFileSpec)
end function


sub logFileShow
  fAppend( logFileFileSpec, "+++ "+cStr(now)+" +++"+vbCrLf)
  showAsTextFileWSX (logFileFileSpec)
  'showAsTextFile(logFileFileSpec)
end sub



'--> ----------------------------------------------

Function fRead(FileSpecRead)
  'result = fRead(FileSpecRead)
  'Given the path to a file, will return entire contents
  ' works with either ANSI or Unicode                  
  Const ForReading = 1
  const TristateUseDefault = -2
  const DoNotCreateFile = False
  With CreateObject("Scripting.FileSystemObject")
    with .OpenTextFile(FileSpecRead, ForReading, 	False, TristateUseDefault)
      fRead = .ReadAll
      .Close
    end with
  End With
End Function


Sub fWrite(FileSpecSave, sData)
  'fWrite FileSpecSave, sData
  'writes sData to FilePath
  Const ForReading = 1
  const TristateUseDefault = -2
  const DoNotCreateFile = False
  With CreateObject("Scripting.FileSystemObject")
    with .OpenTextFile(FileSpecSave, 2, True)
      .Write (sData)
      .Close
    end with
  End With
End Sub


Sub fAppend(fileSpecAppend, sData)
  'fAppend fileSpecAppend, sData
  'Given the path to a file, will append sData to it
  Const ForReading = 1
  const TristateUseDefault = -2
  const DoNotCreateFile = False
  With CreateObject("Scripting.FileSystemObject")
    with .OpenTextFile(fileSpecAppend, 8)
	  .Write (sData+vbCrLf)
	  .Close
    end with
  End With 
End Sub



sub showAsTextFile(fileSpec)
  shell (fileSpec)
end sub


sub showAsTextFileWSX(fileSpec)
  'showAsTextFileWSX fileSpec
  'file:C:\Programme\NoteTab#Pro\NotePro.exe
  dim apo:apo=chr(34)
  with CreateObject("WScript.Shell")
    .Run ("notepad.exe " & fileSpec)
    '.Run apo+"C:\Programme\NoteTab Pro\NotePro.exe"+apo+"  "& fileSpec
  end with
end sub


sub showAsGrid(fileSpec)
  
  '... das wäre durchaus praktisch

  'showAsTextFileWSX fileSpec
  'file:C:\Programme\NoteTab#Pro\NotePro.exe
  dim apo:apo=chr(34)
  with createObject("Grid01.application")
    'set .ScriptCallBack = me

    dim nicName: nicname="esNicName"
    with .newGrid(NicName)
      .navigate (fileSpec)
      '.navigate "C:\Eigene Dateien\es-testOut2.TXT"
      .show
    end with
    'if Err then trace cStr(qID), "ERR: "+err.description
  end with
end sub



'--> ----------------------------------------------


'' sub Open
''   dim fileSpec
''   fileSpec="C:\es-logFile.txt"
''   fWrite (FileSpec, "=== "+cStr(now)+" ==="+vbCrLf)
'' end sub
'' 


'' function Close
''   dim fileSpec:fileSpec="C:\es-logFile.txt"
''   Close= fRead(FileSpec)
'' end function
'' 

'' sub Show
''   dim fileSpec:fileSpec="C:\es-logFile.txt"
''   fAppend (fileSpec, "+++ "+cStr(now)+" +++"+vbCrLf)
''   showAsTextFileWSX (fileSpec)
'' end sub
'' 


'' function callBack_fileFind (Sender, typ, Index, lineData, splitterField)
''   'if index mod 17 =0 then printLine 1, typ, index
''   'if index mod 13 =0 then trace "myX "+typ+"  "+cStr(index), Mid(LineData,1,25)
''   trace "myX "+typ, Mid(LineData,1,50)
'' end function
'' 

'' function callBack_grid (Sender, typ, Index, lineData, splitterField)
''   'trace "callBackGrid ???"
''   trace "grid: "+replace(lineData,vbTab,"|")
''   dim fileSpec
''   dim data
''   data=split(lineData,splitterField)
''   filespec=data(5)+""+data(3)+"."+data(4)	
''   sendPreviewMessage(FileSpec)
''   trace fileSpec
'' end function
'' 


