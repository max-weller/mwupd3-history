
#Para DebugMode Internal

'U:\zPara\script4\i-ST-EtikettenVorlage.txt



Sub ListHeader(simSalaBim, headerData, dsData, listCount, dsFileTime)
  With simSalaBim
    Trace "<<<<<< listHeader >>>>>>>>>>"
    .iniPicUpLine()
    .picUpLine ("1.Zeile: ...ich werde von labelPrinter nicht ausgegeben ")
    .picUpTab ("==============================")

    .picUpTab ("now     "+cStr(now))
    .picUpTab ("DsName: ")
    .picUpTab ("(name)  ")
    .picUpTab ("(Verzeichnis)")
    .picUpTab ("Anzahl: ... "+cStr(listCount))

    .picUpTab ("==============================")
    .picUpTabComplete()
    '''.picUpLine "Anzahl: "+cStr(listCount)
    '''.picUpLine ""
  End With
End Sub



Sub ListBodyLoop(simSalaBim, lineData, dsData, index, Cancel)
  With simSalaBim
    Trace "ListBody", .jjj
    
    Dim text As String
    
    text = "|" + prepareText(.jjj) + "<br>" + .hhh + " " + prepareText(.iii) + "|" + prepareText(.kkk) + "|"
    Trace "Data:",text
    .picUpLine(text)
    
  End With
End Sub

Function prepareText(text As String) AS String
  If String.IsNullOrEmpty(text) Then Return ""
  text = text.Trim("""")
  text = Replace(text, vbNewLine, vbLF)
  text = Replace(text, vbLF, "<br>")
  Dim abPos = text.IndexOf("-")
  if abPos>-1 Then
    trace "Schnipp:",text.Substring(0,abPos+1)
    text = text.SubString(abPos+1).Trim
  End If
  Return text
End Function




Sub ListFooter(simSalaBim, headerData, dsData, listCount, dsFileTime)
  With simSalaBim
    Trace "<<<<<< listFooter >>>>>>>>>>"
    '.picUpLine "Ausdruck: " +cStr(date)+"   "+cStr(time)
  End With
End Sub


Sub ListOut(simSalaBim, OutData, dsData, UserAction, UserFileSpec, cancel)
  With simSalaBim
    ZZ.setClipboardText(OutData)
    
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


