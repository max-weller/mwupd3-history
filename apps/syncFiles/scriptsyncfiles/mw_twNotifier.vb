'--> 
'--> mw_twNotifier

#Reference System.Windows.Forms.dll
#Reference System.Data.dll
#Reference System.Drawing.dll
#Reference siaIDEMain.dll
#Reference WeifenLuo.WinFormsUI.Docking.dll
#Reference TenTec.Windows.iGridLib.iGrid.dll

#Imports System.Windows.Forms
#Imports System.Data
#Imports System.text
#Imports System.Drawing
#Imports System.Drawing.Drawing2D
#Imports System.Drawing.Imaging
#Imports System.Runtime.InteropServices
#Imports ScriptIDE.Core
#Imports ScriptIDE.Main
#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.ScriptWindowHelper
#Imports Tentec.Windows.iGridLib

#Include "sys_htmlToRTF.vb"

'' mw_twNotifier

#runtime EXE


'--> ViewForm1

#WindowData ViewForm1
//	x	y	w	h
			430	320	Form		Width=680|Height=320|Text="TW-Notifier"
	10	10	55	30	Button	btn_LoadData	Text="!!!"
	70	10	55	30	Button	btn_Cancel	Text="Abbruch"|Enabled=false
	140	10	55	20	Label	lab_1	Text="Username"
	140	30	55	30	TextBox	txt_user	Text=""
	200	10	55	20	Label	lab_2	Text="password"
	200	30	55	30	TextBox	txt_pass	Text=""
	260	10	55	20	Label	lab_3	Text="lastTS"
	260	30	55	30	TextBox	txt_lastTS	Text=""
	400	20	160	40	Label	labTime	Text="...Willkommen!"|Font=New Font("Microsoft Sans Serif", 12, FontStyle.Regular, GraphicsUnit.Point)
	580	10	90	30	Button	btn_Ende	Text="Beenden"|Anchor=9|Font=New Font("Segoe UI", 12, FontStyle.Regular, GraphicsUnit.Point)
	10	60	650	220	RichTextBox	richtb	Text="Beenden"|Anchor=15|MultiLine=True

#EndData


'--> GlobVars

Public twLoginuser, twLoginPass, twLoginFullname, twSessID As String


Dim MAIN As ViewForm1
Dim ntfy As NotifyIcon


Dim glob As cls_globPara



sub logIn()
  dim RESULT as boolean
  '' dim userName as string
  '' dim passWord as string
  '' userName="mw"
  '' passWord="**************"
  
  RESULT = doLogin(MAIN.txt_user.Text, MAIN.txt_pass.Text)
  trace "logIn", RESULT
end sub  

Function doLogin(ByVal userName As String, ByVal passWord As String) As Boolean
  Dim postData As String = "u=" + userName + "&p=" + passWord
  Dim RESULT = ZZ.http_post("http://teamwiki.net/php/vb/app_login2.php?", postData)
  Dim lines() = Split(RESULT, vbNewLine)
  ReDim Preserve lines(4)
   If lines(0) = "Login OK" Then
    twLoginuser = userName : twLoginPass = passWord : twLoginFullname = lines(3)
    '' glob.para("twLoginData") = glob.RC4StringEncrypt(userName + ":" + passWord, kData)
    twSessID = lines(2)
    Return True
  Else
    MsgBox("Ungueltige Login-Daten!", MsgBoxStyle.Exclamation)
    Return False
  End If
End Function

Sub btn_LoadData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  logIn()
  
  Dim url As String = "http://teamwiki.net/php/vb/notifier.php?m=get_list&dir_cond=all"
  Dim RESULT As String = ZZ.http_get(url, "twnetSID="+twSessID)
  MAIN.richtb.Text=RESULT
  Dim Lines() As String = Split(RESULT, vbLF)
  Dim Parts() As String
  Dim out as New StringBuilder()
  'MAIN.igView.Rows.Clear
  'MAIN.igView.Rows.Count = Lines.Length-5
  TRACE "ResCount", Lines.Length
  For i As  Integer =4 To Lines.Length-1
    '' TRACE "line",i
    TRACE "line" & i,Lines(i)
    Parts = Split(Lines(i), vbTab)
    If Parts.Length<12 Then out.Append("<br>???<br>"):Continue for
    out.Appendline("<font size='14'>Von "+Parts(2)+", "+Parts(12)+", "+Parts(3)+"</font><br>")
    out.Appendline(""+Parts(11)+"<br><br>")
    '' ZZ.SplitIGridLine(MAIN.igView.Rows(i-4))
  Next
  TRACE "htmlCode",out.tostring
  'MAIN.richtb.HTMLCode= out.toString()
  zoomRTF(MAIN.richtb, out.ToString)
  '' ZZ.IGrid_put(MAIN.igView, RESULT, vbLF)
End Sub

Sub richtb_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkClickedEventArgs) ''... Handles RichTextBox.LinkClicked
  Msgbox(e.LinkText)
End Sub

Sub btn_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  'cancelFlag=True
  'MAIN.btn_Cancel.Enabled=False
End Sub

Sub btn_Ende_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  MAIN.close()
End Sub




Sub OnViewForm1_Formclosing ( ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
  glob.saveFormPos(MAIN)
  glob.saveTuttiFrutti(MAIN)
  glob.saveParaFile()
End Sub

Sub OnViewForm1_Load ( ByVal sender As System.Object, ByVal e As System.EventArgs)
  'msgbox("hallo")
  MAIN.CheckForIllegalCrossThreadCalls= False
  glob = New cls_globPara()
  glob.readFormPos(MAIN)
  glob.readTuttiFrutti(MAIN)
 ' MAIN.igView.Cols.Count=14
End Sub


Sub Autostart()
  ntfy = New NotifyIcon
  ntfy.Visible=True
  
  
  MAIN = New ViewForm1(Me)
  addhandler main.formClosing,addressof OnViewForm1_Formclosing
  ntfy.icon=main.icon
  'MAIN.Show()
  Application.Run(MAIN)
End Sub





