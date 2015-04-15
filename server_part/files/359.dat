
#ifcond runtime != "IDE"

#Imports ScriptIDE.ScriptHost
#Imports ScriptIDE.Core
#Imports System.Runtime.InteropServices
#Imports System

#Const TYPE = "winexe"

Public Module ProgramMain
  Public WithEvents sH As ScriptHost
  'Public Shared iproc As sys_interproc
  Public myZZ As cls_scriptHelper
  
  Public Const CLSCTX_LOCAL_SERVER = 4
  Public Const REGCLS_MULTIPLEUSE = 1

  <DllImport("ole32.dll")> _
  Private Function CoRegisterClassObject(ByRef rclsid As Guid, <MarshalAs(UnmanagedType.[Interface])> ByVal pUnkn As IClassFactory, ByVal dwClsContext As Integer, ByVal flags As Integer, ByRef lpdwRegister As Integer) As Integer
  End Function

  <DllImport("ole32.dll")> _
  Private Function CoRevokeClassObject(ByVal dwRegister As Integer) As Integer
  End Function

  Public Sub Main(args() As String)
    #ifdef COM_CLASS
    Dim clsFct As New mwClsFct
    Dim regID As Integer
    CoRegisterClassObject(New Guid(ScriptClass.__CLASSNAME__.ClsID), clsFct, CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, regID)
    #endif
    Try
      
      'Interproc.Initialize
      'sH.setIdeHelper(cls_IDEHelperMini.GetSingleton())
      cls_IDEHelperMini.GetSingleton().initGlobObject(cls_IDEHelperMini.GetSingleton().GetSettingsFolder()+"scriptPara\{ScriptClass}.para.txt")
      
      ' Initialize ScriptHelper
      myZZ = New cls_scriptHelper()
      myZZ._scriptClassName = "__CLASSNAME__"
      'ref.zz_setHlpObject(myZZ)
      
      Dim ref As Object = New ScriptClass.__CLASSNAME__(myZZ, Nothing)
      myZZ._scriptInst = New System.WeakReference(ref)
      #ifdef DEBUG
      sH = New ScriptHost(RuntimeMode.Debug, "__CLASSNAME__", ref)
      #else
      sH = New ScriptHost(RuntimeMode.Release, "__CLASSNAME__", ref)
      #endif
      myZZ.sH = sH
      
      ' Initialize Breakpoints
      ref.zz_BBreset(1000)
      Dim bbs() As String = myZZ.InterProc.GetData("siaCodeCompiler", "_Debug_QueryBreakpoints", "__CLASSNAME__").Split(","c)
      myZZ.trace(0,"","BBcount:" ,bbs.length)
      For Each bb As String In bbs
        If bb <> "" Then ref.zz_BBsetLine(bb, True)
      Next
      
      #ifdef COM_CLASS
      registerServer()
      #endif
      
      #If TYPE = "library"
      ref.AutoStart()
      #End If
      
      #If TYPE = "winexe2" Or TYPE = "winforms2"
      
      ref.AutoStart()
      #End If
      
      #If TYPE = "winexe"
      ref.AutoStart()
      Dim frm As Object = ScriptIDE.ScriptWindowHelper.ScriptWindowManager.GetFirstScriptWindow()
      If frm IsNot Nothing Then
      '' myZZ.trace(0,"","Form:" + frm.tostring)
        System.Windows.Forms.Application.Run(frm.Form)
      End If
      #End If
    Catch ex As Exception
      Microsoft.VisualBasic.Interaction.MsgBox("Unhandled exception in unsafe code"+Microsoft.VisualBasic.ControlChars.Newline+ex.ToString, 16)
    
    Finally
      
    #ifdef COM_CLASS
      CoRevokeClassObject(regID)
    #endif
    End Try
  End Sub
  
  
#ifdef COM_CLASS
  Sub registerServer()
    Dim regKey, regKey2 As Microsoft.Win32.RegistryKey

    Dim progID = q_scriptClass.__CLASSNAME__.ProgID
    Dim clsID = "{" + q_scriptClass.__CLASSNAME__.ClsID + "}"
    If clsID = "{00000000-0000-0000-0000-000000000000}" Then Exit Sub
    
    regKey = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(progID)
    regKey.SetValue("", progID)

    regKey2 = regKey.CreateSubKey("ClsID")
    regKey2.SetValue("", clsID)
    regKey2.Close()

    regKey.Close()

    regKey = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey("CLSID\" + clsID)
    regKey.SetValue("", progID)

    regKey2 = regKey.CreateSubKey("Implemented Categories")
    regKey2.CreateSubKey("{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}").Close()
    regKey2.Close()

    regKey2 = regKey.CreateSubKey("LocalServer32")
    regKey2.SetValue("", Windows.Forms.Application.ExecutablePath)
    regKey2.Close()

    regKey2 = regKey.CreateSubKey("ProgID")
    regKey2.SetValue("", progID)
    regKey2.Close()

    regKey.CreateSubKey("Programmable").Close()

    'regKey2 = regKey.CreateSubKey("TypeLib")
    'regKey2.SetValue("", typeLibID)
    'regKey2.Close()

    regKey.Close()


  End Sub
#endif
  '' #If DEBUG = True
  '' Private Sub sH_HighlightLineRequested(ByVal className As String, ByVal functionName As String, ByVal lineNumber As Integer, ByVal reason As HighlightLineReason) Handles sH.HighlightLineRequested
  ''   myZZ.InterProc.SendCommand("siaCodeCompiler", "_Debug_HighlightLineRequested", className & "|##|" & functionName & "|##|" & lineNumber & "|##|" & reason & "|##|")
  ''   
  '' End Sub
  '' 
  '' Private Sub sH_BreakModeChanged(ByRef breakState As String) Handles sH.BreakModeChanged
  ''   myZZ.InterProc.SendCommand("siaCodeCompiler", "_Debug_BreakModeChanged", breakState)
  ''   
  '' End Sub
  '' #End If
  
End Module

'''
''' IClassFactory declaration
'''
<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid(mwClsFct.IClassFactory)> _
Friend Interface IClassFactory
  <PreserveSig()> _
  Function CreateInstance(ByVal pUnkOuter As IntPtr, ByRef riid As Guid, ByRef ppvObject As IntPtr) As Integer
  <PreserveSig()> _
  Function LockServer(ByVal fLock As Boolean) As Integer
End Interface

'<Guid(mwClsFct.ClassId)> _
  Public Class mwClsFct

  Implements IClassFactory

  Public Const IClassFactory As String = "00000001-0000-0000-C000-000000000046"
  Public Const IUnknown As String =      "00000000-0000-0000-C000-000000000046"

  'Public Const ClassId As String =       "{$COM-ClassFactory-ID}"


  Public Function CreateInstance(ByVal pUnkOuter As System.IntPtr, ByRef riid As System.Guid, ByRef ppvObject As System.IntPtr) As Integer Implements IClassFactory.CreateInstance
    'ppvObject = Marshal.GetComInterfaceForObject(myZZ._scriptInst.Target, GetType(ScriptClass.__CLASSNAME__.I__CLASSNAME__))
  End Function

  Public Function LockServer(ByVal fLock As Boolean) As Integer Implements IClassFactory.LockServer

  End Function
End Class


#endif


