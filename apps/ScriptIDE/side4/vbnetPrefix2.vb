
:BEGIN
Implements IDisposable

'Public Const ClsID = "{$COM-Class-ID|"00000000-0000-0000-0000-000000000000"}"
Public Const ProgID = "__CLASSNAME__" '"{$COM-Prog-ID|"{ScriptClass}.application"}"

Public Shared zLN As Integer       'Zeilennummer
Public Shared zC As Integer=0, zC2 As Integer=0   'Ausf�hrungscounter
Public Shared ziI As Integer       'Interrupt Schleifenz�hler
Public Shared ziE As String        'Interrupt Execute
Public Shared ziQ As Boolean       'Interrupt Quit
Public Shared zNN As String        'Funktionsname
Public Shared zBB() As Boolean     'Breakpoints
Public Shared NicName As String
Public Shared myZZ As IScriptHelper
Public Shared parentRef As WeakReference

Public ReadOnly Property Name() As String
  Get
    Return "__CLASSNAME__"
  End Get
End Property

Private Shared zHlp As IScriptHelper
Shared Public ReadOnly Property ZZ() As IScriptHelper
  Get
    Return zHlp
  End Get
End Property

Sub zz_setHlpObject(obj As IScriptHelper)
  zHlp = obj
  myZZ = obj
End Sub

Sub zz_BBreset(lineCount As Integer)
  Redim zBB(lineCount)
End Sub

Sub zz_BBsetLine(lineIdx As Integer, stat As Boolean)
  zBB(lineIdx) = stat
End Sub

Sub zz_BBtrace
  Dim i As Integer
  For i=0 To UBound(zBB)
    if ZBB(i) Then ZZ.trace(i,"","Break set in __CLASSNAME__ line",i)
  Next
  ZZ.trace(0,"","Line count",ubound(zBB))
End Sub

Sub New(hlp As IScriptHelper, myParentRef As WeakReference)
  On Error Resume Next
  NicName = Name
  myZZ = hlp
  zHlp = hlp
  parentRef = myParentRef
  DirectCast(Me, Object).onInitialize()
  hlp.trace(0,"Class_Initialize","onIniDone __CLASSNAME__", NicName)
End Sub

Protected Overrides Sub Finalize()
  On Error Resume Next
  ZZ.trace(0,"Class_Finalize","__CLASSNAME__", NicName)
  Dispose(False)
End Sub

Private disposedValue As Boolean = False    ' To detect redundant calls

' IDisposable
Protected Overridable Sub Dispose(ByVal disposing As Boolean)
  If Not Me.disposedValue Then
    Me.disposedValue = True
    If disposing Then
      ' TODO: free other state (managed objects).
      ZZ.releaseMyRef()
    End If

    ZZ.trace(0,"Class_Dispose","__CLASSNAME__", NicName)
    DirectCast(Me, Object).onTerminate()
  End If
End Sub

#Region " IDisposable Support "
' This code added by Visual Basic to correctly implement the disposable pattern.
Public Sub Dispose() Implements IDisposable.Dispose
  ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  Dispose(True)
  GC.SuppressFinalize(Me)
End Sub
#End Region

:END

