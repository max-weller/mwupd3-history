      
      Ref System
      Ref mscorlib
      Ref System.Windows.Forms
      Ref Microsoft.visualbasic
      
      Dim vbCrLf
      FGet vbCrLf = Microsoft.VisualBasic.ControlChars::CrLf
      
      #Dim frm As New Form
      #Dim txt As New TextBox
      #frm.Controls.Add(txt)
      #txt.Dock=DockStyle.Fill
      #txt.MultiLine=True
      #txt.Text=Io.File.ReadAllText("")
      #frm.ShowDialog()
      
      #Variablendeklaration
      Dim FactVar
      Dim FibArray
      
      GoSub FactorialTest
      GoSub FibonacciTest
      GoSub Test1
      GoSub AskUser
      GoSub FormTest
      Return
      Exit
      
      #Funktionen
      
AskUser:
      Dim Result
      Call Result=System.Windows.Forms.messageBox::Show["Hallo"+vbCrLf+"...ich bin eine Hinweisbox","title",{System.Windows.Forms.MessageBoxButtons}1]
      If Result = 2
        Exit
      EndIf
      Return
      
FormTest:
      Var  frm = New System.Windows.Forms.Form[]
      Var  txt = New System.Windows.Forms.TextBox[]
      Dim  ctrls,fileCont
      Get  ctrls=frm.Controls[]
      Call ctrls.Add[txt]
      Set  txt.Dock[{System.Windows.Forms.DockStyle} 5]
      Set  txt.MultiLine[True]
      Call fileCont=System.IO.File::ReadAllText["d:\_test\mybatch1.mbt"]
      Set  txt.Text[fileCont]
      Call frm.ShowDialog[]
      Return
      
      
FactorialTest:
      FactVar = 1
      GoSub Factorial[10]
      Trace FactVar
      Return
      
Factorial:
      Trace "Zähler = "+STR(§[0])+"    Wert = $(FactVar)"
      FactVar = FactVar * §[0]
      If §[0] > 1
        GoSub Factorial[§[0]-1]
      EndIf
      Return
      
      
FibonacciTest:
      GoSub Fibonacci
      
      Dim I
      Dim Out
      Out = "Werte: "
      I = 0
Loop1:Out = Out + " " + STR(FibArray[I])
      I = I+1
      If I < LEN(FibArray)
        Goto Loop1
      EndIf
      TRACE Out
      
      Return
      
Fibonacci:
      If § = 
        FibArray = [1,1]
        GoSub Fibonacci 2
        Return
      EndIf
      FibArray[] = FibArray[§-2] + FibArray[§-1]
      If § > 20
        Return
      EndIf
      GoSub Fibonacci §+1
      Return
      
      #Hauptteil
MAIN:
      
      
      
      #Exit
      
Test1:
      Dim Var1,Var2,Var3,Var4
      
      Var1 = "Hallo"
      Var2 = "$(Var1) Welt!"
      Var3 = 21.83
      Var4 = [1,1,2,3,5,7,9,Var3 + 1]
      Var Var5 = [[Var1, Var2], [Var3, Var4]]
      Var5[] = "XYZ"
      
      TRACE Var1
      TRACE Var2
      TRACE Var3
      TRACE Var4
      TRACE Var5
      
      #CheckVar = Var3
      #GoSub CheckVar
      Return
      