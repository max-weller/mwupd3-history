
#Reference Microsoft.VisualBasic.dll
#Imports Microsoft.VisualBasic
#Reference System.Management.dll
#Imports System.Management
    
    ScriptedPanel panelRef;
    Form formRef;
    ComboBox cmb;
    //AsSystem.Windows.Forms,System.Windows.Forms.ComboBox
    
    public void  // ERROR: Handles clauses are not supported in C#
cmb_SelectedIndexChanged(object sender, System.EventArgs e)
    {
      cmb.SelectionLength = 0;
    }
    
    public void AutoStart()
    {
        panelRef = (ScriptedPanel)ZZ.IDEHelper.CreateDockWindow("cstest.testPanel", 2);
        formRef = panelRef.Form;
        formRef.Show();
        formRef.Width = 400;
        formRef.Height = 200;
        Random r = new Random();
        formRef.BackColor = Color.FromArgb(255, r.Next(0, 255), r.Next(0, 255), r.Next(0, 255));
        formRef.Text = "Hallo Welt -- Hintergrundfarbe: " + ColorTranslator.ToHtml(formRef.BackColor);
        
        panelRef.resetControls(5,5,1);
        
        /*
        cmb = new ComboBox();
        panelRef.Controls.@Add(cmb);
        cmb.Items.Add("111");
        cmb.Items.Add("222");
        */
        
        ZZ.trace(0,"","test1","", "info");
        
    }

