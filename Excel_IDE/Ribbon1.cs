using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.IO;
using System.Text;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;


namespace Excel_IDE
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void runButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SaveAllSheets(true);
        }

        private void saveBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.hasDir)
            {
                string newDir = Globals.ThisAddIn.OpenFileDialog();
            }
            Globals.ThisAddIn.SaveAllSheets(false);
        }

        private void openBtn_Click(object sender, RibbonControlEventArgs e)
        {
            string newDir = Globals.ThisAddIn.OpenFileDialog();

            Globals.ThisAddIn.OpenSheets(newDir, true);
        }

        private void importBtn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void PythonIntBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.hasDir)
            {
                string newDir = Globals.ThisAddIn.OpenFileDialog();
            }
            if (Directory.Exists(Path.Combine(Globals.ThisAddIn.currentDir, "venv")))
            {
                MessageBox.Show("Python intepreter already exist!", "ERROR");
                return;
            }

            Globals.ThisAddIn.CreateVenv();
        }
    }
}
