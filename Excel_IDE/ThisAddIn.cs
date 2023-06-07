using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Excel_IDE;
using System.Drawing;

namespace Excel_IDE
{
    public partial class ThisAddIn
    {

        // Directory
        public bool hasDir = false;
        public string currentDir = null;
        public string[] filesInDir = null;

        // Python Intepreter
        public bool hasPythonIntepreter = false;
        static public string pythonPath = null;
        public string pipPath = null;
        public string scriptPath = null;
        public string pythonVersion = null;

        // Syntax hightligting
        public Dictionary<Excel.Worksheet, string[]> variables = new Dictionary<Excel.Worksheet, string[]>();
        public Dictionary<Excel.Worksheet, string[]> functions = new Dictionary<Excel.Worksheet, string[]>();
        public Dictionary<Excel.Worksheet, string[]> classes = new Dictionary<Excel.Worksheet, string[]>();
        public Dictionary<Excel.Worksheet, string[]> imports = new Dictionary<Excel.Worksheet, string[]>();
        public List<string> keywords = new List<string>();


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Checks if the current directory has a venv
            CheckVenv();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void SaveAllSheets()
        {
            if (!hasDir)
                OpenFileDialog();

            // Saves each sheet as a file in the dir
            foreach (Excel.Worksheet displayWorksheet in this.Application.Worksheets)
            {
                string sheetName = displayWorksheet.Name;
                // Remove invalid charachters from the sheetname
                sheetName = CheckTitleSymbols(sheetName);

                // Keep the output sheet from being saved
                if (sheetName == "Output")
                    continue;

                string fileContent = GetTextFromWorkSheet(displayWorksheet);
                WriteToFile(sheetName + ".py", fileContent);
            }
        }

        public void RunSheets()
        {
            Excel.Worksheet currentWorksheet = this.Application.ActiveSheet;

            string currentSheetName = currentWorksheet.Name;
            currentSheetName = CheckTitleSymbols(currentSheetName);

            cmd.RunPy(currentDir, currentSheetName + ".py");
        }

        string GetTextFromWorkSheet(Excel.Worksheet worksheet)
        {
            string fullText = "";

            // src for code segment: https://www.youtube.com/watch?v=H0wlndQUJiU
            // Get all text from a sheet, and return as a string
            Range usedRange = worksheet.UsedRange;

            if (usedRange.Rows.Count > 0)
            {
                for (int irow = 1; irow <= usedRange.Rows.Count; irow++)
                {
                    for (int jcol = 1; jcol <= usedRange.Columns.Count; jcol++)
                    {
                        Range cell = usedRange.Cells[irow, jcol] as Range;
                        string tabs = String.Concat(Enumerable.Repeat("\t", jcol - 1));
                        if (cell.Value2 != null)
                            fullText += tabs + cell.Value2.ToString();
                    }

                    fullText += "\n";
                }
            }

            // Turn smart quotes into normal quotes and other stuff
            fullText = CheckBodySymbols(fullText);
            return fullText;
        }
        

        public string CheckBodySymbols(string content)
        {
            // https://stackoverflow.com/a/58867897

            string text = content;
            // smart single quotes and apostrophe,  single low-9 quotation mark, single high-reversed-9 quotation mark, prime
            text = Regex.Replace(text, "[\u2018\u2019\u201A\u201B\u2032]", "'");
            // smart double quotes, double prime
            text = Regex.Replace(text, "[\u201C\u201D\u201E\u2033]", "\"");
            // ellipsis
            text = Regex.Replace(text, "\u2026", "...");
            // em dashes
            text = Regex.Replace(text, "[\u2013\u2014]", "-");
            // horizontal bar
            text = Regex.Replace(text, "\u2015", "-");
            // double low line
            text = Regex.Replace(text, "\u2017", "-");
            // circumflex
            text = Regex.Replace(text, "\u02C6", "^");
            // open angle bracket
            text = Regex.Replace(text, "\u2039", "<");
            // close angle bracket
            text = Regex.Replace(text, "\u203A", ">");
            // weird tilde and nonblocking space
            text = Regex.Replace(text, "[\u02DC\u00A0]", " ");
            // half
            text = Regex.Replace(text, "[\u00BD]", "1/2");
            // quarter
            text = Regex.Replace(text, "[\u00BC]", "1/4");
            // dot
            text = Regex.Replace(text, "[\u2022]", "*");
            // degrees 
            text = Regex.Replace(text, "[\u00B0]", " degrees");

            return text;
        }

        public string CheckTitleSymbols(string title)
        {
            string text = title;

            //text = Regex.Replace(text, "[\\.\\/].*", "");
            //text = Path.GetFileName(text);
            text = Path.GetFileNameWithoutExtension(text);

            return text;
        }

        public void WriteToFile(string fileName, string content)
        {
            if (!Directory.Exists(currentDir))
            {
                Output.Error.NullDir();
                return;
            }

            // Append text to an existing file named "WriteLines.txt".
            using (StreamWriter outputFile = new StreamWriter(Path.Combine(currentDir, fileName), false))
            {
                outputFile.Write(content);
            }
        }

        

        public string OpenFileDialog()
        {
            string dirPath = null;

            // https://stackoverflow.com/a/11624322

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filesInDir = Directory.GetFiles(fbd.SelectedPath);
                    currentDir = fbd.SelectedPath;
                    hasDir = true;
                }

            }

            return dirPath;
        }


        public void OpenSheets(string dirPath, bool deleteRest)
        {
            //if (!Directory.Exists(dirPath))
            //{
            //    Output.Error.NullDir();
            //    return;
            //}

            Output.WriteArray(filesInDir);
            
            // Create dummy sheet
            Excel.Worksheet loadingSheet;
            loadingSheet = (Excel.Worksheet)this.Application.Worksheets.Add();
            loadingSheet.Name = "LoAdInG";

            if (deleteRest && filesInDir.Length > 0)
            {
                foreach (Excel.Worksheet worksheet in this.Application.Worksheets)
                {
                    if (worksheet.Name != "LoAdInG")
                        worksheet.Delete();
                }
            }

            foreach (string file in filesInDir)
            {
                Excel.Worksheet newSheet;
                newSheet = (Excel.Worksheet)this.Application.Worksheets.Add();
                string fileName = CheckTitleSymbols(file);
                newSheet.Name = fileName;

                string text = File.ReadAllText(file);
                string[] textArray = text.Replace("\r", "").Split('\n');

                for (int i = 0; i < textArray.Length; i++)
                { 
                    int tabs = textArray[i].Count(f => f == '\t');
                    textArray[i] = textArray[i].Replace("\t", "");
                    newSheet.Cells[i+1, tabs+1].value = textArray[i];
                }
            }

            loadingSheet.Delete();
            Output.outputSheet = null;

            CheckVenv();
        }


        public void CreateVenv()
        {
            // CREATE VENV
            string venvCmd = "python -m venv " + currentDir + "\\venv";
            cmd.RunCmd(venvCmd, true);
            //Process p1 = new Process();
            //ProcessStartInfo info1 = new ProcessStartInfo();
            //info1.FileName = "cmd.exe";
            //info1.RedirectStandardInput = true;
            //info1.RedirectStandardOutput = true;
            //info1.CreateNoWindow = true;
            //info1.UseShellExecute = false;

            //p1.StartInfo = info1;
            //p1.Start();

            //p1.StandardInput.WriteLine(venvCmd);
            //p1.StandardInput.Flush();
            //p1.StandardInput.Close();
            //p1.WaitForExit();
            //if (p1.ExitCode == 0)
            //    Output.WriteLine("Succesfully created python intepreter");

            //Output.WriteString(p1.StandardOutput.ReadToEnd());


            pythonPath = currentDir + "\\venv\\Scripts\\python";
            pipPath = currentDir + "\\venv\\Scripts\\pip";
            scriptPath = currentDir + "\\venv\\Scripts";

            // ACTIVATE THE VENV
            string activateCmd = scriptPath+"\\activate --versíon";
            cmd.RunCmd(activateCmd, true);

            //Process p2 = new Process();
            //ProcessStartInfo info2 = new ProcessStartInfo();
            //info2.FileName = "cmd.exe";
            //info2.RedirectStandardInput = true;
            //info2.RedirectStandardOutput = true;
            //info2.CreateNoWindow = true;
            //info2.UseShellExecute = false;

            //p2.StartInfo = info2;
            //p2.Start();

            //p2.StandardInput.WriteLine(activateCmd);
            //p2.StandardInput.Flush();
            //p2.StandardInput.Close();
            //p2.WaitForExit();

            //if (p2.ExitCode == 0)
            //    Output.WriteLine("Activated python interpreter");


            // GET THE VERSION
            string verCmd = pythonPath + " --versíon";
            cmd.RunCmd(verCmd, true);

            //Process p3 = new Process();
            //ProcessStartInfo info3 = new ProcessStartInfo();
            //info3.FileName = "cmd.exe";
            //info3.RedirectStandardInput = true;
            //info3.RedirectStandardOutput = true;
            //info3.CreateNoWindow = true;
            //info3.UseShellExecute = false;

            //p3.StartInfo = info3;
            //p3.Start();

            //p3.StandardInput.WriteLine(verCmd);
            //p3.StandardInput.Flush();
            //p3.StandardInput.Close();
            //p3.WaitForExit();

            CheckVenv();
        }

        public bool CheckVenv()
        {
            bool hasVenv = false;
            if (Directory.Exists(currentDir + "\\venv"))
            {
                hasPythonIntepreter = true;
                pythonPath = currentDir + "\\venv\\Scripts\\python";
                pipPath = currentDir + "\\venv\\Scripts\\pip";
                scriptPath = currentDir + "\\venv\\Scripts";

                Globals.Ribbons.Ribbon1.PythonIntBtn.Visible = false;
                Globals.Ribbons.Ribbon1.PythonIntBtn.Enabled = false;
                Globals.Ribbons.Ribbon1.runButton.Enabled = true;
                Globals.Ribbons.Ribbon1.importBtn.Enabled = true;
                hasVenv = true;
            }
            else
            {
                hasPythonIntepreter = false;
                pythonPath = null;

                Globals.Ribbons.Ribbon1.PythonIntBtn.Visible = true;
                Globals.Ribbons.Ribbon1.PythonIntBtn.Enabled = true;
                Globals.Ribbons.Ribbon1.runButton.Enabled = false;
                Globals.Ribbons.Ribbon1.importBtn.Enabled = false;
                hasVenv = false;
            }


            return hasVenv;
        }

        


        public void SyntaxHightligt(Excel.Worksheet sheet, int r, int c)
        {
            string cellText = sheet.Cells[r, c].value;



        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }

    public static class cmd
    {
        public static void RunCmd(string cmd, bool writeToOutput)
        {
            Output.WipeOutput();
            Output.outputSheet.Activate();
            Output.WriteLine("Running command: "+cmd);

            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "cmd.exe";
            info.RedirectStandardInput = true;
            info.RedirectStandardOutput = true;
            info.RedirectStandardError = true;
            info.CreateNoWindow = true;
            info.UseShellExecute = false;

            p.OutputDataReceived += (a, b) => Output.WriteLine(b.Data);
            p.ErrorDataReceived += (a, b) => Output.WriteLine(b.Data);

            p.StartInfo = info;
            p.Start();

            

            p.StandardInput.WriteLine(cmd);
            while (!p.HasExited)
            {
                if (writeToOutput)
                {
                    p.BeginErrorReadLine();
                    p.BeginOutputReadLine();
                }
            }
            
            p.StandardInput.Flush();
            p.StandardInput.Close();

            //if (writeToOutput) Output.WriteLine(p.StandardOutput.ReadToEnd());
            //if (writeToOutput) Output.WriteLine(p.StandardError.ReadToEnd());

            p.WaitForExit();

            //if (writeToOutput) Output.WriteConsoleOutput(p);
            //return p.ExitCode;
        }

        public static void RunPy(string dirPath, string fileName)
        {
            Output.WipeOutput();
            Output.outputSheet.Activate();
            Output.WriteLine("Running file \"" + fileName +"\"");


            if (!Directory.Exists(dirPath))
            {
                Output.Error.NullDir();
            }

            // Go to the file's directory
            string cdCmd = "cd " + dirPath;
            string runCmd = ThisAddIn.pythonPath + " " + fileName;
            string clearCmd = "cls";

            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "cmd.exe";
            info.RedirectStandardInput = true;
            info.RedirectStandardOutput = true;
            info.RedirectStandardError = true;
            info.CreateNoWindow = true;
            info.UseShellExecute = false;

            p.OutputDataReceived += (a, b) => Output.WriteLine(b.Data);
            p.ErrorDataReceived += (a, b) => Output.WriteLine(b.Data);

            p.StartInfo = info;
            p.Start();

            p.StandardInput.WriteLine(cdCmd);
            p.StandardInput.WriteLine(clearCmd);
            p.StandardInput.WriteLine(runCmd);

            while (!p.HasExited)
            {
                p.BeginErrorReadLine();
                p.BeginOutputReadLine();
            }

            p.StandardInput.Flush();
            p.StandardInput.Close();
            p.WaitForExit();
        }
    }

    
}

