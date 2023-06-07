using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_IDE
{
    public static class Output
    {
        public static Excel.Worksheet outputSheet = null;
        private static int currentLine = 1;

        private static void CreateOutputWorkSheet()
        {
            bool hasSheet = false;
            foreach (Excel.Worksheet sheet in Globals.ThisAddIn.Application.Worksheets)
            {
                if (sheet.Name == "Output")
                {
                    hasSheet = true;
                    outputSheet = sheet;
                    break;
                }
            }

            if (!hasSheet)
            {
                // Create Output sheet
                Excel.Worksheet output;
                output = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                output.Name = "Output";
                outputSheet = output;
            }
        }
        public static void WriteLine(string text)
        {
            CreateOutputWorkSheet();

            outputSheet.Cells[currentLine, 1].Value = text;
            currentLine++;
        }

        public static void WriteConsoleOutput(Process p)
        {
            CreateOutputWorkSheet();
            WipeOutput();

            string outputText = p.StandardOutput.ReadToEnd();
            string errorText = p.StandardError.ReadToEnd();


            string[] outputLines = (outputText + errorText).Replace("\r", "").Split('\n');

            for (int i = 8; i < outputLines.Length - 2; i++)
            {
                WriteLine(outputLines[i]);
            }
        }

        public static void WriteString(string text)
        {
            CreateOutputWorkSheet();
            WipeOutput();

            string[] outputLines = text.Replace("\r", "").Split('\n');

            for (int i = 0; i < outputLines.Length; i++)
            {
                WriteLine(outputLines[i]);
            }
        }

        public static void WriteArray(string[] array)
        {
            CreateOutputWorkSheet();
            WipeOutput();

            for (int i = 0; i < array.Length; i++)
            {
                WriteLine(array[i]);
            }
        }

        public static void WipeOutput()
        {

            if (outputSheet == null || outputSheet.UsedRange == null)
            {
                CreateOutputWorkSheet();
                return;
            }


            Range usedRange = outputSheet.UsedRange;


            if (usedRange.Rows.Count > 0)
            {
                for (int irow = 1; irow <= usedRange.Rows.Count; irow++)
                {
                    outputSheet.Cells[irow, 1].value = "";
                }
            }

            currentLine = 1;
        }


        public static class Error
        {
            private static void ThrowError(string error)
            {
                WriteLine(error);
                outputSheet.Cells[currentLine, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
            public static void NullDir() { ThrowError("NullDir: Directory doesn't exist"); }

        }
    }
}
