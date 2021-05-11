using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp1
{
    class Excel //Apparently it is a good idea to create an Excel class file
    {
        string path = ""; //initialises path variable, other variables initialised below
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(string path, int Sheet) 
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        
        public string ReadCell(int i, int j) //This is a method
        {
            i++; // This is good because excel starts at 1, not 0
            j++;
            _Excel.Range cells = ws.Cells[i, j]; // Makes cells it's own object. This is to stop problems with the 'double dot' rule on the next line!
            if (cells.Value2 != null)
                return cells.Value2; // Returns a value if cells are not null
            else
                return "No values"; 
        }

        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            _Excel.Range range = (_Excel.Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty];
            for (int p = 1; p <= endi - starti; p++)
            {
                for(int q = 1; q <= endy - starty; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;

        }
        public void WriteToCell(int i, int j, string s)
        {
            i++; // This is good because excel starts at 1, not 0
            j++;
            _Excel.Range cells = ws.Cells[i, j];
            cells.Value2 = s;
        }

        public void Save() //So far this doesn't work
        {
            wb.Save();
        }

        public void SaveAs(string path) //It works! But it doesn't like saving in desktop folders with an absolute path for some reason. 
        {
            wb.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void Close()
        {
            wb.Close();
        }
    }
}
