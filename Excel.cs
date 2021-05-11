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
    }
}
