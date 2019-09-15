using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace EngSTD
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = excel.Worksheets[Sheet];
        }
        public string ReadCell(int range, int column)
        {
            string val = Convert.ToString(ws.Cells[range, column].Value != null ? ws.Cells[range, column].Value : "");
            return val;
        }
        public void Close()
        {
            wb.Close(true, this.path, null);
            excel.Quit();
            this.CloseProcess();
        }
        public void WriteCell(int range, int column, string value)
        {
            ws.Cells[range, column].Value = value;
        }
        public void DeletCell(int range, int column)
        {
            ws.Cells[range, column].Value = "";
        }
        public void Save()
        {
            wb.Save();

        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }
        public void SelectWorkSheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }
        public void DeletWorkSheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delet();
        }
        public void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
    }
}
