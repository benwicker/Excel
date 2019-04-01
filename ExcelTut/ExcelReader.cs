using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTut
{
   public class ExcelReader
   {
      private string path = "";
      private _Application excel = new _Excel.Application();
      private Workbook wb;
      private Worksheet ws;

      public ExcelReader(string path, int sheet)
      {
         this.path = path;
         wb = excel.Workbooks.Open(path);
         ws = wb.Worksheets[sheet];
      }

      public string ReadCell(int i, int j)
      {
         i++;
         j++;
         if (ws.Cells[i, j].Value2 != null)
         {
            return ws.Cells[i, j].Value2;
         }
         else
         {
            return "";
         }
      }

      public void WriteCell(int i, int j, string data)
      {
         i++;
         j++;
         ws.Cells[i, j].Value2 = data;
      }

      public void Save()
      {
         wb.Save();
      }

      public void SaveAs(string path)
      {
         wb.SaveAs(path);
      }

      public void Close()
      {
         wb.Close();
      }

   }
}
