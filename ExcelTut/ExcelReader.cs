using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ExcelTut.Entities;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
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

      public string[,] ReadRange(int startX, int startY, int endX, int endY)
      {
         Range range = (Range)ws.Range[ws.Cells[startX, startY], ws.Cells[endX, endY]];
         object[,] holder = range.Value2;
         string[,] returnstring = new string[endX - startX, endY - startY];
         for (int i = 1; i <= endX - startX; i++)
         {
            for (int j = 1; j <= endY - startY; j++)
            {
               returnstring[i - 1, j - 1] = holder[i, j].ToString();
            }
         }

         return returnstring;
      }

      public void CreateRulesJson()
      {
         // hardcoding
         var x1 = 1;
         var x2 = 138;
         var y1 = 1;
         var y2 = 2;

         // pull data from excel
         Range range = (Range) ws.Range[ws.Cells[x1, y1], ws.Cells[x2, y2]];
         object[,] holder = range.Value2;
         List<MineSubsidenceEntry> data = new List<MineSubsidenceEntry>();

         for (int i = 2; i <= x2 - x1; i++)
         {
            data.Add(new MineSubsidenceEntry()
            {
               State = holder[i, 1].ToString(),
               County = holder[i, 2].ToString()
            });
         }

         // build rules
         var e = Event.GetDefaultEvent();

         var IL = new Rule("Illinois", "Mine Subsidence", "IL", e);
         var IN = new Rule("Indiana", "Mine Subsidence", "IN", e);
         var KY = new Rule("Kentucky", "Mine Subsidence", "KY", e);
         var WV = new Rule("West Virginia", "Mine Subsidence", "WV", e);

         foreach (var entry in data)
         {
            switch (entry.State)
            {
               case "IL":
                  IL.Conditions.Any.Conditions.Add(new Condition("county", entry.County));
                  break;
               case "IN":
                  IN.Conditions.Any.Conditions.Add(new Condition("county", entry.County));
                  break;
               case "KY":
                  KY.Conditions.Any.Conditions.Add(new Condition("county", entry.County));
                  break;
               case "WV":
                  WV.Conditions.Any.Conditions.Add(new Condition("county", entry.County));
                  break;
            }
         }

         var rules = new List<Rule>()
         {
            IL,
            IN,
            KY,
            WV
         };

         // write to file
         var json = JsonConvert.SerializeObject(rules);
         System.IO.File.WriteAllText(@"C:\Users\WickerB\Desktop\Test\ExcelTut\ExcelTut\Assets\mineSubsidenceRules.json", json);
      }
   }
}
