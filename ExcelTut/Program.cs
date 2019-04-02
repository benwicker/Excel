using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTut
{
   class Program
   {
      static void Main(string[] args)
      {

         // open new workbook
         var book = new ExcelReader(@"C:\Users\WickerB\Downloads\States and counties that require Mine Subsidence coverage.xlsx", 1);

         book.CreateRulesJson();

         book.Close();
      }
   }
}
