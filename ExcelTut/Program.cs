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
         var book = new ExcelReader(@"C:\Users\WickerB\Desktop\test.xlsx", 1);

         string[,] data = book.ReadRange(1, 1, 17, 2);

         book.Close();
      }
   }
}
