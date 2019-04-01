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

         var data = book.ReadCell(0, 0);
         Console.WriteLine(data);
         book.WriteCell(0, 0, "newHeader");
         Console.WriteLine(book.ReadCell(0, 0));

         // save and close
         book.Save();
         book.Close();
      }
   }
}
