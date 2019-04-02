using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTut.Entities
{
   public class All
   {
      public List<Condition> Conditions { get; set; }
      public Any Any { get; set; }

      public static All GetDefaultAll(string state)
      {
         var a = new All()
         {
            Conditions = new List<Condition>()
            {
               new Condition()
               {
                  Fact = "state",
                  Operator = "equal",
                  Value = state
               }
            },
            Any = new Any()
         };

         return a;
      }
   }
}
