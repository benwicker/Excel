using System.Collections.Generic;

namespace ExcelTut.Entities
{
   public class Any
   {
      public Any()
      {
         Conditions = new List<Condition>();
      }

      public List<Condition> Conditions { get; set; }
   }
}