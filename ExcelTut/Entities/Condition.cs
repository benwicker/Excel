using Newtonsoft.Json;

namespace ExcelTut.Entities
{
   public class Condition
   {
      public Condition(){}

      public Condition(string fact, string value)
      {
         Fact = fact;
         Value = value;
      }

      [JsonProperty(PropertyName = "fact")]

      public string Fact { get; set; } = "state";
      [JsonProperty(PropertyName = "operator")]
      public string Operator { get; set; } = "equal";
      [JsonProperty(PropertyName = "value")]
      public string Value { get; set; }
   }
}