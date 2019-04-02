using System.Collections.Generic;
using Newtonsoft.Json;

namespace ExcelTut.Entities
{
   public class Rule
   {
      public Rule(string name, string domain, string state, Event e)
      {
         Name = name;
         Domain = domain;
         Event = e;
         Conditions = All.GetDefaultAll(state);
      }

      [JsonProperty(PropertyName = "name")]
      public string Name { get; set; }
      [JsonProperty(PropertyName = "domain")]

      public string Domain { get; set; } = "MineSubsidence";
      [JsonProperty(PropertyName = "conditions")]

      public All Conditions { get; set; }
      [JsonProperty(PropertyName = "event")]

      public Event Event { get; set; }
   }
}
