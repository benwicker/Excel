using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace ExcelTut.Entities
{
   public class Event
   {
      [JsonProperty(PropertyName = "type")]
      public string Type { get; set; } = "addField";
      [JsonProperty(PropertyName = "params")]
      public Params Params { get; set; }

      public static Event GetDefaultEvent()
      {
         var e = new Event()
         {
            Type = "addField",
            Params = new Params()
            {
               Fields = new List<Field>()
               {
                  new Field()
                  {
                     Id = "MineSubsidenceLimit",
                     Label = "Mine Subsidence Limit",
                     Validation = "required",
                     Model = new FieldModel()
                     {
                        Name = "mineSubsidenceLimit",
                        Value = "Default Value"
                     }
                  }
               }
            }
         };

         return e;
      }
   }

   public class Params
   {
      [JsonProperty(PropertyName = "field")]
      public List<Field> Fields { get; set; }
   }

   public class Field
   {
      [JsonProperty(PropertyName = "id")]
      public string Id { get; set; }
      [JsonProperty(PropertyName = "label")]
      public string Label { get; set; }
      [JsonProperty(PropertyName = "model")]
      public FieldModel Model { get; set; }
      [JsonProperty(PropertyName = "validation")]
      public string Validation { get; set; }
   }

   public class FieldModel
   {
      [JsonProperty(PropertyName = "name")]
      public string Name { get; set; }
      [JsonProperty(PropertyName = "value")]
      public string Value { get; set; }
   }
}