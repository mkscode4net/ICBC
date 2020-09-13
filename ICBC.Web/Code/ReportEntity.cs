///*
//namespace ICBC.Web
//{
//  using System;
//  using System.Collections.Generic;

//  using System.Globalization;
//  using Newtonsoft.Json;
//  using Newtonsoft.Json.Converters;


//  [Serializable]
//  public  class Reports
//  {
//      [JsonProperty("Report")]
//      public Report Report { get; set; }
//  }

//  [Serializable]
//  public  class Report
//  {
//      [JsonProperty("Name")]
//      public string Name { get; set; }

//      [JsonProperty("ReportVal")]
//      public List<ReportVal> ReportVal { get; set; }
//  }
//  [Serializable]
//  public  class ReportVal
//  {
//      [JsonProperty("ReportRow")]
//      [JsonConverter(typeof(ParseStringConverter))]
//      public int ReportRow { get; set; }

//      [JsonProperty("ReportCol")]
//      [JsonConverter(typeof(ParseStringConverter))]
//      public int ReportCol { get; set; }

//      [JsonProperty("Val")]
//      [JsonConverter(typeof(ParseStringConverter))]
//      public string Val { get; set; }
//  }




//  internal class ParseStringConverter : JsonConverter
//  {
//      public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

//      public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
//      {
//          if (reader.TokenType == JsonToken.Null) return null;
//          var value = serializer.Deserialize<string>(reader);
//          long l;
//          if (Int64.TryParse(value, out l))
//          {
//              return l;
//          }
//          throw new Exception("Cannot unmarshal type long");
//      }

//      public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
//      {
//          if (untypedValue == null)
//          {
//              serializer.Serialize(writer, null);
//              return;
//          }
//          var value = (long)untypedValue;
//          serializer.Serialize(writer, value.ToString());
//          return;
//      }

//      public static readonly ParseStringConverter Singleton = new ParseStringConverter();
//  }
//}

//  */


//using System;
//    using System.Collections.Generic;
//    using System.Linq;
//    using System.Threading.Tasks;
//    using System.Xml.Serialization;
//    namespace ICBC.Web
//    {
//        [Serializable]
//        public class ReportVal
//        {
//            [XmlElement(ElementName = "ReportRow")]
//            public int ReportRow { get; set; }
//            [XmlElement(ElementName = "ReportCol")]
//            public int ReportCol { get; set; }
//            [XmlElement(ElementName = "Val")]
//            public string Val { get; set; }
//        }




//        [Serializable, XmlRoot(ElementName = "Report")]
//        public class Report
//        {
//            [XmlElement(ElementName = "Name")]
//            public string Name { get; set; }
//            [XmlElement(ElementName = "ReportVal")]
//            public List<ReportVal> ReportVal { get; set; }
//        }
//        [Serializable]
//        public class Reports
//        {
//            [XmlElement(ElementName = "Report")]
//            public Report Report { get; set; }
//        }

//    }
