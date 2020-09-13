using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ICBC.Entity
{
    [Serializable]
    public class ReportVal
    {
        [XmlElement(ElementName = "ReportRow")]
        public int ReportRow { get; set; }
        [XmlElement(ElementName = "ReportCol")]
        public int ReportCol { get; set; }
        [XmlElement(ElementName = "Val")]
        public string Val { get; set; }
    }
}
