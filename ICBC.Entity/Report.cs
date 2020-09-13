using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ICBC.Entity
{
    [Serializable, XmlRoot(ElementName = "Report")]
    public class Report
    {
        [XmlElement(ElementName = "Name")]
        public string Name { get; set; }
        [XmlElement(ElementName = "ReportVal")]
        public List<ReportVal> ReportVal { get; set; }
    }
}
