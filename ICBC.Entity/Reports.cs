using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ICBC.Entity
{
    [Serializable]
    public class Reports
    {
        [XmlElement(ElementName = "Report")]
        public Report Report { get; set; }
    }
}
