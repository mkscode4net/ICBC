using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
namespace ICBC.Utility
{
    public class Serializer
    {
        static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    throw new Exception("XML Validation Failed");
                
            }
        }

        public static bool IsValidXmlFile(string xmlFilePath, string xsdfilePath)
        {
            var stringXml = File.ReadAllText(xmlFilePath);
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add("", xsdfilePath);
            settings.ValidationType = ValidationType.Schema;
            XmlReader reader = XmlReader.Create(new StringReader(stringXml), settings);
            XmlDocument document = new XmlDocument();
            document.Load(reader);
            document.Validate(ValidationEventHandler);
            return true;
        }


        public T Deserialize<T>(string input) where T : class
        {
            XmlRootAttribute xRoot = new XmlRootAttribute();
            string typeName = typeof(T).Name;
            xRoot.ElementName = typeName;
            xRoot.Namespace = "https://" + typeName;
            xRoot.IsNullable = true;
            XmlSerializer ser = new XmlSerializer(typeof(T), new XmlRootAttribute(typeName));
            using (StringReader sr = new StringReader(input))
            {
                return (T)ser.Deserialize(sr);
            }
        }

        public string Serialize<T>(T ObjectToSerialize)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(ObjectToSerialize.GetType());

            using (StringWriter textWriter = new StringWriter())
            {
                xmlSerializer.Serialize(textWriter, ObjectToSerialize);
                return textWriter.ToString();
            }
        }
    }
}
