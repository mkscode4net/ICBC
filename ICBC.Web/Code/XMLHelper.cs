using System;
using System.Collections;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
namespace ICBC.Web
{
    public class Serializer
    {
        static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    Console.WriteLine("Error: {0}", e.Message);
                    break;
                case XmlSeverityType.Warning:
                    Console.WriteLine("Warning {0}", e.Message);
                    break;
            }
        }

        public static bool IsValidXmlFile(string xmlFileName, string xsdFileName)
        {
            try
            {
                var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + xsdFileName);
                var xmlFilePath = System.IO.Path.GetFullPath(ExcelConstants.XMLFolderName + "\\" + xmlFileName);
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
            catch (Exception ex)
            {


            }
            return false;
        }


        public T Deserialize<T>(string input) where T : class
        {
            XmlRootAttribute xRoot = new XmlRootAttribute();
            xRoot.ElementName = "Reports";
            xRoot.Namespace = "http://Reports";
            xRoot.IsNullable = true;

            Type[] types = new Type[1]; types[0] = typeof(T);
            //System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(typeof(T), xRoot);
            XmlSerializer ser = new XmlSerializer(typeof(Reports), new XmlRootAttribute("Reports"));
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
