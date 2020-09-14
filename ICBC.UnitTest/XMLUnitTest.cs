using System.IO;
using ICBC.BL.XmlReport;
using ICBC.Utility;
using ICBC.Utility.Constants;
using Newtonsoft.Json.Linq;
using Xunit;
using Serializer = ICBC.Utility.Serializer;

namespace ICBC.UnitTest
{
    public class XMLUnitTest
    {
        private const string firstSheetName = "F 20.04";
        private const string secondSheetName = "Final";

        [Fact]
        public void Create_Excel_Report_From_XML()
        {
            string excelTemplatefileName = "TestReport.xlsx";
            var excelfilePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + excelTemplatefileName);
            var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + ExcelConstants.XSDFileName);
            var objXmlReport = new XmlReport(new LoggerManager());
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), ExcelConstants.XMLFolderName);
            var xmlFilePath = System.IO.Path.GetFullPath("Data\\SampleXml\\Test.xml");
            var newReportfilename = CreateExcelReport();
            Assert.False(string.IsNullOrWhiteSpace(newReportfilename));
        }

        [Fact]
        public void Read_Excel_Report_From_Excel()
        {
            string excelTemplatefileName = "TestReport.xlsx";
            var excelfilePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + excelTemplatefileName);
            var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + ExcelConstants.XSDFileName);
            var objXmlReport = new XmlReport(new LoggerManager());

            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), ExcelConstants.XMLFolderName);
            var xmlFilePath = System.IO.Path.GetFullPath("Data\\SampleXml\\Test.xml");
            var newReportfilename = CreateExcelReport();
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\" + newReportfilename);
            var reportJsonString = objXmlReport.ReadExcelFile(filePath, firstSheetName);
            Assert.False(string.IsNullOrWhiteSpace(reportJsonString));
        }

        [Fact]
        public void Read_Excel_Report_From_Excel_Check_JSON_DATA()
        {
            string excelTemplatefileName = "TestReport.xlsx";
            var excelfilePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + excelTemplatefileName);
            var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + ExcelConstants.XSDFileName);
            var objXmlReport = new XmlReport(new LoggerManager());

            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), ExcelConstants.XMLFolderName);
            var xmlFilePath = System.IO.Path.GetFullPath("Data\\SampleXml\\Test.xml");
            var newReportfilename = CreateExcelReport();
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\" + newReportfilename);
            var reportJsonString = objXmlReport.ReadExcelFile(filePath, firstSheetName);
            Assert.False(string.IsNullOrWhiteSpace(reportJsonString));
            dynamic jsonObject = JObject.Parse(reportJsonString);
            Assert.True(jsonObject.MessageType == "Success");


        }

        [Fact]
        public void Read_Excel_Report_From_Excel_Check_JSON_DATA_Invalid_File_Type()
        {
            var objXmlReport = new XmlReport(new LoggerManager());

            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\abc.xslx" );
            var reportJsonString = objXmlReport.ReadExcelFile(filePath, firstSheetName);
            Assert.False(string.IsNullOrWhiteSpace(reportJsonString));
            dynamic jsonObject = JObject.Parse(reportJsonString);
            Assert.True(jsonObject.MessageType == "Error");
            Assert.True(jsonObject.Message == "Invalid file type");
        }

        [Fact]
        public void Read_Excel_Report_From_Excel_Check_JSON_DATA_Invalid_File()
        {
            var objXmlReport = new XmlReport(new LoggerManager());
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\abc.xlsx");
            var reportJsonString = objXmlReport.ReadExcelFile(filePath, firstSheetName);
            Assert.False(string.IsNullOrWhiteSpace(reportJsonString));
            dynamic jsonObject = JObject.Parse(reportJsonString);
            Assert.True(jsonObject.MessageType == "Error");
            Assert.True(jsonObject.Message == "File not found");
        }

        [Fact]
        public void Excel_Report_Check_Report_Folder()
        {
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName);
            Assert.True(System.IO.Directory.Exists(filePath));
        }

        [Fact]
        public void Excel_Report_Check_Template_Folder()
        {
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName);
            Assert.True(System.IO.Directory.Exists(filePath));
        }
        [Fact]
        public void Excel_Report_Check_XSD_Folder()
        {
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName);
            Assert.True(System.IO.Directory.Exists(filePath));
        }
        [Fact]
        public void Excel_Report_Check_XML_Folder()
        {
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.XMLFolderName);
            Assert.True(System.IO.Directory.Exists(filePath));
        }



        private string CreateExcelReport()
        {
            string excelTemplatefileName = "TestReport.xlsx";
            var excelfilePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + excelTemplatefileName);

            var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + ExcelConstants.XSDFileName);


            var objXmlReport = new XmlReport(new LoggerManager());
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), ExcelConstants.XMLFolderName);
            var xmlFilePath = System.IO.Path.GetFullPath("Data\\SampleXml\\Test.xml");

            if (Serializer.IsValidXmlFile(xmlFilePath, xsdfilePath))
            {
                Serializer ser = new Serializer();
                string xmlInputData = string.Empty;
                xmlInputData = System.IO.File.ReadAllText(xmlFilePath);
                ICBC.Entity.Reports report = ser.Deserialize<ICBC.Entity.Reports>(xmlInputData);
                string newReportfilename = string.Empty;
                Assert.True(objXmlReport.UpdateSheet(excelfilePath, "F 20.04", report.Report, out newReportfilename));
                Assert.False(string.IsNullOrWhiteSpace(newReportfilename));
                return newReportfilename;
            }
            return string.Empty;
        }



    }
}
