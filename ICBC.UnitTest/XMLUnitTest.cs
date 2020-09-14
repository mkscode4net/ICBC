using System.IO;
using ICBC.BL.XmlReport;
using ICBC.Utility;
using ICBC.Utility.Constants;
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
