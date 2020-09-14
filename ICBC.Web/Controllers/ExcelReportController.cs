using System;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using ICBC.BL.XmlReport;
using ICBC.Entity;
using ICBC.Utility;
using ICBC.Utility.Constants;
using Microsoft.AspNetCore.Mvc;

namespace ICBC.Web.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Produces("application/json")]
    public class ExcelReportController : ControllerBase
    {
        private readonly ILoggerManager _logger;
        private XmlReport objXmlReport = default(XmlReport);
        public ExcelReportController(ILoggerManager logger)
        {
            _logger = logger;
            objXmlReport = new XmlReport(_logger);
        }


        [Route("[action]")]
        [HttpGet]
        public IActionResult GetReportFiles()
        {
            System.Text.StringBuilder reportFiles = new System.Text.StringBuilder();
            _logger.LogInfo($"GetReportFiles");
            try
            {
                string reportFolder = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName);
                if (!System.IO.Directory.Exists(reportFolder))
                {
                    _logger.LogInfo($"GetReportFiles: Directory not found creating directory Path: {reportFolder}");
                    System.IO.Directory.CreateDirectory(reportFolder);
                }
                string[] filePaths = System.IO.Directory.GetFiles(reportFolder);
                foreach (string fileName in filePaths.Reverse())
                {
                    reportFiles.Append(ProcessFile(fileName, "xlsx"));
                }

                return Ok(new { ReportFiles = reportFiles.ToString().TrimEnd(',') });
            }
            catch (Exception ex)
            {
                _logger.LogInfo($"GetReportFiles: {ex.ToString()}");
            }

            return Ok(new { ReportFiles = "NO_REPORT_FOUND" });
        }

        private string ProcessFile(string path, string fileType)
        {
            if (path.EndsWith("." + fileType, StringComparison.CurrentCultureIgnoreCase))
            {
                return path.Substring(path.LastIndexOf('\\') + 1) + ",";
            }
            return string.Empty;
        }


        [Route("[action]/{excelFileName}/{sheetName}")]
        [HttpGet]
        public IActionResult GetReport(string excelFileName, string sheetName)
        {
            _logger.LogInfo($"GetReport: excelFileName: {excelFileName}, sheetName={sheetName}");
            string filePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\" + excelFileName);
            var data = objXmlReport.ReadExcelFile(filePath, sheetName);
            return Ok(data);
        }


        [Route("[action]/{fileName}")]
        [HttpGet]
        public IActionResult GetLogFile(string fileName)
        {
            System.Text.StringBuilder reportFiles = new System.Text.StringBuilder();
            _logger.LogInfo($"GetLogs");
            try
            {
                string logFolder = System.IO.Path.GetFullPath(ExcelConstants.LogFolderName + "\\" + fileName);
                if (System.IO.File.Exists(logFolder))
                {
                    
                    _logger.LogInfo($"GetLogs: Directory not found creating directory Path: {logFolder}");
                    return Ok(new { LogFiles = System.IO.File.ReadAllText(logFolder) });
                }
                return Ok(new { LogFiles = "No file found." });
            }
            catch (Exception ex)
            {
                _logger.LogInfo($"GetLogs: {ex.ToString()}");
            }

            return Ok(new { LogFiles = "NO_Log_FOUND" });
        }

        [Route("[action]")]
        [HttpGet]
        public IActionResult GetLogs()
        {
            System.Text.StringBuilder reportFiles = new System.Text.StringBuilder();
            _logger.LogInfo($"GetLogs");
            try
            {
                string logFolder = System.IO.Path.GetFullPath(ExcelConstants.LogFolderName);
                if (!System.IO.Directory.Exists(logFolder))
                {
                    _logger.LogInfo($"GetLogs: Directory not found creating directory Path: {logFolder}");
                    System.IO.Directory.CreateDirectory(logFolder);
                }
                string[] filePaths = System.IO.Directory.GetFiles(logFolder);
                foreach (string fileName in filePaths.Reverse())
                {
                    reportFiles.Append(ProcessFile(fileName, "log"));
                }

                return Ok(new { LogFiles = reportFiles.ToString().TrimEnd(',') });
            }
            catch (Exception ex)
            {
                _logger.LogInfo($"GetLogs: {ex.ToString()}");
            }

            return Ok(new { LogFiles = "NO_Log_FOUND" });
        }




        [Route("[action]/{templateFileName}/{sheetName}")]
        [HttpPost, DisableRequestSizeLimit]
        public IActionResult XmlFileUpload(string templateFileName, string sheetName)
        {
            try
            {
                var file = Request.Form.Files[0];
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), ExcelConstants.XMLFolderName);
                if (!System.IO.Directory.Exists(pathToSave))
                {
                    _logger.LogInfo($"XmlFileUpload: Creating directory {pathToSave}");
                    System.IO.Directory.CreateDirectory(pathToSave);
                }
                if (file.Length > 0)
                {
                    string newReportfilename = string.Empty;
                    var fileName = DateTime.Now.ToString(ExcelConstants.FileNamePrefixFormat) + ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                    _logger.LogInfo($"XmlFileUpload: Trying to save file {fileName}");
                    if (fileName.EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase))
                    {


                        var fullPath = Path.Combine(pathToSave, fileName);
                        using (var stream = new FileStream(fullPath, FileMode.Create))
                        {
                            file.CopyTo(stream);
                            _logger.LogInfo($"XmlFileUpload: Successfully copied file {fullPath}");
                        }
                        var templateFilePath = System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + templateFileName);
                        if (System.IO.File.Exists(templateFilePath))
                        {
                            var xsdfilePath = System.IO.Path.GetFullPath(ExcelConstants.XSDFolderName + "\\" + ExcelConstants.XSDFileName);
                            if (System.IO.File.Exists(xsdfilePath))
                            {
                                if (Serializer.IsValidXmlFile(fullPath, xsdfilePath))
                                {
                                    _logger.LogInfo($"XmlFileUpload: Trying to create excel report for {fullPath}");
                                    Serializer ser = new Serializer();
                                    string xmlInputData = string.Empty;
                                    xmlInputData = System.IO.File.ReadAllText(fullPath);
                                    var report = ser.Deserialize<Reports>(xmlInputData);
                                    objXmlReport.UpdateSheet(templateFilePath, sheetName, report.Report, out newReportfilename);
                                    _logger.LogInfo($"XmlFileUpload: Successfully created excel report {newReportfilename}");
                                    return Ok(new { FileName = newReportfilename, Message = "Successfully uploaded", Result = "Success" });
                                }
                                else
                                {
                                    _logger.LogError($"XmlFileUpload: No XML validation error");
                                    return Ok(new { FileName = "Upload Error", Message = "XML validation failed", Result = "UploadError" });

                                }
                            }
                            else
                            {
                                //error
                                _logger.LogError($"XmlFileUpload: No XSD template file found");
                                return Ok(new { FileName = "Upload Error", Message = "XSD file does not exist", Result = "UploadError" });

                            }
                        }
                        else
                        {
                            _logger.LogError($"XmlFileUpload: No excel template file found");
                            return Ok(new { FileName = "Upload Error", Message = "Excel template file does not exist", Result = "UploadError" });

                        }
                    }
                    else
                    {
                        _logger.LogError($"XmlFileUpload: No XML file found to upload");
                        return Ok(new { FileName = "No File Found Error", Message = "No XML File Found", Result = "NoFileFound" });
                    }
                }
                else
                {
                    //error
                    _logger.LogError($"XmlFileUpload: No  file found to upload");
                    return Ok(new { FileName = "No File Found Error", Message = "No File Found", Result = "NoFileFound" });
                }
            }
            catch (Exception ex)
            {
                //error
                return Ok(new { FileName = "Upload Error", Message = ex.Message, Result = "UploadError" });
            }
        }
    }
}
