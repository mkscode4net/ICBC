using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ICBC.BL.Base;
using ICBC.Entity;
using ICBC.Utility;
using ICBC.Utility.Constants;
using Newtonsoft.Json.Linq;

namespace ICBC.BL.XmlReport
{
    public class XmlReport : BaseBL, IXmlReport
    {
        public XmlReport(ILoggerManager logger) : base(logger)
        {

        }

        public string ReadExcelFile(string fileName, string sheetName)
        {
            JObject o = new JObject();
            JArray arrayRows = new JArray();
            try
            {
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    Sheet thesheet = (Sheet)thesheetcollection.Where(sh => (sh as Sheet).Name.Equals(sheetName)).FirstOrDefault();
                    //using for each loop to get the sheet from the sheetcollection  

                    //statement to get the worksheet object by using the sheet id  
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();
                    //var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Skip(ExcelConstants.ReportHeaderLastRowIndex);
                    foreach (Row thecurrentrow in thesheetdata.Skip(ExcelConstants.ReportHeaderLastRowIndex))
                    {
                        JObject ocell = new JObject();
                        foreach (Cell thecurrentcell in thecurrentrow)
                        {
                            string currentcellvalue = string.Empty;
                            if (thecurrentcell.DataType != null && thecurrentcell.DataType == CellValues.SharedString)
                            {
                                int id;
                                if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                {
                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                    if (item.Text != null)
                                    {
                                        ocell[string.Concat(thecurrentcell.CellReference.Value.Reverse().Skip(thecurrentrow.RowIndex.ToString().Length).Reverse())] = item.Text.Text;
                                    }
                                    else if (item.InnerText != null)
                                    {
                                        ocell[string.Concat(thecurrentcell.CellReference.Value.Reverse().Skip(thecurrentrow.RowIndex.ToString().Length).Reverse())] = item.InnerText;
                                    }
                                    else if (item.InnerXml != null)
                                    {
                                        ocell[string.Concat(thecurrentcell.CellReference.Value.Reverse().Skip(thecurrentrow.RowIndex.ToString().Length).Reverse())] = item.InnerXml;
                                    }
                                }
                            }
                            else
                            {
                                ocell[string.Concat(thecurrentcell.CellReference.Value.Reverse().Skip(thecurrentrow.RowIndex.ToString().Length).Reverse())] = thecurrentcell.InnerText;
                            }
                        }

                        arrayRows.Add(ocell);
                    }

                }
                o["Data"] = arrayRows;
            }
            catch (Exception)
            {

            }
            return o.ToString();
        }


        public bool UpdateSheet(string templateFilePath, string sheetName, Report report, out string newfileName)
        {
            bool fileExists = System.IO.File.Exists(templateFilePath);
            if (!System.IO.Directory.Exists(System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName)))
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName));
            }

            newfileName = Path.GetFileName(templateFilePath).Replace(".xlsx", DateTime.Now.ToString("_yyyyMMddTHHmmssfff") + ".xlsx");
            string reportFilePath = System.IO.Path.GetFullPath(ExcelConstants.ReportFolderName + "\\" + newfileName);
            System.IO.File.Copy(templateFilePath, reportFilePath);
            UpdateCell(reportFilePath, report);
            return true;
        }

        private void UpdateCell(string fileName, Report report)
        {

            if (report.ReportVal.Count > 0)
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
                {
                    var headerListCells = GetAllHeaderCellList(spreadSheet, report.Name, ExcelConstants.ReportHeaderLastRowIndex);

                    WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, report.Name);

                    if (worksheetPart != null)
                    {
                        var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Skip(ExcelConstants.ReportHeaderLastRowIndex);
                        foreach (Row row in rows)
                        {
                            int id;
                            if (Int32.TryParse((row.ChildElements[1] as Cell).InnerText, out id))
                            {
                                SharedStringItem item = spreadSheet.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                if (item.Text != null)
                                {
                                    int reportValCount = 0;
                                    int cellValue = 0;
                                    Int32.TryParse(item.Text.Text, out cellValue);
                                    foreach (var reportRow in report.ReportVal.Where(rp => rp.ReportRow == cellValue))
                                    {
                                        if (reportRow != null)
                                        {
                                            reportValCount++;
                                            var headerNameKV = headerListCells.Where(hc => hc.Key == reportRow.ReportCol).FirstOrDefault();
                                            Cell cell = row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, headerNameKV.Value + row.RowIndex, true) == 0).First();
                                            cell.CellValue = new CellValue(reportRow.Val);
                                            cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                                        }
                                    }
                                    report.ReportVal.RemoveRange(0, reportValCount);
                                }
                            }
                            if (report.ReportVal.Count == 0)
                            {
                                break;
                            }
                        }
                        // Save the worksheet.
                        worksheetPart.Worksheet.Save();
                    }
                }
            }
        }

        private WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }




        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {

            Row row = GetRow(worksheet, rowIndex);
            if (row == null)
            {
                return null;
            }

            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
        }

        public List<KeyValuePair<int, string>> GetAllHeaderCellList(SpreadsheetDocument worksheet, string sheetName, uint rowIndex)
        {
            List<KeyValuePair<int, string>> celList = new List<KeyValuePair<int, string>>();
            WorksheetPart worksheetPart = GetWorksheetPartByName(worksheet, sheetName);

            Row row = GetRow(worksheetPart.Worksheet, rowIndex);
            foreach (Cell thecurrentcell in row)
            {
                string currentcellvalue = string.Empty;
                if (thecurrentcell.DataType != null)
                {
                    if (thecurrentcell.DataType == CellValues.SharedString)
                    {
                        int id;
                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                        {
                            if (id > 0)
                            {
                                SharedStringItem item = worksheet.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                int idCellVal;
                                if (Int32.TryParse(item.Text.Text, out idCellVal))
                                {
                                    if (idCellVal > 0)
                                    {
                                        celList.Add(new KeyValuePair<int, string>(idCellVal, thecurrentcell.CellReference.Value.Substring(0, thecurrentcell.CellReference.Value.Length - row.RowIndex.ToString().Length)));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return celList;
        }



        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }


    }



}
