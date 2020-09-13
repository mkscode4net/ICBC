using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ICBC.Web.Code
{
    public class ExcelHelper
    {
        public static void WriteExcelFile(string fileName, string tabName, string reportRow, string reportCol, string cellValue)
        {
            StringBuilder excelResult = new StringBuilder();
            try
            {
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(System.IO.Path.GetFullPath(ExcelConstants.TemplateFolderName + "\\" + fileName), false))
                {
                    List<string> dat = new List<string>();
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    Sheet thesheet = (Sheet)thesheetcollection.Where(sh => (sh as Sheet).Name.Equals(tabName)).FirstOrDefault();
                    //using for each loop to get the sheet from the sheetcollection  
                    // foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        var allBCells = thesheetdata.Select(c => int.Parse( c.ChildElements[1].InnerText)).ToArray();
                        foreach (var id in allBCells)
                        {
                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                            if (item.Text != null && dat.Contains(item.Text.Text) || dat.Contains(item.InnerText))
                            {
                                //code to take the string value  
                                excelResult.Append(item.Text.Text + " ");
                            }
                        }

                        //   Row _thecurrentrow = thesheetdata.Where(rw=> (rw as Cell).CellValue);
                        foreach (Row thecurrentrow in thesheetdata)
                        {
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                //statement to take the integer value  
                                string currentcellvalue = string.Empty;
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //code to take the string value  
                                                excelResult.Append(item.Text.Text + " ");
                                            }
                                            else if (item.InnerText != null)
                                            {
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                                }
                            }
                            excelResult.AppendLine();
                        }
                        excelResult.Append("");
                        Console.WriteLine(excelResult.ToString());
                        Console.ReadLine();
                    }
                }
            }
            catch (Exception ex)
            {

            }
           // return excelResult.ToString();
        }

        public static string ReadExcelFile(string fileName, string tabName)
        {
            JObject o = new JObject();
            JArray arrayRows = new JArray();

            StringBuilder excelResult = new StringBuilder();
            try
            {
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    Sheet thesheet = (Sheet)thesheetcollection.Where(sh => (sh as Sheet).Name.Equals(tabName)).FirstOrDefault();
                    //using for each loop to get the sheet from the sheetcollection  
                   // foreach (Sheet thesheet in thesheetcollection)
                    {
                        excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                        excelResult.AppendLine("----------------------------------------------- ");
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                        //var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().Skip(ExcelConstants.ReportHeaderLastRowIndex);
                        foreach (Row thecurrentrow in thesheetdata.Skip(ExcelConstants.ReportHeaderLastRowIndex))
                        {
                            JObject ocell = new JObject();
                            // Cell thecurrentcell = thecurrentrow.ChildElements[ExcelConstants.ReportColumnStartIndex] as Cell;
                            foreach (Cell thecurrentcell in thecurrentrow)
                            {
                                
                                //statement to take the integer value  
                                string currentcellvalue = string.Empty;
                                //if (thecurrentcell.DataType != null)
                                {
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
                                                currentcellvalue = item.InnerText;
                                            }
                                            else if (item.InnerXml != null)
                                            {
                                                currentcellvalue = item.InnerXml;
                                            }
                                        }
                                    }
                                    else //if (thecurrentcell.DataType == CellValues.Number)
                                    {
                                        ocell[string.Concat(thecurrentcell.CellReference.Value.Reverse().Skip(thecurrentrow.RowIndex.ToString().Length).Reverse())] = thecurrentcell.InnerText;
                                    }
                                }
                                //}
                                //else
                                //{
                                //    ocell[thecurrentcell.CellReference] = string.Empty;
                                //}
                               
                            }

                            arrayRows.Add(ocell);
                            excelResult.AppendLine();
                        }
                        excelResult.Append("");
                    }
                }
                o["Data"] = arrayRows;
            }
            catch (Exception ex)
            {

            }
            return o.ToString();
        }

    }
}
