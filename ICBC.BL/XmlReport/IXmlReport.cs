using System;
using System.Collections.Generic;
using System.Text;
using ICBC.Entity;

namespace ICBC.BL.XmlReport
{
    interface IXmlReport
    {
        string ReadExcelFile(string fileName, string sheetName);
        bool UpdateSheet(string templateFilePath, string sheetName, Report report, out string newfileName);
    }
}
