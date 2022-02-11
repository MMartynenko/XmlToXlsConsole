using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using System.Collections.Generic;

namespace XmlToXlsConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                System.Console.WriteLine("Please add path to xml file as first argument.");
                return;
            }

            string xmlFilePath = args[0];
            string customFileName = null;
            if (args.Length > 1)
            {
                customFileName = args[1];
            }
            

            if (!String.IsNullOrEmpty(customFileName) && xmlFilePath != "") // using Custome Xml File Name  
            {
                if (File.Exists(xmlFilePath))
                {
                    string CustXmlFilePath = Path.Combine(new FileInfo(xmlFilePath).DirectoryName, customFileName); // Ceating Path for Xml Files  
                    XmlNodeList dt = CreateDataTableFromXml(xmlFilePath);
                    ExportDataTableToExcel(dt, CustXmlFilePath);

                    System.Console.WriteLine("Conversion completed.");
                }

            }
            else if (String.IsNullOrEmpty(customFileName) || xmlFilePath != "") // Using Default Xml File Name  
            {
                if (File.Exists(xmlFilePath))
                {
                    FileInfo fi = new FileInfo(xmlFilePath);
                    string XlFile = fi.DirectoryName + "\\" + fi.Name.Replace(fi.Extension, ".xlsx");
                    XmlNodeList dt = CreateDataTableFromXml(xmlFilePath);
                    ExportDataTableToExcel(dt, XlFile);

                    System.Console.WriteLine("Conversion completed.");
                }
            } else
            {
                System.Console.WriteLine("Please add correct arguments:");
                System.Console.WriteLine("- File path [Required]");
                System.Console.WriteLine("- Converted file name [Optional]");
            }

            return;
        }

        public static XmlNodeList CreateDataTableFromXml(string XmlFile)
        {

            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(XmlFile);
                return doc.GetElementsByTagName("Worksheet");

            }
            catch (Exception ex)
            {

            }
            return null;
        }

        private static void ExportDataTableToExcel(XmlNodeList table, string Xlfile)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excel.Application.Workbooks.Add(Type.Missing);
            excel.Visible = false;
            excel.DisplayAlerts = false;

            for (int i = 0; i < table.Count; i++)
            {
                Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
                excelWorkSheet.Name = table[i].Attributes["ss:Name"].Value;

                List<XmlNode> rows = new List<XmlNode>();
                XmlNodeList children = table[i].FirstChild.ChildNodes;
                foreach (XmlNode child in children)
                {
                    if (child.Name == "Row") rows.Add(child);
                }
                
                for (int j = 0; j < rows.Count; j++)
                {
                    XmlNode row = rows[j];
                    int column = 1;
                    foreach (XmlNode c in row.ChildNodes)
                    {
                        if (c.Name == "Cell")
                        {
                            if (c.Attributes["ss:Index"] != null) column = Int32.Parse(c.Attributes["ss:Index"].Value);
                            excelWorkSheet.Cells[j + 1, column] = c.InnerText;
                            column++;                            
                        }
                    }
                }

                if (i < table.Count - 1)
                {
                    book.Worksheets.Add();
                }
            }

            book.SaveAs(Xlfile);
            book.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(excel);
        }

    }
}
