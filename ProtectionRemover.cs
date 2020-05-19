using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Packaging;

public static void UnprotectExcel(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(Package.Open(filePath,FileMode.Open)))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                try
                {
                    Debug.Print($"{wbPart.Workbook.WorkbookProtection.OuterXml}");
                    wbPart.Workbook.WorkbookProtection.LockStructure.Value = false;
                }
                catch (Exception)
                {
                    Debug.Print($"This workbook does not contain document-level protection");
                }
                
                IEnumerable<WorksheetPart> worksheetParts = document.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    try
                    {
                        SheetProtection sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().FirstOrDefault();
                        Debug.Print($"{sheetProtection.OuterXml}");
                        worksheetPart.Worksheet.RemoveAllChildren<SheetProtection>();
                    }
                    catch (ArgumentNullException)
                    {
                        Debug.Print($"{worksheetPart.Worksheet.LocalName} does not contain sheetProtection");
                    }
                }
                document.Close();
                Debug.Print("Done...");
                //Open the unprotected document.
                Globals.ThisAddIn.Application.Workbooks.Open(filePath);
            }
        }
