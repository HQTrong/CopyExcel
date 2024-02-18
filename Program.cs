using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static void Main()
    {
        string sourceFilePath = @"C:\Users\TRONG\Desktop\Input\input.xlsx";
        string destinationFilePath = @"C:\Users\TRONG\Desktop\Output\output.xlsx";
        List<string> sheetNames = new List<string>();
        sheetNames.Add("Employee");
        sheetNames.Add("Employee1");
        sheetNames.Add("Employee2");

        try
        {
            foreach(string sheetName in sheetNames)
            {
                AddOrUpdateSheet(sourceFilePath, destinationFilePath, sheetName);
            }
            Console.WriteLine("DONE");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            Console.ReadLine();
        }
    }

    public static void AddOrUpdateSheet(string sourceFilePath, string destinationFilePath, string sheetName)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook;

        try
        {
            if (!File.Exists(destinationFilePath))
            {
                workbook = excelApp.Workbooks.Open(sourceFilePath);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Sheets[1];
                oSheet.Name = sheetName;
                CopyFormat(oSheet, excelApp, 12);
                workbook.SaveAs(destinationFilePath);
                // Save changes and close Excel
                workbook.Close();
                excelApp.Quit();
            }
            else
            {
                workbook = excelApp.Workbooks.Open(destinationFilePath);
                if (!SheetExists(workbook, sheetName))
                {
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Sheets[1];
                    oSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];
                    newSheet.Name = sheetName;
                    CopyFormat(newSheet, excelApp, 12);
                    // Save changes and close Excel
                    workbook.Save();
                }
                workbook.Close();
                excelApp.Quit();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"An error occurred: {ex.Message}");
        }
    }

    public static bool SheetExists(Excel.Workbook workbook, string sheetName)
    {
        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            if (sheet.Name == sheetName)
            {
                return true;
            }
        }
        return false;
    }

    public static void CopyFormat(Excel.Worksheet oSheet, Excel.Application excelApp, int numberOfRows)
    {
        // Get format of row 6
        Excel.Range sourceFormatRange = oSheet.Rows[6];
        sourceFormatRange.Copy();

        // Create N rows
        for (int i = 1; i <= numberOfRows; i++)
        {
            Excel.Range targetRange = oSheet.Rows[6 + i];
            targetRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        }
        excelApp.CutCopyMode = 0;
    }
}
