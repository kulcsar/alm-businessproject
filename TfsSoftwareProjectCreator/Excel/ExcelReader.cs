using System;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace TfsSoftwareProjectCreator.Excel
{
    /// <summary>
    /// Provides Excel reading functionality
    /// </summary>
    public class ExcelReader
    {
        /// <summary>
        /// Read Excel content
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="workSheetName"></param>
        /// <returns></returns>
        public static object[,] GetExcelContent(string filePath, string workSheetName)
        {
            // Create COM Objects
            ExcelInterop.Application xlApp = new ExcelInterop.Application();
            ExcelInterop.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            ExcelInterop.Sheets xlSheets = xlWorkbook.Worksheets;
            ExcelInterop.Worksheet xlWorksheet = (ExcelInterop.Worksheet)xlSheets.get_Item(workSheetName);
            ExcelInterop.Range xlRange = xlWorksheet.UsedRange;

            // Get content
            object[,] valueArray = (object[,])xlRange.get_Value(ExcelInterop.XlRangeValueDataType.xlRangeValueDefault);

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            // Close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            // Quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return valueArray;
        }
    }
}
