using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.IO;

namespace EPPLUS_Bug
{
    class Program
    {
        static void Main()
        {
            #region Test setup
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var source = new FileInfo(@"./file.xlsx");
            var destination = new FileInfo(@"./file_saved.xlsx");
            var destination_expected = new FileInfo(@"./file_saved_expected.xlsx");

            if (File.Exists(destination.FullName))
                File.Delete(destination.FullName);

            if (File.Exists(destination_expected.FullName))
                File.Delete(destination_expected.FullName);
            #endregion

            // EPPLUS TEST
            using ExcelPackage excelPackage = new ExcelPackage(source);
            excelPackage.Workbook.Worksheets[0].View.TopLeftCell = "A56";
            excelPackage.SaveAs(destination);

            // EXCEL COM TEST
            DoExcelComTest(source, destination_expected);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        //Helper method so we don't Zombify the Excel process during debug
        private static void DoExcelComTest(FileInfo source, FileInfo destination_expected)
        {
            Application xlApplication = new Application();
            Workbook workBook = xlApplication.Workbooks.Open(source.FullName);
            Worksheet workSheet = (Worksheet)workBook.Sheets[1];

            workSheet.Range["A56", "A56"].Select();
            workBook.SaveAs(destination_expected.FullName);
            xlApplication.Quit();
        }
    }
}
