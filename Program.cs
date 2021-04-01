using OfficeOpenXml;
using System.IO;

namespace EPPLUS_Bug
{
    class Program
    {
        static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var source = new FileInfo(@"./file.xlsx");
            var destination = new FileInfo(@"./file_saved.xlsx");

            using ExcelPackage excelPackage = new ExcelPackage(source);

            //Setting this breaks the pinned headers when all i want it to do is just scroll the view into A56
            //It also breaks the ability to scroll up for all cells before the 56th
            excelPackage.Workbook.Worksheets[0].View.TopLeftCell = "A56";

            excelPackage.SaveAs(destination);
        }
    }
}
