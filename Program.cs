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
            excelPackage.SaveAs(destination);
        }
    }
}
