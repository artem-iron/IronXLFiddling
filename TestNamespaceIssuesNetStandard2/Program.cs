using System.Drawing;
using IronXL;
using IronXL.Drawing.Images;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel.Charts;
using NPOI.Util;

namespace TestNamespaceIssues
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WorkBook assss = WorkBook.Load("C:\\Users\\artem\\Downloads\\TestData.xlsx");

            WorkSheet ass = assss.DefaultWorkSheet;

            IImage pic = ass.Images[0];

            if (pic != null)

            {

                var bitmap = pic.ImageFormat; //It fails here with an exception, the details of which I sent you earlier.

                assss.SaveAs($"C:\\Users\\artem\\Downloads\\{bitmap}.bmp");
            }
        }
    }
}
