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
            IronXL.Xml.Spreadsheet.CT_Authors authors = new IronXL.Xml.Spreadsheet.CT_Authors();

            WorkSheet ass = WorkBook.Load("C:\\Users\\artem\\Downloads\\TestData.xlsx").DefaultWorkSheet;

            IImage pic = ass.Images[0];

            if (pic != null)

            {

                Bitmap bitmap = pic.ToBitmap(); //It fails here with an exception, the details of which I sent you earlier.

                bitmap.Save("C:\\Users\\artem\\Downloads\\ass.bmp", System.Drawing.Imaging.ImageFormat.Bmp);

                bitmap.Dispose();

            }
        }
    }
}
