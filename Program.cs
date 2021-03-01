using System;
using System.IO;
using System.Linq;
using ClosedXML;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace MultipleFixedImages
{
    class Program
    {
        static void Main(string[] args)
        {
            var image = File.ReadAllBytes("test.jpg");
            WriteSingleImage(image, "SingleImage.xlsx");
            WriteMultipleImages(image, "MultipleImages.xlsx");
        }

        static void WriteSingleImage(byte[] image, string fileName)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("Images");
            var firstCell = sheet.Cell(1, 1);
            var picture = sheet.AddPicture(new MemoryStream(image));
            picture.MoveTo(firstCell);
            picture.Placement = XLPicturePlacement.MoveAndSize;
            workbook.SaveAs(fileName);
        }

        static void WriteMultipleImages(byte[] image, string fileName)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("Images");
            foreach (var rowIndex in Enumerable.Range(1, 5))
            {
                var currentCell = sheet.Row(rowIndex).Cell(1);
                var picture = sheet.AddPicture(new MemoryStream(image));
                picture.MoveTo(currentCell);
                picture.Placement = XLPicturePlacement.MoveAndSize;
            }

            workbook.SaveAs(fileName);
        }
    }
}
