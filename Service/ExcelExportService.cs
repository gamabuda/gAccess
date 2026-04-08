using ClosedXML.Excel;
using gAcss.Models;

namespace gAcss.Service
{
    public class ExcelExportService
    {
        public void Export(IEnumerable<DriveFileEntry> data, string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Drive Analytics");

            var table = ws.Cell(1, 1).InsertTable(data);

            ws.Row(1).Style.Font.Bold = true;
            ws.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#4F81BD");
            ws.Row(1).Style.Font.FontColor = XLColor.White;

            ws.Column(7).Style.DateFormat.Format = "yyyy-mm-dd hh:mm";

            ws.Columns().AdjustToContents();
            workbook.SaveAs(filePath);
        }
    }
}