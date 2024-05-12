using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExperimentCalculations.Utils
{
    internal class XlsxUtils
    {
        public static ExcelPackage CreateExcelPackage()
        {
            return new ExcelPackage();
        }

        public static void FillCell(ExcelWorksheet worksheet, int row, int column, string value, bool isBold)
        {
            var cell = worksheet.Cells[row, column];
            cell.Value = value;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            if (isBold)
            {
                cell.Style.Font.Bold = isBold;
            }
        }
    }
}
