using System.Data;
using System.Linq;
using OfficeOpenXml;

namespace EPPlusExtensions
{
    // ReSharper disable once InconsistentNaming
    public static class EPPlusExtensions
    {
        public static DataSet ToDataSet(this ExcelPackage package, bool firstRowContainsHeader = false)
        {
            var result = new DataSet();

            foreach (var sheet in package.Workbook.Worksheets)
            {
                var table = new DataTable { TableName = sheet.Name };

                var columns = from firstRowCell in sheet.Cells[1, 1, 1, sheet.Dimension.End.Column]
                              select new DataColumn(firstRowContainsHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");

                table.Columns.AddRange(columns.ToArray());

                var startRow = firstRowContainsHeader ? 2 : 1;

                for (var rowIndex = startRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                {
                    var inputRow = sheet.Cells[rowIndex, 1, rowIndex, sheet.Dimension.End.Column];
                    var row = table.Rows.Add();
                    foreach (var cell in inputRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }

                result.Tables.Add(table);
            }

            return result;
        }
    }
}
