using System.Data;
using System.Linq;
using OfficeOpenXml;
using System;

namespace EPPlus.Extensions
{
    // ReSharper disable once InconsistentNaming 
    public static class Extensions
    {
        public static DataSet ToDataSet(this ExcelPackage package, bool firstRowContainsHeader = false)
        {
            var headerRow = firstRowContainsHeader ? 1 : 0;
            return ToDataSet(package, headerRow);
        }


        public static DataSet ToDataSet(this ExcelPackage package, int headerRow = 0)
        {
            if (headerRow < 0)
            {
                throw new ArgumentException("headerRow must be zero or greater.");
            }
            var result = new DataSet();


            foreach (var sheet in package.Workbook.Worksheets)
            {
                var table = new DataTable { TableName = sheet.Name };

                int sheetStartRow = 1;
                if (headerRow > 0)
                {
                    sheetStartRow = headerRow;
                }
                var columns = from firstRowCell in sheet.Cells[sheetStartRow, 1, sheetStartRow, sheet.Dimension.End.Column]
                              select new DataColumn(headerRow > 0 ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");


                table.Columns.AddRange(columns.ToArray());


                var startRow = headerRow > 0 ? sheetStartRow + 1 : sheetStartRow;


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