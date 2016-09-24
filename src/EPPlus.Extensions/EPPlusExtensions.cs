using System.Data;
using System.Linq;
using OfficeOpenXml;
using System;

namespace EPPlus.Extensions
{
    // ReSharper disable once InconsistentNaming 
    public static class Extensions
    {
        /// <summary>
        /// Extracts a DataSet from the ExcelPackage.
        /// </summary>
        /// <param name="package">The Excel package.</param>
        /// <param name="firstRowContainsHeader">if set to <c>true</c> [first row contains header].</param>
        /// <returns></returns>
        public static DataSet ToDataSet(this ExcelPackage package, bool firstRowContainsHeader = false)
        {
            var headerRow = firstRowContainsHeader ? 1 : 0;
            return ToDataSet(package, headerRow);
        }


        /// <summary>
        /// Extracts a DataSet from the ExcelPackage.
        /// </summary>
        /// <param name="package">The Excel package.</param>
        /// <param name="headerRow">The header row. Use 0 if there is no header row. Value must be 0 or greater.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">headerRow must be 0 or greater.</exception>
        public static DataSet ToDataSet(this ExcelPackage package, int headerRow = 0)
        {
            if (headerRow < 0)
            {
                throw new ArgumentException("headerRow must be 0 or greater.");
            }
            var result = new DataSet();


            foreach (var sheet in package.Workbook.Worksheets)
            {
                var table = ToDataTable(sheet, headerRow);


                result.Tables.Add(table);
            }


            return result;
        }

        /// <summary>
        /// Extracts a DataTable from the ExcelWorksheet object.
        /// </summary>
        /// <param name="package">The <see cref="ExcelWorksheet"/> object.</param>
        /// <param name="sheet">if set to <c>true</c> [first row contains header].</param>
        /// <returns></returns>
        public static DataTable ToDataTable(this ExcelWorksheet sheet, bool firstRowContainsHeader = false)
        {
            var headerRow = firstRowContainsHeader ? 1 : 0;
            return ToDataTable(sheet, headerRow);
        }


        /// <summary>
        /// Extracts a DataTable from the ExcelWorksheet object.
        /// </summary>
        /// <param name="sheet">The Excel package.</param>
        /// <param name="headerRow">The header row. Use 0 if there is no header row. Value must be 0 or greater.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">headerRow must be 0 or greater.</exception>
        public static DataTable ToDataTable(this ExcelWorksheet sheet, int headerRow = 0)
        {
            if (headerRow < 0)
            {
                throw new ArgumentException("headerRow must be 0 or greater.");
            }

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

            return table;
        }
    }
}