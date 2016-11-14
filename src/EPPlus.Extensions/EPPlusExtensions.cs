using System.Data;
using System.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

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
            try
            {
                var headerRow = firstRowContainsHeader ? 1 : 0;
                return ToDataSet(package, headerRow);
            }
            catch (Exception)
            {
                throw;
            }
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
                throw new ArgumentOutOfRangeException(nameof(headerRow), headerRow, "Must be 0 or greater.");
            }

            try
            {
                var result = new DataSet();


                foreach (var sheet in package.Workbook.Worksheets)
                {
                    var table = ToDataTable(sheet, headerRow);


                    result.Tables.Add(table);
                }


                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static DataTable ToDataTable(this ExcelWorksheet sheet, int headerRow = 0, Func<ExcelRangeBase, bool> footerRowPredicate = null, int startColumn = 1, int endColumn = -1, List<string> columnNames = null)
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

            int maxColumnIndex = endColumn == -1 ? sheet.Dimension.End.Column : endColumn;

            IEnumerable<DataColumn> columns = null;

            //If user provides column names, we use them
            if (columnNames == null)
            {
                columns = from firstRowCell in sheet.Cells[sheetStartRow, startColumn, sheetStartRow, maxColumnIndex]
                          select new DataColumn(headerRow > 0 ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
            }
            else
            {
                columns = columnNames.Select(columnName => new DataColumn(columnName));
            }

            table.Columns.AddRange(columns.ToArray());

            var startRow = headerRow > 0 ? sheetStartRow + 1 : sheetStartRow;

            for (var rowIndex = startRow; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
            {

                var inputRow = sheet.Cells[rowIndex, startColumn, rowIndex, maxColumnIndex];

                //if user has provided a predicate that determines if a row is a 'footer'
                if (footerRowPredicate != null)
                {
                    //If any cell in the current row satisifies the predicate, we bail out and consider that we have reached the footer row.
                    if (inputRow.Any(cell => footerRowPredicate(cell)))
                    {
                        return table;
                    }
                }

                var row = table.Rows.Add();
                int colIndex = 0;

                foreach (var cell in inputRow)
                {
                    row[colIndex] = cell.Text;
                    colIndex++;
                }
            }

            return table;
        }

    }
}
