using System;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using Shouldly;

namespace EPPlus.Extensions.Tests
{
    public class ExtensionTests
    {
        public void ToDataSetSimple_ShouldHandleHeaderRows_WhenSpecified()
        {
            var package = GetMarvelPackage();

            var result = package.ToDataSet(true);

            result.Tables[0].Rows.Count.ShouldBe(10);
        }

        public void ToDataSetSimple_ShouldHandleHeaderRows_WhenNotSpecified()
        {
            var package = GetMarvelPackage();

            var result = package.ToDataSet(false);

            result.Tables[0].Rows.Count.ShouldBe(11);
        }

        public void ToDataSet_ShouldThrowException_WhenHeaderRowIsLessThanZero()
        {
            var package = new ExcelPackage();

            var exception = Should.Throw<ArgumentOutOfRangeException>(() => package.ToDataSet(-1));
            exception.ParamName.ShouldBe("headerRow");
        }

        public void ToDataSet_ShouldReturnOneTable_WhenOneSheet()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(0);

            result.Tables.Count.ShouldBe(1);
        }

        public void ToDataSet_ShouldReturnTwoTables_WhenTwoSheets()
        {
            var package = GetMarvelPackage();

            var result = package.ToDataSet(0);

            result.Tables.Count.ShouldBe(2);
        }

        public void ToDataSet_ShouldHandleHeaderRows_WhenSetToZero()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(0);

            result.Tables[0].Rows.Count.ShouldBe(50);
        }

        public void ToDataSet_ShouldHandleHeaderRows_WhenSetToOne()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(1);

            result.Tables[0].Rows.Count.ShouldBe(49);
        }

        public void ToDataSet_ShouldHandleHeaderRows_WhenSetToTen()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(10);

            result.Tables[0].Rows.Count.ShouldBe(40);
        }

        public void ToDataSet_ShouldNameColumnsWithHeaderValues_WhenHeaderValuesExist()
        {
            var package = GetMarvelPackage();

            var result = package.ToDataSet(1);

            result.Tables[0].Columns[0].ColumnName.ShouldBe("First Name");
            result.Tables[0].Columns[1].ColumnName.ShouldBe("Last Name");
            result.Tables[0].Columns[2].ColumnName.ShouldBe("Alter Ego");
        }

        public void ToDataSet_ShouldUseGenericColumnNames_WhenHeaderValuesDoNotExist()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(0);

            result.Tables[0].Columns[0].ColumnName.ShouldBe("Column 1");
        }

        public void ToDataSet_ShouldAddColumns_ForEachSourceColumn()
        {
            var package = GetMarvelPackage();

            var result = package.ToDataSet(0);

            result.Tables[0].Columns.Count.ShouldBe(3);
        }

        public void ToDataSet_ShouldAddRows_ForEachSourceRow()
        {
            var package = GetStatesPackage();

            var result = package.ToDataSet(0);

            result.Tables[0].Rows.Count.ShouldBe(50);
            result.Tables[0].Rows[0][0].ToString().ShouldBe("Alabama");
            result.Tables[0].Rows[49][0].ToString().ShouldBe("Wyoming");
        }

        public void ToDataSet_WhenCsv_AllowsDateFormat()
        {
            var package = GetMarvelCsvPackage();

            var result = package.ToDataSet(true);

            var cellValue = result.Tables[0].Rows[0][3];

            DateTime temp;
            DateTime.TryParse(cellValue.ToString(), out temp).ShouldBeTrue($"Input was {cellValue}");

            Convert.ToDateTime(cellValue).ToString("o").ShouldBe(DateTime.Parse("04/22/1950 08:41 PM").ToString("o"));
        }

        public void IsLastRowEmpty_ReturnsFalse_WhenLastRowIsNotEmpty()
        {
            var package = GetMarvelCsvPackage();
            package.Workbook.Worksheets["Marvel"].IsLastRowEmpty().ShouldBeFalse();
        }

        public void IsLastRowEmpty_ReturnsTrue_WhenLastRowIsEmpty()
        {
            var package = GetMarvelWithSpacesCsvPackage();
            package.Workbook.Worksheets["Marvel"].IsLastRowEmpty().ShouldBeTrue();
        }

        public void TrimLastEmptyRows_RemovesEmptyRows()
        {
            var package = GetMarvelWithSpacesCsvPackage();
            var sheet = package.Workbook.Worksheets["Marvel"];

            sheet.TrimLastEmptyRows();

            sheet.IsLastRowEmpty().ShouldBeFalse();
            sheet.Dimension.End.Row.ShouldBe(11);
        }

        public void TrimLastEmptyRows_DoesRemoveNotEmptyRows()
        {
            var package = GetMarvelCsvPackage();
            var sheet = package.Workbook.Worksheets["Marvel"];

            sheet.TrimLastEmptyRows();

            sheet.IsLastRowEmpty().ShouldBeFalse();
            sheet.Dimension.End.Row.ShouldBe(11);
        }

        private static ExcelPackage GetMarvelPackage()
        {
            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Marvel.xlsx");
            var file = new FileInfo(path);
            var package = new ExcelPackage(file);
            return package;
        }

        private static ExcelPackage GetMarvelCsvPackage()
        {
            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Marvel.csv");
            var file = new FileInfo(path);
            var package = new ExcelPackage();

            var textFormat = new ExcelTextFormat { TextQualifier = '"' };

            var sheet = package.Workbook.Worksheets.Add("Marvel");
            sheet.Cells.LoadFromText(File.ReadAllText(file.FullName), textFormat);

            return package;
        }

        private static ExcelPackage GetMarvelWithSpacesCsvPackage()
        {
            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "MarvelWithSpaces.csv");
            var file = new FileInfo(path);
            var package = new ExcelPackage();

            var textFormat = new ExcelTextFormat { TextQualifier = '"' };

            var sheet = package.Workbook.Worksheets.Add("Marvel");
            sheet.Cells.LoadFromText(File.ReadAllText(file.FullName), textFormat);

            return package;
        }

        private static ExcelPackage GetStatesPackage()
        {
            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"States.xlsx");
            var file = new FileInfo(path);
            var package = new ExcelPackage(file);
            return package;
        }
    }
}