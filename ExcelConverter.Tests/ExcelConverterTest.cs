using ExcelConverter.Factories;
using ExcelConverter.Interfaces;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit.Abstractions;

namespace ExcelConverterTest.Tests
{
    public class ExcelConverterTest
    {
        [Fact]
        public void TestConvertExcelToDataSet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            string tableName = "Test1";
            var file = new FileInfo(tempFile);
            using var excelPackage = new OfficeOpenXml.ExcelPackage(file);
            var worksheet = excelPackage.Workbook.Worksheets.Add(tableName); 
            excelPackage.Save();

            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.Excel,tempFile);

            DataSet dataset = convert.GetDataSet();
            Assert.NotNull(dataset);
            Assert.NotNull(dataset.Tables[tableName]);
        }

        [Fact]
        public void TestConvertExcelToMultipleDataSet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            var file = new FileInfo(tempFile);
            using var excelPackage = new OfficeOpenXml.ExcelPackage(file);


            string[] tableNames = new string[]{ "Test1", "Test2", "Test3" };
            foreach(var item in tableNames)
            {
                excelPackage.Workbook.Worksheets.Add(item);
                excelPackage.Save();
            }

            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.Excel, tempFile);

            DataSet dataset = convert.GetDataSet();
            Assert.NotNull(dataset);
            Assert.Equal(tableNames.Length,dataset.Tables.Count);

            foreach (var item in tableNames)
            {
                Assert.NotNull(dataset.Tables[item]);
                Assert.NotNull(convert.GetDataTable(item));
            }
        }

        [Fact]
        public void TestConvertExcelToDataSet_GetValue()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            string tableName = "Test1";
            var file = new FileInfo(tempFile);
            using var excelPackage = new OfficeOpenXml.ExcelPackage(file);
            var worksheet = excelPackage.Workbook.Worksheets.Add(tableName);
            worksheet.Cells[1, 1].Value = "idTest";
            worksheet.Cells[1, 2].Value = "NameTest";
            worksheet.Cells[2, 1].Value = "123";
            worksheet.Cells[2, 2].Value = "Test123";
            excelPackage.Save();

            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.Excel, tempFile);

            DataTable dt = convert.GetDataTable(tableName);
            Assert.Equal(dt.Rows[0]["A"], "idTest");
            Assert.Equal(dt.Rows[0]["B"], "NameTest");
            Assert.Equal(dt.Rows[1]["A"], "123");
            Assert.Equal(dt.Rows[1]["B"], "Test123");
        }
    }
}
