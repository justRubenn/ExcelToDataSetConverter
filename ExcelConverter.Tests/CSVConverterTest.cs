using ExcelConverter.Factories;
using ExcelConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSVConverter.Tests
{
    public class CSVConverterTest
    { 
        private static string GetTempFileName(string ext) => Path.ChangeExtension(Path.GetTempFileName(), ext);
        [Fact]
        public void TestConvertCSVToDataset()
        {
            string filename = GetTempFileName("csv");

            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id,name,account");
                writer.WriteLine("1,name1,account1");
                writer.WriteLine("2,name2,account2");
                 
            };

            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt,filename);

            DataSet dataset = convert.GetDataSet();
            DataTable dt = convert.GetDataTable();
            string tableName = dt.TableName;

            Assert.NotNull(dataset);
            Assert.Single(dataset.Tables);
            Assert.NotNull(dt);
            Assert.Equal(tableName, Path.GetFileNameWithoutExtension(filename));

        }

        [Fact]
        public void TestConvertCSVToDatasetGetByTableName()
        {
            string filename = GetTempFileName("csv");
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id,name,account");
                writer.WriteLine("1,name1,account1");
                writer.WriteLine("2,name2,account2");

            }; 
            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename);

            DataSet dataset = convert.GetDataSet();
            DataTable dt = convert.GetDataTable(Path.GetFileNameWithoutExtension(filename));
            string tableName = dt.TableName;

            Assert.NotNull(dataset);
            Assert.Single(dataset.Tables);
            Assert.NotNull(dt);
            Assert.Equal(tableName, Path.GetFileNameWithoutExtension(filename));
        }

        [Fact]
        public void TestConvertCSVToDatasetGetByTableName_Negatif()
        {
            string filename = GetTempFileName("csv");
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id,name,account");
                writer.WriteLine("1,name1,account1");
                writer.WriteLine("2,name2,account2");

            }
            ;
            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename);

            DataSet dataset = convert.GetDataSet();
            DataTable dt = convert.GetDataTable(Path.GetFileNameWithoutExtension(filename) + "Test");
            string tableName = dt.TableName;

            Assert.NotNull(dataset);
            Assert.Single(dataset.Tables);
            Assert.Empty(dt.Rows); 
        }

        [Fact]
        public void TestConvertCSVToDatasetTableMustHaveExpectedColumnBasedOnDelimiter()
        {
            string filename = GetTempFileName("csv");
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id,name,account");
                writer.WriteLine("1,name1,account1");
                writer.WriteLine("2,name2,account2");

            };
            IConverter convert1 = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename);
            IConverter convert2 = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename,"|");

            Assert.Equal(3, convert1.GetDataTable().Columns.Count);
            Assert.Single(convert2.GetDataTable().Columns); 
        }

        [Fact]
        public void TestConvertCSVToDatasetTableColumnNameBasedOnItsSequence()
        {
            string filename = GetTempFileName("csv");
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id|name|account");
                writer.WriteLine("1|name1|account1");
                writer.WriteLine("2|name2|account2");

            }
            ; 
            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename, "|");
            DataTable dt = convert.GetDataTable();

            Assert.Equal(3, convert.GetDataTable().Columns.Count); 

            for(int i = 1; i <= dt.Columns.Count; i++)
            {
                Assert.Contains(dt.Columns.Cast<DataColumn>(), item => item.ColumnName == i.ToString());  
            }
        }

        [Fact]
        public void TestConvertCSVToDatasetColumnCountTableMustCreateFromTheLongestLine()
        {
            string filename = GetTempFileName("csv");
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Id|name|account");
                writer.WriteLine("1|name1|account1");
                writer.WriteLine("2|name2|account2");
                writer.WriteLine("2|name2|account2|1|2|3|4|5");

            }
            ;
            IConverter convert = ConverterFactory.Create(ConverterFactory.FileType.CsvOrTxt, filename, "|");
            DataTable dt = convert.GetDataTable();

            Assert.Equal(8, convert.GetDataTable().Columns.Count);

            for (int i = 1; i <= dt.Columns.Count; i++)
            {
                Assert.Contains(dt.Columns.Cast<DataColumn>(), item => item.ColumnName == i.ToString());
            }
        }


    }
}
