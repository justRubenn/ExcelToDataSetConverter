using System.Data;
using ExcelConverter.Factories;
using ExcelConverter.Interfaces;
using static ExcelConverter.Factories.ConverterFactory;

namespace ExcelConverter
{

    public class ConvertToDataTable(ConverterFactory.FileType fileType, string filePath, string delimiter = ",")
    {
        private IConverter Converter { get; set; } = ConverterFactory.Create(fileType, filePath, delimiter);
        public DataTable GetDataTable(string tableName) => Converter.GetDataTable(tableName);
        public DataSet GetDataSet() => Converter.GetDataSet();
    }
}
