using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Math;
using ExcelConverter.Interfaces;

namespace ExcelConverter.Factories
{
    public static class ConverterFactory
    { 
        public enum FileType
        {
            Excel,
            CsvOrTxt 
        }
        public static IConverter Create(FileType fileType,string filePath,string delimiter = ",")
        {
            return fileType switch
            {
                FileType.Excel => GetExcelConverter(filePath),
                FileType.CsvOrTxt => CSVConverter(filePath,delimiter),
                _ => throw new NotSupportedException(),
            };
        }

        private static ExcelConverter.Logics.ExcelConverter GetExcelConverter(string filePath)
        {
            var excelByte = File.ReadAllBytes(filePath);
            MemoryStream stream = new MemoryStream(excelByte); 
            return new ExcelConverter.Logics.ExcelConverter(stream);
        }

        private static ExcelConverter.Logics.CSVConverter CSVConverter(string filePath,string delimiter)
        {
            var lines = File.ReadAllLines(filePath);
            return new ExcelConverter.Logics.CSVConverter(lines,delimiter);
        }
    }
}
