using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Math;
using FileToDataset.Interfaces;

namespace FileToDataset.Factories
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

        private static FileToDataset.Logics.ExcelConverter GetExcelConverter(string filePath)
        {
            var excelByte = File.ReadAllBytes(filePath);
            MemoryStream stream = new MemoryStream(excelByte); 
            return new FileToDataset.Logics.ExcelConverter(stream);
        }

        private static FileToDataset.Logics.CSVConverter CSVConverter(string filePath,string delimiter)
        {
            var lines = File.ReadAllLines(filePath);
            return new FileToDataset.Logics.CSVConverter(lines,delimiter,filePath);
        }
    }
}
