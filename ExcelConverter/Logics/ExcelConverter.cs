using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelConverter.Interfaces;

namespace ExcelConverter.Logics
{
    public class ExcelConverter : IConverter
    {
        private DataSet? ExcelDataSet { get; set; } = new DataSet();

        public ExcelConverter(MemoryStream stream)
        {
            ToDataSet(stream);
            //convert_excel(stream);
        }
        public DataTable GetDataTable(string tableName) => ExcelDataSet?.Tables[tableName] ?? new DataTable();
        public DataSet GetDataSet() => ExcelDataSet ?? new DataSet();
        private void ToDataSet(MemoryStream stream)
        {
            using MemoryStream ms = stream;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart!;
                foreach (Sheet sheet in workbookPart.Workbook.Sheets!)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
                    List<Dictionary<string, string>> ExcelListObj = new List<Dictionary<string, string>>();
                    string sheetName = sheet.Name;

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        Cell[] cellArr = row.Elements<Cell>().ToArray();
                        Dictionary<string, string> ExcelObj = new Dictionary<string, string>();
                        foreach (Cell cell in cellArr)
                        {
                            string key = cell.CellReference;
                            string value = cell != null ? GetCellValue(doc, cell) ?? "null" : "null";
                            ExcelObj[key] = value;
                        }

                        string a = "";
                        ExcelListObj.Add(ExcelObj);
                    }
                    DataTable dt = ConvertDictToDataTable(ExcelListObj, sheetName!);
                    ExcelDataSet!.Tables.Add(dt);
                }

            }
        }
        private DataTable ConvertDictToDataTable(List<Dictionary<string, string>> list_dict, string sheetName)
        {
            DataTable ExcelTable = new(sheetName);
            Dictionary<string, string> longest_dict = list_dict.OrderByDescending(item => item.Count)
                                                              .FirstOrDefault()!;
            List<string> longestKeys = longest_dict != null ? [.. longest_dict.Keys] : new List<string>();
            ExcelTable = SetTableColumn(ExcelTable, longestKeys);
            ExcelTable = SetTableData(ExcelTable, list_dict, longestKeys);
            return ExcelTable;
        }
        private DataTable SetTableColumn(DataTable table, List<string> keys)
        {
            foreach (string key in keys)
            {
                string filteredColumn = Regex.Replace(key, @"\d", "");
                table.Columns.Add(filteredColumn);
            }
            return table;
        }
        private DataTable SetTableData(DataTable table, List<Dictionary<string, string>> List_data, List<string> columns)
        {
            foreach (Dictionary<string, string> data in List_data)
            {
                DataRow row = table.NewRow();
                foreach (string key in columns)
                {
                    string column_name = Regex.Replace(key, @"\d", "");
                    string value = data.FirstOrDefault(item => Regex.Replace(item.Key, @"\d", "") == column_name).Value;
                    row[column_name] = value;
                }
                table.Rows.Add(row);
            }
            return table;
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            SharedStringTablePart stringTablePart = doc.WorkbookPart!.SharedStringTablePart!;
            string value = cell.CellValue!.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText?.ToString()!;
            }
            return value?.ToString() ?? "";
        }

    }
}
