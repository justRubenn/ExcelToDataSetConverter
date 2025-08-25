using DocumentFormat.OpenXml.ExtendedProperties;
using ExcelConverter.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Logics
{
    public class CSVConverter : IConverter
    {
        private DataSet CSVDataset { get; set; } = new DataSet();

        public CSVConverter(string[] lines, string delimiter)
        {
            ToDataTable(lines, delimiter);
        }
        public DataSet GetDataSet() => CSVDataset;
        public DataTable GetDataTable(string tableName) => CSVDataset?.Tables[tableName] ?? new DataTable();
        private void ToDataTable(string[] lines, string delimiter)
        { 
            int columnCount = lines.Max(item => item.Split(delimiter).Length);
            DataTable dt = CreateDataTable(columnCount); 
            foreach(var item in lines)
            {
                dt = InsertRows(dt,item,delimiter);
            }
            CSVDataset.Tables.Add(dt);
        }
        private DataTable CreateDataTable(int columnCount)
        {
            DataTable dt = new DataTable();
            for (int i = 1; i <= columnCount; i++)
            {
                dt.Columns.Add(i.ToString(), typeof(string));
            }
            return dt;
        }
        private DataTable InsertRows(DataTable dt, string rows, string delimiter)
        {
            string[] row = rows.Split(delimiter);
            dt.Rows.Add(row);
            return dt;
        }
    }
}
