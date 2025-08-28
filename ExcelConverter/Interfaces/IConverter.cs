using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Interfaces
{
    public interface IConverter
    {
        DataTable GetDataTable(string? tablename = null);
        DataSet GetDataSet();
    }
}
