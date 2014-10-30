using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertCsvToExcel.ExtensionMethods
{
    using NPOI.SS.UserModel;

    public static class NPOIExtensions
    {
        public static void SetCellValue(this ICell cell, object value)
        {
            var type = value.GetType();
            if (value.GetType() == typeof(bool))
            {
                cell.SetCellValue((bool)value);
                return;
            }
            if (value.GetType() == typeof(DateTime))
            {
                cell.SetCellValue((DateTime)value);
                return;
            }
            if (value.GetType() == typeof(double) || value.GetType() == typeof(float) || value.GetType() == typeof(long) || value.GetType() == typeof(int))
            {
                cell.SetCellValue((double)value);
                return;
            }

            cell.SetCellValue(value.ToString());

        }
    }
}
