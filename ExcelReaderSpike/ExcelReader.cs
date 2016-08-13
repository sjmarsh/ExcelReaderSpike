using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderSpike
{
    public class ExcelReader
    {

        public Dictionary<string, PropertyInfo> GetColumnPropertyMap<T>()
        {
            var maps = new Dictionary<string, PropertyInfo>();

            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                var excelColumn = prop.GetCustomAttributes(typeof(ExcelColumn), true).FirstOrDefault() as ExcelColumn;
                if (excelColumn != null)
                {
                    maps.Add(excelColumn.ColumnIdentifier, prop);
                }
            }

            return maps;
        }

        public List<T> ReadIntoList<T>(ExcelWorksheet worksheet, bool hasHeaderRow) where T : new()
        {
            var rows = new List<T>();
            var startRow = hasHeaderRow ? 2 : 3;
            var endRow = worksheet.Dimension.End.Row;

            var maps = GetColumnPropertyMap<T>();

            for (int i = startRow; i < endRow; i++)
            {
                var objectInstance = new T();

                foreach (var map in maps)
                {
                    var cellValue = worksheet.Cells[map.Key + i].Value;
                    var typedValue = Convert.ChangeType(cellValue, map.Value.PropertyType); // this needs better error handling
                    map.Value.SetValue(objectInstance, typedValue);
                }

                rows.Add(objectInstance);
            }

            return rows;
        }
    }
}
