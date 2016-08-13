using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderSpike
{
    public class ExcelColumn : Attribute
    {
        private string _columnIdentifier;

        public ExcelColumn(string columnIdentifier)
        {
            _columnIdentifier = columnIdentifier;
        }

        public virtual string ColumnIdentifier { get { return _columnIdentifier; } }
    }
}
