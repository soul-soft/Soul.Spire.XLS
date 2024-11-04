using System.Collections.Generic;
using System.Linq;

namespace Soul.Spire.XLS
{
    public class XLSRow
    {
        public IReadOnlyCollection<XLSCell> Cells { get; }

        public XLSRow(XLSTable table)
        {
            var cells = new List<XLSCell>();
            foreach (var item in table.Columns)
            {
                cells.Add(new XLSCell(item));
            }
            Cells = cells;
        }

        public object this[string code]
        {
            set
            {
                var cell = Cells.Where(a => a.Column.Code == code).FirstOrDefault();
                cell.Value = value;
            }
            get
            {
                return Cells.Where(a => a.Column.Code == code).FirstOrDefault();
            }
        }
    }
}
