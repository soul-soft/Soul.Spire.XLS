using System.Collections.Generic;

namespace Soul.Spire.XLS
{
    public class XLSTable
    {
        internal XLSColumn Root { get; }

        public IReadOnlyCollection<XLSColumn> Columns { get; private set; }

        public List<XLSRow> Rows { get; }

        public XLSTable(ICollection<XLSColumn> items)
        {
            Rows = new List<XLSRow>();
            Root = new XLSColumn()
            {
                Items = items,
            };
            Root.CalculateRowSpans();
            Columns = Root.GetLeafNodes();
        }

        public XLSRow NewRow()
        {
            return new XLSRow(this);
        }
    }
}
