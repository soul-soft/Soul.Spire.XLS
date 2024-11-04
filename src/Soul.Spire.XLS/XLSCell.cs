namespace Soul.Spire.XLS
{
    public class XLSCell
    {
        public object Value { get; set; }

        public XLSColumn Column { get; }

        public XLSCell(XLSColumn column)
        {
            Column = column;
        }

    }
}
