using Spire.Xls;
using Spire.Xls.Core;
using System;

namespace Soul.Spire.XLS
{
    public class XLSTableInfo
    {
        public IXLSRange Header { get; set; }

        public IXLSRange Body { get; set; }

        public XLSTableInfo(IXLSRange header, IXLSRange body)
        {
            Header = header;
            Body = body;
        }

        public void ApplyDefaultStyle()
        {
            Header.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            Header.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            Header.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            Header.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            Header.Style.Font.IsBold = true;
            Body.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            Body.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            Body.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            Body.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
        }
    }
}
