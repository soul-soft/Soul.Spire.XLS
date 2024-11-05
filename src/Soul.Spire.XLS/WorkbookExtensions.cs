using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Linq;

namespace Soul.Spire.XLS
{
    public static class WorkbookExtensions
    {
        public static XLSTableInfo InsertTable(this IWorksheet sheet, XLSTable table, int rowIndex, int columnIndex = 1)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));
            if (table == null)
                throw new ArgumentNullException(nameof(table));
            InsertHeaders(sheet, table.Root, rowIndex, columnIndex);
            InsertTableRows(sheet, table, rowIndex + table.Root.GetDepth() - 1, columnIndex);
            var body = sheet.Range[rowIndex, columnIndex, rowIndex + table.Rows.Count + table.Root.GetDepth() - 2, columnIndex +  table.Columns.Count - 1];
            var header = sheet.Range[rowIndex, columnIndex, rowIndex + table.Root.GetDepth() - 2, columnIndex + table.Columns.Count - 1];
            return new XLSTableInfo(header, body);
        }

        private static void InsertHeaders(IWorksheet sheet, XLSColumn header, int rowIndex, int columnIndex = 0)
        {
            if (header.Items == null || !header.Items.Any())
            {
                return;
            }
            foreach (var item in header.Items)
            {
                sheet.Range[rowIndex, columnIndex, rowIndex + item.RowSpan - 1, columnIndex + item.ColSpan - 1].Value = item.Name;
                sheet.Range[rowIndex, columnIndex, rowIndex + item.RowSpan - 1, columnIndex + item.ColSpan - 1].Merge();
                sheet.Range[rowIndex, columnIndex, rowIndex + item.RowSpan - 1, columnIndex + item.ColSpan - 1].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[rowIndex, columnIndex, rowIndex + item.RowSpan - 1, columnIndex + item.ColSpan - 1].HorizontalAlignment = HorizontalAlignType.Center;
                InsertHeaders(sheet, item, rowIndex + 1, columnIndex);
                columnIndex = columnIndex + item.ColSpan;
            }
        }

        private static void InsertTableRows(IWorksheet sheet, XLSTable table, int rowIndex, int columnIndex = 0)
        {
            if (table.Columns != null)
            {
                var startColumn = columnIndex;
                foreach (var item in table.Columns)
                {
                    if (item.Width > 0)
                    {
                        sheet.Columns[startColumn].ColumnWidth = item.Width;
                    }
                }
            }
            foreach (var row in table.Rows)
            {
                var startColumn = columnIndex;
                foreach (var cell in row.Cells)
                {
                    sheet.Range[rowIndex, startColumn].Value = cell.Value?.ToString();
                    sheet.Range[rowIndex, startColumn].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[rowIndex, startColumn].HorizontalAlignment = HorizontalAlignType.Center;
                    startColumn++;
                }
                rowIndex++;
            }
        }
    }
}
