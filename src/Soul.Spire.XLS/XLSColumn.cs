using System;
using System.Collections.Generic;
using System.Linq;

namespace Soul.Spire.XLS
{
    public class XLSColumn
    {
        public string Code { get;  }
        public string Name { get;  }
        public double Width { get; set; }
        internal int RowSpan { get; set; }
        internal int ColSpan { get; set; }

        public ICollection<XLSColumn> Items { get; set; }

        internal XLSColumn() 
        {

        }

        public XLSColumn(string name) : this(name, name)
        {

        }

        public XLSColumn(string code, string name)
        {
            Code = code;
            Name = name;
        }

        internal void CalculateRowSpans()
        {
            var maxDepth = GetDepth();
            CalcSpan(this, 0, maxDepth);
        }

        internal List<XLSColumn> GetLeafNodes()
        {
            var nodes = new List<XLSColumn>();
            GetLeafNodes(this, nodes);
            return nodes;
        }

        private void GetLeafNodes(XLSColumn node, List<XLSColumn> leaves)
        {
            if (node.Items == null || node.Items.Count == 0)
            {
                leaves.Add(node);
            }
            else
            {
                foreach (var item in node.Items)
                {
                    GetLeafNodes(item, leaves);
                }
            }
        }

        private void CalcSpan(XLSColumn node, int depth, int maxDepth)
        {
            if (node.Items == null || !node.Items.Any())
            {
                node.RowSpan = maxDepth - depth;
            }
            else
            {
                node.RowSpan = 1;
            }
            node.ColSpan = node.GetLeafNodes().Count;
            foreach (var item in node.Items ?? new List<XLSColumn>())
            {
                CalcSpan(item, depth + 1, maxDepth);
            }
        }

        internal int GetDepth()
        {
            if (Items == null || Items.Count == 0)
            {
                return 1;
            }
            int maxDepth = 0;
            foreach (var item in Items)
            {
                maxDepth = Math.Max(maxDepth, item.GetDepth());
            }
            return maxDepth + 1;
        }
    }
}
