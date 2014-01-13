using System;
using System.Collections;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ListViewSorter
{
    public class ListViewColumnSorter : IComparer
    {
        private int ColumnToSort;
        private SortOrder OrderOfSort;
        private CaseInsensitiveComparer ciCompare;

        public enum SortTypes { stText = 0, stNumeric, stHexNumber, stNone };
        public SortTypes[] columnSortType;
        private bool warned = false;

        public ListViewColumnSorter()
        {
            ColumnToSort = 0;
            OrderOfSort = SortOrder.None;
            ciCompare = new CaseInsensitiveComparer();
        }

        // "0" if equal, - if 'x' is < than 'y' and + if 'x' is > than 'y'
        public int Compare(object x, object y)
        {
            int compareResult = 0;
            ListViewItem listviewX = (ListViewItem)x;
            ListViewItem listviewY = (ListViewItem)y;
            string a, b;

            if (columnSortType == null)
            {
                if (!warned)
                {
                    MessageBox.Show("You have not set the ColumnSortTypes!");
                    warned = true;
                }
                return 0;
            }

            if (ColumnToSort > columnSortType.Length) return 0; //not specified so not handled...

            SortTypes s = columnSortType[ColumnToSort];
            if (s == SortTypes.stNone) return 0;

            if (ColumnToSort == 0)
            {
                a = listviewX.Text;
                b = listviewY.Text;
            }
            else
            {
                a = listviewX.SubItems[ColumnToSort].Text;
                b = listviewY.SubItems[ColumnToSort].Text;
            }

            switch (s)
            {
                case SortTypes.stNone: break;
                case SortTypes.stHexNumber:
                    if (IsHexNumber(a) && IsHexNumber(b))
                    {
                        int ia = Int32.Parse(a, System.Globalization.NumberStyles.HexNumber);
                        int ib = Int32.Parse(b, System.Globalization.NumberStyles.HexNumber);
                        compareResult = ciCompare.Compare(ia, ib);
                    }
                    break;
                case SortTypes.stNumeric:
                    if (IsWholeNumber(a) && IsWholeNumber(b))
                    {
                        compareResult = ciCompare.Compare(System.Convert.ToInt32(a), System.Convert.ToInt32(b));
                    }
                    break;
                case SortTypes.stText:
                    compareResult = ciCompare.Compare(a, b);
                    break;

            }

            if (OrderOfSort == SortOrder.Ascending) return compareResult;
            if (OrderOfSort == SortOrder.Descending) return (-compareResult);
            return 0;

        }

        public int SortColumn
        {
            set { ColumnToSort = value; }
            get { return ColumnToSort; }
        }

        public SortOrder Order
        {
            set { OrderOfSort = value; }
            get { return OrderOfSort; }
        }

        private bool IsWholeNumber(string strNumber)
        {
            Regex objNotWholePattern = new Regex("[^0-9]");
            return !objNotWholePattern.IsMatch(strNumber);
        }

        private bool IsHexNumber(string strNumber)
        {
            try
            {
                int x = Int32.Parse(strNumber, System.Globalization.NumberStyles.HexNumber);
                return true;
            }
            catch (Exception e) { return false; }
        }

    }
}
