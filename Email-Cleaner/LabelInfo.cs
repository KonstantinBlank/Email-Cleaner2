using Microsoft.Office.Interop.Outlook;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Email_Cleaner
{
    public class LabelInfo
    {
        public Label Label { get; private set; }
        public Action<Label, Folder> Update { get; private set; }
        public int RowIndex { get; private set; }

        public LabelInfo(Label label, int rowIndex,  Action<Label, Folder> update = null)
        {
            Label = label;
            RowIndex = rowIndex;
            Update = update;
        }
    }
}
