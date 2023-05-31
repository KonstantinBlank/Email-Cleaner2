using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace Email_Cleaner
{
    internal class FolderInfoUI
    {
        private string _id;
        private List<LabelInfo> _labels;
        private Button _deleteButton;

        public string Id
        {
            get { return _id; }
        }

        public FolderInfoUI(string id, List<LabelInfo> labels, Button deletButton)
        {
            _id = id;
            _labels = labels;
            _deleteButton = deletButton;
        }

        public void UpdateLabels(Folder folder)
        {
            foreach (LabelInfo label in _labels.FindAll(label => label.Update != null))
            {
                label.Update(label.Label, folder);
            }
        }

        public void Delete(TableLayoutPanel layout)
        {
            TableLayoutControlCollection controls = layout.Controls;
            // get the number of rows that are to be delted
            // +1 because there is an empty row before
            int removedRowCount = _labels.Count + 1;

            // get the index of the first row after the rows that are to be deleted
            int rowIndex = _labels[_labels.Count - 1].RowIndex + 1;

            foreach (LabelInfo label in _labels)
            {
                controls.RemoveByKey(label.Label.Name);
            }
            _labels = null;

            controls.RemoveByKey(_deleteButton.Name);
            _deleteButton = null;

            _id = null;

            // remove the rows by shifting the rows below up and than remove the rows at the bottom
            // Shift controls from rows below the removed row
            for (int row = rowIndex; row < layout.RowCount; row++)
            {
                for (int column = 0; column < layout.ColumnCount; column++)
                {
                    Control control = layout.GetControlFromPosition(column, row);
                    if (control != null)
                    {
                        layout.SetRow(control, row - removedRowCount);
                    }
                }
            }

            // Remove the last row
            layout.RowStyles.RemoveAt(layout.RowCount - removedRowCount);
            layout.RowCount -= removedRowCount;
        }
    }
}
