using System.Collections.Generic;
using System.Windows.Forms;
using static System.Windows.Forms.Control;
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

        public void Delete(ControlCollection controls)
        {
            foreach (LabelInfo label in _labels)
            {
                controls.RemoveByKey(label.Label.Name);
            }
            _labels = null;

            controls.RemoveByKey(_deleteButton.Name);
            _deleteButton = null;
            
            _id = null;
        }


        internal int UpdatePosition(int x, int y)
        {
            foreach (LabelInfo label in _labels)
            {
                label.SetLocation(x, y);
                y += 25;
            }            
            y += 25;
            return y;
        }
    }
}
