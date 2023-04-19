using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;


namespace Email_Cleaner
{
    public partial class UserControl_Outlook : UserControl
    {
        private List<Folder> _folders = null;
        private List<FolderInfoUI> _folderInfos = null;
        private Label _summaryLabel = null;
        private Folder _trash = null;
        private int _deletedEmails = -1;
        private int xDefault = 50;
        private int yDefault = 50;
        private int xCoordinate;
        private int yCoordinate;

        public Folder Trash
        {
            set
            {
                if (_trash == null)
                {
                    _trash = value;
                }
            }
        }

        public UserControl_Outlook()
        {
            InitializeComponent();
            _folderInfos = new List<FolderInfoUI>();
            xCoordinate = xDefault;
            yCoordinate = yDefault;
            //TODO: load deleted Emails here
            _deletedEmails = 0;
            _summaryLabel = addSummaryLabel();
            updateSummaryLabel();
            //this.Width = 1000;
        }

        public void SetFolders(List<Folder> folders)
        {
            _folders = folders;
            foreach (Folder folder in _folders)
            {
                FolderInfoUI folderInfo = GetFolderInfo(folder);
                if (folderInfo == null)
                {
                    addFolderInfoUI(folder);
                }
                else
                {
                    folderInfo.UpdateLabels(folder);
                }
            }
            Width = 1000;
        }

        private void addFolderInfoUI(Folder folder)
        {
            List<LabelInfo> labels = new List<LabelInfo>();
            labels.Add(addFolderNameLabel(folder));
            labels.Add(addFolderEmailCountLabel(folder));
            labels.Add(addComparisonLabel(folder));
            Button deleteButton = addDeleteButton(folder);
            FolderInfoUI folderInfo = new FolderInfoUI(folder.EntryID, labels, deleteButton);
            _folderInfos.Add(folderInfo);
            yCoordinate += 25;
        }

        private void updateSummaryLabel()
        {
            _summaryLabel.Text = $"Du hast bereits {_deletedEmails} Plastikstüten gespart.";
        }

        private Label addSummaryLabel()
        {
            string name = "summary";
            if (Controls.ContainsKey(name))
            {
                return (Label)Controls[name];
            }
            return addLabel(name, "", xCoordinate, yCoordinate);
        }

        /* lables */
        private LabelInfo addComparisonLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_comparison";
            string text = $"{folder.Items.Count} Plastiktüten";
            Label label = addLabel(name, text, xCoordinate, yCoordinate);
            return new LabelInfo(label, (label1, folder1) =>
            {
                label1.Text = $"{folder1.Items.Count} Plastiktüten";
            });
        }

        private LabelInfo addFolderEmailCountLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_mail_count";
            string text = folder.Items.Count.ToString() + " E-Mails";
            Label label = addLabel(name, text, xCoordinate, yCoordinate);
            return new LabelInfo(label, (label1, folder1) =>
            {
                label1.Text = $"{folder1.Items.Count} E-Mails";
            });
        }

        private LabelInfo addFolderNameLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_name";
            Label label = addLabel(name, folder.Name, xCoordinate, yCoordinate);
            return new LabelInfo(label);
        }

        private LabelInfo addFolderSizeLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_size_label";
            Label label = addLabel(name, "Größe", xCoordinate, yCoordinate, false);
            return new LabelInfo(label);
        }

        private LabelInfo addFolderSizeValueLabel(Folder folder)
        {
            string size = getFolderSize(folder);
            string name = folder.Name + "_" + folder.EntryID + "_size_value";
            Label label = addLabel(name, size, xCoordinate + 50, yCoordinate);
            return new LabelInfo(label);
        }

        private Label addLabel(string name, string text, int x, int y, bool newLine = true)
        {
            Label label = new Label();
            label.AutoSize = true;
            label.Name = name;
            label.Text = text;
            label.Location = new Point(x, y);
            if (newLine)
            {
                yCoordinate += 25;
            }
            this.Controls.Add(label);
            return label;
        }


        /* buttons */

        private Button addDeleteButton(Folder folder)
        {
            Button button = addButton(folder.Name + "_" + folder.EntryID + "_delete_button", "Löschen", xCoordinate + 150, yCoordinate - 50);
            button.Click += delegate (object sender, EventArgs e)
            {
                delete_mails(sender, e, folder.Items);
                updateSummaryLabel();
                deleteFolderInfoUI(folder);
                updateFolderInfoUIPosition();
            };
            return button;
        }

        private void updateFolderInfoUIPosition()
        {
            xCoordinate = xDefault;
            yCoordinate = yDefault;
            foreach (FolderInfoUI folderInfo in _folderInfos)
            {
                yCoordinate = folderInfo.UpdatePosition(xCoordinate, yCoordinate);
            }
        }

        private void deleteFolderInfoUI(Folder folder)
        {
            FolderInfoUI folderInfo = GetFolderInfo(folder);
            folderInfo.Delete(Controls);
            _folderInfos.Remove(folderInfo);
        }

        private Button addButton(string name, string text, int x, int y)
        {
            Button button = new Button();
            button.Location = new Point(x, y);
            button.Text = text;
            button.Name = name;
            button.AutoSize = true;
            //button.UseVisualStyleBackColor = true;

            this.Controls.Add(button);

            return button;
        }


        /* helper */

        private string getFolderSize(Folder folder)
        {
            double size = 0;
            for (int i = folder.Items.Count; i > 0; i--)
            {
                MailItem item = folder.Items[i];
                //converting bytes into megabytes
                size += item.Size / 1e6;
            }
            return Math.Round(size, 2).ToString() + " megabytes";
        }

        private void delete_mails(object sender, EventArgs e, Items items)
        {
            int itemCount = items.Count;
            List<string> safedEntryIds = new List<string>();

            // safe all trash items
            Items trashItems = _trash.Items;
            int trashItemCount = trashItems.Count;
            for (int i = trashItemCount; i > 0; i--)
            {
                MailItem trashItem = trashItems[i];
                safedEntryIds.Add(trashItem.EntryID);
            }

            // move into trahs
            for (int i = items.Count; i > 0; i--)
            {
                MailItem item = items[i];
                item.Delete();
            }

            // delete from all trash items that are not saved
            for (int i = trashItems.Count; i > 0; i--)
            {
                MailItem item = trashItems[i];
                if (!safedEntryIds.Contains(item.EntryID))
                {
                    item.Delete();
                }
            }

            _deletedEmails += itemCount;
        }

        private FolderInfoUI GetFolderInfo(Folder folder)
        {
            return _folderInfos.Find(f => f.Id.Equals(folder.EntryID));
        }

    }
}
