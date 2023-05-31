using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Exception = System.Exception;

namespace Email_Cleaner
{
    public partial class UserControl_Outlook : UserControl
    {
        private List<Folder> _folders = null;
        private List<FolderInfoUI> _folderInfos = null;
        private Label _summaryLabel = null;
        private Folder _trash = null;
        private int _deletedEmails = -1;
        private TableLayoutPanel tableLayoutPanel;

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
            InitializeTableLayoutPanel();
            _folderInfos = new List<FolderInfoUI>();
            //TODO: load deleted Emails here
            _deletedEmails = 0;
            _summaryLabel = addSummaryLabel();
            updateSummaryLabel();

        }

        private void InitializeTableLayoutPanel()
        {
            // Create the TableLayoutPanel
            tableLayoutPanel = new TableLayoutPanel();
            tableLayoutPanel.Dock = DockStyle.Fill;
            tableLayoutPanel.RowCount = 0;
            tableLayoutPanel.ColumnCount = 2;
            //set border style
            //tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            Controls.Add(tableLayoutPanel);
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
            tableLayoutPanel.RowStyles[tableLayoutPanel.RowCount - 1] = new RowStyle(SizeType.AutoSize, 20F);
        }

        private void addFolderInfoUI(Folder folder)
        {
            addEmtpyRow();

            List<LabelInfo> labels = new List<LabelInfo>
            {
                addFolderNameLabel(folder),
                addFolderEmailCountLabel(folder),
                addComparisonLabel(folder)
            };
            Button deleteButton = addDeleteButton(folder);
            FolderInfoUI folderInfo = new FolderInfoUI(folder.EntryID, labels, deleteButton);
            _folderInfos.Add(folderInfo);
        }

        private void addEmtpyRow()
        {
            tableLayoutPanel.RowCount++;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
        }

        private void updateSummaryLabel()
        {
            _summaryLabel.Text = $"Du hast bereits {_deletedEmails} Plastikstüten gespart.";
        }

        private Label addSummaryLabel()
        {
            addEmtpyRow();
            string name = "summary";
            if (Controls.ContainsKey(name))
            {
                return (Label)Controls[name];
            }

            Label label = addLabel(name, "", true);
            tableLayoutPanel.SetColumnSpan(label, 2);
            return label;
        }

        /* lables */
        private LabelInfo addComparisonLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_comparison";
            string text = $"{folder.Items.Count} Plastiktüten";
            Label label = addLabel(name, text);
            return new LabelInfo(label, tableLayoutPanel.RowCount - 1, (label1, folder1) =>
            {
                label1.Text = $"{folder1.Items.Count} Plastiktüten";
            });
        }

        private LabelInfo addFolderEmailCountLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_mail_count";
            string text = folder.Items.Count.ToString() + " E-Mails";
            Label label = addLabel(name, text);
            return new LabelInfo(label, tableLayoutPanel.RowCount - 1, (label1, folder1) =>
            {
                label1.Text = $"{folder1.Items.Count} E-Mails";
            });
        }

        private LabelInfo addFolderNameLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_name";
            Label label = addLabel(name, folder.Name, true);
            return new LabelInfo(label, tableLayoutPanel.RowCount - 1);
        }

        private LabelInfo addFolderSizeLabel(Folder folder)
        {
            string name = folder.Name + "_" + folder.EntryID + "_size_label";
            Label label = addLabel(name, "Größe", false);
            return new LabelInfo(label, tableLayoutPanel.RowCount - 1);
        }

        private LabelInfo addFolderSizeValueLabel(Folder folder)
        {
            string size = getFolderSize(folder);
            string name = folder.Name + "_" + folder.EntryID + "_size_value";
            Label label = addLabel(name, size);
            return new LabelInfo(label, tableLayoutPanel.RowCount - 1);
        }

        private Label addLabel(string name, string text, bool bold = false)
        {
            Label label = new Label();
            label.AutoSize = true;
            label.Name = name;
            label.Text = text;

            if (bold)
            {
                label.Font = new Font(base.Font.FontFamily, 10, FontStyle.Bold);
            }
            else
            {
                label.Font = new Font(base.Font.FontFamily, 10);
            }
            int rowCount = tableLayoutPanel.RowCount++;
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize, 20F));

            tableLayoutPanel.Controls.Add(label, 0, rowCount);

            return label;
        }


        /* buttons */

        private Button addDeleteButton(Folder folder)
        {
            Button button = addButton(folder.Name + "_" + folder.EntryID + "_delete_button", "Löschen");

            button.Click += delegate (object sender, EventArgs e)
            {
                delete_mails(folder.Items);
                updateSummaryLabel();
                deleteFolderInfoUI(folder);
            };
            button.Anchor = AnchorStyles.None | AnchorStyles.Top;
            return button;
        }


        private void deleteFolderInfoUI(Folder folder)
        {
            FolderInfoUI folderInfo = GetFolderInfo(folder);
            folderInfo.Delete(tableLayoutPanel);
            _folderInfos.Remove(folderInfo);
        }

        private Button addButton(string name, string text)
        {
            Button button = new Button();
            button.Text = text;
            button.Name = name;
            button.AutoSize = true;
            tableLayoutPanel.Controls.Add(button, 1, tableLayoutPanel.RowCount - 1);
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

        private void delete_mails(Items itemsToDelete)
        {
            int itemsToDeleteCount = itemsToDelete.Count;

            // safe all trash items
            List<string> safedEntryIds = new List<string>();
            Items trashItems = _trash.Items;
            foreach (MailItem trashItem in trashItems)
            {
                safedEntryIds.Add(trashItem.EntryID);
            }

            // move into trash
            // The index for the Items collection starts at 1, and the items in the Items collection object are not guaranteed to be in any particular order.
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mapifolder.items?view=outlook-pia#microsoft-office-interop-outlook-mapifolder-items
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem.delete?view=outlook-pia#microsoft-office-interop-outlook-mailitem-delete
            int count = itemsToDelete.Count;
            for (int i = count; i > 0; i--)
            {
                MailItem item = itemsToDelete[i];
                item.Delete();
            }
                        
            count = itemsToDelete.Count;
            if (count > 0)
            {
                //throw new Exception("Somehow not all emails got deleted from the folder.");
            }

            // delete from all trash items that are not saved
            // The index for the Items collection starts at 1, and the items in the Items collection object are not guaranteed to be in any particular order.
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mapifolder.items?view=outlook-pia#microsoft-office-interop-outlook-mapifolder-items
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem.delete?view=outlook-pia#microsoft-office-interop-outlook-mailitem-delete
            trashItems = _trash.Items;
            for (int i = _trash.Items.Count; i > 0; i--)
            {
                MailItem item = trashItems[i];
                if (!safedEntryIds.Contains(item.EntryID))
                {
                    item.Delete();
                }
            }

            trashItems = _trash.Items;
            if (trashItems.Count > safedEntryIds.Count)
            {
                //throw new Exception("Somehow not all emails got deleted from the trash.");
            }

            _deletedEmails += itemsToDeleteCount;
        }

        private FolderInfoUI GetFolderInfo(Folder folder)
        {
            return _folderInfos.Find(f => f.Id.Equals(folder.EntryID));
        }

    }
}
