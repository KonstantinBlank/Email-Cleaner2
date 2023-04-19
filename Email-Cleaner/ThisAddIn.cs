using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Outlook;

namespace Email_Cleaner
{
    public partial class ThisAddIn
    {
        private UserControl_Outlook _userControl_Outlook;
        private CustomTaskPane _taskPane;
        public CustomTaskPane TaskPane
        {
            get
            {
                return _taskPane;
            }
        }

        /// <summary>
        /// Connect with Outlook and add the TaskPane 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _userControl_Outlook = new UserControl_Outlook();
            _taskPane = this.CustomTaskPanes.Add(_userControl_Outlook, "Email Cleaner");
            _taskPane.VisibleChanged += new EventHandler(taskPane_VisibleChanged);
        }

        private void taskPane_VisibleChanged(object sender, System.EventArgs e)
        {
            Application app = new Application();
            NameSpace mapiNameSpace = app.GetNamespace("MAPI");
            List<Folder> folders = getFoldersWithEmails(mapiNameSpace);
            _userControl_Outlook.SetFolders(folders);
            _userControl_Outlook.Trash = getTrashFolder(mapiNameSpace);
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = _taskPane.Visible;
        }

        private List<Folder> getFoldersWithEmails(NameSpace mapiNameSpace)
        {
            Folders folders = mapiNameSpace.Folders;
            Folder accountFolder = (Folder)folders.GetFirst();
            List<Folder> foldersWithEmails = new List<Folder>();
            foreach (Folder folder in accountFolder.Folders)
            {
                string name = folder.Name;
                string path = folder.FullFolderPath;
                int itemCount = folder.Items.Count;
                if (itemCount > 0 && !name.Contains("This computer only"))
                {
                    foldersWithEmails.Add(folder);
                }
            }
            return foldersWithEmails;
        }

        private Folder getTrashFolder(NameSpace nameSpace)
        {
            Folder trash = (Folder) nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            return trash;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
