using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Controls;
using jd.Helper.Configuration;

namespace WDE
{
    public delegate void MyDelegate(object sender, EventArgs e);

    public partial class UserControlExplorer : UserControl
    {
        public event MyDelegate MyEvent;

        private ApplicationSettings applicationSettings = null;
        private string sectionName = "";
        private Section optionSection = null;
        private string LastPath = "";
        private bool ShowDetails = false;
        private bool ShowPreview = false;
        private bool ShowNavigation = true;
        private bool showSelected = false;
        private int savedTabIndex = -1;
        public bool TabControlLocked = false;
        private TabControl currentTabControl;
        private string _lockedPath = "";

        private ToolStripButton contextTSB = null;

        public TabControl CurrentTabControl
        {
            get { return currentTabControl; }
        }

        public string LockedPath
        {
            get { return _lockedPath; }
        }

        public bool ShowSelected
        {
            set
            {
                optionSection.settings.GetItemByString("ShowSelected").Value = value.ToString();
                showSelected = value;
            }
            get
            {
                return showSelected;
            }
        }

        public int SavedTabIndex
        {
            set
            {
                optionSection.settings.GetItemByString("savedTabIndex").Value = value.ToString();
                savedTabIndex = value;
            }
            get
            {
                return savedTabIndex;
            }
        }

        public UserControlExplorer(string newSectionName, ApplicationSettings applicationSettingsValue, string name, TabControl destTabControl, TabPage tabPage)
        {
            InitializeComponent();

            explorerBrowser.NavigationLog.NavigationLogChanged += NavigationLog_NavigationLogChanged;

            currentTabControl = destTabControl;

            if (name == "")
            {
                sectionName = newSectionName; //"UserExplorerControl" + id.ToString();
            }
            else
            {
                sectionName = name;
            }

            applicationSettings = applicationSettingsValue;
            tabPage.Tag = sectionName;

            if (applicationSettings.sections.GetItemIDByString(sectionName, "FAV") == -1)
            {
                applicationSettings.sections.Add(sectionName, "FAV");
            }

            if (applicationSettings.sections.GetItemIDByString(sectionName, "OPTIONS") == -1)
            {
                applicationSettings.sections.Add(sectionName, "OPTIONS");
            }

            optionSection = applicationSettings.sections.GetItemByString(sectionName, "OPTIONS");

            if (optionSection.settings.GetItemByString("LastPath") == null)
            {
                optionSection.settings.Add("LastPath", "");
            }
            else
            {
                LastPath = optionSection.settings.GetItemByString("LastPath").Value;
            }

            tabPage.Text = LastPath;

            //ShowDetails
            if (optionSection.settings.GetItemByString("ShowDetails") == null)
            {
                optionSection.settings.Add("ShowDetails", ShowDetails.ToString());
            }
            else
            {
                ShowDetails = Convert.ToBoolean(optionSection.settings.GetItemByString("ShowDetails").Value);
            }

            tsmShowDetails.Checked = ShowDetails;
            if (ShowDetails)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Details = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Details = PaneVisibilityState.Hide;
            }

            //savedTabIndex
            if (optionSection.settings.GetItemByString("savedTabIndex") == null)
            {
                optionSection.settings.Add("savedTabIndex", (destTabControl.TabPages.Count - 1).ToString());
            }
            else
            {
                SavedTabIndex = Convert.ToInt32(optionSection.settings.GetItemByString("savedTabIndex").Value); //destTabControl.TabPages.Count-1; //not delete
            }

            //ShowSelected
            if (optionSection.settings.GetItemByString("ShowSelected") == null)
            {
                optionSection.settings.Add("ShowSelected", ShowSelected.ToString());
            }
            else
            {
                ShowSelected = Convert.ToBoolean(optionSection.settings.GetItemByString("ShowSelected").Value);
                if (ShowSelected)
                {
                    destTabControl.SelectedTab = tabPage;
                }
            }

            //ShowPreview
            if (optionSection.settings.GetItemByString("ShowPreview") == null)
            {
                optionSection.settings.Add("ShowPreview", ShowPreview.ToString());
            }
            else
            {
                ShowPreview = Convert.ToBoolean(optionSection.settings.GetItemByString("ShowPreview").Value);
            }

            tsmShowPreview.Checked = ShowPreview;
            if (ShowPreview)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Preview = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Preview = PaneVisibilityState.Hide;
            }

            //ShowNavigation
            if (optionSection.settings.GetItemByString("ShowNavigation") == null)
            {
                optionSection.settings.Add("ShowNavigation", ShowNavigation.ToString());
            }
            else
            {
                ShowNavigation = Convert.ToBoolean(optionSection.settings.GetItemByString("ShowNavigation").Value);
            }

            tsmShowNavigation.Checked = ShowNavigation;
            if (ShowNavigation)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Navigation = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Navigation = PaneVisibilityState.Hide;
            }

            //ParentTabControl
            if (optionSection.settings.GetItemByString("ParentTabControl") == null)
            {
                optionSection.settings.Add("ParentTabControl", destTabControl.Name);
            }

            //Locked
            if (optionSection.settings.GetItemByString("Locked") == null)
            {
                optionSection.settings.Add("Locked", "False");
            }
            else
            {
                TabControlLocked = Convert.ToBoolean(optionSection.settings.GetItemByString("Locked").Value);
            }

            if (TabControlLocked)
            {
                tabPage.ImageIndex = 0;
                _lockedPath = LastPath;
            }

            int sectionID = applicationSettings.sections.GetItemIDByString(sectionName, "FAV");

            for (int i = 0; i < applicationSettings.sections[sectionID].settings.Count; i++)
            {
                string fav = applicationSettings.sections[sectionID].settings[i].Name;
                string favValue = applicationSettings.sections[sectionID].settings[i].Value;
                CreateFavButton(Path.GetFileName(fav), favValue);
            }
        }

        void NavigationLog_NavigationLogChanged(object sender, NavigationLogEventArgs args)
        {
            //This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate()
            {
                // calculate button states
                if (args.CanNavigateBackwardChanged)
                {
                    tsbBack.Enabled = explorerBrowser.NavigationLog.CanNavigateBackward;
                }
                if (args.CanNavigateForwardChanged)
                {
                    tsbForward.Enabled = explorerBrowser.NavigationLog.CanNavigateForward;
                }
            }));
        }


        public delegate void PathChanged(string pathname, string name, UserControlExplorer uce);
        public event PathChanged pathChanged;

        private void UserControlExplorer_Load(object sender, EventArgs e)
        {
            try
            {
                if (LastPath == "")
                {
                    explorerBrowser.Navigate((ShellObject)KnownFolders.Desktop);
                }
                else
                {
                    explorerBrowser.Navigate(ShellFileSystemFolder.FromFolderPath(LastPath));
                }

                RefreshApp();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private string GetRootDriveFromPath(string path)
        {
            string driveLetter = path.Substring(0, 2);
            if (driveLetter != "::" && driveLetter != "\\\\")
            {
                DriveInfo driveInfo = new DriveInfo(driveLetter);

                return driveInfo.RootDirectory.FullName;
            }
            return "";
        }


        private void tsbRoot1_Click(object sender, EventArgs e)
        {
            string driveletter = GetRootDriveFromPath(tsddbtn.Text);
            if (driveletter != "")
            {
                try
                {
                    explorerBrowser.Navigate(ShellObject.FromParsingName(driveletter));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }
        }

        private void tsbUp_Click(object sender, EventArgs e)
        {
            try
            {
                if (explorerBrowser.NavigationLog.CurrentLocation.Parent != null)
                {
                    explorerBrowser.Navigate(explorerBrowser.NavigationLog.CurrentLocation.Parent);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tsbBack_Click(object sender, EventArgs e)
        {
            // Move backwards through navigation log
            try
            {
                explorerBrowser.NavigateLogLocation(NavigationLogDirection.Backward);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tsbForward_Click(object sender, EventArgs e)
        {
            // Move forwards through navigation log
            try
            {
                explorerBrowser.NavigateLogLocation(NavigationLogDirection.Forward);

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void explorerBrowser_Enter(object sender, EventArgs e)
        {
            OnMyEvent();
        }


        private void explorerBrowser_NavigationComplete(object sender, NavigationCompleteEventArgs e)
        {
            BeginInvoke(new MethodInvoker(delegate()
            {
                string location = (e.NewLocation == null) ? "(unknown)" : e.NewLocation.ParsingName;
                string locationName = (e.NewLocation == null) ? "(unknown)" : e.NewLocation.Name;

                if (!location.Contains("::") && (!TabControlLocked))
                {
                    optionSection.settings.GetItemByString("LastPath").Value = location;
                }

                tsddbtn.Text = GetRootDriveFromPath(location);

                tssl.Text = GetFreeDiskSpace(location);

                if (location.Contains("::"))
                {
                    navigationHistoryCombo.Text = locationName;

                    if (!navigationHistoryCombo.Items.Contains(locationName))
                    {
                        navigationHistoryCombo.Items.Add(locationName);
                    }

                    SizePathTextbox();

                    if (pathChanged != null)
                    {
                        pathChanged(locationName, locationName, this);
                    }
                }
                else
                {
                    navigationHistoryCombo.Text = location;

                    if (!navigationHistoryCombo.Items.Contains(location))
                    {
                        navigationHistoryCombo.Items.Add(location);
                    }

                    SizePathTextbox();

                    if (pathChanged != null)
                    {
                        pathChanged(location, locationName, this);
                    }
                }
            }));

            OnMyEvent();
        }

        private void explorerBrowser_SelectionChanged(object sender, EventArgs e)
        {
            OnMyEvent();

            tsbViewer.Enabled = explorerBrowser.SelectedItems[0] != null;
            tsbDel.Enabled = tsbViewer.Enabled;
        }

        private void tsbSwitch_Click(object sender, EventArgs e)
        {
            try
            {
                if (FormMain.current_uce1.explorerBrowser.NavigationLog.CurrentLocation ==
                    explorerBrowser.NavigationLog.CurrentLocation)
                {
                    explorerBrowser.Navigate(FormMain.current_uce2.explorerBrowser.NavigationLog.CurrentLocation);
                }
                else
                {
                    explorerBrowser.Navigate(FormMain.current_uce1.explorerBrowser.NavigationLog.CurrentLocation);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tsbDel_Click(object sender, EventArgs e)
        {
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            if (MessageBox.Show("Delete all selected files?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                BeginInvoke(new MethodInvoker(delegate()
                {
                    var itemsText = new StringBuilder();

                    foreach (ShellObject item in explorerBrowser.SelectedItems)
                    {
                        if (item != null)
                        {
                            File.Delete(item.ParsingName);
                        }
                    }
                }));
            }
        }

        private void tsbViewer_Click(object sender, EventArgs e)
        {
            FormViewer formViewer = new FormViewer();

            ShellObject so = explorerBrowser.SelectedItems[0];
            if (so != null)
            {
                formViewer.Show();
                formViewer.SetText(so.ParsingName);
            }
        }

        private void tsmiCopyFullnameToClipboard_Click(object sender, EventArgs e)
        {
            string body = "";
            int i = 0;

            foreach (ShellObject so2 in explorerBrowser.SelectedItems)
            {
                body += so2.ParsingName;

                i++;

                if (i < explorerBrowser.SelectedItems.Count)
                {
                    body += "\n";
                }
            }

            Clipboard.SetText(body);
        }

        private void tsmiCopyShortnameToClipboard_Click(object sender, EventArgs e)
        {
            string body = "";
            int i = 0;

            foreach (ShellObject so2 in explorerBrowser.SelectedItems)
            {
                body += so2.Name;

                i++;

                if (i < explorerBrowser.SelectedItems.Count)
                {
                    body += "\n";
                }
            }

            Clipboard.SetText(body);
        }

        private void tsmShowDetails_Click(object sender, EventArgs e)
        {
            ShowDetails = tsmShowDetails.Checked;

            if (tsmShowDetails.Checked)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Details = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Details = PaneVisibilityState.Hide;
            }

            explorerBrowser.Update();

            optionSection.settings.GetItemByString("ShowDetails").Value = ShowDetails.ToString();
        }

        private void tsmShowPreview_Click(object sender, EventArgs e)
        {
            ShowPreview = tsmShowPreview.Checked;

            if (ShowPreview)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Preview = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Preview = PaneVisibilityState.Hide;
            }

            optionSection.settings.GetItemByString("ShowPreview").Value = ShowPreview.ToString();
        }

        private void tsmShowNavigation_Click(object sender, EventArgs e)
        {
            ShowNavigation = tsmShowNavigation.Checked;

            if (ShowNavigation)
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Navigation = PaneVisibilityState.Show;
            }
            else
            {
                explorerBrowser.NavigationOptions.PaneVisibility.Navigation = PaneVisibilityState.Hide;
            }

            optionSection.settings.GetItemByString("ShowNavigation").Value = ShowNavigation.ToString();
        }

        private void CreateFavButton(string FavText, string FavPath)
        {
            if (FavText == "")
            {
                FavText = FavPath;
            }

            ToolStripButton mybtn = new ToolStripButton(FavText);
            mybtn.Tag = 1;
            mybtn.Image = Properties.Resources.folder_star;
            mybtn.ToolTipText = FavPath;
            mybtn.MouseDown += new MouseEventHandler(toolStripButton1_MouseDown);

            mybtn.Click += new EventHandler(mysbtn_Click);
            ts.Items.Add(mybtn);
        }

        private void mysbtn_Click(object sender, EventArgs e)
        {
            string pfad = (sender as ToolStripButton).ToolTipText;
            if (pfad != "")
            {
                if (Directory.Exists(pfad))
                {
                    try
                    {
                        explorerBrowser.Navigate(ShellObject.FromParsingName(pfad));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                }
                else
                {
                    MessageBox.Show(string.Format("Path [{0}] not exist! Maybe an old temporary drive like an usb stick or network drive.", pfad), "Path Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void toolStripButton1_MouseDown(object sender, MouseEventArgs e)
        {
            contextTSB = (sender as ToolStripButton);
        }

        private void mybtn_Click(object sender, EventArgs e)
        {
            string pfad = toolTip.GetToolTip(sender as Button);
            if (pfad != "")
            {
                if (Directory.Exists(pfad))
                {
                    try
                    {
                        explorerBrowser.Navigate(ShellObject.FromParsingName(pfad));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"Path [{pfad}] not exist! Maybe an old temporary drive like an usb stick or network drive.", "Path Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        // Spend size and free space
        private string GetFreeDiskSpace(string path)
        {
            string driveletter = path.Substring(0, 2);

            if (driveletter != "::" && driveletter != "\\\\")
            {
                var driveInfo = new DriveInfo(driveletter);
                string freeSpace = GetBestDiskSpaceSize(driveInfo.TotalFreeSpace);
                string totalSpace = GetBestDiskSpaceSize(driveInfo.TotalSize);

                return "   [" + driveInfo.VolumeLabel + " " + freeSpace + " / " + totalSpace + "]";
            }
            return "";
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (contextTSB == null)
            {
                return;
            }

            string mypath = contextTSB.ToolTipText;

            string mycaption = contextTSB.Text;


            int i = applicationSettings.sections.GetItemIDByString(sectionName, "FAV");
            if (i > -1)
            {
                int settingID = applicationSettings.sections[i].settings.GetItemIDByString(mycaption);
                Setting setting = applicationSettings.sections[i].settings.GetItemByString(mycaption);
                applicationSettings.sections[i].settings.Remove(setting);
            }

            contextTSB.Dispose();
        }

        private string GetBestDiskSpaceSize(long disksize)
        {
            string ResultString;
            Int64 diskSpaceKB = disksize / 1024;
            Int64 diskSpaceMB = diskSpaceKB / 1024;
            Int64 diskSpaceGB = diskSpaceMB / 1024;

            ResultString = diskSpaceGB + "GB";

            if (diskSpaceGB < 1)
            {
                ResultString = diskSpaceMB + "MB";
            }

            if (diskSpaceMB < 1)
            {
                ResultString = disksize + "Byte";
            }

            return ResultString;
        }

        private void ReadAllDrives(ToolStripDropDownButton btn)
        {
            btn.DropDownItems.Clear();

            foreach (var driveInfo in DriveInfo.GetDrives())
            {
                var newBtn = new ToolStripMenuItem();
                if (driveInfo.IsReady)
                {
                    newBtn.Text = driveInfo.Name + " [" + driveInfo.VolumeLabel + "]";
                    newBtn.ToolTipText = driveInfo.Name;
                    newBtn.Image = Properties.Resources.Drive;
                    newBtn.Tag = btn.Name == "tsddbtn1" ? 1 : 2;

                    newBtn.Click += newBtn_Click;

                    btn.DropDownItems.Add(newBtn);
                }
            }
        }

        private void newBtn_Click(object sender, EventArgs e)
        {
            // Move backwards through navigation log
            string dir = (sender as ToolStripMenuItem).ToolTipText;

            try
            {
                explorerBrowser.Navigate(ShellObject.FromParsingName(dir));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        public void RefreshApp()
        {
            ReadAllDrives(tsddbtn);
        }

        private void toolTip_Popup(object sender, PopupEventArgs e)
        {

        }

        private void explorerBrowser_ItemsChanged(object sender, EventArgs e)
        {
            //System.Media.SystemSounds.Beep.Play();
            // This event is BeginInvoked to decouple the ExplorerBrowser UI from this UI
            BeginInvoke(new MethodInvoker(delegate()
            {
                // update items text box
                StringBuilder itemsText = new StringBuilder();

                //foreach (ShellObject item in explorerBrowser1.Items)
                //{
                //    if (item != null)
                //        itemsText.AppendLine("\tItem = " + item.GetDisplayName(DisplayNameType.Default));
                //}

                //this.itemsTextBox.Text = itemsText.ToString();
                //this.itemsTabControl.TabPages[0].Text = "Items (Count=" + explorerBrowser1.Items.Count.ToString() + ")";

            }));
            OnMyEvent();
        }

        private void OnMyEvent()
        {
            if (MyEvent != null)
            {
                MyEvent(this, null);
            }
        }

        private void toolStrip2_SizeChanged(object sender, EventArgs e)
        {
            SizePathTextbox();
        }

        private void SizePathTextbox()
        {
            navigationHistoryCombo.Width = toolStrip2.Width - tsddbtn.Width - tssl.Width - 10;
        }

        private void toolStripComboBoxPath_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangePath(navigationHistoryCombo.Text);
        }

        private bool ChangePath(string newPath)
        {
            if (Directory.Exists(newPath))
            {
                try
                {
                    explorerBrowser.Navigate(ShellObject.FromParsingName(newPath));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
                return true;
            }

            return false;
        }

        private void toolStripComboBoxPath_KeyDown(object sender, KeyEventArgs e)
        {
            string oldPath = explorerBrowser.NavigationLog.CurrentLocation.ParsingName;
            string oldText = navigationHistoryCombo.Text;

            if (e.KeyCode == Keys.Return)
            {
                ChangePath(navigationHistoryCombo.Text);

                if (navigationHistoryCombo.Text == oldText)
                {
                    navigationHistoryCombo.Text = oldPath;
                }
            }

        }

        private void navigationHistoryCombo_Enter(object sender, EventArgs e)
        {
            navigationHistoryCombo.BackColor = SystemColors.Window;
        }

        private void navigationHistoryCombo_Leave(object sender, EventArgs e)
        {
            navigationHistoryCombo.BackColor = SystemColors.Control;
        }

        private void tsddbtn_DropDownOpening(object sender, EventArgs e)
        {
            RefreshApp();
        }

        private void ts_DragEnter(object sender, DragEventArgs e)
        {
            var s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            e.Effect = DragDropEffects.Link;
        }

        private void ts_DragDrop(object sender, DragEventArgs e)
        {
            String[] Params = (String[])e.Data.GetData(DataFormats.FileDrop);
            string myParam = Params[0];
            string myDir;
            string myDirCaption;

            if (File.Exists(myParam))
            {
                myDir = Path.GetDirectoryName(myParam);
            }
            else
            {
                myDir = Path.GetFullPath(myParam);
            }

            myDirCaption = Path.GetFileName(myDir);

            if (myDirCaption == "")
            {
                myDirCaption = myDir;
            }

            int i = applicationSettings.sections.GetItemIDByString(sectionName, "FAV");
            if (i >= 0)
            {
                applicationSettings.sections[i].settings.Add(myDirCaption, myDir);
            }

            CreateFavButton(myDirCaption, myDir);
        }

        private void ts_MouseDown(object sender, MouseEventArgs e)
        {
            contextTSB = null;
        }

        private void contextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            clearToolStripMenuItem.Enabled = false;

            if (contextTSB == null)
            {
                return;
            }

            clearToolStripMenuItem.Enabled = (contextTSB.Tag.ToString() == "1");
        }

        private void tsb_MouseDown(object sender, MouseEventArgs e)
        {
            contextTSB = null;
        }

        public bool ResetToLockedPath()
        {
            if (Directory.Exists(_lockedPath))
            {
                try
                {
                    explorerBrowser.Navigate(ShellObject.FromParsingName(_lockedPath));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        public delegate void MyPreviewKeyDown(object sender, PreviewKeyDownEventArgs e, TabControl dTabControl);
        public event MyPreviewKeyDown myPreviewKeyDown;

        private void explorerBrowser_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            myPreviewKeyDown?.Invoke(sender, e, currentTabControl);
        }

        private void sendFullnameViaEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShellObject so = explorerBrowser.SelectedItems[0];
            if (so != null)
            {
                string subject = so.Name;
                string body = "";// so.ParsingName;
                int i = 0;

                foreach (ShellObject so2 in explorerBrowser.SelectedItems)
                {
                    body += so2.ParsingName;

                    i++;

                    if (i < explorerBrowser.SelectedItems.Count)
                    {
                        body += "%0D%0A";
                    }
                }

                string s = "mailto:?body=" + body + "&subject=" + subject;

                Process.Start(s);
            }
        }

        private void sendShortnameViaEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShellObject so = explorerBrowser.SelectedItems[0];
            if (so != null)
            {
                string subject = so.Name;
                string body = "";
                int i = 0;

                foreach (ShellObject so2 in explorerBrowser.SelectedItems)
                {
                    body += so2.Name;

                    i++;

                    if (i < explorerBrowser.SelectedItems.Count)
                    {
                        body += "%0D%0A";
                    }
                }

                string s = "mailto:?body=" + body + "&subject=" + subject;

                Process.Start(s);
            }
        }
    }
}
