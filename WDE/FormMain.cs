using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Controls;
using System.Diagnostics;
using System.IO;
using jd.Helper.Configuration;
using System.Net;
using Microsoft.Win32;


namespace WDE
{
    public partial class FormMain : Form
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;


        const int TOPOFEACHOTHER = 1;
        const int NEXTTOEACHOTHER = 2;
        static public UserControlExplorer current_uce = null;
        static public UserControlExplorer current_uce1 = null;
        static public UserControlExplorer current_uce2 = null;

        private string[] _args;
        private ToolStripButton contextTSB = null;

        FormSplash formSplash = new FormSplash();

        ApplicationSettings applicationSettings = new ApplicationSettings(Path.Combine(Application.LocalUserAppDataPath, "wdeConfig.xml"));

        public FormMain(string[] args)
        {
            Stopwatch perfStopwatch = Stopwatch.StartNew();
            InitializeComponent();

            Debug.WriteLine("*** 1. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            formSplash.Show();

            Application.DoEvents();
            Debug.WriteLine("*** 2. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            _args = args;

            applicationSettings.Load();
            Debug.WriteLine("*** 3. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            for (int i = 0; i < applicationSettings.sections.Count; i++)
            {
                if (applicationSettings.sections[i].ClassName == "OPTIONS")
                {
                    string sectionName = applicationSettings.sections[i].Name;
                    TabControl tc = null;

                    string tabControlName = "";
                    if (applicationSettings.sections[i].settings.GetItemByString("ParentTabControl") != null)
                    {
                        tabControlName = applicationSettings.sections[i].settings.GetItemByString("ParentTabControl").Value;
                    }

                    if ((tabControlName == "") || (tabControlName == tabControl1.Name))
                    {
                        tc = tabControl1;
                    }
                    else
                    {
                        tc = tabControl2;
                    }

                    AddExplorer(tc, sectionName, false);
                }
            }
            Debug.WriteLine("*** 4. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            sortPages(tabControl1);
            sortPages(tabControl2);

            Debug.WriteLine("*** 5. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            if (tabControl1.TabCount == 0)
            {
                AddExplorer(tabControl1, "", false);
            }

            if (tabControl2.TabCount == 0)
            {
                AddExplorer(tabControl2, "", false);
            }

            //
            if (Properties.Settings.Default.CallUpgrade)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.CallUpgrade = false;
                Properties.Settings.Default.Save();
            }

            //caption
            Text += System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            Properties.Settings settings = Properties.Settings.Default;
            Size = settings.formsize;

            if (Properties.Settings.Default.Layout == TOPOFEACHOTHER)
            {
                panel1.Height = Properties.Settings.Default.Panel1H;
            }
            else
            {
                toolStripButton2_Click(null, null);
            }

            Left = Properties.Settings.Default.FormL;
            Top = Properties.Settings.Default.FormT;

            if (Left < 0)
            {
                Left = 0;
            }

            if (Top < 0)
            {
                Top = 0;
            }


            if (Height < 50)
            {
                Height = 700;
            }

            if (Width < 50)
            {
                Width = 800;
            }

            WindowState = Properties.Settings.Default.WindowState;

            Debug.WriteLine("*** 6. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            // initialize known folder combo box
            List<string> knownFolderList = new List<string>();
            foreach (IKnownFolder folder in KnownFolders.All)
            {
                knownFolderList.Add(folder.CanonicalName);
            }
            knownFolderList.Sort();
            Debug.WriteLine("*** 7. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            tscbKnownFolder.Items.AddRange(knownFolderList.ToArray());

            foreach (string s in knownFolderList)
            {
                ToolStripMenuItem tsmi = new ToolStripMenuItem(s);
                jumpToToolStripMenuItem.DropDownItems.Add(s, Properties.Resources.folder, jumpTo_Click);
            }

            Debug.WriteLine("*** 8. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            if (Properties.Settings.Default.Favs != null)
            {
                foreach (string fav in Properties.Settings.Default.Favs)
                {
                    CreateFavButton(Path.GetFileName(fav), fav);
                }
            }

            Debug.WriteLine("*** 9. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();

            if (!Properties.Settings.Default.ViewToolbar)
            {
                mainToolbarToolStripMenuItem.Checked = false;
                mainToolbarToolStripMenuItem.CheckState = CheckState.Unchecked;
                tsMain.Visible = false;
            }

            if (!Properties.Settings.Default.ViewExpl1Statusbar)
            {
                tsmiViewStatusbar1.Checked = false;
                tsmiViewStatusbar1.CheckState = CheckState.Unchecked;
            }
            if (!Properties.Settings.Default.ViewExpl1Toolbar)
            {
                tsmiViewToolbar1.Checked = false;
                tsmiViewToolbar1.CheckState = CheckState.Unchecked;
                //flowLayoutPanel1.Visible = false;
            }

            if (!Properties.Settings.Default.ViewFavToolbar)
            {
                favoritesToolbarToolStripMenuItem.Checked = false;
                favoritesToolbarToolStripMenuItem.CheckState = CheckState.Unchecked;
                toolStripFav.Visible = false;
            }

            //
            TabPage tpnew1 = new TabPage();
            tpnew1.ImageIndex = 1;
            tpnew1.ToolTipText = "HALLO";
            tabControl1.TabPages.Add(tpnew1);

            //
            TabPage tpnew2 = new TabPage();
            tpnew2.ImageIndex = 1;
            tabControl2.TabPages.Add(tpnew2);

            tabControl1_SelectedIndexChanged(null, null);
            tabControl2_SelectedIndexChanged(null, null);

            //
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced");
            if (regKey != null)
            {
                object regValue = regKey.GetValue("Hidden");
                if (regValue != null)
                {
                    tsmiShowHiddenFiles.Checked = ((int)regValue != 0);
                }
            }
            Debug.WriteLine("*** 10. duration: {0} ms.", perfStopwatch.ElapsedMilliseconds);
            perfStopwatch.Restart();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            if (_args.Length > 0)
            {

                try
                {
                    string newPath = _args[0];

                    if (Directory.Exists(newPath))
                    {

                        UserControlExplorer curUCE = AddExplorer(tabControl1, "", true);
                        tabControl1.SelectedIndex = tabControl1.TabCount - 2;

                        if (curUCE != null)
                        {
                            curUCE.explorerBrowser.Navigate(ShellObject.FromParsingName(newPath));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }
        }

        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            applicationSettings.Save();

            Properties.Settings settings = Properties.Settings.Default;
            Properties.Settings.Default.Panel1H = panel1.Height;
            Properties.Settings.Default.Panel1W = panel1.Width;
            Properties.Settings.Default.FormL = Left;
            Properties.Settings.Default.FormT = Top;
            Properties.Settings.Default.FormW = Width;
            Properties.Settings.Default.FormH = Height;
            Properties.Settings.Default.ViewToolbar = tsMain.Visible;
            Properties.Settings.Default.WindowState = WindowState;
            settings.formsize = Size;
            settings.Save();

            Properties.Settings.Default.Save();
        }

        private void knownFolderCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Navigate to a known folder
                IKnownFolder kf = KnownFolderHelper.FromCanonicalName(tscbKnownFolder.Items[tscbKnownFolder.SelectedIndex].ToString());

                current_uce.explorerBrowser.Navigate((ShellObject)kf);
            }
            catch (COMException)
            {
                MessageBox.Show("Navigation not possible.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void navigationHistoryCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void tsbExplorer_Click(object sender, EventArgs e)
        {
            string Programmname = "explorer.exe";
            Process.Start(Programmname, current_uce.explorerBrowser.NavigationLog.CurrentLocation.ParsingName);
        }


        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormInfo formInfo = new FormInfo();
            formInfo.ShowDialog();
        }



        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }



        private void ClearLayout()
        {
            panel1.Dock = DockStyle.None;
            splitter.Dock = DockStyle.Bottom;
            panel2.Dock = DockStyle.None;
        }

        private void OrderLayout()
        {
            panel1.BringToFront();
            splitter.BringToFront();
            panel2.BringToFront();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (tsbLayout1.Checked)
            {
                return;
            }

            //Übereinander
            Properties.Settings.Default.Layout = TOPOFEACHOTHER;
            tsbLayout1.Checked = true;
            tsbLayout2.Checked = false;

            ClearLayout();

            if ((tsbShowHide1.Checked) && (tsbShowHide2.Checked))
            {
                panel1.Dock = DockStyle.Top;
                panel1.Height = (Height - menuStripMain.Height - tsMain.Height) / 2;
                splitter.Dock = DockStyle.Top;
                panel2.Dock = DockStyle.Fill;
            }
            else if ((tsbShowHide1.Checked) && (!tsbShowHide2.Checked))
            {
                panel1.Dock = DockStyle.Fill;
            }
            else
            {
                panel2.Dock = DockStyle.Fill;
            }

            OrderLayout();

            layout1ToolStripMenuItem.Checked = tsbLayout1.Checked;
            layout2ToolStripMenuItem.Checked = tsbLayout2.Checked;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (tsbLayout2.Checked)
            {
                return;
            }

            //Hintereinanden
            Properties.Settings.Default.Layout = NEXTTOEACHOTHER;
            tsbLayout2.Checked = true;
            tsbLayout1.Checked = false;

            ClearLayout();

            if ((tsbShowHide1.Checked) && (tsbShowHide2.Checked))
            {
                panel1.Dock = DockStyle.Left;
                panel1.Width = panel1.Width = Properties.Settings.Default.Panel1W;
                if (panel1.Width > (Width - 20))
                {
                    panel1.Width = Width / 2;
                }

                splitter.Dock = DockStyle.Left;
                panel2.Dock = DockStyle.Fill;
            }
            else if ((tsbShowHide1.Checked) && (!tsbShowHide2.Checked))
            {
                panel1.Dock = DockStyle.Fill;
                panel1.Width = panel1.Width = Properties.Settings.Default.Panel1W;
            }
            else
            {
                panel2.Dock = DockStyle.Fill;

            }

            OrderLayout();

            layout1ToolStripMenuItem.Checked = tsbLayout1.Checked;
            layout2ToolStripMenuItem.Checked = tsbLayout2.Checked;
        }


        private void AjustPanel()
        {
            if ((panel1.Visible) && (!panel2.Visible))
            {
                panel1.Dock = DockStyle.Fill;
            }
            else
            {
                if (tsbLayout1.Checked)
                {
                    panel1.Dock = DockStyle.Top;
                }
                else
                {
                    panel1.Dock = DockStyle.Left;
                }
            }

        }

        private void ToogleToolStripMenuItem(ToolStripMenuItem tsmi, Control obj)
        {
            if (tsmi.CheckState == CheckState.Checked)
            {
                tsmi.CheckState = CheckState.Unchecked;
            }
            else
            {
                tsmi.CheckState = CheckState.Checked;
            }

            obj.Visible = tsmi.CheckState == CheckState.Checked;

        }

        private void tsmiViewToolbar1_Click(object sender, EventArgs e)
        {
            //ToogleToolStripMenuItem(tsmiViewToolbar1, flowLayoutPanel1);
        }

        private void tsmiViewStatusbar1_Click(object sender, EventArgs e)
        {
            //ToogleToolStripMenuItem(tsmiViewStatusbar1, statusStrip1);
        }

        private void tsmiViewToolbar2_Click(object sender, EventArgs e)
        {
            //ToogleToolStripMenuItem(tsmiViewToolbar2, flowLayoutPanel2);
        }

        private void tsmiViewStatusbar2_Click(object sender, EventArgs e)
        {
            //ToogleToolStripMenuItem(tsmiViewStatusbar2, statusStrip2);
        }

        private void tsbSwitch_Click_1(object sender, EventArgs e)
        {
            ShellObject so = current_uce1.explorerBrowser.NavigationLog.CurrentLocation;
            current_uce1.explorerBrowser.Navigate(current_uce2.explorerBrowser.NavigationLog.CurrentLocation);
            current_uce2.explorerBrowser.Navigate(so);
        }


        private void path_Changed(string pathname, string name, UserControlExplorer uce)
        {
            uce.Parent.Text = pathname;
            if ((uce.TabControlLocked == true) && (uce.LockedPath == pathname))
            {
                (uce.Parent as TabPage).ImageIndex = 0;
            }
            else if ((uce.TabControlLocked == true) && (uce.LockedPath != pathname))
            {
                (uce.Parent as TabPage).ImageIndex = 2;
            }
        }


        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            Section section = applicationSettings.sections.Add("A1", "");
            section.settings.Add("AA1", "1");
            section.settings.Add("AA2", "2");
            section = applicationSettings.sections.Add("B1", "");
            section.settings.Add("BB1", "3");
        }

        private void tsmiNewTab_Click(object sender, EventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);

            AddExplorer(tc, "", true);
            tc.SelectedIndex = tc.TabCount - 2;
            ReorderTabIDs(tc);
        }

        private void closeTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);
            TabPage tp = tc.SelectedTab;
            if (tp != null)
            {
                string removeSectionName = tp.Tag.ToString();
                applicationSettings.sections.Remove(applicationSettings.sections.GetItemByString(removeSectionName, "FAV"));
                applicationSettings.sections.Remove(applicationSettings.sections.GetItemByString(removeSectionName, "OPTIONS"));
                tc.TabPages.Remove(tp);
            }
            ReorderTabIDs(tc);
        }

        private void cmsTabControl_Opening(object sender, CancelEventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);

            if (tc == null)
            {
                return;
            }

            tsmiLockTab.Enabled = tc.SelectedTab.Text.Contains(":");
            tsmiLockTab.Checked = ((tc.SelectedTab.ImageIndex == 0) || (tc.SelectedTab.ImageIndex == 2));
            tsmiResetTabPath.Enabled = (tsmiLockTab.Checked == true);


            closeTabToolStripMenuItem.Enabled = ((tc.TabCount > 1) && (!tsmiLockTab.Checked));
        }

        private void tsmiLockTab_Click(object sender, EventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);
            TabPage tp = tc.SelectedTab;
            if (tsmiLockTab.Checked)
            {
                tp.ImageIndex = 0;
            }
            else
            {
                tp.ImageIndex = -1;
            }

            string sectionName = tp.Tag.ToString();
            applicationSettings.sections.GetItemByString(sectionName, "OPTIONS").settings.GetItemByString("Locked").Value =
                tsmiLockTab.Checked.ToString();

            current_uce.TabControlLocked = tsmiLockTab.Checked;
        }


        private void jumpTo_Click(object sender, EventArgs e)
        {
            string jumptoFolder = (sender as ToolStripMenuItem).Text;

            try
            {
                // Navigate to a known folder
                IKnownFolder kf = KnownFolderHelper.FromCanonicalName(jumptoFolder);

                current_uce.explorerBrowser.Navigate((ShellObject)kf);
            }
            catch (COMException)
            {
                MessageBox.Show("Navigation not possible.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void flowLayoutPanel_DragEnter(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            e.Effect = DragDropEffects.Link;
        }

        private void flowLayoutPanel_DragDrop(object sender, DragEventArgs e)
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

            Properties.Settings.Default.Favs.Add(myDir);

            CreateFavButton(myDirCaption, myDir);
        }

        private void CreateFavButton(string FavText, string FavPath)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(FavPath, Properties.Resources.StartSmal, myFav_Click);
            ToolStripMenuItem itemDel = new ToolStripMenuItem("Remove from Favorites", Properties.Resources.remove, myFavRemove_Click);
            itemDel.ToolTipText = FavPath;
            itemDel.Tag = FavText;
            item.DropDownItems.Add(itemDel);
            favoritesToolStripMenuItem.DropDownItems.Add(item);

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
            toolStripFav.Items.Add(mybtn);
        }

        private void mysbtn_Click(object sender, EventArgs e)
        {
            NavigateTo((sender as ToolStripButton).ToolTipText);
        }


        private void myFav_Click(object sender, EventArgs e)
        {
            NavigateTo((sender as ToolStripMenuItem).Text);
        }

        private bool NavigateTo(string path)
        {
            if (Directory.Exists(path))
            {
                try
                {
                    current_uce.explorerBrowser.Navigate(ShellObject.FromParsingName(path));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }
            else
            {
                MessageBox.Show(string.Format("Path [{0}] not exist! Maybe an old temporary drive like an usb stick or network drive.", path), "Path Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void mybtn_Click(object sender, EventArgs e)
        {
            NavigateTo(toolTip.GetToolTip(sender as Button));
        }

        private void RemoveFavItem(string name)
        {
            foreach (ToolStripMenuItem item in favoritesToolStripMenuItem.DropDownItems)
            {
                if (item.Text == name)
                {
                    item.DropDownItems.Clear();
                    item.Dispose();
                    return;
                }
            }
        }

        private void RemoveFavItemFromToolstrip(string name)
        {
            foreach (ToolStripItem myControl in toolStripFav.Items)
            {
                if (myControl.Text == name)
                {
                    myControl.Dispose();
                    return;
                }
            }
        }

        private void myFavRemove_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (sender as ToolStripMenuItem);
            string mycaption = item.Tag.ToString();
            string mypath = item.ToolTipText;

            RemoveFavItem(mypath);
            RemoveFavItemFromToolstrip(mycaption);

            Properties.Settings.Default.Favs.Remove(mypath);
        }

        private void mainToolbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsMain.Visible = mainToolbarToolStripMenuItem.Checked;
            if (tsMain.Visible)
            {
                tsMain.BringToFront();
                toolStripFav.BringToFront();
                OrderLayout();
            }
        }

        private void favoritesToolbarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripFav.Visible = favoritesToolbarToolStripMenuItem.Checked;
            Properties.Settings.Default.ViewFavToolbar = favoritesToolbarToolStripMenuItem.Checked;

            if (toolStripFav.Visible)
            {
                tsMain.BringToFront();
                toolStripFav.BringToFront();
                OrderLayout();
            }
        }

        private void tsbShowHide1_Click(object sender, EventArgs e)
        {
            panel1.Visible = tsbShowHide1.Checked;

            if (!panel1.Visible)
            {
                panel2.Visible = true;
                tsbShowHide2.Checked = true;
            }

            showHideExplorer1ToolStripMenuItem.Checked = tsbShowHide1.Checked;
            showHideExplorer2ToolStripMenuItem.Checked = tsbShowHide2.Checked;

            AjustPanel();
        }

        private void tsbShowHide2_Click(object sender, EventArgs e)
        {
            panel2.Visible = tsbShowHide2.Checked;

            if (!panel2.Visible)
            {
                panel1.Visible = true;
                tsbShowHide1.Checked = true;
            }

            showHideExplorer1ToolStripMenuItem.Checked = tsbShowHide1.Checked;
            showHideExplorer2ToolStripMenuItem.Checked = tsbShowHide2.Checked;
            AjustPanel();
        }

        private void showHideExplorer1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbShowHide1.Checked = showHideExplorer1ToolStripMenuItem.Checked;
            tsbShowHide1_Click(sender, null);
        }

        private void showHideExplorer2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tsbShowHide2.Checked = showHideExplorer2ToolStripMenuItem.Checked;
            tsbShowHide2_Click(sender, null);
        }

        private UserControlExplorer AddExplorer(TabControl DestinationTabControl, string name, bool newMode)
        {
            if (DestinationTabControl == null)
            {
                return null;
            }

            string newSectionName = applicationSettings.GetNewSectionName();

            if ((name != "") || (newSectionName != ""))
            {
                TabPage newTabPage = new TabPage();

                if (!newMode)
                {
                    DestinationTabControl.TabPages.Add(newTabPage);
                }
                else
                {
                    int insertPosition = DestinationTabControl.TabCount - 1;

                    DestinationTabControl.TabPages.Insert(insertPosition, newTabPage);   //Add(newTabPage);
                }

                UserControlExplorer uce = new UserControlExplorer(newSectionName, applicationSettings, name, DestinationTabControl, newTabPage);
                uce.Parent = newTabPage;
                uce.Dock = DockStyle.Fill;
                uce.MyEvent += new MyDelegate(GetMyEvent);
                uce.pathChanged += new UserControlExplorer.PathChanged(path_Changed);
                uce.myPreviewKeyDown += new UserControlExplorer.MyPreviewKeyDown(uce_myPreviewKeyDown);

                return uce;
            }

            return null;
        }

        private void GetMyEvent(object sender, EventArgs e)
        {
            try
            {
                current_uce = (sender as UserControlExplorer);

                MarkTabControl(current_uce.Parent.Parent.Name);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void MarkTabControl(string tabControlName)
        {

            if (tabControlName == tabControl1.Name)
            {
                panel1.BackColor = Color.DimGray;
                panel2.BackColor = Color.Transparent;
            }
            else
            {
                panel1.BackColor = Color.Transparent;
                panel2.BackColor = Color.DimGray;
            }
        }

        private void wdecodeplexcomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Run("http://wde.codeplex.com");
        }

        private void Run(string AppName)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = AppName;

            Process.Start(psi);
        }

        public static string GetUrlResponse(string url, string username, string password, bool showMsgOnError)
        {
            string content = "";

            WebRequest webRequest = WebRequest.Create(url);

            if (username == null || password == null)
            {
                NetworkCredential networkCredential = new NetworkCredential(username, password);
                webRequest.PreAuthenticate = true;
                webRequest.Credentials = networkCredential;
            }

            try
            {
                WebResponse webResponse = webRequest.GetResponse();

                StreamReader sr = new StreamReader(webResponse.GetResponseStream(), Encoding.ASCII);
                //StringBuilder contentBuilder = new StringBuilder();
                //while (-1 != sr.Peek())
                //{
                //    contentBuilder.Append(sr.ReadLine());
                //    contentBuilder.Append("\r\n");
                //}
                //content = contentBuilder.ToString();

                content = sr.ReadToEnd();

            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                if (showMsgOnError)
                {
                    MessageBox.Show(e.ToString());
                }
            }

            return content;
        }

        private void tabControl1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string curLocation = current_uce.explorerBrowser.NavigationLog.CurrentLocation.ParsingName;
                AddExplorer(tabControl2, "", true);
                tabControl2.SelectedIndex = tabControl2.TabCount - 2;

                UserControlExplorer cur_uce = tabControl2.SelectedTab.Controls["UserControlExplorer"] as UserControlExplorer;
                cur_uce.explorerBrowser.Navigate(ShellFileSystemFolder.FromFolderPath(curLocation));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tabControl2_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string curLocation = current_uce.explorerBrowser.NavigationLog.CurrentLocation.ParsingName;
                AddExplorer(tabControl1, "", true);
                tabControl1.SelectedIndex = tabControl1.TabCount - 2;

                UserControlExplorer cur_uce = tabControl1.SelectedTab.Controls["UserControlExplorer"] as UserControlExplorer;
                cur_uce.explorerBrowser.Navigate(ShellFileSystemFolder.FromFolderPath(curLocation));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tabControl1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, e.X, e.Y, 0, 0);
            }
            SetCurrentUCE(tabControl1);
        }

        private void tabControl2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, e.X, e.Y, 0, 0);
            }
            SetCurrentUCE(tabControl2);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetCurrentUCE(tabControl1);
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetCurrentUCE(tabControl2);
        }
        private void SetCurrentUCE(TabControl tabControlx)
        {
            if (tabControlx.SelectedTab.ImageIndex == 1)
            {
                AddExplorer(tabControlx, "", true);
                tabControlx.SelectedIndex = tabControlx.TabCount - 2;
            }
            else
            {
                try
                {
                    string sectionName = tabControlx.SelectedTab.Tag.ToString();
                    if (tabControlx == tabControl1)
                    {
                        current_uce1 = (tabControlx.SelectedTab.Controls["UserControlExplorer"] as UserControlExplorer);
                        current_uce = current_uce1;
                    }
                    else if (tabControlx == tabControl2)
                    {
                        current_uce2 = (tabControlx.SelectedTab.Controls["UserControlExplorer"] as UserControlExplorer);
                        current_uce = current_uce2;
                    }

                    foreach (TabPage tp in tabControlx.TabPages)
                    {
                        try
                        {
                            if (tp.Controls["UserControlExplorer"] is UserControlExplorer)
                            {
                                (tp.Controls["UserControlExplorer"] as UserControlExplorer).ShowSelected = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex);
                        }
                    }

                    MarkTabControl(tabControlx.Name);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }
            }
            current_uce.ShowSelected = true;
        }

        private void toolStripFav_DragDrop(object sender, DragEventArgs e)
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

            Properties.Settings.Default.Favs.Add(myDir);

            CreateFavButton(myDirCaption, myDir);
        }

        private void toolStripFav_DragEnter(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            e.Effect = DragDropEffects.Link;
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (contextTSB == null)
            {
                return;
            }

            string mypath = contextTSB.ToolTipText;

            RemoveFavItem(mypath);
            Properties.Settings.Default.Favs.Remove(mypath);
            contextTSB.Dispose();
        }

        private void toolStripFav_MouseDown(object sender, MouseEventArgs e)
        {
            contextTSB = null;
        }

        private void toolStripButton1_MouseDown(object sender, MouseEventArgs e)
        {
            contextTSB = (sender as ToolStripButton);
        }

        private void mapNetworkDriveToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void disconnectDriveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormDisconnctDrive fdd = new FormDisconnctDrive();
            fdd.Show();
            Update();
        }

        private void toolStripButton1_Click_3(object sender, EventArgs e)
        {
            Process P = new Process();
            P.StartInfo.FileName = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase;
            P.Start();
        }

        private void tsmiResetTabPath_Click(object sender, EventArgs e)
        {
            current_uce.ResetToLockedPath();
        }

        private void tsmiMainResetTabPath_Click(object sender, EventArgs e)
        {
            ResetAllTabPages(tabControl1);
            ResetAllTabPages(tabControl2);
        }

        private void ResetAllTabPages(TabControl tc)
        {
            foreach (TabPage item in tc.TabPages)
            {
                if (item.Text != string.Empty)
                {
                    UserControlExplorer my_uce = item.Controls["UserControlExplorer"] as UserControlExplorer;
                    my_uce.ResetToLockedPath();
                }
            }
        }

        private void FormMain_Shown(object sender, EventArgs e)
        {
            formSplash.Close();
        }

        private void tabControl1_DragDrop(object sender, DragEventArgs e)
        {
            tabControlDragDrop(tabControl1, e);
        }

        private void tabControl1_DragEnter(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            e.Effect = DragDropEffects.Link;
        }

        private void tabControl2_DragDrop(object sender, DragEventArgs e)
        {
            tabControlDragDrop(tabControl2, e);
        }

        private void tabControlDragDrop(TabControl tc, DragEventArgs e)
        {
            String[] Params = (String[])e.Data.GetData(DataFormats.FileDrop);
            string myParam = Params[0];
            string myDir;

            if (File.Exists(myParam))
            {
                myDir = Path.GetDirectoryName(myParam);
            }
            else
            {
                myDir = Path.GetFullPath(myParam);
            }

            try
            {
                AddExplorer(tc, "", true);
                tc.SelectedIndex = tc.TabCount - 2;

                UserControlExplorer cur_uce = tc.SelectedTab.Controls["UserControlExplorer"] as UserControlExplorer;
                cur_uce.explorerBrowser.Navigate(ShellFileSystemFolder.FromFolderPath(myDir));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void tsmiShowHiddenFiles_Click(object sender, EventArgs e)
        {
            RegistryKey regKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", true);
            if (regKey != null)
            {
                int value = 0;
                if (tsmiShowHiddenFiles.Checked == true)
                {
                    value = 1;
                }

                regKey.SetValue("Hidden", value, RegistryValueKind.DWord);

                if (current_uce1.explorerBrowser.NavigationLog.CurrentLocation.Parent != null)
                {
                    current_uce1.explorerBrowser.Navigate(current_uce1.explorerBrowser.NavigationLog.CurrentLocation.Parent);
                }

                if (current_uce2.explorerBrowser.NavigationLog.CurrentLocation.Parent != null)
                {
                    current_uce2.explorerBrowser.Navigate(current_uce2.explorerBrowser.NavigationLog.CurrentLocation.Parent);
                }

                current_uce1.explorerBrowser.NavigateLogLocation(NavigationLogDirection.Backward);
                current_uce2.explorerBrowser.NavigateLogLocation(NavigationLogDirection.Backward);
            }
        }

        private void RefreshAllExplorer()
        {
            MessageBox.Show("Please press F5 to refresh");
            // TODO - implement this
        }

        private void tsmiPlus_Click(object sender, EventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);

            if (tc == null)
            {
                return;
            }

            TabUp(tc);
        }

        private void TabUp(TabControl tc)
        {
            TabPage tp = tc.SelectedTab;
            if (tp != null)
            {
                int i = tc.SelectedIndex;

                if (i >= tc.TabPages.Count - 2)
                {
                    return;
                }

                TabPage tp2 = tc.TabPages[tc.SelectedIndex + 1];

                tc.TabPages[tc.SelectedIndex] = tp2;
                tc.TabPages[tc.SelectedIndex + 1] = tp;

                tc.SelectedIndex = i + 1;

                ReorderTabIDs(tc);
            }
        }

        private void tsmiMinus_Click(object sender, EventArgs e)
        {
            TabControl tc = (cmsTabControl.SourceControl as TabControl);

            if (tc == null)
            {
                return;
            }

            TabDown(tc);
        }

        private void TabDown(TabControl tc)
        {
            TabPage tp = tc.SelectedTab;
            if (tp != null)
            {
                int i = tc.SelectedIndex;

                if (i == 0)
                {
                    return;
                }

                TabPage tp2 = tc.TabPages[tc.SelectedIndex - 1];

                tc.TabPages[tc.SelectedIndex] = tp2;
                tc.TabPages[tc.SelectedIndex - 1] = tp;

                tc.SelectedIndex = i - 1;

                ReorderTabIDs(tc);
            }
        }

        private void ReorderTabIDs(TabControl tc)
        {
            for (int j = 0; j < tc.TabPages.Count - 1; j++)
            {
                if (tc.TabPages[j].Controls["UserControlExplorer"] is UserControlExplorer)
                {
                    (tc.TabPages[j].Controls["UserControlExplorer"] as UserControlExplorer).SavedTabIndex = j;
                }
            }
        }

        private void FormMain_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.Add))
            {
                TabUp(current_uce.CurrentTabControl);
            }
            else if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.Subtract))
            {
                TabDown(current_uce.CurrentTabControl);
            }
            else if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.U))
            {
                if (current_uce == current_uce1)
                {
                    current_uce = current_uce2;
                }
                else
                {
                    current_uce = current_uce1;
                }

                MarkTabControl(current_uce.Parent.Parent.Name);
                current_uce.Focus();
            }
        }

        void uce_myPreviewKeyDown(object sender, PreviewKeyDownEventArgs e, TabControl dTabControl)
        {
            if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.Add))
            {
                TabUp(dTabControl);
            }
            else if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.Subtract))
            {
                TabDown(dTabControl);
            }
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabPage CurrentTab = tabControl1.TabPages[e.Index];
            Rectangle ItemRect = tabControl1.GetTabRect(e.Index);
            var TextBrush = new SolidBrush(Color.White);
            var sf = new StringFormat();
            sf.Alignment = StringAlignment.Center;
            sf.LineAlignment = StringAlignment.Center;

            //If we are currently painting the Selected TabItem we'll
            //change the brush colors and inflate the rectangle.
            if (Convert.ToBoolean(e.State & DrawItemState.Selected))
            {
                TextBrush.Color = Color.Red;
                ItemRect.Inflate(2, 2);
            }

            //Set up rotation for left and right aligned tabs
            if (tabControl1.Alignment == TabAlignment.Left || tabControl1.Alignment == TabAlignment.Right)
            {
                float RotateAngle = 90;
                if (tabControl1.Alignment == TabAlignment.Left)
                {
                    RotateAngle = 270;
                }

                PointF cp = new PointF(ItemRect.Left + (ItemRect.Width / 2), ItemRect.Top + (ItemRect.Height / 2));
                e.Graphics.TranslateTransform(cp.X, cp.Y);
                e.Graphics.RotateTransform(RotateAngle);
                ItemRect = new Rectangle(-(ItemRect.Height / 2), -(ItemRect.Width / 2), ItemRect.Height, ItemRect.Width);
            }

            //Now draw the text.
            e.Graphics.DrawString(CurrentTab.Text, e.Font, TextBrush, (RectangleF)ItemRect, sf);

            //Reset any Graphics rotation
            e.Graphics.ResetTransform();

            //Finally, we should Dispose of our brushes.
            TextBrush.Dispose();
        }

        private void sortPages(TabControl tc)
        {
            for (int i = 0; i < tc.TabPages.Count - 1; i++)
            {
                TabPage tp = tc.TabPages[i];
                int id = (tp.Controls["UserControlExplorer"] as UserControlExplorer).SavedTabIndex;

                for (int j = i + 1; j < tc.TabPages.Count; j++)
                {
                    TabPage tp2 = tc.TabPages[j];
                    int id2 = (tp2.Controls["UserControlExplorer"] as UserControlExplorer).SavedTabIndex;

                    if (id > id2)
                    {
                        tc.TabPages[j] = tp;
                        tc.TabPages[i] = tp2;
                        break;
                    }
                }
            }

            for (int i = 0; i < tc.TabPages.Count; i++)
            {
                TabPage tp = tc.TabPages[i];
                if ((tp.Controls["UserControlExplorer"] as UserControlExplorer).ShowSelected == true)
                {
                    tc.SelectedIndex = i;
                    break;
                }
            }
        }

        private void tsmiMoveTabUp_Click(object sender, EventArgs e)
        {
            TabUp(current_uce.CurrentTabControl);
        }

        private void tsmiMoveTabDown_Click(object sender, EventArgs e)
        {
            TabDown(current_uce.CurrentTabControl);
        }
    }
}
