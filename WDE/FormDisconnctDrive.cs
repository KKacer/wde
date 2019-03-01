using System;
using System.Windows.Forms;
using System.IO;

namespace WDE
{
    public partial class FormDisconnctDrive : Form
    {
        /// <summary>
        /// API Call to Drop existing Network Connection
        /// </summary>
        /// <param name="lpName">Pointer to a constant null-terminated string that specifies the name of either the redirected local device or the remote network resource to disconnect from.</param>
        /// <param name="dwFlags">Connection type.  0 to retain only for current session or CONNECT_UPDATE_PROFILE to persist on session being logged out.</param>
        /// <param name="fForce">Specifies whether the disconnection should occur if there are open files or jobs on the connection. If this parameter is FALSE, the function fails if there are open files or jobs.</param>
        /// <returns></returns>
        [System.Runtime.InteropServices.DllImport("Mpr.dll",
            EntryPoint = "WNetCancelConnection2A",
            CharSet = System.Runtime.InteropServices.CharSet.Ansi)]
        private static extern int DropNetworkConnection(string lpName, int dwFlags, bool fForce);
        internal const int NO_ERROR = 0x000;
        internal const int ERROR_BAD_PROFILE = 0x4B6;
        internal const int ERROR_CANNOT_OPEN_PROFILE = 0x4B5;
        internal const int ERROR_DEVICE_IN_USE = 0x964;
        internal const int ERROR_EXTENDED_ERROR = 0x4B8;
        internal const int ERROR_NOT_CONNECTED = 0x8CA;
        internal const int ERROR_OPEN_FILES = 0x961;

        public FormDisconnctDrive()
        {
            InitializeComponent();
        }

        private void FormDisconnctDrive_Load(object sender, EventArgs e)
        {
            ReadAllNetworkDrives();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ReadAllNetworkDrives()
        {
            listViewNetDrives.Items.Clear();

            foreach (var driveInfo in DriveInfo.GetDrives())
            {
                if (driveInfo.DriveType == DriveType.Network)
                {
                    ListViewItem item = listViewNetDrives.Items.Add(driveInfo.Name + " ("+ driveInfo.VolumeLabel+")");
                    item.Tag = driveInfo.Name;
                }
            }
        }

        private void listViewNetDrives_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonOK.Enabled = (listViewNetDrives.SelectedItems.Count > 0);
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewNetDrives.SelectedItems)
            {
                string name = item.Tag.ToString();
                DropNetworkConnection(name.Substring(0, 2), 0, false);
            }
            ReadAllNetworkDrives();
            Close();
        }
    }
}
