using System;
using System.Windows.Forms;
using System.Configuration;

namespace WDE
{
    public partial class FormInfo : Form
    {
        public FormInfo()
        {
            InitializeComponent();
        }

        private void FormInfo_Load(object sender, EventArgs e)
        {
            lblVersion.Text = "Version " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

            textBox1.Clear();
            // ______________
            // USER-Settings:
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            textBox1.AppendText("Path to 'user.config':\r\n" + config.FilePath + "\r\n\r\n");


            // ____________________
            // Application-Settings:
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            textBox1.AppendText("Path to Application-Settings:\r\n" + config.FilePath + "\r\n\r\n");

            // ____________________
            // Machine.config-Pfad:
            config = ConfigurationManager.OpenMachineConfiguration();
            textBox1.AppendText("Path to 'machine.config':\r\n" + config.FilePath + "\r\n\r\n");

            textBox1.AppendText("Application.LocalUserAppDataPath':\r\n" + Application.LocalUserAppDataPath + "\r\n\r\n");
            textBox1.AppendText("Application.UserAppDataPath ':\r\n" + Application.UserAppDataPath + "\r\n\r\n");
        }

        private void label3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
            psi.FileName = "http://www.it-nonstop.de";

            System.Diagnostics.Process.Start(psi);
        }
    }
}
