namespace TaskbarGroups.UserControls
{
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;

    using Classes;

    using Forms;

    public partial class ucShortcut : UserControl
    {
        public ucShortcut()
        {
            InitializeComponent();
        }

        public ProgramShortcut Psc { get; set; }
        public frmMain MotherForm { get; set; }
        public Category ThisCategory { get; set; }

        private void ucShortcut_Load(object sender, EventArgs e)
        {
            Show();
            BringToFront();
            BackColor = MotherForm.BackColor;
            picIcon.BackgroundImage =
                ThisCategory.LoadImageCache(Psc); // Use the local icon cache for the file specified as the icon image
        }

        public void ucShortcut_Click(object sender, EventArgs e)
        {
            if (Psc.IsWindowsApp)
            {
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo
                        {UseShellExecute = true, FileName = $@"shell:appsFolder\{Psc.FilePath}"}
                };
                p.Start();
            }
            else
            {
                if (Path.GetExtension(Psc.FilePath).ToLower() == ".lnk" && Psc.FilePath == MainPath.exeString)
                    MotherForm.OpenFile(Psc.Arguments, Psc.FilePath, MainPath.path);
                else
                    MotherForm.OpenFile(Psc.Arguments, Psc.FilePath, Psc.WorkingDirectory);
            }
        }

        public void ucShortcut_MouseEnter(object sender, EventArgs e)
        {
            BackColor = MotherForm.HoverColor;
        }

        public void ucShortcut_MouseLeave(object sender, EventArgs e)
        {
            BackColor = Color.Transparent;
        }
    }
}