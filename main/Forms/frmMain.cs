namespace TaskbarGroups.Forms
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Runtime;
    using System.Windows.Forms;

    using Classes;

    using UserControls;

    public sealed partial class frmMain : Form
    {
        private readonly string passedDirec;
        public List<ucShortcut> ControlList;
        public Color HoverColor;
        public Point mouseClick;

        public Panel shortcutPanel;

        //
        // endregion
        //
        public PictureBox shortcutPic;

        public Category ThisCategory;

        //------------------------------------------------------------------------------------
        // CTOR AND LOAD
        //
        public frmMain(string passedDirectory, int cursorPosX, int cursorPosY)
        {
            InitializeComponent();

            ProfileOptimization.StartProfile("frmMain.Profile");
            mouseClick = new Point(cursorPosX, cursorPosY); // Consstruct point p based on passed x y mouse values
            passedDirec = passedDirectory;
            FormBorderStyle = FormBorderStyle.None;

            using (var ms = new MemoryStream(File.ReadAllBytes(MainPath.path + "\\config\\" + passedDirec + "\\GroupIcon.ico")))
            {
                Icon = new Icon(ms);
            }

            if (Directory.Exists(MainPath.path + @"\config\" + passedDirec))
            {
                ControlList = new List<ucShortcut>();

                SetStyle(ControlStyles.SupportsTransparentBackColor, true);
                ThisCategory = new Category($"config\\{passedDirec}");
                BackColor = ImageFunctions.FromString(ThisCategory.ColorString);
                Opacity = 1 - ThisCategory.Opacity / 100;

                HoverColor = BackColor.R * 0.2126 + BackColor.G * 0.7152 + BackColor.B * 0.0722 > 255 / 2
                    ? Color.FromArgb(BackColor.A, BackColor.R - 50, BackColor.G - 50, BackColor.B - 50)
                    : Color.FromArgb(BackColor.A, BackColor.R + 50, BackColor.G + 50, BackColor.B + 50);
            }
            else
                Application.Exit();
        }

        // Allow doubleBuffering drawing each frame to memory and then onto screen
        // Solves flickering issues mostly as the entire rendering of the screen is done in 1 operation after being first loaded to memory
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;
                return cp;
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            LoadCategory();
            SetLocation();
        }

        // Sets location of form
        private void SetLocation()
        {
            var taskbarList = FindDockedTaskBars();
            var taskbar = new Rectangle();
            var screen = new Rectangle();

            var i = 0;
            int locationy;
            int locationx;
            if (taskbarList.Count != 0)
            {
                foreach (var scr in Screen.AllScreens) // Get what screen user clicked on
                {
                    if (scr.Bounds.Contains(mouseClick))
                    {
                        screen.X = scr.Bounds.X;
                        screen.Y = scr.Bounds.Y;
                        screen.Width = scr.Bounds.Width;
                        screen.Height = scr.Bounds.Height;
                        taskbar = taskbarList[i];
                    }

                    i++;
                }

                if (taskbar.Contains(mouseClick)) // Click on taskbar
                {
                    if (taskbar.Top == screen.Top && taskbar.Width == screen.Width)
                    {
                        // TOP
                        locationy = screen.Y + taskbar.Height + 10;
                        locationx = mouseClick.X - Width / 2;
                    }
                    else if (taskbar.Bottom == screen.Bottom && taskbar.Width == screen.Width)
                    {
                        // BOTTOM
                        locationy = screen.Y + screen.Height - Height - taskbar.Height - 10;
                        locationx = mouseClick.X - Width / 2;
                    }
                    else if (taskbar.Left == screen.Left)
                    {
                        // LEFT
                        locationy = mouseClick.Y - Height / 2;
                        locationx = screen.X + taskbar.Width + 10;
                    }
                    else
                    {
                        // RIGHT
                        locationy = mouseClick.Y - Height / 2;
                        locationx = screen.X + screen.Width - Width - taskbar.Width - 10;
                    }
                }
                else // not click on taskbar
                {
                    locationy = mouseClick.Y - Height - 20;
                    locationx = mouseClick.X - Width / 2;
                }

                Location = new Point(locationx, locationy);

                // If form goes over screen edge
                if (Left < screen.Left)
                    Left = screen.Left + 10;
                if (Top < screen.Top)
                    Top = screen.Top + 10;
                if (Right > screen.Right)
                    Left = screen.Right - Width - 10;

                // If form goes over taskbar
                if (taskbar.Contains(Left, Top) && taskbar.Contains(Right, Top)) // Top taskbar
                    Top = screen.Top + 10 + taskbar.Height;
                if (taskbar.Contains(Left, Top)) // Left taskbar
                    Left = screen.Left + 10 + taskbar.Width;
                if (taskbar.Contains(Right, Top)) // Right taskbar
                    Left = screen.Right - Width - 10 - taskbar.Width;
            }
            else // Hidden taskbar
            {
                foreach (var scr in Screen.AllScreens) // get what screen user clicked on
                {
                    if (scr.Bounds.Contains(mouseClick))
                    {
                        screen.X = scr.Bounds.X;
                        screen.Y = scr.Bounds.Y;
                        screen.Width = scr.Bounds.Width;
                        screen.Height = scr.Bounds.Height;
                    }

                    i++;
                }

                if (mouseClick.Y > Screen.PrimaryScreen.Bounds.Height - 35)
                    locationy = Screen.PrimaryScreen.Bounds.Height - Height - 45;
                else
                    locationy = mouseClick.Y - Height - 20;
                locationx = mouseClick.X - Width / 2;

                Location = new Point(locationx, locationy);

                // If form goes over screen edge
                if (Left < screen.Left)
                    Left = screen.Left + 10;
                if (Top < screen.Top)
                    Top = screen.Top + 10;
                if (Right > screen.Right)
                    Left = screen.Right - Width - 10;

                // If form goes over taskbar
                if (taskbar.Contains(Left, Top) && taskbar.Contains(Right, Top)) // Top taskbar
                    Top = screen.Top + 10 + taskbar.Height;
                if (taskbar.Contains(Left, Top)) // Left taskbar
                    Left = screen.Left + 10 + taskbar.Width;
                if (taskbar.Contains(Right, Top)) // Right taskbar
                    Left = screen.Right - Width - 10 - taskbar.Width;
            }
        }

        // Search for active taskbars on screen
        public static List<Rectangle> FindDockedTaskBars()
        {
            var dockedRects = new List<Rectangle>();
            foreach (var tmpScrn in Screen.AllScreens)
                if (!tmpScrn.Bounds.Equals(tmpScrn.WorkingArea))
                {
                    var rect = new Rectangle();

                    var leftDockedWidth = Math.Abs(Math.Abs(tmpScrn.Bounds.Left) - Math.Abs(tmpScrn.WorkingArea.Left));
                    var topDockedHeight = Math.Abs(Math.Abs(tmpScrn.Bounds.Top) - Math.Abs(tmpScrn.WorkingArea.Top));
                    var rightDockedWidth = tmpScrn.Bounds.Width - leftDockedWidth - tmpScrn.WorkingArea.Width;
                    var bottomDockedHeight = tmpScrn.Bounds.Height - topDockedHeight - tmpScrn.WorkingArea.Height;
                    if (leftDockedWidth > 0)
                    {
                        rect.X = tmpScrn.Bounds.Left;
                        rect.Y = tmpScrn.Bounds.Top;
                        rect.Width = leftDockedWidth;
                        rect.Height = tmpScrn.Bounds.Height;
                    }
                    else if (rightDockedWidth > 0)
                    {
                        rect.X = tmpScrn.WorkingArea.Right;
                        rect.Y = tmpScrn.Bounds.Top;
                        rect.Width = rightDockedWidth;
                        rect.Height = tmpScrn.Bounds.Height;
                    }
                    else if (topDockedHeight > 0)
                    {
                        rect.X = tmpScrn.WorkingArea.Left;
                        rect.Y = tmpScrn.Bounds.Top;
                        rect.Width = tmpScrn.WorkingArea.Width;
                        rect.Height = topDockedHeight;
                    }
                    else if (bottomDockedHeight > 0)
                    {
                        rect.X = tmpScrn.WorkingArea.Left;
                        rect.Y = tmpScrn.WorkingArea.Bottom;
                        rect.Width = tmpScrn.WorkingArea.Width;
                        rect.Height = bottomDockedHeight;
                    }

                    dockedRects.Add(rect);
                }

            if (dockedRects.Count == 0)
            {
                // Taskbar is set to "Auto-Hide".
            }

            return dockedRects;
        }

        //
        //------------------------------------------------------------------------------------
        //

        // Loading category and building shortcuts
        private void LoadCategory()
        {
            //System.Diagnostics.Debugger.Launch();

            Width = 0;
            Height = 45;
            var x = 0;
            var y = 0;
            var width = ThisCategory.Width;
            var columns = 1;

            // Check if icon caches exist for the category being loaded
            // If not then rebuild the icon cache
            if (!Directory.Exists(MainPath.path + @"\config\" + ThisCategory.Name + @"\Icons\"))
                ThisCategory.CacheIcons();

            foreach (var psc in ThisCategory.ShortcutList)
            {
                if (columns > width) // creating new row if there are more psc than max width
                {
                    x = 0;
                    y += 45;
                    Height += 45;
                    columns = 1;
                }

                if (Width < width * 55)
                    Width += 55;

                // OLD
                //BuildShortcutPanel(x, y, psc);

                // Building shortcut controls
                var pscPanel = new ucShortcut
                {
                    Psc = psc,
                    MotherForm = this,
                    ThisCategory = ThisCategory
                };
                pscPanel.Location = new Point(x, y);
                Controls.Add(pscPanel);
                ControlList.Add(pscPanel);
                pscPanel.Show();
                pscPanel.BringToFront();

                // Reset values
                x += 55;
                columns++;
            }

            Width -= 2; // For some reason the width is 2 pixels larger than the shortcuts. Temporary fix
        }

        // OLD (Having some issues with the uc build, so keeping the old code below)
        private void BuildShortcutPanel(int x, int y, ProgramShortcut psc)
        {
            shortcutPic = new PictureBox();
            shortcutPic.BackColor = Color.Transparent;
            shortcutPic.Location = new Point(25, 15);
            shortcutPic.Size = new Size(25, 25);
            shortcutPic.BackgroundImage =
                ThisCategory.LoadImageCache(psc); // Use the local icon cache for the file specified as the icon image
            shortcutPic.BackgroundImageLayout = ImageLayout.Stretch;
            shortcutPic.TabStop = false;
            shortcutPic.Click += (sender, e) => OpenFile(psc.Arguments, psc.FilePath, psc.WorkingDirectory);
            shortcutPic.Cursor = Cursors.Hand;
            shortcutPanel.Controls.Add(shortcutPic);
            shortcutPic.Show();
            shortcutPic.BringToFront();
            shortcutPic.MouseEnter += (sender, e) => shortcutPanel.BackColor = Color.Black;
            shortcutPic.MouseLeave += (sender, e) => shortcutPanel.BackColor = Color.Transparent;
        }

        // Click handler for shortcuts
        public void OpenFile(string arguments, string path, string workingDirec)
        {
            // starting program from psc panel click
            var proc = new ProcessStartInfo
            {
                Arguments = arguments,
                FileName = path,
                WorkingDirectory = workingDirec
            };

            /*
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = path;
            */

            try
            {
                Process.Start(proc);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Closes application upon deactivation
        private void frmMain_Deactivate(object sender, EventArgs e)
        {
            // closes program if user clicks outside form
            Close();
        }

        // Keyboard shortcut handlers
        private void frmMain_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                switch (e.KeyCode)
                {
                    case Keys.D1:
                        ControlList[0].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D2:
                        ControlList[1].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D3:
                        ControlList[2].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D4:
                        ControlList[3].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D5:
                        ControlList[4].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D6:
                        ControlList[5].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D7:
                        ControlList[6].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D8:
                        ControlList[7].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D9:
                        ControlList[8].ucShortcut_MouseEnter(sender, e);
                        break;

                    case Keys.D0:
                        ControlList[9].ucShortcut_MouseEnter(sender, e);
                        break;
                }
            }
            catch
            {
            }
        }

        private void frmMain_KeyUp(object sender, KeyEventArgs e)
        {
            //System.Diagnostics.Debugger.Launch();
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.Enter && ThisCategory.allowOpenAll)
                foreach (var usc in ControlList)
                    usc.ucShortcut_Click(sender, e);

            try
            {
                switch (e.KeyCode)
                {
                    case Keys.D1:
                        ControlList[0].ucShortcut_MouseLeave(sender, e);
                        ControlList[0].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D2:
                        ControlList[1].ucShortcut_MouseLeave(sender, e);
                        ControlList[1].ucShortcut_Click(sender, e);

                        break;

                    case Keys.D3:
                        ControlList[2].ucShortcut_MouseLeave(sender, e);
                        ControlList[2].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D4:
                        ControlList[3].ucShortcut_MouseLeave(sender, e);
                        ControlList[3].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D5:
                        ControlList[4].ucShortcut_MouseLeave(sender, e);
                        ControlList[4].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D6:
                        ControlList[5].ucShortcut_MouseLeave(sender, e);
                        ControlList[5].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D7:
                        ControlList[6].ucShortcut_MouseLeave(sender, e);
                        ControlList[6].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D8:
                        ControlList[7].ucShortcut_MouseLeave(sender, e);
                        ControlList[7].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D9:
                        ControlList[8].ucShortcut_MouseLeave(sender, e);
                        ControlList[8].ucShortcut_Click(sender, e);
                        break;

                    case Keys.D0:
                        ControlList[9].ucShortcut_MouseLeave(sender, e);
                        ControlList[9].ucShortcut_Click(sender, e);
                        break;
                }
            }
            catch
            {
            }
        }
        //
        // END OF CLASS
        //
    }
}