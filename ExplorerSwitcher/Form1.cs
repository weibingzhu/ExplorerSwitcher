using SHDocVw;
using Shell32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer; 

namespace ExplorerSwitcher
{
    internal class UnsafeNativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
    }

    public partial class Form1 : Form
    {
        private IWebBrowserApp _win;
        private Timer _timer;

        private string coolStr = "锁定该目录";
        private string selectRootpathStr = "请选择导出协议的路径";

        private string rootPath;

        public Form1()
        {
            InitializeComponent();
         

            var tmp = Path.GetTempFileName().Replace("\\", "/");
            File.Delete(tmp);
            Directory.CreateDirectory(tmp);

            var sh = new ShellClass();
            sh.ShellExecute(tmp);

            var sv = new ShellWindowsClass();
            do
            {
                foreach (IWebBrowserApp win in sv)
                {
                    if (0 == string.Compare(win.LocationURL, "file:///" + tmp, true))
                    {
                        _win = win;
                    }
                }
                if (_win == null) Thread.Sleep(500);
            }
            while (_win == null);

            _win.Navigate("c:\\");

            Directory.Delete(tmp);

            _timer = new Timer();
            _timer.Tick += _timer_Tick;
            _timer.Interval = 200;
            _timer.Start();

            this.Click += Form1_Click;
        }

        private void Form1_Click(object sender, EventArgs e)
        {
            ContextMenuStrip.Show(MousePosition);
        }

        private void _timer_Tick(object sender, EventArgs e)
        {
            Top = -_win.Top+100;
            Left = _win.Width-256;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            UnsafeNativeMethods.SetParent(Handle, new IntPtr(_win.HWND));
            Top = 0;
            Width = 64;
            BringToFront();
            BackgroundImage = Icon.ToBitmap();


            ContextMenuStrip = new ContextMenuStrip();
            foreach (var info in DriveInfo.GetDrives())
            {
                ContextMenuStrip.Items.Add(info.Name);
            }
            ContextMenuStrip.Items.Add(coolStr);

            ContextMenuStrip.ItemClicked += (o, args) =>
            {
                var item = (ToolStripMenuItem)args.ClickedItem;
                if (item.Text == coolStr)
                {
                    var fbd = new FolderBrowserDialog {Description = selectRootpathStr};
                    fbd.ShowDialog();
                    rootPath = fbd.SelectedPath;
                    item.ToolTipText = rootPath;
                }
                if (Directory.Exists(item.Text))
                {

                    _win.Navigate(item.Text);
                }

            };
        }
        
    }
}