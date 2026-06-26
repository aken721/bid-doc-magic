using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace BidDocMagic
{
    public class ProgressForm : Form
    {
        private TextBox _statusLabel;
        private Panel _progressPanel;
        private Panel[] _blocks;
        private Button _actionButton;
        private Button _closeButton;
        private Timer _marqueeTimer;
        private int _marqueePos;
        private const int TotalBlocks = 20;
        private const int BlockSize = 14;
        private const int BlockGap = 4;
        private const int MarqueeLitCount = 5;

        public bool CancelRequested { get; private set; }
        public System.Threading.CancellationTokenSource CancellationTokenSource { get; set; }

        public ProgressForm()
        {
            CancelRequested = false;
            _marqueePos = 0;
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(420, 140);
            this.Text = "双层PDF转换";
            this.ShowInTaskbar = false;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("Microsoft YaHei UI", 9f);

            this.Shown += (s, e) =>
            {
                try
                {
                    var wordHwnd = (IntPtr)ThisAddIn.app.ActiveWindow.Hwnd;
                    if (wordHwnd != IntPtr.Zero)
                        NativeMethods.SetWindowLongPtr(this.Handle, NativeMethods.GWLP_HWNDPARENT, wordHwnd);
                }
                catch { }
            };

            _statusLabel = new TextBox
            {
                Location = new Point(20, 18),
                Size = new Size(380, 24),
                Text = "正在准备...",
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                Multiline = true,
                WordWrap = true,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                ShortcutsEnabled = true
            };
            this.Controls.Add(_statusLabel);

            _progressPanel = new Panel
            {
                Location = new Point(20, 52),
                Size = new Size(TotalBlocks * (BlockSize + BlockGap) - BlockGap, BlockSize)
            };

            _blocks = new Panel[TotalBlocks];
            for (int i = 0; i < TotalBlocks; i++)
            {
                _blocks[i] = new Panel
                {
                    Location = new Point(i * (BlockSize + BlockGap), 0),
                    Size = new Size(BlockSize, BlockSize),
                    BackColor = Color.FromArgb(230, 230, 230),
                    BorderStyle = BorderStyle.FixedSingle
                };
                _progressPanel.Controls.Add(_blocks[i]);
            }
            this.Controls.Add(_progressPanel);

            _actionButton = new Button
            {
                Location = new Point(110, 90),
                Size = new Size(90, 30),
                Text = "取消",
                FlatStyle = FlatStyle.System
            };
            _actionButton.Click += (s, e) =>
            {
                if (_actionButton.Tag is string filePath && File.Exists(filePath))
                {
                    try { Process.Start(filePath); } catch { }
                    this.Close();
                }
                else if (!CancelRequested)
                {
                    CancelRequested = true;
                    _actionButton.Enabled = false;
                    _statusLabel.Text = "正在取消...";
                    try { CancellationTokenSource?.Cancel(); } catch { }
                    _closeButton.Visible = true;
                }
            };
            this.Controls.Add(_actionButton);

            _closeButton = new Button
            {
                Location = new Point(220, 90),
                Size = new Size(90, 30),
                Text = "关闭",
                FlatStyle = FlatStyle.System,
                Visible = false
            };
            _closeButton.Click += (s, e) => this.Close();
            this.Controls.Add(_closeButton);

            _marqueeTimer = new Timer { Interval = 120 };
            _marqueeTimer.Tick += (s, e) =>
            {
                _marqueePos = (_marqueePos + 1) % TotalBlocks;
                for (int i = 0; i < TotalBlocks; i++)
                {
                    int dist = (i - _marqueePos + TotalBlocks) % TotalBlocks;
                    _blocks[i].BackColor = dist < MarqueeLitCount
                        ? Color.FromArgb(0, 120, 215)
                        : Color.FromArgb(230, 230, 230);
                }
            };
            _marqueeTimer.Start();

            this.FormClosed += (s, e) =>
            {
                _marqueeTimer.Stop();
                _marqueeTimer.Dispose();
            };
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Escape) return true;
            return base.ProcessDialogKey(keyData);
        }

        public void UpdateStatus(string text)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            try
            {
                if (this.InvokeRequired)
                    this.BeginInvoke(new Action(() => UpdateStatus(text)));
                else
                    _statusLabel.Text = text;
            }
            catch { }
        }

        public void ShowResult(string message, string filePath)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => ShowResult(message, filePath)));
                    return;
                }

                _marqueeTimer.Stop();
                _progressPanel.Visible = false;

                _statusLabel.Text = message;
                _statusLabel.Size = new Size(380, 40);

                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    _actionButton.Text = "打开文件";
                    _actionButton.Tag = filePath;
                    _actionButton.Enabled = true;
                }
                else
                {
                    _actionButton.Visible = false;
                }

                _closeButton.Visible = true;
            }
            catch { }
        }

        public void ShowError(string message)
        {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => ShowError(message)));
                    return;
                }

                _marqueeTimer.Stop();
                _progressPanel.Visible = false;

                _statusLabel.Text = message;
                _statusLabel.Size = new Size(380, 40);
                _statusLabel.ForeColor = Color.FromArgb(200, 50, 50);
                _statusLabel.BackColor = Color.FromArgb(255, 240, 240);

                _actionButton.Visible = false;
                _closeButton.Visible = true;
            }
            catch { }
        }

        public void CloseProgress()
        {
            if (this.IsDisposed) return;
            try
            {
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => CloseProgress()));
                else
                    this.Close();
            }
            catch { }
        }

        private static class NativeMethods
        {
            public const int GWLP_HWNDPARENT = -8;

            [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
            public static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

            [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLong")]
            public static extern IntPtr SetWindowLongPtr32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

            public static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
            {
                if (IntPtr.Size == 8)
                    return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
                else
                    return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
            }
        }
    }
}
