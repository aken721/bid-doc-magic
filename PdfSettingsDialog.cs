using System;
using System.Drawing;
using System.Windows.Forms;

namespace BidDocMagic
{
    public class PdfSettingsDialog : Form
    {
        private NumericUpDown numDpi;
        private CheckBox chkOpenAfterConvert;

        public int Dpi { get; private set; }
        public bool OpenAfterConvert { get; private set; }

        public int InitialDpi { get; set; } = 300;
        public bool InitialOpenAfterConvert { get; set; }

        public PdfSettingsDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "双层PDF设置";
            this.Size = new Size(520, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = new Font("微软雅黑", 9f);

            Panel scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                AutoScrollMargin = new Size(0, 20),
                Padding = new Padding(5)
            };

            int yPos = 10;

            // ===== DPI设置 =====
            var grpDpi = new GroupPanel("图片质量", 15, yPos, 470, 60);
            int dy = 22;

            var lblDpi = new Label { Text = "渲染DPI：", Location = new Point(20, dy + 3), AutoSize = true };
            numDpi = new NumericUpDown { Location = new Point(110, dy), Size = new Size(70, 25), Minimum = 150, Maximum = 1200, Value = 300, Increment = 50 };
            var lblDpiTip = new Label { Text = "推荐：300（快速）/ 600（标准）/ 1200（高清）", Location = new Point(195, dy + 3), AutoSize = true, ForeColor = Color.Gray };

            grpDpi.Controls.AddRange(new Control[] { lblDpi, numDpi, lblDpiTip });
            scrollPanel.Controls.Add(grpDpi);
            yPos += 70;

            // ===== 其他选项 =====
            var grpOptions = new GroupPanel("其他选项", 15, yPos, 470, 45);
            int oy = 15;

            chkOpenAfterConvert = new CheckBox { Text = "转换完成后自动打开PDF", Location = new Point(20, oy), AutoSize = true, Checked = false };

            grpOptions.Controls.Add(chkOpenAfterConvert);
            scrollPanel.Controls.Add(grpOptions);
            yPos += 55;

            // ===== 关于 =====
            var grpAbout = new GroupPanel("关于", 15, yPos, 470, 120);
            int ay = 15;

            var lblAppName = new Label { Text = "BidDocMagic", Location = new Point(20, ay), AutoSize = true, Font = new Font("微软雅黑", 12f, FontStyle.Bold) };
            ay += 28;

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var lblVersion = new Label { Text = $"版本：{version.Major}.{version.Minor}.{version.Build}", Location = new Point(20, ay), AutoSize = true };
            ay += 22;

            var lblCompany = new Label { Text = "软件开发：海南纽沃数字科技有限公司", Location = new Point(20, ay), AutoSize = true };
            ay += 22;

            var lblCopyright = new Label { Text = "Copyright © 海南纽沃数字科技有限公司 2026-present", Location = new Point(20, ay), AutoSize = true, ForeColor = Color.Gray };

            grpAbout.Controls.AddRange(new Control[] { lblAppName, lblVersion, lblCompany, lblCopyright });
            scrollPanel.Controls.Add(grpAbout);
            yPos += 130;

            // ===== 底部按钮面板 =====
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            var btnSave = new Button { Text = "保存", Location = new Point(280, 10), Size = new Size(90, 32), BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnApply_Click;

            var btnCancel = new Button { Text = "取消", Location = new Point(385, 10), Size = new Size(90, 32), FlatStyle = FlatStyle.Flat };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            buttonPanel.Controls.Add(btnSave);
            buttonPanel.Controls.Add(btnCancel);

            this.Controls.Add(buttonPanel);
            this.Controls.Add(scrollPanel);

            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;

            this.Load += (s, e) => LoadInitialValues();
        }

        private void LoadInitialValues()
        {
            numDpi.Value = InitialDpi;
            chkOpenAfterConvert.Checked = InitialOpenAfterConvert;
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            Dpi = (int)numDpi.Value;
            OpenAfterConvert = chkOpenAfterConvert.Checked;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }

    public class GroupPanel : GroupBox
    {
        public GroupPanel(string title, int x, int y, int width, int height)
        {
            this.Text = title;
            this.Location = new Point(x, y);
            this.Size = new Size(width, height);
        }
    }
}
