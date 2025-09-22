using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;
using System.Text.Json;
using System.IO;

namespace WinFormsApp4
{
    // 修改：在 PiPSettings 中添加工具列狀態
    public class PiPSettings
    {
        public int WindowX { get; set; } = 100;
        public int WindowY { get; set; } = 100;
        public int WindowWidth { get; set; } = 480;
        public int WindowHeight { get; set; } = 320;
        public int CutX { get; set; } = 0;
        public int CutY { get; set; } = 0;
        public int CutWidth { get; set; } = 480;
        public int CutHeight { get; set; } = 270;
        public string LastSelectedWindow { get; set; } = "";
        public int WindowOpacity { get; set; } = 100;
        public bool ToolbarVisible { get; set; } = true; // 新增：工具列顯示狀態
    }

    public partial class PiPForm : Form
    {
        private PiPSettings settings;
        private string settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PiPTool",
            "settings.json"
        );
        // 新增：工具列隱藏/顯示相關變數
        private Button btnRereshWindowsList;
        private Button btnToggleToolbar;
        private bool isToolbarVisible = true;
        private int originalFormHeight;
        private int toolbarHeight;

        // 新增：直接儲存裁切區域的變數
        private int cutX = 0;
        private int cutY = 0;
        private int cutWidth = 480;
        private int cutHeight = 270;

        // 新增：透明度控制項
        private TrackBar trackBarOpacity;
        private Label lblOpacity;

        // 新增：區塊選取相關變數
        private Button btnToggleRegionSelect;
        private bool isRegionSelectMode = false;
        private bool isSelecting = false;
        private Point selectionStart;
        private Point selectionEnd;
        private Bitmap originalCapturedImage = null;
        private Timer saveDelayTimer;

        // WinAPI 宣告 (保持原有的)
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x20;
        private const int WS_EX_LAYERED = 0x80000;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        private Dictionary<string, IntPtr> windows = new Dictionary<string, IntPtr>();
        private IntPtr targetHwnd = IntPtr.Zero;
        private bool isClickThrough = false;

        private Panel pnlTitleBar;
        private ComboBox comboWindows;
        private Button btnToggleClickThrough;
        private PictureBox pictureBox1;
        private Timer captureTimer;

        public PiPForm()
        {
            LoadSettings();
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.Width = settings?.WindowWidth ?? 720;
            this.Height = settings?.WindowHeight ?? 480;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(settings?.WindowX ?? 100, settings?.WindowY ?? 100);
            this.BackColor = Color.LimeGreen;
            this.TransparencyKey = Color.LimeGreen;

            // 載入裁切設定
            if (settings != null)
            {
                cutX = settings.CutX;
                cutY = settings.CutY;
                cutWidth = settings.CutWidth;
                cutHeight = settings.CutHeight;
                this.Opacity = settings.WindowOpacity / 100.0;
            }

            // 記錄原始高度
            originalFormHeight = this.Height;

            // 增加標題列高度來容納所有控制項
            pnlTitleBar = new Panel()
            {
                Height = 60,
                Dock = DockStyle.Top,
                BackColor = Color.DimGray,
                Cursor = Cursors.SizeAll
            };
            pnlTitleBar.MouseDown += PnlTitleBar_MouseDown;
            this.Controls.Add(pnlTitleBar);

            // 記錄工具列高度
            toolbarHeight = pnlTitleBar.Height;

            // 第一行控制項
            Button btnClose = new Button()
            {
                Text = "✕",
                Font = new Font("Arial", 10, FontStyle.Bold),
                Width = 30,
                Height = 25,
                Left = pnlTitleBar.Width - 35,
                Top = 5,
                BackColor = Color.Red,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatAppearance.MouseOverBackColor = Color.DarkRed;
            btnClose.Click += BtnClose_Click;
            pnlTitleBar.Controls.Add(btnClose);
            
            // 新增：隱藏/顯示工具列按鈕（放在關閉按鈕旁邊）
            btnToggleToolbar = new Button()
            {
                Text = "▲", // 上箭頭表示隱藏
                Font = new Font("Arial", 8, FontStyle.Bold),
                Width = 25,
                Height = 25,
                Left = btnClose.Left - 30, // 放在關閉按鈕左邊
                Top = 5,
                BackColor = Color.Gray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand                
            };
            btnToggleToolbar.FlatAppearance.BorderSize = 0;
            btnToggleToolbar.FlatAppearance.MouseOverBackColor = Color.DarkGray;
            btnToggleToolbar.Click += BtnToggleToolbar_Click;
            pnlTitleBar.Controls.Add(btnToggleToolbar);
            btnRereshWindowsList = new Button()
            {
                Text = "⟳", // 上箭頭表示隱藏
                Font = new Font("Arial", 8, FontStyle.Bold),
                Width = 25,
                Height = 25,
                Left = btnToggleToolbar.Left - 30, // 放在關閉按鈕左邊
                Top = 5,
                BackColor = Color.Gray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnRereshWindowsList.FlatAppearance.BorderSize = 0;
            btnRereshWindowsList.FlatAppearance.MouseDownBackColor = Color.Aqua;
            btnRereshWindowsList.Click += RefreshWindowList_Click;
            pnlTitleBar.Controls.Add(btnRereshWindowsList);
            comboWindows = new ComboBox()
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Height = 25,
                Width = 140, // 再縮小一點給隱藏按鈕留空間
                Left = 5,
                Top = 5,
            };
            comboWindows.SelectedIndexChanged += ComboWindows_SelectedIndexChanged;
            pnlTitleBar.Controls.Add(comboWindows);

            btnToggleClickThrough = new Button()
            {
                Text = "切換穿透",
                Width = 70,
                Height = 25,
                Left = comboWindows.Right + 5,
                Top = 5,
                BackColor = Color.LightGray
            };
            btnToggleClickThrough.Click += BtnToggleClickThrough_Click;
            pnlTitleBar.Controls.Add(btnToggleClickThrough);

            btnToggleRegionSelect = new Button()
            {
                Text = "區塊選取",
                Width = 70,
                Height = 25,
                Left = btnToggleClickThrough.Right + 5,
                Top = 5,
                BackColor = Color.LightBlue
            };
            btnToggleRegionSelect.Click += BtnToggleRegionSelect_Click;
            pnlTitleBar.Controls.Add(btnToggleRegionSelect);

            // 第二行：透明度控制項
            Label lblOpacityText = new Label()
            {
                Text = "透明度:",
                Width = 50,
                Height = 20,
                Left = 5,
                Top = 35,
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlTitleBar.Controls.Add(lblOpacityText);

            trackBarOpacity = new TrackBar()
            {
                Minimum = 20,
                Maximum = 100,
                Value = settings?.WindowOpacity ?? 100,
                Width = 120,
                Height = 25,
                Left = lblOpacityText.Right + 5,
                Top = 32,
                TickFrequency = 10,
                SmallChange = 5,
                LargeChange = 10,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            trackBarOpacity.ValueChanged += TrackBarOpacity_ValueChanged;
            pnlTitleBar.Controls.Add(trackBarOpacity);

            lblOpacity = new Label()
            {
                Text = $"{trackBarOpacity.Value}%",
                Width = 40,
                Height = 20,
                Left = trackBarOpacity.Right + 5,
                Top = 35,
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            pnlTitleBar.Controls.Add(lblOpacity);

            Button btnResetOpacity = new Button()
            {
                Text = "重設",
                Width = 45,
                Height = 20,
                Left = lblOpacity.Right + 5,
                Top = 35,
                BackColor = Color.Gray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Arial", 8),
                Cursor = Cursors.Hand
            };
            btnResetOpacity.FlatAppearance.BorderSize = 0;
            btnResetOpacity.Click += BtnResetOpacity_Click;
            pnlTitleBar.Controls.Add(btnResetOpacity);

            pictureBox1 = new PictureBox()
            {
                BackColor = Color.Black,
                SizeMode = PictureBoxSizeMode.Zoom,
                Left = 2,
                Top = pnlTitleBar.Bottom + 2,
                Width = this.ClientSize.Width - 4,
                Height = this.ClientSize.Height - pnlTitleBar.Height - 4,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Top
            };

            // 添加滑鼠事件處理器
            pictureBox1.MouseClick += PictureBox1_MouseClick;
            pictureBox1.MouseDown += PictureBox1_MouseDown;
            pictureBox1.MouseMove += PictureBox1_MouseMove;
            pictureBox1.MouseUp += PictureBox1_MouseUp;
            pictureBox1.Paint += PictureBox1_Paint;

            this.Controls.Add(pictureBox1);
            pictureBox1.BringToFront();

            SetClickThrough(isClickThrough);

            captureTimer = new Timer();
            captureTimer.Interval = 33;
            captureTimer.Tick += CaptureTimer_Tick;
            captureTimer.Start();

            RefreshWindowList();
            ApplySettings();
        }
        // 新增：切換工具列顯示/隱藏的事件處理
        private void BtnToggleToolbar_Click(object sender, EventArgs e)
        {
            ToggleToolbar();
        }
        private void RefreshWindowList_Click(object sender, EventArgs e)
        {
            RefreshWindowList();
        }
        // 新增：切換工具列的主要方法
        private void ToggleToolbar()
        {
            isToolbarVisible = !isToolbarVisible;

            if (isToolbarVisible)
            {
                // 顯示工具列
                ShowToolbar();
            }
            else
            {
                // 隱藏工具列
                HideToolbar();
            }
        }
        // 新增：隱藏工具列
        private void HideToolbar()
        {
            // 隱藏工具列中的所有控制項，但保留隱藏按鈕
            foreach (Control control in pnlTitleBar.Controls)
            {
                if (control != btnToggleToolbar)
                {
                    control.Visible = false;
                }
            }

            // 改變按鈕文字和位置
            btnToggleToolbar.Text = "▼"; // 下箭頭表示顯示
            btnToggleToolbar.Left = 5; // 移到左邊
            btnToggleToolbar.Top = 2;
            btnToggleToolbar.Width = 30;

            // 縮小工具列高度
            pnlTitleBar.Height = 25;

            // 調整視窗高度（變小）
            int newHeight = originalFormHeight - (toolbarHeight - 25);
            this.Height = newHeight;

            // 調整 PictureBox 位置
            pictureBox1.Top = pnlTitleBar.Bottom + 2;
            pictureBox1.Height = this.ClientSize.Height - pnlTitleBar.Height - 4;

            // 移除拖拽功能（因為工具列太小）
            pnlTitleBar.Cursor = Cursors.Default;
        }

        // 新增：顯示工具列
        private void ShowToolbar()
        {
            // 顯示工具列中的所有控制項
            foreach (Control control in pnlTitleBar.Controls)
            {
                control.Visible = true;
            }

            // 恢復按鈕文字和位置
            btnToggleToolbar.Text = "▲"; // 上箭頭表示隱藏
            btnToggleToolbar.Left = pnlTitleBar.Width - 65; // 回到原位（關閉按鈕左邊）
            btnToggleToolbar.Top = 5;
            btnToggleToolbar.Width = 25;
            btnToggleToolbar.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // 恢復工具列高度
            pnlTitleBar.Height = toolbarHeight; // 60

            // 恢復視窗高度
            this.Height = originalFormHeight;

            // 調整 PictureBox 位置
            pictureBox1.Top = pnlTitleBar.Bottom + 2;
            pictureBox1.Height = this.ClientSize.Height - pnlTitleBar.Height - 4;

            // 恢復拖拽功能
            pnlTitleBar.Cursor = Cursors.SizeAll;
        }
        // 新增：透明度滑桿事件處理
        private void TrackBarOpacity_ValueChanged(object sender, EventArgs e)
        {
            // 更新視窗透明度 (TrackBar 值轉換為 0.0-1.0 範圍)
            this.Opacity = trackBarOpacity.Value / 100.0;

            // 更新標籤顯示
            lblOpacity.Text = $"{trackBarOpacity.Value}%";

            // 即時儲存設定
            if (settings != null)
            {
                settings.WindowOpacity = trackBarOpacity.Value;
                SaveSettingsDelayed();
            }
        }

        // 新增：重設透明度按鈕事件
        private void BtnResetOpacity_Click(object sender, EventArgs e)
        {
            trackBarOpacity.Value = 100;
            // ValueChanged 事件會自動觸發，更新透明度和標籤
        }

        // 新增：延遲儲存設定
        private void SaveSettingsDelayed()
        {
            // 使用計時器延遲儲存，避免拖拉時頻繁寫入
            if (saveDelayTimer == null)
            {
                saveDelayTimer = new Timer();
                saveDelayTimer.Interval = 500; // 500ms 後儲存
                saveDelayTimer.Tick += (s, e) =>
                {
                    saveDelayTimer.Stop();
                    SaveSettings();
                };
            }

            saveDelayTimer.Stop();
            saveDelayTimer.Start();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(settingsPath))
                {
                    string json = File.ReadAllText(settingsPath);
                    settings = JsonSerializer.Deserialize<PiPSettings>(json) ?? new PiPSettings();
                }
                else
                {
                    settings = new PiPSettings();
                }
            }
            catch
            {
                settings = new PiPSettings();
            }
        }

        private void SaveSettings()
        {
            try
            {
                if (settings == null)
                    settings = new PiPSettings();

                Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));

                settings.WindowX = this.Location.X;
                settings.WindowY = this.Location.Y;
                settings.WindowWidth = this.Width;
                settings.WindowHeight = this.Height;
                settings.CutX = cutX;
                settings.CutY = cutY;
                settings.CutWidth = cutWidth;
                settings.CutHeight = cutHeight;
                settings.LastSelectedWindow = comboWindows.SelectedItem?.ToString() ?? "";
                settings.WindowOpacity = trackBarOpacity?.Value ?? 100;
                settings.ToolbarVisible = isToolbarVisible; // 儲存工具列狀態

                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(settingsPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"無法儲存設定: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ApplySettings()
        {
            if (settings == null) return;

            Rectangle screenBounds = Screen.PrimaryScreen.WorkingArea;

            int x = Math.Max(0, Math.Min(settings.WindowX, screenBounds.Width - settings.WindowWidth));
            int y = Math.Max(0, Math.Min(settings.WindowY, screenBounds.Height - settings.WindowHeight));

            this.Location = new Point(x, y);
            this.Width = Math.Max(300, Math.Min(settings.WindowWidth, screenBounds.Width));
            this.Height = Math.Max(200, Math.Min(settings.WindowHeight, screenBounds.Height));

            cutX = settings.CutX;
            cutY = settings.CutY;
            cutWidth = settings.CutWidth;
            cutHeight = settings.CutHeight;

            int opacity = Math.Max(20, Math.Min(100, settings.WindowOpacity));
            this.Opacity = opacity / 100.0;
            if (trackBarOpacity != null)
            {
                trackBarOpacity.Value = opacity;
                lblOpacity.Text = $"{opacity}%";
            }

            // 套用工具列顯示狀態
            isToolbarVisible = settings.ToolbarVisible;
            if (!isToolbarVisible)
            {
                // 延遲隱藏工具列，確保所有控制項都已初始化
                this.BeginInvoke(new Action(() => HideToolbar()));
            }

            if (!string.IsNullOrEmpty(settings.LastSelectedWindow))
            {
                this.BeginInvoke(new Action(() =>
                {
                    for (int i = 0; i < comboWindows.Items.Count; i++)
                    {
                        if (comboWindows.Items[i].ToString() == settings.LastSelectedWindow)
                        {
                            comboWindows.SelectedIndex = i;
                            break;
                        }
                    }
                }));
            }
        }

        // 修改：ProcessCmdKey 新增透明度快捷鍵
        // 修改：ProcessCmdKey 新增工具列切換快捷鍵
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Shift | Keys.T))
            {
                if (isClickThrough)
                {
                    isClickThrough = false;
                    SetClickThrough(false);

                    // 取消穿透時自動顯示工具列
                    if (!isToolbarVisible)
                    {
                        ShowToolbar();
                    }
                }
                else
                {
                    isClickThrough = true;
                    SetClickThrough(true);

                    // 啟用穿透時自動隱藏工具列
                    if (isToolbarVisible)
                    {
                        HideToolbar();
                    }
                }
                return true;
            }
            // 新增：工具列切換快捷鍵
            else if (keyData == (Keys.Control | Keys.Shift | Keys.H))
            {
                ToggleToolbar();
                return true;
            }
            // 透明度快捷鍵
            else if (keyData == (Keys.Control | Keys.Add) || keyData == (Keys.Control | Keys.Oemplus))
            {
                if (trackBarOpacity.Value < trackBarOpacity.Maximum)
                {
                    trackBarOpacity.Value = Math.Min(trackBarOpacity.Maximum, trackBarOpacity.Value + 10);
                }
                return true;
            }
            else if (keyData == (Keys.Control | Keys.Subtract) || keyData == (Keys.Control | Keys.OemMinus))
            {
                if (trackBarOpacity.Value > trackBarOpacity.Minimum)
                {
                    trackBarOpacity.Value = Math.Max(trackBarOpacity.Minimum, trackBarOpacity.Value - 10);
                }
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // 修改：計算選取區域並更新裁切變數
        private void CalculateSelectionRegion()
        {
            if (pictureBox1.Image == null || originalCapturedImage == null) return;

            var pictureBoxSize = pictureBox1.Size;
            var imageSize = originalCapturedImage.Size;

            // 計算縮放比例和偏移
            float scaleX = (float)pictureBoxSize.Width / imageSize.Width;
            float scaleY = (float)pictureBoxSize.Height / imageSize.Height;
            float scale = Math.Min(scaleX, scaleY);

            int displayWidth = (int)(imageSize.Width * scale);
            int displayHeight = (int)(imageSize.Height * scale);
            int offsetX = (pictureBoxSize.Width - displayWidth) / 2;
            int offsetY = (pictureBoxSize.Height - displayHeight) / 2;

            // 轉換選取座標為圖片座標
            int startX = Math.Min(selectionStart.X, selectionEnd.X) - offsetX;
            int startY = Math.Min(selectionStart.Y, selectionEnd.Y) - offsetY;
            int endX = Math.Max(selectionStart.X, selectionEnd.X) - offsetX;
            int endY = Math.Max(selectionStart.Y, selectionEnd.Y) - offsetY;

            // 限制在圖片範圍內
            startX = Math.Max(0, Math.Min(startX, displayWidth));
            startY = Math.Max(0, Math.Min(startY, displayHeight));
            endX = Math.Max(0, Math.Min(endX, displayWidth));
            endY = Math.Max(0, Math.Min(endY, displayHeight));

            // 轉換為原始完整視窗的座標
            int imageStartX = (int)(startX / scale);
            int imageStartY = (int)(startY / scale);
            int imageEndX = (int)(endX / scale);
            int imageEndY = (int)(endY / scale);

            // 計算選取區域的寬高
            int selectionWidth = imageEndX - imageStartX;
            int selectionHeight = imageEndY - imageStartY;

            // 確保選取區域至少有 1 像素
            if (selectionWidth < 1) selectionWidth = 1;
            if (selectionHeight < 1) selectionHeight = 1;

            // 更新裁切變數（不使用數值控制項）
            cutX = Math.Max(0, imageStartX);
            cutY = Math.Max(0, imageStartY);
            cutWidth = selectionWidth;
            cutHeight = selectionHeight;

            // 顯示選取結果
            MessageBox.Show($"已選取區域:\nX: {cutX}\nY: {cutY}\nWidth: {cutWidth}\nHeight: {cutHeight}\n\n" +
                          $"原始視窗尺寸: {imageSize.Width} x {imageSize.Height}",
                          "區塊選取完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 修改：CaptureTimer_Tick 使用裁切變數而非數值控制項
        private void CaptureTimer_Tick(object sender, EventArgs e)
        {
            if (targetHwnd == IntPtr.Zero) return;
            if (isRegionSelectMode) return;

            if (GetWindowRect(targetHwnd, out RECT rect))
            {
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                if (width <= 0 || height <= 0) return;

                try
                {
                    using Bitmap fullBmp = new Bitmap(width, height);
                    using Graphics g = Graphics.FromImage(fullBmp);
                    const uint PW_RENDERFULLCONTENT = 0x00000002;
                    bool success = PrintWindow(targetHwnd, g.GetHdc(), PW_RENDERFULLCONTENT);
                    g.ReleaseHdc();

                    if (success)
                    {
                        // 使用裁切變數而非數值控制項
                        int currentCutX = Math.Min(cutX, width - 1);
                        int currentCutY = Math.Min(cutY, height - 1);
                        int currentCutW = Math.Min(cutWidth, width - currentCutX);
                        int currentCutH = Math.Min(cutHeight, height - currentCutY);

                        Rectangle cutRect = new Rectangle(currentCutX, currentCutY, currentCutW, currentCutH);
                        Bitmap croppedBmp = fullBmp.Clone(cutRect, fullBmp.PixelFormat);

                        pictureBox1.Image?.Dispose();
                        pictureBox1.Image = croppedBmp;
                    }
                }
                catch
                {
                    // 擷取或裁切錯誤忽略或記錄
                }
            }
        }

        // 修改：點擊座標顯示功能使用裁切變數
        private void PictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (isRegionSelectMode) return;
            if (pictureBox1.Image == null) return;

            var pictureBoxSize = pictureBox1.Size;
            var imageSize = pictureBox1.Image.Size;

            float scaleX = (float)pictureBoxSize.Width / imageSize.Width;
            float scaleY = (float)pictureBoxSize.Height / imageSize.Height;
            float scale = Math.Min(scaleX, scaleY);

            int displayWidth = (int)(imageSize.Width * scale);
            int displayHeight = (int)(imageSize.Height * scale);
            int offsetX = (pictureBoxSize.Width - displayWidth) / 2;
            int offsetY = (pictureBoxSize.Height - displayHeight) / 2;

            int relativeX = e.X - offsetX;
            int relativeY = e.Y - offsetY;

            if (relativeX >= 0 && relativeY >= 0 && relativeX < displayWidth && relativeY < displayHeight)
            {
                int imageX = (int)(relativeX / scale);
                int imageY = (int)(relativeY / scale);
                int originalX = imageX + cutX; // 使用裁切變數
                int originalY = imageY + cutY; // 使用裁切變數

                //MessageBox.Show($"點擊位置: ({originalX}, {originalY})", "座標", MessageBoxButtons.OK);
            }
            else
            {
                //MessageBox.Show("點擊在圖片範圍外", "提醒", MessageBoxButtons.OK);
            }
        }

        // 保持其他方法不變...
        private void BtnToggleRegionSelect_Click(object sender, EventArgs e)
        {
            isRegionSelectMode = !isRegionSelectMode;

            if (isRegionSelectMode)
            {
                btnToggleRegionSelect.BackColor = Color.Orange;
                btnToggleRegionSelect.Text = "取消選取";
                captureTimer.Stop();
                CaptureFullWindowForSelection();
                pictureBox1.Cursor = Cursors.Cross;
            }
            else
            {
                btnToggleRegionSelect.BackColor = Color.LightBlue;
                btnToggleRegionSelect.Text = "區塊選取";
                pictureBox1.Cursor = Cursors.Default;
                captureTimer.Start();
                originalCapturedImage?.Dispose();
                originalCapturedImage = null;
            }
        }

        private void CaptureFullWindowForSelection()
        {
            if (targetHwnd == IntPtr.Zero) return;

            if (GetWindowRect(targetHwnd, out RECT rect))
            {
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                if (width <= 0 || height <= 0) return;

                try
                {
                    using Bitmap fullBmp = new Bitmap(width, height);
                    using Graphics g = Graphics.FromImage(fullBmp);
                    const uint PW_RENDERFULLCONTENT = 0x00000002;
                    bool success = PrintWindow(targetHwnd, g.GetHdc(), PW_RENDERFULLCONTENT);
                    g.ReleaseHdc();

                    if (success)
                    {
                        originalCapturedImage?.Dispose();
                        originalCapturedImage = new Bitmap(fullBmp);
                        pictureBox1.Image?.Dispose();
                        pictureBox1.Image = new Bitmap(fullBmp);

                        MessageBox.Show($"已載入完整視窗畫面 ({width} x {height})\n請拖拉滑鼠選取想要擷取的區域",
                                      "區塊選取模式", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("無法擷取完整視窗畫面", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        isRegionSelectMode = false;
                        btnToggleRegionSelect.BackColor = Color.LightBlue;
                        btnToggleRegionSelect.Text = "區塊選取";
                        pictureBox1.Cursor = Cursors.Default;
                        captureTimer.Start();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"擷取視窗時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    isRegionSelectMode = false;
                    btnToggleRegionSelect.BackColor = Color.LightBlue;
                    btnToggleRegionSelect.Text = "區塊選取";
                    pictureBox1.Cursor = Cursors.Default;
                    captureTimer.Start();
                }
            }
        }

        private void PictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (isRegionSelectMode && e.Button == MouseButtons.Left)
            {
                isSelecting = true;
                selectionStart = e.Location;
                selectionEnd = e.Location;
            }
        }

        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isRegionSelectMode && isSelecting && e.Button == MouseButtons.Left)
            {
                selectionEnd = e.Location;
                pictureBox1.Invalidate();
            }
        }

        private void PictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (isRegionSelectMode && isSelecting && e.Button == MouseButtons.Left)
            {
                isSelecting = false;
                selectionEnd = e.Location;

                int minX = Math.Min(selectionStart.X, selectionEnd.X);
                int minY = Math.Min(selectionStart.Y, selectionEnd.Y);
                int maxX = Math.Max(selectionStart.X, selectionEnd.X);
                int maxY = Math.Max(selectionStart.Y, selectionEnd.Y);

                if (Math.Abs(maxX - minX) < 5 || Math.Abs(maxY - minY) < 5)
                {
                    pictureBox1.Invalidate();
                    return;
                }

                CalculateSelectionRegion();

                isRegionSelectMode = false;
                btnToggleRegionSelect.BackColor = Color.LightBlue;
                btnToggleRegionSelect.Text = "區塊選取";
                pictureBox1.Cursor = Cursors.Default;
                captureTimer.Start();
                pictureBox1.Invalidate();
                originalCapturedImage?.Dispose();
                originalCapturedImage = null;
            }
        }

        private void PictureBox1_Paint(object sender, PaintEventArgs e)
        {
            if (isRegionSelectMode && isSelecting)
            {
                Rectangle selectRect = new Rectangle(
                    Math.Min(selectionStart.X, selectionEnd.X),
                    Math.Min(selectionStart.Y, selectionEnd.Y),
                    Math.Abs(selectionStart.X - selectionEnd.X),
                    Math.Abs(selectionStart.Y - selectionEnd.Y)
                );

                using (Pen pen = new Pen(Color.Red, 2))
                {
                    e.Graphics.DrawRectangle(pen, selectRect);
                }

                using (Brush brush = new SolidBrush(Color.FromArgb(50, 255, 0, 0)))
                {
                    e.Graphics.FillRectangle(brush, selectRect);
                }
            }
        }

        private void PnlTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void RefreshWindowList()
        {
            comboWindows.Items.Clear();
            windows.Clear();

            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    StringBuilder title = new StringBuilder(256);
                    GetWindowText(hWnd, title, title.Capacity);
                    string windowTitle = title.ToString();
                    if (!string.IsNullOrEmpty(windowTitle))
                    {
                        var lowerWindowTitle = windowTitle.ToLower();
                        if (!windows.ContainsKey(windowTitle) && 
                        (lowerWindowTitle.Contains("brave") || 
                        lowerWindowTitle.Contains("chrome") || 
                        lowerWindowTitle.Contains("edge") ||
                        lowerWindowTitle.Contains("firefox")))
                        {
                            windows[windowTitle] = hWnd;
                            comboWindows.Items.Add(windowTitle);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            if (comboWindows.Items.Count > 0)
                comboWindows.SelectedIndex = 0;
        }

        private void ComboWindows_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboWindows.SelectedItem is string selectedTitle
                && windows.TryGetValue(selectedTitle, out IntPtr hwnd))
            {
                targetHwnd = hwnd;
            }
        }

        // 修改 BtnToggleClickThrough_Click 方法
        private void BtnToggleClickThrough_Click(object sender, EventArgs e)
        {
            isClickThrough = !isClickThrough;
            SetClickThrough(isClickThrough);

            // 新增：自動切換工具列狀態
            if (isClickThrough)
            {
                // 啟用穿透時自動隱藏工具列
                if (isToolbarVisible)
                {
                    HideToolbar();
                }
            }
            else
            {
                // 取消穿透時自動顯示工具列
                if (!isToolbarVisible)
                {
                    ShowToolbar();
                }
            }
        }

        private void SetClickThrough(bool enabled)
        {
            int exStyle = GetWindowLong(this.Handle, GWL_EXSTYLE);
            if (enabled)
                exStyle |= WS_EX_TRANSPARENT | WS_EX_LAYERED;
            else
                exStyle = (exStyle & ~WS_EX_TRANSPARENT) | WS_EX_LAYERED;
            SetWindowLong(this.Handle, GWL_EXSTYLE, exStyle);

            // 更新按鈕顯示狀態
            UpdateClickThroughButtonText(enabled);
        }
        // 新增：更新穿透按鈕文字的方法
        private void UpdateClickThroughButtonText(bool enabled)
        {
            if (btnToggleClickThrough != null)
            {
                btnToggleClickThrough.Text = enabled ? "取消穿透" : "切換穿透";
                btnToggleClickThrough.BackColor = enabled ? Color.Orange : Color.LightGray;
            }
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_NCHITTEST = 0x0084;
            const int HTCLIENT = 1;
            const int HTLEFT = 10;
            const int HTRIGHT = 11;
            const int HTTOP = 12;
            const int HTTOPLEFT = 13;
            const int HTTOPRIGHT = 14;
            const int HTBOTTOM = 15;
            const int HTBOTTOMLEFT = 16;
            const int HTBOTTOMRIGHT = 17;
            const int RESIZE_HANDLE_SIZE = 10;

            if (m.Msg == WM_NCHITTEST)
            {
                base.WndProc(ref m);

                Point cursor = PointToClient(Cursor.Position);
                if (cursor.X <= RESIZE_HANDLE_SIZE && cursor.Y <= RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTTOPLEFT;
                else if (cursor.X >= ClientSize.Width - RESIZE_HANDLE_SIZE && cursor.Y <= RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTTOPRIGHT;
                else if (cursor.X <= RESIZE_HANDLE_SIZE && cursor.Y >= ClientSize.Height - RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTBOTTOMLEFT;
                else if (cursor.X >= ClientSize.Width - RESIZE_HANDLE_SIZE && cursor.Y >= ClientSize.Height - RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTBOTTOMRIGHT;
                else if (cursor.X <= RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTLEFT;
                else if (cursor.X >= ClientSize.Width - RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTRIGHT;
                else if (cursor.Y <= RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTTOP;
                else if (cursor.Y >= ClientSize.Height - RESIZE_HANDLE_SIZE)
                    m.Result = (IntPtr)HTBOTTOM;
                else
                    m.Result = (IntPtr)HTCLIENT;
                return;
            }
            base.WndProc(ref m);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            SaveSettings();
            base.OnFormClosed(e);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            SaveSettings();
            this.Close();
        }
    }
}
