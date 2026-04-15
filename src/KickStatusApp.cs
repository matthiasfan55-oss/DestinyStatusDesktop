using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace KickStatusApp
{
    public class PriorityRule
    {
        public List<string> channels { get; set; }
        public string displayMode { get; set; }
        public string displayValue { get; set; }

        public PriorityRule()
        {
            channels = new List<string>();
            displayMode = "Text";
            displayValue = "";
        }
    }

    public class AppConfig
    {
        public List<PriorityRule> priorityRules { get; set; }
        public bool openKickOnClick { get; set; }
        public bool openBigscreenInChromeOnClick { get; set; }
        public double refreshMinutes { get; set; }

        public string primaryChannel { get; set; }
        public string secondaryChannel { get; set; }
        public List<string> priorityChannels { get; set; }
        public string priorityGifPath { get; set; }

        public AppConfig()
        {
            priorityRules = CreateDefaultRules();
            openKickOnClick = true;
            openBigscreenInChromeOnClick = false;
            refreshMinutes = 1.0;

            primaryChannel = "destiny";
            secondaryChannel = "anythingelse";
            priorityChannels = new List<string>();
            priorityGifPath = "";
        }

        public static List<PriorityRule> CreateDefaultRules()
        {
            return new List<PriorityRule>
            {
                new PriorityRule
                {
                    channels = new List<string> { "destiny" },
                    displayMode = "Text",
                    displayValue = "D"
                },
                new PriorityRule
                {
                    channels = new List<string> { "anythingelse" },
                    displayMode = "Text",
                    displayValue = "AE"
                },
                new PriorityRule
                {
                    channels = new List<string>(),
                    displayMode = "Gif",
                    displayValue = ""
                }
            };
        }
    }

    public class WindowState
    {
        public int x { get; set; }
        public int y { get; set; }
        public double scale { get; set; }
    }

    public class StateFile
    {
        public Dictionary<string, WindowState> layouts { get; set; }

        public StateFile()
        {
            layouts = new Dictionary<string, WindowState>();
        }
    }

    public class AppUpdateManifest
    {
        public string version { get; set; }
        public string packageUrl { get; set; }
        public string notes { get; set; }
    }

    public class GitHubContentResponse
    {
        public string content { get; set; }
        public string encoding { get; set; }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern int SetCurrentProcessExplicitAppUserModelID(string appID);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
    }

    internal class CheckCandidate
    {
        public int RuleIndex { get; set; }
        public string ChannelName { get; set; }
    }

    public class StatusForm : Form
    {
        private const string UpdateManifestApiUrl = "https://api.github.com/repos/matthiasfan55-oss/DestinyStatusDesktop/contents/update.json?ref=main";
        private const string UpdateManifestRawUrl = "https://raw.githubusercontent.com/matthiasfan55-oss/DestinyStatusDesktop/main/update.json";
        private const string UpdatePackageApiUrl = "https://api.github.com/repos/matthiasfan55-oss/DestinyStatusDesktop/contents/dist/KickStatusSquare-package.zip?ref=main";
        private const string GitHubUserAgent = "KickStatusSquare-Updater";

        private readonly string appDir;
        private readonly string configPath;
        private readonly string statePath;
        private readonly string settingsScriptPath;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();

        private AppConfig config;
        private StateFile state;

        private readonly Panel statusPanel;
        private readonly PictureBox gifBox;
        private readonly Label markerLabel;
        private readonly SpinnerControl spinnerControl;
        private readonly WebBrowser browser;
        private readonly Form browserHost;
        private readonly Timer settleTimer;
        private readonly Timer watchdogTimer;
        private readonly Timer refreshTimer;
        private readonly ContextMenuStrip menu;

        private bool isChecking;
        private DateTime loadStartedAt;
        private string currentCheckName;
        private string currentCheckUrl;
        private string lastOpenUrl;
        private string lastOpenChannelName;
        private int lastOpenRuleIndex = -1;
        private string currentMarker = "";
        private string currentGifPath = "";
        private int currentCheckQueueIndex;
        private int currentStateRetryCount;
        private bool hasResolvedVisualState;
        private bool spinnerInitialMode;
        private List<CheckCandidate> checkQueue = new List<CheckCandidate>();
        private Image gifImage;

        private bool leftMouseDown;
        private bool dragStarted;
        private Point mouseDownPoint;
        private double currentScale = 1.0;
        private int updateCheckQueued;
        private bool updateInProgress;

        public StatusForm()
        {
            appDir = AppDomain.CurrentDomain.BaseDirectory;
            string dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KickStatusAppData");
            Directory.CreateDirectory(dataDir);
            configPath = Path.Combine(dataDir, "config.json");
            statePath = Path.Combine(dataDir, "state.json");
            settingsScriptPath = Path.Combine(appDir, "EditKickStatusConfig.ps1");

            config = LoadConfig();
            state = LoadState();

            InitializeFallbackOpenTarget();

            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ApplyScale(1.0);
            ShowInTaskbar = true;
            ShowIcon = false;
            Text = "Kick Status";
            BackColor = Color.FromArgb(24, 24, 24);
            TopMost = true;

            statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Gray,
                Cursor = Cursors.Hand,
                Margin = Padding.Empty,
                BackgroundImageLayout = ImageLayout.Zoom
            };
            Controls.Add(statusPanel);

            gifBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                SizeMode = PictureBoxSizeMode.Zoom,
                Visible = false
            };
            statusPanel.Controls.Add(gifBox);

            markerLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 20, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                Text = ""
            };
            markerLabel.Paint += MarkerLabel_Paint;
            statusPanel.Controls.Add(markerLabel);
            markerLabel.BringToFront();

            spinnerControl = new SpinnerControl
            {
                Size = new Size(16, 16),
                BackColor = Color.Transparent,
                Visible = false
            };
            statusPanel.Controls.Add(spinnerControl);
            PositionSpinner();
            HookPointer(spinnerControl);
            ApplyScale(currentScale);
            SetInitialLoadingState();

            browserHost = new Form
            {
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                Opacity = 0,
                Size = new Size(1, 1),
                Location = new Point(-2000, -2000)
            };

            browser = new WebBrowser
            {
                ScriptErrorsSuppressed = true,
                ScrollBarsEnabled = false,
                Dock = DockStyle.Fill
            };
            browser.DocumentCompleted += Browser_DocumentCompleted;
            browserHost.Controls.Add(browser);

            settleTimer = new Timer { Interval = 1500 };
            settleTimer.Tick += delegate
            {
                settleTimer.Stop();
                CompleteStatusCheck();
            };

            watchdogTimer = new Timer { Interval = 1000 };
            watchdogTimer.Tick += delegate
            {
                if (isChecking && (DateTime.Now - loadStartedAt).TotalSeconds >= 20)
                {
                    isChecking = false;
                    settleTimer.Stop();
                    SetOfflineState();
                }
            };
            watchdogTimer.Start();

            refreshTimer = new Timer { Interval = GetRefreshIntervalMilliseconds(config) };
            refreshTimer.Tick += delegate { StartStatusCheck(); };
            refreshTimer.Stop();

            menu = new ContextMenuStrip();
            var settingsItem = new ToolStripMenuItem("Settings");
            settingsItem.Click += delegate
            {
                if (File.Exists(settingsScriptPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + settingsScriptPath + "\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                }
            };
            menu.Items.Add(settingsItem);

            var updateItem = new ToolStripMenuItem("Check for Updates");
            updateItem.Click += delegate { BeginUpdateCheck(true); };
            menu.Items.Add(updateItem);

            var sizeItem = new ToolStripMenuItem("Size");
            var smallItem = new ToolStripMenuItem("Small");
            smallItem.Click += delegate { ApplyScale(0.3); };
            var mediumItem = new ToolStripMenuItem("Medium");
            mediumItem.Click += delegate { ApplyScale(1.0); };
            var largeItem = new ToolStripMenuItem("Large");
            largeItem.Click += delegate { ApplyScale(2.0); };
            sizeItem.DropDownItems.Add(smallItem);
            sizeItem.DropDownItems.Add(mediumItem);
            sizeItem.DropDownItems.Add(largeItem);
            menu.Items.Add(sizeItem);

            menu.Items.Add(new ToolStripSeparator());

            var closeItem = new ToolStripMenuItem("Close App");
            closeItem.Font = new Font(closeItem.Font, FontStyle.Bold);
            closeItem.BackColor = Color.FromArgb(238, 238, 238);
            closeItem.ForeColor = Color.FromArgb(70, 70, 70);
            closeItem.Click += delegate { Close(); };
            menu.Items.Add(closeItem);

            HookPointer(statusPanel);
            HookPointer(markerLabel);
            HookPointer(gifBox);

            var saved = GetSavedLocation();
            if (saved != null)
            {
                if (saved.scale > 0)
                {
                    ApplyScale(saved.scale);
                }
                var restored = new Point(saved.x, saved.y);
                if (IsLocationVisible(restored))
                {
                    Location = restored;
                }
                else
                {
                    CenterToVisibleScreen();
                }
            }
            else
            {
                CenterToVisibleScreen();
            }

            Move += delegate { SaveWindowState(); };
            FormClosing += delegate { SaveWindowState(); };
            SizeChanged += delegate { UpdateRoundedRegion(); PositionSpinner(); };
            Shown += delegate
            {
                browserHost.Show();
                StartStatusCheck();
                BeginUpdateCheck(false);
            };
        }

        private void HookPointer(Control control)
        {
            control.MouseDown += Control_MouseDown;
            control.MouseMove += Control_MouseMove;
            control.MouseUp += Control_MouseUp;
        }

        private void Control_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                leftMouseDown = true;
                dragStarted = false;
                mouseDownPoint = e.Location;
            }
            else if (e.Button == MouseButtons.Right)
            {
                var control = sender as Control;
                menu.Show(control ?? this, e.Location);
            }
        }

        private void Control_MouseMove(object sender, MouseEventArgs e)
        {
            if (!leftMouseDown)
            {
                return;
            }

            int dx = Math.Abs(e.X - mouseDownPoint.X);
            int dy = Math.Abs(e.Y - mouseDownPoint.Y);
            if (!dragStarted && (dx >= 4 || dy >= 4))
            {
                dragStarted = true;
                NativeMethods.ReleaseCapture();
                NativeMethods.SendMessage(Handle, 0xA1, 0x2, 0);
            }
        }

        private void Control_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            bool wasDrag = dragStarted;
            leftMouseDown = false;
            dragStarted = false;

            if (!wasDrag)
            {
                OpenClickTargets();
            }
        }

        private void OpenClickTargets()
        {
            config = LoadConfig();

            if (config.openKickOnClick && !string.IsNullOrWhiteSpace(lastOpenUrl))
            {
                try
                {
                    Process.Start(lastOpenUrl);
                }
                catch
                {
                }
            }

            if (config.openBigscreenInChromeOnClick)
            {
                string bigscreenUrl = "https://www.destiny.gg/bigscreen";
                if (lastOpenRuleIndex >= 2 && !string.IsNullOrWhiteSpace(lastOpenChannelName))
                {
                    bigscreenUrl = "https://www.destiny.gg/bigscreen#kick/" + lastOpenChannelName;
                }

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "chrome.exe",
                        Arguments = "--new-tab " + bigscreenUrl,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                }
                catch
                {
                    try
                    {
                        Process.Start(bigscreenUrl);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void Browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (!isChecking)
            {
                return;
            }

            if (browser.ReadyState == WebBrowserReadyState.Complete)
            {
                settleTimer.Stop();
                settleTimer.Start();
            }
        }

        private void StartStatusCheck()
        {
            if (isChecking)
            {
                return;
            }

            refreshTimer.Stop();
            config = LoadConfig();
            checkQueue = BuildCheckQueue(config);
            currentCheckQueueIndex = 0;

            if (checkQueue.Count == 0)
            {
                SetOfflineState();
                return;
            }

            isChecking = true;
            BeginCheckAt(currentCheckQueueIndex);
        }

        private void BeginCheckAt(int index)
        {
            if (index < 0 || index >= checkQueue.Count)
            {
                SetOfflineState();
                isChecking = false;
                return;
            }

            currentCheckQueueIndex = index;
            currentStateRetryCount = 0;
            currentCheckName = checkQueue[index].ChannelName;
            currentCheckUrl = BuildKickUrl(currentCheckName);
            loadStartedAt = DateTime.Now;
            settleTimer.Interval = 1500;
            SetStatusUi("Checking", "", null);

            try
            {
                browser.Navigate(currentCheckUrl);
            }
            catch
            {
                isChecking = false;
                SetOfflineState();
            }
        }

        private void ApplyScale(double scale)
        {
            currentScale = scale;
            int side = Math.Max(20, (int)Math.Round(64 * scale));
            ClientSize = new Size(side, side);
            MinimumSize = new Size(side, side);
            MaximumSize = new Size(side, side);
            if (markerLabel != null)
            {
                int fontSize = Math.Max(7, (int)Math.Round(20 * scale));
                markerLabel.Font = new Font("Segoe UI", fontSize, FontStyle.Bold);
            }
            if (spinnerControl != null)
            {
                UpdateSpinnerLayout();
            }
            UpdateRoundedRegion();
        }

        private void MarkerLabel_Paint(object sender, PaintEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(currentMarker))
            {
                return;
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            using (var path = new GraphicsPath())
            {
                float emSize = e.Graphics.DpiY * markerLabel.Font.Size / 72f;
                var rect = markerLabel.ClientRectangle;
                var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };

                path.AddString(
                    currentMarker,
                    markerLabel.Font.FontFamily,
                    (int)markerLabel.Font.Style,
                    emSize,
                    rect,
                    format);

                using (var pen = new Pen(Color.Black, Math.Max(1.5f, markerLabel.Font.Size / 10f)) { LineJoin = LineJoin.Round })
                {
                    e.Graphics.DrawPath(pen, path);
                }

                using (var brush = new SolidBrush(markerLabel.ForeColor))
                {
                    e.Graphics.FillPath(brush, path);
                }
            }
        }

        private void UpdateRoundedRegion()
        {
            int radius = Math.Max(8, (int)Math.Round(12 * currentScale));
            IntPtr hrgn = NativeMethods.CreateRoundRectRgn(0, 0, Width + 1, Height + 1, radius, radius);
            Region = Region.FromHrgn(hrgn);
        }

        private void CompleteStatusCheck()
        {
            bool continueChecking = false;

            try
            {
                if (browser.Document == null || browser.Document.Body == null)
                {
                    continueChecking = TryAdvanceToNextCandidate();
                    if (!continueChecking)
                    {
                        SetOfflineState();
                    }
                    return;
                }

                string text = browser.Document.Body.InnerText ?? "";
                string normalized = NormalizeWhitespace(text);
                string channelState = GetChannelState(normalized, currentCheckName);

                if (channelState == "Unknown")
                {
                    if (currentStateRetryCount < 2)
                    {
                        currentStateRetryCount++;
                        continueChecking = true;
                        BeginInvoke((Action)delegate
                        {
                            settleTimer.Stop();
                            settleTimer.Interval = 2000;
                            settleTimer.Start();
                        });
                        return;
                    }

                    continueChecking = TryAdvanceToNextCandidate();
                    if (!continueChecking)
                    {
                        SetOfflineState();
                        ScheduleNextCheck(false);
                    }
                    return;
                }

                if (channelState == "Online")
                {
                    var candidate = checkQueue[currentCheckQueueIndex];
                    var rule = config.priorityRules[candidate.RuleIndex];

                    if (string.Equals(rule.displayMode, "Gif", StringComparison.OrdinalIgnoreCase))
                    {
                        string assetPath = NormalizeDisplayAssetPath(rule.displayValue);
                        EnsureGifLoaded(assetPath);
                        if (gifImage == null)
                        {
                            continueChecking = TryAdvanceToNextCandidate();
                            if (!continueChecking)
                            {
                                lastOpenUrl = currentCheckUrl;
                                lastOpenChannelName = currentCheckName;
                                lastOpenRuleIndex = candidate.RuleIndex;
                                SetStatusUi("Online", "", null);
                                ScheduleNextCheck(true);
                            }
                            return;
                        }

                        lastOpenUrl = currentCheckUrl;
                        lastOpenChannelName = currentCheckName;
                        lastOpenRuleIndex = candidate.RuleIndex;
                        SetStatusUi("Online", "", assetPath);
                        ScheduleNextCheck(true);
                    }
                    else
                    {
                        lastOpenUrl = currentCheckUrl;
                        lastOpenChannelName = currentCheckName;
                        lastOpenRuleIndex = candidate.RuleIndex;
                        SetStatusUi("Online", rule.displayValue, null);
                        ScheduleNextCheck(true);
                    }
                }
                else
                {
                    continueChecking = TryAdvanceToNextCandidate();
                    if (!continueChecking)
                    {
                        SetOfflineState();
                        ScheduleNextCheck(false);
                    }
                }
            }
            catch
            {
                continueChecking = TryAdvanceToNextCandidate();
                if (!continueChecking)
                {
                    SetOfflineState();
                    ScheduleNextCheck(false);
                }
            }
            finally
            {
                settleTimer.Stop();
                if (!continueChecking)
                {
                    isChecking = false;
                }
            }
        }

        private bool TryAdvanceToNextCandidate()
        {
            int nextIndex = currentCheckQueueIndex + 1;
            if (nextIndex < checkQueue.Count)
            {
                BeginCheckAt(nextIndex);
                return true;
            }

            return false;
        }

        private void SetOfflineState()
        {
            InitializeFallbackOpenTarget();
            SetStatusUi("Offline", "", null);
        }

        private void SetInitialLoadingState()
        {
            hasResolvedVisualState = false;
            spinnerInitialMode = true;
            statusPanel.BackColor = Color.Gray;
            statusPanel.BackgroundImage = null;
            currentMarker = "";
            markerLabel.Text = "";
            markerLabel.Visible = false;
            gifBox.Visible = false;
            gifBox.Image = null;
            spinnerControl.Visible = true;
            SyncSpinnerHost();
            UpdateSpinnerLayout();
            spinnerControl.BringToFront();
        }

        private void ScheduleNextCheck(bool foundLive)
        {
            int interval = foundLive ? GetRefreshIntervalMilliseconds(config) : 15000;
            refreshTimer.Stop();
            refreshTimer.Interval = interval;
            refreshTimer.Start();
        }

        private void SetStatusUi(string stateName, string marker, string gifPath)
        {
            bool isCheckingState = string.Equals(stateName, "Checking", StringComparison.OrdinalIgnoreCase);
            bool isInitialCheckState = isCheckingState && !hasResolvedVisualState;
            switch (stateName)
            {
                case "Online":
                    statusPanel.BackColor = Color.LimeGreen;
                    hasResolvedVisualState = true;
                    break;
                case "Offline":
                    statusPanel.BackColor = Color.Red;
                    hasResolvedVisualState = true;
                    break;
                default:
                    if (isInitialCheckState)
                    {
                        statusPanel.BackColor = Color.Gray;
                    }
                    break;
            }

            spinnerControl.Visible = isCheckingState;
            spinnerInitialMode = isInitialCheckState;
            SyncSpinnerHost();
            UpdateSpinnerLayout();
            if (spinnerControl.Visible)
            {
                spinnerControl.BringToFront();
            }

            if (isCheckingState && !isInitialCheckState)
            {
                return;
            }

            currentMarker = marker ?? "";
            markerLabel.Text = "";
            markerLabel.Visible = !string.IsNullOrWhiteSpace(currentMarker);
            markerLabel.Invalidate();

            if (!string.IsNullOrWhiteSpace(gifPath))
            {
                EnsureGifLoaded(gifPath);
                if (gifImage != null)
                {
                    statusPanel.BackgroundImage = null;
                    gifBox.Image = gifImage;
                    gifBox.Visible = true;
                    gifBox.BringToFront();
                    if (ImageAnimator.CanAnimate(gifImage))
                    {
                        ImageAnimator.Animate(gifImage, delegate { BeginInvoke((Action)(delegate { gifBox.Invalidate(); })); });
                    }
                    if (markerLabel.Visible)
                    {
                        markerLabel.BringToFront();
                    }
                    if (spinnerControl.Visible)
                    {
                        spinnerControl.BringToFront();
                    }
                    return;
                }
            }

            if (gifImage != null && ImageAnimator.CanAnimate(gifImage))
            {
                ImageAnimator.StopAnimate(gifImage, null);
            }
            statusPanel.BackgroundImage = null;
            gifBox.Visible = false;
            gifBox.Image = null;
            if (markerLabel.Visible)
            {
                markerLabel.BringToFront();
            }
            if (spinnerControl.Visible)
            {
                spinnerControl.BringToFront();
            }
        }

        private void PositionSpinner()
        {
            if (spinnerControl == null)
            {
                return;
            }

            var host = spinnerControl.Parent ?? statusPanel;

            if (spinnerInitialMode)
            {
                spinnerControl.Location = new Point(
                    Math.Max(0, (host.ClientSize.Width - spinnerControl.Width) / 2),
                    Math.Max(0, (host.ClientSize.Height - spinnerControl.Height) / 2));
            }
            else
            {
                int margin = Math.Max(3, (int)Math.Round(4 * currentScale));
                spinnerControl.Location = new Point(
                    Math.Max(0, host.ClientSize.Width - spinnerControl.Width - margin),
                    margin);
            }
        }

        private void UpdateSpinnerLayout()
        {
            if (spinnerControl == null)
            {
                return;
            }

            int spinnerSize;
            if (spinnerInitialMode)
            {
                spinnerSize = Math.Max(16, (int)Math.Round(26 * currentScale));
            }
            else
            {
                spinnerSize = Math.Max(10, (int)Math.Round(16 * currentScale));
            }

            spinnerControl.Size = new Size(spinnerSize, spinnerSize);
            PositionSpinner();
        }

        private void SyncSpinnerHost()
        {
            if (spinnerControl == null)
            {
                return;
            }

            Control host = statusPanel;
            if (!spinnerInitialMode && gifBox.Visible)
            {
                host = gifBox;
            }
            else if (!spinnerInitialMode && markerLabel.Visible)
            {
                host = markerLabel;
            }

            if (spinnerControl.Parent != host)
            {
                spinnerControl.Parent = host;
            }
        }

        private void EnsureGifLoaded(string gifPath)
        {
            gifPath = NormalizeDisplayAssetPath(gifPath);
            if (string.Equals(currentGifPath, gifPath, StringComparison.OrdinalIgnoreCase) && gifImage != null)
            {
                return;
            }

            try
            {
                if (gifImage != null)
                {
                    gifImage.Dispose();
                }
            }
            catch
            {
            }

            gifImage = null;
            currentGifPath = gifPath ?? "";

            try
            {
                if (!string.IsNullOrWhiteSpace(gifPath) && File.Exists(gifPath))
                {
                    gifImage = LoadDisplayImage(gifPath);
                }
            }
            catch
            {
                gifImage = null;
            }
        }

        private static Image LoadDisplayImage(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var loaded = Image.FromStream(stream))
                    {
                        return new Bitmap(loaded);
                    }
                }
            }
            catch
            {
            }

            try
            {
                var uri = new Uri(path, UriKind.Absolute);
                var decoder = BitmapDecoder.Create(uri, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                var frame = decoder.Frames.FirstOrDefault();
                if (frame == null)
                {
                    return null;
                }

                var encoder = new BmpBitmapEncoder();
                encoder.Frames.Add(frame);
                using (var ms = new MemoryStream())
                {
                    encoder.Save(ms);
                    ms.Position = 0;
                    using (var loaded = Image.FromStream(ms))
                    {
                        return new Bitmap(loaded);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        private static string NormalizeDisplayAssetPath(string rawPath)
        {
            if (string.IsNullOrWhiteSpace(rawPath))
            {
                return "";
            }

            string value = rawPath.Trim();
            if (value.StartsWith("\"") && value.EndsWith("\"") && value.Length >= 2)
            {
                value = value.Substring(1, value.Length - 2);
            }

            int markdownStart = value.IndexOf("](");
            if (markdownStart >= 0 && value.EndsWith(")"))
            {
                value = value.Substring(markdownStart + 2, value.Length - markdownStart - 3);
            }

            return value.Trim();
        }

        private void BeginUpdateCheck(bool force)
        {
            if (!force && System.Threading.Interlocked.Exchange(ref updateCheckQueued, 1) == 1)
            {
                return;
            }

            System.Threading.ThreadPool.QueueUserWorkItem(delegate
            {
                if (force)
                {
                    System.Threading.Interlocked.Exchange(ref updateCheckQueued, 1);
                }

                try
                {
                    CheckForUpdates();
                }
                finally
                {
                    System.Threading.Interlocked.Exchange(ref updateCheckQueued, 0);
                }
            });
        }

        private void CheckForUpdates()
        {
            try
            {
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Tls12 |
                    SecurityProtocolType.Tls11 |
                    SecurityProtocolType.Tls;

                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.UserAgent] = GitHubUserAgent;
                    string manifestJson = DownloadManifestJson(client);
                    var manifest = serializer.Deserialize<AppUpdateManifest>(manifestJson);
                    if (manifest == null ||
                        string.IsNullOrWhiteSpace(manifest.version) ||
                        string.IsNullOrWhiteSpace(manifest.packageUrl))
                    {
                        return;
                    }

                    if (!IsNewerVersion(manifest.version, GeneratedVersion.Number))
                    {
                        return;
                    }

                    DownloadAndInstallUpdate(manifest);
                }
            }
            catch
            {
            }
        }

        private string DownloadManifestJson(WebClient client)
        {
            string nonce = DateTime.UtcNow.Ticks.ToString();

            try
            {
                client.Headers[HttpRequestHeader.Accept] = "application/vnd.github+json";
                string apiResponse = client.DownloadString(UpdateManifestApiUrl + "&v=" + nonce);
                var contentResponse = serializer.Deserialize<GitHubContentResponse>(apiResponse);
                if (contentResponse != null &&
                    string.Equals(contentResponse.encoding, "base64", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(contentResponse.content))
                {
                    string base64 = contentResponse.content.Replace("\r", "").Replace("\n", "");
                    return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(base64));
                }
            }
            catch
            {
            }

            client.Headers[HttpRequestHeader.Accept] = "application/json";
            return client.DownloadString(UpdateManifestRawUrl + "?v=" + nonce);
        }

        private static bool IsNewerVersion(string remoteVersion, string currentVersion)
        {
            if (string.IsNullOrWhiteSpace(remoteVersion) || string.IsNullOrWhiteSpace(currentVersion))
            {
                return false;
            }

            Version remote;
            Version current;
            if (Version.TryParse(remoteVersion, out remote) && Version.TryParse(currentVersion, out current))
            {
                return remote > current;
            }

            return !string.Equals(remoteVersion, currentVersion, StringComparison.OrdinalIgnoreCase);
        }

        private void DownloadAndInstallUpdate(AppUpdateManifest manifest)
        {
            if (updateInProgress)
            {
                return;
            }

            updateInProgress = true;

            try
            {
                string tempRoot = Path.Combine(Path.GetTempPath(), "KickStatusSquareUpdate", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempRoot);

                string packagePath = Path.Combine(tempRoot, "KickStatusSquare-package.zip");
                DownloadUpdatePackage(manifest, packagePath);

                string updaterScriptPath = WriteUpdaterScript(tempRoot);
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = string.Format(
                        "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File \"{0}\" -PackagePath \"{1}\" -InstallDir \"{2}\" -ExeName \"{3}\"",
                        updaterScriptPath,
                        packagePath,
                        appDir.TrimEnd('\\'),
                        Path.GetFileName(Application.ExecutablePath)),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = tempRoot
                };

                BeginInvoke((Action)delegate
                {
                    refreshTimer.Stop();
                    watchdogTimer.Stop();
                    settleTimer.Stop();
                    try
                    {
                        browser.Stop();
                    }
                    catch
                    {
                    }

                    try
                    {
                        browserHost.Close();
                    }
                    catch
                    {
                    }

                    Process.Start(startInfo);
                    Close();
                });
            }
            catch
            {
                updateInProgress = false;
            }
        }

        private void DownloadUpdatePackage(AppUpdateManifest manifest, string packagePath)
        {
            string versionNonce = manifest != null && !string.IsNullOrWhiteSpace(manifest.version)
                ? Uri.EscapeDataString(manifest.version)
                : DateTime.UtcNow.Ticks.ToString();

            using (var client = new WebClient())
            {
                client.Headers[HttpRequestHeader.UserAgent] = GitHubUserAgent;

                try
                {
                    client.Headers[HttpRequestHeader.Accept] = "application/vnd.github+json";
                    string apiResponse = client.DownloadString(UpdatePackageApiUrl + "&v=" + versionNonce);
                    var contentResponse = serializer.Deserialize<GitHubContentResponse>(apiResponse);
                    if (contentResponse != null &&
                        string.Equals(contentResponse.encoding, "base64", StringComparison.OrdinalIgnoreCase) &&
                        !string.IsNullOrWhiteSpace(contentResponse.content))
                    {
                        string base64 = contentResponse.content.Replace("\r", "").Replace("\n", "");
                        File.WriteAllBytes(packagePath, Convert.FromBase64String(base64));
                        return;
                    }
                }
                catch
                {
                }

                if (manifest == null || string.IsNullOrWhiteSpace(manifest.packageUrl))
                {
                    throw new InvalidOperationException("Update manifest did not include a package URL.");
                }

                client.Headers[HttpRequestHeader.Accept] = "application/octet-stream";
                client.DownloadFile(manifest.packageUrl + "?v=" + versionNonce, packagePath);
            }
        }

        private static string WriteUpdaterScript(string tempRoot)
        {
            string updaterPath = Path.Combine(tempRoot, "apply-update.ps1");
            string script = @"param(
    [string]$PackagePath,
    [string]$InstallDir,
    [string]$ExeName
)

$ErrorActionPreference = 'Stop'
$exePath = Join-Path $InstallDir $ExeName
$extractDir = Join-Path ([System.IO.Path]::GetDirectoryName($PackagePath)) 'package'

for ($i = 0; $i -lt 60; $i++) {
    try {
        if (Test-Path $exePath) {
            $stream = [System.IO.File]::Open($exePath, 'Open', 'ReadWrite', 'None')
            $stream.Close()
        }
        break
    }
    catch {
        Start-Sleep -Milliseconds 500
    }
}

if (Test-Path $extractDir) {
    Remove-Item -LiteralPath $extractDir -Recurse -Force
}

Expand-Archive -LiteralPath $PackagePath -DestinationPath $extractDir -Force
Get-ChildItem -Path $extractDir -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $InstallDir $_.Name) -Recurse -Force
}

Start-Process -FilePath (Join-Path $InstallDir $ExeName)
";
            File.WriteAllText(updaterPath, script);
            return updaterPath;
        }

        private AppConfig LoadConfig()
        {
            AppConfig cfg = null;

            try
            {
                if (File.Exists(configPath))
                {
                    cfg = serializer.Deserialize<AppConfig>(File.ReadAllText(configPath));
                }
            }
            catch
            {
            }

            if (cfg == null)
            {
                try
                {
                    string legacyConfig = Path.Combine(appDir, "config.json");
                    if (File.Exists(legacyConfig))
                    {
                        cfg = serializer.Deserialize<AppConfig>(File.ReadAllText(legacyConfig));
                    }
                }
                catch
                {
                }
            }

            cfg = NormalizeConfig(cfg ?? new AppConfig());

            try
            {
                File.WriteAllText(configPath, serializer.Serialize(cfg));
            }
            catch
            {
            }

            return cfg;
        }

        private static AppConfig NormalizeConfig(AppConfig cfg)
        {
            if (cfg == null)
            {
                cfg = new AppConfig();
            }

            if (cfg.refreshMinutes <= 0)
            {
                cfg.refreshMinutes = 1.0;
            }

            if (cfg.priorityRules == null || cfg.priorityRules.Count == 0)
            {
                cfg.priorityRules = new List<PriorityRule>();

                if (!string.IsNullOrWhiteSpace(cfg.primaryChannel))
                {
                    cfg.priorityRules.Add(new PriorityRule
                    {
                        channels = new List<string> { cfg.primaryChannel.Trim() },
                        displayMode = "Text",
                        displayValue = "D"
                    });
                }

                if (!string.IsNullOrWhiteSpace(cfg.secondaryChannel))
                {
                    cfg.priorityRules.Add(new PriorityRule
                    {
                        channels = new List<string> { cfg.secondaryChannel.Trim() },
                        displayMode = "Text",
                        displayValue = "AE"
                    });
                }

                cfg.priorityRules.Add(new PriorityRule
                {
                    channels = cfg.priorityChannels != null ? cfg.priorityChannels.Where(IsNonEmpty).Select(TrimValue).ToList() : new List<string>(),
                    displayMode = "Gif",
                    displayValue = cfg.priorityGifPath ?? ""
                });
            }

            foreach (var rule in cfg.priorityRules)
            {
                if (rule.channels == null)
                {
                    rule.channels = new List<string>();
                }

                rule.channels = rule.channels.Where(IsNonEmpty).Select(TrimValue).ToList();

                if (string.IsNullOrWhiteSpace(rule.displayMode))
                {
                    rule.displayMode = "Text";
                }

                if (!string.Equals(rule.displayMode, "Gif", StringComparison.OrdinalIgnoreCase))
                {
                    rule.displayMode = "Text";
                }

                if (rule.displayValue == null)
                {
                    rule.displayValue = "";
                }
            }

            while (cfg.priorityRules.Count < 3)
            {
                cfg.priorityRules.Add(new PriorityRule());
            }

            if (cfg.priorityRules.Count > 0 && string.IsNullOrWhiteSpace(cfg.priorityRules[0].displayValue))
            {
                cfg.priorityRules[0].displayMode = "Text";
                cfg.priorityRules[0].displayValue = "D";
            }

            if (cfg.priorityRules.Count > 1 && string.IsNullOrWhiteSpace(cfg.priorityRules[1].displayValue))
            {
                cfg.priorityRules[1].displayMode = "Text";
                cfg.priorityRules[1].displayValue = "AE";
            }

            cfg.primaryChannel = cfg.priorityRules.Count > 0 && cfg.priorityRules[0].channels.Count > 0 ? cfg.priorityRules[0].channels[0] : "destiny";
            cfg.secondaryChannel = cfg.priorityRules.Count > 1 && cfg.priorityRules[1].channels.Count > 0 ? cfg.priorityRules[1].channels[0] : "anythingelse";
            cfg.priorityChannels = cfg.priorityRules.Count > 2 ? cfg.priorityRules.Skip(2).SelectMany(r => r.channels).ToList() : new List<string>();
            cfg.priorityGifPath = cfg.priorityRules.Count > 2 && string.Equals(cfg.priorityRules[2].displayMode, "Gif", StringComparison.OrdinalIgnoreCase) ? cfg.priorityRules[2].displayValue : "";

            return cfg;
        }

        private StateFile LoadState()
        {
            try
            {
                if (File.Exists(statePath))
                {
                    return serializer.Deserialize<StateFile>(File.ReadAllText(statePath)) ?? new StateFile();
                }
            }
            catch
            {
            }

            return new StateFile();
        }

        private void SaveWindowState()
        {
            try
            {
                state = LoadState();
                state.layouts[GetDisplayLayoutKey()] = new WindowState { x = Location.X, y = Location.Y, scale = currentScale };
                File.WriteAllText(statePath, serializer.Serialize(state));
            }
            catch
            {
            }
        }

        private WindowState GetSavedLocation()
        {
            string key = GetDisplayLayoutKey();
            if (state.layouts != null && state.layouts.ContainsKey(key))
            {
                return state.layouts[key];
            }
            return null;
        }

        private bool IsLocationVisible(Point location)
        {
            var rect = new Rectangle(location, Size);
            foreach (var screen in Screen.AllScreens)
            {
                var working = screen.WorkingArea;
                if (working.IntersectsWith(rect))
                {
                    return true;
                }
            }

            return false;
        }

        private void CenterToVisibleScreen()
        {
            var working = Screen.PrimaryScreen != null ? Screen.PrimaryScreen.WorkingArea : new Rectangle(0, 0, 800, 600);
            Location = new Point(
                working.Left + Math.Max(0, (working.Width - Width) / 2),
                working.Top + Math.Max(0, (working.Height - Height) / 2));
        }

        private void InitializeFallbackOpenTarget()
        {
            var fallbackQueue = BuildCheckQueue(config);
            if (fallbackQueue.Count > 0)
            {
                lastOpenRuleIndex = fallbackQueue[0].RuleIndex;
                lastOpenChannelName = fallbackQueue[0].ChannelName;
                lastOpenUrl = BuildKickUrl(lastOpenChannelName);
            }
            else
            {
                lastOpenRuleIndex = -1;
                lastOpenChannelName = "";
                lastOpenUrl = "";
            }
        }

        private static List<CheckCandidate> BuildCheckQueue(AppConfig cfg)
        {
            var result = new List<CheckCandidate>();
            if (cfg == null || cfg.priorityRules == null)
            {
                return result;
            }

            for (int i = 0; i < cfg.priorityRules.Count; i++)
            {
                var rule = cfg.priorityRules[i];
                if (rule.channels == null)
                {
                    continue;
                }

                foreach (var channel in rule.channels.Where(IsNonEmpty).Select(TrimValue))
                {
                    result.Add(new CheckCandidate { RuleIndex = i, ChannelName = channel });
                }
            }

            return result;
        }

        private static string GetDisplayLayoutKey()
        {
            var screens = Screen.AllScreens
                .OrderBy(s => s.Bounds.X)
                .ThenBy(s => s.Bounds.Y)
                .ThenBy(s => s.Bounds.Width)
                .ThenBy(s => s.Bounds.Height)
                .Select(s => string.Format("{0},{1},{2},{3}", s.Bounds.X, s.Bounds.Y, s.Bounds.Width, s.Bounds.Height));
            return string.Format("count={0}|{1}", Screen.AllScreens.Length, string.Join(";", screens));
        }

        private static string BuildKickUrl(string channelName)
        {
            return string.Format("https://kick.com/{0}", channelName);
        }

        private static string GetChannelState(string normalizedText, string channelName)
        {
            if (string.IsNullOrWhiteSpace(normalizedText))
            {
                return "Unknown";
            }

            string lower = normalizedText.ToLowerInvariant();
            string channel = string.IsNullOrWhiteSpace(channelName) ? "" : channelName.Trim().ToLowerInvariant();

            if (lower.Contains("chatoffline") ||
                lower.Contains(" is offline") ||
                (!string.IsNullOrWhiteSpace(channel) && lower.Contains(channel + " is offline")))
            {
                return "Offline";
            }

            if ((!string.IsNullOrWhiteSpace(channel) && (lower.Contains("live " + channel) || lower.Contains("live" + channel))) ||
                (lower.Contains("live") && lower.Contains("viewers") && !string.IsNullOrWhiteSpace(channel) && lower.Contains(channel)))
            {
                return "Online";
            }

            return "Unknown";
        }

        private static string NormalizeWhitespace(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return "";
            }

            var chars = input.Where(c => !char.IsControl(c) || c == ' ' || c == '\n' || c == '\r' || c == '\t').ToArray();
            var compact = new string(chars);
            return System.Text.RegularExpressions.Regex.Replace(compact, @"\s+", " ").Trim();
        }

        private static bool IsNonEmpty(string value)
        {
            return !string.IsNullOrWhiteSpace(value);
        }

        private static string TrimValue(string value)
        {
            return value == null ? "" : value.Trim();
        }

        private static int GetRefreshIntervalMilliseconds(AppConfig cfg)
        {
            double minutes = 1.0;
            if (cfg != null && cfg.refreshMinutes > 0)
            {
                minutes = cfg.refreshMinutes;
            }

            double milliseconds = minutes * 60d * 1000d;
            if (milliseconds < 1000d)
            {
                milliseconds = 1000d;
            }
            if (milliseconds > int.MaxValue)
            {
                milliseconds = int.MaxValue;
            }

            return (int)Math.Round(milliseconds);
        }
    }

    internal class SpinnerControl : Control
    {
        private readonly Timer animationTimer;
        private float angle;

        public SpinnerControl()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            animationTimer = new Timer { Interval = 80 };
            animationTimer.Tick += delegate
            {
                angle += 30f;
                if (angle >= 360f)
                {
                    angle -= 360f;
                }
                Invalidate();
            };
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (Visible)
            {
                animationTimer.Start();
            }
            else
            {
                animationTimer.Stop();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new RectangleF(2f, 2f, Width - 5f, Height - 5f);
            using (var outline = new Pen(Color.Black, Math.Max(2f, Width / 7f)))
            using (var stroke = new Pen(Color.White, Math.Max(1.2f, Width / 10f)))
            {
                e.Graphics.DrawArc(outline, rect, angle, 270f);
                e.Graphics.DrawArc(stroke, rect, angle, 270f);
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            NativeMethods.SetCurrentProcessExplicitAppUserModelID("Seths.KickStatusSquare");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new StatusForm());
        }
    }
}
