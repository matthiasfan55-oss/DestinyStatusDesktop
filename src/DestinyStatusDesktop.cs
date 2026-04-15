using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace DestinyStatusDesktop
{
    public class AppConfig
    {
        public string primaryChannel { get; set; }
        public string secondaryChannel { get; set; }
        public List<string> priorityChannels { get; set; }
        public string primaryImagePath { get; set; }
        public string secondaryImagePath { get; set; }
        public string priorityImagePath { get; set; }
        public string priorityGifPath { get; set; }

        public AppConfig()
        {
            primaryChannel = "destiny";
            secondaryChannel = "anythingelse";
            priorityChannels = new List<string>();
            primaryImagePath = "";
            secondaryImagePath = "";
            priorityImagePath = "";
            priorityGifPath = "";
        }
    }

    public class WindowState
    {
        public int x { get; set; }
        public int y { get; set; }
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

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern int SetCurrentProcessExplicitAppUserModelID(string appID);
    }

    public class StatusForm : Form
    {
        private const string AppDisplayName = "Destiny Status Desktop";
        private const string AppDataFolderName = "DestinyStatusDesktopData";
        private const string LegacyAppDataFolderName = "KickStatusAppData";
        private const string UpdateManifestUrl = "https://raw.githubusercontent.com/matthiasfan55-oss/DestinyStatusDesktop/main/update.json";
        private const string GitHubUserAgent = "DestinyStatusDesktop-Updater";

        private readonly string appDir;
        private readonly string dataDir;
        private readonly string legacyDataDir;
        private readonly string configPath;
        private readonly string statePath;
        private readonly string settingsLauncherPath;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();

        private AppConfig config;
        private StateFile state;

        private readonly Panel statusPanel;
        private readonly PictureBox gifBox;
        private readonly Label markerLabel;
        private readonly WebBrowser browser;
        private readonly Form browserHost;
        private readonly System.Windows.Forms.Timer settleTimer;
        private readonly System.Windows.Forms.Timer watchdogTimer;
        private readonly System.Windows.Forms.Timer refreshTimer;
        private readonly ContextMenuStrip menu;

        private bool isChecking;
        private DateTime loadStartedAt;
        private string currentCheckName;
        private string currentCheckUrl;
        private string lastOpenUrl;
        private int priorityIndex;
        private Image primaryImage;
        private Image secondaryImage;
        private Image priorityImage;
        private Image animatedImage;

        private bool leftMouseDown;
        private bool dragStarted;
        private Point mouseDownPoint;
        private int updateCheckQueued;
        private bool updateInProgress;

        public StatusForm()
        {
            appDir = AppDomain.CurrentDomain.BaseDirectory;
            dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppDataFolderName);
            legacyDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), LegacyAppDataFolderName);
            Directory.CreateDirectory(dataDir);
            MigrateLegacyDataFiles();
            configPath = Path.Combine(dataDir, "config.json");
            statePath = Path.Combine(dataDir, "state.json");
            settingsLauncherPath = Path.Combine(appDir, "Edit Destiny Status Channels.cmd");

            config = LoadConfig();
            state = LoadState();

            currentCheckName = config.primaryChannel;
            currentCheckUrl = BuildKickUrl(currentCheckName);
            lastOpenUrl = currentCheckUrl;

            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(64, 64);
            MinimumSize = new Size(64, 64);
            MaximumSize = new Size(64, 64);
            ShowInTaskbar = true;
            ShowIcon = false;
            Text = AppDisplayName;
            BackColor = Color.FromArgb(24, 24, 24);
            TopMost = true;

            statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Gray,
                Cursor = Cursors.Hand,
                Margin = Padding.Empty
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
            statusPanel.Controls.Add(markerLabel);
            markerLabel.BringToFront();

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

            settleTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            settleTimer.Tick += (_, __) =>
            {
                settleTimer.Stop();
                CompleteStatusCheck();
            };

            watchdogTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            watchdogTimer.Tick += (_, __) =>
            {
                if (isChecking && (DateTime.Now - loadStartedAt).TotalSeconds >= 20)
                {
                    isChecking = false;
                    settleTimer.Stop();
                    SetStatusUi("Offline", "", null);
                }
            };
            watchdogTimer.Start();

            refreshTimer = new System.Windows.Forms.Timer { Interval = 60000 };
            refreshTimer.Tick += (_, __) => StartStatusCheck();
            refreshTimer.Start();

            menu = new ContextMenuStrip();
            var settingsItem = new ToolStripMenuItem("Settings");
            settingsItem.Click += (_, __) =>
            {
                if (File.Exists(settingsLauncherPath))
                {
                    Process.Start(settingsLauncherPath);
                }
            };
            menu.Items.Add(settingsItem);

            var updateItem = new ToolStripMenuItem("Check for Updates");
            updateItem.Click += (_, __) => BeginUpdateCheck(true);
            menu.Items.Add(updateItem);

            HookPointer(statusPanel);
            HookPointer(markerLabel);
            HookPointer(gifBox);

            var saved = GetSavedLocation();
            if (saved != null)
            {
                Location = new Point(saved.x, saved.y);
            }
            else
            {
                CenterToScreen();
            }

            Move += (_, __) => SaveWindowState();
            FormClosing += (_, __) =>
            {
                SaveWindowState();
                browserHost.Close();
                DisposeImage(ref primaryImage);
                DisposeImage(ref secondaryImage);
                DisposeImage(ref priorityImage);
            };
            Shown += (_, __) =>
            {
                browserHost.Show();
                StartStatusCheck();
                BeginUpdateCheck(false);
            };

            LoadLevelImages();
        }

        private void MigrateLegacyDataFiles()
        {
            try
            {
                if (!Directory.Exists(legacyDataDir))
                {
                    return;
                }

                string legacyConfigPath = Path.Combine(legacyDataDir, "config.json");
                string legacyStatePath = Path.Combine(legacyDataDir, "state.json");

                string newConfigPath = Path.Combine(dataDir, "config.json");
                string newStatePath = Path.Combine(dataDir, "state.json");

                if (!File.Exists(newConfigPath) && File.Exists(legacyConfigPath))
                {
                    File.Copy(legacyConfigPath, newConfigPath, false);
                }

                if (!File.Exists(newStatePath) && File.Exists(legacyStatePath))
                {
                    File.Copy(legacyStatePath, newStatePath, false);
                }
            }
            catch
            {
            }
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
                menu.Show(this, e.Location);
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

            if (!wasDrag && !string.IsNullOrWhiteSpace(lastOpenUrl))
            {
                Process.Start(lastOpenUrl);
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

            config = LoadConfig();
            LoadLevelImages();

            isChecking = true;
            currentCheckName = config.primaryChannel;
            currentCheckUrl = BuildKickUrl(currentCheckName);
            priorityIndex = 0;
            loadStartedAt = DateTime.Now;
            SetStatusUi("Checking", "", null);

            try
            {
                browser.Navigate(currentCheckUrl);
            }
            catch
            {
                isChecking = false;
                SetStatusUi("Offline", "", null);
            }
        }

        private void CompleteStatusCheck()
        {
            bool continueChecking = false;
            try
            {
                if (browser.Document == null || browser.Document.Body == null)
                {
                    SetStatusUi("Offline", "", null);
                    return;
                }

                string text = browser.Document.Body.InnerText ?? "";
                string normalized = NormalizeWhitespace(text);
                string channelState = GetChannelState(normalized);

                if (channelState == "Unknown")
                {
                    SetStatusUi("Offline", "", null);
                }
                else if (channelState == "Online")
                {
                    lastOpenUrl = currentCheckUrl;
                    if (string.Equals(currentCheckName, config.primaryChannel, StringComparison.OrdinalIgnoreCase))
                    {
                        SetStatusUi("Online", "D", primaryImage);
                    }
                    else if (string.Equals(currentCheckName, config.secondaryChannel, StringComparison.OrdinalIgnoreCase))
                    {
                        SetStatusUi("Online", "AE", secondaryImage);
                    }
                    else
                    {
                        SetStatusUi("Online", "", priorityImage);
                    }
                }
                else if (string.Equals(currentCheckName, config.primaryChannel, StringComparison.OrdinalIgnoreCase))
                {
                    currentCheckName = config.secondaryChannel;
                    currentCheckUrl = BuildKickUrl(currentCheckName);
                    loadStartedAt = DateTime.Now;
                    continueChecking = true;
                    SetStatusUi("Checking", "", null);
                    browser.Navigate(currentCheckUrl);
                    return;
                }
                else if (string.Equals(currentCheckName, config.secondaryChannel, StringComparison.OrdinalIgnoreCase))
                {
                    if (config.priorityChannels != null && config.priorityChannels.Count > 0)
                    {
                        priorityIndex = 0;
                        currentCheckName = config.priorityChannels[priorityIndex];
                        currentCheckUrl = BuildKickUrl(currentCheckName);
                        loadStartedAt = DateTime.Now;
                        continueChecking = true;
                        SetStatusUi("Checking", "", null);
                        browser.Navigate(currentCheckUrl);
                        return;
                    }

                    lastOpenUrl = BuildKickUrl(config.primaryChannel);
                    SetStatusUi("Offline", "", null);
                }
                else if (config.priorityChannels != null && priorityIndex < config.priorityChannels.Count - 1)
                {
                    priorityIndex++;
                    currentCheckName = config.priorityChannels[priorityIndex];
                    currentCheckUrl = BuildKickUrl(currentCheckName);
                    loadStartedAt = DateTime.Now;
                    continueChecking = true;
                    SetStatusUi("Checking", "", null);
                    browser.Navigate(currentCheckUrl);
                    return;
                }
                else
                {
                    lastOpenUrl = BuildKickUrl(config.primaryChannel);
                    SetStatusUi("Offline", "", null);
                }
            }
            catch
            {
                SetStatusUi("Offline", "", null);
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

        private void SetStatusUi(string stateName, string marker, Image displayImage)
        {
            switch (stateName)
            {
                case "Online":
                    statusPanel.BackColor = Color.LimeGreen;
                    break;
                case "Offline":
                    statusPanel.BackColor = Color.Red;
                    break;
                default:
                    statusPanel.BackColor = Color.Goldenrod;
                    break;
            }

            markerLabel.Text = marker ?? "";
            if (animatedImage != null && !ReferenceEquals(animatedImage, displayImage) && ImageAnimator.CanAnimate(animatedImage))
            {
                ImageAnimator.StopAnimate(animatedImage, null);
                animatedImage = null;
            }

            if (displayImage != null)
            {
                gifBox.Image = displayImage;
                gifBox.Visible = true;
                markerLabel.Visible = false;

                if (ImageAnimator.CanAnimate(displayImage))
                {
                    animatedImage = displayImage;
                    ImageAnimator.Animate(displayImage, (_, __) => BeginInvoke((Action)(() => gifBox.Invalidate())));
                }
            }
            else
            {
                gifBox.Visible = false;
                gifBox.Image = null;
                markerLabel.Visible = true;
                markerLabel.BringToFront();
            }
        }

        private void DisposeImage(ref Image image)
        {
            if (image == null)
            {
                return;
            }

            if (ReferenceEquals(animatedImage, image) && ImageAnimator.CanAnimate(animatedImage))
            {
                ImageAnimator.StopAnimate(animatedImage, null);
                animatedImage = null;
            }

            image.Dispose();
            image = null;
        }

        private static string NormalizeImagePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            string clean = path.Trim();
            if (clean.Length >= 2)
            {
                bool startsWithQuote = clean.StartsWith("\"") || clean.StartsWith("'");
                bool endsWithQuote = clean.EndsWith("\"") || clean.EndsWith("'");
                if (startsWithQuote && endsWithQuote)
                {
                    clean = clean.Substring(1, clean.Length - 2).Trim();
                }
            }

            return clean;
        }

        private Image LoadImageFromPath(string path)
        {
            try
            {
                string normalizedPath = NormalizeImagePath(path);
                if (!string.IsNullOrWhiteSpace(normalizedPath) && File.Exists(normalizedPath))
                {
                    return Image.FromFile(normalizedPath);
                }
            }
            catch
            {
            }

            return null;
        }

        private void LoadLevelImages()
        {
            DisposeImage(ref primaryImage);
            DisposeImage(ref secondaryImage);
            DisposeImage(ref priorityImage);

            primaryImage = LoadImageFromPath(config.primaryImagePath);
            secondaryImage = LoadImageFromPath(config.secondaryImagePath);

            string priorityPath = !string.IsNullOrWhiteSpace(config.priorityImagePath)
                ? config.priorityImagePath
                : config.priorityGifPath;
            priorityImage = LoadImageFromPath(priorityPath);
        }

        private static List<string> ConvertToStringList(object value)
        {
            var results = new List<string>();
            var values = value as object[];
            if (values == null)
            {
                return results;
            }

            foreach (object item in values)
            {
                string text = item as string;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    results.Add(text.Trim());
                }
            }

            return results;
        }

        private string GetRuleDisplayValue(string rawJson, IEnumerable<string> channels)
        {
            try
            {
                var targetChannels = new HashSet<string>(
                    channels.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()),
                    StringComparer.OrdinalIgnoreCase);

                if (targetChannels.Count == 0)
                {
                    return string.Empty;
                }

                var root = serializer.DeserializeObject(rawJson) as Dictionary<string, object>;
                if (root == null || !root.ContainsKey("priorityRules"))
                {
                    return string.Empty;
                }

                var rules = root["priorityRules"] as object[];
                if (rules == null)
                {
                    return string.Empty;
                }

                foreach (object ruleObject in rules)
                {
                    var rule = ruleObject as Dictionary<string, object>;
                    if (rule == null || !rule.ContainsKey("channels") || !rule.ContainsKey("displayValue"))
                    {
                        continue;
                    }

                    List<string> ruleChannels = ConvertToStringList(rule["channels"]);
                    if (ruleChannels.Count == 1 && targetChannels.Count == 1 && targetChannels.Contains(ruleChannels[0]))
                    {
                        return NormalizeImagePath(rule["displayValue"] as string);
                    }
                }

                foreach (object ruleObject in rules)
                {
                    var rule = ruleObject as Dictionary<string, object>;
                    if (rule == null || !rule.ContainsKey("channels") || !rule.ContainsKey("displayValue"))
                    {
                        continue;
                    }

                    List<string> ruleChannels = ConvertToStringList(rule["channels"]);
                    if (ruleChannels.Any(targetChannels.Contains))
                    {
                        return NormalizeImagePath(rule["displayValue"] as string);
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private AppConfig NormalizeConfig(AppConfig cfg, string rawJson)
        {
            cfg = cfg ?? new AppConfig();
            cfg.priorityChannels = cfg.priorityChannels ?? new List<string>();
            cfg.primaryImagePath = NormalizeImagePath(cfg.primaryImagePath);
            cfg.secondaryImagePath = NormalizeImagePath(cfg.secondaryImagePath);
            cfg.priorityImagePath = NormalizeImagePath(cfg.priorityImagePath);
            cfg.priorityGifPath = NormalizeImagePath(cfg.priorityGifPath);

            if (string.IsNullOrWhiteSpace(cfg.primaryImagePath) && !string.IsNullOrWhiteSpace(rawJson))
            {
                cfg.primaryImagePath = GetRuleDisplayValue(rawJson, new[] { cfg.primaryChannel });
            }
            if (string.IsNullOrWhiteSpace(cfg.secondaryImagePath) && !string.IsNullOrWhiteSpace(rawJson))
            {
                cfg.secondaryImagePath = GetRuleDisplayValue(rawJson, new[] { cfg.secondaryChannel });
            }
            if (string.IsNullOrWhiteSpace(cfg.priorityImagePath) && !string.IsNullOrWhiteSpace(rawJson))
            {
                cfg.priorityImagePath = GetRuleDisplayValue(rawJson, cfg.priorityChannels);
            }

            if (string.IsNullOrWhiteSpace(cfg.priorityImagePath) && !string.IsNullOrWhiteSpace(cfg.priorityGifPath))
            {
                cfg.priorityImagePath = cfg.priorityGifPath;
            }
            if (string.IsNullOrWhiteSpace(cfg.priorityGifPath) && !string.IsNullOrWhiteSpace(cfg.priorityImagePath))
            {
                cfg.priorityGifPath = cfg.priorityImagePath;
            }

            return cfg;
        }

        private void BeginUpdateCheck(bool force)
        {
            if (!force && Interlocked.Exchange(ref updateCheckQueued, 1) == 1)
            {
                return;
            }

            ThreadPool.QueueUserWorkItem(_ =>
            {
                if (force)
                {
                    Interlocked.Exchange(ref updateCheckQueued, 1);
                }

                try
                {
                    CheckForUpdates();
                }
                finally
                {
                    Interlocked.Exchange(ref updateCheckQueued, 0);
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
                    string manifestJson = client.DownloadString(UpdateManifestUrl + "?v=" + DateTime.UtcNow.Ticks);
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
                string tempRoot = Path.Combine(Path.GetTempPath(), "DestinyStatusDesktopUpdate", Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempRoot);

                string packagePath = Path.Combine(tempRoot, "DestinyStatusDesktop-package.zip");
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.UserAgent] = GitHubUserAgent;
                    client.DownloadFile(manifest.packageUrl + "?v=" + Uri.EscapeDataString(manifest.version), packagePath);
                }

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

                BeginInvoke((Action)(() =>
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

                    Process.Start(startInfo);
                    Close();
                }));
            }
            catch
            {
                updateInProgress = false;
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
            try
            {
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var cfg = NormalizeConfig(serializer.Deserialize<AppConfig>(json) ?? new AppConfig(), json);
                    File.WriteAllText(configPath, serializer.Serialize(cfg));
                    return cfg;
                }
            }
            catch
            {
            }

            try
            {
                string legacyConfig = Path.Combine(legacyDataDir, "config.json");
                if (File.Exists(legacyConfig))
                {
                    string json = File.ReadAllText(legacyConfig);
                    var cfg = NormalizeConfig(serializer.Deserialize<AppConfig>(json) ?? new AppConfig(), json);
                    File.WriteAllText(configPath, serializer.Serialize(cfg));
                    return cfg;
                }
            }
            catch
            {
            }

            try
            {
                string legacyConfig = Path.Combine(appDir, "config.json");
                if (File.Exists(legacyConfig))
                {
                    string json = File.ReadAllText(legacyConfig);
                    var cfg = NormalizeConfig(serializer.Deserialize<AppConfig>(json) ?? new AppConfig(), json);
                    File.WriteAllText(configPath, serializer.Serialize(cfg));
                    return cfg;
                }
            }
            catch
            {
            }

            return new AppConfig();
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

            try
            {
                string legacyStatePath = Path.Combine(legacyDataDir, "state.json");
                if (File.Exists(legacyStatePath))
                {
                    return serializer.Deserialize<StateFile>(File.ReadAllText(legacyStatePath)) ?? new StateFile();
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
                state.layouts[GetDisplayLayoutKey()] = new WindowState { x = Location.X, y = Location.Y };
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

        private static string GetChannelState(string normalizedText)
        {
            if (string.IsNullOrWhiteSpace(normalizedText))
            {
                return "Unknown";
            }
            if (normalizedText.IndexOf("is offline", StringComparison.OrdinalIgnoreCase) >= 0 ||
                normalizedText.IndexOf("Offline", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Offline";
            }
            return "Online";
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
    }

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            NativeMethods.SetCurrentProcessExplicitAppUserModelID("matthiasfan55-oss.DestinyStatusDesktop");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new StatusForm());
        }
    }
}
