using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using TeamsDownloadWatcher.Properties;

namespace TeamsDownloadWatcher
{
	static class Program
    {
        private const string AppName = "TeamsDownloadWatcher";

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new WatcherContext());
        }

        class WatcherContext : ApplicationContext
        {
            private readonly NotifyIcon _notifyIcon;
            private readonly FileSystemWatcher _fileSystemWatcher;
            private static readonly TimeSpan WaitForFilePollInterval = TimeSpan.FromMilliseconds(100);

            public WatcherContext()
            {
                _notifyIcon = new NotifyIcon
                {
                    Icon = Resources.Icon,
                    Visible = true,
                    Text = Resources.ApplicationName,
                    ContextMenuStrip = new ContextMenuStrip
                    {
                        ShowCheckMargin = true,
                        ShowImageMargin = false,
                        Items =
                        {
                            new ToolStripMenuItem(Resources.SetDownloadLocationMenuItem, null, OnSetDownloadLocation) { ToolTipText = Settings.Default.DownloadLocation },
                            new ToolStripMenuItem(Resources.OpenImmediatelyMenuItem, null, OnChangeOpenImmediately) { Checked = Settings.Default.OpenImmediately, CheckOnClick = true },
                            new ToolStripMenuItem(Resources.StartWithWindowsMenuItem, null, OnChangeStartWithWindows) { Checked = StartWithWindows, CheckOnClick = true },
                            new ToolStripSeparator(),
                            new ToolStripMenuItem(Resources.ExitMenuItem, null, OnExitClicked)
                        }
                    }
                };

                _fileSystemWatcher = new FileSystemWatcher(GetDownloadsFolder()) {NotifyFilter = NotifyFilters.FileName};

                _fileSystemWatcher.Created += OnNewDownload;
                _fileSystemWatcher.EnableRaisingEvents = true;

                _notifyIcon.Click += OnIconClick;
            }

            private static string GetDownloadsFolder()
			{
                return Path.Combine(Path.GetTempPath(), "TeamsDownload");
            }

            private bool StartWithWindows
            {
                set
                {
                    using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                    {
                        if (value)
                        {
                            key.SetValue(AppName, "\"" + Assembly.GetExecutingAssembly().Location + "\"");
                        }
                        else
                        {
                            key.DeleteValue(AppName, false);
                        }
                    }
                }
                get
                {
                    using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false))
                    {
                        return key.GetValue(AppName) != null;
                    }
                }
            }

            private void OnChangeStartWithWindows(object sender, EventArgs e)
            {
                StartWithWindows = ((ToolStripMenuItem)sender).Checked;
            }

            private void OnIconClick(object sender, EventArgs e)
            {
                // Hack to make context menu appear on left click too
                typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic)?.Invoke(_notifyIcon, null);
            }

            private void OnChangeOpenImmediately(object sender, EventArgs e)
            {
                Settings.Default.OpenImmediately = ((ToolStripMenuItem)sender).Checked;
                Settings.Default.Save();
            }

            private void OnSetDownloadLocation(object sender, EventArgs e)
            {
                using (var folderBrowser = new FolderBrowserDialog
                {
                    Description = Resources.SetDownloadLocationDescription,
                    SelectedPath = Settings.Default.DownloadLocation,
                })
                {
                    if (folderBrowser.ShowDialog() == DialogResult.OK)
                    {
                        if (folderBrowser.SelectedPath.Equals(GetDownloadsFolder(), StringComparison.OrdinalIgnoreCase))
                        {
                            Settings.Default.DownloadLocation = null;
                        }
                        else
                        {
                            Settings.Default.DownloadLocation = folderBrowser.SelectedPath;
                        }
                        Settings.Default.Save();

                        ((ToolStripMenuItem)sender).ToolTipText = Settings.Default.DownloadLocation;
                    }
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon.Dispose();
                    _fileSystemWatcher.Dispose();
                }

                base.Dispose(disposing);
            }

            private static void OnExitClicked(object sender, EventArgs e)
            {
                Application.Exit();
            }

            private static void OnNewDownload(object sender, FileSystemEventArgs e)
            {
                var destFileName = e.FullPath;
                var destPath = Settings.Default.DownloadLocation;
                if (!string.IsNullOrEmpty(destPath))
                {
                    try
                    {
                        Directory.CreateDirectory(destPath);

                        destFileName = Path.Combine(destPath, e.Name);

                        var i = 0;
                        while (File.Exists(destFileName))
                        {
                            destFileName = Path.Combine(destPath, $"{Path.GetFileNameWithoutExtension(e.Name)} ({++i}){Path.GetExtension(e.Name)}");
                        }

                        var canOpenFile = false;
                        do
                        {
                            Thread.Sleep(WaitForFilePollInterval);
                            try
                            {
                                // Check if we can get exclusive access to the file, to move it. If not, it may be still being downloaded
                                using (File.Open(e.FullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                                {
                                    canOpenFile = true;
                                }
                            }
							catch (IOException)
                            {
                                canOpenFile = false;
                            }
                        } while (!canOpenFile ||
                                 DateTime.UtcNow - File.GetLastWriteTimeUtc(e.FullPath) < WaitForFilePollInterval); // If it's been modified since we last checked, keep waiting until Teams has finished saving

                        File.Move(e.FullPath, destFileName);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine("Could not move file: " + ex.Message);
                    }
                }

                if (Settings.Default.OpenImmediately)
                {
					try
					{
						Process.Start(destFileName);
					}
					catch (Exception ex)
					{
                        Debug.WriteLine("Could not open file: " + ex.Message);
                    }
                }
            }
        }
    }
}
