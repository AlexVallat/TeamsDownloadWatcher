using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using TeamsDownloadWatcher.Properties;

namespace TeamsDownloadWatcher
{
	static class Program
    {
        private const string AppName = "TeamsDownloadWatcher";

        private static readonly Guid DownloadsFolderRfid = new Guid("374DE290-123F-4565-9164-39C4925E467B");
		[DllImport("Shell32.dll")]
        private static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);
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

                _fileSystemWatcher = new FileSystemWatcher(GetDownloadsFolder(), "????????-????-????-????-????????????.tmp") {NotifyFilter = NotifyFilters.FileName};

                _fileSystemWatcher.Renamed += OnNewDownload;
                _fileSystemWatcher.EnableRaisingEvents = true;

                _notifyIcon.Click += OnIconClick;
            }

            private static string GetDownloadsFolder()
			{
                const uint DontVerify = 0x4000;
                var hResult = SHGetKnownFolderPath(DownloadsFolderRfid, DontVerify, IntPtr.Zero, out var ppszPath);
                if (hResult < 0)
				{
                    throw new ExternalException("Could not obtain download folder", hResult);
				}

                var result = Marshal.PtrToStringUni(ppszPath);
                Marshal.FreeCoTaskMem(ppszPath);
                return result;
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
                        Settings.Default.DownloadLocation = folderBrowser.SelectedPath;
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

            private static void OnNewDownload(object sender, RenamedEventArgs e)
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
