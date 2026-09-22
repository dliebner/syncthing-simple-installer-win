// SyncthingMonitor.cs
// Minimal system-tray indicator for Syncthing ("Syncthing Monitor").
//   Windows folder icon + GREEN badge = Syncthing is running and answering.
//   Windows folder icon + RED badge   = Syncthing is not running / not answering.
// Right-click menu (one item, changes with state):
//   "Sync Now"        when running     -> rescans all folders
//   "Start Syncthing" when not running -> starts it via the scheduled task
// No window. No Exit item for the end user. (Optional hidden Shift+Exit, below.)
//
// Install-SyncthingTray.ps1 compiles this into SyncthingMonitor.exe next to
// syncthing.exe using the C# compiler that ships with Windows (.NET Framework),
// and registers a logon task that runs it. Being a real Windows program rather
// than a script, it needs no interpreter, no execution-policy bypass and shows
// no console window; allow-list antivirus only ever has to allow this one exe.
//
// Usage:
//   SyncthingMonitor.exe                      run the tray icon (normal use)
//   SyncthingMonitor.exe --export-icon <path> write the shortcut icon (badge half
//                                             green, half red) as a multi-size .ico
//                                             and exit (used by the installer)
//
// Written for the C# 5 compiler in the .NET Framework, so it stays deliberately
// old-fashioned: no string interpolation, no ?. operator, no expression bodies.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace SyncthingMonitor
{
    // --------------------------- CONFIG (edit if needed) ---------------------------
    static class Config
    {
        // Folder that contains config.xml (Syncthing's home for this user).
        public static readonly string SyncthingHome =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Syncthing");

        // Path to syncthing.exe (only used as a fallback to read the API key).
        // The installer puts this program next to syncthing.exe; fall back to the
        // default install location if it isn't there.
        public static string SyncthingExe
        {
            get
            {
                string beside = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "syncthing.exe");
                if (File.Exists(beside)) return beside;
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    @"Programs\Syncthing\syncthing.exe");
            }
        }

        // Scheduled task that starts Syncthing (as created by Install-Syncthing.ps1).
        public static readonly string TaskName = "Syncthing - Start at Logon (" + Environment.UserName + ")";

        // How often to check, in seconds.
        public const int PollSeconds = 15;

        // Pop a small notification when it goes down (what actually reaches the end user).
        public const bool NotifyOnDown = true;

        // Hidden Exit for admins: hold Shift while right-clicking to reveal an Exit item.
        // Invisible to the end user on a normal right-click. Set false to remove entirely.
        public const bool EnableShiftExit = true;

        // Badge colors.
        public static readonly Color ColorUp   = Color.FromArgb(40, 170, 70);
        public static readonly Color ColorDown = Color.FromArgb(200, 60, 60);

        // Single-instance mutex, and the named event the installer/uninstaller signal
        // to ask a running instance to exit cleanly (so its icon is removed, not left
        // as a ghost). The event name must match Install-SyncthingTray.ps1.
        public static readonly string MutexName     = "SyncthingTrayMonitor_" + Environment.UserName;
        public static readonly string ExitEventName = "SyncthingTrayMonitor_Exit_" + Environment.UserName;
    }
    // -------------------------------------------------------------------------------

    static class Program
    {
        // Kept in a static so the GC can never collect (and thereby release) it.
        static Mutex instanceMutex;

        [STAThread]
        static int Main(string[] args)
        {
            // --- Export mode: write the shortcut icon and leave ---
            // The shortcut badge is half green / half red on purpose: a solid green
            // one on the desktop would read as "Syncthing is running" to someone who
            // only knows the tray icon. Half-and-half says "this shows green or red".
            if (args.Length == 2 && args[0] == "--export-icon")
            {
                try
                {
                    int[] sizes = { 16, 32, 48, 256 };
                    Bitmap[] bitmaps = new Bitmap[sizes.Length];
                    for (int i = 0; i < sizes.Length; i++)
                        bitmaps[i] = IconFactory.NewStatusBitmap(sizes[i], Config.ColorUp, Config.ColorDown);
                    IconFactory.WriteIco(args[1], bitmaps);
                    foreach (Bitmap b in bitmaps) b.Dispose();
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Couldn't write icon: " + ex.Message);
                    return 1;
                }
            }

            // --- Single instance: if one is already running for this user, quietly exit. ---
            // Makes it safe to re-launch (e.g. from the Start-menu shortcut) if the icon
            // ever goes missing: a live instance blocks a duplicate, a dead one gets replaced.
            bool createdNew;
            instanceMutex = new Mutex(true, Config.MutexName, out createdNew);
            if (!createdNew) return 0;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApp());
            GC.KeepAlive(instanceMutex);
            return 0;
        }
    }

    class TrayApp : ApplicationContext
    {
        // Talks to Syncthing over its local API. Self-signed certs (GUI TLS on) are
        // accepted, same as "curl -k".
        string baseUrl = "http://127.0.0.1:8384";
        string[] fallbackUrls = new string[0];   // other forms of baseUrl to try if that doesn't answer
        string apiKey  = "";

        NotifyIcon notify;
        ContextMenuStrip menu;
        System.Windows.Forms.Timer timer;
        System.Windows.Forms.Timer controlTimer;
        EventWaitHandle exitEvent;
        Icon iconUp, iconDown;
        Icon registeredIcon;   // the icon the shell has on file for our tray entry (see ShowToast)

        // State
        bool?    lastUp      = null;                 // last confirmed up/down
        bool?    menuState   = null;                 // what the menu currently reflects
        int      missCount   = 0;                    // consecutive failed health checks
        DateTime lastBalloon = DateTime.MinValue;    // for the "stopped" balloon cooldown
        DateTime fastUntil   = DateTime.MinValue;    // end of any fast-poll burst
        readonly int normalInterval = Math.Max(3, Config.PollSeconds) * 1000;

        public TrayApp()
        {
            ServicePointManager.ServerCertificateValidationCallback =
                delegate { return true; };
            // Syncthing's GUI needs TLS 1.2 when its TLS is on, and older .NET Framework
            // defaults don't offer it. Numeric values (Tls12 | Tls11) so this compiles
            // against any 4.x.
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072 | (SecurityProtocolType)768; }
            catch (NotSupportedException) { }

            // --- Stop request from the installer/uninstaller (see Config.ExitEventName) ---
            exitEvent = new EventWaitHandle(false, EventResetMode.ManualReset, Config.ExitEventName);
            exitEvent.Reset();   // ignore any stale signal left over from a previous stop

            ReadConfig();

            // --- One-time sanity checks, so a misconfiguration surfaces immediately ---
            // (rather than only when someone clicks a button that silently does nothing).
            string warnings = "";
            if (apiKey.Length == 0)
                warnings += "Couldn't read the Syncthing API key, so 'Sync Now' won't work. ";
            if (RunHidden("schtasks.exe", "/Query /TN \"" + Config.TaskName + "\"") != 0)
                warnings += "Scheduled task '" + Config.TaskName + "' not found, so 'Start Syncthing' won't work.";

            iconUp   = IconFactory.NewStatusIcon(Config.ColorUp);
            iconDown = IconFactory.NewStatusIcon(Config.ColorDown);

            menu = new ContextMenuStrip();
            if (Config.EnableShiftExit) menu.Opening += OnMenuOpening;

            notify = new NotifyIcon();
            // Register with the real state rather than a placeholder: the shell keeps
            // the icon it sees here for toasts (see ShowToast), and this also avoids
            // a red flash at startup.
            notify.Icon = TestUp() ? iconUp : iconDown;
            notify.Text = "Syncthing: checking...";
            notify.ContextMenuStrip = menu;
            notify.Visible = true;
            registeredIcon = notify.Icon;

            timer = new System.Windows.Forms.Timer();
            timer.Interval = normalInterval;
            timer.Tick += delegate { try { UpdateStatus(); } catch { } };
            timer.Start();

            // Watches for a stop request from the installer/uninstaller.
            controlTimer = new System.Windows.Forms.Timer();
            controlTimer.Interval = 500;
            controlTimer.Tick += delegate { if (exitEvent.WaitOne(0)) StopTray(); };
            controlTimer.Start();

            // First check, guarded so a startup hiccup can't kill the app before the loop runs.
            try { UpdateStatus(); } catch { }

            // Surface any config problems found at startup, once.
            if (warnings.Length > 0)
                ShowToast(8000, "Syncthing Monitor: check setup", warnings.Trim(), ToolTipIcon.Warning);
        }

        // --- Read address, scheme, and API key from config.xml (best-effort) ---
        // Defaults match a fresh Install-Syncthing.ps1 setup (GUI TLS off); config.xml
        // wins whenever it can be read.
        void ReadConfig()
        {
            string address = "127.0.0.1:8384";
            string scheme  = "http";
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(Path.Combine(Config.SyncthingHome, "config.xml"));
                XmlNode gui = doc.SelectSingleNode("/configuration/gui");
                if (gui != null)
                {
                    XmlNode n = gui.SelectSingleNode("address");
                    if (n != null && n.InnerText.Trim().Length > 0) address = n.InnerText.Trim();
                    XmlAttribute tls = gui.Attributes["tls"];
                    if (tls != null) scheme = (tls.Value == "true") ? "https" : "http";
                    n = gui.SelectSingleNode("apikey");
                    if (n != null) apiKey = n.InnerText.Trim();
                }
            }
            catch { }
            if (apiKey.Length == 0 && File.Exists(Config.SyncthingExe))
            {
                try { apiKey = RunHiddenCapture(Config.SyncthingExe, "cli config gui apikey get").Trim(); }
                catch { }
            }
            baseUrl = scheme + "://" + address;

            // Other URLs to try when the configured one doesn't answer:
            // - Loopback on the same port. The GUI address is a *listen* address:
            //   "0.0.0.0:8384" (all interfaces, the usual setting for LAN access) or
            //   "[::]:8384" can't be *connected* to on Windows, and "localhost" may
            //   resolve to ::1 first. A browser gets away with that; HttpWebRequest
            //   doesn't. Loopback is what they all mean for a check from this machine.
            // - The other scheme. config.xml is only read at startup, so a TLS toggle
            //   in the GUI would otherwise read as "down" until the tray is restarted.
            string host = address, port = "8384";
            int colon = address.LastIndexOf(':');
            if (colon >= 0 && address.IndexOf(']') < colon)
            {
                host = address.Substring(0, colon);
                port = address.Substring(colon + 1);
            }
            host = host.Trim('[', ']').Trim();
            port = port.Trim();
            if (port.Length == 0) port = "8384";
            string other = (scheme == "https") ? "http" : "https";
            List<string> alts = new List<string>();
            if (host != "127.0.0.1") alts.Add(scheme + "://127.0.0.1:" + port);
            alts.Add(other + "://" + address);
            if (host != "127.0.0.1") alts.Add(other + "://127.0.0.1:" + port);
            fallbackUrls = alts.ToArray();
        }

        // All API calls go through here. No system proxy: a configured proxy or VPN
        // client would otherwise be handed our loopback requests and drop them.
        // Browsers bypass the proxy for local addresses; HttpWebRequest does not.
        HttpWebRequest NewRequest(string url, int timeoutMs)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Proxy = null;
            req.Timeout = timeoutMs;
            req.ReadWriteTimeout = timeoutMs;
            return req;
        }

        bool Probe(string url)
        {
            try
            {
                HttpWebRequest req = NewRequest(url + "/rest/noauth/health", 3000);
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader sr = new StreamReader(resp.GetResponseStream()))
                {
                    string body = sr.ReadToEnd();
                    return Regex.IsMatch(body, "\"status\"\\s*:\\s*\"OK\"");
                }
            }
            catch { return false; }
        }

        bool TestUp()
        {
            if (Probe(baseUrl)) return true;
            for (int i = 0; i < fallbackUrls.Length; i++)
            {
                if (!Probe(fallbackUrls[i])) continue;
                // That one answers: use it from now on (for Sync Now as well), and
                // keep the old one as a fallback in case things change back.
                string old = baseUrl;
                baseUrl = fallbackUrls[i];
                fallbackUrls[i] = old;
                return true;
            }
            return false;
        }

        void StartFastPoll()
        {
            // Poll every second for up to 30s so the badge reacts within ~1s of a change.
            fastUntil = DateTime.Now.AddSeconds(30);
            timer.Interval = 1000;
        }

        void DoStart()
        {
            if (RunHidden("schtasks.exe", "/Run /TN \"" + Config.TaskName + "\"") != 0)
            {
                ShowToast(4000, "Syncthing",
                    "Couldn't start Syncthing: the scheduled task '" + Config.TaskName + "' wasn't found or couldn't run.",
                    ToolTipIcon.Warning);
                return;
            }
            notify.Text = "Syncthing: starting...";
            StartFastPoll();
        }

        void DoSyncNow()
        {
            bool ok = false;
            try
            {
                HttpWebRequest req = NewRequest(baseUrl + "/rest/db/scan", 10000);
                req.Method = "POST";
                req.ContentLength = 0;
                req.Headers["X-API-Key"] = apiKey;
                // GetResponse throws on HTTP errors (e.g. 403 from a bad API key), so an
                // auth failure isn't reported as a successful sync.
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                {
                    ok = (int)resp.StatusCode >= 200 && (int)resp.StatusCode < 300;
                }
            }
            catch { ok = false; }

            if (ok)
            {
                ShowToast(2000, "Syncthing", "Sync started.", ToolTipIcon.Info);
            }
            else
            {
                string why = (apiKey.Length == 0) ? "no API key was found." : "Syncthing may have stopped.";
                ShowToast(3000, "Syncthing", "Couldn't trigger a sync: " + why, ToolTipIcon.Warning);
                StartFastPoll();   // confirm red quickly if it really has gone down
            }
        }

        // Balloon tips are toasts on Windows 10/11, and the picture on a toast is the
        // icon the shell had when the tray entry was *registered*, not the current
        // one: after a status change, toasts would keep showing the old badge. If the
        // badge has changed since registration, re-register first so they match.
        void ShowToast(int timeoutMs, string title, string text, ToolTipIcon kind)
        {
            if (notify.Icon != registeredIcon)
            {
                notify.Visible = false;
                notify.Visible = true;
                registeredIcon = notify.Icon;
            }
            notify.ShowBalloonTip(timeoutMs, title, text, kind);
        }

        void SetMenu(bool up)
        {
            menu.Items.Clear();
            if (up)
            {
                ToolStripItem item = menu.Items.Add("Sync Now");
                item.Click += delegate { DoSyncNow(); };
            }
            else
            {
                ToolStripItem item = menu.Items.Add("Start Syncthing");
                item.Click += delegate { DoStart(); };
            }
        }

        void UpdateStatus()
        {
            bool probe = TestUp();
            if (probe) missCount = 0; else missCount++;

            // Debounce: require two consecutive misses before declaring "down". A single
            // blip (e.g. the laptop just woke from sleep) shouldn't trigger a false alarm.
            bool up;
            if (probe)                                   up = true;
            else if (missCount >= 2 || lastUp != true)   up = false;
            else                                         up = true;   // one miss while up: hold for now

            bool fast = DateTime.Now < fastUntil;

            // Only touch the tray icon when the state actually changes.
            if (lastUp != up) notify.Icon = up ? iconUp : iconDown;

            if (up)        notify.Text = "Syncthing: running";
            else if (fast) notify.Text = "Syncthing: starting...";
            else           notify.Text = "Syncthing: NOT running";

            // Rebuild the menu when needed, but never while it's open under the mouse.
            if (menuState != up && !menu.Visible)
            {
                SetMenu(up);
                menuState = up;
            }

            // Balloon on a confirmed up->down transition, with a cooldown to avoid spam
            // if Syncthing is flapping.
            if (Config.NotifyOnDown && lastUp == true && !up)
            {
                if ((DateTime.Now - lastBalloon).TotalSeconds >= 120)
                {
                    ShowToast(5000, "Syncthing stopped",
                        "File sync is not running. Right-click the tray icon and choose 'Start Syncthing'.",
                        ToolTipIcon.Warning);
                    lastBalloon = DateTime.Now;
                }
            }

            lastUp = up;

            // Leave any fast-poll burst once it's up, or once the 30s window expires.
            if (up || !fast)
            {
                fastUntil = DateTime.MinValue;
                if (timer.Interval != normalInterval) timer.Interval = normalInterval;
            }
        }

        void OnMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            for (int i = menu.Items.Count - 1; i >= 0; i--)
                if (menu.Items[i].Text == "Exit") menu.Items.RemoveAt(i);
            if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                ToolStripItem exit = menu.Items.Add("Exit");
                exit.Click += delegate { StopTray(); };
            }
        }

        void StopTray()
        {
            // Clean shutdown: stop polling, remove the icon, leave the message loop.
            timer.Stop();
            controlTimer.Stop();
            notify.Visible = false;
            notify.Dispose();
            ExitThread();
        }

        // Runs a console program with no window; returns its exit code (-1 if it couldn't start).
        static int RunHidden(string file, string args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(file, args);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                using (Process p = Process.Start(psi))
                {
                    p.WaitForExit(15000);
                    return p.HasExited ? p.ExitCode : -1;
                }
            }
            catch { return -1; }
        }

        static string RunHiddenCapture(string file, string args)
        {
            ProcessStartInfo psi = new ProcessStartInfo(file, args);
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            using (Process p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit(15000);
                return output;
            }
        }
    }

    // ============================== ICON DRAWING ==================================
    // Shared by the tray icon and by --export-icon, so the shortcut icon the installer
    // creates is drawn by exactly the same code as what sits in the tray (only the
    // badge colouring differs).
    static class IconFactory
    {
        // The stock Windows folder icon as a bitmap of the requested size.
        // Prefers an exact-size render from shell32.dll (crisp at 16..256); falls back
        // to the shell's 32px "folder" icon, then to null (caller draws one).
        static Bitmap GetFolderBitmap(int size)
        {
            try
            {
                IntPtr[] handles = new IntPtr[1];
                uint[]   ids     = new uint[1];
                uint n = NativeMethods.PrivateExtractIcons("shell32.dll", 3, size, size, handles, ids, 1, 0);
                if (n >= 1 && handles[0] != IntPtr.Zero)
                {
                    using (Icon ficon = Icon.FromHandle(handles[0]))
                    {
                        Bitmap bmp = ficon.ToBitmap();
                        NativeMethods.DestroyIcon(handles[0]);
                        return bmp;
                    }
                }
            }
            catch { }
            try
            {
                NativeMethods.SHFILEINFO info = new NativeMethods.SHFILEINFO();
                const uint FILE_ATTRIBUTE_DIRECTORY = 0x10, SHGFI_ICON = 0x100, SHGFI_USEFILEATTRIBUTES = 0x10;
                NativeMethods.SHGetFileInfo("folder", FILE_ATTRIBUTE_DIRECTORY, ref info,
                    (uint)Marshal.SizeOf(info), SHGFI_ICON | SHGFI_USEFILEATTRIBUTES);
                if (info.hIcon != IntPtr.Zero)
                {
                    using (Icon ficon = Icon.FromHandle(info.hIcon))
                    {
                        Bitmap bmp = ficon.ToBitmap();
                        NativeMethods.DestroyIcon(info.hIcon);
                        return bmp;
                    }
                }
            }
            catch { }
            return null;
        }

        // Folder icon with a colored status badge in the top-right, at any size.
        // Badge geometry is proportional to the 32px original (14px dot, 2px halo, 2px inset).
        public static Bitmap NewStatusBitmap(int size, Color dot)
        {
            return NewStatusBitmap(size, dot, dot);
        }

        // Same, with the dot split down the middle: left half in one color, right
        // half in the other. Used for the shortcut icon.
        public static Bitmap NewStatusBitmap(int size, Color dotLeft, Color dotRight)
        {
            Bitmap bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode     = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.Clear(Color.Transparent);

                Bitmap folder = GetFolderBitmap(size);
                if (folder != null)
                {
                    g.DrawImage(folder, 0, 0, size, size);
                    folder.Dispose();
                }
                else
                {
                    // Fallback: a simple drawn folder if no shell icon was available.
                    using (Brush body = new SolidBrush(Color.FromArgb(255, 222, 184, 100)))
                    {
                        g.FillRectangle(body, size * 3 / 32, size * 12 / 32, size * 26 / 32, size * 16 / 32);
                        g.FillRectangle(body, size * 3 / 32, size *  8 / 32, size * 12 / 32, size *  6 / 32);
                    }
                }

                // Status badge: white halo + colored dot, tucked into the top-right corner.
                int bd     = Math.Max(5, (int)Math.Round(size * 14.0 / 32));
                int ringW  = Math.Max(1, (int)Math.Round(size *  2.0 / 32));
                int inset  = Math.Max(1, (int)Math.Round(size *  2.0 / 32));
                int outerD = bd + 2 * ringW;
                int cx     = size - bd / 2 - inset;
                int cy     = bd / 2 + inset;
                using (Brush white = new SolidBrush(Color.White))
                using (Brush left  = new SolidBrush(dotLeft))
                using (Brush right = new SolidBrush(dotRight))
                {
                    g.FillEllipse(white, cx - outerD / 2, cy - outerD / 2, outerD, outerD);
                    g.FillEllipse(left, cx - bd / 2, cy - bd / 2, bd, bd);
                    if (dotRight != dotLeft)
                    {
                        // Paint the right half over the top, clipped to a rectangle: no
                        // anti-aliased seam, unlike two half-pies meeting in the middle.
                        g.SetClip(new Rectangle(cx, cy - outerD / 2, outerD, outerD));
                        g.FillEllipse(right, cx - bd / 2, cy - bd / 2, bd, bd);
                        g.ResetClip();
                    }
                }
            }
            return bmp;
        }

        public static Icon NewStatusIcon(Color dot)
        {
            // GetHicon() returns an independent HICON. The two status icons live for
            // the whole process, so the handle is simply never destroyed.
            using (Bitmap bmp = NewStatusBitmap(32, dot))
            {
                return Icon.FromHandle(bmp.GetHicon());
            }
        }

        // Writes bitmaps to a classic .ico file (32-bit BGRA entries with an AND mask
        // derived from alpha), which Explorer renders correctly at every size.
        public static void WriteIco(string path, Bitmap[] bitmaps)
        {
            byte[][] entries = new byte[bitmaps.Length][];
            for (int i = 0; i < bitmaps.Length; i++)
            {
                Bitmap bmp = bitmaps[i];
                int w = bmp.Width, h = bmp.Height;
                BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                int stride = data.Stride;
                byte[] src = new byte[stride * h];
                Marshal.Copy(data.Scan0, src, 0, src.Length);
                bmp.UnlockBits(data);

                int rowBytes   = w * 4;
                int maskStride = ((w + 31) / 32) * 4;
                byte[] xor = new byte[rowBytes * h];       // BGRA, bottom-up
                byte[] and = new byte[maskStride * h];     // 1bpp, bottom-up, 1 = transparent
                for (int y = 0; y < h; y++)
                {
                    int srcRow = y * stride;
                    int dstY   = h - 1 - y;
                    Array.Copy(src, srcRow, xor, dstY * rowBytes, rowBytes);
                    for (int x = 0; x < w; x++)
                        if (src[srcRow + x * 4 + 3] == 0)
                            and[dstY * maskStride + (x >> 3)] |= (byte)(0x80 >> (x & 7));
                }

                using (MemoryStream ms = new MemoryStream())
                using (BinaryWriter bw = new BinaryWriter(ms))
                {
                    // BITMAPINFOHEADER (height is doubled: XOR + AND masks)
                    bw.Write(40); bw.Write(w); bw.Write(h * 2);
                    bw.Write((ushort)1); bw.Write((ushort)32); bw.Write((uint)0);
                    bw.Write((uint)(xor.Length + and.Length));
                    bw.Write(0); bw.Write(0); bw.Write((uint)0); bw.Write((uint)0);
                    bw.Write(xor); bw.Write(and);
                    bw.Flush();
                    entries[i] = ms.ToArray();
                }
            }

            using (FileStream fs = File.Create(path))
            using (BinaryWriter bw = new BinaryWriter(fs))
            {
                // ICONDIR
                bw.Write((ushort)0); bw.Write((ushort)1); bw.Write((ushort)entries.Length);
                uint offset = (uint)(6 + 16 * entries.Length);
                for (int i = 0; i < entries.Length; i++)
                {
                    int w = bitmaps[i].Width, h = bitmaps[i].Height;
                    // ICONDIRENTRY (0 means 256)
                    bw.Write((byte)(w >= 256 ? 0 : w));
                    bw.Write((byte)(h >= 256 ? 0 : h));
                    bw.Write((byte)0); bw.Write((byte)0);
                    bw.Write((ushort)1); bw.Write((ushort)32);
                    bw.Write((uint)entries[i].Length); bw.Write(offset);
                    offset += (uint)entries[i].Length;
                }
                for (int i = 0; i < entries.Length; i++) bw.Write(entries[i]);
            }
        }
    }

    static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int    iIcon;
            public uint   dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]  public string szTypeName;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        // Extracts an icon from a DLL rendered at an exact pixel size (up to 256).
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern uint PrivateExtractIcons(string lpszFile, int nIconIndex, int cxIcon, int cyIcon, IntPtr[] phicon, uint[] piconid, uint nIcons, uint flags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);
    }
}
