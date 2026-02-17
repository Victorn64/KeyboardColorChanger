using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using HidSharp;
using Microsoft.Win32;

namespace keybord_color_changer
{
    public partial class Form1 : Form
    {
        // Константы устройства DEXP OMNI
        private const int VID = 0x320F;
        private const int PID = 0x505B;

        // WinAPI для работы с языками
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);
        [DllImport("user32.dll")]
        static extern IntPtr GetKeyboardLayout(uint threadId);
        [DllImport("user32.dll")]
        static extern int GetKeyboardLayoutList(int nBuff, [Out] IntPtr[] lpList);

        [StructLayout(LayoutKind.Sequential)]
        struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public System.Drawing.Rectangle rcCaret;
        }

        [DllImport("user32.dll")]
        static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private System.Windows.Forms.Timer monitorTimer;
        private Dictionary<string, Color> settings = new Dictionary<string, Color>();
        private string lastLang = null;

        private static readonly Dictionary<int, string> LANG_MAP = new Dictionary<int, string>
        {
            { 0x0419, "RU" },
            { 0x0409, "EN" },
            { 0x0422, "UA" },
            { 0x0407, "DE" },
            { 0x040c, "FR" },
            { 0x0410, "IT" },
            { 0x040a, "ES" }
        };

        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Visible = false;

            LoadSettings();
            SetupTray();

            monitorTimer = new System.Windows.Forms.Timer();
            monitorTimer.Interval = 100;
            monitorTimer.Tick += MonitorTimer_Tick;
            monitorTimer.Start();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.Hide(); // Дополнительная гарантия скрытия окна
        }

        private void SetupTray()
        {
            trayMenu = new ContextMenuStrip();
            UpdateTrayMenu();

            trayIcon = new NotifyIcon();
            trayIcon.Text = "DEXP OMNI White Color Changer";
            trayIcon.Icon = CreateTrayIcon();
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
        }

        private void UpdateTrayMenu()
        {
            trayMenu.Items.Clear();
            trayMenu.Items.Add(new ToolStripMenuItem("--- Настройка цветов ---") { Enabled = false });

            var installedLangs = GetInstalledLanguages();
            foreach (var lang in installedLangs)
            {
                Color c = settings.ContainsKey(lang) ? settings[lang] : Color.White;
                var item = new ToolStripMenuItem($"{lang}: {c.R} {c.G} {c.B}");
                item.Click += (s, e) => PickColor(lang);
                trayMenu.Items.Add(item);
            }

            trayMenu.Items.Add(new ToolStripSeparator());
            
            var autostartItem = new ToolStripMenuItem("Включить автозагрузку", null, OnAutostartClick);
            trayMenu.Items.Add(autostartItem);

            trayMenu.Items.Add(new ToolStripMenuItem("Выход", null, (s, e) => Application.Exit()));
        }

        private Icon CreateTrayIcon()
        {
            Bitmap bmp = new Bitmap(32, 32);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.FillEllipse(new SolidBrush(Color.FromArgb(30, 30, 30)), 2, 2, 28, 28);
                g.FillPie(Brushes.Red, 2, 2, 28, 28, 0, 120);
                g.FillPie(Brushes.Green, 2, 2, 28, 28, 120, 120);
                g.FillPie(Brushes.Blue, 2, 2, 28, 28, 240, 120);
            }
            return Icon.FromHandle(bmp.GetHicon());
        }

        private List<string> GetInstalledLanguages()
        {
            int count = GetKeyboardLayoutList(0, null);
            IntPtr[] buffer = new IntPtr[count];
            GetKeyboardLayoutList(count, buffer);
            
            List<string> langs = new List<string>();
            foreach (var hkl in buffer)
            {
                int langId = (int)hkl & 0xFFFF;
                string name = LANG_MAP.ContainsKey(langId) ? LANG_MAP[langId] : $"ID:{langId:X4}";
                if (!langs.Contains(name)) langs.Add(name);
            }
            return langs;
        }

        private string GetCurrentLanguage()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero) return "EN";

                uint threadId = 0;

                // Попробуем получить информацию о потоке с фокусом через GetGUIThreadInfo
                GUITHREADINFO guiInfo = new GUITHREADINFO();
                guiInfo.cbSize = Marshal.SizeOf(guiInfo);
                if (GetGUIThreadInfo(0, ref guiInfo) && guiInfo.hwndFocus != IntPtr.Zero)
                {
                    threadId = GetWindowThreadProcessId(guiInfo.hwndFocus, IntPtr.Zero);
                }
                else
                {
                    // Если не вышло (например, защищенное окно), берем поток активного окна
                    threadId = GetWindowThreadProcessId(hwnd, IntPtr.Zero);
                }

                if (threadId == 0) return "EN";

                IntPtr hkl = GetKeyboardLayout(threadId);
                int langId = (int)hkl & 0xFFFF;
                return LANG_MAP.ContainsKey(langId) ? LANG_MAP[langId] : "EN";
            }
            catch { return "EN"; }
        }

        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
            string currentLang = GetCurrentLanguage();
            if (currentLang != lastLang)
            {
                if (settings.ContainsKey(currentLang))
                {
                    Color c = settings[currentLang];
                    SetKeyboardColor(c.R, c.G, c.B);
                    lastLang = currentLang;
                }
            }
        }

        private void SetKeyboardColor(byte r, byte g, byte b)
        {
            Task.Run(() =>
            {
                try
                {
                    var loader = DeviceList.Local;
                    // Получаем все HID устройства с нашими VID/PID
                    var devices = loader.GetHidDevices(VID, PID).ToList();
                    
                    HidDevice targetDevice = null;

                    foreach (var dev in devices)
                    {
                        try
                        {
                            // Проверяем UsagePage. 
                            // В HidSharp это можно достать через дескрипторы или свойства
                            // 65308 (0xFF1C) - это то, что искал Python скрипт
                            if (dev.ToString().Contains("UsagePage=0xFF1C") || dev.ToString().Contains("UsagePage=65308"))
                            {
                                targetDevice = dev;
                                break;
                            }
                            
                            // Запасной вариант: проверка по длине репорта (обычно 64 или 65)
                            if (targetDevice == null && dev.GetMaxOutputReportLength() >= 64)
                            {
                                targetDevice = dev;
                            }
                        }
                        catch { }
                    }

                    // Если совсем не нашли по признакам, пробуем 3-е устройство (как в kbd_devs[2])
                    if (targetDevice == null && devices.Count > 2) targetDevice = devices[2];
                    if (targetDevice == null && devices.Count > 0) targetDevice = devices[0];

                    if (targetDevice != null && targetDevice.TryOpen(out HidStream stream))
                    {
                        using (stream)
                        {
                            int reportLength = targetDevice.GetMaxOutputReportLength();
                            
                            // ВАЖНО: В Python было device.write([0x04, 0x01, 0x01, ...])
                            // В Windows/HidSharp первый байт (index 0) — это Report ID.
                            // Если команда начинается с 0x04, значит Report ID = 0x04.
                            
                            byte[] cmd_s = new byte[reportLength];
                            cmd_s[0] = 0x04; // Report ID = 0x04
                            cmd_s[1] = 0x01; cmd_s[2] = 0x01;
                            cmd_s[7] = r; cmd_s[8] = g; cmd_s[9] = b;
                            stream.Write(cmd_s);
                            
                            Thread.Sleep(20);

                            byte[] cmd1 = new byte[reportLength];
                            cmd1[0] = 0x04; cmd1[1] = 0x01; cmd1[3] = 0x01;
                            stream.Write(cmd1);

                            byte[] cmd2 = new byte[reportLength];
                            cmd2[0] = 0x04; cmd2[1] = 0x29; cmd2[2] = 0x02; cmd2[3] = 0x06; 
                            cmd2[4] = 0x22; cmd2[9] = 0x06; cmd2[10] = 0x04; cmd2[11] = 0x02;
                            cmd2[14] = r; cmd2[15] = g; cmd2[16] = b;
                            stream.Write(cmd2);

                            byte[] cmd3 = new byte[reportLength];
                            cmd3[0] = 0x04; cmd3[1] = 0x02; cmd3[3] = 0x02;
                            stream.Write(cmd3);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Временно для отладки можно раскомментировать
                    // MessageBox.Show("Error: " + ex.Message);
                }
            });
        }

        private void PickColor(string lang)
        {
            using (ColorDialog cd = new ColorDialog())
            {
                if (settings.ContainsKey(lang)) cd.Color = settings[lang];
                if (cd.ShowDialog() == DialogResult.OK)
                {
                    settings[lang] = cd.Color;
                    SaveSettings();
                    UpdateTrayMenu();
                    if (GetCurrentLanguage() == lang) SetKeyboardColor(cd.Color.R, cd.Color.G, cd.Color.B);
                }
            }
        }

        private void LoadSettings()
        {
            try
            {
                foreach (var l in GetInstalledLanguages()) settings[l] = Color.White;

                string saved = Properties.Settings.Default.ColorSettings;
                if (!string.IsNullOrEmpty(saved))
                {
                    var pairs = saved.Split('|');
                    foreach (var pair in pairs)
                    {
                        var parts = pair.Split(':');
                        if (parts.Length == 2)
                        {
                            string lang = parts[0].ToUpper();
                            var rgb = parts[1].Split(',');
                            if (rgb.Length == 3)
                            {
                                settings[lang] = Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]));
                            }
                        }
                    }
                }
                else
                {
                    // МИГРАЦИЯ из settings.txt
                    string oldPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.txt");
                    if (File.Exists(oldPath))
                    {
                        foreach (var line in File.ReadAllLines(oldPath))
                        {
                            var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            if (parts.Length >= 4)
                            {
                                settings[parts[0].ToUpper()] = Color.FromArgb(int.Parse(parts[1]), int.Parse(parts[2]), int.Parse(parts[3]));
                            }
                        }
                        SaveSettings();
                    }
                }
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                List<string> pairs = new List<string>();
                foreach (var kvp in settings)
                {
                    pairs.Add($"{kvp.Key}:{kvp.Value.R},{kvp.Value.G},{kvp.Value.B}");
                }
                Properties.Settings.Default.ColorSettings = string.Join("|", pairs);
                Properties.Settings.Default.Save();
            }
            catch { }
        }

        private void OnAutostartClick(object sender, EventArgs e)
        {
            try
            {
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);
                rk.SetValue("DEXP_OMNI_White_ColorChanger", Application.ExecutablePath);
                MessageBox.Show("Автозагрузка для DEXP OMNI White включена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}
