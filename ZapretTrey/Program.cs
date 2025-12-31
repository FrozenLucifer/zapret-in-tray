using System.Diagnostics;
using System.Reflection;

namespace ZapretTrey;

class TrayBatLauncher : ApplicationContext
{
    private readonly NotifyIcon _trayIcon;
    private readonly string _baseDir;
    private readonly string _zapretDir;
    private readonly string _serviceBatPath;
    private readonly string _logFilePath;
    private readonly string _repoUrl = "https://github.com/Flowseal/zapret-discord-youtube";
    private readonly System.Windows.Forms.Timer _updateTimer;

    private TrayBatLauncher()
    {
        _baseDir = AppDomain.CurrentDomain.BaseDirectory;
        _zapretDir = Path.Combine(_baseDir, "zapret-discord");
        _serviceBatPath = Path.Combine(_zapretDir, "service.bat");
        _logFilePath = Path.Combine(_baseDir, "tray_errors.log");
        
        if (!File.Exists(_logFilePath))
            File.WriteAllText(_logFilePath, "");

        SetAutoStart();
        EnsureZapretExistsAsync().GetAwaiter().GetResult();
        
        _trayIcon = new NotifyIcon
        {
            Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("ZapretTrey.Resources.tray.ico") ?? throw new InvalidOperationException()),
            ContextMenuStrip = BuildMenu(),
            Visible = true,
            Text = "General BAT launcher"
        };

        _updateTimer = new System.Windows.Forms.Timer();
        _updateTimer.Interval = 24 * 60 * 60 * 1000;
        _updateTimer.Tick += async (_, _) => await CheckForUpdatesAsync(silent: true);
        _updateTimer.Start();

        var startupTimer = new System.Windows.Forms.Timer();
        startupTimer.Interval = 30000;
        startupTimer.Tick += async (_, _) =>
        {
            await CheckForUpdatesAsync(silent: true);
            startupTimer.Stop();
            startupTimer.Dispose();
        };
        startupTimer.Start();
    }
    
    private async Task EnsureZapretExistsAsync()
    {
        try
        {
            if (File.Exists(_serviceBatPath))
                return;

            Directory.CreateDirectory(_zapretDir);

            ShowSilent("Zapret не найден. Идёт первичная загрузка…", "Инициализация");

            string tempPath = Path.Combine(Path.GetTempPath(), "zapret_init");
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);

            Directory.CreateDirectory(tempPath);

            string zipUrl = $"{_repoUrl}/archive/refs/heads/main.zip";
            string zipPath = Path.Combine(tempPath, "zapret.zip");

            using (var http = new HttpClient())
            {
                var resp = await http.GetAsync(zipUrl);
                resp.EnsureSuccessStatusCode();

                using var fs = new FileStream(zipPath, FileMode.Create);
                await resp.Content.CopyToAsync(fs);
            }

            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, tempPath);

            string extractedDir = Path.Combine(tempPath, "zapret-discord-youtube-main");
            if (!Directory.Exists(extractedDir))
                throw new Exception("Не удалось распаковать zapret");

            foreach (var file in Directory.GetFiles(extractedDir, "*", SearchOption.AllDirectories))
            {
                var rel = file[(extractedDir.Length + 1)..];
                var dst = Path.Combine(_zapretDir, rel);

                var dstDir = Path.GetDirectoryName(dst);
                if (!Directory.Exists(dstDir))
                    Directory.CreateDirectory(dstDir);

                File.Copy(file, dst, true);
            }

            ShowSilent("Zapret успешно загружен", "Готово");
        }
        catch (Exception ex)
        {
            LogError("EnsureZapretExistsAsync", ex);
            ShowSilent($"Ошибка загрузки zapret: {ex.Message}", "Ошибка");
        }
    }


    private void LogError(string context, Exception ex)
    {
        try
        {
            File.AppendAllText(
                _logFilePath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}\n{ex}\n\n"
            );
        }
        catch
        {
        }
    }

    private void OpenLogs()
    {
        if (!File.Exists(_logFilePath))
        {
            ShowSilent("Файл логов пока не создан", "Логи");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = _logFilePath,
            UseShellExecute = true
        });
    }

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();

        var generalMenu = new ToolStripMenuItem("General scripts");
        LoadGeneralBats(generalMenu);
        menu.Items.Add(generalMenu);

        menu.Items.Add(new ToolStripSeparator());

        var serviceItem = new ToolStripMenuItem("Service");
        serviceItem.DropDownItems.Add("Запуск Service.bat", null, (_, _) => RunServiceBat());
        serviceItem.DropDownItems.Add("Проверить обновления", null, async (_, _) => await CheckForUpdatesAsync(silent: false));
        serviceItem.DropDownItems.Add("Обновить списки", null, async (_, _) => await UpdateListsAsync());
        menu.Items.Add(serviceItem);

        menu.Items.Add("Сбросить кеш Discord", null, (_, _) => ClearDiscordCache());

        menu.Items.Add(new ToolStripSeparator());

        var autostartItem = new ToolStripMenuItem("Автозапуск");
        autostartItem.Checked = IsAutoStartEnabled();
        autostartItem.Click += (s, _) =>
        {
            var item = s as ToolStripMenuItem;

            if (item != null)
            {
                item.Checked = !item.Checked;
                SetAutoStart(item.Checked);
            }
        };
        menu.Items.Add(autostartItem);
        menu.Items.Add("Открыть логи", null, (_, _) => OpenLogs());
        menu.Items.Add("Выход", null, (_, _) => Exit());

        return menu;
    }

    private void LoadGeneralBats(ToolStripMenuItem root)
    {
        root.DropDownItems.Clear();

        var files = Directory
            .GetFiles(_zapretDir, "general*.bat", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f);

        foreach (var file in files)
        {
            var name = Path.GetFileName(file);
            root.DropDownItems.Add(name, null, (_, _) => RunBat(name));
        }

        if (!files.Any())
            root.DropDownItems.Add("(не найдено)").Enabled = false;
    }

    private void RunBat(string name)
    {
        var path = Path.Combine(_zapretDir, name);
        if (!File.Exists(path)) return;

        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            WorkingDirectory = _baseDir,
            UseShellExecute = true
        });
    }

    private void RunServiceBat()
    {
        if (!File.Exists(_serviceBatPath))
        {
            ShowSilent("Файл service.bat не найден", "Ошибка");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = _serviceBatPath,
            WorkingDirectory = _zapretDir,
            UseShellExecute = true
        });
    }

    private async Task CheckForUpdatesAsync(bool silent = true)
    {
        try
        {
            if (!File.Exists(_serviceBatPath))
                return;

            var process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = $"/c \"\"{_serviceBatPath}\" check_updates soft\"";
            process.StartInfo.WorkingDirectory = _zapretDir;
            process.StartInfo.CreateNoWindow = silent;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;

            process.Start();
            var output = await process.StandardOutput.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (!silent && output.Contains("New version available"))
            {
                var result = AskSilent("Доступна новая версия. Скачать и обновить?", "Обновление");

                if (result == DialogResult.Yes)
                {
                    await DownloadAndUpdateAsync();
                }
            }
        }
        catch (Exception ex)
        {
            LogError("CheckForUpdatesAsync", ex);

            if (!silent)
            {
                ShowSilent($"Ошибка при проверке обновлений: {ex.Message}", "Ошибка");
            }
        }
    }

    private async Task UpdateListsAsync()
    {
        try
        {
            if (!File.Exists(_serviceBatPath))
                return;

            var process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = $"/c \"\"{_serviceBatPath}\"\"";
            process.StartInfo.WorkingDirectory = _zapretDir;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;

            process.Start();

            await using (var sw = process.StandardInput)
            {
                if (sw.BaseStream.CanWrite)
                {
                    await sw.WriteLineAsync("8");
                    await sw.WriteLineAsync();
                }
            }

            await process.WaitForExitAsync();

            ShowSilent("Списки успешно обновлены", "Обновление");
        }
        catch (Exception ex)
        {
            LogError("UpdateListsAsync", ex);

            ShowSilent($"Ошибка при обновлении списков: {ex.Message}", "Ошибка");
        }
    }

    private async Task DownloadAndUpdateAsync()
    {
        try
        {
            string tempPath = Path.Combine(Path.GetTempPath(), "zapret_update");
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);

            Directory.CreateDirectory(tempPath);

            string zipUrl = $"{_repoUrl}/archive/refs/heads/main.zip";
            string zipPath = Path.Combine(tempPath, "update.zip");

            using (var httpClient = new HttpClient())
            {
                httpClient.Timeout = TimeSpan.FromMinutes(5);

                var response = await httpClient.GetAsync(zipUrl);
                if (!response.IsSuccessStatusCode)
                    throw new Exception($"Ошибка загрузки: {response.StatusCode}");

                await using (var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }

            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, tempPath, true);

            string extractedDir = Path.Combine(tempPath, "zapret-discord-youtube-main");
            if (!Directory.Exists(extractedDir))
                throw new Exception("Не удалось найти распакованные файлы");

            var exeName = Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName);

            foreach (var file in Directory.GetFiles(extractedDir, "*", SearchOption.AllDirectories))
            {
                var relativePath = file[(extractedDir.Length + 1)..];
                var destPath = Path.Combine(_zapretDir, relativePath);

                if (Path.GetFileName(file).Equals(exeName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var destDir = Path.GetDirectoryName(destPath);
                if (!Directory.Exists(destDir))
                    Directory.CreateDirectory(destDir);

                File.Copy(file, destPath, true);
            }

            ShowSilent("Обновление успешно завершено. Перезапустите приложение для применения изменений.", "Обновление");
        }
        catch (Exception ex)
        {
            LogError("DownloadAndUpdateAsync", ex);

            ShowSilent($"Ошибка при обновлении: {ex.Message}", "Ошибка");
        }
    }

    private bool IsAutoStartEnabled()
    {
        try
        {
            var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
                false);

            if (key == null) return false;

            var value = key.GetValue("ZapretTrayLauncher") as string;
            key.Close();

            return !string.IsNullOrEmpty(value) &&
                   value.Equals($"\"{Application.ExecutablePath}\"", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private void SetAutoStart(bool enabled = true)
    {
        try
        {
            var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",
                true);

            if (key == null) return;

            if (enabled)
            {
                key.SetValue("ZapretTrayLauncher", $"\"{Application.ExecutablePath}\"");
            }
            else
            {
                key.DeleteValue("ZapretTrayLauncher", false);
            }

            key.Close();
        }
        catch (Exception ex)
        {
            LogError("SetAutoStart", ex);

            ShowSilent($"Ошибка при настройке автозапуска: {ex.Message}", "Ошибка");
        }
    }

    private void ClearDiscordCache()
    {
        try
        {
            Process[] discordProcesses = Process.GetProcessesByName("Discord");

            if (discordProcesses.Length > 0)
            {
                foreach (Process discordProcess in discordProcesses)
                {
                    try
                    {
                        discordProcess.Kill();
                        discordProcess.WaitForExit(5000);
                    }
                    catch (Exception ex)
                    {
                        LogError("ClearDiscordCache", ex);
                    }
                }

                Thread.Sleep(2000);
            }

            string[] paths =
            [
                Environment.ExpandEnvironmentVariables("%AppData%\\Discord\\Cache"),
                Environment.ExpandEnvironmentVariables("%AppData%\\Discord\\Code Cache"),
                Environment.ExpandEnvironmentVariables("%AppData%\\Discord\\GPUCache"),
            ];

            bool success = true;
            List<string> failedPaths = new List<string>();

            foreach (var p in paths)
            {
                if (!Directory.Exists(p)) continue;

                try
                {
                    Directory.Delete(p, true);
                }
                catch (Exception)
                {
                    success = false;
                    failedPaths.Add(p);
                }
            }

            if (success)
            {
                ShowSilent("Кеш Discord успешно очищен", "OK");
            }
            else
            {
                MessageBox.Show($"Частично очищено. Не удалось очистить:\n{string.Join("\n", failedPaths)}",
                    "Внимание",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            LogError("ClearDiscordCache", ex);

            ShowSilent($"Ошибка при очистке кеша: {ex.Message}", "Ошибка");
        }
    }
    
    private void ShowSilent(string text, string caption)
    {
        MessageBox.Show(
            text,
            caption,
            MessageBoxButtons.OK,
            MessageBoxIcon.None
        );
    }

    private DialogResult AskSilent(string text, string caption)
    {
        return MessageBox.Show(
            text,
            caption,
            MessageBoxButtons.YesNo,
            MessageBoxIcon.None
        );
    }


    private void Exit()
    {
        _updateTimer.Stop();
        _updateTimer.Dispose();
        _trayIcon.Visible = false;
        Application.Exit();
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new TrayBatLauncher());
    }
}