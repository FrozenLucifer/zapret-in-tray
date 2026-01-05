using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace ZapretTray;

class TrayBatLauncher : ApplicationContext
{
    private readonly string _baseDir;
    private readonly string _zapretDir;
    private readonly string _serviceBatPath;
    private readonly string _logFilePath;
    private readonly string _repoUrl = "https://github.com/Flowseal/zapret-discord-youtube";
    private readonly System.Windows.Forms.Timer _updateTimer;
    private const string ServiceRegValue = "zapret-discord-youtube";
    private ToolStripMenuItem? _installServiceMenu;
    private const string MutexName = "ZapretTray_SingleInstance_Mutex";
    private const string ExitEventName = "ZapretTray_Exit_Old_Instance";

    private static Mutex? _mutex;
    private static EventWaitHandle? _exitEvent;


    private TrayBatLauncher()
    {
        Task.Run(() =>
        {
            while (_exitEvent!.WaitOne())
            {
                Exit();
                break;
            }
        });

        _baseDir = AppDomain.CurrentDomain.BaseDirectory;
        _zapretDir = Path.Combine(_baseDir, "zapret-discord");
        _serviceBatPath = Path.Combine(_zapretDir, "service.bat");
        _logFilePath = Path.Combine(_baseDir, "tray_errors.log");

        if (!File.Exists(_logFilePath))
            File.Create(_logFilePath);

        SetAutoStart();
        EnsureZapretExistsAsync().GetAwaiter().GetResult();

        new NotifyIcon
        {
            Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("ZapretTray.Resources.tray.ico") ??
                            throw new InvalidOperationException()),
            ContextMenuStrip = BuildMenu(),
            Visible = true,
            Text = "Zapret Tray"
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

            var tempPath = Path.Combine(Path.GetTempPath(), "zapret_init");
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);

            Directory.CreateDirectory(tempPath);

            var zipUrl = $"{_repoUrl}/archive/refs/heads/main.zip";
            var zipPath = Path.Combine(tempPath, "zapret.zip");

            using (var http = new HttpClient())
            {
                var resp = await http.GetAsync(zipUrl);
                resp.EnsureSuccessStatusCode();

                await using var fs = new FileStream(zipPath, FileMode.Create);
                await resp.Content.CopyToAsync(fs);
            }

            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, tempPath);

            var extractedDir = Path.Combine(tempPath, "zapret-discord-youtube-main");
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
            Console.WriteLine(context, ex);
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

    private void OpenFolder()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = _baseDir,
            UseShellExecute = true
        });
    }

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();

        menu.Items.Add(BuildServiceMenu());

        var scripts = new ToolStripMenuItem(".bat скрипты");
        LoadGeneralBats(scripts);
        menu.Items.Add(scripts);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildMiscMenu());

        menu.Items.Add("Закрыть", null, (_, _) => Exit());

        return menu;
    }

    private ToolStripMenuItem BuildServiceMenu()
    {
        var serviceItem = new ToolStripMenuItem("Сервис");

        serviceItem.DropDownItems.Add(BuildInstallServiceMenu());
        serviceItem.DropDownItems.Add("Удалить сервис", null, (_, _) => RemoveService());
        serviceItem.DropDownItems.Add("Запустить service.bat", null, (_, _) => RunServiceBat());

        return serviceItem;
    }

    private ToolStripMenuItem BuildInstallServiceMenu()
    {
        _installServiceMenu = new ToolStripMenuItem("Установить сервис");

        var currentBat = GetInstalledServiceBat();

        foreach (var bat in Directory.GetFiles(_zapretDir, "*.bat")
                     .Where(b => !Path.GetFileName(b).StartsWith("service")))
        {
            var name = Path.GetFileName(bat);
            var item = new ToolStripMenuItem(name)
            {
                Checked = Path.GetFileNameWithoutExtension(name) == currentBat
            };

            item.Click += (_, _) => InstallServiceFromBat(name);
            _installServiceMenu.DropDownItems.Add(item);
        }

        return _installServiceMenu;
    }

    private ToolStripMenuItem BuildMiscMenu()
    {
        var miscMenu = new ToolStripMenuItem("Прочее");

        var autostartItem = new ToolStripMenuItem("Автозапуск")
        {
            Checked = IsAutoStartEnabled()
        };

        autostartItem.Click += (s, _) =>
        {
            if (s is not ToolStripMenuItem item) return;

            item.Checked = !item.Checked;
            SetAutoStart(item.Checked);
        };

        miscMenu.DropDownItems.Add(autostartItem);
        miscMenu.DropDownItems.Add("Проверить обновления", null, async (_, _) => await CheckForUpdatesAsync(silent: false));
        miscMenu.DropDownItems.Add("Сбросить кеш Discord", null, (_, _) => ClearDiscordCache());
        miscMenu.DropDownItems.Add("Открыть логи", null, (_, _) => OpenLogs());
        miscMenu.DropDownItems.Add("Открыть папку", null, (_, _) => OpenFolder());

        return miscMenu;
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

    private string? GetInstalledServiceBat()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                @"SYSTEM\CurrentControlSet\Services\zapret");

            return key?.GetValue(ServiceRegValue) as string;
        }
        catch
        {
            return null;
        }
    }

    private void RemoveService()
    {
        try
        {
            RunAdmin("sc stop zapret");
            RunAdmin("sc delete zapret");

            RunAdmin("sc stop WinDivert");
            RunAdmin("sc delete WinDivert");

            RunAdmin("sc stop WinDivert14");
            RunAdmin("sc delete WinDivert14");

            RunAdmin("taskkill /IM winws.exe /F");

            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(
                @"SYSTEM\CurrentControlSet\Services\zapret",
                false
            );

            ShowSilent("Сервис zapret удалён", "Service");

            UpdateServiceMenuChecks();
        }
        catch (Exception ex)
        {
            LogError("RemoveService", ex);
            ShowSilent(ex.Message, "Ошибка");
        }
    }

    private void InstallServiceFromBat(string batName)
    {
        try
        {
            var batPath = Path.Combine(_zapretDir, batName);
            if (!File.Exists(batPath))
                throw new FileNotFoundException(batName);

            var args = ParseWinwsArgs(batPath);
            var bin = Path.Combine(_zapretDir, "bin", "winws.exe\\");

            RunAdmin("net stop zapret");
            RunAdmin("sc delete zapret");
            var quotedBinPath = $"\"\\\"{bin}\" {args}\"";

            RunAdmin(
                $"sc create zapret binPath= {quotedBinPath} start= auto DisplayName= \"zapret\""
            );


            RunAdmin("sc description zapret \"Zapret DPI bypass software\"");
            RunAdmin("sc start zapret");

            using var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(
                @"SYSTEM\CurrentControlSet\Services\zapret");

            key.SetValue(ServiceRegValue, Path.GetFileNameWithoutExtension(batName));

            ShowSilent($"Сервис установлен из {batName}", "Service");

            UpdateServiceMenuChecks();
        }
        catch (Exception ex)
        {
            LogError("InstallServiceFromBat", ex);
            ShowSilent(ex.Message, "Ошибка");
        }
    }

    private string ParseWinwsArgs(string batPath)
    {
        var binPath = Path.Combine(_zapretDir, "bin") + Path.DirectorySeparatorChar;
        var listsPath = Path.Combine(_zapretDir, "lists") + Path.DirectorySeparatorChar;
        var gameFilter = "12";

        var sb = new StringBuilder();
        bool found = false;

        foreach (var raw in File.ReadLines(batPath))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith("::"))
                continue;

            if (!found)
            {
                var idx = line.IndexOf("winws.exe\"", StringComparison.OrdinalIgnoreCase);
                if (idx < 0)
                    continue;

                found = true;
                line = line[(idx + "winws.exe\"".Length)..];
            }

            if (line.EndsWith("^"))
                line = line[..^1];

            sb.Append(' ');
            sb.Append(line);
        }

        if (!found)
            throw new Exception("winws.exe не найден в bat");

        var result = sb.ToString()
            .Replace("%BIN%", binPath)
            .Replace("%LISTS%", listsPath)
            .Replace("%GameFilter%", gameFilter);

        result = System.Text.RegularExpressions.Regex.Replace(result, @"--(\S+?)=", "--$1 ");

        result = result.Replace("\"", "\\\"");

        result = result.Replace("^", "");

        return result.Trim();
    }


    private void RunAdmin(string cmd)
    {
        var ps = $@"
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = 'cmd.exe'
$psi.Arguments = '/c {cmd}'
$psi.Verb = 'runas'
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true

$p = New-Object System.Diagnostics.Process
$p.StartInfo = $psi
$p.Start() | Out-Null
$p.WaitForExit()

Write-Output 'EXIT=' + $p.ExitCode
Write-Output $p.StandardOutput.ReadToEnd()
Write-Output $p.StandardError.ReadToEnd()
";

        var tmp = Path.GetTempFileName() + ".ps1";
        File.WriteAllText(tmp, ps);

        var p = Process.Start(new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-ExecutionPolicy Bypass -File \"{tmp}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        });

        var output = p!.StandardOutput.ReadToEnd();
        var error = p.StandardError.ReadToEnd();
        p.WaitForExit();

        File.Delete(tmp);
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

    private void UpdateServiceMenuChecks()
    {
        if (_installServiceMenu == null)
            return;

        var currentBat = GetInstalledServiceBat();

        foreach (ToolStripMenuItem item in _installServiceMenu.DropDownItems)
        {
            item.Checked = Path.GetFileNameWithoutExtension(item.Text) == currentBat;
        }
    }


    private async Task DownloadAndUpdateAsync()
    {
        try
        {
            var tempPath = Path.Combine(Path.GetTempPath(), "zapret_update");
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);

            Directory.CreateDirectory(tempPath);

            var zipUrl = $"{_repoUrl}/archive/refs/heads/main.zip";
            var zipPath = Path.Combine(tempPath, "update.zip");

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

            var extractedDir = Path.Combine(tempPath, "zapret-discord-youtube-main");
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
            var discordProcesses = Process.GetProcessesByName("Discord");

            if (discordProcesses.Length > 0)
            {
                foreach (var discordProcess in discordProcesses)
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

            var success = true;
            var failedPaths = new List<string>();

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
        Application.Exit();
    }

    [STAThread]
    static void Main()
    {
        bool created;

        _mutex = new Mutex(true, MutexName, out created);
        _exitEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ExitEventName);

        if (!created)
        {
            _exitEvent.Set();
            Thread.Sleep(500);
            _mutex = new Mutex(true, MutexName, out _);
        }

        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayBatLauncher());
        }
        finally
        {
            _mutex.ReleaseMutex();
        }
    }
}