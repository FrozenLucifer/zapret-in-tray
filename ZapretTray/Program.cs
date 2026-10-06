using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ZapretTray;

class TrayBatLauncher : ApplicationContext
{
    private const string RepoUrl = "https://github.com/Flowseal/zapret-discord-youtube";
    private const string VersionUrl = "https://raw.githubusercontent.com/Flowseal/zapret-discord-youtube/main/.service/version.txt";
    private const string ServiceRegValue = "zapret-discord-youtube";
    private const string MutexName = "ZapretTray_SingleInstance_Mutex";
    private const string ExitEventName = "ZapretTray_Exit_Old_Instance";

    private readonly string _baseDir = AppDomain.CurrentDomain.BaseDirectory;
    private readonly string _appDataDir;
    private readonly string _zapretDir;
    private readonly string _serviceBatPath;
    private readonly string _logFilePath;
    
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(30) };
    private readonly NotifyIcon _notifyIcon;
    private ToolStripMenuItem? _installServiceMenu;

    private static Mutex? _mutex;
    private static EventWaitHandle? _exitEvent;

    private TrayBatLauncher()
    {
        _appDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ZapretTray");

        _zapretDir = Path.Combine(_appDataDir, "zapret-discord");
        _serviceBatPath = Path.Combine(_zapretDir, "service.bat");
        _logFilePath = Path.Combine(_appDataDir, "tray_errors.log");

        Directory.CreateDirectory(_appDataDir);

        EnsureLogFileExists();
        EnsureZapretExistsAsync().GetAwaiter().GetResult();

        _notifyIcon = new NotifyIcon
        {
            Icon = new Icon(
                Assembly.GetExecutingAssembly().GetManifestResourceStream("ZapretTray.Resources.tray.ico")
                ?? throw new InvalidOperationException("Не удалось загрузить значок приложения.")),
            ContextMenuStrip = BuildMenu(),
            Text = "Zapret Tray",
            Visible = true
        };
    }

    private void EnsureLogFileExists()
    {
        try
        {
            if (!File.Exists(_logFilePath))
            {
                using var _ = File.Create(_logFilePath);
            }
        }
        catch
        {
        }
    }

private async Task EnsureZapretExistsAsync()
{
    if (File.Exists(_serviceBatPath))
        return;

    var tempPath = Path.Combine(
        Path.GetTempPath(),
        "ZapretTray",
        Guid.NewGuid().ToString("N"));

    try
    {
        Directory.CreateDirectory(_zapretDir);
        Directory.CreateDirectory(tempPath);

        var zipPath = Path.Combine(tempPath, "zapret.zip");

        using var response = await _http.GetAsync(
            $"{RepoUrl}/archive/refs/heads/main.zip",
            HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using (var file = new FileStream(
            zipPath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None))
        {
            await response.Content.CopyToAsync(file);
        }

        System.IO.Compression.ZipFile.ExtractToDirectory(
            zipPath,
            tempPath);

        var sourceDir = Path.Combine(
            tempPath,
            "zapret-discord-youtube-main");

        if (!Directory.Exists(sourceDir))
            throw new InvalidOperationException(
                $"После распаковки не найдена папка:\n{sourceDir}");

        CopyDirectory(sourceDir, _zapretDir);

        if (!File.Exists(_serviceBatPath))
            throw new InvalidOperationException(
                $"После копирования не найден service.bat:\n{_serviceBatPath}");
    }
    catch (Exception ex)
    {
        LogError("EnsureZapretExistsAsync ERROR", ex);

        ShowError(
            $"Не удалось загрузить zapret.\n\n" +
            $"{ex.GetType().Name}\n" +
            $"{ex.Message}\n\n" +
            $"Подробности записаны в:\n{_logFilePath}");
    }
    finally
    {
        TryDeleteDirectory(tempPath);
    }
}

    private static void CopyDirectory(string sourceDir, string destinationDir)
    {
        foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
        {
            var destination = Path.Combine(destinationDir, Path.GetRelativePath(sourceDir, file));
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(file, destination, true);
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path)) Directory.Delete(path, true);
        }
        catch
        {
        }
    }

    private void LogError(string context, Exception ex)
    {
        try
        {
            File.AppendAllText(_logFilePath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}{Environment.NewLine}{ex}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
        }
    }

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add(BuildServiceMenu());
        var strategies = new ToolStripMenuItem(".bat-стратегии");
        LoadGeneralBats(strategies);
        menu.Items.Add(strategies);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(BuildMiscMenu());
        menu.Items.Add("Закрыть", null, (_, _) => Exit());
        return menu;
    }

    private ToolStripMenuItem BuildServiceMenu()
    {
        var menu = new ToolStripMenuItem("Сервис");
        menu.DropDownItems.Add(BuildInstallServiceMenu());
        menu.DropDownItems.Add("Удалить сервис zapret", null, (_, _) => RemoveService());
        menu.DropDownItems.Add("Открыть service.bat", null, (_, _) => RunServiceBat());
        return menu;
    }

    private ToolStripMenuItem BuildInstallServiceMenu()
    {
        _installServiceMenu = new ToolStripMenuItem("Установить сервис");
        var bats = GetRootBatFiles();
        var installed = GetInstalledServiceBat();

        foreach (var bat in bats)
        {
            var name = Path.GetFileName(bat);
            var item = new ToolStripMenuItem(name)
            {
                Checked = string.Equals(Path.GetFileNameWithoutExtension(name), installed, StringComparison.OrdinalIgnoreCase)
            };
            item.Click += (_, _) => InstallServiceFromBat(name);
            _installServiceMenu.DropDownItems.Add(item);
        }

        if (bats.Count == 0)
        {
            _installServiceMenu.DropDownItems.Add("(стратегии не найдены)").Enabled = false;
            _installServiceMenu.Enabled = false;
        }

        return _installServiceMenu;
    }

    private List<string> GetRootBatFiles() => !Directory.Exists(_zapretDir)
        ? []
        : Directory.GetFiles(_zapretDir, "*.bat")
            .Where(path => !Path.GetFileName(path).StartsWith("service", StringComparison.OrdinalIgnoreCase))
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private ToolStripMenuItem BuildMiscMenu()
    {
        var menu = new ToolStripMenuItem("Прочее");
        var autostart = new ToolStripMenuItem("Автозапуск") { Checked = IsAutoStartEnabled() };
        autostart.Click += (sender, _) =>
        {
            if (sender is not ToolStripMenuItem item) return;

            var enabled = !item.Checked;
            if (SetAutoStart(enabled)) item.Checked = enabled;
        };
        menu.DropDownItems.Add(autostart);
        menu.DropDownItems.Add("Проверить обновления Zapret", null, async (_, _) => await CheckForUpdatesAsync());
        menu.DropDownItems.Add("Сбросить кеш Discord", null, async (_, _) => await ClearDiscordCacheAsync());
        menu.DropDownItems.Add("Открыть логи", null, (_, _) => OpenWithShell(_logFilePath));
        menu.DropDownItems.Add("Открыть папку", null, (_, _) => OpenWithShell(_appDataDir));
        return menu;
    }

    private void LoadGeneralBats(ToolStripMenuItem root)
    {
        foreach (var file in GetRootBatFiles().Where(path => Path.GetFileName(path).StartsWith("general", StringComparison.OrdinalIgnoreCase)))
        {
            var name = Path.GetFileName(file);
            root.DropDownItems.Add(name, null, (_, _) => RunBat(name));
        }

        if (root.DropDownItems.Count == 0) root.DropDownItems.Add("(не найдено)").Enabled = false;
    }

    private void RunBat(string name)
    {
        var path = Path.Combine(_zapretDir, name);

        if (!File.Exists(path))
        {
            ShowError($"Файл стратегии не найден: {name}");
            return;
        }

        Process.Start(new ProcessStartInfo { FileName = path, WorkingDirectory = _zapretDir, UseShellExecute = true });
    }

    private string? GetInstalledServiceBat()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\zapret");
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
            // Не трогаем WinDivert и чужие winws.exe: они могут принадлежать другой программе.
            RunAdminCommand(
                "sc stop zapret >nul 2>&1 & sc delete zapret >nul 2>&1 & reg delete \"HKLM\\SYSTEM\\CurrentControlSet\\Services\\zapret\" /v zapret-discord-youtube /f >nul 2>&1 & exit /b 0");
            UpdateServiceMenuChecks();
            ShowNotification("Сервис zapret удалён.", "Zapret Tray");
        }
        catch (Exception ex)
        {
            LogError("RemoveService", ex);
            ShowError($"Не удалось удалить сервис zapret.\n\n{ex.Message}");
        }
    }

    private void InstallServiceFromBat(string batName)
    {
        try
        {
            var batPath = Path.Combine(_zapretDir, batName);
            var winwsPath = Path.Combine(_zapretDir, "bin", "winws.exe");

            if (!File.Exists(batPath))
                throw new FileNotFoundException("Файл стратегии не найден.", batPath);

            if (!File.Exists(winwsPath))
                throw new FileNotFoundException("Не найден winws.exe.", winwsPath);
            
            var args = ParseWinwsArgs(batPath);
            var binPath = $"\"\\\"{winwsPath}\\\" {args}\"";

            var createCommand =
                $"sc create zapret " +
                $"binPath= {binPath} " +
                $"start= auto " +
                $"DisplayName= \"zapret\"";

            RunAdminCommand(
                $"sc stop zapret >nul 2>&1 & " +
                $"sc delete zapret >nul 2>&1 & " +
                $"{createCommand} && " +
                $"sc description zapret \"Zapret DPI bypass software\" && " +
                $"sc start zapret"
            );

            RunAdminCommand(
                $"reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Services\\zapret\" " +
                $"/v {ServiceRegValue} " +
                $"/t REG_SZ " +
                $"/d \"{Path.GetFileNameWithoutExtension(batName)}\" " +
                "/f"
            );

            UpdateServiceMenuChecks();
            ShowNotification($"Сервис установлен из {batName}.", "Zapret Tray");
        }
        catch (Exception ex)
        {
            LogError("InstallServiceFromBat", ex);
            ShowError($"Не удалось установить сервис.\n\n{ex.Message}");
        }
    }

    private string ParseWinwsArgs(string batPath)
    {
        var binPath = Path.Combine(_zapretDir, "bin") + Path.DirectorySeparatorChar;
        var listsPath = Path.Combine(_zapretDir, "lists") + Path.DirectorySeparatorChar;

        var (tcp, udp) = GetGameFilterPorts();

        var sb = new StringBuilder();
        var found = false;

        foreach (var raw in File.ReadLines(batPath))
        {
            var line = raw.Trim();

            if (line.Length == 0 ||
                line.StartsWith("::") ||
                line.StartsWith("REM ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!found)
            {
                var idx = line.IndexOf("winws.exe\"", StringComparison.OrdinalIgnoreCase);

                if (idx < 0)
                    continue;

                found = true;
                line = line[(idx + "winws.exe\"".Length)..];
            }

            line = line.TrimEnd();

            if (line.EndsWith('^'))
                line = line[..^1].TrimEnd();

            if (line.Length > 0)
            {
                sb.Append(' ');
                sb.Append(line);
            }
        }

        if (!found)
            throw new InvalidOperationException(
                "В выбранной стратегии не найдена команда запуска winws.exe.");

        var result = sb.ToString().Trim();

        // Подстановка переменных из bat.
        result = result
            .Replace("%BIN%", binPath, StringComparison.OrdinalIgnoreCase)
            .Replace("%LISTS%", listsPath, StringComparison.OrdinalIgnoreCase)
            .Replace("%GameFilterTCP%", tcp, StringComparison.OrdinalIgnoreCase)
            .Replace("%GameFilterUDP%", udp, StringComparison.OrdinalIgnoreCase)
            .Replace("%GameFilter%", tcp, StringComparison.OrdinalIgnoreCase);

        // КРИТИЧЕСКИ ВАЖНО:
        // старый код преобразовывал:
        //
        // --wf-tcp=80
        //
        // в:
        //
        // --wf-tcp 80
        //
        // Для запуска через Service Control Manager это необходимо.
        result = Regex.Replace(
            result,
            @"--(\S+?)=",
            "--$1 "
        );

        // Убираем batch-символ продолжения строки.
        result = result.Replace("^", string.Empty);

        // Экранируем кавычки перед передачей команды в cmd/sc.
        result = result.Replace("\"", "\\\"");

        return result.Trim();
    }

    private (string Tcp, string Udp) GetGameFilterPorts()
    {
        const string disabled = "12";
        var path = Path.Combine(_zapretDir, "utils", "game_filter.enabled");
        if (!File.Exists(path)) return (disabled, disabled);

        var values = File.ReadLines(path)
            .Select(line => line.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);
        values.TryGetValue("mode", out var mode);
        values.TryGetValue("tcp", out var tcp);
        values.TryGetValue("udp", out var udp);
        tcp = IsValidPortRange(tcp) ? tcp! : "1024-65535";
        udp = IsValidPortRange(udp) ? udp! : "1024-65535";
        return mode?.ToLowerInvariant() switch
        {
            "all" => (tcp, udp), "tcp" => (tcp, disabled), "udp" => (disabled, udp), _ => (disabled, disabled)
        };
    }

    private static bool IsValidPortRange(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;

        foreach (var item in value.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            var match = Regex.Match(item, "^(\\d+)(?:-(\\d+))?$");
            if (!match.Success || !int.TryParse(match.Groups[1].Value, out var start) || start is < 1 or > 65535) return false;

            var end = match.Groups[2].Success && int.TryParse(match.Groups[2].Value, out var parsed) ? parsed : start;
            if (end is < 1 or > 65535 || start > end) return false;
        }

        return true;
    }

    private void RunAdminCommand(string cmd)
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
            ShowError("Файл service.bat не найден.");
            return;
        }

        Process.Start(new ProcessStartInfo { FileName = _serviceBatPath, WorkingDirectory = _zapretDir, UseShellExecute = true });
    }

    private async Task CheckForUpdatesAsync()
    {
        try
        {
            var local = GetLocalVersion() ?? throw new InvalidOperationException("Не удалось определить установленную версию Zapret.");
            var remote = (await _http.GetStringAsync(VersionUrl)).Trim();
            if (string.IsNullOrWhiteSpace(remote)) throw new InvalidOperationException("GitHub вернул пустой номер версии.");

            if (string.Equals(local, remote, StringComparison.OrdinalIgnoreCase))
            {
                ShowNotification($"Установлена актуальная версия Zapret ({local}).", "Zapret Tray");
            }
            else if (Ask($"Доступна версия Zapret {remote} (установлена {local}). Открыть официальную страницу релиза?", "Обновление Zapret") ==
                     DialogResult.Yes)
            {
                OpenWithShell($"{RepoUrl}/releases/tag/{Uri.EscapeDataString(remote)}");
            }
        }
        catch (Exception ex)
        {
            LogError("CheckForUpdatesAsync", ex);
            ShowError($"Не удалось проверить обновления Zapret.\n\n{ex.Message}");
        }
    }

    private string? GetLocalVersion()
    {
        if (!File.Exists(_serviceBatPath)) return null;

        foreach (var line in File.ReadLines(_serviceBatPath))
        {
            var match = Regex.Match(line, "^set \\\"LOCAL_VERSION=(.+?)\\\"$", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value.Trim();
        }

        return null;
    }

    private void UpdateServiceMenuChecks()
    {
        if (_installServiceMenu == null) return;

        var current = GetInstalledServiceBat();
        foreach (var item in _installServiceMenu.DropDownItems.OfType<ToolStripMenuItem>())
            item.Checked = string.Equals(Path.GetFileNameWithoutExtension(item.Text), current, StringComparison.OrdinalIgnoreCase);
    }

    private bool IsAutoStartEnabled()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", false);
            return string.Equals(key?.GetValue("ZapretTrayLauncher") as string,
                $"\"{Application.ExecutablePath}\"",
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private bool SetAutoStart(bool enabled)
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)
                            ?? throw new InvalidOperationException("Не удалось открыть настройки автозапуска Windows.");
            if (enabled) key.SetValue("ZapretTrayLauncher", $"\"{Application.ExecutablePath}\"");
            else key.DeleteValue("ZapretTrayLauncher", false);
            ShowNotification(enabled ? "Автозапуск включён." : "Автозапуск выключен.", "Zapret Tray");
            return true;
        }
        catch (Exception ex)
        {
            LogError("SetAutoStart", ex);
            ShowError($"Не удалось изменить автозапуск.\n\n{ex.Message}");
            return false;
        }
    }

    private async Task ClearDiscordCacheAsync()
    {
        if (Ask("Discord будет закрыт, а его кеш очищен. Несохранённые данные в Discord могут быть потеряны. Продолжить?", "Очистка кеша Discord") !=
            DialogResult.Yes) return;

        try
        {
            var installations = new[]
            {
                ("Discord", "Discord"), ("DiscordPTB", "discordptb"), ("DiscordCanary", "discordcanary"), ("DiscordDevelopment", "discorddevelopment")
            };

            foreach (var installation in installations)
            foreach (var process in Process.GetProcessesByName(installation.Item1))
            {
                using (process)
                {
                    process.Kill();
                    await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                }
            }

            await Task.Delay(TimeSpan.FromSeconds(1));
            var removed = 0;

            foreach (var installation in installations)
            {
                var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), installation.Item2);

                foreach (var cache in new[] { "Cache", "Code Cache", "GPUCache" })
                {
                    var path = Path.Combine(root, cache);
                    if (!Directory.Exists(path)) continue;

                    Directory.Delete(path, true);
                    removed++;
                }
            }

            ShowNotification(removed > 0 ? "Кеш Discord очищен." : "Папки кеша Discord не найдены.", "Zapret Tray");
        }
        catch (Exception ex)
        {
            LogError("ClearDiscordCacheAsync", ex);
            ShowError($"Не удалось полностью очистить кеш Discord.\n\n{ex.Message}");
        }
    }

    private void ShowNotification(string text, string title)
    {
        _notifyIcon.BalloonTipTitle = title;
        _notifyIcon.BalloonTipText = text;
        _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
        _notifyIcon.ShowBalloonTip(3000);
    }

    private static void OpenWithShell(string path) => Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
    private static void ShowError(string text) => MessageBox.Show(text, "Zapret Tray — ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
    private static DialogResult Ask(string text, string title) => MessageBox.Show(text, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

    private void Exit()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _http.Dispose();
        Application.Exit();
    }

    [STAThread]
    private static void Main()
    {
        _mutex = new Mutex(true, MutexName, out var created);

        if (!created)
        {
            try
            {
                using var exitEvent = EventWaitHandle.OpenExisting(ExitEventName);
                exitEvent.Set();
            }
            catch
            {
            }

            _mutex.WaitOne();
        }

        _exitEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ExitEventName);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        using var context = new TrayBatLauncher();
        Task.Run(() =>
        {
            _exitEvent.WaitOne();
            context.Exit();
        });
        Application.Run(context);
        _mutex.ReleaseMutex();
    }
}