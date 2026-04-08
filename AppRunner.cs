using DriveAnalytic.Services;
using gAcss.Models;
using gAcss.Service;
using Spectre.Console;

public class AppRunner
{
    private readonly IDriveService _driveProvider;
    private readonly ExcelExportService _excelService;

    public AppRunner(IDriveService driveProvider, ExcelExportService excelService)
    {
        _driveProvider = driveProvider;
        _excelService = excelService;
    }

    public async Task RunAsync()
    {
        PrintWelcome();
        
        ShowHelp();

        while (true)
        {
            AnsiConsole.Markup("\n[white]gAcss[/] [grey]>[/] ");
            var input = Console.ReadLine()?.Trim();

            if (string.IsNullOrWhiteSpace(input)) continue;

            var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var command = parts[0].ToLower();

            if (command is "close" or "c" or "exit") break;

            try
            {
                switch (command)
                {
                    case "help": ShowHelp(); break;
                    case "clear": Console.Clear(); PrintWelcome(); ShowHelp(); break;
                    case "logout": HandleLogout(); break;
                    case "v": await HandleViewCommand(parts.Skip(1).ToArray()); break;
                    case "xlsx": await HandleExportCommand(parts.Skip(1).ToArray()); break;
                    default:
                        AnsiConsole.MarkupLine("[grey]Неизвестная команда. Введите[/] [white]'help'[/].");
                        break;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Ошибка:[/] {ex.Message}");
            }
        }
    }

    private void HandleLogout()
    {
        if (AnsiConsole.Confirm("[grey]Вы уверены, что хотите выйти из аккаунта?[/]"))
        {
            GoogleClientFactory.Logout();
            AnsiConsole.MarkupLine("[white]✅ Выход выполнен успешно. Перезапустите программу для смены аккаунта.[/]");
        }
    }

    private async Task HandleViewCommand(string[] args)
    {
        if (args.Length == 0) { ShowHelp(); return; }

        var type = args[0].ToLower();
        var data = await GetDataWithProgress(type, args.Skip(1).ToArray());

        if (data.Any()) PrintPaged(data);
        else AnsiConsole.MarkupLine("[grey]Файлы не найдены.[/]");
    }

    private async Task HandleExportCommand(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[grey]Ошибка: Недостаточно аргументов.[/]");
            return;
        }

        string rawPath = args.Last();
        string type = args[0].ToLower();
        string[] filterArgs = args.Skip(1).Take(args.Length - 2).ToArray();

        string finalPath = ValidateAndFixPath(rawPath);
        if (string.IsNullOrEmpty(finalPath)) return;

        var data = await GetDataWithProgress(type, filterArgs);

        if (data.Any())
        {
            _excelService.Export(data, finalPath);
            AnsiConsole.MarkupLine($"[white]✅ Данные ({data.Count}) сохранены в:[/] [grey]{finalPath}[/]");
        }
    }

    private async Task<List<DriveFileEntry>> GetDataWithProgress(string type, string[] args)
    {
        return await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .StartAsync($"[grey]Чтение данных ({type})...[/]", async ctx =>
            {
                return await FetchDataBasedOnType(type, args);
            });
    }

    private async Task<List<DriveFileEntry>> FetchDataBasedOnType(string type, string[] args)
    {
        return type switch
        {
            "disk" => await _driveProvider.GetAllFilesAsync(),
            "folder" => await HandleFolderFilter(args),
            "email" => (await _driveProvider.GetAllFilesAsync())
                        .Where(f => args.Any(a => string.Equals(f.Email, a, StringComparison.OrdinalIgnoreCase))).ToList(),
            _ => throw new ArgumentException($"Тип '{type}' не поддерживается.")
        };
    }

    private async Task<List<DriveFileEntry>> HandleFolderFilter(string[] args)
    {
        if (!args.Any()) throw new ArgumentException("Укажите ссылку или ID папки.");
        var result = new List<DriveFileEntry>();
        foreach (var input in args)
        {
            var id = input.Contains("/") ? input.Split('/').Last().Split('?').First() : input;
            result.AddRange(await _driveProvider.GetFolderFilesAsync(id));
        }
        return result;
    }

    private string ValidateAndFixPath(string path)
    {
        if (Directory.Exists(path))
        {
            AnsiConsole.MarkupLine($"[grey]Путь[/] [white]{path}[/] [grey]является папкой.[/]");
            var fileName = $"gAcss_Report_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            var newPath = Path.Combine(path, fileName);

            if (AnsiConsole.Confirm($"Создать файл {fileName}?")) return newPath;
            return string.Empty;
        }
        if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) path += ".xlsx";
        return path;
    }

    private void PrintPaged(List<DriveFileEntry> files, int pageSize = 10)
    {
        for (int i = 0; i < files.Count; i += pageSize)
        {
            var page = files.Skip(i).Take(pageSize).ToList();
            var table = new Table().Border(TableBorder.MinimalHeavyHead).Expand();

            table.AddColumn("Роль");
            table.AddColumn("Имя файла / Ссылка");
            table.AddColumn("Email");
            table.AddColumn("Размер (Б)");

            foreach (var file in page)
            {
                table.AddRow(
                    file.Role ?? "-",
                    $"{file.Name}\n[grey]{file.Link}[/]",
                    file.Email ?? "-",
                    file.Size?.ToString("N0") ?? "0"
                );
            }

            AnsiConsole.Write(table);

            if (i + pageSize < files.Count)
            {
                AnsiConsole.MarkupLine($"[grey]Отображено {i + page.Count} из {files.Count}. [white]ENTER[/] - далее, [white]ESC[/] - выход.[/]");
                if (Console.ReadKey(true).Key == ConsoleKey.Escape) break;
            }
        }
    }

    private void PrintWelcome()
    {
        AnsiConsole.Write(new Rule("[white]gAcss (Google drive Access) Audit Tool[/]").LeftJustified());
    }

    private void ShowHelp()
    {
        var table = new Table().Border(TableBorder.None).HideHeaders();
        table.AddColumn("Команда");
        table.AddColumn("Описание");

        table.AddRow("[white]v disk[/]", "Просмотр всех файлов");
        table.AddRow("[white]v folder <ссылка>[/]", "Просмотр файлов в конкретной папке");
        table.AddRow("[white]v email <почта>[/]", "Фильтр по адресу почты");
        table.AddRow("[white]xlsx disk <путь>[/]", "Выгрузка всего диска в Excel");
        table.AddRow("[white]xlsx folder <ссылка> <путь>[/]", "Выгрузка папки в Excel");
        table.AddRow("[white]xlsx email <почта> <путь>[/]", "Выгрузка данных по почте в Excel");
        table.AddRow("[white]logout[/]", "Выйти из текущего аккаунта Google");
        table.AddRow("[white]clear / close / c[/]", "Очистка / Закрыть программу");

        AnsiConsole.Write(new Panel(table).Header("Доступные команды").BorderColor(Color.Grey));
    }
}