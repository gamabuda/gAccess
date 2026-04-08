using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using ClosedXML.Excel;

class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            Console.Title = "DriveAnalytic";
            PrintWelcome();

            // Создаем сервис Google Drive с подпапкой data
            var service = await CreateDriveService();

            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.Write("driveanalytic > ");
                Console.ResetColor();

                var input = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(input)) continue;

                var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var mainCommand = parts[0].ToLower();

                try
                {
                    switch (mainCommand)
                    {
                        case "help": ShowHelp(); break;
                        case "exit": return;
                        case "clear": Console.Clear(); PrintWelcome(); break;
                        case "v": await HandleView(service, parts); break;
                        case "xlsx": await HandleExcel(service, parts); break;
                        default:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Неизвестная команда. Введите 'help'");
                            Console.ResetColor();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Ошибка: " + ex.Message);
                    Console.ResetColor();
                }
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Критическая ошибка: " + ex.Message);
            Console.ResetColor();
        }

        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
    }

    // ===================== Приветствие и помощь =====================
    static void PrintWelcome()
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("══════════════════════════════════════");
        Console.WriteLine("        Доступ к Google Drive");
        Console.WriteLine("══════════════════════════════════════");
        Console.ResetColor();
        Console.WriteLine("Мини-программа для просмотра и экспорта доступа к файлам");
        Console.WriteLine("Введите 'help' для списка команд\n");
    }

    static void ShowHelp()
    {
        Console.WriteLine("Команды:");
        Console.WriteLine(" v disk");
        Console.WriteLine(" v email <email1> <email2> ...");
        Console.WriteLine(" v folder <folder_link1> <folder_link2> ...");
        Console.WriteLine(" v ignore <email1> <email2> ...");
        Console.WriteLine(" xlsx disk <path_to_file.xlsx>");
        Console.WriteLine(" xlsx email <email1> ... <path_to_file.xlsx>");
        Console.WriteLine(" xlsx folder <folder_link1> ... <path_to_file.xlsx>");
        Console.WriteLine(" xlsx ignore <email1> ... <path_to_file.xlsx>");
        Console.WriteLine(" clear - Очистить консоль");
        Console.WriteLine(" exit - Выйти из программы\n");
    }

    // ===================== Создание сервиса Google Drive (old) =====================
    static async Task<DriveService> CreateDriveService()
    {
        string dataFolder = Path.Combine(AppContext.BaseDirectory, "data");
        Directory.CreateDirectory(dataFolder);

        string credentialsPath = Path.Combine(dataFolder, "credentials.json");
        string tokenPath = Path.Combine(dataFolder, "token");

        if (!File.Exists(credentialsPath))
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Файл credentials.json не найден!");
            Console.WriteLine("Скачайте OAuth 2.0 Client ID (Desktop app) и поместите его в папку 'data'");
            Console.ResetColor();
            throw new FileNotFoundException("credentials.json отсутствует в папке data");
        }

        var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.FromFile(credentialsPath).Secrets,
            new[] { DriveService.Scope.DriveReadonly },
            "user",
            CancellationToken.None,
            new FileDataStore(tokenPath, true));

        return new DriveService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "DriveAnalytic"
        });
    }

    // ===================== Анимация "три точки" =====================
    static async Task ShowProgress(Func<Task> action, string message)
    {
        var cts = new CancellationTokenSource();
        var token = cts.Token;

        var animationTask = Task.Run(async () =>
        {
            int dots = 0;
            while (!token.IsCancellationRequested)
            {
                Console.Write($"\r{message}{new string('.', dots)}   ");
                dots = (dots + 1) % 4;
                await Task.Delay(500);
            }
            Console.WriteLine();
        });

        try
        {
            await action();
        }
        finally
        {
            cts.Cancel();
            await animationTask;
        }
    }

    // ===================== Работа с файлами =====================
    static async Task<List<Google.Apis.Drive.v3.Data.File>> GetAllFiles(DriveService service)
    {
        var files = new List<Google.Apis.Drive.v3.Data.File>();
        string pageToken = null;

        do
        {
            var request = service.Files.List();
            request.Fields = "nextPageToken, files(id,name,mimeType,permissions(emailAddress,role,type),webViewLink)";
            request.PageSize = 1000;
            request.PageToken = pageToken;

            var result = await request.ExecuteAsync();
            files.AddRange(result.Files);
            pageToken = result.NextPageToken;

        } while (pageToken != null);

        return files;
    }

    static async Task<List<Google.Apis.Drive.v3.Data.File>> GetFolderFilesRecursive(DriveService service, string folderId)
    {
        var result = new List<Google.Apis.Drive.v3.Data.File>();

        var request = service.Files.List();
        request.Q = $"'{folderId}' in parents";
        request.Fields = "files(id,name,mimeType,permissions(emailAddress,role,type),webViewLink)";
        request.PageSize = 1000;

        var response = await request.ExecuteAsync();

        foreach (var file in response.Files)
        {
            result.Add(file);
            if (file.MimeType == "application/vnd.google-apps.folder")
            {
                var inner = await GetFolderFilesRecursive(service, file.Id);
                result.AddRange(inner);
            }
        }

        return result;
    }

    static string ExtractFolderId(string url) => url.Split('/').Last();

    // ===================== Обработка команд view (v) =====================
    static async Task HandleView(DriveService service, string[] parts)
    {
        if (parts.Length < 2)
        {
            Console.WriteLine("Использование: v disk/email/folder/ignore ...");
            return;
        }

        var subCommand = parts[1].ToLower();

        switch (subCommand)
        {
            case "disk":
                await ShowProgress(async () =>
                {
                    var filesDisk = await GetAllFiles(service);
                    PrintPaged(filesDisk);
                }, "Загрузка всех файлов");
                break;

            case "email":
                if (parts.Length < 3) { Console.WriteLine("Использование: v email <email1> ..."); return; }
                var emails = parts.Skip(2).ToArray();
                await ShowProgress(async () =>
                {
                    var filesEmail = await GetAllFiles(service);
                    var filteredEmail = filesEmail
                        .Where(f => f.Permissions != null && f.Permissions.Any(p => emails.Contains(p.EmailAddress)))
                        .ToList();
                    PrintPaged(filteredEmail);
                }, $"Фильтрация файлов по email: {string.Join(", ", emails)}");
                break;

            case "folder":
                if (parts.Length < 3) { Console.WriteLine("Использование: v folder <folder_link1> ..."); return; }
                var folderLinks = parts.Skip(2).ToArray();
                await ShowProgress(async () =>
                {
                    var allFolderFiles = new List<Google.Apis.Drive.v3.Data.File>();
                    foreach (var link in folderLinks)
                    {
                        var folderId = ExtractFolderId(link);
                        var folderFiles = await GetFolderFilesRecursive(service, folderId);
                        allFolderFiles.AddRange(folderFiles);
                    }
                    PrintPaged(allFolderFiles);
                }, "Загрузка файлов из папок");
                break;

            case "ignore":
                if (parts.Length < 3) { Console.WriteLine("Использование: v ignore <email1> ..."); return; }
                var ignoreEmails = parts.Skip(2).ToArray();
                await ShowProgress(async () => await CommandIgnoreEmail(service, ignoreEmails),
                    $"Фильтрация файлов (игнорируем): {string.Join(", ", ignoreEmails)}");
                break;

            default:
                Console.WriteLine("Неизвестная команда view. Введите 'help'");
                break;
        }
    }

    // ===================== Обработка команд xlsx =====================
    static async Task HandleExcel(DriveService service, string[] parts)
    {
        if (parts.Length < 3) { Console.WriteLine("Использование: xlsx disk/email/folder/ignore <args> <path_to_file.xlsx>"); return; }

        var subCommand = parts[1].ToLower();
        var path = parts.Last();
        List<Google.Apis.Drive.v3.Data.File> filesForExcel = new List<Google.Apis.Drive.v3.Data.File>();

        switch (subCommand)
        {
            case "disk":
                await ShowProgress(async () => filesForExcel = await GetAllFiles(service), "Загрузка всех файлов");
                break;

            case "email":
                var emails = parts.Skip(2).Take(parts.Length - 3).ToArray();
                await ShowProgress(async () =>
                {
                    var allFiles = await GetAllFiles(service);
                    filesForExcel = allFiles
                        .Where(f => f.Permissions != null && f.Permissions.Any(p => emails.Contains(p.EmailAddress)))
                        .ToList();
                }, $"Фильтрация файлов по email: {string.Join(", ", emails)}");
                break;

            case "folder":
                var folderLinks = parts.Skip(2).Take(parts.Length - 3).ToArray();
                await ShowProgress(async () =>
                {
                    foreach (var link in folderLinks)
                    {
                        var folderId = ExtractFolderId(link);
                        var folderFiles = await GetFolderFilesRecursive(service, folderId);
                        filesForExcel.AddRange(folderFiles);
                    }
                }, "Загрузка файлов из папок");
                break;

            case "ignore":
                var ignoreEmails = parts.Skip(2).Take(parts.Length - 3).ToArray();
                await ShowProgress(async () =>
                {
                    var allFiles = await GetAllFiles(service);
                    filesForExcel = allFiles
                        .Where(f => f.Permissions == null || !f.Permissions.Any(p => ignoreEmails.Contains(p.EmailAddress)))
                        .ToList();
                }, $"Фильтрация файлов (игнорируем): {string.Join(", ", ignoreEmails)}");
                break;

            default:
                Console.WriteLine("Неизвестная команда xlsx. Введите 'help'");
                return;
        }

        ExportToExcel(filesForExcel, path);
    }

    // ===================== Печать файлов в консоль =====================
    static void PrintPaged(List<Google.Apis.Drive.v3.Data.File> files, int pageSize = 20)
    {
        int page = 0;
        while (true)
        {
            var items = files.Skip(page * pageSize).Take(pageSize).ToList();
            if (!items.Any()) break;

            foreach (var file in items)
                PrintFile(file);

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("Нажмите ENTER для следующей страницы...");
            Console.ResetColor();
            Console.ReadLine();
            page++;
        }
    }

    static void PrintFile(Google.Apis.Drive.v3.Data.File file)
    {
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine(file.Name);

        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine(file.WebViewLink);

        if (file.Permissions != null)
        {
            foreach (var p in file.Permissions)
            {
                Console.ForegroundColor = p.Role switch
                {
                    "owner" => ConsoleColor.Red,
                    "writer" => ConsoleColor.Yellow,
                    "reader" => ConsoleColor.Green,
                    _ => ConsoleColor.White
                };
                Console.WriteLine($"  {p.EmailAddress} - {p.Role}");
            }
        }

        Console.WriteLine();
        Console.ResetColor();
    }

    // ===================== Команды ignore =====================
    static async Task CommandIgnoreEmail(DriveService service, string[] ignoreEmails)
    {
        Console.WriteLine($"Показываем все файлы кроме: {string.Join(", ", ignoreEmails)}\n");
        var files = await GetAllFiles(service);
        var filtered = files
            .Where(f => f.Permissions == null || !f.Permissions.Any(p => ignoreEmails.Contains(p.EmailAddress)))
            .ToList();
        PrintPaged(filtered);
    }

    // ===================== Экспорт в Excel =====================
    static void ExportToExcel(List<Google.Apis.Drive.v3.Data.File> files, string path)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("DriveAnalytic");

        ws.Cell(1, 1).Value = "File Name";
        ws.Cell(1, 2).Value = "Link";
        ws.Cell(1, 3).Value = "Email";
        ws.Cell(1, 4).Value = "Role";
        ws.Cell(1, 5).Value = "Type";

        int row = 2;
        foreach (var file in files)
        {
            if (file.Permissions != null && file.Permissions.Count > 0)
            {
                foreach (var p in file.Permissions)
                {
                    ws.Cell(row, 1).Value = file.Name;
                    ws.Cell(row, 2).Value = file.WebViewLink;
                    ws.Cell(row, 3).Value = p.EmailAddress;
                    ws.Cell(row, 4).Value = p.Role;
                    ws.Cell(row, 5).Value = p.Type;
                    row++;
                }
            }
            else
            {
                ws.Cell(row, 1).Value = file.Name;
                ws.Cell(row, 2).Value = file.WebViewLink;
                row++;
            }
        }

        wb.SaveAs(path);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Файл Excel сохранён: {path}");
        Console.ResetColor();
    }


}