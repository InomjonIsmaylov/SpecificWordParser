using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SpecificWordParser;

public static class Worker
{
    public static string? FilePath { get; set; }

    public static int BatchSize { get; set; } = 5000;

    public static async Task WorkAsync()
    {
        using var wordDocument = OpenFile();

        if (wordDocument.MainDocumentPart?.Document.Body is null)
            throw new ArgumentNullException(nameof(wordDocument), "Invalid Word Document");

        var count = wordDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().Count();

        var taskCount = count / BatchSize;

        PrintLine.Info($"Количество параграфов: {count}");
        PrintLine.Info($"Количество потоков: {taskCount}");
        PrintLine.Info($"Количество параграфов обрабатываемых каждым потоком: {BatchSize}");

        PrintLine.Info("Начало обработки\n\n...\n");
        
        var tasks = new Task[taskCount];

        SetProcessTasks(taskCount, wordDocument, BatchSize, tasks);

        await Task.Delay(1000);

        await Task.WhenAll(tasks);

        PrintLine.Info("\nОбработка завершена");

        wordDocument.Dispose();
    }

    private static WordprocessingDocument OpenFile()
    {
        var openSettings = new OpenSettings
        {
            RelationshipErrorHandlerFactory = _ => new UriRelationshipErrorHandler()
        };

        return WordprocessingDocument.Open(FilePath!, true, openSettings);
    }

    private static void SetProcessTasks(int taskCount, WordprocessingDocument wordDocument, int batchSize, IList<Task> tasks)
    {
        for (int i = 0; i < taskCount; i++)
        {
            var paragraphs = wordDocument.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Skip(i * batchSize)
                .Take(batchSize)
                .ToList();

            var process = new Process(i, paragraphs);

            tasks[i] = Task.Run(() => { process.Run(); });
        }
    }

    public static void GetPathFromUserInput()
    {
        Console.OutputEncoding = Console.InputEncoding = Encoding.UTF8;

        Console.Write("Укажите путь до файла (");
        Print.Success("нажмите Enter если файл data.docx находится в папке с приложением");
        Console.Write("): ");

        FilePath = Console.ReadLine();

        FilePath = FilePath!.Trim();

        FilePath = string.IsNullOrEmpty(FilePath) ? Path.Combine(Directory.GetCurrentDirectory(), "data.docx") : FilePath;
    }

    public static void GetBatchSizeFromUserInput()
    {
        Console.Write("Укажите сколько параграфов обработать каждому патоку: (");
        Print.Success("нажмите Enter чтобы оставить значение по умолчанию 5000");
        Console.Write("): ");

        var temp = Console.ReadLine();

        try
        {
            temp = temp!.Trim();
            if (string.IsNullOrEmpty(temp))
            {
                PrintLine.Info("Учитивается значение по умолчанию 5000");

                return;
            }

            _ = int.TryParse(temp, out var batchSize);

            if (batchSize >= 100)
            {
                Console.WriteLine();

                BatchSize = batchSize;

                return;
            }

            PrintLine.Error("Значение не может быть меньше 100");
            PrintLine.Info("Учитивается значение по умолчанию 5000");

            BatchSize = 5000;
        }
        catch
        {
            PrintLine.Error("Неверное значение");
            PrintLine.Info("Значение учитивается значение по умолчанию 5000");

            BatchSize = 5000;
        }
    }

    public static bool CheckFileExtension()
    {
        return FilePath!.EndsWith(".docx") || FilePath.EndsWith(".doc");
    }
}