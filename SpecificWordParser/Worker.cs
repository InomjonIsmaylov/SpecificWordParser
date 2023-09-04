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

        var tasks = new Task[taskCount];

        SetProcessTasks(taskCount, wordDocument, BatchSize, tasks);

        await Task.WhenAll(tasks);

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

    //private static string GetFilepath(bool test = true)
    //{
    //    var fileName = test ? "test.docx" : "data.docx";

    //    return Path.Combine(@"C:\Users\i_ismaylov\source\repos\SpecificWordParser\SpecificWordParser\wwwroot", fileName);

    //    //return "";
    //}

    public static void GetPathFromUserInput()
    {
        Console.OutputEncoding = Console.InputEncoding = Encoding.UTF8;

        Console.Write("Укажите путь до файла (");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("нажмите Enter если файл data.docx находится в папке с приложением");
        Console.ResetColor();
        Console.Write("): ");

        FilePath = Console.ReadLine();

        FilePath = FilePath!.Trim();

        FilePath = string.IsNullOrEmpty(FilePath) ? Path.Combine(Directory.GetCurrentDirectory(), "data.docx") : FilePath;
    }

    public static void GetBatchSizeFromUserInput()
    {
        Console.Write("Укажите сколько параграфов обработать каждому патоку: (");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("нажмите Enter чтобы оставить значение по умолчанию 5000");
        Console.ResetColor();
        Console.Write("): ");

        var temp = Console.ReadLine();

        try
        {
            temp = temp!.Trim();
            _ = int.TryParse(temp, out var batchSize);
            BatchSize = batchSize > 99 ? batchSize : 5000;
        }
        catch
        {
            BatchSize = 5000;
        }
    }
}