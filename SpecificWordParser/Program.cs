using SpecificWordParser;

Worker.GetPathFromUserInput();

if (!Path.Exists(Worker.FilePath))
{
    var hasDocExtension = Path.HasExtension("docx") || Path.HasExtension("doc");

    if (hasDocExtension)
    {
        Console.WriteLine("Указанный путь не является файлом doc или docx");
        return;
    }

    var newPath = Path.Combine(Worker.FilePath!, "data.docx");
    var hasDataDocxOnProvidedPath = Path.Exists(newPath);

    if (hasDataDocxOnProvidedPath)
    {
        Worker.FilePath = newPath;
    }
    else
    {
        Console.WriteLine("Файл не найден");
    }
}

Console.WriteLine($"Путь до файла: {Worker.FilePath}");

await Worker.WorkAsync();