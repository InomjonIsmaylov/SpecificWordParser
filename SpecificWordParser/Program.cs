using SpecificWordParser;

Worker.GetPathFromUserInput();


if (!Path.Exists(Worker.FilePath))
{
    PrintLine.Error("Файл не найден");
    return;
}

if (!Worker.CheckFileExtension())
{
    PrintLine.Error("Неверное расширение файла");
    return;
}

PrintLine.Info($"Путь до файла: {Worker.FilePath}");

Worker.GetBatchSizeFromUserInput();

await Worker.WorkAsync();

PrintLine.Info("Работа завершена");

Console.ReadKey();