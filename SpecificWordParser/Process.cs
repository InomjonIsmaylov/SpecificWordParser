using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SpecificWordParser;

public class Process
{
    private readonly int _id;
    private readonly List<Paragraph> _paragraphs;

    public Process(int id, List<Paragraph> paragraphs)
    {
        _id = id;
        _paragraphs = paragraphs;
    }

    public List<Translation> Run()
    {
        var foundIndexes = CollectIndexes();

        var translations = new List<Translation>(_paragraphs.Count / 4);

        ConvertToTranslations(foundIndexes, translations);

        var invalidOnes = GetValidAndInvalidOnes(translations, out var validOnes);

        CreateJsonFile(validOnes);

        if (invalidOnes.Any())
        {
            CreateJsonFile(invalidOnes, false);
        }

        PrintLine.Info($"Invalid ones: {invalidOnes.Count}, Valid ones: {validOnes.Count}");

        return translations;
    }

    private void CreateJsonFile(List<Translation> validOnes, bool valid = true)
    {
        var json = JsonSerializer.Serialize(validOnes, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
        });

        var fileName = valid ? $"valid{_id}.json" : $"invalid{_id}.json";

        PrintLine.Info("Генерируется файл " + fileName);

        File.WriteAllText(fileName, json);
    }

    private void ConvertToTranslations(IReadOnlyList<int> foundIndexes, ICollection<Translation> translations)
    {
        for (int i = 0; i < foundIndexes.Count - 1; i++)
        {
            var index = foundIndexes[i];
            var nextIndex = foundIndexes[i + 1];

            var chunk = _paragraphs.Skip(index).Take(nextIndex - index).ToList();

            translations.Add(new Translation
            {
                Header = chunk[0].InnerText.Trim(),
                English = chunk[1].InnerText.Trim(),
                Rus = chunk[2].InnerText.Trim(),
                Content = string.Join("\n", chunk.Skip(3).Select(x => x.InnerText.Trim()))
            });
        }
    }

    private List<int> CollectIndexes()
    {
        var limit = _paragraphs.Count - 3;

        var foundIndexes = new List<int>(limit / 4);

        for (int currentIndex = 1; currentIndex < limit; currentIndex++)
        {
            var index = FindIndexOfNextStart(currentIndex, limit);

            if (index == -1)
                break;

            foundIndexes.Add(index);
            currentIndex = index + 3;
        }

        return foundIndexes;
    }

    private int FindIndexOfNextStart(int index, int limit)
    {
        if (index > limit)
            throw new ArgumentOutOfRangeException(nameof(index));

        while (true)
        {
            if (index == limit)
                return -1;

            if (_paragraphs[index].InnerText.TrimStart().StartsWith("ing.") && _paragraphs[index + 1].InnerText.TrimStart().StartsWith("rus."))
            {
                return index - 1;
            }

            index++;
        }
    }

    private static List<Translation> GetValidAndInvalidOnes(IReadOnlyCollection<Translation> translations, out List<Translation> validOnes)
    {
        var invalidOnes = translations
            .Where(x =>
                x.Header is null || x.English is null || x.Rus is null || x.Content is null
                || !x.English.StartsWith("ing.")
                || !x.Rus.StartsWith("rus.")
            )
            .ToList();

        validOnes = translations
            .Where(x => !invalidOnes.Contains(x))
            .ToList();

        return invalidOnes;
    }
}