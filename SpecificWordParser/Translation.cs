namespace SpecificWordParser;

public class Translation
{
    public string? Header { get; set; }

    public string? English { get; set; }

    public string? Rus { get; set; }

    public string? Content { get; set; }

    public override string ToString()
    {
        return $"Translation ({{{Header}}}, {{{English}}}, {{{Rus}}})";
    }
}