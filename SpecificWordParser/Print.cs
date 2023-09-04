namespace SpecificWordParser;

public static class PrintLine
{
    public static void Success(string message)
    {
        Result(message, ConsoleColor.Green);
    }

    public static void Info(string message)
    {
        Result(message, ConsoleColor.Blue);
    }

    public static void Error(string message)
    {
        Result(message, ConsoleColor.Red);
    }

    public static void Result(string message, ConsoleColor color)
    {
        Console.ForegroundColor = color;

        Console.WriteLine(message);

        Console.ResetColor();
    }
}

public static class Print
{
    public static void Success(string message)
    {
        Result(message, ConsoleColor.Green);
    }

    public static void Info(string message)
    {
        Result(message, ConsoleColor.Blue);
    }

    public static void Error(string message)
    {
        Result(message, ConsoleColor.Red);
    }

    public static void Result(string message, ConsoleColor color)
    {
        Console.ForegroundColor = color;

        Console.Write(message);

        Console.ResetColor();
    }
}