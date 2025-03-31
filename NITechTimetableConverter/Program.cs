using NITechTimetableConverter.Model;
using NITechTimetableConverter.Properties;
using NITechTimetableConverter.Utility;

internal class Program
{
    private static void Main(string[]? args)
    {
        InitialUserInputModel model = new();
        model.ShowInitialMessage();
        string path = model.GetUserPathInput();
        GenerationProcess(path);
    }

    private static void GenerationProcess(string path)
    {
        try
        {
            GenerateTimetableFromDegradedFormatModel generator = new(path);
            generator.GenerateTimetable(path.ToLower().Replace(".xlsx", "_変換済.xlsx"));
        }
        catch (IOException)
        {
            ConsoleUtil.WriteLineWithColor(Resources.ErrorIOException, ConsoleUtil.ColorMode.Error);
            Console.WriteLine(Resources.MessageRetry);
            while (Console.ReadKey().Key != ConsoleKey.Enter) { }
            GenerationProcess(path);
        }
        catch (Exception e)
        {
            ConsoleUtil.WriteLineWithColor(string.Format(Resources.ErrorGeneric, e.Message), ConsoleUtil.ColorMode.Error);
            Console.WriteLine(Resources.MessageRetry);
            while (Console.ReadKey().Key != ConsoleKey.Enter) { }
            Main(Array.Empty<string>());
        }
    }
}