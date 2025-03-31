using NITechTimetableConverter.Model;

internal class Program
{
    private static void Main(string[] args)
    {
        InitialUserInputModel model = new();
        model.ShowInitialMessage();
        string path = model.GetUserPathInput();
        GenerateTimetableFromDegradedFormatModel generator = new(path);
        generator.GenerateTimetable(path.ToLower().Replace(".xlsx","_変換済.xlsx"));
    }
}