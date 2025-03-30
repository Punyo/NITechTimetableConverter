using ClosedXML.Excel;
using NITechTimetableConverter.Data;
using NITechTimetableConverter.Model;
using NITechTimetableConverter.Properties;
using System.Diagnostics;
using System.Resources;
using System.Text.RegularExpressions;

internal class Program
{
    private static void Main(string[] args)
    {
        var folder = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
        Console.WriteLine(Resources.MessageCredit);
        Console.WriteLine(Resources.MessageCaution);
        Console.WriteLine($"{Resources.MessageDivider}{Environment.NewLine}");
        GenerateTimetableFromDegradedFormatModel model = new($@"{folder}\Assets\b3.xlsx");
        model.GenerateTimetable($@"{folder}\Assets\b3_result.xlsx");
        ProcessStartInfo result = new()
        {
            FileName = $@"{folder}\Assets\b3_result.xlsx",
            UseShellExecute = true
        };
        Process.Start(result);
    }
}