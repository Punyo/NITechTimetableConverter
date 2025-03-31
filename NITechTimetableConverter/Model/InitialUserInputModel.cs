using NITechTimetableConverter.Properties;
using NITechTimetableConverter.Utility;

namespace NITechTimetableConverter.Model
{
    internal class InitialUserInputModel
    {
        public string GetUserPathInput()
        {
            while (true)
            {
                Console.Write(Resources.MessageEnterPath);
                string? path = Console.ReadLine();
                if (string.IsNullOrEmpty(path) || !path.EndsWith(".xlsx"))
                {
                    ConsoleUtil.WriteLineWithColor(Resources.ErrorInputInvalidOrNonXLSX, ConsoleUtil.ColorMode.Error);
                    continue;
                }
                if (!File.Exists(path))
                {
                    ConsoleUtil.WriteLineWithColor(Resources.ErrorFileNotFound, ConsoleUtil.ColorMode.Error);
                    continue;
                }
                return path;
            }
        }

        public void ShowInitialMessage()
        {
            Console.WriteLine(Resources.MessageCredit);
            Console.WriteLine(Resources.MessageGitHubURL);
            Console.WriteLine(Resources.MessageXURL);
            Console.WriteLine(Resources.MessageDivider);
            ConsoleUtil.WriteLineWithColor(Resources.MessageWarning, ConsoleUtil.ColorMode.Error);
            ConsoleUtil.WriteLineWithColor(Resources.MessageCaution, ConsoleUtil.ColorMode.Warning);
            Console.WriteLine($"{Resources.MessageDivider}{Environment.NewLine}");
        }
    }
}
